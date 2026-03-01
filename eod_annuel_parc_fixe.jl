
include("extraction_donnees_excel.jl")
using JuMP
using HiGHS
using Dates
using JSON3

data_file = "data/Donnees_etude_de_cas_ETE305.xlsx"
parc_de_prod = "results/annual/parc_annuel.json"

data = open(parc_de_prod, "r") do f
    str = read(f, String)
    JSON3.read(str)
end

# -------- Extraction des hypothèses du problèmes --------
config = extraire_donnees_config(data_file)

PU_cost_h2_CCG = config["H2"]["CCG"]["PU_cost"] #€/MWh basé sur le tarif de prod de la centrale CCG gaz A MODIFIER
Pmin_CCG_h2 = config["H2"]["CCG"]["Pmin"] #MW idem
Pmax_CCG_h2 = config["H2"]["CCG"]["Pmax"] #MW idem
dmin_CCG = config["H2"]["CCG"]["dmin"]#hours idem

PU_cost_h2_TAC = config["H2"]["TAC"]["PU_cost"] #€/MWh basé sur le tarif de prod de la centrale TAC gaz A MODIFIER
Pmin_TAC_h2 = config["H2"]["TAC"]["Pmin"] #MW idem
Pmax_TAC_h2 = config["H2"]["TAC"]["Pmax"] #MW idem
dmin_TAC = config["H2"]["TAC"]["dmin"] #hours idem

RendementElectrolyse = XLSX.readdata(data_file, "Rendements", "B8") # Rendement de l'électrolyse
RendementCombustion = XLSX.readdata(data_file, "Rendements", "B11") # Rendement de la combustion de l'hydrogène

# Hydro
Pmin_hy_lacs = 0
Pmax_hy_lacs = config["hydro"]["lacs"]["Pmax"] #MW
Pmax_STEP = XLSX.readdata(data_file, "Parc électrique", "C21") #MW
rSTEP = XLSX.readdata(data_file, "Rendements", "B10") #rendement au pompage 

#battery
rbattery = config["battery"]["rendement"] # rendement de la batterie au sockage ou au destockage
d_battery = config["battery"]["d_battery"] #hours

# Defailance
cuns = config["defaillance"]["cost_unsupplied"]  #cost of unsupplied energy €/MWh
cexc = config["defaillance"]["cost_excess"] #cost of in excess energy €/MWh

# ----------------- Définition des variables annuelles -----------------
# Nombre de semaines et d'heures totales
Nweeks = 52
Nhours_per_week = 7*24
Nhours = Nweeks * Nhours_per_week

# Tableaux horaires annuels pour stocker les résultats de dispatch
solar_annual      = zeros(Nhours)
offshore_annual   = zeros(Nhours)
onshore_annual    = zeros(Nhours)
load_annual       = zeros(Nhours)
battery_stock_annual   = zeros(Nhours)
battery_charge_annual  = zeros(Nhours)
battery_discharge_annual = zeros(Nhours)
STEP_stock_annual      = zeros(Nhours)
STEP_charge_annual     = zeros(Nhours)
STEP_discharge_annual  = zeros(Nhours)

# Variables H2
NH2_CCG = data["H2"]["CCG"]["nombre_installees"]
NH2_TAC = data["H2"]["TAC"]["nombre_installees"]
CCG_H2_running_annual = zeros(Nhours, NH2_CCG)
TAC_H2_running_annual = zeros(Nhours, NH2_TAC)
PH2_CCG_annual        = zeros(Nhours, NH2_CCG)
PH2_TAC_annual        = zeros(Nhours, NH2_TAC)

# Hydro / defaillance / excès / hydro et thermique fatal
Phy_annual  = zeros(Nhours)
Puns_annual = zeros(Nhours)
Pexc_annual = zeros(Nhours)
Pres_annual = zeros(Nhours)

# Conditions initiales pour la première semaine
global stock_battery_initial = 0
global stock_STEP_initial    = 0

global installed_CCG_H2 = zeros(Int, NH2_CCG)
global installed_TAC_H2 = zeros(Int, NH2_TAC)

# Capacités
CapaSolar = data["capacites_MW"]["solar"]
CapaOnshore = data["capacites_MW"]["onshore"]
CapaOffshore = data["capacites_MW"]["offshore"]
CapaBattery = data["capacites_MW"]["battery"]

# Initialize installed H2 units from JSON data
for i in 1:NH2_CCG
    if i <= data["H2"]["CCG"]["nombre_installees"]
        installed_CCG_H2[i] = 1
    end
end

for i in 1:NH2_TAC
    if i <= data["H2"]["TAC"]["nombre_installees"]
        installed_TAC_H2[i] = 1
    end
end

LAST_WEEK = 51

for week in 1:LAST_WEEK
    t_start = now()
    println("Début de l'EOD semaine $week à $t_start ...")

    time_series = extraire_donnees_semaine(data_file, week, AnneauGarde=24)

    Tmax = time_series["Tmax"]

    load = time_series["load"]
    offshore_load_factor = time_series["offshore_load_factor"]
    onshore_load_factor = time_series["onshore_load_factor"]
    solar_load_factor = time_series["solar_load_factor"]
    hydro_fatal = time_series["hydro_fatal"]
    e_hy_lacs = time_series["Usable_per_week_hydro_lacs"]
    thermique_fatal = time_series["thermique_fatal"]

    Pres = hydro_fatal + thermique_fatal

    ########## Defining model ##########
    model = Model(HiGHS.Optimizer)
    set_optimizer_attribute(model, "mip_rel_gap", 0.005) # S'arrête à 0.5%

    ########## Defining variables ##########
    #H2 generation variables
    #CCG H2
    @variable(model, CCG_H2_running[1:Tmax, 1:NH2_CCG], Bin)         # 1 si ON à t
    @variable(model, PH2_CCG[1:Tmax, 1:NH2_CCG] >= 0)

    @variable(model, CCG_H2_start[1:Tmax,1:NH2_CCG], Bin)
    @variable(model, CCG_H2_stop[1:Tmax,1:NH2_CCG], Bin)

    #TAC H2
    @variable(model, TAC_H2_running[1:Tmax, 1:NH2_TAC], Bin)         # 1 si ON à t
    @variable(model, PH2_TAC[1:Tmax, 1:NH2_TAC] >= 0)

    @variable(model, TAC_H2_start[1:Tmax,1:NH2_TAC], Bin)
    @variable(model, TAC_H2_stop[1:Tmax,1:NH2_TAC], Bin)

    #hydro generation variables
    @variable(model, Phy[1:Tmax] >= 0)
    #unsupplied energy variables
    @variable(model, Puns[1:Tmax] >= 0)
    #in excess energy variables
    @variable(model, Pexc[1:Tmax] >= 0)
    #weekly STEP variables
    @variable(model, Pcharge_STEP[1:Tmax] >= 0)
    @variable(model, Pdecharge_STEP[1:Tmax] >= 0)
    @variable(model, stock_STEP[1:Tmax] >= 0)
    # #battery variables
    @variable(model, Pcharge_battery[1:Tmax] >= 0)
    @variable(model, Pdecharge_battery[1:Tmax] >= 0)
    @variable(model, stock_battery[1:Tmax] >= 0)


    for g in 1:NH2_TAC
        @constraint(model, TAC_H2_running[1,g] - TAC_H2_running_initial[g] == TAC_H2_start[1,g] - TAC_H2_stop[1,g])
    end

    ####### Defining objective function #######
    @objective(model, Min, (sum(PH2_CCG[t,g] for t in 1:Tmax, g in 1:NH2_CCG)*PU_cost_h2_CCG + sum(PH2_TAC[t,g] for t in 1:Tmax, g in 1:NH2_TAC)*PU_cost_h2_TAC + sum(Puns[t] for t in 1:Tmax)*cuns + sum(Pexc[t] for t in 1:Tmax)*cexc)/1e4)

    ######## Defining constraints ########
    # Initial conditions
    @constraint(model, stock_battery[1] == stock_battery_initial)
    @constraint(model, stock_STEP[1] == stock_STEP_initial)

    #balance constraint
    @constraint(model, balance[t in 1:Tmax], sum(PH2_CCG[t,g] for g in 1:NH2_CCG) + sum(PH2_TAC[t,g] for g in 1:NH2_TAC) + Phy[t] + Pres[t] + CapaSolar * solar_load_factor[t] + CapaOffshore * offshore_load_factor[t] + CapaOnshore * onshore_load_factor[t] - Pcharge_STEP[t] + Pdecharge_STEP[t] - Pcharge_battery[t] + Pdecharge_battery[t] + Puns[t] - load[t] - Pexc[t]  == 0)

    # H2 Power constraints
    @constraint(model, max_CCG_H2[t in 1:Tmax, i in 1:NH2_CCG], PH2_CCG[t,i] <= Pmax_CCG_h2*CCG_H2_running[t,i]) #Pmax constraints
    @constraint(model, min_CCG_H2[t in 1:Tmax, i in 1:NH2_CCG], Pmin_CCG_h2*CCG_H2_running[t,i] <= PH2_CCG[t,i]) #Pmin constraints

    @constraint(model, max_TAC_H2[t in 1:Tmax, i in 1:NH2_TAC], PH2_TAC[t,i] <= Pmax_TAC_h2*TAC_H2_running[t,i]) #Pmax constraints
    @constraint(model, min_TAC_H2[t in 1:Tmax, i in 1:NH2_TAC], Pmin_TAC_h2*TAC_H2_running[t,i] <= PH2_TAC[t,i]) #Pmin constraints

    # H2 instalation constraints
    @constraint(model, [t in 1:Tmax, g in 1:NH2_CCG], CCG_H2_running[t,g] <= installed_CCG_H2[g]) # only produce if installed
    @constraint(model, [t in 1:Tmax, g in 1:NH2_TAC], TAC_H2_running[t,g] <= installed_TAC_H2[g]) # only produce if installed

    # H2 duration constraints
    for g in 1:NH2_CCG
            if (dmin_CCG > 1)
                @constraint(model, [t in 2:Tmax], CCG_H2_running[t,g]-CCG_H2_running[t-1,g]==CCG_H2_start[t,g]-CCG_H2_stop[t,g],  base_name = "stateH2_$g") # detect start and stop
                @constraint(model, [t in 1:Tmax], CCG_H2_start[t,g]+CCG_H2_stop[t,g]<=1,  base_name = "exclusiveH2_$g") # avoid starting and stoping at same step

                # Initial conditions
                @constraint(model, CCG_H2_start[1,g]==0,  base_name = "iniStartH2_$g")
                @constraint(model, CCG_H2_stop[1,g]==0,  base_name = "iniStopH2_$g")
                @constraint(model, [t in 1:dmin_CCG-1], CCG_H2_running[t,g] >= sum(CCG_H2_start[i,g] for i in 1:t), base_name = "dminStartH2_$(g)_init")
                @constraint(model, [t in 1:dmin_CCG-1], CCG_H2_running[t,g] <= 1-sum(CCG_H2_stop[i,g] for i in 1:t), base_name = "dminStopH2_$(g)_init")

                # Minimum up and down time constraints
                @constraint(model, [t in dmin_CCG:Tmax], CCG_H2_running[t,g] >= sum(CCG_H2_start[i,g] for i in (t-dmin_CCG+1):t),  base_name = "dminStartH2_$g")
                @constraint(model, [t in dmin_CCG:Tmax], CCG_H2_running[t,g] <= 1 - sum(CCG_H2_stop[i,g] for i in (t-dmin_CCG+1):t),  base_name = "dminStopH2_$g")

            end
        end

    # H2 volume constraints
    @constraint(model, sum(PH2_CCG[t,g] for t in 1:Tmax, g in 1:NH2_CCG) + sum(PH2_TAC[t,g] for t in 1:Tmax, g in 1:NH2_TAC) <= sum(Pexc[t] for t in 1:Tmax)*RendementCombustion*RendementElectrolyse)

    # hydro unit constraints
    @constraint(model, [t=1:Tmax], Phy[t] >= Pmin_hy_lacs)
    @constraint(model, [t=1:Tmax], Phy[t] <= Pmax_hy_lacs)
    # hydro stock constraint
    @constraint(model, sum(Phy[t] for t in 1:Tmax) <= e_hy_lacs)

    # weekly STEP
    @constraint(model, [t in 1:Tmax], Pcharge_STEP[t] <= Pmax_STEP)
    @constraint(model, [t in 1:Tmax], Pdecharge_STEP[t] <= Pmax_STEP)
    @constraint(model, stock_STEP[1] == stock_STEP_initial)
    @constraint(model, Pdecharge_STEP[Tmax] <= stock_STEP[Tmax])
    #@constraint(model, stock_STEP[Tmax] == stock_STEP[1])
    #@constraint(model, Pdecharge_STEP[1] == 0)
    @constraint(model, [t in 1:Tmax-1], stock_STEP[t+1]-stock_STEP[t]- rSTEP*Pcharge_STEP[t]+Pdecharge_STEP[t]== 0)
    @constraint(model, [t in 1:Tmax], stock_STEP[t] <= 24*7*Pmax_STEP)

    # #battery
    @constraint(model, [t in 1:Tmax], Pcharge_battery[t] <= CapaBattery)
    @constraint(model, [t in 1:Tmax], Pdecharge_battery[t] <= CapaBattery)
    @constraint(model, stock_battery[1] == stock_battery_initial)
    @constraint(model, Pdecharge_battery[Tmax] <= stock_battery[Tmax])
    #@constraint(model, stock_battery[Tmax] == stock_battery[1])
    #@constraint(model, Pdecharge_battery[1] == 0)
    @constraint(model, [t in 1:Tmax-1], stock_battery[t+1]-stock_battery[t]- rbattery*Pcharge_battery[t]+1/rbattery*Pdecharge_battery[t]== 0)
    @constraint(model, [t in 1:Tmax], stock_battery[t] <= d_battery*CapaBattery)

    optimize!(model)

    #Results
    # @show termination_status(model)
    # @show objective_value(model)


    
    # --- Stockage des résultats horaires dans les tableaux annuels ---
    t_start = (week-1)*Nhours_per_week + 1
    t_end   = week*Nhours_per_week
    idx = t_start:t_end
    
    # Exclure l'anneau de garde pour sauvegarder
    # @views eviter de faire des copies inutiles lors de l'affectation dans les tableaux annuels
    @views solar_annual[idx]    .= value(CapaSolar) .* solar_load_factor[1:Nhours_per_week]
    @views onshore_annual[idx]  .= value(CapaOnshore) .* onshore_load_factor[1:Nhours_per_week]
    @views offshore_annual[idx] .= value(CapaOffshore) .* offshore_load_factor[1:Nhours_per_week]

    @views battery_stock_annual[idx]     .= value.(stock_battery[1:Nhours_per_week])
    @views battery_charge_annual[idx]    .= value.(Pcharge_battery[1:Nhours_per_week])
    @views battery_discharge_annual[idx] .= value.(Pdecharge_battery[1:Nhours_per_week])

    @views STEP_stock_annual[idx]        .= value.(stock_STEP[1:Nhours_per_week])
    @views STEP_charge_annual[idx]       .= value.(Pcharge_STEP[1:Nhours_per_week])
    @views STEP_discharge_annual[idx]    .= value.(Pdecharge_STEP[1:Nhours_per_week])

    @views PH2_CCG_annual[idx, :]        .= value.(PH2_CCG[1:Nhours_per_week, :])
    @views PH2_TAC_annual[idx, :]        .= value.(PH2_TAC[1:Nhours_per_week, :])
    @views CCG_H2_running_annual[idx, :] .= value.(CCG_H2_running[1:Nhours_per_week, :])
    @views TAC_H2_running_annual[idx, :] .= value.(TAC_H2_running[1:Nhours_per_week, :])

    @views Phy_annual[idx]  .= value.(Phy[1:Nhours_per_week])
    @views Puns_annual[idx] .= value.(Puns[1:Nhours_per_week])
    @views Pexc_annual[idx] .= value.(Pexc[1:Nhours_per_week])
    @views Pres_annual[idx] .= Pres[1:Nhours_per_week]
    
    @views load_annual[idx] .= value.(load[1:Nhours_per_week])
    
    # --- Stock initial pour la semaine suivante ---
    global stock_battery_initial = value(stock_battery[Tmax])
    global stock_STEP_initial    = value(stock_STEP[Tmax])

    last_CCG_H2_running = value.(CCG_H2_running[Tmax, :])
    last_TAC_H2_running = value.(TAC_H2_running[Tmax, :])

    t_end = now()
    println("Fin de l'EOD semaine $week, durée : $(t_end)")
    
end

# --- Export CSV annuel (comme dans ton code) ---
open("results/annual_fixed_parc/results.csv", "w") do f
    write(f, "t;Solar;Onshore;Offshore;Battery_stock;Battery_charge;Battery_discharge;STEP_stock;STEP_charge;STEP_discharge;H2_CCG;H2_TAC;Hydro;hy_th_fatal;Load;Defailance;Exces\n")
    for t in 1:Nhours
        write(f,
            string(
                t, ";",
                round(solar_annual[t], digits=2), ";",
                round(onshore_annual[t], digits=2), ";",
                round(offshore_annual[t], digits=2), ";",
                round(battery_stock_annual[t], digits=2), ";",
                round(battery_charge_annual[t], digits=2), ";",
                round(battery_discharge_annual[t], digits=2), ";",
                round(STEP_stock_annual[t], digits=2), ";",
                round(STEP_charge_annual[t], digits=2), ";",
                round(STEP_discharge_annual[t], digits=2), ";",
                round(sum(PH2_CCG_annual[t, :]), digits=2), ";",
                round(sum(PH2_TAC_annual[t, :]), digits=2), ";",
                round(Phy_annual[t], digits=2), ";",
                round(Pres_annual[t], digits=2), ";",
                round(load_annual[t], digits=2), ";",
                round(Puns_annual[t], digits=2), ";",
                round(Pexc_annual[t], digits=2),
                "\n"
            )
        )
    end
end