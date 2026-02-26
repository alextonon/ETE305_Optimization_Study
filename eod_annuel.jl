include("extraction_donnees_excel.jl")
using JuMP
using HiGHS
using Dates

data_file = "data/Donnees_etude_de_cas_ETE305.xlsx"

# -------- Extraction des hypothèses du problèmes --------
config = extraire_donnees_config(data_file)

# Data for h2 clusters
capex_CCG_H2 = config["H2"]["CCG"]["capex"] #€/MW
opex_CCG_H2 = config["H2"]["CCG"]["opex"] #€/MW/year
PU_cost_h2_CCG = config["H2"]["CCG"]["PU_cost"] #€/MWh basé sur le tarif de prod de la centrale CCG gaz A MODIFIER
NH2_CCG_max = config["H2"]["CCG"]["gisement"] # Nombre de centrales CCG H2 disponibles
Pmin_CCG_h2 = config["H2"]["CCG"]["Pmin"] #MW idem
Pmax_CCG_h2 = config["H2"]["CCG"]["Pmax"] #MW idem
dmin_CCG = config["H2"]["CCG"]["dmin"]#hours idem

capex_TAC_H2 = config["H2"]["TAC"]["capex"] #€/MW
opex_TAC_H2 = config["H2"]["TAC"]["opex"] #€/MW/year
PU_cost_h2_TAC = config["H2"]["TAC"]["PU_cost"] #€/MWh basé sur le tarif de prod de la centrale TAC gaz A MODIFIER
Pmin_TAC_h2 = config["H2"]["TAC"]["Pmin"] #MW idem
Pmax_TAC_h2 = config["H2"]["TAC"]["Pmax"] #MW idem
dmin_TAC = config["H2"]["TAC"]["dmin"] #hours idem
NH2_TAC_max = config["H2"]["TAC"]["gisement"] # Nombre de centrales TAC H2 disponibles

# Renewables
capex_onshore = config["enr"]["onshore"]["capex"] #€/MW
capex_offshore = config["enr"]["offshore_pose"]["capex"] #€/MW
capex_solar = config["enr"]["pv_pose"]["capex"] #€/MW

opex_onshore = config["enr"]["onshore"]["opex"] #€/MW
opex_offshore = config["enr"]["offshore_pose"]["opex"] #€/MW
opex_solar = config["enr"]["pv_pose"]["opex"] #€/MW

CapaSolar_max = config["enr"]["gisement"]["Solar"] +70000#MW
CapaOnshore_max = config["enr"]["gisement"]["Onshore"] + 20000 #MW
CapaOffshore_max = config["enr"]["gisement"]["Offshore"] + 20000#MW


# Hydro
Pmin_hy_lacs = 0
Pmax_hy_lacs = config["hydro"]["lacs"]["Pmax"] #MW
Pmax_STEP = XLSX.readdata(data_file, "Parc électrique", "C21") #MW
rSTEP = XLSX.readdata(data_file, "Rendements", "B10") #rendement au pompage 

#battery
capex_2h_battery = config["battery"]["capex"] #€/MW
opex_2h_battery = config["battery"]["opex"] #€/MW
rbattery = config["battery"]["rendement"] # rendement de la batterie au sockage ou au destockage
d_battery = config["battery"]["d_battery"] #hours

# Defailance
cuns = config["defaillance"]["cost_unsupplied"]  #cost of unsupplied energy €/MWh
cexc = config["defaillance"]["cost_excess"] #cost of in excess energy €/MWh

# Initial capacities 
CapaSolar_init = config["capacites_init"]["Solar"] #MW
CapaOffshore_init = config["capacites_init"]["Offshore"] #MW
CapaOnshore_init = config["capacites_init"]["Onshore"] #MW
CapaBattery_max = config["battery"]["gisement"] #MW
CapaBattery_init = 0 #MW

# ----------------- Définition des variables annuelles -----------------
# Nombre de semaines et d'heures totales
Nweeks = 52
Nhours_per_week = 7*24
Nhours = Nweeks * Nhours_per_week

# Capacités initiales par technologie
solar_capacities = fill(CapaSolar_init, Nweeks+1)
offshore_capacities = fill(CapaOffshore_init, Nweeks+1)
onshore_capacities = fill(CapaOnshore_init, Nweeks+1)
battery_capacities = fill(CapaBattery_init, Nweeks+1)
installed_CCG_H2 = fill(0, Nweeks+1)
installed_TAC_H2 = fill(0, Nweeks+1)

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
CCG_H2_running_annual = zeros(Nhours, NH2_CCG_max)
TAC_H2_running_annual = zeros(Nhours, NH2_TAC_max)
PH2_CCG_annual        = zeros(Nhours, NH2_CCG_max)
PH2_TAC_annual        = zeros(Nhours, NH2_TAC_max)

# Hydro / defaillance / excès
Phy_annual  = zeros(Nhours)
Puns_annual = zeros(Nhours)
Pexc_annual = zeros(Nhours)

# Conditions initiales pour la première semaine
global stock_battery_initial = 0
global stock_STEP_initial    = 0

global CCG_H2_running_initial = zeros(Int, NH2_CCG_max)
global TAC_H2_running_initial = zeros(Int, NH2_TAC_max)

global CCG_H2_installed_initial = zeros(Int, NH2_CCG_max)
global TAC_H2_installed_initial = zeros(Int, NH2_TAC_max)

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
    #energie renouvelables
    @variable(model, onshore_capacities[week] <= CapaOnshore <= CapaOnshore_max, start = onshore_capacities[week])
    @variable(model, offshore_capacities[week] <= CapaOffshore <= CapaOffshore_max, start = offshore_capacities[week])
    @variable(model, solar_capacities[week] <= CapaSolar <= CapaSolar_max, start = solar_capacities[week])

    #H2 generation variables
    #CCG H2
    @variable(model, CCG_H2_installed[1:NH2_CCG_max], Bin)                 # 1 si la centrale i est construite
    @variable(model, CCG_H2_running[1:Tmax, 1:NH2_CCG_max], Bin)         # 1 si ON à t
    @variable(model, PH2_CCG[1:Tmax, 1:NH2_CCG_max] >= 0)

    @variable(model, CCG_H2_start[1:Tmax,1:NH2_CCG_max], Bin)
    @variable(model, CCG_H2_stop[1:Tmax,1:NH2_CCG_max], Bin)

    #TAC H2
    @variable(model, TAC_H2_installed[1:NH2_TAC_max], Bin)                 # 1 si la centrale i est construite
    @variable(model, TAC_H2_running[1:Tmax, 1:NH2_TAC_max], Bin)         # 1 si ON à t
    @variable(model, PH2_TAC[1:Tmax, 1:NH2_TAC_max] >= 0)

    @variable(model, TAC_H2_start[1:Tmax,1:NH2_TAC_max], Bin)
    @variable(model, TAC_H2_stop[1:Tmax,1:NH2_TAC_max], Bin)

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
    @variable(model, battery_capacities[week] <= CapaBattery <= CapaBattery_max, start = battery_capacities[week])
    @variable(model, Pcharge_battery[1:Tmax] >= 0)
    @variable(model, Pdecharge_battery[1:Tmax] >= 0)
    @variable(model, stock_battery[1:Tmax] >= 0)


    set_start_value.(CapaOnshore, onshore_capacities[week])
    set_start_value.(CapaOffshore, offshore_capacities[week])
    set_start_value.(CapaSolar, solar_capacities[week])
    set_start_value.(CapaBattery, battery_capacities[week])

    for g in 1:NH2_CCG_max
        if CCG_H2_installed_initial[g] > 0.5
            @constraint(model, CCG_H2_installed[g] == 1) # Déjà construite !
        end
    end

    for g in 1:NH2_TAC_max
        if TAC_H2_installed_initial[g] > 0.5
            @constraint(model, TAC_H2_installed[g] == 1) # Déjà construite !
        end
    end


    for g in 1:NH2_TAC_max
        @constraint(model, TAC_H2_running[1,g] - TAC_H2_running_initial[g] == TAC_H2_start[1,g] - TAC_H2_stop[1,g])
    end

    ####### Defining objective function #######
    @objective(model, Min, 
                        (CapaSolar*(capex_solar + opex_solar) + CapaOffshore*(capex_offshore + opex_offshore) + CapaOnshore*(capex_onshore + opex_onshore) 
                        + CapaBattery*(capex_2h_battery + opex_2h_battery) 
                        + sum(CCG_H2_installed[g]*Pmax_CCG_h2 for g in 1:NH2_CCG_max)*(capex_CCG_H2 + opex_CCG_H2) + sum(PH2_CCG[t,g] for t in 1:Tmax, g in 1:NH2_CCG_max)*PU_cost_h2_CCG
                        + sum(TAC_H2_installed[g]*Pmax_TAC_h2 for g in 1:NH2_TAC_max)*(capex_TAC_H2 + opex_TAC_H2) + sum(PH2_TAC[t,g] for t in 1:Tmax, g in 1:NH2_TAC_max)*PU_cost_h2_TAC
                        + sum(Puns[t] for t in 1:Tmax)*cuns + sum(Pexc[t] for t in 1:Tmax)*cexc)/1e4
    )

    ######## Defining constraints ########
    # Initial conditions
    @constraint(model, stock_battery[1] == stock_battery_initial)
    @constraint(model, stock_STEP[1] == stock_STEP_initial)

    #balance constraint
    @constraint(model, balance[t in 1:Tmax], sum(PH2_CCG[t,g] for g in 1:NH2_CCG_max) + sum(PH2_TAC[t,g] for g in 1:NH2_TAC_max) + Phy[t] + CapaSolar * solar_load_factor[t] + CapaOffshore * offshore_load_factor[t] + CapaOnshore * onshore_load_factor[t] + Puns[t] - load[t] - Pexc[t] - Pcharge_STEP[t] + Pdecharge_STEP[t] - Pcharge_battery[t] + Pdecharge_battery[t] == 0)

    # H2 Power constraints
    @constraint(model, max_CCG_H2[t in 1:Tmax, i in 1:NH2_CCG_max], PH2_CCG[t,i] <= Pmax_CCG_h2*CCG_H2_running[t,i]) #Pmax constraints
    @constraint(model, min_CCG_H2[t in 1:Tmax, i in 1:NH2_CCG_max], Pmin_CCG_h2*CCG_H2_running[t,i] <= PH2_CCG[t,i]) #Pmin constraints

    @constraint(model, max_TAC_H2[t in 1:Tmax, i in 1:NH2_TAC_max], PH2_TAC[t,i] <= Pmax_TAC_h2*TAC_H2_running[t,i]) #Pmax constraints
    @constraint(model, min_TAC_H2[t in 1:Tmax, i in 1:NH2_TAC_max], Pmin_TAC_h2*TAC_H2_running[t,i] <= PH2_TAC[t,i]) #Pmin constraints

    # H2 instalation constraints
    @constraint(model, [t in 1:Tmax, g in 1:NH2_CCG_max], CCG_H2_running[t,g] <= CCG_H2_installed[g]) # only produce if installed
    @constraint(model, [t in 1:Tmax, g in 1:NH2_TAC_max], TAC_H2_running[t,g] <= TAC_H2_installed[g]) # only produce if installed

    # H2 duration constraints
    for g in 1:NH2_CCG_max
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
    RendementElectrolyse = 0.7 # Rendement de l'électrolyse
    RendementCombustion = 0.5 # Rendement de la combustion de l'hydrogène
    @constraint(model, sum(PH2_CCG[t,g] for t in 1:Tmax, g in 1:NH2_CCG_max) + sum(PH2_TAC[t,g] for t in 1:Tmax, g in 1:NH2_TAC_max) <= sum(Pexc[t] for t in 1:Tmax)*RendementCombustion*RendementElectrolyse)

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
    @constraint(model, stock_STEP[Tmax] == stock_STEP[1])
    @constraint(model, Pdecharge_STEP[1] == 0)
    @constraint(model, [t in 1:Tmax-1], stock_STEP[t+1]-stock_STEP[t]- rSTEP*Pcharge_STEP[t]+Pdecharge_STEP[t]== 0)
    @constraint(model, [t in 1:Tmax], stock_STEP[t] <= 24*7*Pmax_STEP)

    # #battery
    @constraint(model, [t in 1:Tmax], Pcharge_battery[t] <= CapaBattery)
    @constraint(model, [t in 1:Tmax], Pdecharge_battery[t] <= CapaBattery)
    @constraint(model, stock_battery[1] == stock_battery_initial)
    @constraint(model, Pdecharge_battery[Tmax] <= stock_battery[Tmax])
    @constraint(model, stock_battery[Tmax] == stock_battery[1])
    @constraint(model, Pdecharge_battery[1] == 0)
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

    @views installed_CCG_H2[week] = round(Int64, sum(value.(CCG_H2_installed)))
    @views installed_TAC_H2[week] = round(Int64, sum(value.(TAC_H2_installed)))

    @views Phy_annual[idx]  .= value.(Phy[1:Nhours_per_week])
    @views Puns_annual[idx] .= value.(Puns[1:Nhours_per_week])
    @views Pexc_annual[idx] .= value.(Pexc[1:Nhours_per_week])
    
    @views load_annual[idx] .= value.(load[1:Nhours_per_week])
    
    # --- Mise à jour des capacités pour la semaine suivante (évolution du parc) ---
    solar_capacities[week:week+1]    .= round(Int, value(CapaSolar))
    onshore_capacities[week:week+1]  .= round(Int, value(CapaOnshore))
    offshore_capacities[week:week+1] .= round(Int, value(CapaOffshore))
    battery_capacities[week:week+1]  .= round(Int, value(CapaBattery))

    
    # --- Stock initial pour la semaine suivante ---
    global stock_battery_initial = value(stock_battery[Tmax])
    global stock_STEP_initial    = value(stock_STEP[Tmax])

    last_CCG_H2_running = value.(CCG_H2_running[Tmax, :])
    last_TAC_H2_running = value.(TAC_H2_running[Tmax, :])
    
    global CCG_H2_running_initial = round.(Int, last_CCG_H2_running)
    global TAC_H2_running_initial = round.(Int, last_TAC_H2_running)

    global CCG_H2_installed_initial = round.(Int, value.(CCG_H2_installed))
    global TAC_H2_installed_initial = round.(Int, value.(TAC_H2_installed))

    t_end = now()
    println("Fin de l'EOD semaine $week, durée : $(t_end)")
    
end

# --- Export CSV annuel (comme dans ton code) ---
open("results/annual/results.csv", "w") do f
    write(f, "t;Solar;Onshore;Offshore;Battery_stock;Battery_charge;Battery_discharge;STEP_stock;STEP_charge;STEP_discharge;H2_CCG;H2_TAC;Hydro;Load;Defailance;Exces\n")
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
                round(load_annual[t], digits=2), ";",
                round(Puns_annual[t], digits=2), ";",
                round(Pexc_annual[t], digits=2),
                "\n"
            )
        )
    end
end

using JSON3

# Construction du dictionnaire annuel
parc = Dict(
    "capacites_MW" => Dict(
        "onshore" => round(value(onshore_capacities[LAST_WEEK]), digits=2),
        "offshore" => round(value(offshore_capacities[LAST_WEEK]), digits=2),
        "solar" => round(value(solar_capacities[LAST_WEEK]), digits=2),
        "battery" => round(value(battery_capacities[LAST_WEEK]), digits=2)
    ),

"H2" => Dict(
    "CCG" => Dict(
        "nombre_installees" => sum(CCG_H2_installed_initial), 
        "production_totale_MWh" => round(sum(PH2_CCG_annual), digits=2)
    ),

    "TAC" => Dict(
        "nombre_installees" => sum(TAC_H2_installed_initial),
        "production_totale_MWh" => round(sum(PH2_TAC_annual), digits=2)
    )
),

    "energie_totale_MWh" => Dict(
        "defaillance" => round(sum(value.(Puns_annual)), digits=2),
        "exces" => round(sum(value.(Pexc_annual)), digits=2),
        "production_H2_totale" => round(sum(value.(PH2_CCG_annual)) + sum(value.(PH2_TAC_annual)), digits=2),
        "production_hydro" => round(sum(value.(Phy_annual)), digits=2)
    )
)


# Écriture dans fichier JSON
open("results/annual/parc_annuel.json", "w") do f
    JSON3.write(f, parc; indent=4)
end

println("Fichier parc_annuel.json généré avec succès ✅")

evolution_parc = Dict(
    "semaine" => 1:LAST_WEEK,
    "capacites_MW" => Dict(
        "onshore" => onshore_capacities[1:LAST_WEEK],
        "offshore" => offshore_capacities[1:LAST_WEEK],
        "solar" => solar_capacities[1:LAST_WEEK],
        "battery" => battery_capacities[1:LAST_WEEK]
    ),
    "H2" => Dict(
        "CCG" => Dict(
            "nombre_installees" => installed_CCG_H2[1:LAST_WEEK],
            "production_totale_MWh" => [round(sum(value.(PH2_CCG_annual)), digits=2) for week in 1:LAST_WEEK]
        ),

        "TAC" => Dict(
            "nombre_installees" => installed_TAC_H2[1:LAST_WEEK],
            "production_totale_MWh" => [round(sum(value.(PH2_TAC_annual)), digits=2) for week in 1:LAST_WEEK]
        )
    )
)

# Écriture dans fichier JSON
open("results/annual/evolution_parc.json", "w") do f
    JSON3.write(f, evolution_parc; indent=4)
end