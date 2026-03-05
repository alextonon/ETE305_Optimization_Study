include("utils/extraction_donnees_excel.jl")
using JuMP
using Dates
using Random

# -------- Configuration ---------

FIRST_WEEK = 40 # Semaine début de la simulation (1 à 52)

H2_ANNUAL_STOCK = false
H2_NO_LIMIT = false # false : ne pas cumuler au stockage annuel...
GISEMENTS = false
HYDRO_STOCK_REMAINING = false 

solver_list = ["HiGHS", "Gurobi", "SCIP", "CBC"]  # Gurobi nécessite une license
solver = solver_list[3] # Choix du solveur
TIMING_COMPUTATION = true

# -------- Extraction des hypothèses du problèmes --------
data_file = "data/Donnees_etude_de_cas_ETE305.xlsx"
base_de_resultats = "results/base_de_données_résultats.csv"

config = extraire_donnees_config(data_file)

# Data for h2 clusters
capex_CCG_H2 = config["H2"]["CCG"]["capex"] #€/MW
opex_CCG_H2 = config["H2"]["CCG"]["opex"] #€/MW/year
PU_cost_h2_CCG = config["H2"]["CCG"]["PU_cost"] #€/MWh basé sur le tarif de prod de la centrale CCG gaz A MODIFIER
NH2_CCG_max = config["H2"]["CCG"]["gisement"]
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

# Gestion H2
RendementElectrolyse = config["rendements"]["electrolyse"] # Rendement de l'électrolyse
RendementCombustion = config["rendements"]["combustion"] # Rendement de la combustion de l'électrolyse

capex_electrolyzer = config["electrolyzer"]["capex"] #€/MW
opex_electrolyzer = config["electrolyzer"]["opex"] #€/MW/year
d_electrolyzer = config["electrolyzer"]["duree_vie"] # years

# Renewables
capex_onshore = config["enr"]["onshore"]["capex"] #€/MW
capex_offshore = config["enr"]["offshore_pose"]["capex"] #€/MW
capex_solar = config["enr"]["pv_pose"]["capex"] #€/MW

opex_onshore = config["enr"]["onshore"]["opex"] #€/MW
opex_offshore = config["enr"]["offshore_pose"]["opex"] #€/MW
opex_solar = config["enr"]["pv_pose"]["opex"] #€/MW

CapaSolar_max = config["enr"]["gisement"]["Solar"] #MW
CapaOnshore_max = config["enr"]["gisement"]["Onshore"] #MW
CapaOffshore_max = config["enr"]["gisement"]["Offshore"] #MW

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
CapaBattery_max = config["battery"]["gisement"] #MW

CapaElectrolyzer_max = 100000 # à Implémenter....

if GISEMENTS == false
    CapaSolar_max = CapaSolar_max + 25000
    CapaOnshore_max = CapaOnshore_max + 25000
    CapaBattery_max = CapaBattery_max + 25000
    CapaOffshore_max = CapaOffshore_max + 25000
    NH2_CCG_max = NH2_CCG_max + 5
    NH2_TAC_max = NH2_TAC_max + 5
end

# Defailance
cuns = config["defaillance"]["cost_unsupplied"] #cost of unsupplied energy €/MWh
cexc = config["defaillance"]["cost_excess"] #cost of in excess energy €/MWh

# Initial capacities 
CapaSolar_init = config["capacites_init"]["Solar"] #MW
CapaOffshore_init = config["capacites_init"]["Offshore"] #MW
CapaOnshore_init = config["capacites_init"]["Onshore"] #MW
CapaBattery_init = 0 #MW

# ----------------- Paramètre optim -----------------
# Monitoring du temps
if TIMING_COMPUTATION
    timing_file = joinpath("results/", "timing_log.csv")
    open(timing_file, "w") do f
        write(f, "iteration;semaine_calendaire;temps_secondes;status\n")
    end
end

if solver == "Gurobi"
    println("Attention : tu as choisi Gurobi comme solveur, assure-toi que la licence est bien configurée sur ta machine !")
    using Gurobi
elseif solver == "HiGHS"
    using HiGHS
elseif solver == "SCIP"
    using SCIP
elseif solver == "CBC"
    using Cbc
else
    error("Solveur inconnu : $solver. Choisis entre 'Gurobi', 'HiGHS', 'SCIP'.")
end

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
electrolyzer_capacities = fill(0, Nweeks+1)
installed_CCG_H2 = fill(0, Nweeks+1)
installed_TAC_H2 = fill(0, Nweeks+1)

# Stock Hydro
hydro_utilization_rate = zeros(Nweeks+1)
hydro_utilization_annual = zeros(Nhours)

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
PH2_CCG_annual        = zeros(Nhours, NH2_CCG_max)
PH2_TAC_annual        = zeros(Nhours, NH2_TAC_max)

stock_H2_annual = zeros(Nhours)
electrolyzer_annual = zeros(Nhours)

# Hydro / defaillance / excès / hydro et thermique fatal
Phy_annual  = zeros(Nhours)
Puns_annual = zeros(Nhours)
Pexc_annual = zeros(Nhours)
Pres_annual = zeros(Nhours)

# Conditions initiales pour la première semaine
global stock_battery_initial = 0
global stock_STEP_initial    = 0
global stock_hydro_lac_initial = 0.0
global stock_H2_initial = 1e7

global CCG_H2_running_initial = zeros(Int, NH2_CCG_max)
global TAC_H2_running_initial = zeros(Int, NH2_TAC_max)

global CCG_H2_installed_initial = zeros(Int, NH2_CCG_max)
global TAC_H2_installed_initial = zeros(Int, NH2_TAC_max)

LAST_WEEK = FIRST_WEEK + Nweeks - 1

for (i, w) in enumerate(FIRST_WEEK:LAST_WEEK)
    t_start_proc = time() # Début du chrono
    week = i # Simulation week
    current_week = (w - 1) % 52 + 1 # Annual week
    
    t_start = now()
    println("Itération $i/$Nweeks | Semaine calendaire $current_week | Début : $t_start")

    time_series = extraire_donnees_semaine(data_file, current_week, AnneauGarde=10)

    Tmax = time_series["Tmax"]

    load = time_series["load"]
    offshore_load_factor = time_series["offshore_load_factor"]
    onshore_load_factor = time_series["onshore_load_factor"]
    solar_load_factor = time_series["solar_load_factor"]
    hydro_fatal = time_series["hydro_fatal"]
    e_hy_lacs = time_series["Usable_per_week_hydro_lacs"] + stock_hydro_lac_initial
    thermique_fatal = time_series["thermique_fatal"]

    Pres = hydro_fatal + thermique_fatal

    ########## Defining model ##########

    if solver == "Gurobi"
        model = Model(Gurobi.Optimizer)
        set_optimizer_attribute(model, "MIPGap", 0.01)       # S'arrête à 1% de l'optimum 
        set_optimizer_attribute(model, "OutputFlag", 1)      # Outputs
        set_optimizer_attribute(model, "Threads", 0)         # maximise les threads CPU utilisés
    elseif solver == "HiGHS"
        model = Model(HiGHS.Optimizer)
        set_optimizer_attribute(model, "mip_rel_gap", 0.01)
        set_optimizer_attribute(model, "parallel", "on")
        set_optimizer_attribute(model, "threads", 0)
    elseif solver == "SCIP"
        model = Model(SCIP.Optimizer)
        set_optimizer_attribute(model, "limits/gap", 0.01)       # Gap de 1%
        set_optimizer_attribute(model, "display/verblevel", 1)   # Niveau de log (0 à 5)
        set_optimizer_attribute(model, "parallel/mode", 1)       # Active le mode parallèle
        set_optimizer_attribute(model, "lp/threads", 0)    # Utilise tous les threads disponibles
    elseif solver == "CBC"
        model = Model(Cbc.Optimizer)
        set_optimizer_attribute(model, "ratio", 0.01)       # Gap de 1%
        set_optimizer_attribute(model, "logLevel", 1)   # Niveau de log (0 à 5)
        set_optimizer_attribute(model, "threads", 0)    # Utilise tous les threads disponibles
    end

    ########## Defining variables ##########
    #energie renouvelables
    @variable(model, onshore_capacities[week] <= CapaOnshore <= CapaOnshore_max)
    @variable(model, offshore_capacities[week] <= CapaOffshore <= CapaOffshore_max)
    @variable(model, solar_capacities[week] <= CapaSolar <= CapaSolar_max)

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

    @variable(model, Pcharge_electrolyzer[1:Tmax] >= 0)
    @variable(model, electrolyzer_capacities[week] <= CapaElectrolyzer<= CapaElectrolyzer_max)

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
    @variable(model, battery_capacities[week] <= CapaBattery <= CapaBattery_max)
    @variable(model, Pcharge_battery[1:Tmax] >= 0)
    @variable(model, Pdecharge_battery[1:Tmax] >= 0)
    @variable(model, stock_battery[1:Tmax] >= 0)
    @variable(model, charging_battery[1:Tmax], Bin) # 1 si la batterie charge à t, 0 sinon (pour éviter de charger et décharger en même temps)

    # Stockage et volume H2
    if H2_ANNUAL_STOCK
        @variable(model, stock_H2[1:Tmax] >= 0)
        @constraint(model, [t in 1:Tmax], Pcharge_electrolyzer[t] <= CapaElectrolyzer)

        @constraint(model, stock_H2[1] == stock_H2_initial + (Pcharge_electrolyzer[1]*RendementElectrolyse) - (sum(PH2_CCG[1,g] for g in 1:NH2_CCG_max) + sum(PH2_TAC[1,g] for g in 1:NH2_TAC_max))/RendementCombustion)
        @constraint(model, stock_H2_balance[t in 2:Tmax], 
            stock_H2[t] == stock_H2[t-1] 
            + (Pcharge_electrolyzer[t] * RendementElectrolyse)      # Ce qu'on transforme en H2
            - (sum(PH2_CCG[t,g] for g in 1:NH2_CCG_max) + sum(PH2_TAC[t,g] for g in 1:NH2_TAC_max)) / RendementCombustion           # Ce qu'on puise pour faire de l'élec
        )

    elseif H2_NO_LIMIT == false
        @constraint(model, sum(PH2_CCG[t,g] for t in 1:Tmax, g in 1:NH2_CCG_max) + sum(PH2_TAC[t,g] for t in 1:Tmax, g in 1:NH2_TAC_max) <= sum(Pexc[t] for t in 1:Tmax)*RendementCombustion*RendementElectrolyse)
        @constraint(model, Pcharge_electrolyzer == 0)
    else
        @constraint(model, Pcharge_electrolyzer == 0) # On s'assure que l'électrolyse ne peut pas permettre d'éviter Pexc
    end


    # set_start_value.(CapaOnshore, onshore_capacities[week])
    # set_start_value.(CapaOffshore, offshore_capacities[week])
    # set_start_value.(CapaSolar, solar_capacities[week])
    # set_start_value.(CapaBattery, battery_capacities[week])
    # set_start_value.(CapaElectrolyzer, electrolyzer_capacities[week])

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

    for g in 1:NH2_CCG_max
        @constraint(model, CCG_H2_running[1,g] - CCG_H2_running_initial[g] == CCG_H2_start[1,g] - CCG_H2_stop[1,g])
    end

    ####### Defining objective function #######
    @objective(model, Min, 
                        (CapaSolar*(capex_solar + opex_solar) + CapaOffshore*(capex_offshore + opex_offshore) + CapaOnshore*(capex_onshore + opex_onshore) 
                        + CapaBattery*(capex_2h_battery + opex_2h_battery) 
                        + sum(CCG_H2_installed[g]*Pmax_CCG_h2 for g in 1:NH2_CCG_max)*(capex_CCG_H2 + opex_CCG_H2) + sum(PH2_CCG[t,g] for t in 1:Tmax, g in 1:NH2_CCG_max)*PU_cost_h2_CCG
                        + sum(TAC_H2_installed[g]*Pmax_TAC_h2 for g in 1:NH2_TAC_max)*(capex_TAC_H2 + opex_TAC_H2) + sum(PH2_TAC[t,g] for t in 1:Tmax, g in 1:NH2_TAC_max)*PU_cost_h2_TAC
                        + CapaElectrolyzer*(capex_electrolyzer + opex_electrolyzer)
                        + sum(Puns[t] for t in 1:Tmax)*cuns + sum(Pexc[t] for t in 1:Tmax)*cexc)/1e4
    )

    ######## Defining constraints ########
    # Initial conditions
    @constraint(model, stock_battery[1] == stock_battery_initial)
    @constraint(model, stock_STEP[1] == stock_STEP_initial)

    #balance constraint
    @constraint(model, balance[t in 1:Tmax], sum(PH2_CCG[t,g] for g in 1:NH2_CCG_max) + sum(PH2_TAC[t,g] for g in 1:NH2_TAC_max) + Phy[t] + Pres[t] + CapaSolar * solar_load_factor[t] + CapaOffshore * offshore_load_factor[t] + CapaOnshore * onshore_load_factor[t] - Pcharge_STEP[t] + Pdecharge_STEP[t] - Pcharge_battery[t] + Pdecharge_battery[t] + Puns[t] - load[t] - Pcharge_electrolyzer[t] - Pexc[t]  == 0)

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

    # hydro unit constraints
    @constraint(model, [t=1:Tmax], Phy[t] >= Pmin_hy_lacs)
    @constraint(model, [t=1:Tmax], Phy[t] <= Pmax_hy_lacs)
    # hydro stock constraint
    @constraint(model, sum(Phy[t] for t in 1:Tmax) <= e_hy_lacs)

    # weekly STEP
    @constraint(model, [t in 1:Tmax], Pcharge_STEP[t] <= Pmax_STEP)
    @constraint(model, [t in 1:Tmax], Pdecharge_STEP[t] <= Pmax_STEP)
    @constraint(model, Pdecharge_STEP[Tmax] <= stock_STEP[Tmax])
    #@constraint(model, stock_STEP[Tmax] == stock_STEP[1])
    #@constraint(model, Pdecharge_STEP[1] == 0)
    @constraint(model, [t in 1:Tmax-1], stock_STEP[t+1]-stock_STEP[t]- rSTEP*Pcharge_STEP[t]+Pdecharge_STEP[t]== 0)
    @constraint(model, [t in 1:Tmax], stock_STEP[t] <= 24*7*Pmax_STEP)

    # #battery
    @constraint(model, [t in 1:Tmax], Pcharge_battery[t] <= CapaBattery)
    @constraint(model, [t in 1:Tmax], Pdecharge_battery[t] <= CapaBattery)
    @constraint(model, Pdecharge_battery[Tmax] <= stock_battery[Tmax])
    #@constraint(model, stock_battery[Tmax] == stock_battery[1])
    #@constraint(model, Pdecharge_battery[1] == 0)
    @constraint(model, [t in 1:Tmax-1], stock_battery[t+1]-stock_battery[t]- rbattery*Pcharge_battery[t]+1/rbattery*Pdecharge_battery[t]== 0)
    @constraint(model, [t in 1:Tmax], stock_battery[t] <= d_battery*CapaBattery)

    # @constraint(model, [t in 1:Tmax], Pcharge_battery[t] <= CapaBattery_max * charging_battery[t])
    # @constraint(model, [t in 1:Tmax], Pdecharge_battery[t] <= CapaBattery_max * (1 - charging_battery[t]))
    
    optimize!(model)

    # Calcul du temps
    t_end_proc = time() # Fin du chrono
    duree = t_end_proc - t_start_proc
    statut = termination_status(model)

    if TIMING_COMPUTATION 
        open(timing_file, "a") do f
            write(f, "$i;$current_week;$duree;$statut\n")
        end
    end

    println("Itération $i terminée en $(round(duree, digits=2))s | Statut : $statut")

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

    @views installed_CCG_H2[week] = round(Int64, sum(value.(CCG_H2_installed)))
    @views installed_TAC_H2[week] = round(Int64, sum(value.(TAC_H2_installed)))

    @views Phy_annual[idx]  .= value.(Phy[1:Nhours_per_week])
    @views Puns_annual[idx] .= value.(Puns[1:Nhours_per_week])
    @views Pexc_annual[idx] .= value.(Pexc[1:Nhours_per_week])
    @views Pres_annual[idx] .= Pres[1:Nhours_per_week]

    @views electrolyzer_annual[idx] .= value.(Pcharge_electrolyzer[1:Nhours_per_week])
    
    @views load_annual[idx] .= value.(load[1:Nhours_per_week])

    if H2_ANNUAL_STOCK
        @views stock_H2_annual[idx] .= value.(stock_H2[1:Nhours_per_week])
    end
    if HYDRO_STOCK_REMAINING
        production_hydro_week = sum(value.(Phy))
        remaining_hydro = e_hy_lacs - production_hydro_week
        hydro_utilization_rate[week] = production_hydro_week / e_hy_lacs
        @views hydro_utilization_annual[idx] .= hydro_utilization_rate[week]

        global stock_hydro_lac_initial = remaining_hydro
    end
    
    # --- Mise à jour des capacités pour la semaine suivante (évolution du parc) ---
    solar_capacities[week:week+1]    .= round(Int, value(CapaSolar))
    onshore_capacities[week:week+1]  .= round(Int, value(CapaOnshore))
    offshore_capacities[week:week+1] .= round(Int, value(CapaOffshore))
    battery_capacities[week:week+1]  .= round(Int, value(CapaBattery))
    electrolyzer_capacities[week:week+1] .= round(Int, value(CapaElectrolyzer))
    
    # --- Stock initial pour la semaine suivante ---
    global stock_battery_initial = value(stock_battery[Tmax])
    global stock_STEP_initial    = value(stock_STEP[Tmax])
    if H2_ANNUAL_STOCK
        global stock_H2_initial = value(stock_H2[Tmax])
    end

    last_CCG_H2_running = value.(CCG_H2_running[Tmax, :])
    last_TAC_H2_running = value.(TAC_H2_running[Tmax, :])
    
    global CCG_H2_running_initial = round.(Int, last_CCG_H2_running)
    global TAC_H2_running_initial = round.(Int, last_TAC_H2_running)

    global CCG_H2_installed_initial = round.(Int, value.(CCG_H2_installed))
    global TAC_H2_installed_initial = round.(Int, value.(TAC_H2_installed))

    t_end = now()
    println("Fin de l'EOD semaine $week, durée : $(t_end)")
    
end

open(base_de_resultats, "r") do f
    lignes = readlines(f)
    lignes_non_vides = filter(!isempty, lignes)
    
    if isempty(lignes_non_vides)
        global id = 0
    else
        derniere_ligne = lignes_non_vides[end]
        valeurs = split(derniere_ligne, ";")
        
        id_temp = tryparse(Int, valeurs[1])
        global id = isnothing(id_temp) ? 0 : id_temp
    end
end

global id_hex
id_hex = randstring(['0':'9'; 'A':'F'], 8)

FIRST_WEEK_PARC_FIXE = -1

open(base_de_resultats, "a") do f
    write(f,
        string(
            id_hex, ";",
            H2_ANNUAL_STOCK, ";",
            H2_NO_LIMIT, ";",
            GISEMENTS, ";",
            HYDRO_STOCK_REMAINING, ";",
            FIRST_WEEK, ";",
            FIRST_WEEK_PARC_FIXE, ";",
            "\n"
        )
    )
end

result_file_path = "results/$(id_hex)/results.csv"
parc_file_path = "results/$(id_hex)/parc_annuel.json"
evolution_parc_file_path = "results/$(id_hex)/evolution_parc.json"

dir_path = "results/$(id_hex)"
mkpath(dir_path)

result_file_path = joinpath(dir_path, "results.csv")
parc_file_path = joinpath(dir_path, "parc_annuel.json")
evolution_parc_file_path = joinpath(dir_path, "evolution_parc.json")

# --- Export CSV annuel (comme dans ton code) ---
open(result_file_path, "w") do f
    write(f, "t;Solar;Onshore;Offshore;Battery_stock;Battery_charge;Battery_discharge;STEP_stock;STEP_charge;STEP_discharge;H2_CCG;H2_TAC;H2_Stock;Hydro;Hydro_lac_utilization_rate;hy_th_fatal;Load;Defailance;Exces;Electrolyse\n")
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
                round(stock_H2_annual[t], digits=2), ";",
                round(Phy_annual[t], digits=2), ";",
                round(hydro_utilization_annual[t], digits=2), ";",
                round(Pres_annual[t], digits=2), ";",
                round(load_annual[t], digits=2), ";",
                round(Puns_annual[t], digits=2), ";",
                round(Pexc_annual[t], digits=2), ";",
                round(electrolyzer_annual[t], digits=2),
                "\n"
            )
        )
    end
end

println("Fichier results_$(id_hex).csv généré avec succès ✅")
using JSON3

# Construction du dictionnaire annuel
parc = Dict(
    "capacites_MW" => Dict(
        "onshore" => round(value(onshore_capacities[Nweeks]), digits=2),
        "offshore" => round(value(offshore_capacities[Nweeks]), digits=2),
        "solar" => round(value(solar_capacities[Nweeks]), digits=2),
        "battery" => round(value(battery_capacities[Nweeks]), digits=2),
        "electrolyzer" => round(value(electrolyzer_capacities[Nweeks]), digits=2)
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
        "production_hydro" => round(sum(value.(Phy_annual)), digits=2),
        "p_charge_electrolyser" => round(sum(value.(electrolyzer_annual)), digits=2)
    )
)


# Écriture dans fichier JSON
open(parc_file_path, "w") do f
    JSON3.write(f, parc; indent=4)
end

println("Fichier parc_annuel_$(id_hex).json généré avec succès ✅")

evolution_parc = Dict(
    "semaine" => 1:Nweeks,
    "capacites_MW" => Dict(
        "onshore" => onshore_capacities[1:Nweeks],
        "offshore" => offshore_capacities[1:Nweeks],
        "solar" => solar_capacities[1:Nweeks],
        "battery" => battery_capacities[1:Nweeks],
        "electrolyzer" => electrolyzer_capacities[1:Nweeks]
        ),
    "H2" => Dict(
        "CCG" => Dict(
            "nombre_installees" => installed_CCG_H2[1:Nweeks],
            "production_totale_MWh" => [round(sum(value.(PH2_CCG_annual)), digits=2) for week in 1:Nweeks]
        ),

        "TAC" => Dict(
            "nombre_installees" => installed_TAC_H2[1:Nweeks],
            "production_totale_MWh" => [round(sum(value.(PH2_TAC_annual)), digits=2) for week in 1:Nweeks]
        )
    )
)

# Écriture dans fichier JSON
open(evolution_parc_file_path, "w") do f
    JSON3.write(f, evolution_parc; indent=4)
end

println("Fichier evolution_parc_$(id_hex).json généré avec succès ✅")