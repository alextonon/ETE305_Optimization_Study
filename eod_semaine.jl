#packages
using JuMP
#use the solver you want
using HiGHS
#package to read excel files
using XLSX

Tweek = 168 #optimization for 1 week (7*24=168 hours)
Tguard = 24

Tmax = Tweek + Tguard

#data for load and fatal generation
data_file = "data/Donnees_etude_de_cas_ETE305.xlsx"
#data for load and fatal generation
load = XLSX.readdata(data_file, "Résumé", "H2:H193")
offshore_load_factor = XLSX.readdata(data_file, "Résumé", "J2:J193")  
onshore_load_factor = XLSX.readdata(data_file, "Résumé", "I2:I193")
solar_load_factor = XLSX.readdata(data_file, "Résumé", "K2:K193")
hydro_fatal = XLSX.readdata(data_file, "Résumé", "L2:L193")
Pres = hydro_fatal 
thermique_fatal = XLSX.readdata(data_file, "Résumé", "M2:M193")

# data per week for hydro use (reference)
Usable_per_week_hydro_lacs= XLSX.readdata(data_file, "Résumé", "D2") #quantité d'hydro disponible à la semaine (en MWh) pour les lacs
Usable_per_week_hydro_STEP= XLSX.readdata(data_file, "Résumé", "D4") #quantité d'hydro disponible à la semaine (en MWh) pour les STEP

#initial capacities 
CapaSolar_init = XLSX.readdata(data_file, "Parc électrique", "C24") #MW
CapaOffshore_init = XLSX.readdata(data_file, "Parc électrique", "C23") #MW
CapaOnshore_init = XLSX.readdata(data_file, "Parc électrique", "C22") #MW

#max gistances
CapaSolar_max = XLSX.readdata(data_file, "Gisements", "B3") #MW
CapaOnshore_max = XLSX.readdata(data_file, "Gisements", "B4") #MW
CapaOffshore_max = XLSX.readdata(data_file, "Gisements", "B5") #MW

#### Loading of CAPEX/OPEX data

types_centrales=["onshore", "offshore_pose", "offshore_flot", "pv_pose", "pv_gd_toit", "pv_pet_toit", "CCG_H2", "TAC_H2", "electrolyseur", "batterie"]
CAPEX=XLSX.readdata(data_file, "Investissements", "B2:B11") # CAPEX des différentes technologies en €/kW
OPEX=XLSX.readdata(data_file, "Investissements", "C2:C11") # OPEX des différentes technologies en €/kW/an
Duree_vie=XLSX.readdata(data_file, "Investissements", "D2:D11") # Durée de vie des différentes technologies en années

#data for h2 clusters
#CCG H2
NH2_CCG_max = XLSX.readdata(data_file, "Gisements", "B12") 
NH2_CCG_max = Int(NH2_CCG_max)
capex_CCG_H2 = CAPEX[7]*1000 #€/MW
opex_CCG_H2 = OPEX[7]*1000 #€/MW/year
PU_cost_h2_CCG = XLSX.readdata(data_file, "H2", "B5") #€/MWh basé sur le tarif de prod de la centrale CCG gaz A MODIFIER
Pmin_CCG_h2 = XLSX.readdata(data_file, "Parc électrique", "F9")*ones(NH2_CCG_max) #MW idem
Pmax_CCG_h2 = XLSX.readdata(data_file, "Parc électrique", "E9")*ones(NH2_CCG_max) #MW idem
dmin_CCG_h2 = XLSX.readdata(data_file, "Parc électrique", "G9")*ones(Int, NH2_CCG_max) #hours idem

#TAC H2
NH2_TAC_max = XLSX.readdata(data_file, "Gisements", "B13") 
NH2_TAC_max = Int(NH2_TAC_max)
capex_TAC_H2 = CAPEX[8]*1000 #€/MW
opex_TAC_H2 = OPEX[8]*1000 #€/MW/year
PU_cost_h2_TAC = XLSX.readdata(data_file, "H2", "B5") #€/MWh basé sur le tarif de prod de la centrale TAC gaz A MODIFIER
Pmin_TAC_h2 = XLSX.readdata(data_file, "Parc électrique", "F10")*ones(NH2_TAC_max) #MW idem
Pmax_TAC_h2 = XLSX.readdata(data_file, "Parc électrique", "E10")*ones(NH2_TAC_max) #MW idem
dmin_TAC_h2 = XLSX.readdata(data_file, "Parc électrique", "G10")*ones(Int, NH2_TAC_max) #hours idem


#data for hydro reservoir "lacs"
Nhy = 1 #number of hydro generation units
Pmin_hy_lacs = zeros(Nhy)
Pmax_hy_lacs = XLSX.readdata(data_file, "Parc électrique", "C20") *ones(Nhy) #MW
e_hy_lacs = Usable_per_week_hydro_lacs*ones(Nhy) #MWh


# variable costs
cuns = XLSX.readdata(data_file, "Defaillance", "B2") * 1000 #cost of unsupplied energy €/MWh
cexc = XLSX.readdata(data_file, "Defaillance", "B3") #cost of in excess energy €/MWh

 #investment costs 
capex_onshore = CAPEX[1]*1000 #€/MW
capex_offshore = CAPEX[2]*1000 #€/MW
capex_solar = CAPEX[4]*1000 #€/MW
capex_2h_battery = CAPEX[10]*1000 #€/MW

opex_onshore = OPEX[1]*1000 #€/MW
opex_offshore = OPEX[2]*1000 #€/MW
opex_solar = OPEX[4]*1000 #€/MW
opex_2h_battery = OPEX[10]*1000 #€/MW


#data for STEP/battery
#weekly STEP
Pmax_STEP = XLSX.readdata(data_file, "Parc électrique", "C21") #MW
rSTEP = XLSX.readdata(data_file, "Rendements", "B10") #rendement au pompage 

#battery
rbattery = XLSX.readdata(data_file, "Rendements", "B9") # rendement de la batterie au sockage ou au destockage
d_battery = XLSX.readdata(data_file, "Investissements", "E11") #hours
CapaBattery_init = 0 #MW
CapaBattery_max = XLSX.readdata(data_file, "Gisements", "B8") #MW





#############################
# Verification of loaded data

println("=== Vérification des données chargées ===")

# Charges et facteurs de charge
@show load
@show offshore_load_factor
@show onshore_load_factor
@show solar_load_factor
@show hydro_fatal
@show Pres
@show thermique_fatal

# Hydro utilisable par semaine
@show Usable_per_week_hydro_lacs
@show Usable_per_week_hydro_STEP

# Capacités initiales et max
@show CapaSolar_init
@show CapaOffshore_init
@show CapaOnshore_init
@show CapaSolar_max
@show CapaOnshore_max
@show CapaOffshore_max

# CAPEX, OPEX, durée de vie
@show CAPEX
@show OPEX
@show Duree_vie

# H2 CCG et TAC
@show NH2_CCG_max
@show capex_CCG_H2
@show opex_CCG_H2
@show PU_cost_h2_CCG
@show Pmin_CCG_h2
@show Pmax_CCG_h2
@show dmin_CCG_h2

@show NH2_TAC_max
@show capex_TAC_H2
@show opex_TAC_H2
@show PU_cost_h2_TAC
@show Pmin_TAC_h2
@show Pmax_TAC_h2
@show dmin_TAC_h2

# Hydro
@show Nhy
@show Pmin_hy_lacs
@show Pmax_hy_lacs
@show e_hy_lacs

# Coûts variables
@show cuns
@show cexc

# Coûts d'investissement
@show capex_onshore
@show capex_offshore
@show capex_solar
@show capex_2h_battery

@show opex_onshore
@show opex_offshore
@show opex_solar
@show opex_2h_battery

# STEP et batteries
@show Pmax_STEP
@show rSTEP

@show rbattery
@show d_battery
@show CapaBattery_init
@show CapaBattery_max

println("=== Fin de la vérification ===")




#############################
#create the optimization model
#############################
model = Model(HiGHS.Optimizer)

#############################
#define the variables
#############################
#energie renouvelables

@variable(model, 0 <= CapaOnshore <= CapaOnshore_max, start = CapaOnshore_init)
@variable(model, 0 <= CapaOffshore <= CapaOffshore_max, start = CapaOffshore_init)
@variable(model, 0 <= CapaSolar <= CapaSolar_max, start = CapaSolar_init)

# @variable(model, capacity_offshore_fixed >= 0)
# @variable(model, capacity_offshore_floating >= 0)
# @variable(model, capacity_solar_ground >= 0)
# @variable(model, capacity_solar_big_roof >= 0)
# @variable(model, capacity_solar_small_roof >= 0)

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
@variable(model, Phy[1:Tmax,1:Nhy] >= 0)
#unsupplied energy variables
@variable(model, Puns[1:Tmax] >= 0)
#in excess energy variables
@variable(model, Pexc[1:Tmax] >= 0)
#weekly STEP variables
@variable(model, Pcharge_STEP[1:Tmax] >= 0)
@variable(model, Pdecharge_STEP[1:Tmax] >= 0)
@variable(model, stock_STEP[1:Tmax] >= 0)
# #battery variables
@variable(model, 0 <= CapaBattery <= CapaBattery_max, start = CapaBattery_init)
@variable(model, Pcharge_battery[1:Tmax] >= 0)
@variable(model, Pdecharge_battery[1:Tmax] >= 0)
@variable(model, stock_battery[1:Tmax] >= 0)
# #############################
#define the objective function
#############################
@objective(model, Min, 
                        CapaSolar*(capex_solar + opex_solar) + CapaOffshore*(capex_offshore + opex_offshore) + CapaOnshore*(capex_onshore + opex_onshore) 
                        + CapaBattery*(capex_2h_battery + opex_2h_battery) 
                        + sum(CCG_H2_installed[g]*Pmax_CCG_h2[g] for g in 1:NH2_CCG_max)*(capex_CCG_H2 + opex_CCG_H2) + sum(PH2_CCG[t,g] for t in 1:Tmax, g in 1:NH2_CCG_max)*PU_cost_h2_CCG
                        + sum(TAC_H2_installed[g]*Pmax_TAC_h2[g] for g in 1:NH2_TAC_max)*(capex_TAC_H2 + opex_TAC_H2) + sum(PH2_TAC[t,g] for t in 1:Tmax, g in 1:NH2_TAC_max)*PU_cost_h2_TAC
                        + sum(Puns[t] for t in 1:Tmax)*cuns + sum(Pexc[t] for t in 1:Tmax)*cexc
)

#cout de l'ecretement et prod H2 gratuit
# @objective(model, Min, 
#                         CapaSolar*(capex_solar + opex_solar) + CapaOffshore*(capex_offshore + opex_offshore) + CapaOnshore*(capex_onshore + opex_onshore) 
#                         + CapaBattery*(capex_2h_battery + opex_2h_battery) 
#                         + sum(CCG_H2_installed[g]*Pmax_CCG_h2[g] for g in 1:NH2_CCG_max)*(capex_CCG_H2 + opex_CCG_H2) + sum(PH2_CCG[t,g] for t in 1:Tmax, g in 1:NH2_CCG_max)*0
#                         + sum(TAC_H2_installed[g]*Pmax_TAC_h2[g] for g in 1:NH2_TAC_max)*(capex_TAC_H2 + opex_TAC_H2) + sum(PH2_TAC[t,g] for t in 1:Tmax, g in 1:NH2_TAC_max)*0
#                         + sum(Puns[t] for t in 1:Tmax)*cuns + sum(Pexc[t] for t in 1:Tmax)*PU_cost_h2_CCG
# )


#############################
#define the constraints
#############################
#balance constraint
@constraint(model, balance[t in 1:Tmax], sum(PH2_CCG[t,g] for g in 1:NH2_CCG_max) + sum(PH2_TAC[t,g] for g in 1:NH2_TAC_max) + sum(Phy[t,h] for h in 1:Nhy) + CapaSolar * solar_load_factor[t] + CapaOffshore * offshore_load_factor[t] + CapaOnshore * onshore_load_factor[t] + Puns[t] - load[t] - Pexc[t] - Pcharge_STEP[t] + Pdecharge_STEP[t] - Pcharge_battery[t] + Pdecharge_battery[t] == 0)

# H2 Power constraints
@constraint(model, max_CCG_H2[t in 1:Tmax, i in 1:NH2_CCG_max], PH2_CCG[t,i] <= Pmax_CCG_h2[i]*CCG_H2_running[t,i]) #Pmax constraints
@constraint(model, min_CCG_H2[t in 1:Tmax, i in 1:NH2_CCG_max], Pmin_CCG_h2[i]*CCG_H2_running[t,i] <= PH2_CCG[t,i]) #Pmin constraints

@constraint(model, max_TAC_H2[t in 1:Tmax, i in 1:NH2_TAC_max], PH2_TAC[t,i] <= Pmax_TAC_h2[i]*TAC_H2_running[t,i]) #Pmax constraints
@constraint(model, min_TAC_H2[t in 1:Tmax, i in 1:NH2_TAC_max], Pmin_TAC_h2[i]*TAC_H2_running[t,i] <= PH2_TAC[t,i]) #Pmin constraints

# H2 instalation constraints
@constraint(model, [t in 1:Tmax, g in 1:NH2_CCG_max], CCG_H2_running[t,g] <= CCG_H2_installed[g]) # only produce if installed
@constraint(model, [t in 1:Tmax, g in 1:NH2_TAC_max], TAC_H2_running[t,g] <= TAC_H2_installed[g]) # only produce if installed

# H2 duration constraints
for g in 1:NH2_CCG_max
        if (dmin_CCG_h2[g] > 1)
            @constraint(model, [t in 2:Tmax], CCG_H2_running[t,g]-CCG_H2_running[t-1,g]==CCG_H2_start[t,g]-CCG_H2_stop[t,g],  base_name = "stateH2_$g") # detect start and stop
            @constraint(model, [t in 1:Tmax], CCG_H2_start[t,g]+CCG_H2_stop[t,g]<=1,  base_name = "exclusiveH2_$g") # avoid starting and stoping at same step

            # Initial conditions
            @constraint(model, CCG_H2_start[1,g]==0,  base_name = "iniStartH2_$g")
            @constraint(model, CCG_H2_stop[1,g]==0,  base_name = "iniStopH2_$g")
            @constraint(model, [t in 1:dmin_CCG_h2[g]-1], CCG_H2_running[t,g] >= sum(CCG_H2_start[i,g] for i in 1:t), base_name = "dminStartH2_$(g)_init")
            @constraint(model, [t in 1:dmin_CCG_h2[g]-1], CCG_H2_running[t,g] <= 1-sum(CCG_H2_stop[i,g] for i in 1:t), base_name = "dminStopH2_$(g)_init")

            # Minimum up and down time constraints
            @constraint(model, [t in dmin_CCG_h2[g]:Tmax], CCG_H2_running[t,g] >= sum(CCG_H2_start[i,g] for i in (t-dmin_CCG_h2[g]+1):t),  base_name = "dminStartH2_$g")
            @constraint(model, [t in dmin_CCG_h2[g]:Tmax], CCG_H2_running[t,g] <= 1 - sum(CCG_H2_stop[i,g] for i in (t-dmin_CCG_h2[g]+1):t),  base_name = "dminStopH2_$g")

        end
end

for g in 1:NH2_TAC_max
        if (dmin_TAC_h2[g] > 1)
            @constraint(model, [t in 2:Tmax], TAC_H2_running[t,g]-TAC_H2_running[t-1,g]==TAC_H2_start[t,g]-TAC_H2_stop[t,g],  base_name = "stateH2_TAC_$g") # detect start and stop
            @constraint(model, [t in 1:Tmax], TAC_H2_start[t,g]+TAC_H2_stop[t,g]<=1,  base_name = "exclusiveH2_TAC_$g") # avoid starting and stoping at same step

            # Initial conditions
            @constraint(model, TAC_H2_start[1,g]==0,  base_name = "iniStartH2_TAC_$g")
            @constraint(model, TAC_H2_stop[1,g]==0,  base_name = "iniStopH2_TAC_$g")
            @constraint(model, [t in 1:dmin_TAC_h2[g]-1], TAC_H2_running[t,g] >= sum(TAC_H2_start[i,g] for i in 1:t), base_name = "dminStartH2_TAC_$(g)_init")
            @constraint(model, [t in 1:dmin_TAC_h2[g]-1], TAC_H2_running[t,g] <= 1-sum(TAC_H2_stop[i,g] for i in 1:t), base_name = "dminStopH2_TAC_$(g)_init")

            # Minimum up and down time constraints
            @constraint(model, [t in dmin_TAC_h2[g]:Tmax], TAC_H2_running[t,g] >= sum(TAC_H2_start[i,g] for i in (t-dmin_TAC_h2[g]+1):t),  base_name = "dminStartH2_TAC_$g")
            @constraint(model, [t in dmin_TAC_h2[g]:Tmax], TAC_H2_running[t,g] <= 1 - sum(TAC_H2_stop[i,g] for i in (t-dmin_TAC_h2[g]+1):t),  base_name = "dminStopH2_TAC_$g")
        end
    
end

# H2 volume constraints
RendementElectrolyse = 0.7 # Rendement de l'électrolyse
RendementCombustion = 0.5 # Rendement de la combustion de l'hydrogène
@constraint(model, sum(PH2_CCG[t,g] for t in 1:Tmax, g in 1:NH2_CCG_max) + sum(PH2_TAC[t,g] for t in 1:Tmax, g in 1:NH2_TAC_max) <= sum(Pexc[t] for t in 1:Tmax)*RendementCombustion*RendementElectrolyse)

# hydro unit constraints
@constraint(model, [t in 1:Tmax, h in 1:Nhy], Pmin_hy_lacs[h] <= Phy[t,h] <= Pmax_hy_lacs[h])
# hydro stock constraint
@constraint(model, [h in 1:Nhy], sum(Phy[t,h] for t in 1:Tmax) <= e_hy_lacs[h])

# weekly STEP
@constraint(model, [t in 1:Tmax], Pcharge_STEP[t] <= Pmax_STEP)
@constraint(model, [t in 1:Tmax], Pdecharge_STEP[t] <= Pmax_STEP)
@constraint(model, stock_STEP[1] == 0)
@constraint(model, Pdecharge_STEP[Tmax] <= stock_STEP[Tmax])
@constraint(model, stock_STEP[Tmax] == stock_STEP[1])
@constraint(model, Pdecharge_STEP[1] == 0)
@constraint(model, [t in 1:Tmax-1], stock_STEP[t+1]-stock_STEP[t]- rSTEP*Pcharge_STEP[t]+Pdecharge_STEP[t]== 0)
@constraint(model, [t in 1:Tmax], stock_STEP[t] <= 24*7*Pmax_STEP)

# #battery
@constraint(model, [t in 1:Tmax], Pcharge_battery[t] <= CapaBattery)
@constraint(model, [t in 1:Tmax], Pdecharge_battery[t] <= CapaBattery)
@constraint(model, stock_battery[1] == 0)
@constraint(model, Pdecharge_battery[Tmax] <= stock_battery[Tmax])
@constraint(model, stock_battery[Tmax] == stock_battery[1])
@constraint(model, Pdecharge_battery[1] == 0)
@constraint(model, [t in 1:Tmax-1], stock_battery[t+1]-stock_battery[t]- rbattery*Pcharge_battery[t]+1/rbattery*Pdecharge_battery[t]== 0)
@constraint(model, [t in 1:Tmax], stock_battery[t] <= d_battery*CapaBattery)



# solve the model
optimize!(model)

#Results
@show termination_status(model)
@show objective_value(model)


#exports results as csv file
th_gen = value.(PH2_CCG)
hy_gen = value.(Phy)
STEP_charge = value.(Pcharge_STEP)
STEP_decharge = value.(Pdecharge_STEP)
battery_charge = value.(Pcharge_battery)
battery_decharge = value.(Pdecharge_battery)


println("\n==================== RÉSULTATS ====================")
println("Status : ", termination_status(model))
println("Coût total : ", round(objective_value(model), digits=2), " €")

println("\n--- Capacités optimales ---")
println("Onshore  : ", round(value(CapaOnshore), digits=2), " MW")
println("Offshore : ", round(value(CapaOffshore), digits=2), " MW")
println("Solar    : ", round(value(CapaSolar), digits=2), " MW")
println("Battery  : ", round(value(CapaBattery), digits=2), " MW")

println("\n--- CCG H2 installées ---")
for g in 1:NH2_CCG_max
    if value(CCG_H2_installed[g]) > 0.5
        println("Unité CCG H2 ", g, " installée")
    end
end

println("\n--- TACH2 installées ---")
for g in 1:NH2_TAC_max
    if value(TAC_H2_installed[g]) > 0.5
        println("Unité TAC H2 ", g, " installée")
    end
end

println("\n--- Génération horaire ---")
println("Défaillance : ", round(sum(value.(Puns)), digits=2), " MWh")
println("Excès d'énergie : ", round(sum(value.(Pexc)), digits=2), " MWh")

println("===================================================\n")

using JSON3

# Construction du dictionnaire du parc
parc = Dict(
    "status" => string(termination_status(model)),
    "cout_total_euros" => round(objective_value(model), digits=2),

    "capacites_MW" => Dict(
        "onshore" => round(value(CapaOnshore), digits=2),
        "offshore" => round(value(CapaOffshore), digits=2),
        "solar" => round(value(CapaSolar), digits=2),
        "battery" => round(value(CapaBattery), digits=2)
    ),

    "H2" => Dict(
        "CCG" => Dict(
            "nombre_installees" => sum(value.(CCG_H2_installed) .> 0.5),
            "unites_installees" => [
                g for g in 1:NH2_CCG_max if value(CCG_H2_installed[g]) > 0.5
            ],
            "production_totale_MWh" => round(sum(value.(PH2_CCG)), digits=2)
        ),

        "TAC" => Dict(
            "nombre_installees" => sum(value.(TAC_H2_installed) .> 0.5),
            "unites_installees" => [
                g for g in 1:NH2_TAC_max if value(TAC_H2_installed[g]) > 0.5
            ],
            "production_totale_MWh" => round(sum(value.(PH2_TAC)), digits=2)
        )
    ),

    "energie_totale_MWh" => Dict(
        "defaillance" => round(sum(value.(Puns)), digits=2),
        "exces" => round(sum(value.(Pexc)), digits=2),
        "production_H2_totale" => round(sum(value.(PH2_CCG)) + sum(value.(PH2_TAC)), digits=2),
        "production_hydro" => round(sum(value.(Phy)), digits=2)
    )
)

# Écriture dans fichier
open("parc_resultats.json", "w") do f
    JSON3.write(f, parc; indent=4)
end

println("Fichier parc_resultats.json généré avec succès ✅")

using DelimitedFiles

CCG_gen = value.(PH2_CCG)
TAC_gen = value.(PH2_TAC)
hy_gen = value.(Phy)

STEP_charge = value.(Pcharge_STEP)
STEP_decharge = value.(Pdecharge_STEP)

battery_charge = value.(Pcharge_battery)
battery_decharge = value.(Pdecharge_battery)

solar_gen = value(CapaSolar) .* solar_load_factor
onshore_gen = value(CapaOnshore) .* onshore_load_factor
offshore_gen = value(CapaOffshore) .* offshore_load_factor

open("results.csv", "w") do f

    # Header
    write(f, "t;H2_CCG;H2_TAC;Hydro;Solaire;Onshore;Offshore;")
    write(f, "STEP_charge;STEP_decharge;")
    write(f, "Battery_charge;Battery_decharge;")
    write(f, "Load;Net_load\n")

    for t in 1:Tmax

        # Somme des H2
        H2_CCG = sum(CCG_gen[t,g] for g in 1:NH2_CCG_max)
        H2_TAC = sum(TAC_gen[t,g] for g in 1:NH2_TAC_max)

        write(f,
            string(
                t, ";",
                round(H2_CCG, digits=2), ";",
                round(H2_TAC, digits=2), ";",
                round(hy_gen[t,1], digits=2), ";",
                round(solar_gen[t], digits=2), ";",
                round(onshore_gen[t], digits=2), ";",
                round(offshore_gen[t], digits=2), ";",
                round(-STEP_charge[t], digits=2), ";",
                round(STEP_decharge[t], digits=2), ";",
                round(-battery_charge[t], digits=2), ";",
                round(battery_decharge[t], digits=2), ";",
                round(load[t], digits=2), ";",
                round(load[t] - solar_gen[t] - onshore_gen[t] - offshore_gen[t], digits=2),
                "\n"
            )
        )
    end
end

println("Fichier results.csv généré avec succès ✅")

