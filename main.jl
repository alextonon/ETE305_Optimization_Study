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

#### Loading of CAPEX/OPEX data

types_centrales=["onshore", "offshore_pose", "offshore_flot", "pv_pose", "pv_gd_toit", "pv_pet_toit", "CCG_H2", "TAC_H2", "electrolyseur", "batterie"]
CAPEX=XLSX.readdata(data_file, "Investissements", "B2:B11") # CAPEX des différentes technologies en €/kW
OPEX=XLSX.readdata(data_file, "Investissements", "C2:C11") # OPEX des différentes technologies en €/kW/an
Duree_vie=XLSX.readdata(data_file, "Investissements", "D2:D11") # Durée de vie des différentes technologies en années

#data for h2 clusters
capex_CCG_H2 = CAPEX[7]*1000 #€/MW
opex_CCG_H2 = OPEX[7]*1000 #€/MW/year
PU_cost_h2_CCG = XLSX.readdata(data_file, "Parc électrique", "H9") #€/MWh basé sur le tarif de prod de la centrale CCG gaz A MODIFIER
Pmin_h2_CCG = XLSX.readdata(data_file, "Parc électrique", "F9") #MW idem
Pmax_h2_CCG = XLSX.readdata(data_file, "Parc électrique", "E9") #MW idem
dmin_CCG = XLSX.readdata(data_file, "Parc électrique", "G9") #hours idem

capex_TAC_H2 = CAPEX[8]*1000 #€/MW
opex_TAC_H2 = OPEX[8]*1000 #€/MW/year
PU_cost_h2_TAC = XLSX.readdata(data_file, "Parc électrique", "H10") #€/MWh basé sur le tarif de prod de la centrale TAC gaz A MODIFIER
Pmin_h2_TAC = XLSX.readdata(data_file, "Parc électrique", "F10") #MW idem
Pmax_h2_TAC = XLSX.readdata(data_file, "Parc électrique", "E10") #MW idem
dmin_TAC = XLSX.readdata(data_file, "Parc électrique", "G10") #hours idem

## temporary
# NH2_max_CCG = 10

# Pmin_h2_CCG = Pmin_h2_CCG*ones(NH2_max_CCG)
# Pmax_h2_CCG = Pmax_h2_CCG*ones(NH2_max_CCG)
# dmin_CCG = dmin_CCG*ones(Int, NH2_max_CCG)

NH2_max = 30

Pmin_h2 = Pmin_h2_CCG*ones(NH2_max)
Pmax_h2 = Pmax_h2_CCG*ones(NH2_max)
dmin = dmin_CCG*ones(Int, NH2_max)


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





#############################
# Verification of loaded data

println("=== Vérification des données chargées ===")

# @show load
# @show offshore_load_factor
# @show onshore_load_factor
# @show solar_load_factor
# @show hydro_fatal
# @show Pres
# @show thermique_fatal

# @show Usable_per_week_hydro_lacs
# @show Usable_per_week_hydro_STEP

# @show CapaSolar_init
# @show CapaOffshore_init
# @show CapaOnshore_init

# @show CAPEX
# @show OPEX
# @show Duree_vie

# @show capex_CCG_H2
# @show opex_CCG_H2
# @show PU_cost_h2_CCG
# @show Pmin_h2_CCG
# @show Pmax_h2_CCG
# @show dmin_CCG

# @show capex_TAC_H2
# @show opex_TAC_H2
# @show PU_cost_h2_TAC
# @show Pmin_h2_TAC
# @show Pmax_h2_TAC
# @show dmin_TAC

# @show Pmin_h2
# @show Pmax_h2
# @show dmin

# @show Nhy
# @show Pmin_hy_lacs
# @show Pmax_hy_lacs
# @show e_hy_lacs

# @show cuns
# @show cexc

# @show capex_onshore
# @show capex_offshore
# @show capex_solar
# @show capex_2h_battery

# @show opex_onshore
# @show opex_offshore
# @show opex_solar
# @show opex_2h_battery

# @show Pmax_STEP
# @show rSTEP

# @show rbattery
# @show d_battery
# @show CapaBattery_init

println("=== Fin de la vérification ===")


#############################
#create the optimization model
#############################
model = Model(HiGHS.Optimizer)

#############################
#define the variables
#############################
#energie renouvelables

@variable(model, CapaOnshore >= 0, start = CapaOnshore_init)
@variable(model, CapaOffshore >= 0, start = CapaOffshore_init)
@variable(model, CapaSolar >= 0, start = CapaSolar_init)

# @variable(model, capacity_offshore_fixed >= 0)
# @variable(model, capacity_offshore_floating >= 0)
# @variable(model, capacity_solar_ground >= 0)
# @variable(model, capacity_solar_big_roof >= 0)
# @variable(model, capacity_solar_small_roof >= 0)

#H2 generation variables

@variable(model, H2installed[1:NH2_max], Bin)                 # 1 si la centrale i est construite
@variable(model, H2running[1:Tmax, 1:NH2_max], Bin)         # 1 si ON à t
@variable(model, PH2[1:Tmax, 1:NH2_max] >= 0)

@variable(model, H2start[1:Tmax,1:NH2_max], Bin)
@variable(model, H2stop[1:Tmax,1:NH2_max], Bin)

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
@variable(model, CapaBattery >= 0, start = CapaBattery_init)
@variable(model, Pcharge_battery[1:Tmax] >= 0)
@variable(model, Pdecharge_battery[1:Tmax] >= 0)
@variable(model, stock_battery[1:Tmax] >= 0)
# #############################
#define the objective function
#############################
@objective(model, Min, 
                        CapaSolar*(capex_solar + opex_solar) + CapaOffshore*(capex_offshore + opex_offshore) + CapaOnshore*(capex_onshore + opex_onshore) 
                        + CapaBattery*(capex_2h_battery + opex_2h_battery) 
                        + sum(H2installed[g]*Pmax_h2[g] for g in 1:NH2_max)*(capex_CCG_H2 + opex_CCG_H2) + sum(PH2[t,g] for t in 1:Tmax, g in 1:NH2_max)*PU_cost_h2_CCG
                        + sum(Puns[t] for t in 1:Tmax)*cuns + sum(Pexc[t] for t in 1:Tmax)*cexc
)


#############################
#define the constraints
#############################
#balance constraint
@constraint(model, balance[t in 1:Tmax], sum(PH2[t,g] for g in 1:NH2_max) + sum(Phy[t,h] for h in 1:Nhy) + CapaSolar * solar_load_factor[t] + CapaOffshore * offshore_load_factor[t] + CapaOnshore * onshore_load_factor[t] + Puns[t] - load[t] - Pexc[t] - Pcharge_STEP[t] + Pdecharge_STEP[t] - Pcharge_battery[t] + Pdecharge_battery[t] == 0)

# H2 Power constraints
@constraint(model, max_H2[t in 1:Tmax, i in 1:NH2_max], PH2[t,i] <= Pmax_h2[i]*H2running[t,i]) #Pmax constraints
@constraint(model, min_H2[t in 1:Tmax, i in 1:NH2_max], Pmin_h2[i]*H2running[t,i] <= PH2[t,i]) #Pmin constraints

# H2 instalation constraints
@constraint(model, [t in 1:Tmax, g in 1:NH2_max], H2running[t,g] <= H2installed[g]) # only produce if installed

# H2 duration constraints
for g in 1:NH2_max
        if (dmin[g] > 1)
            @constraint(model, [t in 2:Tmax], H2running[t,g]-H2running[t-1,g]==H2start[t,g]-H2stop[t,g],  base_name = "stateH2_$g") # detect start and stop
            @constraint(model, [t in 1:Tmax], H2start[t,g]+H2stop[t,g]<=1,  base_name = "exclusiveH2_$g") # avoid starting and stoping at same step

            # Initial conditions
            @constraint(model, H2start[1,g]==0,  base_name = "iniStartH2_$g")
            @constraint(model, H2stop[1,g]==0,  base_name = "iniStopH2_$g")
            @constraint(model, [t in 1:dmin[g]-1], H2running[t,g] >= sum(H2start[i,g] for i in 1:t), base_name = "dminStartH2_$(g)_init")
            @constraint(model, [t in 1:dmin[g]-1], H2running[t,g] <= 1-sum(H2stop[i,g] for i in 1:t), base_name = "dminStopH2_$(g)_init")

            # Minimum up and down time constraints
            @constraint(model, [t in dmin[g]:Tmax], H2running[t,g] >= sum(H2start[i,g] for i in (t-dmin[g]+1):t),  base_name = "dminStartH2_$g")
            @constraint(model, [t in dmin[g]:Tmax], H2running[t,g] <= 1 - sum(H2stop[i,g] for i in (t-dmin[g]+1):t),  base_name = "dminStopH2_$g")
end

# H2 volume constraints
RendementElectrolyse = 0.7 # Rendement de l'électrolyse
RendementCombustion = 0.5 # Rendement de la combustion de l'hydrogène
@constraint(model, sum(PH2[t,g] for t in 1:Tmax, g in 1:NH2_max) <= sum(Pexc[t] for t in 1:Tmax)*RendementCombustion*RendementElectrolyse)

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
th_gen = value.(PH2)
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

println("\n--- H2 installées ---")
for g in 1:NH2_max
    if value(H2installed[g]) > 0.5
        println("Unité H2 ", g, " installée")
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
        "nombre_installees" => sum(value.(H2installed) .> 0.5),
        "unites_installees" => [
            g for g in 1:NH2_max if value(H2installed[g]) > 0.5
        ]
    ),

    "energie_totale_MWh" => Dict(
        "defaillance" => round(sum(value.(Puns)), digits=2),
        "exces" => round(sum(value.(Pexc)), digits=2),
        "production_H2" => round(sum(value.(PH2)), digits=2),
        "production_hydro" => round(sum(value.(Phy)), digits=2)
    )
)

# Écriture dans fichier
open("parc_resultats.json", "w") do f
    JSON3.write(f, parc; indent=4)
end

println("Fichier parc_resultats.json généré avec succès ✅")

using DelimitedFiles

th_gen = value.(PH2)
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
    write(f, "t;H2_total;Hydro;Solaire;Onshore;Offshore;")
    write(f, "STEP_charge;STEP_decharge;")
    write(f, "Battery_charge;Battery_decharge;")
    write(f, "Load;Net_load\n")

    for t in 1:Tmax

        # Somme des H2
        H2_total = sum(th_gen[t,g] for g in 1:NH2_max)

        write(f,
            string(
                t, ";",
                round(H2_total, digits=2), ";",
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

end