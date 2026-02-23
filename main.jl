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
offshore_load_factor = XLSX.readdata(data_file, "Résumé", "I2:I193")  
onshore_load_factor = XLSX.readdata(data_file, "Résumé", "I2:I193")
solar_load_factor = XLSX.readdata(data_file, "Résumé", "J2:J193")
hydro_fatal = XLSX.readdata(data_file, "Résumé", "K2:K193")
Pres = hydro_fatal 
thermique_fatal = XLSX.readdata(data_file, "Résumé", "L2:L193")

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
CAPEX_H2_CCG = CAPEX[7] #k€/MW
OPEX_H2_CCG = OPEX[7] #k€/MW/year
PU_cost_h2_CCG = XLSX.readdata(data_file, "Parc électrique", "H9") #€/MWh basé sur le tarif de prod de la centrale CCG gaz A MODIFIER
Pmin_h2_CCG = XLSX.readdata(data_file, "Parc électrique", "F9") #MW idem
Pmax_h2_CCG = XLSX.readdata(data_file, "Parc électrique", "E9") #MW idem
dmin_CCG = XLSX.readdata(data_file, "Parc électrique", "G9") #hours idem

CAPEX_H2_TAC = CAPEX[8] #k€/MW
OPEX_H2_TAC = OPEX[8] #k€/MW/year
PU_cost_h2_TAC = XLSX.readdata(data_file, "Parc électrique", "H10") #€/MWh basé sur le tarif de prod de la centrale TAC gaz A MODIFIER
Pmin_h2_TAC = XLSX.readdata(data_file, "Parc électrique", "F10") #MW idem
Pmax_h2_TAC = XLSX.readdata(data_file, "Parc électrique", "E10") #MW idem
dmin_TAC = XLSX.readdata(data_file, "Parc électrique", "G10") #hours idem

#data for hydro reservoir "lacs"
Nhy = 1 #number of hydro generation units
Pmin_hy_lacs = zeros(Nhy)
Pmax_hy_lacs = XLSX.readdata(data_file, "Parc électrique", "C20") *ones(Nhy) #MW
e_hy_lacs = Usable_per_week_hydro_lacs*ones(Nhy) #MWh


# variable costs
ch2 = repeat(costs_h2', Tmax) #cost of hydrogen generation €/MWh
cuns = 5000*ones(Tmax) #cost of unsupplied energy €/MWh
cexc = 0*ones(Tmax) #cost of in excess energy €/MWh

# #investment costs
# capex_onshore = XLSX.readdata(data_file, "Investissements", "B2") #€/MW
# capex_offshore_fixed = XLSX.readdata(data_file, "Investissements", "B3") #€/MW
# capex_offshore_floating = XLSX.readdata(data_file, "Investissements", "B4") #€/MW
# capex_solar_ground = XLSX.readdata(data_file, "Investissements", "B5") #€/MW
# capex_solar_big_roof = XLSX.readdata(data_file, "Investissements", "B6") #€/MW
# capex_solar_small_roof = XLSX.readdata(data_file, "Investissements", "B7") #€/MW
# capex_CCG_H2 = XLSX.readdata(data_file, "Investissements", "B8") #€/MW
# capex_TAC_H2 = XLSX.readdata(data_file, "Investissements", "B9") #€/MW
# capex_electrolyser_H2 = XLSX.readdata(data_file, "Investissements", "B10") #€/MW
# capex_2h_battery = XLSX.readdata(data_file, "Investissements", "B11") #€/MW

# #fixed opex 
# opex_onshore = XLSX.readdata(data_file, "Investissements", "C2") #€/MW/year
# opex_offshore_fixed = XLSX.readdata(data_file, "Investissements", "C3") #€/MW/year
# opex_offshore_floating = XLSX.readdata(data_file, "Investissements", "C4") #€/MW/year
# opex_solar_ground = XLSX.readdata(data_file, "Investissements", "C5") #€/MW/year
# opex_solar_big_roof = XLSX.readdata(data_file, "Investissements", "C6") #€/MW/year
# opex_solar_small_roof = XLSX.readdata(data_file, "Investissements", "C7") #€/MW/year
# opex_CCG_H2 = XLSX.readdata(data_file, "Investissements", "C8") #€/MW/year
# opex_TAC_H2 = XLSX.readdata(data_file, "Investissements", "C9") #€/MW/year
# opex_electrolyser_H2 = XLSX.readdata(data_file, "Investissements", "C10") #€/MW/year
# opex_2h_battery = XLSX.readdata(data_file, "Investissements", "C11") #€/MW/year

#data for STEP/battery
#weekly STEP
Pmax_STEP = XLSX.readdata(data_file, "Parc électrique", "C21") #MW
rSTEP = Usable_per_week_hydro_STEP

#battery
Capa_batteries = 280 #MW
rbattery = XLSX.readdata(data_file, "Rendements", "B9") # rendement de la batterie au sockage ou au destockage
d_battery = XLSX.readdata(data_file, "Investissements", "E11") #hours


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
NH2_max = 10 # nombre maximum d'unités de production d'hydrogène

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
@variable(model, Pcharge_battery[1:Tmax] >= 0)
@variable(model, Pdecharge_battery[1:Tmax] >= 0)
@variable(model, stock_battery[1:Tmax] >= 0)
# #############################
#define the objective function
#############################
@objective(model, Min, sum(Phy[t,h]*chy[h] for t in 1:Tmax, h in 1:Nhy) + sum(PH2[t,g]*cH2[g] for t in 1:Tmax, g in 1:NH2_max) + sum(Puns[t]*cuns[t] for t in 1:Tmax) + sum(Pexc[t]*cexc[t] for t in 1:Tmax))

#############################
#define the constraints
#############################
#balance constraint
@constraint(model, balance[t in 1:Tmax], sum(PH2[t,g] for g in 1:NH2_max) + sum(Phy[t,h] for h in 1:Nhy) + CapaSolaire * solar_load_factor[t] + CapaOffshore * offshore_load_factor[t] + CapaOnshore * onshore_load_factor[t] + Puns[t] - load[t] - Pexc[t] - Pcharge_STEP[t] + Pdecharge_STEP[t] - Pcharge_battery[t] + Pdecharge_battery[t] == 0)

# H2 Power constraints
@constraint(model, max_H2[t in 1:Tmax, i in 1:NH2_max], PH2[t,i] <= Pmax_h2[i]*H2running[t,i]) #Pmax constraints
@constraint(model, min_H2[t in 1:Tmax, i in 1:NH2_max], Pmin_h2[i]*H2running[t,i] <= PH2[t,i]) #Pmin constraints

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
@constraint(model, volume_H2, sum(PH2[t,g] for t in 1:Tmax, g in 1:NH2_max) <= sum(Pexc[t] for t in 1:Tmax)*RendementCombustion*RendementElectrolyse)

# hydro unit constraints
@constraint(model, bounds_hy[t in 1:Tmax, h in 1:Nhy], Pmin_hy[h] <= Phy[t,h] <= Pmax_hy[h])
# hydro stock constraint
@constraint(model, stock_hy[h in 1:Nhy], sum(Phy[t,h] for t in 1:Tmax) <= e_hy_week[h])

# weekly STEP
@constraint(model, Pcharge_max_STEP[t in 1:Tmax], Pcharge_STEP[t] <= Pmax_STEP)
@constraint(model, Pdecharge_max_STEP[t in 1:Tmax], Pdecharge_STEP[t] <= Pmax_STEP)
@constraint(model, init_stock_STEP, stock_STEP[1] == 0)
@constraint(model, end_Pdecharge_STEP, Pdecharge_STEP[Tmax] <= stock_STEP[Tmax])
@constraint(model, Tmax_stock_STEP, stock_STEP[Tmax] == stock_STEP[1])
@constraint(model, init_Pdecharge_STEP, Pdecharge_STEP[1] == 0)
@constraint(model, evol_stock_STEP[t in 1:Tmax-1], stock_STEP[t+1]-stock_STEP[t]- rSTEP*Pcharge_STEP[t]+Pdecharge_STEP[t]== 0)
@constraint(model, stock_max_STEP[t in 1:Tmax], stock_STEP[t] <= 24*7*Pmax_STEP)

# #battery
@constraint(model, Pcharge_max_battery[t in 1:Tmax], Pcharge_battery[t] <= Pmax_battery)
@constraint(model, Pdecharge_max_battery[t in 1:Tmax], Pdecharge_battery[t] <= Pmax_battery)
@constraint(model, init_stock_battery, stock_battery[1] == 0)
@constraint(model, end_Pdecharge_battery, Pdecharge_battery[Tmax] <= stock_battery[Tmax])
@constraint(model, Tmax_stock_battery, stock_battery[Tmax] == stock_battery[1])
@constraint(model, init_Pdecharge_battery, Pdecharge_battery[1] == 0)
@constraint(model, evol_stock_battery[t in 1:Tmax-1], stock_battery[t+1]-stock_battery[t]- rbattery*Pcharge_battery[t]+1/rbattery*Pdecharge_battery[t]== 0)
@constraint(model, stock_max_battery[t in 1:Tmax], stock_battery[t] <= d_battery*Pmax_battery)



##############################################################################################################################################################
#
# #############################
#define the objective function
#############################
@objective(model, Min, sum(Pth.*cth)+sum(Phy.*chy)+sum(PH2.*cH2)+Puns'cuns+Pexc'cexc)

#############################
#define the constraints
#############################
#balance constraint
@constraint(model, balance[t in 1:Tmax], sum(Pth[t,g] for g in 1:Nth) + sum(Phy[t,h] for h in 1:Nhy) + Pres[t] + Puns[t] - load[t] - Pexc[t] - Pcharge_STEP[t] + Pdecharge_STEP[t] - Pcharge_battery[t] + Pdecharge_battery[t] == 0)
# H2 constraints
@constraint(model, max_th[t in 1:Tmax, g in 1:Nth], Pth[t,g] <= Pmax_th[g]*UCth[t,g])
#thermal unit Pmin constraints
@constraint(model, min_th[t in 1:Tmax, g in 1:Nth], Pmin_th[g]*UCth[t,g] <= Pth[t,g])
#thermal unit Dmin constraints
for g in 1:Nth
        if (dmin[g] > 1)
            @constraint(model, [t in 2:Tmax], UCth[t,g]-UCth[t-1,g]==UPth[t,g]-DOth[t,g],  base_name = "fct_th_$g")
            @constraint(model, [t in 1:Tmax], UPth[t]+DOth[t]<=1,  base_name = "UPDOth_$g")
            @constraint(model, UPth[1,g]==0,  base_name = "iniUPth_$g")
            @constraint(model, DOth[1,g]==0,  base_name = "iniDOth_$g")
            @constraint(model, [t in dmin[g]:Tmax], UCth[t,g] >= sum(UPth[i,g] for i in (t-dmin[g]+1):t),  base_name = "dminUPth_$g")
            @constraint(model, [t in dmin[g]:Tmax], UCth[t,g] <= 1 - sum(DOth[i,g] for i in (t-dmin[g]+1):t),  base_name = "dminDOth_$g")
            @constraint(model, [t in 1:dmin[g]-1], UCth[t,g] >= sum(UPth[i,g] for i in 1:t), base_name = "dminUPth_$(g)_init")
            @constraint(model, [t in 1:dmin[g]-1], UCth[t,g] <= 1-sum(DOth[i,g] for i in 1:t), base_name = "dminDOth_$(g)_init")
    end
end

#hydro unit constraints
@constraint(model, bounds_hy[t in 1:Tmax, h in 1:Nhy], Pmin_hy[h] <= Phy[t,h] <= Pmax_hy[h])
#hydro stock constraint
@constraint(model, stock_hy[h in 1:Nhy], sum(Phy[t,h] for t in 1:Tmax) <= e_hy[h])

#weekly STEP
@constraint(model, Pcharge_max_STEP[t in 1:Tmax], Pcharge_STEP[t] <= Pmax_STEP)
@constraint(model, Pdecharge_max_STEP[t in 1:Tmax], Pdecharge_STEP[t] <= Pmax_STEP)
@constraint(model, init_stock_STEP, stock_STEP[1] == 0)
@constraint(model, end_Pdecharge_STEP, Pdecharge_STEP[Tmax] <= stock_STEP[Tmax])
@constraint(model, Tmax_stock_STEP, stock_STEP[Tmax] == stock_STEP[1])
@constraint(model, init_Pdecharge_STEP, Pdecharge_STEP[1] == 0)
@constraint(model, evol_stock_STEP[t in 1:Tmax-1], stock_STEP[t+1]-stock_STEP[t]- rSTEP*Pcharge_STEP[t]+Pdecharge_STEP[t]== 0)
@constraint(model, stock_max_STEP[t in 1:Tmax], stock_STEP[t] <= 24*7*Pmax_STEP)

# #battery
@constraint(model, Pcharge_max_battery[t in 1:Tmax], Pcharge_battery[t] <= Pmax_battery)
@constraint(model, Pdecharge_max_battery[t in 1:Tmax], Pdecharge_battery[t] <= Pmax_battery)
@constraint(model, init_stock_battery, stock_battery[1] == 0)
@constraint(model, end_Pdecharge_battery, Pdecharge_battery[Tmax] <= stock_battery[Tmax])
@constraint(model, Tmax_stock_battery, stock_battery[Tmax] == stock_battery[1])
@constraint(model, init_Pdecharge_battery, Pdecharge_battery[1] == 0)
@constraint(model, evol_stock_battery[t in 1:Tmax-1], stock_battery[t+1]-stock_battery[t]- rbattery*Pcharge_battery[t]+1/rbattery*Pdecharge_battery[t]== 0)
@constraint(model, stock_max_battery[t in 1:Tmax], stock_battery[t] <= d_battery*Pmax_battery)




#TODO: solve and analyse the results
#solve the model
optimize!(model)
#------------------------------
#Results
@show termination_status(model)
@show objective_value(model)


#exports results as csv file
th_gen = value.(Pth)
hy_gen = value.(Phy)
STEP_charge = value.(Pcharge_STEP)
STEP_decharge = value.(Pdecharge_STEP)
battery_charge = value.(Pcharge_battery)
battery_decharge = value.(Pdecharge_battery)


# new file created
touch("results.csv")

# file handling in write mode
f = open("results.csv", "w")

for name in names
    write(f, "$name ;")
end
write(f, "Hydro;STEP pompage;STEP turbinage;Batterie injection;Batterie soutirage;RES;load;Net load\n")

for t in 1:Tmax
    for g in 1:Nth
        write(f, "$(th_gen[t,g]);")
    end
    for h in 1:Nhy
        write(f, "$(hy_gen[t,h]) ;")
    end
    write(f, "$(STEP_charge[t]);$(STEP_decharge[t]) ;")
    write(f, "$(battery_charge[t]);$(battery_decharge[t]) ;")
    write(f, "$(Pres[t]); $(load[t]);$(load[t]-Pres[t]) \n")

end

close(f)