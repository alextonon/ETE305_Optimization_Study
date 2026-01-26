#packages
using JuMP
#use the solver you want
using HiGHS


# ###########################
# ### 1. EOD tres simple
# ###########################
# #create the optimization model
# model = Model(HiGHS.Optimizer)
# #------------------------------
# #define the variables
# #Nucleaire 1
# @variable(model, Pnuc1 >= 0)
# #Nucleaire 2
# @variable(model, Pnuc2 >= 0)
# #CCG
# @variable(model, Pccg >= 0)
# #Hydro
# @variable(model, Phydro >= 0)
# #Eolien
# @variable(model, Peolien >= 0)
# #------------------------------
# #define the objective function
# @objective(model, Min, 14*Pnuc1+16*Pnuc2+45*Pccg+48*Phydro+0*Peolien)
# #------------------------------
# #define the constraints
# #la demande de 2200MWh doit être satisfaite
# @constraint(model, eod, Pnuc1+Pnuc2+Pccg+Phydro+Peolien==2200)
# #contraintes de production
# @constraint(model, maxnuc1, Pnuc1 <= 900)
# @constraint(model, maxnuc2, Pnuc2 <= 900)
# @constraint(model, maxccg, Pccg <= 300)
# @constraint(model, maxhydro, Phydro <= 300)
# @constraint(model, maxeolien, Peolien <= 300)
# #------------------------------
# #print the model
# print(model)
# #------------------------------
# #solve the model
# optimize!(model)
# #------------------------------
# #Results
# @show termination_status(model)
# @show objective_value(model)
# @show value(Pnuc1)
# @show value(Pnuc2)
# @show value(Pccg)
# @show value(Phydro)
# @show value(Peolien)

# ###########################
# ### 2. EOD avec contrainte de minimum de fonctionnement => introduction des variables binaires
# ###########################
# #create the optimization model
# model = Model(HiGHS.Optimizer)
# #------------------------------
# #define the variables
# #Nucleaire 1
# @variable(model, Pnuc1 >= 0)
# @variable(model, UCnuc1, Bin) #Nucleaire 1 on (UCnuc1=1) ou off (UCnuc1=0)
# #Nucleaire 2
# @variable(model, Pnuc2 >= 0)
# @variable(model, UCnuc2, Bin)
# #CCG
# @variable(model, Pccg >= 0)
# @variable(model, UCccg, Bin)
# #Hydro
# @variable(model, Phydro >= 0)
# #Eolien
# @variable(model, Peolien >= 0)
# #------------------------------
# #define the objective function
# @objective(model, Min, 14*Pnuc1+16*Pnuc2+45*Pccg+48*Phydro+0*Peolien)
# #------------------------------
# #define the constraints
# #la demande de 2200MWh doit être satisfaite
# @constraint(model, eod, Pnuc1+Pnuc2+Pccg+Phydro+Peolien==2200)
# #contraintes de production
# @constraint(model, maxnuc1, Pnuc1 <= UCnuc1*900)
# @constraint(model, maxnuc2, Pnuc2 <= UCnuc2*900)
# @constraint(model, maxccg, Pccg <= UCccg*300)

# @constraint(model, minnuc1, Pnuc1>=UCnuc1*300)
# @constraint(model, minnuc2, Pnuc2>=UCnuc2*300)
# @constraint(model, minccg, Pccg>=UCccg*150)
# #...
# @constraint(model, maxhydro, Phydro <= 300)
# @constraint(model, maxeolien, Peolien <= 300)

# #------------------------------
# #print the model
# print(model)
# #------------------------------
# #solve the model
# optimize!(model)
# #------------------------------
# #Results
# @show termination_status(model)
# @show objective_value(model)
# @show value(Pnuc1)
# @show value(Pnuc2)
# @show value(Pccg)
# @show value(Phydro)
# @show value(Peolien)

# ###########################
# ### 3. EOD avec contrainte de rampe => problème d'optimisation sur plusieurs pas de temps et contraintes supplémentaires
# ###########################
# #create the optimization model
# model = Model(HiGHS.Optimizer)
# #------------------------------
# #define the variables
# #Nombre de pas de temps
# T = 3
# #Demande pour chaque pas de temps
# demande = [2200, 2450, 1900]
# #Nucleaire 1
# @variable(model, Pnuc1[1:T] >= 0)
# #Nucleaire 1 on (UCnuc1=1) ou off (UCnuc1=0)
# @variable(model, UCnuc1[1:T], Bin)
# #Nucleaire 2
# @variable(model, Pnuc2[1:T] >= 0)
# @variable(model, UCnuc2[1:T], Bin)
# #CCG
# @variable(model, Pccg[1:T] >= 0)
# @variable(model, UCccg[1:T], Bin)
# #Hydro
# @variable(model, Phydro[1:T] >= 0)
# #Eolien
# @variable(model, Peolien[1:T] >= 0)
# #------------------------------
# #define the objective function
# @objective(model, Min, sum(14*Pnuc1[t]+16*Pnuc2[t]+45*Pccg[t]+48*Phydro[t]+0*Peolien[t] for t in 1:T))
# #------------------------------
# #define the constraints
# #la demande doit être satisfaite à chaque heure t
# @constraint(model, eod[t in 1:T], Pnuc1[t]+Pnuc2[t]+Pccg[t]+Phydro[t]+Peolien[t]==demande[t])
# #contraintes de production
# @constraint(model, maxnuc1[t in 1:T], Pnuc1[t] <= UCnuc1[t]*900)
# @constraint(model, maxnuc2[t in 1:T], Pnuc2[t] <= UCnuc2[t]*900)
# @constraint(model, maxccg[t in 1:T], Pccg[t]<= UCccg[t]*300)

# @constraint(model, minnuc1[t in 1:T], Pnuc1[t]>=UCnuc1[t]*300)
# @constraint(model, minnuc2[t in 1:T], Pnuc2[t]>=UCnuc2[t]*300)
# @constraint(model, minccg[t in 1:T], Pccg[t]>=UCccg[t]*150)

# @constraint(model, maxhydro[t in 1:T], Phydro[t] <= 300)
# @constraint(model, maxeolien[t in 1:T], Peolien[t] <= 300)
# #contraintes de limitation de variation de la puissance
# @constraint(model, rampnuc1[t in 1:T-1], -350 <= Pnuc1[t+1] - Pnuc1[t] <= 350)
# @constraint(model, rampnuc2[t in 1:T-1], -350 <= Pnuc2[t+1] - Pnuc2[t] <= 350)
# @constraint(model, rampccg[t in 1:T-1], -200 <= Pccg[t+1] - Pccg[t] <= 200)
# #...
# #------------------------------
# #print the model
# print(model)
# #------------------------------
# #solve the model
# optimize!(model)
# #------------------------------
# #Results
# @show termination_status(model)
# @show objective_value(model)
# @show value.(Pnuc1)
# @show value.(Pnuc2)
# @show value.(Pccg)
# @show value.(Phydro)
# @show value.(Peolien)

###########################
### 4. EOD avec contrainte de Dmin => ajout de variables binaires supplémentaires
###########################
#create the optimization model
model = Model(HiGHS.Optimizer)
#------------------------------
#define the variables
#Nombre de pas de temps
T = 5
#Demande pour chaque pas de temps
demande = [2200, 2450, 1900, 1450, 1200]
#Nucleaire 1
@variable(model, Pnuc1[1:T] >= 0)
#Nucleaire 1 on (UCnuc1=1) ou off (UCnuc1=0)
@variable(model, UCnuc1[1:T], Bin)
@variable(model, UPnuc1[1:T], Bin)
@variable(model, DOnuc1[1:T], Bin)
#Nucleaire 2
@variable(model, Pnuc2[1:T] >= 0)
@variable(model, UCnuc2[1:T], Bin)
@variable(model, UPnuc2[1:T], Bin)
@variable(model, DOnuc2[1:T], Bin)
#CCG
@variable(model, Pccg[1:T] >= 0)
@variable(model, UCccg[1:T], Bin)
@variable(model, UPccg[1:T], Bin)
@variable(model, DOccg[1:T], Bin)
#Hydro
@variable(model, Phydro[1:T] >= 0)
#Eolien
@variable(model, Peolien[1:T] >= 0)
#------------------------------
#define the objective function
@objective(model, Min, sum(14*Pnuc1[t]+16*Pnuc2[t]+45*Pccg[t]+48*Phydro[t]+0*Peolien[t] for t in 1:T))
#------------------------------
#define the constraints
#la demande doit être satisfaite à chaque heure t
@constraint(model, eod[t in 1:T], Pnuc1[t]+Pnuc2[t]+Pccg[t]+Phydro[t]+Peolien[t]==demande[t])
#contraintes de production
@constraint(model, maxnuc1[t in 1:T], Pnuc1[t] <= UCnuc1[t]*900)
@constraint(model, maxnuc2[t in 1:T], Pnuc2[t] <= UCnuc2[t]*900)
@constraint(model, maxccg[t in 1:T], Pccg[t]<= UCccg[t]*300)
# 
@constraint(model, minnuc1[t in 1:T], Pnuc1[t]>=UCnuc1[t]*300)
@constraint(model, minnuc2[t in 1:T], Pnuc2[t]>=UCnuc2[t]*300)
@constraint(model, minccg[t in 1:T], Pccg[t]>=UCccg[t]*150)
# 
@constraint(model, maxhydro[t in 1:T], Phydro[t] <= 300)
@constraint(model, maxeolien[t in 1:T], Peolien[t] <= 300)
#contraintes de limitation de variation de la puissance
@constraint(model, rampnuc1[t in 1:T-1], -350 <= Pnuc1[t+1] - Pnuc1[t] <= 350)
@constraint(model, rampnuc2[t in 1:T-1], -350 <= Pnuc2[t+1] - Pnuc2[t] <= 350)
@constraint(model, rampccg[t in 1:T-1], -200 <= Pccg[t+1] - Pccg[t] <= 200)
#contraintes liant UC, UP et DO
#TODO: ajouter contraintes liant UC, UP et DO pour nuc1, nuc2 et CCG (attention penser à l'initialisation de UP et DO)
#Il y a 4 contraintes pour chaque moyen de production donc 12 contraintes au total
@constraint(model, UPnuc1_t[t in 2:T], UPnuc1[t] >= UCnuc1[t]-UCnuc1[t-1] )
@constraint(model, DOnuc1_t[t in 2:T], DOnuc1[t] >= UCnuc1[t-1]-UCnuc1[t] )

@constraint(model, UPnuc2_t[t in 2:T], UPnuc2[t] >= UCnuc2[t]-UCnuc2[t-1] )
@constraint(model, DOnuc2_t[t in 2:T], DOnuc2[t] >=  UCnuc2[t-1]-UCnuc2[t] )

@constraint(model, UPccg_t[t in 2:T], UPccg[t] >= UCccg[t]-UCccg[t-1] )
@constraint(model, DOccg_t[t in 2:T], DOccg[t] >= UCccg[t-1]-UCccg[t] )
#initialisation de UP et DO à t=1
@constraint(model, UPnuc1_0, UPnuc1[1] == 0)
@constraint(model, DOnuc1_0, DOnuc1[1] == 0 )
@constraint(model, UPnuc2_0, UPnuc2[1] == 0 )
@constraint(model, DOnuc2_0, DOnuc2[1] == 0 )
@constraint(model, UPccg_0, UPccg[1] == 0 )
@constraint(model, DOccg_0, DOccg[1] == 0 )
#contraintes de Dmin de fonctionnement
#TODO: ajouter contraintes de dmin pour la marche pour nuc1, nuc2 et CCG
@constraint(model, test[t in 1:T-2], (UCnuc1[t] + UCnuc1[t+1]+ UCnuc1[t+2]) >= UPnuc1[t]*3) #faudrait remplacer par 24
#...
#------------------------------
#print the model
print(model)
#------------------------------
#solve the model
optimize!(model)
#------------------------------
#Results
@show termination_status(model)
@show objective_value(model)
@show value.(Pnuc1)
@show value.(Pnuc2)
@show value.(Pccg)
@show value.(Phydro)
@show value.(Peolien)
