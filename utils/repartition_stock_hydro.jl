using JuMP
using Ipopt
#package to read excel files
using XLSX

data_file = "data/Donnees_etude_de_cas_ETE305.xlsx"

#data for load and fatal generation
Prod_hydro_lacs = XLSX.readdata(data_file, "Détails historique hydro", "N3:N8737")
apports_mois=XLSX.readdata(data_file, "Détails historique hydro", "C2:C13")

A=24*31 # nombre d'heures dans un mois de 31 jours
B=24*28 # nombre d'heures dans un mois de 28 jours
C=24*30 # nombre d'heures dans un mois de 30 jours
T = length(Prod_hydro_lacs)                # nombre d'heures
nb_mois = length(apports_mois)
month_hours = [
    1:24*31-1,
    24*31:48,
    49:72
]
model = Model(Ipopt.Optimizer)

# --------------------
# VARIABLES
# --------------------

@variable(model, A[1:T] >= 0)          # apports horaires
@variable(model, S[1:T])               # stock horaire

# --------------------
# CONDITIONS INITIALES
# --------------------

@constraint(model, S[1] == S0)

# --------------------
# BILAN HORAIRE
# --------------------

@constraint(model, [t=1:T-1],
    S[t+1] == S[t] + A[t] - Prod_hydro_lacs[t]
)

# --------------------
# CONTRAINTE MENSUELLE
# --------------------

for m in 1:nb_months
    hours_m = month_hours[m]   # vecteur des indices horaires du mois m
    @constraint(model,
        sum(A[t] for t in hours_m) == A_month[m]
    )
end

# --------------------
# BORNES JOURNALIERES
# --------------------

for d in 1:nb_days
    hours_d = day_hours[d]   # indices horaires du jour d
    for t in hours_d
        @constraint(model, S[t] >= Smin[d])
        @constraint(model, S[t] <= Smax[d])
    end
end

# --------------------
# OBJECTIF : LISSAGE
# --------------------

@objective(model, Min,
    sum((A[t] - A[t-1])^2 for t in 2:T)
)

optimize!(model)

A_opt = value.(A)
S_opt = value.(S)