#packages
using JuMP
#use the solver you want
using HiGHS
#test XLSX
using XLSX

model = Model(HiGHS.Optimizer)
@variable(model, 0<= P_nuc1 <=900)
@variable(model, 0<= P_nuc2 <=600)
@variable(model, 0<= P_ccg <=300)
@variable(model, 0<= P_hydro <=300)
@variable(model, 0<= P_eolien <=300)

@objective(model, Min, 14*P_nuc1 + 16*P_nuc2 + 45*P_ccg + 48*P_hydro + 0*P_eolien)
@constraint(model, demand, P_nuc1 + P_nuc2 + P_ccg + P_hydro + P_eolien == 2200)

optimize!(model)
#------------------------------
#Results
@show termination_status(model)
@show objective_value(model)
@show value(P_nuc1)
@show value(P_nuc2)
@show value(P_ccg)
@show value(P_hydro)
@show value(P_eolien)