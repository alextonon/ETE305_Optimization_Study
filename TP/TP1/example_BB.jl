#packages
using JuMP
#use the solver you want
using HiGHS
#test XLSX
using XLSX



model = Model(HiGHS.Optimizer)
@variable(model, x>=0)
@variable(model, y>=0)
@objective(model, Min, x-2*y)
@constraint(model, ctr1,-4*x+6*y<=9)
@constraint(model, ctr2, x+y<=4)
@constraint(model, b, y<=2)
optimize!(model)
#------------------------------
#Results
@show termination_status(model)
@show objective_value(model)
@show value(x)
@show value(y)
