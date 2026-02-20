#packages
using JuMP
#use the solver you want
using HiGHS
#package to read excel files
using XLSX

num_semaine=1 #numéro de la semaine à extraire (semaine 1 = du 04/07/2020 au 10/07/2020, semaine 2 = du 11/07/2020 au 17/07/2020, etc.) 

Tmax = 192 # extraction for 1 week + 1 day (8*24=192 hours)
col = "C"  # colonne de consommation électrique totale dans le fichier excel
row_start = 3+(num_semaine-1)*7*24 # ligne de début de la semaine à extraire
row_end = row_start+Tmax-1 # ligne de fin de la semaine à extraire

range = string(col, row_start, ":", col, row_end)
#data for load and fatal generation
data_file = "data/Donnees_etude_de_cas_ETE305.xlsx"

#data for load and fatal generation
conso = XLSX.readdata(data_file, "Consommation_elec_totale", range)
@show conso


