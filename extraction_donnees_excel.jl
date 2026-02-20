#packages
using JuMP
#use the solver you want
using HiGHS
#package to read excel files
using XLSX


############## EXTRACTION DES CONSO A LA SEMAINE + ANNEAU DE GARDE (1 JOUR)

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

################### EXTRACTION DES STOCKS HYDRO ##########

# On résonne sur le destockage à la semaine et non sur la différence de stock (qui n'est pas accessible au pas de temps horaire) 
range_hydro_lac = string("N", row_start, ":", "N", row_end)
capa_hydro_lac = XLSX.readdata(data_file, "Détails historique hydro",range_hydro_lac) # capacité de stockage du lac en MWh
utilisable_hydro_lac = sum(capa_hydro_lac[1:Tmax-24]) # quantité totale d'énergie hydroélectrique disponible sur la semaine (en MWh)


range_hydro_STEP = string("O", row_start, ":", "O", row_end)
capa_hydro_STEP = XLSX.readdata(data_file, "Détails historique hydro",range_hydro_STEP) # capacité de stockage du lac STEP en MWh
utilisable_hydro_STEP = sum(capa_hydro_STEP[1:Tmax-24]) # quantité totale d'énergie hydroélectrique disponible sur la semaine (en MWh)


################### EXTRACTION DES CAPEX ET OPEX DES INSTALLATIONS ##########

