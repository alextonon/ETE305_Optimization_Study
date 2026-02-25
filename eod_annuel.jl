include("extraction_donnees_excel.jl")

data_file = "data/Donnees_etude_de_cas_ETE305.xlsx"



## Extraction des données de la semaine
time_series = extraire_donnees_semaine(data_file, 1, AnneauGarde=24)

load = time_series["load"]
offshore_load_factor = time_series["offshore_load_factor"]
onshore_load_factor = time_series["onshore_load_factor"]
solar_load_factor = time_series["solar_load_factor"]
hydro_fatal = time_series["hydro_fatal"]
thermique_fatal = time_series["thermique_fatal"]

Pres = hydro_fatal 