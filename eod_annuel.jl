include("extraction_donnees_excel.jl")

data_file = "data/Donnees_etude_de_cas_ETE305.xlsx"

# -------- Extraction des hypothèses du problèmes --------
config = extraire_donnees_config(data_file)

# Data for h2 clusters
capex_CCG_H2 = config["H2"]["CCG"]["capex"] #€/MW
opex_CCG_H2 = config["H2"]["CCG"]["opex"] #€/MW/year
PU_cost_h2_CCG = config["H2"]["CCG"]["PU_cost"] #€/MWh basé sur le tarif de prod de la centrale CCG gaz A MODIFIER
Pmin_h2_CCG = config["H2"]["CCG"]["Pmin"] #MW idem
Pmax_h2_CCG = config["H2"]["CCG"]["Pmax"] #MW idem
dmin_CCG = config["H2"]["CCG"]["dmin"] #hours idem

capex_TAC_H2 = config["H2"]["TAC"]["capex"] #€/MW
opex_TAC_H2 = config["H2"]["TAC"]["opex"] #€/MW/year
PU_cost_h2_TAC = config["H2"]["TAC"]["PU_cost"] #€/MWh basé sur le tarif de prod de la centrale TAC gaz A MODIFIER
Pmin_h2_TAC = config["H2"]["TAC"]["Pmin"] #MW idem
Pmax_h2_TAC = config["H2"]["TAC"]["Pmax"] #MW idem
dmin_TAC = config["H2"]["TAC"]["dmin"] #hours idem

# Renewables
capex_onshore = config["enr"]["onshore"]["capex"] #€/MW
capex_offshore = config["enr"]["offshore_pose"]["capex"] #€/MW
capex_solar = config["enr"]["pv_pose"]["capex"] #€/MW

opex_onshore = config["enr"]["onshore"]["opex"] #€/MW
opex_offshore = config["enr"]["offshore_pose"]["opex"] #€/MW
opex_solar = config["enr"]["pv_pose"]["opex"] #€/MW


# Hydro
Pmin_hy_lacs = zeros(1)
Pmax_hy_lacs = config["hydro"]["lacs"]["Pmax"] *ones(1) #MW
Pmax_STEP = XLSX.readdata(data_file, "Parc électrique", "C21") #MW
rSTEP = XLSX.readdata(data_file, "Rendements", "B10") #rendement au pompage 

#battery
capex_2h_battery = config["battery"]["capex"] #€/MW
opex_2h_battery = config["battery"]["opex"] #€/MW
rbattery = config["battery"]["rendement"] # rendement de la batterie au sockage ou au destockage
d_battery = config["battery"]["d_battery"] #hours

# Initial capacities 
CapaSolar_init = config["capacites_init"]["Solar"] #MW
CapaOffshore_init = config["capacites_init"]["Offshore"] #MW
CapaOnshore_init = config["capacites_init"]["Onshore"] #MW
CapaBattery_init = 0 #MW


# ----------------- Définition des variables annuelles -----------------
Nweeks = 52
Nhours = Nweeks * 7 * 24

solar_capacities = Capasolar_init * ones(Nweeks)
offshore_capacities = CapaOffshore_init * ones(Nweeks)
onshore_capacities = CapaOnshore_init * ones(Nweeks)

solar_annual = zeros(Nhours)
offshore_annual= zeros(Nhours)
onshore_annual = zeros(Nhours)
battery_stock_annual = zeros(Nhours)
battery_charge_annual = zeros(Nhours)
battery_discharge_annual = zeros(Nhours)

stock_battery_initial = 0
stock_STEP_initial = 0

for week in 0:Nweeks
    time_series = extraire_donnees_semaine(data_file, week, AnneauGarde=24)

    load = time_series["load"]
    offshore_load_factor = time_series["offshore_load_factor"]
    onshore_load_factor = time_series["onshore_load_factor"]
    solar_load_factor = time_series["solar_load_factor"]
    hydro_fatal = time_series["hydro_fatal"]
    thermique_fatal = time_series["thermique_fatal"]

    Pres = hydro_fatal + thermique_fatal