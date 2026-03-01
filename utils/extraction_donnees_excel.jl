using XLSX

col_conso = "C"  # colonne de consommation électrique totale dans le fichier excel
col_fc_onshore = "E" # colonne de la capacité de production onshore
col_fc_offshore = "F" # colonne de la capacité de production offshore
col_fc_solaire = "H" # colonne de la capacité de production solaire
col_fil_eau = "I" # colonne de la capacité de production hydroélectrique (fil de l'eau)
col_hydro_lac = "J" # colonne de la capacité de production hydroélectrique (lacs)
col_th_fatal = "L"


function extraire_donnees_semaine(data_file::String, num_semaine::Int; AnneauGarde::Int=24)
    # Calcul des lignes à extraire
    Tmax = 7*24 + AnneauGarde
    row_start = 3 + (num_semaine-1)*7*24
    row_end = row_start + Tmax - 1

    # --- Consommation ---
    range_conso = string(col_conso, row_start, ":", col_conso, row_end)
    conso = XLSX.readdata(data_file, "Consommation_elec_totale", range_conso)

    # --- Hydro lac ---
    range_hydro_lac = string(col_hydro_lac, row_start, ":", col_hydro_lac, row_end)
    capa_hydro_lac = XLSX.readdata(data_file, "Consommation_elec_totale", range_hydro_lac)
    hydro_lac_dispo = sum(capa_hydro_lac[1:Tmax-24])

    # --- Hydro fil de l'eau ---
    range_fil_eau = string(col_fil_eau, row_start, ":", col_fil_eau, row_end)
    hydro_fil_eau = XLSX.readdata(data_file, "Consommation_elec_totale", range_fil_eau)

    # --- Facteurs de charge ---
    range_fc_onshore = string(col_fc_onshore, row_start, ":", col_fc_onshore, row_end)
    fc_onshore = XLSX.readdata(data_file, "Consommation_elec_totale", range_fc_onshore)

    range_fc_offshore = string(col_fc_offshore, row_start, ":", col_fc_offshore, row_end)
    fc_offshore = XLSX.readdata(data_file, "Consommation_elec_totale", range_fc_offshore)

    range_fc_solaire = string(col_fc_solaire, row_start, ":", col_fc_solaire, row_end)
    fc_solaire = XLSX.readdata(data_file, "Consommation_elec_totale", range_fc_solaire)

    # --- Thermique ---
    range_th_fatal = string(col_th_fatal, row_start, ":", col_th_fatal, row_end)
    th_fatal = XLSX.readdata(data_file, "Consommation_elec_totale", range_th_fatal)

    return Dict(
        "load" => conso,
        "Usable_per_week_hydro_lacs" => hydro_lac_dispo,
        "hydro_fatal" => hydro_fil_eau,
        "onshore_load_factor" => fc_onshore,
        "offshore_load_factor" => fc_offshore,
        "solar_load_factor" => fc_solaire,
        "thermique_fatal" => th_fatal,
        "Tmax" => Tmax
    )
end

function extraire_donnees_config(data_file::String)
    types_centrales = ["onshore", "offshore_pose", "offshore_flot", "pv_pose", 
                       "pv_gd_toit", "pv_pet_toit", "CCG_H2", "TAC_H2", 
                       "electrolyseur", "batterie"]

    # --- Capacités initiales ---
    capacites_init = Dict(
        "Solar" => XLSX.readdata(data_file, "Parc électrique", "C24"),
        "Offshore" => XLSX.readdata(data_file, "Parc électrique", "C23"),
        "Onshore" => XLSX.readdata(data_file, "Parc électrique", "C22")
    )

    # --- CAPEX / OPEX / Durée de vie ---
    CAPEX = XLSX.readdata(data_file, "Investissements", "B2:B11")
    OPEX = XLSX.readdata(data_file, "Investissements", "C2:C11")
    Duree_vie = XLSX.readdata(data_file, "Investissements", "D2:D11")

    # Créer un dictionnaire centralisé pour toutes les centrales
    centrales = Dict()
    for (i, type_c) in enumerate(types_centrales)
        centrales[type_c] = Dict(
            "opex" => OPEX[i]*1000/52,     # €/MW/an
            "duree_vie" => Duree_vie[i], # années
            "capex" => CAPEX[i]*1000/Duree_vie[i]/52 # €/MW
        )
    end

    gisement_solaire = XLSX.readdata(data_file, "Gisements", "B3") # MW
    gisement_onshore = XLSX.readdata(data_file, "Gisements", "B4") # MW
    gisement_offshore = XLSX.readdata(data_file, "Gisements", "B5") # MW

    centrales["gisement"] = Dict(
        "Solar" => gisement_solaire,
        "Onshore" => gisement_onshore,
        "Offshore" => gisement_offshore
    )


    RendementElectrolyse = XLSX.readdata(data_file, "Rendements", "B8") # Rendement de l'électrolyse
    RendementCombustion = XLSX.readdata(data_file, "Rendements", "B11") # Rendement de la combustion de l'hydrogène

    rendements = Dict(
        "electrolyse" => RendementElectrolyse,
        "combustion" => RendementCombustion
    )

    # CAPEX / OPEX / Durée de vie
    CAPEX = XLSX.readdata(data_file, "Investissements", "B2:B11")
    OPEX = XLSX.readdata(data_file, "Investissements", "C2:C11")
    Duree_vie = XLSX.readdata(data_file, "Investissements", "D2:D11")

    # H2 clusters
    H2 = Dict(
        "CCG" => Dict(
            "opex" => OPEX[7]*1000/52, #€/MW/ year
            "duree_vie" => Duree_vie[7], # years
            "capex" => CAPEX[7]*1000/Duree_vie[7]/52, #€/MW
            "PU_cost" => XLSX.readdata(data_file, "Parc électrique", "H9"),  #€/MWh basé sur le tarif de prod de la centrale CCG gaz A MODIFIER
            "Pmin" => XLSX.readdata(data_file, "Parc électrique", "F9"), #MW
            "Pmax" => XLSX.readdata(data_file, "Parc électrique", "E9"), #MW
            "dmin" => XLSX.readdata(data_file, "Parc électrique", "G9"), #h
            "gisement" => XLSX.readdata(data_file, "Gisements", "B12") # Nombre
            
        ),
        "TAC" => Dict(
            "opex" => OPEX[8]*1000/52, #€/MW/ year
            "duree_vie" => Duree_vie[8], # years
            "capex" => CAPEX[8]*1000/Duree_vie[8]/52, #€/MW
            "PU_cost" => XLSX.readdata(data_file, "Parc électrique", "H10"),
            "Pmin" => XLSX.readdata(data_file, "Parc électrique", "F10"),
            "Pmax" => XLSX.readdata(data_file, "Parc électrique", "E10"),
            "dmin" => XLSX.readdata(data_file, "Parc électrique", "G10"),
            "gisement" => XLSX.readdata(data_file, "Gisements", "B13") # Nombre
           
        )
    )

    # Défaillance
    cuns = XLSX.readdata(data_file, "Defaillance", "B2")  #cost of unsupplied energy €/MWh
    cexc = XLSX.readdata(data_file, "Defaillance", "B3") #cost of in excess energy €/MWh

    Pmax_STEP = XLSX.readdata(data_file, "Parc électrique", "C21") #MW
    rSTEP = XLSX.readdata(data_file, "Rendements", "B10") #rendement au pompage 
    Pmax_hy_lacs = XLSX.readdata(data_file, "Parc électrique", "C20") #MW

    hydro = Dict(
        "lacs" => Dict(
            "Pmax" => Pmax_hy_lacs,
        ),
        "STEP" => Dict(
            "Pmax" => Pmax_STEP,
            "rendement_pompage" => rSTEP
        )
    )

    #battery
    rbattery = XLSX.readdata(data_file, "Rendements", "B9") # rendement de la batterie au sockage ou au destockage
    d_battery = XLSX.readdata(data_file, "Investissements", "E11") #hours
    CapaBattery_init = 0 #MW

    battery = Dict(
        "opex" => centrales["batterie"]["opex"]/52, #€/MW/semaine
        "duree_vie" => centrales["batterie"]["duree_vie"],
        "capex" => centrales["batterie"]["capex"]/centrales["batterie"]["duree_vie"]/52, #€/MW/semaine
        "rendement" => rbattery,
        "d_battery" => d_battery,
        "CapaBattery_init" => CapaBattery_init,
        "gisement" => XLSX.readdata(data_file, "Gisements", "B8") #MW
    )

    defaillance = Dict(
        "cost_unsupplied" => cuns,
        "cost_excess" => cexc
    )

    return Dict(
        "capacites_init" => capacites_init,
        "enr" => centrales,
        "H2" => H2,
        "hydro" => hydro,
        "battery" => battery,
        "defaillance" => defaillance,
        "rendements" => rendements
    )
end