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
        "thermique_fatal" => th_fatal
    )
end