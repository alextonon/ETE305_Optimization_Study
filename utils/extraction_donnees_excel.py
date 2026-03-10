import openpyxl


def extraire_donnees_config(data_file: str):

    wb = openpyxl.load_workbook(data_file, data_only=True)

    types_centrales = [
        "onshore", "offshore_pose", "offshore_flot", "pv_pose",
        "pv_gd_toit", "pv_pet_toit", "CCG_H2", "TAC_H2",
        "electrolyseur", "batterie"
    ]

    # ==============================
    # --- Capacités initiales ---
    # ==============================

    ws_parc = wb["Parc électrique"]

    capacites_init = {
        "Solar": ws_parc["C24"].value,
        "Offshore": ws_parc["C23"].value,
        "Onshore": ws_parc["C22"].value
    }

    # ==============================
    # --- CAPEX / OPEX / Durée de vie ---
    # ==============================

    ws_inv = wb["Investissements"]

    CAPEX = [ws_inv[f"B{i}"].value for i in range(2, 12)]
    OPEX = [ws_inv[f"C{i}"].value for i in range(2, 12)]
    Duree_vie = [ws_inv[f"D{i}"].value for i in range(2, 12)]

    centrales = {}

    for i, type_c in enumerate(types_centrales):
        centrales[type_c] = {
            "opex": OPEX[i] * 1000,
            "duree_vie": Duree_vie[i],
            "capex": CAPEX[i] * 1000
        }

    # ==============================
    # --- Gisements ---
    # ==============================

    ws_gis = wb["Gisements"]

    centrales["gisement"] = {
        "Solar": ws_gis["B3"].value,
        "Onshore": ws_gis["B4"].value,
        "Offshore": ws_gis["B5"].value
    }

    # ==============================
    # --- Rendements ---
    # ==============================

    ws_rend = wb["Rendements"]

    rendements = {
        "electrolyse": ws_rend["B8"].value,
        "CCG_H2": ws_rend["B2"].value,
        "TAC_H2": ws_rend["B3"].value
    }

    # ==============================
    # --- H2 ---
    # ==============================

    H2 = {
        "CCG": {
            "opex": OPEX[6] * 1000,
            "duree_vie": Duree_vie[6],
            "capex": CAPEX[6] * 1000,
            "PU_cost": ws_parc["H9"].value,
            "Pmin": ws_parc["F9"].value,
            "Pmax": ws_parc["E9"].value,
            "dmin": ws_parc["G9"].value,
            "gisement": ws_gis["B12"].value
        },
        "TAC": {
            "opex": OPEX[7] * 1000 ,
            "duree_vie": Duree_vie[7],
            "capex": CAPEX[7] * 1000,
            "PU_cost": ws_parc["H10"].value,
            "Pmin": ws_parc["F10"].value,
            "Pmax": ws_parc["E10"].value,
            "dmin": ws_parc["G10"].value,
            "gisement": ws_gis["B13"].value
        }
    }

    # ==============================
    # --- Défaillance ---
    # ==============================

    ws_def = wb["Defaillance"]

    defaillance = {
        "cost_unsupplied": ws_def["B2"].value,
        "cost_excess": ws_def["B3"].value
    }

    # ==============================
    # --- Hydro ---
    # ==============================

    hydro = {
        "lacs": {
            "Pmax": ws_parc["C20"].value
        },
        "STEP": {
            "Pmax": ws_parc["C21"].value,
            "rendement_pompage": ws_rend["B10"].value
        }
    }

    # ==============================
    # --- Battery ---
    # ==============================

    rbattery = ws_rend["B9"].value
    d_battery = ws_inv["E11"].value
    CapaBattery_init = 0

    battery = {
        "opex": centrales["batterie"]["opex"],
        "duree_vie": centrales["batterie"]["duree_vie"],
        "capex": centrales["batterie"]["capex"],
        "rendement": rbattery,
        "d_battery": d_battery,
        "CapaBattery_init": CapaBattery_init,
        "gisement": ws_gis["B8"].value
    }

    # ==============================
    # --- Electrolyzer ---
    # ==============================

    electrolyzer = {
        "opex": centrales["electrolyseur"]["opex"],
        "duree_vie": centrales["electrolyseur"]["duree_vie"],
        "capex": centrales["electrolyseur"]["capex"]
    }

    return {
        "capacites_init": capacites_init,
        "enr": centrales,
        "H2": H2,
        "hydro": hydro,
        "battery": battery,
        "defaillance": defaillance,
        "rendements": rendements,
        "electrolyzer": electrolyzer
    }