# Simulateur d'Optimisation du Mix Énergétique (Julia)

Ce projet est un outil de simulation et d'optimisation (Dispatch et Planification) pour un système électrique multi-énergies. Il utilise la programmation linéaire (MILP) via **JuMP** pour minimiser les coûts totaux du système tout en garantissant l'équilibre offre-demande.

## 🚀 Installation

1.  **Installer Julia** (version 1.6 ou supérieure recommandée).
2.  **Cloner le dépôt** et s'assurer que la structure des dossiers existe :
    ```bash
    mkdir -p results
    ```
3.  **Installer les dépendances** :
    Lancez Julia dans le dossier du projet et exécutez :
    ```julia
    using Pkg
    Pkg.add(["JuMP", "Dates", "Random", "JSON3", "XLSX", "CSV"])
    # Choisissez votre solveur (HiGHS est gratuit et performant)
    Pkg.add("HiGHS")
    ```

## 📁 Structure du Projet

* `eod_annuel_master.jl` : Script principal de simulation (boucle hebdomadaire).
* `utils/extraction_donnees_excel.jl` : Fonctions de lecture des paramètres.
* `data/` : Contient le fichier Excel d'hypothèses (`Donnees_etude_de_cas_ETE305.xlsx`).
* `results/` : Dossier de sortie (créé automatiquement) contenant les simulations par ID unique.

## ⚙️ Configuration

Toute la configuration se fait en début de script `main.jl` :
* `solver` : "HiGHS", "Gurobi", "SCIP", ou "CBC".
* `GISEMENTS` : `true`/`false` pour activer les limites de potentiel maximum.
* `H2_ANNUAL_STOCK` : Si `true`, active la gestion long-terme du réservoir d'hydrogène.

## 📊 Sorties de simulation

Chaque exécution génère un dossier dans `results/` nommé par un ID hexadécimal (ex: `8F2A1B34`) contenant :

| Fichier | Description |
| :--- | :--- |
| `results.csv` | Chroniques horaires (production, stocks, défaillance, charge). |
| `parc_annuel.json` | État final du parc et bilan énergétique annuel. |
| `evolution_parc.json` | Historique de la construction des capacités semaine par semaine. |

## 🛠 Modèle Mathématique

Le script optimise le mix en tenant compte de :
* **Contraintes techniques** : Puissance Min/Max, temps de fonctionnement minimal (`dmin`) pour le thermique H2.
* **Stockage** : Rendements de cycle pour les Batteries, les STEP et l'électrolyse.
* **Équilibre** : `Production + Déstockage = Charge + Stockage + Excès - Défaillance`.

---
**Note :** Pour utiliser Gurobi, une licence valide doit être installée sur votre machine.