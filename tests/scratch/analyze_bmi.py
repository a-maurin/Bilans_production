# Copyright (C) 2026 Aguirre MAURIN
#
# Ce programme est un logiciel libre : vous pouvez le redistribuer et/ou le modifier
# selon les termes de la Licence Publique Générale GNU (GPL) telle que publiée par
# la Free Software Foundation, version 3 de la licence, ou (à votre choix) toute version ultérieure.
#
# Ce programme est distribué dans l'espoir qu'il sera utile, mais SANS AUCUNE GARANTIE ;
# sans même la garantie implicite de QUALITÉ MARCHANDE ou D'ADÉQUATION À UN USAGE PARTICULIER.
# Voir la Licence Publique Générale GNU pour plus de détails.
#
# CONDITIONS SUPPLÉMENTAIRES D'ATTRIBUTION (SECTION 7(b) DE LA GPL v3) :
# Conformément à la section 7(b) de la GNU GPL v3, vous devez expressément conserver
# intactes et lisibles toutes les mentions d'auteur, notices de copyright et la présente
# clause dans chaque fichier source ou interface utilisateur redistribué. Toute version modifiée
# doit clairement indiquer qu'elle a été altérée et ne doit en aucun cas supprimer le nom
# de l'auteur original (Aguirre MAURIN).

#
import pandas as pd
from pathlib import Path
import re

pve_file = Path("data/sources/Stats_PVe_OFB au 02.06.2026.xlsx")
if not pve_file.exists():
    print(f"Fichier introuvable: {pve_file.resolve()}")
    exit(1)

print("Chargement du fichier Excel (cela peut prendre quelques secondes)...")
df = pd.read_excel(pve_file, dtype=str, engine="openpyxl")

if "nom_site" not in df.columns:
    print("La colonne 'nom_site' est introuvable.")
    exit(1)

# Extraire les valeurs uniques de nom_site qui contiennent 'BMI'
sites = df["nom_site"].dropna().astype(str)
bmi_sites = sites[sites.str.upper().str.contains("BMI")].unique()

# Catégorisation
categories = {
    "BMI Sud-Ouest (SO)": ["SUD", "OUEST", "SO"],
    "BMI Sud-Est (SE)": ["SUD", "EST", "SE"],
    "BMI Nord-Ouest (NO)": ["NORD", "OUEST", "NO"],
    "BMI Nord-Est-Centre (NEC)": ["NORD", "EST", "CENTRE", "NEC"],
    "BMI Île-de-France (IFO/IFE)": ["ILE", "FRANCE", "IFO", "IFE", "IDF"],
    "Autres / Non catégorisés": []
}

results = {k: [] for k in categories}

for site in bmi_sites:
    site_upper = site.upper()
    categorized = False
    
    # Conditions strictes pour éviter les faux positifs
    if "SUD" in site_upper and "OUEST" in site_upper:
        results["BMI Sud-Ouest (SO)"].append(site)
        categorized = True
    elif "SUD" in site_upper and "EST" in site_upper:
        results["BMI Sud-Est (SE)"].append(site)
        categorized = True
    elif "NORD" in site_upper and "OUEST" in site_upper:
        results["BMI Nord-Ouest (NO)"].append(site)
        categorized = True
    elif "NORD" in site_upper and "EST" in site_upper and "CENTRE" in site_upper:
        results["BMI Nord-Est-Centre (NEC)"].append(site)
        categorized = True
    elif any(k in site_upper for k in ["ILE", "IDF", "IFO", "IFE", "FRANCE"]):
        results["BMI Île-de-France (IFO/IFE)"].append(site)
        categorized = True
    
    if not categorized:
        results["Autres / Non catégorisés"].append(site)

print("\n=== ANALYSE DES LIBELLÉS BMI DANS LA SOURCE PVe ===\n")
for cat, sites_list in results.items():
    if sites_list:
        print(f"[{cat}]")
        for s in sorted(sites_list):
            print(f"  - {s}")
        print("")