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
from core.engine.agregations_region import analyse_region_par_departement

def test_analyse_region_par_departement(tmp_path, monkeypatch):
    import core.engine.agregations_region as mod
    
    # Mock config to test with Region BFC (27)
    monkeypatch.setattr(mod, "get_departements_pour_perimetre", lambda e, c: ["21", "25"] if e == "region" and c == "27" else [])
    
    point = pd.DataFrame([
        {"num_depart": "21", "domaine": "Eau", "theme": "Peche", "fc_id": "A"},
        {"num_depart": "21", "domaine": "Eau", "theme": "Peche", "fc_id": "A"},
        {"num_depart": "25", "domaine": "Nature", "theme": "Foret", "fc_id": "B"}
    ])
    
    pej = pd.DataFrame([
        {"ENTITE_ORIGINE_PROCEDURE": "SD21", "DOMAINE": "Eau", "THEME": "Peche"}
    ])
    
    pa = pd.DataFrame()
    pve = pd.DataFrame([
        {"INF-INSEE": "21000", "DOMAINE": "Nature", "THEME": "Chasse"}
    ])
    
    out_dir = tmp_path
    
    analyse_region_par_departement(point, pa, pej, pve, "region", "27", out_dir)
    
    out_file = out_dir / "region_detail_par_dept.csv"
    assert out_file.exists()
    
    df = pd.read_csv(out_file, sep=";", encoding="utf-8")
    df["departement"] = df["departement"].astype(str)
    
    assert "departement" in df.columns
    assert "nb_localisations" in df.columns
    
    # 21 Eau Peche has 2 localisations, 1 operation, 1 pej, 0 pve
    row21 = df[(df["departement"] == "21") & (df["domaine"] == "Eau") & (df["theme"] == "Peche")]
    assert not row21.empty
    assert row21["nb_localisations"].iloc[0] == 2
    assert row21["nb_operations"].iloc[0] == 1
    assert row21["nb_pej"].iloc[0] == 1
    
    # 21 Nature Chasse has 1 pve
    row21pv = df[(df["departement"] == "21") & (df["domaine"] == "Nature")]
    assert not row21pv.empty
    assert row21pv["nb_pve"].iloc[0] == 1