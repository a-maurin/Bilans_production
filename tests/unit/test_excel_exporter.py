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

# Ce programme est un logiciel libre : vous pouvez le redistribuer et/ou le modifier
# selon les termes de la Licence Publique Générale GNU (GPL) telle que publiée par
# la Free Software Foundation, version 3 de la licence, ou (à votre choix) toute version ultérieure.

import pytest
import pandas as pd
from pathlib import Path
from core.common.excel_exporter import export_synthese_region_excel, main, load_input_file

def test_export_synthese_region_excel(tmp_path: Path):
    df_detail = pd.DataFrame({
        "departement": ["21", "21", "71"],
        "domaine": ["Eau", "Faune", "Eau"],
        "theme": ["Pollution", "Chasse", "Pollution"],
        "nb_operations": [10, 5, 8],
        "nb_localisations": [12, 6, 9],
        "nb_pej": [2, 1, 3],
        "nb_pa": [1, 0, 0],
        "nb_pve": [3, 2, 1]
    })
    
    depts = ["21", "71", "58"]
    res = export_synthese_region_excel(tmp_path, df_detail, depts)
    
    assert res is not None
    assert res.exists()
    assert res.name == "Synthese_Region.xlsx"


def test_excel_exporter_cli_main(tmp_path: Path):
    csv_file = tmp_path / "sample.csv"
    csv_file.write_text("departement;domaine;theme;nb_operations\n21;Eau;Pollution;10\n71;Faune;Chasse;5\n", encoding="utf-8")
    
    out_dir = tmp_path / "output"
    
    # Test CLI échelle région (déduction automatique des depts)
    ret = main(["-i", str(csv_file), "-o", str(out_dir), "-f", "Bilan_CLI.xlsx", "-e", "region"])
    assert ret == 0
    assert (out_dir / "Bilan_CLI.xlsx").exists()

    # Test CLI échelle département avec --code 21
    ret_dept = main(["-i", str(csv_file), "-o", str(out_dir), "-f", "Bilan_Dept21.xlsx", "-e", "departement", "-c", "21"])
    assert ret_dept == 0
    assert (out_dir / "Bilan_Dept21.xlsx").exists()

