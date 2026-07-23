# Copyright (C) 2026 Aguirre MAURIN
#
# Ce programme est un logiciel libre : vous pouvez le redistribuer et/ou le modifier
# selon les termes de la Licence Publique Générale GNU (GPL) telle que publiée par
# la Free Software Foundation, version 3 de la licence, ou (à votre choix) toute version ultérieure.

import pytest
import pandas as pd
from pathlib import Path
from core.common.excel_exporter import export_synthese_region_excel

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
