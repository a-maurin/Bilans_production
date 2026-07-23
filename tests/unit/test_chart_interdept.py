# Copyright (C) 2026 Aguirre MAURIN
#
# Ce programme est un logiciel libre : vous pouvez le redistribuer et/ou le modifier
# selon les termes de la Licence Publique Générale GNU (GPL) telle que publiée par
# la Free Software Foundation, version 3 de la licence, ou (à votre choix) toute version ultérieure.

import pytest
from pathlib import Path
from core.common.rendus_graphiques import chart_interdept_stacked_bar

def test_chart_interdept_stacked_bar(tmp_path: Path):
    depts = ["21 - Côte-d'Or", "58 - Nièvre", "71 - Saône-et-Loire"]
    categories = ["PEJ", "PA", "PVe"]
    data_by_category = {
        "PEJ": [10, 4, 15],
        "PA": [2, 1, 5],
        "PVe": [8, 3, 12]
    }
    
    out = chart_interdept_stacked_bar(
        depts,
        categories,
        data_by_category,
        tmp_path,
        "test_interdept_bar.png"
    )
    
    assert Path(out).exists()
    assert Path(out).stat().st_size > 0
