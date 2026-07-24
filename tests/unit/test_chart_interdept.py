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
