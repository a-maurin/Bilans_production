# Copyright (C) 2026 Aguirre MAURIN
#
# Ce programme est un logiciel libre : vous pouvez le redistribuer et/ou le modifier
# selon les termes de la Licence Publique Générale GNU (GPL) telle que publiée par
# la Free Software Foundation, version 3 de la licence, ou (à votre choix) toute version ultérieure.

import pandas as pd
from reportlab.platypus import Table
from core.common.pdf_report_builder import PDFReportBuilder
from core.engine.generation_pdf_synthese_brochure import (
    _build_treemap_placeholder_banner,
    _build_matrice_themes_table,
)


def test_build_treemap_placeholder_banner(tmp_path):
    pdf_file = tmp_path / "test.pdf"
    builder = PDFReportBuilder(pdf_path=pdf_file, header_title="Test Header", title="Test Treemap Banner")
    banner = _build_treemap_placeholder_banner(builder, 200.0)
    assert banner is not None
    assert banner.box_width == 200.0


def test_build_matrice_themes_table_empty():
    tbl = _build_matrice_themes_table(None, 200.0)
    assert isinstance(tbl, Table)


def test_build_matrice_themes_table_with_data():
    df = pd.DataFrame([
        {"theme": "Police de la chasse", "nb_pa": 0, "nb_pej": 8, "nb_pve": 18, "nb_total": 26},
        {"theme": "Qualité de l'eau", "nb_pa": 8, "nb_pej": 21, "nb_pve": 0, "nb_total": 29},
        {"theme": "Faune sauvage captive", "nb_pa": 0, "nb_pej": 0, "nb_pve": 0, "nb_total": 0},
    ])
    tbl = _build_matrice_themes_table(df, 200.0)
    assert isinstance(tbl, Table)
