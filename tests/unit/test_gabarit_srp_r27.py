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


def test_generate_synthese_brochure_pdf_report_gabarit_srp_r27(tmp_path):
    from core.engine.generation_pdf_synthese_brochure import generate_synthese_brochure_pdf_report
    generate_synthese_brochure_pdf_report(
        tmp_path,
        gabarit="srp_r27",
        cartes=False,
    )
    generated_pdf = tmp_path / "synthese_activite_PA_PJ_brochure_ext.pdf"
    assert generated_pdf.exists()
    assert generated_pdf.stat().st_size > 0


def test_generate_profile_pdf_report_brochure_routing(tmp_path):
    from core.engine.generation_pdf_profil import generate_profile_pdf_report
    profile = {"id": "global", "presentation_scope": "global"}
    date_deb = pd.Timestamp("2025-01-01")
    date_fin = pd.Timestamp("2025-12-31")
    generate_profile_pdf_report(
        tmp_path,
        profile=profile,
        date_deb=date_deb,
        date_fin=date_fin,
        echelle="departement",
        code="21",
        cartes=False,
        cli_options={"gabarit": "srp_r27"},
    )
    generated_pdf = tmp_path / "global_brochure_int.pdf"
    assert generated_pdf.exists()
    assert generated_pdf.stat().st_size > 0


def test_brochure_resultat_pastilles():
    from core.engine.generation_pdf_synthese_brochure import BrochureResultatPastilles
    widget = BrochureResultatPastilles(200.0, 549, 90, 3, 7)
    assert widget.height == 85.0


def test_brochure_badges_suites():
    from core.engine.generation_pdf_synthese_brochure import BrochureBadgesSuites
    widget = BrochureBadgesSuites(200.0, 11, 37, 44)
    assert widget.height == 100.0


def test_load_annuaire_contact():
    from core.engine.generation_pdf_synthese_brochure import _load_annuaire_contact
    line1, line2 = _load_annuaire_contact("region", "27")
    assert "Office français de la biodiversité" in line1
    assert "Bourgogne-Franche-Comté" in line1
    assert "Dijon" in line2


def test_build_matrice_themes_table_srp_other_row():
    from core.engine.generation_pdf_synthese_brochure import _build_matrice_themes_table_srp
    data = [{"theme": f"Thème {i}", "nb_pa": 1, "nb_pej": 2, "nb_pve": 3, "nb_total": 6} for i in range(15)]
    df = pd.DataFrame(data)
    tbl = _build_matrice_themes_table_srp(df, 200.0, max_top_rows=10)
    assert isinstance(tbl, Table)



