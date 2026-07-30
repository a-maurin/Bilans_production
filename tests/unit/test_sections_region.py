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
from core.engine.pdf_context import PdfContext
from core.common.pdf_report_builder import PDFReportBuilder
from core.engine.sections_region import (
    render_sec_region_dashboard,
    render_sec_region_fiches,
    render_sec_region_detail
)

def test_sections_region_pipeline(tmp_path: Path):
    csv_path = tmp_path / "region_detail_par_dept.csv"
    df = pd.DataFrame({
        "departement": ["21", "21", "58"],
        "domaine": ["Eau", "Faune", "Eau"],
        "theme": ["Pollution", "Chasse", "Pollution"],
        "nb_operations": [10, 5, 0],
        "nb_localisations": [12, 6, 0],
        "nb_pej": [2, 1, 0],
        "nb_pa": [1, 0, 0],
        "nb_pve": [3, 2, 0]
    })
    df.to_csv(csv_path, sep=";", index=False, encoding="utf-8")

    pdf_file = tmp_path / "Bilan_Regional_Test.pdf"
    builder = PDFReportBuilder(pdf_file, "Test Régional", title="Test Régional")
    
    ctx = PdfContext(
        builder=builder,
        profile={"id": "global"},
        presentation_cfg={},
        behavior_cfg={},
        show_placeholder=False,
        date_deb=pd.Timestamp("2025-01-01"),
        date_fin=pd.Timestamp("2025-12-31"),
        dept_code="r27",
        dept_name_typo="Région Bourgogne-Franche-Comté",
        diffusion="interne",
        ventilation_mode="annee",
        out_dir=tmp_path,
        avail_w=builder.avail_w,
        tmp_dir=tmp_path,
        chart_bar_w=0.72,
        legend_fontsize=8.0,
        legend_ncol_max=4,
        figure_scale=1.0,
        ref_pie_w=0.34,
        ref_pie_fs=8.0,
        ref_pie_legend_fs=7.0,
        split_by_row=False,
        tables_layout={},
        section_title={
            "sec_region_dashboard": "1. Synthèse régionale",
            "sec_region_fiches": "2. Fiches départementales",
            "secregion": "3. Annexe technique"
        },
        nb_ops=15,
        nb_localisations=18,
        nb_pej=3,
        nb_pa=1,
        nb_pve=5
    )

    render_sec_region_dashboard(ctx)
    render_sec_region_fiches(ctx)
    render_sec_region_detail(ctx)
    builder.build()

    assert pdf_file.exists()
    
    # Vérification du sous-dossier et des fiches individuelles
    dept_dir = tmp_path / "departements"
    assert dept_dir.exists()
    assert (dept_dir / "Fiche_Dept_21.pdf").exists()
    assert (dept_dir / "Fiche_Dept_58.pdf").exists()

def test_dynamic_note_and_vignette_ratio(tmp_path: Path):
    from core.engine.sections_region import _generate_dept_vignette
    from PIL import Image
    vignette = _generate_dept_vignette("21", tmp_path, tmp_path, "vignette_test_21.png")
    assert vignette.exists()
    with Image.open(vignette) as im:
        w, h = im.size
        # Ratio carré 1:1 parfait
        assert w == h
