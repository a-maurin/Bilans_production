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
"""Tests d'intégration TOC PDF — profil global."""
from __future__ import annotations

import shutil
from pathlib import Path

import pandas as pd

from core.common.pdf_presentation_config import (
    apply_diffusion_pdf_suffix,
    resolve_pdf_presentation_config,
    resolve_section_titles,
)
from core.common.pdf_toc_inspection import (
    assert_section_headings_order,
    extract_pdf_section_headings,
)
from pdf_toc_test_support import patch_pdf_charts

FIXTURES = Path(__file__).resolve().parent.parent / "fixtures" / "pdf_toc_global"
PROJECT_ROOT = Path(__file__).resolve().parents[2]


def _seed_global_out_dir(out_dir: Path) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    for name in FIXTURES.glob("*.csv"):
        shutil.copy(name, out_dir / name.name)


def test_global_scope_section_titles_from_yaml() -> None:
    resolved = resolve_pdf_presentation_config(
        PROJECT_ROOT, scope="global", profile_id=None, diffusion="interne"
    )
    effective = resolved["effective"]
    titles = effective.get("sections", {}).get("titles", {})
    assert "3." in str(titles.get("sec4", ""))
    assert "4." in str(titles.get("sec3", ""))

    section_defs = [
        ("sec4", "3. Activité par type d'usager"),
        ("sec3", "4. Procédures (PEJ, PA, PVe)"),
    ]
    resolved_defs = resolve_section_titles(effective, section_defs)
    by_id = dict(resolved_defs)
    assert "3." in by_id["sec4"]
    assert "4." in by_id["sec3"]


def test_global_pdf_section_headings_order(tmp_path: Path, monkeypatch) -> None:
    import core.engine.generation_pdf_profil as global_pdf

    out_dir = tmp_path / "out_global"
    _seed_global_out_dir(out_dir)

    patch_pdf_charts(monkeypatch, global_pdf)

    global_pdf.generate_pdf_report(
        out_dir,
        profile={"id": "global", "presentation_scope": "global"},
        date_deb=pd.Timestamp("2025-01-01"),
        date_fin=pd.Timestamp("2025-12-31"),
        echelle="departement",
        code="21",
        ventilation_mode="globale",
        diffusion="interne",
        cartes=False,
        output_filename="bilan_global_test.pdf",
    )

    pdf_path = apply_diffusion_pdf_suffix(out_dir / "bilan_global_test.pdf", "interne")
    assert pdf_path.is_file(), f"PDF absent : {list(out_dir.glob('*.pdf'))}"

    headings = extract_pdf_section_headings(pdf_path)
    assert_section_headings_order(
        headings,
        [
            "2.2.",
            "3. Activité",
            "4. Procédures",
        ],
    )