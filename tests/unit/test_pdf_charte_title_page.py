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
"""Non-régression charte OFB : page de garde et pages intérieures."""

from __future__ import annotations

from pathlib import Path

from reportlab.lib.units import mm

from core.common.ofb_charte import IMG_FILIGRANE
from core.common.pdf_report_builder import PDFReportBuilder


def test_title_page_builds_with_charte_config(tmp_path: Path) -> None:
    pdf_path = tmp_path / "cover.pdf"
    builder = PDFReportBuilder(
        pdf_path=pdf_path,
        header_title="Bilan test",
        charte_config={
            "assets": {
                "banner": "image5.jpg",
                "title_page_deco": "image6.jpeg",
            },
            "title_page": {
                "banner_height_mm": 40,
                "deco_height_ratio": 0.45,
            },
        },
        diffusion="externe",
    )
    builder.add_title_page(["Bilan test"], "Période : 01/01/2025 au 31/12/2025")
    out = builder.build()
    assert out.exists()
    assert out.stat().st_size > 500


def test_content_page_builds_with_charte_config(tmp_path: Path) -> None:
    pdf_path = tmp_path / "content.pdf"
    builder = PDFReportBuilder(
        pdf_path=pdf_path,
        header_title="Bilan test",
        charte_config={
            "content_page": {
                "watermark_enabled": True,
                "filigrane_height_ratio": 0.45,
                "filigrane_align": "bottom_right",
                "footer_deco_enabled": False,
            },
        },
    )
    builder.add_title_page(["Bilan test"], "Période : 01/01/2025 au 31/12/2025")
    builder.add_paragraph("Paragraphe de contrôle sur page intérieure.")
    out = builder.build()
    assert out.exists()
    assert out.stat().st_size > 800


def test_filigrane_size_bottom_right_clamped_to_page_width() -> None:
    builder = PDFReportBuilder(
        pdf_path=Path("unused3.pdf"),
        header_title="Test",
        charte_config={"content_page": {"filigrane_height_ratio": 0.50}},
    )
    width_pt, height_pt = builder._filigrane_image_size_pt(IMG_FILIGRANE)
    assert width_pt <= builder._page_w + 0.01
    assert height_pt > 0


def test_footer_text_zone_top_unchanged() -> None:
    builder = PDFReportBuilder(pdf_path=Path("unused2.pdf"), header_title="Test")
    y_foot = 8 * mm
    assert abs(builder._footer_text_zone_top_pt() - (y_foot + 12 + 7)) < 0.01


def test_title_page_banner_height_pt_from_charte() -> None:
    builder = PDFReportBuilder(
        pdf_path=Path("unused.pdf"),
        header_title="Test",
        charte_config={"title_page": {"banner_height_mm": 38}},
    )

    assert abs(builder._title_page_banner_height_pt() - 38 * mm) < 0.01