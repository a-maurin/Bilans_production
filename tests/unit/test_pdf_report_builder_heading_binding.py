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
"""Non-régression pagination: titre lié au début de contenu."""

from __future__ import annotations

from pathlib import Path

from PIL import Image
from reportlab.platypus import KeepTogether

from core.common.pdf_report_builder import PDFReportBuilder


def _make_png(path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    Image.new("RGB", (120, 80), color=(230, 240, 250)).save(path)
    return path


def _has_image(flowables: list) -> bool:
    for f in flowables:
        if type(f).__name__ == "Image":
            return True
        if type(f).__name__ == "KeepTogether":
            if _has_image(getattr(f, "_content", [])):
                return True
def _has_binding(story: list) -> bool:
    """Vérifie que la section commence par un KeepTogether ou un CondPageBreak protecteur."""
    for f in story:
        name = type(f).__name__
        if name in ("KeepTogether", "CondPageBreak"):
            return True
    return False


def test_section_heading_stays_bound_to_first_content_item_for_image(tmp_path: Path) -> None:
    pdf_path = tmp_path / "out.pdf"
    img_path = _make_png(tmp_path / "img.png")
    builder = PDFReportBuilder(pdf_path=pdf_path, header_title="Test")

    builder.add_section("sec_test", "Section test")
    builder.add_image(img_path, width_ratio=0.4)

    assert builder.story, "Story vide inattendue."
    assert _has_binding(builder.story), "Le titre doit être lié (KeepTogether ou CondPageBreak)."


def test_subsection_heading_stays_bound_when_table_and_image_can_split(tmp_path: Path) -> None:
    pdf_path = tmp_path / "out2.pdf"
    img_path = _make_png(tmp_path / "img2.png")
    builder = PDFReportBuilder(
        pdf_path=pdf_path,
        header_title="Test",
        tables_layout={"split_by_row": True},
    )

    builder.add_section("sec_sub", "Sous-section test", level=2)
    builder.add_table_and_image_keep_together(
        [["Col A", "Col B"], ["x", "1"], ["y", "2"]],
        table_caption="Table test",
        image_path=img_path,
        image_width_ratio=0.5,
    )

    assert builder.story, "Story vide inattendue."
    assert _has_binding(builder.story), "Le titre doit être lié (KeepTogether ou CondPageBreak)."


def test_local_heading_chart_table_keeps_heading_with_first_content(tmp_path: Path) -> None:
    pdf_path = tmp_path / "out3.pdf"
    img_path = _make_png(tmp_path / "img3.png")
    builder = PDFReportBuilder(pdf_path=pdf_path, header_title="Test")
    builder.add_section("sec_usagers", "Section usagers")
    builder.add_heading_chart_table_keep_together(
        heading_text="Résultats des contrôles par type d'usager",
        heading_style="Heading2",
        chart_path=img_path,
        chart_width_ratio=0.5,
        table_rows=[["Type", "Nb"], ["A", "1"], ["B", "2"]],
        table_caption="Résultats des contrôles par type d'usager",
    )

    assert builder.story, "Story vide inattendue."
    assert _has_binding(builder.story), "Le titre local doit être lié par KeepTogether ou CondPageBreak."
    
    # Vérifie qu'on n'a pas inclus l'image lourde dans le KeepTogether du titre (s'il y en a un)
    for f in builder.story:
        if type(f).__name__ == "KeepTogether":
            content = getattr(f, "_content", [])
            # Si le KeepTogether commence par un Paragraph (le titre), il ne doit pas engloutir l'image
            if content and type(content[0]).__name__ == "Paragraph":
                assert not _has_image(content), "Le KeepTogether lié au titre ne doit pas inclure l'image lourde."


def test_local_heading_without_pending_still_keeps_first_content(tmp_path: Path) -> None:
    pdf_path = tmp_path / "out4.pdf"
    img_path = _make_png(tmp_path / "img4.png")
    builder = PDFReportBuilder(pdf_path=pdf_path, header_title="Test")

    builder.add_heading_chart_table_keep_together(
        heading_text="Résultats des contrôles par type d'usager",
        heading_style="Heading2",
        chart_path=img_path,
        chart_width_ratio=0.5,
        table_rows=[["Type", "Nb"], ["A", "1"]],
    )

    assert builder.story, "Story vide inattendue."
    assert _has_binding(builder.story), "Le titre local doit être lié par KeepTogether ou CondPageBreak."
    
    for f in builder.story:
        if type(f).__name__ == "KeepTogether":
            content = getattr(f, "_content", [])
            if content and type(content[0]).__name__ == "Paragraph":
                assert not _has_image(content), "Le KeepTogether lié au titre ne doit pas inclure l'image lourde."