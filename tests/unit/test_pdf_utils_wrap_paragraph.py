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
from core.common.pdf_utils import wrap_plain_text_for_pdf_paragraph


def test_wrap_plain_text_for_pdf_paragraph_inserts_breaks() -> None:
    s = "un deux trois quatre cinq six sept huit neuf dix onze douze"
    out = wrap_plain_text_for_pdf_paragraph(s, wrap_width=12, max_lines=20)
    assert "<br/>" in out
    assert "un" in out


def test_wrap_plain_text_for_pdf_paragraph_escapes_html() -> None:
    out = wrap_plain_text_for_pdf_paragraph("A < B et C > D", wrap_width=20, max_lines=10)
    assert "&lt;" in out
    assert "&gt;" in out


def test_wrap_plain_text_for_pdf_paragraph_caps_lines() -> None:
    long = "mot " * 80
    out = wrap_plain_text_for_pdf_paragraph(long, wrap_width=10, max_lines=5)
    assert out.endswith("…")
    assert out.count("<br/>") <= 4