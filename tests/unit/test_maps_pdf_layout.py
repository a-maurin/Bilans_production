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
"""Mise en page des cartes PDF (une carte par page)."""

from __future__ import annotations

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

from core.common.ofb_charte import MARGIN_BOTTOM, MARGIN_LEFT, MARGIN_RIGHT, MARGIN_TOP
from core.common.pdf_report_builder import (
    _MAP_PAGE_WIDTH_FRACTION,
    compute_stacked_maps_width,
)


def test_stacked_maps_width_reduced_when_too_tall() -> None:
    avail_w = A4[0] - MARGIN_LEFT - MARGIN_RIGHT
    frame_h = A4[1] - MARGIN_TOP - MARGIN_BOTTOM
    # Deux cartes portrait (ratio > 1) : à pleine largeur elles dépasseraient la page.
    ratios = [1.3, 1.3]
    map_w = compute_stacked_maps_width(avail_w, frame_h, ratios)
    assert map_w < avail_w * 0.98
    total_h = map_w * sum(ratios) + 2 * mm
    assert total_h <= frame_h - 28 * mm


def test_single_map_width_fits_height_budget_for_portrait_map() -> None:
    avail_w = A4[0] - MARGIN_LEFT - MARGIN_RIGHT
    frame_h = A4[1] - MARGIN_TOP - MARGIN_BOTTOM
    ratio = 1.6
    map_w = compute_stacked_maps_width(avail_w, frame_h, [ratio])
    assert map_w < avail_w * 0.98
    assert map_w * ratio <= frame_h - 28 * mm


def test_single_map_width_equal_when_computed_individually() -> None:
    avail_w = A4[0] - MARGIN_LEFT - MARGIN_RIGHT
    frame_h = A4[1] - MARGIN_TOP - MARGIN_BOTTOM
    budget = frame_h - 28 * mm
    ratio_landscape = 0.65
    ratio_portrait = 1.35

    width_landscape = compute_stacked_maps_width(avail_w, frame_h, [ratio_landscape])
    width_portrait = compute_stacked_maps_width(avail_w, frame_h, [ratio_portrait])

    expected_landscape = min(avail_w * 0.98, budget / ratio_landscape)
    expected_portrait = min(avail_w * 0.98, budget / ratio_portrait)

    assert width_landscape == expected_landscape
    assert width_portrait == expected_portrait


def test_single_map_pdf_width_fraction_strictly_below_frame_width() -> None:
    """Même logique que ``_map_display_width`` : marge horizontale pour éviter la coupe à droite."""
    avail_w = A4[0] - MARGIN_LEFT - MARGIN_RIGHT
    frame_h = A4[1] - MARGIN_TOP - MARGIN_BOTTOM
    ratio = 0.35
    map_w = compute_stacked_maps_width(
        avail_w, frame_h, [ratio], width_fraction=_MAP_PAGE_WIDTH_FRACTION
    )
    assert map_w < avail_w
    assert map_w <= avail_w * _MAP_PAGE_WIDTH_FRACTION + 1e-6