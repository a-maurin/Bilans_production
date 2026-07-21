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
"""Tests configuration d'affichage des graphiques PDF (ratios, garde-fous)."""

from __future__ import annotations

from core.chemins_projet import PROJECT_ROOT
from core.common.chart_display_config import (
    DEFAULT_CHART_DISPLAY_CONFIG,
    clamp_uniform_pie_ratio,
    compute_pdf_ratios,
    load_chart_display_config,
    resolve_reference_pie_display,
)


def test_compute_pdf_ratios_sec4_activite_defaults() -> None:
    ratios = compute_pdf_ratios(DEFAULT_CHART_DISPLAY_CONFIG)
    assert ratios["thematique_sec4_activite_pie_width_ratio_mult"] == 1.0
    assert ratios["thematique_sec4_activite_pie_figure_scale_mult"] == 1.0


def test_compute_pdf_ratios_sec4_activite_clamped() -> None:
    cfg = {
        "pdf": {
            "thematique_sec4_activite_pie_width_ratio_mult": 99.0,
            "thematique_sec4_activite_pie_figure_scale_mult": -1.0,
        }
    }
    ratios = compute_pdf_ratios(cfg)
    assert ratios["thematique_sec4_activite_pie_width_ratio_mult"] == 1.5
    assert ratios["thematique_sec4_activite_pie_figure_scale_mult"] == 0.5


def test_resolve_reference_pie_display_matches_sec22_agrainage_formula() -> None:
    ratios = compute_pdf_ratios(load_chart_display_config(PROJECT_ROOT))
    pie_base = clamp_uniform_pie_ratio(
        ratios,
        uniform_key="thematique_uniform_pie",
        min_key="thematique_uniform_pie_min_ratio",
        max_key="thematique_uniform_pie_max_ratio",
    )
    disp = resolve_reference_pie_display(ratios, pie_base)
    width_mult = ratios["thematique_sec22_resultats_pie_width_ratio_mult"]
    figure_mult = ratios["thematique_sec22_resultats_pie_figure_scale_mult"]
    assert disp["width_ratio"] == min(0.95, pie_base * width_mult)
    assert disp["figure_scale"] == ratios["figure_scale"] * figure_mult
    assert disp["legend_fontsize"] >= 9.0