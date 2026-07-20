#
# Copyright (C) 2026 Aguirre MAURIN
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.

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
