# Copyright (C) 2026 Aguirre MAURIN
#
# Ce programme est un logiciel libre : vous pouvez le redistribuer et/ou le modifier
# selon les termes de la Licence Publique Générale GNU (GPL) telle que publiée par
# la Free Software Foundation, version 3 de la licence, ou (à votre choix) toute version ultérieure.
#
# Ce programme est distribué dans l'espoir qu'il sera utile, mais SANS AUCUNE GARANTIE ;
# sans même la garantie implicite de QUALITÉ MARCHANDE ou D'ADÉQUATION À UN USAGE PARTICULIER.
# Voir la Licence Publique Générale GNU pour plus de détails.

"""
========================================================================================
MODULE : MOTEUR DE RENDU DECLARATIF PAR WIDGETS (`moteur_layout_widgets.py`)
========================================================================================
Ce module interprète les grilles de mise en page définies dans les gabarits YAML
(`layout.pages` -> `rows` -> `columns` -> `widget`) et les transforme en éléments ReportLab
(Flowables, Tableaux invisibles) pour la génération des bilans et brochures PDF.

Widgets supportés :
  1. `map` : Cartographie avec options de masquage.
  2. `section_group` : Groupe de chapitres/sections.
  3. `stat_kpi_grid` : Grille de chiffres clés / synthèses.
  4. `theme_breakdown_table` : Tableau de ventilation par usager / thématique.
  5. `evolution_chart` : Graphique d'évolution ou camembert.
  6. `custom_text_box` : Encart textuel / méthodologie.
========================================================================================
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from typing import Any, Callable

from reportlab.platypus import Flowable, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors as rl_colors

logger = logging.getLogger(__name__)


@dataclass
class WidgetContext:
    """Contexte de données et de style transmis aux widgets lors du rendu."""
    gabarit_data: dict[str, Any] = field(default_factory=dict)
    profil_data: dict[str, Any] = field(default_factory=dict)
    options: dict[str, Any] = field(default_factory=dict)
    render_handlers: dict[str, Callable[[dict[str, Any], WidgetContext, float], Flowable | list[Flowable]]] = field(default_factory=dict)


def _parse_column_width(width_spec: str | float | int, usable_width: float) -> float:
    """Calcule la largeur en points PDF à partir de la spécification (ex: '50%' ou points bruts)."""
    if isinstance(width_spec, (int, float)):
        return float(width_spec)
    w_str = str(width_spec).strip()
    if w_str.endswith("%"):
        try:
            pct = float(w_str.rstrip("%"))
            return (pct / 100.0) * usable_width
        except ValueError:
            pass
    try:
        return float(w_str)
    except ValueError:
        return usable_width


def render_widget_fallback(
    widget_config: dict[str, Any],
    ctx: WidgetContext,
    target_width: float,
) -> Flowable:
    """Générateur de secours lorsqu'aucun handler spécifique n'est fourni pour un widget."""
    w_type = widget_config.get("type", "inconnu")
    txt = f"<b>[Widget: {w_type}]</b>"
    return Paragraph(txt)


def compile_row_flowables(
    row_config: dict[str, Any],
    usable_width: float,
    ctx: WidgetContext,
) -> Flowable | list[Flowable]:
    """Compile une rangée du gabarit (`rows`) sous forme d'un tableau ReportLab invisible multi-colonnes."""
    columns = row_config.get("columns", [])
    if not columns:
        return Spacer(1, 1)

    if len(columns) == 1:
        col = columns[0]
        w_width = _parse_column_width(col.get("width", "100%"), usable_width)
        w_config = col.get("widget", {})
        w_type = w_config.get("type", "")
        handler = ctx.render_handlers.get(w_type, render_widget_fallback)
        return handler(w_config, ctx, w_width)

    col_widths: list[float] = []
    col_flowables: list[Any] = []

    for col in columns:
        c_width = _parse_column_width(col.get("width", "100%"), usable_width)
        col_widths.append(c_width)
        w_config = col.get("widget", {})
        w_type = w_config.get("type", "")
        handler = ctx.render_handlers.get(w_type, render_widget_fallback)
        rendered = handler(w_config, ctx, c_width)
        col_flowables.append(rendered)

    table_data = [col_flowables]
    tbl = Table(table_data, colWidths=col_widths)
    tbl.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ]))
    return tbl


def compile_page_layout(
    page_config: dict[str, Any],
    usable_width: float,
    usable_height: float,
    ctx: WidgetContext,
) -> list[Flowable]:
    """Compile l'ensemble des rangées d'une page du gabarit en une liste de Flowables ReportLab."""
    story: list[Flowable] = []
    rows = page_config.get("rows", [])
    total_estimated_height = 0.0

    for r_idx, row in enumerate(rows, start=1):
        flowable_or_list = compile_row_flowables(row, usable_width, ctx)
        if isinstance(flowable_or_list, list):
            story.extend(flowable_or_list)
        else:
            story.append(flowable_or_list)
        story.append(Spacer(1, 4))

    if usable_height > 0 and total_estimated_height > usable_height:
        logger.warning(
            f"Page {page_config.get('page_number', '?')} : la hauteur estimée des widgets ({total_estimated_height:.1f}pt) "
            f"dépasse la hauteur utile de la page ({usable_height:.1f}pt)."
        )

    return story
