"""Graphiques matplotlib pour les bilans (camemberts, barres, cartes)."""
from pathlib import Path

import geopandas as gpd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

from scripts.common.ofb_charte import (
    COLOR_PRIMARY,
    CHART_PIE_COLORS,
    CHART_BAR_GROUPED_COLORS,
    COLOR_CHART_4,
)


def _pick_mpl_font() -> str:
    """Pick the best matplotlib-compatible font from the OFB chain."""
    from matplotlib.font_manager import fontManager
    available = {f.name for f in fontManager.ttflist}
    for name in ("Marianne", "Arial", "Liberation Sans", "Helvetica", "DejaVu Sans"):
        if name in available:
            return name
    return "sans-serif"


def apply_mpl_style() -> None:
    """Style matplotlib pour les graphiques exportés en PNG."""
    plt.rcParams["font.family"] = _pick_mpl_font()
    plt.rcParams["axes.titlesize"] = 12
    plt.rcParams["axes.labelsize"] = 10
    plt.rcParams["figure.facecolor"] = "white"


def save_chart(fig, tmp_dir: Path, name: str) -> str:
    path = str(tmp_dir / name)
    fig.savefig(path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return path


def chart_pie(data: dict, title: str, tmp_dir: Path, name: str) -> str:
    apply_mpl_style()
    labels = list(data.keys())
    values = list(data.values())
    # Légende volumineuse : on élargit légèrement la figure pour ne pas tasser
    # exagérément le camembert.
    fig, ax = plt.subplots(figsize=(6, 4))
    # Palette : on force une couleur rouge douce (COLOR_CHART_4) pour les parts
    # correspondant aux infractions / non-conformes, les autres utilisent la
    # palette standard CHART_PIE_COLORS.
    colors_pie = []
    keywords_inf = ("infraction", "infractions", "non conforme", "non conformes", "non-conforme", "non-conformes")
    for i, lb in enumerate(labels):
        base = CHART_PIE_COLORS[i % len(CHART_PIE_COLORS)]
        lbl = str(lb).lower()
        if any(k in lbl for k in keywords_inf):
            colors_pie.append(COLOR_CHART_4)
        else:
            colors_pie.append(base)
    total = sum(values)
    legend_labels = [
        f"{lb} : {v} ({v / total:.1%})" if total else f"{lb} : {v} (0.0%)"
        for lb, v in zip(labels, values)
    ]
    wedges, _ = ax.pie(
        values, startangle=90, colors=colors_pie
    )
    ax.set_aspect("equal")
    ax.legend(
        wedges, legend_labels, loc="center left", bbox_to_anchor=(1.0, 0.5),
        fontsize=9, frameon=False,
    )
    ax.set_title(title, fontsize=11, fontweight="bold", color=COLOR_PRIMARY, pad=10)
    fig.subplots_adjust(left=0.05, right=0.58, top=0.90, bottom=0.05)
    return save_chart(fig, tmp_dir, name)


def chart_bar(
    categories: list,
    values: list,
    title: str,
    ylabel: str,
    tmp_dir: Path,
    name: str,
    color=COLOR_PRIMARY,
) -> str:
    apply_mpl_style()
    fig, ax = plt.subplots(figsize=(5, 3))
    x = np.arange(len(categories))
    bars = ax.bar(x, values, color=color, width=0.5)
    ax.set_xticks(x)
    ax.set_xticklabels(categories, fontsize=9)
    ax.set_ylabel(ylabel, fontsize=9)
    ax.set_title(title, fontsize=11, fontweight="bold", color=COLOR_PRIMARY, pad=10)
    ax.bar_label(bars, fmt="%g", fontsize=9, fontweight="bold")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    fig.tight_layout()
    return save_chart(fig, tmp_dir, name)


def chart_bar_grouped(
    group_labels,
    series: dict,
    title: str,
    ylabel: str,
    tmp_dir: Path,
    name: str,
) -> str:
    apply_mpl_style()
    fig, ax = plt.subplots(figsize=(6, 3.5))
    x = np.arange(len(group_labels))
    n = len(series)
    w = 0.30
    keywords_inf = ("infraction", "infractions", "non conforme", "non conformes", "non-conforme", "non-conformes")
    for i, (label, vals) in enumerate(series.items()):
        offset = (i - n / 2 + 0.5) * w
        lbl = str(label).lower()
        if any(k in lbl for k in keywords_inf):
            color = COLOR_CHART_4
        else:
            color = CHART_BAR_GROUPED_COLORS[i % len(CHART_BAR_GROUPED_COLORS)]
        bars = ax.bar(
            x + offset, vals, w, label=label,
            color=color,
        )
        ax.bar_label(bars, fmt="%g", fontsize=8, fontweight="bold")
    ax.set_xticks(x)
    ax.set_xticklabels(group_labels, fontsize=9)
    ax.set_ylabel(ylabel, fontsize=9)
    ax.set_title(title, fontsize=11, fontweight="bold", color=COLOR_PRIMARY, pad=10)
    ax.legend(fontsize=8)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    fig.tight_layout()
    return save_chart(fig, tmp_dir, name)


def chart_bar_stacked(
    group_labels: list,
    series: dict,
    title: str,
    ylabel: str,
    tmp_dir: Path,
    name: str,
) -> str:
    """Barres empilées : chaque série est empilée sur la précédente."""
    apply_mpl_style()
    fig, ax = plt.subplots(figsize=(max(5, len(group_labels) * 1.1), 4))
    x = np.arange(len(group_labels))
    bottom = np.zeros(len(group_labels))
    keywords_inf = ("infraction", "infractions", "non conforme", "non conformes", "non-conforme", "non-conformes")
    for i, (label, vals) in enumerate(series.items()):
        vals_arr = np.array(vals, dtype=float)
        lbl = str(label).lower()
        if any(k in lbl for k in keywords_inf):
            color = COLOR_CHART_4
        else:
            color = CHART_BAR_GROUPED_COLORS[i % len(CHART_BAR_GROUPED_COLORS)]
        bars = ax.bar(x, vals_arr, 0.55, label=label, bottom=bottom, color=color)
        for bar, val in zip(bars, vals_arr):
            if val > 0:
                ax.text(
                    bar.get_x() + bar.get_width() / 2,
                    bar.get_y() + bar.get_height() / 2,
                    f"{int(val)}",
                    ha="center", va="center", fontsize=7, fontweight="bold",
                    color="white",
                )
        bottom += vals_arr
    ax.set_xticks(x)
    ax.set_xticklabels(group_labels, fontsize=9)
    ax.set_ylabel(ylabel, fontsize=9)
    ax.set_title(title, fontsize=11, fontweight="bold", color=COLOR_PRIMARY, pad=10)
    ax.legend(fontsize=8, loc="upper left")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    fig.tight_layout()
    return save_chart(fig, tmp_dir, name)


def chart_line_evolution(
    x_labels: list,
    series: dict,
    title: str,
    ylabel: str,
    tmp_dir: Path,
    name: str,
) -> str:
    """Courbe d'évolution multi-séries (un trait par indicateur)."""
    apply_mpl_style()
    fig, ax = plt.subplots(figsize=(max(5, len(x_labels) * 1.1), 3.5))
    x = np.arange(len(x_labels))
    keywords_inf = ("infraction", "infractions", "non conforme", "non conformes", "non-conforme", "non-conformes")
    for i, (label, vals) in enumerate(series.items()):
        lbl = str(label).lower()
        if any(k in lbl for k in keywords_inf):
            color = COLOR_CHART_4
        else:
            color = CHART_BAR_GROUPED_COLORS[i % len(CHART_BAR_GROUPED_COLORS)]
        ax.plot(x, vals, marker="o", label=label, color=color, linewidth=2, markersize=6)
        for xi, v in zip(x, vals):
            ax.annotate(
                str(int(v)), (xi, v),
                textcoords="offset points", xytext=(0, 8),
                ha="center", fontsize=8, fontweight="bold", color=color,
            )
    ax.set_xticks(x)
    ax.set_xticklabels(x_labels, fontsize=9)
    ax.set_ylabel(ylabel, fontsize=9)
    ax.set_title(title, fontsize=11, fontweight="bold", color=COLOR_PRIMARY, pad=10)
    ax.legend(fontsize=8)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    fig.tight_layout()
    return save_chart(fig, tmp_dir, name)


def _make_map(
    communes_gdf: gpd.GeoDataFrame,
    agg_data,
    insee_col: str,
    column: str,
    cmap: str,
    legend_label: str,
    title: str,
    tmp_dir: Path,
    name: str,
    points_gdf=None,
    points_label=None,
) -> str:
    """Génère une carte choroplèthe PNG pour intégration dans le PDF."""
    apply_mpl_style()
    communes_simple = communes_gdf.copy()
    communes_simple["geometry"] = communes_simple.geometry.simplify(
        tolerance=0.0005, preserve_topology=True
    )

    if agg_data is not None and not agg_data.empty:
        merge_col = (
            "insee_comm" if "insee_comm" in agg_data.columns else agg_data.columns[0]
        )
        geo = communes_simple.merge(
            agg_data, left_on=insee_col, right_on=merge_col, how="left"
        )
    else:
        geo = communes_simple.copy()
        geo[column] = 0

    xmin, ymin, xmax, ymax = communes_gdf.total_bounds
    marge = 0.02 * max(xmax - xmin, ymax - ymin)

    fig, ax = plt.subplots(figsize=(7, 7))
    geo[column] = geo[column].fillna(0)
    geo.plot(
        column=column,
        cmap=cmap,
        linewidth=0.3,
        edgecolor="white",
        legend=True,
        ax=ax,
        legend_kwds={"label": legend_label, "shrink": 0.6, "aspect": 25},
        missing_kwds={"color": "lightgrey", "label": "Aucune donnée"},
        rasterized=True,
    )

    if points_gdf is not None and not points_gdf.empty:
        if (
            communes_gdf.crs is not None
            and points_gdf.crs != communes_gdf.crs
        ):
            points_gdf = points_gdf.to_crs(communes_gdf.crs)
        points_gdf.plot(
            ax=ax,
            color="#E76F51",
            markersize=18,
            alpha=0.8,
            label=points_label or "Points",
            edgecolor="white",
            linewidth=0.3,
            rasterized=True,
        )
        ax.legend(fontsize=8, loc="lower left")

    ax.set_xlim(xmin - marge, xmax + marge)
    ax.set_ylim(ymin - marge, ymax + marge)
    ax.set_aspect("equal")
    ax.set_title(title, fontsize=13, fontweight="bold", color=COLOR_PRIMARY, pad=12)
    ax.axis("off")

    scale_len_m = 20000
    if communes_gdf.crs and communes_gdf.crs.is_geographic:
        scale_len_deg = scale_len_m / 111320
    else:
        scale_len_deg = scale_len_m
    sx = xmin + marge
    sy = ymin + marge * 0.5
    ax.plot([sx, sx + scale_len_deg], [sy, sy], color="black", linewidth=2)
    ax.text(
        sx + scale_len_deg / 2,
        sy + marge * 0.3,
        "20 km",
        ha="center",
        fontsize=8,
        fontweight="bold",
    )

    nx = xmax - marge * 0.5
    ny = ymax - marge * 1.5
    ax.annotate(
        "N",
        xy=(nx, ny),
        xytext=(nx, ny - marge * 2),
        arrowprops=dict(arrowstyle="->", lw=1.5, color="black"),
        fontsize=10,
        fontweight="bold",
        ha="center",
        va="bottom",
    )

    fig.tight_layout()
    return save_chart(fig, tmp_dir, name)
