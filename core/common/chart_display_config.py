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
# Conformément à la section 7(b) DE LA GPL v3, vous devez expressément conserver
# intactes et lisibles toutes les mentions d'auteur, notices de copyright et la présente
# clause dans chaque fichier source ou interface utilisateur redistribué. Toute version modifiée
# doit clairement indiquer qu'elle a été altérée et ne doit en aucun cas supprimer le nom
# de l'auteur original (Aguirre MAURIN).

"""
========================================================================================
MODULE : CONFIGURATION ET DIMENSIONNEMENT DES GRAPHIQUES PDF (`chart_display_config.py`)
========================================================================================
Ce module contrôle les proportions, échelles et polices des graphiques insérés dans les rapports.

Points clés :
  1. Configuration par défaut des ratios de largeur (camemberts et histogrammes).
  2. Presets prédéfinis ('compact', 'standard', 'large') ajustant l'ensemble des figures.
  3. Chargement dynamique des surcharges YAML (`charts_config.yaml` et `pdf_presentation.yaml`).
  4. Calcul sécurisé des ratios avec bornes de garde-fous pour éviter tout débordement de page.
========================================================================================
"""
from __future__ import annotations

from pathlib import Path
from typing import Any

from core.common.pdf_report_builder import (
    THEMATIC_CHART_WIDTH_RATIO,
    THEMATIC_PIE_CHART_WIDTH_RATIO,
)
from core.common.pdf_presentation_config import load_pdf_presentation_raw_config

# ========================================================================================
# VALEURS PAR DÉFAUT ET PRESETS DE RENDU DES GRAPHIQUES
# ========================================================================================

DEFAULT_CHART_DISPLAY_CONFIG: dict[str, Any] = {
    "pdf": {
        # Ratios de base pour l'insertion ReportLab
        "pie_width_ratio_base": float(THEMATIC_PIE_CHART_WIDTH_RATIO),
        "chart_width_ratio_base": float(THEMATIC_CHART_WIDTH_RATIO),
        # Ajustements pour la section usagers / activités
        "activite_usagers_controles_pie_scale": 3.0,
        "activite_usagers_resultats_bar_scale": 0.70,
        # Multiplicateurs pour les camemberts globaux
        "global_resultats_pie_scale": 1.0,
        "global_usagers_pie_scale": 1.0,
        "global_domaine_pie_scale": 1.0,
        "global_theme_pie_scale": 1.0,
        # Uniformisation des dimensions de camemberts
        "global_uniform_pie_scale": 1.0,
        "global_uniform_pie_min_ratio": 0.70,
        "global_uniform_pie_max_ratio": 0.82,
        "global_type_usager_bar_scale": 1.25,
        # Uniformisation du moteur thématique
        "thematique_uniform_pie_scale": 1.0,
        "thematique_uniform_pie_min_ratio": 0.70,
        "thematique_uniform_pie_max_ratio": 0.82,
        "thematique_uniform_chart_scale": 1.0,
        # Multiplicateurs Matplotlib (hauteur de figure)
        "thematique_sec21_figure_scale_mult": 1.55,
        "thematique_sec22_resultats_pie_figure_scale_mult": 1.22,
        "thematique_sec22_resultats_pie_width_ratio_mult": 1.12,
        "thematique_sec4_activite_pie_width_ratio_mult": 1.0,
        "thematique_sec4_activite_pie_figure_scale_mult": 1.0,
        # Tailles et légendes Matplotlib
        "figure_scale": 1.0,
        "legend_fontsize": 8.0,
        "legend_ncol_max": 4.0,
    }
}

# Profils préconfigurés pour s'adapter aux différents gabarits de documents
CHART_PRESETS: dict[str, dict[str, float]] = {
    "compact": {
        "pie_width_ratio_base": 0.30,
        "chart_width_ratio_base": 0.62,
        "activite_usagers_controles_pie_scale": 2.0,
        "activite_usagers_resultats_bar_scale": 0.55,
        "global_resultats_pie_scale": 0.90,
        "global_usagers_pie_scale": 0.90,
        "global_domaine_pie_scale": 0.90,
        "global_theme_pie_scale": 0.90,
        "global_uniform_pie_scale": 0.90,
        "global_uniform_pie_min_ratio": 0.60,
        "global_uniform_pie_max_ratio": 0.78,
        "global_type_usager_bar_scale": 1.15,
        "thematique_uniform_pie_scale": 0.90,
        "thematique_uniform_pie_min_ratio": 0.60,
        "thematique_uniform_pie_max_ratio": 0.78,
        "thematique_uniform_chart_scale": 0.90,
        "figure_scale": 0.95,
        "legend_fontsize": 7.0,
        "legend_ncol_max": 3.0,
    },
    "standard": {
        "pie_width_ratio_base": 0.34,
        "chart_width_ratio_base": 0.72,
        "activite_usagers_controles_pie_scale": 3.0,
        "activite_usagers_resultats_bar_scale": 1.20,
        "global_resultats_pie_scale": 1.00,
        "global_usagers_pie_scale": 1.00,
        "global_domaine_pie_scale": 1.00,
        "global_theme_pie_scale": 1.00,
        "global_uniform_pie_scale": 1.00,
        "global_uniform_pie_min_ratio": 0.70,
        "global_uniform_pie_max_ratio": 0.82,
        "global_type_usager_bar_scale": 1.25,
        "thematique_uniform_pie_scale": 1.00,
        "thematique_uniform_pie_min_ratio": 0.70,
        "thematique_uniform_pie_max_ratio": 0.82,
        "thematique_uniform_chart_scale": 1.00,
        "figure_scale": 1.00,
        "legend_fontsize": 8.0,
        "legend_ncol_max": 4.0,
    },
    "large": {
        "pie_width_ratio_base": 0.40,
        "chart_width_ratio_base": 0.85,
        "activite_usagers_controles_pie_scale": 3.0,
        "activite_usagers_resultats_bar_scale": 0.90,
        "global_resultats_pie_scale": 1.10,
        "global_usagers_pie_scale": 1.10,
        "global_domaine_pie_scale": 1.10,
        "global_theme_pie_scale": 1.10,
        "global_uniform_pie_scale": 1.10,
        "global_uniform_pie_min_ratio": 0.74,
        "global_uniform_pie_max_ratio": 0.90,
        "global_type_usager_bar_scale": 1.35,
        "thematique_uniform_pie_scale": 1.10,
        "thematique_uniform_pie_min_ratio": 0.74,
        "thematique_uniform_pie_max_ratio": 0.90,
        "thematique_uniform_chart_scale": 1.10,
        "figure_scale": 1.08,
        "legend_fontsize": 9.0,
        "legend_ncol_max": 4.0,
    },
}


def _clamp_ratio(value: float) -> float:
    """Restreint une valeur de ratio dans l'intervalle valide [0.1, 1.0]."""
    return max(0.1, min(1.0, float(value)))


# ========================================================================================
# CHARGEMENT ET CALCUL DES PROPORTIONS DE GRAPHIQUES
# ========================================================================================

def load_chart_display_config(root: Path, preset: str | None = None) -> dict[str, Any]:
    """Charge la configuration d'affichage depuis les fichiers YAML de configuration."""
    cfg = DEFAULT_CHART_DISPLAY_CONFIG.copy()
    presentation_cfg = load_pdf_presentation_raw_config(root)
    presentation_charts = (presentation_cfg.get("defaults") or {}).get("charte", {}).get("charts", {})
    if presentation_charts and "pie_width_ratio_base" in presentation_charts:
        cfg["pdf"]["pie_width_ratio_base"] = float(presentation_charts["pie_width_ratio_base"])

    candidates = [
        root / "config" / "charts" / "charts_config.yaml",
        root / "ref" / "programme" / "charts_config.yaml",
    ]
    path = next((p for p in candidates if p.exists()), None)
    if path is None:
        return cfg
    try:
        import yaml
    except ImportError:
        return cfg
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}
    except Exception:
        return cfg

    pdf_data = data.get("pdf", {}) if isinstance(data, dict) else {}
    if not isinstance(pdf_data, dict):
        return cfg

    out = dict(cfg["pdf"])
    for key in out.keys():
        if key in pdf_data:
            out[key] = pdf_data[key]
    cfg["pdf"] = out
    if preset:
        preset_key = str(preset).strip().lower()
        if preset_key in CHART_PRESETS:
            cfg["pdf"].update(CHART_PRESETS[preset_key])
    return cfg


def compute_pdf_ratios(cfg: dict[str, Any]) -> dict[str, float]:
    """Calcule tous les ratios réels appliqués aux graphiques PDF avec sécurité de bornage."""
    pdf_cfg = cfg.get("pdf", {})
    pie_base = _clamp_ratio(pdf_cfg.get("pie_width_ratio_base", THEMATIC_PIE_CHART_WIDTH_RATIO))
    chart_base = _clamp_ratio(pdf_cfg.get("chart_width_ratio_base", THEMATIC_CHART_WIDTH_RATIO))

    return {
        "pie_base": pie_base,
        "chart_base": chart_base,
        "activite_usagers_controles_pie": _clamp_ratio(
            pie_base * float(pdf_cfg.get("activite_usagers_controles_pie_scale", 3.0))
        ),
        "activite_usagers_resultats_bar": _clamp_ratio(
            chart_base * float(pdf_cfg.get("activite_usagers_resultats_bar_scale", 0.70))
        ),
        "global_resultats_pie": _clamp_ratio(
            pie_base * float(pdf_cfg.get("global_resultats_pie_scale", 1.0))
        ),
        "global_usagers_pie": _clamp_ratio(
            pie_base * float(pdf_cfg.get("global_usagers_pie_scale", 1.0))
        ),
        "global_domaine_pie": _clamp_ratio(
            pie_base * float(pdf_cfg.get("global_domaine_pie_scale", 1.0))
        ),
        "global_theme_pie": _clamp_ratio(
            pie_base * float(pdf_cfg.get("global_theme_pie_scale", 1.0))
        ),
        "global_uniform_pie": _clamp_ratio(
            pie_base * float(pdf_cfg.get("global_uniform_pie_scale", 1.0))
        ),
        "global_uniform_pie_min_ratio": _clamp_ratio(
            float(pdf_cfg.get("global_uniform_pie_min_ratio", 0.70))
        ),
        "global_uniform_pie_max_ratio": _clamp_ratio(
            float(pdf_cfg.get("global_uniform_pie_max_ratio", 0.82))
        ),
        "global_type_usager_bar_ratio": _clamp_ratio(
            chart_base * float(pdf_cfg.get("global_type_usager_bar_scale", 1.25))
        ),
        "thematique_uniform_pie": _clamp_ratio(
            pie_base * float(pdf_cfg.get("thematique_uniform_pie_scale", 1.0))
        ),
        "thematique_uniform_pie_min_ratio": _clamp_ratio(
            float(pdf_cfg.get("thematique_uniform_pie_min_ratio", 0.70))
        ),
        "thematique_uniform_pie_max_ratio": _clamp_ratio(
            float(pdf_cfg.get("thematique_uniform_pie_max_ratio", 0.82))
        ),
        "thematique_uniform_chart": _clamp_ratio(
            chart_base * float(pdf_cfg.get("thematique_uniform_chart_scale", 1.0))
        ),
        "thematique_sec21_figure_scale_mult": max(
            0.5, min(2.5, float(pdf_cfg.get("thematique_sec21_figure_scale_mult", 1.55)))
        ),
        "thematique_sec22_resultats_pie_figure_scale_mult": max(
            0.5, min(2.5, float(pdf_cfg.get("thematique_sec22_resultats_pie_figure_scale_mult", 1.22)))
        ),
        "thematique_sec22_resultats_pie_width_ratio_mult": max(
            0.5, min(1.5, float(pdf_cfg.get("thematique_sec22_resultats_pie_width_ratio_mult", 1.12)))
        ),
        "thematique_sec4_activite_pie_width_ratio_mult": max(
            0.5, min(1.5, float(pdf_cfg.get("thematique_sec4_activite_pie_width_ratio_mult", 1.0)))
        ),
        "thematique_sec4_activite_pie_figure_scale_mult": max(
            0.5, min(2.0, float(pdf_cfg.get("thematique_sec4_activite_pie_figure_scale_mult", 1.0)))
        ),
        "figure_scale": max(0.7, min(1.6, float(pdf_cfg.get("figure_scale", 1.0)))),
        "legend_fontsize": max(6.0, min(12.0, float(pdf_cfg.get("legend_fontsize", 8.0)))),
        "legend_ncol_max": max(1.0, min(6.0, float(pdf_cfg.get("legend_ncol_max", 4.0)))),
    }


def clamp_uniform_pie_ratio(
    chart_ratios: dict[str, float],
    *,
    uniform_key: str,
    min_key: str,
    max_key: str,
    fallback_key: str = "pie_base",
) -> float:
    """Borne le ratio d'un camembert entre un minimum et un maximum autorisé."""
    pie_min = float(chart_ratios.get(min_key, 0.70))
    pie_max = float(chart_ratios.get(max_key, 0.82))
    if pie_min > pie_max:
        pie_min, pie_max = pie_max, pie_min
    raw = float(chart_ratios.get(uniform_key, chart_ratios.get(fallback_key, pie_min)))
    return min(pie_max, max(pie_min, raw))


def resolve_reference_pie_display(
    chart_ratios: dict[str, float],
    pie_ratio_base: float,
) -> dict[str, float]:
    """Retourne la configuration de référence du camembert principal pour harmoniser l'ensemble des figures."""
    width_mult = float(
        chart_ratios.get("thematique_sec22_resultats_pie_width_ratio_mult", 1.12)
    )
    figure_mult = float(
        chart_ratios.get("thematique_sec22_resultats_pie_figure_scale_mult", 1.22)
    )
    base_fs = float(chart_ratios.get("figure_scale", 1.0))
    return {
        "width_ratio": min(0.95, float(pie_ratio_base) * width_mult),
        "figure_scale": base_fs * figure_mult,
        "legend_fontsize": float(chart_ratios.get("legend_fontsize", 8.0)),
    }