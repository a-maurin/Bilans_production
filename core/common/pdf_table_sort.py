# Copyright (C) 2026 Aguirre MAURIN
#
# Ce programme est un logiciel libre : vous pouvez le redistribuer et/ou le modifier
# selon les termes de la Licence Publique Générale GNU (GPL) telle que publiée par
# la Free Software Foundation, version 3 de la licence, ou (à votre choix) toute version ultérieure.
#
# Ce programme est distribué dans l'espoir qu'il sera utile, me SANS AUCUNE GARANTIE ;
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
MODULE : GESTION DU TRI ET NORMES D'AFFICHAGE DES TABLEAUX PDF (`pdf_table_sort.py`)
========================================================================================
Ce module régit le classement des données et le nommage des en-têtes dans les tableaux PDF.

Principes clés :
  1. Tri par ordre décroissant de volume d'activité (nombre de contrôles ou de procédures).
  2. Respect d'ordres métier fixes pour certains tableaux (ex: résultats de contrôle ou zones PEJ).
  3. Renommage harmonisé des colonnes et mise en forme des légendes explicatives.
========================================================================================
"""

from __future__ import annotations

import pandas as pd

from core.common.utilitaires_metier import ZONE_PEJ_LECTEUR_TABLE_ORDER

# Colonnes numériques de référence pour déterminer le tri par défaut
_COUNT_COLUMNS_PRIORITY = (
    "nb_localisations",
    "nb_total",
    "nb",
    "nb_pej",
    "nb_pa",
    "nb_pve",
    "Total",
)

# Libellés officiels des en-têtes de colonnes dans les rapports PDF
PDF_LABEL_PEJ = "PEJ"
PDF_LABEL_PEJ_COUNT = "Nombre de PEJ"
PDF_LABEL_CTRL_LOCATIONS = "Localisations de contrôle"
PDF_LABEL_CTRL_LOCATIONS_SHORT = "Loc. de contrôle"
PDF_LABEL_NON_CONFORME_LOCATIONS = "Loc. non-conformes"
PDF_LABEL_EFFECTIFS = "Effectifs"

_PDF_COLUMN_LABELS: dict[str, str] = {
    "domaine": "Domaine",
    "theme": "Thème",
    "type_usager": "Type d'usager",
    "nb_pej": PDF_LABEL_PEJ,
    "nb_pa": "PA",
    "nb_pve": "PVe",
    "nb_localisations": PDF_LABEL_CTRL_LOCATIONS_SHORT,
    "resultat": "Résultat",
    "coeur_hors_coeur": "Zone",
}

_METRIC_SUFFIX_CTRL = " (localisations de contrôle)"
_METRIC_SUFFIX_PROC = " (nombre de procédures)"
_METRIC_SUFFIX_EFFECTIFS = " (effectifs d'usagers)"


# ========================================================================================
# HARMONISATION DES TITRES ET LIBELLES DE COLONNES
# ========================================================================================

def pdf_metric_caption(title: str, metric: str) -> str:
    """Complète le titre d'un graphique ou d'un tableau en ajoutant la métrique (ex: '(nombre de procédures)')."""
    t = str(title).strip()
    low = t.lower()
    if metric == "ctrl" and "localisation" not in low:
        return t + _METRIC_SUFFIX_CTRL
    if metric == "proc" and "procédure" not in low and "pej" not in low and "pve" not in low:
        return t + _METRIC_SUFFIX_PROC
    if metric == "effectifs" and "effectif" not in low:
        return t + _METRIC_SUFFIX_EFFECTIFS
    return t


def pdf_column_label(col: str) -> str:
    """Convertit le nom technique d'une colonne DataFrame en libellé lisible pour le PDF."""
    key = str(col).strip()
    if key in _PDF_COLUMN_LABELS:
        return _PDF_COLUMN_LABELS[key]
    return key.replace("_", " ").title()


# Ordre fixe d'affichage des lignes du tableau de bilan des contrôles
_RESULTATS_CONTROLES_ROW_ORDER: tuple[str, ...] = (
    "Conforme",
    "Non-conforme",
    "Dont manquement",
    "Dont infraction",
    "En attente",
)


def resultat_controle_label_for_pdf(resultat: object) -> str:
    """Ajoute des espaces incécables d'indentation HTML pour les sous-lignes 'Dont...'."""
    label = str(resultat or "").strip()
    if label in ("Dont infraction", "Dont manquement"):
        return f"&nbsp;&nbsp;&nbsp;{label}"
    return label


def sort_tab_resultats_controles_for_pdf(df: pd.DataFrame) -> pd.DataFrame:
    """Conserve l'ordre métier réglementaire fixe des résultats de contrôle."""
    if df.empty or "resultat" not in df.columns:
        return df

    def _rank(val: object) -> int:
        key = str(val or "").strip()
        try:
            return _RESULTATS_CONTROLES_ROW_ORDER.index(key)
        except ValueError:
            return len(_RESULTATS_CONTROLES_ROW_ORDER)

    tmp = df.copy()
    tmp["_pdf_row_order"] = tmp["resultat"].map(_rank)
    return (
        tmp.sort_values("_pdf_row_order", ascending=True, kind="stable")
        .drop(columns=["_pdf_row_order"])
        .reset_index(drop=True)
    )


# ========================================================================================
# FONCTIONS GENERALES DE TRI DECROISSANT DE DATAFRAMES
# ========================================================================================

def sort_dataframe_desc(df: pd.DataFrame | None, columns: list[str]) -> pd.DataFrame | None:
    """Trie un DataFrame selon la première colonne disponible de la liste (ordre décroissant)."""
    if df is None or df.empty:
        return df
    for col in columns:
        if col in df.columns:
            return df.sort_values(by=col, ascending=False, kind="stable").reset_index(drop=True)
    return df


def sort_dataframe_desc_auto(df: pd.DataFrame | None) -> pd.DataFrame | None:
    """Détecte automatiquement la colonne numérique principale et trie en ordre décroissant."""
    return sort_dataframe_desc(df, list(_COUNT_COLUMNS_PRIORITY))


def sort_dataframe_desc_by_sum(
    df: pd.DataFrame | None,
    sum_columns: list[str] | None = None,
) -> pd.DataFrame | None:
    """Trie en ordre décroissant selon la somme calculée de plusieurs colonnes numériques `nb_*`."""
    if df is None or df.empty:
        return df
    cols = sum_columns or [c for c in df.columns if str(c).startswith("nb_")]
    if not cols:
        return sort_dataframe_desc_auto(df)
    tmp = df.copy()
    tmp["_pdf_sort"] = tmp[cols].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)
    return sort_dataframe_desc(tmp, ["_pdf_sort"]).drop(columns=["_pdf_sort"])


def sort_detail_dataframe_by_date_desc(
    df: pd.DataFrame | None,
    date_col: str = "date",
) -> pd.DataFrame | None:
    """Trie les enregistrements par date décroissante (les événements récents apparaissent en premier)."""
    if df is None or df.empty or date_col not in df.columns:
        return df
    tmp = df.copy()
    tmp["_pdf_date"] = pd.to_datetime(tmp[date_col], errors="coerce")
    return (
        tmp.sort_values("_pdf_date", ascending=False, kind="stable")
        .drop(columns=["_pdf_date"])
        .reset_index(drop=True)
    )


def sort_zone_pej_for_pdf(df: pd.DataFrame) -> pd.DataFrame:
    """Ordre d'affichage spécifique pour le tableau PEJ par zone géographique."""
    if df is None or df.empty or "zone" not in df.columns:
        return df
    zones = df["zone"].astype(str)
    if zones.isin(ZONE_PEJ_LECTEUR_TABLE_ORDER).any():
        ordered = pd.Categorical(zones, categories=list(ZONE_PEJ_LECTEUR_TABLE_ORDER), ordered=True)
        return df.assign(_zone_ord=ordered).sort_values("_zone_ord").drop(columns="_zone_ord")
    return sort_dataframe_desc(df, ["nb"])


# ========================================================================================
# TRI ET FORMATAGE DE TOUS LES TABLEAUX DES SECTIONS PDF
# ========================================================================================

def prepare_pdf_results_sec23_sorting(results: dict) -> None:
    """Parcourt l'ensemble des DataFrames du dictionnaire de résultats et leur applique le tri approprié."""
    column_sorts: list[tuple[str, list[str]]] = [
        ("usager_effectifs", ["nb"]),
        ("tab_resultats", ["nb"]),
        ("pve_top_infractions", ["nb"]),
        ("pve_natinf_analysis", ["nb"]),
        ("zone_pve", ["nb"]),
        ("pej_top_infractions", ["nb"]),
        ("pej_par_theme", ["nb_pej", "nb"]),
        ("pej_natinf_analysis", ["nb"]),
        ("pej_clotur", ["nb"]),
        ("pej_suite", ["nb"]),
        ("res_par_usager_domaine", ["nb_localisations", "nb_conforme", "nb_manquement", "nb_infraction"]),
        ("zone_ctrl", ["nb_total"]),
        ("agg_theme", ["nb"]),
        ("agg_commune", ["nb_localisations", "nb"]),
        ("pa_par_theme", ["nb_pa", "nb"]),
    ]
    for key, cols in column_sorts:
        value = results.get(key)
        if isinstance(value, pd.DataFrame) and not value.empty:
            results[key] = sort_dataframe_desc(value, cols)

    zone_pej = results.get("zone_pej")
    if isinstance(zone_pej, pd.DataFrame) and not zone_pej.empty:
        results["zone_pej"] = sort_zone_pej_for_pdf(zone_pej)

    tab_res_ctrl = results.get("tab_resultats_controles")
    if isinstance(tab_res_ctrl, pd.DataFrame) and not tab_res_ctrl.empty:
        results["tab_resultats_controles"] = sort_tab_resultats_controles_for_pdf(tab_res_ctrl)

    for key in ("proc_par_usager_domaine", "proc_par_usager_theme"):
        value = results.get(key)
        if isinstance(value, pd.DataFrame) and not value.empty:
            results[key] = sort_dataframe_desc_by_sum(value)

    for key in ("pve_detail", "pej_detail", "pa_detail"):
        value = results.get(key)
        if isinstance(value, pd.DataFrame) and not value.empty:
            results[key] = sort_detail_dataframe_by_date_desc(value)


def build_resultats_par_usager_domaine_pdf_rows(
    df: pd.DataFrame,
    *,
    is_single_usager: bool,
    max_rows: int = 15,
) -> tuple[list[str], list[list[str]], bool]:
    """Prépare les lignes de données formatées du tableau 'Résultats par Domaine'."""
    if df is None or df.empty:
        return [], [], False

    work = df.copy()
    for col in ("nb_conforme", "nb_manquement", "nb_infraction", "nb_en_attente"):
        if col not in work.columns:
            work[col] = 0
    work["_sort_res"] = work["nb_localisations"].fillna(0).astype(int)
    work = work.sort_values("_sort_res", ascending=False, kind="stable").head(max_rows)

    res_field_cols: list[tuple[str, str]] = [
        ("Conforme", "nb_conforme"),
        ("Manquement", "nb_manquement"),
        ("Infraction", "nb_infraction"),
    ]
    if int(work["nb_en_attente"].fillna(0).sum()) > 0:
        res_field_cols.append(("En attente", "nb_en_attente"))
    res_cols = [label for label, _ in res_field_cols]

    with_type_col = not (
        is_single_usager
        and "type_usager" in work.columns
        and work["type_usager"].nunique() == 1
    )
    if with_type_col:
        header = ["Type d'usager", "Domaine", *res_cols]
    else:
        work = work.drop(columns=["type_usager"])
        header = ["Domaine", *res_cols]

    body: list[list[str]] = []
    for _, row in work.iterrows():
        base = [str(row.get("domaine", ""))] + [
            str(int(row.get(field, 0))) for _, field in res_field_cols
        ]
        if with_type_col:
            body.append([str(row.get("type_usager", "")), *base])
        else:
            body.append(base)

    return header, body, with_type_col