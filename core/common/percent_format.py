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
"""Affichage des pourcentages en entiers (PDF, tableaux, légendes de graphiques)."""
from __future__ import annotations

from typing import Sequence

import pandas as pd


def format_pct_int_from_rate(rate: float | None, *, na: str = "n.d.") -> str:
    """Formate un taux dans [0, 1] (ou proche) en pourcentage entier, ex. ``42 %``."""
    if rate is None or pd.isna(rate):
        return na
    try:
        r = float(rate)
    except (TypeError, ValueError):
        return na
    p = int(round(r * 100.0))
    p = max(0, min(100, p))
    return f"{p} %"


def int_percents_largest_remainder(counts: Sequence[int]) -> list[int]:
    """
    Répartit 100 points de pourcentage sur des effectifs entiers (méthode des plus grands restes).

    La somme des pourcentages retournés vaut exactement 100 dès que ``sum(counts) > 0``.
    """
    counts = [int(max(0, c)) for c in counts]
    total = sum(counts)
    n = len(counts)
    if n == 0:
        return []
    if total <= 0:
        return [0] * n
    longs = [100 * c for c in counts]
    floors = [ln // total for ln in longs]
    rem = [ln % total for ln in longs]
    deficit = 100 - sum(floors)
    order = sorted(range(n), key=lambda i: (rem[i], counts[i], i), reverse=True)
    for k in range(deficit):
        floors[order[k]] += 1
    return floors


def format_partition_pct_strings(counts: Sequence[int]) -> list[str]:
    """Même logique que ``int_percents_largest_remainder``, avec suffixe ``%``."""
    return [f"{p} %" for p in int_percents_largest_remainder(counts)]


def tab_counts_to_pct_strings(nbs: Sequence[int]) -> list[str]:
    """Tableau « Nombre / Taux » : taux entiers dont la somme vaut 100 % des lignes."""
    counts = [int(max(0, x)) for x in nbs]
    if not counts or sum(counts) <= 0:
        return ["n.d."] * len(counts)
    return format_partition_pct_strings(counts)