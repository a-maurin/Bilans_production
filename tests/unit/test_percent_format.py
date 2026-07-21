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
"""Tests pour l'affichage des pourcentages entiers (somme 100 % sur les répartitions)."""
from core.common.percent_format import (
    format_pct_int_from_rate,
    int_percents_largest_remainder,
    tab_counts_to_pct_strings,
)


def test_partition_sums_to_100():
    assert sum(int_percents_largest_remainder([33, 33, 33])) == 100
    assert sum(int_percents_largest_remainder([1, 1, 98])) == 100
    assert sum(int_percents_largest_remainder([40, 60])) == 100


def test_tab_counts_strings_sum_parsed():
    s = tab_counts_to_pct_strings([10, 30, 60])
    vals = [int(x.replace("%", "").strip()) for x in s]
    assert sum(vals) == 100


def test_format_rate_int():
    assert format_pct_int_from_rate(0.333) == "33 %"
    assert format_pct_int_from_rate(1.0) == "100 %"
    assert format_pct_int_from_rate(None) == "n.d."