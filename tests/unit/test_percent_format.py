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
