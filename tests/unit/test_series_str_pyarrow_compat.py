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

"""Compatibilité filtrage texte pandas / PyArrow (Python QGIS)."""

import pandas as pd
import pytest

from core.common.utilitaires_metier import (
    coalesced_insee_series,
    extract_insee_code_series,
    series_str_contains,
)


def test_series_str_contains_insensitive_literal():
    s = pd.Series(["Agrain 2025", "chasse", None])
    mask = series_str_contains(s, "agrain", regex=False)
    assert mask.tolist() == [True, False, False]


def test_series_str_contains_regex_on_lowered_series():
    s = pd.Series(["Police sanitaire", "TUBERCULOSE", "ok"])
    mask = series_str_contains(s, r"tubercul|grippe", regex=True)
    assert mask.tolist() == [False, True, False]


def test_extract_insee_code_series():
    s = pd.Series([" 21054 ", "invalid", "1234", None])
    got = extract_insee_code_series(s)
    assert got.iloc[0] == "21054"
    assert pd.isna(got.iloc[1])
    assert got.iloc[2] == "01234"
    assert pd.isna(got.iloc[3])


def test_coalesced_insee_series_from_columns():
    df = pd.DataFrame({"insee_comm": [pd.NA, "21054"], "INF-INSEE": ["", "21999"]})
    got = coalesced_insee_series(df)
    assert pd.isna(got.iloc[0])
    assert str(got.iloc[1]) == "21054"


def test_series_str_contains_avoids_pyarrow_string_dtype():
    try:
        s = pd.Series(["Agrain"], dtype=pd.StringDtype(storage="pyarrow"))
    except (TypeError, ImportError):
        pytest.skip("string[pyarrow] indisponible sur cet environnement")
    mask = series_str_contains(s, "agrain", regex=False)
    assert bool(mask.iloc[0])
