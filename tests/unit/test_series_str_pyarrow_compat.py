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