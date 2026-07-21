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
"""Tests des défauts interactifs de période (ask_periode_perimetre)."""

from __future__ import annotations

from datetime import datetime
from unittest.mock import patch

import pytest

from core.common import prompt_periode as pp


def test_default_date_deb_is_first_day_of_current_year() -> None:
    fixed = datetime(2026, 5, 28, 12, 0, 0)
    with patch("core.common.prompt_periode.dt.datetime") as mock_dt:
        mock_dt.now.return_value = fixed
        mock_dt.strptime = datetime.strptime
        assert pp._default_date_deb() == "2026-01-01"


def test_default_date_fin_is_today() -> None:
    fixed = datetime(2026, 5, 28, 12, 0, 0)
    with patch("core.common.prompt_periode.dt.datetime") as mock_dt:
        mock_dt.now.return_value = fixed
        mock_dt.strptime = datetime.strptime
        assert pp._default_date_fin() == "2026-05-28"


def test_ask_periode_perimetre_uses_dynamic_defaults_on_empty_input() -> None:
    fixed = datetime(2026, 5, 28, 12, 0, 0)
    inputs = iter(["", "", "", ""])
    with (
        patch("core.common.prompt_periode.dt.datetime") as mock_dt,
        patch.object(pp, "_is_interactive", return_value=True),
        patch("builtins.input", side_effect=lambda _prompt: next(inputs)),
    ):
        mock_dt.now.return_value = fixed
        mock_dt.strptime = datetime.strptime
        deb, fin, echelle, code = pp.ask_periode_perimetre()

    assert deb == "2026-01-01"
    assert fin == "2026-05-28"
    assert echelle == "departement"
    assert code == "21"


def test_ask_periode_perimetre_non_interactive_requires_dates() -> None:
    with patch.object(pp, "_is_interactive", return_value=False):
        with pytest.raises(ValueError, match="non interactif"):
            pp.ask_periode_perimetre()