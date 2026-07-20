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

"""Classification des résultats de contrôle (section 2.2)."""

from __future__ import annotations

import pandas as pd

from core.common.utilitaires_metier import (
    build_tab_resultats_controles,
    classify_resultat_controle,
)


def test_classify_autres_libelles_en_attente() -> None:
    assert classify_resultat_controle("Conforme") == "Conforme"
    assert classify_resultat_controle("Infraction") == "Infraction"
    assert classify_resultat_controle("Manquement") == "Manquement"
    assert classify_resultat_controle("") == "En attente"
    assert classify_resultat_controle("Non-conforme") == "En attente"
    assert classify_resultat_controle("En cours") == "En attente"


def test_build_tab_resultats_controles_agrege_en_attente() -> None:
    point = pd.DataFrame(
        {
            "resultat": [
                "Conforme",
                "Infraction",
                "Manquement",
                "Non-conforme",
                "",
                "Autre",
            ],
        }
    )
    tab = build_tab_resultats_controles(point)
    assert [str(x).strip() for x in tab["resultat"]] == [
        "Conforme",
        "Non-conforme",
        "Dont manquement",
        "Dont infraction",
        "En attente",
    ]
    assert int(tab.loc[tab["resultat"] == "Conforme", "nb"].sum()) == 1
    assert int(tab.loc[tab["resultat"] == "Non-conforme", "nb"].sum()) == 2
    assert int(tab.loc[tab["resultat"] == "En attente", "nb"].sum()) == 3
