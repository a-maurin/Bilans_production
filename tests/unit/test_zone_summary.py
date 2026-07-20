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

"""Non-régression : décompte zone « Département » = hors TUB et hors PNF."""

from __future__ import annotations

import pandas as pd

from core.common.utilitaires_metier import (
    ZONE_KEY_DEPARTEMENT,
    ZONE_LABEL_DEPARTEMENT_HORS,
    _zone_count,
    _zone_summary,
    zone_table_display_label,
)


def test_zone_summary_departement_counts_hors_tub_pnf() -> None:
    df = pd.DataFrame(
        {
            "insee": ["01001", "01002", "01003", "01004"],
            "resultat": ["Conforme", "Non conforme", "Conforme", "Conforme"],
        }
    )
    tub = {"01001", "01002"}
    pnf = {"01002", "01003"}
    summary = _zone_summary(df, "insee", tub, pnf)
    dep = summary.loc[summary["zone"] == ZONE_KEY_DEPARTEMENT].iloc[0]
    assert int(dep["nb_total"]) == 1
    assert int(dep["nb_non_conforme"]) == 0


def test_zone_count_departement_hors_tub_pnf() -> None:
    df = pd.DataFrame({"insee": ["01001", "01002", "01003"]})
    tub = {"01001"}
    pnf = {"01002"}
    counts = _zone_count(df, "insee", tub, pnf)
    dep_nb = int(counts.loc[counts["zone"] == ZONE_KEY_DEPARTEMENT, "nb"].iloc[0])
    assert dep_nb == 1


def test_zone_table_display_label() -> None:
    assert zone_table_display_label("Département") == ZONE_LABEL_DEPARTEMENT_HORS
    assert zone_table_display_label("Zone TUB") == "Zone TUB"
