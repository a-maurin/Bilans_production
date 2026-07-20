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

from __future__ import annotations

import pandas as pd

from core.engine.agregations_profil import _resultats_par_domaine_pour_pdf


def test_resultats_par_domaine_conforme_nc_attente() -> None:
    pt = pd.DataFrame(
        {
            "domaine": ["A", "A", "B", "B", "B", "C"],
            "resultat": [
                "Conforme",
                "Infraction",
                "Conforme",
                "Manquement",
                "Conforme",
                "Autre",
            ],
        }
    )
    out = _resultats_par_domaine_pour_pdf(pt)
    row_b = out.loc[out["domaine"] == "B"].iloc[0]
    assert int(row_b["Conforme"]) == 2
    assert int(row_b["Non-conforme"]) == 1
    assert int(row_b["En attente"]) == 0

    row_c = out.loc[out["domaine"] == "C"].iloc[0]
    assert int(row_c["Conforme"]) == 0
    assert int(row_c["Non-conforme"]) == 0
    assert int(row_c["En attente"]) == 1
