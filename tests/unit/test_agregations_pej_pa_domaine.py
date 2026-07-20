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

"""Régression : PA dérivées des contrôles à manquement (bilan global)."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from core.engine.agregations_profil import analyse_pej_pa_global


def test_pa_par_domaine_depuis_controles_manquement(tmp_path: Path) -> None:
    point = pd.DataFrame(
        {
            "dc_id": ["dc1", "dc2", "dc3"],
            "date_ctrl": pd.to_datetime(["2025-01-01", "2025-02-01", "2025-03-01"]),
            "resultat": ["Manquement", "Manquement et infraction", "Conforme"],
            "code_pa": ["PA1", "PA2", None],
            "domaine": ["Eau", "Faune", "Flore"],
            "theme": ["Th1", "Th2", "Th3"],
        }
    )
    pej = pd.DataFrame(
        {
            "ENTITE_ORIGINE_PROCEDURE": ["SD21", "SD21"],
            "DC_ID": ["p1", "p2"],
            "DOMAINE": ["A", "B"],
        }
    )
    pa_ods = pd.DataFrame({"DC_ID": ["dc1"]})
    analyse_pej_pa_global(
        tmp_path,
        point,
        pa_ods,
        pej,
        tmp_path,
        echelle="departement",
        code="21",
    )
    pa_dom = pd.read_csv(tmp_path / "pa_global_par_domaine.csv", sep=";")
    assert len(pa_dom) == 2
    assert int(pa_dom["nb_pa"].sum()) == 2
    resume = pd.read_csv(tmp_path / "pa_global_resume.csv", sep=";")
    assert int(resume.iloc[0]["nb_pa_global"]) == 2
