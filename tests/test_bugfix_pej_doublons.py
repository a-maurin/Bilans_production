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

import pandas as pd
import pytest

def test_load_pej_preserves_multiple_nans(monkeypatch, tmp_path):
    from core.common import chargeurs_donnees
    
    # Jeu de données simulé : 1 procédure avec ID, 2 sans ID de liaison (NaN/None)
    mock_df = pd.DataFrame({
        "DC_ID": ["123", None, pd.NA],
        "DATE_CONSTATATION": ["2023-01-01", "2023-01-02", "2023-01-03"],
        "DATE_OUVERTURE_PROCEDURE": [pd.NA, pd.NA, pd.NA],
        "RECAP_DATE_INIT_PJ": [pd.NA, pd.NA, pd.NA]
    })
    
    # Mock des fonctions de lecture de fichiers pour injecter notre DataFrame
    monkeypatch.setattr(chargeurs_donnees, "_find_latest_dated_file", lambda *args, **kwargs: tmp_path / "fake.ods")
    monkeypatch.setattr(chargeurs_donnees, "_read_spreadsheet", lambda *args, **kwargs: mock_df.copy())
    
    # Exécution du chargeur
    res = chargeurs_donnees.load_pej(tmp_path)
    
    # Vérification : on s'attend à récupérer les 3 procédures
    assert len(res) == 3, "Le dédoublonnage a supprimé par erreur les procédures sans DC_ID."
