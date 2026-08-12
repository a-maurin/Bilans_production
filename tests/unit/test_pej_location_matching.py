import pytest
import pandas as pd
import numpy as np
from pathlib import Path
from core.common.chargeurs_donnees import merge_pej_faits_locations

def test_merge_pej_faits_locations_of_prefix_and_numproc(tmp_path):
    # Mock PEJ dataframe
    pej = pd.DataFrame([
        {
            "DC_ID": "OF20260707-101",
            "NUMERO_PROCEDURE": "SD21 2026 PJ 0054",
            "ENTITE_ORIGINE_PROCEDURE": "SD21",
            "DATE_REF": pd.Timestamp("2026-07-07"),
            "NOM_COM": np.nan,
        },
        {
            "DC_ID": "OF20260707-102",
            "NUMERO_PROCEDURE": "SD21 2026 PJ 0055",
            "ENTITE_ORIGINE_PROCEDURE": "SD21",
            "DATE_REF": pd.Timestamp("2026-07-07"),
            "NOM_COM": "Recey-sur-Ource",
        }
    ])

    # Mock FAITS layer dataframe
    gdf_faits = pd.DataFrame([
        # Match via cleaned ID (20260707-101) without OF prefix
        {
            "dossier": "20260707-101",
            "entite": "SD21 - Côte-d'Or",
            "x_infrac": 5.0,
            "y_infrac": 47.0,
            "commune_fait": "Recey-sur-Ource",
        }
    ])

    merged = merge_pej_faits_locations(pej, tmp_path, "dept", "21", gdf_faits=gdf_faits)
    
    assert len(merged) == 2
    row1 = merged[merged["DC_ID"] == "OF20260707-101"].iloc[0]
    assert row1["x_faits"] == 5.0
    assert row1["y_faits"] == 47.0
    assert row1["precision_loc"] == "GPS Fait (Exacte)"
    assert row1["NOM_COM"] == "Recey-sur-Ource"
