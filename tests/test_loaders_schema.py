from pathlib import Path
from typing import Any

import pandas as pd


class DummyPath(Path):
    _flavour = type(Path())._flavour


def test_load_point_ctrl_missing_required_columns(monkeypatch, tmp_path: Path) -> None:
    """
    Vérifie que load_point_ctrl lève une erreur explicite si une colonne
    obligatoire est absente du GPKG.
    """
    import scripts.common.loaders as loaders

    root = tmp_path
    sources_sig = root / "sources" / "sig"
    year_dir = sources_sig / "points_de_ctrl_OSCEAN_2025"
    year_dir.mkdir(parents=True)
    fake_file = year_dir / "point_ctrl_20250101.gpkg"
    fake_file.write_bytes(b"")  # le contenu est ignoré par le mock

    # Monkeypatch du répertoire sources/sig utilisé par load_point_ctrl
    def fake_sources_sig() -> Path:
        return sources_sig

    monkeypatch.setattr(loaders, "_GPKG_ENGINE", "pyogrio", raising=False)

    # Mock de geopandas.read_file pour renvoyer un DataFrame incomplet
    import geopandas as gpd

    def fake_read_file(*args: Any, **kwargs: Any) -> gpd.GeoDataFrame:
        df = pd.DataFrame(
            [
                {
                    # "date_ctrl" manquant volontairement
                    "dc_id": 1,
                    "num_depart": "21",
                }
            ]
        )
        return gpd.GeoDataFrame(df)

    monkeypatch.setattr("scripts.common.loaders.gpd.read_file", fake_read_file)

    # Appel : on s'attend à une KeyError pour date_ctrl manquant.
    with pytest.raises(KeyError):
        loaders.load_point_ctrl(root, dept_code="21", date_deb="2025-01-01", date_fin="2025-12-31")


def test_load_communes_centroides_missing_insee_column(monkeypatch, tmp_path: Path) -> None:
    """
    Vérifie que load_communes_centroides signale proprement l'absence de
    colonne de code INSEE dans le CSV.
    """
    import scripts.common.loaders as loaders

    root = tmp_path
    sig_dir = root / "ref" / "sig"
    sig_dir.mkdir(parents=True)
    csv_path = sig_dir / "communes-france-2025.csv"
    # CSV volontairement sans colonne code_insee / CODE_INSEE / insee
    csv_path.write_text("foo,latitude_centre,longitude_centre\n1,47.0,5.0\n", encoding="utf-8")

    with pytest.raises(KeyError):
        loaders.load_communes_centroides(root)

