import pandas as pd
import geopandas as gpd
from pathlib import Path

root = Path(r"c:\Users\aguirre.maurin\Documents\GitHub\OFBilan-Plugin-QGIS")
sources = root / "data" / "sources"

# 1. PVe
pve_files = list(sources.glob("Stats_PVe_OFB*.xlsx")) + list(sources.glob("Stats_PVe_OFB*.ods")) + list(sources.glob("Stats_PVe_OFB*.csv"))
if pve_files:
    latest_pve = sorted(pve_files, key=lambda f: f.stat().st_mtime)[-1]
    print("PVe file:", latest_pve.name)
    df_pve = pd.read_excel(latest_pve) if latest_pve.suffix == ".xlsx" else pd.read_csv(latest_pve, sep=";")
    cols = [c for c in df_pve.columns if "UNITE" in c.upper() or "ENTITE" in c.upper() or "SERVICE" in c.upper() or "AGENT" in c.upper()]
    print("PVe columns matching UNITE/ENTITE:", cols)
    if "UNITE_libelle" in df_pve.columns:
        print("PVe UNITE_libelle sample values:")
        print(df_pve["UNITE_libelle"].dropna().value_counts().head(20))

# 2. PEJ (localisation_infrac_FAITS_*.gpkg)
pej_files = list((sources / "sig" / "point_infraction_PJ").glob("localisation_infrac_FAITS_*.gpkg"))
if pej_files:
    latest_pej = sorted(pej_files, key=lambda f: f.stat().st_mtime)[-1]
    print("\nPEJ file:", latest_pej.name)
    gdf_pej = gpd.read_file(latest_pej)
    print("PEJ columns matching entite:", [c for c in gdf_pej.columns if "entite" in c.lower()])
    if "entite" in gdf_pej.columns:
        print("PEJ entite sample values:")
        print(gdf_pej["entite"].dropna().value_counts().head(20))

# 3. Points ctrl (point_ctrl_*_wgs84.gpkg)
ctrl_dirs = list((sources / "sig").glob("points_de_ctrl_OSCEAN_*"))
ctrl_files = []
for d in ctrl_dirs:
    ctrl_files.extend(d.glob("point_ctrl_*_wgs84.gpkg"))
if ctrl_files:
    latest_ctrl = sorted(ctrl_files, key=lambda f: f.stat().st_mtime)[-1]
    print("\nCtrl file:", latest_ctrl.name)
    gdf_ctrl = gpd.read_file(latest_ctrl)
    print("Ctrl columns matching entite:", [c for c in gdf_ctrl.columns if "entite" in c.lower() or "entit" in c.lower()])
    for col in [c for c in gdf_ctrl.columns if "entit" in c.lower()]:
        print(f"Ctrl {col} sample values:")
        print(gdf_ctrl[col].dropna().value_counts().head(20))
