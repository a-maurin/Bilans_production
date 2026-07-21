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
import pandas as pd
import geopandas as gpd
from shapely.geometry import Point
from pathlib import Path
import math

pve_tub_path = r'data\out\bilan_tub\pve_tub.csv'
pve_tub = pd.read_csv(pve_tub_path, sep=';', dtype=str, encoding='utf-8')
kept_ids = pve_tub['INF-ID'].dropna().unique().tolist()

raw_pve_path = r'data\sources\Stats_PVe_OFB au 07.04.2026.csv'
try:
    raw_pve = pd.read_csv(raw_pve_path, sep=';', encoding='cp1252', dtype=str)
except:
    raw_pve = pd.read_csv(raw_pve_path, sep=';', encoding='iso-8859-1', dtype=str)

tub_poly_path = r'ref\hors_programme\sig\TUB\2025 - 2026\COMMUNE_InterdAgrain_2026.shp'
tub_poly = gpd.read_file(tub_poly_path)
if tub_poly.crs is None:
    tub_poly.set_crs(epsg=2154, inplace=True)

df = raw_pve[raw_pve['INF-ID'].isin(kept_ids)].copy()

case_orphans = []
case_inside = []
case_outside = []

for _, row in df.iterrows():
    inf_id = row['INF-ID']
    lat_str = str(row.get('inf_gps_lat', '')).replace(',', '.')
    lon_str = str(row.get('inf_gps_long', '')).replace(',', '.')
    
    try:
        lat = float(lat_str)
        lon = float(lon_str)
        if lat == 0.0 or lon == 0.0 or math.isnan(lat) or math.isnan(lon):
            raise ValueError()
        valid_coords = True
    except:
        valid_coords = False

    if not valid_coords:
        case_orphans.append(inf_id)
        continue
        
    pt = gpd.GeoSeries([Point(lon, lat)], crs="EPSG:4326").to_crs(tub_poly.crs).iloc[0]
    intersects = tub_poly.geometry.contains(pt).any()
    
    if intersects:
        case_inside.append(inf_id)
    else:
        case_outside.append(inf_id)

print(f"Case 1 (Dans le polygone TUB avec GPS) - {len(case_inside)} PVe :")
print(", ".join(case_inside))
print(f"\nCase 2 (Hors du polygone mais rattachés à une commune TUB) - {len(case_outside)} PVe :")
print(", ".join(case_outside))
print(f"\nCase 3 (Orphelins sans coordonnées GPS valides) - {len(case_orphans)} PVe :")
print(", ".join(case_orphans))