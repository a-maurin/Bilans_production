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
import math

pve_tub_path = r'data\out\bilan_tub\pve_tub.csv'
pve_tub = pd.read_csv(pve_tub_path, sep=';', dtype=str, encoding='utf-8')
kept_ids = pve_tub['INF-ID'].dropna().unique().tolist()

try:
    raw_pve = pd.read_csv(r'data\sources\Stats_PVe_OFB au 07.04.2026.csv', sep=';', encoding='cp1252', dtype=str)
except:
    raw_pve = pd.read_csv(r'data\sources\Stats_PVe_OFB au 07.04.2026.csv', sep=';', encoding='iso-8859-1', dtype=str)

df = raw_pve[raw_pve['INF-ID'].isin(kept_ids)]
for _, row in df.iterrows():
    try:
        lat = float(str(row.get('inf_gps_lat')).replace(',', '.'))
        lon = float(str(row.get('inf_gps_long')).replace(',', '.'))
        if lat != 0 and lon != 0 and not math.isnan(lat) and not math.isnan(lon):
            pt = gpd.GeoSeries([Point(lon, lat)], crs='EPSG:4326').to_crs('EPSG:2154').iloc[0]
            print(f"ID {row['INF-ID']}: {pt.x}, {pt.y}")
    except Exception as e:
        pass