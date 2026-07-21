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
import geopandas as gpd
import pandas as pd
from pathlib import Path
import sys

# Charger les codes INSEE de la zone TUB
sys.path.append(r'c:\Users\aguirre.maurin\Documents\GitHub\Bilans_production\src')
from ofbilan.engine.orchestrateur_profils import load_tub_pnf_codes
tub_codes, _ = load_tub_pnf_codes(Path(r'c:\Users\aguirre.maurin\Documents\GitHub\Bilans_production'))

path = r'c:\Users\aguirre.maurin\Documents\GitHub\Bilans_production\data\sources\sig\CARTO\pve_infractions.gpkg'
gdf = gpd.read_file(path)

# Simulate QGIS filter
gdf['mif_date'] = pd.to_datetime(gdf['PVe_INF-DATE-MIF'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
mask_date = (gdf['mif_date'] >= '2025-07-01') & (gdf['mif_date'] <= '2026-06-30')
gdf_filtered = gdf[mask_date].copy()

natinf_str = gdf_filtered['PVe_INF-NATINF'].astype(str)
mask_natinf = natinf_str.str.contains('27742|25001', regex=True)
gdf_filtered = gdf_filtered[mask_natinf]

# Ne garder que ceux dont le code INSEE est dans la zone TUB (pour simuler ce qui se passe géographiquement dans la zone)
if 'INSEE_COM' in gdf_filtered.columns:
    insee = gdf_filtered['INSEE_COM'].str.zfill(5)
    gdf_tub = gdf_filtered[insee.isin(tub_codes)]
    print(f'Total PVe in TUB zone (via INSEE): {len(gdf_tub)}')
    unique = gdf_tub.geometry.apply(lambda g: g.wkt).nunique()
    print(f'Unique spatial points in TUB zone: {unique}')
    
    # Detail overlaps
    counts = gdf_tub.geometry.apply(lambda g: g.wkt).value_counts()
    for wkt, count in counts[counts > 1].items():
        print(f'OVERLAP: {count} PVe au point {wkt}')
else:
    print('Pas de colonne INSEE_COM')