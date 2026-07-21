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
import sys
import json
import traceback
import os
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

out_path = project_root / "test_output.txt"
with open(out_path, "w", encoding="utf-8") as f:
    try:
        from ofbilan.common.chargeurs_donnees import load_pnf_aoa_gdf
        gdf_boundary = load_pnf_aoa_gdf(Path(project_root))
        if not gdf_boundary.empty:
            if gdf_boundary.crs is None:
                gdf_boundary.set_crs(epsg=2154, inplace=True)
            gdf_boundary_wgs84 = gdf_boundary.to_crs("EPSG:4326")
            geojson_data = json.loads(gdf_boundary_wgs84.to_json())
            f.write(f"SUCCESS! Length of features: {len(geojson_data.get('features', []))}\n")
        else:
            f.write("EMPTY GDF\n")
    except Exception as e:
        f.write("ERROR:\n")
        traceback.print_exc(file=f)