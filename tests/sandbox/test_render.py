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
import os

osgeo4w_root = r"C:\Program Files\QGIS 3.44.8"
sys.path.insert(0, os.path.join(osgeo4w_root, r"apps\qgis-ltr\python"))
sys.path.insert(0, os.path.join(osgeo4w_root, r"apps\Python312\Lib\site-packages"))
sys.path.insert(0, r"C:\Users\aguirre.maurin\Documents\GitHub\Bilans_production\src\bilans\cartographie")

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QT_QPA_FONTDIR", r"C:\Windows\Fonts")

from qgis.core import QgsApplication, QgsProject
from production_cartographique import get_effective_config, resolve_profile_layers, apply_layer_symbology, apply_date_filter
from config_cartes import CONFIG

app = QgsApplication([], False)
app.initQgis()

project_path = r"C:\Users\aguirre.maurin\Documents\GitHub\Bilans_production\ref\programme\sig\bilans_carte.qgz"

proj = QgsProject.instance()
proj.read(project_path)

config = get_effective_config()
prof = config.profiles["global_domaines"]

layer = proj.mapLayersByName("point_ctrl_20251231_wgs84")[0]
layer_cfg = prof.layers["point_ctrl_20260205_wgs84"]

apply_layer_symbology(layer, layer_cfg)
apply_date_filter(layer, layer_cfg, prof.date_deb, prof.date_fin, config=config, profile=prof)

renderer = layer.renderer()
print("Renderer type:", renderer.type())
if renderer.type() == "categorizedSymbol":
    print("Expression:", renderer.classAttribute())
    for cat in renderer.categories():
        print("  - Category:", cat.value(), cat.label(), cat.symbol().color().name())

print("Filter:", layer.subsetString())

# Let's count features!
features = [f for f in layer.getFeatures()]
print("Features count:", len(features))

app.exitQgis()