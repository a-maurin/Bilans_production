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

import sys
import os

osgeo4w_root = r"C:\Program Files\QGIS 3.44.8"
sys.path.insert(0, os.path.join(osgeo4w_root, r"apps\qgis-ltr\python"))
sys.path.insert(0, os.path.join(osgeo4w_root, r"apps\Python312\Lib\site-packages"))

from qgis.core import QgsApplication, QgsVectorLayer, QgsFeatureRequest

qgs = QgsApplication([], False)
qgs.initQgis()

layer = QgsVectorLayer(r"C:\Users\aguirre.maurin\Documents\GitHub\Bilans_production\data\sources\sig\points_infractions_pj\localisation_infrac_FAITS_20260505.gpkg", "test", "ogr")
req = QgsFeatureRequest()
for i, f in enumerate(layer.getFeatures(req)):
    print(list(f.fields().names()))
    break
qgs.exitQgis()
