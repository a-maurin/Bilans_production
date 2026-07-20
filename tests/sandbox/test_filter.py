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

from qgis.core import QgsApplication, QgsVectorLayer, QgsFeatureRequest, QgsExpressionContext, QgsExpressionContextUtils

qgs = QgsApplication([], False)
qgs.initQgis()

layer = QgsVectorLayer(r"C:\Users\aguirre.maurin\Documents\GitHub\Bilans_production\data\sources\sig\points_de_ctrl_OSCEAN_2026\point_ctrl_20260505_wgs84.gpkg", "test", "ogr")

print("Layer valid?", layer.isValid())
print("Total features:", layer.featureCount())

# Test filter global
expr = '"num_depart" = \'21\' AND to_date("date_ctrl") >= to_date(\'2025-01-01\') AND to_date("date_ctrl") <= to_date(\'2025-12-31\')'
req = QgsFeatureRequest().setFilterExpression(expr)
count = 0
for f in layer.getFeatures(req):
    count += 1
print(f"Filter 1 ({expr}): {count} matches")

expr2 = '"num_depart" = 21'
req2 = QgsFeatureRequest().setFilterExpression(expr2)
count2 = 0
for f in layer.getFeatures(req2):
    count2 += 1
print(f"Filter 2 ({expr2}): {count2} matches")

expr3 = 'to_date("date_ctrl") >= to_date(\'2025-01-01\') AND to_date("date_ctrl") <= to_date(\'2025-12-31\')'
req3 = QgsFeatureRequest().setFilterExpression(expr3)
count3 = 0
for f in layer.getFeatures(req3):
    count3 += 1
print(f"Filter 3 ({expr3}): {count3} matches")

qgs.exitQgis()
