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

import os
from qgis.core import QgsApplication, QgsProject, QgsVectorLayer, QgsLegendPatchShape, QgsStyle

QgsApplication.setPrefixPath("C:/OSGeo4W/apps/qgis", True)
qgs = QgsApplication([], False)
qgs.initQgis()

print("QGIS Init OK")
layer = QgsVectorLayer("Polygon", "test", "memory")
print("Layer created")

legend = layer.legend()
print(type(legend))

try:
    style = QgsStyle.defaultStyle()
    print("Style default:", style)
    patch = style.legendPatchShape(0, style.legendPatchShapeNames()[0])
    print("Patch shape:", patch)
except Exception as e:
    print("Error:", e)

qgs.exitQgis()
