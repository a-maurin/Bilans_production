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

"""Package principal pour la génération des bilans."""

import warnings

# --- CONFIGURATION GLOBALE PANDAS ---
# Solution pérenne pour désactiver le backend PyArrow pour les chaînes de caractères.
# Dans l'environnement QGIS (Pandas 2.1+), PyArrow est souvent activé par défaut, mais la
# version de PyArrow fournie manque de certaines fonctionnalités regex (ex: replace_substring_regex).
# En forçant le mode "python", Pandas utilisera le moteur natif (module re) pour toutes les séries.
try:
    import pandas as pd
    
    # Pandas 2.0+
    try:
        if hasattr(pd.options.mode, "string_storage"):
            pd.options.mode.string_storage = "python"
    except Exception:
        pass
        
    # Option future pour Pandas 2.1+
    try:
        if hasattr(pd.options.future, "infer_string"):
            pd.options.future.infer_string = False
    except Exception:
        pass
except ImportError:
    pass
except Exception as e:
    warnings.warn(f"Impossible de configurer globalement le backend string de Pandas : {e}")
