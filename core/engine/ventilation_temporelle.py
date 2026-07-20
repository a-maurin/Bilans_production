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

"""Règles de ventilation temporelle (mode ``auto``)."""

from __future__ import annotations

# Bornes en jours (durée inclusive date_fin − date_deb).
VENTILATION_JOURS_SIX_MOIS = 183  # ~6 mois civil
VENTILATION_JOURS_UN_AN = 366
VENTILATION_JOURS_DEUX_ANS = 730


def resolve_ventilation_auto(duree_jours: int, *, seuil_jours: int = 366) -> str:
    """Ventilation en mode ``auto`` selon la durée de la période d'analyse (en jours).

    - ≤ 6 mois (183 j) : hebdomadaire (libellé de période ``YYYY-Sww``)
    - > 6 mois et ≤ 1 an (366 j) : mensuelle (libellé ``YYYY-MM``)
    - > 1 an et < 2 ans (730 j) : mensuelle
    - ≥ 2 ans (730 j) : trimestrielle (2 ans exact inclus)
    - au-delà du ``seuil_jours`` du profil (si ≥ 2 ans) : annuelle
    """
    if duree_jours <= VENTILATION_JOURS_SIX_MOIS:
        return "hebdomadaire"
    if duree_jours <= VENTILATION_JOURS_UN_AN:
        return "mensuelle"
    if duree_jours < VENTILATION_JOURS_DEUX_ANS:
        return "mensuelle"
    if seuil_jours >= VENTILATION_JOURS_DEUX_ANS and duree_jours > seuil_jours:
        return "annuelle"
    return "trimestrielle"
