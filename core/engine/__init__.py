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

"""API publique du moteur profilé des bilans."""

from core.engine.catalogue_profils import list_profiles, resolve_profile_ids
from core.engine.registre_sections_pdf import SectionRegistry
from core.engine.execution_lots_profils import run_profile, run_profiles_batch

__all__ = [
    "list_profiles",
    "resolve_profile_ids",
    "run_profile",
    "run_profiles_batch",
    "SectionRegistry",
]
