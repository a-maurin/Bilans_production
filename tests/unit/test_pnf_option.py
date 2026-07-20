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

"""Option PNF et distinction cœur / hors-cœur."""

from __future__ import annotations

from core.chemins_projet import PROJECT_ROOT
from core.engine.orchestrateur_profils import load_profile_config, resolve_options


def test_types_usager_cible_pnf_desactive_par_defaut() -> None:
    profile = load_profile_config(PROJECT_ROOT, "types_usager_cible")
    opts = resolve_options(profile, {})
    assert opts.get("pnf") is False
    assert profile.get("options", {}).get("pnf", {}).get("ask") is True


def test_chasse_pnf_active_par_defaut_sans_question() -> None:
    profile = load_profile_config(PROJECT_ROOT, "chasse")
    opts = resolve_options(profile, {})
    assert opts.get("pnf") is True
    assert profile.get("options", {}).get("pnf", {}).get("ask") is False
