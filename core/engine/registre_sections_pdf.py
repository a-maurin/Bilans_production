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

"""
Contrat pour un futur registre de sections PDF partagé global / thématique.

Aujourd'hui, le rendu PDF reste implémenté dans
``bilan_global.analyse_global`` et ``bilan_thematique.bilan_thematique_engine``.
Ce module fournit un squelette minimal (enregistrement / résolution) pour
factoriser progressivement les sections sans disperser des ``if profil``.

Usage envisagé :

    from core.engine.registre_sections_pdf import SectionRegistry

    registry = SectionRegistry()
    registry.register("sec4", render_sec4_thematic)
    registry.render("sec4", context)
"""

from __future__ import annotations

from collections.abc import Callable, Mapping
from typing import Any

SectionRenderer = Callable[[Mapping[str, Any]], None]


class SectionRegistry:
    """Registre ``section_id -> fonction de rendu``."""

    __slots__ = ("_renderers",)

    def __init__(self) -> None:
        self._renderers: dict[str, SectionRenderer] = {}

    def register(self, section_id: str, renderer: SectionRenderer) -> None:
        sid = str(section_id).strip()
        if not sid:
            raise ValueError("section_id vide")
        self._renderers[sid] = renderer

    def get(self, section_id: str) -> SectionRenderer | None:
        return self._renderers.get(str(section_id).strip())

    def render(self, section_id: str, context: Mapping[str, Any]) -> None:
        renderer = self.get(section_id)
        if renderer is None:
            raise KeyError(f"Aucun renderer enregistré pour la section {section_id!r}")
        renderer(context)

    def render_many(
        self,
        section_ids: list[str],
        context: Mapping[str, Any],
        *,
        skip_unknown: bool = True,
    ) -> None:
        """
        Rend plusieurs sections dans l'ordre fourni.

        - `skip_unknown=True` : ignore silencieusement les sections non enregistrées.
        - `skip_unknown=False` : lève `KeyError` pour la première section inconnue.
        """
        for sid in section_ids:
            renderer = self.get(sid)
            if renderer is None:
                if skip_unknown:
                    continue
                raise KeyError(f"Aucun renderer enregistré pour la section {sid!r}")
            renderer(context)
