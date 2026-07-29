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

"""
========================================================================================
MODULE : REGISTRE DES SECTIONS PDF (`registre_sections_pdf.py`)
========================================================================================
Ce module implémente le pattern d'architecture Registre (`SectionRegistry`) pour le rendu
des différentes sections de rapports PDF.

Objectifs :
  1. Permettre aux différents générateurs d'enregistrer dynamiquement leurs fonctions de rendu
     associées à chaque identifiant de section (ex: `sec1`, `sec2`, `sec4`).
  2. Offrir une exécution modulaire (`render()` et `render_many()`) évitant la multiplication
     des conditions `if/else` monolithiques.
========================================================================================
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