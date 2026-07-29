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
MODULE : CONFIGURATION DES PARAMETRES DE BILAN (`bilan_config.py`)
========================================================================================
Ce module centralise tous les paramètres d'un bilan sous forme d'une classe de données (`BilanConfig`).
Il contient :
  - La période temporelle d'analyse (date de début, date de fin).
  - Le périmètre géographique (échelle nationale/régionale/départementale et code du territoire).
  - Les méthodes de résolution des dossiers de sortie et des unités de services (Service Départemental).
========================================================================================
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import pandas as pd

from core.common.utilitaires_metier import get_perimetre_name
from core.chemins_projet import PROJECT_ROOT, get_out_dir


# ========================================================================================
# FONCTIONS DE RESOLUTION DES ARGUMENTS DE PERIMETRE
# ========================================================================================

def resolve_perimetre_kwargs(
    *,
    echelle: str | None = None,
    code: str | None = None,
    dept_code: str | None = None,
) -> tuple[str, str]:
    """Résout l'échelle géographique et le code territoire transmis en paramètres CLI ou Python.

    Assure la rétrocompatibilité si l'ancien paramètre `dept_code` est utilisé à la place d'échelle/code.
    """
    if echelle is not None and code is not None:
        return str(echelle).strip(), str(code).strip()
    if dept_code is not None:
        return "departement", str(dept_code).strip()
    return "departement", "21"


# ========================================================================================
# CLASSE PRINCIPALE DE CONFIGURATION D'UN BILAN
# ========================================================================================

@dataclass
class BilanConfig:
    """Structure de données (Dataclass) contenant tous les paramètres d'exécution d'un bilan.

    - `date_deb` : Horodatage de début de période.
    - `date_fin` : Horodatage de fin de période.
    - `echelle` : 'departement', 'region' ou 'national'.
    - `code` : Code INSEE du département (ex: '21') ou de la région (ex: 'r27').
    - `out_dir` : Dossier optionnel de destination pour l'enregistrement du rapport.
    """
    date_deb: pd.Timestamp
    date_fin: pd.Timestamp
    echelle: str
    code: str
    root: Path = field(default_factory=lambda: PROJECT_ROOT)
    out_dir: Optional[Path] = None

    @property
    def entity_sds(self) -> list[str]:
        """Retourne la liste des codes de Services Départementaux (ex: ['SD21', 'SD58']) inclus dans le périmètre."""
        from core.common.utilitaires_metier import get_departements_pour_perimetre
        codes = get_departements_pour_perimetre(self.echelle, self.code)
        if codes and "FR" not in codes:
            return [f"SD{c}" for c in codes]
        return []

    @property
    def dept_code(self) -> str:
        """Alias rétro-compatible retournant le code du territoire (ex: '21')."""
        return self.code

    @property
    def perimetre_name(self) -> str:
        """Retourne le nom complet lisible du territoire (ex: 'Côte-d\'Or' ou 'Bourgogne-Franche-Comté')."""
        return get_perimetre_name(self.echelle, self.code)

    def get_out(self, programme: str) -> Path:
        """Retourne le dossier de sortie des résultats ; crée le dossier s'il n'existe pas encore."""
        if self.out_dir is not None:
            self.out_dir.mkdir(parents=True, exist_ok=True)
            return self.out_dir
        return get_out_dir(programme)

    @classmethod
    def from_strings(
        cls,
        date_deb: str,
        date_fin: str,
        echelle: str = "departement",
        code: str = "21",
        root: Optional[Path] = None,
        out_dir: Optional[Path] = None,
    ) -> "BilanConfig":
        """Instancie la configuration à partir de simples chaînes de caractères (ex: '2026-01-01')."""
        return cls(
            date_deb=pd.to_datetime(date_deb),
            date_fin=pd.to_datetime(date_fin),
            echelle=str(echelle).strip(),
            code=str(code).strip(),
            root=root or PROJECT_ROOT,
            out_dir=out_dir,
        )