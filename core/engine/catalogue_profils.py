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

#
"""
Liste et résolution des profils bilans.

Source de vérité : fichiers YAML dans config/profils_bilan/ et ref_themes_ctrl.
"""
from __future__ import annotations

from core.common.chargeurs_donnees import load_ref_themes_ctrl
from core.chemins_projet import PROJECT_ROOT

_HIDDEN_PROFILES: frozenset[str] = frozenset({"pnf_foret", "_defaults", "types_usager", "synthese_activite_PA_PJ"})


def list_profiles() -> list[str]:
    """
    Identifiants de profils disponibles, avec ordre console :
    chasse, agrainage, types_usager, types_usager_cible, puis alphabétique, hors_theme en dernier.
    """
    profils_dir = PROJECT_ROOT / "config" / "profils_bilan"
    id_to_label: dict[str, str] = {}

    themes = load_ref_themes_ctrl(PROJECT_ROOT)
    if themes:
        for t in themes:
            pid = str(t.get("id", "")).strip()
            if not pid or pid in _HIDDEN_PROFILES:
                continue
            label = str(t.get("label", pid)).strip() or pid
            id_to_label[pid] = label
    if profils_dir.exists():
        for p in profils_dir.glob("*.yaml"):
            pid = p.stem
            if pid in _HIDDEN_PROFILES:
                continue
            id_to_label.setdefault(pid, pid)

    if not id_to_label:
        return []

    types_usager_cible_id = "types_usager_cible"
    if types_usager_cible_id not in id_to_label:
        yaml_path = profils_dir / f"{types_usager_cible_id}.yaml"
        if yaml_path.exists() and types_usager_cible_id not in _HIDDEN_PROFILES:
            id_to_label[types_usager_cible_id] = "Types d'usagers – ciblé"

    priority_order: dict[str, int] = {
        "global": 0,
        "synthese_activite_PA_PJ": 1,
        "agrainage": 2,
        "tub": 3,
        "controles_secheresse": 4,
        "pnf": 5,
    }

    def _sort_key(pid: str) -> tuple[int, str]:
        if pid == "hors_theme":
            return (1000, "")
        base_rank = priority_order.get(pid, 10)
        label = id_to_label.get(pid, pid)
        return (base_rank, label.lower())

    all_ids = list(id_to_label.keys())
    all_ids.sort(key=_sort_key)
    return all_ids


def resolve_profile_ids(raw_ids: list[str]) -> list[str]:
    """Résout les numéros (1, 2, …) en identifiants selon list_profiles()."""
    themes = list_profiles()
    if not themes:
        return raw_ids
    resolved: list[str] = []
    for p in raw_ids:
        p = str(p).strip()
        if not p:
            continue
        if p.isdigit():
            n = int(p)
            if 1 <= n <= len(themes):
                resolved.append(themes[n - 1])
            else:
                resolved.append(p)
        else:
            resolved.append(p)
    return resolved