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
"""Module de chargement et de normalisation des profils YAML."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

_log = logging.getLogger(__name__)


def _deep_merge_dicts(base: dict, override: dict) -> dict:
    """Fusion récursive de dictionnaires (override prioritaire)."""
    merged: dict[str, Any] = dict(base)
    for key, value in override.items():
        if (
            key in merged
            and isinstance(merged[key], dict)
            and isinstance(value, dict)
        ):
            merged[key] = _deep_merge_dicts(merged[key], value)
        else:
            merged[key] = value
    return merged


def _normalize_profile(data: dict, profil_id: str) -> dict:
    """Assure la présence de toutes les clés attendues par le moteur."""
    data.setdefault("id", profil_id)
    data.setdefault("label", profil_id)
    data.setdefault("title_label", data.get("label", profil_id))
    pipeline = str(data.get("pipeline", "")).strip().lower()
    if not pipeline:
        raise ValueError(f"Profil {profil_id}: clé YAML requise manquante: pipeline")
    data["pipeline"] = pipeline
    data.setdefault("out_subdir", f"bilan_{profil_id}")
    data.setdefault("analyse_PVe", True)

    # --- filter ---
    if "filter" not in data:
        data["filter"] = {
            "type": data.pop("filter_type", "keywords"),
            "keywords": data.get("keywords", []),
            "columns": ["theme", "type_actio", "nom_dossie"],
            "exclude_patterns": [],
            "type_usager_target": [],
        }
    filt = data["filter"]
    filt.setdefault("type", "keywords")
    filt.setdefault("keywords", data.get("keywords", []))
    filt.setdefault("type_actions", [])
    filt.setdefault("columns", ["theme", "type_actio", "nom_dossie"])
    filt.setdefault("exclude_patterns", [])
    filt.setdefault("type_usager_target", [])

    # --- natinf ---
    data.setdefault("natinf_pve", [])
    data.setdefault("natinf_pej", [])
    if isinstance(data["natinf_pve"], str):
        data["natinf_pve"] = [x.strip() for x in data["natinf_pve"].split(",") if x.strip()]
    if isinstance(data["natinf_pej"], str):
        data["natinf_pej"] = [x.strip() for x in data["natinf_pej"].split(",") if x.strip()]

    return data


def load_profile_config(root: Path, profil_id: str) -> dict:
    """Charge et normalise un profil depuis config/profils_bilan/<id>.yaml."""
    try:
        import yaml
    except ImportError:
        yaml = None

    profiles_dir = root / "config" / "profils_bilan"
    defaults_path = profiles_dir / "_defaults.yaml"
    path = profiles_dir / f"{profil_id}.yaml"
    if not path.exists():
        raise FileNotFoundError(
            f"Profil introuvable: {profil_id} (attendu: {path})"
        )

    if yaml is not None:
        def yaml_include_constructor(loader, node):
            value = loader.construct_scalar(node)
            include_path = Path(loader.stream.name).parent / value
            if not include_path.exists():
                include_path = root / value
            with open(include_path, "r", encoding="utf-8") as f_inc:
                return yaml.safe_load(f_inc)

        yaml.add_constructor("!include", yaml_include_constructor, Loader=yaml.SafeLoader)

        defaults_data: dict[str, Any] = {}
        if defaults_path.exists():
            with open(defaults_path, "r", encoding="utf-8") as f:
                loaded_defaults = yaml.safe_load(f) or {}
                if isinstance(loaded_defaults, dict):
                    defaults_data = loaded_defaults
        with open(path, "r", encoding="utf-8") as f:
            loaded_profile = yaml.safe_load(f) or {}
        data = _deep_merge_dicts(defaults_data, loaded_profile if isinstance(loaded_profile, dict) else {})
    else:
        raise ImportError(
            "PyYAML est requis pour lire les profils bilan (config/profils_bilan/*.yaml)."
        ) from None

    return _normalize_profile(data, profil_id)
