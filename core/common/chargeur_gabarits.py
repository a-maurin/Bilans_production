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
# Conformément à la section 7(b) DE LA GPL v3, vous devez expressément conserver
# intactes et lisibles toutes les mentions d'auteur, notices de copyright et la présente
# clause dans chaque fichier source ou interface utilisateur redistribué. Toute version modifiée
# doit clairement indiquer qu'elle a été altérée et ne doit en aucun cas supprimer le nom
# de l'auteur original (Aguirre MAURIN).

"""
========================================================================================
MODULE : CHARGEUR ET GESTIONNAIRE DE GABARITS DE PRESENTATION (`chargeur_gabarits.py`)
========================================================================================
Ce module gère la découverte, le chargement et l'attribution des gabarits personnalisés (fichiers YAML).

Fonctions assurées :
  1. Découverte des répertoires de gabarits (locaux et globaux).
  2. Chargement et validation de la structure YAML d'un gabarit (gabarit_id, libellés, cibles).
  3. Vérification de la compatibilité d'un gabarit avec un profil de bilan.
  4. Résolution automatique du gabarit approprié selon l'organisation (Région / Service).
  5. Détermination des éléments de carte à masquer (pochoirs cartographiques).
========================================================================================
"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Any
import yaml

from core.chemins_projet import PROJECT_ROOT

logger = logging.getLogger(__name__)


# ========================================================================================
# RECHERCHE ET LISTAGE DES DOSSIERS DE GABARITS
# ========================================================================================

def get_gabarits_dirs(root: Path | None = None) -> list[Path]:
    """Retourne la liste des répertoires où sont stockés les fichiers YAML de gabarits.

    Consulte dans l'ordre :
      - Le dossier du projet (`config/presentation/gabarits/`).
      - Le dossier utilisateur local (`~/.ofbilan/gabarits/`).
    """
    base_root = root or PROJECT_ROOT
    dirs = [
        base_root / "config" / "presentation" / "gabarits",
        PROJECT_ROOT / "config" / "presentation" / "gabarits",
        Path.home() / ".ofbilan" / "gabarits",
    ]
    unique_dirs: list[Path] = []
    seen: set[Path] = set()
    for d in dirs:
        try:
            resolved = d.resolve()
        except Exception:
            resolved = d
        if d.exists() and d.is_dir() and resolved not in seen:
            seen.add(resolved)
            unique_dirs.append(d)
    return unique_dirs


def load_gabarit_from_path(file_path: Path) -> dict[str, Any] | None:
    """Lit un fichier YAML de gabarit, le valide et complète ses valeurs par défaut."""
    try:
        with file_path.open("r", encoding="utf-8") as f:
            content = yaml.safe_load(f)
        if not isinstance(content, dict):
            logger.warning(f"Fichier de gabarit invalide (non-dictionnaire) : {file_path}")
            return None
        gid = content.get("gabarit_id") or file_path.stem
        content["gabarit_id"] = str(gid).strip()
        content.setdefault("label", content["gabarit_id"])
        content.setdefault("description", "")
        return content
    except Exception as e:
        logger.warning(f"Erreur lors de la lecture du gabarit {file_path} : {e}")
        return None


def list_gabarits(root: Path | None = None) -> list[dict[str, Any]]:
    """Scanne tous les dossiers et retourne la liste des gabarits disponibles avec leurs métadonnées."""
    seen: set[str] = set()
    gabarits: list[dict[str, Any]] = []

    for d in get_gabarits_dirs(root):
        for p in sorted(d.glob("*.yaml")):
            data = load_gabarit_from_path(p)
            if not data:
                continue
            gid = data["gabarit_id"]
            if gid in seen:
                continue
            seen.add(gid)
            gabarits.append({
                "gabarit_id": gid,
                "label": data.get("label", gid),
                "description": data.get("description", ""),
                "cible": data.get("cible", "les_deux"),
                "profils_compatibles": data.get("profils_compatibles"),
                "organisation": data.get("organisation", {}),
            })

    return gabarits


# ========================================================================================
# CHARGEMENT ET VERIFICATION DE COMPATIBILITE DES GABARITS
# ========================================================================================

def load_gabarit(gabarit_id: str, root: Path | None = None) -> dict[str, Any] | None:
    """Charge la configuration complète d'un gabarit spécifique par son identifiant."""
    gid_clean = str(gabarit_id).strip()
    if not gid_clean or gid_clean.lower() in ("none", "null", "standard", "default"):
        return None

    for d in get_gabarits_dirs(root):
        candidate = d / f"{gid_clean}.yaml"
        if candidate.exists():
            data = load_gabarit_from_path(candidate)
            if data and data.get("gabarit_id") == gid_clean:
                return data

        # Recherche par parcours de tous les fichiers YAML si le nom de fichier diffère
        for p in d.glob("*.yaml"):
            data = load_gabarit_from_path(p)
            if data and data.get("gabarit_id") == gid_clean:
                return data

    logger.warning(f"Gabarit de présentation introuvable : '{gabarit_id}'. Bascule sur le profil standard.")
    return None


def is_gabarit_compatible(
    gabarit: dict[str, Any],
    profile_id: str | None = None,
    cible: str = "bilan",
) -> bool:
    """Vérifie si un gabarit peut s'appliquer au type de document (bilan ou brochure) et au profil choisi."""
    g_cible = str(gabarit.get("cible", "les_deux")).strip().lower()
    if g_cible not in ("les_deux", cible.lower()):
        return False

    compat_profiles = gabarit.get("profils_compatibles")
    if compat_profiles and isinstance(compat_profiles, list):
        if profile_id and profile_id not in compat_profiles:
            return False

    return True


def resolve_gabarit_for_service(
    code_region: str | None = None,
    code_service: str | None = None,
    root: Path | None = None,
) -> str | None:
    """Détermine automatiquement le gabarit à utiliser selon la région et le service demandeur.

    1. Priorité 1 : Gabarit spécifique au service dans la région.
    2. Priorité 2 : Gabarit régional général.
    """
    reg = str(code_region or "").strip().lower()
    srv = str(code_service or "").strip().lower()
    if not reg and not srv:
        return None

    available = list_gabarits(root)
    match_region_service: str | None = None
    match_region_only: str | None = None

    for g in available:
        org = g.get("organisation", {})
        if not isinstance(org, dict):
            continue
        g_reg = str(org.get("code_region", "")).strip().lower()
        g_srv = str(org.get("service", "")).strip().lower()

        if reg and srv and g_reg == reg and g_srv == srv:
            match_region_service = g["gabarit_id"]
            break
        if reg and g_reg == reg and not g_srv:
            match_region_only = g["gabarit_id"]

    return match_region_service or match_region_only


def resolve_items_masques_carte(
    profil_data: dict[str, Any] | None = None,
    gabarit_data: dict[str, Any] | None = None,
    *,
    is_brochure: bool = False,
) -> list[str]:
    """Détermine la liste des éléments de la carte à masquer (ex: légendes ou logos spécifiques)."""
    g_carto = (gabarit_data or {}).get("cartographie", {}) if isinstance(gabarit_data, dict) else {}
    if isinstance(g_carto, dict):
        if is_brochure and "items_masques_brochure" in g_carto:
            res = g_carto.get("items_masques_brochure")
            if isinstance(res, list):
                return [str(x) for x in res]
        elif not is_brochure and "items_masques" in g_carto:
            res = g_carto.get("items_masques")
            if isinstance(res, list):
                return [str(x) for x in res]

    p_carto = (profil_data or {}).get("cartographie", {}) if isinstance(profil_data, dict) else {}
    if isinstance(p_carto, dict):
        if is_brochure and "items_masques_brochure" in p_carto:
            res = p_carto.get("items_masques_brochure")
            if isinstance(res, list):
                return [str(x) for x in res]
        res = p_carto.get("items_masques_defaut", p_carto.get("items_masques", []))
        if isinstance(res, list):
            return [str(x) for x in res]

    return []
