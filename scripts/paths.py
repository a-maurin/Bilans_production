"""Chemins du projet — partagés par tous les programmes."""
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent


def get_out_dir(programme: str) -> Path:
    """Dossier de sortie du programme (out/bilan_agrainage, out/bilan_chasse, out/generateur_de_cartes)."""
    d = PROJECT_ROOT / "out" / programme
    d.mkdir(parents=True, exist_ok=True)
    return d


def get_cartes_dir() -> Path:
    """Dossier des cartes générées (pour les bilans qui les intègrent)."""
    return PROJECT_ROOT / "out" / "generateur_de_cartes"


def get_sources_dir() -> Path:
    """Dossier des sources (CSV OSCEAN, etc.)."""
    return PROJECT_ROOT / "sources"


def get_ref_dir() -> Path:
    """Dossier des références (modele_ofb, etc.)."""
    return PROJECT_ROOT / "ref"


def get_config_dir() -> Path:
    """Dossier des configurations versionnées (profils YAML, configs cartes, etc.)."""
    return PROJECT_ROOT / "config"


def get_sig_dir() -> Path:
    """Dossier des données SIG (pve_communes.shp, etc.)."""
    return PROJECT_ROOT / "ref" / "sig"


def get_sources_sig_dir() -> Path:
    """Dossier des données SIG construites à partir des sources (PVe, points de contrôle, etc.)."""
    return PROJECT_ROOT / "sources" / "sig"
