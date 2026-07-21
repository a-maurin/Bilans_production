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
"""Utilitaires transverses pour les moteurs PDF (global, profil, etc.)."""

import pandas as pd
import yaml
from pathlib import Path
from core.chemins_projet import PROJECT_ROOT
from core.common.percent_format import format_pct_int_from_rate

def truncate_with_dash(value: str, max_len: int) -> str:
    txt = str(value or "")
    if len(txt) <= max_len:
        return txt
    if max_len <= 1:
        return "-"
    return txt[: max_len - 1].rstrip() + "-"

def nb_non_conformes_brut(tab_resultats: pd.DataFrame | None) -> int:
    """Somme Infraction + Manquement (aligné OSCEAN / bilan thématique)."""
    if tab_resultats is None or tab_resultats.empty:
        return 0
    m = tab_resultats["resultat"].astype(str).str.strip()
    return int(tab_resultats.loc[m.isin(["Infraction", "Manquement"]), "nb"].sum())

def pct_table_cell(n: int | float, denom: float) -> str:
    if denom is None or denom <= 0:
        return "n.d."
    return format_pct_int_from_rate(float(n) / float(denom))

def get_region_name_for_footer(echelle: str, code: str) -> str | None:
    """Récupère le nom de la Direction régionale depuis annuaire_ofb.yaml."""
    if echelle != "region" or not code:
        return None
    # Enlever le 'r' éventuel (ex: 'r27' -> '27')
    code_num = code.lstrip('r')
    yaml_path = PROJECT_ROOT / "config" / "annuaire_ofb.yaml"
    if yaml_path.exists():
        with open(yaml_path, "r", encoding="utf-8") as f:
            annuaire = yaml.safe_load(f) or {}
            regions = annuaire.get("regions", {})
            region_info = regions.get(code_num, {})
            nom = region_info.get("nom")
            if nom:
                return f"Office français de la biodiversité – {nom}"
    return None