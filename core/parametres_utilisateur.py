# -*- coding: utf-8 -*-
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
"""Gestion de la persistance des paramètres utilisateur du plugin OFBilan."""

import os
import json
from pathlib import Path
from typing import Dict, Any

# Valeurs par défaut globales
DEFAUT_PARAMETRES: Dict[str, Any] = {
    "profil": {
        "nom": "",
        "prenom": "",
        "service": ""
    },
    "geo": {
        "code_geo_defaut": "",
        "annee_reference": 2024,
        "gabarit_defaut": "gabarit_defaut"
    },
    "ui": {
        "vue_lancement": "explorer",
        "theme": "clair",
        "zoom_defaut": 8,
        "auto_select_gabarit": False
    },
    "carto": {
        "fond_plan": "OSM",
        "options_infobulles": {}
    },
    "export": {
        "dpi": 300,
        "inclure_donnees_brutes": False
    },
    "systeme": {
        "dossier_export": str(Path.home() / "Documents" / "OFBilan_Exports"),
        "proxy": ""
    },
    "tech": {
        "port_serveur": 5000,
        "mode_debug": False
    }
}

def get_settings_file_path() -> Path:
    """Retourne le chemin absolu vers le fichier de paramètres utilisateur."""
    # Stockage dans le profil utilisateur Windows (~/.ofbilan/user_settings.json)
    base_dir = Path.home() / ".ofbilan"
    # Création du dossier s'il n'existe pas
    base_dir.mkdir(parents=True, exist_ok=True)
    return base_dir / "user_settings.json"

def _fusion_recursif(dict_base: Dict[str, Any], dict_mise_a_jour: Dict[str, Any]) -> Dict[str, Any]:
    """Fusionne récursivement deux dictionnaires."""
    resultat = dict_base.copy()
    for cle, valeur in dict_mise_a_jour.items():
        if isinstance(valeur, dict) and isinstance(resultat.get(cle), dict):
            resultat[cle] = _fusion_recursif(resultat[cle], valeur)
        else:
            resultat[cle] = valeur
    return resultat


def lire_parametres() -> Dict[str, Any]:
    """Lit les paramètres depuis le fichier JSON. Renvoie les valeurs par défaut si absent."""
    fichier = get_settings_file_path()
    parametres = DEFAUT_PARAMETRES.copy()

    if fichier.is_file():
        try:
            with fichier.open("r", encoding="utf-8") as f:
                donnees_json = json.load(f)
                if isinstance(donnees_json, dict):
                    parametres = _fusion_recursif(parametres, donnees_json)
        except (OSError, json.JSONDecodeError) as e:
            print(f"Erreur lors de la lecture des paramètres : {e}")

    return parametres


def sauvegarder_parametres(nouveaux_parametres: Dict[str, Any]) -> None:
    """Sauvegarde les paramètres fournis dans le fichier JSON."""
    fichier = get_settings_file_path()
    parametres_actuels = lire_parametres()
    parametres_fusionnes = _fusion_recursif(parametres_actuels, nouveaux_parametres)

    try:
        fichier.parent.mkdir(parents=True, exist_ok=True)
        with fichier.open("w", encoding="utf-8") as f:
            json.dump(parametres_fusionnes, f, indent=4, ensure_ascii=False)
    except OSError as e:
        print(f"Erreur lors de la sauvegarde des paramètres : {e}")