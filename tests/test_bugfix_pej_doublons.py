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
TEST UNITAIRE : NON-DOUBLONNEMENT DES PROCEDURES PEJ (`test_bugfix_pej_doublons.py`)
========================================================================================
Ce fichier de test valide la résolution du bug de comptage en double des Procédures Enquête
Judiciaire (PEJ) lorsque des faits multiples sont enregistrés sur une même fiche.
========================================================================================
"""
import pandas as pd
import pytest

def test_load_pej_preserves_multiple_nans(monkeypatch, tmp_path):
    from core.common import chargeurs_donnees
    
    # Jeu de données simulé : 1 procédure avec ID, 2 sans ID de liaison (NaN/None)
    mock_df = pd.DataFrame({
        "DC_ID": ["123", None, pd.NA],
        "DATE_CONSTATATION": ["2023-01-01", "2023-01-02", "2023-01-03"],
        "DATE_OUVERTURE_PROCEDURE": [pd.NA, pd.NA, pd.NA],
        "RECAP_DATE_INIT_PJ": [pd.NA, pd.NA, pd.NA]
    })
    
    # Mock des fonctions de lecture de fichiers pour injecter notre DataFrame
    monkeypatch.setattr(chargeurs_donnees, "_find_latest_dated_file", lambda *args, **kwargs: tmp_path / "fake.ods")
    monkeypatch.setattr(chargeurs_donnees, "_read_spreadsheet", lambda *args, **kwargs: mock_df.copy())
    
    # Exécution du chargeur
    res = chargeurs_donnees.load_pej(tmp_path)
    
    # Vérification : on s'attend à récupérer les 3 procédures
    assert len(res) == 3, "Le dédoublonnage a supprimé par erreur les procédures sans DC_ID."