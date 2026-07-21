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
"""Package principal pour la génération des bilans."""

import warnings

# --- CONFIGURATION GLOBALE PANDAS ---
# Solution pérenne pour désactiver le backend PyArrow pour les chaînes de caractères.
# Dans l'environnement QGIS (Pandas 2.1+), PyArrow est souvent activé par défaut, mais la
# version de PyArrow fournie manque de certaines fonctionnalités regex (ex: replace_substring_regex).
# En forçant le mode "python", Pandas utilisera le moteur natif (module re) pour toutes les séries.
try:
    import pandas as pd
    
    # Pandas 2.0+
    try:
        if hasattr(pd.options.mode, "string_storage"):
            pd.options.mode.string_storage = "python"
    except Exception:
        pass
        
    # Option future pour Pandas 2.1+
    try:
        if hasattr(pd.options.future, "infer_string"):
            pd.options.future.infer_string = False
    except Exception:
        pass
except ImportError:
    pass
except Exception as e:
    warnings.warn(f"Impossible de configurer globalement le backend string de Pandas : {e}")