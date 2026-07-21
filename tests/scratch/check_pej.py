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
import sys
from pathlib import Path
from core.engine.orchestrateur_profils import load_profile_config
from core.common.chargeurs_donnees import load_pej

project_root = Path('.')
try:
    p = load_profile_config(project_root, 'ppp')
    df_pej = load_pej(project_root, echelle='national', code='', date_deb='2025-01-01', date_fin='2025-12-31')
    
    natinf_pej = p.get('natinf_pej', [])
    print("NATINF in profile:", natinf_pej[:5], "...")
    
    import re
    from core.common.utilitaires_metier import series_str_contains
    
    pattern = "|".join(rf"(?:^|_){re.escape(c)}(?:_|$)" for c in natinf_pej)
    natinf_col = "NATINF_PEJ" if "NATINF_PEJ" in df_pej.columns else "NATINF"
    print("NATINF column:", natinf_col)
    
    if natinf_col in df_pej.columns:
        res = df_pej[series_str_contains(df_pej[natinf_col], pattern, regex=True)]
        print("Filtered PEJ length:", len(res))
        if len(res) == 0:
            print("Why 0?")
            print("Sample NATINF_PEJ values:", df_pej[natinf_col].dropna().head().tolist())
    else:
        print("No natinf col found.")
except Exception as e:
    print("Error:", e)