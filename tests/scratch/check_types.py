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
import os
from pathlib import Path
import pandas as pd
from core.common.chargeurs_donnees import load_point_ctrl, load_pej

project_root = Path(os.getcwd())
df_pts = load_point_ctrl(project_root, echelle="national", code="FR")
print("Unique type_actio in point_ctrl:")
print(df_pts["type_actio"].dropna().unique())

df_pej = load_pej(project_root, echelle="national", code="FR")
print("\nUnique type_action in PEJ:")
if "TYPE_ACTION" in df_pej.columns:
    print(df_pej["TYPE_ACTION"].dropna().unique())
elif "type_action" in df_pej.columns:
    print(df_pej["type_action"].dropna().unique())