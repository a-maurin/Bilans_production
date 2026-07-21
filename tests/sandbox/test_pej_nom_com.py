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
sys.path.insert(0, 'src')
import pandas as pd
from core.common.chargeurs_donnees import load_pej, load_point_ctrl, merge_pej_faits_locations
from core.engine.agregations_profil import _build_global_proc_detail

root = Path('.')
point = load_point_ctrl(root)
pej = load_pej(root)

print("Total PEJ rows raw:", len(pej))
# Simulate orchestrateur merging
pej_dept = merge_pej_faits_locations(pej, root, dept_code="21")

pej_detail_old = _build_global_proc_detail(
    pej_dept, "PEJ", ["DC_ID"], ["DATE_REF"], ["COMMUNE", "nom_commune", "INF-LIEU", "INF-INSEE"], ["THEME"], ["DOMAINE"]
)

pej_detail_new = _build_global_proc_detail(
    pej_dept, "PEJ", ["DC_ID"], ["DATE_REF"], ["COMMUNE", "NOM_COM", "nom_commune", "INF-LIEU", "INF-INSEE"], ["THEME"], ["DOMAINE"]
)

mask_old = pej_detail_old["commune"].isin(["n.d.", "INC", "ND"])
mask_new = pej_detail_new["commune"].isin(["n.d.", "INC", "ND"])

print("Without NOM_COM, missing communes:", mask_old.sum())
print("With NOM_COM, missing communes:", mask_new.sum())