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
from core.common.chargeurs_donnees import load_pve, load_pej
from core.engine.agregations_profil import _build_global_proc_detail

root = Path('.')
pve = load_pve(root)
pej = load_pej(root)

print(pve[['INF-DATE-INTG', 'INF-DATE-MIF']].head())

pve_detail = _build_global_proc_detail(
    pve, 'PVe', ['INF-ID'], 
    ['INF-DATE', 'INF-DATE-INTG', 'INF-DATE-MIF', 'INF-DATE-I', 'INF_DATE', 'DATE_FAITS'], 
    ['COMMUNE_LIB', 'INF-LIEU', 'COMMUNE', 'NOM_COM', 'INF-INSEE', 'INSEE_DEP'], 
    ['INF-NATINF'], ['DOMAINE']
)

print("PVE DETAIL:")
print(pve_detail[['date', 'commune']].head())

pej_detail = _build_global_proc_detail(
    pej, 'PEJ', ['DC_ID'], 
    ['DATE_REF'], 
    ['COMMUNE', 'nom_commune', 'INF-LIEU', 'INF-INSEE'], 
    ['THEME'], ['DOMAINE']
)
print("PEJ DETAIL:")
print(pej_detail[['date', 'commune']].head())