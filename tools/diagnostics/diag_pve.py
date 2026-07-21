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
import pandas as pd
from pathlib import Path
from ofbilan.common.chargeurs_donnees import load_pve, load_point_ctrl
from ofbilan.engine.orchestrateur_profils import _filter_pve, load_tub_pnf_codes, _apply_restrict_geo_tub
import yaml
import logging

root = Path('.')
tub_codes, _ = load_tub_pnf_codes(root)

with open('config/profils_bilan/tub.yaml', 'r', encoding='utf-8') as f:
    tub_cfg = yaml.safe_load(f)

pve = load_pve(root, dept_code='21', date_deb='2025-07-01', date_fin='2026-06-01')
print('Total PVe après filtre période (INF-DATE-MIF):', len(pve))

pve_filtered = _filter_pve(pve, tub_cfg)
print('Total PVe après filtre NATINF:', len(pve_filtered))

pve_tub = _apply_restrict_geo_tub(pve_filtered, pd.DataFrame(), root, tub_codes, 'PVE', log=logging.getLogger())
print('Total PVe dans zone TUB (ou match contrôle):', len(pve_tub))

if not pve_tub.empty:
    for _, r in pve_tub.iterrows():
        natinf = r.get('INF-NATINF')
        date_mif = r.get('INF-DATE-MIF')
        insee = r.get('INF-INSEE')
        print(f"INSEE: {insee} | NATINF: {natinf} | DATE_MIF: {date_mif}")