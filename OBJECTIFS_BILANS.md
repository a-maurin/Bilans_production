# Objectifs des bilans

Ce document décrit les deux objectifs des scripts de bilans et le point d'entrée unique associé.

## Point d'entrée unique

**Un seul script** assure la génération des bilans : **`scripts/run_bilan.py`**.  
Il est **dynamique** selon le choix de l'utilisateur :
- **Mode global** : bilan d'activité du service (tous domaines/thèmes, PA, PEJ, PVe).
- **Mode thématique** : un ou plusieurs bilans ciblés (agrainage, chasse, piégeage, types d'usagers, procédures, etc.).

Le batch **lancer_bilans.bat** appelle uniquement ce script (choix 1 → `--mode global`, choix 2 → `--mode thematique` + liste de `--profil`).  
Les scripts `scripts/bilan_global/analyse_global.py` et `scripts/bilan_thematique/run_bilan_thematique.py` restent utilisables en direct pour compatibilité.

---

## Objectif 1 : Bilan global

**Description** : Produire un **bilan global** de l'activité du service sur une période donnée (tous domaines/thèmes, contrôles, PA, PEJ, PVe).

**Paramètres** :
- **Période** : date de début et date de fin (format YYYY-MM-DD)
- **Département** : code département (ex. 21)

**Point d'entrée** :
- Script unique : `scripts/run_bilan.py --mode global` (recommandé)
- Ou directement : `scripts/bilan_global/analyse_global.py`
- Batch : **lancer_bilans.bat** (choix 1)

**Sorties** : `out/bilan_global/` (PDF, CSV par domaine/thème).

---

## Objectif 2 : Bilan thématique (adaptable)

**Description** : Produire un **bilan sur un domaine, un thème ou tout autre objet** choisi par l'utilisateur sur une période donnée (ex. bilan agrainage, chasse, professions agricoles). Le programme s'adapte via le paramètre « profil » (thème).

**Paramètres** :
- **Période** : date de début et date de fin (YYYY-MM-DD)
- **Département** : code département
- **Sujet** : un ou plusieurs profils ; la liste est définie dans **`ref/ref_themes_ctrl.csv`** et affichée par `scripts/run_bilan.py --list-themes` (ou par le batch). Elle inclut notamment :
  - les thèmes de contrôles (ex. `agrainage`, `chasse`, `piegeage`) ;
  - le bilan **types d'usagers** (`types_usager`) ;
  - le bilan **procédures judiciaires et PVe** (`procedures_pve`).
  - Possibilité de générer un rapport par profil ou un **bilan combiné** (option `--combine`).

**Point d'entrée** :
- Script unique : `scripts/run_bilan.py --mode thematique --profil <id> [--profil <id> ...]` (recommandé)
- Ou directement : `scripts/bilan_thematique/run_bilan_thematique.py` (avec `--profil`, option `--combine`)
- Batch : **lancer_bilans.bat** (choix 2)
- La liste officielle des thèmes est dans **`ref/ref_themes_ctrl.csv`** ; le détail de chaque profil (filtres, NATINF, script) est dans `ref/profils_bilan/` (fichiers YAML).

**Sorties** : `out/bilan_<profil>/` ou un rapport combiné selon le choix de l'utilisateur.

---

## Résumé

| Objectif              | Point d'entrée recommandé              | Rôle |
|-----------------------|----------------------------------------|------|
| **1. Bilan global**   | **lancer_bilans.bat** (choix 1) / `scripts/run_bilan.py --mode global` | Bilan d'activité du service (tous domaines/thèmes) sur une période |
| **2. Bilan thématique** | **lancer_bilans.bat** (choix 2) / `scripts/run_bilan.py --mode thematique --profil <id> [...]` | Bilan ciblé sur un ou plusieurs thèmes (agrainage, chasse, etc.) sur une période |
