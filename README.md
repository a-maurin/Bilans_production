# Bilans de production 2025–2026

Projet de génération des bilans d'activité de contrôle pour le SD Côte-d'Or, à partir des données OSCEAN.

## Architecture

**Point d'entrée unique** : tous les bilans passent par le script **`scripts/run_bilan.py`**, qui adapte son comportement au choix (global ou un/plusieurs thèmes).

**Moteur thématique unifié** : tous les bilans thématiques (agrainage, chasse, piégeage, types d'usagers, procédures, etc.) sont générés par un **seul moteur** (`scripts/bilan_thematique/bilan_thematique_engine.py`), piloté par les profils YAML dans `ref/profils_bilan/`.

## Points d'entrée utilisateur (trois batch uniquement)

À la racine du projet, **trois** fichiers batch permettent toute l'exécution :

| Batch | Rôle |
|-------|------|
| **lancer_bilans.bat** | Générer un bilan global (choix 1) ou un ou plusieurs bilans thématiques (choix 2 : liste exhaustive des profils). |
| **generer_cartes.bat** | Générer une ou plusieurs cartes pour une période et un département. |
| **parametrer_cartes.bat** | Ouvrir la fenêtre de configuration des profils de cartes (couches, symbologie). |

## Structure du projet

```
Bilans_production/
├── README.md
├── OBJECTIFS_BILANS.md
├── .gitignore                         # Exclut données sources + sorties générées
├── requirements.txt / pyproject.toml  # Dépendances Python
├── bilans/                            # Package Python principal
│   ├── __init__.py
│   ├── cli.py                         # Point d'entrée CLI (`python -m bilans`)
│   ├── common/                        # Modules partagés
│   ├── bilan_global/                  # Bilan global (objectif 1)
│   └── bilan_thematique/              # Moteur bilans thématiques (objectif 2)
├── scripts/
│   ├── run_bilan.py                   # Point d'entrée unique (--mode global | thematique)
│   ├── paths.py
│   ├── generateur_de_cartes/          # Génération de cartes QGIS
│   ├── generer_cartes.py              # Wrapper Python vers le générateur de cartes
│   └── windows/                       # Wrappers batch Windows
│       ├── lancer_bilans.bat
│       ├── generer_cartes.bat
│       └── parametrer_cartes.bat
├── config/                            # Configurations versionnées
│   ├── README.md
│   ├── profils_bilan/                 # Profils YAML (agrainage, chasse, piégeage, etc.)
│   ├── cartes/                        # Profils de cartes / symbologies
│   └── ref_themes_ctrl.csv            # Liste ordonnée des thèmes
├── ref/                               # Référentiels (SIG, modèle OFB, glossaire…)
├── sources/                           # Données sources locales (non versionnées)
├── out/                               # Sorties générées (non versionnées)
├── docs/                              # Documentation détaillée
└── legacy/                            # Ancien code conservé à titre d’archive
```

## Profils YAML (schema v2)

Chaque profil dans `ref/profils_bilan/` décrit un bilan thématique :

```yaml
schema_version: 2
id: chasse
label: "Bilan chasse"
out_subdir: bilan_chasse

filter:
  type: chasse              # chasse | agrainage | keywords | type_usager | procedures | all
  keywords: [...]
  columns: [theme, type_actio, nom_dossie]
  exclude_patterns: []
  type_usager_target: []

natinf_pve: [...]
natinf_pej: [...]

sources:
  point_ctrl: true
  pej: true
  pa: true
  pve: true

options:
  pnf:
    label: "Analyse PNF / hors PNF"
    default: true
    ask: true                # question interactive avant le lancement
  tub:
    label: "Analyse zones TUB"
    default: false
    ask: true
  cartes:
    label: "Intégration des cartes"
    default: true
```

### Options configurables

Les options sont des paramètres modifiables par l'utilisateur (console interactive ou CLI). Pour ajouter une nouvelle option :
1. Ajouter la clé dans le bloc `options` du profil YAML.
2. Ajouter le traitement correspondant dans le moteur (`bilan_thematique_engine.py`).

Options actuelles :
- **pnf** : découpage PNF / hors PNF
- **tub** : découpage zones TUB (tuberculose bovine)
- **cartes** : intégration des cartes générées par QGIS
- **par_commune** : indicateurs par commune
- **synthese_croisee** : synthèse croisée Ctrl × PVe × PEJ par zone

### Surcharge CLI

```batch
REM Activer PNF, désactiver TUB
python scripts/run_bilan.py --mode thematique --profil chasse --with-pnf --no-tub

REM Option générique
python scripts/bilan_thematique/run_bilan_thematique.py --profil agrainage --option synthese_croisee=true
```

### Profil `types_usager_cible` (sélection interactive)

Le profil `types_usager_cible` permet de générer un bilan ciblé sur un ou plusieurs types d’usagers :

- **Mode interactif** (par défaut) :

  ```batch
  python scripts\run_bilan.py --mode thematique --profil types_usager_cible
  ```

  Le moteur affiche la liste des types d’usagers (d’après `ref/types_usagers.csv`) et vous invite
  à saisir un ou plusieurs numéros (ex. `1,3,5`) ou `*` pour tous les types.

- **Mode non interactif** (pilotage par la CLI) :

  ```batch
  python scripts\bilan_thematique\run_bilan_thematique.py --profil types_usager_cible ^
      --date-deb 2025-01-01 --date-fin 2025-12-31 --dept-code 21 ^
      --option type_usager_target="Agriculteur et autres acteurs agricoles" ^
      --option type_usager_target="Collectivité"
  ```

  Dans ce cas, aucune question n’est posée et seuls les types explicitement fournis sont analysés.

#### Cas particulier : bilan ciblé sur un seul type d’usager

- Lorsque **un seul** type d’usager est ciblé (ex. uniquement les agriculteurs) :
  - le bandeau de chiffres clés indique toujours le **nombre de localisations de contrôle**, mais l’effectif s’intitule explicitement
    *« Effectifs – &lt;type ciblé&gt; »* (par ex. « Effectifs – Agriculteur et autres acteurs agricoles ») ;
  - la section \"Contrôles par type d’usager\" :
    - ne comporte plus de camembert (inutile à 100 %),
    - présente un tableau résumé spécifique au type ciblé,
    - supprime la colonne `type_usager` des tableaux lorsqu’elle serait identique sur toutes les lignes
      (les tableaux deviennent alors simplement \"Répartition par domaine\", \"Résultats par domaine\", etc.).
- Lorsque **plusieurs** types d’usagers sont sélectionnés, l’architecture complète (camembert + colonnes `type_usager`) est conservée.

## Prérequis

- **Python 3.10+** (pandas, geopandas, matplotlib, reportlab, Pillow)
- **QGIS** avec Python intégré (pour le générateur de cartes uniquement)

## Ordre d'exécution recommandé

1. **Générateur de cartes** — produit les PNG dans `out/generateur_de_cartes/`
2. **Bilans** — consomment les cartes et produisent les PDF

## Exécution

### Bilans (global ou thématiques)

```batch
lancer_bilans.bat
```

En ligne de commande :

```batch
REM Liste des thèmes disponibles
python scripts\run_bilan.py --list-themes

REM Bilan global
python scripts\run_bilan.py --mode global --date-deb 2025-01-01 --date-fin 2026-02-05 --dept-code 21

REM Bilans thématiques (un ou plusieurs profils)
python scripts\run_bilan.py --mode thematique --profil agrainage --profil chasse --date-deb 2025-01-01 --date-fin 2026-02-05 --dept-code 21

REM Bilan thématique avec options
python scripts\run_bilan.py --mode thematique --profil chasse --with-pnf --with-tub --date-deb 2025-09-01 --date-fin 2026-03-01 --dept-code 21
```

### Génération de cartes

```batch
generer_cartes.bat
```

### Configuration des profils de cartes

```batch
parametrer_cartes.bat
```

## Sorties

| Programme / objectif | Fichiers principaux |
|---------------------|---------------------|
| **Bilan global** (objectif 1) | `out/bilan_global/` (PDF, CSV par domaine/thème) |
| **Bilan thématique** (objectif 2) | `out/bilan_<profil>/` (PDF, CSV filtré, graphiques) |
| **Cartes** (generer_cartes.bat) | `out/generateur_de_cartes/carte_*.png` |
