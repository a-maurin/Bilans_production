# Paramétrage des cartes (générateur de cartes)

Ce dossier contient les fichiers de configuration optionnels pour la production cartographique.
S’ils sont présents, ils sont chargés **en priorité** sur `config_cartes.py` (rétrocompatibilité).

## Fichiers

- **profils_cartes.yaml** — Définition des profils de cartes (agrainage, chasse, piégeage, etc.) et des couches à afficher pour chaque type de carte. Obligatoire pour activer le paramétrage par fichiers.
- **symbologies.yaml** — (Optionnel) Symbologies nommées réutilisables. Les couches dans `profils_cartes.yaml` peuvent y faire référence via `symbology_ref: nom_symbologie`.

## Dépendance

Le chargement YAML nécessite **PyYAML** (`pip install pyyaml`). Si PyYAML est absent, le script utilise uniquement `config_cartes.py`.

## Structure minimale

**profils_cartes.yaml** :
```yaml
profiles:
  mon_profil:
    id: mon_profil
    title: "Titre de la carte"
    layout_name: "Nom du layout dans le projet QGIS"
    output_filename: carte_mon_profil.png
    date_deb: "2025-01-01"
    date_fin: "2026-02-05"
    layers:
      nom_couche_qgis:
        layer_name: nom_couche_qgis
        legend_label: "Libellé en légende"
        filter_type: ""   # ou point_ctrl_agrainage, point_ctrl_chasse, etc.
```

Pour réutiliser une symbologie définie dans **symbologies.yaml** :
```yaml
      ma_couche:
        symbology_ref: points_agrainage
        layer_name: point_ctrl_20260205_wgs84
```

Les clés possibles pour une couche (ou une entrée de symbologies) sont : `layer_name`, `legend_label`, `filter_type`, `geometry_mode`, `renderer_type`, `field`, `classification_mode`, `num_classes`, `palette`, `symbol_size_mm`, `symbol_shape`, `visible`.
