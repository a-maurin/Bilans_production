![OFBilan Banner](ref/programme/logos/bandeau_ofbilan.svg)

# OFBilan (Extension QGIS)

**OFBilan** est un outil d'aide à la décision, d'exploration de données et de communication. Il s'appuie sur les données de contrôles (OSCEAN) et les procédures (PVe / PEJ / PA) de l'Office Français de la Biodiversité (OFB).

Conçu pour allier la puissance d'analyse spatiale de QGIS à la flexibilité d'une interface web moderne, **OFBilan est distribué sous la forme d'une extension QGIS**.

---

## Fonctionnalités Principales

Le programme s'articule autour de deux modules majeurs :

### 1. OFBilan Explorer (Analyse interactive)
Une interface web embarquée pour l'exploration fluide de vos données :
*   **Cartographie interactive (Leaflet)** : Visualisation géographique instantanée des points de contrôle, clustering dynamique et génération de cartes de chaleur (Heatmaps) filtrables par année.
*   **Tableaux de bord (Chart.js)** : Statistiques clés actualisées en temps réel selon l'emprise spatiale et les filtres actifs (Top 5 thématiques, répartition des suites données, etc.).
*   **Filtrage multi-critères à la volée** : Affinez instantanément vos données par période, département, BMI, type d'usager, thématique ou nature de l'infraction.
*   **Interface ergonomique** : Légende dynamique, console d'édition intégrée avec outils de copie rapide, et gestion optimisée des regroupements cartographiques.

### 2. Éditeur de Bilans PDF (Génération de rapports)
Un puissant moteur de rendu pour automatiser vos rapports d'activité :
*   **Catalogue sur mesure** : Bilans globaux, thématiques (eau, chasse, espèces, pollutions...) ou ciblés par type d'usager, paramétrables via des profils YAML.
*   **Mise en page professionnelle** : Génération de rapports détaillés ou de brochures A4 synthétiques (4 pages).
*   **Gestion de la confidentialité** : Double périmètre de diffusion (versions *Internes* détaillées vs *Externes* anonymisées pour les partenaires).
*   **Cartographie native automatisée** : Communication directe avec le moteur QGIS pour insérer automatiquement des cartes territoriales de haute qualité dans les rapports PDF.

---

## Avantages de l'intégration QGIS

*   **Portabilité totale** : Aucun environnement Python complexe à configurer. L'outil exploite directement l'interpréteur et les bibliothèques embarqués de QGIS.
*   **Déploiement simplifié** : Installation rapide depuis le gestionnaire d'extensions QGIS.
*   **Moteur cartographique robuste** : Exploitation native des capacités de rendu de QGIS pour les exports statiques.

---

## ⚙️ Installation et Configuration

La mise en place s'effectue en deux étapes.

### Étape 1 : Installation du cœur de l'extension
**Méthode A : Installation par fichier ZIP (Recommandée)**
1.  Téléchargez la version packagée au format `.zip`.
2.  Ouvrez QGIS > Menu **Extensions > Installer/Gérer les extensions**.
3.  Onglet **Installer depuis un ZIP** > Sélectionnez le fichier et cliquez sur **Installer**.

**Méthode B : Copie manuelle**
1. Copiez le dossier `OFBilan-Plugin-QGIS` dans vos extensions QGIS.
2. Chemin Windows standard : `%APPDATA%\QGIS\QGIS3\profiles\default\python\plugins\OFBilan-Plugin-QGIS`
3. Redémarrez QGIS.

### Étape 2 : Pack de configuration (Référentiels & Données)
Pour des raisons de confidentialité, les données sources (`data/sources/`) et les référentiels géographiques (`ref/programme/`) ne sont pas inclus dans le dépôt public.
1.  Contactez l'auteur pour obtenir l'archive `pack_configuration_referentiels.zip`.
2.  Placez ce ZIP et le script `installer_pack.bat` à la racine de votre dossier de plugin.
3.  Exécutez `installer_pack.bat` pour déployer automatiquement l'arborescence requise.

---

## 📖 Guide d'Utilisation

### Mode 1 : Depuis QGIS (Recommandé)
1. Activez l'extension **OFBilan**.
2. Cliquez sur l'icône OFBilan dans la barre d'outils, ou via le menu **Extensions > OFBilan > scripts > Lancer OFBilan Explorer**.
3. Un serveur local démarre de façon transparente et ouvre `http://localhost:8000/explorer.html` dans votre navigateur.
4. À la fermeture de QGIS, le serveur s'arrête automatiquement.

### Mode 2 : Autonome (Sans lancer QGIS)
1. Exécutez le script `demarrer_serveur_OFBilan.bat` situé à la racine du projet.
2. Le script localise l'interpréteur Python de QGIS et lance le serveur.
3. Accédez à l'interface via `http://localhost:8000/explorer.html`.
4. Fermez la console de commande pour arrêter le serveur.

### Mode 3 : Ligne de commande (CLI / Automatisation)
Exploitez l'API pour des traitements par lots (adapter le chemin vers QGIS) :
```bash
# Générer un bilan global
"C:\Program Files\QGIS 3.40.11\bin\python.exe" -m ofbilan --profil global --date-deb 2026-01-01 --date-fin 2026-12-31 --code 21

# Lister les paramètres disponibles
python -m ofbilan --list-themes
python -m ofbilan --list-type-usagers

# Générer un bilan anonymisé pour l'externe (ex: PNF) sans cartes
python -m ofbilan --profil pnf --code 21 --diffusion externe --no-cartes
```

---

## 📂 Structure du Projet

*   `ofbilan_plugin.py` : Point d'entrée de l'extension QGIS.
*   `core/` : Moteur de l'application (serveur FastAPI, calculs statistiques, génération PDF ReportLab).
*   `config/` : Profils YAML (paramétrage des bilans) et configuration géographique.
*   `data/` : Répertoire des données d'entrée (`sources/`) et rapports générés (`out/`).
*   `ref/` : Référentiels cartographiques, chartes graphiques et logos.

---

## ⚖️ Licence et Droits d'Auteur

Ce projet est développé par **Aguirre Maurin** (Service Départemental de la Côte-d'Or, OFB).

Le code source est distribué sous la licence **GNU General Public License v3.0 (GPLv3)**.

**Clause stricte d'attribution (Article 7(b) de la GPLv3) :**
Conformément à la licence, il est **strictement interdit** de supprimer, altérer ou masquer les mentions de droits d'auteur, les notices de licence, ou les informations d'identification de l'auteur d'origine présentes dans les fichiers sources, les interfaces utilisateurs (logos, mentions légales, console), et les documents générés. Toute modification ou redistribution du code doit clairement conserver l'attribution à l'auteur initial.

**Contact :** [aguirre.maurin@ofb.gouv.fr](mailto:aguirre.maurin@ofb.gouv.fr)
