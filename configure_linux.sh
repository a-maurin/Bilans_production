#!/bin/bash

# Script de configuration pour Linux
# Installe les dépendances nécessaires pour le programme Bilan_production

echo "Configuration de l'environnement pour Bilan_production sous Linux"
echo "=================================================================="

# Mettre à jour les paquets
 echo "Mise à jour des paquets..."
sudo apt update

# Installer Python et les bibliothèques de base
echo "Installation de Python et des bibliothèques de base..."
sudo apt install -y python3 python3-pip python3-venv

# Installer les bibliothèques Python nécessaires
echo "Installation des bibliothèques Python nécessaires..."
pip3 install --user pandas geopandas reportlab Pillow matplotlib numpy shapely fiona pyyaml odfpy

# Installer QGIS et les bibliothèques associées
echo "Installation de QGIS et des bibliothèques associées..."
sudo apt install -y qgis python3-qgis

# Configurer les variables d'environnement
echo "Configuration des variables d'environnement..."
echo "export PATH=$PATH:~/.local/bin" >> ~/.bashrc
echo "export PYTHONPATH=$PYTHONPATH:/usr/share/qgis/python" >> ~/.bashrc

# Configurer les permissions pour les scripts shell
echo "Configuration des permissions pour les scripts shell..."
chmod +x lancer_bilans.sh
echo "chmod +x generer_cartes.sh"
echo "chmod +x parametrer_cartes.sh"

echo ""
echo "Configuration terminée."
echo "Veuillez redémarrer votre session ou exécuter 'source ~/.bashrc' pour appliquer les changements."
echo "Vous pouvez maintenant lancer les scripts avec :"
echo "  ./lancer_bilans.sh"
echo "  ./generer_cartes.sh"
echo "  ./parametrer_cartes.sh"

exit 0