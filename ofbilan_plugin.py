# -*- coding: utf-8 -*-
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

"""
========================================================================================
EXTENSION QGIS - INTERFACE DU PLUGIN (`ofbilan_plugin.py`)
========================================================================================
Ce fichier implémente la classe principale `OFBilanPlugin` requise par l'architecture QGIS.

Rôles principaux :
  1. `initGui()` : enregistrement du bouton "Lancer OFBilan Explorer" dans la barre d'outils
     et le menu Extensions de QGIS.
  2. `run()` : démarrage en arrière-plan du serveur Web Python (`serveur.py`) sans bloquer
     l'interface utilisateur QGIS, puis ouverture automatique du navigateur web.
  3. `unload()` : nettoyage des boutons et arrêt propre du serveur web à la fermeture de QGIS.
========================================================================================
"""
from qgis.PyQt.QtCore import Qt
from qgis.PyQt.QtGui import QIcon
from qgis.PyQt.QtWidgets import QAction, QMessageBox
import os
import sys
import subprocess
import webbrowser
import threading
import time

class OFBilanPlugin:
    """Plugin QGIS pour lancer OFBilan."""

    def __init__(self, iface):
        self.iface = iface
        self.plugin_dir = os.path.dirname(__file__)
        self.action = None
        self.server_process = None

    def initGui(self):
        """Initialise l'interface QGIS (bouton de barre d'outils et menu)."""
        icon_path = ':/images/themes/default/mActionStart.svg' # Fallback QGIS icon
        
        # Si OFBilan possède un icône, on l'utilise
        local_icon = os.path.join(self.plugin_dir, 'icon.svg')
        if os.path.exists(local_icon):
            icon_path = local_icon
            
        icon = QIcon(icon_path)
        self.action = QAction(icon, "Lancer OFBilan Explorer", self.iface.mainWindow())
        self.action.triggered.connect(self.run)

        self.iface.addToolBarIcon(self.action)
        self.iface.addPluginToMenu("&OFBilan", self.action)

    def unload(self):
        """Nettoyage lors du déchargement du plugin."""
        if self.action:
            self.iface.removePluginMenu("&OFBilan", self.action)
            self.iface.removeToolBarIcon(self.action)
        
        # Arrêter le serveur s'il tourne encore
        if self.server_process and self.server_process.poll() is None:
            self.server_process.terminate()

    def run(self):
        """Logique exécutée au clic sur le bouton."""
        port = 8000
        try:
            from .core.parametres_utilisateur import lire_parametres
            port = int(lire_parametres().get("tech", {}).get("port_serveur", 8000))
        except Exception:
            pass

        if self.server_process and self.server_process.poll() is None:
            QMessageBox.information(self.iface.mainWindow(), "OFBilan", "Le serveur OFBilan est déjà en cours d'exécution.\nOuverture du navigateur...")
            webbrowser.open(f'http://localhost:{port}/explorer.html')
            return

        # Configuration de l'environnement pour importer 'core'
        env = os.environ.copy()
        env["PYTHONPATH"] = self.plugin_dir + os.pathsep + env.get("PYTHONPATH", "")
        
        serveur_script = os.path.join(self.plugin_dir, 'core', 'web', 'serveur.py')
        
        try:
            python_exe = sys.executable
            if os.name == 'nt' and "qgis" in python_exe.lower():
                bin_dir = os.path.dirname(python_exe)
                if os.path.exists(os.path.join(bin_dir, "python.exe")):
                    python_exe = os.path.join(bin_dir, "python.exe")
                elif os.path.exists(os.path.join(bin_dir, "python3.exe")):
                    python_exe = os.path.join(bin_dir, "python3.exe")
                    
            # Lancement en arrière-plan sans bloquer QGIS
            self.server_process = subprocess.Popen(
                [python_exe, serveur_script],
                env=env,
                cwd=self.plugin_dir,
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
            )
            
            # Ouverture immédiate de la page de chargement (qui attendra le serveur)
            loading_html = os.path.join(self.plugin_dir, 'core', 'web', 'loading.html')
            webbrowser.open(f"file:///{loading_html.replace(os.sep, '/')}?port={port}")
            
            self.iface.messageBar().pushMessage("OFBilan", "Démarrage du serveur web...", level=0, duration=3)
            
        except Exception as e:
            QMessageBox.critical(self.iface.mainWindow(), "Erreur OFBilan", f"Impossible de lancer le serveur :\n{e}")