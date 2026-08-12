"""
Module de vérification et d'installation douce des accélérateurs Python (pyogrio, python-calamine).
"""

import sys
import subprocess
import logging
import importlib.util
from typing import Callable, Optional

logger = logging.getLogger("OFBilan.Dependances")


def verifier_et_installer_accelerateurs(log_callback: Optional[Callable[[str], None]] = None) -> None:
    """
    Vérifie la présence de pyogrio et python-calamine.
    Tente une installation silencieuse si manquants, sans jamais bloquer l'application.
    """
    targets = [
        ("pyogrio", "pyogrio"),
        ("calamine", "python-calamine"),
    ]

    for mod_name, pkg_name in targets:
        if importlib.util.find_spec(mod_name) is None:
            msg = f"  [INFO] Accélérateur '{pkg_name}' manquant. Tentative d'installation..."
            logger.info(msg)
            if log_callback:
                log_callback(msg)
            
            try:
                subprocess.run(
                    [sys.executable, "-m", "pip", "install", "--quiet", pkg_name],
                    check=False,
                    timeout=30,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                )
                if importlib.util.find_spec(mod_name) is not None:
                    msg_ok = f"  [OK] Accélérateur '{pkg_name}' installé avec succès !"
                    logger.info(msg_ok)
                    if log_callback:
                        log_callback(msg_ok)
                else:
                    logger.warning("  [WARN] Échec de l'installation de %s (droits ou réseau).", pkg_name)
            except Exception as e:
                logger.warning("  [WARN] Impossible d'installer %s : %s", pkg_name, e)
