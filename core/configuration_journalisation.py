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
import logging
import sys
from pathlib import Path


def configure_logging(console_level: int = logging.WARNING) -> None:
    """
    Configure le logging pour les scripts de bilans.

    - Le logger 'ofbilan' est réglé sur DEBUG pour propager tous les messages.
    - Le StreamHandler (console) filtre selon console_level (WARNING par défaut).
    - Sortie console : stderr
    """
    logger = logging.getLogger("ofbilan")
    
    # Si déjà configuré, on ajuste simplement le niveau console existant
    if logger.handlers:
        for h in logger.handlers:
            if not isinstance(h, logging.FileHandler):
                h.setLevel(console_level)
        logger.setLevel(logging.DEBUG)
        return

    logger.setLevel(logging.DEBUG)

    # Handler console
    sh = logging.StreamHandler(sys.stderr)
    sh.setLevel(console_level)
    formatter = logging.Formatter("%(levelname)s - %(message)s")
    sh.setFormatter(formatter)
    logger.addHandler(sh)

    # Empêcher la propagation au root logger pour éviter les doublons si basicConfig est utilisé
    logger.propagate = False


def add_file_handler(out_dir: Path) -> None:
    """
    Ajoute un FileHandler pour enregistrer tous les logs techniques (DEBUG)
    dans un fichier 'debug_run.log' situé dans out_dir.
    """
    logger = logging.getLogger("ofbilan")

    # Éviter d'ajouter plusieurs FileHandlers identiques
    for h in logger.handlers:
        if isinstance(h, logging.FileHandler):
            return

    try:
        out_dir.mkdir(parents=True, exist_ok=True)
        log_file = out_dir / "debug_run.log"
        if log_file.exists():
            try:
                log_file.unlink()
            except Exception:
                pass
        
        fh = logging.FileHandler(log_file, encoding="utf-8")
        fh.setLevel(logging.DEBUG)
        formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
        fh.setFormatter(formatter)
        logger.addHandler(fh)
    except Exception as e:
        # En cas d'erreur de création du dossier ou du fichier, on n'interrompt pas le programme
        logger.warning("Impossible de créer le fichier journal de debug dans %s : %s", out_dir, e)

