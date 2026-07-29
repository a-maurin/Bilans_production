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


def configure_logging(console_level: int = logging.ERROR) -> None:
    """
    Configure le logging pour les scripts de bilans.

    - Les loggers 'ofbilan' et 'core' sont réglés sur DEBUG pour tout enregistrer.
    - Le StreamHandler (console) filtre selon console_level (ERROR par défaut en mode normal).
    - Sortie console : stderr
    """
    for logger_name in ("ofbilan", "core"):
        logger = logging.getLogger(logger_name)
        logger.setLevel(logging.DEBUG)
        
        # Ajustement des handlers existants
        for h in list(logger.handlers):
            if not isinstance(h, logging.FileHandler):
                h.setLevel(console_level)

        if not any(not isinstance(h, logging.FileHandler) for h in logger.handlers):
            sh = logging.StreamHandler(sys.stderr)
            sh.setLevel(console_level)
            formatter = logging.Formatter("%(levelname)s - %(message)s")
            sh.setFormatter(formatter)
            logger.addHandler(sh)

        logger.propagate = False

    # Neutraliser l'affichage console bruyant sur le root logger en mode normal
    root_logger = logging.getLogger()
    if console_level >= logging.WARNING:
        root_logger.setLevel(logging.ERROR)


def add_file_handler(out_dir: Path) -> None:
    """
    Ajoute un FileHandler pour enregistrer tous les logs techniques (DEBUG)
    dans un fichier 'debug_run.log' situé dans out_dir.
    """
    try:
        out_dir.mkdir(parents=True, exist_ok=True)
        log_file = out_dir / "debug_run.log"
        if log_file.exists():
            try:
                log_file.unlink()
            except Exception:
                pass

        for logger_name in ("ofbilan", "core"):
            logger = logging.getLogger(logger_name)
            has_fh = any(isinstance(h, logging.FileHandler) for h in logger.handlers)
            if not has_fh:
                fh = logging.FileHandler(log_file, encoding="utf-8")
                fh.setLevel(logging.DEBUG)
                formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
                fh.setFormatter(formatter)
                logger.addHandler(fh)
    except Exception as e:
        logger = logging.getLogger("ofbilan")
        logger.warning("Impossible de créer le fichier journal de debug dans %s : %s", out_dir, e)

