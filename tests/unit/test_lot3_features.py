# Copyright (C) 2026 Aguirre MAURIN
#
# Ce programme est un logiciel libre : vous pouvez le redistribuer et/ou le modifier
# selon les termes de la Licence Publique Générale GNU (GPL) telle que publiée par
# la Free Software Foundation, version 3 de la licence, ou (à votre choix) toute version ultérieure.

from pathlib import Path
from core.engine.orchestration.loader import load_profile_config
from core.common.pdf_presentation_config import REALISATION

def test_profile_versioning_defaults(tmp_path: Path) -> None:
    root = Path(__file__).resolve().parents[2]
    profile = load_profile_config(root, "chasse")
    assert "version" in profile
    assert "date_modification" in profile
    assert profile["version"] == "1.0.0"

def test_realisation_string_format() -> None:
    assert "Réalisation :" in REALISATION
    assert "Aguirre MAURIN" in REALISATION
