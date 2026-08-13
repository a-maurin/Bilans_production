# Copyright (C) 2026 Aguirre MAURIN
#
# Ce programme est un logiciel libre : vous pouvez le redistribuer et/ou le modifier
# selon les termes de la Licence Publique Générale GNU (GPL) telle que publiée par
# la Free Software Foundation, version 3 de la licence, ou (à votre choix) toute version ultérieure.

from dataclasses import dataclass
from pathlib import Path
import pandas as pd
from core.engine.validation_audit import validate_post_generation_assertions, export_notice_sources

@dataclass
class DummyCtx:
    date_deb: pd.Timestamp = pd.Timestamp("2025-01-01")
    date_fin: pd.Timestamp = pd.Timestamp("2025-12-31")
    profile_id: str = "global"
    dept_code: str = "21"
    dept_name_typo: str = "Côte d'Or"
    diffusion: str = "interne"
    nb_localisations: int = 100
    nb_pej: int = 10
    nb_pa: int = 5
    nb_pve: int = 20
    tab_resultats: pd.DataFrame | None = None

def test_validate_post_generation_assertions_clean() -> None:
    ctx = DummyCtx()
    errors = validate_post_generation_assertions(ctx)
    assert len(errors) == 0

def test_validate_post_generation_assertions_negative() -> None:
    ctx = DummyCtx(nb_pve=-5)
    errors = validate_post_generation_assertions(ctx)
    assert len(errors) == 1
    assert "PVe est négatif" in errors[0]

def test_export_notice_sources(tmp_path: Path) -> None:
    ctx = DummyCtx()
    out_dir = tmp_path / "out"
    notice_file = export_notice_sources(out_dir, ctx, tmp_path)
    assert notice_file.exists()
    content = notice_file.read_text(encoding="utf-8")
    assert "NOTICE DE TRAÇABILITÉ DES SOURCES" in content
    assert "Profil utilisé      : global" in content
