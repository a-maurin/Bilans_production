# Copyright (C) 2026 Aguirre MAURIN
#
# Ce programme est un logiciel libre : vous pouvez le redistribuer et/ou le modifier
# selon les termes de la Licence Publique Générale GNU (GPL) telle que publiée par
# la Free Software Foundation, version 3 de la licence, ou (à votre choix) toute version ultérieure.

from dataclasses import dataclass
from pathlib import Path
import pandas as pd
from core.engine.sections_profil import render_sec1
from core.common.pdf_report_builder import PDFReportBuilder

@dataclass
class DummyCtxH3:
    builder: PDFReportBuilder
    section_title: dict
    presentation_cfg: dict
    nb_localisations: int = 0
    nb_ops: int = 0
    nb_pej: int = 0
    nb_pa: int = 0
    nb_pve: int = 0
    tab_resultats_controles: pd.DataFrame | None = None
    tab_resultats: pd.DataFrame | None = None

def test_null_report_rendering(tmp_path: Path) -> None:
    builder = PDFReportBuilder(tmp_path / "test_null.pdf", "Bilan Test")
    ctx = DummyCtxH3(
        builder=builder,
        section_title={"sec1": "1. Synthèse de l'activité"},
        presentation_cfg={},
    )
    render_sec1(ctx)
    assert len(builder.story) > 0
