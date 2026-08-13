# Copyright (C) 2026 Aguirre MAURIN
#
# Ce programme est un logiciel libre : vous pouvez le redistribuer et/ou le modifier
# selon les termes de la Licence Publique Générale GNU (GPL) telle que publiée par
# la Free Software Foundation, version 3 de la licence, ou (à votre choix) toute version ultérieure.

from pathlib import Path
from core.common.chargeurs_donnees import get_source_files_metadata

def test_get_source_files_metadata_structure(tmp_path: Path) -> None:
    sources_dir = tmp_path / "data" / "sources"
    sources_dir.mkdir(parents=True, exist_ok=True)
    
    test_file = sources_dir / "Stats_PVe_OFB_test.csv"
    test_file.write_text("col1;col2\nval1;val2\n", encoding="utf-8")
    
    meta = get_source_files_metadata(tmp_path)
    assert len(meta) == 1
    assert meta[0]["nom"] == "Stats_PVe_OFB_test.csv"
    assert "taille" in meta[0]
    assert "date" in meta[0]
    assert "empreinte" in meta[0]
    assert len(meta[0]["empreinte"]) == 12
