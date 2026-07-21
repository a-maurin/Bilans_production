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
"""Tests des utilitaires du script tools/audit_pve_totaux.py (sans données sources)."""
from __future__ import annotations

import importlib.util
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[2]


@pytest.fixture(scope="module")
def audit_pve_mod():
    path = ROOT / "tools" / "audit" / "audit_pve_totaux.py"
    spec = importlib.util.spec_from_file_location("audit_pve_totaux", path)
    mod = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(mod)
    return mod


def test_list_pve_detail_csv_paths_excludes_zone_exports(tmp_path: Path, audit_pve_mod) -> None:
    (tmp_path / "pve_foo_detail.csv").write_text("h\n1\n", encoding="utf-8")
    (tmp_path / "pve_foo_par_zone.csv").write_text("h\n2\n", encoding="utf-8")
    names = [p.name for p in audit_pve_mod.list_pve_detail_csv_paths(tmp_path)]
    assert names == ["pve_foo_detail.csv"]


def test_count_csv_body_rows_utf8_with_header(tmp_path: Path, audit_pve_mod) -> None:
    p = tmp_path / "t.csv"
    p.write_text("a;b\n1;2\n3;4\n", encoding="utf-8")
    assert audit_pve_mod.count_csv_body_rows(p) == 2