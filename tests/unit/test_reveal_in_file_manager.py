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
from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock

import pytest


def test_reveal_path_skips_when_ci(monkeypatch, tmp_path: Path) -> None:
    import core.common.reveal_in_file_manager as mod

    monkeypatch.setenv("CI", "true")
    spy = MagicMock()
    monkeypatch.setattr(mod.os, "startfile", spy, raising=False)
    d = tmp_path / "out"
    d.mkdir()
    mod.reveal_path_in_file_manager(d)
    spy.assert_not_called()


def test_reveal_path_skips_when_bilans_open_output_dir_off(monkeypatch, tmp_path: Path) -> None:
    import core.common.reveal_in_file_manager as mod

    monkeypatch.delenv("CI", raising=False)
    monkeypatch.setenv("BILANS_OPEN_OUTPUT_DIR", "0")
    spy = MagicMock()
    monkeypatch.setattr(mod.os, "startfile", spy, raising=False)
    d = tmp_path / "out"
    d.mkdir()
    mod.reveal_path_in_file_manager(d)
    spy.assert_not_called()


@pytest.mark.parametrize("val", ("FALSE", "no", "OFF"))
def test_reveal_path_skips_variant_flags(monkeypatch, tmp_path: Path, val: str) -> None:
    import core.common.reveal_in_file_manager as mod

    monkeypatch.delenv("CI", raising=False)
    monkeypatch.setenv("BILANS_OPEN_OUTPUT_DIR", val)
    spy = MagicMock()
    monkeypatch.setattr(mod.os, "startfile", spy, raising=False)
    d = tmp_path / "out"
    d.mkdir()
    mod.reveal_path_in_file_manager(d)
    spy.assert_not_called()


def test_reveal_path_windows_calls_startfile(monkeypatch, tmp_path: Path) -> None:
    import core.common.reveal_in_file_manager as mod

    monkeypatch.delenv("CI", raising=False)
    monkeypatch.delenv("BILANS_OPEN_OUTPUT_DIR", raising=False)
    monkeypatch.setattr(mod.sys, "platform", "win32")
    spy = MagicMock()
    monkeypatch.setattr(mod.os, "startfile", spy, raising=False)
    d = tmp_path / "out"
    d.mkdir()
    mod.reveal_path_in_file_manager(d)
    spy.assert_called_once()
    assert spy.call_args[0][0] == d.resolve()


def test_reveal_path_file_calls_startfile(monkeypatch, tmp_path: Path) -> None:
    import core.common.reveal_in_file_manager as mod

    monkeypatch.delenv("CI", raising=False)
    monkeypatch.delenv("BILANS_OPEN_OUTPUT_DIR", raising=False)
    monkeypatch.setattr(mod.sys, "platform", "win32")
    spy = MagicMock()
    monkeypatch.setattr(mod.os, "startfile", spy, raising=False)
    f = tmp_path / "file.txt"
    f.write_text("x", encoding="utf-8")
    mod.reveal_path_in_file_manager(f)
    spy.assert_called_once()
    assert spy.call_args[0][0] == f.resolve()