#
# Copyright (C) 2026 Aguirre MAURIN
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.

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
