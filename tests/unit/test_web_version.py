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

import sys
from pathlib import Path

# Add src/ to sys.path
sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "core"))

from core.web.serveur import get_latest_version

def test_get_latest_version():
    version = get_latest_version()
    # It should return a string starting with 'v' and containing version numbers
    assert isinstance(version, str)
    assert version.startswith("v")
    # Clean check that it has digits and dots
    parts = version[1:].split(".")
    assert len(parts) >= 2
    for p in parts:
        assert p.isdigit()
