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

from reportlab.pdfbase import pdfmetrics

from core.common.ofb_charte import FONT_FAMILY
from core.common.pdf_utils import truncate_text_to_width


def test_truncate_text_to_width_fits_one_line() -> None:
    long_label = (
        "27745 – NON RESPECT DES PRESCRIPTIONS DU SCHEMA DEPARTEMENTAL DE GESTION "
        "DES POPULATIONS DE GRAND GIBIER"
    )
    width_pt = 320.0
    out = truncate_text_to_width(long_label, width_pt)
    assert "\n" not in out
    assert pdfmetrics.stringWidth(out, FONT_FAMILY, 9.0) <= width_pt - 8.0
    assert out.endswith("…")
