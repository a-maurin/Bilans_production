# Copyright (C) 2026 Aguirre MAURIN
#
# Ce programme est un logiciel libre : vous pouvez le redistribuer et/ou le modifier
# selon les termes de la Licence Publique Générale GNU (GPL) telle que publiée par
# la Free Software Foundation, version 3 de la licence, ou (à votre choix) toute version ultérieure.

import sys
import pytest
from core.point_entree_cli import main


def test_cli_list_gabarits(capsys):
    test_args = ["point_entree_cli.py", "--list-gabarits"]
    sys.argv = test_args
    exit_code = main()
    assert exit_code == 0
    captured = capsys.readouterr()
    assert "srp_r27" in captured.out
    assert "Service Régional Police" in captured.out
