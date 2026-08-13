# Copyright (C) 2026 Aguirre MAURIN
#
# Ce programme est un logiciel libre : vous pouvez le redistribuer et/ou le modifier
# selon les termes de la Licence Publique Générale GNU (GPL) telle que publiée par
# la Free Software Foundation, version 3 de la licence, ou (à votre choix) toute version ultérieure.

from core.engine.agregations_profil import compute_n1_deltas

def test_compute_n1_deltas_augmentation() -> None:
    res = compute_n1_deltas(150, 100)
    assert res["delta_pct"] == 50.0
    assert res["delta_str"] == "+50.0%"
    assert res["alerte_baisse"] is False

def test_compute_n1_deltas_baisse_normale() -> None:
    res = compute_n1_deltas(80, 100)
    assert res["delta_pct"] == -20.0
    assert res["delta_str"] == "-20.0%"
    assert res["alerte_baisse"] is False

def test_compute_n1_deltas_baisse_anormale_alerte() -> None:
    res = compute_n1_deltas(50, 100)
    assert res["delta_pct"] == -50.0
    assert res["alerte_baisse"] is True
    assert "Alerte statistique" in res["message_alerte"]

def test_compute_n1_deltas_previous_zero() -> None:
    res = compute_n1_deltas(50, 0)
    assert res["delta_pct"] is None
    assert res["delta_str"] == "N/A"
    assert res["alerte_baisse"] is False
