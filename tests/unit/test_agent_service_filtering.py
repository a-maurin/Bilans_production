import pytest
import pandas as pd
from pathlib import Path
from core.engine.orchestrateur_profils import filter_by_agent_service, load_profile_config

def test_filter_by_agent_service_basic():
    df = pd.DataFrame({
        "entite_ctrl": ["SD21", "PNF - Agents", "SD52", None, "PARC NATIONAL"],
        "nb": [10, 20, 30, 40, 50]
    })
    
    cfg = {
        "agent_rules": {
            "pnf_keywords": ["PNF", "PARC"],
            "ofb_keywords": ["OFB", "SD"]
        }
    }
    
    # Mode "tous"
    df_tous = filter_by_agent_service(df, ["entite_ctrl"], "tous", cfg)
    assert len(df_tous) == 5
    
    # Mode "pnf" -> seules les entités PNF/PARC
    df_pnf = filter_by_agent_service(df, ["entite_ctrl"], "pnf", cfg)
    assert len(df_pnf) == 2
    assert set(df_pnf["nb"]) == {20, 50}
    
    # Mode "ofb" -> SD21, SD52 et NULL (fallback OFB)
    df_ofb = filter_by_agent_service(df, ["entite_ctrl"], "ofb", cfg)
    assert len(df_ofb) == 3
    assert set(df_ofb["nb"]) == {10, 30, 40}


def test_filter_by_agent_service_pve_column():
    df = pd.DataFrame({
        "UNITE_libelle": ["Brigade SD21", "Unité PNF Forets", "Agent SD52"],
        "val": [1, 2, 3]
    })
    cfg = {"agent_rules": {"pnf_keywords": ["PNF"]}}
    
    df_pnf = filter_by_agent_service(df, ["UNITE_libelle", "unite_libelle"], "pnf", cfg)
    assert len(df_pnf) == 1
    assert df_pnf.iloc[0]["val"] == 2

    df_ofb = filter_by_agent_service(df, ["UNITE_libelle", "unite_libelle"], "ofb", cfg)
    assert len(df_ofb) == 2
    assert set(df_ofb["val"]) == {1, 3}


def test_explorer_html_has_pnf_agent_select():
    html_path = Path(__file__).resolve().parents[2] / "core" / "web" / "explorer.html"
    content = html_path.read_text(encoding="utf-8")
    assert "pnf-agent-select" in content
    assert "Tous les agents" in content
    assert "OFB uniquement" in content
    assert "PNF uniquement" in content


def test_explorer_js_has_agent_service_params():
    js_path = Path(__file__).resolve().parents[2] / "core" / "web" / "explorer.js"
    content = js_path.read_text(encoding="utf-8")
    assert "pnf-agent-select" in content
    assert "agent_service" in content
