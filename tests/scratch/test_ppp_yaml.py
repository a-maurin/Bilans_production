import sys
from pathlib import Path
from core.engine.orchestrateur_profils import load_profile_config

try:
    project_root = Path('.')
    p = load_profile_config(project_root, 'ppp')
    with open('test_ppp.txt', 'w', encoding='utf-8') as f:
        f.write(f"sources: {p.get('sources')}\n")
        f.write(f"natinf_pej len: {len(p.get('natinf_pej', []))}\n")
        f.write(f"natinf_pej: {p.get('natinf_pej')}\n")
except Exception as e:
    with open('test_ppp.txt', 'w', encoding='utf-8') as f:
        f.write(f"Error: {e}\n")
