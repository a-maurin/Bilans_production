import pytest
import sys
from pathlib import Path

root = Path(r"c:\Users\aguirre.maurin\Documents\GitHub\OFBilan-Plugin-QGIS")
sys.path.insert(0, str(root))

res = pytest.main([str(root / "tests" / "unit" / "test_agent_service_filtering.py"), "-v"])
with open(root / "tests" / "scratch" / "test_output.txt", "w", encoding="utf-8") as f:
    f.write(f"Pytest exit code: {res}\n")
