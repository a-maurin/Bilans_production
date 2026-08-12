from pathlib import Path
import pandas as pd
import pytest

from core.common.chargeurs_donnees import (
    _compute_files_signature,
    _load_disk_cache,
    _save_disk_cache,
)


def test_disk_cache_roundtrip(tmp_path: Path):
    df_test = pd.DataFrame({"col1": [1, 2, 3], "col2": ["a", "b", "c"]})
    fake_file = tmp_path / "test_file.txt"
    fake_file.write_text("hello", encoding="utf-8")
    
    sig = _compute_files_signature([fake_file])
    assert sig is not None and len(sig) > 0
    
    # Doit retourner None si pas encore en cache
    assert _load_disk_cache(tmp_path, "test_key", sig) is None
    
    # Sauvegarder dans le cache
    _save_disk_cache(tmp_path, "test_key", sig, df_test)
    
    # Relire depuis le cache
    df_cached = _load_disk_cache(tmp_path, "test_key", sig)
    assert df_cached is not None
    assert len(df_cached) == 3
    assert list(df_cached.columns) == ["col1", "col2"]
    
    # Test d'invalidation (signature différente)
    assert _load_disk_cache(tmp_path, "test_key", "invalid_sig") is None


def test_disk_cache_corrupted(tmp_path: Path):
    fake_file = tmp_path / "test_file.txt"
    fake_file.write_text("hello", encoding="utf-8")
    sig = _compute_files_signature([fake_file])
    
    cache_dir = tmp_path / "data" / "cache"
    cache_dir.mkdir(parents=True, exist_ok=True)
    
    meta_file = cache_dir / "corrupt_key.meta.json"
    data_file = cache_dir / "corrupt_key.pkl.gz"
    
    meta_file.write_text(f'{{"files_sig": "{sig}"}}', encoding="utf-8")
    data_file.write_text("invalid_gzip_content", encoding="utf-8")
    
    # Doit retourner None et nettoyer le cache corrompu
    assert _load_disk_cache(tmp_path, "corrupt_key", sig) is None
    assert not data_file.exists()
