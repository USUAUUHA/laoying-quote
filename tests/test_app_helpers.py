from __future__ import annotations

import json

import pandas as pd

from core.app import load_saved_price_params, save_price_params
from core.calculator import needs_migration
from core.config import get_default_price_params


def test_load_saved_price_params_returns_defaults_when_file_missing(tmp_path):
    result = load_saved_price_params(tmp_path / "missing.json")

    assert result == get_default_price_params()


def test_save_and_load_price_params_roundtrip_with_sanitization(tmp_path):
    path = tmp_path / "price_params.json"
    raw = get_default_price_params()
    raw["cutting_rate"] = -1
    raw["tap_al"] = 1.23

    save_price_params(raw, path)
    loaded = load_saved_price_params(path)
    on_disk = json.loads(path.read_text(encoding="utf-8"))

    assert loaded["cutting_rate"] == get_default_price_params()["cutting_rate"]
    assert loaded["tap_al"] == 1.23
    assert on_disk["tap_al"] == 1.23


def test_needs_migration_detects_missing_columns():
    assert needs_migration(pd.DataFrame([{"序号": "1"}])) is True
