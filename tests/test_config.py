from __future__ import annotations

from pathlib import Path

from core import config


def test_price_params_可持久化读写(tmp_path: Path, monkeypatch) -> None:
    cfg_path = tmp_path / "price_params.local.json"
    monkeypatch.setattr(config, "PRICE_PARAMS_FILE", cfg_path)

    config.save_price_params({"cutting_rate": 9.9, "tap_al": 1.1})
    loaded = config.load_price_params()

    assert loaded["cutting_rate"] == 9.9
    assert loaded["tap_al"] == 1.1
    assert loaded["grinding_small"] == config.get_default_price_params()["grinding_small"]


def test_load_price_params_坏文件回退默认值(tmp_path: Path, monkeypatch) -> None:
    cfg_path = tmp_path / "price_params.local.json"
    cfg_path.write_text("{not-json", encoding="utf-8")
    monkeypatch.setattr(config, "PRICE_PARAMS_FILE", cfg_path)

    loaded = config.load_price_params()

    assert loaded == config.get_default_price_params()

