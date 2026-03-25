from __future__ import annotations

import pandas as pd
import pytest

from core.calculator import recalc_table
from core.excel_export import compute_ly_export_breakdown, export_table_excel


def test_compute_ly_export_breakdown_reuses_calculation_totals(valid_quote_row, default_price_params):
    breakdown = compute_ly_export_breakdown(pd.Series(valid_quote_row), 1.3, default_price_params)

    assert breakdown is not None
    assert breakdown["料费"] == pytest.approx(22.71, rel=1e-3)
    assert breakdown["合计"] == pytest.approx(63.88, rel=1e-3)
    assert breakdown["折弯费"] == pytest.approx(10.0, rel=1e-6)
    assert breakdown["_surface"] == "喷粉"


def test_export_table_excel_writes_file_to_target_directory(tmp_path, valid_quote_row, default_price_params, monkeypatch):
    df, _, _ = recalc_table(pd.DataFrame([valid_quote_row]), 1.3, default_price_params)
    monkeypatch.setattr("core.excel_export.EXPORT_REFERENCE_ROOT", tmp_path / "exports")

    out_path = export_table_excel(df, 1.3, default_price_params)

    assert out_path.exists()
    assert out_path.suffix == ".xlsx"


def test_export_table_excel_raises_clear_error_when_directory_unwritable(
    tmp_path, valid_quote_row, default_price_params, monkeypatch
):
    df, _, _ = recalc_table(pd.DataFrame([valid_quote_row]), 1.3, default_price_params)
    monkeypatch.setattr("core.excel_export.EXPORT_REFERENCE_ROOT", tmp_path / "exports")

    def raise_unwritable(_directory):
        raise RuntimeError("导出目录不可写：mock")

    monkeypatch.setattr("core.excel_export._ensure_directory_writable", raise_unwritable)

    with pytest.raises(RuntimeError, match="导出目录不可写"):
        export_table_excel(df, 1.3, default_price_params)
