from __future__ import annotations

import pandas as pd
import pytest

from core.calculator import calc_quote_row, merge_preserve_calc, migrate_dataframe, recalc_table
from core.config import COL_WEIGHT_AFTER_THICK


def test_calc_quote_row_valid_input_returns_expected_totals(valid_quote_row, default_price_params):
    result, errors = calc_quote_row(pd.Series(valid_quote_row), 1.3, default_price_params)

    assert errors == []
    assert result[COL_WEIGHT_AFTER_THICK] == pytest.approx(0.07098, rel=1e-6)
    assert result["_total_weight_display_kg"] == pytest.approx(0.7098, rel=1e-6)
    assert result["材料费(元)"] == pytest.approx(22.7136, rel=1e-6)
    assert result["_cutting_cost"] == pytest.approx(3.6, rel=1e-6)
    assert result["含税总价(元)"] == pytest.approx(63.8836, rel=1e-6)


def test_calc_quote_row_missing_required_field_returns_errors(valid_quote_row, default_price_params):
    valid_quote_row["料品名称"] = ""

    result, errors = calc_quote_row(pd.Series(valid_quote_row), 1.3, default_price_params)

    assert result == {}
    assert "料品名称为空" in errors


def test_recalc_table_marks_invalid_rows_and_keeps_valid_rows(valid_quote_row, default_price_params):
    invalid_row = dict(valid_quote_row)
    invalid_row["厚度(mm)"] = 0
    df = pd.DataFrame([valid_quote_row, invalid_row])

    out, invalid_rows, error_msgs = recalc_table(df, 1.3, default_price_params)

    assert invalid_rows == [2]
    assert out.iloc[0]["序号"] == "1"
    assert out.iloc[1]["序号"] == "🔴2"
    assert out.iloc[1]["含税总价(元)"] == 0.0
    assert error_msgs[0].startswith("第2行：")


def test_merge_preserve_calc_keeps_previous_calculated_columns(valid_quote_row):
    edited_df = pd.DataFrame([valid_quote_row])
    prev_df = migrate_dataframe(pd.DataFrame([valid_quote_row]))
    prev_df.at[0, COL_WEIGHT_AFTER_THICK] = 9.99
    prev_df.at[0, "含税总价(元)"] = 88.88

    merged = merge_preserve_calc(edited_df, prev_df)

    assert merged.at[0, COL_WEIGHT_AFTER_THICK] == 9.99
    assert merged.at[0, "含税总价(元)"] == 88.88
