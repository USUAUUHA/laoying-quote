from __future__ import annotations

import pandas as pd
import pytest

from core.data_io import build_quote_from_sheet_df, map_headers_to_fields, normalize_material_name


@pytest.mark.parametrize(
    ("raw_name", "expected"),
    [
        ("5052", "5052铝板"),
        ("6061-T6", "6061铝板"),
        ("304", "SUS304"),
        ("Q235", "Q235"),
        ("未知材质", "未知材质"),
    ],
)
def test_normalize_material_name_maps_common_aliases(raw_name, expected):
    assert normalize_material_name(raw_name) == expected


def test_map_headers_to_fields_matches_common_headers():
    headers = ["零件名称", "材质", "板厚", "零件尺寸(mm)", "切割长度", "折弯次数"]

    mapped = map_headers_to_fields(headers)

    assert mapped["料品名称"] == 0
    assert mapped["材质"] == 1
    assert mapped["厚度(mm)"] == 2
    assert mapped["零件尺寸"] == 3
    assert mapped["米数(切割长度)"] == 4
    assert mapped["折弯次数"] == 5


def test_build_quote_from_sheet_df_raises_when_material_column_missing():
    raw = pd.DataFrame(
        [
            ["零件名称", "板厚", "长(mm)", "宽(mm)"],
            ["A件", 2, 100, 80],
        ]
    )

    with pytest.raises(ValueError, match="未识别到「材质」列"):
        build_quote_from_sheet_df(raw)


def test_build_quote_from_sheet_df_parses_size_and_material():
    raw = pd.DataFrame(
        [
            ["零件名称", "材质", "板厚", "零件尺寸(mm)", "切割长度", "折弯次数"],
            ["A件", "5052", 2, "100x80", 1.5, 2],
        ]
    )

    result = build_quote_from_sheet_df(raw)

    assert len(result) == 1
    assert result.at[0, "料品名称"] == "A件"
    assert result.at[0, "材质"] == "5052铝板"
    assert result.at[0, "长(mm)"] == 100
    assert result.at[0, "宽(mm)"] == 80
    assert result.at[0, "米数(切割长度)"] == 1.5
