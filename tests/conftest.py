from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.calculator import _blank_row
from core.config import get_default_price_params


@pytest.fixture
def default_price_params():
    return get_default_price_params()


@pytest.fixture
def valid_quote_row():
    row = _blank_row(1)
    row.update(
        {
            "料品名称": "测试件",
            "材质": "6061铝板",
            "长(mm)": 100.0,
            "宽(mm)": 100.0,
            "厚度(mm)": 2.0,
            "米数(切割长度)": 1.5,
            "穿孔数量": 2.0,
            "焊接长度(mm)": 0.0,
            "沉孔数量": 1.0,
            "折弯次数": 2.0,
            "压铆数量": 3.0,
            "攻丝数量": 4.0,
            "激光打标": 0.0,
            "表面处理": "喷粉",
            "加工数量": 10.0,
        }
    )
    return row
