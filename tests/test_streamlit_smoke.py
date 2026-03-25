from __future__ import annotations

from pathlib import Path

import pandas as pd
from streamlit.testing.v1 import AppTest

from core.calculator import _blank_row
from core.config import QUOTE_DF_SCHEMA_VERSION

PROJECT_ROOT = Path(__file__).resolve().parents[1]


def _build_valid_quote_df() -> pd.DataFrame:
    row = _blank_row(1)
    row.update(
        {
            "料品名称": "页面冒烟验证件",
            "材质": "6061铝板",
            "长(mm)": 100.0,
            "宽(mm)": 100.0,
            "厚度(mm)": 2.0,
            "米数(切割长度)": 1.5,
            "折弯次数": 2.0,
            "表面处理": "喷粉",
            "加工数量": 10.0,
        }
    )
    return pd.DataFrame([row])


def test_app_home_smoke_renders_homepage_navigation() -> None:
    app = AppTest.from_file(str(PROJECT_ROOT / "app_home.py"))

    app.run(timeout=20)

    assert len(app.exception) == 0
    markdown_text = "\n".join(getattr(item, "value", "") for item in app.markdown)
    assert "青岛宏泰铭润机械" in markdown_text
    assert "多客户专属版" in markdown_text
    assert len(app.markdown) >= 4


def test_laoying_page_smoke_renders_core_widgets() -> None:
    app = AppTest.from_file(str(PROJECT_ROOT / "pages" / "2_崂应报价.py"))

    app.run(timeout=20)

    assert len(app.exception) == 0
    assert len(app.expander) >= 2
    assert len(app.number_input) >= 10
    assert any(button.key == "btn_update_calc" for button in app.button)
    assert any("整单汇总" in subheader.value for subheader in app.subheader)


def test_laoying_page_smoke_recalc_with_prefilled_valid_data() -> None:
    app = AppTest.from_file(str(PROJECT_ROOT / "pages" / "2_崂应报价.py"))
    app.session_state["quote_df"] = _build_valid_quote_df()
    app.session_state["quote_df_schema_version"] = QUOTE_DF_SCHEMA_VERSION

    app.run(timeout=20)
    next(button for button in app.button if button.key == "btn_update_calc").click()
    app.run(timeout=20)

    assert len(app.exception) == 0
    assert len(app.error) == 0
    metric_values = [metric.value for metric in app.metric]
    assert any("10" in value for value in metric_values)
    assert any("0.710" in value or "0.709" in value for value in metric_values)
    assert any("45.83" in value for value in metric_values)
