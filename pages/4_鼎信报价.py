"""
页面：鼎信报价（占位）

作用：
- 预留鼎信客户专属报价模板入口；
- 当前仅展示占位信息，后续可直接扩展；
- 提供空白明细模板导出，便于线下先建表。
"""

import sys
from datetime import datetime
from pathlib import Path

import streamlit as st

_ROOT = Path(__file__).resolve().parents[1]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))
from common_quote_export import build_placeholder_quote_workbook_bytes

st.set_page_config(page_title="鼎信报价", page_icon="🏗️", layout="wide")

st.title("鼎信报价")
st.info("📌 鼎信专属报价模板，后续更新上线，敬请期待")

_TEMPLATE_COLS = [
    "序号",
    "料号",
    "物料名称",
    "规格型号",
    "材质",
    "长(mm)",
    "宽(mm)",
    "厚(mm)",
    "数量",
    "备注",
]
st.download_button(
    "导出报价明细模板（空白可编辑）",
    data=build_placeholder_quote_workbook_bytes(
        customer_name="鼎信",
        columns=_TEMPLATE_COLS,
        note="鼎信报价页尚未接入计算逻辑。本模板仅提供通用列头，可在 Excel 中自行扩展公式；正式上线后此处将导出完整明细。",
    ),
    file_name=f"鼎信-报价明细模板_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

st.divider()
st.page_link("报价系统首页.py", label="返回首页", icon="🏠")
