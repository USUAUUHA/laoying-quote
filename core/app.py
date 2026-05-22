from __future__ import annotations

import json
import sys
from datetime import datetime
from pathlib import Path
from typing import Dict, Optional

import pandas as pd
import streamlit as st

from .calculator import (
    _blank_row,
    merge_preserve_calc,
    migrate_dataframe,
    needs_migration,
    recalc_table,
    sync_material_unit_price,
)
from .config import (
    APP_SUBTITLE,
    APP_TITLE,
    CALC_COLS,
    COL_MATERIAL_PRICE,
    COL_WEIGHT_AFTER_THICK,
    DEFAULT_WEIGHT_COEFFICIENT,
    MATERIAL_LIBRARY,
    QUOTE_DF_SCHEMA_VERSION,
    SURFACE_OPTIONS,
    TABLE_COLUMNS,
    get_default_price_params,
    load_price_params,
    save_price_params as save_price_params_to_config,
)
from .data_io import parse_parts_upload
from .excel_export import export_table_excel

_ROOT_APP = Path(__file__).resolve().parent.parent
if str(_ROOT_APP) not in sys.path:
    sys.path.insert(0, str(_ROOT_APP))

try:
    from common_quote_export import build_laoying_quote_detail_workbook_bytes
except ImportError:
    build_laoying_quote_detail_workbook_bytes = None  # type: ignore[assignment]


def _sanitize_price_params(raw: object) -> Dict[str, float]:
    defaults = get_default_price_params()
    if not isinstance(raw, dict):
        return defaults.copy()

    out: Dict[str, float] = {}
    for key, default in defaults.items():
        try:
            value = float(raw.get(key, default))
        except (TypeError, ValueError):
            value = default
        out[key] = value if value >= 0 else default
    return out


def load_saved_price_params(path: Path | None = None) -> Dict[str, float]:
    if path is None:
        return load_price_params()
    if not path.is_file():
        return get_default_price_params()
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return get_default_price_params()
    return _sanitize_price_params(raw)


def save_price_params(price_params: Dict[str, float], path: Path | None = None) -> None:
    if path is None:
        save_price_params_to_config(price_params)
        return
    safe = _sanitize_price_params(price_params)
    path.write_text(json.dumps(safe, ensure_ascii=False, indent=2), encoding="utf-8")


def _persist_price_params_if_needed(price_params: Dict[str, float], path: Path | None = None) -> Optional[str]:
    safe = _sanitize_price_params(price_params)
    if st.session_state.get("_last_saved_price_params") == safe:
        return None
    try:
        save_price_params(safe, path=path)
        st.session_state["_last_saved_price_params"] = safe
        return None
    except Exception as e:
        return str(e)


def init_table() -> None:
    if "quote_df" not in st.session_state:
        st.session_state["quote_df"] = pd.DataFrame([_blank_row(i + 1) for i in range(5)])
        st.session_state["quote_df_schema_version"] = QUOTE_DF_SCHEMA_VERSION
        return

    quote_df = st.session_state["quote_df"]
    if (
        st.session_state.get("quote_df_schema_version") != QUOTE_DF_SCHEMA_VERSION
        or needs_migration(quote_df)
    ):
        st.session_state["quote_df"] = migrate_dataframe(quote_df)
        st.session_state["quote_df_schema_version"] = QUOTE_DF_SCHEMA_VERSION


def init_pp_widget_keys() -> None:
    saved = load_saved_price_params()
    for key, default in get_default_price_params().items():
        widget_key = f"pp_{key}"
        if widget_key not in st.session_state:
            st.session_state[widget_key] = float(saved.get(key, default))
    if "_last_saved_price_params" not in st.session_state:
        st.session_state["_last_saved_price_params"] = _sanitize_price_params(saved)


def price_params_for_calc() -> Dict[str, float]:
    return {key: float(st.session_state[f"pp_{key}"]) for key in get_default_price_params()}


def main(*, set_page_config: bool = True) -> None:
    if set_page_config:
        st.set_page_config(page_title=APP_TITLE, layout="wide")

    st.title(APP_TITLE)
    st.caption(APP_SUBTITLE)

    if "weight_coef" not in st.session_state:
        st.session_state["weight_coef"] = DEFAULT_WEIGHT_COEFFICIENT

    init_table()
    init_pp_widget_keys()

    st.sidebar.subheader("总重系数")
    st.sidebar.caption("与 Excel 一致：理重×系数；材料费与按重量计费项目都基于总重(kg)计算。")
    coef = st.sidebar.number_input(
        "总重系数（默认1.3，与 Excel 中×1.3 对应）",
        min_value=0.01,
        max_value=99.0,
        value=float(st.session_state["weight_coef"]),
        step=0.01,
        format="%.2f",
        key="weight_coef_input",
    )
    st.session_state["weight_coef"] = coef

    with st.sidebar.expander("单价与公式说明（可直接改单价）", expanded=False):
        st.markdown(
            """
**重量 / 材料费（与 Excel 一致）**
`重量(kg) = 长×宽×厚×密度×总重系数 ÷ 1e6`
`总重(kg) = 重量(kg) × 加工数量`
`材料费(元) = 总重(kg) × 材料单价(元/kg)`

**切割**
`切割费 = 米数(切割长度) × 厚度(mm) × 切割系数`

**打磨**
若 `长<200` 或 `宽<200` 用小件基数，否则用大件基数。

**折弯**
铝板、钢/不锈钢用不同单价；若长或宽大于 100，则统一用大件单价。
"""
        )
        st.divider()
        st.caption("切割 / 穿孔 / 焊接 / 沉孔")
        st.number_input("切割系数（米×厚×）", min_value=0.0, step=0.001, format="%.3f", key="pp_cutting_rate")
        st.number_input("穿孔：数量×厚×系数", min_value=0.0, step=0.001, format="%.3f", key="pp_pierce_per_t")
        c1, c2 = st.columns(2)
        with c1:
            st.number_input("焊接系数A", min_value=0.0, step=1.0, format="%.1f", key="pp_weld_a")
        with c2:
            st.number_input("焊接系数B", min_value=0.0, step=1.0, format="%.1f", key="pp_weld_b")
        st.caption("焊接费 = 焊接长度(mm) × (A + B)")
        st.number_input("沉孔单价/个", min_value=0.0, step=0.01, format="%.2f", key="pp_countersink")

        st.divider()
        st.caption("打磨（基数 × 数量 × 倍率）")
        st.number_input("打磨基数：长或宽<200(mm) 时（每件）", min_value=0.0, step=0.01, format="%.2f", key="pp_grinding_small")
        st.number_input("打磨基数：否则（每件）", min_value=0.0, step=0.01, format="%.2f", key="pp_grinding_large")
        st.number_input("打磨倍率（默认1.13）", min_value=0.0, step=0.001, format="%.3f", key="pp_grinding_mult")

        st.divider()
        st.caption("折弯单价（元/次·件，再×折弯次数×数量）")
        st.number_input("折弯：铝板（长宽均≤100 时）", min_value=0.0, step=0.01, format="%.2f", key="pp_bend_al")
        st.number_input("折弯：钢/不锈钢（长宽均≤100 时）", min_value=0.0, step=0.01, format="%.2f", key="pp_bend_steel")
        st.number_input("折弯：长或宽>100 时", min_value=0.0, step=0.01, format="%.2f", key="pp_bend_large")

        st.divider()
        st.caption("压铆 / 攻丝")
        c3, c4 = st.columns(2)
        with c3:
            st.number_input("压铆单价A/个", min_value=0.0, step=0.01, format="%.2f", key="pp_pem_a")
        with c4:
            st.number_input("压铆单价B/个", min_value=0.0, step=0.01, format="%.2f", key="pp_pem_b")
        st.number_input("攻丝：铝板 单价", min_value=0.0, step=0.01, format="%.2f", key="pp_tap_al")
        st.number_input("攻丝：钢/不锈钢 单价", min_value=0.0, step=0.01, format="%.2f", key="pp_tap_steel")

        if st.button("恢复加工单价默认值", key="reset_pp_defaults"):
            defaults = get_default_price_params()
            for key, value in defaults.items():
                st.session_state[f"pp_{key}"] = float(value)
            save_price_params(defaults)
            st.session_state["_last_saved_price_params"] = defaults.copy()
            st.rerun()

    pp = price_params_for_calc()
    persist_error = _persist_price_params_if_needed(pp)
    if persist_error:
        st.sidebar.warning(f"单价参数保存失败：{persist_error}")

    with st.expander("导入零件表（Excel / CSV / 截图）", expanded=False):
        st.caption("自动识别零件名称、材质、板厚、长宽/尺寸、切割长度、轮廓数量、沉孔/折弯/压铆等。")
        up = st.file_uploader(
            "选择文件",
            type=["xlsx", "xls", "csv", "png", "jpg", "jpeg"],
            key="parts_import_uploader",
        )
        if up is not None and st.button("解析并覆盖当前表格", key="btn_import_parts"):
            try:
                ext = Path(up.name).suffix.lower()
                msg = "正在解析表格..." if ext not in (".png", ".jpg", ".jpeg") else "正在初始化 OCR 并识别图片..."
                with st.spinner(msg):
                    new_df = parse_parts_upload(up.getvalue(), up.name)
                st.session_state["quote_df"] = new_df
                st.session_state["quote_df_schema_version"] = QUOTE_DF_SCHEMA_VERSION
                st.session_state["last_recalc_errors"] = ([], [])
                st.success(f"已导入 {len(new_df)} 行，请向下滚动点击「更新计算」。")
                st.rerun()
            except Exception as e:
                st.error(str(e))

    st.caption("提示：改完数据后点表格下方的「更新计算」再导出；仅填表时不会自动重算。")

    prev_snapshot = st.session_state["quote_df"].copy()
    df_in = st.session_state["quote_df"].copy()

    edited_df = st.data_editor(
        df_in,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        height=560,
        disabled=["序号", COL_MATERIAL_PRICE, COL_WEIGHT_AFTER_THICK] + CALC_COLS,
        column_config={
            "选择": st.column_config.CheckboxColumn("选择", help="勾选后可删除"),
            "序号": st.column_config.TextColumn("序号"),
            "料品名称": st.column_config.TextColumn("料品名称"),
            "料号": st.column_config.TextColumn("料号"),
            "料品规格": st.column_config.TextColumn("料品规格"),
            "材质": st.column_config.SelectboxColumn("材质", options=list(MATERIAL_LIBRARY.keys())),
            COL_MATERIAL_PRICE: st.column_config.NumberColumn(COL_MATERIAL_PRICE, format="%.2f"),
            "长(mm)": st.column_config.NumberColumn("长(mm)", step=1.0),
            "宽(mm)": st.column_config.NumberColumn("宽(mm)", step=1.0),
            "厚度(mm)": st.column_config.NumberColumn("厚度(mm)", step=0.1),
            COL_WEIGHT_AFTER_THICK: st.column_config.NumberColumn("重量(kg)", format="%.6f"),
            "米数(切割长度)": st.column_config.NumberColumn("米数(切割长度)", step=0.001, format="%.3f"),
            "穿孔数量": st.column_config.NumberColumn("穿孔数量", step=1.0),
            "焊接长度(mm)": st.column_config.NumberColumn("焊接长度(mm)", step=1.0),
            "沉孔数量": st.column_config.NumberColumn("沉孔数量", step=1.0),
            "折弯次数": st.column_config.NumberColumn("折弯次数", step=1.0),
            "压铆数量": st.column_config.NumberColumn("压铆数量", step=1.0),
            "攻丝数量": st.column_config.NumberColumn("攻丝数量", step=1.0),
            "激光打标": st.column_config.NumberColumn("激光打标", step=1.0),
            "表面处理": st.column_config.SelectboxColumn("表面处理", options=SURFACE_OPTIONS),
            "加工数量": st.column_config.NumberColumn("加工数量", step=1.0),
            "总重(kg)": st.column_config.NumberColumn("总重(kg)", format="%.6f"),
            "打磨(元)": st.column_config.NumberColumn("打磨(元)", format="%.3f"),
            "材料费(元)": st.column_config.NumberColumn("材料费(元)", format="%.2f"),
            "总加工费(元)": st.column_config.NumberColumn("总加工费(元)", format="%.2f"),
            "表面处理费(元)": st.column_config.NumberColumn("表面处理费(元)", format="%.2f"),
            "含税单价(元/件)": st.column_config.NumberColumn("含税单价(元/件)", format="%.2f"),
            "含税总价(元)": st.column_config.NumberColumn("含税总价(元)", format="%.2f"),
        },
        key="quote_editor",
    )

    c_up1, c_up2, _ = st.columns([1, 2, 5])
    with c_up1:
        do_update = st.button(
            "更新计算",
            type="primary",
            use_container_width=True,
            help="点击后才会按公式重算重量、材料费、加工费等。",
            key="btn_update_calc",
        )
    with c_up2:
        st.caption("填完或粘贴后点此，再导出。")

    if do_update:
        calc_df, invalid_rows, error_msgs = recalc_table(edited_df, weight_coef=coef, price_params=pp)
        st.session_state["quote_df"] = calc_df
        st.session_state["quote_df_schema_version"] = QUOTE_DF_SCHEMA_VERSION
        st.session_state["last_recalc_errors"] = (invalid_rows, error_msgs)
        st.rerun()
    else:
        merged = sync_material_unit_price(merge_preserve_calc(edited_df, prev_snapshot))
        st.session_state["quote_df"] = merged[TABLE_COLUMNS].copy()
        st.session_state["quote_df_schema_version"] = QUOTE_DF_SCHEMA_VERSION
        invalid_rows, error_msgs = st.session_state.get("last_recalc_errors", ([], []))

    calc_df = migrate_dataframe(st.session_state["quote_df"])
    if not isinstance(invalid_rows, list):
        invalid_rows = []
    if not isinstance(error_msgs, list):
        error_msgs = []

    b1, b2, _ = st.columns([1, 1, 6])
    with b1:
        if st.button("插入新行", use_container_width=True):
            df2 = calc_df.copy()
            df2.loc[len(df2)] = _blank_row(len(df2) + 1)
            st.session_state["quote_df"] = df2
            st.session_state["quote_df_schema_version"] = QUOTE_DF_SCHEMA_VERSION
            st.rerun()
    with b2:
        if st.button("删除选中行", use_container_width=True):
            df2 = calc_df.copy()
            kept = df2[~df2["选择"]].copy()
            if kept.empty:
                kept = pd.DataFrame([_blank_row(1)])
            st.session_state["quote_df"] = migrate_dataframe(kept)
            st.session_state["quote_df_schema_version"] = QUOTE_DF_SCHEMA_VERSION
            st.rerun()

    if invalid_rows:
        st.error(f"当前有 {len(invalid_rows)} 行存在必填项缺失或数据错误（序号已标红）。")
        with st.expander("查看错误明细"):
            for msg in error_msgs[:100]:
                st.write(f"- {msg}")

    valid_df = calc_df[~calc_df["序号"].astype(str).str.startswith("🔴")].copy()
    total_parts = int(pd.to_numeric(valid_df["加工数量"], errors="coerce").fillna(0).sum())
    total_weight = float(pd.to_numeric(valid_df["总重(kg)"], errors="coerce").fillna(0).sum())
    total_price = float(pd.to_numeric(valid_df["含税总价(元)"], errors="coerce").fillna(0).sum())

    st.markdown("---")
    st.subheader("整单汇总")
    s1, s2, s3, s4 = st.columns([2, 2, 3, 2])
    s1.metric("零件总数量", f"{total_parts} 件")
    s2.metric("整单总重量(×系数)", f"{total_weight:.3f} kg")
    s3.metric("整单最终含税总价", f"{total_price:.2f} 元")
    with s4:
        if st.button("崂应表格导出", type="primary", use_container_width=True):
            try:
                with st.spinner("正在生成崂应导出文件..."):
                    out_path = export_table_excel(calc_df, weight_coef=coef, price_params=pp)
                st.success(f"导出成功：{out_path}")
            except Exception as e:
                st.error(f"导出失败：{e}")
        if build_laoying_quote_detail_workbook_bytes is not None:
            ly_detail_name = f"崂应-报价明细_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            st.download_button(
                label="导出报价明细（含加工单价+材料库+公式说明）",
                data=build_laoying_quote_detail_workbook_bytes(
                    calc_df,
                    price_params=pp,
                    weight_coef=float(coef),
                    material_library=MATERIAL_LIBRARY,
                ),
                file_name=ly_detail_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                help="导出当前表格数值快照；另附加工单价、材质库及说明，便于线下改价对账。",
            )
