from __future__ import annotations

from typing import Any, Dict, List, Tuple

import pandas as pd

from .config import (
    CALC_COLS,
    COL_MATERIAL_PRICE,
    COL_WEIGHT_AFTER_THICK,
    COLS_RECALC_ONLY,
    INPUT_COLS,
    MANDATORY_COLS,
    MATERIAL_LIBRARY,
    NUMERIC_INPUT_COLS,
    SURFACE_OPTIONS,
    TABLE_COLUMNS,
)


def _to_float(v: Any, default: float = 0.0) -> Tuple[float, bool]:
    if v is None:
        return default, True
    t = str(v).strip()
    if t == "":
        return default, True
    try:
        return float(t), True
    except Exception:
        return default, False


def _blank_row(seq: int) -> Dict[str, Any]:
    row: Dict[str, Any] = {"选择": False, "序号": str(seq)}
    first_mat = list(MATERIAL_LIBRARY.keys())[0]
    for c in INPUT_COLS:
        if c == COL_WEIGHT_AFTER_THICK:
            row[c] = 0.0
        elif c == "材质":
            row[c] = first_mat
        elif c == COL_MATERIAL_PRICE:
            row[c] = float(MATERIAL_LIBRARY[first_mat]["unit_price"])
        elif c == "表面处理":
            row[c] = "无"
        elif c == "加工数量":
            row[c] = 1
        else:
            row[c] = "" if c in ("料品名称", "料号", "料品规格") else 0.0
    for c in CALC_COLS:
        row[c] = 0.0
    return row


def needs_migration(df: pd.DataFrame | None) -> bool:
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return True
    cols = set(df.columns)
    if "单重(kg)" in cols:
        return True
    return any(col not in cols for col in TABLE_COLUMNS)


def migrate_dataframe(df: pd.DataFrame | None) -> pd.DataFrame:
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return pd.DataFrame([_blank_row(1)])

    out = df.copy()
    blank = _blank_row(1)
    for c in TABLE_COLUMNS:
        if c not in out.columns:
            if c == "选择":
                out[c] = False
            elif c in (COL_WEIGHT_AFTER_THICK, COL_MATERIAL_PRICE) or c in CALC_COLS:
                out[c] = 0.0
            else:
                out[c] = blank.get(c, "")

    if "单重(kg)" in out.columns:
        out = out.drop(columns=["单重(kg)"], errors="ignore")

    return out[[c for c in TABLE_COLUMNS if c in out.columns]]


def merge_preserve_calc(edited_df: pd.DataFrame, prev_full: pd.DataFrame) -> pd.DataFrame:
    out = migrate_dataframe(edited_df)
    prev_full = migrate_dataframe(prev_full)
    for i in range(len(out)):
        if i < len(prev_full):
            for c in COLS_RECALC_ONLY:
                out.at[i, c] = prev_full.iloc[i][c]
        else:
            for c in COLS_RECALC_ONLY:
                out.at[i, c] = 0.0
    return out


def sync_material_unit_price(df: pd.DataFrame) -> pd.DataFrame:
    out = migrate_dataframe(df)
    for i in range(len(out)):
        name = str(out.at[i, "材质"]).strip()
        out.at[i, COL_MATERIAL_PRICE] = float(MATERIAL_LIBRARY.get(name, {}).get("unit_price", 0.0))
    return out


def calc_quote_row(
    row: pd.Series,
    weight_coef: float,
    price_params: Dict[str, float],
) -> Tuple[Dict[str, float], List[str]]:
    errs: List[str] = []

    for col in MANDATORY_COLS:
        if col not in row.index:
            errs.append(f"缺少列：{col}")
            return {}, errs

    for col in MANDATORY_COLS:
        if str(row.get(col, "")).strip() == "":
            errs.append(f"{col}为空")

    material_name = str(row.get("材质", "")).strip()
    if material_name not in MATERIAL_LIBRARY:
        errs.append("材质未匹配")

    nums: Dict[str, float] = {}
    for col in NUMERIC_INPUT_COLS:
        v, ok = _to_float(row.get(col, 0.0), 0.0)
        if not ok:
            errs.append(f"{col}非数字")
        if v < 0:
            errs.append(f"{col}不能为负")
        nums[col] = v

    if nums["长(mm)"] <= 0 or nums["宽(mm)"] <= 0 or nums["厚度(mm)"] <= 0:
        errs.append("长宽厚必须大于0")
    if nums["加工数量"] <= 0:
        errs.append("加工数量必须大于0")

    surface = str(row.get("表面处理", "")).strip()
    if surface not in SURFACE_OPTIONS:
        errs.append("表面处理无效")

    if errs:
        return {}, errs

    density = float(MATERIAL_LIBRARY[material_name]["density"])
    unit_price = float(MATERIAL_LIBRARY[material_name]["unit_price"])
    is_aluminum = material_name in {"6061铝板", "1060铝板", "5052铝板"}

    length_mm = nums["长(mm)"]
    width_mm = nums["宽(mm)"]
    thickness_mm = nums["厚度(mm)"]
    meter = nums["米数(切割长度)"]
    pierce_count = nums["穿孔数量"]
    weld_length_mm = nums["焊接长度(mm)"]
    countersink_count = nums["沉孔数量"]
    bend_times = nums["折弯次数"]
    pem_count = nums["压铆数量"]
    tapping_count = nums["攻丝数量"]
    qty = int(nums["加工数量"])

    p = price_params
    single_weight_physical_kg = length_mm * width_mm * thickness_mm * density / 1_000_000
    single_weight_kg = single_weight_physical_kg * weight_coef
    total_weight_display_kg = single_weight_kg * qty

    material_cost = total_weight_display_kg * unit_price
    cutting_cost = meter * thickness_mm * p["cutting_rate"] * p["grinding_mult"]
    grinding_base = p["grinding_small"] if (length_mm <= 200 or width_mm <= 200) else p["grinding_large"]
    grinding_cost = grinding_base * qty * p["grinding_mult"]
    pierce_unit = thickness_mm * p["pierce_per_t"]
    pierce_cost = pierce_count * pierce_unit
    welding_cost = weld_length_mm * (p["weld_a"] + p["weld_b"])
    countersink_unit = p["countersink"]
    countersink_cost = countersink_count * countersink_unit

    bend_unit_price = p["bend_al"] if is_aluminum else p["bend_steel"]
    if length_mm > 100 or width_mm > 100:
        bend_unit_price = p["bend_large"]
    bend_qty_total = bend_times * qty
    bend_cost = bend_unit_price * bend_times * qty

    pem_unit = p["pem_a"] + p["pem_b"]
    pem_cost = pem_count * pem_unit
    tapping_unit_price = p["tap_al"] if is_aluminum else p["tap_steel"]
    tapping_qty_total = tapping_count * qty
    tapping_cost = tapping_unit_price * tapping_count * qty
    laser_mark_cost = nums["激光打标"] * p["grinding_mult"]

    total_process_cost = (
        cutting_cost
        + grinding_cost
        + pierce_cost
        + welding_cost
        + countersink_cost
        + bend_cost
        + pem_cost
        + tapping_cost
        + laser_mark_cost
    )

    unfold_area_mm2 = length_mm * width_mm * 2
    if surface == "无":
        surface_cost = 0.0
    elif surface == "喷粉":
        surface_cost = (unfold_area_mm2 / 1_000_000) * 25 * qty
    elif surface == "喷砂":
        surface_cost = (unfold_area_mm2 / 10_000) * 1.2 * qty
    elif surface == "氧化":
        surface_cost = (unfold_area_mm2 / 10_000) * 0.8 * qty
    elif surface == "抛光":
        surface_cost = (unfold_area_mm2 / 10_000) * 4 * qty
    else:
        surface_cost = total_weight_display_kg * 15

    final_total_price = material_cost + total_process_cost + surface_cost
    unit_piece_price = final_total_price / qty if qty else 0.0

    return {
        COL_WEIGHT_AFTER_THICK: single_weight_kg,
        "_total_weight_display_kg": total_weight_display_kg,
        "_material_name": material_name,
        "_material_cost": material_cost,
        "_qty": qty,
        "_surface": surface,
        "_cutting_cost": cutting_cost,
        "_grinding_cost": grinding_cost,
        "_pierce_unit": pierce_unit,
        "_pierce_cost": pierce_cost,
        "_welding_cost": welding_cost,
        "_countersink_unit": countersink_unit,
        "_countersink_cost": countersink_cost,
        "_bend_unit_price": bend_unit_price,
        "_bend_qty_total": bend_qty_total,
        "_bend_cost": bend_cost,
        "_pem_cost": pem_cost,
        "_tapping_unit_price": tapping_unit_price,
        "_tapping_qty_total": tapping_qty_total,
        "_tapping_cost": tapping_cost,
        "_surface_cost": surface_cost,
        "打磨(元)": grinding_cost,
        "材料费(元)": material_cost,
        "总加工费(元)": total_process_cost,
        "表面处理费(元)": surface_cost,
        "含税单价(元/件)": unit_piece_price,
        "含税总价(元)": final_total_price,
    }, []


def recalc_table(
    df: pd.DataFrame,
    weight_coef: float,
    price_params: Dict[str, float],
) -> Tuple[pd.DataFrame, List[int], List[str]]:
    if df is None or df.empty:
        return pd.DataFrame([_blank_row(1)]), [1], ["无有效数据"]

    out = migrate_dataframe(df)
    invalid_rows: List[int] = []
    error_msgs: List[str] = []
    seq_values: List[str] = []

    for idx in range(len(out)):
        row_no = idx + 1
        calc_res, errs = calc_quote_row(out.iloc[idx], weight_coef=weight_coef, price_params=price_params)
        if errs:
            invalid_rows.append(row_no)
            seq_values.append(f"🔴{row_no}")
            error_msgs.append(f"第{row_no}行：{'；'.join(errs)}")
            out.at[idx, COL_WEIGHT_AFTER_THICK] = 0.0
            for c in CALC_COLS:
                out.at[idx, c] = 0.0
            continue

        seq_values.append(str(row_no))
        out.at[idx, COL_WEIGHT_AFTER_THICK] = round(float(calc_res[COL_WEIGHT_AFTER_THICK]), 6)
        out.at[idx, "总重(kg)"] = round(float(calc_res["_total_weight_display_kg"]), 6)
        out.at[idx, "打磨(元)"] = round(float(calc_res["打磨(元)"]), 3)
        out.at[idx, "材料费(元)"] = round(float(calc_res["材料费(元)"]), 2)
        out.at[idx, "总加工费(元)"] = round(float(calc_res["总加工费(元)"]), 2)
        out.at[idx, "表面处理费(元)"] = round(float(calc_res["表面处理费(元)"]), 2)
        out.at[idx, "含税单价(元/件)"] = round(float(calc_res["含税单价(元/件)"]), 2)
        out.at[idx, "含税总价(元)"] = round(float(calc_res["含税总价(元)"]), 2)

    out["序号"] = seq_values
    out["选择"] = out["选择"].fillna(False).astype(bool)
    out = sync_material_unit_price(out)
    return out[[c for c in TABLE_COLUMNS if c in out.columns]], invalid_rows, error_msgs
