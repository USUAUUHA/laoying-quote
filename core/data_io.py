from __future__ import annotations

import io
import re
from functools import lru_cache
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

from .calculator import _blank_row, _to_float, migrate_dataframe, sync_material_unit_price
from .config import MATERIAL_LIBRARY


def normalize_material_name(s: Any) -> str:
    t = str(s).strip() if s is not None else ""
    if not t:
        return list(MATERIAL_LIBRARY.keys())[0]
    if t in MATERIAL_LIBRARY:
        return t
    if re.fullmatch(r"5052.*", t):
        return "5052铝板"
    if re.fullmatch(r"6061.*", t):
        return "6061铝板"
    if re.fullmatch(r"1060.*", t):
        return "1060铝板"
    if re.search(r"5052.*铝", t, re.I):
        return "5052铝板"
    if re.search(r"6061.*铝", t, re.I):
        return "6061铝板"
    if re.search(r"1060.*铝", t, re.I):
        return "1060铝板"
    if t in ("304", "SUS304", "不锈钢304"):
        return "SUS304"
    if t in ("316", "SUS316"):
        return "SUS316"
    if t in ("201", "SUS201"):
        return "SUS201"
    if t in ("430", "SUS430"):
        return "SUS430"
    if t.upper() == "Q235":
        return "Q235"
    if t.upper() == "SGCC":
        return "SGCC"
    return t


def normalize_surface_name(s: Any) -> str:
    t = str(s).strip() if s is not None else ""
    if not t or t == "无":
        return "无"
    for opt in ("喷粉", "喷砂", "氧化", "抛光", "镀锌"):
        if opt in t:
            return opt
    return t


def _parse_dim_two(s: Any) -> Tuple[Optional[float], Optional[float]]:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return None, None
    t = str(s).strip()
    if not t:
        return None, None
    t = t.replace("×", "x").replace("X", "x").replace("*", "x")
    m = re.search(r"([\d.]+)\s*[x×*]\s*([\d.]+)", t, re.I)
    if m:
        try:
            return float(m.group(1)), float(m.group(2))
        except ValueError:
            return None, None
    return None, None


def _find_header_row(df: pd.DataFrame) -> int:
    for i in range(min(45, len(df))):
        row = " ".join(str(x) for x in df.iloc[i].tolist())
        if "材质" in row:
            return i
    return 0


def map_headers_to_fields(headers: List[str]) -> Dict[str, int]:
    out: Dict[str, int] = {}
    for j, h in enumerate(headers):
        h = str(h).replace("\n", "").strip()
        if not h:
            continue
        if ("零件名称" in h) or (h.endswith("名称") and "零件" in h):
            out.setdefault("料品名称", j)
        elif h == "材质" or h.startswith("材质"):
            out.setdefault("材质", j)
        elif "板厚" in h or h == "厚度" or ("厚度" in h and "mm" in h):
            out.setdefault("厚度(mm)", j)
        elif "零件尺寸" in h or (h.startswith("尺寸") and "mm" in h):
            out.setdefault("零件尺寸", j)
        elif "切割" in h and "长度" in h:
            out.setdefault("米数(切割长度)", j)
        elif ("长" in h and "宽" not in h and "尺寸" not in h) or h == "长(mm)":
            out.setdefault("长(mm)", j)
        elif ("宽" in h and "长" not in h and "尺寸" not in h) or h == "宽(mm)":
            out.setdefault("宽(mm)", j)
        elif "轮廓" in h:
            out.setdefault("穿孔数量", j)
        elif "沉孔" in h:
            out.setdefault("沉孔数量", j)
        elif "折弯" in h:
            out.setdefault("折弯次数", j)
        elif "压铆" in h:
            out.setdefault("压铆数量", j)
        elif "攻丝" in h:
            out.setdefault("攻丝数量", j)
        elif "焊接" in h:
            out.setdefault("焊接长度(mm)", j)
        elif "表面处理" in h:
            out.setdefault("表面处理", j)
        elif "缩略图" in h or "图" == h:
            continue
    return out


def _cell_at(df: pd.DataFrame, r: int, j: Optional[int]) -> Any:
    if j is None or j >= df.shape[1]:
        return None
    v = df.iloc[r, j]
    if pd.isna(v):
        return None
    return v


def _load_sheet_bytes(name: str, raw: bytes) -> pd.DataFrame:
    ext = Path(name).suffix.lower()
    bio = io.BytesIO(raw)
    if ext == ".csv":
        try:
            return pd.read_csv(bio, header=None, encoding="utf-8-sig")
        except UnicodeDecodeError:
            bio.seek(0)
            return pd.read_csv(bio, header=None, encoding="gbk")
    xl = pd.ExcelFile(bio)
    if "零件总表" in xl.sheet_names:
        return pd.read_excel(xl, sheet_name="零件总表", header=None)
    return pd.read_excel(xl, header=None)


def _cluster_ocr_boxes_to_rows(boxes: List[Any]) -> List[List[str]]:
    items: List[Tuple[float, float, str]] = []
    for b in boxes:
        if not b:
            continue
        box = b[0]
        txtconf = b[1]
        txt = txtconf[0] if isinstance(txtconf, (list, tuple)) else str(txtconf)
        if not txt or not str(txt).strip():
            continue
        xs = [float(p[0]) for p in box]
        ys = [float(p[1]) for p in box]
        cx = sum(xs) / len(xs)
        cy = sum(ys) / len(ys)
        items.append((cy, cx, str(txt).strip()))
    items.sort(key=lambda t: (t[0], t[1]))
    rows: List[List[str]] = []
    cur: List[Tuple[float, str]] = []
    last_y: Optional[float] = None
    y_thr = 22.0
    for cy, cx, txt in items:
        if last_y is None or abs(cy - last_y) <= y_thr:
            cur.append((cx, txt))
            last_y = cy if last_y is None else (last_y * 0.55 + cy * 0.45)
        else:
            cur.sort(key=lambda x: x[0])
            rows.append([t[1] for t in cur])
            cur = [(cx, txt)]
            last_y = cy
    if cur:
        cur.sort(key=lambda x: x[0])
        rows.append([t[1] for t in cur])
    return rows


@lru_cache(maxsize=1)
def _get_paddle_ocr():
    try:
        from paddleocr import PaddleOCR
    except ImportError as e:
        raise RuntimeError("图片识别需安装：pip install paddleocr（首次会自动下载模型，较慢）") from e
    return PaddleOCR(use_angle_cls=True, lang="ch", show_log=False)


def ocr_image_to_dataframe(file_bytes: bytes) -> pd.DataFrame:
    import numpy as np
    from PIL import Image

    ocr = _get_paddle_ocr()
    img = Image.open(io.BytesIO(file_bytes)).convert("RGB")
    arr = np.asarray(img)
    res = ocr.ocr(arr, cls=True)
    if not res or res[0] is None:
        raise ValueError("OCR 未识别到文字")
    rows = _cluster_ocr_boxes_to_rows(res[0])
    if not rows:
        raise ValueError("OCR 未识别到有效行")
    max_len = max(len(x) for x in rows)
    pad = [r + [""] * (max_len - len(r)) for r in rows]
    return pd.DataFrame(pad)


def build_quote_from_sheet_df(df: pd.DataFrame) -> pd.DataFrame:
    hr = _find_header_row(df)
    headers = [str(df.iloc[hr, j]).strip() for j in range(df.shape[1])]
    cmap = map_headers_to_fields(headers)
    if "材质" not in cmap:
        raise ValueError(f"未识别到「材质」列，请检查表头或使用 Excel 导出后再导入。\n识别的表头：{headers}")

    out_rows: List[Dict[str, Any]] = []
    for r in range(hr + 1, len(df)):
        line = " ".join(str(x) for x in df.iloc[r].tolist() if str(x).strip() and str(x) != "nan")
        if not line.strip():
            continue
        mv = _cell_at(df, r, cmap["材质"])
        if mv is None or str(mv).strip() == "":
            continue

        br = _blank_row(len(out_rows) + 1)
        if "料品名称" in cmap:
            v = _cell_at(df, r, cmap["料品名称"])
            if v is not None:
                br["料品名称"] = str(v).strip()

        br["材质"] = normalize_material_name(mv)

        if "厚度(mm)" in cmap:
            v = _cell_at(df, r, cmap["厚度(mm)"])
            fv, ok = _to_float(v, 0.0)
            if ok and fv > 0:
                br["厚度(mm)"] = fv
        if "长(mm)" in cmap and "宽(mm)" in cmap:
            la, ok1 = _to_float(_cell_at(df, r, cmap["长(mm)"]), 0.0)
            wi, ok2 = _to_float(_cell_at(df, r, cmap["宽(mm)"]), 0.0)
            if ok1 and la > 0:
                br["长(mm)"] = la
            if ok2 and wi > 0:
                br["宽(mm)"] = wi
        elif "零件尺寸" in cmap:
            length_mm, width_mm = _parse_dim_two(_cell_at(df, r, cmap["零件尺寸"]))
            if length_mm and width_mm:
                br["长(mm)"] = length_mm
                br["宽(mm)"] = width_mm

        for key, col in [
            ("米数(切割长度)", "米数(切割长度)"),
            ("穿孔数量", "穿孔数量"),
            ("沉孔数量", "沉孔数量"),
            ("折弯次数", "折弯次数"),
            ("压铆数量", "压铆数量"),
            ("攻丝数量", "攻丝数量"),
            ("焊接长度(mm)", "焊接长度(mm)"),
        ]:
            if key in cmap:
                v = _cell_at(df, r, cmap[key])
                fv, ok = _to_float(v, 0.0)
                if ok:
                    br[col] = fv
        if "表面处理" in cmap:
            sv = _cell_at(df, r, cmap["表面处理"])
            if sv is not None:
                br["表面处理"] = normalize_surface_name(sv)
        out_rows.append(br)

    if not out_rows:
        raise ValueError("未解析到有效数据行（可能表头未识别，请尝试导出为 Excel 再导入）。")

    for i, row in enumerate(out_rows, 1):
        row["序号"] = str(i)
    qdf = pd.DataFrame(out_rows)
    return sync_material_unit_price(migrate_dataframe(qdf))


def parse_parts_upload(file_bytes: bytes, name: str) -> pd.DataFrame:
    ext = Path(name).suffix.lower()
    raw = ocr_image_to_dataframe(file_bytes) if ext in (".png", ".jpg", ".jpeg") else _load_sheet_bytes(name, file_bytes)
    return build_quote_from_sheet_df(raw)
