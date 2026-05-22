from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, List

APP_TITLE = "青岛宏泰铭润机械 · 钣金批量智能报价系统"
APP_SUBTITLE = "支持Excel/柏楚激光软件数据整列复制粘贴，每行自动独立计算报价"

PROJECT_ROOT = Path(__file__).resolve().parent.parent
EXPORT_REFERENCE_ROOT = PROJECT_ROOT / "03-导出报价文件"
EXPORT_FOLDER_SUFFIX = "-给客户报价-盖章文件对比"
EXPORT_LY_NAME_CORE = "崂应-给客户报价-盖章"

SEAL_IMAGE_DIR = PROJECT_ROOT / "04-盖章图"
SEAL_IMAGE_PATH = SEAL_IMAGE_DIR / "盖章.png"
SEAL_IMAGE_ALT = SEAL_IMAGE_DIR / "seal.png"
SEAL_EXPORT_WIDTH_PX = 180
SEAL_EXPORT_HEIGHT_PX = 180

PRICE_PARAMS_FILE = PROJECT_ROOT / "price_params.json"
QUOTE_DF_SCHEMA_VERSION = "2026-03-25-core-v1"
DEFAULT_WEIGHT_COEFFICIENT = 1.3


def get_default_price_params() -> Dict[str, float]:
    return {
        "cutting_rate": 1.2,
        "grinding_small": 0.4,
        "grinding_large": 0.5,
        "grinding_mult": 1.13,
        "pierce_per_t": 0.1,
        "weld_a": 40.0,
        "weld_b": 20.0,
        "countersink": 0.3,
        "bend_al": 0.5,
        "bend_steel": 0.6,
        "bend_large": 0.7,
        "pem_a": 0.3,
        "pem_b": 0.15,
        "tap_al": 0.4,
        "tap_steel": 0.6,
    }


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


def load_price_params() -> Dict[str, float]:
    if not PRICE_PARAMS_FILE.is_file():
        return get_default_price_params()
    try:
        raw = json.loads(PRICE_PARAMS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return get_default_price_params()
    return _sanitize_price_params(raw)


def save_price_params(price_params: Dict[str, float]) -> None:
    safe = _sanitize_price_params(price_params)
    PRICE_PARAMS_FILE.write_text(json.dumps(safe, ensure_ascii=False, indent=2), encoding="utf-8")


SURFACE_OPTIONS = ["无", "喷粉", "喷砂", "氧化", "抛光", "镀锌"]

MATERIAL_LIBRARY: Dict[str, Dict[str, Any]] = {
    "6061铝板": {"density": 2.73, "unit_price": 32.0, "category": "铝板"},
    "1060铝板": {"density": 2.73, "unit_price": 27.0, "category": "铝板"},
    "5052铝板": {"density": 2.73, "unit_price": 27.0, "category": "铝板"},
    "SUS304": {"density": 7.93, "unit_price": 22.0, "category": "不锈钢"},
    "SUS316": {"density": 7.93, "unit_price": 40.0, "category": "不锈钢"},
    "304 双面拉丝": {"density": 7.93, "unit_price": 22.0, "category": "不锈钢"},
    "拉丝板304": {"density": 7.93, "unit_price": 22.0, "category": "不锈钢"},
    "SUS201": {"density": 7.93, "unit_price": 22.0, "category": "不锈钢"},
    "316覆膜板": {"density": 7.93, "unit_price": 22.0, "category": "不锈钢"},
    "430拉丝板": {"density": 7.93, "unit_price": 13.0, "category": "不锈钢"},
    "SUS430": {"density": 7.93, "unit_price": 13.0, "category": "不锈钢"},
    "马氏体沉淀硬化不锈钢17-4 PH": {"density": 7.93, "unit_price": 22.0, "category": "不锈钢"},
    "304抛光板": {"density": 7.93, "unit_price": 22.0, "category": "不锈钢"},
    "SGCC": {"density": 7.85, "unit_price": 8.0, "category": "碳钢/镀锌板"},
    "Q235": {"density": 7.85, "unit_price": 8.0, "category": "碳钢/镀锌板"},
    "冷板": {"density": 7.85, "unit_price": 8.0, "category": "碳钢/镀锌板"},
}

COL_MATERIAL_PRICE = "材料单价(元/kg)"
COL_WEIGHT_AFTER_THICK = "重量(kg)"

INPUT_COLS_BEFORE_WEIGHT: List[str] = [
    "料品名称",
    "料号",
    "料品规格",
    "材质",
    COL_MATERIAL_PRICE,
    "长(mm)",
    "宽(mm)",
    "厚度(mm)",
]

INPUT_COLS_AFTER_WEIGHT: List[str] = [
    "米数(切割长度)",
    "穿孔数量",
    "焊接长度(mm)",
    "沉孔数量",
    "折弯次数",
    "压铆数量",
    "攻丝数量",
    "激光打标",
    "表面处理",
    "加工数量",
]

INPUT_COLS = INPUT_COLS_BEFORE_WEIGHT + [COL_WEIGHT_AFTER_THICK] + INPUT_COLS_AFTER_WEIGHT

CALC_COLS = [
    "总重(kg)",
    "打磨(元)",
    "材料费(元)",
    "总加工费(元)",
    "表面处理费(元)",
    "含税单价(元/件)",
    "含税总价(元)",
]

MANDATORY_COLS = ["料品名称", "材质", "长(mm)", "宽(mm)", "厚度(mm)", "表面处理", "加工数量"]

NUMERIC_INPUT_COLS = [
    "长(mm)",
    "宽(mm)",
    "厚度(mm)",
    "米数(切割长度)",
    "穿孔数量",
    "焊接长度(mm)",
    "沉孔数量",
    "折弯次数",
    "压铆数量",
    "攻丝数量",
    "激光打标",
    "加工数量",
]

COLS_RECALC_ONLY = [COL_WEIGHT_AFTER_THICK] + CALC_COLS
TABLE_COLUMNS = ["选择", "序号"] + INPUT_COLS + CALC_COLS
