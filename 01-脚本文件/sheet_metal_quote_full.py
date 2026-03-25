from __future__ import annotations

"""
青岛宏泰铭润机械 · 钣金批量智能报价系统

兼容入口：
  streamlit run "G:\\workspace\\崂应报价系统\\01-脚本文件\\sheet_metal_quote_full.py"
"""

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from core.app import main


if __name__ == "__main__":
    main()
