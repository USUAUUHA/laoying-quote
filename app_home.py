"""
ASCII entry point for Windows batch launcher.

Purpose:
- Avoid Chinese file name issues in cmd/bat environments.
- Delegate to the real homepage script without changing business logic.
"""

from __future__ import annotations

from pathlib import Path
import runpy


ROOT = Path(__file__).resolve().parent
HOME_SCRIPT = ROOT / "报价系统首页.py"

if not HOME_SCRIPT.is_file():
    raise FileNotFoundError(f"Homepage script not found: {HOME_SCRIPT}")

runpy.run_path(str(HOME_SCRIPT), run_name="__main__")
