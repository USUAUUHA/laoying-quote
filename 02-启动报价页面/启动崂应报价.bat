@echo off
setlocal

rem Avoid hard-coded Chinese path names; discover script folder automatically.
set "ROOT=%~dp0.."
set "SCRIPT_DIR="
for /d %%I in ("%ROOT%\01-*") do (
  set "SCRIPT_DIR=%%~fI"
  goto :found_dir
)

echo [ERROR] Cannot find script directory under "%ROOT%".
echo Expected a folder like "01-...".
pause
exit /b 1

:found_dir
cd /d "%SCRIPT_DIR%"
if errorlevel 1 (
  echo [ERROR] Cannot enter script directory: "%SCRIPT_DIR%"
  pause
  exit /b 1
)

set "STREAMLIT_BROWSER_GATHER_USAGE_STATS=false"
if exist "requirements_streamlit.txt" (
  python -m pip install -r "requirements_streamlit.txt"
) else (
  python -m pip install streamlit pandas openpyxl
)

python -m streamlit run "sheet_metal_quote_full.py" --server.port 8511
pause
