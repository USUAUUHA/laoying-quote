@echo off
setlocal

set "ROOT=%~dp0.."
cd /d "%ROOT%"
if errorlevel 1 (
  echo [ERROR] Cannot enter project root: "%ROOT%"
  pause
  exit /b 1
)

where python >nul 2>nul
if errorlevel 1 (
  echo [ERROR] python is not found in PATH.
  echo Please install Python or add python.exe to PATH.
  pause
  exit /b 1
)

set "STREAMLIT_BROWSER_GATHER_USAGE_STATS=false"
python -m pip install streamlit pandas openpyxl
if errorlevel 1 (
  echo [ERROR] pip install failed.
  pause
  exit /b 1
)

set "PORT="
for %%P in (8512 8513 8514 8515) do (
  netstat -ano | findstr ":%%P" >nul
  if errorlevel 1 (
    set "PORT=%%P"
    goto :port_ready
  )
)

echo [ERROR] No available port in 8512/8513/8514/8515.
echo Please close one occupied port or edit launcher port list.
pause
exit /b 1

:port_ready
echo [INFO] Starting Streamlit on port %PORT% ...
python -m streamlit run "app_home.py" --server.port %PORT%
if errorlevel 1 (
  echo [ERROR] Streamlit start failed.
)
pause
