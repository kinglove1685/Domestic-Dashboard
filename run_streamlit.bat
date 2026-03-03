@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
set "PORT=8501"

if not exist ".venv\Scripts\python.exe" (
  echo [INFO] Creating virtual environment...
  python -m venv .venv
)

call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt

netstat -ano | findstr /R /C:":%PORT% .*LISTENING" >nul
if %errorlevel%==0 (
  echo [INFO] Streamlit is already running on http://localhost:%PORT%
  start "" "http://localhost:%PORT%"
  goto :end
)

echo [INFO] Starting Streamlit...
start "" "http://localhost:%PORT%"
streamlit run app.py --server.headless false --server.port %PORT%

:end
endlocal
