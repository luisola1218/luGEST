@echo off
cd /d %~dp0
if not exist .venv\Scripts\python.exe (
  python -m venv .venv
)
call .venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
python verify_installation.py
if errorlevel 1 (
  echo.
  echo [AVISO] A instalacao terminou, mas o preflight falhou. Corrija o ficheiro .env antes de arrancar a API.
  pause
  exit /b 1
)
echo.
echo [OK] API instalada. Edite o ficheiro .env se necessario e execute arrancar_impulse_api.bat
pause
