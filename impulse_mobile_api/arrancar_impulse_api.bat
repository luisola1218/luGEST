@echo off
cd /d %~dp0
if not exist .venv\Scripts\python.exe (
  echo [ERRO] Crie primeiro a venv em impulse_mobile_api\.venv e instale requirements.
  pause
  exit /b 1
)
call .venv\Scripts\activate
python verify_installation.py
if errorlevel 1 (
  echo.
  echo [ERRO] O preflight falhou. Reveja o ficheiro .env e a ligacao MySQL antes de arrancar a API.
  pause
  exit /b 1
)
python run_server.py
