@echo off
cd /d %~dp0
if exist .venv\Scripts\python.exe (
    .venv\Scripts\python.exe -c "import PySide6" >nul 2>nul
    if errorlevel 1 (
        echo [ERRO] A .venv local existe, mas nao tem o ambiente Qt pronto.
        echo.
        echo Execute:
        echo   .venv\Scripts\python.exe -m pip install --upgrade pip
        echo   .venv\Scripts\python.exe -m pip install -r requirements-qt.txt
        exit /b 1
    )
    start "" .venv\Scripts\python.exe main.py
    exit /b 0
)
if exist dist_qt_stable\lugest_qt\lugest_qt.exe (
    start "" dist_qt_stable\lugest_qt\lugest_qt.exe
    exit /b 0
)
if exist dist\lugest_qt\lugest_qt.exe (
    start "" dist\lugest_qt\lugest_qt.exe
    exit /b 0
)
if exist dist\main.exe (
    start "" dist\main.exe
    exit /b 0
)
if exist main.exe (
    start "" main.exe
    exit /b 0
)
py main.py
