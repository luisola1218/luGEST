@echo off
cd /d %~dp0
powershell -ExecutionPolicy Bypass -File "%~dp0status_impulse_api_startup_task.ps1"
pause
