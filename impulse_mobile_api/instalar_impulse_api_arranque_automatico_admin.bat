@echo off
cd /d %~dp0
powershell -ExecutionPolicy Bypass -File "%~dp0install_impulse_api_startup_task.ps1" -OpenFirewall
pause
