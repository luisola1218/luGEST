@echo off
cd /d %~dp0
powershell -ExecutionPolicy Bypass -File "%~dp0remove_impulse_api_startup_task.ps1" -RemoveFirewall
pause
