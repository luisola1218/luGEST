@echo off
cd /d %~dp0
echo Ajuste primeiro os parametros no comando abaixo se necessario.
powershell -ExecutionPolicy Bypass -File "%~dp0backup_lugest_mysql.ps1" -DbHost 127.0.0.1 -Port 3306 -Database lugest
pause
