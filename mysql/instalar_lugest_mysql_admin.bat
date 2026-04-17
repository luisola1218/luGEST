@echo off
cd /d %~dp0
echo Ajuste primeiro os parametros no comando abaixo se necessario.
powershell -ExecutionPolicy Bypass -File "%~dp0install_lugest_mysql.ps1" -DbHost 127.0.0.1 -Port 3306 -AdminUser root -Database lugest -AppUser lugest_user -AppHost localhost -ResetDatabase
pause
