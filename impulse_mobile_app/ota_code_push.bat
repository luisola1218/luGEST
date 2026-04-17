@echo off
setlocal
cd /d "%~dp0"

where shorebird >nul 2>nul
if errorlevel 1 (
  echo [ERRO] Shorebird CLI nao encontrado.
  echo Instalar: dart pub global activate shorebird_cli
  echo Depois: shorebird login
  exit /b 1
)

if "%~1"=="" goto usage

if /I "%~1"=="release" goto release
if /I "%~1"=="patch" goto patch
goto usage

:release
echo [OTA] Primeira release base para permitir patches OTA.
echo [OTA] Este passo deve ser feito uma vez por versao base.
shorebird release android --artifact apk
exit /b %errorlevel%

:patch
echo [OTA] A publicar patch Dart/UI sem nova APK.
echo [OTA] Limite: alteracoes nativas (Android/iOS/plugins) exigem nova release.
shorebird patch android
exit /b %errorlevel%

:usage
echo Uso:
echo   ota_code_push.bat release   ^(primeira release base^)
echo   ota_code_push.bat patch     ^(publicar patch OTA^)
exit /b 1

endlocal
