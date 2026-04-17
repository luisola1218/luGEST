@echo off
setlocal
cd /d "%~dp0"

where flutter >nul 2>nul
if errorlevel 1 (
  echo [ERRO] Flutter nao encontrado no PATH.
  echo Instala o Flutter SDK e reinicia o terminal.
  exit /b 1
)

echo [1/2] A instalar dependencias...
flutter pub get
if errorlevel 1 (
  echo [ERRO] Falha no flutter pub get.
  exit /b 1
)

echo [2/2] A arrancar app em DEBUG com hot reload...
echo Dica: no VS Code, ao guardar, o app atualiza automaticamente.
flutter run --debug

endlocal
