@echo off
setlocal
cd /d "%~dp0"

where adb >nul 2>nul
if errorlevel 1 (
  echo [ERRO] adb nao encontrado no PATH.
  echo Instala Android Platform Tools e reinicia o terminal.
  exit /b 1
)

where flutter >nul 2>nul
if errorlevel 1 (
  echo [ERRO] Flutter nao encontrado no PATH.
  echo Instala o Flutter SDK e reinicia o terminal.
  exit /b 1
)

echo [PASSO 1] Liga o telemovel por USB e confirma depuracao USB ativa.
adb devices

set /p DEVICE_IP=IP do telemovel na rede local (ex: 192.168.1.120): 
if "%DEVICE_IP%"=="" (
  echo [ERRO] IP em falta.
  exit /b 1
)

set "PORT=5555"
set /p PORT=Porta ADB [5555]: 
if "%PORT%"=="" set "PORT=5555"

echo [PASSO 2] A ativar ADB por Wi-Fi...
adb tcpip %PORT%

echo [PASSO 3] A ligar ao telemovel por Wi-Fi...
adb connect %DEVICE_IP%:%PORT%
if errorlevel 1 (
  echo [ERRO] Falha na ligacao ADB Wi-Fi.
  echo Verifica se PC e telemovel estao na mesma rede.
  exit /b 1
)

echo [PASSO 4] A instalar dependencias...
flutter pub get
if errorlevel 1 (
  echo [ERRO] Falha no flutter pub get.
  exit /b 1
)

echo [PASSO 5] A arrancar em DEBUG com hot reload no device %DEVICE_IP%:%PORT%...
echo Ao guardar ficheiros no VS Code, o reload e automatico.
flutter run --debug -d %DEVICE_IP%:%PORT%

endlocal
