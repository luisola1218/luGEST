$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$desktopRoot = Join-Path $env:USERPROFILE 'Desktop'
$releaseDate = Get-Date
$releaseStamp = $releaseDate.ToString('yyyyMMdd')
$releaseDateTxt = $releaseDate.ToString('dd/MM/yyyy HH:mm')
$releaseName = 'App LuisGEST - Revis' + [char]0x00E3 + 'o Final'
$releaseRoot = Join-Path $desktopRoot $releaseName

$desktopEnvExample = Join-Path $repoRoot 'lugest.env.example'
$desktopServerEnvExample = Join-Path $repoRoot 'lugest.env.servidor.example'
$desktopPostEnvExample = Join-Path $repoRoot 'lugest.env.posto.example'
$desktopBranding = Join-Path $repoRoot 'lugest_branding.json'
$desktopQtConfig = Join-Path $repoRoot 'lugest_qt_config.json'
$desktopIcon = Join-Path $repoRoot 'app.ico'
$desktopLogo = Join-Path $repoRoot 'logo.jpg'
$desktopLogosDir = Join-Path $repoRoot 'Logos'
$desktopAdminSetup = Join-Path $repoRoot 'scripts\setup_lugest_admin.ps1'
$apiSource = Join-Path $repoRoot 'impulse_mobile_api'
$mobileSource = Join-Path $repoRoot 'impulse_mobile_app'
$databaseSource = Join-Path $repoRoot 'mysql'
$securityPlan = Join-Path $repoRoot 'SECURITY_TEST_PLAN.md'
$installGuide = Join-Path $repoRoot 'GUIA_INSTALACAO_OUTRO_PC.md'
$fullInstallGuide = Join-Path $repoRoot 'GUIA_INSTALACAO_TOTAL.md'
$simpleGuide = Join-Path $repoRoot 'GUIA_MUITO_SIMPLES.md'
$billingPlan = Join-Path $repoRoot 'FATURACAO_PLAN.md'

function Resolve-MobileApkPath {
    $candidates = @()

    $repoApk = Join-Path $repoRoot 'impulse_mobile_app\build\app\outputs\flutter-apk\app-release.apk'
    if (Test-Path $repoApk) {
        $candidates += Get-Item $repoApk
    }

    $candidates = $candidates | Sort-Object LastWriteTime -Descending
    if ($candidates) {
        return $candidates[0].FullName
    }
    return $null
}

function Resolve-DesktopExePath {
    $candidates = @()

    foreach ($relativePath in @(
        'dist_dashboard_release\main.exe',
        'dist_transportes_tarifario_release\main.exe',
        'dist_transportes_pro_release\main.exe',
        'dist_transportes_ui_release\main.exe',
        'dist_transportes_release\main.exe',
        'dist_multifix_release\main.exe',
        'dist_multinet_release\main.exe',
        'dist_billing_release\main.exe',
        'dist\main.exe',
        'dist_pack\main.exe'
    )) {
        $candidate = Join-Path $repoRoot $relativePath
        if (Test-Path $candidate) {
            $candidates += Get-Item $candidate
        }
    }

    $candidates = $candidates | Sort-Object LastWriteTime -Descending
    if ($candidates) {
        return $candidates[0].FullName
    }
    return $null
}

$desktopExe = Resolve-DesktopExePath
if (-not $desktopExe) {
    throw "Nao foi encontrado nenhum main.exe valido para a release."
}

$mobileApk = Resolve-MobileApkPath
$mobileApkSource = $null
if ($mobileApk) {
    if (-not (Test-Path $mobileApk)) {
        throw "APK invalida ou inacessivel: $mobileApk"
    }
    $mobileApkSource = $mobileApk
    if ($mobileApkSource -like (Join-Path $desktopRoot 'App LuisGEST*')) {
        $tempApk = Join-Path $env:TEMP ("lugest_mobile_release_" + $releaseStamp + ".apk")
        Copy-Item $mobileApkSource $tempApk -Force
        $mobileApkSource = $tempApk
    }
}

foreach ($requiredPath in @(
    $desktopExe,
    $desktopEnvExample,
    $desktopServerEnvExample,
    $desktopPostEnvExample,
    $desktopBranding,
    $desktopQtConfig,
    $desktopIcon,
    $desktopLogo,
    $desktopLogosDir,
    $desktopAdminSetup,
    $apiSource,
    $mobileSource,
    $databaseSource,
    $securityPlan,
    $installGuide,
    $fullInstallGuide,
    $simpleGuide,
    $billingPlan
)) {
    if (-not (Test-Path $requiredPath)) {
        throw "Falta ficheiro/pasta obrigatoria para a release: $requiredPath"
    }
}

Get-ChildItem $desktopRoot -ErrorAction SilentlyContinue |
    Where-Object { $_.PSIsContainer -and $_.Name -like 'App LuisGEST*' } |
    ForEach-Object { Remove-Item $_.FullName -Recurse -Force }

if (Test-Path $releaseRoot) {
    Remove-Item $releaseRoot -Recurse -Force
}

$desktopDir = Join-Path $releaseRoot 'Desktop App'
$apiDir = Join-Path $releaseRoot 'Mobile API'
$mobileApkDir = Join-Path $releaseRoot 'Mobile APK'
$mobileSrcDir = Join-Path $releaseRoot 'Mobile App Fonte'
$dbDir = Join-Path $releaseRoot 'Base de Dados'

New-Item -ItemType Directory -Force -Path $desktopDir, $apiDir, $mobileApkDir, $mobileSrcDir, $dbDir | Out-Null

Copy-Item $desktopExe $desktopDir
Copy-Item $desktopEnvExample (Join-Path $desktopDir 'lugest.env.example')
Copy-Item $desktopEnvExample (Join-Path $desktopDir 'lugest.env')
Copy-Item $desktopServerEnvExample (Join-Path $desktopDir 'lugest.env.servidor.example')
Copy-Item $desktopPostEnvExample (Join-Path $desktopDir 'lugest.env.posto.example')
Copy-Item $desktopBranding $desktopDir
Copy-Item $desktopQtConfig $desktopDir
Copy-Item $desktopIcon $desktopDir
Copy-Item $desktopLogo $desktopDir
Copy-Item $desktopAdminSetup (Join-Path $desktopDir 'setup_lugest_admin.ps1')
Copy-Item -Recurse $desktopLogosDir $desktopDir

$trialTemplate = [ordered]@{
    enabled = $false
    company_name = ''
    device_fingerprint = ''
    started_at = ''
    duration_days = 60
    created_at = ''
    created_by = ''
    updated_at = ''
    updated_by = ''
    last_success_at = ''
    last_success_user = ''
    last_owner_auth_at = ''
    last_owner_auth_user = ''
    notes = ''
}
$trialTemplate | ConvertTo-Json -Depth 4 | Set-Content -Path (Join-Path $desktopDir 'lugest_trial.json') -Encoding UTF8

$desktopLauncher = @'
@echo off
cd /d %~dp0
if not exist lugest.env (
    echo Falta o ficheiro lugest.env.
    echo Copia lugest.env.example para lugest.env e ajusta os parametros MySQL antes de arrancar.
    pause
    exit /b 1
)
start "" "%~dp0main.exe"
'@
Set-Content -Path (Join-Path $desktopDir 'Arrancar LuisGEST Desktop.bat') -Value $desktopLauncher -Encoding ASCII

$desktopAdminLauncher = @'
@echo off
cd /d %~dp0
powershell -ExecutionPolicy Bypass -File "%~dp0setup_lugest_admin.ps1"
pause
'@
Set-Content -Path (Join-Path $desktopDir 'Criar Administrador Inicial.bat') -Value $desktopAdminLauncher -Encoding ASCII

$desktopAdminResetLauncher = @'
@echo off
cd /d %~dp0
powershell -ExecutionPolicy Bypass -File "%~dp0setup_lugest_admin.ps1" -Reset
pause
'@
Set-Content -Path (Join-Path $desktopDir 'Repor Password Administrador.bat') -Value $desktopAdminResetLauncher -Encoding ASCII

$desktopAdminReadme = @"
ADMIN LOCAL DO LUGEST
=====================

O login OWNER no lugest.env nao e o admin normal da aplicacao.
O OWNER serve para trial/licenciamento.

Para criar o admin local da app:
1. Configura primeiro o ficheiro lugest.env com os dados MySQL corretos.
2. Corre Criar Administrador Inicial.bat
3. Abre o LuGEST e entra com esse username/password

Se precisares de trocar a password do admin:
- corre Repor Password Administrador.bat

Se importares lugest_instalacao_unica.sql, existe tambem um login temporario:
- admin / Trocar#Admin2026

Troca essa password logo apos a instalacao.
"@
Set-Content -Path (Join-Path $desktopDir 'LEIA-ME - ADMIN.txt') -Value $desktopAdminReadme -Encoding UTF8

robocopy $apiSource $apiDir /E /XD '.venv' '__pycache__' 'app\__pycache__' 'app\services\__pycache__' /XF '.env' 'api_runtime.log' 'api_stdout.log' 'api_stderr.log' '.server.pid' | Out-Null
if ($LASTEXITCODE -gt 7) {
    throw "Falha a copiar Mobile API (robocopy exit $LASTEXITCODE)"
}
Copy-Item (Join-Path $apiDir '.env.example') (Join-Path $apiDir '.env')

robocopy $mobileSource $mobileSrcDir /E /XD 'build' '.dart_tool' '.idea' '.vscode' 'android\.gradle' 'android\.kotlin' /XF '.flutter-plugins-dependencies' '*.iml' 'local.properties' | Out-Null
if ($LASTEXITCODE -gt 7) {
    throw "Falha a copiar Mobile App Fonte (robocopy exit $LASTEXITCODE)"
}
foreach ($generatedDir in @(
    (Join-Path $mobileSrcDir 'android\.gradle'),
    (Join-Path $mobileSrcDir 'android\.kotlin')
)) {
    if (Test-Path $generatedDir) {
        Remove-Item $generatedDir -Recurse -Force
    }
}
if ($mobileApkSource) {
    Copy-Item $mobileApkSource (Join-Path $mobileApkDir ("LuisGEST-Impulse-release-" + $releaseStamp + ".apk"))
}
else {
    $apkPending = @"
Nao foi encontrada nenhuma APK release local para copiar.

Esta entrega segue com o codigo fonte Flutter, mas a APK precisa de ser gerada quando existir Flutter SDK neste posto:

    cd impulse_mobile_app
    flutter pub get
    flutter build apk --release

Opcional para deixar a APK ja apontada para um servidor especifico:

    flutter build apk --release --dart-define=LUGEST_DEFAULT_API_HOST=http://IP_DO_SERVIDOR:8050
"@
    Set-Content -Path (Join-Path $mobileApkDir 'APK-PENDENTE.txt') -Value $apkPending -Encoding UTF8
}

Copy-Item -Recurse $databaseSource $dbDir
Get-ChildItem -Path $dbDir -Recurse -Directory -Filter '__pycache__' -ErrorAction SilentlyContinue |
    ForEach-Object { Remove-Item $_.FullName -Recurse -Force -ErrorAction SilentlyContinue }
Copy-Item $securityPlan (Join-Path $releaseRoot 'CHECKLIST - Seguranca e Testes.md')
Copy-Item $fullInstallGuide (Join-Path $releaseRoot 'GUIA - INSTALACAO TOTAL PASSO A PASSO.md')
Copy-Item $installGuide (Join-Path $releaseRoot 'GUIA - Instalacao Outro Computador.md')
Copy-Item $simpleGuide (Join-Path $releaseRoot 'GUIA - MUITO SIMPLES.md')
Copy-Item $billingPlan (Join-Path $releaseRoot 'PLANO - Faturacao.md')

$readme = @"
# App LuisGEST - Revisao Final

Data de preparacao: $releaseDateTxt

## Conteudo
- Desktop App: executavel principal main.exe, launcher Qt, lugest.env, branding, lugest_qt_config.json e lugest_trial.json placeholder.
- Mobile API: API FastAPI com .env placeholder e scripts de instalacao/arranque.
- Mobile APK: APK Android release atual, quando disponivel; caso contrario segue uma nota APK-PENDENTE.txt.
- Mobile App Fonte: codigo Flutter limpo para manutencao futura.
- Base de Dados\mysql: dump principal lugest.sql, instalacao unica e patches auxiliares.
- GUIA - INSTALACAO TOTAL PASSO A PASSO.md: guia principal para instalar servidor e postos do zero.
- CHECKLIST - Seguranca e Testes.md: guia de validacao antes da entrega.

## Comecar por aqui
1. Se nunca fizeste uma instalacao, le primeiro GUIA - INSTALACAO TOTAL PASSO A PASSO.md.
2. No servidor, usar Desktop App\lugest.env.servidor.example como base do lugest.env.
3. Nos postos, usar Desktop App\lugest.env.posto.example como base do lugest.env.

## Arranque desktop
1. No servidor, usar Desktop App\lugest.env.servidor.example como base do lugest.env.
2. Nos postos, usar Desktop App\lugest.env.posto.example como base do lugest.env.
3. Se o ambiente for multi-posto, definir tambem LUGEST_SHARED_STORAGE_ROOT com uma pasta UNC partilhada para desenhos, PDFs e anexos.
4. Se a base estiver sem utilizadores locais, correr Desktop App\Criar Administrador Inicial.bat.
5. Ajustar Desktop App\lugest_branding.json, Desktop App\lugest_qt_config.json e os logos se quiser personalizacao.
6. Se fores usar trial, editar/gerir Desktop App\lugest_trial.json apenas na maquina final, nunca copiando o trial desta maquina.
7. Executar Desktop App\Arrancar LuisGEST Desktop.bat.
8. O desktop arranca sempre em Qt e trabalha apenas com MySQL.

## Arranque API mobile
1. Entrar em Mobile API.
2. Editar .env com os dados reais MySQL e segredo API.
3. Executar instalar_impulse_api.bat.
4. Executar arrancar_impulse_api.bat.
5. Para producao em Windows Server, usar instalar_impulse_api_arranque_automatico_admin.bat.

## Mobile Android
1. Se existir Mobile APK\LuisGEST-Impulse-release-$releaseStamp.apk, instalar essa APK.
2. Se a pasta trouxer APK-PENDENTE.txt, gerar primeiro a APK a partir de Mobile App Fonte.
3. No primeiro login indicar o IP/URL do servidor API.
4. A app guarda o ultimo servidor e utilizador usados.

## Base de dados
- Importar Base de Dados\mysql\lugest.sql no MySQL da empresa.
- Em instalacao nova com logins temporarios e arranque rapido, podes importar Base de Dados\mysql\lugest_instalacao_unica.sql.
- Aplicar patches adicionais da mesma pasta se o ambiente os exigir.
- Para automatizar a instalacao inicial, usar Base de Dados\mysql\install_lugest_mysql.ps1 ou o atalho instalar_lugest_mysql_admin.bat.

## Notas finais
- py main.py e main.exe abrem a desktop Qt principal.
- O desktop suporta um ficheiro externo lugest.env, por isso nao e preciso alterar codigo para mudar empresa/servidor.
- Em multi-posto, usa LUGEST_SHARED_STORAGE_ROOT com um caminho UNC comum a todos os postos.
- O login OWNER no lugest.env serve para trial/licenciamento; o admin normal da app cria-se ou repoe-se pelos scripts da pasta Desktop App.
- Usa uma conta MySQL dedicada para a aplicacao em vez de root.
- O menu Faturacao faz seguimento de vendidos, faturas, pagamentos e comprovativos.
- O menu Transportes faz agendamento de viagens, afeta encomendas e gera folha de rota PDF.
- Este pacote nao leva os segredos da maquina atual: lugest.env, .env da API e o trial seguem em modo placeholder.
- Depois de instalar, usa CHECKLIST - Seguranca e Testes.md para validar o ambiente novo.
"@
Set-Content -Path (Join-Path $releaseRoot 'README - Revisao Final.md') -Value $readme -Encoding UTF8

Write-Output $releaseRoot
