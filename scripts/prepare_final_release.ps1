$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$desktopRoot = Join-Path $env:USERPROFILE 'Desktop'
$releaseDate = Get-Date
$releaseStamp = $releaseDate.ToString('yyyyMMdd')
$releaseDateTxt = $releaseDate.ToString('dd/MM/yyyy HH:mm')
$releaseName = 'App LuisGEST - Revis' + [char]0x00E3 + 'o Final'
$releaseRoot = Join-Path $desktopRoot $releaseName

$desktopEnvExample = Join-Path $repoRoot 'config\examples\lugest.env.example'
$desktopServerEnvExample = Join-Path $repoRoot 'config\examples\lugest.env.servidor.example'
$desktopPostEnvExample = Join-Path $repoRoot 'config\examples\lugest.env.posto.example'
$desktopBranding = Join-Path $repoRoot 'lugest_branding.json'
$desktopQtConfig = Join-Path $repoRoot 'lugest_qt_config.json'
$desktopIcon = Join-Path $repoRoot 'app.ico'
$desktopLogo = Join-Path $repoRoot 'logo.jpg'
$desktopLogosDir = Join-Path $repoRoot 'Logos'
$desktopAdminSetup = Join-Path $repoRoot 'scripts\setup_lugest_admin.ps1'
$desktopEnvDiagnose = Join-Path $repoRoot 'scripts\diagnose_lugest_env.ps1'
$desktopStartLauncher = Join-Path $repoRoot 'scripts\start_lugest_desktop.ps1'
$desktopStartLauncherHidden = Join-Path $repoRoot 'scripts\start_lugest_desktop_hidden.vbs'
$desktopUpdaterScript = Join-Path $repoRoot 'scripts\lugest_update.ps1'
$desktopRepairUpdaterScript = Join-Path $repoRoot 'scripts\repair_installed_updater.ps1'
$desktopInstallerScript = Join-Path $repoRoot 'scripts\install_lugest_desktop.ps1'
$desktopVersionFile = Join-Path $repoRoot 'VERSION'
$apiSource = Join-Path $repoRoot 'impulse_mobile_api'
$databaseSource = Join-Path $repoRoot 'mysql'
$securityPlan = Join-Path $repoRoot 'docs\plans\SECURITY_TEST_PLAN.md'
$installGuide = Join-Path $repoRoot 'docs\install\GUIA_INSTALACAO_OUTRO_PC.md'
$fullInstallGuide = Join-Path $repoRoot 'docs\install\GUIA_INSTALACAO_TOTAL.md'
$simpleGuide = Join-Path $repoRoot 'docs\install\GUIA_MUITO_SIMPLES.md'
$emailUpdatesGuide = Join-Path $repoRoot 'docs\install\GUIA_EMAIL_E_ATUALIZACOES.md'

function Write-Utf8NoBomFile {
    param(
        [string]$Path,
        [string]$Content
    )
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Content, $utf8NoBom)
}

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
    foreach ($relativePath in @(
        'dist_qt_stable\lugest_qt\lugest_qt.exe',
        'dist\lugest_qt\lugest_qt.exe',
        'dist\lugest_qt.exe',
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
            return (Get-Item $candidate).FullName
        }
    }
    return $null
}

$desktopExe = Resolve-DesktopExePath
if (-not $desktopExe) {
    throw "Nao foi encontrado nenhum executavel desktop valido para a release."
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
    $desktopEnvDiagnose,
    $desktopStartLauncher,
    $desktopStartLauncherHidden,
    $desktopUpdaterScript,
    $desktopInstallerScript,
    $desktopVersionFile,
    $apiSource,
    $databaseSource,
    $securityPlan,
    $installGuide,
    $fullInstallGuide,
    $simpleGuide,
    $emailUpdatesGuide
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
$dbDir = Join-Path $releaseRoot 'Base de Dados'
$updatesDir = Join-Path $releaseRoot 'Atualizacoes'
$docsDir = Join-Path $releaseRoot 'Documentacao'

New-Item -ItemType Directory -Force -Path $desktopDir, $apiDir, $mobileApkDir, $dbDir, $updatesDir, $docsDir | Out-Null

$desktopExeItem = Get-Item $desktopExe
$desktopExeName = $desktopExeItem.Name
$desktopExeParent = $desktopExeItem.Directory.FullName
$desktopConsoleExe = Join-Path $repoRoot 'dist\main.exe'
if ($desktopExeParent -like (Join-Path $repoRoot 'dist\lugest_qt') -or $desktopExeParent -like (Join-Path $repoRoot 'dist_qt_stable\lugest_qt')) {
    Copy-Item -Path (Join-Path $desktopExeParent '*') -Destination $desktopDir -Recurse -Force
}
else {
    Copy-Item $desktopExe $desktopDir
}
if ((Test-Path $desktopConsoleExe) -and -not (Test-Path (Join-Path $desktopDir 'main.exe'))) {
    Copy-Item $desktopConsoleExe (Join-Path $desktopDir 'main.exe') -Force
}
Copy-Item $desktopEnvExample (Join-Path $desktopDir 'lugest.env.example')
Copy-Item $desktopEnvExample (Join-Path $desktopDir 'lugest.env')
Copy-Item $desktopServerEnvExample (Join-Path $desktopDir 'lugest.env.servidor.example')
Copy-Item $desktopPostEnvExample (Join-Path $desktopDir 'lugest.env.posto.example')
Copy-Item $desktopBranding $desktopDir
Copy-Item $desktopQtConfig $desktopDir
Copy-Item $desktopIcon $desktopDir
Copy-Item $desktopLogo $desktopDir
Copy-Item $desktopAdminSetup (Join-Path $desktopDir 'setup_lugest_admin.ps1')
Copy-Item $desktopEnvDiagnose (Join-Path $desktopDir 'Diagnosticar Ligacao LuisGEST.ps1')
Copy-Item $desktopStartLauncher (Join-Path $desktopDir 'Arrancar LuisGEST Desktop.ps1')
Copy-Item $desktopStartLauncherHidden (Join-Path $desktopDir 'Arrancar LuisGEST Desktop.vbs')
Copy-Item $desktopUpdaterScript (Join-Path $desktopDir 'Atualizar LuisGEST.ps1')
Copy-Item $desktopRepairUpdaterScript (Join-Path $desktopDir 'Reparar Atualizador Instalado.ps1')
Copy-Item $desktopInstallerScript (Join-Path $desktopDir 'Instalar LuisGEST no computador.ps1')
Copy-Item $desktopVersionFile (Join-Path $desktopDir 'VERSION')
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

$appVersion = (Get-Content $desktopVersionFile -Raw -Encoding UTF8).Trim()
if (-not $appVersion) {
    $appVersion = $releaseStamp
}
$updateConfig = [ordered]@{
    current_version = $appVersion
    manifest_url = 'https://github.com/luisola1218/luGEST/releases/latest/download/latest.json'
    channel = 'stable'
    github_token = ''
    auto_check = $false
}
$updateConfig | ConvertTo-Json -Depth 5 | Set-Content -Path (Join-Path $desktopDir 'update_config.json') -Encoding UTF8

$desktopLauncher = @"
@echo off
cd /d %~dp0
if not exist lugest.env (
    echo Falta o ficheiro lugest.env.
    echo Copia lugest.env.example para lugest.env e ajusta os parametros MySQL antes de arrancar.
    pause
    exit /b 1
)
if exist "%~dp0Arrancar LuisGEST Desktop.vbs" (
    wscript.exe "%~dp0Arrancar LuisGEST Desktop.vbs" "%~dp0"
    exit /b 0
)
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Arrancar LuisGEST Desktop.ps1"
"@
Set-Content -Path (Join-Path $desktopDir 'Arrancar LuisGEST Desktop.bat') -Value $desktopLauncher -Encoding ASCII

$desktopUpdateLauncher = @'
@echo off
cd /d %~dp0
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Atualizar LuisGEST.ps1" %*
pause
'@
Set-Content -Path (Join-Path $desktopDir 'Atualizar LuisGEST.bat') -Value $desktopUpdateLauncher -Encoding ASCII

$desktopRepairUpdaterLauncher = @'
@echo off
cd /d %~dp0
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Reparar Atualizador Instalado.ps1" %*
pause
'@
Set-Content -Path (Join-Path $desktopDir 'Reparar Atualizador Instalado.bat') -Value $desktopRepairUpdaterLauncher -Encoding ASCII

$desktopInstallerLauncher = @'
@echo off
cd /d %~dp0
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Instalar LuisGEST no computador.ps1"
pause
'@
Set-Content -Path (Join-Path $desktopDir 'Instalar LuisGEST no computador.bat') -Value $desktopInstallerLauncher -Encoding ASCII

$desktopAdminLauncher = @'
@echo off
cd /d %~dp0
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -ExecutionPolicy Bypass -File "%~dp0setup_lugest_admin.ps1"
pause
'@
Set-Content -Path (Join-Path $desktopDir 'Criar Administrador Inicial.bat') -Value $desktopAdminLauncher -Encoding ASCII

$desktopAdminResetLauncher = @'
@echo off
cd /d %~dp0
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -ExecutionPolicy Bypass -File "%~dp0setup_lugest_admin.ps1" -Reset
pause
'@
Set-Content -Path (Join-Path $desktopDir 'Repor Password Administrador.bat') -Value $desktopAdminResetLauncher -Encoding ASCII

$desktopDiagnoseLauncher = @'
@echo off
cd /d %~dp0
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Diagnosticar Ligacao LuisGEST.ps1"
pause
'@
Set-Content -Path (Join-Path $desktopDir 'Diagnosticar Ligacao LuisGEST.bat') -Value $desktopDiagnoseLauncher -Encoding ASCII

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

if ($mobileApkSource) {
    Copy-Item $mobileApkSource (Join-Path $mobileApkDir ("LuisGEST-Impulse-release-" + $releaseStamp + ".apk"))
}
else {
    $apkPending = @"
Nao foi encontrada nenhuma APK release local para copiar.

Esta entrega segue com o codigo fonte Flutter, mas a APK precisa de ser gerada quando existir Flutter SDK neste posto:

    A APK deve ser gerada pelo programador a partir do projeto Flutter original.

Opcional para deixar a APK ja apontada para um servidor especifico:

    flutter build apk --release --dart-define=LUGEST_DEFAULT_API_HOST=http://IP_DO_SERVIDOR:8050
"@
    Set-Content -Path (Join-Path $mobileApkDir 'APK-PENDENTE.txt') -Value $apkPending -Encoding UTF8
}

$mysqlClientDir = Join-Path $dbDir 'mysql'
$mysqlMigrationsDir = Join-Path $mysqlClientDir 'Migracoes'
New-Item -ItemType Directory -Force -Path $mysqlClientDir, $mysqlMigrationsDir | Out-Null
foreach ($relativePath in @(
    'README.md',
    'lugest.sql',
    'lugest_instalacao_unica.sql',
    'install_lugest_mysql.py',
    'install_lugest_mysql.ps1',
    'instalar_lugest_mysql_admin.bat',
    'validate_lugest_mysql.ps1',
    'validar_lugest_mysql.bat',
    'backup_lugest_mysql.py',
    'backup_lugest_mysql.ps1',
    'backup_lugest_mysql_admin.bat',
    'restore_lugest_mysql.py',
    'restore_lugest_mysql.ps1',
    'restaurar_lugest_mysql_admin.bat',
    'apply_lugest_migrations.py',
    'apply_lugest_migrations.ps1',
    'aplicar_migracoes_lugest_admin.bat',
    'estado_migracoes_lugest_admin.bat',
    'mysql_tooling.py'
)) {
    $source = Join-Path $databaseSource $relativePath
    if (Test-Path $source) {
        Copy-Item $source $mysqlClientDir -Force
    }
}
Copy-Item (Join-Path $databaseSource 'lugest_instalacao_unica.sql') (Join-Path $mysqlClientDir 'IMPORTAR_NO_HEIDI.sql') -Force
Get-ChildItem $databaseSource -Filter 'patch_*.sql' -File | ForEach-Object {
    Copy-Item $_.FullName $mysqlMigrationsDir -Force
}
Copy-Item $securityPlan (Join-Path $docsDir 'CHECKLIST - Seguranca e Testes.md')
Copy-Item $fullInstallGuide (Join-Path $docsDir 'GUIA - INSTALACAO TOTAL PASSO A PASSO.md')
Copy-Item $installGuide (Join-Path $docsDir 'GUIA - Instalacao Outro Computador.md')
Copy-Item $simpleGuide (Join-Path $docsDir 'GUIA - MUITO SIMPLES.md')
Copy-Item $emailUpdatesGuide (Join-Path $docsDir 'GUIA - Email e Atualizacoes.md')

$updateZipName = "LuisGEST-Desktop-$($appVersion.Replace('.', '-')).zip"
$updateZipPath = Join-Path $updatesDir $updateZipName
$bootstrapAssetName = "Reparar Atualizador Instalado.ps1"
$bootstrapAssetPath = Join-Path $updatesDir $bootstrapAssetName
if (Test-Path $updateZipPath) {
    Remove-Item $updateZipPath -Force
}
if (Test-Path $bootstrapAssetPath) {
    Remove-Item $bootstrapAssetPath -Force
}
Compress-Archive -Path (Join-Path $desktopDir '*') -DestinationPath $updateZipPath -Force
Copy-Item $desktopRepairUpdaterScript $bootstrapAssetPath -Force
$updateHash = (Get-FileHash $updateZipPath -Algorithm SHA256).Hash
$updateManifest = [ordered]@{
    product = 'LuisGEST Desktop'
    version = $appVersion
    channel = 'stable'
    package_url = $updateZipName
    bootstrap_url = $bootstrapAssetName
    sha256 = $updateHash
    created_at = $releaseDate.ToString('s')
    notes = "Atualizacao LuisGEST $appVersion preparada em $releaseDateTxt."
    requires_app_backup = $true
    requires_database_backup = $true
}
$updateManifestJson = $updateManifest | ConvertTo-Json -Depth 5
Write-Utf8NoBomFile -Path (Join-Path $updatesDir 'latest.json') -Content $updateManifestJson

$readme = @"
# App LuisGEST - Revisao Final

Data de preparacao: $releaseDateTxt

## Conteudo
- Desktop App: executavel principal $desktopExeName, launcher Qt, lugest.env, branding, lugest_qt_config.json e lugest_trial.json placeholder.
- Desktop App\Instalar LuisGEST no computador.bat: instala como software normal e cria atalhos com icon.
- Mobile API: API FastAPI com .env placeholder e scripts de instalacao/arranque.
- Mobile APK: APK Android release atual, quando disponivel; caso contrario segue uma nota APK-PENDENTE.txt.
- Base de Dados\mysql: apenas instalacao, backup/restauro, validacao e migracoes necessarias.
- Atualizacoes: pacote ZIP desktop, bootstrap remoto e latest.json com checksum SHA256 para atualizar clientes.
- Documentacao: guias de instalacao, email, atualizacoes e checklist.

## Comecar por aqui
1. Se nunca fizeste uma instalacao, le primeiro Documentacao\GUIA - INSTALACAO TOTAL PASSO A PASSO.md.
2. No servidor, usar Desktop App\lugest.env.servidor.example como base do lugest.env.
3. Nos postos, usar Desktop App\lugest.env.posto.example como base do lugest.env.

## Arranque desktop
1. Opcional: correr Desktop App\Instalar LuisGEST no computador.bat para instalar e criar atalhos.
2. No servidor, usar Desktop App\lugest.env.servidor.example como base do lugest.env.
3. Nos postos, usar Desktop App\lugest.env.posto.example como base do lugest.env.
4. Se o ambiente for multi-posto, definir tambem LUGEST_SHARED_STORAGE_ROOT com uma pasta UNC partilhada para desenhos, PDFs e anexos.
5. Se a base estiver sem utilizadores locais, correr Desktop App\Criar Administrador Inicial.bat.
6. Ajustar Desktop App\lugest_branding.json, Desktop App\lugest_qt_config.json e os logos se quiser personalizacao.
7. Se fores usar trial, editar/gerir Desktop App\lugest_trial.json apenas na maquina final, nunca copiando o trial desta maquina.
8. Executar Desktop App\Arrancar LuisGEST Desktop.bat ou o atalho LuisGEST criado pelo instalador.
9. O desktop arranca sempre em Qt e trabalha apenas com MySQL.

## Arranque API mobile
1. Entrar em Mobile API.
2. Editar .env com os dados reais MySQL e segredo API.
3. Executar instalar_impulse_api.bat.
4. Executar arrancar_impulse_api.bat.
5. Para producao em Windows Server, usar instalar_impulse_api_arranque_automatico_admin.bat.

## Mobile Android
1. Se existir Mobile APK\LuisGEST-Impulse-release-$releaseStamp.apk, instalar essa APK.
2. Se a pasta trouxer APK-PENDENTE.txt, gerar primeiro a APK no ambiente do programador.
3. No primeiro login indicar o IP/URL do servidor API.
4. A app guarda o ultimo servidor e utilizador usados.

## Base de dados
- Para HeidiSQL, importar Base de Dados\mysql\IMPORTAR_NO_HEIDI.sql em instalacao nova.
- Em alternativa, importar Base de Dados\mysql\lugest.sql no MySQL da empresa, ou usar o instalador PowerShell.
- Em instalacao nova com logins temporarios e arranque rapido, podes importar Base de Dados\mysql\lugest_instalacao_unica.sql.
- Aplicar patches adicionais da mesma pasta se o ambiente os exigir.
- Para automatizar a instalacao inicial, usar Base de Dados\mysql\install_lugest_mysql.ps1 ou o atalho instalar_lugest_mysql_admin.bat.

## Atualizacoes desktop
1. Em cada cliente, abrir o LuisGEST como admin e ir a Extras > Atualizacoes.
2. O Manifest ja segue preconfigurado para a release estavel no GitHub privado.
3. O cliente apenas precisa de colar um token valido e carregar em Verificar.
4. Se preferires atualizacao por servidor proprio ou pasta partilhada, podes trocar manualmente o Manifest para outro latest.json.
5. O atualizador valida SHA256, cria backup da pasta Desktop App e tenta criar backup MySQL com mysqldump antes de instalar.

## Notas finais
- A release desktop deve arrancar preferencialmente por lugest_qt.exe; main.exe fica apenas como fallback tecnico quando existir.
- O desktop suporta um ficheiro externo lugest.env, por isso nao e preciso alterar codigo para mudar empresa/servidor.
- Em multi-posto, usa LUGEST_SHARED_STORAGE_ROOT com um caminho UNC comum a todos os postos.
- O login OWNER no lugest.env serve para trial/licenciamento; o admin normal da app cria-se ou repoe-se pelos scripts da pasta Desktop App.
- Usa uma conta MySQL dedicada para a aplicacao em vez de root.
- O menu Faturacao faz seguimento de vendidos, faturas, pagamentos e comprovativos.
- O menu Transportes faz agendamento de viagens, afeta encomendas e gera folha de rota PDF.
- Este pacote nao leva os segredos da maquina atual: lugest.env, .env da API e o trial seguem em modo placeholder.
- Depois de instalar, usa Documentacao\CHECKLIST - Seguranca e Testes.md para validar o ambiente novo.
- Para email e atualizacoes, usa Documentacao\GUIA - Email e Atualizacoes.md.
"@
Set-Content -Path (Join-Path $releaseRoot 'README - Revisao Final.md') -Value $readme -Encoding UTF8

Write-Output $releaseRoot
