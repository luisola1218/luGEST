$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$desktopRoot = Join-Path $env:USERPROFILE 'Desktop'
$releaseDate = Get-Date
$releaseDateTxt = $releaseDate.ToString('dd/MM/yyyy HH:mm')
$releaseName = 'App LuisGEST - Revis' + [char]0x00E3 + 'o Final'
$releaseRoot = Join-Path $desktopRoot $releaseName

$envExample = Join-Path $repoRoot 'config\examples\lugest.env.example'
$serverEnvExample = Join-Path $repoRoot 'config\examples\lugest.env.servidor.example'
$postEnvExample = Join-Path $repoRoot 'config\examples\lugest.env.posto.example'
$brandingFile = Join-Path $repoRoot 'lugest_branding.json'
$qtConfigFile = Join-Path $repoRoot 'lugest_qt_config.json'
$iconFile = Join-Path $repoRoot 'app.ico'
$logoFile = Join-Path $repoRoot 'logo.jpg'
$logosDir = Join-Path $repoRoot 'Logos'
$databaseSource = Join-Path $repoRoot 'mysql'
$versionFile = Join-Path $repoRoot 'VERSION'
$securityPlan = Join-Path $repoRoot 'docs\plans\SECURITY_TEST_PLAN.md'
$localGuide = Join-Path $repoRoot 'docs\install\GUIA_ARRANQUE_QT_LOCAL.md'
$updateGuide = Join-Path $repoRoot 'docs\install\UPDATE_FLOW_CLIENTE.md'

function Resolve-DesktopExePath {
    foreach ($relativePath in @(
        'dist\lugest_qt\lugest_qt.exe',
        'dist_qt_stable\lugest_qt\lugest_qt.exe',
        'dist\lugest_qt.exe'
    )) {
        $candidate = Join-Path $repoRoot $relativePath
        if (Test-Path $candidate) {
            return (Get-Item $candidate).FullName
        }
    }
    return $null
}

function Write-Utf8NoBomFile {
    param(
        [string]$Path,
        [string]$Content
    )
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Content, $utf8NoBom)
}

$desktopExe = Resolve-DesktopExePath
if (-not $desktopExe) {
    throw "Nao foi encontrado o executavel desktop Qt. Gera primeiro o build PyInstaller."
}
$desktopExeItem = Get-Item $desktopExe
$desktopExeParent = $desktopExeItem.Directory.FullName
$pySideRuntime = Join-Path $desktopExeParent '_internal\PySide6'
if ($desktopExeParent -like (Join-Path $repoRoot 'dist\lugest_qt') -or $desktopExeParent -like (Join-Path $repoRoot 'dist_qt_stable\lugest_qt')) {
    if (-not (Test-Path $pySideRuntime)) {
        throw "Build invalido: falta PySide6 no runtime PyInstaller. Recria o build com .\.venv\Scripts\python.exe -m PyInstaller lugest_qt.spec --noconfirm"
    }
}

foreach ($requiredPath in @(
    $desktopExe,
    $envExample,
    $serverEnvExample,
    $postEnvExample,
    $brandingFile,
    $qtConfigFile,
    $iconFile,
    $logoFile,
    $logosDir,
    $databaseSource,
    $versionFile,
    $securityPlan,
    $localGuide,
    $updateGuide
)) {
    if (-not (Test-Path $requiredPath)) {
        throw "Falta ficheiro/pasta obrigatoria para a release: $requiredPath"
    }
}

Get-ChildItem $desktopRoot -ErrorAction SilentlyContinue |
    Where-Object { $_.PSIsContainer -and $_.Name -like 'App LuisGEST*' } |
    ForEach-Object { Remove-Item $_.FullName -Recurse -Force }

$dbDir = Join-Path $releaseRoot 'Base de Dados'
$mysqlDir = Join-Path $dbDir 'mysql'
$migrationsDir = Join-Path $mysqlDir 'Migracoes'
$docsDir = Join-Path $releaseRoot 'Documentacao'

New-Item -ItemType Directory -Force -Path $releaseRoot, $mysqlDir, $migrationsDir, $docsDir | Out-Null

if ($desktopExeParent -like (Join-Path $repoRoot 'dist\lugest_qt') -or $desktopExeParent -like (Join-Path $repoRoot 'dist_qt_stable\lugest_qt')) {
    Copy-Item -Path (Join-Path $desktopExeParent '*') -Destination $releaseRoot -Recurse -Force
}
else {
    Copy-Item $desktopExe (Join-Path $releaseRoot 'lugest_qt.exe') -Force
}

$releasedExe = Join-Path $releaseRoot 'lugest_qt.exe'
if (Test-Path $releasedExe) {
    Rename-Item -LiteralPath $releasedExe -NewName 'LuisGEST.exe' -Force
}
$internalDir = Join-Path $releaseRoot '_internal'
if (Test-Path $internalDir) {
    $internalItem = Get-Item -LiteralPath $internalDir
    $internalItem.Attributes = $internalItem.Attributes -band (-bnot [System.IO.FileAttributes]::Hidden)
}

Copy-Item $envExample (Join-Path $releaseRoot 'lugest.env.example') -Force
Copy-Item $envExample (Join-Path $releaseRoot 'lugest.env') -Force
Copy-Item $serverEnvExample (Join-Path $releaseRoot 'lugest.env.servidor.example') -Force
Copy-Item $postEnvExample (Join-Path $releaseRoot 'lugest.env.posto.example') -Force
Copy-Item $brandingFile $releaseRoot -Force
Copy-Item $qtConfigFile $releaseRoot -Force
Copy-Item $iconFile $releaseRoot -Force
Copy-Item $logoFile $releaseRoot -Force
Copy-Item $versionFile $releaseRoot -Force
Copy-Item -Recurse $logosDir $releaseRoot -Force

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
$trialTemplate | ConvertTo-Json -Depth 4 | Set-Content -Path (Join-Path $releaseRoot 'lugest_trial.json') -Encoding UTF8

$installer = @"
`$ErrorActionPreference = 'Stop'

`$sourceRoot = Split-Path -Parent `$MyInvocation.MyCommand.Path
`$installRoot = 'C:\LuisGEST'
`$desktopShortcut = Join-Path `$env:USERPROFILE 'Desktop\LuisGEST.lnk'

if (-not (Test-Path (Join-Path `$sourceRoot 'LuisGEST.exe'))) {
    throw 'Esta pasta nao contem LuisGEST.exe.'
}
if (-not (Test-Path (Join-Path `$sourceRoot '_internal\PySide6'))) {
    throw 'Instalacao incompleta: falta a pasta _internal\PySide6. Copia a pasta LuisGEST completa.'
}

New-Item -ItemType Directory -Force -Path `$installRoot | Out-Null
Get-ChildItem `$installRoot -Force -ErrorAction SilentlyContinue |
    Where-Object { `$_.Name -notin @('generated', 'backups') } |
    ForEach-Object { Remove-Item `$_.FullName -Recurse -Force }

Copy-Item -Path (Join-Path `$sourceRoot '*') -Destination `$installRoot -Recurse -Force

`$shell = New-Object -ComObject WScript.Shell
`$shortcut = `$shell.CreateShortcut(`$desktopShortcut)
`$shortcut.TargetPath = Join-Path `$installRoot 'LuisGEST.exe'
`$shortcut.WorkingDirectory = `$installRoot
`$shortcut.IconLocation = Join-Path `$installRoot 'app.ico'
`$shortcut.Description = 'LuisGEST ERP industrial'
`$shortcut.Save()

Write-Host ''
Write-Host 'LuisGEST instalado com sucesso em C:\LuisGEST'
Write-Host 'Atalho criado no Ambiente de Trabalho: LuisGEST'
Write-Host ''
Write-Host 'Se este posto for cliente, confirma o ficheiro C:\LuisGEST\lugest.env com o IP/nome do servidor MySQL.'
Pause
"@
Write-Utf8NoBomFile -Path (Join-Path $releaseRoot 'INSTALAR_LUISGEST.ps1') -Content $installer

foreach ($relativePath in @(
    'README.md',
    'lugest.sql',
    'lugest_instalacao_unica.sql',
    'install_lugest_mysql.py',
    'install_lugest_mysql.ps1',
    'validate_lugest_mysql.ps1',
    'backup_lugest_mysql.py',
    'backup_lugest_mysql.ps1',
    'restore_lugest_mysql.py',
    'restore_lugest_mysql.ps1',
    'apply_lugest_migrations.py',
    'apply_lugest_migrations.ps1',
    'mysql_tooling.py'
)) {
    $source = Join-Path $databaseSource $relativePath
    if (Test-Path $source) {
        Copy-Item $source $mysqlDir -Force
    }
}
Copy-Item (Join-Path $databaseSource 'lugest_instalacao_unica.sql') (Join-Path $mysqlDir 'IMPORTAR_NO_HEIDI.sql') -Force
Get-ChildItem $databaseSource -Filter 'patch_*.sql' -File | ForEach-Object {
    Copy-Item $_.FullName $migrationsDir -Force
}

Copy-Item $securityPlan (Join-Path $docsDir 'CHECKLIST - Seguranca e Testes.md') -Force
Copy-Item $localGuide (Join-Path $docsDir 'GUIA - Arranque Desktop Local.md') -Force
Copy-Item $updateGuide (Join-Path $docsDir 'GUIA - Atualizacao Cliente.md') -Force

$readme = @"
# LuisGEST Desktop

Preparado em: $releaseDateTxt

Esta pasta foi simplificada para testes no cliente. Nesta fase segue apenas a aplicacao desktop.

## O que interessa
- LuisGEST.exe: aplicacao principal.
- INSTALAR_LUISGEST.ps1: instala a app em C:\LuisGEST e cria atalho no Ambiente de Trabalho.
- _internal: motor interno do executavel; nao apagar nem copiar o LuisGEST.exe sozinho.
- lugest.env: configuracao da ligacao MySQL deste posto.
- lugest.env.servidor.example: exemplo para o computador servidor.
- lugest.env.posto.example: exemplo para os outros postos.
- lugest_branding.json, lugest_qt_config.json e Logos: configuracao visual e PDFs.
- Base de Dados\mysql: SQL, migracoes e ferramentas PowerShell/Python sem atalhos .bat.
- Documentacao: guias essenciais de instalacao e checklist.

## Como arrancar
1. No cliente, clicar com botao direito em INSTALAR_LUISGEST.ps1 e escolher Executar com PowerShell.
2. Confirmar o ficheiro C:\LuisGEST\lugest.env com os dados MySQL corretos.
3. Em multiutilizador, todos os postos devem apontar para a mesma base MySQL.
4. Se houver ficheiros partilhados, usar em lugest.env uma pasta UNC comum em LUGEST_SHARED_STORAGE_ROOT.
5. Abrir pelo atalho LuisGEST no Ambiente de Trabalho.

## Notas
- Os atalhos .bat foram removidos para reduzir confusao; a instalacao usa apenas PowerShell e depois arranca pelo atalho.
- Para instalar base nova no HeidiSQL, importar Base de Dados\mysql\IMPORTAR_NO_HEIDI.sql.
"@
Write-Utf8NoBomFile -Path (Join-Path $releaseRoot 'README.md') -Content $readme

Write-Output $releaseRoot
