param(
    [switch]$Commercial
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$venvPython = Join-Path $repoRoot '.venv\Scripts\python.exe'
$desktopRoot = Join-Path $env:USERPROFILE 'Desktop'
$releaseDate = Get-Date
$releaseDateTxt = $releaseDate.ToString('dd/MM/yyyy HH:mm')
$releaseName = if ($Commercial) { 'luGEST - Pacote Comercial Piloto' } else { 'App luGEST - Revis' + [char]0x00E3 + 'o Final' }
$releaseRoot = Join-Path $desktopRoot $releaseName

$envExample = Join-Path $repoRoot 'config\examples\lugest.env.example'
$activeEnvFile = Join-Path $repoRoot 'lugest.env'
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
$readinessGuide = Join-Path $repoRoot 'docs\install\PRONTIDAO_COMERCIAL.md'
$fluidityAudit = Join-Path $repoRoot 'docs\plans\AUDITORIA_FLUIDEZ_2026-07-23.md'
$profileGuide = Join-Path $repoRoot 'docs\PERFIS_ESTRUTURAIS.md'
$localAiGuide = Join-Path $repoRoot 'docs\install\IA_LOCAL_OLLAMA.md'
$preservedConfigRoot = Join-Path $env:TEMP ("lugest_release_config_" + [guid]::NewGuid().ToString('N'))
$preservedEnv = Join-Path $preservedConfigRoot 'lugest.env'
$preservedTrial = Join-Path $preservedConfigRoot 'lugest_trial.json'
$preservedBackups = Join-Path $preservedConfigRoot 'Backups'

function Resolve-DesktopExePath {
    foreach ($relativePath in @(
        'dist\lugest_qt\lugest_qt.exe',
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
    $updateGuide,
    $readinessGuide,
    $fluidityAudit,
    $profileGuide,
    $localAiGuide
)) {
    if (-not (Test-Path $requiredPath)) {
        throw "Falta ficheiro/pasta obrigatoria para a release: $requiredPath"
    }
}

$existingRelease = Get-ChildItem $desktopRoot -ErrorAction SilentlyContinue |
    Where-Object { $_.PSIsContainer -and $_.Name -eq $releaseName } |
    Select-Object -First 1
if ($existingRelease -and -not $Commercial) {
    New-Item -ItemType Directory -Force -Path $preservedConfigRoot | Out-Null
    foreach ($name in @('lugest.env', 'lugest_trial.json')) {
        $source = Join-Path $existingRelease.FullName $name
        if (Test-Path $source) {
            Copy-Item $source (Join-Path $preservedConfigRoot $name) -Force
        }
    }
    $existingBackups = Join-Path $existingRelease.FullName 'Base de Dados\Backups'
    if (Test-Path $existingBackups) {
        Copy-Item $existingBackups $preservedBackups -Recurse -Force
    }
}

if (Test-Path $releaseRoot) {
    Remove-Item $releaseRoot -Recurse -Force
}

$dbDir = Join-Path $releaseRoot 'Base de Dados'
$mysqlDir = Join-Path $dbDir 'mysql'
$docsDir = Join-Path $releaseRoot 'Documentacao'

New-Item -ItemType Directory -Force -Path $releaseRoot, $mysqlDir, $docsDir | Out-Null

if ($desktopExeParent -like (Join-Path $repoRoot 'dist\lugest_qt') -or $desktopExeParent -like (Join-Path $repoRoot 'dist_qt_stable\lugest_qt')) {
    Copy-Item -Path (Join-Path $desktopExeParent '*') -Destination $releaseRoot -Recurse -Force
}
else {
    Copy-Item $desktopExe (Join-Path $releaseRoot 'lugest_qt.exe') -Force
}

if ($Commercial) {
    # A valid smoke test can create mutable runtime data beside the executable.
    # Never transfer drawings, generated documents, backups or workstation state
    # from the build machine into a clean commercial package.
    foreach ($relativePath in @(
        'generated',
        'backups',
        'lugest_runtime_state.json',
        'lugest_supplier_seq.json'
    )) {
        $runtimeArtifact = Join-Path $releaseRoot $relativePath
        if (Test-Path -LiteralPath $runtimeArtifact) {
            Remove-Item -LiteralPath $runtimeArtifact -Recurse -Force
        }
    }
}

$releasedExe = Join-Path $releaseRoot 'lugest_qt.exe'
if (Test-Path $releasedExe) {
    Rename-Item -LiteralPath $releasedExe -NewName 'luGEST.exe' -Force
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

if ($Commercial) {
    foreach ($commercialEnvName in @(
        'lugest.env',
        'lugest.env.example',
        'lugest.env.servidor.example',
        'lugest.env.posto.example'
    )) {
        $commercialEnvPath = Join-Path $releaseRoot $commercialEnvName
        $commercialEnv = Get-Content $commercialEnvPath -Raw
        $commercialEnv = [regex]::Replace(
            $commercialEnv,
            '(?m)^LUGEST_OWNER_USERNAME=.*$',
            'LUGEST_OWNER_USERNAME='
        )
        $commercialEnv = [regex]::Replace(
            $commercialEnv,
            '(?m)^LUGEST_OWNER_PASSWORD=.*$',
            'LUGEST_OWNER_PASSWORD='
        )
        Write-Utf8NoBomFile -Path $commercialEnvPath -Content $commercialEnv
    }

    $commercialBranding = [ordered]@{
        logo_candidates = @('logo.jpg', 'Logos\logo.png', 'Logos\lg.png')
        empresa_info_rodape = @()
        guia_emitente = [ordered]@{
            nome = ''
            nif = ''
            morada = ''
            local_carga = ''
        }
        guia_info_extra = @()
        primary_color = '#454567'
        logo_scale_pct = 90
    }
    $commercialBranding |
        ConvertTo-Json -Depth 6 |
        Set-Content -Path (Join-Path $releaseRoot 'lugest_branding.json') -Encoding UTF8

    $sourceQtConfig = Get-Content $qtConfigFile -Raw | ConvertFrom-Json
    $commercialQtConfig = [ordered]@{
        ui_options = [ordered]@{
            operator_show_client_name = $true
        }
        user_profiles = [ordered]@{}
        material_assistant_feedback = [ordered]@{}
        material_assistant_checks = [ordered]@{}
        pulse_plan_delay_reasons = @()
        update_settings = $sourceQtConfig.update_settings
        pdf = $sourceQtConfig.pdf
    }
    $commercialQtConfig |
        ConvertTo-Json -Depth 20 |
        Set-Content -Path (Join-Path $releaseRoot 'lugest_qt_config.json') -Encoding UTF8
}

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
    last_trusted_at = ''
    last_time_source = ''
    notes = ''
}
$trialTemplate | ConvertTo-Json -Depth 4 | Set-Content -Path (Join-Path $releaseRoot 'lugest_trial.json') -Encoding UTF8
if (-not $Commercial -and (Test-Path $preservedEnv)) {
    $preservedHash = (Get-FileHash $preservedEnv -Algorithm SHA256).Hash
    $exampleHash = (Get-FileHash $envExample -Algorithm SHA256).Hash
    if ($preservedHash -ne $exampleHash) {
        Copy-Item $preservedEnv (Join-Path $releaseRoot 'lugest.env') -Force
    }
    elseif (Test-Path $activeEnvFile) {
        Copy-Item $activeEnvFile (Join-Path $releaseRoot 'lugest.env') -Force
    }
}
elseif (-not $Commercial -and (Test-Path $activeEnvFile)) {
    Copy-Item $activeEnvFile (Join-Path $releaseRoot 'lugest.env') -Force
}
if (-not $Commercial -and (Test-Path $preservedTrial)) {
    Copy-Item $preservedTrial (Join-Path $releaseRoot 'lugest_trial.json') -Force
}
if (-not $Commercial -and (Test-Path $preservedBackups)) {
    $releaseBackups = Join-Path $releaseRoot 'Base de Dados\Backups'
    New-Item -ItemType Directory -Force -Path $releaseBackups | Out-Null
    Copy-Item (Join-Path $preservedBackups '*') $releaseBackups -Recurse -Force
}
if (Test-Path $preservedConfigRoot) {
    Remove-Item $preservedConfigRoot -Recurse -Force
}

$installer = @"
param(
    [string]`$InstallRoot = '',
    [switch]`$NoShortcut,
    [switch]`$NoPause
)

`$ErrorActionPreference = 'Stop'

function Resolve-SafeDirectory {
    param([string]`$RequestedPath)

    `$candidate = `$RequestedPath
    if ([string]::IsNullOrWhiteSpace(`$candidate)) {
        `$candidate = Join-Path `$env:LOCALAPPDATA 'luGEST'
    }
    `$fullPath = [IO.Path]::GetFullPath([Environment]::ExpandEnvironmentVariables(`$candidate))
    `$driveRoot = [IO.Path]::GetPathRoot(`$fullPath).TrimEnd('\')
    if (`$fullPath.TrimEnd('\') -eq `$driveRoot) {
        throw "O destino nao pode ser a raiz do disco: `$fullPath"
    }
    return `$fullPath.TrimEnd('\')
}

function Assert-WritableDirectory {
    param([string]`$PathToTest)

    try {
        New-Item -ItemType Directory -Force -Path `$PathToTest | Out-Null
        `$probe = Join-Path `$PathToTest ('.lugest_write_' + [guid]::NewGuid().ToString('N') + '.tmp')
        [IO.File]::WriteAllText(`$probe, 'ok')
        Remove-Item -LiteralPath `$probe -Force
    }
    catch {
        throw "Nao foi possivel escrever em '`$PathToTest'. Escolhe outra pasta com -InstallRoot. Detalhe: `$(`$_.Exception.Message)"
    }
}

`$sourceRoot = [IO.Path]::GetFullPath((Split-Path -Parent `$MyInvocation.MyCommand.Path)).TrimEnd('\')
`$installRoot = Resolve-SafeDirectory -RequestedPath `$InstallRoot
`$sourcePrefix = `$sourceRoot + '\'
`$installPrefix = `$installRoot + '\'
if (`$installRoot -eq `$sourceRoot -or `$installPrefix.StartsWith(`$sourcePrefix, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'O destino da instalacao nao pode ser a pasta do pacote nem uma subpasta dela.'
}
if (-not (Test-Path -LiteralPath (Join-Path `$sourceRoot 'luGEST.exe'))) {
    throw 'Pacote incompleto: nao foi encontrado luGEST.exe.'
}
if (-not (Test-Path -LiteralPath (Join-Path `$sourceRoot '_internal\PySide6'))) {
    throw 'Pacote incompleto: falta _internal\PySide6. Copia sempre a pasta completa.'
}

Assert-WritableDirectory -PathToTest `$installRoot
`$preserveRoot = Join-Path `$env:TEMP ('lugest_install_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force -Path `$preserveRoot | Out-Null

foreach (`$name in @('lugest.env', 'lugest_trial.json', 'lugest_branding.json', 'lugest_qt_config.json', 'Logos', 'generated', 'backups')) {
    `$existing = Join-Path `$installRoot `$name
    if (Test-Path -LiteralPath `$existing) {
        Copy-Item -LiteralPath `$existing -Destination (Join-Path `$preserveRoot `$name) -Recurse -Force
    }
}
try {
    Get-ChildItem -LiteralPath `$installRoot -Force -ErrorAction SilentlyContinue |
        ForEach-Object { Remove-Item -LiteralPath `$_.FullName -Recurse -Force }
    Copy-Item -Path (Join-Path `$sourceRoot '*') -Destination `$installRoot -Recurse -Force
    Get-ChildItem -LiteralPath `$preserveRoot -Force -ErrorAction SilentlyContinue |
        ForEach-Object {
            Copy-Item -LiteralPath `$_.FullName -Destination (Join-Path `$installRoot `$_.Name) -Recurse -Force
        }
}
finally {
    Remove-Item -LiteralPath `$preserveRoot -Recurse -Force -ErrorAction SilentlyContinue
}

if (-not `$NoShortcut) {
    `$desktopShortcut = Join-Path ([Environment]::GetFolderPath('Desktop')) 'luGEST.lnk'
    `$startMenuDir = Join-Path ([Environment]::GetFolderPath('Programs')) 'luGEST'
    New-Item -ItemType Directory -Force -Path `$startMenuDir | Out-Null
    `$shell = New-Object -ComObject WScript.Shell
    foreach (`$shortcutPath in @(`$desktopShortcut, (Join-Path `$startMenuDir 'luGEST.lnk'))) {
        `$shortcut = `$shell.CreateShortcut(`$shortcutPath)
        `$shortcut.TargetPath = Join-Path `$installRoot 'luGEST.exe'
        `$shortcut.WorkingDirectory = `$installRoot
        `$shortcut.IconLocation = Join-Path `$installRoot 'app.ico'
        `$shortcut.Description = 'luGEST ERP industrial'
        `$shortcut.Save()
    }
}

Write-Host ''
Write-Host "luGEST instalado com sucesso em `$installRoot" -ForegroundColor Green
if (-not `$NoShortcut) { Write-Host 'Atalhos criados no Ambiente de Trabalho e no menu Iniciar.' }
Write-Host "Confirma agora a ligacao MySQL em `$installRoot\lugest.env."
Write-Host 'A instalacao foi efetuada no perfil do utilizador e nao exige privilegios de administrador.'
if (-not `$NoPause) { Pause }
"@
Write-Utf8NoBomFile -Path (Join-Path $releaseRoot 'INSTALAR_LUISGEST.ps1') -Content $installer

$ownerConfigurator = @'
param(
    [string]$InstallRoot = ''
)

$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrWhiteSpace($InstallRoot)) {
    $InstallRoot = Join-Path $env:LOCALAPPDATA 'luGEST'
}
$InstallRoot = [IO.Path]::GetFullPath([Environment]::ExpandEnvironmentVariables($InstallRoot))
$envPath = Join-Path $InstallRoot 'lugest.env'
if (-not (Test-Path -LiteralPath $envPath)) {
    throw "Nao foi encontrado $envPath. Instala primeiro o LuisGEST."
}

function ConvertTo-PlainText {
    param([Security.SecureString]$SecureValue)
    $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureValue)
    try {
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
    }
    finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
    }
}

function ConvertTo-Base64Url {
    param([byte[]]$Bytes)
    return [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
}

$ownerUser = Read-Host 'Utilizador OWNER [lugestowner]'
if ([string]::IsNullOrWhiteSpace($ownerUser)) {
    $ownerUser = 'lugestowner'
}
$ownerUser = $ownerUser.Trim()
if ($ownerUser -match '\s') {
    throw 'O utilizador OWNER nao pode conter espacos.'
}

$secureOne = Read-Host 'Nova password OWNER (minimo 12 caracteres)' -AsSecureString
$secureTwo = Read-Host 'Confirmar password OWNER' -AsSecureString
$plainOne = ConvertTo-PlainText $secureOne
$plainTwo = ConvertTo-PlainText $secureTwo
try {
    if ($plainOne -cne $plainTwo) {
        throw 'As passwords nao coincidem.'
    }
    if ($plainOne.Length -lt 12) {
        throw 'A password OWNER deve ter pelo menos 12 caracteres.'
    }
    if ($plainOne -notmatch '[A-Z]' -or $plainOne -notmatch '[a-z]' -or
        $plainOne -notmatch '[0-9]' -or $plainOne -notmatch '[^A-Za-z0-9]') {
        throw 'Usa maiusculas, minusculas, numeros e um simbolo na password OWNER.'
    }

    $salt = New-Object byte[] 16
    $random = [Security.Cryptography.RandomNumberGenerator]::Create()
    try {
        $random.GetBytes($salt)
    }
    finally {
        $random.Dispose()
    }
    $iterations = 260000
    $derive = [Security.Cryptography.Rfc2898DeriveBytes]::new(
        $plainOne,
        $salt,
        $iterations,
        [Security.Cryptography.HashAlgorithmName]::SHA256
    )
    try {
        $digest = $derive.GetBytes(32)
    }
    finally {
        $derive.Dispose()
    }
    $passwordHash = 'pbkdf2_sha256${0}${1}${2}' -f (
        $iterations,
        (ConvertTo-Base64Url $salt),
        (ConvertTo-Base64Url $digest)
    )

    $envText = Get-Content -LiteralPath $envPath -Raw
    foreach ($entry in @(
        @('LUGEST_OWNER_USERNAME', $ownerUser),
        @('LUGEST_OWNER_PASSWORD', $passwordHash)
    )) {
        $key = $entry[0]
        $value = $entry[1]
        if ($envText -match "(?m)^$([regex]::Escape($key))=") {
            $envText = [regex]::Replace(
                $envText,
                "(?m)^$([regex]::Escape($key))=.*$",
                "$key=$value"
            )
        }
        else {
            $envText = $envText.TrimEnd() + [Environment]::NewLine + "$key=$value" + [Environment]::NewLine
        }
    }
    $utf8NoBom = New-Object Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($envPath, $envText, $utf8NoBom)
}
finally {
    $plainOne = $null
    $plainTwo = $null
}

Write-Host ''
Write-Host 'OWNER configurado com sucesso.' -ForegroundColor Green
Write-Host "Utilizador: $ownerUser"
Write-Host 'A password nao foi guardada em texto simples. Guarda-a no teu gestor de passwords.'
Write-Host 'Ja podes abrir o luGEST e gerir o trial em Extras > Trial / licenca.'
Write-Host ''
Pause
'@
Write-Utf8NoBomFile -Path (Join-Path $releaseRoot 'CONFIGURAR_OWNER_TRIAL.ps1') -Content $ownerConfigurator

$uninstaller = @'
param(
    [string]$InstallRoot = '',
    [switch]$RemoveLocalData,
    [switch]$NoPause
)

$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrWhiteSpace($InstallRoot)) {
    $InstallRoot = Join-Path $env:LOCALAPPDATA 'luGEST'
}
$installRoot = [IO.Path]::GetFullPath([Environment]::ExpandEnvironmentVariables($InstallRoot)).TrimEnd('\')
$driveRoot = [IO.Path]::GetPathRoot($installRoot).TrimEnd('\')
if ($installRoot -eq $driveRoot -or -not (Test-Path -LiteralPath (Join-Path $installRoot 'luGEST.exe'))) {
    throw "Destino de desinstalacao invalido ou sem luGEST.exe: $installRoot"
}

$backupRoot = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'luGEST Backups'
$backupDir = Join-Path $backupRoot ('Desinstalacao_' + (Get-Date -Format 'yyyyMMdd_HHmmss'))
$itemsToPreserve = @(
    'lugest.env',
    'lugest_trial.json',
    'lugest_branding.json',
    'lugest_qt_config.json',
    'Logos',
    'generated',
    'backups'
)
$foundData = $false
if (-not $RemoveLocalData) {
    foreach ($name in $itemsToPreserve) {
        if (Test-Path -LiteralPath (Join-Path $installRoot $name)) {
            if (-not $foundData) {
                New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
                $foundData = $true
            }
            Copy-Item -LiteralPath (Join-Path $installRoot $name) -Destination (Join-Path $backupDir $name) -Recurse -Force
        }
    }
}

foreach ($shortcutPath in @(
    (Join-Path ([Environment]::GetFolderPath('Desktop')) 'luGEST.lnk'),
    (Join-Path ([Environment]::GetFolderPath('Programs')) 'luGEST\luGEST.lnk')
)) {
    if (Test-Path -LiteralPath $shortcutPath) {
        Remove-Item -LiteralPath $shortcutPath -Force
    }
}
$startMenuDir = Join-Path ([Environment]::GetFolderPath('Programs')) 'luGEST'
if (Test-Path -LiteralPath $startMenuDir) {
    Remove-Item -LiteralPath $startMenuDir -Recurse -Force
}

Remove-Item -LiteralPath $installRoot -Recurse -Force
Write-Host ''
Write-Host 'luGEST desinstalado com sucesso.' -ForegroundColor Green
if ($foundData) {
    Write-Host "Configuracao e dados locais preservados em: $backupDir"
}
if (-not $NoPause) { Pause }
'@
Write-Utf8NoBomFile -Path (Join-Path $releaseRoot 'DESINSTALAR_LUGEST.ps1') -Content $uninstaller

$canonicalSchema = Join-Path $databaseSource 'lugest.sql'
$releaseSchema = Join-Path $mysqlDir 'IMPORTAR_NO_HEIDI.sql'
if (Test-Path $venvPython) {
    $temporarySchema = Join-Path $env:TEMP ("lugest_release_schema_" + [guid]::NewGuid().ToString('N') + '.sql')
    try {
        & $venvPython (Join-Path $databaseSource 'export_current_schema_sql.py') --with-starter-users --output $temporarySchema | Out-Null
        Copy-Item $temporarySchema $releaseSchema -Force
    }
    finally {
        if (Test-Path -LiteralPath $temporarySchema) {
            Remove-Item -LiteralPath $temporarySchema -Force
        }
    }
}
else {
    Copy-Item $canonicalSchema $releaseSchema -Force
}
$dbReadme = @"
# Base de Dados luGEST

Para instalar uma base nova no cliente, importar no HeidiSQL:

1. IMPORTAR_NO_HEIDI.sql

Este ficheiro contem todas as tabelas atuais e utilizadores iniciais temporarios.

Utilizadores temporarios:
- admin / Trocar#Admin2026
- operador / Trocar#Operador2026
- orcamentista / Trocar#Orc2026
- planeamento / Trocar#Planeamento2026

Trocar as passwords logo apos a instalacao.
"@
Write-Utf8NoBomFile -Path (Join-Path $mysqlDir 'README_BASE_DADOS.txt') -Content $dbReadme

Copy-Item $securityPlan (Join-Path $docsDir 'CHECKLIST - Seguranca e Testes.md') -Force
Copy-Item $localGuide (Join-Path $docsDir 'GUIA - Arranque Desktop Local.md') -Force
Copy-Item $updateGuide (Join-Path $docsDir 'GUIA - Atualizacao Cliente.md') -Force
Copy-Item $readinessGuide (Join-Path $docsDir 'PRONTIDAO COMERCIAL.md') -Force
Copy-Item $fluidityAudit (Join-Path $docsDir 'AUDITORIA - Fluidez e Robustez.md') -Force
Copy-Item $profileGuide (Join-Path $docsDir 'GUIA - Perfis Estruturais.md') -Force
Copy-Item $localAiGuide (Join-Path $docsDir 'GUIA - Copiloto IA Local.md') -Force

$presentationReadme = @"
# Guia de Apresentacao Comercial

## Antes da demonstracao
1. Importar `Base de Dados\mysql\IMPORTAR_NO_HEIDI.sql` numa instalacao MySQL limpa.
2. Configurar `lugest.env` com o servidor, utilizador e password dedicados ao cliente.
3. Executar `INSTALAR_LUISGEST.ps1` com PowerShell.
4. Entrar com `admin / Trocar#Admin2026` e alterar imediatamente as passwords temporarias.
5. Configurar logotipo, dados da empresa, operadores e parametros comerciais antes de usar em producao.

## Ativar o trial no cliente
1. Instalar primeiro o luGEST.
2. Executar `CONFIGURAR_OWNER_TRIAL.ps1` na pasta onde o luGEST foi instalado (por omissao, `%LOCALAPPDATA%\luGEST`).
3. Guardar o utilizador e a password OWNER num gestor de passwords; a password nao fica gravada em texto simples.
4. Garantir acesso HTTPS (porta 443) a `www.google.com` e `www.cloudflare.com`; `www.microsoft.com` e a contingencia.
5. Abrir o luGEST e entrar com a conta OWNER.
6. Abrir `Extras > Trial / licenca`, validar a hora online, indicar empresa e dias, e clicar em `Ativar / reiniciar`.
7. Sair da sessao OWNER e entrar com o utilizador normal do cliente.

## Percurso recomendado da demonstracao
1. Dashboard e Pulse: visao global da operacao.
2. Clientes e Fornecedores: fichas, contactos e localizacao.
3. Orcamentos: calculo, referencias, PDF e aprovacao.
4. Encomendas: ordem de fabrico, materiais, pecas e montagem.
5. Assistente MP: prioridades, plano de separacao, stock e alertas.
6. Planeamento, Operador e OPP: execucao industrial.
7. Expedicao e Transportes: guia, carga, rota e comprovativo de entrega.
8. Notas Encomenda e Faturacao: compra, rececao, documentos e controlo financeiro.

## Seguranca comercial
- Este pacote nao inclui credenciais, dados da base de desenvolvimento, perfis pessoais ou passwords internas.
- O SQL cria apenas a estrutura e utilizadores temporarios.
- Usar uma conta MySQL dedicada; nao usar `root` nos postos cliente.
- Definir uma pasta partilhada segura para documentos quando existirem varios postos.
- O trial exige Internet para validar a hora portuguesa por fontes HTTPS independentes e bloqueia se nao conseguir validar.
"@
Write-Utf8NoBomFile -Path (Join-Path $docsDir 'GUIA - Apresentacao Comercial.md') -Content $presentationReadme

$trialGuide = @"
# Ativacao segura do trial

## Primeira ativacao
1. Instalar o luGEST com `INSTALAR_LUISGEST.ps1`.
2. Confirmar a ligacao MySQL no ficheiro `lugest.env` da pasta instalada (por omissao, `%LOCALAPPDATA%\luGEST`).
3. Executar `CONFIGURAR_OWNER_TRIAL.ps1` nessa pasta.
4. Definir um utilizador OWNER e uma password forte, diferente das contas normais.
5. Guardar estas credenciais num gestor de passwords. O ficheiro guarda apenas o hash PBKDF2-SHA256.
6. Garantir Internet e HTTPS/443 para `www.google.com` e `www.cloudflare.com`; `www.microsoft.com` e a contingencia.
7. Abrir o LuisGEST e autenticar com a conta OWNER.
8. Abrir `Extras > Trial / licenca`.
9. Clicar em `Validar hora agora`, preencher a empresa e a duracao em dias.
10. Clicar em `Ativar / reiniciar`.
11. Sair da sessao OWNER e entrar com um utilizador normal.

## Prolongar ou terminar
- Para prolongar: entrar como OWNER, abrir `Extras > Trial / licenca`, indicar os dias e clicar em `Prolongar`.
- Para terminar: entrar como OWNER e usar `Desativar trial`.
- A conta OWNER serve apenas para administracao comercial/licenciamento; nao deve ser entregue aos utilizadores operacionais.

## Relogio e Internet
- O trial nao usa o relogio editavel do Windows como autoridade.
- A aplicacao compara fontes HTTPS independentes, usa UTC e apresenta a hora no fuso de Portugal.
- Sem uma validacao online coerente, ou perante regressao temporal, o acesso do trial e bloqueado.
- Mudar manualmente a data/hora do computador nao prolonga o prazo.
"@
Write-Utf8NoBomFile -Path (Join-Path $docsDir 'GUIA - Ativacao Trial.md') -Content $trialGuide

$readme = @"
# luGEST Desktop

Preparado em: $releaseDateTxt

Esta pasta contem a aplicacao desktop luGEST preparada para instalacao e validacao controlada.

## O que interessa
- luGEST.exe: aplicacao principal.
- INSTALAR_LUISGEST.ps1: instala por omissao em `%LOCALAPPDATA%\luGEST`, sem exigir administrador; aceita outro destino com `-InstallRoot`.
- DESINSTALAR_LUGEST.ps1: remove a aplicacao e preserva configuracao/dados locais em `Documentos\luGEST Backups`.
- CONFIGURAR_OWNER_TRIAL.ps1: cria, depois da instalacao, as credenciais privadas para gerir o trial.
- _internal: motor interno do executavel; nao apagar nem copiar o luGEST.exe sozinho.
- lugest.env: configuracao da ligacao MySQL deste posto.
- lugest.env.servidor.example: exemplo para o computador servidor.
- lugest.env.posto.example: exemplo para os outros postos.
- lugest_branding.json, lugest_qt_config.json e Logos: configuracao visual e PDFs.
- Base de Dados\mysql: SQL unico para importar no HeidiSQL.
- Documentacao: guias essenciais de instalacao e checklist.

## Como arrancar
1. No cliente, clicar com botao direito em INSTALAR_LUISGEST.ps1 e escolher Executar com PowerShell.
2. Confirmar o ficheiro `lugest.env` na pasta instalada com os dados MySQL corretos.
3. Executar `CONFIGURAR_OWNER_TRIAL.ps1` nessa pasta e guardar as credenciais OWNER em local seguro.
4. Abrir pelo atalho luGEST, entrar como OWNER e ativar em Extras > Trial / licenca.
5. Em multiutilizador, todos os postos devem apontar para a mesma base MySQL.
6. Se houver ficheiros partilhados, usar em lugest.env uma pasta UNC comum em LUGEST_SHARED_STORAGE_ROOT.

## Requisito do trial
- O computador precisa de Internet enquanto o trial estiver ativo.
- Permitir HTTPS/443 para www.google.com e www.cloudflare.com; www.microsoft.com serve de contingencia.
- A hora e validada online e convertida automaticamente para o fuso de Portugal.
- Alterar manualmente o relogio do Windows nao prolonga o trial.

## Notas
- Os atalhos .bat foram removidos para reduzir confusao; a instalacao usa apenas PowerShell e depois arranca pelo atalho.
- A pasta por omissao e do utilizador atual. Para escolher outra: `powershell -ExecutionPolicy Bypass -File .\INSTALAR_LUISGEST.ps1 -InstallRoot "D:\Aplicacoes\luGEST"`.
- Para instalar base nova no HeidiSQL, importar Base de Dados\mysql\IMPORTAR_NO_HEIDI.sql.
"@
Write-Utf8NoBomFile -Path (Join-Path $releaseRoot 'README.md') -Content $readme

Write-Output $releaseRoot
