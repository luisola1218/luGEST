param(
    [switch]$Commercial
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$venvPython = Join-Path $repoRoot '.venv\Scripts\python.exe'
$desktopRoot = Join-Path $env:USERPROFILE 'Desktop'
$releaseDate = Get-Date
$releaseDateTxt = $releaseDate.ToString('dd/MM/yyyy HH:mm')
$releaseName = if ($Commercial) { 'LuisGEST - Pacote Comercial Piloto' } else { 'App LuisGEST - Revis' + [char]0x00E3 + 'o Final' }
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
$preservedConfigRoot = Join-Path $env:TEMP ("lugest_release_config_" + [guid]::NewGuid().ToString('N'))
$preservedEnv = Join-Path $preservedConfigRoot 'lugest.env'
$preservedTrial = Join-Path $preservedConfigRoot 'lugest_trial.json'
$preservedBackups = Join-Path $preservedConfigRoot 'Backups'

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
    $updateGuide,
    $readinessGuide,
    $fluidityAudit
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

if ($Commercial) {
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
    [string]`$InstallRoot = 'C:\LuisGEST',
    [switch]`$NoShortcut,
    [switch]`$NoPause
)

`$ErrorActionPreference = 'Stop'

`$sourceRoot = Split-Path -Parent `$MyInvocation.MyCommand.Path
`$installRoot = `$InstallRoot
`$desktopShortcut = Join-Path `$env:USERPROFILE 'Desktop\LuisGEST.lnk'
`$preserveRoot = Join-Path `$env:TEMP ('lugest_install_' + [guid]::NewGuid().ToString('N'))

if (-not (Test-Path (Join-Path `$sourceRoot 'LuisGEST.exe'))) {
    throw 'Esta pasta nao contem LuisGEST.exe.'
}
if (-not (Test-Path (Join-Path `$sourceRoot '_internal\PySide6'))) {
    throw 'Instalacao incompleta: falta a pasta _internal\PySide6. Copia a pasta LuisGEST completa.'
}

New-Item -ItemType Directory -Force -Path `$installRoot, `$preserveRoot | Out-Null
foreach (`$name in @('lugest.env', 'lugest_trial.json', 'lugest_branding.json', 'lugest_qt_config.json', 'Logos')) {
    `$existing = Join-Path `$installRoot `$name
    if (Test-Path `$existing) {
        Copy-Item `$existing (Join-Path `$preserveRoot `$name) -Recurse -Force
    }
}
try {
    Get-ChildItem `$installRoot -Force -ErrorAction SilentlyContinue |
        Where-Object { `$_.Name -notin @('generated', 'backups') } |
        ForEach-Object { Remove-Item `$_.FullName -Recurse -Force }

    Copy-Item -Path (Join-Path `$sourceRoot '*') -Destination `$installRoot -Recurse -Force
    Get-ChildItem `$preserveRoot -Force -ErrorAction SilentlyContinue |
        ForEach-Object { Copy-Item `$_.FullName (Join-Path `$installRoot `$_.Name) -Recurse -Force }
}
finally {
    Remove-Item `$preserveRoot -Recurse -Force -ErrorAction SilentlyContinue
}

if (-not `$NoShortcut) {
    `$shell = New-Object -ComObject WScript.Shell
    `$shortcut = `$shell.CreateShortcut(`$desktopShortcut)
    `$shortcut.TargetPath = Join-Path `$installRoot 'LuisGEST.exe'
    `$shortcut.WorkingDirectory = `$installRoot
    `$shortcut.IconLocation = Join-Path `$installRoot 'app.ico'
    `$shortcut.Description = 'LuisGEST ERP industrial'
    `$shortcut.Save()
}

Write-Host ''
Write-Host "LuisGEST instalado com sucesso em `$installRoot"
if (-not `$NoShortcut) { Write-Host 'Atalho criado no Ambiente de Trabalho: LuisGEST' }
Write-Host ''
Write-Host "Se este posto for cliente, confirma o ficheiro `$installRoot\lugest.env com o IP/nome do servidor MySQL."
if (-not `$NoPause) { Pause }
"@
Write-Utf8NoBomFile -Path (Join-Path $releaseRoot 'INSTALAR_LUISGEST.ps1') -Content $installer

if (Test-Path $venvPython) {
    & $venvPython (Join-Path $databaseSource 'export_current_schema_sql.py') --with-starter-users --output (Join-Path $databaseSource 'lugest.sql') | Out-Null
}
Copy-Item (Join-Path $databaseSource 'lugest.sql') (Join-Path $mysqlDir 'IMPORTAR_NO_HEIDI.sql') -Force
$dbReadme = @"
# Base de Dados LuisGEST

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

$presentationReadme = @"
# Guia de Apresentacao Comercial

## Antes da demonstracao
1. Importar `Base de Dados\mysql\IMPORTAR_NO_HEIDI.sql` numa instalacao MySQL limpa.
2. Configurar `lugest.env` com o servidor, utilizador e password dedicados ao cliente.
3. Executar `INSTALAR_LUISGEST.ps1` com PowerShell.
4. Entrar com `admin / Trocar#Admin2026` e alterar imediatamente as passwords temporarias.
5. Configurar logotipo, dados da empresa, operadores e parametros comerciais antes de usar em producao.

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
"@
Write-Utf8NoBomFile -Path (Join-Path $docsDir 'GUIA - Apresentacao Comercial.md') -Content $presentationReadme

$readme = @"
# LuisGEST Desktop

Preparado em: $releaseDateTxt

Esta pasta contem a aplicacao desktop LuisGEST preparada para instalacao e validacao controlada.

## O que interessa
- LuisGEST.exe: aplicacao principal.
- INSTALAR_LUISGEST.ps1: instala a app em C:\LuisGEST e cria atalho no Ambiente de Trabalho.
- _internal: motor interno do executavel; nao apagar nem copiar o LuisGEST.exe sozinho.
- lugest.env: configuracao da ligacao MySQL deste posto.
- lugest.env.servidor.example: exemplo para o computador servidor.
- lugest.env.posto.example: exemplo para os outros postos.
- lugest_branding.json, lugest_qt_config.json e Logos: configuracao visual e PDFs.
- Base de Dados\mysql: SQL unico para importar no HeidiSQL.
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
