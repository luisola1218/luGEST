param(
    [string]$InstallDir = "",
    [string]$ManifestUrl = "",
    [string]$GitHubToken = ""
)

$ErrorActionPreference = "Stop"

function Resolve-InstallDir {
    param([string]$Preferred)
    $candidates = @()
    if ($Preferred) {
        $candidates += $Preferred
    }
    $candidates += @(
        (Join-Path $env:ProgramFiles "LuisGEST"),
        (Join-Path $env:LOCALAPPDATA "LuisGEST"),
        (Join-Path $env:LOCALAPPDATA "LuisGEST-Cliente")
    )
    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path (Join-Path $candidate "VERSION"))) {
            return (Resolve-Path $candidate).Path
        }
    }
    throw "Nao foi encontrada uma instalacao LuisGEST valida."
}

function Read-JsonFile {
    param([string]$Path)
    if (-not (Test-Path $Path)) {
        return @{}
    }
    $raw = Get-Content $Path -Raw -Encoding UTF8
    if (-not $raw.Trim()) {
        return @{}
    }
    return $raw | ConvertFrom-Json
}

function Get-JsonValue {
    param($Obj, [string]$Name, $Default = "")
    if ($null -eq $Obj) {
        return $Default
    }
    $prop = $Obj.PSObject.Properties[$Name]
    if ($null -eq $prop -or $null -eq $prop.Value) {
        return $Default
    }
    return $prop.Value
}

function Save-JsonFile {
    param([string]$Path, $Payload)
    $Payload | ConvertTo-Json -Depth 8 | Set-Content -Path $Path -Encoding UTF8
}

$installRoot = Resolve-InstallDir $InstallDir
$sourceRoot = (Resolve-Path $PSScriptRoot).Path
$qtConfigPath = Join-Path $installRoot "lugest_qt_config.json"
$updateConfigPath = Join-Path $installRoot "update_config.json"
$updatePs1Target = Join-Path $installRoot "Atualizar LuisGEST.ps1"
$updateBatTarget = Join-Path $installRoot "Atualizar LuisGEST.bat"

$qtConfig = Read-JsonFile $qtConfigPath
$qtUpdate = Get-JsonValue $qtConfig "update_settings" @{}

if (-not $ManifestUrl) {
    $ManifestUrl = [string](Get-JsonValue $qtUpdate "manifest_url" "")
}
if (-not $GitHubToken) {
    $GitHubToken = [string](Get-JsonValue $qtUpdate "github_token" "")
}

if (-not $ManifestUrl) {
    throw "Nao foi possivel determinar o manifest_url. Abre a app, guarda primeiro o URL em Extras > Atualizacoes, ou passa -ManifestUrl."
}

$version = "0.0.0"
$versionPath = Join-Path $installRoot "VERSION"
if (Test-Path $versionPath) {
    $version = (Get-Content $versionPath -Raw -Encoding UTF8).Trim()
}

$releaseUpdatePs1 = Join-Path $sourceRoot "Atualizar LuisGEST.ps1"
$releaseUpdateBat = Join-Path $sourceRoot "Atualizar LuisGEST.bat"
if (-not (Test-Path $releaseUpdatePs1)) {
    throw "Nao foi encontrado 'Atualizar LuisGEST.ps1' nesta pasta. Executa este reparador a partir da Desktop App da release."
}
if (-not (Test-Path $releaseUpdateBat)) {
    throw "Nao foi encontrado 'Atualizar LuisGEST.bat' nesta pasta. Executa este reparador a partir da Desktop App da release."
}

Copy-Item $releaseUpdatePs1 $updatePs1Target -Force
Copy-Item $releaseUpdateBat $updateBatTarget -Force

$payload = @{
    current_version = $version
    manifest_url = $ManifestUrl
    channel = "stable"
    github_token = $GitHubToken
    auto_check = [bool](Get-JsonValue $qtUpdate "auto_check" $false)
}
Save-JsonFile $updateConfigPath $payload

Write-Host ""
Write-Host "Atualizador reparado em: $installRoot"
Write-Host "Manifest: $ManifestUrl"
if ($GitHubToken) {
    Write-Host "Token GitHub: configurado"
}
else {
    Write-Host "Token GitHub: vazio"
}
Write-Host ""
Write-Host "Fecha o LuisGEST se ainda estiver aberto."
Read-Host "Carrega Enter para lancar o atualizador reparado"

$powershellExe = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
& $powershellExe -NoProfile -ExecutionPolicy Bypass -File $updatePs1Target -AppDir $installRoot
