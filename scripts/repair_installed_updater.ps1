param(
    [string]$InstallDir = "",
    [string]$ManifestUrl = "",
    [string]$GitHubToken = "",
    [string]$CurrentVersion = ""
)

$ErrorActionPreference = "Stop"

function Write-Info($message) {
    Write-Host "[LuisGEST Repair] $message"
}

function Read-JsonFile {
    param([string]$Path)
    if (-not (Test-Path $Path)) {
        return [pscustomobject]@{}
    }
    $raw = Get-Content $Path -Raw -Encoding UTF8
    if (-not $raw.Trim()) {
        return [pscustomobject]@{}
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

function Copy-ItemIfDifferent {
    param(
        [string]$Source,
        [string]$Destination
    )
    $sourceResolved = ""
    $destinationResolved = ""
    if (Test-Path $Source) {
        $sourceResolved = (Resolve-Path $Source).Path
    }
    if (Test-Path $Destination) {
        $destinationResolved = (Resolve-Path $Destination).Path
    }
    if ($sourceResolved -and $destinationResolved -and ($sourceResolved -ieq $destinationResolved)) {
        return
    }
    Copy-Item $Source $Destination -Force
}

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

function Resolve-RelativePathOrUrl {
    param([string]$Value, [string]$BaseDir)
    $txt = [string]$Value
    $txt = $txt.Trim()
    if (-not $txt) {
        return ""
    }
    if ($txt -match '^https?://') {
        return $txt
    }
    if ($txt -match '^file:///') {
        return ([System.Uri]$txt).LocalPath
    }
    $baseTxt = [string]$BaseDir
    $baseTxt = $baseTxt.Trim()
    if ($baseTxt -match '^https?://') {
        return ([System.Uri]::new([System.Uri]$baseTxt, $txt)).AbsoluteUri
    }
    if ([System.IO.Path]::IsPathRooted($txt)) {
        return $txt
    }
    return (Join-Path $BaseDir $txt)
}

function Get-DownloadHeaders {
    param([string]$Url, [string]$Token, [switch]$BinaryAsset)
    $headers = @{}
    if ($Token) {
        $headers['Authorization'] = "Bearer $Token"
    }
    if ($BinaryAsset) {
        $headers['Accept'] = 'application/octet-stream'
    }
    return $headers
}

function Resolve-GitHubReleaseAssetApiUrl {
    param([string]$Url, [string]$Token)
    if (-not $Token) {
        return $null
    }
    $txt = [string]$Url
    $tagMatch = [regex]::Match($txt, '^https://github\.com/(?<owner>[^/]+)/(?<repo>[^/]+)/releases/download/(?<tag>[^/]+)/(?<asset>[^/?#]+)$')
    $latestMatch = [regex]::Match($txt, '^https://github\.com/(?<owner>[^/]+)/(?<repo>[^/]+)/releases/latest/download/(?<asset>[^/?#]+)$')
    $match = if ($tagMatch.Success) { $tagMatch } elseif ($latestMatch.Success) { $latestMatch } else { $null }
    if ($null -eq $match) {
        return $null
    }
    $owner = $match.Groups['owner'].Value
    $repo = $match.Groups['repo'].Value
    $assetName = [System.Uri]::UnescapeDataString($match.Groups['asset'].Value)
    $headers = Get-DownloadHeaders "https://api.github.com/" $Token
    $headers['User-Agent'] = 'LuisGEST-Updater'
    $headers['Accept'] = 'application/vnd.github+json'
    $releaseApiUrl = "https://api.github.com/repos/$owner/$repo/releases/latest"
    if ($tagMatch.Success) {
        $tag = $tagMatch.Groups['tag'].Value
        $releaseApiUrl = "https://api.github.com/repos/$owner/$repo/releases/tags/$tag"
    }
    $release = Invoke-RestMethod -Uri $releaseApiUrl -Headers $headers -Method Get
    foreach ($asset in $release.assets) {
        if ([string]$asset.name -eq $assetName) {
            return [string]$asset.url
        }
    }
    throw "Asset GitHub nao encontrado na release: $assetName"
}

function Download-RemoteFile {
    param([string]$Url, [string]$TargetPath, [string]$Token)
    $txt = [string]$Url
    $assetApiUrl = Resolve-GitHubReleaseAssetApiUrl $txt $Token
    if ($assetApiUrl) {
        $headers = Get-DownloadHeaders $assetApiUrl $Token -BinaryAsset
        $headers['User-Agent'] = 'LuisGEST-Updater'
        Invoke-WebRequest -Uri $assetApiUrl -OutFile $TargetPath -Headers $headers -UseBasicParsing
        return
    }
    if ($txt -match '^https://api\.github\.com/repos/.+/releases/assets/\d+$') {
        $headers = Get-DownloadHeaders $txt $Token -BinaryAsset
        $headers['User-Agent'] = 'LuisGEST-Updater'
        Invoke-WebRequest -Uri $txt -OutFile $TargetPath -Headers $headers -UseBasicParsing
        return
    }
    $headers = Get-DownloadHeaders $txt $Token
    if ($Token) {
        $headers['User-Agent'] = 'LuisGEST-Updater'
    }
    Invoke-WebRequest -Uri $txt -OutFile $TargetPath -Headers $headers -UseBasicParsing
}

$installRoot = Resolve-InstallDir $InstallDir
$qtConfigPath = Join-Path $installRoot "lugest_qt_config.json"
$updateConfigPath = Join-Path $installRoot "update_config.json"
$versionPath = Join-Path $installRoot "VERSION"
$updatePs1Target = Join-Path $installRoot "Atualizar LuisGEST.ps1"
$updateBatTarget = Join-Path $installRoot "Atualizar LuisGEST.bat"
$repairPs1Target = Join-Path $installRoot "Reparar Atualizador Instalado.ps1"
$repairBatTarget = Join-Path $installRoot "Reparar Atualizador Instalado.bat"

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
if (-not $CurrentVersion -and (Test-Path $versionPath)) {
    $CurrentVersion = (Get-Content $versionPath -Raw -Encoding UTF8).Trim()
}
if (-not $CurrentVersion) {
    $CurrentVersion = "0.0.0"
}

$workDir = Join-Path $env:TEMP ("lugest_repair_" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workDir -Force | Out-Null

try {
    $sourceRoot = ""
    $releaseUpdatePs1 = Join-Path $PSScriptRoot "Atualizar LuisGEST.ps1"
    $releaseUpdateBat = Join-Path $PSScriptRoot "Atualizar LuisGEST.bat"
    $releaseRepairPs1 = Join-Path $PSScriptRoot "Reparar Atualizador Instalado.ps1"
    $releaseRepairBat = Join-Path $PSScriptRoot "Reparar Atualizador Instalado.bat"

    if ((Test-Path $releaseUpdatePs1) -and (Test-Path $releaseUpdateBat) -and (Test-Path $releaseRepairPs1) -and (Test-Path $releaseRepairBat)) {
        $sourceRoot = (Resolve-Path $PSScriptRoot).Path
        Write-Info "A usar os scripts locais da release."
    }
    else {
        Write-Info "A descarregar a release remota para reparar o atualizador instalado..."
        $manifestPath = Join-Path $workDir "latest.json"
        Download-RemoteFile $ManifestUrl $manifestPath $GitHubToken
        $manifest = Read-JsonFile $manifestPath
        $packageRef = [string](Get-JsonValue $manifest "package_url" "")
        if (-not $packageRef) {
            throw "Manifest sem package_url."
        }
        $packageResolved = Resolve-RelativePathOrUrl $packageRef $ManifestUrl
        $packagePath = Join-Path $workDir "package.zip"
        Download-RemoteFile $packageResolved $packagePath $GitHubToken

        $expectedHash = [string](Get-JsonValue $manifest "sha256" "")
        if ($expectedHash) {
            $actualHash = (Get-FileHash $packagePath -Algorithm SHA256).Hash
            if ($actualHash.ToLowerInvariant() -ne $expectedHash.ToLowerInvariant()) {
                throw "Checksum invalido. Esperado $expectedHash, obtido $actualHash."
            }
        }

        $extractDir = Join-Path $workDir "extract"
        Expand-Archive -Path $packagePath -DestinationPath $extractDir -Force
        $desktopCandidate = Join-Path $extractDir "Desktop App"
        $sourceRoot = if (Test-Path $desktopCandidate) { $desktopCandidate } else { $extractDir }
        $releaseUpdatePs1 = Join-Path $sourceRoot "Atualizar LuisGEST.ps1"
        $releaseUpdateBat = Join-Path $sourceRoot "Atualizar LuisGEST.bat"
        $releaseRepairPs1 = Join-Path $sourceRoot "Reparar Atualizador Instalado.ps1"
        $releaseRepairBat = Join-Path $sourceRoot "Reparar Atualizador Instalado.bat"
    }

    if (-not (Test-Path $releaseUpdatePs1)) {
        throw "Nao foi encontrado 'Atualizar LuisGEST.ps1' na release."
    }
    if (-not (Test-Path $releaseUpdateBat)) {
        throw "Nao foi encontrado 'Atualizar LuisGEST.bat' na release."
    }
    if (-not (Test-Path $releaseRepairPs1)) {
        throw "Nao foi encontrado 'Reparar Atualizador Instalado.ps1' na release."
    }
    if (-not (Test-Path $releaseRepairBat)) {
        throw "Nao foi encontrado 'Reparar Atualizador Instalado.bat' na release."
    }

    Copy-ItemIfDifferent $releaseUpdatePs1 $updatePs1Target
    Copy-ItemIfDifferent $releaseUpdateBat $updateBatTarget
    Copy-ItemIfDifferent $releaseRepairPs1 $repairPs1Target
    Copy-ItemIfDifferent $releaseRepairBat $repairBatTarget

    $payload = @{
        current_version = $CurrentVersion
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
    $updaterArgs = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", $releaseUpdatePs1,
        "-AppDir", $installRoot,
        "-ManifestUrl", $ManifestUrl,
        "-CurrentVersion", $CurrentVersion
    )
    if ($GitHubToken) {
        $updaterArgs += @("-GitHubToken", $GitHubToken)
    }
    & $powershellExe @updaterArgs
}
catch {
    Write-Host ""
    Write-Host "ERRO ao reparar o atualizador do LuisGEST:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Read-Host "Carrega Enter para fechar"
    exit 1
}
