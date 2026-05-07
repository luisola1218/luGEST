param(
    [string]$InstallDir = "",
    [string]$ManifestUrl = "",
    [string]$GitHubToken = "",
    [string]$CurrentVersion = ""
)

$ErrorActionPreference = "Stop"

function Write-Info($message) {
    Write-Host "[LuisGEST Bootstrap] $message"
}

function Read-JsonFile($path) {
    if (-not (Test-Path $path)) {
        return [pscustomobject]@{}
    }
    $raw = Get-Content $path -Raw -Encoding UTF8
    if (-not $raw.Trim()) {
        return [pscustomobject]@{}
    }
    return $raw | ConvertFrom-Json
}

function Save-JsonFile($path, $payload) {
    $payload | ConvertTo-Json -Depth 8 | Set-Content -Path $path -Encoding UTF8
}

function Get-JsonValue($obj, $name, $default = "") {
    if ($null -eq $obj) {
        return $default
    }
    $prop = $obj.PSObject.Properties[$name]
    if ($null -eq $prop -or $null -eq $prop.Value) {
        return $default
    }
    return $prop.Value
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

function Resolve-RelativePathOrUrl($value, $baseDir) {
    $txt = [string]$value
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
    $baseTxt = [string]$baseDir
    $baseTxt = $baseTxt.Trim()
    if ($baseTxt -match '^https?://') {
        return ([System.Uri]::new([System.Uri]$baseTxt, $txt)).AbsoluteUri
    }
    if ([System.IO.Path]::IsPathRooted($txt)) {
        return $txt
    }
    return (Join-Path $baseDir $txt)
}

function Get-DownloadHeaders($url, $token, [switch]$BinaryAsset) {
    $headers = @{}
    if ($token) {
        $headers['Authorization'] = "Bearer $token"
    }
    if ($BinaryAsset) {
        $headers['Accept'] = 'application/octet-stream'
    }
    return $headers
}

function Resolve-GitHubReleaseAssetApiUrl($url, $token) {
    if (-not $token) {
        return $null
    }
    $txt = [string]$url
    $tagMatch = [regex]::Match($txt, '^https://github\.com/(?<owner>[^/]+)/(?<repo>[^/]+)/releases/download/(?<tag>[^/]+)/(?<asset>[^/?#]+)$')
    $latestMatch = [regex]::Match($txt, '^https://github\.com/(?<owner>[^/]+)/(?<repo>[^/]+)/releases/latest/download/(?<asset>[^/?#]+)$')
    $match = if ($tagMatch.Success) { $tagMatch } elseif ($latestMatch.Success) { $latestMatch } else { $null }
    if ($null -eq $match) {
        return $null
    }
    $owner = $match.Groups['owner'].Value
    $repo = $match.Groups['repo'].Value
    $assetName = [System.Uri]::UnescapeDataString($match.Groups['asset'].Value)
    $headers = Get-DownloadHeaders "https://api.github.com/" $token
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
    throw "Asset GitHub nao encontrado na release ${tag}: $assetName"
}

function Download-RemoteFile($url, $targetPath, $token) {
    $txt = [string]$url
    $assetApiUrl = Resolve-GitHubReleaseAssetApiUrl $txt $token
    if ($assetApiUrl) {
        $headers = Get-DownloadHeaders $assetApiUrl $token -BinaryAsset
        $headers['User-Agent'] = 'LuisGEST-Updater'
        Invoke-WebRequest -Uri $assetApiUrl -OutFile $targetPath -Headers $headers -UseBasicParsing
        return
    }
    if ($txt -match '^https://api\.github\.com/repos/.+/releases/assets/\d+$') {
        $headers = Get-DownloadHeaders $txt $token -BinaryAsset
        $headers['User-Agent'] = 'LuisGEST-Updater'
        Invoke-WebRequest -Uri $txt -OutFile $targetPath -Headers $headers -UseBasicParsing
        return
    }
    $headers = Get-DownloadHeaders $txt $token
    if ($token) {
        $headers['User-Agent'] = 'LuisGEST-Updater'
    }
    Invoke-WebRequest -Uri $txt -OutFile $targetPath -Headers $headers -UseBasicParsing
}

try {
    $installRoot = Resolve-InstallDir $InstallDir
    $versionPath = Join-Path $installRoot "VERSION"
    if (-not $CurrentVersion -and (Test-Path $versionPath)) {
        $CurrentVersion = (Get-Content $versionPath -Raw -Encoding UTF8).Trim()
    }
    if (-not $CurrentVersion) {
        $CurrentVersion = "0.0.0"
    }

    $qtConfigPath = Join-Path $installRoot "lugest_qt_config.json"
    $updateConfigPath = Join-Path $installRoot "update_config.json"
    $qtConfig = Read-JsonFile $qtConfigPath
    $qtUpdate = Get-JsonValue $qtConfig "update_settings" @{}
    if (-not $ManifestUrl) {
        $ManifestUrl = [string](Get-JsonValue $qtUpdate "manifest_url" "")
    }
    if (-not $GitHubToken) {
        $GitHubToken = [string](Get-JsonValue $qtUpdate "github_token" "")
    }
    if (-not $ManifestUrl) {
        throw "Nao foi possivel determinar o manifest_url."
    }

    $workDir = Join-Path $env:TEMP ("lugest_remote_bootstrap_" + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $workDir -Force | Out-Null

    $manifestPath = Join-Path $workDir "latest.json"
    Write-Info "A descarregar manifest remoto..."
    Download-RemoteFile $ManifestUrl $manifestPath $GitHubToken
    $manifest = Read-JsonFile $manifestPath
    $latestVersion = [string](Get-JsonValue $manifest "version" "")
    if (-not $latestVersion) {
        throw "Manifest sem campo version."
    }
    $packageRef = [string](Get-JsonValue $manifest "package_url" "")
    if (-not $packageRef) {
        throw "Manifest sem package_url."
    }
    $packageResolved = Resolve-RelativePathOrUrl $packageRef $ManifestUrl
    $packagePath = Join-Path $workDir "package.zip"
    Write-Info "A descarregar pacote da release $latestVersion..."
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
    $desktopDir = Join-Path $extractDir "Desktop App"
    if (-not (Test-Path $desktopDir)) {
        $desktopDir = $extractDir
    }
    $updaterScript = Join-Path $desktopDir "Atualizar LuisGEST.ps1"
    if (-not (Test-Path $updaterScript)) {
        throw "Pacote sem Atualizar LuisGEST.ps1."
    }

    $payload = @{
        current_version = $CurrentVersion
        manifest_url = $ManifestUrl
        channel = "stable"
        github_token = $GitHubToken
        auto_check = [bool](Get-JsonValue $qtUpdate "auto_check" $false)
    }
    Save-JsonFile $updateConfigPath $payload

    Write-Host ""
    Write-Info "Bootstrap preparado. O atualizador mais recente vai arrancar agora."
    Write-Info "Fecha o LuisGEST quando o atualizador pedir."
    $powershellExe = Join-Path $env:SystemRoot "System32\\WindowsPowerShell\\v1.0\\powershell.exe"
    & $powershellExe `
        -NoProfile `
        -ExecutionPolicy Bypass `
        -File $updaterScript `
        -AppDir $installRoot `
        -ManifestUrl $ManifestUrl `
        -GitHubToken $GitHubToken `
        -CurrentVersion $CurrentVersion `
        -PackageFileOverride $packagePath
}
catch {
    Write-Host ""
    Write-Host "ERRO no bootstrap remoto do LuisGEST:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Read-Host "Carrega Enter para fechar"
    exit 1
}
