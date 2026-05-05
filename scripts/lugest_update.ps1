param(
    [string]$AppDir = "",
    [string]$ConfigPath = "",
    [switch]$CheckOnly,
    [switch]$Force,
    [switch]$NoRestart
)

$ErrorActionPreference = 'Stop'

function Write-Info($message) {
    Write-Host "[LuisGEST Update] $message"
}

function Read-JsonFile($path) {
    if (-not (Test-Path $path)) {
        return [pscustomobject]@{}
    }
    $raw = Get-Content $path -Raw -Encoding UTF8
    if (-not $raw.Trim()) {
        return @{}
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

function Resolve-AppDir {
    if ($AppDir -and (Test-Path $AppDir)) {
        return (Resolve-Path $AppDir).Path
    }
    return (Resolve-Path $PSScriptRoot).Path
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
    $match = [regex]::Match($txt, '^https://github\.com/(?<owner>[^/]+)/(?<repo>[^/]+)/releases/download/(?<tag>[^/]+)/(?<asset>[^/?#]+)$')
    if (-not $match.Success) {
        return $null
    }
    $owner = $match.Groups['owner'].Value
    $repo = $match.Groups['repo'].Value
    $tag = $match.Groups['tag'].Value
    $assetName = [System.Uri]::UnescapeDataString($match.Groups['asset'].Value)
    $headers = Get-DownloadHeaders "https://api.github.com/" $token
    $headers['User-Agent'] = 'LuisGEST-Updater'
    $headers['Accept'] = 'application/vnd.github+json'
    $release = Invoke-RestMethod -Uri "https://api.github.com/repos/$owner/$repo/releases/tags/$tag" -Headers $headers -Method Get
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

function Read-Manifest($manifestRef, $config, $workDir) {
    $manifestPath = Join-Path $workDir 'latest.json'
    if ($manifestRef -match '^https?://') {
        $token = [string](Get-JsonValue $config 'github_token' '')
        Download-RemoteFile $manifestRef $manifestPath $token
        return [pscustomobject]@{ Manifest = (Read-JsonFile $manifestPath); LocalPath = $manifestPath; SourceRef = $manifestRef }
    }
    $resolved = Resolve-RelativePathOrUrl $manifestRef $appDir
    if (-not (Test-Path $resolved)) {
        throw "Manifest nao encontrado: $resolved"
    }
    Copy-Item $resolved $manifestPath -Force
    return [pscustomobject]@{ Manifest = (Read-JsonFile $manifestPath); LocalPath = (Resolve-Path $resolved).Path; SourceRef = (Resolve-Path $resolved).Path }
}

function Version-Parts($value) {
    $txt = [string]$value
    $matches = [regex]::Matches($txt, '\d+')
    $parts = @()
    foreach ($m in $matches) {
        $parts += [int]$m.Value
    }
    while ($parts.Count -lt 4) {
        $parts += 0
    }
    return $parts[0..3]
}

function Compare-Version($left, $right) {
    $a = Version-Parts $left
    $b = Version-Parts $right
    for ($i = 0; $i -lt 4; $i++) {
        if ($a[$i] -lt $b[$i]) { return -1 }
        if ($a[$i] -gt $b[$i]) { return 1 }
    }
    return 0
}

function Load-EnvFile($path) {
    $envMap = @{}
    if (-not (Test-Path $path)) {
        return $envMap
    }
    foreach ($line in Get-Content $path -Encoding UTF8) {
        $txt = [string]$line
        if (-not $txt.Trim() -or $txt.Trim().StartsWith('#') -or $txt -notmatch '=') {
            continue
        }
        $key, $value = $txt.Split('=', 2)
        $envMap[$key.Trim()] = $value.Trim()
    }
    return $envMap
}

function Backup-DatabaseBestEffort($appDir, $backupDir) {
    $envPath = Join-Path $appDir 'lugest.env'
    $envMap = Load-EnvFile $envPath
    if (-not $envMap.ContainsKey('LUGEST_DB_NAME')) {
        Write-Info "Backup MySQL ignorado: lugest.env sem LUGEST_DB_NAME."
        return
    }
    $dump = Get-Command mysqldump -ErrorAction SilentlyContinue
    if (-not $dump) {
        Write-Info "Backup MySQL ignorado: mysqldump nao esta no PATH."
        return
    }
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $target = Join-Path $backupDir "mysql_backup_$stamp.sql"
    $hostName = '127.0.0.1'
    if ($envMap.ContainsKey('LUGEST_DB_HOST')) { $hostName = [string]$envMap['LUGEST_DB_HOST'] }
    $port = '3306'
    if ($envMap.ContainsKey('LUGEST_DB_PORT')) { $port = [string]$envMap['LUGEST_DB_PORT'] }
    $user = ''
    if ($envMap.ContainsKey('LUGEST_DB_USER')) { $user = [string]$envMap['LUGEST_DB_USER'] }
    $pass = ''
    if ($envMap.ContainsKey('LUGEST_DB_PASS')) { $pass = [string]$envMap['LUGEST_DB_PASS'] }
    $db = ''
    if ($envMap.ContainsKey('LUGEST_DB_NAME')) { $db = [string]$envMap['LUGEST_DB_NAME'] }
    if (-not $user -or -not $db) {
        Write-Info "Backup MySQL ignorado: utilizador/base nao configurados."
        return
    }
    Write-Info "A criar backup MySQL em $target"
    & $dump.Source -h $hostName -P $port -u $user "-p$pass" $db --result-file="$target"
}

function Assert-AppClosed($appDir) {
    $exeNames = @('lugest_qt.exe', 'main.exe')
    foreach ($name in $exeNames) {
        $running = Get-Process -Name ([System.IO.Path]::GetFileNameWithoutExtension($name)) -ErrorAction SilentlyContinue
        if ($running) {
            Write-Host ""
            Write-Host "Fecha o LuisGEST antes de continuar. Processo detetado: $name"
            if (-not $Force) {
                Read-Host "Depois de fechar, carrega Enter"
            }
        }
    }
}

$appDir = Resolve-AppDir
if (-not $ConfigPath) {
    $ConfigPath = Join-Path $appDir 'update_config.json'
}
$config = Read-JsonFile $ConfigPath
$versionPath = Join-Path $appDir 'VERSION'
$currentVersion = [string](Get-JsonValue $config 'current_version' '')
if ((Test-Path $versionPath) -and -not $currentVersion) {
    $currentVersion = (Get-Content $versionPath -Raw -Encoding UTF8).Trim()
}
if (-not $currentVersion) {
    $currentVersion = "0.0.0"
}
$manifestRef = [string](Get-JsonValue $config 'manifest_url' '')
if (-not $manifestRef) {
    throw "Configura manifest_url em update_config.json."
}

$workDir = Join-Path $env:TEMP ("lugest_update_" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workDir -Force | Out-Null

try {
    $manifestResult = Read-Manifest $manifestRef $config $workDir
    $manifest = $manifestResult.Manifest
    $manifestLocalPath = $manifestResult.LocalPath
    $manifestSourceRef = [string]$manifestResult.SourceRef
    $latestVersion = [string](Get-JsonValue $manifest 'version' '')
    if (-not $latestVersion) {
        throw "Manifest sem campo version."
    }
    $cmp = Compare-Version $currentVersion $latestVersion
    Write-Info "Versao instalada: $currentVersion"
    Write-Info "Versao disponivel: $latestVersion"
    if ($cmp -ge 0) {
        Write-Info "Nao existem atualizacoes novas."
        Save-JsonFile (Join-Path $appDir 'update_last_result.json') @{
            checked_at = (Get-Date).ToString('s')
            current_version = $currentVersion
            latest_version = $latestVersion
            update_available = $false
            status = 'up_to_date'
        }
        exit 0
    }
    if ($CheckOnly) {
        Write-Info "Existe atualizacao disponivel."
        exit 0
    }
    Write-Host ""
    Write-Host "Vai ser instalada a versao $latestVersion."
    Write-Host "Notas: $([string](Get-JsonValue $manifest 'notes' ''))"
    if (-not $Force) {
        $answer = Read-Host "Continuar? (S/N)"
        if ($answer.Trim().ToUpperInvariant() -ne 'S') {
            Write-Info "Atualizacao cancelada."
            exit 2
        }
    }

    $manifestBase = Split-Path $manifestLocalPath -Parent
    if ($manifestSourceRef -match '^https?://') {
        $manifestBase = $manifestSourceRef
    }
    $packageRef = [string](Get-JsonValue $manifest 'package_url' '')
    if (-not $packageRef) {
        throw "Manifest sem package_url."
    }
    $packageResolved = Resolve-RelativePathOrUrl $packageRef $manifestBase
    $packagePath = Join-Path $workDir 'package.zip'
    if ($packageResolved -match '^https?://') {
        $token = [string](Get-JsonValue $config 'github_token' '')
        Download-RemoteFile $packageResolved $packagePath $token
    }
    else {
        if (-not (Test-Path $packageResolved)) {
            throw "Pacote nao encontrado: $packageResolved"
        }
        Copy-Item $packageResolved $packagePath -Force
    }

    $expectedHash = [string](Get-JsonValue $manifest 'sha256' '')
    if ($expectedHash) {
        $actualHash = (Get-FileHash $packagePath -Algorithm SHA256).Hash
        if ($actualHash.ToLowerInvariant() -ne $expectedHash.ToLowerInvariant()) {
            throw "Checksum invalido. Esperado $expectedHash, obtido $actualHash."
        }
    }

    $backupRoot = Join-Path (Split-Path $appDir -Parent) 'Backups LuisGEST'
    New-Item -ItemType Directory -Path $backupRoot -Force | Out-Null
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $backupZip = Join-Path $backupRoot "DesktopApp_$currentVersion`_$stamp.zip"
    Write-Info "A criar backup da aplicacao em $backupZip"
    Compress-Archive -Path (Join-Path $appDir '*') -DestinationPath $backupZip -Force
    Backup-DatabaseBestEffort $appDir $backupRoot

    $extractDir = Join-Path $workDir 'extract'
    Expand-Archive -Path $packagePath -DestinationPath $extractDir -Force
    $sourceDir = $extractDir
    $desktopCandidate = Join-Path $extractDir 'Desktop App'
    if (Test-Path $desktopCandidate) {
        $sourceDir = $desktopCandidate
    }

    Assert-AppClosed $appDir
    Write-Info "A instalar ficheiros..."
    robocopy $sourceDir $appDir /E /XD "Backups LuisGEST" /XF "lugest.env" "lugest_trial.json" "update_config.json" "update_last_result.json" | Out-Null
    if ($LASTEXITCODE -gt 7) {
        throw "Falha a copiar ficheiros da atualizacao (robocopy exit $LASTEXITCODE)."
    }

    if ($null -eq $config.PSObject.Properties['current_version']) {
        $config | Add-Member -NotePropertyName 'current_version' -NotePropertyValue $latestVersion
    }
    else {
        $config.current_version = $latestVersion
    }
    Save-JsonFile $ConfigPath $config
    if (Test-Path $versionPath) {
        Set-Content -Path $versionPath -Value $latestVersion -Encoding UTF8
    }
    Save-JsonFile (Join-Path $appDir 'update_last_result.json') @{
        installed_at = (Get-Date).ToString('s')
        previous_version = $currentVersion
        installed_version = $latestVersion
        backup = $backupZip
        status = 'installed'
    }
    Write-Info "Atualizacao concluida."
    $exe = Join-Path $appDir 'lugest_qt.exe'
    if (-not (Test-Path $exe)) {
        $exe = Join-Path $appDir 'main.exe'
    }
    if ((-not $NoRestart) -and (Test-Path $exe)) {
        Start-Process -FilePath $exe -WorkingDirectory $appDir
    }
}
finally {
    try {
        if (Test-Path $workDir) {
            Remove-Item $workDir -Recurse -Force
        }
    }
    catch {
    }
}
