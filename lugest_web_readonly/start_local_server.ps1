param(
    [int]$Port = 8088,
    [switch]$NoBrowser
)

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$publicDir = Join-Path $projectRoot 'public'
$phpIni = Join-Path $projectRoot 'php.ini'
$pidFile = Join-Path $projectRoot 'local-server.pid'
$stdoutLog = Join-Path $projectRoot 'local-server.out.log'
$stderrLog = Join-Path $projectRoot 'local-server.err.log'

function Find-PhpExe {
    $command = Get-Command php -ErrorAction SilentlyContinue
    if ($command -and $command.Source) {
        return $command.Source
    }

    $candidates = @(
        (Join-Path $env:LOCALAPPDATA 'Microsoft\WinGet\Packages\PHP.PHP.8.3_Microsoft.Winget.Source_8wekyb3d8bbwe\php.exe'),
        (Join-Path $env:LOCALAPPDATA 'Microsoft\WinGet\Packages\PHP.PHP.8.4_Microsoft.Winget.Source_8wekyb3d8bbwe\php.exe')
    )

    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) {
            return $candidate
        }
    }

    $found = Get-ChildItem (Join-Path $env:LOCALAPPDATA 'Microsoft\WinGet\Packages') -Recurse -Filter php.exe -ErrorAction SilentlyContinue |
        Select-Object -First 1 -ExpandProperty FullName
    if ($found) {
        return $found
    }

    throw 'Nao encontrei php.exe. Instala PHP 8.3 com winget e volta a correr este script.'
}

function Stop-ListeningProcess([int]$ListenPort) {
    $connections = Get-NetTCPConnection -LocalPort $ListenPort -State Listen -ErrorAction SilentlyContinue
    if (-not $connections) {
        return
    }

    $pids = $connections | Select-Object -ExpandProperty OwningProcess -Unique
    foreach ($pid in $pids) {
        Stop-Process -Id $pid -Force -ErrorAction SilentlyContinue
    }
}

if (-not (Test-Path $publicDir)) {
    throw "Nao encontrei a pasta public em $publicDir"
}

if (-not (Test-Path $phpIni)) {
    throw "Nao encontrei o ficheiro php.ini em $phpIni"
}

if (Test-Path $pidFile) {
    $existingPid = Get-Content $pidFile -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($existingPid -and (Get-Process -Id $existingPid -ErrorAction SilentlyContinue)) {
        Stop-Process -Id $existingPid -Force -ErrorAction SilentlyContinue
    }
    Remove-Item $pidFile -Force -ErrorAction SilentlyContinue
}

Stop-ListeningProcess -ListenPort $Port

$phpExe = Find-PhpExe
$arguments = @(
    '-c', $phpIni,
    '-S', "127.0.0.1:$Port",
    '-t', $publicDir
)

$process = Start-Process -FilePath $phpExe `
    -ArgumentList $arguments `
    -WorkingDirectory $projectRoot `
    -RedirectStandardOutput $stdoutLog `
    -RedirectStandardError $stderrLog `
    -PassThru

Start-Sleep -Seconds 2

$listening = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue |
    Where-Object { $_.OwningProcess -eq $process.Id }

if (-not $listening) {
    throw "O servidor nao arrancou corretamente. Verifica $stderrLog"
}

Set-Content -Path $pidFile -Value $process.Id -Encoding ASCII

$url = "http://127.0.0.1:$Port/login.php"
Write-Output "Portal luGEST de consulta a correr em $url"
Write-Output "PID: $($process.Id)"
Write-Output "Logs: $stdoutLog e $stderrLog"

if (-not $NoBrowser) {
    Start-Process $url
}
