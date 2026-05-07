param(
    [string]$InstallDir = "",
    [switch]$Force,
    [switch]$ReplaceEnv
)

$ErrorActionPreference = 'Stop'

function Test-IsAdmin {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function New-Shortcut($shortcutPath, $targetPath, $workingDirectory, $iconPath) {
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $targetPath
    $shortcut.WorkingDirectory = $workingDirectory
    if (Test-Path $iconPath) {
        $shortcut.IconLocation = $iconPath
    }
    $shortcut.Description = 'LuisGEST ERP industrial'
    $shortcut.Save()
}

function New-ShortcutWithArgs($shortcutPath, $targetPath, $arguments, $workingDirectory, $iconPath) {
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $targetPath
    $shortcut.Arguments = $arguments
    $shortcut.WorkingDirectory = $workingDirectory
    if (Test-Path $iconPath) {
        $shortcut.IconLocation = $iconPath
    }
    $shortcut.Description = 'LuisGEST ERP industrial'
    $shortcut.Save()
}

$sourceDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $InstallDir) {
    if (Test-IsAdmin) {
        $InstallDir = Join-Path $env:ProgramFiles 'LuisGEST'
    }
    else {
        $InstallDir = Join-Path $env:LOCALAPPDATA 'LuisGEST'
    }
}

$exeName = 'main.exe'
$sourceExe = Join-Path $sourceDir $exeName
if (-not (Test-Path $sourceExe)) {
    $exeName = 'lugest_qt.exe'
    $sourceExe = Join-Path $sourceDir $exeName
}
if (-not (Test-Path $sourceExe)) {
    throw 'Nao foi encontrado lugest_qt.exe nem main.exe na pasta de origem.'
}

if (Test-Path $InstallDir) {
    if (-not $Force) {
        $answer = Read-Host "A pasta $InstallDir ja existe. Atualizar ficheiros mantendo lugest.env? (S/N)"
        if ($answer.Trim().ToUpperInvariant() -ne 'S') {
            Write-Host 'Instalacao cancelada.'
            exit 2
        }
    }
    $backupRoot = Join-Path (Split-Path $InstallDir -Parent) 'LuisGEST Backups'
    New-Item -ItemType Directory -Path $backupRoot -Force | Out-Null
    $backupZip = Join-Path $backupRoot ("DesktopApp_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + ".zip")
    Compress-Archive -Path (Join-Path $InstallDir '*') -DestinationPath $backupZip -Force
    Write-Host "Backup criado: $backupZip"
}

New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
robocopy $sourceDir $InstallDir /E /XF 'lugest.env' 'lugest_trial.json' 'update_config.json' 'update_last_result.json' | Out-Null
if ($LASTEXITCODE -gt 7) {
    throw "Falha ao copiar ficheiros (robocopy exit $LASTEXITCODE)."
}

$sourceEnv = Join-Path $sourceDir 'lugest.env'
$targetEnv = Join-Path $InstallDir 'lugest.env'
$shouldReplaceEnv = $false
if ((Test-Path $sourceEnv) -and (Test-Path $targetEnv)) {
    if ($ReplaceEnv) {
        $shouldReplaceEnv = $true
    }
    elseif (-not $Force) {
        $answerEnv = Read-Host "Substituir o lugest.env ja instalado pelo lugest.env desta pasta? (S/N)"
        $shouldReplaceEnv = $answerEnv.Trim().ToUpperInvariant() -eq 'S'
    }
}
elseif ((Test-Path $sourceEnv) -and -not (Test-Path $targetEnv)) {
    $shouldReplaceEnv = $true
}

if ($shouldReplaceEnv -and (Test-Path $sourceEnv)) {
    Copy-Item $sourceEnv $targetEnv -Force
    Write-Host "lugest.env atualizado: $targetEnv"
}

foreach ($fileName in @('lugest_trial.json', 'update_config.json')) {
    $sourceFile = Join-Path $sourceDir $fileName
    $targetFile = Join-Path $InstallDir $fileName
    if ((Test-Path $sourceFile) -and -not (Test-Path $targetFile)) {
        Copy-Item $sourceFile $targetFile -Force
    }
}

$targetExe = Join-Path $InstallDir $exeName
$targetLauncherBat = Join-Path $InstallDir 'Arrancar LuisGEST Desktop.bat'
$targetLauncherPs1 = Join-Path $InstallDir 'Arrancar LuisGEST Desktop.ps1'
$targetLauncherVbs = Join-Path $InstallDir 'Arrancar LuisGEST Desktop.vbs'
$iconPath = Join-Path $InstallDir 'app.ico'
$desktopShortcut = Join-Path ([Environment]::GetFolderPath('Desktop')) 'LuisGEST.lnk'
$startMenuDir = Join-Path ([Environment]::GetFolderPath('Programs')) 'LuisGEST'
New-Item -ItemType Directory -Path $startMenuDir -Force | Out-Null
$startMenuShortcut = Join-Path $startMenuDir 'LuisGEST.lnk'
$powershellExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$wscriptExe = Join-Path $env:SystemRoot 'System32\wscript.exe'
if ((Test-Path $targetLauncherVbs) -and (Test-Path $wscriptExe)) {
    $launcherArgs = "`"$targetLauncherVbs`" `"$InstallDir`""
    New-ShortcutWithArgs $desktopShortcut $wscriptExe $launcherArgs $InstallDir $iconPath
    New-ShortcutWithArgs $startMenuShortcut $wscriptExe $launcherArgs $InstallDir $iconPath
}
elseif (Test-Path $targetLauncherBat) {
    New-Shortcut $desktopShortcut $targetLauncherBat $InstallDir $iconPath
    New-Shortcut $startMenuShortcut $targetLauncherBat $InstallDir $iconPath
}
elseif ((Test-Path $targetLauncherPs1) -and (Test-Path $powershellExe)) {
    $launcherArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$targetLauncherPs1`" -InstallDir `"$InstallDir`""
    New-ShortcutWithArgs $desktopShortcut $powershellExe $launcherArgs $InstallDir $iconPath
    New-ShortcutWithArgs $startMenuShortcut $powershellExe $launcherArgs $InstallDir $iconPath
}
else {
    New-Shortcut $desktopShortcut $targetExe $InstallDir $iconPath
    New-Shortcut $startMenuShortcut $targetExe $InstallDir $iconPath
}

Write-Host ''
Write-Host 'LuisGEST instalado com sucesso.'
Write-Host "Pasta: $InstallDir"
Write-Host "Atalho Ambiente de Trabalho: $desktopShortcut"
Write-Host "Atalho Menu Iniciar: $startMenuShortcut"
