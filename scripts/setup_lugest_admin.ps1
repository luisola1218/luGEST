param(
    [string]$ExecutablePath = '',
    [string]$Username = '',
    [string]$Password = '',
    [string]$Role = 'Admin',
    [switch]$Reset
)

$ErrorActionPreference = 'Stop'

function Read-RequiredText {
    param(
        [string]$Prompt,
        [string]$DefaultValue = ''
    )
    while ($true) {
        $suffix = ''
        if ($DefaultValue) {
            $suffix = " [$DefaultValue]"
        }
        $value = Read-Host "$Prompt$suffix"
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            return $value.Trim()
        }
        if (-not [string]::IsNullOrWhiteSpace($DefaultValue)) {
            return $DefaultValue.Trim()
        }
    }
}

function Read-RequiredPassword {
    param(
        [string]$Prompt
    )
    while ($true) {
        $secureA = Read-Host $Prompt -AsSecureString
        $secureB = Read-Host 'Confirmar password' -AsSecureString
        $plainA = [System.Net.NetworkCredential]::new('', $secureA).Password
        $plainB = [System.Net.NetworkCredential]::new('', $secureB).Password
        if ([string]::IsNullOrWhiteSpace($plainA)) {
            Write-Host 'A password nao pode ficar vazia.' -ForegroundColor Yellow
            continue
        }
        if ($plainA -ne $plainB) {
            Write-Host 'As passwords nao coincidem. Tenta outra vez.' -ForegroundColor Yellow
            continue
        }
        return $plainA
    }
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $ExecutablePath) {
    foreach ($candidateName in @('lugest_qt.exe', 'main.exe')) {
        $candidate = Join-Path $scriptDir $candidateName
        if (Test-Path $candidate) {
            $ExecutablePath = $candidate
            break
        }
    }
}

if (-not $ExecutablePath) {
    throw 'Nao foi encontrado lugest_qt.exe nem main.exe. Indica -ExecutablePath explicitamente.'
}

if (-not (Test-Path $ExecutablePath)) {
    throw "Executavel invalido: $ExecutablePath"
}

if (-not $Username) {
    $Username = Read-RequiredText -Prompt 'Username do administrador' -DefaultValue 'admin'
}

if (-not $Password) {
    $Password = Read-RequiredPassword -Prompt 'Password do administrador'
}

$argList = @(
    '--setup-admin',
    '--admin-username', $Username,
    '--admin-password', $Password,
    '--admin-role', $Role
)
if ($Reset) {
    $argList += '--reset-admin'
}

Write-Host ''
Write-Host "A criar/repor o administrador local '$Username'..." -ForegroundColor Cyan
& $ExecutablePath @argList
$exitCode = if ($LASTEXITCODE -is [int]) { $LASTEXITCODE } else { 0 }
if ($exitCode -ne 0) {
    throw "Falha ao criar/repor o administrador (exit $exitCode)."
}

Write-Host ''
Write-Host "Administrador local pronto: $Username" -ForegroundColor Green
Write-Host 'Este login e diferente do OWNER no lugest.env.' -ForegroundColor DarkCyan
