[CmdletBinding()]
param(
    [string]$HostName = "127.0.0.1",
    [int]$Port = 0,
    [switch]$Background,
    [switch]$NoHotkey,
    [string]$LogLevel = "INFO"
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
Push-Location $Root
try {
    $activate = Join-Path $Root ".venv\Scripts\Activate.ps1"
    if (Test-Path -LiteralPath $activate) {
        . $activate
    }

    $serverArgs = @("-m", "office_bridge.server", "--host", $HostName, "--log-level", $LogLevel)
    if ($Port -gt 0) {
        $serverArgs += @("--port", $Port)
    }
    if ($NoHotkey) {
        $serverArgs += "--no-hotkey"
    }

    if ($Background) {
        $logDir = Join-Path $Root ".office-bridge"
        New-Item -ItemType Directory -Force -Path $logDir | Out-Null
        $stdout = Join-Path $logDir "bridge.stdout.log"
        $stderr = Join-Path $logDir "bridge.stderr.log"
        $process = Start-Process -FilePath "python" `
            -ArgumentList $serverArgs `
            -WorkingDirectory $Root `
            -PassThru `
            -WindowStyle Hidden `
            -RedirectStandardOutput $stdout `
            -RedirectStandardError $stderr
        Write-Output $process.Id
        exit 0
    }

    & python @serverArgs
    exit $LASTEXITCODE
}
finally {
    Pop-Location
}
