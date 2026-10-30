[CmdletBinding()]
param(
    [int]$Port = 0,
    [switch]$NoHotkey,
    [switch]$KeepBridge,
    [switch]$NoSkillInstall,
    [string]$Agent = "",
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$AgentArgs = @()
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
Push-Location $Root
$bridgeProcess = $null
$startedBridge = $false

function Test-Bridge {
    param([int]$TimeoutSeconds = 3)
    try {
        $response = Invoke-RestMethod `
            -Uri "$env:OFFICE_BRIDGE_URL/health" `
            -TimeoutSec $TimeoutSeconds `
            -ErrorAction Stop
        return [bool]$response.ok
    }
    catch {
        return $false
    }
}

function Show-BridgeLogs {
    $logDir = Join-Path $Root ".office-bridge"
    foreach ($name in @("bridge.stderr.log", "bridge.stdout.log")) {
        $path = Join-Path $logDir $name
        if (Test-Path -LiteralPath $path) {
            Write-Host "--- $name ---" -ForegroundColor Yellow
            Get-Content -LiteralPath $path -Tail 40
        }
    }
}

function Invoke-AgentConfig {
    param([string[]]$Arguments)
    $output = & $python -m office_bridge.agent_config @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "office-agent failed with exit code $LASTEXITCODE."
    }
    return $output
}

try {
    $activate = Join-Path $Root ".venv\Scripts\Activate.ps1"
    if (Test-Path -LiteralPath $activate) {
        . $activate
    }

    $python = Join-Path $Root ".venv\Scripts\python.exe"
    if (-not (Test-Path -LiteralPath $python)) {
        $python = "python"
    }

    if ($Agent) {
        $env:OFFICE_BRIDGE_AGENT = $Agent
    }

    $effectivePort = if ($Port -gt 0) {
        $Port
    }
    elseif ($env:OFFICE_BRIDGE_PORT) {
        [int]$env:OFFICE_BRIDGE_PORT
    }
    else {
        8733
    }
    $env:OFFICE_BRIDGE_URL = "http://127.0.0.1:$effectivePort"
    try {
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    }
    catch {
    }

    if (-not $NoSkillInstall) {
        $installed = Invoke-AgentConfig @("install-skill")
        Write-Host "Skill: $installed"
    }

    $specJson = (Invoke-AgentConfig @("resolve", "--json")) -join "`n"
    $spec = $specJson | ConvertFrom-Json
    if (-not $spec.ok) {
        throw $spec.error
    }
    if ($spec.skill_dir -and -not $spec.skill_installed -and -not $NoSkillInstall) {
        throw "Office skill is not installed at $($spec.skill_dir)."
    }

    if (Test-Bridge) {
        Write-Host "Bridge already running at $env:OFFICE_BRIDGE_URL"
    }
    else {
        $logDir = Join-Path $Root ".office-bridge"
        New-Item -ItemType Directory -Force -Path $logDir | Out-Null
        $stdout = Join-Path $logDir "bridge.stdout.log"
        $stderr = Join-Path $logDir "bridge.stderr.log"
        $serverArgs = @(
            "-m", "office_bridge.server",
            "--host", "127.0.0.1",
            "--port", "$effectivePort"
        )
        if ($NoHotkey) {
            $serverArgs += "--no-hotkey"
        }

        Write-Host "Starting bridge on port $effectivePort ..."
        $bridgeProcess = Start-Process -FilePath $python `
            -ArgumentList $serverArgs `
            -WorkingDirectory $Root `
            -PassThru `
            -WindowStyle Hidden `
            -RedirectStandardOutput $stdout `
            -RedirectStandardError $stderr
        $startedBridge = $true

        $deadline = (Get-Date).AddSeconds(20)
        $healthy = $false
        while ((Get-Date) -lt $deadline) {
            if ($bridgeProcess.HasExited) {
                break
            }
            if (Test-Bridge) {
                $healthy = $true
                break
            }
            Start-Sleep -Milliseconds 400
        }
        if (-not $healthy) {
            Write-Host "The bridge did not become healthy within 20 seconds." -ForegroundColor Red
            Show-BridgeLogs
            if ($bridgeProcess -and -not $bridgeProcess.HasExited) {
                Stop-Process -Id $bridgeProcess.Id -Force
            }
            exit 1
        }
        Write-Host "Bridge is healthy."
    }

    $launch = @($spec.launch_argv)
    if ($launch.Count -eq 0) {
        throw "No launch command is configured."
    }
    $launchExecutable = [string]$launch[0]
    $launchArguments = if ($launch.Count -gt 1) {
        $launch[1..($launch.Count - 1)]
    }
    else {
        @()
    }

    Write-Host "Starting $($spec.display_name) with the Office skill available."
    & $launchExecutable @launchArguments @AgentArgs
    exit $LASTEXITCODE
}
catch {
    Write-Host "start-agent.ps1: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
finally {
    if ($startedBridge -and -not $KeepBridge -and $bridgeProcess) {
        if (-not $bridgeProcess.HasExited) {
            Write-Host "Stopping the bridge started by this launcher."
            Stop-Process -Id $bridgeProcess.Id -Force
        }
    }
    elseif ($startedBridge -and $KeepBridge) {
        Write-Host "Bridge left running (pid $($bridgeProcess.Id)) because -KeepBridge was specified."
    }
    Pop-Location
}
