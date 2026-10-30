[CmdletBinding()]
param(
    [string]$SourceRoot = "",
    [switch]$Global,
    [string]$Agent = "",
    [string]$Scope = ""
)

$ErrorActionPreference = "Stop"
$Root = if ($SourceRoot) { (Resolve-Path -LiteralPath $SourceRoot).Path } else { Split-Path -Parent $PSScriptRoot }

$python = Join-Path $Root ".venv\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $python)) {
    $python = "python"
}

$arguments = @("-m", "office_bridge.agent_config", "install-skill")
if ($Agent) {
    $arguments += @("--agent", $Agent)
}
if ($Scope) {
    $arguments += @("--scope", $Scope)
}
elseif ($Global) {
    $arguments += @("--scope", "user")
}

$source = Join-Path $Root "skills\office\SKILL.md"
if (Test-Path -LiteralPath $source) {
    $arguments += @("--source", $source)
}

& $python @arguments
exit $LASTEXITCODE
