[CmdletBinding()]
param(
    [string]$SourceRoot = "",
    [switch]$Global
)

$ErrorActionPreference = "Stop"
$Root = if ($SourceRoot) { (Resolve-Path -LiteralPath $SourceRoot).Path } else { Split-Path -Parent $PSScriptRoot }
$source = Join-Path $Root "skills\office\SKILL.md"
if (-not (Test-Path -LiteralPath $source)) {
    throw "Skill source not found: $source"
}

if ($Global) {
    $destination = Join-Path $env:USERPROFILE ".config\opencode\skills\office"
}
else {
    $destination = Join-Path $Root ".opencode\skills\office"
}

New-Item -ItemType Directory -Force -Path $destination | Out-Null
Copy-Item -LiteralPath $source -Destination (Join-Path $destination "SKILL.md") -Force
Write-Output (Join-Path $destination "SKILL.md")
