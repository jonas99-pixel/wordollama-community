param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release',
    [switch]$SkipBuild
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot

if (-not $SkipBuild) {
    & (Join-Path $PSScriptRoot 'build-open-source.ps1') -Configuration $Configuration
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

& (Join-Path $PSScriptRoot 'verify-localization.ps1')
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

if ($Configuration -eq 'Release') {
    & (Join-Path $PSScriptRoot 'run-core-smoke-test.ps1')
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

$legacyIcons = Get-ChildItem -LiteralPath (Join-Path $root 'WordOllama\Resources') -Filter 'icons8-*' -ErrorAction SilentlyContinue
if ($legacyIcons) {
    throw 'Legacy Icons8 assets remain in WordOllama/Resources.'
}

Write-Host "Release verification passed for $Configuration." -ForegroundColor Green
