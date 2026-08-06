param(
    [string]$ProjectRoot = (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path))
)

$ErrorActionPreference = 'Stop'
$releaseDir = Join-Path $ProjectRoot 'WordOllama\bin\Release'
$legacyIcons = Get-ChildItem -LiteralPath (Join-Path $ProjectRoot 'WordOllama\Resources') -Filter 'icons8-*' -ErrorAction SilentlyContinue
if ($legacyIcons) {
    throw 'Legacy Icons8 assets remain in WordOllama/Resources.'
}

& (Join-Path $PSScriptRoot 'verify-localization.ps1')
if (-not $?) { throw 'Localization verification failed.' }

& (Join-Path $PSScriptRoot 'run-core-smoke-test.ps1')
if (-not $?) { throw 'Core smoke test failed.' }

& (Join-Path $PSScriptRoot 'run-ui-thread-smoke-test.ps1')
if (-not $?) { throw 'Office UI thread smoke test failed.' }

$required = @(
    'WordOllama.dll',
    'WordOllama.vsto',
    'WordOllama.dll.manifest',
    'models.json',
    'en\WordOllama.resources.dll'
)

foreach ($name in $required) {
    $path = Join-Path $releaseDir $name
    if (-not (Test-Path -LiteralPath $path)) { throw "Release artifact is missing: $path" }
}

$assemblyPath = Join-Path $releaseDir 'WordOllama.dll'
$deploymentManifestPath = Join-Path $releaseDir 'WordOllama.vsto'
$applicationManifestPath = Join-Path $releaseDir 'WordOllama.dll.manifest'
$assemblyVersion = [Reflection.AssemblyName]::GetAssemblyName($assemblyPath).Version.ToString()
$fileVersion = (Get-Item -LiteralPath $assemblyPath).VersionInfo.FileVersion
$sourceVersionLine = Get-Content -LiteralPath (Join-Path $ProjectRoot 'WordOllama\Properties\AssemblyInfo.cs') |
    Where-Object { $_ -notmatch '^\s*//' -and $_ -match '^\s*\[assembly:\s*AssemblyVersion\("[^"]+"\)\]' } |
    Select-Object -First 1
$sourceVersion = ([regex]::Match($sourceVersionLine, 'AssemblyVersion\("([^"]+)"\)')).Groups[1].Value
if ([string]::IsNullOrWhiteSpace($sourceVersion)) { throw 'AssemblyVersion was not found in Properties\AssemblyInfo.cs.' }
if ($assemblyVersion -ne $sourceVersion -or $fileVersion -ne $sourceVersion) { throw "Release DLL version ($assemblyVersion / $fileVersion) does not match source version ($sourceVersion)." }

$sourceFiles = Get-ChildItem -Path (Join-Path $ProjectRoot 'WordOllama') -Recurse -File -Include *.cs,*.xaml,*.resx,*.config |
    Where-Object { $_.FullName -notmatch '\\(bin|obj)\\' }
$newestSource = $sourceFiles | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
if ($newestSource.LastWriteTimeUtc -gt (Get-Item -LiteralPath $assemblyPath).LastWriteTimeUtc) {
    throw "Release DLL is older than source file: $($newestSource.FullName). Build Release again before packaging."
}

[xml]$deploymentManifest = Get-Content -LiteralPath $deploymentManifestPath -Raw
[xml]$applicationManifest = Get-Content -LiteralPath $applicationManifestPath -Raw
$deploymentXml = $deploymentManifest.OuterXml
$applicationXml = $applicationManifest.OuterXml
if ($deploymentXml -notmatch 'codebase="WordOllama\.dll\.manifest"') { throw 'Deployment manifest does not reference WordOllama.dll.manifest.' }
if ($deploymentXml -notmatch ('name="WordOllama\.dll" version="' + [regex]::Escape($assemblyVersion) + '"')) { throw 'Deployment manifest DLL version does not match the Release DLL.' }
if ($applicationXml -notmatch ('name="WordOllama\.dll" version="' + [regex]::Escape($assemblyVersion) + '"')) { throw 'Application manifest DLL version does not match the Release DLL.' }
Write-Host "Release artifacts: OK"
Write-Host "Assembly version: $assemblyVersion"
Write-Host "File version: $fileVersion"
Write-Host "Newest source: $($newestSource.FullName)"
Write-Host "Source folder for installer sync: $releaseDir"
