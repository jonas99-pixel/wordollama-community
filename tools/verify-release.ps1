param(
    [string]$ProjectRoot = (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)),
    [switch]$SkipBuild
)

$ErrorActionPreference = 'Stop'
$releaseDir = Join-Path $ProjectRoot 'WordOllama\bin\Release'
$legacyIcons = Get-ChildItem -LiteralPath (Join-Path $ProjectRoot 'WordOllama\Resources') -Filter 'icons8-*' -ErrorAction SilentlyContinue
if ($legacyIcons) {
    throw 'Legacy Icons8 assets remain in WordOllama/Resources.'
}

if (-not $SkipBuild) {
    & (Join-Path $ProjectRoot 'tools\build-open-source.ps1') -Configuration Release
    if (-not $?) { throw 'Community Release build failed.' }
}

& (Join-Path $PSScriptRoot 'verify-localization.ps1')
if (-not $?) { throw 'Localization verification failed.' }

& (Join-Path $PSScriptRoot 'run-core-smoke-test.ps1')
if (-not $?) { throw 'Core smoke test failed.' }

& (Join-Path $PSScriptRoot 'run-ui-thread-smoke-test.ps1')
if (-not $?) { throw 'Office UI thread smoke test failed.' }

$required = @(
    'WordOllama.dll',
    'WordOllama.dll.config',
    'en\WordOllama.resources.dll'
)

foreach ($name in $required) {
    $path = Join-Path $releaseDir $name
    if (-not (Test-Path -LiteralPath $path)) { throw "Release artifact is missing: $path" }
}

$assemblyPath = Join-Path $releaseDir 'WordOllama.dll'
$assemblyVersion = [Reflection.AssemblyName]::GetAssemblyName($assemblyPath).Version.ToString()
$fileVersion = (Get-Item -LiteralPath $assemblyPath).VersionInfo.FileVersion
$sourceVersionLine = Get-Content -LiteralPath (Join-Path $ProjectRoot 'WordOllama\Properties\AssemblyInfo.cs') |
    Where-Object { $_ -notmatch '^\s*//' -and $_ -match '^\s*\[assembly:\s*AssemblyVersion\("[^"]+"\)\]' } |
    Select-Object -First 1
$sourceVersion = ([regex]::Match($sourceVersionLine, 'AssemblyVersion\("([^"]+)"\)')).Groups[1].Value
if ([string]::IsNullOrWhiteSpace($sourceVersion)) { throw 'AssemblyVersion was not found in Properties\AssemblyInfo.cs.' }
if ($assemblyVersion -ne $sourceVersion -or $fileVersion -ne $sourceVersion) { throw "Release DLL version ($assemblyVersion / $fileVersion) does not match source version ($sourceVersion)." }

[xml]$projectXml = [IO.File]::ReadAllText((Join-Path $ProjectRoot 'WordOllama\WordOllama.csproj'), [Text.Encoding]::UTF8)
$applicationVersion = $projectXml.SelectSingleNode("//*[local-name()='ApplicationVersion']").InnerText
if ($applicationVersion -ne $sourceVersion) { throw "Project version ($applicationVersion) does not match source version ($sourceVersion)." }

$sourceFiles = Get-ChildItem -Path (Join-Path $ProjectRoot 'WordOllama') -Recurse -File -Include *.cs,*.xaml,*.resx,*.config,*.txt,*.csproj |
    Where-Object { $_.FullName -notmatch '\\(bin|obj)\\' }
$newestSource = $sourceFiles | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
if ($newestSource.LastWriteTimeUtc -gt (Get-Item -LiteralPath $assemblyPath).LastWriteTimeUtc) {
    throw "Release DLL is older than source file: $($newestSource.FullName). Build Release again before packaging."
}

$assembly = [Reflection.Assembly]::LoadFrom($assemblyPath)
$releaseVersion = ([Version]$sourceVersion).ToString(3)
foreach ($logName in @('updateLog.txt', 'updateLog.en.txt')) {
    $logPath = Join-Path $ProjectRoot ("WordOllama\Resources\" + $logName)
    $sourceBytes = [IO.File]::ReadAllBytes($logPath)
    $sourceText = [Text.Encoding]::UTF8.GetString($sourceBytes).TrimStart([char]0xFEFF)
    if (($sourceText -split '\r?\n')[0] -notmatch ('\bV' + [regex]::Escape($releaseVersion) + '\s')) {
        throw "The original $logName does not start with version $releaseVersion."
    }
    $stream = $assembly.GetManifestResourceStream("WordOllama.Resources.$logName")
    if ($null -eq $stream) { throw "Original update log is not embedded: $logName" }
    $buffer = New-Object IO.MemoryStream
    try {
        $stream.CopyTo($buffer)
        if ([Convert]::ToBase64String($sourceBytes) -cne [Convert]::ToBase64String($buffer.ToArray())) {
            throw "Embedded $logName differs from the source TXT. Rebuild Release."
        }
    }
    finally {
        $stream.Dispose()
        $buffer.Dispose()
    }
}
Write-Host 'Original Chinese/English TXT resources: OK'
Write-Host 'Community Release assemblies: OK (ClickOnce manifests are excluded from source builds)'
Write-Host "Assembly version: $assemblyVersion"
Write-Host "File version: $fileVersion"
Write-Host "Newest source: $($newestSource.FullName)"
Write-Host "Community build output: $releaseDir"
