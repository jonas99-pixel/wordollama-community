param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release'
)

$ErrorActionPreference = 'Stop'
if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
    throw 'The VSTO build requires Windows, Visual Studio, and the Office developer tools.'
}
$root = Split-Path -Parent $PSScriptRoot
$solution = Join-Path $root 'WordOllama.sln'
$msbuild = Get-Command MSBuild.exe -ErrorAction SilentlyContinue

if (-not $msbuild) {
    $programFilesX86 = [Environment]::GetEnvironmentVariable('ProgramFiles(x86)')
    $vswhereCandidates = @(
        (Join-Path $programFilesX86 'Microsoft Visual Studio\Installer\vswhere.exe'),
        (Join-Path ${env:ProgramFiles} 'Microsoft Visual Studio\Installer\vswhere.exe')
    ) | Where-Object { Test-Path -LiteralPath $_ }
    $vswhereCandidates = @($vswhereCandidates)

    if ($vswhereCandidates.Count -gt 0) {
        $installationPath = & ($vswhereCandidates[0]) -latest -products * -requires Microsoft.Component.MSBuild -property installationPath 2>$null
        if ($installationPath) {
            $candidate = Join-Path $installationPath 'MSBuild\Current\Bin\MSBuild.exe'
            if (Test-Path -LiteralPath $candidate) {
                $msbuild = Get-Item -LiteralPath $candidate
            }
        }
    }
}

if (-not $msbuild) {
    throw 'MSBuild.exe was not found. Install Visual Studio with the .NET desktop workload.'
}

$msbuildPath = $msbuild.Source
if (-not $msbuildPath) { $msbuildPath = $msbuild.Path }
if (-not $msbuildPath) { $msbuildPath = $msbuild.FullName }
& $msbuildPath $solution /t:Restore,Build /p:Configuration=$Configuration /p:OpenSourceBuild=true
if ($LASTEXITCODE -ne 0) {
    throw "WordOllama open-source $Configuration build failed with exit code $LASTEXITCODE."
}

Write-Host "WordOllama open-source $Configuration build completed without ClickOnce or assembly signing."
