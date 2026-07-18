param(
    [string]$ProjectRoot = (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path))
)

$ErrorActionPreference = 'Stop'
$releaseDir = Join-Path $ProjectRoot 'WordOllama\bin\Release'
$sourceDir = Join-Path $ProjectRoot 'tools\mcp-smoke-test'
$workDir = Join-Path ([System.IO.Path]::GetTempPath()) 'WordOllamaMcpSmokeBuild'
$mainDll = Join-Path $releaseDir 'WordOllama.dll'
$jsonDll = Join-Path $releaseDir 'Newtonsoft.Json.dll'
$programFilesX86 = [Environment]::GetEnvironmentVariable('ProgramFiles(x86)')
$piaDir = Join-Path $programFilesX86 'Microsoft Visual Studio\Shared\Visual Studio Tools for Office\PIA\Office15'
$wordPia = Join-Path $piaDir 'Microsoft.Office.Interop.Word.dll'
$officePia = Join-Path $piaDir 'Office.dll'

$vswhereCandidates = @(
    (Join-Path $programFilesX86 'Microsoft Visual Studio\Installer\vswhere.exe'),
    (Join-Path ${env:ProgramFiles} 'Microsoft Visual Studio\Installer\vswhere.exe')
) | Where-Object { Test-Path -LiteralPath $_ }
$vswhereCandidates = @($vswhereCandidates)
$csc = $null
if ($vswhereCandidates.Count -gt 0) {
    $installationPath = & ($vswhereCandidates[0]) -latest -products * -requires Microsoft.Component.MSBuild -property installationPath 2>$null
    if ($installationPath) {
        $candidate = Join-Path $installationPath 'MSBuild\Current\Bin\Roslyn\csc.exe'
        if (Test-Path -LiteralPath $candidate) { $csc = $candidate }
    }
}

if (-not (Test-Path -LiteralPath $mainDll)) { throw "Release assembly not found: $mainDll" }
if (-not $csc) { throw 'Visual Studio Roslyn C# compiler was not found.' }
if (-not (Test-Path -LiteralPath $wordPia) -or -not (Test-Path -LiteralPath $officePia)) {
    throw 'Office 15 PIA reference assemblies were not found. Install the Visual Studio Office developer tools.'
}
if (-not (Get-Command node -ErrorAction SilentlyContinue)) { throw 'Node.js is required for the MCP smoke servers.' }

New-Item -ItemType Directory -Path $workDir -Force | Out-Null
Copy-Item -Path (Join-Path $releaseDir '*.dll') -Destination $workDir -Force
Copy-Item -LiteralPath $wordPia,$officePia -Destination $workDir -Force
if (Test-Path -LiteralPath (Join-Path $releaseDir 'en')) {
    Copy-Item -LiteralPath (Join-Path $releaseDir 'en') -Destination $workDir -Recurse -Force
}

function Build-SmokeTest([string]$name, [string[]]$references) {
    $source = Join-Path $sourceDir ($name + '.cs')
    $output = Join-Path $workDir ($name + '.exe')
    $arguments = @('/nologo', '/target:exe', '/platform:anycpu', "/out:$output", "/reference:$mainDll", "/reference:$jsonDll")
    foreach ($reference in $references) { $arguments += "/reference:$reference" }
    $arguments += $source
    & $csc @arguments
    if ($LASTEXITCODE -ne 0) { throw "$name compilation failed." }
    return $output
}

function Invoke-SmokeTest([string]$executable, [string[]]$arguments) {
    & $executable @arguments
    if ($LASTEXITCODE -ne 0) { throw "$(Split-Path -Leaf $executable) failed with exit code $LASTEXITCODE." }
}

$bridgeTest = Build-SmokeTest 'McpBridgeSmokeTest' @()
$managerTest = Build-SmokeTest 'McpManagerSmokeTest' @($wordPia, $officePia)
$httpTest = Build-SmokeTest 'McpHttpSmokeTest' @('System.Net.Http.dll')

Push-Location $workDir
try {
    Invoke-SmokeTest $bridgeTest @((Join-Path $sourceDir 'strict-server.js'))
    Invoke-SmokeTest $managerTest @((Join-Path $sourceDir 'strict-server.cmd'))
    Invoke-SmokeTest $httpTest @((Join-Path $sourceDir 'http-server.js'))
}
finally {
    Pop-Location
}

Write-Host 'All MCP smoke tests passed.'
