param(
    [string]$ProjectRoot = (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path))
)

$ErrorActionPreference = 'Stop'
$releaseDir = Join-Path $ProjectRoot 'WordOllama\bin\Release'
$mainDll = Join-Path $releaseDir 'WordOllama.dll'
$jsonDll = Join-Path $releaseDir 'Newtonsoft.Json.dll'
$source = Join-Path $ProjectRoot 'tools\core-smoke-test\Program.cs'
$workDir = Join-Path ([System.IO.Path]::GetTempPath()) 'WordOllamaCoreSmokeBuild'
$exe = Join-Path $workDir 'CoreSmokeTest.exe'
$csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe'

if (-not (Test-Path -LiteralPath $mainDll)) { throw "Release assembly not found: $mainDll" }
if (-not (Test-Path -LiteralPath $csc)) { throw "C# compiler not found: $csc" }

New-Item -ItemType Directory -Path $workDir -Force | Out-Null
Copy-Item -LiteralPath $mainDll -Destination $workDir -Force
Copy-Item -LiteralPath $jsonDll -Destination $workDir -Force
Copy-Item -LiteralPath (Join-Path $releaseDir 'WordOllama.dll.config') -Destination (Join-Path $workDir 'CoreSmokeTest.exe.config') -Force
Copy-Item -Path (Join-Path $releaseDir '*.dll') -Destination $workDir -Force
Copy-Item -LiteralPath (Join-Path $releaseDir 'en') -Destination $workDir -Recurse -Force
& $csc /nologo /target:exe /platform:anycpu /out:$exe /reference:$mainDll /reference:$jsonDll $source
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Push-Location $workDir
try {
    & $exe
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}
finally {
    Pop-Location
}
