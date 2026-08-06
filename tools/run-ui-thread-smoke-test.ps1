param(
    [string]$ProjectRoot = (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path))
)

$ErrorActionPreference = 'Stop'
$releaseDir = Join-Path $ProjectRoot 'WordOllama\bin\Release'
$mainDll = Join-Path $releaseDir 'WordOllama.dll'
$source = Join-Path $ProjectRoot 'tools\ui-thread-smoke-test\Program.cs'
$workDir = Join-Path ([System.IO.Path]::GetTempPath()) 'WordOllamaUiThreadSmokeBuild'
$exe = Join-Path $workDir 'UiThreadSmokeTest.exe'
$csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe'

if (-not (Test-Path -LiteralPath $mainDll)) { throw "Release assembly not found: $mainDll" }
if (-not (Test-Path -LiteralPath $csc)) { throw "C# compiler not found: $csc" }

New-Item -ItemType Directory -Path $workDir -Force | Out-Null
Copy-Item -Path (Join-Path $releaseDir '*.dll') -Destination $workDir -Force
& $csc /nologo /target:exe /platform:anycpu /out:$exe /reference:$mainDll /reference:System.Windows.Forms.dll $source
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Push-Location $workDir
try
{
    & $exe
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}
finally { Pop-Location }
