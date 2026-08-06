param(
    [string]$ProjectRoot = (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path))
)

$ErrorActionPreference = 'Stop'
$basePath = Join-Path $ProjectRoot 'WordOllama\Properties\Resources.resx'
$englishPath = Join-Path $ProjectRoot 'WordOllama\Properties\Resources.en.resx'

function Get-ResourceKeys([string]$path) {
    # Resource files are UTF-8 without a BOM. Windows PowerShell otherwise
    # falls back to the active ANSI code page and can corrupt Chinese text.
    [xml]$xml = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
    return @($xml.root.data | ForEach-Object { $_.name })
}

$baseKeys = Get-ResourceKeys $basePath
$englishKeys = Get-ResourceKeys $englishPath
$missingEnglish = @($baseKeys | Where-Object { $_ -cnotin $englishKeys })
$missingBase = @($englishKeys | Where-Object { $_ -cnotin $baseKeys })

$xamlKeys = @(
    Get-ChildItem -LiteralPath (Join-Path $ProjectRoot 'WordOllama') -Recurse -Filter *.xaml |
        Select-String -Pattern '\{[A-Za-z][A-Za-z0-9]*:Loc\s+(I18n_[A-Za-z0-9_]+)' -AllMatches -CaseSensitive |
        ForEach-Object { $_.Matches | ForEach-Object { $_.Groups[1].Value } }
) | Sort-Object -Unique -CaseSensitive
$missingXaml = @($xamlKeys | Where-Object { $_ -cnotin $baseKeys })

$codeKeys = @(
    Get-ChildItem -LiteralPath (Join-Path $ProjectRoot 'WordOllama') -Recurse -Filter *.cs |
        Where-Object { $_.FullName -notmatch '\\(bin|obj)\\' } |
        Select-String -Pattern '(?:Loc|UiText)\.(?:T|F)\("(I18n_[A-Za-z0-9_]+)"' -AllMatches -CaseSensitive |
        ForEach-Object { $_.Matches | ForEach-Object { $_.Groups[1].Value } }
) | Sort-Object -Unique -CaseSensitive
$missingCode = @($codeKeys | Where-Object { $_ -cnotin $baseKeys })

$toolNames = @(
    Get-ChildItem -LiteralPath (Join-Path $ProjectRoot 'WordOllama\lib\Tools') -Filter *.cs |
        Select-String -Pattern 'public string Name\s*=>\s*"([a-z0-9_]+)"' -AllMatches -CaseSensitive |
        ForEach-Object { $_.Matches | ForEach-Object { $_.Groups[1].Value } }
) | Sort-Object -Unique -CaseSensitive
$missingToolDescriptions = @($toolNames | Where-Object { "I18n_AgentToolDesc_$_" -cnotin $baseKeys })

[xml]$projectXml = [System.IO.File]::ReadAllText(
    (Join-Path $ProjectRoot 'WordOllama\WordOllama.csproj'),
    [System.Text.Encoding]::UTF8)
$compileItems = @($projectXml.SelectNodes("//*[local-name()='Compile']") | ForEach-Object { $_.Include })
$toolSourceItems = @(
    Get-ChildItem -LiteralPath (Join-Path $ProjectRoot 'WordOllama\lib\Tools') -Filter *.cs |
        ForEach-Object { "lib\Tools\$($_.Name)" }
)
$missingToolCompileItems = @($toolSourceItems | Where-Object { $_ -cnotin $compileItems })

if ($missingEnglish.Count -gt 0) {
    Write-Host "Missing English resource keys:" -ForegroundColor Red
    $missingEnglish | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
}
if ($missingBase.Count -gt 0) {
    Write-Host "English-only resource keys:" -ForegroundColor Red
    $missingBase | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
}
if ($missingXaml.Count -gt 0) {
    Write-Host "XAML keys missing from Resources.resx:" -ForegroundColor Red
    $missingXaml | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
}
if ($missingCode.Count -gt 0) {
    Write-Host "C# localization keys missing from Resources.resx:" -ForegroundColor Red
    $missingCode | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
}
if ($missingToolDescriptions.Count -gt 0) {
    Write-Host "Agent tool descriptions missing from Resources.resx:" -ForegroundColor Red
    $missingToolDescriptions | ForEach-Object { Write-Host "  I18n_AgentToolDesc_$_" -ForegroundColor Red }
}
if ($missingToolCompileItems.Count -gt 0) {
    Write-Host "Agent tool source files missing from WordOllama.csproj:" -ForegroundColor Red
    $missingToolCompileItems | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
}

if ($missingEnglish.Count -or $missingBase.Count -or $missingXaml.Count -or $missingCode.Count -or $missingToolDescriptions.Count -or $missingToolCompileItems.Count) {
    exit 1
}

Write-Host "Localization check passed: $($baseKeys.Count) resource keys, $($xamlKeys.Count) XAML Loc keys, and $($toolNames.Count) Agent tool descriptions." -ForegroundColor Green
exit 0
