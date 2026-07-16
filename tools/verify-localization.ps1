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
$missingEnglish = @($baseKeys | Where-Object { $_ -notin $englishKeys })
$missingBase = @($englishKeys | Where-Object { $_ -notin $baseKeys })

$xamlKeys = @(
    Get-ChildItem -LiteralPath (Join-Path $ProjectRoot 'WordOllama') -Recurse -Filter *.xaml |
        Select-String -Pattern '\{local:Loc\s+(I18n_[A-Za-z0-9_]+)' |
        ForEach-Object { $_.Matches | ForEach-Object { $_.Groups[1].Value } }
) | Sort-Object -Unique
$missingXaml = @($xamlKeys | Where-Object { $_ -notin $baseKeys })

$codeKeys = @(
    Get-ChildItem -LiteralPath (Join-Path $ProjectRoot 'WordOllama') -Recurse -Filter *.cs |
        Where-Object { $_.FullName -notmatch '\\(bin|obj)\\' } |
        Select-String -Pattern '(?:Loc|UiText)\.(?:T|F)\("(I18n_[A-Za-z0-9_]+)"' |
        ForEach-Object { $_.Matches | ForEach-Object { $_.Groups[1].Value } }
) | Sort-Object -Unique
$missingCode = @($codeKeys | Where-Object { $_ -notin $baseKeys })

$toolNames = @(
    Get-ChildItem -LiteralPath (Join-Path $ProjectRoot 'WordOllama\lib\Tools') -Filter *.cs |
        Select-String -Pattern 'public string Name\s*=>\s*"([a-z0-9_]+)"' |
        ForEach-Object { $_.Matches | ForEach-Object { $_.Groups[1].Value } }
) | Sort-Object -Unique
$missingToolDescriptions = @($toolNames | Where-Object { "I18n_AgentToolDesc_$_" -notin $baseKeys })

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

if ($missingEnglish.Count -or $missingBase.Count -or $missingXaml.Count -or $missingCode.Count -or $missingToolDescriptions.Count) {
    exit 1
}

Write-Host "Localization check passed: $($baseKeys.Count) resource keys, $($xamlKeys.Count) XAML Loc keys, and $($toolNames.Count) Agent tool descriptions." -ForegroundColor Green
