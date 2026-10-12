param(
    [string]$ProjectRoot = (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path))
)

$ErrorActionPreference = 'Stop'
$assemblyPath = Join-Path $ProjectRoot 'WordOllama\bin\Release\WordOllama.dll'
if (-not (Test-Path -LiteralPath $assemblyPath)) { throw "Release assembly not found: $assemblyPath" }
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') { throw 'Run with Windows PowerShell -STA.' }

$assembly = [Reflection.Assembly]::LoadFrom($assemblyPath)
$paneType = $assembly.GetType('WordOllama.NewUI.AgentTaskPaneUI', $true)
$indexMethod = $paneType.GetMethod('FindParagraphIndexByRange', [Reflection.BindingFlags]'NonPublic,Static')
if ($null -eq $indexMethod) { throw 'Paragraph index helper was not found.' }

Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class ReviewWordTestWindow {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
}
'@

$existingWordPids = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$wordTestApp = $null
$wordTestDoc = $null
$ownsTestApp = $false
try {
    $wordTestApp = New-Object -ComObject Word.Application
    $wordTestDoc = $wordTestApp.Documents.Add()
    [uint32]$wordTestPid = 0
    [void][ReviewWordTestWindow]::GetWindowThreadProcessId([IntPtr]$wordTestDoc.ActiveWindow.Hwnd, [ref]$wordTestPid)
    $ownsTestApp = $wordTestPid -gt 0 -and $existingWordPids -notcontains [int]$wordTestPid
    if (-not $ownsTestApp) { throw 'Refusing to operate an existing Word process.' }
    $wordTestApp.Visible = $false
    $wordTestApp.DisplayAlerts = 0

    $paragraphs = @('First paragraph.') + (1..118 | ForEach-Object { "Body paragraph $_." }) + @('Last paragraph.')
    $wordTestDoc.Content.Text = ($paragraphs -join "`r")
    foreach ($expected in @(1, 2, 80, 81, 85, 120)) {
        $paragraph = $wordTestDoc.Paragraphs.Item($expected)
        $range = $paragraph.Range
        try {
            $actual = $indexMethod.Invoke($null, @($wordTestDoc, $range))
            if ($actual -ne $expected) { throw "Paragraph-start index: expected $expected, got $actual." }
            $range.SetRange($range.Start + 2, $range.End - 1)
            $actual = $indexMethod.Invoke($null, @($wordTestDoc, $range))
            if ($actual -ne $expected) { throw "Inside-paragraph index: expected $expected, got $actual." }
        }
        finally {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($range)
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($paragraph)
        }
    }
    Write-Output 'PASS: Word paragraph indexes at starts and inside paragraphs, including 80, 81, 85 and 120.'

    $wordTestDoc.Content.Text = ('Long paragraph sentence. ' * 5000)
    $range = $wordTestDoc.Content.Duplicate
    try {
        $range.SetRange(6100, 100100)
        if ($indexMethod.Invoke($null, @($wordTestDoc, $range)) -ne 1) { throw 'Long selection index mismatch.' }
        if ($range.Text.Length -ne 94000) { throw 'Exact selection clipping mismatch.' }
    }
    finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($range) }
    Write-Output 'PASS: exact 94000-character Word selection inside an oversized paragraph.'
}
finally {
    if ($null -ne $wordTestDoc) {
        $wordTestDoc.Close(0)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordTestDoc)
    }
    if ($null -ne $wordTestApp) {
        if ($ownsTestApp) { $wordTestApp.Quit() }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordTestApp)
    }
}
