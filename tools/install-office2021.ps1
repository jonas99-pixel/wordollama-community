[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$ManifestPath)
$ErrorActionPreference = 'Stop'
if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) { throw 'Windows required' }
if (-not (Test-Path -LiteralPath $ManifestPath -PathType Leaf)) { throw 'Manifest not found' }
$manifest = (Resolve-Path -LiteralPath $ManifestPath).Path
if ([IO.Path]::GetExtension($manifest) -ine '.vsto') { throw 'Expected .vsto deployment manifest' }
if (-not (Test-Path 'HKLM:\SOFTWARE\Microsoft\Office\16.0\Word\InstallRoot') -and -not (Test-Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Word\InstallRoot')) { throw 'Word desktop 16.0 not detected' }
$common32 = [Environment]::GetFolderPath('CommonProgramFilesX86')
$common = [Environment]::GetFolderPath('CommonProgramFiles')
$paths = @((Join-Path $common 'Microsoft Shared\VSTO\10.0\VSTOInstaller.exe'),(Join-Path $common32 'Microsoft Shared\VSTO\10.0\VSTOInstaller.exe'))
$installer = $paths | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $installer) { throw 'Microsoft VSTO Runtime not detected' }
Write-Host "Microsoft VSTOInstaller: $installer"
Write-Host "Manifest: $manifest"
Write-Host 'Review publisher trust prompts carefully. Cancel unexpected prompts. No certificate import is performed by this script.'
$proc = Start-Process -FilePath $installer -ArgumentList @('/Install',('"' + $manifest + '"')) -PassThru -Wait
if ($proc.ExitCode -ne 0) { throw "VSTO installation failed with code $($proc.ExitCode)" }
Write-Host 'Installer completed. Verify ribbon and actual document editing in Word.'
