[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$SourceManifest,[string]$OutputDirectory='dist\WordOllama-Office2021')
$ErrorActionPreference = 'Stop'
if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) { throw 'Windows build host required' }
$root = Split-Path -Parent $PSScriptRoot
$manifest = (Resolve-Path -LiteralPath $SourceManifest -ErrorAction Stop).Path
if ([IO.Path]::GetExtension($manifest) -ine '.vsto') { throw 'Expected .vsto' }
[xml]$xml = Get-Content -LiteralPath $manifest -Raw
if ($null -eq $xml.SelectSingleNode("//*[local-name()='Signature']")) { throw 'Unsigned deployment manifest rejected' }
$source = Split-Path -Parent $manifest
$appFiles = Join-Path $source 'Application Files'
if (-not (Test-Path -LiteralPath $appFiles -PathType Container)) { throw 'Missing published Application Files folder' }
if ([IO.Path]::IsPathRooted($OutputDirectory)) { $dest = $OutputDirectory } else { $dest = Join-Path $root $OutputDirectory }
if (Test-Path -LiteralPath $dest) { throw 'Target exists: refusing overwrite' }
New-Item -ItemType Directory -Path $dest | Out-Null
Copy-Item -LiteralPath $manifest -Destination $dest
Copy-Item -LiteralPath $appFiles -Destination $dest -Recurse
Copy-Item -LiteralPath (Join-Path $root 'tools\install-office2021.ps1') -Destination $dest
Get-ChildItem -LiteralPath $dest -Recurse -File | Get-FileHash -Algorithm SHA256 | Export-Csv -LiteralPath (Join-Path $dest 'SHA256.csv') -NoTypeInformation
Write-Host "Staged package: $dest"
Write-Warning 'Presence of a Signature element is not signature validation. VSTOInstaller must validate publisher trust on the tablet.'
