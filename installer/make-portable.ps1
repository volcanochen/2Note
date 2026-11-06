# make-portable.ps1 - build release\2Note-1.0.0-portable.zip
$ErrorActionPreference = 'Stop'
$ws      = Split-Path $PSScriptRoot -Parent
$src     = Join-Path $ws 'src'
$release = Join-Path $ws 'release'
$stage   = Join-Path $env:TEMP '2Note-portable'
if (Test-Path $stage) { Remove-Item $stage -Recurse -Force }
New-Item -ItemType Directory -Force -Path $stage | Out-Null

Copy-Item (Join-Path $src 'MD2OneNote.Addin.dll') $stage -Force
Copy-Item (Join-Path $src 'set-hotkey.ps1')       $stage -Force
Copy-Item (Join-Path $src 'hotkey.txt')           $stage -Force
Copy-Item (Join-Path $ws  'README.md')            $stage -Force
Copy-Item (Join-Path $PSScriptRoot 'INSTALL.ps1')   $stage -Force
Copy-Item (Join-Path $PSScriptRoot 'UNINSTALL.ps1') $stage -Force

$zip = Join-Path $release '2Note-1.0.0-portable.zip'
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path "$stage\*" -DestinationPath $zip -Force
$z = Get-Item $zip
Write-Output ("portable zip OK: " + $z.FullName + " (" + $z.Length + " bytes)")
