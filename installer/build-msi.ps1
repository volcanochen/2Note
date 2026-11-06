# build-msi.ps1 - build the 2Note per-user MSI into release\
# Requires wix311 binaries extracted at installer\wix311 (candle.exe/light.exe)
$ErrorActionPreference = 'Stop'
$ws      = Split-Path $PSScriptRoot -Parent                          # repo root
$src     = Join-Path $ws 'src'
$wix     = Join-Path $PSScriptRoot 'wix311'
$candle  = Join-Path $wix 'candle.exe'
$light   = Join-Path $wix 'light.exe'
$release = Join-Path $ws 'release'
$obj     = Join-Path $env:TEMP '2Note-msi-obj'

# 0. fresh copy of README for packaging
Copy-Item (Join-Path $ws 'README.md') (Join-Path $src 'README.md') -Force

# 1. build the DLL
& (Join-Path $src 'build.ps1')
if (-not (Test-Path (Join-Path $src 'MD2OneNote.Addin.dll'))) { throw "DLL build failed" }

# 2. compile + link MSI
New-Item -ItemType Directory -Force -Path $obj, $release | Out-Null
& $candle -nologo -arch x64 "-dSourceDir=$src" -out "$obj\2Note.wixobj" (Join-Path $PSScriptRoot '2Note.wxs')
if ($LASTEXITCODE -ne 0) { throw "candle failed ($LASTEXITCODE)" }

$msi = Join-Path $release '2Note-Setup-1.0.0.msi'
& $light -nologo -out $msi "$obj\2Note.wixobj"
if ($LASTEXITCODE -ne 0) { throw "light failed ($LASTEXITCODE)" }

$m = Get-Item $msi
Write-Output ("MSI OK: " + $m.FullName + " (" + $m.Length + " bytes)")
