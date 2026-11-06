# INSTALL.ps1 - register the portable 2Note add-in for the current user (no admin)
$ErrorActionPreference = 'Stop'
$dir      = $PSScriptRoot
$codeBase = Join-Path $dir 'MD2OneNote.Addin.dll'
if (-not (Test-Path $codeBase)) { throw "MD2OneNote.Addin.dll not found next to this script" }

$clsid = '{4A9D4FBD-D100-41EF-9C9C-7CD1DB2EE81B}'
$k = "HKCU:\Software\Classes\CLSID\$clsid\InprocServer32"
New-Item -Path $k -Force | Out-Null
Set-ItemProperty -Path $k -Name '(default)'      -Value 'mscoree.dll'
Set-ItemProperty -Path $k -Name 'Assembly'       -Value 'MD2OneNote.Addin, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null'
Set-ItemProperty -Path $k -Name 'Class'          -Value 'MD2OneNote.Connect'
Set-ItemProperty -Path $k -Name 'RuntimeVersion' -Value 'v4.0.30319'
Set-ItemProperty -Path $k -Name 'CodeBase'       -Value $codeBase
Set-ItemProperty -Path $k -Name 'ThreadingModel' -Value 'Both'

$a = 'HKCU:\Software\Microsoft\Office\OneNote\AddIns\MD2OneNote.Connect'
New-Item -Path $a -Force | Out-Null
Set-ItemProperty -Path $a -Name 'LoadBehavior'  -Value 3 -Type DWord
Set-ItemProperty -Path $a -Name 'FriendlyName'  -Value '2Note'
Set-ItemProperty -Path $a -Name 'Description'   -Value 'Paste HTML/Markdown as native OneNote structure'

Write-Output '2Note registered for the current user.'
Write-Output 'Restart OneNote to load it. Uninstall: run UNINSTALL.ps1'
