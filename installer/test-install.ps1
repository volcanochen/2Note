# test-install.ps1 - install the built MSI quietly and verify registration
$ErrorActionPreference = 'Continue'
$msi = "C:\Users\adminlocal\WorkBuddy\2026-11-04-15-46-33\release\2Note-Setup-1.0.0.msi"
$p = Start-Process msiexec.exe -ArgumentList "/i", "`"$msi`"", "/qn", "/l*v", "$env:TEMP\2note-install.log" -Wait -PassThru
Write-Output ("msiexec exit code: " + $p.ExitCode)
Write-Output ("dll installed: " + (Test-Path "$env:LOCALAPPDATA\2Note\MD2OneNote.Addin.dll"))
Get-ChildItem "$env:LOCALAPPDATA\2Note" | ForEach-Object { Write-Output ("  " + $_.Name + " " + $_.Length) }
$k1 = Get-ItemProperty 'HKCU:\Software\Classes\CLSID\{4A9D4FBD-D100-41EF-9C9C-7CD1DB2EE81B}\InprocServer32' -ErrorAction SilentlyContinue
if ($k1) { Write-Output ("CLSID default=" + $k1.'(default)'); Write-Output ("CLSID CodeBase=" + $k1.CodeBase); Write-Output ("CLSID Assembly=" + $k1.Assembly) } else { Write-Output "CLSID MISSING" }
$k2 = Get-ItemProperty 'HKCU:\Software\Microsoft\Office\OneNote\AddIns\MD2OneNote.Connect' -ErrorAction SilentlyContinue
if ($k2) { Write-Output ("AddIns FriendlyName=" + $k2.FriendlyName + " LoadBehavior=" + $k2.LoadBehavior) } else { Write-Output "AddIns MISSING" }
