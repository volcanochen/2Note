# cleanup-legacy.ps1 - remove the pre-MSI legacy HKLM registration that shadows the
# per-user HKCU install. DCOM launches the add-in surrogate (DllHost) in the SYSTEM
# context, whose merged HKCR view only contains HKLM -> the OLD DLL under
# C:\ProgramData\MD2OneNote keeps getting loaded even after the MSI installs the new
# one under %LOCALAPPDATA%\2Note. Run this once (accept the UAC prompt), then
# restart OneNote.
# Self-elevating.
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    exit
}

$clsid = '{4A9D4FBD-D100-41EF-9C9C-7CD1DB2EE81B}'

Write-Host "== removing legacy HKLM CLSID registration =="
Remove-Item -Path "HKLM:\Software\Classes\CLSID\$clsid" -Recurse -Force -ErrorAction Continue
Remove-Item -Path "HKLM:\Software\Microsoft\Office\OneNote\AddIns\MD2OneNote.Connect" -Force -ErrorAction Continue

Write-Host "== removing legacy install folder C:\ProgramData\MD2OneNote =="
if (Test-Path 'C:\ProgramData\MD2OneNote') {
    Remove-Item 'C:\ProgramData\MD2OneNote' -Recurse -Force -ErrorAction Continue
}

Write-Host ""
Write-Host "Done. Please restart OneNote now (fully quit, then reopen)."
Write-Host "Press Enter to close..."
Read-Host
