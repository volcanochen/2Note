# update-installed-dll.ps1 - rebuild src DLL and deploy to the MSI install location
$ErrorActionPreference = 'Stop'
$ws  = Split-Path $PSScriptRoot -Parent
& (Join-Path $ws 'src\build.ps1')
# kill the add-in surrogate so the DLL is not locked
$sur = Get-CimInstance Win32_Process -Filter "Name='dllhost.exe'" | Where-Object { $_.CommandLine -match '4A9D4FBD' }
foreach ($t in $sur) { Stop-Process -Id $t.ProcessId -Force -Confirm:$false }
Start-Sleep -Seconds 1
$dst = Join-Path $env:LOCALAPPDATA '2Note\MD2OneNote.Addin.dll'
if (-not (Test-Path $dst)) { $dst = 'C:\ProgramData\MD2OneNote\MD2OneNote.Addin.dll' }  # legacy fallback
Copy-Item (Join-Path $ws 'src\MD2OneNote.Addin.dll') $dst -Force -ErrorAction Stop
$d = Get-Item $dst
Write-Output ("deployed: " + $d.FullName + " (" + $d.Length + " bytes, " + $d.LastWriteTime.ToString("HH:mm:ss") + ")")
