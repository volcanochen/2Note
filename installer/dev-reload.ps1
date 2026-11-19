# dev-reload.ps1 - one-command dev iteration: rebuild + deploy to BOTH install
# locations + kill the DllHost surrogate. OneNote must still be restarted manually
# (the assembly is cached in the surrogate and OneNote holds a live reference).
$ErrorActionPreference = 'Stop'
$ws = Split-Path $PSScriptRoot -Parent
& (Join-Path $ws 'src\build.ps1')
$sur = Get-CimInstance Win32_Process -Filter "Name='dllhost.exe'" | Where-Object { $_.CommandLine -match '4A9D4FBD' }
foreach ($t in $sur) { Stop-Process -Id $t.ProcessId -Force -Confirm:$false }
Start-Sleep -Seconds 1
Copy-Item (Join-Path $ws 'src\MD2OneNote.Addin.dll') (Join-Path $env:LOCALAPPDATA '2Note\MD2OneNote.Addin.dll') -Force
try { Copy-Item (Join-Path $ws 'src\MD2OneNote.Addin.dll') 'C:\ProgramData\MD2OneNote\MD2OneNote.Addin.dll' -Force -ErrorAction Stop; Write-Output 'legacy location updated (HKLM cleanup pending)' } catch { Write-Output 'WARN: legacy ProgramData copy failed (locked?)' }
Write-Output 'deployed. Now restart OneNote to load the new code.'
