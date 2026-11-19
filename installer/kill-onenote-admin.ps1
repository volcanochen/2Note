# kill-onenote-admin.ps1 - run ELEVATED: terminate all OneNote processes
Get-Process ONENOTE, ONENOTEM -ErrorAction SilentlyContinue | Stop-Process -Force -Confirm:$false
"done" | Out-File "$env:TEMP\2note-kill-done.txt" -Encoding ascii
