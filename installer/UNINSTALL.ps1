# UNINSTALL.ps1 - remove the portable 2Note registration (does not delete files)
$clsid = '{4A9D4FBD-D100-41EF-9C9C-7CD1DB2EE81B}'
Remove-Item "HKCU:\Software\Classes\CLSID\$clsid" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item 'HKCU:\Software\Microsoft\Office\OneNote\AddIns\MD2OneNote.Connect' -Recurse -Force -ErrorAction SilentlyContinue
Write-Output '2Note registration removed. Restart OneNote. You may now delete this folder.'
