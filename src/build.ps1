# build MD2OneNote.Addin.dll - pure late binding, NO PIA dependency
$ErrorActionPreference = 'Stop'
$build = $PSScriptRoot
$csc = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$out = "$build\MD2OneNote.Addin.dll"
if (Test-Path $out) { Remove-Item $out -Force }

& $csc /nologo /target:library /platform:x64 /optimize+ `
    "/out:$out" `
    /r:System.dll /r:System.Core.dll `
    "$build\MD2OneNote.cs"
if ($LASTEXITCODE -eq 0 -and (Test-Path $out)) { Write-Output "BUILD OK (no PIA)"; exit 0 }
throw "compile failed"
