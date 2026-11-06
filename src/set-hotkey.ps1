# set-hotkey.ps1 - configure MD2OneNote custom hotkey
# Usage:  .\set-hotkey.ps1 Ctrl+Alt+P
#         .\set-hotkey.ps1 -Check Ctrl+Shift+V   (only test availability, don't save)
# Format: modifiers Ctrl/Alt/Shift/Win + a key (A-Z, 0-9, F1-F24, Home/End/Insert/
#         Delete/PageUp/PageDown/arrows/Space/Tab/Enter/Backspace/Esc)
#         e.g. Ctrl+Alt+P | Ctrl+Shift+V | Win+F9 | Ctrl+Alt+Shift+O
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Hotkey,
    [switch]$Check
)
$ErrorActionPreference = 'Stop'
$configPath = Join-Path $PSScriptRoot 'hotkey.txt'

# ---- validate + availability check via RegisterHotKey ----
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class HkCheck {
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool RegisterHotKey(IntPtr h, int id, uint mod, uint vk);
    [DllImport("user32.dll")]
    public static extern bool UnregisterHotKey(IntPtr h, int id);
}
'@

# parse (mirrors HotkeyConfig.TryParse in the add-in)
function Parse-Hotkey([string]$spec) {
    [uint32]0; $key = $null
    foreach ($raw in $spec.Split('+')) {
        $u = $raw.Trim().ToUpperInvariant()
        if ($u.Length -eq 0) { return $null }
        switch -Regex ($u) {
            '^(CTRL|CONTROL)$' { $mod = $mod -bor 2; continue }
            '^(ALT)$'          { $mod = $mod -bor 1; continue }
            '^(SHIFT)$'        { $mod = $mod -bor 4; continue }
            '^(WIN|WINDOWS)$'  { $mod = $mod -bor 8; continue }
        }
        if ($null -ne $key) { return $null }   # two keys
        $key = $u
    }
    if ($null -eq $key -or $mod -eq 0) { return $null }
    if ($key -match '^[A-Z0-9]$') { $vk = [uint32][char]$key[0] }
    elseif ($key -match '^F([1-9]|1[0-9]|2[0-4])$') { $vk = [uint32]0x70 - 1 + [uint32]$Matches[1] }
    else {
        $named = @{ SPACE = 0x20; TAB = 0x09; ENTER = 0x0D; BACKSPACE = 0x08;
                    ESC = 0x1B; ESCAPE = 0x1B; INSERT = 0x2D; DELETE = 0x2E; DEL = 0x2E;
                    HOME = 0x24; END = 0x23; PAGEUP = 0x21; PAGEDOWN = 0x22;
                    LEFT = 0x25; UP = 0x26; RIGHT = 0x27; DOWN = 0x28 }
        if ($named.ContainsKey($key)) { $vk = [uint32]$named[$key] } else { return $null }
    }
    return @{ mod = $mod; vk = $vk }
}

$parsed = Parse-Hotkey $Hotkey
if ($null -eq $parsed) {
    Write-Output "INVALID: '$Hotkey' is not a supported combo."
    Write-Output "Examples: Ctrl+Alt+P | Ctrl+Shift+V | Win+F9 | Ctrl+Alt+Shift+O"
    exit 1
}

# availability: try to register globally (thread hotkey, NULL hwnd)
$free = [HkCheck]::RegisterHotKey([IntPtr]::Zero, 0x4D21, $parsed.mod, $parsed.vk)
if ($free) {
    [HkCheck]::UnregisterHotKey([IntPtr]::Zero, 0x4D21) | Out-Null
    Write-Output "OK: $Hotkey is currently AVAILABLE."
} else {
    Write-Output "BUSY: $Hotkey is registered by another app right now (error $([Runtime.InteropServices.Marshal]::GetLastWin32Error()))."
    Write-Output "Pick another combo, e.g. Ctrl+Alt+P / Ctrl+Shift+J / Win+F9."
}
if ($Check) { exit 0 }

if (-not $free) {
    Write-Output "NOT SAVED. Choose an available combo (use -Check to test)."
    exit 2
}

# ---- save ----
$dir = Split-Path $configPath
if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
$content = "# MD2OneNote hotkey config" + [Environment]::NewLine +
           "# first non-# line = combo; restart OneNote to apply" + [Environment]::NewLine +
           $Hotkey + [Environment]::NewLine
[System.IO.File]::WriteAllText($configPath, $content, (New-Object System.Text.UTF8Encoding($false)))
Write-Output ("SAVED: $Hotkey -> " + $configPath)
Write-Output "Restart OneNote to apply."
