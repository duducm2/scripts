$scriptsDir = Split-Path -Parent $PSScriptRoot
taskkill /f /im AutoHotkey.exe
Start-Sleep -Milliseconds 500
Start-Process AutoHotkey.exe -ArgumentList (Join-Path $scriptsDir "mousemaster.ahk") -WindowStyle Hidden
