taskkill /f /im AutoHotkey.exe
Start-Sleep -Milliseconds 500
Start-Process AutoHotkey.exe -ArgumentList "mousemaster.ahk" -WindowStyle Hidden