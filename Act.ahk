;---------------------------------------- Github ----------------------------------------------------
; https://github.com/duducm2/zmk-sofle/blob/main/keymap-drawer/eyelash_sofle.svg
;---------------------------------------- Scripts ---------------------------------------------------

#Include env.ahk

Run GetScriptPath("Shift keys.ahk")
Run GetScriptPath("AppLaunchers.ahk")
Run GetScriptPath("Spotify.ahk")
Run GetScriptPath("ChatGPT.ahk")
Run GetScriptPath("WindowManagement.ahk")
Run GetScriptPath("Utils.ahk")

if (IS_WORK_ENVIRONMENT) {

    Run "C:\Users\fie7ca\Documents\HuntAndPack\hap.exe"
    Run "C:\Users\fie7ca\Documents\Caffeine\caffeine64.exe"
    Run "C:\Users\fie7ca\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\apps do Chrome\Mobills"
    Run "C:\Users\fie7ca\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\apps do Chrome\Gmail"
    Run "c:\Users\fie7ca\Documents\Settle Up.lnk"
    Run "c:\ProgramData\Microsoft\Windows\Start Menu\Programs\OneNote.lnk"
    Run "cc:\Users\fie7ca\Documents\Atalhos\Microsoft Teams - Shortcut.lnk"
    Run GetScriptPath("Microsoft Teams.ahk")
    Run GetScriptPath("Outlook.ahk")
} else {
    Run "C:\Users\eduev\OneDrive\Documentos\HuntAndPeck\hap.exe"
    Run "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Chrome Apps\Mobills.lnk"
    Run "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Chrome Apps\Settle Up.lnk"
    Run "https://keep.google.com/u/0/#NOTE/1YCVkrriqNRyRhRyz1PV0gJ9Eor66ARhb1i9uLpvXTZ2j79nDScAUOIK4pBAMwHY"
    Run "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\apps do Chrome\WhatsApp Web.lnk"
}

Sleep 10000

; Run #!+i for both environments
Send "#!+i"