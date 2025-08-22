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

    Run "C:\Users\fie7ca\Documents\HuntAndPeck\HuntAndPeck-1.7\hap.exe"
    Run "C:\Users\fie7ca\Documents\Shortcuts\Mobills.lnk"
    Run "C:\Users\fie7ca\Documents\Shortcuts\Gmail.lnk"
    Run "C:\Users\fie7ca\Documents\Shortcuts\Settle Up.lnk"
    Run "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\OneNote.lnk"
    Run "https://keep.google.com/u/0/#NOTE/1YCVkrriqNRyRhRyz1PV0gJ9Eor66ARhb1i9uLpvXTZ2j79nDScAUOIK4pBAMwHY"
    Run GetScriptPath("Microsoft Teams.ahk")
    Run GetScriptPath("Outlook.ahk")
} else {
    Run "C:\Users\eduev\OneDrive\Documentos\HuntAndPeck\hap.exe"
    Run "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Chrome Apps\Mobills.lnk"
    Run "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Chrome Apps\Settle Up.lnk"
    Run "https://keep.google.com/u/0/#NOTE/1YCVkrriqNRyRhRyz1PV0gJ9Eor66ARhb1i9uLpvXTZ2j79nDScAUOIK4pBAMwHY"
    Run "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\apps do Chrome\WhatsApp Web.lnk"
    Run "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Chrome Apps\Gmail.lnk"
}

Sleep 10000

; Run #!+i for both environments
Send "#!+i"