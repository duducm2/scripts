;---------------------------------------- Github ----------------------------------------------------
; https://github.com/duducm2/zmk-sofle/blob/main/keymap-drawer/eyelash_sofle.svg
;---------------------------------------- Scripts -------------------------------

#Include env.ahk

if (IS_WORK_ENVIRONMENT) {
    response := MsgBox("Can we proceed with Act?", "Act automation", "YesNo")
    if (response = "No") {
        return
    }

    ; TODO: Replace with the actual scripts folder path on the work laptop
    scriptsFolder := "C:\Users\fie7ca\Documents\scripts"
} else {
    scriptsFolder := "C:\Users\eduev\Meu Drive\12 - Scripts"
}

; Ensure the scripts folder is up to date before launching any scripts
SetWorkingDir(scriptsFolder)
RunWait("git fetch", scriptsFolder, "Hide")
RunWait("git pull", scriptsFolder, "Hide")

Sleep 10000

if (IS_WORK_ENVIRONMENT) {
    ; TODO: Update with actual work environment path
    notesFolder := "C:\Users\fie7ca\OneDrive - Bosch Group\14-my-notes"
} else {
    notesFolder := "C:\Users\eduev\Meu Drive\14 - Notes"
}

; Ensure the notes folder is up to date before working with habits
SetWorkingDir(notesFolder)
RunWait("git fetch", notesFolder, "Hide")
RunWait("git pull", notesFolder, "Hide")

Sleep 10000

Run GetScriptPath("Shift keys.ahk")
Run GetScriptPath("Gemini.ahk")
Run GetScriptPath("AppLaunchers.ahk")
Run GetScriptPath("Spotify.ahk")
Run GetScriptPath("Hotstrings.ahk")
Run GetScriptPath("ChatGPT.ahk")
Run GetScriptPath("WindowManagement.ahk")
Run GetScriptPath("Utils.ahk")

if (IS_WORK_ENVIRONMENT) {

    Run "C:\Users\fie7ca\Documents\HuntAndPeck\HuntAndPeck-1.7\hap.exe"
    Run "C:\Users\fie7ca\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Cursor\Cursor.lnk"
    Run "C:\Users\fie7ca\Documents\Shortcuts\Mobills.lnk"
    Run "C:\Users\fie7ca\Documents\Shortcuts\Settle Up.lnk"

    Run GetScriptPath("Microsoft Teams.ahk")
    Run GetScriptPath("Outlook.ahk")
} else {
    Run "C:\Users\eduev\OneDrive\Documentos\HuntAndPeck\hap.exe"
    Run "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Chrome Apps\Mobills.lnk"
    Run "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Chrome Apps\Settle Up.lnk"
    Run "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\apps do Chrome\WhatsApp Web.lnk"
    Run "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Chrome Apps\Gmail.lnk"
}

habitsFolder := notesFolder . "\habits"
excelFile := habitsFolder . "\habit_sleep_food_tracker.xlsx"

Sleep 1000

; Open the Excel file
Run(excelFile)

; Run #!+i for both environments
Send "#!+i"