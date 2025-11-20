;---------------------------------------- Github ----------------------------------------------------
; https://github.com/duducm2/zmk-sofle/blob/main/keymap-drawer/eyelash_sofle.svg
;---------------------------------------- Scripts -------------------------------

#Include env.ahk

if (IS_WORK_ENVIRONMENT) {
    ; TODO: Replace with the actual scripts folder path on the work laptop
    scriptsFolder := "C:\Path\To\Work\Scripts"
} else {
    scriptsFolder := "C:\Users\eduev\Meu Drive\12 - Scripts"
}

; Ensure the scripts folder is up to date before launching any scripts
SetWorkingDir(scriptsFolder)
RunWait("git fetch", scriptsFolder, "Hide")
RunWait("git pull", scriptsFolder, "Hide")

Run GetScriptPath("Shift keys.ahk")
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

Sleep 10000

; Run #!+i for both environments
Send "#!+i"

; Ask user if they want to update the Habits Sheet
result := MsgBox("Do you want to update the Habits Sheet?", "Update Habits Sheet", "YesNo")
if (result = "No") {
    return
}

; If Yes, proceed with updating habits
if (IS_WORK_ENVIRONMENT) {
    ; TODO: Update with actual work environment path
    notesFolder := "C:\Users\fie7ca\OneDrive - Bosch Group\14-my-notes"
} else {
    notesFolder := "C:\Users\eduev\Meu Drive\14 - Notes"
}

habitsFolder := notesFolder . "\habits"
excelFile := habitsFolder . "\habit_sleep_food_tracker.xlsx"

; Ensure the notes folder is up to date before opening the habits sheet
SetWorkingDir(notesFolder)
RunWait("git pull", notesFolder, "Hide")

Sleep 3000

; Change to the habits folder and run git pull
SetWorkingDir(habitsFolder)
RunWait("git pull", habitsFolder, "Hide")

Sleep 3000

; Open the Excel file
Run(excelFile)

; Ask user if they want to update Notes folder
result := MsgBox("Do you want to git pull your Notes folder?", "Update Notes Folder", "YesNo")
if (result = "No") {
    return
}

; If Yes, proceed with updating notes
SetWorkingDir(notesFolder)
RunWait("git pull", notesFolder, "Hide")