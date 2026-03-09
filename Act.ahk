;---------------------------------------- Github ----------------------------------------------------
; https://github.com/duducm2/zmk-sofle/blob/main/keymap-drawer/eyelash_sofle.svg
;---------------------------------------- Scripts -------------------------------

#Include env.ahk
#Include %A_ScriptDir%\Utils.ahk

if (IS_WORK_ENVIRONMENT) {
    response := MsgBox("Can we proceed with Act?", "Act automation", "YesNo")
    if (response = "No") {
        return
    }

    ; TODO: Replace with the actual scripts folder path on the work laptop
    scriptsFolder := "C:\Users\fie7ca\Documents\scripts"
} else {
    scriptsFolder := "C:\Users\eduev\Meu Drive\17 - Projects\scripts"
}

StandardLoadingBar_Show("⏳ Updating scripts...", BANNER_ACCENT_INTERMEDIATE)
SetWorkingDir(scriptsFolder)
RunWait("git fetch", scriptsFolder, "Hide")
RunWait("git pull", scriptsFolder, "Hide")
StandardLoadingBar_Update("⏳ Waiting...")
Sleep 10000

if (IS_WORK_ENVIRONMENT) {
    ; TODO: Update with actual work environment path
    notesFolder := "C:\Users\fie7ca\OneDrive - Bosch Group\14-my-notes"
} else {
    notesFolder := "C:\Users\eduev\Meu Drive\17 - Projects\notes"
}

StandardLoadingBar_Update("⏳ Updating notes...")
SetWorkingDir(notesFolder)
RunWait("git fetch", notesFolder, "Hide")
RunWait("git pull", notesFolder, "Hide")
StandardLoadingBar_Update("🚀 Launching apps...")
Sleep 10000

Run GetScriptPath("Shift keys.ahk")
Run GetScriptPath("Gemini.ahk")
Run GetScriptPath("AppLaunchers.ahk")
Run GetScriptPath("WindowManagement.ahk")
Run GetScriptPath("Utils.ahk")

if (IS_WORK_ENVIRONMENT) {

    Run "C:\Users\fie7ca\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Cursor\Cursor.lnk"
    Run "C:\Users\fie7ca\Documents\Shortcuts\Mobills.lnk"
    Run "C:\Users\fie7ca\Documents\Shortcuts\Settle Up.lnk"

    Run GetScriptPath("Microsoft Teams.ahk")
    Run GetScriptPath("Outlook.ahk")
} else {
    Run "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Chrome Apps\Mobills.lnk"
    Run "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Chrome Apps\Settle Up.lnk"
    Run "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\apps do Chrome\WhatsApp Web.lnk"
    Run "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Chrome Apps\Gmail.lnk"
}

habitsFolder := notesFolder . "\habits"
excelFile := habitsFolder . "\habit_sleep_food_tracker.xlsx"
StandardLoadingBar_Update("✅ Done", BANNER_ACCENT_SUCCESS)
StandardLoadingBar_Hide(500)
Sleep 1000
Run(excelFile)