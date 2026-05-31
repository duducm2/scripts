;---------------------------------------- Github ----------------------------------------------------
; https://github.com/duducm2/zmk-sofle/blob/main/keymap-drawer/eyelash_sofle.svg
;---------------------------------------- Scripts -------------------------------

#Include env.ahk
#Include %A_ScriptDir%\Utils.ahk

; Run a command with a timeout.
; Returns the process exit code, or 124 on timeout.
RunWaitWithTimeout(cmd, workingDir := "", options := "", timeoutMs := 120000) {
    ; AHK v2 ProcessWaitClose() does not return exit code; it returns 0 on success and PID on timeout.
    ; Use PowerShell to enforce timeout and propagate the real ExitCode.
    safeWorkDir := StrReplace(workingDir, "'", "''")
    safeCmd := StrReplace(cmd, "'", "''")

    ps := ""
        . "$ErrorActionPreference='Stop';"
        . "$cmd='" . safeCmd . "';"
        . "$wd='" . safeWorkDir . "';"
        . "$t=[int]" . timeoutMs . ";"
        .
        "$p=Start-Process -FilePath 'cmd.exe' -ArgumentList @('/v:on','/c',$cmd) -WorkingDirectory $wd -PassThru -WindowStyle Hidden;"
        . "if(-not $p.WaitForExit($t)){try{$p.Kill()}catch{}; exit 124};"
        . "exit $p.ExitCode"

    ; RunWait returns the exit code of PowerShell, which we set to either 124 or the child's ExitCode.
    try return RunWait("powershell.exe -NoProfile -ExecutionPolicy Bypass -Command " . Chr(34) . ps . Chr(34),
    workingDir, options)
    catch as e {
        return 1
    }
}

GitInRepoOrFail(repoDir, gitArgs, timeoutMs := 120000) {
    ; Prevent silent hangs: disable interactive credential prompts (GCM + terminal),
    ; and allow longer operations on slow links.
    cmd := 'set GIT_TERMINAL_PROMPT=0& set GCM_INTERACTIVE=Never& git ' . gitArgs
    displayCmd := A_ComSpec . ' /v:on /c "' . cmd . '"'
    exitCode := RunWaitWithTimeout(cmd, repoDir, "Min", timeoutMs)
    if (exitCode = 124) {
        StandardLoadingBar_Hide(0)
        MsgBox(
            "Timed out while running:`n`n" . displayCmd . "`n`nIn:`n" . repoDir . "`n`n" .
            "This usually means git is waiting for input (credentials/2FA) or is blocked by network/proxy.`n" .
            "Try running the same command in a terminal in that folder to see the prompt/output.",
            "Act automation", "Icon!"
        )
        ExitApp
    }
    if (exitCode != 0) {
        StandardLoadingBar_Hide(0)
        MsgBox(
            "Git command failed (exit code " . exitCode . "):`n`n" . displayCmd . "`n`nIn:`n" . repoDir . "`n`n" .
            "Run it in a terminal in that folder to see the error output.",
            "Act automation", "Icon!"
        )
        ExitApp
    }
}

if (IS_WORK_ENVIRONMENT) {
    response := MsgBox("Can we proceed with Act?", "Act automation", "YesNo")
    if (response = "No") {
        return
    }
}

scriptsFolder := GetScriptsRepoPath()
if (!scriptsFolder) {
    MsgBox("Scripts repo folder not found. Check WORK_SCRIPTS_PATH and PERSONAL_SCRIPTS_PATH in env.ahk.",
        "Act automation", "Icon!")
    return
}

StandardLoadingBar_Show("⏳ Updating scripts...", BANNER_ACCENT_INTERMEDIATE)
SetWorkingDir(scriptsFolder)
GitInRepoOrFail(scriptsFolder, "fetch --prune", 900000)
GitInRepoOrFail(scriptsFolder, "pull", 900000)
StandardLoadingBar_Update("⏳ Waiting...")
Sleep 10000

notesFolder := GetNotesRepoPath()
if (!notesFolder) {
    StandardLoadingBar_Hide(0)
    MsgBox("Notes repo folder not found. Check NOTES_REPO_PATH_* in env.ahk.", "Act automation", "Icon!")
    return
}

StandardLoadingBar_Update("⏳ Updating notes...")
SetWorkingDir(notesFolder)
GitInRepoOrFail(notesFolder, "fetch --prune", 900000)
GitInRepoOrFail(notesFolder, "pull", 900000)
; MyNotes technique prompts are read from disk when you use them in Utility Shortcuts (Utils.ahk), not by Act.
StandardLoadingBar_Update("🚀 Launching apps...")
Sleep 10000

; Start QuickLook (study viewer) based on the current environment in env.ahk
quicklookExe := GetQuickLookExePath()
if (quicklookExe)
    Run quicklookExe

Run GetScriptPath("Shift keys.ahk")
Run GetScriptPath("Gemini.ahk")
Run GetScriptPath("AppLaunchers.ahk")
Run GetScriptPath("WindowManagement.ahk")
; Do not Run Utils.ahk here: AppLaunchers.ahk already #includes Utils.ahk. A second Utils process duplicates
; keyboard hooks (e.g. global Escape) and breaks modals that rely on g_OnEscapePressed / I10 in AppLaunchers.
Run GetScriptPath("Mousemaster.ahk")

if (IS_WORK_ENVIRONMENT) {
    Run "C:\Users\fie7ca\Documents\caffeine\caffeine64.exe"
    Run "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Visual Studio Code\Visual Studio Code.lnk"
    Run "C:\Users\fie7ca\Documents\Shortcuts\Mobills.lnk"
    Run "C:\Users\fie7ca\Documents\Shortcuts\Settle Up.lnk"

    Run GetScriptPath("Microsoft Teams.ahk")
    Run GetScriptPath("Outlook.ahk")
} else {
    Run "C:\Users\eduev\Documents\Atalhos\Mobills.lnk"
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