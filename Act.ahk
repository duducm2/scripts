;---------------------------------------- Github ----------------------------------------------------
; https://github.com/duducm2/zmk-sofle/blob/main/keymap-drawer/eyelash_sofle.svg
;---------------------------------------- Scripts -------------------------------
#Requires AutoHotkey v2.0+
#SingleInstance Force

#Include env.ahk

; Thin bootstrap only: do NOT #include full Utils.ahk.
; Full Utils would take D2C_Dictation_Hotkey_Owner / Escape I10 before git pull, then
; leave Act owning post-dictation UX with stale in-memory code while children load new Utils.
; See Utils/dictation_toggle.ahk Dictation_IsOwnerProcess (AppLaunchers only).
#Include %A_ScriptDir%\Utils\git_cli.ahk
#Include %A_ScriptDir%\Utils\standard_loading_bar.ahk

; standard_loading_bar keys-overlay paths reference this; Act never uses ShowWithKeys.
; Stub avoids pulling print_screen_escape (second Escape host breaks AppLaunchers modals).
Utils_EnsureGlobalEscapeHotkey() {
}

; Preflight: catch env.ahk drift before launching hosts that load lib/CopilotWeb.ahk
verifyPs1 := A_ScriptDir "\infra\ipc\Verify-EnvAhk.ps1"
if FileExist(verifyPs1) {
    exitCode := RunWait(
        'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "' verifyPs1 '" -ScriptsDir "' A_ScriptDir '"',
        , "Hide")
    if (exitCode != 0) {
        MsgBox(
            "env.ahk failed preflight (exit " exitCode ").`n`nRun:`n  powershell -File infra\ipc\Verify-EnvAhk.ps1`n`nCompare env.ahk with env.ahk.example.",
            "Act automation", "Icon!")
        ExitApp exitCode
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
        ExitApp
    }
}

scriptsFolder := GetScriptsRepoPath()
if (!scriptsFolder) {
    MsgBox("Scripts repo folder not found. Check WORK_SCRIPTS_PATH and PERSONAL_SCRIPTS_PATH in env.ahk.",
        "Act automation", "Icon!")
    ExitApp
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
    ExitApp
}

StandardLoadingBar_Update("⏳ Updating notes...")
SetWorkingDir(notesFolder)
GitInRepoOrFail(notesFolder, "fetch --prune", 900000)
GitInRepoOrFail(notesFolder, "pull", 900000)
; MyNotes technique prompts are read from disk when you use them in Utility Shortcuts (Utils.ahk), not by Act.
StandardLoadingBar_Update("🚀 Launching apps...")
; Pre-start Tasks (:8766) and Memory Palace (:8767) so first open is not cold.
Run GetScriptPath("Utils\web_servers_warmup.ahk")
Sleep 10000

; Start QuickLookThat's interesting (study viewer) based on the current environment in env.ahk
quicklookExe := GetQuickLookExePath()
if (quicklookExe)
    Run quicklookExe

Run GetScriptPath("Shift keys.ahk")
Run GetScriptPath("Gemini.ahk")
Run GetScriptPath("AppLaunchers.ahk")
Run GetScriptPath("WindowManagement.ahk")
Run GetScriptPath("Spotify.ahk")
; Do not Run Utils.ahk here: AppLaunchers.ahk already #includes Utils.ahk. A second Utils process duplicates
; keyboard hooks (e.g. global Escape) and breaks modals that rely on g_OnEscapePressed / I10 in AppLaunchers.
; Dictation (~#!+0 / Send dictation?) is owned only by AppLaunchers (Dictation_IsOwnerProcess).
Run GetScriptPath("Mousemaster.ahk")

if (IS_WORK_ENVIRONMENT) {
    Run "C:\Users\fie7ca\Documents\caffeine\caffeine64.exe"
    Run "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Visual Studio Code\Visual Studio Code.lnk"
    Run "C:\Users\fie7ca\Documents\Shortcuts\Settle Up.lnk"

    Run GetScriptPath("Microsoft Teams.ahk")
    Run GetScriptPath("Outlook.ahk")
} else {
    Run "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Chrome Apps\Settle Up.lnk"
    Run "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\apps do Chrome\WhatsApp Web.lnk"
    Run "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Chrome Apps\Gmail.lnk"
}

habitsFolder := notesFolder . "\habits"
StandardLoadingBar_Update("✅ Done", BANNER_ACCENT_SUCCESS)
StandardLoadingBar_Hide(500)
Sleep 1000
; Exit so Act never remains as a lingering Utils/hotkey host after bootstrap.
ExitApp