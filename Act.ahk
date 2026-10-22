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

; Thin Clip Angel stack (not full Utils.ahk) so Act can guarantee Clip Angel is up.
#Include %A_ScriptDir%\Utils\clip_angel_act_bootstrap.ahk

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

personalFolder := GetPersonalRepoPath()
if (!personalFolder) {
    StandardLoadingBar_Hide(0)
    MsgBox("Personal repo folder not found. Check PERSONAL_REPO_PATH_WORK and PERSONAL_REPO_PATH_PERSONAL in env.ahk.",
        "Act automation", "Icon!")
    ExitApp
}

StandardLoadingBar_Update("⏳ Updating personal repo...")
SetWorkingDir(personalFolder)
Act_PersonalRepoStashAndPull(personalFolder)
StandardLoadingBar_Update("🚀 Launching apps...")
; Persistent keep-alive: Tasks (:8766), Memory Palace (:8767), Finance dashboard (:8765).
; Heartbeats every 30s; Utils reload no longer kills these Python servers.
Run GetScriptPath("Utils\web_servers_warmup.ahk")

; Start QuickLook (study viewer) early so a long warmup sleep cannot skip it if Act is stopped.
quicklookExe := GetQuickLookExePath()
if (quicklookExe = "" || !FileExist(quicklookExe))
    quicklookExe := "C:\QuickLook\QuickLook.exe"
if (FileExist(quicklookExe)) {
    StandardLoadingBar_Update("🚀 Starting QuickLook...")
    try {
        if !ProcessExist("QuickLook.exe")
            Run '"' quicklookExe '"'
    } catch as e {
        StandardLoadingBar_Update("⚠ QuickLook launch failed", BANNER_ACCENT_ERROR)
        Sleep 800
    }
} else {
    StandardLoadingBar_Update("⚠ QuickLook.exe not found", BANNER_ACCENT_ERROR)
    Sleep 800
}

Sleep 10000

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
    ; Headset / other endpoints often are not connected yet. Fill a dead input or
    ; output with this PC's Realtek mic and speakers (the devices marked in the
    ; Audio / Bluetooth menu) so the session does not start silent.
    StandardLoadingBar_Update("🔊 Checking audio…")
    Act_EnsurePersonalAudio()
    Run "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Chrome Apps\Settle Up.lnk"
    Run "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\apps do Chrome\WhatsApp Web.lnk"
    Run "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Chrome Apps\Gmail.lnk"
}

; Guarantee Clip Angel is open and usable (soft activate, else Macros [R] hard restart).
StandardLoadingBar_Update("⏳ Ensuring Clip Angel...")
clipAngelOk := ClipAngel_EnsureRunning()

; Open AI companion (same chord as Gemini/gemini_open.ahk #!+i).
StandardLoadingBar_Update("⏳ Opening AI companion...")
Sleep 3000  ; Gemini.ahk must finish auto-exec so #!+i is registered
Send "#!+i"

habitsFolder := notesFolder . "\habits"
if (clipAngelOk)
    StandardLoadingBar_Update("✅ Done", BANNER_ACCENT_SUCCESS)
StandardLoadingBar_Hide(clipAngelOk ? 500 : 0)
Sleep 1000
; Exit so Act never remains as a lingering Utils/hotkey host after bootstrap.
ExitApp

; Stash, fetch, pull --ff-only, stash pop. Same gates as editor Alt+S, hidden.
; Each git step inside the script is capped at 900s. Outer cap covers stash, fetch,
; pull, one fetch/pull retry, and stash pop.
Act_PersonalRepoStashAndPull(personalFolder) {
    flowPs1 := A_ScriptDir "\infra\tools\Editor-GitStashFetchPull.ps1"
    if !FileExist(flowPs1) {
        StandardLoadingBar_Hide(0)
        MsgBox("Editor-GitStashFetchPull.ps1 not found:`n`n" flowPs1, "Act automation", "Icon!")
        ExitApp
    }
    resultPath := A_Temp "\act-personal-repo-" A_TickCount ".json"
    try FileDelete(resultPath)
    catch {
    }
    ; RunWaitWithTimeout wraps the command in powershell -Command "...".
    ; Double quotes in that string are eaten, so a -File path with spaces never starts
    ; and PowerShell does exit $null, which AutoHotkey reports as -196608.
    ; -EncodedCommand carries the script without quotes on the command line.
    ps := Act_PersonalRepoPullScript(flowPs1, personalFolder, resultPath, 900, 5500000)
    encoded := Act_PowerShellEncodedCommand(ps)
    if (encoded = "") {
        StandardLoadingBar_Hide(0)
        MsgBox("Could not encode the personal repo stash/pull command.", "Act automation", "Icon!")
        ExitApp
    }
    exitCode := 1
    try exitCode := RunWait("powershell.exe -NoProfile -ExecutionPolicy Bypass -EncodedCommand " encoded,
        personalFolder, "Hide")
    catch {
        exitCode := 1
    }
    detail := Act_PersonalRepoPullError(resultPath)
    try FileDelete(resultPath)
    catch {
    }
    if (exitCode = 0)
        return
    StandardLoadingBar_Hide(0)
    if (exitCode = 124) {
        MsgBox(
            "Timed out while stashing and pulling the personal repo.`n`nIn:`n" personalFolder "`n`n"
            . "This usually means git is waiting for input (credentials/2FA) or is blocked by network/proxy.",
            "Act automation", "Icon!")
        ExitApp
    }
    MsgBox(
        "Personal repo stash/pull failed (exit code " exitCode ").`n`n"
        . (detail != "" ? detail "`n`n" : "")
        . "In:`n" personalFolder "`n`n"
        . "Run git status in that folder to see the error.",
        "Act automation", "Icon!")
    ExitApp
}

; Single-quoted PowerShell string. Paths here have no apostrophes; still escape any.
Act_PsQuote(s) {
    return "'" StrReplace(s, "'", "''") "'"
}

; Hidden powershell -File of the stash/pull script.
; ArgumentList must be one string. The array form splits paths on spaces, the child
; powershell exits -196608, and no result file is written.
; Second WaitForExit fills ExitCode (the timed overload can leave it null, and exit $null is -196608).
Act_PersonalRepoPullScript(flowPs1, personalFolder, resultPath, timeoutSec, outerTimeoutMs) {
    arg := "-NoProfile -ExecutionPolicy Bypass -File `"" flowPs1 "`" -RepoDir `"" personalFolder
        . "`" -ResultPath `"" resultPath "`" -TimeoutSec " timeoutSec
    return "$ErrorActionPreference='Stop';"
    . "$arg=" Act_PsQuote(arg) ";"
    . "$p=Start-Process -FilePath 'powershell.exe' -ArgumentList `$arg -WorkingDirectory " Act_PsQuote(personalFolder)
    . " -PassThru -WindowStyle Hidden;"
    . "if(-not $p.WaitForExit([int]" outerTimeoutMs ")){try{$p.Kill()}catch{}; exit 124};"
    . "$p.WaitForExit()|Out-Null;"
    . "if($null -eq $p.ExitCode){exit 1};"
    . "exit $p.ExitCode"
}

; UTF-16LE base64 for powershell -EncodedCommand. AHK "UTF-16" is little-endian.
; Drops the trailing null.
Act_PowerShellEncodedCommand(script) {
    bytes := StrPut(script, "UTF-16")
    if (bytes <= 2)
        return ""
    buf := Buffer(bytes)
    StrPut(script, buf, "UTF-16")
    payload := bytes - 2
    flags := 0x40000001  ; CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF
    chars := 0
    if !DllCall("crypt32\CryptBinaryToStringW", "Ptr", buf.Ptr, "UInt", payload, "UInt", flags, "Ptr", 0, "UInt*", &
        chars)
        return ""
    out := Buffer(chars * 2)
    if !DllCall("crypt32\CryptBinaryToStringW", "Ptr", buf.Ptr, "UInt", payload, "UInt", flags, "Ptr", out.Ptr, "UInt*", &
        chars)
        return ""
    return Trim(StrGet(out, "UTF-16"), "`r`n `t")
}

Act_PersonalRepoPullError(resultPath) {
    if !FileExist(resultPath)
        return ""
    raw := ""
    try raw := FileRead(resultPath, "UTF-8")
    catch {
        return ""
    }
    failed := Act_JsonStringField(raw, "failedStep")
    err := Act_JsonStringField(raw, "error")
    if (failed != "" && err != "")
        return failed ": " err
    if (err != "")
        return err
    return failed
}

; Personal PC endpoints from the Audio / Bluetooth menu (💻):
;   In  Grupo de microfones (2- Realtek(R) Audio)
;   Out Alto-falantes (2- Realtek(R) Audio)
; Ids are the stable Core Audio ids; names still match if Windows recreates the id.
Act_EnsurePersonalAudio() {
    global IS_WORK_ENVIRONMENT
    if (IsSet(IS_WORK_ENVIRONMENT) && IS_WORK_ENVIRONMENT)
        return
    ps1 := A_ScriptDir "\infra\tools\AudioBt.ps1"
    if !FileExist(ps1)
        return
    listed := Act_AudioBtRun(ps1, "list", "")
    if (!listed.ok) {
        StandardLoadingBar_Update("⚠ Could not read audio devices", BANNER_ACCENT_ERROR)
        Sleep 800
        return
    }
    rows := Act_AudioParseList(listed.text)
    fixed := ""
    missed := ""
    for kind in ["In", "Out"] {
        current := Act_AudioFindDefault(rows, kind)
        if Act_AudioDefaultIsLive(rows, current)
            continue
        pc := Act_AudioFindComputerRealtek(rows, kind)
        if !IsObject(pc) {
            missed .= (missed = "" ? "" : " and ") . (kind = "In" ? "mic" : "speakers")
            continue
        }
        if (pc.isDefault && !Act_AudioStateIsDead(pc.state))
            continue
        set := Act_AudioBtRun(ps1, "default", pc.id)
        if (set.ok)
            fixed .= (fixed = "" ? "" : " and ") . (kind = "In" ? "mic" : "speakers")
        else
            missed .= (missed = "" ? "" : " and ") . (kind = "In" ? "mic" : "speakers")
    }
    if (fixed != "")
        StandardLoadingBar_Update("🔊 Computer " . fixed . " (Realtek)")
    if (missed != "") {
        StandardLoadingBar_Update("⚠ Could not set computer " . missed, BANNER_ACCENT_ERROR)
        Sleep 800
    }
}

Act_AudioBtRun(ps1, action, id) {
    outFile := A_Temp "\act-audiobt-" A_TickCount "-" action ".txt"
    try FileDelete(outFile)
    catch {
    }
    cmd := 'powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "' ps1 '" -Action ' action
    if (id != "")
        cmd .= ' -Id "' id '"'
    cmd .= ' -OutFile "' outFile '"'
    timeout := (action = "list") ? 90000 : 30000
    exitCode := RunWaitWithTimeout(cmd, A_ScriptDir, "Hide", timeout)
    text := ""
    try {
        if FileExist(outFile)
            text := Trim(FileRead(outFile, "UTF-8"), "`r`n")
    } catch {
    }
    try FileDelete(outFile)
    catch {
    }
    ok := (exitCode = 0 && text != "" && SubStr(text, 1, 4) != "ERR`t")
    return { ok: ok, text: text }
}

Act_AudioParseList(text) {
    rows := []
    if (text = "" || SubStr(text, 1, 4) = "ERR`t")
        return rows
    for line in StrSplit(text, "`n", "`r") {
        line := Trim(line)
        if (line = "")
            continue
        parts := StrSplit(line, "`t")
        if (parts.Length < 6)
            continue
        if (A_Index = 1 && StrLower(parts[1]) = "id")
            continue
        rows.Push({
            id: parts[1],
            kind: parts[2],
            name: parts[3],
            state: parts[4],
            isDefault: parts[5] = "1",
            canConnect: parts[6] = "1"
        })
    }
    return rows
}

Act_AudioFindDefault(rows, kind) {
    for row in rows {
        if (row.kind = kind && row.isDefault)
            return row
    }
    return ""
}

; A default that is still the Windows default while its radio is down counts as empty.
Act_AudioDefaultIsLive(rows, row) {
    if !IsObject(row) || !row.isDefault
        return false
    if Act_AudioStateIsDead(row.state)
        return false
    if (row.canConnect) {
        bt := Act_AudioBtForEndpoint(rows, row)
        if !IsObject(bt) || !Act_AudioBtIsConnected(bt.state)
            return false
    }
    return true
}

Act_AudioStateIsDead(state) {
    for dead in ["Disabled", "Unplugged", "Disconnected", "Not present"] {
        if (state = dead || InStr(state, dead))
            return true
    }
    return false
}

Act_AudioBtIsConnected(state) {
    if InStr(state, "Disconnected")
        return false
    return InStr(state, "Connected")
}

Act_AudioBtForEndpoint(rows, ep) {
    best := ""
    bestLen := 0
    for row in rows {
        if (row.kind != "BT" || row.name = "")
            continue
        if InStr(ep.name, row.name) && StrLen(row.name) > bestLen {
            best := row
            bestLen := StrLen(row.name)
        }
    }
    return best
}

Act_AudioFindComputerRealtek(rows, kind) {
    wantId := (kind = "In")
        ? "{0.0.1.00000000}.{9236f67f-79b6-4c55-a573-df6e0f8399c3}"
        : "{0.0.0.00000000}.{d2cda8f3-6330-4d0b-b560-d788366b7240}"
    idHit := ""
    nameHit := ""
    for row in rows {
        if (row.kind != kind)
            continue
        if (StrLower(row.id) = StrLower(wantId))
            idHit := row
        else if Act_AudioNameIsComputerRealtek(row)
            nameHit := row
    }
    if (IsObject(idHit) && !Act_AudioStateIsDead(idHit.state))
        return idHit
    if (IsObject(nameHit) && !Act_AudioStateIsDead(nameHit.state))
        return nameHit
    if IsObject(idHit)
        return idHit
    return nameHit
}

Act_AudioNameIsComputerRealtek(row) {
    if !InStr(row.name, "Realtek")
        return false
    if (row.kind = "In")
        return InStr(row.name, "Grupo de microfones")
    if (row.kind = "Out")
        return InStr(row.name, "Alto-falantes")
    return false
}

Act_JsonStringField(raw, key) {
    needle := '"' key '":"'
    pos := InStr(raw, needle)
    if !pos {
        needle := '"' key '": "'
        pos := InStr(raw, needle)
    }
    if !pos
        return ""
    i := pos + StrLen(needle)
    out := ""
    while (i <= StrLen(raw)) {
        ch := SubStr(raw, i, 1)
        if (ch = "\") {
            n := SubStr(raw, i + 1, 1)
            if (n = "n")
                out .= "`n"
            else if (n = "r")
                out .= "`r"
            else if (n = "t")
                out .= "`t"
            else
                out .= n
            i += 2
            continue
        }
        if (ch = '"')
            break
        out .= ch
        i += 1
    }
    return out
}
