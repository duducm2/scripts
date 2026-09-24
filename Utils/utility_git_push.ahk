; =============================================================================
; Utils module: utility_git_push.ahk
; Commit and push scripts + notes (+ personal when present) from Utility [G]
; Exports personal → main/punctual.md and habits → main/habits.md (never work);
; syncs Palace MD when mnemonics/data is dirty.
; Runs in the background; no mid-run loading bar — final 3s banner on all monitors.
; =============================================================================

global g_UtilityGitPushBusy := false
global UTILITY_GIT_RESULT_BANNER_MS := 3000
global UTILITY_GIT_PUSH_WATCHDOG_MS := 180000
global g_UtilityGitPushResultMsg := ""
global g_UtilityGitPushResultAccent := ""
global g_UtilityGitPushWatchdogArmed := false

Utility_GitChimeStart() {
    try ScriptSoundPlay(A_ScriptDir . "\assets\sounds\commit-start.wav")
    catch {
    }
}

Utility_GitChimeEnd(ok := true) {
    path := ok
        ? (A_ScriptDir . "\assets\sounds\pull-successful.wav")
            : (A_ScriptDir . "\assets\sounds\quick-update-failure.wav")
    try ScriptSoundPlay(path)
    catch {
    }
}

; Final status on every monitor for exactly UTILITY_GIT_RESULT_BANNER_MS (default 3s).
Utility_GitNotify(msg, ms := 0, accent := "") {
    global UTILITY_GIT_RESULT_BANNER_MS
    if (ms <= 0)
        ms := UTILITY_GIT_RESULT_BANNER_MS
    if (accent = "")
        accent := BANNER_ACCENT_INFO
    shown := false
    try shown := StandardLoadingBar_ShowPassiveAllMonitors(msg, ms, accent)
    catch {
        shown := false
    }
    if (!shown) {
        ; Fallback: single-monitor overlay, then tray tip.
        try {
            ShowCenteredOverlay_Utils(msg, ms, accent)
            shown := true
        } catch {
        }
    }
    if (!shown)
        TrayTip("Git push", msg)
}

Utility_GitFirstErrorLine(r) {
    text := ""
    if (IsObject(r)) {
        if (r.HasProp("stderr") && r.stderr != "")
            text := r.stderr
        else if (r.HasProp("stdout") && r.stdout != "")
            text := r.stdout
    }
    if (text = "")
        return "unknown error"
    text := StrReplace(StrReplace(text, "`r`n", "`n"), "`r", "`n")
    for line in StrSplit(text, "`n") {
        t := Trim(line)
        if (t != "")
            return SubStr(t, 1, 120)
    }
    return SubStr(Trim(text), 1, 120)
}

Utility_GitFormatPushError(r) {
    line := Utility_GitFirstErrorLine(r)
    lower := StrLower(line)
    if (InStr(lower, "non-fast-forward") || InStr(lower, "fetch first") || InStr(lower, "rejected"))
        return line . " — pull first (Alt+S or Act)"
    if (r.HasProp("exitCode") && r.exitCode = 124)
        return "timed out — check network or run git push in a terminal"
    if (InStr(lower, "authentication") || InStr(lower, "permission") || InStr(lower, "could not read"))
        return line . " — sign in to git in a terminal"
    return line
}

Utility_GitPassiveBar(msg) {
    ; Used by discard (Main Repos); push path intentionally does not call this.
    try StandardLoadingBar_Update(msg)
    catch {
        try StandardLoadingBar_Show(msg, BANNER_ACCENT_INTERMEDIATE)
        catch {
        }
    }
}

; True if porcelain status mentions a path under prefix (forward or backslash).
Utility_GitStatusHasPathPrefix(porcelain, prefixFwd) {
    pref := StrLower(StrReplace(prefixFwd, "\", "/"))
    prefAlt := StrReplace(pref, "/", "\")
    for line in StrSplit(porcelain, "`n", "`r") {
        t := Trim(line)
        if (t = "")
            continue
        ; XY<space>path  or  XY path / rename with " -> "
        low := StrLower(t)
        if (InStr(low, pref) || InStr(low, prefAlt))
            return true
    }
    return false
}

Utility_GitExportPhoneTasksMd(scriptsRoot, notesRoot) {
    ; Phone mirror: personal + habits only (work stays local CSV).
    py := scriptsRoot . "\tasks\python\export_to_md.py"
    if (!FileExist(py))
        return "error:export_to_md.py not found"
    pyCmd := ""
    try pyCmd := Task_FindPythonCmd()
    catch {
        pyCmd := ""
    }
    if (pyCmd = "")
        return "error:Python not found for Tasks MD export"
    dataDir := scriptsRoot . "\tasks\data"
    punctual := notesRoot . "\main\punctual.md"
    habits := notesRoot . "\main\habits.md"
    cmd := pyCmd . ' "' . py . '" --data-dir "' . dataDir
        . '" --punctual "' . punctual . '" --habits "' . habits . '"'
    ; Temp .cmd avoids RunWaitWithTimeout PowerShell -Command breaking on quoted
    ; paths that contain spaces (e.g. "Meu Drive").
    bat := A_Temp . "\utility-git-export-" . A_TickCount . ".cmd"
    try FileDelete(bat)
    catch {
    }
    try FileAppend("@echo off`r`n" . cmd . "`r`n", bat, "CP0")
    catch as e {
        return "error:Tasks MD export failed: " . e.Message
    }
    exitCode := 0
    try {
        exitCode := RunWaitWithTimeout('"' . bat . '"', scriptsRoot, "Hide", 60000)
    } catch as e {
        try FileDelete(bat)
        catch {
        }
        return "error:Tasks MD export failed: " . e.Message
    }
    try FileDelete(bat)
    catch {
    }
    if (exitCode = 124)
        return "error:Tasks MD export timed out"
    if (exitCode != 0)
        return "error:Tasks MD export failed (exit " . exitCode . ")"
    return "ok"
}

; Backward-compatible alias (#!+9 / Utility [G] path).
Utility_GitExportPunctualMd(scriptsRoot, notesRoot) {
    return Utility_GitExportPhoneTasksMd(scriptsRoot, notesRoot)
}

Utility_GitPrepareExports(scriptsRoot, notesRoot) {
    status := GitCli_Run(scriptsRoot, "status --porcelain", 30000)
    if (status.exitCode != 0)
        return "error:Scripts status failed: " . Utility_GitFirstErrorLine(status)

    export := Utility_GitExportPhoneTasksMd(scriptsRoot, notesRoot)
    if (SubStr(export, 1, 6) = "error:")
        return export

    porcelain := status.stdout
    needPalace := Utility_GitStatusHasPathPrefix(porcelain, "mnemonics/data/")

    if (needPalace) {
        ; Soft-fail: Drive-locked prune must not abort scripts/notes push
        try Palace_SyncAllPracticeMd(false)
        catch {
        }
        try Palace_SyncAllPlansMd(false)
        catch {
        }
    }
    return "ok"
}

; Returns "pushed", "noop", or "error:…"
Utility_GitSyncPushOne(repoDir, label, commitMsg) {
    repo := GitCli_RevParseTopLevel(repoDir)
    if (repo = "")
        return "error:" . label . " not a git repository"

    status := GitCli_Run(repo, "status --porcelain", 30000)
    if (status.exitCode != 0)
        return "error:" . label . " status failed: " . Utility_GitFirstErrorLine(status)

    changed := 0
    for line in StrSplit(status.stdout, "`n", "`r") {
        if (Trim(line) != "")
            changed += 1
    }
    if (changed = 0)
        return "noop"

    add := GitCli_Run(repo, "add -A", 60000)
    if (add.exitCode != 0)
        return "error:" . label . " add failed: " . Utility_GitFirstErrorLine(add)

    msgFile := A_Temp . "\utility-git-msg-" . A_TickCount . "-" . label . ".txt"
    try FileDelete(msgFile)
    catch {
    }
    try FileAppend(commitMsg, msgFile, "UTF-8")
    catch as e {
        return "error:" . label . " could not write commit message: " . e.Message
    }
    commit := GitCli_Run(repo, 'commit -F "' . StrReplace(msgFile, '"', '') . '"', 120000)
    try FileDelete(msgFile)
    catch {
    }
    if (commit.exitCode != 0) {
        err := Utility_GitFirstErrorLine(commit)
        if (InStr(StrLower(err), "nothing to commit"))
            return "noop"
        return "error:" . label . " commit failed: " . err
    }

    push := GitCli_Run(repo, "push", 120000)
    if (push.exitCode != 0) {
        branch := GitCli_CaptureStdout(repo, "branch --show-current", 15000)
        if (branch != "") {
            pushUp := GitCli_Run(repo, 'push -u origin "' . StrReplace(branch, '"', '') . '"', 120000)
            if (pushUp.exitCode != 0)
                return "error:" . label . " push failed: " . Utility_GitFormatPushError(pushUp)
        } else {
            return "error:" . label . " push failed: " . Utility_GitFormatPushError(push)
        }
    }
    return "pushed"
}

Utility_GitPushCancelWatchdog() {
    global g_UtilityGitPushWatchdogArmed
    try SetTimer(Utility_GitPushWatchdog, 0)
    catch {
    }
    g_UtilityGitPushWatchdogArmed := false
}

Utility_GitPushArmWatchdog() {
    global g_UtilityGitPushWatchdogArmed, UTILITY_GIT_PUSH_WATCHDOG_MS
    Utility_GitPushCancelWatchdog()
    SetTimer(Utility_GitPushWatchdog, -UTILITY_GIT_PUSH_WATCHDOG_MS)
    g_UtilityGitPushWatchdogArmed := true
}

; If worker still busy after ~180s, force an explicit timeout result (never silent).
Utility_GitPushWatchdog(*) {
    global g_UtilityGitPushBusy, g_UtilityGitPushWatchdogArmed
    global g_UtilityGitPushResultMsg, g_UtilityGitPushResultAccent
    g_UtilityGitPushWatchdogArmed := false
    if (!g_UtilityGitPushBusy)
        return
    g_UtilityGitPushBusy := false
    g_UtilityGitPushResultMsg := "❌ Push timed out — check network or git"
    g_UtilityGitPushResultAccent := BANNER_ACCENT_ERROR
    SetTimer(Utility_GitSyncPushShowResult, -50)
}

; Fresh-timer delivery so GUIs paint after a long blocking worker.
Utility_GitSyncPushShowResult(*) {
    global g_UtilityGitPushResultMsg, g_UtilityGitPushResultAccent
    msg := g_UtilityGitPushResultMsg
    accent := g_UtilityGitPushResultAccent
    g_UtilityGitPushResultMsg := ""
    g_UtilityGitPushResultAccent := ""
    if (msg = "")
        return
    if (accent = "")
        accent := BANNER_ACCENT_INFO
    ok := (accent != BANNER_ACCENT_ERROR)
    Utility_GitChimeEnd(ok)
    Utility_GitNotify(msg, , accent)
}

Utility_GitSyncPushScheduleResult(msg, accent) {
    global g_UtilityGitPushResultMsg, g_UtilityGitPushResultAccent
    if (msg = "")
        return
    g_UtilityGitPushResultMsg := msg
    g_UtilityGitPushResultAccent := accent
    try SetTimer(Utility_GitSyncPushShowResult, 0)
    catch {
    }
    SetTimer(Utility_GitSyncPushShowResult, -50)
}

; Entry from Utility Shortcuts [G] — arms background worker immediately.
Utility_GitSyncPush() {
    global g_UtilityGitPushBusy
    if (g_UtilityGitPushBusy) {
        Utility_GitNotify("ℹ Push already running", , BANNER_ACCENT_INFO)
        return
    }
    g_UtilityGitPushBusy := true
    Utility_GitChimeStart()
    Utility_GitPushArmWatchdog()
    SetTimer(Utility_GitSyncPushWorker, -1)
}

Utility_GitSyncPushWorker() {
    global g_UtilityGitPushBusy
    resultMsg := ""
    resultAccent := BANNER_ACCENT_INFO
    try {
        ; No mid-run loading bar — only the final all-monitors result banner below.

        commitMsg := Format("{:04d}-{:02d}-{:02d} {:02d}:{:02d}:{:02d}",
            A_YYYY, A_MM, A_DD, A_Hour, A_Min, A_Sec)

        scriptsRoot := GitCli_RevParseTopLevel(A_ScriptDir)
        if (scriptsRoot = "") {
            resultMsg := "❌ Scripts not a git repository"
            resultAccent := BANNER_ACCENT_ERROR
            return
        }

        notesRoot := ""
        try notesRoot := GetNotesRepoPath()
        catch {
            notesRoot := ""
        }
        if (notesRoot = "" || !DirExist(notesRoot)) {
            resultMsg := "❌ Notes repo folder not found"
            resultAccent := BANNER_ACCENT_ERROR
            return
        }

        prep := Utility_GitPrepareExports(scriptsRoot, notesRoot)
        if (SubStr(prep, 1, 6) = "error:") {
            resultMsg := "❌ " . SubStr(prep, 7)
            resultAccent := BANNER_ACCENT_ERROR
            return
        }

        scriptsResult := Utility_GitSyncPushOne(scriptsRoot, "Scripts", commitMsg)
        if (SubStr(scriptsResult, 1, 6) = "error:") {
            resultMsg := "❌ " . SubStr(scriptsResult, 7)
            resultAccent := BANNER_ACCENT_ERROR
            return
        }

        notesResult := Utility_GitSyncPushOne(notesRoot, "Notes", commitMsg)
        if (SubStr(notesResult, 1, 6) = "error:") {
            resultMsg := "❌ " . SubStr(notesResult, 7)
            resultAccent := BANNER_ACCENT_ERROR
            return
        }

        ; Soft-skip when PERSONAL_REPO_PATH missing (typical on work PC).
        personalRoot := ""
        try personalRoot := GetPersonalRepoPath()
        catch {
            personalRoot := ""
        }
        personalResult := "noop"
        if (personalRoot != "") {
            personalResult := Utility_GitSyncPushOne(personalRoot, "Personal", commitMsg)
            if (SubStr(personalResult, 1, 6) = "error:") {
                resultMsg := "❌ " . SubStr(personalResult, 7)
                resultAccent := BANNER_ACCENT_ERROR
                return
            }
        }

        if (scriptsResult = "noop" && notesResult = "noop" && personalResult = "noop") {
            resultMsg := "ℹ Nothing to commit"
            resultAccent := BANNER_ACCENT_INFO
            return
        }

        parts := []
        if (scriptsResult = "pushed")
            parts.Push("Scripts")
        if (notesResult = "pushed")
            parts.Push("Notes")
        if (personalResult = "pushed")
            parts.Push("Personal")
        if (parts.Length = 1)
            resultMsg := "✅ " . parts[1] . " pushed"
        else if (parts.Length = 2)
            resultMsg := "✅ " . parts[1] . " + " . parts[2] . " pushed"
        else
            resultMsg := "✅ " . parts[1] . " + " . parts[2] . " + " . parts[3] . " pushed"
        resultAccent := BANNER_ACCENT_SUCCESS
    } catch as e {
        resultMsg := "❌ Push failed: " . e.Message
        resultAccent := BANNER_ACCENT_ERROR
    } finally {
        Utility_GitPushCancelWatchdog()
        g_UtilityGitPushBusy := false
        if (resultMsg != "")
            Utility_GitSyncPushScheduleResult(resultMsg, resultAccent)
    }
}
