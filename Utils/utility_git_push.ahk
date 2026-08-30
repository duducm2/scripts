; =============================================================================
; Utils module: utility_git_push.ahk
; Commit and push scripts + notes repos from Utility Shortcuts top-level [G]
; Writes empty main/punctual.md stub in notes; syncs Palace MD when mnemonics/data is dirty.
; Runs in the background so the UI stays usable.
; =============================================================================

global g_UtilityGitPushBusy := false

Utility_GitNotify(msg, ms := 1800, accent := "") {
    if (accent = "")
        accent := BANNER_ACCENT_INFO
    try ShowCenteredOverlay_Utils(msg, ms, accent)
    catch {
        TrayTip("Git push", msg)
    }
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
    ; Milestone updates on the active Loading Indication (animated bar).
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

Utility_GitWritePunctualMdStub(notesRoot) {
    if (notesRoot = "" || !DirExist(notesRoot))
        return "error:Notes repo folder not found"
    mainDir := RTrim(notesRoot, "\") . "\main"
    path := mainDir . "\punctual.md"
    try {
        if (!DirExist(mainDir))
            DirCreate(mainDir)
    } catch as e {
        return "error:Could not create main folder: " . e.Message
    }
    try FileDelete(path)
    catch {
    }
    try FileAppend("", path, "UTF-8")
    catch as e {
        return "error:Could not write punctual.md stub: " . e.Message
    }
    return "ok"
}

Utility_GitPrepareExports(scriptsRoot, notesRoot) {
    status := GitCli_Run(scriptsRoot, "status --porcelain", 30000)
    if (status.exitCode != 0)
        return "error:Scripts status failed: " . Utility_GitFirstErrorLine(status)

    Utility_GitPassiveBar("⏳ Ensuring punctual.md stub…")
    stub := Utility_GitWritePunctualMdStub(notesRoot)
    if (SubStr(stub, 1, 6) = "error:")
        return stub

    porcelain := status.stdout
    needPalace := Utility_GitStatusHasPathPrefix(porcelain, "mnemonics/data/")

    if (needPalace) {
        Utility_GitPassiveBar("⏳ Syncing Memory Palace Markdown…")
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
    prefix := label . ": "
    Utility_GitPassiveBar("⏳ " . prefix . "Checking git status…")

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

    Utility_GitPassiveBar("⏳ " . prefix . "Staging changes…")
    add := GitCli_Run(repo, "add -A", 60000)
    if (add.exitCode != 0)
        return "error:" . label . " add failed: " . Utility_GitFirstErrorLine(add)

    Utility_GitPassiveBar("⏳ " . prefix . "Committing…")
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

    Utility_GitPassiveBar("⏳ " . prefix . "Pushing to remote…")
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

; Entry from Utility Shortcuts [G] — arms background worker immediately.
Utility_GitSyncPush() {
    global g_UtilityGitPushBusy
    if (g_UtilityGitPushBusy) {
        Utility_GitNotify("ℹ Push already running", 1800, BANNER_ACCENT_INFO)
        return
    }
    g_UtilityGitPushBusy := true
    SetTimer(Utility_GitSyncPushWorker, -1)
}

Utility_GitSyncPushWorker() {
    global g_UtilityGitPushBusy
    resultMsg := ""
    resultAccent := BANNER_ACCENT_INFO
    try {
        ; Loading Indication for the whole push (animated bar)
        try StandardLoadingBar_Show("⏳ Preparing push…", BANNER_ACCENT_INTERMEDIATE)
        catch {
        }

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

        Utility_GitPassiveBar("⏳ Exporting Markdown if needed…")
        prep := Utility_GitPrepareExports(scriptsRoot, notesRoot)
        if (SubStr(prep, 1, 6) = "error:") {
            resultMsg := "❌ " . SubStr(prep, 7)
            resultAccent := BANNER_ACCENT_ERROR
            return
        }

        scriptsResult := Utility_GitSyncPushOne(scriptsRoot, "Scripts", commitMsg)
        notesResult := Utility_GitSyncPushOne(notesRoot, "Notes", commitMsg)

        errors := []
        if (SubStr(scriptsResult, 1, 6) = "error:")
            errors.Push(SubStr(scriptsResult, 7))
        if (SubStr(notesResult, 1, 6) = "error:")
            errors.Push(SubStr(notesResult, 7))
        if (errors.Length > 0) {
            resultMsg := "❌ " . errors[1]
            resultAccent := BANNER_ACCENT_ERROR
            return
        }

        if (scriptsResult = "noop" && notesResult = "noop") {
            resultMsg := "ℹ Nothing to commit"
            resultAccent := BANNER_ACCENT_INFO
            return
        }

        parts := []
        if (scriptsResult = "pushed")
            parts.Push("Scripts")
        if (notesResult = "pushed")
            parts.Push("Notes")
        if (parts.Length = 2)
            resultMsg := "✅ Scripts + Notes pushed"
        else
            resultMsg := "✅ " . parts[1] . " pushed"
        resultAccent := BANNER_ACCENT_SUCCESS
    } catch as e {
        resultMsg := "❌ Push failed: " . e.Message
        resultAccent := BANNER_ACCENT_ERROR
    } finally {
        g_UtilityGitPushBusy := false
        try StandardLoadingBar_Hide(0)
        catch {
        }
        ; Information Only AFTER Hide so Hide does not wipe the result toast
        if (resultMsg != "")
            Utility_GitNotify(resultMsg, 2800, resultAccent)
    }
}
