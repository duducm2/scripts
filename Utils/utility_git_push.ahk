; =============================================================================
; Utils module: utility_git_push.ahk
; Commit and push scripts + notes repos from Utility Shortcuts top-level [G]
; =============================================================================

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

; Returns "pushed", "noop", or "error:…"
Utility_GitSyncPushOne(repoDir, label, commitMsg) {
    prefix := label . ": "
    try StandardLoadingBar_Update("⏳ " . prefix . "Checking git status…")
    catch {
        try StandardLoadingBar_Show("⏳ " . prefix . "Checking git status…", BANNER_ACCENT_INTERMEDIATE)
        catch {
        }
    }

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

    try StandardLoadingBar_Update("⏳ " . prefix . "Staging changes…")
    catch {
    }
    add := GitCli_Run(repo, "add -A", 60000)
    if (add.exitCode != 0)
        return "error:" . label . " add failed: " . Utility_GitFirstErrorLine(add)

    try StandardLoadingBar_Update("⏳ " . prefix . "Committing…")
    catch {
    }
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

    try StandardLoadingBar_Update("⏳ " . prefix . "Pushing to remote…")
    catch {
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

Utility_GitSyncPush() {
    try {
        try StandardLoadingBar_Show("⏳ Checking git status…", BANNER_ACCENT_INTERMEDIATE)
        catch {
        }

        commitMsg := Format("{:04d}-{:02d}-{:02d} {:02d}:{:02d}:{:02d}",
            A_YYYY, A_MM, A_DD, A_Hour, A_Min, A_Sec)

        scriptsRoot := GitCli_RevParseTopLevel(A_ScriptDir)
        if (scriptsRoot = "") {
            Utility_GitNotify("❌ Scripts not a git repository", 2500, BANNER_ACCENT_ERROR)
            return
        }

        notesRoot := ""
        try notesRoot := GetNotesRepoPath()
        catch {
            notesRoot := ""
        }
        if (notesRoot = "" || !DirExist(notesRoot)) {
            Utility_GitNotify("❌ Notes repo folder not found", 2500, BANNER_ACCENT_ERROR)
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
            Utility_GitNotify("❌ " . errors[1], 4000, BANNER_ACCENT_ERROR)
            return
        }

        if (scriptsResult = "noop" && notesResult = "noop") {
            Utility_GitNotify("ℹ Nothing to commit", 1800, BANNER_ACCENT_INFO)
            return
        }

        parts := []
        if (scriptsResult = "pushed")
            parts.Push("Scripts")
        if (notesResult = "pushed")
            parts.Push("Notes")
        if (parts.Length = 2)
            Utility_GitNotify("✅ Scripts + Notes pushed", 2500, BANNER_ACCENT_SUCCESS)
        else
            Utility_GitNotify("✅ " . parts[1] . " pushed", 2500, BANNER_ACCENT_SUCCESS)
    } catch as e {
        Utility_GitNotify("❌ Push failed: " . e.Message, 3500, BANNER_ACCENT_ERROR)
    } finally {
        try StandardLoadingBar_Hide(0)
        catch {
        }
    }
}
