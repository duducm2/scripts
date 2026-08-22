; =============================================================================
; Utils module: finance_git_sync.ahk
; Commit and push the scripts repo from Finance main menu [P]
; =============================================================================

Finance_GitFirstErrorLine(r) {
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

Finance_GitFormatPushError(r) {
    line := Finance_GitFirstErrorLine(r)
    lower := StrLower(line)
    if (InStr(lower, "non-fast-forward") || InStr(lower, "fetch first") || InStr(lower, "rejected"))
        return line . " — pull first (Alt+S or Act)"
    if (r.HasProp("exitCode") && r.exitCode = 124)
        return "timed out — check network or run git push in a terminal"
    if (InStr(lower, "authentication") || InStr(lower, "permission") || InStr(lower, "could not read"))
        return line . " — sign in to git in a terminal"
    return line
}

Finance_GitSyncPush() {
    Finance_CloseGui()
    repo := ""
    success := false
    try {
        try StandardLoadingBar_Show("⏳ Checking git status…", BANNER_ACCENT_INTERMEDIATE)
        catch {
        }

        repo := GitCli_RevParseTopLevel(A_ScriptDir)
        if (repo = "") {
            Finance_Notify("❌ Not a git repository", 2500, BANNER_ACCENT_ERROR)
            return
        }

        status := GitCli_Run(repo, "status --porcelain", 30000)
        if (status.exitCode != 0) {
            Finance_Notify("❌ Git status failed: " . Finance_GitFirstErrorLine(status), 3500, BANNER_ACCENT_ERROR)
            return
        }

        changed := 0
        for line in StrSplit(status.stdout, "`n", "`r") {
            if (Trim(line) != "")
                changed += 1
        }
        if (changed = 0) {
            Finance_Notify("ℹ Nothing to commit", 1800, BANNER_ACCENT_INFO)
            return
        }

        commitMsg := "Finance update "
            . Format("{:04d}-{:02d}-{:02d} {:02d}:{:02d}:{:02d}", A_YYYY, A_MM, A_DD, A_Hour, A_Min, A_Sec)

        try StandardLoadingBar_Update("⏳ Staging changes…")
        catch {
            try StandardLoadingBar_Show("⏳ Staging changes…", BANNER_ACCENT_INTERMEDIATE)
            catch {
            }
        }

        add := GitCli_Run(repo, "add -A", 60000)
        if (add.exitCode != 0) {
            Finance_Notify("❌ Git add failed: " . Finance_GitFirstErrorLine(add), 3500, BANNER_ACCENT_ERROR)
            return
        }

        try StandardLoadingBar_Update("⏳ Committing…")
        catch {
        }
        msgFile := A_Temp . "\finance-git-msg-" . A_TickCount . ".txt"
        try FileDelete(msgFile)
        catch {
        }
        try FileAppend(commitMsg, msgFile, "UTF-8")
        catch as e {
            Finance_Notify("❌ Could not write commit message: " . e.Message, 3000, BANNER_ACCENT_ERROR)
            return
        }
        commit := GitCli_Run(repo, 'commit -F "' . StrReplace(msgFile, '"', '') . '"', 120000)
        try FileDelete(msgFile)
        catch {
        }
        if (commit.exitCode != 0) {
            err := Finance_GitFirstErrorLine(commit)
            if (InStr(StrLower(err), "nothing to commit"))
                Finance_Notify("ℹ Nothing to commit", 1800, BANNER_ACCENT_INFO)
            else
                Finance_Notify("❌ Commit failed: " . err, 3500, BANNER_ACCENT_ERROR)
            return
        }

        try StandardLoadingBar_Update("⏳ Pushing to remote…")
        catch {
        }
        push := GitCli_Run(repo, "push", 120000)
        if (push.exitCode != 0) {
            branch := GitCli_CaptureStdout(repo, "branch --show-current", 15000)
            if (branch != "") {
                pushUp := GitCli_Run(repo, 'push -u origin "' . StrReplace(branch, '"', '') . '"', 120000)
                if (pushUp.exitCode != 0) {
                    Finance_Notify("❌ Push failed: " . Finance_GitFormatPushError(pushUp), 4000, BANNER_ACCENT_ERROR)
                    return
                }
            } else {
                Finance_Notify("❌ Push failed: " . Finance_GitFormatPushError(push), 4000, BANNER_ACCENT_ERROR)
                return
            }
        }
        success := true
    } catch as e {
        Finance_Notify("❌ Push failed: " . e.Message, 3500, BANNER_ACCENT_ERROR)
    } finally {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        if (success) {
            Finance_Notify("✅ Updates pushed to cloud", 2500, BANNER_ACCENT_SUCCESS)
            Finance_CloseGui()
        } else {
            Finance_ShowMainMenu()
        }
    }
}
