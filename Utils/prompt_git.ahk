; =============================================================================
; Utils module: prompt_git.ahk
; Prompt-aware git history, diff, rollback, optional commit on save
; =============================================================================

global g_PromptGitHistoryGui := false

PromptGit_RepoRoot() {
    return GitCli_RevParseTopLevel(A_ScriptDir)
}

PromptGit_RelPath(absPath) {
    root := RTrim(PromptGit_RepoRoot(), "\")
    if (root = "" || absPath = "")
        return ""
    norm := RTrim(absPath, "\")
    prefix := StrLower(root) "\"
    if (StrLower(SubStr(norm, 1, StrLen(prefix))) = prefix)
        return SubStr(norm, StrLen(root) + 2)
    return norm
}

PromptGit_LogEntries(relPath, limit := 12) {
    repo := PromptGit_RepoRoot()
    if (repo = "" || relPath = "")
        return []
    fmt := 'log -n ' limit ' --format=%H|%s|%ci -- "' relPath '"'
    out := GitCli_CaptureStdout(repo, fmt, 30000)
    if (out = "")
        return []
    entries := []
    for line in StrSplit(out, "`n", "`r") {
        line := Trim(line)
        if (line = "")
            continue
        parts := StrSplit(line, "|", , 3)
        if (parts.Length < 2)
            continue
        hash := parts[1]
        subj := parts.Length >= 2 ? parts[2] : ""
        date := parts.Length >= 3 ? parts[3] : ""
        short := SubStr(hash, 1, 8)
        entries.Push({ hash: hash, short: short, subject: subj, date: date })
    }
    return entries
}

PromptGit_DiffText(relPath, commitHash := "") {
    repo := PromptGit_RepoRoot()
    if (repo = "" || relPath = "")
        return ""
    args := commitHash != "" ? 'show ' commitHash ' -- "' relPath '"' : 'diff HEAD -- "' relPath '"'
    return GitCli_CaptureStdout(repo, args, 45000)
}

PromptGit_CommitPromptFiles(promptName, bodyPath, commitMsg := "") {
    repo := PromptGit_RepoRoot()
    if (repo = "")
        return { ok: false, msg: "Not a git repository" }
    relBody := PromptGit_RelPath(bodyPath)
    relIni := PromptGit_RelPath(PromptData_IniPath())
    if (relBody = "")
        return { ok: false, msg: "Prompt file outside repo" }
    msg := Trim(commitMsg)
    if (msg = "")
        msg := "update prompt"
    safeMsg := StrReplace(msg, '"', "'")
    subject := "[prompt] " promptName ": " safeMsg
    addArgs := 'add -- "' relBody '"'
    if (relIni != "" && FileExist(PromptData_IniPath()))
        addArgs .= ' "' relIni '"'
    addResult := GitCli_Run(repo, addArgs, 20000)
    if (addResult.exitCode != 0)
        return { ok: false, msg: "git add failed" }
    staged := GitCli_Run(repo, "diff --cached --name-only", 20000)
    if (staged.exitCode != 0 || Trim(staged.stdout) = "")
        return { ok: true, msg: "Nothing to commit" }
    commitArgs := 'commit -m "' subject '"'
    commitResult := GitCli_Run(repo, commitArgs, 30000)
    if (commitResult.exitCode != 0)
        return { ok: false, msg: commitResult.stdout != "" ? commitResult.stdout : "git commit failed" }
    return { ok: true, msg: commitResult.stdout != "" ? commitResult.stdout : "Committed" }
}

PromptGit_RollbackFile(absPath, commitHash) {
    repo := PromptGit_RepoRoot()
    rel := PromptGit_RelPath(absPath)
    if (repo = "" || rel = "" || commitHash = "")
        return false
    result := GitCli_Run(repo, 'checkout ' commitHash ' -- "' rel '"', 30000)
    return (result.exitCode = 0)
}

PromptGit_ShowHistory(absPath) {
    global g_PromptGitHistoryGui
    rel := PromptGit_RelPath(absPath)
    if (rel = "") {
        UtilitySelector_Notify("File is not inside the git repo.")
        return
    }
    entries := PromptGit_LogEntries(rel)
    if (entries.Length = 0) {
        UtilitySelector_Notify("No git history for this prompt file.")
        return
    }
    g_PromptGitHistoryGui := Gui("+AlwaysOnTop +ToolWindow", "Prompt history")
    g_PromptGitHistoryGui.SetFont("s10", "Segoe UI")
    g_PromptGitHistoryGui.Add("Text", "xm w560", rel)
    lv := g_PromptGitHistoryGui.Add("ListView", "xm w560 r10", ["Date", "Commit", "Subject"])
    lv.ModifyCol(1, 140)
    lv.ModifyCol(2, 70)
    lv.ModifyCol(3, 340)
    for e in entries
        lv.Add("", e.date, e.short, e.subject)
    if (entries.Length)
        lv.Modify(1, "Select Focus Vis")
    diffCtrl := g_PromptGitHistoryGui.Add("Edit", "xm w560 r12 ReadOnly -Wrap", "")
    lv.OnEvent("ItemFocus", (ctrl, info) => PromptGit_OnHistoryFocus(lv, diffCtrl, rel, info))
    PromptGit_OnHistoryFocus(lv, diffCtrl, rel, { EventInfo: 1 })
    g_PromptGitHistoryGui.Add("Button", "xm w100", "Rollback").OnEvent("Click", (*) => PromptGit_OnRollback(lv,
        diffCtrl, rel, absPath))
    g_PromptGitHistoryGui.Add("Button", "x+8 yp w100", "Close").OnEvent("Click", PromptGit_CloseHistory)
    g_PromptGitHistoryGui.OnEvent("Close", PromptGit_CloseHistory)
    g_PromptGitHistoryGui.OnEvent("Escape", PromptGit_CloseHistory)
    mon := UtilitySelector_ActiveMonitorWorkArea()
    g_PromptGitHistoryGui.Show("Hide")
    g_PromptGitHistoryGui.GetPos(, , &gw, &gh)
    g_PromptGitHistoryGui.Show("x" (mon.left + (mon.width - gw) // 2) " y" (mon.top + (mon.height - gh) // 2))
}

PromptGit_OnHistoryFocus(lv, diffCtrl, relPath, info) {
    row := 1
    try row := Integer(info.EventInfo)
    catch {
        row := 1
    }
    if (row < 1)
        return
    hash := ""
    try hash := lv.GetText(row, 2)
    catch {
        return
    }
    if (hash = "")
        return
    entries := PromptGit_LogEntries(relPath, 50)
    fullHash := ""
    for e in entries {
        if (e.short = hash) {
            fullHash := e.hash
            break
        }
    }
    if (fullHash = "")
        return
    diff := PromptGit_DiffText(relPath, fullHash)
    try diffCtrl.Value := diff != "" ? diff : "(no diff for this commit)"
    catch {
    }
}

PromptGit_OnRollback(lv, diffCtrl, relPath, absPath) {
    row := lv.GetNext(0, "F")
    if (!row) {
        UtilitySelector_Notify("Select a commit to roll back to.")
        return
    }
    short := lv.GetText(row, 2)
    entries := PromptGit_LogEntries(relPath, 50)
    fullHash := ""
    for e in entries {
        if (e.short = short) {
            fullHash := e.hash
            break
        }
    }
    if (fullHash = "") {
        UtilitySelector_Notify("Could not resolve commit.")
        return
    }
    if (MsgBox("Restore prompt file from commit " short "?", "Rollback", "YesNo Icon!") != "Yes")
        return
    if PromptGit_RollbackFile(absPath, fullHash) {
        UtilitySelector_Notify("Rolled back to " short ".")
        PromptGit_CloseHistory()
    } else {
        UtilitySelector_Notify("Rollback failed.")
    }
}

PromptGit_CloseHistory(*) {
    global g_PromptGitHistoryGui
    if (IsObject(g_PromptGitHistoryGui)) {
        try g_PromptGitHistoryGui.Destroy()
        catch {
        }
    }
    g_PromptGitHistoryGui := false
}

PromptGit_PromptCommitMessage(defaultName := "") {
    ib := InputBox("Optional commit message (leave blank for default):", "Git commit prompt", "w420", "")
    if (ib.Result = "Cancel")
        return false
    return Trim(ib.Value)
}
