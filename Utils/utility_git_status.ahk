; =============================================================================
; Utils module: utility_git_status.ahk
; Main Repos window — glance dirty files in scripts + notes; push via Utility [G].
; Entry: Win+Alt+Shift+Y double-tap (see focus_mode.ahk).
; =============================================================================

global g_UtilityGitStatusGui := false
global g_UtilityGitStatusLv := false
global g_UtilityGitStatusSummary := false
global g_UtilityGitStatusHotkeys := []
global g_UtilityGitStatusActive := false

Utility_GitStatusXyLabel(xy) {
    x := SubStr(xy, 1, 1)
    y := SubStr(xy, 2, 1)
    if (xy = "??")
        return "Untracked"
    if (x = "R" || y = "R" || x = "C" || y = "C")
        return "Renamed"
    if (x = "D" || y = "D")
        return "Deleted"
    if (x = "A" || y = "A")
        return "Added"
    if (x = "M" || y = "M")
        return "Modified"
    if (x = "U" || y = "U")
        return "Conflict"
    if (Trim(xy) = "")
        return "Changed"
    return Trim(xy)
}

; Parse porcelain lines into {repo, xy, status, path} rows.
Utility_GitStatusParsePorcelain(porcelain, repoLabel) {
    rows := []
    for line in StrSplit(porcelain, "`n", "`r") {
        t := line
        if (StrLen(t) < 3)
            continue
        xy := SubStr(t, 1, 2)
        path := Trim(SubStr(t, 3))
        if (path = "")
            continue
        rows.Push({ repo: repoLabel, xy: xy, status: Utility_GitStatusXyLabel(xy), path: path })
    }
    return rows
}

Utility_GitStatusCollectRows() {
    rows := []
    errMsgs := []

    scriptsRoot := GitCli_RevParseTopLevel(A_ScriptDir)
    if (scriptsRoot = "") {
        errMsgs.Push("Scripts: not a git repository")
    } else {
        status := GitCli_Run(scriptsRoot, "status --porcelain", 30000)
        if (status.exitCode != 0)
            errMsgs.Push("Scripts: status failed — " . Utility_GitFirstErrorLine(status))
        else {
            for r in Utility_GitStatusParsePorcelain(status.stdout, "Scripts")
                rows.Push(r)
        }
    }

    notesRoot := ""
    try notesRoot := GetNotesRepoPath()
    catch {
        notesRoot := ""
    }
    if (notesRoot = "" || !DirExist(notesRoot)) {
        errMsgs.Push("Notes: repo folder not found")
    } else {
        notesTop := GitCli_RevParseTopLevel(notesRoot)
        if (notesTop = "") {
            errMsgs.Push("Notes: not a git repository")
        } else {
            status := GitCli_Run(notesTop, "status --porcelain", 30000)
            if (status.exitCode != 0)
                errMsgs.Push("Notes: status failed — " . Utility_GitFirstErrorLine(status))
            else {
                for r in Utility_GitStatusParsePorcelain(status.stdout, "Notes")
                    rows.Push(r)
            }
        }
    }

    return { rows: rows, errors: errMsgs }
}

Utility_GitStatusSummaryText(rows, errors) {
    scriptsN := 0
    notesN := 0
    for r in rows {
        if (r.repo = "Scripts")
            scriptsN += 1
        else if (r.repo = "Notes")
            notesN += 1
    }
    text := "Scripts: " . scriptsN . "  ·  Notes: " . notesN
    if (scriptsN = 0 && notesN = 0 && errors.Length = 0)
        text .= "  —  clean"
    if (errors.Length > 0)
        text .= "  |  " . errors[1]
    return text
}

Utility_GitStatusIsOpen() {
    global g_UtilityGitStatusGui, g_UtilityGitStatusActive
    if (!g_UtilityGitStatusActive)
        return false
    try {
        return IsObject(g_UtilityGitStatusGui) && WinExist("ahk_id " g_UtilityGitStatusGui.Hwnd)
    } catch {
        return false
    }
}

Utility_GitStatusCleanup() {
    global g_UtilityGitStatusGui, g_UtilityGitStatusLv, g_UtilityGitStatusSummary
    global g_UtilityGitStatusActive, g_OnEscapePressed
    Utility_GitStatusUnbindHotkeys()
    if (IsSet(g_OnEscapePressed) && g_OnEscapePressed = Utility_GitStatusOnEscape)
        g_OnEscapePressed := ""
    try {
        if (IsObject(g_UtilityGitStatusGui))
            g_UtilityGitStatusGui.Destroy()
    } catch {
    }
    g_UtilityGitStatusGui := false
    g_UtilityGitStatusLv := false
    g_UtilityGitStatusSummary := false
    g_UtilityGitStatusActive := false
}

Utility_GitStatusOnEscape(*) {
    if (Utility_GitStatusIsOpen()) {
        Utility_GitStatusCleanup()
        return true
    }
    return false
}

Utility_GitStatusHotIf(*) {
    global g_UtilityGitStatusGui
    if (!IsObject(g_UtilityGitStatusGui))
        return false
    try {
        return WinActive("ahk_id " g_UtilityGitStatusGui.Hwnd)
    } catch {
        return false
    }
}

Utility_GitStatusUnbindHotkeys() {
    global g_UtilityGitStatusHotkeys
    try HotIf(Utility_GitStatusHotIf)
    catch {
    }
    for key in g_UtilityGitStatusHotkeys {
        try Hotkey(key, "Off")
        catch {
        }
    }
    g_UtilityGitStatusHotkeys := []
    try HotIf()
    catch {
    }
}

Utility_GitStatusBindHotkeys() {
    global g_UtilityGitStatusHotkeys
    Utility_GitStatusUnbindHotkeys()
    pairs := [
        ["g", Utility_GitStatusPush],
        ["G", Utility_GitStatusPush],
        ["p", Utility_GitStatusPush],
        ["P", Utility_GitStatusPush],
        ["d", Utility_GitStatusDiscard],
        ["D", Utility_GitStatusDiscard],
        ["r", Utility_GitStatusRefresh],
        ["R", Utility_GitStatusRefresh],
        ["Escape", (*) => Utility_GitStatusCleanup()]
    ]
    try HotIf(Utility_GitStatusHotIf)
    catch {
        return
    }
    for p in pairs {
        try {
            Hotkey(p[1], p[2], "On")
            g_UtilityGitStatusHotkeys.Push(p[1])
        } catch {
        }
    }
    try HotIf()
    catch {
    }
}

Utility_GitStatusCenter(guiObj, w := 820, h := 520) {
    ml := 0
    mt := 0
    mr := 0
    mb := 0
    try GetActiveMonitorWorkArea_StandardBar(&ml, &mt, &mr, &mb)
    catch {
        MonitorGetWorkArea(MonitorGetPrimary(), &ml, &mt, &mr, &mb)
    }
    x := ml + ((mr - ml) - w) // 2
    y := mt + ((mb - mt) - h) // 2
    guiObj.Show("x" . x . " y" . y . " w" . w . " h" . h)
}

Utility_GitStatusPopulate() {
    global g_UtilityGitStatusLv, g_UtilityGitStatusSummary
    if (!IsObject(g_UtilityGitStatusLv))
        return
    collected := Utility_GitStatusCollectRows()
    try g_UtilityGitStatusLv.Delete()
    catch {
    }
    for r in collected.rows
        g_UtilityGitStatusLv.Add("", r.repo, r.status, r.path)
    if (collected.rows.Length > 0)
        try g_UtilityGitStatusLv.Modify(1, "Select Focus Vis")
        catch {
        }
    if (IsObject(g_UtilityGitStatusSummary)) {
        try g_UtilityGitStatusSummary.Value := Utility_GitStatusSummaryText(collected.rows, collected.errors)
        catch {
        }
    }
}

Utility_GitStatusRefresh(*) {
    Utility_GitStatusPopulate()
}

Utility_GitStatusPush(*) {
    ; Close first so the push loading bar is usable and the user can keep working.
    Utility_GitStatusCleanup()
    ; Same path as #!+9 hold / Utility Shortcuts [G]
    Utility_GitSyncPush()
}

; reset --hard + clean -fd for one repo. Returns "ok" or "error:…".
Utility_GitStatusDiscardOne(repoDir, label) {
    repo := GitCli_RevParseTopLevel(repoDir)
    if (repo = "")
        return "error:" . label . " not a git repository"
    Utility_GitPassiveBar("⏳ " . label . ": reset --hard…")
    reset := GitCli_Run(repo, "reset --hard HEAD", 60000)
    if (reset.exitCode != 0)
        return "error:" . label . " reset failed: " . Utility_GitFirstErrorLine(reset)
    Utility_GitPassiveBar("⏳ " . label . ": clean -fd…")
    clean := GitCli_Run(repo, "clean -fd", 60000)
    if (clean.exitCode != 0)
        return "error:" . label . " clean failed: " . Utility_GitFirstErrorLine(clean)
    return "ok"
}

; [D] Discard all local changes in scripts + notes (after MsgBox confirm).
Utility_GitStatusDiscard(*) {
    global g_UtilityGitStatusGui

    ownerOpt := ""
    try {
        if (IsObject(g_UtilityGitStatusGui)) {
            g_UtilityGitStatusGui.Opt("-AlwaysOnTop")
            ownerOpt := " Owner" . g_UtilityGitStatusGui.Hwnd
        }
    } catch {
    }

    answer := MsgBox(
        "Discard ALL local changes in Scripts and Notes?`n`n"
        . "Runs git reset --hard and git clean -fd on both main repos.`n"
        . "Tracked modifications and untracked files are permanently removed.`n"
        . "Nothing will be pushed.`n`n"
        . "This cannot be undone.",
        "Main Repos — Discard",
        "YesNo Icon! Default2" . ownerOpt)

    if (answer != "Yes") {
        try {
            if (IsObject(g_UtilityGitStatusGui))
                g_UtilityGitStatusGui.Opt("+AlwaysOnTop")
        } catch {
        }
        return
    }

    Utility_GitStatusCleanup()

    resultMsg := ""
    resultAccent := BANNER_ACCENT_INFO
    try {
        try StandardLoadingBar_Show("⏳ Discarding local changes…", BANNER_ACCENT_INTERMEDIATE, { passive: false })
        catch {
        }

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

        scriptsResult := Utility_GitStatusDiscardOne(scriptsRoot, "Scripts")
        notesResult := Utility_GitStatusDiscardOne(notesRoot, "Notes")

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

        resultMsg := "✅ Scripts + Notes discarded"
        resultAccent := BANNER_ACCENT_SUCCESS
    } catch as e {
        resultMsg := "❌ Discard failed: " . e.Message
        resultAccent := BANNER_ACCENT_ERROR
    } finally {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        if (resultMsg != "")
            Utility_GitNotify(resultMsg, 2800, resultAccent)
    }
}

; Toggle: open Main Repos window, or close if already open.
; Loading Indication while git status runs (see docs/standard_information_display.md).
Utility_GitStatusShow() {
    global g_UtilityGitStatusGui, g_UtilityGitStatusLv, g_UtilityGitStatusSummary
    global g_UtilityGitStatusActive, g_OnEscapePressed

    if (Utility_GitStatusIsOpen()) {
        Utility_GitStatusCleanup()
        return
    }

    try StandardLoadingBar_Show("⏳ Opening main repos...", BANNER_ACCENT_INTERMEDIATE, { passive: false })
    catch {
    }
    try {
        Utility_GitStatusCleanup()
        g_UtilityGitStatusActive := true

        g_UtilityGitStatusGui := Gui("+AlwaysOnTop +ToolWindow", "Main Repos — scripts + notes")
        g_UtilityGitStatusGui.SetFont("s10", "Segoe UI")
        g_UtilityGitStatusSummary := g_UtilityGitStatusGui.Add("Text", "x12 y12 w790", "Loading…")
        g_UtilityGitStatusGui.Add("Text", "x12 y36 w790 c666666",
            "[G]/[P] Push   [D] Discard (reset+clean)   [R] Refresh   Esc close")
        g_UtilityGitStatusLv := g_UtilityGitStatusGui.Add("ListView", "x12 y60 w790 h400 Grid", ["Repo", "Status",
            "Path"])
        g_UtilityGitStatusLv.ModifyCol(1, 80)
        g_UtilityGitStatusLv.ModifyCol(2, 90)
        g_UtilityGitStatusLv.ModifyCol(3, 600)

        g_UtilityGitStatusGui.Add("Button", "x12 y472 w100", "Push [G]").OnEvent("Click", Utility_GitStatusPush)
        g_UtilityGitStatusGui.Add("Button", "x122 y472 w100", "Discard [D]").OnEvent("Click", Utility_GitStatusDiscard)
        g_UtilityGitStatusGui.Add("Button", "x232 y472 w100", "Refresh [R]").OnEvent("Click", Utility_GitStatusRefresh)
        g_UtilityGitStatusGui.Add("Button", "x342 y472 w100", "Close").OnEvent("Click", (*) => Utility_GitStatusCleanup())

        g_UtilityGitStatusGui.OnEvent("Close", (*) => Utility_GitStatusCleanup())
        g_UtilityGitStatusGui.OnEvent("Escape", (*) => Utility_GitStatusCleanup())

        g_OnEscapePressed := Utility_GitStatusOnEscape
        Utility_GitStatusBindHotkeys()
        try StandardLoadingBar_Update("⏳ Checking git status…", BANNER_ACCENT_INTERMEDIATE)
        catch {
        }
        Utility_GitStatusPopulate()
        ; Hide before ListView Show so the bar does not sit on top of the window.
        try StandardLoadingBar_Hide(0)
        catch {
        }
        Utility_GitStatusCenter(g_UtilityGitStatusGui, 820, 520)
        try g_UtilityGitStatusLv.Focus()
        catch {
        }
    } catch {
        Utility_GitStatusCleanup()
    } finally {
        try StandardLoadingBar_Hide(0)
        catch {
        }
    }
}
