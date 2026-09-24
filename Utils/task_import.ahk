; =============================================================================
; Utils module: task_import.ahk
; TASK_PACK Desktop import via Python CLI + Palace-style confirm
; Entry: Import Management [T] (#!+X / Utility [J])
; =============================================================================

Task_ImportPackScript() {
    return Task_PythonDir() . "\task_pack_import.py"
}

Task_ImportAiFixPath() {
    return A_Desktop . "\TASK_AI_FIX.txt"
}

; Palace-style confirm: label ListView + Import / Cancel.
Task_ImportConfirmPreview(title, labels) {
    global g_ImportMgmtGui
    owner := ""
    try {
        if (IsObject(g_ImportMgmtGui))
            owner := " +Owner" . g_ImportMgmtGui.Hwnd
    } catch {
        owner := ""
    }
    try {
        if (IsObject(g_ImportMgmtGui))
            g_ImportMgmtGui.Opt("-AlwaysOnTop")
    } catch {
    }
    g := Gui("+AlwaysOnTop +ToolWindow" . owner, title)
    g.SetFont("s10", "Segoe UI")
    lv := g.Add("ListView", "w720 h420 Grid", ["Row"])
    for lab in labels
        lv.Add("", lab)
    lv.ModifyCol(1, "AutoHdr")
    ok := false
    g.Add("Button", "y+8 w120 Default", "Import").OnEvent("Click", (*) => (ok := true, g.Destroy()))
    g.Add("Button", "x+8 w100", "Cancel").OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())
    g.Show()
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    try {
        if (IsObject(g_ImportMgmtGui))
            g_ImportMgmtGui.Opt("+AlwaysOnTop")
    } catch {
    }
    return ok
}

Task_ImportReadLabels(path) {
    labels := []
    if (path = "" || !FileExist(path))
        return labels
    text := ""
    try {
        f := FileOpen(path, "r", "UTF-8")
        if (f) {
            text := f.Read()
            f.Close()
            if (SubStr(text, 1, 1) = Chr(0xFEFF))
                text := SubStr(text, 2)
        }
    } catch {
        return labels
    }
    text := StrReplace(text, "`r`n", "`n")
    text := StrReplace(text, "`r", "`n")
    for line in StrSplit(text, "`n") {
        if (Trim(line) != "")
            labels.Push(line)
    }
    return labels
}

Task_ImportFailAi(errorMsg := "") {
    path := Task_ImportAiFixPath()
    if (FileExist(path))
        return ImportMgmt_OnAiFixReady(path, "TASK_AI_FIX.txt")
    if (errorMsg != "")
        Task_Notify(errorMsg, 5000, BANNER_ACCENT_ERROR)
    return false
}

; Import Management [T]: preview → confirm → commit (no browser).
Task_ImportPackFromDesktop(*) {
    Task_EnsureData()
    pyCmd := Task_FindPythonCmd()
    if (pyCmd = "") {
        Task_Notify("Python not found for Task import", 2500, BANNER_ACCENT_ERROR)
        return false
    }
    script := Task_ImportPackScript()
    if (!FileExist(script)) {
        Task_Notify("task_pack_import.py not found", 2200, BANNER_ACCENT_ERROR)
        return false
    }
    dataDir := Task_DataDir()
    stamp := A_TickCount
    jsonOut := A_Temp . "\task_import_preview_" . stamp . ".json"
    labelsOut := A_Temp . "\task_import_labels_" . stamp . ".txt"
    try FileDelete(jsonOut)
    catch {
    }
    try FileDelete(labelsOut)
    catch {
    }

    cmdPrev := pyCmd . ' "' . script . '" --data-dir "' . dataDir
        . '" preview --json-out "' . jsonOut . '" --labels-out "' . labelsOut . '"'
    try StandardLoadingBar_Show("⏳ Previewing TASK_PACK…", BANNER_ACCENT_INTERMEDIATE, { passive: false })
    catch {
    }
    exitPrev := 1
    try exitPrev := RunWait(A_ComSpec . ' /c ' . cmdPrev, Task_PythonDir(), "Hide")
    catch as e {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        Task_Notify("Preview failed: " . e.Message, 2800, BANNER_ACCENT_ERROR)
        return false
    }
    try StandardLoadingBar_Hide(200)
    catch {
    }

    if (exitPrev != 0 || !FileExist(jsonOut)) {
        Task_ImportFailAi("TASK_PACK preview failed")
        return false
    }

    labels := Task_ImportReadLabels(labelsOut)
    if (!labels.Length)
        labels.Push("(No preview rows — pack may be empty)")
    title := "Import TASK_PACK?"
    if (!Task_ImportConfirmPreview(title, labels))
        return false

    cmdCommit := pyCmd . ' "' . script . '" --data-dir "' . dataDir
        . '" commit --pack-json "' . jsonOut . '"'
    try StandardLoadingBar_Show("⏳ Importing TASK_PACK…", BANNER_ACCENT_INTERMEDIATE, { passive: false })
    catch {
    }
    exitCommit := 1
    try exitCommit := RunWait(A_ComSpec . ' /c ' . cmdCommit, Task_PythonDir(), "Hide")
    catch as e {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        Task_Notify("Commit failed: " . e.Message, 2800, BANNER_ACCENT_ERROR)
        return false
    }
    try StandardLoadingBar_Hide(200)
    catch {
    }

    ; 0 = full success; 2 = success with partial row errors (fix file written)
    ; preview writes foo.json; commit writes foo.commit.json next to it
    commitResultPath := RegExReplace(jsonOut, "i)\.json$", ".commit.json")
    if (exitCommit = 0) {
        Task_Notify(Task_ImportSuccessMessage(commitResultPath), 3200, BANNER_ACCENT_SUCCESS)
        ImportMgmt_OnImportSuccess()
        return true
    }
    if (exitCommit = 2) {
        Task_Notify(
            Task_ImportSuccessMessage(commitResultPath, true),
            3600,
            BANNER_ACCENT_INTERMEDIATE
        )
        if (FileExist(Task_ImportAiFixPath()))
            ImportMgmt_OnAiFixReady(Task_ImportAiFixPath(), "TASK_AI_FIX.txt")
        else
            ImportMgmt_OnImportSuccess()
        return true
    }
    Task_ImportFailAi("TASK_PACK commit failed")
    return false
}

; Build toast from commit JSON: counts + where to look (e.g. Habits → Dog).
Task_ImportSuccessMessage(commitJsonPath, withErrors := false) {
    prefix := withErrors ? "Imported with row errors" : "TASK_PACK imported"
    if (commitJsonPath = "" || !FileExist(commitJsonPath))
        return withErrors ? prefix . " — see AI fix" : prefix
    text := ""
    try {
        f := FileOpen(commitJsonPath, "r", "UTF-8")
        if (f) {
            text := f.Read()
            f.Close()
        }
    } catch {
        text := ""
    }
    if (text = "")
        return withErrors ? prefix . " — see AI fix" : prefix
    addedTasks := 0
    addedProj := 0
    addedInfo := 0
    hintFilt := ""
    hintProj := ""
    if RegExMatch(text, '"tasks"\s*:\s*(\d+)', &m)
        addedTasks := Integer(m[1])
    if RegExMatch(text, '"projects"\s*:\s*(\d+)', &m)
        addedProj := Integer(m[1])
    if RegExMatch(text, '"info"\s*:\s*(\d+)', &m)
        addedInfo := Integer(m[1])
    if RegExMatch(text, '"hint"\s*:\s*\{[^}]*"filter"\s*:\s*"([^"]*)"', &m)
        hintFilt := m[1]
    if RegExMatch(text, '"hint"\s*:\s*\{[^}]*"project"\s*:\s*"([^"]*)"', &m)
        hintProj := m[1]
    counts := ""
    if (addedTasks || addedProj || addedInfo)
        counts := addedTasks . " task(s)"
            . (addedProj ? ", " . addedProj . " project(s)" : "")
            . (addedInfo ? ", " . addedInfo . " info" : "")
    where := ""
    if (hintFilt != "" || hintProj != "") {
        filtLabel := ""
        if (hintFilt = "habits")
            filtLabel := "Habits"
        else if (hintFilt = "personal")
            filtLabel := "Personal"
        else if (hintFilt = "work")
            filtLabel := "Work"
        else if (hintFilt != "")
            filtLabel := hintFilt
        if (filtLabel != "" && hintProj != "")
            where := filtLabel . " → " . hintProj
        else if (filtLabel != "")
            where := filtLabel
        else
            where := hintProj
    }
    msg := prefix
    if (counts != "")
        msg .= ": " . counts
    if (where != "")
        msg .= " · look in " . where
    if (withErrors)
        msg .= " — see AI fix"
    return msg
}
