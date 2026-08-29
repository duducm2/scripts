; =============================================================================
; Utils module: task_migrate.ahk
; Run Python migrate_from_md.py against notes MD sources
; =============================================================================

Task_NotesCandidates() {
    roots := []
    try {
        r := GetNotesRepoPath()
        if (r != "")
            roots.Push(r)
    } catch {
    }
    ; Common sibling of scripts repo
    roots.Push(A_ScriptDir . "\..\notes")
    roots.Push("C:\Users\eduev\Meu Drive\17 - Projects\notes")
    out := []
    seen := Map()
    for r in roots {
        full := ""
        try full := DirExist(r) ? r : ""
        catch {
            full := ""
        }
        if (full = "")
            continue
        ; normalize
        try full := RTrim(full, "\")
        catch {
        }
        key := StrLower(full)
        if (seen.Has(key))
            continue
        seen[key] := true
        out.Push(full)
    }
    return out
}

Task_ResolveMdPaths(&workPath, &punctualPath, &habitsPath) {
    workPath := ""
    punctualPath := ""
    habitsPath := ""
    for root in Task_NotesCandidates() {
        w := root . "\work\work.md"
        p := root . "\main\punctual.md"
        h := root . "\main\habits.md"
        if (FileExist(w) && FileExist(p) && FileExist(h)) {
            workPath := w
            punctualPath := p
            habitsPath := h
            return true
        }
    }
    return false
}

Task_RunMigrateFromMd() {
    Task_EnsureData()
    workPath := ""
    punctualPath := ""
    habitsPath := ""
    if (!Task_ResolveMdPaths(&workPath, &punctualPath, &habitsPath)) {
        Task_Alert(
            "Could not find work.md, punctual.md, and habits.md under notes repo candidates.",
            "Migrate")
        return
    }
    existing := Task_Load("tasks")
    if (existing.Length > 0) {
        if (!Task_Confirm(
            "This will OVERWRITE tasks CSV data with a fresh migration from:`n"
            . workPath . "`n" . punctualPath . "`n" . habitsPath
            . "`n`nContinue?", "Migrate"))
            return
    } else {
        if (!Task_Confirm(
            "Import Markdown into CSV?`n"
            . workPath . "`n" . punctualPath . "`n" . habitsPath, "Migrate"))
            return
    }

    py := Task_PythonDir() . "\migrate_from_md.py"
    if (!FileExist(py)) {
        Task_Notify("migrate_from_md.py not found", 2000, BANNER_ACCENT_ERROR)
        return
    }
    pyCmd := Task_FindPythonCmd()
    if (pyCmd = "") {
        Task_Notify("Python not found.", 3500, BANNER_ACCENT_ERROR)
        return
    }
    dataDir := Task_DataDir()
    try StandardLoadingBar_Show("Migrating Markdown…", BANNER_ACCENT_INTERMEDIATE)
    catch {
    }
    cmd := pyCmd . ' "' . py . '" --data-dir "' . dataDir
        . '" --work "' . workPath
        . '" --punctual "' . punctualPath
        . '" --habits "' . habitsPath . '"'
    exitCode := 0
    try {
        exitCode := RunWait(A_ComSpec . ' /c ' . cmd, A_ScriptDir, "Hide")
    } catch as e {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        Task_Notify("Migrate failed: " . e.Message, 2500, BANNER_ACCENT_ERROR)
        return
    }
    try StandardLoadingBar_Hide(400)
    catch {
    }
    if (exitCode != 0) {
        Task_Notify("Migrate exit " . exitCode, 3000, BANNER_ACCENT_ERROR)
        return
    }
    projects := Task_Load("projects")
    tasks := Task_Load("tasks")
    infos := Task_Load("info_points")
    atts := Task_Load("attachments")
    Task_Notify(
        "Migrated " . projects.Length . " proj / " . tasks.Length . " tasks / "
        . infos.Length . " info / " . atts.Length . " att",
        3500, BANNER_ACCENT_SUCCESS)
    Task_ShowMainMenu()
}
