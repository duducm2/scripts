; =============================================================================
; Utils module: mnemonic_palace_practice_md.ahk
; Practice Markdown sync (Python study_practice_md.py)
; =============================================================================

Palace_PracticeDir() {
    dir := Palace_OutputDir() . "\practice"
    if (!DirExist(dir))
        DirCreate(dir)
    return dir
}

Palace_StudyIdForPalace(palaceId) {
    if (Trim(palaceId) = "")
        return ""
    p := Palace_FindById(Palace_Load("palaces"), palaceId)
    return IsObject(p) ? p["study_id"] : ""
}

Palace_StudyIdForBeast(beastId) {
    if (Trim(beastId) = "")
        return ""
    b := Palace_FindById(Palace_Load("beasts"), beastId)
    if (!IsObject(b))
        return ""
    return Palace_StudyIdForPalace(b["palace_id"])
}

Palace_StudyIdForAtom(atomId) {
    if (Trim(atomId) = "")
        return ""
    a := Palace_FindById(Palace_Load("atoms"), atomId)
    if (!IsObject(a))
        return ""
    return Palace_StudyIdForBeast(a["beast_id"])
}

Palace_SyncPracticeMd(studyIds := "", deleteSlugs := "") {
    py := Palace_PythonDir() . "\study_practice_md.py"
    if (!FileExist(py)) {
        Palace_Notify("study_practice_md.py not found", 2200, BANNER_ACCENT_ERROR)
        return false
    }
    pyCmd := Palace_FindPythonCmd()
    if (pyCmd = "") {
        Palace_Notify("Python not found for practice sync", 2500, BANNER_ACCENT_ERROR)
        return false
    }

    ids := []
    if (IsObject(studyIds)) {
        seen := Map()
        for sid in studyIds {
            s := Trim(String(sid))
            if (s != "" && !seen.Has(s)) {
                seen[s] := true
                ids.Push(s)
            }
        }
    }

    slugs := []
    if (IsObject(deleteSlugs)) {
        seenSlug := Map()
        for slug in deleteSlugs {
            s := Trim(String(slug))
            if (s != "" && !seenSlug.Has(s)) {
                seenSlug[s] := true
                slugs.Push(s)
            }
        }
    } else if (Trim(String(deleteSlugs)) != "") {
        slugs.Push(Trim(String(deleteSlugs)))
    }

    if (!ids.Length && !slugs.Length)
        return true

    dataDir := Palace_DataDir()
    outDir := Palace_OutputDir()
    notesRoot := Palace_NotesStudiesRoot()
    cmd := pyCmd . ' "' . py . '" --data-dir "' . dataDir . '" --output-dir "' . outDir . '"'
    if (notesRoot != "")
        cmd .= ' --notes-root "' . notesRoot . '"'
    for sid in ids
        cmd .= ' --study-id "' . sid . '"'
    for slug in slugs
        cmd .= ' --delete-slug "' . slug . '"'

    try StandardLoadingBar_Show("Syncing practice notes…", BANNER_ACCENT_INTERMEDIATE)
    catch {
    }
    exitCode := 0
    try {
        exitCode := RunWait(A_ComSpec . ' /c ' . cmd, A_ScriptDir, "Hide")
    } catch as e {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        Palace_Notify("Practice sync failed: " . e.Message, 2800, BANNER_ACCENT_ERROR)
        return false
    }
    if (exitCode != 0) {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        Palace_Notify("Practice sync failed (exit " . exitCode . ")", 2800, BANNER_ACCENT_ERROR)
        return false
    }
    try StandardLoadingBar_Hide(400)
    catch {
    }
    return true
}

Palace_SyncAllPracticeMd() {
    py := Palace_PythonDir() . "\study_practice_md.py"
    if (!FileExist(py)) {
        Palace_Notify("study_practice_md.py not found", 2200, BANNER_ACCENT_ERROR)
        return false
    }
    pyCmd := Palace_FindPythonCmd()
    if (pyCmd = "") {
        Palace_Notify("Python not found for practice sync", 2500, BANNER_ACCENT_ERROR)
        return false
    }
    dataDir := Palace_DataDir()
    outDir := Palace_OutputDir()
    notesRoot := Palace_NotesStudiesRoot()
    cmd := pyCmd . ' "' . py . '" --data-dir "' . dataDir . '" --output-dir "' . outDir
        . '" --sync-all --migrate-image-paths'
    if (notesRoot != "")
        cmd .= ' --notes-root "' . notesRoot . '"'
    try StandardLoadingBar_Show("Syncing all practice notes…", BANNER_ACCENT_INTERMEDIATE)
    catch {
    }
    exitCode := 0
    try {
        exitCode := RunWait(A_ComSpec . ' /c ' . cmd, A_ScriptDir, "Hide")
    } catch as e {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        Palace_Notify("Practice sync failed: " . e.Message, 2800, BANNER_ACCENT_ERROR)
        return false
    }
    if (exitCode != 0) {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        Palace_Notify("Practice sync failed (exit " . exitCode . ")", 2800, BANNER_ACCENT_ERROR)
        return false
    }
    try StandardLoadingBar_Hide(400)
    catch {
    }
    return true
}
