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

    try StandardLoadingBar_Show("⏳ Syncing practice notes…", BANNER_ACCENT_INTERMEDIATE)
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
    try StandardLoadingBar_Hide(0)
    catch {
    }
    return true
}

Palace_SyncAllPracticeMd(showUi := true) {
    py := Palace_PythonDir() . "\study_practice_md.py"
    if (!FileExist(py)) {
        if (showUi)
            Palace_Notify("study_practice_md.py not found", 2200, BANNER_ACCENT_ERROR)
        return false
    }
    pyCmd := Palace_FindPythonCmd()
    if (pyCmd = "") {
        if (showUi)
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
    if (showUi) {
        try StandardLoadingBar_Show("Syncing all practice notes…", BANNER_ACCENT_INTERMEDIATE)
        catch {
        }
    }
    exitCode := 0
    try {
        exitCode := RunWait(A_ComSpec . ' /c ' . cmd, A_ScriptDir, "Hide")
    } catch as e {
        if (showUi) {
            try StandardLoadingBar_Hide(0)
            catch {
            }
            Palace_Notify("Practice sync failed: " . e.Message, 2800, BANNER_ACCENT_ERROR)
        }
        return false
    }
    if (exitCode != 0) {
        if (showUi) {
            try StandardLoadingBar_Hide(0)
            catch {
            }
            Palace_Notify("Practice sync failed (exit " . exitCode . ")", 2800, BANNER_ACCENT_ERROR)
        }
        return false
    }
    if (showUi) {
        try StandardLoadingBar_Hide(400)
        catch {
        }
    }
    return true
}

Palace_SyncAllPlansMd(showUi := true) {
    py := Palace_PythonDir() . "\study_plans_md.py"
    if (!FileExist(py)) {
        if (showUi)
            Palace_Notify("study_plans_md.py not found", 2200, BANNER_ACCENT_ERROR)
        return false
    }
    pyCmd := Palace_FindPythonCmd()
    if (pyCmd = "") {
        if (showUi)
            Palace_Notify("Python not found for plans sync", 2500, BANNER_ACCENT_ERROR)
        return false
    }
    dataDir := Palace_DataDir()
    outDir := Palace_OutputDir()
    notesRoot := Palace_NotesStudiesRoot()
    cmd := pyCmd . ' "' . py . '" --data-dir "' . dataDir . '" --output-dir "' . outDir
        . '" --sync-all'
    if (notesRoot != "")
        cmd .= ' --studies-root "' . notesRoot . '"'
    if (showUi) {
        try StandardLoadingBar_Show("Syncing all study plans…", BANNER_ACCENT_INTERMEDIATE)
        catch {
        }
    }
    exitCode := 0
    try {
        exitCode := RunWait(A_ComSpec . ' /c ' . cmd, A_ScriptDir, "Hide")
    } catch as e {
        if (showUi) {
            try StandardLoadingBar_Hide(0)
            catch {
            }
            Palace_Notify("Plans sync failed: " . e.Message, 2800, BANNER_ACCENT_ERROR)
        }
        return false
    }
    if (exitCode != 0) {
        if (showUi) {
            try StandardLoadingBar_Hide(0)
            catch {
            }
            Palace_Notify("Plans sync failed (exit " . exitCode . ")", 2800, BANNER_ACCENT_ERROR)
        }
        return false
    }
    if (showUi) {
        try StandardLoadingBar_Hide(400)
        catch {
        }
    }
    return true
}

; Sync plan Markdown for one or more study ids (CSV → output/plans/).
Palace_SyncPlansMd(studyIds := "") {
    py := Palace_PythonDir() . "\study_plans_md.py"
    if (!FileExist(py))
        return false
    pyCmd := Palace_FindPythonCmd()
    if (pyCmd = "")
        return false
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
    } else if (Trim(String(studyIds)) != "") {
        ids.Push(Trim(String(studyIds)))
    }
    if (!ids.Length)
        return true
    dataDir := Palace_DataDir()
    outDir := Palace_OutputDir()
    notesRoot := Palace_NotesStudiesRoot()
    cmd := pyCmd . ' "' . py . '" --data-dir "' . dataDir . '" --output-dir "' . outDir . '"'
    if (notesRoot != "")
        cmd .= ' --studies-root "' . notesRoot . '"'
    for sid in ids
        cmd .= ' --study-id "' . sid . '"'
    try RunWait(A_ComSpec . ' /c ' . cmd, A_ScriptDir, "Hide")
    catch {
        return false
    }
    return true
}

Palace_ForceRegenAllMarkdown() {
    ; Manual fallback: force-regenerate practice + plan Markdown (independent of CRUD hooks)
    pyCmd := Palace_FindPythonCmd()
    if (pyCmd = "") {
        Palace_Notify("Python not found for Markdown regen", 2800, BANNER_ACCENT_ERROR)
        return false
    }
    try StandardLoadingBar_Show("Regenerating practice Markdown…", BANNER_ACCENT_INTERMEDIATE)
    catch {
    }
    okPractice := Palace_SyncAllPracticeMd(false)
    if (!okPractice) {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        Palace_Notify("Practice Markdown regen failed", 3000, BANNER_ACCENT_ERROR)
        return false
    }
    try StandardLoadingBar_Update("Regenerating study plan Markdown…", BANNER_ACCENT_INTERMEDIATE)
    catch {
        try StandardLoadingBar_Show("Regenerating study plan Markdown…", BANNER_ACCENT_INTERMEDIATE)
        catch {
        }
    }
    okPlans := Palace_SyncAllPlansMd(false)
    if (!okPlans) {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        Palace_Notify("Plan Markdown regen failed (practice OK)", 3200, BANNER_ACCENT_ERROR)
        return false
    }
    try StandardLoadingBar_Hide(400)
    catch {
    }
    Palace_Notify("All Markdown regenerated (practice + plans)", 2800, BANNER_ACCENT_SUCCESS)
    return true
}
