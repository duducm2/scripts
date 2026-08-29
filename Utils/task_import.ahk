; =============================================================================
; Utils module: task_import.ahk
; Import AI-generated TASK_PACK from Desktop into tasks/data CSV
; Agent docs: docs/prompt-data-output-and-finance-packs.md
; =============================================================================

Task_DesktopNewest(pattern) {
    newest := ""
    newestTime := 0
    loop files A_Desktop . "\" . pattern, "F" {
        ts := Number(A_LoopFileTimeModified)
        if (ts > newestTime) {
            newestTime := ts
            newest := A_LoopFileFullPath
        }
    }
    return newest
}

Task_MdFence() {
    bt := Chr(96)
    return bt . bt . bt
}

Task_ExtractPackFileSection(text, fileName) {
    names := [fileName]
    if (RegExMatch(fileName, "i)\.csv$"))
        names.Push(RegExReplace(fileName, "i)\.csv$", ""))
    else
        names.Push(fileName . ".csv")
    for name in names {
        for style in ["===", "---"] {
            needle := style . "FILE: " . name . style
            endNeedle := style . "END_FILE" . style
            pos := InStr(text, needle, false)
            if (!pos)
                continue
            rest := SubStr(text, pos + StrLen(needle))
            endPos := InStr(rest, endNeedle, false)
            if (endPos)
                return Trim(SubStr(rest, 1, endPos - 1), "`r`n `t")
            nextPos := 0
            for style2 in ["===", "---"] {
                n2 := InStr(rest, "`n" . style2 . "FILE:", false)
                if (n2 && (!nextPos || n2 < nextPos))
                    nextPos := n2
            }
            if (nextPos)
                return Trim(SubStr(rest, 1, nextPos - 1), "`r`n `t")
            return Trim(rest, "`r`n `t")
        }
    }
    return ""
}

Task_StripMarkdownFences(body) {
    fence := Task_MdFence()
    body := Trim(body, "`r`n `t")
    if (SubStr(body, 1, 3) = fence) {
        nl := InStr(body, "`n")
        body := nl ? SubStr(body, nl + 1) : ""
    }
    body := Trim(body, "`r`n `t")
    if (SubStr(body, -3) = fence)
        body := Trim(SubStr(body, 1, StrLen(body) - 3), "`r`n `t")
    return body
}

Task_ParseCsvBody(body) {
    body := Task_StripMarkdownFences(body)
    if (Trim(body) = "")
        return []
    tmp := A_Temp . "\task_pack_section_" . A_TickCount . "_" . Random(1000, 9999) . ".csv"
    Task_WriteUtf8(tmp, body)
    rows := Task_ReadCsv(tmp)
    try FileDelete(tmp)
    catch {
    }
    ; Drop accidental duplicate header rows
    cleaned := []
    for r in rows {
        skip := false
        if (r.Has("title") && StrLower(Trim(r["title"])) = "title")
            skip := true
        if (r.Has("project_title") && StrLower(Trim(r["project_title"])) = "project_title")
            skip := true
        if (r.Has("attach_to") && StrLower(Trim(r["attach_to"])) = "attach_to")
            skip := true
        if (!skip)
            cleaned.Push(r)
    }
    return cleaned
}

Task_WriteAiCompanionImportError(errorMsg, extraNotes := "") {
    errorMsg := Trim(errorMsg)
    if (errorMsg = "")
        return ""
    guidance := Task_AiCompanionFixGuidance(errorMsg)
    body := "The Desktop Tasks importer rejected my last TASK_PACK.txt. Fix and re-deliver.`r`n`r`n"
        . "IMPORT ERROR`r`n"
        . errorMsg . "`r`n`r`n"
    if (Trim(extraNotes) != "")
        body .= "EXTRA NOTES`r`n" . Trim(extraNotes) . "`r`n`r`n"
    body .= "WHAT YOU MUST DO`r`n"
        . guidance . "`r`n`r`n"
        . "DELIVERY RULES (mandatory)`r`n"
        . "- Deliver one complete TASK_PACK.txt (download chip preferred; else one marked code fence).`r`n"
        . "- Never claim you saved to Desktop / disk. I save the file myself.`r`n"
        . "- Pack must include ===PREVIEW=== … ===END_PREVIEW=== and FILE sections:`r`n"
        . "  ===FILE: TASK_PROJECTS.csv=== / TASK_TASKS.csv / TASK_INFO.csv ===END_FILE===`r`n"
        . "- FILE bodies are pure CSV with the exact headers from the Convert to Task prompt.`r`n"
        . "- filter = work|personal|habits; kind = punctual|habitual.`r`n"
        . "- Re-deliver using the exact canonical filename (TASK_PACK.txt). Overwrite any prior Desktop copy.`r`n"
        . "- Never add updated, corrected, v2, or similar suffixes to the filename.`r`n`r`n"
        . "After you fix it, I will save TASK_PACK.txt to Desktop and re-import.`r`n"
    path := A_Desktop . "\TASK_AI_FIX.txt"
    try {
        Task_WriteUtf8(path, body)
        return path
    } catch {
        return ""
    }
}

Task_AiCompanionFixGuidance(errorMsg) {
    e := StrLower(errorMsg)
    if (InStr(e, "no task_pack")) {
        return "- No pack file was found on Desktop.`r`n"
        . "- Re-deliver TASK_PACK.txt via download chip or one marked fence so I can save it and import."
    }
    if (InStr(e, "no usable") || InStr(e, "no data") || InStr(e, "no rows")) {
        return "- The pack had no usable task rows.`r`n"
        . "- Re-emit TASK_TASKS.csv with header:`r`n"
        . "  project_title,filter,title,emoji,kind,recurrence,due_date,next_due,section_path`r`n"
        . "- filter = work|personal|habits; kind = punctual|habitual; title required."
    }
    if (InStr(e, "invalid filter") || InStr(e, "invalid kind")) {
        return "- One or more rows used invalid filter/kind values.`r`n"
        . "- filter must be work|personal|habits; kind must be punctual|habitual."
    }
    return "- Read the IMPORT ERROR above and reframe as one complete, valid TASK_PACK.`r`n"
    . "- Prefer download chip; else one marked fence. Never claim a disk save."
}

Task_FailAiImport(errorMsg, notifyMs := 2200, extraNotes := "") {
    path := Task_WriteAiCompanionImportError(errorMsg, extraNotes)
    notify := errorMsg
    if (path != "")
        notify .= " — AI fix → Desktop TASK_AI_FIX.txt"
    Task_Notify(notify, notifyMs, BANNER_ACCENT_ERROR)
    return path != ""
}

Task_ArchiveImported(path) {
    destDir := Task_DataDir() . "\imported"
    if (!DirExist(destDir))
        DirCreate(destDir)
    SplitPath(path, &name)
    dest := destDir . "\" . FormatTime(, "yyyyMMdd-HHmmss") . "_" . name
    try FileMove(path, dest, 1)
    catch {
        try FileCopy(path, dest, 1)
        catch {
        }
    }
}

Task_DesktopNewestTaskPackCodeDump() {
    newest := ""
    newestTime := 0
    loop files A_Desktop . "\gemini-code*.txt", "F" {
        body := Task_ReadUtf8(A_LoopFileFullPath)
        if (body = "")
            continue
        lower := StrLower(body)
        hasPack := InStr(lower, "file: task_tasks") || InStr(lower, "file: task_projects")
        || InStr(lower, "project_title,filter,title")
        if (!hasPack)
            continue
        ts := Number(A_LoopFileTimeModified)
        if (ts > newestTime) {
            newestTime := ts
            newest := A_LoopFileFullPath
        }
    }
    return newest
}

Task_NormalizeFilter(raw) {
    f := StrLower(Trim(raw))
    if (f = "work" || f = "personal" || f = "habits")
        return f
    return ""
}

Task_NormalizeKind(raw) {
    k := StrLower(Trim(raw))
    if (k = "habitual")
        return "habitual"
    if (k = "" || k = "punctual")
        return "punctual"
    return ""
}

Task_ValidRecurrence(raw) {
    r := StrLower(Trim(raw))
    if (r = "")
        return true
    allowed := "daily|weekly|monthly|quarterly|biannual|yearly|every_2y|every_3y|every_5y|every_10y"
    return InStr("|" . allowed . "|", "|" . r . "|") > 0
}

Task_InboxTitleForFilter(filt) {
    switch filt {
        case "personal":
            return "Personal inbox"
        case "habits":
            return "Habits & Health"
        default:
            return "Work inbox"
    }
}

Task_FindProjectByTitleFilter(projects, title, filt) {
    want := StrLower(Trim(title))
    for p in projects {
        if (p["filter"] != filt)
            continue
        if (StrLower(Trim(p["title"])) = want)
            return p
    }
    return false
}

Task_NextIdAvoiding(prefix, existingRows, pendingRows, pad := 4) {
    maxN := 0
    for row in existingRows {
        id := row.Has("id") ? row["id"] : ""
        if (SubStr(id, 1, StrLen(prefix)) != prefix)
            continue
        rest := SubStr(id, StrLen(prefix) + 1)
        if (rest != "" && IsDigit(rest) && Integer(rest) > maxN)
            maxN := Integer(rest)
    }
    for row in pendingRows {
        id := row.Has("id") ? row["id"] : ""
        if (SubStr(id, 1, StrLen(prefix)) != prefix)
            continue
        rest := SubStr(id, StrLen(prefix) + 1)
        if (rest != "" && IsDigit(rest) && Integer(rest) > maxN)
            maxN := Integer(rest)
    }
    return prefix . Format("{:0" . pad . "d}", maxN + 1)
}

Task_EnsureProjectId(projects, newProjects, stagedProjByKey, projectTitle, filt) {
    title := Trim(projectTitle)
    if (title = "")
        title := Task_InboxTitleForFilter(filt)
    key := filt . "|" . StrLower(title)
    if (stagedProjByKey.Has(key))
        return stagedProjByKey[key]
    existing := Task_FindProjectByTitleFilter(projects, title, filt)
    if (IsObject(existing)) {
        stagedProjByKey[key] := existing["id"]
        return existing["id"]
    }
    row := Map(
        "id", Task_NextIdAvoiding("PROJ_", projects, newProjects),
        "title", title,
        "filter", filt,
        "section_path", "",
        "sort_order", Task_NextSortOrder(projects),
        "active", "1",
        "created_at", Task_NowStamp()
    )
    newProjects.Push(row)
    projects.Push(row)
    stagedProjByKey[key] := row["id"]
    return row["id"]
}

Task_ImportConfirm(title, previewRows) {
    global g_TaskGui
    owner := ""
    try {
        if (IsObject(g_TaskGui))
            owner := " +Owner" . g_TaskGui.Hwnd
    } catch {
        owner := ""
    }
    Task_DialogsBegin()
    g := Gui("+AlwaysOnTop +ToolWindow" . owner, title)
    g.SetFont("s10", "Segoe UI")
    g.Add("Text", "w780", "Confirm import of " . previewRows.Length . " row(s).")
    lv := g.Add("ListView", "w780 r14 Grid", ["Kind", "Filter", "Project / Parent", "Title"])
    for r in previewRows
        lv.Add("", r["kind"], r["filter"], r["project"], r["title"])
    lv.ModifyCol(1, 70)
    lv.ModifyCol(2, 80)
    lv.ModifyCol(3, 200)
    lv.ModifyCol(4, 400)
    ok := false
    g.Add("Button", "y+10 w120 Default", "Import").OnEvent("Click", ConfirmYes)
    g.Add("Button", "x+8 w120", "Cancel").OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())
    g.Show()
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    Task_DialogsEnd()
    return ok

    ConfirmYes(*) {
        ok := true
        g.Destroy()
    }
}

Task_ImportPackFromDesktop(*) {
    Task_ImportPackFromPath("")
}

Task_ImportPackFromPath(path := "", autoConfirm := false) {
    Task_EnsureData()
    if (path = "") {
        path := Task_DesktopNewest("TASK_PACK*.txt")
        if (path = "")
            path := Task_DesktopNewest("TASK_PACK*.csv")
        if (path = "")
            path := Task_DesktopNewestTaskPackCodeDump()
    }
    if (path = "" || !FileExist(path)) {
        Task_FailAiImport("No TASK_PACK file on Desktop", 2000)
        return false
    }
    path := PackImport_NormalizeDesktopSource(path, "TASK_PACK.txt")
    if (path = "" || !FileExist(path)) {
        Task_FailAiImport("No TASK_PACK file on Desktop", 2000)
        return false
    }
    sourcePath := path
    text := Task_ReadUtf8(path)
    if (text = "") {
        Task_FailAiImport("TASK_PACK is empty", 2000)
        return false
    }

    projBody := Task_ExtractPackFileSection(text, "TASK_PROJECTS.csv")
    taskBody := Task_ExtractPackFileSection(text, "TASK_TASKS.csv")
    infoBody := Task_ExtractPackFileSection(text, "TASK_INFO.csv")

    packProjects := projBody != "" ? Task_ParseCsvBody(projBody) : []
    packTasks := taskBody != "" ? Task_ParseCsvBody(taskBody) : []
    packInfos := infoBody != "" ? Task_ParseCsvBody(infoBody) : []

    if (!packTasks.Length && !packProjects.Length && !packInfos.Length) {
        Task_FailAiImport("File has no usable rows (need TASK_TASKS / TASK_PROJECTS / TASK_INFO)", 2400)
        return false
    }

    errors := []
    projects := Task_Load("projects")
    tasks := Task_Load("tasks")
    infos := Task_Load("info_points")

    ; Resolve / stage projects from pack
    stagedProjByKey := Map() ; filter|titleLower -> id
    newProjects := []
    for r in packProjects {
        title := Trim(r.Has("title") ? r["title"] : "")
        filt := Task_NormalizeFilter(r.Has("filter") ? r["filter"] : "")
        if (title = "") {
            errors.Push("PROJECT: missing title")
            continue
        }
        if (filt = "") {
            errors.Push("PROJECT invalid filter: " . title)
            continue
        }
        existing := Task_FindProjectByTitleFilter(projects, title, filt)
        if (IsObject(existing)) {
            stagedProjByKey[filt . "|" . StrLower(title)] := existing["id"]
            continue
        }
        ; also check newly staged
        key := filt . "|" . StrLower(title)
        if (stagedProjByKey.Has(key))
            continue
        row := Map(
            "id", Task_NextIdAvoiding("PROJ_", projects, newProjects),
            "title", title,
            "filter", filt,
            "section_path", Trim(r.Has("section_path") ? r["section_path"] : ""),
            "sort_order", Task_NextSortOrder(projects),
            "active", "1",
            "created_at", Task_NowStamp()
        )
        newProjects.Push(row)
        projects.Push(row)
        stagedProjByKey[key] := row["id"]
    }

    newTasks := []
    taskTitleIndex := Map() ; filter|titleLower -> id (latest)
    for t in tasks {
        if (t.Has("title") && t.Has("filter"))
            taskTitleIndex[t["filter"] . "|" . StrLower(Trim(t["title"]))] := t["id"]
    }

    for r in packTasks {
        title := Trim(r.Has("title") ? r["title"] : "")
        filt := Task_NormalizeFilter(r.Has("filter") ? r["filter"] : "work")
        kind := Task_NormalizeKind(r.Has("kind") ? r["kind"] : "punctual")
        recurrence := StrLower(Trim(r.Has("recurrence") ? r["recurrence"] : ""))
        if (title = "") {
            errors.Push("TASK: missing title")
            continue
        }
        if (filt = "") {
            errors.Push("TASK invalid filter: " . title)
            continue
        }
        if (kind = "") {
            errors.Push("TASK invalid kind: " . title)
            continue
        }
        if (!Task_ValidRecurrence(recurrence)) {
            errors.Push("TASK invalid recurrence (" . recurrence . "): " . title)
            continue
        }
        if (kind = "punctual")
            recurrence := ""
        emoji := Trim(r.Has("emoji") ? r["emoji"] : "")
        if (emoji = "")
            emoji := Task_DefaultEmoji()
        projTitle := Trim(r.Has("project_title") ? r["project_title"] : "")
        projId := Task_EnsureProjectId(projects, newProjects, stagedProjByKey, projTitle, filt)
        tid := Task_NextIdAvoiding("TASK_", tasks, newTasks)
        row := Map(
            "id", tid,
            "project_id", projId,
            "title", title,
            "emoji", emoji,
            "kind", kind,
            "recurrence", recurrence,
            "due_date", Trim(r.Has("due_date") ? r["due_date"] : ""),
            "next_due", Trim(r.Has("next_due") ? r["next_due"] : ""),
            "section_path", Trim(r.Has("section_path") ? r["section_path"] : ""),
            "filter", filt,
            "sort_order", Task_NextSortOrder(tasks),
            "completed_at", "",
            "created_at", Task_NowStamp(),
            "active", "1"
        )
        newTasks.Push(row)
        tasks.Push(row)
        taskTitleIndex[filt . "|" . StrLower(title)] := tid
    }

    newInfos := []
    for r in packInfos {
        attachTo := StrLower(Trim(r.Has("attach_to") ? r["attach_to"] : "task"))
        parentTitle := Trim(r.Has("parent_title") ? r["parent_title"] : "")
        filt := Task_NormalizeFilter(r.Has("filter") ? r["filter"] : "work")
        title := Trim(r.Has("title") ? r["title"] : "")
        body := r.Has("body") ? r["body"] : ""
        if (title = "" && Trim(body) = "") {
            errors.Push("INFO: empty title and body")
            continue
        }
        if (title = "")
            title := "Note"
        if (filt = "") {
            errors.Push("INFO invalid filter: " . title)
            continue
        }
        if (attachTo != "project" && attachTo != "task")
            attachTo := "task"
        parentId := ""
        if (attachTo = "project") {
            if (parentTitle = "")
                parentTitle := Task_InboxTitleForFilter(filt)
            parentId := Task_EnsureProjectId(projects, newProjects, stagedProjByKey, parentTitle, filt)
        } else {
            key := filt . "|" . StrLower(parentTitle)
            if (parentTitle != "" && taskTitleIndex.Has(key))
                parentId := taskTitleIndex[key]
            else if (newTasks.Length)
                parentId := newTasks[newTasks.Length]["id"]
            else {
                errors.Push("INFO: no parent task for " . title)
                continue
            }
        }
        iid := Task_NextIdAvoiding("INFO_", infos, newInfos)
        row := Map(
            "id", iid,
            "parent_type", attachTo,
            "parent_id", parentId,
            "title", title,
            "body", body,
            "emoji", Task_InfoEmoji(),
            "section_path", Trim(r.Has("section_path") ? r["section_path"] : ""),
            "sort_order", Task_NextSortOrder(infos),
            "created_at", Task_NowStamp()
        )
        newInfos.Push(row)
        infos.Push(row)
    }

    if (errors.Length && !newTasks.Length && !newProjects.Length && !newInfos.Length) {
        extra := ""
        for e in errors
            extra .= (extra = "" ? "" : "`r`n") . e
        Task_FailAiImport("No rows to import", 2400, extra)
        return false
    }

    preview := []
    for p in newProjects
        preview.Push(Map("kind", "project", "filter", p["filter"], "project", "", "title", p["title"]))
    for t in newTasks {
        pt := ""
        pr := Task_FindById(projects, t["project_id"])
        if (IsObject(pr))
            pt := pr["title"]
        preview.Push(Map("kind", t["kind"], "filter", t["filter"], "project", pt, "title", t["emoji"] . " " . t["title"
            ]))
    }
    for i in newInfos
        preview.Push(Map("kind", "info", "filter", "", "project", i["parent_type"] . ":" . i["parent_id"], "title", i[
            "title"]))

    if (!preview.Length) {
        Task_FailAiImport("No rows to import", 2200)
        return false
    }

    if (!autoConfirm && !Task_ImportConfirm("Import TASK_PACK", preview))
        return false

    ; Reload fresh and apply only new rows (projects already may include staged — save full arrays we mutated)
    ; We mutated projects/tasks/infos in memory including existing — safe to save.
    Task_Save("projects", projects)
    Task_Save("tasks", tasks)
    Task_Save("info_points", infos)
    Task_ArchiveImported(sourcePath)

    msg := "Imported " . newProjects.Length . " proj / " . newTasks.Length . " tasks / " . newInfos.Length . " info"
    if (errors.Length)
        msg .= " (" . errors.Length . " skipped)"
    Task_Notify(msg, 2800, BANNER_ACCENT_SUCCESS)
    try Task_ShowMainMenu()
    catch {
    }
    return true
}
