; =============================================================================
; Utils module: task_projects.ahk
; Projects ListView CRUD
; =============================================================================

Task_ShowProjects() {
    global g_TaskGui, g_TaskLv, g_TaskBrowseProjectId, g_TaskBrowseTaskId
    Task_CloseGui()
    Task_EnsureData()
    g_TaskBrowseProjectId := ""
    g_TaskBrowseTaskId := ""
    g_TaskGui := Gui("+AlwaysOnTop +ToolWindow", "Tasks — Projects")
    g_TaskGui.SetFont("s10", "Segoe UI")
    lvY := Task_AddBrowseChrome(g_TaskGui, "Projects")
    g_TaskLv := g_TaskGui.Add("ListView", "x12 y" . lvY . " w860 h440 Grid Background2D2D30",
        ["Title", "Filter", "Section", "Order", "Active"])
    Task_StyleDarkListView(g_TaskLv)
    g_TaskLv.OnEvent("DoubleClick", (*) => Task_ProjectDrill())
    g_TaskGui.OnEvent("Close", (*) => Task_CloseGui())
    g_TaskGui.OnEvent("Escape", (*) => Task_ShowMainMenu())
    Task_ProjectRefresh()
    Task_BindHotkeys([
        ["a", (*) => Task_ProjectAdd()],
        ["Insert", (*) => Task_ProjectAdd()],
        ["e", (*) => Task_ProjectEdit()],
        ["Delete", (*) => Task_ProjectDelete()],
        ["n", (*) => Task_ProjectShowInfo()],
        ["v", (*) => Task_ProjectPasteImage()],
        ["Enter", (*) => Task_ProjectDrill()],
        ["1", (*) => Task_BrowseSetFilter("work")],
        ["2", (*) => Task_BrowseSetFilter("personal")],
        ["3", (*) => Task_BrowseSetFilter("habits")],
        ["Backspace", (*) => Task_ShowMainMenu()],
        ["Escape", (*) => Task_ShowMainMenu()]
    ])
    Task_LetterJumpStart((entry) => entry["title"])
    Task_CenterGui(g_TaskGui, 890, 540)
}

Task_BrowseSetFilter(filt) {
    Task_SetFilter(filt)
    Task_Notify("Filter: " . Task_FilterLabel(filt), 900, BANNER_ACCENT_INTERMEDIATE)
    global g_TaskBrowseProjectId, g_TaskBrowseTaskId
    if (g_TaskBrowseProjectId = "" && g_TaskBrowseTaskId = "")
        Task_ShowProjects()
    else if (g_TaskBrowseTaskId = "")
        Task_ShowTasksForProject()
    else
        Task_ShowInfoForParent("task", g_TaskBrowseTaskId)
}

Task_ProjectRefresh() {
    global g_TaskLv, g_TaskRows
    if (!IsObject(g_TaskLv))
        return
    rows := Task_Load("projects")
    g_TaskLv.Delete()
    g_TaskRows := []
    for s in rows {
        if (!Task_MatchesFilter(s))
            continue
        if (s.Has("active") && s["active"] = "0")
            continue
        g_TaskRows.Push(s)
        g_TaskLv.Add("", s["title"], s["filter"], s["section_path"], s["sort_order"], s["active"])
    }
    loop 5
        g_TaskLv.ModifyCol(A_Index, "AutoHdr")
    Task_StyleDarkListView(g_TaskLv)
}

Task_ProjectSelected() {
    global g_TaskLv, g_TaskRows
    row := g_TaskLv.GetNext()
    if (!row || row > g_TaskRows.Length)
        return false
    return g_TaskRows[row]
}

Task_ProjectDrill(*) {
    global g_TaskBrowseProjectId
    p := Task_ProjectSelected()
    if (!p) {
        Task_Notify("Select a project", 1200, BANNER_ACCENT_ERROR)
        return
    }
    g_TaskBrowseProjectId := p["id"]
    Task_SetSetting("General", "LastProjectId", p["id"])
    Task_ShowTasksForProject()
}

Task_ProjectAdd(*) {
    Task_ProjectForm(false)
}
Task_ProjectEdit(*) {
    p := Task_ProjectSelected()
    if (!p) {
        Task_Notify("Select a project", 1200, BANNER_ACCENT_ERROR)
        return
    }
    Task_ProjectForm(p)
}

Task_ProjectDelete(*) {
    p := Task_ProjectSelected()
    if (!p)
        return
    taskIds := Map()
    taskCount := 0
    for t in Task_Load("tasks") {
        if (t["project_id"] = p["id"]) {
            taskIds[t["id"]] := true
            taskCount += 1
        }
    }
    infoCount := 0
    for i in Task_Load("info_points") {
        if (i["parent_type"] = "project" && i["parent_id"] = p["id"])
            infoCount += 1
        else if (i["parent_type"] = "task" && taskIds.Has(i["parent_id"]))
            infoCount += 1
    }
    attCount := 0
    for a in Task_Load("attachments") {
        if (a["parent_type"] = "project" && a["parent_id"] = p["id"])
            attCount += 1
        else if (a["parent_type"] = "task" && taskIds.Has(a["parent_id"]))
            attCount += 1
        else if (a["parent_type"] = "info") {
            ; counted via info cascade below only if we track — skip for summary
        }
    }
    msg := "Delete project " . p["title"] . "?"
    if (taskCount || infoCount || attCount)
        msg .= "`nAlso deletes " . taskCount . " task(s), " . infoCount . " info, attachments."
    if (!Task_Confirm(msg, "Projects"))
        return
    Task_CascadeDeleteProject(p["id"])
    Task_ProjectRefresh()
    Task_Notify("Project removed", 1200, BANNER_ACCENT_SUCCESS)
}

Task_CascadeDeleteProject(projectId) {
    taskIds := Map()
    for t in Task_Load("tasks") {
        if (t["project_id"] = projectId)
            taskIds[t["id"]] := true
    }
    infoIds := Map()
    infoOut := []
    for i in Task_Load("info_points") {
        drop := false
        if (i["parent_type"] = "project" && i["parent_id"] = projectId)
            drop := true
        else if (i["parent_type"] = "task" && taskIds.Has(i["parent_id"]))
            drop := true
        if (drop)
            infoIds[i["id"]] := true
        else
            infoOut.Push(i)
    }
    Task_Save("info_points", infoOut)
    Task_PurgeAttachments((a) => (
        (a["parent_type"] = "project" && a["parent_id"] = projectId)
        || (a["parent_type"] = "task" && taskIds.Has(a["parent_id"]))
        || (a["parent_type"] = "info" && infoIds.Has(a["parent_id"]))
    ))
    taskOut := []
    for t in Task_Load("tasks") {
        if (t["project_id"] != projectId)
            taskOut.Push(t)
    }
    Task_Save("tasks", taskOut)
    projOut := []
    for p in Task_Load("projects") {
        if (p["id"] != projectId)
            projOut.Push(p)
    }
    Task_Save("projects", projOut)
}

Task_ProjectForm(existing) {
    global g_TaskGui, g_TaskFilter
    isEdit := IsObject(existing)
    owner := ""
    try {
        if (IsObject(g_TaskGui))
            owner := " +Owner" . g_TaskGui.Hwnd
    } catch {
        owner := ""
    }
    Task_DialogsBegin()
    g := Gui("+AlwaysOnTop +ToolWindow" . owner, isEdit ? "Edit project" : "Add project")
    g.SetFont("s10", "Segoe UI")
    g.Add("Text", , "Title")
    eTitle := g.Add("Edit", "w360", isEdit ? existing["title"] : "")
    g.Add("Text", "y+8", "Filter (work / personal / habits)")
    defFilt := isEdit ? existing["filter"] : (g_TaskFilter = "all" || g_TaskFilter = "" ? "work" : g_TaskFilter)
    eFilt := g.Add("Edit", "w360", defFilt)
    g.Add("Text", "y+8", "Section path (optional)")
    eSec := g.Add("Edit", "w360", isEdit ? existing["section_path"] : "")
    chk := g.Add("CheckBox", "y+8 Checked" . (isEdit ? (existing["active"] = "0" ? "0" : "1") : "1"), "Active")
    saved := false
    g.Add("Button", "y+16 w100 Default", "Save").OnEvent("Click", SaveProj)
    g.Add("Button", "x+8 w100", "Cancel").OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())
    g.Show()
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    Task_DialogsEnd()
    if (saved)
        Task_ProjectRefresh()

    SaveProj(*) {
        title := Trim(eTitle.Value)
        if (title = "") {
            Task_Alert("Title is required.", "Projects")
            return
        }
        filt := StrLower(Trim(eFilt.Value))
        if (filt != "work" && filt != "personal" && filt != "habits")
            filt := "work"
        projects := Task_Load("projects")
        row := Map(
            "id", isEdit ? existing["id"] : Task_NextId("PROJ_", projects),
        "title", title,
        "filter", filt,
        "section_path", Trim(eSec.Value),
        "sort_order", isEdit ? existing["sort_order"] : Task_NextSortOrder(projects),
        "active", chk.Value ? "1" : "0",
        "created_at", isEdit ? existing["created_at"] : Task_NowStamp())
        if (isEdit) {
            out := []
            for r in projects {
                if (r["id"] = existing["id"])
                    out.Push(row)
                else
                    out.Push(r)
            }
            projects := out
        } else {
            projects.Push(row)
        }
        Task_Save("projects", projects)
        saved := true
        g.Destroy()
        Task_Notify(isEdit ? "Project updated" : "Project saved", 1200, BANNER_ACCENT_SUCCESS)
    }
}

Task_ProjectShowInfo(*) {
    global g_TaskBrowseProjectId
    p := Task_ProjectSelected()
    if (!p) {
        Task_Notify("Select a project", 1200, BANNER_ACCENT_ERROR)
        return
    }
    g_TaskBrowseProjectId := p["id"]
    Task_ShowInfoForParent("project", p["id"])
}

Task_ProjectPasteImage(*) {
    p := Task_ProjectSelected()
    if (!p) {
        Task_Notify("Select a project", 1200, BANNER_ACCENT_ERROR)
        return
    }
    Task_PasteClipboardAttachment("project", p["id"])
}
