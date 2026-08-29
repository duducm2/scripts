; =============================================================================
; Utils module: task_tasks.ahk
; Tasks ListView CRUD, inbox, habits view, emoji actions
; =============================================================================

Task_ShowTasksForProject() {
    global g_TaskGui, g_TaskLv, g_TaskBrowseProjectId, g_TaskBrowseTaskId
    if (g_TaskBrowseProjectId = "") {
        Task_ShowProjects()
        return
    }
    Task_CloseGui()
    Task_EnsureData()
    g_TaskBrowseTaskId := ""
    g_TaskGui := Gui("+AlwaysOnTop +ToolWindow", "Tasks — Tasks")
    g_TaskGui.SetFont("s10", "Segoe UI")
    lvY := Task_AddBrowseChrome(g_TaskGui, "Tasks")
    g_TaskLv := g_TaskGui.Add("ListView", "x12 y" . lvY . " w860 h440 Grid Background2D2D30",
        ["", "Title", "Kind", "Recur", "Due/Next", "Section"])
    Task_StyleDarkListView(g_TaskLv)
    g_TaskLv.OnEvent("DoubleClick", (*) => Task_TaskDrill())
    g_TaskGui.OnEvent("Close", (*) => Task_CloseGui())
    g_TaskGui.OnEvent("Escape", (*) => Task_ShowProjects())
    Task_TaskRefresh(g_TaskBrowseProjectId, false, false)
    Task_BindTaskListHotkeys(() => Task_ShowProjects())
    Task_LetterJumpStart((entry) => entry["title"])
    Task_CenterGui(g_TaskGui, 890, 540)
}

Task_ShowTasksInbox(habitsOnly) {
    global g_TaskGui, g_TaskLv, g_TaskBrowseProjectId, g_TaskBrowseTaskId
    Task_CloseGui()
    Task_EnsureData()
    g_TaskBrowseProjectId := ""
    g_TaskBrowseTaskId := ""
    title := habitsOnly ? "Tasks — Habits due" : "Tasks — Inbox"
    g_TaskGui := Gui("+AlwaysOnTop +ToolWindow", title)
    g_TaskGui.SetFont("s10", "Segoe UI")
    lvY := Task_AddBrowseChrome(g_TaskGui, habitsOnly ? "Habits" : "Inbox")
    g_TaskLv := g_TaskGui.Add("ListView", "x12 y" . lvY . " w860 h440 Grid Background2D2D30",
        ["", "Title", "Project", "Kind", "Recur", "Due/Next", "Filter"])
    Task_StyleDarkListView(g_TaskLv)
    g_TaskLv.OnEvent("DoubleClick", (*) => Task_TaskDrill())
    g_TaskGui.OnEvent("Close", (*) => Task_CloseGui())
    g_TaskGui.OnEvent("Escape", (*) => Task_ShowMainMenu())
    Task_TaskRefresh("", true, habitsOnly)
    Task_BindTaskListHotkeys(() => Task_ShowMainMenu())
    Task_LetterJumpStart((entry) => entry["title"])
    Task_CenterGui(g_TaskGui, 890, 540)
}

Task_BindTaskListHotkeys(backFn) {
    Task_BindHotkeys([
        ["a", (*) => Task_TaskAdd()],
        ["Insert", (*) => Task_TaskAdd()],
        ["e", (*) => Task_TaskEdit()],
        ["Delete", (*) => Task_TaskDelete()],
        ["c", (*) => Task_TaskSetEmoji("done")],
        ["w", (*) => Task_TaskSetEmoji("waiting")],
        ["i", (*) => Task_TaskSetEmoji("important")],
        ["u", (*) => Task_TaskSetEmoji("doubt")],
        ["g", (*) => Task_TaskSetEmoji("general")],
        ["x", (*) => Task_TaskCustomEmoji()],
        ["n", (*) => Task_TaskShowInfo()],
        ["v", (*) => Task_TaskPasteImage()],
        ["Enter", (*) => Task_TaskDrill()],
        ["f", (*) => Task_OnFilterFromBrowse()],
        ["Backspace", backFn],
        ["Escape", backFn]
    ])
}

Task_TaskRefresh(projectId, inboxMode, habitsOnly) {
    global g_TaskLv, g_TaskRows
    if (!IsObject(g_TaskLv))
        return
    projects := Task_Load("projects")
    projTitle := Map()
    for p in projects
        projTitle[p["id"]] := p["title"]
    rows := Task_Load("tasks")
    g_TaskLv.Delete()
    g_TaskRows := []
    for t in rows {
        if (!Task_MatchesFilter(t))
            continue
        if (t.Has("active") && t["active"] = "0")
            continue
        if (projectId != "" && t["project_id"] != projectId)
            continue
        if (inboxMode && !Task_IsOpenEmoji(t["emoji"]))
            continue
        if (habitsOnly && t["kind"] != "habitual")
            continue
        g_TaskRows.Push(t)
        due := Trim(t["next_due"]) != "" ? t["next_due"] : t["due_date"]
        if (inboxMode) {
            pt := projTitle.Has(t["project_id"]) ? projTitle[t["project_id"]] : t["project_id"]
            g_TaskLv.Add("", t["emoji"], t["title"], pt, t["kind"], t["recurrence"], due, t["filter"])
        } else {
            g_TaskLv.Add("", t["emoji"], t["title"], t["kind"], t["recurrence"], due, t["section_path"])
        }
    }
    colN := inboxMode ? 7 : 6
    loop colN
        g_TaskLv.ModifyCol(A_Index, "AutoHdr")
    Task_StyleDarkListView(g_TaskLv)
}

Task_TaskSelected() {
    global g_TaskLv, g_TaskRows
    row := g_TaskLv.GetNext()
    if (!row || row > g_TaskRows.Length)
        return false
    return g_TaskRows[row]
}

Task_TaskReloadCurrentView() {
    global g_TaskBrowseProjectId
    if (g_TaskBrowseProjectId != "")
        Task_ShowTasksForProject()
    else {
        title := ""
        try title := WinGetTitle("A")
        catch {
        }
        Task_ShowTasksInbox(InStr(title, "Habits") > 0)
    }
}

Task_TaskDrill(*) {
    global g_TaskBrowseTaskId, g_TaskBrowseProjectId
    t := Task_TaskSelected()
    if (!t) {
        Task_Notify("Select a task", 1200, BANNER_ACCENT_ERROR)
        return
    }
    g_TaskBrowseTaskId := t["id"]
    if (g_TaskBrowseProjectId = "")
        g_TaskBrowseProjectId := t["project_id"]
    Task_ShowInfoForParent("task", t["id"])
}

Task_TaskShowInfo(*) {
    Task_TaskDrill()
}

Task_TaskPasteImage(*) {
    t := Task_TaskSelected()
    if (!t) {
        Task_Notify("Select a task", 1200, BANNER_ACCENT_ERROR)
        return
    }
    Task_PasteClipboardAttachment("task", t["id"])
}

Task_TaskSetEmoji(key) {
    t := Task_TaskSelected()
    if (!t) {
        Task_Notify("Select a task", 1200, BANNER_ACCENT_ERROR)
        return
    }
    emojis := Task_StatusEmojis()
    if (key = "done") {
        t := Task_CompleteTask(t)
    } else {
        t["emoji"] := emojis[key]
        if (key != "done")
            t["completed_at"] := ""
    }
    Task_UpsertTask(t)
    Task_TaskReloadCurrentView()
    Task_Notify(t["emoji"] . " updated", 900, BANNER_ACCENT_SUCCESS)
}

Task_TaskCustomEmoji(*) {
    t := Task_TaskSelected()
    if (!t) {
        Task_Notify("Select a task", 1200, BANNER_ACCENT_ERROR)
        return
    }
    res := Task_InputBox("Custom emoji for this task:", "Custom emoji", t["emoji"])
    if (res.Result != "OK")
        return
    em := Trim(res.Value)
    if (em = "")
        return
    t["emoji"] := em
    Task_UpsertTask(t)
    Task_TaskReloadCurrentView()
    Task_Notify(em . " set", 900, BANNER_ACCENT_SUCCESS)
}

Task_UpsertTask(row) {
    tasks := Task_Load("tasks")
    out := []
    found := false
    for t in tasks {
        if (t["id"] = row["id"]) {
            out.Push(row)
            found := true
        } else {
            out.Push(t)
        }
    }
    if (!found)
        out.Push(row)
    Task_Save("tasks", out)
}

Task_TaskAdd(*) {
    global g_TaskBrowseProjectId
    if (g_TaskBrowseProjectId = "") {
        ; Pick project first
        projects := []
        for p in Task_Load("projects") {
            if (Task_MatchesFilter(p) && !(p.Has("active") && p["active"] = "0"))
                projects.Push(p)
        }
        if (!projects.Length) {
            Task_Alert("Create a project first.", "Tasks")
            return
        }
        pick := Task_PickProject(projects)
        if (pick = "")
            return
        g_TaskBrowseProjectId := pick
    }
    Task_TaskForm(false)
}

Task_PickProject(projects) {
    Task_DialogsBegin()
    names := ""
    for i, p in projects
        names .= i . " — " . p["title"] . " [" . p["filter"] . "]`n"
    res := Task_InputBox("Enter project number:`n" . names, "Pick project", "1")
    Task_DialogsEnd()
    if (res.Result != "OK")
        return ""
    n := Integer(Trim(res.Value) || 0)
    if (n < 1 || n > projects.Length)
        return ""
    return projects[n]["id"]
}

Task_TaskEdit(*) {
    t := Task_TaskSelected()
    if (!t) {
        Task_Notify("Select a task", 1200, BANNER_ACCENT_ERROR)
        return
    }
    Task_TaskForm(t)
}

Task_TaskDelete(*) {
    t := Task_TaskSelected()
    if (!t)
        return
    infoCount := 0
    for i in Task_Load("info_points") {
        if (i["parent_type"] = "task" && i["parent_id"] = t["id"])
            infoCount += 1
    }
    msg := "Delete task " . t["title"] . "?"
    if (infoCount)
        msg .= "`nAlso deletes " . infoCount . " info point(s) and attachments."
    if (!Task_Confirm(msg, "Tasks"))
        return
    Task_CascadeDeleteTask(t["id"])
    Task_TaskReloadCurrentView()
    Task_Notify("Task removed", 1200, BANNER_ACCENT_SUCCESS)
}

Task_CascadeDeleteTask(taskId) {
    infoIds := Map()
    infoOut := []
    for i in Task_Load("info_points") {
        if (i["parent_type"] = "task" && i["parent_id"] = taskId)
            infoIds[i["id"]] := true
        else
            infoOut.Push(i)
    }
    Task_Save("info_points", infoOut)
    attOut := []
    for a in Task_Load("attachments") {
        drop := (a["parent_type"] = "task" && a["parent_id"] = taskId)
        || (a["parent_type"] = "info" && infoIds.Has(a["parent_id"]))
        if (!drop)
            attOut.Push(a)
    }
    Task_Save("attachments", attOut)
    taskOut := []
    for t in Task_Load("tasks") {
        if (t["id"] != taskId)
            taskOut.Push(t)
    }
    Task_Save("tasks", taskOut)
}

Task_TaskForm(existing) {
    global g_TaskGui, g_TaskBrowseProjectId, g_TaskFilter
    isEdit := IsObject(existing)
    owner := ""
    try {
        if (IsObject(g_TaskGui))
            owner := " +Owner" . g_TaskGui.Hwnd
    } catch {
        owner := ""
    }
    projId := isEdit ? existing["project_id"] : g_TaskBrowseProjectId
    proj := Task_FindById(Task_Load("projects"), projId)
    defFilt := IsObject(proj) ? proj["filter"] : (g_TaskFilter = "all" ? "work" : g_TaskFilter)

    Task_DialogsBegin()
    g := Gui("+AlwaysOnTop +ToolWindow" . owner, isEdit ? "Edit task" : "Add task")
    g.SetFont("s10", "Segoe UI")
    g.Add("Text", , "Title")
    eTitle := g.Add("Edit", "w400", isEdit ? existing["title"] : "")
    g.Add("Text", "y+8", "Emoji")
    eEmoji := g.Add("Edit", "w80", isEdit ? existing["emoji"] : Task_DefaultEmoji())
    g.Add("Text", "y+8", "Kind (punctual / habitual)")
    eKind := g.Add("Edit", "w200", isEdit ? existing["kind"] : "punctual")
    g.Add("Text", "y+8", "Recurrence (empty if punctual)")
    eRec := g.Add("Edit", "w200", isEdit ? existing["recurrence"] : "")
    g.Add("Text", "y+8", "Due date (optional yyyy-MM-dd)")
    eDue := g.Add("Edit", "w200", isEdit ? existing["due_date"] : "")
    g.Add("Text", "y+8", "Next due (habitual)")
    eNext := g.Add("Edit", "w200", isEdit ? existing["next_due"] : "")
    g.Add("Text", "y+8", "Section path")
    eSec := g.Add("Edit", "w400", isEdit ? existing["section_path"] : "")
    g.Add("Text", "y+8", "Filter")
    eFilt := g.Add("Edit", "w200", isEdit ? existing["filter"] : defFilt)
    chk := g.Add("CheckBox", "y+8 Checked" . (isEdit ? (existing["active"] = "0" ? "0" : "1") : "1"), "Active")
    saved := false
    g.Add("Button", "y+16 w100 Default", "Save").OnEvent("Click", SaveTask)
    g.Add("Button", "x+8 w100", "Cancel").OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())
    g.Show()
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    Task_DialogsEnd()
    if (saved)
        Task_TaskReloadCurrentView()

    SaveTask(*) {
        title := Trim(eTitle.Value)
        if (title = "") {
            Task_Alert("Title is required.", "Tasks")
            return
        }
        kind := StrLower(Trim(eKind.Value))
        if (kind != "habitual")
            kind := "punctual"
        filt := StrLower(Trim(eFilt.Value))
        if (filt != "work" && filt != "personal" && filt != "habits")
            filt := defFilt
        emoji := Trim(eEmoji.Value)
        if (emoji = "")
            emoji := Task_DefaultEmoji()
        tasks := Task_Load("tasks")
        row := Map(
            "id", isEdit ? existing["id"] : Task_NextId("TASK_", tasks),
        "project_id", projId,
        "title", title,
        "emoji", emoji,
        "kind", kind,
        "recurrence", Trim(eRec.Value),
        "due_date", Trim(eDue.Value),
        "next_due", Trim(eNext.Value),
        "section_path", Trim(eSec.Value),
        "filter", filt,
        "sort_order", isEdit ? existing["sort_order"] : Task_NextSortOrder(tasks),
        "completed_at", isEdit ? existing["completed_at"] : "",
        "created_at", isEdit ? existing["created_at"] : Task_NowStamp(),
        "active", chk.Value ? "1" : "0"
        )
        Task_UpsertTask(row)
        saved := true
        g.Destroy()
        Task_Notify(isEdit ? "Task updated" : "Task saved", 1200, BANNER_ACCENT_SUCCESS)
    }
}
