; =============================================================================
; Utils module: task_info.ahk
; Info points ListView, large text editor, clipboard image attachments
; =============================================================================

Task_ShowInfoForParent(parentType, parentId) {
    global g_TaskGui, g_TaskLv, g_TaskBrowseProjectId, g_TaskBrowseTaskId
    Task_CloseGui()
    Task_EnsureData()
    if (parentType = "task")
        g_TaskBrowseTaskId := parentId
    else
        g_TaskBrowseTaskId := ""

    label := parentType = "task" ? "Info (task)" : "Info (project)"
    g_TaskGui := Gui("+AlwaysOnTop +ToolWindow", "Tasks — " . label)
    g_TaskGui.SetFont("s10", "Segoe UI")
    lvY := Task_AddBrowseChrome(g_TaskGui, label)
    g_TaskLv := g_TaskGui.Add("ListView", "x12 y" . lvY . " w860 h400 Grid Background2D2D30",
        ["", "Title", "Section", "Preview"])
    Task_StyleDarkListView(g_TaskLv)
    g_TaskLv.OnEvent("DoubleClick", (*) => Task_InfoOpenEditor())
    g_TaskGui.OnEvent("Close", (*) => Task_CloseGui())
    backFn := parentType = "task"
        ? (*) => Task_ShowTasksForProjectOrInbox()
            : (*) => Task_ShowProjects()
    g_TaskGui.OnEvent("Escape", backFn)
    Task_InfoRefresh(parentType, parentId)
    Task_BindHotkeys([
        ["a", (*) => Task_InfoAdd(parentType, parentId)],
        ["Insert", (*) => Task_InfoAdd(parentType, parentId)],
        ["e", (*) => Task_InfoOpenEditor()],
        ["Enter", (*) => Task_InfoOpenEditor()],
        ["Delete", (*) => Task_InfoDelete(parentType, parentId)],
        ["v", (*) => Task_InfoPasteImage(parentType, parentId)],
        ["o", (*) => Task_InfoOpenAttachments(parentType, parentId)],
        ["Backspace", backFn],
        ["Escape", backFn]
    ])
    Task_LetterJumpStart((entry) => entry["title"])
    Task_CenterGui(g_TaskGui, 890, 520)
}

Task_ShowTasksForProjectOrInbox() {
    global g_TaskBrowseProjectId
    if (g_TaskBrowseProjectId != "")
        Task_ShowTasksForProject()
    else
        Task_ShowProjects()
}

Task_InfoRefresh(parentType, parentId) {
    global g_TaskLv, g_TaskRows
    if (!IsObject(g_TaskLv))
        return
    g_TaskLv.Delete()
    g_TaskRows := []
    for i in Task_Load("info_points") {
        if (i["parent_type"] != parentType || i["parent_id"] != parentId)
            continue
        g_TaskRows.Push(i)
        preview := i["body"]
        preview := StrReplace(preview, "`n", " ")
        if (StrLen(preview) > 80)
            preview := SubStr(preview, 1, 77) . "…"
        em := Trim(i["emoji"]) != "" ? i["emoji"] : Task_InfoEmoji()
        g_TaskLv.Add("", em, i["title"], i["section_path"], preview)
    }
    loop 4
        g_TaskLv.ModifyCol(A_Index, "AutoHdr")
    Task_StyleDarkListView(g_TaskLv)
}

Task_InfoSelected() {
    global g_TaskLv, g_TaskRows
    row := g_TaskLv.GetNext()
    if (!row || row > g_TaskRows.Length)
        return false
    return g_TaskRows[row]
}

Task_InfoAdd(parentType, parentId) {
    Task_InfoForm(false, parentType, parentId)
}

Task_InfoDelete(parentType, parentId) {
    info := Task_InfoSelected()
    if (!info)
        return
    if (!Task_Confirm("Delete info point " . info["title"] . "?", "Info"))
        return
    out := []
    for i in Task_Load("info_points") {
        if (i["id"] != info["id"])
            out.Push(i)
    }
    Task_Save("info_points", out)
    infoId := info["id"]
    Task_PurgeAttachments((a) => a["parent_type"] = "info" && a["parent_id"] = infoId)
    Task_InfoRefresh(parentType, parentId)
    Task_Notify("Info removed", 1200, BANNER_ACCENT_SUCCESS)
}

Task_InfoPasteImage(parentType, parentId) {
    info := Task_InfoSelected()
    if (IsObject(info))
        Task_PasteClipboardAttachment("info", info["id"])
    else
        Task_PasteClipboardAttachment(parentType, parentId)
}

Task_InfoOpenAttachments(parentType, parentId) {
    info := Task_InfoSelected()
    pid := IsObject(info) ? info["id"] : parentId
    ptype := IsObject(info) ? "info" : parentType
    lines := ""
    n := 0
    for a in Task_Load("attachments") {
        if (a["parent_type"] = ptype && a["parent_id"] = pid) {
            n += 1
            lines .= n . ". [" . a["kind"] . "] " . a["description"] . " — " . a["ref"] . "`n"
        }
    }
    if (n = 0) {
        Task_Notify("No attachments", 1200, BANNER_ACCENT_INTERMEDIATE)
        return
    }
    Task_Alert(lines, "Attachments")
}

Task_InfoForm(existing, parentType, parentId) {
    global g_TaskGui
    isEdit := IsObject(existing)
    owner := ""
    try {
        if (IsObject(g_TaskGui))
            owner := " +Owner" . g_TaskGui.Hwnd
    } catch {
        owner := ""
    }
    Task_DialogsBegin()
    g := Gui("+AlwaysOnTop +ToolWindow" . owner, isEdit ? "Edit info" : "Add info")
    g.SetFont("s10", "Segoe UI")
    g.Add("Text", , "Title")
    eTitle := g.Add("Edit", "w420", isEdit ? existing["title"] : "")
    g.Add("Text", "y+8", "Section path")
    eSec := g.Add("Edit", "w420", isEdit ? existing["section_path"] : "")
    g.Add("Text", "y+8", "Emoji")
    eEmoji := g.Add("Edit", "w80", isEdit ? existing["emoji"] : Task_InfoEmoji())
    saved := false
    g.Add("Button", "y+16 w100 Default", "Save").OnEvent("Click", SaveInfo)
    g.Add("Button", "x+8 w100", "Cancel").OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())
    g.Show()
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    Task_DialogsEnd()
    if (saved) {
        Task_InfoRefresh(parentType, parentId)
        if (!isEdit) {
            ; open editor for body
            infos := Task_Load("info_points")
            newest := infos[infos.Length]
            Task_InfoOpenEditorFor(newest, parentType, parentId)
        }
    }

    SaveInfo(*) {
        title := Trim(eTitle.Value)
        if (title = "") {
            Task_Alert("Title is required.", "Info")
            return
        }
        infos := Task_Load("info_points")
        row := Map(
            "id", isEdit ? existing["id"] : Task_NextId("INFO_", infos),
        "parent_type", parentType,
        "parent_id", parentId,
        "title", title,
        "body", isEdit ? existing["body"] : "",
        "emoji", Trim(eEmoji.Value) != "" ? Trim(eEmoji.Value) : Task_InfoEmoji(),
        "section_path", Trim(eSec.Value),
        "sort_order", isEdit ? existing["sort_order"] : Task_NextSortOrder(infos),
        "created_at", isEdit ? existing["created_at"] : Task_NowStamp())
        if (isEdit) {
            out := []
            for r in infos {
                if (r["id"] = existing["id"])
                    out.Push(row)
                else
                    out.Push(r)
            }
            infos := out
        } else {
            infos.Push(row)
        }
        Task_Save("info_points", infos)
        saved := true
        g.Destroy()
        Task_Notify(isEdit ? "Info updated" : "Info saved", 1000, BANNER_ACCENT_SUCCESS)
    }
}

Task_InfoOpenEditor(*) {
    info := Task_InfoSelected()
    if (!info) {
        Task_Notify("Select an info point", 1200, BANNER_ACCENT_ERROR)
        return
    }
    Task_InfoOpenEditorFor(info, info["parent_type"], info["parent_id"])
}

Task_InfoOpenEditorFor(info, parentType, parentId) {
    global g_TaskGui
    owner := ""
    try {
        if (IsObject(g_TaskGui))
            owner := " +Owner" . g_TaskGui.Hwnd
    } catch {
        owner := ""
    }
    Task_DialogsBegin()
    g := Gui("+AlwaysOnTop +Resize" . owner, "Info — " . info["title"])
    g.BackColor := "1E1E1E"
    g.SetFont("s11 cWhite", "Consolas")
    g.Add("Text", "x12 y10 w760 cF1C40F", info["emoji"] . "  " . info["title"])
    edit := g.Add("Edit", "x12 y36 w760 h420 Multi WantReturn VScroll Background2D2D30 cWhite", info["body"])
    saved := false
    g.Add("Button", "x12 y470 w100 Default", "Save").OnEvent("Click", SaveBody)
    g.Add("Button", "x+8 w100", "Cancel").OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())
    g.OnEvent("Size", (*) => ResizeEd())
    Task_CenterGui(g, 800, 540)
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    Task_DialogsEnd()
    if (saved)
        Task_InfoRefresh(parentType, parentId)

    ResizeEd(*) {
        try {
            g.GetPos(, , &gw, &gh)
            edit.Move(12, 36, gw - 40, gh - 120)
        } catch {
        }
    }

    SaveBody(*) {
        infos := Task_Load("info_points")
        out := []
        for r in infos {
            if (r["id"] = info["id"]) {
                r["body"] := edit.Value
                out.Push(r)
            } else {
                out.Push(r)
            }
        }
        Task_Save("info_points", out)
        saved := true
        g.Destroy()
        Task_Notify("Info body saved", 1000, BANNER_ACCENT_SUCCESS)
    }
}

; --- Attachments / clipboard paste ---

Task_PasteClipboardAttachment(parentType, parentId) {
    clipText := ""
    try clipText := A_Clipboard
    catch {
        clipText := ""
    }
    clipText := Trim(clipText)
    hasImg := Task_ClipboardHasImage()

    if (hasImg) {
        desc := clipText
        if (desc = "") {
            res := Task_InputBox("Description for this clipboard image:", "Image description", "")
            if (res.Result != "OK")
                return
            desc := Trim(res.Value)
            if (desc = "") {
                Task_Alert("A description is required when pasting an image without text.", "Attachments")
                return
            }
        }
        fname := FormatTime(, "yyyyMMdd-HHmmss") . "-" . parentId . ".png"
        dest := Task_AttachmentsDir() . "\" . fname
        if (!Task_SaveClipboardImage(dest)) {
            Task_Notify("Could not save clipboard image", 2000, BANNER_ACCENT_ERROR)
            return
        }
        Task_AddAttachment(parentType, parentId, "image", "attachments\" . fname, desc)
        Task_Notify("Image attached", 1200, BANNER_ACCENT_SUCCESS)
        return
    }

    if (clipText = "") {
        Task_Notify("Clipboard has no image or text", 1600, BANNER_ACCENT_ERROR)
        return
    }

    kind := "text"
    ref := clipText
    desc := clipText
    if (RegExMatch(clipText, "i)^https?://")) {
        kind := "url"
        desc := clipText
    } else if (RegExMatch(clipText, "i)^[A-Za-z]:\\") || RegExMatch(clipText, "^\\\\")) {
        kind := "file"
        desc := clipText
    } else if (StrLen(clipText) > 200) {
        kind := "text"
        res := Task_InputBox("Short description for this text blob:", "Text attachment", SubStr(clipText, 1, 60))
        if (res.Result != "OK")
            return
        desc := Trim(res.Value)
        if (desc = "")
            desc := SubStr(clipText, 1, 80)
        ; store long text as sidecar .txt under attachments
        fname := FormatTime(, "yyyyMMdd-HHmmss") . "-" . parentId . ".txt"
        dest := Task_AttachmentsDir() . "\" . fname
        Task_WriteUtf8(dest, clipText)
        ref := "attachments\" . fname
    }
    Task_AddAttachment(parentType, parentId, kind, ref, desc)
    Task_Notify("Attachment saved (" . kind . ")", 1200, BANNER_ACCENT_SUCCESS)
}

Task_AddAttachment(parentType, parentId, kind, ref, description) {
    rows := Task_Load("attachments")
    row := Map(
        "id", Task_NextId("ATT_", rows),
        "parent_type", parentType,
        "parent_id", parentId,
        "kind", kind,
        "ref", ref,
        "description", description,
        "sort_order", Task_NextSortOrder(rows)
    )
    rows.Push(row)
    Task_Save("attachments", rows)
}
