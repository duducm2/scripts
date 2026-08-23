; =============================================================================
; Utils module: mnemonic_palace_studies.ahk
; Studies CRUD
; =============================================================================

global g_PalaceStudyLv := false
global g_PalaceStudyRows := []

Palace_ShowStudies() {
    global g_PalaceGui, g_PalaceStudyLv
    Palace_CloseGui()
    Palace_EnsureData()
    g_PalaceGui := Gui("+AlwaysOnTop +ToolWindow", Palace_BrowseWindowTitle("Studies"))
    g_PalaceGui.SetFont("s10", "Segoe UI")
    lvY := Palace_AddBrowseChrome(g_PalaceGui, "Studies")
    g_PalaceStudyLv := g_PalaceGui.Add("ListView", "x12 y" . lvY . " w860 h440 Grid Background2D2D30",
        ["Title", "Notes path", "Order", "Active"])
    Palace_StyleDarkListView(g_PalaceStudyLv)
    g_PalaceStudyLv.OnEvent("DoubleClick", (*) => Palace_BrowseInto())
    g_PalaceGui.OnEvent("Close", (*) => Palace_CloseGui())
    g_PalaceGui.OnEvent("Escape", (*) => Palace_BrowseUp())
    Palace_StudyRefresh()
    Palace_BindHotkeys([
        ["a", (*) => Palace_StudyAdd()],
        ["Insert", (*) => Palace_StudyAdd()],
        ["e", (*) => Palace_StudyEdit()],
        ["Delete", (*) => Palace_StudyDelete()],
        ["Enter", (*) => Palace_BrowseInto()],
        ["Backspace", (*) => Palace_BrowseUp()],
        ["Escape", (*) => Palace_BrowseUp()]
    ])
    Palace_CenterGui(g_PalaceGui, 890, 540)
}

Palace_StudyRefresh() {
    global g_PalaceStudyLv, g_PalaceStudyRows
    if (!IsObject(g_PalaceStudyLv))
        return
    rows := Palace_Load("studies")
    g_PalaceStudyLv.Delete()
    g_PalaceStudyRows := []
    for s in rows {
        g_PalaceStudyRows.Push(s)
        g_PalaceStudyLv.Add("", s["title"], s["notes_rel_path"],
            s["sort_order"], s["active"])
    }
    loop 4
        g_PalaceStudyLv.ModifyCol(A_Index, "AutoHdr")
    Palace_StyleDarkListView(g_PalaceStudyLv)
}

Palace_StudySelected() {
    global g_PalaceStudyLv, g_PalaceStudyRows
    row := g_PalaceStudyLv.GetNext()
    if (!row || row > g_PalaceStudyRows.Length)
        return false
    return g_PalaceStudyRows[row]
}

Palace_StudyAdd(*) {
    Palace_StudyForm(false)
}
Palace_StudyEdit(*) {
    s := Palace_StudySelected()
    if (!s) {
        Palace_Notify("Select a study", 1200, BANNER_ACCENT_ERROR)
        return
    }
    Palace_StudyForm(s)
}

Palace_StudyDelete(*) {
    s := Palace_StudySelected()
    if (!s)
        return
    palaces := Palace_FilterBy(Palace_Load("palaces"), "study_id", s["id"])
    if (palaces.Length) {
        Palace_Alert("Remove or reassign " . palaces.Length . " Memory Palace(s) first.", "Studies")
        return
    }
    if (!Palace_Confirm("Delete study " . s["title"] . "?", "Studies"))
        return
    out := []
    for r in Palace_Load("studies") {
        if (r["id"] != s["id"])
            out.Push(r)
    }
    Palace_Save("studies", out)
    Palace_StudyRefresh()
    Palace_Notify("Study removed", 1200, BANNER_ACCENT_SUCCESS)
}

Palace_StudyForm(existing) {
    global g_PalaceGui
    isEdit := IsObject(existing)
    owner := ""
    try {
        if (IsObject(g_PalaceGui))
            owner := " +Owner" . g_PalaceGui.Hwnd
    } catch {
        owner := ""
    }
    Palace_DialogsBegin()
    g := Gui("+AlwaysOnTop +ToolWindow" . owner, isEdit ? "Edit study" : "Add study")
    g.SetFont("s10", "Segoe UI")
    g.Add("Text", , "Title")
    eTitle := g.Add("Edit", "w320", isEdit ? existing["title"] : "")
    g.Add("Text", "y+8", "Notes relative path (folder under notes/studies)")
    ePath := g.Add("Edit", "w320", isEdit ? existing["notes_rel_path"] : "")
    g.Add("Text", "y+8", "Sort order")
    eOrder := g.Add("Edit", "w80", isEdit ? existing["sort_order"] : "1")
    chk := g.Add("CheckBox", "y+8 Checked" . (isEdit ? (existing["active"] = "0" ? "0" : "1") : "1"), "Active")
    saved := false
    g.Add("Button", "y+16 w100 Default", "Save").OnEvent("Click", SaveStudy)
    g.Add("Button", "x+8 w100", "Cancel").OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())
    g.Show()
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    Palace_DialogsEnd()
    if (saved)
        Palace_StudyRefresh()

    SaveStudy(*) {
        title := Trim(eTitle.Value)
        npath := Trim(ePath.Value)
        if (title = "" || npath = "") {
            Palace_Alert("Title and notes path are required.", "Studies")
            return
        }
        studies := Palace_Load("studies")
        row := Map(
            "id", isEdit ? existing["id"] : Palace_SlugId("STUDY_", npath, studies),
        "title", title,
        "notes_rel_path", npath,
        "sort_order", Trim(eOrder.Value),
        "active", chk.Value ? "1" : "0"
        )
        if (isEdit) {
            out := []
            for r in studies {
                if (r["id"] = existing["id"])
                    out.Push(row)
                else
                    out.Push(r)
            }
            studies := out
        } else {
            studies.Push(row)
        }
        Palace_Save("studies", studies)
        Palace_SetSetting("General", "LastStudyId", row["id"])
        saved := true
        g.Destroy()
    }
}
