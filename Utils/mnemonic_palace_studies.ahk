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
    palaceIds := Map()
    for p in palaces
        palaceIds[p["id"]] := true
    beastIds := Map()
    beastCount := 0
    for b in Palace_Load("beasts") {
        if (palaceIds.Has(b["palace_id"])) {
            beastIds[b["id"]] := true
            beastCount += 1
        }
    }
    atomCount := 0
    for a in Palace_Load("atoms") {
        if (beastIds.Has(a["beast_id"]))
            atomCount += 1
    }
    msg := "Delete study " . s["title"] . "?"
    if (palaces.Length || beastCount || atomCount)
        msg .= "`nAlso deletes " . palaces.Length . " palace(s), "
            . beastCount . " beast(s), and " . atomCount . " atom(s)."
    if (!Palace_Confirm(msg, "Studies"))
        return
    atomOut := []
    for a in Palace_Load("atoms") {
        if (!beastIds.Has(a["beast_id"]))
            atomOut.Push(a)
    }
    Palace_Save("atoms", atomOut)
    beastOut := []
    for b in Palace_Load("beasts") {
        if (!palaceIds.Has(b["palace_id"]))
            beastOut.Push(b)
    }
    Palace_Save("beasts", beastOut)
    palaceOut := []
    for r in Palace_Load("palaces") {
        if (r["study_id"] != s["id"])
            palaceOut.Push(r)
    }
    Palace_Save("palaces", palaceOut)
    studyOut := []
    for r in Palace_Load("studies") {
        if (r["id"] != s["id"])
            studyOut.Push(r)
    }
    Palace_Save("studies", studyOut)
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
        if (title = "") {
            Palace_Alert("Title is required.", "Studies")
            return
        }
        ; Notes folder is attributed under the hood (not shown in the form)
        npath := isEdit ? Trim(existing["notes_rel_path"]) : Palace_NotesFolderSlug(title)
        if (npath = "")
            npath := Palace_NotesFolderSlug(title)
        npath := StrReplace(npath, "/", "\")
        npath := Trim(npath, "\")
        if (npath = "" || InStr(npath, "..") || InStr(npath, ":")) {
            Palace_Alert("Could not derive a notes folder from the title.", "Studies")
            return
        }
        if (!Palace_EnsureStudyNotesFolder(npath)) {
            root := Palace_NotesStudiesRoot(true)
            if (root = "")
                Palace_Alert(
                    "Could not resolve notes/studies root. Set NotesStudiesRoot in settings or ensure the notes repo path exists.",
                    "Studies")
            else
                Palace_Alert("Could not create notes folder for this study.", "Studies")
            return
        }
        studies := Palace_Load("studies")
        sortOrder := isEdit ? existing["sort_order"] : Palace_NextSortOrder(studies)
        row := Map(
            "id", isEdit ? existing["id"] : Palace_SlugId("STUDY_", npath, studies),
        "title", title,
        "notes_rel_path", npath,
        "sort_order", sortOrder,
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
        Palace_Notify(isEdit ? "Study updated" : "Study saved", 1200, BANNER_ACCENT_SUCCESS)
    }
}
