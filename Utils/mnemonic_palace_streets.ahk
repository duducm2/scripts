; =============================================================================
; Utils module: mnemonic_palace_streets.ahk
; Streets (Memory Palaces) CRUD
; =============================================================================

global g_PalaceStreetLv := false
global g_PalaceStreetRows := []

Palace_ShowStreets() {
    global g_PalaceGui, g_PalaceStreetLv, g_PalaceFilterStudyId
    Palace_CloseGui()
    Palace_EnsureData()
    if (g_PalaceFilterStudyId = "") {
        pick := Palace_PickStudy()
        if (pick = "") {
            Palace_ShowMainMenu()
            return
        }
        g_PalaceFilterStudyId := pick
    }
    g_PalaceGui := Gui("+AlwaysOnTop +ToolWindow", "Memory Palace — Streets")
    g_PalaceGui.SetFont("s10", "Segoe UI")
    g_PalaceGui.Add("Text", "x12 y10 w860",
        Palace_StudyTitle(g_PalaceFilterStudyId)
        . "   [A] add   [E] edit   Delete   [F] filter study   Backspace menu")
    g_PalaceStreetLv := g_PalaceGui.Add("ListView", "x12 y36 w860 h460 Grid",
        ["#", "Title", "Character", "Image", "Slots"])
    g_PalaceStreetLv.OnEvent("DoubleClick", (*) => Palace_StreetEdit())
    g_PalaceGui.OnEvent("Close", (*) => Palace_CloseGui())
    g_PalaceGui.OnEvent("Escape", (*) => Palace_ShowMainMenu())
    Palace_StreetRefresh()
    Palace_BindHotkeys([
        ["a", (*) => Palace_StreetAdd()],
        ["Insert", (*) => Palace_StreetAdd()],
        ["e", (*) => Palace_StreetEdit()],
        ["Delete", (*) => Palace_StreetDelete()],
        ["f", (*) => Palace_StreetFilter()],
        ["Backspace", (*) => Palace_ShowMainMenu()],
        ["Escape", (*) => Palace_ShowMainMenu()]
    ])
    Palace_CenterGui(g_PalaceGui, 890, 540)
}

Palace_StreetFilter(*) {
    global g_PalaceFilterStudyId
    pick := Palace_PickStudy()
    if (pick = "")
        return
    g_PalaceFilterStudyId := pick
    Palace_SetSetting("General", "LastStudyId", pick)
    Palace_ShowStreets()
}

Palace_StreetRefresh() {
    global g_PalaceStreetLv, g_PalaceStreetRows, g_PalaceFilterStudyId
    if (!IsObject(g_PalaceStreetLv))
        return
    rows := Palace_FilterBy(Palace_Load("streets"), "study_id", g_PalaceFilterStudyId)
    g_PalaceStreetLv.Delete()
    g_PalaceStreetRows := []
    for st in rows {
        g_PalaceStreetRows.Push(st)
        img := st["image_rel_path"]
        abs := Palace_ResolveImagePath(img)
        mark := (abs != "" && FileExist(abs)) ? "✓ " : "? "
        g_PalaceStreetLv.Add("", st["street_number"], st["title"], st["character_name"],
            mark . img, st["depth_slots_used"])
    }
    loop 5
        g_PalaceStreetLv.ModifyCol(A_Index, "AutoHdr")
}

Palace_StreetSelected() {
    global g_PalaceStreetLv, g_PalaceStreetRows
    row := g_PalaceStreetLv.GetNext()
    if (!row || row > g_PalaceStreetRows.Length)
        return false
    return g_PalaceStreetRows[row]
}

Palace_StreetAdd(*) {
    Palace_StreetForm(false)
}
Palace_StreetEdit(*) {
    st := Palace_StreetSelected()
    if (!st) {
        Palace_Notify("Select a street", 1200, BANNER_ACCENT_ERROR)
        return
    }
    Palace_StreetForm(st)
}

Palace_StreetDelete(*) {
    st := Palace_StreetSelected()
    if (!st)
        return
    beasts := Palace_FilterBy(Palace_Load("beasts"), "street_id", st["id"])
    if (beasts.Length) {
        Palace_Alert("Remove " . beasts.Length . " beast(s) on this street first.", "Streets")
        return
    }
    if (!Palace_Confirm("Delete Street " . st["street_number"] . ": " . st["title"] . "?", "Streets"))
        return
    out := []
    for r in Palace_Load("streets") {
        if (r["id"] != st["id"])
            out.Push(r)
    }
    Palace_Save("streets", out)
    Palace_StreetRefresh()
    Palace_Notify("Street removed", 1200, BANNER_ACCENT_SUCCESS)
}

Palace_StreetForm(existing) {
    global g_PalaceGui, g_PalaceFilterStudyId
    isEdit := IsObject(existing)
    studyId := isEdit ? existing["study_id"] : g_PalaceFilterStudyId
    study := Palace_FindById(Palace_Load("studies"), studyId)
    slug := study ? study["notes_rel_path"] : ""
    owner := ""
    try {
        if (IsObject(g_PalaceGui))
            owner := " +Owner" . g_PalaceGui.Hwnd
    } catch {
        owner := ""
    }
    Palace_DialogsBegin()
    g := Gui("+AlwaysOnTop +ToolWindow" . owner, isEdit ? "Edit street" : "Add street")
    g.SetFont("s10", "Segoe UI")
    g.Add("Text", , "Street number")
    eNum := g.Add("Edit", "w80", isEdit ? existing["street_number"] : "")
    g.Add("Text", "y+8", "Title")
    eTitle := g.Add("Edit", "w360", isEdit ? existing["title"] : "")
    g.Add("Text", "y+8", "Character (from characters.json)")
    eChar := g.Add("Edit", "w360", isEdit ? existing["character_name"] : "")
    g.Add("Text", "y+8", "Image relative path (e.g. english/images/1.png)")
    defaultImg := ""
    if (isEdit)
        defaultImg := existing["image_rel_path"]
    else if (slug != "")
        defaultImg := slug . "/images/"
    eImg := g.Add("Edit", "w360", defaultImg)
    g.Add("Text", "y+8", "Depth slots used (0–5)")
    eSlots := g.Add("Edit", "w80", isEdit ? existing["depth_slots_used"] : "0")
    saved := false
    g.Add("Button", "y+16 w100 Default", "Save").OnEvent("Click", SaveStreet)
    g.Add("Button", "x+8 w100", "Cancel").OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())
    g.Show()
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    Palace_DialogsEnd()
    if (saved)
        Palace_StreetRefresh()

    SaveStreet(*) {
        num := Trim(eNum.Value)
        title := Trim(eTitle.Value)
        charName := Trim(eChar.Value)
        img := Trim(eImg.Value)
        if (num = "" || title = "") {
            Palace_Alert("Street number and title are required.", "Streets")
            return
        }
        if (charName = "") {
            Palace_Alert("Character is required (one per street).", "Streets")
            return
        }
        streets := Palace_Load("streets")
        for r in streets {
            if (r["study_id"] = studyId && r["character_name"] = charName
                && (!isEdit || r["id"] != existing["id"])) {
                Palace_Alert("Character already used on another street in this study.", "Streets")
                return
            }
        }
        pad := Format("{:02d}", Integer(num))
        id := isEdit ? existing["id"] : "STREET_" . Palace_Slug(studyId) . "_" . pad
        if (!isEdit && Palace_IdExists(streets, id))
            id := Palace_SlugId("STREET_", studyId . "_" . num, streets)
        row := Map(
            "id", id,
            "study_id", studyId,
            "street_number", num,
            "title", title,
            "character_name", charName,
            "image_rel_path", img,
            "depth_slots_used", Trim(eSlots.Value)
        )
        if (isEdit) {
            out := []
            for r in streets {
                if (r["id"] = existing["id"])
                    out.Push(row)
                else
                    out.Push(r)
            }
            streets := out
        } else {
            streets.Push(row)
        }
        Palace_Save("streets", streets)
        saved := true
        g.Destroy()
    }
}
