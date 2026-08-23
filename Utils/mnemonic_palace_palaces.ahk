; =============================================================================
; Utils module: mnemonic_palace_palaces.ahk
; Memory Palaces CRUD (images + image prompts)
; =============================================================================

global g_PalacePalaceLv := false
global g_PalacePalaceRows := []

Palace_ShowPalaces() {
    global g_PalaceGui, g_PalacePalaceLv, g_PalaceFilterStudyId
    Palace_CloseGui()
    Palace_EnsureData()
    if (g_PalaceFilterStudyId = "") {
        Palace_ShowStudies()
        return
    }
    g_PalaceGui := Gui("+AlwaysOnTop +ToolWindow", Palace_BrowseWindowTitle("Palaces"))
    g_PalaceGui.SetFont("s10", "Segoe UI")
    lvY := Palace_AddBrowseChrome(g_PalaceGui, "Palaces")
    g_PalacePalaceLv := g_PalaceGui.Add("ListView", "x12 y" . lvY . " w860 h440 Grid Background2D2D30",
        ["#", "Title", "Character", "Image", "Slots", "Prompt"])
    Palace_StyleDarkListView(g_PalacePalaceLv)
    g_PalacePalaceLv.OnEvent("DoubleClick", (*) => Palace_BrowseInto())
    g_PalaceGui.OnEvent("Close", (*) => Palace_CloseGui())
    g_PalaceGui.OnEvent("Escape", (*) => Palace_BrowseUp())
    Palace_PalaceRefresh()
    Palace_BindHotkeys([
        ["a", (*) => Palace_PalaceAdd()],
        ["Insert", (*) => Palace_PalaceAdd()],
        ["e", (*) => Palace_PalaceEdit()],
        ["Delete", (*) => Palace_PalaceDelete()],
        ["Enter", (*) => Palace_BrowseInto()],
        ["Backspace", (*) => Palace_BrowseUp()],
        ["Escape", (*) => Palace_BrowseUp()]
    ])
    Palace_CenterGui(g_PalaceGui, 890, 540)
}

Palace_PalaceFilter(*) {
    global g_PalaceFilterStudyId
    pick := Palace_PickStudy()
    if (pick = "")
        return
    g_PalaceFilterStudyId := pick
    Palace_SetSetting("General", "LastStudyId", pick)
    Palace_ShowPalaces()
}

Palace_PalaceRefresh() {
    global g_PalacePalaceLv, g_PalacePalaceRows, g_PalaceFilterStudyId
    if (!IsObject(g_PalacePalaceLv))
        return
    rows := Palace_FilterBy(Palace_Load("palaces"), "study_id", g_PalaceFilterStudyId)
    g_PalacePalaceLv.Delete()
    g_PalacePalaceRows := []
    for st in rows {
        g_PalacePalaceRows.Push(st)
        img := st["image_rel_path"]
        abs := Palace_ResolveImagePath(img)
        mark := (abs != "" && FileExist(abs)) ? "✓ " : "? "
        prompt := st.Has("image_prompt") ? Trim(st["image_prompt"]) : ""
        promptMark := prompt != "" ? "yes" : "—"
        g_PalacePalaceLv.Add("", st["palace_number"], st["title"], st["character_name"],
            mark . img, st["depth_slots_used"], promptMark)
    }
    loop 6
        g_PalacePalaceLv.ModifyCol(A_Index, "AutoHdr")
    Palace_StyleDarkListView(g_PalacePalaceLv)
}

Palace_PalaceSelected() {
    global g_PalacePalaceLv, g_PalacePalaceRows
    row := g_PalacePalaceLv.GetNext()
    if (!row || row > g_PalacePalaceRows.Length)
        return false
    return g_PalacePalaceRows[row]
}

Palace_PalaceAdd(*) {
    Palace_PalaceForm(false)
}
Palace_PalaceEdit(*) {
    st := Palace_PalaceSelected()
    if (!st) {
        Palace_Notify("Select a Memory Palace", 1200, BANNER_ACCENT_ERROR)
        return
    }
    Palace_PalaceForm(st)
}

Palace_PalaceDelete(*) {
    st := Palace_PalaceSelected()
    if (!st)
        return
    beasts := Palace_FilterBy(Palace_Load("beasts"), "palace_id", st["id"])
    beastIds := Map()
    for b in beasts
        beastIds[b["id"]] := true
    atomCount := 0
    for a in Palace_Load("atoms") {
        if (beastIds.Has(a["beast_id"]))
            atomCount += 1
    }
    msg := "Delete Memory Palace " . st["palace_number"] . ": " . st["title"] . "?"
    if (beasts.Length || atomCount)
        msg .= "`nAlso deletes " . beasts.Length . " beast(s) and " . atomCount . " atom(s)."
    if (!Palace_Confirm(msg, "Palaces"))
        return
    atomOut := []
    for a in Palace_Load("atoms") {
        if (!beastIds.Has(a["beast_id"]))
            atomOut.Push(a)
    }
    Palace_Save("atoms", atomOut)
    beastOut := []
    for r in Palace_Load("beasts") {
        if (r["palace_id"] != st["id"])
            beastOut.Push(r)
    }
    Palace_Save("beasts", beastOut)
    palaceOut := []
    for r in Palace_Load("palaces") {
        if (r["id"] != st["id"])
            palaceOut.Push(r)
    }
    Palace_Save("palaces", palaceOut)
    Palace_PalaceRefresh()
    Palace_Notify("Memory Palace removed", 1200, BANNER_ACCENT_SUCCESS)
}

Palace_PalaceForm(existing) {
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
    g := Gui("+AlwaysOnTop +ToolWindow" . owner, isEdit ? "Edit Memory Palace" : "Add Memory Palace")
    g.SetFont("s10", "Segoe UI")
    g.Add("Text", , "Palace number")
    eNum := g.Add("Edit", "w80", isEdit ? existing["palace_number"] : "")
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
    g.Add("Text", "y+8", "Image prompt (optional; empty for legacy)")
    ePrompt := g.Add("Edit", "w360 r5", isEdit && existing.Has("image_prompt") ? existing["image_prompt"] : "")
    saved := false
    g.Add("Button", "y+16 w100 Default", "Save").OnEvent("Click", SavePalace)
    g.Add("Button", "x+8 w100", "Cancel").OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())
    g.Show()
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    Palace_DialogsEnd()
    if (saved)
        Palace_PalaceRefresh()

    SavePalace(*) {
        num := Trim(eNum.Value)
        title := Trim(eTitle.Value)
        charName := Trim(eChar.Value)
        img := Trim(eImg.Value)
        if (num = "" || title = "") {
            Palace_Alert("Palace number and title are required.", "Palaces")
            return
        }
        if (charName = "") {
            Palace_Alert("Character is required (one per Memory Palace).", "Palaces")
            return
        }
        palaces := Palace_Load("palaces")
        for r in palaces {
            if (r["study_id"] = studyId && r["character_name"] = charName
                && (!isEdit || r["id"] != existing["id"])) {
                Palace_Alert("Character already used on another Memory Palace in this study.", "Palaces")
                return
            }
        }
        pad := Format("{:02d}", Integer(num))
        id := isEdit ? existing["id"] : "PALACE_" . Palace_Slug(studyId) . "_" . pad
        if (!isEdit && Palace_IdExists(palaces, id))
            id := Palace_SlugId("PALACE_", studyId . "_" . num, palaces)
        row := Map(
            "id", id,
            "study_id", studyId,
            "palace_number", num,
            "title", title,
            "character_name", charName,
            "image_rel_path", img,
            "depth_slots_used", Trim(eSlots.Value),
            "image_prompt", ePrompt.Value
        )
        if (isEdit) {
            out := []
            for r in palaces {
                if (r["id"] = existing["id"])
                    out.Push(row)
                else
                    out.Push(r)
            }
            palaces := out
        } else {
            palaces.Push(row)
        }
        Palace_Save("palaces", palaces)
        saved := true
        g.Destroy()
    }
}
