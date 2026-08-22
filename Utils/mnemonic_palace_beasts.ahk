; =============================================================================
; Utils module: mnemonic_palace_beasts.ahk
; Beasts CRUD
; =============================================================================

global g_PalaceBeastLv := false
global g_PalaceBeastRows := []

Palace_ShowBeasts() {
    global g_PalaceGui, g_PalaceBeastLv, g_PalaceFilterStreetId, g_PalaceFilterStudyId
    Palace_CloseGui()
    Palace_EnsureData()
    if (g_PalaceFilterStreetId = "") {
        studyId := g_PalaceFilterStudyId
        if (studyId = "")
            studyId := Palace_PickStudy()
        if (studyId = "") {
            Palace_ShowMainMenu()
            return
        }
        g_PalaceFilterStudyId := studyId
        pick := Palace_PickStreet(studyId)
        if (pick = "") {
            Palace_ShowMainMenu()
            return
        }
        g_PalaceFilterStreetId := pick
    }
    g_PalaceGui := Gui("+AlwaysOnTop +ToolWindow", "Memory Palace — Beasts")
    g_PalaceGui.SetFont("s10", "Segoe UI")
    g_PalaceGui.Add("Text", "x12 y10 w860",
        Palace_StreetLabel(g_PalaceFilterStreetId)
        . "   [A] add   [E] edit   Delete   [F] filter   Backspace menu")
    g_PalaceBeastLv := g_PalaceGui.Add("ListView", "x12 y36 w860 h460 Grid",
        ["Peg", "Name", "Source", "Sensory", "Smashed", "Order"])
    g_PalaceBeastLv.OnEvent("DoubleClick", (*) => Palace_BeastEdit())
    g_PalaceGui.OnEvent("Close", (*) => Palace_CloseGui())
    g_PalaceGui.OnEvent("Escape", (*) => Palace_ShowMainMenu())
    Palace_BeastRefresh()
    Palace_BindHotkeys([
        ["a", (*) => Palace_BeastAdd()],
        ["Insert", (*) => Palace_BeastAdd()],
        ["e", (*) => Palace_BeastEdit()],
        ["Delete", (*) => Palace_BeastDelete()],
        ["f", (*) => Palace_BeastFilter()],
        ["Backspace", (*) => Palace_ShowMainMenu()],
        ["Escape", (*) => Palace_ShowMainMenu()]
    ])
    Palace_CenterGui(g_PalaceGui, 890, 540)
}

Palace_BeastFilter(*) {
    global g_PalaceFilterStreetId, g_PalaceFilterStudyId
    studyId := Palace_PickStudy()
    if (studyId = "")
        return
    g_PalaceFilterStudyId := studyId
    pick := Palace_PickStreet(studyId)
    if (pick = "")
        return
    g_PalaceFilterStreetId := pick
    Palace_ShowBeasts()
}

Palace_BeastRefresh() {
    global g_PalaceBeastLv, g_PalaceBeastRows, g_PalaceFilterStreetId
    if (!IsObject(g_PalaceBeastLv))
        return
    rows := Palace_FilterBy(Palace_Load("beasts"), "street_id", g_PalaceFilterStreetId)
    g_PalaceBeastLv.Delete()
    g_PalaceBeastRows := []
    for b in rows {
        g_PalaceBeastRows.Push(b)
        g_PalaceBeastLv.Add("", b["peg_code"], b["beast_name"], b["beast_source"],
            b["sensory_channel"], b["is_smashed"] = "1" ? "yes" : "no", b["sort_order"])
    }
    loop 6
        g_PalaceBeastLv.ModifyCol(A_Index, "AutoHdr")
}

Palace_BeastSelected() {
    global g_PalaceBeastLv, g_PalaceBeastRows
    row := g_PalaceBeastLv.GetNext()
    if (!row || row > g_PalaceBeastRows.Length)
        return false
    return g_PalaceBeastRows[row]
}

Palace_BeastAdd(*) {
    Palace_BeastForm(false)
}
Palace_BeastEdit(*) {
    b := Palace_BeastSelected()
    if (!b) {
        Palace_Notify("Select a beast", 1200, BANNER_ACCENT_ERROR)
        return
    }
    Palace_BeastForm(b)
}

Palace_BeastDelete(*) {
    b := Palace_BeastSelected()
    if (!b)
        return
    atoms := Palace_FilterBy(Palace_Load("atoms"), "beast_id", b["id"])
    msg := "Delete [" . b["peg_code"] . "] " . b["beast_name"] . "?"
    if (atoms.Length)
        msg .= "`nAlso deletes " . atoms.Length . " knowledge atom(s)."
    if (!Palace_Confirm(msg, "Beasts"))
        return
    atomOut := []
    for a in Palace_Load("atoms") {
        if (a["beast_id"] != b["id"])
            atomOut.Push(a)
    }
    Palace_Save("atoms", atomOut)
    out := []
    for r in Palace_Load("beasts") {
        if (r["id"] != b["id"])
            out.Push(r)
    }
    Palace_Save("beasts", out)
    Palace_BeastRefresh()
    Palace_Notify("Beast removed", 1200, BANNER_ACCENT_SUCCESS)
}

Palace_BeastForm(existing) {
    global g_PalaceGui, g_PalaceFilterStreetId
    isEdit := IsObject(existing)
    streetId := isEdit ? existing["street_id"] : g_PalaceFilterStreetId
    owner := ""
    try {
        if (IsObject(g_PalaceGui))
            owner := " +Owner" . g_PalaceGui.Hwnd
    } catch {
        owner := ""
    }
    Palace_DialogsBegin()
    g := Gui("+AlwaysOnTop +ToolWindow" . owner, isEdit ? "Edit beast" : "Add beast")
    g.SetFont("s10", "Segoe UI")
    g.Add("Text", , "Peg code (A, B, Aa… — never numbers)")
    ePeg := g.Add("Edit", "w120", isEdit ? existing["peg_code"] : "")
    g.Add("Text", "y+8", "Beast name")
    eName := g.Add("Edit", "w360", isEdit ? existing["beast_name"] : "")
    g.Add("Text", "y+8", "Source (Lynne Kelly | Custom)")
    eSrc := g.Add("Edit", "w200", isEdit ? existing["beast_source"] : "Lynne Kelly")
    g.Add("Text", "y+8", "Sensory channel (visual, auditory, …)")
    eSens := g.Add("Edit", "w200", isEdit ? existing["sensory_channel"] : "visual")
    chk := g.Add("CheckBox", "y+8 Checked" . (isEdit && existing["is_smashed"] = "1" ? "1" : "0"),
        "Smashed (up to 4 subtopics; header carries no atom)")
    g.Add("Text", "y+8", "Sort order")
    eOrder := g.Add("Edit", "w80", isEdit ? existing["sort_order"] : "1")
    saved := false
    g.Add("Button", "y+16 w100 Default", "Save").OnEvent("Click", SaveBeast)
    g.Add("Button", "x+8 w100", "Cancel").OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())
    g.Show()
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    Palace_DialogsEnd()
    if (saved)
        Palace_BeastRefresh()

    SaveBeast(*) {
        peg := Trim(ePeg.Value)
        name := Trim(eName.Value)
        if (peg = "" || name = "") {
            Palace_Alert("Peg and name are required.", "Beasts")
            return
        }
        if (RegExMatch(peg, "^\d+$")) {
            Palace_Alert("Peg codes must not be numeric.", "Beasts")
            return
        }
        beasts := Palace_Load("beasts")
        onStreet := Palace_FilterBy(beasts, "street_id", streetId)
        count := 0
        for r in onStreet {
            if (!isEdit || r["id"] != existing["id"])
                count += 1
        }
        if (!isEdit && count >= 5) {
            Palace_Alert("Maximum 5 beasts per street.", "Beasts")
            return
        }
        id := isEdit ? existing["id"] : "BEAST_" . Palace_Slug(streetId) . "_" . Palace_Slug(peg)
        if (!isEdit && Palace_IdExists(beasts, id))
            id := Palace_SlugId("BEAST_", peg . "_" . name, beasts)
        row := Map(
            "id", id,
            "street_id", streetId,
            "peg_code", peg,
            "beast_name", name,
            "beast_source", Trim(eSrc.Value),
            "sensory_channel", Trim(eSens.Value),
            "is_smashed", chk.Value ? "1" : "0",
            "sort_order", Trim(eOrder.Value)
        )
        if (isEdit) {
            out := []
            for r in beasts {
                if (r["id"] = existing["id"])
                    out.Push(row)
                else
                    out.Push(r)
            }
            beasts := out
        } else {
            beasts.Push(row)
        }
        Palace_Save("beasts", beasts)
        saved := true
        g.Destroy()
    }
}
