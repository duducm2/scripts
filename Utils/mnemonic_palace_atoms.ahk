; =============================================================================
; Utils module: mnemonic_palace_atoms.ahk
; Knowledge Atoms / Topics / Subtopics CRUD
; =============================================================================

global g_PalaceAtomLv := false
global g_PalaceAtomRows := []

Palace_ShowAtoms() {
    global g_PalaceGui, g_PalaceAtomLv, g_PalaceFilterBeastId, g_PalaceFilterStreetId, g_PalaceFilterStudyId
    Palace_CloseGui()
    Palace_EnsureData()
    if (g_PalaceFilterBeastId = "") {
        if (g_PalaceFilterStreetId = "") {
            studyId := g_PalaceFilterStudyId
            if (studyId = "")
                studyId := Palace_PickStudy()
            if (studyId = "") {
                Palace_ShowMainMenu()
                return
            }
            g_PalaceFilterStudyId := studyId
            streetId := Palace_PickStreet(studyId)
            if (streetId = "") {
                Palace_ShowMainMenu()
                return
            }
            g_PalaceFilterStreetId := streetId
        }
        pick := Palace_PickBeast(g_PalaceFilterStreetId)
        if (pick = "") {
            Palace_ShowMainMenu()
            return
        }
        g_PalaceFilterBeastId := pick
    }
    g_PalaceGui := Gui("+AlwaysOnTop +ToolWindow", "Memory Palace — Knowledge Atoms")
    g_PalaceGui.SetFont("s10", "Segoe UI")
    g_PalaceGui.Add("Text", "x12 y10 w860",
        Palace_BeastLabel(g_PalaceFilterBeastId)
        . "   [A] add   [E] edit   Delete   [F] filter   Backspace menu")
    g_PalaceAtomLv := g_PalaceGui.Add("ListView", "x12 y36 w860 h460 Grid",
        ["Kind", "Zone", "Label", "Context", "Sensory"])
    g_PalaceAtomLv.OnEvent("DoubleClick", (*) => Palace_AtomEdit())
    g_PalaceGui.OnEvent("Close", (*) => Palace_CloseGui())
    g_PalaceGui.OnEvent("Escape", (*) => Palace_ShowMainMenu())
    Palace_AtomRefresh()
    Palace_BindHotkeys([
        ["a", (*) => Palace_AtomAdd()],
        ["Insert", (*) => Palace_AtomAdd()],
        ["e", (*) => Palace_AtomEdit()],
        ["Delete", (*) => Palace_AtomDelete()],
        ["f", (*) => Palace_AtomFilter()],
        ["Backspace", (*) => Palace_ShowMainMenu()],
        ["Escape", (*) => Palace_ShowMainMenu()]
    ])
    Palace_CenterGui(g_PalaceGui, 890, 540)
}

Palace_AtomFilter(*) {
    global g_PalaceFilterBeastId, g_PalaceFilterStreetId, g_PalaceFilterStudyId
    studyId := Palace_PickStudy()
    if (studyId = "")
        return
    g_PalaceFilterStudyId := studyId
    streetId := Palace_PickStreet(studyId)
    if (streetId = "")
        return
    g_PalaceFilterStreetId := streetId
    pick := Palace_PickBeast(streetId)
    if (pick = "")
        return
    g_PalaceFilterBeastId := pick
    Palace_ShowAtoms()
}

Palace_AtomRefresh() {
    global g_PalaceAtomLv, g_PalaceAtomRows, g_PalaceFilterBeastId
    if (!IsObject(g_PalaceAtomLv))
        return
    rows := Palace_FilterBy(Palace_Load("atoms"), "beast_id", g_PalaceFilterBeastId)
    g_PalaceAtomLv.Delete()
    g_PalaceAtomRows := []
    for a in rows {
        g_PalaceAtomRows.Push(a)
        ctx := a["context"]
        if (StrLen(ctx) > 60)
            ctx := SubStr(ctx, 1, 57) . "..."
        g_PalaceAtomLv.Add("", a["kind"], a["zone"], a["zone_label"], ctx, a["sensory_channel"])
    }
    loop 5
        g_PalaceAtomLv.ModifyCol(A_Index, "AutoHdr")
}

Palace_AtomSelected() {
    global g_PalaceAtomLv, g_PalaceAtomRows
    row := g_PalaceAtomLv.GetNext()
    if (!row || row > g_PalaceAtomRows.Length)
        return false
    return g_PalaceAtomRows[row]
}

Palace_AtomAdd(*) {
    Palace_AtomForm(false)
}
Palace_AtomEdit(*) {
    a := Palace_AtomSelected()
    if (!a) {
        Palace_Notify("Select an atom", 1200, BANNER_ACCENT_ERROR)
        return
    }
    Palace_AtomForm(a)
}

Palace_AtomDelete(*) {
    a := Palace_AtomSelected()
    if (!a)
        return
    if (!Palace_Confirm("Delete this knowledge atom?", "Atoms"))
        return
    out := []
    for r in Palace_Load("atoms") {
        if (r["id"] != a["id"])
            out.Push(r)
    }
    Palace_Save("atoms", out)
    Palace_AtomRefresh()
    Palace_Notify("Atom removed", 1200, BANNER_ACCENT_SUCCESS)
}

Palace_AtomForm(existing) {
    global g_PalaceGui, g_PalaceFilterBeastId
    isEdit := IsObject(existing)
    beastId := isEdit ? existing["beast_id"] : g_PalaceFilterBeastId
    beast := Palace_FindById(Palace_Load("beasts"), beastId)
    smashed := beast && beast["is_smashed"] = "1"
    owner := ""
    try {
        if (IsObject(g_PalaceGui))
            owner := " +Owner" . g_PalaceGui.Hwnd
    } catch {
        owner := ""
    }
    Palace_DialogsBegin()
    g := Gui("+AlwaysOnTop +ToolWindow" . owner, isEdit ? "Edit atom" : "Add atom")
    g.SetFont("s10", "Segoe UI")
    defaultKind := smashed ? "subtopic" : "single"
    if (isEdit)
        defaultKind := existing["kind"]
    g.Add("Text", , "Kind (single | subtopic)")
    eKind := g.Add("Edit", "w120", defaultKind)
    g.Add("Text", "y+8", "Zone (Z1–Z4 for subtopic; empty for single)")
    eZone := g.Add("Edit", "w80", isEdit ? existing["zone"] : (smashed ? "Z1" : ""))
    g.Add("Text", "y+8", "Zone label")
    eLabel := g.Add("Edit", "w360", isEdit ? existing["zone_label"] : "")
    g.Add("Text", "y+8", "Context (💡 definition)")
    eCtx := g.Add("Edit", "w480 r2", isEdit ? existing["context"] : "")
    g.Add("Text", "y+8", "Quote")
    eQuote := g.Add("Edit", "w480 r2", isEdit ? existing["quote"] : "")
    g.Add("Text", "y+8", "Narrative")
    eNarr := g.Add("Edit", "w480 r3", isEdit ? existing["narrative"] : "")
    g.Add("Text", "y+8", "IPA (optional)")
    eIpa := g.Add("Edit", "w480", isEdit ? existing["ipa"] : "")
    g.Add("Text", "y+8", "Sensory channel")
    eSens := g.Add("Edit", "w200", isEdit ? existing["sensory_channel"] : (smashed ? "visual" : ""))
    g.Add("Text", "y+8", "Sort order")
    eOrder := g.Add("Edit", "w80", isEdit ? existing["sort_order"] : "1")
    saved := false
    g.Add("Button", "y+16 w100 Default", "Save").OnEvent("Click", SaveAtom)
    g.Add("Button", "x+8 w100", "Cancel").OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())
    g.Show()
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    Palace_DialogsEnd()
    if (saved)
        Palace_AtomRefresh()

    SaveAtom(*) {
        kind := StrLower(Trim(eKind.Value))
        if (kind != "single" && kind != "subtopic") {
            Palace_Alert("Kind must be single or subtopic.", "Atoms")
            return
        }
        zone := Trim(eZone.Value)
        if (kind = "subtopic" && !RegExMatch(zone, "^Z[1-4]$")) {
            Palace_Alert("Subtopics require zone Z1, Z2, Z3, or Z4.", "Atoms")
            return
        }
        if (kind = "single")
            zone := ""
        atoms := Palace_Load("atoms")
        id := isEdit ? existing["id"] : Palace_NextId("ATOM_", atoms, 4)
        row := Map(
            "id", id,
            "beast_id", beastId,
            "kind", kind,
            "zone", zone,
            "zone_label", Trim(eLabel.Value),
            "context", eCtx.Value,
            "quote", eQuote.Value,
            "narrative", eNarr.Value,
            "ipa", Trim(eIpa.Value),
            "sensory_channel", Trim(eSens.Value),
            "sort_order", Trim(eOrder.Value)
        )
        proposed := []
        for r in atoms {
            if (isEdit && r["id"] = existing["id"])
                continue
            if (r["beast_id"] = beastId)
                proposed.Push(r)
        }
        proposed.Push(row)
        err := Palace_ValidateBeastAtoms(beastId, proposed)
        if (err != "") {
            Palace_Alert(err, "Atoms")
            return
        }
        if (isEdit) {
            out := []
            for r in atoms {
                if (r["id"] = existing["id"])
                    out.Push(row)
                else
                    out.Push(r)
            }
            atoms := out
        } else {
            atoms.Push(row)
        }
        Palace_Save("atoms", atoms)
        if (kind = "subtopic" && beast) {
            beasts := Palace_Load("beasts")
            bout := []
            for b in beasts {
                if (b["id"] = beastId) {
                    b["is_smashed"] := "1"
                    bout.Push(b)
                } else {
                    bout.Push(b)
                }
            }
            Palace_Save("beasts", bout)
        }
        saved := true
        g.Destroy()
    }
}
