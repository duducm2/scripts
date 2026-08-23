; =============================================================================
; Utils module: mnemonic_palace_atoms.ahk
; Knowledge Atoms CRUD
; =============================================================================

global g_PalaceAtomLv := false
global g_PalaceAtomRows := []

Palace_ShowAtoms() {
    global g_PalaceGui, g_PalaceAtomLv, g_PalaceFilterBeastId, g_PalaceFilterPalaceId, g_PalaceFilterStudyId
    Palace_CloseGui()
    Palace_EnsureData()
    if (g_PalaceFilterBeastId = "") {
        if (g_PalaceFilterPalaceId = "") {
            if (g_PalaceFilterStudyId = "")
                Palace_ShowStudies()
            else
                Palace_ShowPalaces()
        } else {
            Palace_ShowBeasts()
        }
        return
    }
    g_PalaceGui := Gui("+AlwaysOnTop +ToolWindow", Palace_BrowseWindowTitle("Knowledge Atoms"))
    g_PalaceGui.SetFont("s10", "Segoe UI")
    g_PalaceGui.Add("Text", "x12 y10 w860", Palace_BrowseBarHints("Knowledge Atoms"))
    g_PalaceAtomLv := g_PalaceGui.Add("ListView", "x12 y36 w860 h460 Grid",
        ["Kind", "Zone", "Label", "Concept", "Sensory"])
    g_PalaceAtomLv.OnEvent("DoubleClick", (*) => Palace_AtomEdit())
    g_PalaceGui.OnEvent("Close", (*) => Palace_CloseGui())
    g_PalaceGui.OnEvent("Escape", (*) => Palace_BrowseUp())
    Palace_AtomRefresh()
    Palace_BindHotkeys([
        ["a", (*) => Palace_AtomAdd()],
        ["Insert", (*) => Palace_AtomAdd()],
        ["e", (*) => Palace_AtomEdit()],
        ["Delete", (*) => Palace_AtomDelete()],
        ["Enter", (*) => Palace_AtomEdit()],
        ["Backspace", (*) => Palace_BrowseUp()],
        ["Escape", (*) => Palace_BrowseUp()]
    ])
    Palace_CenterGui(g_PalaceGui, 890, 540)
}


Palace_AtomFilter(*) {
    global g_PalaceFilterBeastId, g_PalaceFilterPalaceId, g_PalaceFilterStudyId
    studyId := Palace_PickStudy()
    if (studyId = "")
        return
    g_PalaceFilterStudyId := studyId
    palaceId := Palace_PickPalace(studyId)
    if (palaceId = "")
        return
    g_PalaceFilterPalaceId := palaceId
    pick := Palace_PickBeast(palaceId)
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
        ctx := a.Has("concept") ? a["concept"] : (a.Has("context") ? a["context"] : "")
        if (StrLen(ctx) > 60)
            ctx := SubStr(ctx, 1, 57) . "..."
        sens := a.Has("sensory") ? a["sensory"] : (a.Has("sensory_channel") ? a["sensory_channel"] : "")
        g_PalaceAtomLv.Add("", a["kind"], a["zone"], a["zone_label"], ctx, sens)
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
    defaultKind := smashed ? "zoned" : "single"
    if (isEdit) {
        defaultKind := existing["kind"]
        if (defaultKind = "subtopic")
            defaultKind := "zoned"
    }
    g.Add("Text", , "Kind (single | zoned Knowledge Atom)")
    eKind := g.Add("Edit", "w220", defaultKind)
    g.Add("Text", "y+8", "Zone (Z1–Z4 for zoned; empty for single)")
    eZone := g.Add("Edit", "w80", isEdit ? existing["zone"] : (smashed ? "Z1" : ""))
    g.Add("Text", "y+8", "Zone label")
    eLabel := g.Add("Edit", "w360", isEdit ? existing["zone_label"] : "")
    g.Add("Text", "y+8", "Concept (rehearsal definition)")
    eCtx := g.Add("Edit", "w480 r2", isEdit ? (existing.Has("concept") ? existing["concept"] : existing["context"]) :
        "")
    g.Add("Text", "y+8", "Quote")
    eQuote := g.Add("Edit", "w480 r2", isEdit ? existing["quote"] : "")
    g.Add("Text", "y+8", "Story (mnemonic action)")
    eNarr := g.Add("Edit", "w480 r3", isEdit ? (existing.Has("story") ? existing["story"] : existing["narrative"]) : ""
    )
    g.Add("Text", "y+8", "IPA (optional)")
    eIpa := g.Add("Edit", "w480", isEdit ? existing["ipa"] : "")
    g.Add("Text", "y+8", "Sensory (modality the story emphasizes)")
    eSens := g.Add("Edit", "w200", isEdit ? (existing.Has("sensory") ? existing["sensory"] : existing["sensory_channel"
        ]) : (smashed ? "visual" : ""))
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
        if (kind = "subtopic")
            kind := "zoned"
        if (kind != "single" && kind != "zoned") {
            Palace_Alert("Kind must be single or zoned.", "Atoms")
            return
        }
        zone := Trim(eZone.Value)
        if (kind = "zoned" && !RegExMatch(zone, "^Z[1-4]$")) {
            Palace_Alert("Zoned Knowledge Atoms require zone Z1, Z2, Z3, or Z4.", "Atoms")
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
            "concept", eCtx.Value,
            "quote", eQuote.Value,
            "story", eNarr.Value,
            "ipa", Trim(eIpa.Value),
            "sensory", Trim(eSens.Value),
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
        if (kind = "zoned" && beast) {
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
