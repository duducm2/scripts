; =============================================================================
; Utils module: mnemonic_palace_plans.ahk
; Study Plans CRUD (CSV) — plans + plan_items under a Study
; =============================================================================

global g_PalacePlanLv := false
global g_PalacePlanRows := []
global g_PalacePlanItemLv := false
global g_PalacePlanItemRows := []

Palace_OpenPlansForSelectedStudy(*) {
    global g_PalaceFilterStudyId, g_PalaceFilterPlanId
    s := Palace_StudySelected()
    if (!s) {
        Palace_Notify("Select a study", 1200, BANNER_ACCENT_ERROR)
        return
    }
    g_PalaceFilterStudyId := s["id"]
    g_PalaceFilterPlanId := ""
    Palace_ShowPlans()
}

Palace_OpenPlansForCurrentStudy(*) {
    global g_PalaceFilterStudyId, g_PalaceFilterPlanId, g_PalaceFilterPalaceId, g_PalaceFilterBeastId
    if (g_PalaceFilterStudyId = "") {
        Palace_Notify("Open a study first", 1200, BANNER_ACCENT_ERROR)
        return
    }
    g_PalaceFilterPalaceId := ""
    g_PalaceFilterBeastId := ""
    g_PalaceFilterPlanId := ""
    Palace_ShowPlans()
}

Palace_ShowPlans() {
    global g_PalaceGui, g_PalacePlanLv, g_PalaceFilterStudyId, g_PalaceFilterPlanId
    Palace_CloseGui()
    Palace_EnsureData()
    if (g_PalaceFilterStudyId = "") {
        Palace_ShowStudies()
        return
    }
    g_PalaceFilterPlanId := ""
    g_PalaceGui := Gui("+AlwaysOnTop +ToolWindow", Palace_BrowseWindowTitle("Plans"))
    g_PalaceGui.SetFont("s10", "Segoe UI")
    lvY := Palace_AddBrowseChrome(g_PalaceGui, "Plans")
    g_PalacePlanLv := g_PalaceGui.Add("ListView", "x12 y" . lvY . " w860 h440 Grid Background2D2D30",
        ["Title", "Items", "Active", "Id"])
    Palace_StyleDarkListView(g_PalacePlanLv)
    g_PalacePlanLv.OnEvent("DoubleClick", (*) => Palace_PlanOpenItems())
    g_PalaceGui.OnEvent("Close", (*) => Palace_CloseGui())
    g_PalaceGui.OnEvent("Escape", (*) => Palace_BrowseUpFromPlans())
    Palace_PlanRefresh()
    Palace_BindHotkeys([
        ["a", (*) => Palace_PlanAdd()],
        ["Insert", (*) => Palace_PlanAdd()],
        ["e", (*) => Palace_PlanEdit()],
        ["Delete", (*) => Palace_PlanDelete()],
        ["Enter", (*) => Palace_PlanOpenItems()],
        ["Backspace", (*) => Palace_BrowseUpFromPlans()],
        ["Escape", (*) => Palace_BrowseUpFromPlans()]
    ])
    Palace_CenterGui(g_PalaceGui, 890, 540)
}

Palace_BrowseUpFromPlans() {
    global g_PalaceFilterPlanId
    g_PalaceFilterPlanId := ""
    Palace_ShowStudies()
}

Palace_PlanRefresh() {
    global g_PalacePlanLv, g_PalacePlanRows, g_PalaceFilterStudyId
    if (!IsObject(g_PalacePlanLv))
        return
    rows := Palace_FilterBy(Palace_Load("plans"), "study_id", g_PalaceFilterStudyId)
    items := Palace_Load("plan_items")
    g_PalacePlanLv.Delete()
    g_PalacePlanRows := []
    for p in rows {
        g_PalacePlanRows.Push(p)
        n := 0
        for it in items {
            if (it["plan_id"] = p["id"])
                n += 1
        }
        g_PalacePlanLv.Add("", p["title"], n, p["active"], p["id"])
    }
    loop 4
        g_PalacePlanLv.ModifyCol(A_Index, "AutoHdr")
    Palace_StyleDarkListView(g_PalacePlanLv)
}

Palace_PlanSelected() {
    global g_PalacePlanLv, g_PalacePlanRows
    row := g_PalacePlanLv.GetNext()
    if (!row || row > g_PalacePlanRows.Length)
        return false
    return g_PalacePlanRows[row]
}

Palace_PlanOpenItems(*) {
    global g_PalaceFilterPlanId
    p := Palace_PlanSelected()
    if (!p) {
        Palace_Notify("Select a plan", 1200, BANNER_ACCENT_ERROR)
        return
    }
    g_PalaceFilterPlanId := p["id"]
    Palace_ShowPlanItems()
}

Palace_PlanAdd(*) {
    global g_PalaceFilterStudyId
    existing := Palace_FilterBy(Palace_Load("plans"), "study_id", g_PalaceFilterStudyId)
    activeCount := 0
    for p in existing {
        if (p["active"] != "0")
            activeCount += 1
    }
    if (activeCount > 0) {
        Palace_Notify("Study already has an active plan — edit or deactivate it first", 2800, BANNER_ACCENT_ERROR)
        return
    }
    Palace_PlanForm(false)
}

Palace_PlanEdit(*) {
    p := Palace_PlanSelected()
    if (!p) {
        Palace_Notify("Select a plan", 1200, BANNER_ACCENT_ERROR)
        return
    }
    Palace_PlanForm(p)
}

Palace_PlanDelete(*) {
    p := Palace_PlanSelected()
    if (!p)
        return
    itemCount := 0
    for it in Palace_Load("plan_items") {
        if (it["plan_id"] = p["id"])
            itemCount += 1
    }
    msg := "Delete plan " . p["title"] . "?"
    if (itemCount)
        msg .= "`nAlso deletes " . itemCount . " item(s)."
    if (!Palace_Confirm(msg, "Plans"))
        return
    itemOut := []
    for it in Palace_Load("plan_items") {
        if (it["plan_id"] != p["id"])
            itemOut.Push(it)
    }
    Palace_Save("plan_items", itemOut)
    resOut := []
    for r in Palace_Load("plan_resources") {
        if (r["plan_id"] != p["id"])
            resOut.Push(r)
    }
    Palace_Save("plan_resources", resOut)
    planOut := []
    for r in Palace_Load("plans") {
        if (r["id"] != p["id"])
            planOut.Push(r)
    }
    Palace_Save("plans", planOut)
    Palace_SyncPlansMd([p["study_id"]])
    Palace_PlanRefresh()
    Palace_Notify("Plan removed", 1200, BANNER_ACCENT_SUCCESS)
}

Palace_PlanForm(existing) {
    global g_PalaceGui, g_PalaceFilterStudyId
    isEdit := IsObject(existing)
    studyId := isEdit ? existing["study_id"] : g_PalaceFilterStudyId
    owner := ""
    try {
        if (IsObject(g_PalaceGui))
            owner := " +Owner" . g_PalaceGui.Hwnd
    } catch {
        owner := ""
    }
    Palace_DialogsBegin()
    g := Gui("+AlwaysOnTop +ToolWindow" . owner, isEdit ? "Edit plan" : "Add plan")
    g.SetFont("s10", "Segoe UI")
    g.Add("Text", , "Title")
    eTitle := g.Add("Edit", "w360", isEdit ? existing["title"] : (Palace_StudyTitle(studyId) . " Plan"))
    chk := g.Add("CheckBox", "y+8 Checked" . (isEdit ? (existing["active"] = "0" ? "0" : "1") : "1"), "Active")
    saved := false
    g.Add("Button", "y+16 w100 Default", "Save").OnEvent("Click", SavePlan)
    g.Add("Button", "x+8 w100", "Cancel").OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())
    g.Show()
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    Palace_DialogsEnd()
    if (saved)
        Palace_PlanRefresh()
    return

    SavePlan(*) {
        title := Trim(eTitle.Value)
        if (title = "") {
            Palace_Notify("Title required", 1500, BANNER_ACCENT_ERROR)
            return
        }
        rows := Palace_Load("plans")
        if (isEdit) {
            out := []
            for r in rows {
                if (r["id"] = existing["id"]) {
                    out.Push(Map(
                        "id", existing["id"],
                        "study_id", studyId,
                        "title", title,
                        "sort_order", existing["sort_order"],
                        "active", chk.Value ? "1" : "0"
                    ))
                } else {
                    out.Push(r)
                }
            }
            rows := out
        } else {
            pid := Palace_NextId("PLAN_", rows, 4)
            rows.Push(Map(
                "id", pid,
                "study_id", studyId,
                "title", title,
                "sort_order", "1",
                "active", chk.Value ? "1" : "0"
            ))
        }
        Palace_Save("plans", rows)
        Palace_SyncPlansMd([studyId])
        saved := true
        g.Destroy()
        Palace_Notify(isEdit ? "Plan updated" : "Plan created", 1200, BANNER_ACCENT_SUCCESS)
    }
}

Palace_ShowPlanItems() {
    global g_PalaceGui, g_PalacePlanItemLv, g_PalaceFilterPlanId, g_PalaceFilterStudyId
    Palace_CloseGui()
    Palace_EnsureData()
    if (g_PalaceFilterPlanId = "") {
        Palace_ShowPlans()
        return
    }
    plan := Palace_FindById(Palace_Load("plans"), g_PalaceFilterPlanId)
    label := IsObject(plan) ? plan["title"] : "Plan items"
    g_PalaceGui := Gui("+AlwaysOnTop +ToolWindow", Palace_BrowseWindowTitle(label))
    g_PalaceGui.SetFont("s10", "Segoe UI")
    lvY := Palace_AddBrowseChrome(g_PalaceGui, "Plan items")
    g_PalacePlanItemLv := g_PalaceGui.Add("ListView", "x12 y" . lvY . " w860 h440 Grid Background2D2D30",
        ["Done", "Section", "Text", "Id"])
    Palace_StyleDarkListView(g_PalacePlanItemLv)
    g_PalacePlanItemLv.OnEvent("DoubleClick", (*) => Palace_PlanItemEdit())
    g_PalaceGui.OnEvent("Close", (*) => Palace_CloseGui())
    g_PalaceGui.OnEvent("Escape", (*) => Palace_ShowPlans())
    Palace_PlanItemRefresh()
    Palace_BindHotkeys([
        ["a", (*) => Palace_PlanItemAdd()],
        ["Insert", (*) => Palace_PlanItemAdd()],
        ["e", (*) => Palace_PlanItemEdit()],
        ["Delete", (*) => Palace_PlanItemDelete()],
        ["Enter", (*) => Palace_PlanItemEdit()],
        ["Backspace", (*) => Palace_ShowPlans()],
        ["Escape", (*) => Palace_ShowPlans()]
    ])
    Palace_CenterGui(g_PalaceGui, 890, 540)
}

Palace_PlanItemRefresh() {
    global g_PalacePlanItemLv, g_PalacePlanItemRows, g_PalaceFilterPlanId
    if (!IsObject(g_PalacePlanItemLv))
        return
    rows := Palace_FilterBy(Palace_Load("plan_items"), "plan_id", g_PalaceFilterPlanId)
    ; sort by sort_order
    sorted := []
    for r in rows
        sorted.Push(r)
    loop sorted.Length {
        i := A_Index
        while (i > 1) {
            a := Integer(sorted[i - 1].Has("sort_order") ? sorted[i - 1]["sort_order"] : 0)
            b := Integer(sorted[i].Has("sort_order") ? sorted[i]["sort_order"] : 0)
            if (a <= b)
                break
            tmp := sorted[i - 1]
            sorted[i - 1] := sorted[i]
            sorted[i] := tmp
            i -= 1
        }
    }
    g_PalacePlanItemLv.Delete()
    g_PalacePlanItemRows := []
    for it in sorted {
        g_PalacePlanItemRows.Push(it)
        txt := it["text"]
        if (StrLen(txt) > 70)
            txt := SubStr(txt, 1, 67) . "..."
        done := (it["checked"] = "1") ? "✓" : ""
        g_PalacePlanItemLv.Add("", done, it["section_path"], txt, it["id"])
    }
    loop 4
        g_PalacePlanItemLv.ModifyCol(A_Index, "AutoHdr")
    Palace_StyleDarkListView(g_PalacePlanItemLv)
}

Palace_PlanItemSelected() {
    global g_PalacePlanItemLv, g_PalacePlanItemRows
    row := g_PalacePlanItemLv.GetNext()
    if (!row || row > g_PalacePlanItemRows.Length)
        return false
    return g_PalacePlanItemRows[row]
}

Palace_PlanItemAdd(*) {
    Palace_PlanItemForm(false)
}
Palace_PlanItemEdit(*) {
    it := Palace_PlanItemSelected()
    if (!it) {
        Palace_Notify("Select an item", 1200, BANNER_ACCENT_ERROR)
        return
    }
    Palace_PlanItemForm(it)
}

Palace_PlanItemDelete(*) {
    it := Palace_PlanItemSelected()
    if (!it)
        return
    if (!Palace_Confirm("Delete this plan item?", "Plan items"))
        return
    out := []
    for r in Palace_Load("plan_items") {
        if (r["id"] != it["id"])
            out.Push(r)
    }
    Palace_Save("plan_items", out)
    plan := Palace_FindById(Palace_Load("plans"), it["plan_id"])
    if (IsObject(plan))
        Palace_SyncPlansMd([plan["study_id"]])
    Palace_PlanItemRefresh()
    Palace_Notify("Item removed", 1200, BANNER_ACCENT_SUCCESS)
}

Palace_PlanItemForm(existing) {
    global g_PalaceGui, g_PalaceFilterPlanId
    isEdit := IsObject(existing)
    planId := isEdit ? existing["plan_id"] : g_PalaceFilterPlanId
    owner := ""
    try {
        if (IsObject(g_PalaceGui))
            owner := " +Owner" . g_PalaceGui.Hwnd
    } catch {
        owner := ""
    }
    Palace_DialogsBegin()
    g := Gui("+AlwaysOnTop +ToolWindow" . owner, isEdit ? "Edit plan item" : "Add plan item")
    g.SetFont("s10", "Segoe UI")
    g.Add("Text", , "Section path (Backlog or Phase 1 > Month 1 > Topic)")
    eSec := g.Add("Edit", "w520", isEdit ? existing["section_path"] : "Backlog")
    g.Add("Text", "y+8", "Text")
    eText := g.Add("Edit", "w520 r3", isEdit ? existing["text"] : "")
    chk := g.Add("CheckBox", "y+8 Checked" . (isEdit ? (existing["checked"] = "1" ? "1" : "0") : "0"), "Checked")
    saved := false
    g.Add("Button", "y+16 w100 Default", "Save").OnEvent("Click", SaveItem)
    g.Add("Button", "x+8 w100", "Cancel").OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())
    g.Show()
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    Palace_DialogsEnd()
    if (saved)
        Palace_PlanItemRefresh()
    return

    SaveItem(*) {
        sec := Trim(eSec.Value)
        text := Trim(eText.Value)
        if (text = "") {
            Palace_Notify("Text required", 1500, BANNER_ACCENT_ERROR)
            return
        }
        if (sec = "")
            sec := "Backlog"
        rows := Palace_Load("plan_items")
        if (isEdit) {
            out := []
            for r in rows {
                if (r["id"] = existing["id"]) {
                    out.Push(Map(
                        "id", existing["id"],
                        "plan_id", planId,
                        "section_path", sec,
                        "text", text,
                        "checked", chk.Value ? "1" : "0",
                        "sort_order", existing["sort_order"]
                    ))
                } else {
                    out.Push(r)
                }
            }
            rows := out
        } else {
            iid := Palace_NextId("PITEM_", rows, 4)
            maxSort := 0
            for r in rows {
                if (r["plan_id"] = planId) {
                    try maxSort := Max(maxSort, Integer(r["sort_order"]))
                    catch {
                    }
                }
            }
            rows.Push(Map(
                "id", iid,
                "plan_id", planId,
                "section_path", sec,
                "text", text,
                "checked", chk.Value ? "1" : "0",
                "sort_order", String(maxSort + 1)
            ))
        }
        Palace_Save("plan_items", rows)
        plan := Palace_FindById(Palace_Load("plans"), planId)
        if (IsObject(plan))
            Palace_SyncPlansMd([plan["study_id"]])
        saved := true
        g.Destroy()
        Palace_Notify(isEdit ? "Item updated" : "Item added", 1200, BANNER_ACCENT_SUCCESS)
    }
}
