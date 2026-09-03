; =============================================================================
; Utils module: finance_goals.ahk
; Goals list, sort, CRUD
; =============================================================================

global g_FinanceGoalLv := false
global g_FinanceGoalRows := []
global g_FinanceGoalSort := "name"

Finance_ShowGoals() {
    global g_FinanceGui, g_FinanceGoalLv, g_FinanceGoalSort
    Finance_CloseGui()
    Finance_EnsureData()
    g_FinanceGui := Gui("+AlwaysOnTop +ToolWindow", "Goals")
    g_FinanceGui.SetFont("s10", "Segoe UI")
    g_FinanceGui.Add("Text", "x12 y10 w860",
        "Sort [1] name  [2] date  [3] %   [Shift+A]/Insert add  [Shift+E] edit  Delete  Backspace")
    g_FinanceGoalLv := g_FinanceGui.Add("ListView", "x12 y40 w860 h480 Grid",
        ["Purpose", "Current", "Target", "%", "Date"])
    g_FinanceGoalLv.OnEvent("DoubleClick", (*) => Finance_GoalEdit())
    g_FinanceGui.OnEvent("Close", (*) => Finance_CloseGui())
    g_FinanceGui.OnEvent("Escape", (*) => Finance_ShowMainMenu())
    Finance_GoalRefresh()
    Finance_BindHotkeys([
        ["1", Finance_GoalSortName],
        ["2", Finance_GoalSortDate],
        ["3", Finance_GoalSortPct],
        ["+a", (*) => Finance_GoalAdd()],
        ["Insert", (*) => Finance_GoalAdd()],
        ["+e", (*) => Finance_GoalEdit()],
        ["Delete", (*) => Finance_GoalDelete()],
        ["Backspace", (*) => Finance_ShowMainMenu()],
        ["Escape", (*) => Finance_ShowMainMenu()]
    ])
    Finance_CenterGui(g_FinanceGui, 890, 560)
}

Finance_GoalSortName(*) {
    global g_FinanceGoalSort
    g_FinanceGoalSort := "name"
    Finance_GoalRefresh()
}
Finance_GoalSortDate(*) {
    global g_FinanceGoalSort
    g_FinanceGoalSort := "date"
    Finance_GoalRefresh()
}
Finance_GoalSortPct(*) {
    global g_FinanceGoalSort
    g_FinanceGoalSort := "pct"
    Finance_GoalRefresh()
}

Finance_GoalPct(g) {
    t := Finance_ParseDecimal(g["target_amount"])
    if (t <= 0)
        return 0
    return Round(Finance_ParseDecimal(g["current_amount"]) / t * 100, 1)
}

Finance_GoalRefresh() {
    global g_FinanceGoalLv, g_FinanceGoalRows, g_FinanceGoalSort
    rows := Finance_Load("goals")
    loop rows.Length {
        i := A_Index
        loop rows.Length - i {
            j := i + A_Index
            swap := false
            if (g_FinanceGoalSort = "name" && StrCompare(rows[i]["name"], rows[j]["name"], 1) > 0)
                swap := true
            else if (g_FinanceGoalSort = "date" && StrCompare(rows[i]["target_date"], rows[j]["target_date"]) > 0)
                swap := true
            else if (g_FinanceGoalSort = "pct" && Finance_GoalPct(rows[i]) < Finance_GoalPct(rows[j]))
                swap := true
            if (swap) {
                tmp := rows[i]
                rows[i] := rows[j]
                rows[j] := tmp
            }
        }
    }
    g_FinanceGoalLv.Delete()
    g_FinanceGoalRows := []
    for g in rows {
        cur := Finance_ParseDecimal(g["current_amount"])
        tgt := Finance_ParseDecimal(g["target_amount"])
        g_FinanceGoalRows.Push(g)
        g_FinanceGoalLv.Add("", g["name"], Finance_FormatBrl(cur), Finance_FormatBrl(tgt),
        Finance_GoalPct(g) . "%", g["target_date"])
    }
    loop 5
        g_FinanceGoalLv.ModifyCol(A_Index, "AutoHdr")
}

Finance_GoalSelected() {
    global g_FinanceGoalLv, g_FinanceGoalRows
    row := g_FinanceGoalLv.GetNext()
    if (!row || row > g_FinanceGoalRows.Length)
        return false
    return g_FinanceGoalRows[row]
}

Finance_GoalAdd(*) {
    Finance_GoalForm(false)
}

Finance_GoalEdit(*) {
    g := Finance_GoalSelected()
    if (!g)
        return
    Finance_GoalForm(g)
}

Finance_GoalDelete(*) {
    gsel := Finance_GoalSelected()
    if (!gsel)
        return
    if (!Finance_Confirm("Delete " . gsel["name"] . "?", "Goals"))
        return
    out := []
    for r in Finance_Load("goals") {
        if (r["id"] != gsel["id"])
            out.Push(r)
    }
    Finance_Save("goals", out)
    Finance_GoalRefresh()
}

Finance_GoalForm(existing) {
    global g_FinanceGui
    isEdit := IsObject(existing)
    owner := ""
    try {
        if (IsObject(g_FinanceGui))
            owner := " +Owner" . g_FinanceGui.Hwnd
    } catch {
        owner := ""
    }
    Finance_DialogsBegin()
    g := Gui("+AlwaysOnTop +ToolWindow" . owner, isEdit ? "Edit goal" : "Add goal")
    g.SetFont("s10", "Segoe UI")
    g.Add("Text", , "Purpose")
    eName := g.Add("Edit", "w360", isEdit ? existing["name"] : "")
    g.Add("Text", "y+8", "Current amount")
    eCur := g.Add("Edit", "w160", isEdit ? existing["current_amount"] : "0,00")
    g.Add("Text", "y+8", "Target amount")
    eTgt := g.Add("Edit", "w160", isEdit ? existing["target_amount"] : "0,00")
    g.Add("Text", "y+8", "Target date (YYYY-MM-DD)")
    eDate := g.Add("Edit", "w160", isEdit ? existing["target_date"] : "")
    saved := false
    g.Add("Button", "y+16 w100 Default", "Save").OnEvent("Click", SaveGoal)
    g.Add("Button", "x+8 w100", "Cancel").OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())
    g.Show()
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    Finance_DialogsEnd()
    if (saved)
        Finance_GoalRefresh()

    SaveGoal(*) {
        name := Trim(eName.Value)
        if (name = "") {
            Finance_Alert("Purpose is required.", "Goals")
            return
        }
        rows := Finance_Load("goals")
        row := Map(
            "id", isEdit ? existing["id"] : Finance_SlugId("GOAL_", name, rows),
        "name", name,
        "current_amount", Finance_FormatCsvDecimal(Finance_ParseDecimal(eCur.Value)),
        "target_amount", Finance_FormatCsvDecimal(Finance_ParseDecimal(eTgt.Value)),
        "target_date", Trim(eDate.Value))
        if (isEdit) {
            out := []
            for r in rows {
                if (r["id"] = existing["id"])
                    out.Push(row)
                else
                    out.Push(r)
            }
            rows := out
        } else {
            rows.Push(row)
        }
        Finance_Save("goals", rows)
        saved := true
        g.Destroy()
    }
}
