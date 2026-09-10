; =============================================================================
; Utils module: finance_categories.ahk
; Category CRUD, filter, 50-main limit
; =============================================================================

global g_FinanceCatLv := false
global g_FinanceCatRows := []
global g_FinanceCatType := "all"

Finance_ShowCategories() {
    global g_FinanceGui, g_FinanceCatLv, g_FinanceCatType
    Finance_CloseGui()
    Finance_EnsureData()
    g_FinanceCatType := "all"
    g_FinanceGui := Gui("+AlwaysOnTop +ToolWindow", "Categories")
    g_FinanceGui.SetFont("s10", "Segoe UI")
    g_FinanceGui.Add("Text", "x12 y12 w860",
        "[Shift+A] all  [Shift+X] expense  [Shift+N] income  [Shift+I]/Insert add  [Shift+E] edit  Delete  Backspace")
    g_FinanceCatLv := g_FinanceGui.Add("ListView", "x12 y40 w860 h480 Grid",
        ["Name", "Type", "Color"])
    g_FinanceCatLv.OnEvent("DoubleClick", (*) => Finance_CatEdit())
    g_FinanceGui.OnEvent("Close", (*) => Finance_CloseGui())
    g_FinanceGui.OnEvent("Escape", (*) => Finance_ShowMainMenu())
    Finance_CatRefresh()
    Finance_BindHotkeys([
        ["+a", Finance_CatFilterAll],
        ["+x", Finance_CatFilterExpense],
        ["+n", Finance_CatFilterIncome],
        ["+i", (*) => Finance_CatAdd()],
        ["Insert", (*) => Finance_CatAdd()],
        ["+e", (*) => Finance_CatEdit()],
        ["Delete", (*) => Finance_CatDelete()],
        ["Backspace", (*) => Finance_ShowMainMenu()],
        ["Escape", (*) => Finance_ShowMainMenu()]
    ])
    Finance_CenterGui(g_FinanceGui, 890, 560)
}

Finance_CatFilterAll(*) {
    global g_FinanceCatType
    g_FinanceCatType := "all"
    Finance_CatRefresh()
}
Finance_CatFilterExpense(*) {
    global g_FinanceCatType
    g_FinanceCatType := "expense"
    Finance_CatRefresh()
}
Finance_CatFilterIncome(*) {
    global g_FinanceCatType
    g_FinanceCatType := "income"
    Finance_CatRefresh()
}

Finance_CatRefresh() {
    global g_FinanceCatLv, g_FinanceCatRows, g_FinanceCatType
    if (!IsObject(g_FinanceCatLv))
        return
    cats := Finance_Load("categories")
    g_FinanceCatLv.Delete()
    g_FinanceCatRows := []
    for c in cats {
        if (c["parent_id"] != "")
            continue
        if (g_FinanceCatType != "all" && c["type"] != g_FinanceCatType)
            continue
        g_FinanceCatRows.Push(c)
        g_FinanceCatLv.Add("", Finance_CatLabel(c), c["type"], c["color"])
    }
    loop 3
        g_FinanceCatLv.ModifyCol(A_Index, "AutoHdr")
}

Finance_CatSelected() {
    global g_FinanceCatLv, g_FinanceCatRows
    row := g_FinanceCatLv.GetNext()
    if (!row || row > g_FinanceCatRows.Length)
        return false
    return g_FinanceCatRows[row]
}

Finance_CatAdd(*) {
    Finance_CatForm(false)
}

Finance_CatEdit(*) {
    c := Finance_CatSelected()
    if (!c) {
        Finance_Notify("Select a category", 1200, BANNER_ACCENT_ERROR)
        return
    }
    Finance_CatForm(c)
}

Finance_CatDelete(*) {
    c := Finance_CatSelected()
    if (!c)
        return
    if (!Finance_Confirm("Delete " . c["name"] . "?", "Categories"))
        return
    cats := Finance_Load("categories")
    out := []
    for r in cats {
        if (r["id"] != c["id"])
            out.Push(r)
    }
    Finance_Save("categories", out)
    Finance_CatRefresh()
}

Finance_CatForm(existing) {
    global g_FinanceGui
    cats := Finance_Load("categories")
    isEdit := IsObject(existing)
    owner := ""
    try {
        if (IsObject(g_FinanceGui))
            owner := " +Owner" . g_FinanceGui.Hwnd
    } catch {
        owner := ""
    }
    Finance_DialogsBegin()
    g := Gui("+AlwaysOnTop +ToolWindow" . owner, isEdit ? "Edit category" : "Add category")
    g.SetFont("s10", "Segoe UI")
    g.Add("Text", , "Name")
    eName := g.Add("Edit", "w280", isEdit ? existing["name"] : "")
    tIdx := (isEdit && existing["type"] = "income") ? 2 : 1
    g.Add("Text", "y+8", "Type")
    ddType := g.Add("DropDownList", "w180 Choose" . tIdx, ["expense", "income"])
    g.Add("Text", "y+8", "Color (#RRGGBB)")
    eColor := g.Add("Edit", "w120", isEdit ? existing["color"] : Finance_ColorForIndex(cats.Length + 1))
    g.Add("Text", "y+8", "Icon (emoji)")
    eIcon := g.Add("Edit", "w180", isEdit ? existing["icon"] : Finance_DefaultCatIcon(""))
    saved := false
    g.Add("Button", "y+16 w100 Default", "Save").OnEvent("Click", SaveCat)
    g.Add("Button", "x+8 w100", "Cancel").OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())
    g.Show()
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    Finance_DialogsEnd()
    if (saved)
        Finance_CatRefresh()

    SaveCat(*) {
        name := Trim(eName.Value)
        if (name = "") {
            Finance_Alert("Name is required.", "Categories")
            return
        }
        t := ddType.Text
        if (!isEdit && !Finance_CanAddMainCategory(cats)) {
            Finance_Alert("Hard limit: 50 main categories.", "Categories")
            return
        }
        color := Trim(eColor.Value)
        if (!RegExMatch(color, "^#[0-9A-Fa-f]{6}$"))
            color := "#7F8C8D"
        id := isEdit ? existing["id"] : Finance_SlugId("CAT_", name, cats)
        newRow := Map("id", id, "name", name, "type", t, "parent_id", "",
            "color", color, "icon", Trim(eIcon.Value))
        if (isEdit) {
            out := []
            for r in cats {
                if (r["id"] = id)
                    out.Push(newRow)
                else
                    out.Push(r)
            }
            cats := out
        } else {
            cats.Push(newRow)
        }
        Finance_Save("categories", cats)
        Finance_SyncCategoryToIni(name, t)
        saved := true
        g.Destroy()
    }
}
