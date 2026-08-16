; =============================================================================
; Utils module: finance_categories.ahk
; Category CRUD, search, filter, 50/10 limits
; =============================================================================

global g_FinanceCatLv := false
global g_FinanceCatRows := []
global g_FinanceCatFilter := false
global g_FinanceCatType := "all"

Finance_ShowCategories() {
    global g_FinanceGui, g_FinanceCatLv, g_FinanceCatFilter, g_FinanceCatType
    Finance_CloseGui()
    Finance_EnsureData()
    g_FinanceCatType := "all"
    g_FinanceGui := Gui("+AlwaysOnTop +ToolWindow", "Categories")
    g_FinanceGui.SetFont("s10", "Segoe UI")
    g_FinanceGui.Add("Text", "x12 y12", "Search")
    g_FinanceCatFilter := g_FinanceGui.Add("Edit", "x70 y8 w260")
    g_FinanceCatFilter.OnEvent("Change", (*) => Finance_CatRefresh())
    g_FinanceGui.Add("Text", "x350 y12", "[A] all  [E] expense  [N] income  Insert add  F2 edit  Delete")
    g_FinanceCatLv := g_FinanceGui.Add("ListView", "x12 y40 w860 h480 Grid",
        ["Id", "Name", "Type", "Parent", "Color", "Icon"])
    g_FinanceCatLv.OnEvent("DoubleClick", (*) => Finance_CatEdit())
    g_FinanceGui.OnEvent("Close", (*) => Finance_CloseGui())
    g_FinanceGui.OnEvent("Escape", (*) => Finance_ShowMainMenu())
    Finance_CatRefresh()
    Finance_BindHotkeys([
        ["a", Finance_CatFilterAll],
        ["e", Finance_CatFilterExpense],
        ["n", Finance_CatFilterIncome],
        ["Insert", (*) => Finance_CatAdd()],
        ["F2", (*) => Finance_CatEdit()],
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
    global g_FinanceCatLv, g_FinanceCatRows, g_FinanceCatFilter, g_FinanceCatType
    if (!IsObject(g_FinanceCatLv))
        return
    q := IsObject(g_FinanceCatFilter) ? Trim(g_FinanceCatFilter.Value) : ""
    cats := Finance_Load("categories")
    g_FinanceCatLv.Delete()
    g_FinanceCatRows := []
    for c in cats {
        if (g_FinanceCatType != "all" && c["type"] != g_FinanceCatType)
            continue
        if (q != "" && !InStr(c["name"], q) && !InStr(c["id"], q))
            continue
        parentName := ""
        if (c["parent_id"] != "")
            parentName := Finance_CatName(cats, c["parent_id"])
        g_FinanceCatRows.Push(c)
        g_FinanceCatLv.Add("", c["id"], c["name"], c["type"], parentName, c["color"], c["icon"])
    }
    loop 6
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
    cats := Finance_Load("categories")
    if (c["parent_id"] = "") {
        for s in Finance_Subcategories(cats, c["id"]) {
            Finance_Notify("Delete subcategories first", 1800, BANNER_ACCENT_ERROR)
            return
        }
    }
    if (MsgBox("Delete " . c["name"] . "?", "Categories", "YesNo Icon?") != "Yes")
        return
    out := []
    for r in cats {
        if (r["id"] != c["id"])
            out.Push(r)
    }
    Finance_Save("categories", out)
    Finance_CatRefresh()
}

Finance_CatForm(existing) {
    cats := Finance_Load("categories")
    isEdit := IsObject(existing)
    g := Gui("+AlwaysOnTop +ToolWindow", isEdit ? "Edit category" : "Add category")
    g.SetFont("s10", "Segoe UI")
    g.Add("Text", , "Name")
    eName := g.Add("Edit", "w280", isEdit ? existing["name"] : "")
    types := ["expense", "income"]
    tIdx := (isEdit && existing["type"] = "income") ? 2 : 1
    g.Add("Text", "y+8", "Type")
    ddType := g.Add("DropDownList", "w180 Choose" . tIdx, ["expense", "income"])
    mains := Finance_MainCategories(cats)
    parentCombo := Finance_ComboFromRows(mains, "id", "name", true)
    pIdx := Finance_ComboIndex(parentCombo.ids, isEdit ? existing["parent_id"] : "")
    g.Add("Text", "y+8", "Parent (empty = main)")
    ddParent := g.Add("DropDownList", "w280 Choose" . pIdx, parentCombo.names)
    g.Add("Text", "y+8", "Color (#RRGGBB)")
    eColor := g.Add("Edit", "w120", isEdit ? existing["color"] : Finance_ColorForIndex(cats.Length + 1))
    g.Add("Text", "y+8", "Icon")
    eIcon := g.Add("Edit", "w180", isEdit ? existing["icon"] : "tag")
    saved := false
    g.Add("Button", "y+16 w100 Default", "Save").OnEvent("Click", SaveCat)
    g.Add("Button", "x+8 w100", "Cancel").OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())
    g.Show()
    try WinWaitClose("ahk_id " g.Hwnd)
    catch {
    }
    if (saved)
        Finance_CatRefresh()

    SaveCat(*) {
        name := Trim(eName.Value)
        if (name = "") {
            MsgBox("Name is required.", "Categories", "Icon!")
            return
        }
        parentId := parentCombo.ids[ddParent.Value]
        t := ddType.Text
        if (!isEdit) {
            if (parentId = "") {
                if (!Finance_CanAddMainCategory(cats)) {
                    MsgBox("Hard limit: 50 main categories.", "Categories", "Icon!")
                    return
                }
            } else if (!Finance_CanAddSubcategory(cats, parentId)) {
                MsgBox("Hard limit: 10 subcategories per main category.", "Categories", "Icon!")
                return
            }
        }
        color := Trim(eColor.Value)
        if (!RegExMatch(color, "^#[0-9A-Fa-f]{6}$"))
            color := "#7F8C8D"
        id := isEdit ? existing["id"] : Finance_SlugId("CAT_", name, cats)
        newRow := Map("id", id, "name", name, "type", t, "parent_id", parentId,
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
        saved := true
        g.Destroy()
    }
}
