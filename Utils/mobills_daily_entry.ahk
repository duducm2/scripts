; =============================================================================
; Utils module: mobills_daily_entry.ahk
; #!+U Macros [m] — parse Desktop MOBILLS_V1.txt (or matching .ini), confirm, enter in Mobills, scrape.
; =============================================================================

global g_MobillsDailyConfirmResult := ""
global g_MobillsDailyConfirmGui := ""
global g_MobillsDailyConfirmRows := ""
global g_MobillsDailyConfirmLv := ""
global g_MobillsDailyDdlAccount := ""
global g_MobillsDailyDdlCard := ""
global g_MobillsDailyFavOrigAccount := ""
global g_MobillsDailyFavOrigCard := ""
global g_MobillsDailyDdlBusy := false
global g_MobillsDailyEditResult := ""
global g_MobillsDailyEditing := false
global g_MobillsDailyReviewResult := ""
global g_MobillsDailyReviewGui := ""
global g_MobillsDailySkips := []

MobillsDaily_FavoritesPath() {
    return A_ScriptDir "\assets\data\mobills-favorites.ini"
}

MobillsDaily_CatalogNames(path) {
    names := []
    catalog := MobillsDaily_LoadIniCatalog(path)
    for section, _ in catalog
        names.Push(section)
    return names
}

MobillsDaily_LoadFavorites() {
    global MOBILLS_DEFAULT_ACCOUNT, MOBILLS_CARD_NAME
    acc := MOBILLS_DEFAULT_ACCOUNT
    card := MOBILLS_CARD_NAME
    path := MobillsDaily_FavoritesPath()
    if FileExist(path) {
        try {
            a := Trim(IniRead(path, "Favorites", "Account", acc))
            if (a != "" && a != "ERROR")
                acc := a
        } catch {
        }
        try {
            c := Trim(IniRead(path, "Favorites", "Card", card))
            if (c != "" && c != "ERROR")
                card := c
        } catch {
        }
    }
    MOBILLS_DEFAULT_ACCOUNT := acc
    MOBILLS_CARD_NAME := card
    return { account: acc, card: card }
}

MobillsDaily_SaveFavorites(account, card) {
    global MOBILLS_DEFAULT_ACCOUNT, MOBILLS_CARD_NAME
    account := Trim(account)
    card := Trim(card)
    if (account = "")
        account := MOBILLS_DEFAULT_ACCOUNT
    if (card = "")
        card := MOBILLS_CARD_NAME
    path := MobillsDaily_FavoritesPath()
    dir := A_ScriptDir "\assets\data"
    if !DirExist(dir) {
        try DirCreate(dir)
        catch {
        }
    }
    try IniWrite(account, path, "Favorites", "Account")
    catch {
    }
    try IniWrite(card, path, "Favorites", "Card")
    catch {
    }
    MOBILLS_DEFAULT_ACCOUNT := account
    MOBILLS_CARD_NAME := card
}

MobillsDaily_StripStar(text) {
    text := Trim(text)
    if (SubStr(text, 1, 2) = "★ ")
        return Trim(SubStr(text, 3))
    if (SubStr(text, 1, 1) = "★")
        return Trim(SubStr(text, 2))
    return text
}

MobillsDaily_FillStarredDdl(ddl, names, selected) {
    ddl.Delete()
    choose := 1
    i := 1
    list := []
    found := false
    for n in names {
        list.Push(n)
        if (n = selected)
            found := true
    }
    if (selected != "" && !found)
        list.Push(selected)
    items := []
    for n in list {
        items.Push((n = selected) ? ("★ " n) : n)
        if (n = selected)
            choose := i
        i++
    }
    if items.Length
        ddl.Add(items)
    if items.Length {
        try ddl.Choose(choose)
        catch {
        }
    }
}

MobillsDaily_FillPlainDdl(ddl, names, selected, includeBlank := false) {
    ddl.Delete()
    items := []
    blankLabel := "(none)"
    sel := selected
    if includeBlank {
        items.Push(blankLabel)
        if (selected = "")
            sel := blankLabel
    }
    found := (sel = blankLabel || sel = "")
    for n in names {
        items.Push(n)
        if (n = selected)
            found := true
    }
    if (selected != "" && !found)
        items.Push(selected)
    choose := 1
    i := 1
    for it in items {
        if (it = sel)
            choose := i
        i++
    }
    if items.Length
        ddl.Add(items)
    if items.Length {
        try ddl.Choose(choose)
        catch {
        }
    }
}

MobillsDaily_RemapDefaultedRows(rows, oldAcc, newAcc, oldCard, newCard) {
    if (newAcc = "")
        newAcc := oldAcc
    if (newCard = "")
        newCard := oldCard
    for row in rows {
        src := row.HasProp("origSource") ? row.origSource : row.source
        tgt := row.HasProp("origTarget") ? row.origTarget : row.target
        if (row.type = "EXPENSE") {
            row.source := (src = "" || src = oldAcc) ? newAcc : src
            row.target := tgt
        } else if (row.type = "INCOME") {
            row.source := src
            if (tgt = "" || tgt = oldAcc)
                row.target := newAcc
            else
                row.target := tgt
            if (row.target = "" && (src = "" || src = oldAcc))
                row.source := newAcc
        } else if (row.type = "TRANSFER") {
            row.source := (src = "" || src = oldAcc) ? newAcc : src
            row.target := (tgt = "" || tgt = oldAcc) ? newAcc : tgt
        } else if (row.type = "CARD") {
            row.source := (src = "" || src = oldCard) ? newCard : src
            row.target := tgt
        } else {
            row.source := src
            row.target := tgt
        }
    }
}

MobillsDaily_RefreshConfirmLv() {
    global g_MobillsDailyConfirmLv, g_MobillsDailyConfirmRows
    lv := g_MobillsDailyConfirmLv
    rows := g_MobillsDailyConfirmRows
    if (!IsObject(lv) || !IsObject(rows))
        return
    try lv.Delete()
    catch {
    }
    i := 1
    for row in rows {
        mark := row.flags.Length ? "! " : ""
        lv.Add("", mark i, row.type, row.description, row.value, row.source, row.target, MobillsDaily_CategoryLabel(row
        ))
        i++
    }
}

MobillsDaily_LoadIniCatalog(path) {
    catalog := Map()
    if (!FileExist(path))
        return catalog
    raw := ""
    try raw := FileRead(path, "UTF-8")
    catch {
        return catalog
    }
    current := ""
    loop parse raw, "`n", "`r" {
        line := Trim(A_LoopField)
        if (line = "" || SubStr(line, 1, 1) = ";" || SubStr(line, 1, 1) = "#")
            continue
        if (RegExMatch(line, "^\[(.+)\]$", &m)) {
            current := m[1]
            catalog[current] := { desc: "", subs: Map() }
            continue
        }
        if (current = "")
            continue
        eq := InStr(line, "=")
        if (!eq)
            continue
        k := Trim(SubStr(line, 1, eq - 1))
        v := Trim(SubStr(line, eq + 1))
        if (k = "geral")
            catalog[current].desc := v
        else
            catalog[current].subs[k] := v
    }
    return catalog
}

MobillsDaily_RenderIniCatalog(path, heading) {
    catalog := MobillsDaily_LoadIniCatalog(path)
    if (!catalog.Count)
        return heading . ":`n(file missing: " path ")"
    out := heading . ":"
    for section, rec in catalog {
        out .= "`n- " . section
        if (rec.desc != "")
            out .= " — " . rec.desc
        for sub, sdesc in rec.subs {
            out .= "`n  - " . section . " > " . sub
            if (sdesc != "")
                out .= " — " . sdesc
        }
    }
    return out
}

MobillsDaily_CatalogHasSection(catalog, name) {
    if (name = "" || !catalog)
        return false
    if catalog.Has(name)
        return true
    for section, rec in catalog {
        if (StrLower(section) = StrLower(name))
            return true
        for sub, _ in rec.subs {
            if (StrLower(sub) = StrLower(name))
                return true
        }
    }
    return false
}

MobillsDaily_MatchAccountName(catalog, uiName) {
    uiName := Trim(uiName)
    if (uiName = "")
        return ""
    uiName := RegExReplace(uiName, "\.{2,}$", "")
    uiName := Trim(uiName)
    if catalog.Has(uiName)
        return uiName
    best := ""
    bestLen := 0
    for section, _ in catalog {
        if (InStr(section, uiName) = 1 || InStr(uiName, section) = 1) {
            if (StrLen(section) > bestLen) {
                best := section
                bestLen := StrLen(section)
            }
        }
    }
    return best
}

MobillsDaily_DesktopFolders() {
    folders := []
    seen := Map()
    candidates := []
    try candidates.Push(GetDesktopToRecyclePath())
    catch {
    }
    candidates.Push(A_Desktop)
    try candidates.Push(EnvGet("USERPROFILE") "\OneDrive\Desktop")
    catch {
    }
    try candidates.Push(EnvGet("USERPROFILE") "\Desktop")
    catch {
    }
    for p in candidates {
        p := RTrim(p, "\")
        if (p = "" || !DirExist(p))
            continue
        key := StrLower(p)
        if seen.Has(key)
            continue
        seen[key] := true
        folders.Push(p)
    }
    return folders
}

MobillsDaily_DesktopPath() {
    folders := MobillsDaily_DesktopFolders()
    return folders.Length ? folders[1] : A_Desktop
}

; Desktop payload: MOBILLS_V1 pipe table (what Gemini used to save as .ini) or classic INI.
MobillsDaily_FileLooksLikeTable(path) {
    sz := 0
    try sz := FileGetSize(path)
    catch {
        return false
    }
    if (sz < 8 || sz > 524288)
        return false
    raw := MobillsDaily_ReadText(path)
    return (InStr(raw, "MOBILLS_V1") || InStr(raw, "TYPE|DESCRIPTION"))
}

MobillsDaily_FileLooksLikeIniSyntax(path) {
    if MobillsDaily_FileLooksLikeTable(path)
        return true
    sz := 0
    try sz := FileGetSize(path)
    catch {
        return false
    }
    if (sz < 8 || sz > 524288)
        return false
    raw := MobillsDaily_ReadText(path)
    hasSection := false
    hasAssign := false
    loop parse raw, "`n", "`r" {
        line := Trim(A_LoopField)
        if (line = "" || SubStr(line, 1, 1) = ";" || SubStr(line, 1, 1) = "#")
            continue
        if RegExMatch(line, "^\[.+\]$") {
            hasSection := true
            continue
        }
        if (InStr(line, "=") && !InStr(line, "|"))
            hasAssign := true
        parts := StrSplit(line, "|")
        if (parts.Length >= 7)
            return true
    }
    return hasSection && hasAssign
}

; Newest Desktop file for glob (e.g. "*.ini"). skipDesktopIni; if requireIniSyntax, only valid payloads.
MobillsDaily_NewestDesktopByGlob(glob, requireIniSyntax := false) {
    newestPath := ""
    newestStamp := ""
    for folder in MobillsDaily_DesktopFolders() {
        loop files folder "\" glob, "F" {
            if (StrLower(A_LoopFileName) = "desktop.ini")
                continue
            stamp := ""
            try stamp := FileGetTime(A_LoopFileFullPath, "M")
            catch {
                continue
            }
            if (requireIniSyntax && !MobillsDaily_FileLooksLikeIniSyntax(A_LoopFileFullPath))
                continue
            if (newestStamp = "" || stamp > newestStamp) {
                newestStamp := stamp
                newestPath := A_LoopFileFullPath
            }
        }
    }
    return newestPath
}

; 1) Desktop .ini with valid table/INI payload (skip desktop.ini).
; 2) Else .txt with the same payload. Option B: parse as-is (do not rename — Gemini .ini downloads were truncated).
MobillsDaily_FindLatestFile() {
    iniMatch := MobillsDaily_NewestDesktopByGlob("*.ini", true)
    if (iniMatch != "")
        return iniMatch
    txtMatch := MobillsDaily_NewestDesktopByGlob("*.txt", true)
    if (txtMatch != "")
        return txtMatch
    iniAny := MobillsDaily_NewestDesktopByGlob("*.ini", false)
    if (iniAny != "")
        return iniAny
    return MobillsDaily_NewestDesktopByGlob("*.txt", false)
}

MobillsDaily_ReadText(path) {
    raw := ""
    try raw := FileRead(path, "UTF-8")
    catch {
        raw := ""
    }
    if (InStr(raw, Chr(0)) || (raw != "" && InStr(raw, "|") = 0)) {
        try {
            raw16 := FileRead(path, "UTF-16")
            if (InStr(raw16, "|") && (!InStr(raw, "|") || StrLen(raw16) > StrLen(raw)))
                raw := raw16
        } catch {
        }
    }
    return raw
}

MobillsDaily_ParseFile(path) {
    result := { rows: [], unmapped: [], error: "", rawLen: 0, firstLine: "" }
    raw := MobillsDaily_ReadText(path)
    result.rawLen := StrLen(raw)
    loop parse raw, "`n", "`r" {
        if (Trim(A_LoopField) != "") {
            result.firstLine := Trim(A_LoopField)
            break
        }
    }
    if (raw = "") {
        result.error := "File is empty: " path
        return result
    }
    parsed := MobillsDaily_ParseRaw(raw)
    parsed.rawLen := result.rawLen
    parsed.firstLine := result.firstLine
    return parsed
}

MobillsDaily_ParseRaw(raw) {
    result := { rows: [], unmapped: [], error: "", rawLen: StrLen(raw), firstLine: "" }
    fav := MobillsDaily_LoadFavorites()
    accounts := MobillsDaily_LoadIniCatalog(A_ScriptDir "\accounts.ini")
    expenses := MobillsDaily_LoadIniCatalog(A_ScriptDir "\categories-expenses.ini")
    incomes := MobillsDaily_LoadIniCatalog(A_ScriptDir "\categories-income.ini")
    defAcc := fav.account
    defCard := fav.card
    started := true
    loop parse raw, "`n", "`r" {
        line := Trim(A_LoopField)
        if (line = "")
            continue
        if (line = "MOBILLS_V1") {
            started := true
            continue
        }
        if (line = "END")
            break
        if (SubStr(line, 1, 1) = "#") {
            result.unmapped.Push(line)
            continue
        }
        if (InStr(line, "TYPE|DESCRIPTION|VALUE") = 1)
            continue
        parts := StrSplit(line, "|")
        if (parts.Length < 7) {
            result.unmapped.Push("# UNMAPPED: bad column count: " line)
            continue
        }
        typ := StrUpper(Trim(parts[1]))
        desc := Trim(parts[2])
        val := Trim(parts[3])
        source := Trim(parts[4])
        target := Trim(parts[5])
        cat := Trim(parts[6])
        sub := Trim(parts[7])
        if (typ = "CARD" && source = "")
            source := defCard
        if (typ = "EXPENSE" && source = "")
            source := defAcc
        if (typ = "INCOME" && target = "" && source = "")
            target := defAcc
        if (typ = "TRANSFER") {
            if (source = "")
                source := defAcc
            if (target = "")
                target := defAcc
        }
        flags := []
        if !(typ = "EXPENSE" || typ = "INCOME" || typ = "CARD" || typ = "TRANSFER")
            flags.Push("bad type")
        if !RegExMatch(val, "^[0-9]+([.][0-9]+)?$")
            flags.Push("bad value")
        if (typ = "EXPENSE" || typ = "TRANSFER") {
            if (source != "" && !MobillsDaily_CatalogHasSection(accounts, source))
                flags.Push("unknown source account")
        }
        if (typ = "INCOME") {
            dest := (target != "") ? target : source
            if (dest != "" && !MobillsDaily_CatalogHasSection(accounts, dest))
                flags.Push("unknown dest account")
        }
        if (typ = "TRANSFER") {
            if (target != "" && !MobillsDaily_CatalogHasSection(accounts, target))
                flags.Push("unknown dest account")
        }
        if (typ = "EXPENSE" || typ = "CARD") {
            if (cat != "" && !MobillsDaily_CatalogHasSection(expenses, cat) && !MobillsDaily_CatalogHasSection(expenses,
                sub))
                flags.Push("unknown expense category")
        }
        if (typ = "INCOME") {
            if (cat != "" && !MobillsDaily_CatalogHasSection(incomes, cat) && !MobillsDaily_CatalogHasSection(incomes,
                sub))
                flags.Push("unknown income category")
        }
        result.rows.Push({
            type: typ, description: desc, value: val, source: source, target: target,
            origSource: source, origTarget: target,
            category: cat, subcategory: sub, flags: flags, raw: line
        })
    }
    return result
}

MobillsDaily_CategoryLabel(row) {
    if (row.subcategory != "")
        return row.category . " > " . row.subcategory
    return row.category
}

MobillsDaily_ShowConfirmTable(rows, unmapped, filePath) {
    global g_MobillsDailyConfirmResult, g_MobillsDailyConfirmGui, g_MobillsDailyConfirmRows, g_MobillsDailyConfirmLv
    global g_MobillsDailyDdlAccount, g_MobillsDailyDdlCard, g_MobillsDailyFavOrigAccount, g_MobillsDailyFavOrigCard
    global g_MobillsDailyDdlBusy
    g_MobillsDailyConfirmResult := ""
    g_MobillsDailyDdlBusy := false
    if (IsObject(g_MobillsDailyConfirmGui)) {
        try g_MobillsDailyConfirmGui.Destroy()
        catch {
        }
    }
    fav := MobillsDaily_LoadFavorites()
    g_MobillsDailyFavOrigAccount := fav.account
    g_MobillsDailyFavOrigCard := fav.card
    g_MobillsDailyConfirmRows := rows
    for row in rows {
        if !row.HasProp("origSource")
            row.origSource := row.source
        if !row.HasProp("origTarget")
            row.origTarget := row.target
    }
    hint := "File: " filePath "`nDouble-click a row or Edit to fix a cell."
    if (unmapped.Length)
        hint .= "`nUnmapped lines: " unmapped.Length . " (see rows marked ! )"
    dlg := Gui("+AlwaysOnTop +ToolWindow", "Mobills daily entry")
    dlg.SetFont("s10", "Segoe UI")
    dlg.Add("Text", "w900", hint)
    dlg.Add("Text", "xm w140 Section", "★ Main account")
    ddlAcc := dlg.Add("DropDownList", "ys w360")
    dlg.Add("Text", "ys w160", "★ Main credit card")
    ddlCard := dlg.Add("DropDownList", "ys w280")
    g_MobillsDailyDdlAccount := ddlAcc
    g_MobillsDailyDdlCard := ddlCard
    lv := dlg.Add("ListView", "xm w900 h360 -Multi", ["#", "Type", "Description", "Value", "Source", "Target",
        "Category"])
    g_MobillsDailyConfirmLv := lv
    try lv.ModifyCol(1, 40)
    try lv.ModifyCol(2, 90)
    try lv.ModifyCol(3, 220)
    try lv.ModifyCol(4, 80)
    try lv.ModifyCol(5, 180)
    try lv.ModifyCol(6, 160)
    try lv.ModifyCol(7, 160)
    g_MobillsDailyDdlBusy := true
    MobillsDaily_FillStarredDdl(ddlAcc, MobillsDaily_CatalogNames(A_ScriptDir "\accounts.ini"), fav.account)
    MobillsDaily_FillStarredDdl(ddlCard, MobillsDaily_CatalogNames(A_ScriptDir "\cards.ini"), fav.card)
    g_MobillsDailyDdlBusy := false
    ddlAcc.OnEvent("Change", MobillsDaily_ConfirmFavChanged)
    ddlCard.OnEvent("Change", MobillsDaily_ConfirmFavChanged)
    lv.OnEvent("DoubleClick", MobillsDaily_ConfirmLvDoubleClick)
    MobillsDaily_RefreshConfirmLv()
    dlg.Add("Button", "xm w100 Section Default", "OK").OnEvent("Click", MobillsDaily_ConfirmOk)
    dlg.Add("Button", "w100 ys", "Edit").OnEvent("Click", MobillsDaily_ConfirmEditClick)
    dlg.Add("Button", "w100 ys", "Cancel").OnEvent("Click", MobillsDaily_ConfirmCancel)
    dlg.OnEvent("Close", MobillsDaily_ConfirmCancel)
    dlg.OnEvent("Escape", MobillsDaily_ConfirmCancel)
    g_MobillsDailyConfirmGui := dlg
    dlg.Show()
    try lv.Focus()
    catch {
    }
    start := A_TickCount
    while (g_MobillsDailyConfirmResult = "") {
        if ((A_TickCount - start) >= 300000) {
            g_MobillsDailyConfirmResult := "cancel"
            break
        }
        Sleep 50
    }
    res := g_MobillsDailyConfirmResult
    g_MobillsDailyConfirmResult := ""
    try dlg.Destroy()
    catch {
    }
    g_MobillsDailyConfirmGui := ""
    g_MobillsDailyConfirmLv := ""
    g_MobillsDailyDdlAccount := ""
    g_MobillsDailyDdlCard := ""
    return res = "ok"
}

MobillsDaily_ConfirmLvDoubleClick(lv, rowNum, *) {
    if (rowNum < 1) {
        try rowNum := lv.GetNext(0, "Focused")
        catch {
            rowNum := 0
        }
    }
    if (rowNum < 1)
        return
    MobillsDaily_EditConfirmRow(rowNum)
}

MobillsDaily_ConfirmEditClick(*) {
    global g_MobillsDailyConfirmLv
    rowNum := 0
    if IsObject(g_MobillsDailyConfirmLv) {
        try rowNum := g_MobillsDailyConfirmLv.GetNext(0, "Focused")
        catch {
        }
        if (rowNum < 1) {
            try rowNum := g_MobillsDailyConfirmLv.GetNext(0, "Selected")
            catch {
            }
        }
    }
    if (rowNum < 1) {
        ShowCenteredOverlay_Utils("Select a row to edit", 1500, BANNER_ACCENT_INTERMEDIATE)
        return
    }
    MobillsDaily_EditConfirmRow(rowNum)
}

MobillsDaily_ParseCategoryLabel(label) {
    label := Trim(label)
    sep := InStr(label, " > ")
    if (sep)
        return { category: Trim(SubStr(label, 1, sep - 1)), subcategory: Trim(SubStr(label, sep + 3)) }
    return { category: label, subcategory: "" }
}

MobillsDaily_EditConfirmRow(rowNum) {
    global g_MobillsDailyConfirmRows, g_MobillsDailyEditResult, g_MobillsDailyConfirmLv
    global g_MobillsDailyConfirmGui, g_MobillsDailyEditing
    if (!IsObject(g_MobillsDailyConfirmRows) || rowNum < 1 || rowNum > g_MobillsDailyConfirmRows.Length)
        return
    row := g_MobillsDailyConfirmRows[rowNum]
    g_MobillsDailyEditResult := ""
    g_MobillsDailyEditing := true
    if IsObject(g_MobillsDailyConfirmGui) {
        try g_MobillsDailyConfirmGui.Opt("+Disabled")
        catch {
        }
    }
    accNames := MobillsDaily_CatalogNames(A_ScriptDir "\accounts.ini")
    cardNames := MobillsDaily_CatalogNames(A_ScriptDir "\cards.ini")
    sourceNames := []
    for n in accNames
        sourceNames.Push(n)
    for n in cardNames
        sourceNames.Push(n)
    dlg := Gui("+AlwaysOnTop +Owner", "Edit transaction #" rowNum)
    dlg.SetFont("s10", "Segoe UI")
    dlg.Add("Text", "w80 Section", "Type")
    types := ["EXPENSE", "INCOME", "CARD", "TRANSFER"]
    ddlType := dlg.Add("DropDownList", "ys w200", types)
    ti := 1
    for t in types {
        if (t = row.type) {
            try ddlType.Choose(ti)
            catch {
            }
            break
        }
        ti++
    }
    dlg.Add("Text", "xm w80 Section", "Description")
    edDesc := dlg.Add("Edit", "ys w400", row.description)
    dlg.Add("Text", "xm w80 Section", "Value")
    edVal := dlg.Add("Edit", "ys w120", row.value)
    dlg.Add("Text", "xm w80 Section", "Source")
    ddlSrc := dlg.Add("DropDownList", "ys w400")
    MobillsDaily_FillPlainDdl(ddlSrc, sourceNames, row.source, false)
    dlg.Add("Text", "xm w80 Section", "Target")
    ddlTgt := dlg.Add("DropDownList", "ys w400")
    MobillsDaily_FillPlainDdl(ddlTgt, accNames, row.target, true)
    dlg.Add("Text", "xm w80 Section", "Category")
    edCat := dlg.Add("Edit", "ys w400", MobillsDaily_CategoryLabel(row))
    dlg.Add("Text", "xm w480", "Category as Section or Section > subcategory")
    dlg.Add("Button", "xm w100 Section Default", "OK").OnEvent("Click", MobillsDaily_EditOk)
    dlg.Add("Button", "w100 ys", "Cancel").OnEvent("Click", MobillsDaily_EditCancel)
    dlg.OnEvent("Close", MobillsDaily_EditCancel)
    dlg.OnEvent("Escape", MobillsDaily_EditCancel)
    dlg.Show()
    try edDesc.Focus()
    catch {
    }
    start := A_TickCount
    while (g_MobillsDailyEditResult = "") {
        if ((A_TickCount - start) >= 300000) {
            g_MobillsDailyEditResult := "cancel"
            break
        }
        Sleep 50
    }
    if (g_MobillsDailyEditResult = "ok") {
        row.type := StrUpper(Trim(ddlType.Text))
        row.description := Trim(edDesc.Text)
        row.value := Trim(edVal.Text)
        row.source := Trim(ddlSrc.Text)
        tgt := Trim(ddlTgt.Text)
        row.target := (tgt = "(none)") ? "" : tgt
        parsedCat := MobillsDaily_ParseCategoryLabel(edCat.Text)
        row.category := parsedCat.category
        row.subcategory := parsedCat.subcategory
        row.origSource := row.source
        row.origTarget := row.target
        row.flags := []
        MobillsDaily_RefreshConfirmLv()
        if IsObject(g_MobillsDailyConfirmLv) {
            try g_MobillsDailyConfirmLv.Modify(rowNum, "Select Focus Vis")
            catch {
            }
        }
    }
    g_MobillsDailyEditResult := ""
    g_MobillsDailyEditing := false
    try dlg.Destroy()
    catch {
    }
    if IsObject(g_MobillsDailyConfirmGui) {
        try g_MobillsDailyConfirmGui.Opt("-Disabled")
        catch {
        }
        try WinActivate("ahk_id " g_MobillsDailyConfirmGui.Hwnd)
        catch {
        }
    }
}

MobillsDaily_EditOk(*) {
    global g_MobillsDailyEditResult
    if (g_MobillsDailyEditResult != "")
        return
    g_MobillsDailyEditResult := "ok"
}

MobillsDaily_EditCancel(*) {
    global g_MobillsDailyEditResult
    if (g_MobillsDailyEditResult != "")
        return
    g_MobillsDailyEditResult := "cancel"
}

MobillsDaily_ConfirmFavChanged(*) {
    global g_MobillsDailyDdlBusy
    if g_MobillsDailyDdlBusy
        return
    MobillsDaily_ApplyConfirmFavorites(false)
}

MobillsDaily_ApplyConfirmFavorites(save) {
    global g_MobillsDailyConfirmRows, g_MobillsDailyDdlAccount, g_MobillsDailyDdlCard
    global g_MobillsDailyFavOrigAccount, g_MobillsDailyFavOrigCard, g_MobillsDailyDdlBusy
    acc := MobillsDaily_StripStar(IsObject(g_MobillsDailyDdlAccount) ? g_MobillsDailyDdlAccount.Text : "")
    card := MobillsDaily_StripStar(IsObject(g_MobillsDailyDdlCard) ? g_MobillsDailyDdlCard.Text : "")
    if (acc = "")
        acc := g_MobillsDailyFavOrigAccount
    if (card = "")
        card := g_MobillsDailyFavOrigCard
    if IsObject(g_MobillsDailyConfirmRows)
        MobillsDaily_RemapDefaultedRows(g_MobillsDailyConfirmRows, g_MobillsDailyFavOrigAccount, acc,
            g_MobillsDailyFavOrigCard, card)
    MobillsDaily_RefreshConfirmLv()
    g_MobillsDailyDdlBusy := true
    if IsObject(g_MobillsDailyDdlAccount)
        MobillsDaily_FillStarredDdl(g_MobillsDailyDdlAccount, MobillsDaily_CatalogNames(A_ScriptDir "\accounts.ini"),
        acc)
    if IsObject(g_MobillsDailyDdlCard)
        MobillsDaily_FillStarredDdl(g_MobillsDailyDdlCard, MobillsDaily_CatalogNames(A_ScriptDir "\cards.ini"), card)
    g_MobillsDailyDdlBusy := false
    if save
        MobillsDaily_SaveFavorites(acc, card)
}

MobillsDaily_ConfirmOk(*) {
    global g_MobillsDailyConfirmResult, g_MobillsDailyEditing
    if g_MobillsDailyEditing
        return
    if (g_MobillsDailyConfirmResult != "")
        return
    MobillsDaily_ApplyConfirmFavorites(true)
    g_MobillsDailyConfirmResult := "ok"
}

MobillsDaily_ConfirmCancel(*) {
    global g_MobillsDailyConfirmResult, g_MobillsDailyEditing
    if g_MobillsDailyEditing {
        MobillsDaily_EditCancel()
        return
    }
    if (g_MobillsDailyConfirmResult != "")
        return
    g_MobillsDailyConfirmResult := "cancel"
}

MobillsDaily_Halt(msg) {
    try FileAppend(FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss") "`t" msg "`n", A_ScriptDir "\mobills-run-error.log",
    "UTF-8")
    catch {
    }
    dlg := Gui("+AlwaysOnTop +ToolWindow", "Mobills daily entry — halted")
    dlg.SetFont("s10", "Segoe UI")
    dlg.Add("Text", "w720", "The run stopped. Nothing after this transaction was entered.")
    edit := dlg.Add("Edit", "w720 h260 ReadOnly -Wrap", msg)
    dlg.Add("Button", "w100 Default", "Close").OnEvent("Click", (*) => dlg.Destroy())
    dlg.OnEvent("Close", (*) => dlg.Destroy())
    dlg.OnEvent("Escape", (*) => dlg.Destroy())
    dlg.Show()
    try edit.Focus()
    catch {
    }
    WinWaitClose("ahk_id " dlg.Hwnd)
}

MobillsDaily_SkipRow(uia, row, idx, field, gate, wanted := "", extra := "") {
    global g_MobillsDailySkips
    note := "Enter this " StrLower(row.type) " manually; " field " failed"
    if (wanted != "")
        note .= " (wanted " wanted ")"
    skip := {
        idx: idx, type: row.type, description: row.description, value: row.value,
        field: field, wanted: wanted, note: note, gate: gate
    }
    if !IsObject(g_MobillsDailySkips)
        g_MobillsDailySkips := []
    g_MobillsDailySkips.Push(skip)
    msg := "SKIP #" idx " " row.type " " row.description " / " field " / " gate
    if (wanted != "")
        msg .= " wanted=" wanted
    if (extra != "")
        msg .= " " extra
    try FileAppend(FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss") "`t" msg "`n", A_ScriptDir "\mobills-run-error.log",
    "UTF-8")
    catch {
    }
    MobillsDaily_DismissDialog(uia)
    return "skip"
}

MobillsDaily_DismissDialog(uia) {
    hwnd := 0
    try hwnd := MobillsAuto_FindHwnd()
    catch {
    }
    if hwnd {
        try WinActivate("ahk_id " hwnd)
        catch {
        }
    }
    Send "{Escape}"
    Sleep MOBILLS_STEP_MS
    if MobillsDaily_DialogGone(uia)
        return
    Send "{Escape}"
    Sleep MOBILLS_STEP_MS
    if MobillsDaily_DialogGone(uia)
        return
    scope := ""
    try scope := MobillsAuto_FindDialog(uia)
    catch {
    }
    if !scope
        scope := uia
    for nm in ["Cancel", "CANCEL", "Close", "CLOSE"] {
        btn := ""
        try btn := scope.FindFirst({ Type: 50000, Name: nm })
        catch {
        }
        if btn {
            try MobillsAuto_ClickLeft(btn)
            catch {
                try MobillsAuto_Click(btn)
                catch {
                }
            }
            Sleep MOBILLS_STEP_MS
            if MobillsDaily_DialogGone(uia)
                return
        }
    }
    MobillsAuto_WaitFor(MobillsDaily_DialogGone.Bind(uia))
}

MobillsDaily_ExpectedTitle(typ) {
    switch typ {
        case "EXPENSE": return "New expense"
        case "INCOME": return "New income"
        case "CARD": return "New credit card expense"
        case "TRANSFER": return "New transfer"
        default: return ""
    }
}

MobillsDaily_MenuItem(typ) {
    switch typ {
        case "EXPENSE": return "Expense"
        case "INCOME": return "Income"
        case "CARD": return "Credit card expense"
        case "TRANSFER": return "Transfer"
        default: return ""
    }
}

MobillsDaily_DialogGone(uia) {
    return (MobillsAuto_DialogTitle(uia) = "") ? true : ""
}

MobillsDaily_TitleIfExpected(uia, expected) {
    d := MobillsAuto_FindDialog(uia)
    t := MobillsAuto_DialogTitle(d)
    if (t = "")
        t := MobillsAuto_DialogTitle(uia)
    return (t = expected) ? expected : ""
}

MobillsDaily_SaveIfEnabled(scope) {
    b := MobillsAuto_FindSaveButton(scope)
    if (b && !MobillsAuto_IsDisabled(b))
        return b
    return ""
}

MobillsDaily_EnterRow(uia, row, idx, total) {
    StandardLoadingBar_Update("⏳ " idx "/" total " — " row.type ": " row.description " R$ " row.value)
    menuName := MobillsDaily_MenuItem(row.type)
    opened := MobillsAuto_SelectNewMenuItem(uia, menuName)
    if (!opened.ok)
        return MobillsDaily_SkipRow(uia, row, idx, "New menu", "Could not click + then " menuName, menuName)
    expected := MobillsDaily_ExpectedTitle(row.type)
    title := MobillsAuto_WaitFor(MobillsDaily_TitleIfExpected.Bind(uia, expected))
    if (title != expected)
        return MobillsDaily_SkipRow(uia, row, idx, "Dialog title", "Expected '" expected "'", expected, "got '" title "'"
        )
    dialog := MobillsAuto_FindDialog(uia)
    scope := dialog ? dialog : uia
    Sleep MOBILLS_STEP_MS

    amount := MobillsAuto_FindAmountEdit(scope)
    if !amount
        return MobillsDaily_SkipRow(uia, row, idx, "Value", "Amount edit not found", row.value)
    setAmt := MobillsAuto_SetEditVerified(amount, row.value)
    if (!setAmt.ok)
        return MobillsDaily_SkipRow(uia, row, idx, "Value", "Read-back mismatch", row.value, "got " setAmt.got)

    if (row.type = "TRANSFER") {
        combos := MobillsAuto_DialogCombos(scope)
        originEdit := MobillsAuto_FindNamedEdit(scope, "Origin account")
        destEdit := MobillsAuto_FindNamedEdit(scope, "Destination account")
        originCombo := (combos.Length >= 1) ? combos[1] : ""
        destCombo := (combos.Length >= 2) ? combos[2] : ""
        if (originCombo) {
            pk := MobillsAuto_PickAutocomplete(originCombo, row.source, uia)
            if (!pk.ok)
                return MobillsDaily_SkipRow(uia, row, idx, "Origin account", "Autocomplete mismatch", row.source,
                    "got " pk.got)
        } else if (originEdit) {
            setO := MobillsAuto_SetEditVerified(originEdit, row.source)
            if (!setO.ok)
                return MobillsDaily_SkipRow(uia, row, idx, "Origin account", "Edit mismatch", row.source, "got " setO.got
                )
        } else {
            return MobillsDaily_SkipRow(uia, row, idx, "Origin account", "Field not found", row.source)
        }
        if (destCombo) {
            pk := MobillsAuto_PickAutocomplete(destCombo, row.target, uia)
            if (!pk.ok)
                return MobillsDaily_SkipRow(uia, row, idx, "Destination account", "Autocomplete mismatch", row.target,
                    "got " pk.got)
        } else if (destEdit) {
            setD := MobillsAuto_SetEditVerified(destEdit, row.target)
            if (!setD.ok)
                return MobillsDaily_SkipRow(uia, row, idx, "Destination account", "Edit mismatch", row.target, "got " setD
                    .got)
        } else {
            return MobillsDaily_SkipRow(uia, row, idx, "Destination account", "Field not found", row.target)
        }
    } else {
        descEl := MobillsAuto_FindNamedEdit(scope, "Description")
        if !descEl
            return MobillsDaily_SkipRow(uia, row, idx, "Description", "Edit not found", row.description)
        setD := MobillsAuto_SetEditVerified(descEl, row.description)
        if (!setD.ok)
            return MobillsDaily_SkipRow(uia, row, idx, "Description", "Read-back mismatch", row.description, "got " setD
                .got)
        combos := MobillsAuto_DialogCombos(scope)
        catCombo := (combos.Length >= 2) ? combos[2] : ""
        acctCombo := (combos.Length >= 3) ? combos[3] : ""
        catWanted := row.subcategory != "" ? row.subcategory : row.category
        if (catCombo && catWanted != "") {
            pk := MobillsAuto_PickAutocomplete(catCombo, catWanted, uia)
            if (!pk.ok && row.subcategory != "" && row.category != "")
                pk := MobillsAuto_PickAutocomplete(catCombo, row.category, uia)
            if (!pk.ok)
                return MobillsDaily_SkipRow(uia, row, idx, "Category", "Autocomplete mismatch", catWanted, "got " pk.got
                )
        }
        acctWanted := (row.type = "CARD") ? MOBILLS_CARD_NAME : ((row.type = "INCOME") ? (row.target != "" ? row.target :
            row.source) : row.source)
        if (acctCombo && acctWanted != "") {
            pk := MobillsAuto_PickAutocomplete(acctCombo, acctWanted, uia)
            if (!pk.ok)
                return MobillsDaily_SkipRow(uia, row, idx, (row.type = "CARD") ? "Card" : "Account",
                "Autocomplete mismatch",
                acctWanted, "got " pk.got)
        }
    }

    saveBtn := MobillsAuto_WaitFor(MobillsDaily_SaveIfEnabled.Bind(scope))
    if !saveBtn
        return MobillsDaily_SkipRow(uia, row, idx, "SAVE", "Button stayed disabled or missing", row.description)
    if !MobillsAuto_Click(saveBtn)
        return MobillsDaily_SkipRow(uia, row, idx, "SAVE", "Click failed", row.description)
    gone := MobillsAuto_WaitFor(MobillsDaily_DialogGone.Bind(uia))
    if !gone
        return MobillsDaily_SkipRow(uia, row, idx, "SAVE", "Dialog still open after SAVE", expected)
    Sleep MOBILLS_STEP_MS
    return true
}

MobillsDaily_IsMoneyText(nm) {
    return (InStr(nm, "R$") = 1 || RegExMatch(nm, "^R\$"))
}

MobillsDaily_ScrapeAccounts(uia) {
    catalog := MobillsDaily_LoadIniCatalog(A_ScriptDir "\accounts.ini")
    nav := MobillsAuto_ClickMenu(uia, "menu-accounts-item", "Accounts")
    if (!nav.ok)
        return { ok: false, rows: [], error: "Could not open Accounts page", attempted: nav.attempted }
    Sleep MOBILLS_STEP_MS * 2
    all := []
    seen := Map()
    seenPages := 0
    loop 8 {
        seenPages++
        StandardLoadingBar_Update("⏳ Scraping accounts page " seenPages "...")
        pageRows := MobillsDaily_ScrapeAccountsPage(uia, catalog)
        for r in pageRows {
            key := StrLower(r.name)
            if seen.Has(key)
                continue
            seen[key] := true
            all.Push(r)
        }
        fp := MobillsDaily_FirstAccountTitle(uia, catalog)
        nxt := MobillsAuto_NextPage(uia, fp)
        if (nxt.done || !nxt.ok)
            break
        try {
            att := MobillsAuto_AttachBrowser(false)
            if (att.uia)
                uia := att.uia
        } catch {
        }
        Sleep MOBILLS_STEP_MS * 2
    }
    return { ok: true, rows: all, error: "" }
}

MobillsDaily_FirstAccountTitle(uia, catalog) {
    try {
        texts := uia.FindAll({ Type: 50020 })
        if texts {
            for t in texts {
                nm := ""
                try nm := t.Name
                cand := MobillsDaily_MatchAccountName(catalog, nm)
                if (cand != "")
                    return cand
            }
        }
    } catch {
    }
    return ""
}

MobillsDaily_ScrapeAccountsPage(uia, catalog) {
    rows := []
    texts := ""
    try texts := uia.FindAll({ Type: 50020 })
    catch {
        return rows
    }
    if !texts
        return rows
    names := []
    classes := []
    i := 1
    while (i <= texts.Length) {
        nm := ""
        cls := ""
        try nm := texts[i].Name
        try cls := texts[i].ClassName
        names.Push(nm)
        classes.Push(cls)
        i++
    }
    i := 1
    while (i <= names.Length) {
        if (names[i] = "Current balance") {
            money := ""
            moneyCls := ""
            if (i < names.Length && MobillsDaily_IsMoneyText(names[i + 1])) {
                money := names[i + 1]
                moneyCls := classes[i + 1]
            }
            if InStr(moneyCls, "MuiTypography-h5") || InStr(classes[i], "MuiTypography-h5") {
                i++
                continue
            }
            acct := ""
            j := i - 1
            while (j >= 1) {
                prev := names[j]
                if (prev = "Current balance" || prev = "Predicted balance" || prev = "ADD EXPENSE" || prev = "Balance")
                    break
                if MobillsDaily_IsMoneyText(prev) {
                    j--
                    continue
                }
                cand := MobillsDaily_MatchAccountName(catalog, prev)
                if (cand != "") {
                    acct := cand
                    break
                }
                j--
            }
            if (acct != "" && money != "")
                rows.Push({ kind: "account", name: acct, label: "Current balance", value: money })
        }
        i++
    }
    return rows
}

MobillsDaily_ScrapeCards(uia) {
    nav := MobillsAuto_ClickMenu(uia, "menu-creditCards-item", "Credit cards")
    if (!nav.ok)
        return { ok: false, rows: [], error: "Could not open Credit cards page", attempted: nav.attempted }
    Sleep MOBILLS_STEP_MS * 2
    StandardLoadingBar_Update("⏳ Scraping credit cards...")
    rows := []
    texts := ""
    try texts := uia.FindAll({ Type: 50020 })
    catch {
        return { ok: true, rows: rows, error: "" }
    }
    if !texts
        return { ok: true, rows: rows, error: "" }
    names := []
    for t in texts {
        try names.Push(t.Name)
        catch {
            names.Push("")
        }
    }
    card := ""
    sawPartialLabel := false
    sawClosing := false
    openVal := ""
    partialVal := ""
    i := 1
    while (i <= names.Length) {
        nm := names[i]
        if (nm = MOBILLS_CARD_NAME && card = "")
            card := nm
        if (nm = "Partial value")
            sawPartialLabel := true
        if (nm = "Closing on")
            sawClosing := true
        if MobillsDaily_IsMoneyText(nm) {
            if (sawPartialLabel && !sawClosing && partialVal = "")
                partialVal := nm
            else if (sawClosing && openVal = "")
                openVal := nm
        }
        if (nm = "Available Limit" || nm = "PAY INVOICE" || nm = "Total amount")
            break
        i++
    }
    if (card = "")
        card := MOBILLS_CARD_NAME
    if (openVal != "")
        rows.Push({ kind: "card", name: card, label: "Open invoice", value: openVal })
    if (partialVal != "")
        rows.Push({ kind: "card", name: card, label: "Partial value", value: partialVal })
    return { ok: true, rows: rows, error: "" }
}

MobillsDaily_WriteBalancesCsv(rows) {
    desktop := MobillsDaily_DesktopPath()
    name := "mobills-balances-" FormatTime(A_Now, "yyyy-MM-dd-HHmmss") ".csv"
    path := desktop "\" name
    body := "Kind,Name,Label,Value`n"
    for r in rows {
        body .= r.kind "," '"' StrReplace(r.name, '"', '""') '"' "," '"' StrReplace(r.label, '"', '""') '"' "," '"' StrReplace(
            r.value, '"', '""') '"`n'
    }
    try FileAppend(body, path, "UTF-8")
    catch as e {
        return { ok: false, path: "", error: e.Message }
    }
    return { ok: true, path: path, error: "" }
}

MobillsDaily_ShowReviewTable(rows, csvPath, skips := "") {
    global g_MobillsDailyReviewResult, g_MobillsDailyReviewGui
    if !IsObject(skips)
        skips := []
    g_MobillsDailyReviewResult := ""
    title := skips.Length ? "Mobills balances — " skips.Length " to fix manually" : "Mobills balances"
    dlg := Gui("+AlwaysOnTop +ToolWindow", title)
    dlg.SetFont("s10", "Segoe UI")
    hint := "Saved: " csvPath
    if skips.Length
        hint .= "`n" skips.Length " skipped — do these manually in Mobills (not saved)."
    dlg.Add("Text", "w900", hint)
    if skips.Length {
        lvSkip := dlg.Add("ListView", "w900 h180 -Multi", ["#", "Type", "Description", "Value", "Field", "Wanted",
            "Note"])
        for s in skips
            lvSkip.Add("", s.idx, s.type, s.description, s.value, s.field, s.wanted, s.note)
        try lvSkip.ModifyCol(1, 36)
        try lvSkip.ModifyCol(2, 90)
        try lvSkip.ModifyCol(3, 160)
        try lvSkip.ModifyCol(4, 70)
        try lvSkip.ModifyCol(5, 110)
        try lvSkip.ModifyCol(6, 140)
        try lvSkip.ModifyCol(7, 280)
    }
    lv := dlg.Add("ListView", "w900 h280 -Multi", ["Kind", "Name", "Label", "Value"])
    for r in rows
        lv.Add("", r.kind, r.name, r.label, r.value)
    try lv.ModifyCol(1, 80)
    try lv.ModifyCol(2, 220)
    try lv.ModifyCol(3, 160)
    try lv.ModifyCol(4, 140)
    dlg.Add("Button", "w100 Default", "Close").OnEvent("Click", MobillsDaily_ReviewClose)
    dlg.OnEvent("Close", MobillsDaily_ReviewClose)
    dlg.OnEvent("Escape", MobillsDaily_ReviewClose)
    g_MobillsDailyReviewGui := dlg
    dlg.Show()
    start := A_TickCount
    while (g_MobillsDailyReviewResult = "") {
        if ((A_TickCount - start) >= 300000)
            break
        Sleep 50
    }
    g_MobillsDailyReviewResult := ""
    try dlg.Destroy()
    catch {
    }
    g_MobillsDailyReviewGui := ""
}

MobillsDaily_ReviewClose(*) {
    global g_MobillsDailyReviewResult
    g_MobillsDailyReviewResult := "done"
}

MobillsDaily_Run() {
    path := MobillsDaily_FindLatestFile()
    if (path = "") {
        ShowCenteredOverlay_Utils("❌ No MOBILLS_V1.txt (or .ini) on Desktop", 2500, BANNER_ACCENT_ERROR)
        return
    }
    parsed := MobillsDaily_ParseFile(path)
    if (parsed.error != "" && !InStr(parsed.error, "empty")) {
        ShowCenteredOverlay_Utils("❌ " parsed.error, 2500, BANNER_ACCENT_ERROR)
        return
    }
    sourceLabel := path
    if (!parsed.rows.Length) {
        clip := ""
        try clip := A_Clipboard
        catch {
        }
        clipParsed := MobillsDaily_ParseRaw(clip)
        if (clipParsed.rows.Length) {
            parsed := clipParsed
            sourceLabel := path " + clipboard"
        }
    }
    if (!parsed.rows.Length) {
        sz := 0
        try sz := FileGetSize(path)
        catch {
        }
        hint := parsed.firstLine != "" ? parsed.firstLine : "(empty)"
        ShowCenteredOverlay_Utils("❌ Desktop .txt/.ini has no rows (" sz " bytes: " hint "). Save MOBILLS_V1.txt or copy the table, then run Macros [m] again.",
            4500, BANNER_ACCENT_ERROR)
        return
    }
    if !MobillsDaily_ShowConfirmTable(parsed.rows, parsed.unmapped, sourceLabel) {
        ShowCenteredOverlay_Utils("Mobills daily entry cancelled", 1500, BANNER_ACCENT_INTERMEDIATE)
        return
    }

    StandardLoadingBar_Show("⏳ Opening Mobills...", BANNER_ACCENT_INTERMEDIATE, { fontSize: 17, trackActiveMonitor: true })
    barOwned := true
    try {
        att := MobillsAuto_AttachBrowser()
        if (att.error != "") {
            StandardLoadingBar_Hide(0)
            barOwned := false
            MobillsDaily_Halt(att.error)
            return
        }
        uia := att.uia
        total := parsed.rows.Length
        global g_MobillsDailySkips
        g_MobillsDailySkips := []
        i := 1
        for row in parsed.rows {
            MobillsDaily_EnterRow(uia, row, i, total)
            i++
            try uia := MobillsAuto_AttachBrowser(false).uia
            catch {
            }
        }

        StandardLoadingBar_Update("⏳ Scraping accounts...")
        acc := MobillsDaily_ScrapeAccounts(uia)
        if (!acc.ok) {
            StandardLoadingBar_Hide(0)
            barOwned := false
            MobillsDaily_Halt("Accounts scrape failed: " acc.error)
            return
        }
        try uia := MobillsAuto_AttachBrowser(false).uia
        catch {
        }
        cards := MobillsDaily_ScrapeCards(uia)
        if (!cards.ok) {
            StandardLoadingBar_Hide(0)
            barOwned := false
            MobillsDaily_Halt("Credit cards scrape failed: " cards.error)
            return
        }
        scraped := []
        for r in acc.rows
            scraped.Push(r)
        for r in cards.rows
            scraped.Push(r)
        written := MobillsDaily_WriteBalancesCsv(scraped)
        StandardLoadingBar_Hide(0)
        barOwned := false
        if (!written.ok) {
            MobillsDaily_Halt("Could not write balances CSV: " written.error)
            return
        }
        if (g_MobillsDailySkips.Length)
            ShowCenteredOverlay_Utils("Mobills done — " g_MobillsDailySkips.Length " to enter manually", 2500,
                BANNER_ACCENT_INTERMEDIATE)
        else
            ShowCenteredOverlay_Utils("✅ Mobills daily entry done", 1500, BANNER_ACCENT_SUCCESS)
        MobillsDaily_ShowReviewTable(scraped, written.path, g_MobillsDailySkips)
    } catch as e {
        try StandardLoadingBar_Hide(0)
        barOwned := false
        MobillsDaily_Halt("Unexpected error: " e.Message)
    } finally {
        if (barOwned) {
            try StandardLoadingBar_Hide(0)
        }
    }
}
