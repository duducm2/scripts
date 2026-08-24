; =============================================================================
; Utils module: finance_import.ahk
; Import AI-generated FINANCE_DAILY / FINANCE_MONTHLY files from Desktop
; =============================================================================

Finance_DesktopNewest(pattern) {
    newest := ""
    newestTime := 0
    loop files A_Desktop . "\" . pattern, "F" {
        ts := Number(A_LoopFileTimeModified)
        if (ts > newestTime) {
            newestTime := ts
            newest := A_LoopFileFullPath
        }
    }
    return newest
}

; Skip Gemini preambles (e.g. "FILE: FINANCE_DAILY.csv") so the header row is first.
Finance_ReadAiImportCsv(path) {
    text := Finance_ReadUtf8(path)
    if (text = "")
        return []
    text := StrReplace(text, "`r`n", "`n")
    text := StrReplace(text, "`r", "`n")
    lines := StrSplit(text, "`n")
    start := 0
    for idx, line in lines {
        t := Trim(line)
        if (t = "")
            continue
        lower := StrLower(t)
        if (InStr(lower, "file:") = 1 || SubStr(t, 1, 1) = "#")
            continue
        if (InStr(lower, "description") && InStr(lower, "amount")) {
            start := idx
            break
        }
        if (InStr(lower, "entity_type") && InStr(lower, "entity_id")) {
            start := idx
            break
        }
    }
    if (!start)
        return Finance_ReadCsv(path)
    cleaned := ""
    loop lines.Length {
        if (A_Index < start)
            continue
        if (cleaned != "")
            cleaned .= "`n"
        cleaned .= lines[A_Index]
    }
    tmp := A_Temp . "\finance_ai_import_norm.csv"
    Finance_WriteUtf8(tmp, cleaned)
    rows := Finance_ReadCsv(tmp)
    try FileDelete(tmp)
    catch {
    }
    cleaned_rows := []
    for r in rows {
        d := r.Has("description") ? StrLower(Trim(r["description"])) : ""
        if (d = "description" || d = "amount" || d = "type")
            continue
        et := r.Has("entity_type") ? StrLower(Trim(r["entity_type"])) : ""
        if (et = "entity_type")
            continue
        cleaned_rows.Push(r)
    }
    return cleaned_rows
}

; Extract CSV body between ===FILE: name=== / ---FILE: name--- and matching END_FILE.
Finance_ExtractPackFileSection(text, fileName) {
    for style in ["===", "---"] {
        needle := style . "FILE: " . fileName . style
        endNeedle := style . "END_FILE" . style
        pos := InStr(text, needle, false)
        if (!pos)
            continue
        rest := SubStr(text, pos + StrLen(needle))
        endPos := InStr(rest, endNeedle, false)
        if (!endPos)
            continue
        return Trim(SubStr(rest, 1, endPos - 1), "`r`n `t")
    }
    return ""
}

Finance_StripMarkdownFences(body) {
    body := Trim(body, "`r`n `t")
    if (SubStr(body, 1, 3) = "```") {
    nl := InStr(body, "`n")
    body := nl ? SubStr(body, nl + 1) : ""
    }
        body := Trim(body, "`r`n `t")
    if (SubStr(body, -3) = "```")
    body := Trim(SubStr(body, 1, StrLen(body) - 3), "`r`n `t")
    return body
    }
        Finance_StripPackPreview(text) {
        for style in ["===", "---"] {
            needle := style . "PREVIEW" . style
            endNeedle := style . "END_PREVIEW" . style
            pos := InStr(text, needle, false)
            if (!pos)
                continue
            rest := SubStr(text, pos + StrLen(needle))
            endPos := InStr(rest, endNeedle, false)
            if (!endPos)
                continue
            text := SubStr(text, 1, pos - 1) . SubStr(rest, endPos + StrLen(endNeedle))
            break
        }
        return text
    }

    ; Keep from the first CSV header row (daily or monthly).
    Finance_ExtractCsvFromHeader(text) {
        text := StrReplace(text, "`r`n", "`n")
        text := StrReplace(text, "`r", "`n")
        lines := StrSplit(text, "`n")
        start := 0
        for idx, line in lines {
            t := Trim(line)
            if (t = "")
                continue
            lower := StrLower(t)
            if (InStr(lower, "```") = 1)
                continue
                if (InStr(lower, "file:") = 1 || SubStr(t, 1, 1) = "#")
                continue
                if (InStr(lower, "description") && InStr(lower, "amount")) {
                start := idx
                break
            }
            if (InStr(lower, "entity_type") && InStr(lower, "entity_id")) {
                start := idx
                break
            }
            }
            if (!start)
                return ""
            cleaned := ""
            loop lines.Length {
                if (A_Index < start)
                    continue
                line := lines[A_Index]
                if (Trim(line) = "```")
                break
                if (cleaned != "")
                cleaned .= "`n"
                cleaned .= line
                }
                    return Trim(cleaned, "`r`n `t")
            }

            ; Convert an AI pack (.txt with markers/fences/preview) to a temp .csv path; else return path.
            Finance_MaterializeAiCsv(path, expectedFileName) {
                text := Finance_ReadUtf8(path)
                if (text = "")
                    return path
                stem := expectedFileName
                stem := RegExReplace(stem, "i)\.(csv|txt|ini)$", "")
                body := Finance_ExtractPackFileSection(text, stem . ".csv")
                if (body = "")
                    body := Finance_ExtractPackFileSection(text, stem . ".txt")
                if (body = "")
                    body := Finance_ExtractCsvFromHeader(Finance_StripPackPreview(text))
                if (body = "")
                    return path
                body := Finance_StripMarkdownFences(body)
                if (Trim(body) = "")
                    return path
                tmp := A_Temp . "\finance_pack_" . StrReplace(expectedFileName, ".", "_") . ".csv"
                Finance_WriteUtf8(tmp, body)
                return tmp
            }

            ; Newest Desktop dump that looks like a daily finance CSV (Gemini code downloads).
            Finance_DesktopNewestDailyCodeDump() {
                newest := ""
                newestTime := 0
                loop files A_Desktop . "\gemini-code*.txt", "F" {
                    body := Finance_ReadUtf8(A_LoopFileFullPath)
                    if (body = "")
                        continue
                    lower := StrLower(body)
                    hasDaily := InStr(lower, "description,amount") || InStr(lower, "date,description,amount")
                    || InStr(lower, "file: finance_daily")
                    if (!hasDaily)
                        continue
                    ts := Number(A_LoopFileTimeModified)
                    if (ts > newestTime) {
                        newestTime := ts
                        newest := A_LoopFileFullPath
                    }
                }
                return newest
            }

            Finance_DesktopNewestMonthlyCodeDump() {
                newest := ""
                newestTime := 0
                loop files A_Desktop . "\gemini-code*.txt", "F" {
                    body := Finance_ReadUtf8(A_LoopFileFullPath)
                    if (body = "")
                        continue
                    lower := StrLower(body)
                    hasMonthly := InStr(lower, "entity_type,entity_id") || InStr(lower, "file: finance_monthly")
                    if (!hasMonthly)
                        continue
                    ts := Number(A_LoopFileTimeModified)
                    if (ts > newestTime) {
                        newestTime := ts
                        newest := A_LoopFileFullPath
                    }
                }
                return newest
            }

            Finance_ArchiveImported(path) {
                destDir := Finance_DataDir() . "\imported"
                if (!DirExist(destDir))
                    DirCreate(destDir)
                SplitPath(path, &name)
                dest := destDir . "\" . FormatTime(, "yyyyMMdd-HHmmss") . "_" . name
                try FileMove(path, dest, 1)
                catch {
                    try FileCopy(path, dest, 1)
                    catch {
                    }
                }
            }

            ; Editable import preview. parsed is an array of Maps (mutated in place).
            ; Returns true if user confirms import, false if cancelled.
            Finance_ImportConfirmEditable(title, parsed) {
                global g_FinanceGui
                owner := ""
                try {
                    if (IsObject(g_FinanceGui))
                        owner := " +Owner" . g_FinanceGui.Hwnd
                } catch {
                    owner := ""
                }
                accs := Finance_Load("accounts")
                cats := Finance_Load("categories")
                cards := Finance_Load("credit_cards")
                Finance_DialogsBegin()
                prevGui := g_FinanceGui
                g := Gui("+AlwaysOnTop +ToolWindow" . owner, title)
                g_FinanceGui := g
                g.SetFont("s10", "Segoe UI")
                hdr := g.Add("Text", "x12 y8 w880", "Import preview — " . parsed.Length .
                    " row(s).  [E] edit  [A] add  [Delete] remove")
                lv := g.Add("ListView", "x12 y32 w896 r14 Grid", ["Date", "Type", "Description", "Amount", "Category",
                    "Account"])
                ok := false

                RefreshLv() {
                    lv.Delete()
                    for p in parsed {
                        cat := Finance_CatName(cats, p["category_id"])
                        if (p["subcategory"] != "")
                            cat .= " / " . p["subcategory"]
                        acc := Finance_ImportAccountLabel(p, accs, cards)
                        lv.Add("", p["date"], Finance_TypeLabel(p["type"]), p["description"],
                        Finance_FormatBrl(Finance_ParseDecimal(p["amount"])), cat, acc)
                    }
                    lv.ModifyCol(1, 90)
                    lv.ModifyCol(2, 90)
                    lv.ModifyCol(3, 250)
                    lv.ModifyCol(4, 100)
                    lv.ModifyCol(5, 180)
                    lv.ModifyCol(6, 200)
                    hdr.Value := "Import preview — " . parsed.Length . " row(s).  [E] edit  [A] add  [Delete] remove"
                }

                EditSelected() {
                    row := lv.GetNext()
                    if (!row || row > parsed.Length)
                        return
                    if (Finance_ImportRowForm(g, parsed[row], cats, accs))
                        RefreshLv()
                }

                DeleteSelected() {
                    row := lv.GetNext()
                    if (!row || row > parsed.Length)
                        return
                    parsed.RemoveAt(row)
                    RefreshLv()
                }

                AddRow() {
                    newRow := Map(
                        "date", Finance_Yesterday(),
                        "description", "",
                        "amount", "0,00",
                        "type", "expense",
                        "category_id", "",
                        "subcategory", "",
                        "account_id", "",
                        "card_id", "",
                        "transfer_account_id", ""
                    )
                    if (Finance_ImportRowForm(g, newRow, cats, accs)) {
                        parsed.Push(newRow)
                        RefreshLv()
                    }
                }

                RefreshLv()
                lv.OnEvent("DoubleClick", (*) => EditSelected())

                btnY := "y+10"
                g.Add("Button", btnY . " x12 w80", "Edit").OnEvent("Click", (*) => EditSelected())
                g.Add("Button", "x+6 w80", "Delete").OnEvent("Click", (*) => DeleteSelected())
                g.Add("Button", "x+6 w80", "Add").OnEvent("Click", (*) => AddRow())
                g.Add("Button", "x+30 w100 Default", "Import").OnEvent("Click", ConfirmYes)
                g.Add("Button", "x+6 w100", "Cancel").OnEvent("Click", (*) => g.Destroy())

                CleanupImportDialog(*) {
                    Finance_UnbindHotkeys()
                    g_FinanceGui := prevGui
                    g.Destroy()
                }

                Finance_BindHotkeys([
                    ["e", (*) => EditSelected()],
                    ["a", (*) => AddRow()],
                    ["Insert", (*) => AddRow()],
                    ["Delete", (*) => DeleteSelected()],
                    ["Escape", (*) => CleanupImportDialog()]
                ])

                g.OnEvent("Escape", (*) => CleanupImportDialog())
                g.OnEvent("Close", (*) => CleanupImportDialog())
                g.Show("w920")
                try WinWaitClose("ahk_id " g.Hwnd)
                catch {
                }
                Finance_UnbindHotkeys()
                g_FinanceGui := prevGui
                Finance_DialogsEnd()
                return ok

                ConfirmYes(*) {
                    ok := true
                    CleanupImportDialog()
                }
            }

            ; Read-only confirm for monthly import. rows: Maps with type, name, amount, description.
            Finance_ImportConfirm(title, rows) {
                global g_FinanceGui
                owner := ""
                try {
                    if (IsObject(g_FinanceGui))
                        owner := " +Owner" . g_FinanceGui.Hwnd
                } catch {
                    owner := ""
                }
                Finance_DialogsBegin()
                g := Gui("+AlwaysOnTop +ToolWindow" . owner, title)
                g.SetFont("s10", "Segoe UI")
                g.Add("Text", "w720", "Confirm import of " . rows.Length . " row(s).")
                lv := g.Add("ListView", "w720 r14 Grid", ["Type", "Name", "Amount", "Description"])
                for r in rows
                    lv.Add("", r["type"], r["name"], r["amount"], r["description"])
                lv.ModifyCol(1, 90)
                lv.ModifyCol(2, 180)
                lv.ModifyCol(3, 120)
                lv.ModifyCol(4, 300)
                ok := false
                g.Add("Button", "y+10 w120 Default", "Import").OnEvent("Click", ConfirmYes)
                g.Add("Button", "x+8 w120", "Cancel").OnEvent("Click", (*) => g.Destroy())
                g.OnEvent("Escape", (*) => g.Destroy())
                g.Show()
                try WinWaitClose("ahk_id " g.Hwnd)
                catch {
                }
                Finance_DialogsEnd()
                return ok

                ConfirmYes(*) {
                    ok := true
                    g.Destroy()
                }
            }

            ; Edit a single import row Map in-place. Returns true if saved.
            Finance_ImportRowForm(ownerGui, row, cats, accs) {
                cards := Finance_Load("credit_cards")
                owner := " +Owner" . ownerGui.Hwnd
                Finance_DialogsBegin()
                g := Gui("+AlwaysOnTop +ToolWindow" . owner, "Edit import row")
                g.SetFont("s10", "Segoe UI")

                g.Add("Text", "x10 y10", "Description")
                eDesc := g.Add("Edit", "x10 y28 w420", row["description"])
                g.Add("Text", "x10 y58", "Amount (comma decimal)")
                eAmt := g.Add("Edit", "x10 y76 w200", row["amount"])

                types := ["expense", "income", "transfer", "card_expense", "adjustment"]
                curType := row["type"]
                typeIdx := 1
                loop types.Length {
                    if (types[A_Index] = curType)
                        typeIdx := A_Index
                }
                g.Add("Text", "x230 y58", "Type")
                ddType := g.Add("DropDownList", "x230 y76 w200 Choose" . typeIdx,
                    ["Expense", "Income", "Transfer", "Credit card", "Adjustment"])

                y1 := 112
                y1c := 130
                y2 := 168
                y2c := 186

                catCombo := Finance_ComboFromRows(Finance_MainCategories(cats), "id", "name", true, "icon")
                catIdx := Finance_ComboIndex(catCombo.ids, row["category_id"])
                lblCat := g.Add("Text", "x10 y" . y1, "Category")
                ddCat := g.Add("DropDownList", "x10 y" . y1c . " w220 Choose" . catIdx, catCombo.names)
                lblSub := g.Add("Text", "x242 y" . y1, "Subcategory")
                eSub := g.Add("Edit", "x242 y" . y1c . " w180", row["subcategory"])

                accCombo := Finance_ComboFromRows(accs)
                accIdx := Finance_ComboIndex(accCombo.ids, row["account_id"] != "" ? row["account_id"]
                    : Finance_Setting("General", "DefaultAccountId", ""))
                lblAcc := g.Add("Text", "x10 y" . y2, "Account")
                ddAcc := g.Add("DropDownList", "x10 y" . y2c . " w220 Choose" . accIdx, accCombo.names)

                destCombo := Finance_ComboFromRows(accs, "id", "name", true)
                destIdx := Finance_ComboIndex(destCombo.ids, row["transfer_account_id"])
                lblDest := g.Add("Text", "x242 y" . y2, "To account")
                ddDest := g.Add("DropDownList", "x242 y" . y2c . " w180 Choose" . destIdx, destCombo.names)

                primaryCardId := Finance_Setting("General", "PrimaryCardId", "")
                cardRows := []
                if (primaryCardId != "") {
                    prow := Finance_FindById(cards, primaryCardId)
                    if (prow)
                        cardRows.Push(prow)
                }
                for c in cards {
                    if (c["id"] != primaryCardId)
                        cardRows.Push(c)
                }
                cardCombo := Finance_ComboFromRows(cardRows, "id", "name", false)
                defaultCardId := row["card_id"] != "" ? row["card_id"] : primaryCardId
                if (defaultCardId = "" && cardCombo.ids.Length)
                    defaultCardId := cardCombo.ids[1]
                cardIdx := Finance_ComboIndex(cardCombo.ids, defaultCardId)
                lblCard := g.Add("Text", "x10 y" . y2, "Credit card")
                ddCard := g.Add("DropDownList", "x10 y" . y2c . " w220 Choose" . cardIdx, cardCombo.names)

                saved := false
                g.Add("Button", "x10 y230 w100 Default", "Save").OnEvent("Click", SaveRow)
                g.Add("Button", "x118 y230 w100", "Cancel").OnEvent("Click", (*) => g.Destroy())
                g.OnEvent("Escape", (*) => g.Destroy())
                ddType.OnEvent("Change", (*) => ApplyFields(types[ddType.Value]))
                ApplyFields(curType)
                g.Show("w460 h280")
                try WinWaitClose("ahk_id " g.Hwnd)
                catch {
                }
                Finance_DialogsEnd()
                return saved

                ApplyFields(t) {
                    showCat := (t != "transfer")
                    showAcc := (t != "card_expense")
                    showDest := (t = "transfer")
                    showCard := (t = "card_expense")
                    lblCat.Visible := showCat
                    ddCat.Visible := showCat
                    lblSub.Visible := showCat
                    eSub.Visible := showCat
                    lblAcc.Visible := showAcc
                    ddAcc.Visible := showAcc
                    lblDest.Visible := showDest
                    ddDest.Visible := showDest
                    lblCard.Visible := showCard
                    ddCard.Visible := showCard
                    lblAcc.Text := (t = "transfer") ? "From account" : "Account"
                    if (t = "transfer") {
                        lblAcc.Move(10, y1)
                        ddAcc.Move(10, y1c)
                        lblDest.Move(242, y2)
                        ddDest.Move(242, y2c)
                    } else {
                        lblAcc.Move(10, y2)
                        ddAcc.Move(10, y2c)
                    }
                    catFilter := ""
                    if (t = "expense" || t = "card_expense")
                        catFilter := "expense"
                    else if (t = "income")
                        catFilter := "income"
                    keepId := ""
                    try keepId := catCombo.ids[ddCat.Value]
                    catch {
                        keepId := ""
                    }
                    catCombo := Finance_ComboFromRows(Finance_MainCategories(cats, catFilter), "id", "name", true,
                    "icon")
                    ddCat.Delete()
                    ddCat.Add(catCombo.names)
                    ddCat.Choose(Finance_ComboIndex(catCombo.ids, keepId))
                }

                SaveRow(*) {
                    desc := Trim(eDesc.Value)
                    if (desc = "") {
                        Finance_Alert("Description is required.", "Import")
                        return
                    }
                    rawAmt := Finance_NormalizeDot(eAmt.Value)
                    amt := Finance_ParseDecimal(rawAmt)
                    if (amt < 0)
                        amt := -amt
                    t := types[ddType.Value]
                    catId := ""
                    subVal := ""
                    accId := ""
                    destId := ""
                    cardId := ""
                    if (t != "transfer") {
                        catId := catCombo.ids[ddCat.Value]
                        subVal := Trim(eSub.Value)
                    }
                    if (t != "card_expense")
                        accId := accCombo.ids[ddAcc.Value]
                    if (t = "transfer")
                        destId := destCombo.ids[ddDest.Value]
                    if (t = "card_expense") {
                        cardId := cardCombo.ids[ddCard.Value]
                        if (cardId = "")
                            cardId := Finance_Setting("General", "PrimaryCardId", "CARD_MP")
                    }
                    row["description"] := desc
                    row["amount"] := Finance_FormatCsvDecimal(amt)
                    row["type"] := t
                    row["category_id"] := catId
                    row["subcategory"] := subVal
                    row["account_id"] := accId
                    row["card_id"] := cardId
                    row["transfer_account_id"] := destId
                    saved := true
                    g.Destroy()
                }
            }

            Finance_ImportDaily(*) {
                Finance_ImportDailyFromPath("", false)
            }

            ; path empty = discover on Desktop. autoConfirm skips the confirm dialog.
            Finance_ImportDailyFromPath(path := "", autoConfirm := false) {
                if (path = "") {
                    path := Finance_DesktopNewest("FINANCE_DAILY*.txt")
                    if (path = "")
                        path := Finance_DesktopNewest("FINANCE_DAILY*.csv")
                    if (path = "")
                        path := Finance_DesktopNewest("FINANCE_DAILY*.ini")
                    if (path = "")
                        path := Finance_DesktopNewestDailyCodeDump()
                }
                if (path = "" || !FileExist(path)) {
                    Finance_Notify("No FINANCE_DAILY file on Desktop", 2000, BANNER_ACCENT_ERROR)
                    return false
                }
                sourcePath := path
                csvPath := Finance_MaterializeAiCsv(path, "FINANCE_DAILY.csv")
                rows := Finance_ReadAiImportCsv(csvPath)
                if (csvPath != sourcePath) {
                    try FileDelete(csvPath)
                    catch {
                    }
                }
                if (!rows.Length) {
                    Finance_Notify("File has no data rows", 1800, BANNER_ACCENT_ERROR)
                    return false
                }
                cats := Finance_Load("categories")
                parsed := []
                for r in rows {
                    desc := r.Has("description") ? r["description"] : ""
                    amt := r.Has("amount") ? r["amount"] : "0"
                    t := r.Has("type") ? r["type"] : "expense"
                    date := Finance_Yesterday()
                    resolved := Finance_ResolveImportCategory(cats, t,
                        r.Has("category_id") ? r["category_id"] : "",
                        r.Has("subcategory") ? r["subcategory"] : "")
                    parsed.Push(Map(
                        "date", date,
                        "description", desc,
                        "amount", Finance_FormatCsvDecimal(Abs(Finance_ParseDecimal(amt))),
                        "type", t,
                        "category_id", resolved["category_id"],
                        "subcategory", resolved["subcategory"],
                        "account_id", r.Has("account_id") ? r["account_id"] : "",
                        "card_id", r.Has("card_id") ? r["card_id"] : "",
                        "transfer_account_id", r.Has("transfer_account_id") ? r["transfer_account_id"] : ""
                    ))
                }
                if (!autoConfirm && !Finance_ImportConfirmEditable("Import daily transactions", parsed))
                    return false
                if (!parsed.Length) {
                    Finance_Notify("No rows to import", 1800, BANNER_ACCENT_ERROR)
                    return false
                }
                txs := Finance_Load("transactions")
                for p in parsed {
                    p["id"] := Finance_NextId("TX", txs)
                    txs.Push(p)
                    Finance_ApplyTransactionToBalances(p, false)
                }
                Finance_Save("transactions", txs)
                Finance_AfterDailyImport(parsed, autoConfirm)
                Finance_ArchiveImported(sourcePath)
                Finance_Notify("Imported " . parsed.Length . " transactions", 1800, BANNER_ACCENT_SUCCESS)
                return true
            }

            Finance_AfterDailyImport(parsed, autoConfirm) {
                global g_FinanceMonth, g_FinanceTxFilter, g_FinanceGui
                months := Map()
                latest := ""
                cur := Finance_CurrentYearMonth()
                hasCurrent := false
                for p in parsed {
                    ym := SubStr(p["date"], 1, 7)
                    if (StrLen(ym) < 7)
                        continue
                    months[ym] := true
                    if (ym = cur)
                        hasCurrent := true
                    if (latest = "" || StrCompare(p["date"], latest) > 0)
                        latest := p["date"]
                }
                nMonths := 0
                onlyYm := ""
                for ym, _ in months {
                    nMonths += 1
                    onlyYm := ym
                    Finance_RecomputeBudgetSpent(ym)
                }
                if (!nMonths)
                    Finance_RecomputeBudgetSpent(cur)
                if (nMonths = 1)
                    g_FinanceMonth := onlyYm
                else if (hasCurrent)
                    g_FinanceMonth := cur
                else if (latest != "")
                    g_FinanceMonth := SubStr(latest, 1, 7)
                else
                    g_FinanceMonth := cur
                g_FinanceTxFilter := "all"
                guiOpen := false
                try guiOpen := IsObject(g_FinanceGui)
                catch {
                    guiOpen := false
                }
                if (!autoConfirm || guiOpen)
                    Finance_ShowTransactions()
            }

            Finance_ImportMonthly(*) {
                path := Finance_DesktopNewest("FINANCE_MONTHLY*.txt")
                if (path = "")
                    path := Finance_DesktopNewest("FINANCE_MONTHLY*.csv")
                if (path = "")
                    path := Finance_DesktopNewest("FINANCE_MONTHLY*.ini")
                if (path = "")
                    path := Finance_DesktopNewestMonthlyCodeDump()
                if (path = "") {
                    Finance_Notify("No FINANCE_MONTHLY file on Desktop", 2000, BANNER_ACCENT_ERROR)
                    return
                }
                sourcePath := path
                csvPath := Finance_MaterializeAiCsv(path, "FINANCE_MONTHLY.csv")
                rows := Finance_ReadAiImportCsv(csvPath)
                if (csvPath != sourcePath) {
                    try FileDelete(csvPath)
                    catch {
                    }
                }
                if (!rows.Length) {
                    Finance_Notify("File has no data rows", 1800, BANNER_ACCENT_ERROR)
                    return
                }
                accs := Finance_Load("accounts")
                goals := Finance_Load("goals")
                previewRows := []
                for r in rows {
                    et := r.Has("entity_type") ? StrLower(Trim(r["entity_type"])) : ""
                    eid := r.Has("entity_id") ? Trim(r["entity_id"]) : ""
                    adj := Finance_ParseDecimal(r.Has("adjustment_amount") ? r["adjustment_amount"] : "0")
                    desc := r.Has("description") ? r["description"] : ""
                    typeLabel := et = "goal" ? "Goal" : (et = "account" ? "Account" : et)
                    name := et = "goal" ? Finance_NameOrUnknown(goals, eid) : Finance_NameOrUnknown(accs, eid)
                    previewRows.Push(Map(
                        "type", typeLabel,
                        "name", name,
                        "amount", Finance_FormatBrl(adj),
                        "description", desc
                    ))
                }
                if (!Finance_ImportConfirm("Import monthly adjustments", previewRows))
                    return
                txs := Finance_Load("transactions")
                n := 0
                for r in rows {
                    et := r.Has("entity_type") ? StrLower(Trim(r["entity_type"])) : ""
                    eid := r.Has("entity_id") ? Trim(r["entity_id"]) : ""
                    adj := Finance_ParseDecimal(r.Has("adjustment_amount") ? r["adjustment_amount"] : "0")
                    desc := r.Has("description") ? r["description"] : "Monthly adjustment"
                    if (et = "account") {
                        acc := Finance_FindById(accs, eid)
                        if (!acc)
                            continue
                        Finance_AdjustAccount(accs, eid, adj)
                        tx := Map(
                            "id", Finance_NextId("TX", txs),
                            "date", Finance_Today(),
                            "description", desc,
                            "amount", Finance_FormatCsvDecimal(Abs(adj)),
                            "type", adj >= 0 ? "income" : "expense",
                            "category_id", adj >= 0 ? Finance_CatIdByName("Investments") : Finance_CatIdByName(
                                "Adjustment"),
                        "subcategory", "",
                        "account_id", eid,
                        "card_id", "",
                        "transfer_account_id", ""
                        )
                        txs.Push(tx)
                        n += 1
                    } else if (et = "goal") {
                        g := Finance_FindById(goals, eid)
                        if (!g)
                            continue
                        cur := Finance_ParseDecimal(g["current_amount"]) + adj
                        g["current_amount"] := Finance_FormatCsvDecimal(cur)
                        n += 1
                    }
                }
                Finance_Save("accounts", accs)
                Finance_Save("goals", goals)
                Finance_Save("transactions", txs)
                Finance_RecomputeBudgetSpent(Finance_CurrentYearMonth())
                Finance_ArchiveImported(sourcePath)
                Finance_Notify("Applied " . n . " monthly adjustments", 1800, BANNER_ACCENT_SUCCESS)
                Finance_ShowMainMenu()
            }
