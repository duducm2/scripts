; =============================================================================
; Shift keys module: hotif_excel_mspaint.ahk
; Excel and Paint hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

#HotIf WinActive("ahk_exe mspaint.exe")

; Shift + Y : Resize and Skew (Ctrl+W)
+y:: Send "^w"

#HotIf

;-------------------------------------------------------------------
; Excel Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe EXCEL.EXE") && SafeWinGetClass() != "#32770"

; Helper function: Convert CSV delimited by semicolon into columns
; autoSelectSemicolon: If true, automatically selects semicolon without showing dialog. If false, shows confirmation dialog.
Excel_CSVToColumns(autoSelectSemicolon := false) {
    Send "{Alt}"
    Sleep 100
    Send "0"
    Sleep 100
    Send "5"
    Sleep 100
    Send "d"
    Sleep 100
    Send "{Enter}"
    Sleep 100
    if (autoSelectSemicolon) {
        ; Automatically select semicolon without dialog
        Send "m"
        Sleep 100
    } else {
        ; Show confirmation dialog for user to decide
        if MsgBox("If 'semicolon' is not selected, hit yes", "Confirm", "YesNo Icon?") = "Yes" {
            Send "m"
            Sleep 100
        }
    }
    Send "{Enter}"
    Sleep 100
    Send "{Enter}"
}

; Shift + W : Select White Color (Up-Arrow, Ctrl-Home, Ctrl-Home)
+w:: {
    Send "^{PgUp}"
}

; Shift + E : Click Enable Editing button
+e:: {
    try {
        root := UIA.ElementFromHandle(WinExist("A"))
        if (btn := WaitForButton(root, "Enable Editing", 3000)) {
            btn.Invoke()
        } else {
            MsgBox "Couldn't find the Enable Editing button."
        }
    } catch Error as err {
        MsgBox "Error:`n" err.Message
    }
}

; Shift + I : From Text/CSV ← clipboard path → shared import pipeline → Save CSV UTF-8
; (includes former Shift+U save step — U left free for a future Excel shortcut)
+i:: {
    Excel_ImportCsvFromClipboardPath()
}

; Shift + C : Turn CSV delimited by semicolon into columns (Alt, 0, 5, D, Enter, M, Enter, Enter)
+c:: {
    Excel_CSVToColumns()
}

; Convert 1-based column index to A1 letter(s) (1 -> A, 27 -> AA).
Excel_ColLetter(col) {
    s := ""
    while (col > 0) {
        col -= 1
        s := Chr(65 + Mod(col, 26)) . s
        col := col // 26
    }
    return s
}

; Read a single cell via A1 address (avoids AHK COM multi-arg Cells binding issues).
Excel_CellText(ws, row, col) {
    try return Trim(String(ws.Range(Excel_ColLetter(col) . row).Value2))
    catch {
        return ""
    }
}

; True when name is Excel's default Column1 / Coluna1 style header.
Excel_IsGenericColumnHeader(name) {
    try name := Trim(String(name))
    catch {
        return false
    }
    if (name = "")
        return false
    return RegExMatch(name, "i)^(Column|Coluna)\d+$")
}

; First header generic AND at least half of non-empty headers generic.
Excel_HeadersLookGeneric(headers) {
    if (!IsObject(headers) || headers.Length = 0)
        return false
    if (!Excel_IsGenericColumnHeader(headers[1]))
        return false
    nonEmpty := 0
    generic := 0
    for h in headers {
        try hStr := Trim(String(h))
        catch {
            continue
        }
        if (hStr = "")
            continue
        nonEmpty++
        if (Excel_IsGenericColumnHeader(hStr))
            generic++
    }
    if (nonEmpty = 0)
        return false
    return generic * 2 >= nonEmpty
}

; If table/sheet headers are Column1/Column2/…, promote real header row into the table.
; After CSV Load, a junk ColumnN data row may sit above the real names — scan past it.
; Silent no-op when headers are not generic. MsgBox only on unexpected COM errors mid-promote.
Excel_PromoteGenericHeaders() {
    try {
        xl := ComObjActive("Excel.Application")
        ws := xl.ActiveSheet
    } catch {
        return
    }

    try {
        loCount := 0
        try loCount := ws.ListObjects.Count
        catch {
        }
        if (loCount >= 1) {
            lo := ws.ListObjects(1)
            colCount := lo.ListColumns.Count
            headers := []
            loop colCount
                headers.Push(lo.ListColumns(A_Index).Name)
            if (!Excel_HeadersLookGeneric(headers))
                return
            body := lo.DataBodyRange
            bodyRows := 0
            try bodyRows := body.Rows.Count
            catch {
            }
            if (!body || bodyRows < 1)
                return
            hdrRow := lo.HeaderRowRange.Row
            startCol := lo.Range.Column
            ; First body row may be junk Column1 values; real headers can be on a later row
            headerSourceRow := 0
            loop bodyRows {
                r := hdrRow + A_Index
                cell1 := Excel_CellText(ws, r, startCol)
                if (cell1 != "" && !Excel_IsGenericColumnHeader(cell1)) {
                    headerSourceRow := r
                    break
                }
            }
            if (headerSourceRow = 0)
                return
            applied := 0
            loop colCount {
                newName := Excel_CellText(ws, headerSourceRow, startCol + A_Index - 1)
                if (newName != "" && !Excel_IsGenericColumnHeader(newName)) {
                    lo.ListColumns(A_Index).Name := newName
                    applied++
                }
            }
            ; Delete body rows from first through header-source (junk + header-as-data)
            listRowsToDelete := headerSourceRow - hdrRow
            if (applied > 0 && listRowsToDelete > 0) {
                loop listRowsToDelete {
                    try lo.ListRows(1).Delete()
                    catch {
                        break
                    }
                }
            }
            return
        }

        ur := ws.UsedRange
        if (!ur || ur.Rows.Count < 2)
            return
        colCount := ur.Columns.Count
        startCol := ur.Column
        startRow := ur.Row
        rowCount := ur.Rows.Count
        headers := []
        loop colCount
            headers.Push(Excel_CellText(ws, startRow, startCol + A_Index - 1))
        if (!Excel_HeadersLookGeneric(headers))
            return
        headerSourceRow := 0
        loop rowCount - 1 {
            r := startRow + A_Index
            cell1 := Excel_CellText(ws, r, startCol)
            if (cell1 != "" && !Excel_IsGenericColumnHeader(cell1)) {
                headerSourceRow := r
                break
            }
        }
        if (headerSourceRow = 0)
            return
        applied := 0
        loop colCount {
            c := startCol + A_Index - 1
            newName := Excel_CellText(ws, headerSourceRow, c)
            if (newName != "" && !Excel_IsGenericColumnHeader(newName)) {
                ws.Range(Excel_ColLetter(c) . startRow).Value2 := newName
                applied++
            }
        }
        if (applied > 0) {
            ; Delete from row under header through header-source row (bottom-up)
            r := headerSourceRow
            while (r > startRow) {
                ws.Rows(r).Delete()
                r--
            }
        }
    } catch Error as err {
        MsgBox("Could not promote generic headers:`n" err.Message)
    }
}

; Shift + N : Cycle layout pillars
; Balanced · Scan · Narrative · Reference · Titles · Triage · Immersive
global Excel_LayoutPillarIndex := -1
global Excel_LayoutPillarCount := 7

Excel_ReadableLayout_CellStrLen(v) {
    if (v = "")
        return 0
    try return StrLen(Trim(String(v)))
    catch {
        return 0
    }
}

Excel_ReadableLayout_ResolveRange(xl, ws, cell) {
    try {
        loCount := 0
        try loCount := ws.ListObjects.Count
        catch {
        }
        loop loCount {
            lo := ws.ListObjects(A_Index)
            try {
                if xl.Intersect(cell, lo.Range)
                    return lo.Range
            } catch {
            }
        }
    } catch {
    }
    try return ws.UsedRange
    catch {
        return 0
    }
}

Excel_Layout_PillarName(pillar) {
    switch pillar {
        case 0: return "Balanced"
        case 1: return "Scan"
        case 2: return "Narrative"
        case 3: return "Reference"
        case 4: return "Titles"
        case 5: return "Triage"
        case 6: return "Immersive"
        default: return "Balanced"
    }
}

; Returns array of Maps: header, headerLen, maxLen, isUrlish, kind ("short"|"wrap"|"single")
Excel_Layout_Measure(ws, tableRange) {
    static SHORT_MAX_LEN := 18
    static WRAP_MAX_LEN := 160

    startRow := tableRange.Row
    startCol := tableRange.Column
    rowCount := tableRange.Rows.Count
    colCount := tableRange.Columns.Count
    endRow := startRow + rowCount - 1
    metrics := []

    loop colCount {
        c := startCol + A_Index - 1
        headerText := ""
        headerLen := 0
        maxLen := 0
        try {
            headerText := Trim(String(ws.Cells(startRow, c).Text))
            headerLen := StrLen(headerText)
            maxLen := headerLen
        } catch {
        }
        r := startRow
        while (r <= endRow) {
            try {
                len := Excel_ReadableLayout_CellStrLen(ws.Cells(r, c).Text)
                if (len > maxLen)
                    maxLen := len
            } catch {
            }
            r++
        }

        headerLower := StrLower(headerText)
        isUrlish := InStr(headerLower, "url") || InStr(headerLower, "link")
        if !isUrlish {
            sampleR := startRow + (rowCount > 1 ? 1 : 0)
            try {
                sample := Trim(String(ws.Cells(sampleR, c).Text))
                if (SubStr(sample, 1, 7) = "http://" || SubStr(sample, 1, 8) = "https://")
                    isUrlish := true
            } catch {
            }
        }

        kind := "wrap"
        if (isUrlish || maxLen > WRAP_MAX_LEN)
            kind := "single"
        else if (maxLen <= SHORT_MAX_LEN)
            kind := "short"

        metrics.Push(Map(
            "header", headerText,
            "headerLen", headerLen,
            "maxLen", maxLen,
            "isUrlish", isUrlish,
            "kind", kind
        ))
    }
    return metrics
}

Excel_Layout_BasePrefer(m) {
    static SHORT_FLOOR := 4
    static SHORT_CAP := 14
    static WRAP_TARGET_MIN := 28
    static WRAP_TARGET_MAX := 45
    static SINGLE_WIDTH := 28

    kind := m["kind"]
    if (kind = "single")
        return SINGLE_WIDTH
    if (kind = "short") {
        snug := Max(m["maxLen"] + 2, m["headerLen"] + 1)
        return Min(SHORT_CAP, Max(SHORT_FLOOR, snug))
    }
    return Min(WRAP_TARGET_MAX, Max(WRAP_TARGET_MIN, Round(m["maxLen"] / 4)))
}

; Returns Map(prefer, doWrap, wrapCount, shortCount, singleCount, focusHeader, maxRowHeight)
Excel_Layout_PrefsForPillar(metrics, pillar) {
    static SCAN_ROW_HEIGHT := 22
    static DEFAULT_ROW_HEIGHT := 72

    colCount := metrics.Length
    prefer := []
    doWrap := []
    wrapCount := 0
    shortCount := 0
    singleCount := 0
    focusHeader := ""
    maxRowHeight := DEFAULT_ROW_HEIGHT

    ; Balanced base prefs
    loop colCount {
        m := metrics[A_Index]
        base := Excel_Layout_BasePrefer(m)
        prefer.Push(base)
        wrap := (m["kind"] = "wrap")
        doWrap.Push(wrap)
        if (m["kind"] = "single")
            singleCount++
        else if (m["kind"] = "short")
            shortCount++
        else
            wrapCount++
    }

    if (pillar = 0) {
        ; Balanced — already set
        return Map("prefer", prefer, "doWrap", doWrap, "wrapCount", wrapCount,
            "shortCount", shortCount, "singleCount", singleCount,
            "focusHeader", "", "maxRowHeight", maxRowHeight)
    }

    if (pillar = 1) {
        ; Scan — all single-line, snug, compact rows
        prefer := []
        doWrap := []
        wrapCount := 0
        shortCount := 0
        singleCount := 0
        loop colCount {
            m := metrics[A_Index]
            snug := Max(m["maxLen"] + 1, m["headerLen"] + 1)
            prefer.Push(Min(16, Max(3.5, snug)))
            doWrap.Push(false)
            shortCount++
        }
        return Map("prefer", prefer, "doWrap", doWrap, "wrapCount", 0,
            "shortCount", shortCount, "singleCount", 0,
            "focusHeader", "", "maxRowHeight", SCAN_ROW_HEIGHT)
    }

    if (pillar = 2) {
        ; Narrative — focus longest wrap-eligible column
        focusIdx := 0
        focusLen := -1
        loop colCount {
            m := metrics[A_Index]
            if (m["kind"] = "wrap" && m["maxLen"] > focusLen) {
                focusLen := m["maxLen"]
                focusIdx := A_Index
            }
        }
        if (focusIdx = 0) {
            return Map("prefer", prefer, "doWrap", doWrap, "wrapCount", wrapCount,
                "shortCount", shortCount, "singleCount", singleCount,
                "focusHeader", "", "maxRowHeight", maxRowHeight)
        }
        prefer2 := []
        doWrap2 := []
        wrapCount := 0
        shortCount := 0
        singleCount := 0
        loop colCount {
            m := metrics[A_Index]
            base := Excel_Layout_BasePrefer(m)
            if (A_Index = focusIdx) {
                prefer2.Push(base * 3.5)
                doWrap2.Push(true)
                wrapCount++
            } else {
                prefer2.Push(base * 0.55)
                doWrap2.Push(false)
                if (m["kind"] = "single")
                    singleCount++
                else
                    shortCount++
            }
        }
        focusHeader := metrics[focusIdx]["header"]
        return Map("prefer", prefer2, "doWrap", doWrap2, "wrapCount", wrapCount,
            "shortCount", shortCount, "singleCount", singleCount,
            "focusHeader", focusHeader, "maxRowHeight", maxRowHeight)
    }

    if (pillar = 3) {
        ; Reference — focus largest URL/extreme column
        focusIdx := 0
        focusLen := -1
        loop colCount {
            m := metrics[A_Index]
            if (m["kind"] = "single" && m["maxLen"] > focusLen) {
                focusLen := m["maxLen"]
                focusIdx := A_Index
            }
        }
        if (focusIdx = 0) {
            return Map("prefer", prefer, "doWrap", doWrap, "wrapCount", wrapCount,
                "shortCount", shortCount, "singleCount", singleCount,
                "focusHeader", "", "maxRowHeight", maxRowHeight)
        }
        prefer3 := []
        doWrap3 := []
        wrapCount := 0
        shortCount := 0
        singleCount := 0
        loop colCount {
            m := metrics[A_Index]
            base := Excel_Layout_BasePrefer(m)
            if (A_Index = focusIdx) {
                prefer3.Push(base * 2.5)
                doWrap3.Push(false)
                singleCount++
            } else if (m["kind"] = "wrap") {
                prefer3.Push(base * 0.7)
                doWrap3.Push(true)
                wrapCount++
            } else {
                prefer3.Push(base)
                doWrap3.Push(false)
                if (m["kind"] = "single")
                    singleCount++
                else
                    shortCount++
            }
        }
        focusHeader := metrics[focusIdx]["header"]
        return Map("prefer", prefer3, "doWrap", doWrap3, "wrapCount", wrapCount,
            "shortCount", shortCount, "singleCount", singleCount,
            "focusHeader", focusHeader, "maxRowHeight", maxRowHeight)
    }

    if (pillar = 4) {
        ; Titles — focus shortest wrap-eligible column (role/title skimming)
        focusIdx := 0
        focusLen := 0x7FFFFFFF
        loop colCount {
            m := metrics[A_Index]
            if (m["kind"] = "wrap" && m["maxLen"] < focusLen) {
                focusLen := m["maxLen"]
                focusIdx := A_Index
            }
        }
        if (focusIdx = 0) {
            return Map("prefer", prefer, "doWrap", doWrap, "wrapCount", wrapCount,
                "shortCount", shortCount, "singleCount", singleCount,
                "focusHeader", "", "maxRowHeight", maxRowHeight)
        }
        prefer4 := []
        doWrap4 := []
        wrapCount := 0
        shortCount := 0
        singleCount := 0
        loop colCount {
            m := metrics[A_Index]
            base := Excel_Layout_BasePrefer(m)
            if (A_Index = focusIdx) {
                prefer4.Push(base * 3.0)
                doWrap4.Push(true)
                wrapCount++
            } else if (m["kind"] = "short") {
                prefer4.Push(base * 1.15)
                doWrap4.Push(false)
                shortCount++
            } else {
                prefer4.Push(base * 0.45)
                doWrap4.Push(false)
                if (m["kind"] = "single")
                    singleCount++
                else
                    wrapCount++
            }
        }
        focusHeader := metrics[focusIdx]["header"]
        return Map("prefer", prefer4, "doWrap", doWrap4, "wrapCount", wrapCount,
            "shortCount", shortCount, "singleCount", singleCount,
            "focusHeader", focusHeader, "maxRowHeight", maxRowHeight)
    }

    if (pillar = 5) {
        ; Triage — boost short/categorical cols; one title readable; crush extremes
        titleIdx := 0
        titleLen := -1
        loop colCount {
            m := metrics[A_Index]
            if (m["kind"] = "wrap" && m["maxLen"] > titleLen) {
                titleLen := m["maxLen"]
                titleIdx := A_Index
            }
        }
        ; Prefer shorter wrap as title label when multiple exist
        if (titleIdx > 0) {
            shortestWrap := 0x7FFFFFFF
            loop colCount {
                m := metrics[A_Index]
                if (m["kind"] = "wrap" && m["maxLen"] < shortestWrap) {
                    shortestWrap := m["maxLen"]
                    titleIdx := A_Index
                }
            }
        }
        prefer5 := []
        doWrap5 := []
        wrapCount := 0
        shortCount := 0
        singleCount := 0
        loop colCount {
            m := metrics[A_Index]
            base := Excel_Layout_BasePrefer(m)
            if (m["kind"] = "short") {
                prefer5.Push(base * 2.4)
                doWrap5.Push(false)
                shortCount++
            } else if (A_Index = titleIdx) {
                prefer5.Push(base * 2.0)
                doWrap5.Push(true)
                wrapCount++
            } else {
                prefer5.Push(Max(3.5, base * 0.35))
                doWrap5.Push(false)
                if (m["kind"] = "single")
                    singleCount++
                else
                    shortCount++
            }
        }
        focusHeader := (titleIdx > 0) ? metrics[titleIdx]["header"] : ""
        return Map("prefer", prefer5, "doWrap", doWrap5, "wrapCount", wrapCount,
            "shortCount", shortCount, "singleCount", singleCount,
            "focusHeader", focusHeader, "maxRowHeight", SCAN_ROW_HEIGHT + 8)
    }

    if (pillar = 6) {
        ; Immersive — deep-read longest non-URL prose (incl. extreme); force wrap
        focusIdx := 0
        focusLen := -1
        loop colCount {
            m := metrics[A_Index]
            if (m["isUrlish"])
                continue
            if (m["maxLen"] > focusLen && m["maxLen"] > 18) {
                focusLen := m["maxLen"]
                focusIdx := A_Index
            }
        }
        if (focusIdx = 0) {
            return Map("prefer", prefer, "doWrap", doWrap, "wrapCount", wrapCount,
                "shortCount", shortCount, "singleCount", singleCount,
                "focusHeader", "", "maxRowHeight", maxRowHeight)
        }
        prefer6 := []
        doWrap6 := []
        wrapCount := 0
        shortCount := 0
        singleCount := 0
        loop colCount {
            m := metrics[A_Index]
            base := Excel_Layout_BasePrefer(m)
            if (A_Index = focusIdx) {
                prefer6.Push(Max(base, 36) * 5.0)
                doWrap6.Push(true)
                wrapCount++
            } else {
                prefer6.Push(base * 0.4)
                doWrap6.Push(false)
                if (m["kind"] = "single")
                    singleCount++
                else
                    shortCount++
            }
        }
        focusHeader := metrics[focusIdx]["header"]
        return Map("prefer", prefer6, "doWrap", doWrap6, "wrapCount", wrapCount,
            "shortCount", shortCount, "singleCount", singleCount,
            "focusHeader", focusHeader, "maxRowHeight", 96)
    }

    ; Unknown pillar → Balanced
    return Map("prefer", prefer, "doWrap", doWrap, "wrapCount", wrapCount,
        "shortCount", shortCount, "singleCount", singleCount,
        "focusHeader", "", "maxRowHeight", maxRowHeight)
}

Excel_Layout_Apply(xl, ws, tableRange, prefer, doWrap, maxRowHeight) {
    static MIN_COL_CHARS := 3.5
    static GUTTER_PTS := 36

    startRow := tableRange.Row
    startCol := tableRange.Column
    rowCount := tableRange.Rows.Count
    colCount := tableRange.Columns.Count
    endRow := startRow + rowCount - 1

    ptsPerChar := 7.0
    try {
        probe := ws.Columns(startCol)
        cw := probe.ColumnWidth
        pw := probe.Width
        if (cw > 0 && pw > 0)
            ptsPerChar := pw / cw
    } catch {
    }

    usablePts := 0
    try usablePts := xl.ActiveWindow.UsableWidth
    catch {
        try usablePts := xl.UsableWidth
        catch {
            usablePts := 0
        }
    }
    budgetPts := Max(120.0, usablePts - GUTTER_PTS)

    sumPrefer := 0.0
    for w in prefer
        sumPrefer += w
    if (sumPrefer <= 0)
        sumPrefer := colCount * 10.0

    scale := budgetPts / (sumPrefer * ptsPerChar)
    allocated := []
    sumAlloc := 0.0
    wrapCount := 0
    loop colCount {
        w := Max(MIN_COL_CHARS, prefer[A_Index] * scale)
        allocated.Push(w)
        sumAlloc += w
        if doWrap[A_Index]
            wrapCount++
    }
    leftoverChars := (budgetPts / ptsPerChar) - sumAlloc
    if (leftoverChars > 0.5 && wrapCount > 0) {
        addEach := leftoverChars / wrapCount
        loop colCount {
            if doWrap[A_Index]
                allocated[A_Index] += addEach
        }
    } else if (leftoverChars > 0.5) {
        addEach := leftoverChars / colCount
        loop colCount
            allocated[A_Index] += addEach
    }

    prevScreen := true
    try prevScreen := xl.ScreenUpdating
    try {
        xl.ScreenUpdating := false
        loop colCount {
            c := startCol + A_Index - 1
            colRng := ws.Range(ws.Cells(startRow, c), ws.Cells(endRow, c))
            colRng.WrapText := doWrap[A_Index]
            ws.Columns(c).ColumnWidth := allocated[A_Index]
        }
        tableRange.Rows.AutoFit()
        r := startRow
        while (r <= endRow) {
            try {
                if (ws.Rows(r).RowHeight > maxRowHeight)
                    ws.Rows(r).RowHeight := maxRowHeight
            } catch {
            }
            r++
        }
        tableRange.Select()
    } finally {
        try xl.ScreenUpdating := prevScreen
    }
    return true
}

Excel_Layout_CyclePillar() {
    global Excel_LayoutPillarIndex, Excel_LayoutPillarCount

    try {
        xl := ComObjActive("Excel.Application")
        ws := xl.ActiveSheet
        cell := xl.ActiveCell
    } catch {
        ShowCenteredOverlay_Utils("❌ Excel COM unavailable", 2200, BANNER_ACCENT_ERROR)
        return false
    }

    tableRange := Excel_ReadableLayout_ResolveRange(xl, ws, cell)
    if !tableRange {
        ShowCenteredOverlay_Utils("❌ Nothing to layout on sheet", 2000, BANNER_ACCENT_ERROR)
        return false
    }

    try {
        if (tableRange.Columns.Count < 1 || tableRange.Rows.Count < 1) {
            ShowCenteredOverlay_Utils("❌ Nothing to layout on sheet", 2000, BANNER_ACCENT_ERROR)
            return false
        }
    } catch as e {
        ShowCenteredOverlay_Utils("❌ Layout range failed`n" e.Message, 2500, BANNER_ACCENT_ERROR)
        return false
    }

    n := Excel_LayoutPillarCount
    if (n < 1)
        n := 7
    Excel_LayoutPillarIndex := Mod(Excel_LayoutPillarIndex + 1, n)
    pillar := Excel_LayoutPillarIndex
    name := Excel_Layout_PillarName(pillar)

    try {
        metrics := Excel_Layout_Measure(ws, tableRange)
        prefs := Excel_Layout_PrefsForPillar(metrics, pillar)
        Excel_Layout_Apply(xl, ws, tableRange, prefs["prefer"], prefs["doWrap"], prefs["maxRowHeight"])
    } catch as e {
        ShowCenteredOverlay_Utils("❌ Layout failed`n" e.Message, 2500, BANNER_ACCENT_ERROR)
        return false
    }

    msg := "📐 " . (pillar + 1) . "/" . n . " " . name
    if (prefs["focusHeader"] != "")
        msg .= " · " . prefs["focusHeader"]
    ShowCenteredOverlay_Utils(msg, 1800, BANNER_ACCENT_SUCCESS)
    return true
}

; After CSV Load: autofit, cap widths >20 → 20, wrap text, center H/V on imported table.
; Soft no-op on COM failure so a successful import is not undone.
Excel_FormatImportedTable(maxWidth := 20) {
    try {
        xl := ComObjActive("Excel.Application")
        ws := xl.ActiveSheet
    } catch {
        return
    }
    try {
        tableRange := 0
        loCount := 0
        try loCount := ws.ListObjects.Count
        catch {
        }
        if (loCount >= 1) {
            try tableRange := ws.ListObjects(1).Range
            catch {
                tableRange := 0
            }
        }
        if (!tableRange) {
            try tableRange := ws.UsedRange
            catch {
                tableRange := 0
            }
        }
        if (!tableRange)
            return
        tableRange.Columns.AutoFit()
        colCount := tableRange.Columns.Count
        startCol := tableRange.Column
        loop colCount {
            col := ws.Columns(startCol + A_Index - 1)
            if (col.ColumnWidth > maxWidth)
                col.ColumnWidth := maxWidth
        }
        tableRange.WrapText := true
        tableRange.HorizontalAlignment := -4108  ; xlCenter
        tableRange.VerticalAlignment := -4108    ; xlCenter
        tableRange.Select()
    } catch {
    }
}

; After CSV Load: paint entire sheet dark gray, clear fill on imported table only.
; Soft no-op on COM failure so a successful import is not undone.
Excel_ShadeOutsideImportedTable(color := 0x505050) {
    try {
        xl := ComObjActive("Excel.Application")
        ws := xl.ActiveSheet
    } catch {
        return
    }
    try {
        tableRange := 0
        loCount := 0
        try loCount := ws.ListObjects.Count
        catch {
        }
        if (loCount >= 1) {
            try tableRange := ws.ListObjects(1).Range
            catch {
                tableRange := 0
            }
        }
        if (!tableRange) {
            try tableRange := ws.UsedRange
            catch {
                tableRange := 0
            }
        }
        if (!tableRange)
            return
        ws.Cells.Interior.Color := color
        tableRange.Interior.ColorIndex := -4142  ; xlColorIndexNone
        tableRange.Select()
    } catch {
    }
}

; Open Data → From Text/CSV (ExecuteMso, then UIA name fallback).
Excel_OpenFromTextCsv() {
    try {
        xl := ComObjActive("Excel.Application")
        xl.CommandBars.ExecuteMso("PowerQueryNewFromTextCsv")
        return true
    } catch {
    }
    excelHwnd := WinExist("ahk_exe EXCEL.EXE")
    if !excelHwnd
        return false
    try {
        root := UIA.ElementFromHandle(excelHwnd)
        for name in ["From Text/CSV", "De Texto/CSV", "Texto/CSV"] {
            btn := 0
            try btn := WaitForButton(root, name, 1500)
            catch {
                btn := 0
            }
            if btn {
                try {
                    btn.Invoke()
                    return true
                } catch {
                }
                try {
                    btn.Click()
                    return true
                } catch {
                }
            }
        }
    } catch {
    }
    return false
}

; Wait for Excel's file-open dialog (#32770).
Excel_WaitForFileDialog(timeoutMs := 15000) {
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        hwnd := 0
        try hwnd := WinExist("ahk_class #32770 ahk_exe EXCEL.EXE")
        catch {
            hwnd := 0
        }
        if !hwnd {
            try {
                for id in WinGetList("ahk_class #32770") {
                    try {
                        if (WinGetProcessName("ahk_id " id) = "EXCEL.EXE") {
                            hwnd := id
                            break
                        }
                    } catch {
                    }
                }
            } catch {
            }
        }
        if hwnd
            return hwnd
        Sleep 200
    }
    return 0
}

; Clipboard holds CSV path → From Text/CSV → paste → FileDialog_ImportCsvLoad.
Excel_ImportCsvFromClipboardPath() {
    hwndExcel := WinExist("A")
    try {
        StandardLoadingBar_Show("⏳ Opening From Text/CSV…", BANNER_ACCENT_INTERMEDIATE, {
            passive: false, centerOnHwnd: hwndExcel, textWidth: 480 })
    } catch {
    }
    if !Excel_OpenFromTextCsv() {
        try StandardLoadingBar_Update("❌ From Text/CSV not found", BANNER_ACCENT_ERROR)
        catch {
        }
        try StandardLoadingBar_Hide(800)
        catch {
        }
        return
    }
    try StandardLoadingBar_Update("⏳ Waiting for file dialog…", BANNER_ACCENT_INTERMEDIATE)
    catch {
    }
    dlgHwnd := Excel_WaitForFileDialog()
    if !dlgHwnd {
        try StandardLoadingBar_Update("❌ File dialog not found", BANNER_ACCENT_ERROR)
        catch {
        }
        try StandardLoadingBar_Hide(800)
        catch {
        }
        return
    }
    try StandardLoadingBar_Update("⏳ Pasting CSV path…", BANNER_ACCENT_INTERMEDIATE)
    catch {
    }
    try WinActivate("ahk_id " dlgHwnd)
    catch {
    }
    Sleep 200
    FileDialog_FocusFileNameField()
    Sleep 50
    Send "^v"
    Sleep 300
    try StandardLoadingBar_Hide(200)
    catch {
    }
    FileDialog_ImportCsvLoad()
    ; Same as Shift+U: F12 → paste clipboard path → CSV UTF-8
    Sleep 400
    Excel_SaveCsvUtf8FromClipboardPath()
}

; Clipboard holds CSV path → F12 Save As → paste → FileDialog_SaveAsCsvUtf8.
Excel_SaveCsvUtf8FromClipboardPath() {
    hwndExcel := WinExist("A")
    try {
        StandardLoadingBar_Show("⏳ Opening Save As…", BANNER_ACCENT_INTERMEDIATE, {
            passive: false, centerOnHwnd: hwndExcel, textWidth: 480 })
    } catch {
    }
    Send "{F12}"
    try StandardLoadingBar_Update("⏳ Waiting for file dialog…", BANNER_ACCENT_INTERMEDIATE)
    catch {
    }
    dlgHwnd := Excel_WaitForFileDialog()
    if !dlgHwnd {
        try StandardLoadingBar_Update("❌ File dialog not found", BANNER_ACCENT_ERROR)
        catch {
        }
        try StandardLoadingBar_Hide(800)
        catch {
        }
        return
    }
    try StandardLoadingBar_Update("⏳ Pasting CSV path…", BANNER_ACCENT_INTERMEDIATE)
    catch {
    }
    try WinActivate("ahk_id " dlgHwnd)
    catch {
    }
    Sleep 200
    FileDialog_FocusFileNameField()
    Sleep 50
    saveName := FileDialog_NormalizeCsvSaveName(A_Clipboard)
    if (saveName != "")
        SendText(saveName)
    else
        Send "^v"
    Sleep 300
    try StandardLoadingBar_Update("⏳ Saving CSV UTF-8…", BANNER_ACCENT_INTERMEDIATE)
    catch {
    }
    try StandardLoadingBar_Hide(200)
    catch {
    }
    FileDialog_SaveAsCsvUtf8()
}

; Shift + N : Cycle layout pillars (7 views)
+n:: {
    Excel_Layout_CyclePillar()
}

; Resolve http(s) URL from the active cell: Hyperlinks collection, =HYPERLINK()
; formula, or plain cell text. Returns Map("ok", "err", "url").
Excel_ResolveActiveCellUrl() {
    try {
        xl := ComObjActive("Excel.Application")
        cell := xl.ActiveCell
    } catch {
        return Map("ok", false, "err", "Excel COM unavailable", "url", "")
    }
    if !cell
        return Map("ok", false, "err", "No active cell", "url", "")

    try {
        if (cell.Hyperlinks.Count >= 1) {
            addr := Trim(String(cell.Hyperlinks(1).Address))
            if StudyLink_IsValidHttpUrl(addr)
                return Map("ok", true, "err", "", "url", addr)
            extracted := StudyLink_ExtractUrlFromClipboardText(addr)
            if (extracted != "")
                return Map("ok", true, "err", "", "url", extracted)
        }
    } catch {
    }

    try {
        formula := Trim(String(cell.Formula))
        if RegExMatch(formula, 'i)HYPERLINK\s*\(\s*"([^"]+)"', &m) {
            cand := Trim(m[1])
            if StudyLink_IsValidHttpUrl(cand)
                return Map("ok", true, "err", "", "url", cand)
            extracted := StudyLink_ExtractUrlFromClipboardText(cand)
            if (extracted != "")
                return Map("ok", true, "err", "", "url", extracted)
        }
    } catch {
    }

    try {
        val := ""
        try val := Trim(String(cell.Value2))
        catch {
            try val := Trim(String(cell.Text))
            catch {
                val := ""
            }
        }
        extracted := StudyLink_ExtractUrlFromClipboardText(val)
        if (extracted != "")
            return Map("ok", true, "err", "", "url", extracted)
    } catch {
    }

    return Map("ok", false, "err", "No http(s) link in active cell", "url", "")
}

; Shift + L : Open active cell hyperlink / URL in a new Chrome window (COM + StudyLink)
Excel_OpenActiveCellHyperlinkInChrome() {
    r := Excel_ResolveActiveCellUrl()
    if !r["ok"] {
        ShowCenteredOverlay_Utils("❌ " . r["err"], 2200, BANNER_ACCENT_ERROR)
        return false
    }
    if !StudyLink_OpenUrlInChrome(r["url"], true) {
        ShowCenteredOverlay_Utils("❌ Could not open Chrome.", 2000, BANNER_ACCENT_ERROR)
        return false
    }
    ShowCenteredOverlay_Utils("🌐 Opening link in new Chrome…", 1800, BANNER_ACCENT_SUCCESS)
    return true
}

+l:: Excel_OpenActiveCellHyperlinkInChrome()

; Shift + F : Fill series downward (fill-handle equivalent) to last target row.
; Target = max(UsedRange last row, enclosing ListObject last row) so empty table
; body rows still count. Selection = seed (1 cell, or 2+ for step); COM AutoFill.
; Confirm first: COM AutoFill can clear Excel's undo stack, so ask before committing.
Excel_FillSeriesToUsedRange() {
    static XL_FILL_DEFAULT := 0
    try {
        xl := ComObjActive("Excel.Application")
        ws := xl.ActiveSheet
        sel := xl.Selection
    } catch {
        ShowCenteredOverlay_Utils("❌ Excel COM unavailable", 2200, BANNER_ACCENT_ERROR)
        return false
    }

    try {
        src := sel
        try {
            if (sel.Areas.Count > 1)
                src := sel.Areas(1)
        } catch {
        }
        startRow := src.Row
        startCol := src.Column
        srcRows := src.Rows.Count
        srcCols := src.Columns.Count
        srcEndRow := startRow + srcRows - 1
    } catch {
        ShowCenteredOverlay_Utils("❌ Select cells to fill from", 2000, BANNER_ACCENT_ERROR)
        return false
    }

    ; UsedRange often stops at last value and skips empty Excel Table rows.
    lastRow := 0
    try {
        ur := ws.UsedRange
        lastRow := ur.Row + ur.Rows.Count - 1
    } catch {
    }
    try {
        loCount := 0
        try loCount := ws.ListObjects.Count
        catch {
        }
        loop loCount {
            lo := ws.ListObjects(A_Index)
            try {
                if xl.Intersect(src, lo.Range) {
                    tableLast := lo.Range.Row + lo.Range.Rows.Count - 1
                    if (tableLast > lastRow)
                        lastRow := tableLast
                    break
                }
            } catch {
            }
        }
    } catch {
    }
    if (lastRow = 0) {
        ShowCenteredOverlay_Utils("❌ No used range on sheet", 2000, BANNER_ACCENT_ERROR)
        return false
    }

    if (lastRow <= srcEndRow) {
        ShowCenteredOverlay_Utils("❌ Nothing to fill — already at last used row", 2200, BANNER_ACCENT_ERROR)
        return false
    }

    filled := lastRow - srcEndRow
    if MsgBox(
        "Fill series down +" . filled . " row(s) to last table/used row?`n`n"
        . "This may clear Excel's undo history (Ctrl+Z).",
        "Confirm fill series",
        "Icon? YesNo Default2"
    ) != "Yes"
        return false

    prevScreen := true
    try prevScreen := xl.ScreenUpdating
    try {
        dest := ws.Range(ws.Cells(startRow, startCol), ws.Cells(lastRow, startCol + srcCols - 1))
        xl.ScreenUpdating := false
        src.AutoFill(dest, XL_FILL_DEFAULT)
    } catch as e {
        ShowCenteredOverlay_Utils("❌ Fill failed`n" e.Message, 2500, BANNER_ACCENT_ERROR)
        return false
    } finally {
        try xl.ScreenUpdating := prevScreen
    }

    ShowCenteredOverlay_Utils("🔢 Filled series +" . filled . " row(s)", 1600, BANNER_ACCENT_SUCCESS)
    return true
}

+f:: Excel_FillSeriesToUsedRange()

; Shift + O : Quick organize — select all (table incl. headers, else UsedRange),
; center align, font size 11. COM only (no Ctrl+A / ribbon).
Excel_QuickOrganizeCenterFont11() {
    static XL_CENTER := -4108
    try {
        xl := ComObjActive("Excel.Application")
        ws := xl.ActiveSheet
        cell := xl.ActiveCell
    } catch {
        ShowCenteredOverlay_Utils("❌ Excel COM unavailable", 2200, BANNER_ACCENT_ERROR)
        return false
    }

    rng := ""
    try {
        loCount := 0
        try loCount := ws.ListObjects.Count
        catch {
        }
        loop loCount {
            lo := ws.ListObjects(A_Index)
            try {
                if xl.Intersect(cell, lo.Range) {
                    rng := lo.Range
                    break
                }
            } catch {
            }
        }
    } catch {
    }
    if (rng = "") {
        try rng := ws.UsedRange
        catch {
            ShowCenteredOverlay_Utils("❌ Nothing to format on sheet", 2000, BANNER_ACCENT_ERROR)
            return false
        }
    }
    if !rng {
        ShowCenteredOverlay_Utils("❌ Nothing to format on sheet", 2000, BANNER_ACCENT_ERROR)
        return false
    }

    prevScreen := true
    try prevScreen := xl.ScreenUpdating
    try {
        xl.ScreenUpdating := false
        rng.HorizontalAlignment := XL_CENTER
        rng.Font.Size := 11
        rng.Select()
    } catch as e {
        ShowCenteredOverlay_Utils("❌ Organize failed`n" e.Message, 2500, BANNER_ACCENT_ERROR)
        return false
    } finally {
        try xl.ScreenUpdating := prevScreen
    }

    ShowCenteredOverlay_Utils("📐 Centered · font 11", 1600, BANNER_ACCENT_SUCCESS)
    return true
}

+o:: Excel_QuickOrganizeCenterFont11()

; Shift + R : Read active cell in a dark centered modal (toggle; Esc also closes)
global g_ExcelCellPreviewGui := 0
global g_ExcelCellPreviewBorderGui := 0

Excel_CellPreview_IsOpen() {
    global g_ExcelCellPreviewGui
    try return IsObject(g_ExcelCellPreviewGui) && !!g_ExcelCellPreviewGui.Hwnd
    catch {
        return false
    }
}

Excel_CellPreview_Close(*) {
    global g_ExcelCellPreviewGui, g_ExcelCellPreviewBorderGui
    try {
        if IsObject(g_ExcelCellPreviewGui)
            g_ExcelCellPreviewGui.Destroy()
    } catch {
    }
    g_ExcelCellPreviewGui := 0
    try {
        if IsObject(g_ExcelCellPreviewBorderGui)
            g_ExcelCellPreviewBorderGui.Destroy()
    } catch {
    }
    g_ExcelCellPreviewBorderGui := 0
}

Excel_CellPreview_Toggle() {
    if Excel_CellPreview_IsOpen() {
        Excel_CellPreview_Close()
        return
    }

    try {
        xl := ComObjActive("Excel.Application")
        cell := xl.ActiveCell
    } catch {
        ShowCenteredOverlay_Utils("❌ Excel COM unavailable", 2200, BANNER_ACCENT_ERROR)
        return
    }
    if !cell {
        ShowCenteredOverlay_Utils("❌ No active cell", 2000, BANNER_ACCENT_ERROR)
        return
    }

    text := ""
    try text := String(cell.Text)
    catch {
        try text := String(cell.Value2)
        catch {
            text := ""
        }
    }
    text := Trim(text, " `t`r`n")
    if (text = "")
        text := "∅ Empty cell"
    else if (StrLen(text) > 4000)
        text := SubStr(text, 1, 3997) . "…"

    hwndExcel := WinExist("A")
    ml := 0, mt := 0, mr := 0, mb := 0
    workArea := ""
    try workArea := GetWorkAreaForWindow_StandardBar(hwndExcel)
    catch {
        workArea := ""
    }
    if (IsObject(workArea)) {
        ml := workArea.left
        mt := workArea.top
        mr := workArea.right
        mb := workArea.bottom
    } else {
        GetActiveMonitorWorkArea_StandardBar(&ml, &mt, &mr, &mb)
    }
    monW := mr - ml
    monH := mb - mt
    maxW := Max(280, Min(900, Floor(monW * 0.7)))
    maxH := Max(120, Floor(monH * 0.7))

    Excel_CellPreview_Close()
    global g_ExcelCellPreviewGui, g_ExcelCellPreviewBorderGui

    g := Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale")
    g.BackColor := "1E1E2E"
    g.MarginX := 36
    g.MarginY := 32
    g.SetFont("s24 cFFFFFF", "Segoe UI")
    g.Add("Text", "w" . maxW . " Center Wrap", text)
    g.OnEvent("Close", Excel_CellPreview_Close)
    g.OnEvent("Escape", Excel_CellPreview_Close)
    g.Show("AutoSize Hide")
    g.GetPos(, , &gw, &gh)
    if (gh > maxH) {
        try g.Destroy()
        catch {
        }
        g := Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale")
        g.BackColor := "1E1E2E"
        g.MarginX := 36
        g.MarginY := 32
        g.SetFont("s24 cFFFFFF", "Segoe UI")
        bodyH := maxH - 64
        if (bodyH < 80)
            bodyH := 80
        g.Add("Text", "w" . maxW . " h" . bodyH . " Center Wrap", text)
        g.OnEvent("Close", Excel_CellPreview_Close)
        g.OnEvent("Escape", Excel_CellPreview_Close)
        g.Show("AutoSize Hide")
        g.GetPos(, , &gw, &gh)
    }

    guiX := Round(ml + (monW - gw) / 2)
    guiY := Round(mt + (monH - gh) / 2)
    if (guiX < ml)
        guiX := ml
    if (guiY < mt)
        guiY := mt
    if (guiX + gw > mr)
        guiX := mr - gw
    if (guiY + gh > mb)
        guiY := mb - gh

    borderWidth := 4
    borderGui := Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale")
    borderGui.BackColor := "3D3D5C"
    borderGui.Show("NA x" . (guiX - borderWidth) . " y" . (guiY - borderWidth)
    . " w" . (gw + 2 * borderWidth) . " h" . (gh + 2 * borderWidth))
    g_ExcelCellPreviewBorderGui := borderGui

    g.Show("x" . guiX . " y" . guiY)
    WinSetTransparent(245, g)
    g_ExcelCellPreviewGui := g
    try WinActivate("ahk_id " g.Hwnd)
    catch {
    }
}

+r:: Excel_CellPreview_Toggle()

#HotIf Excel_CellPreview_IsOpen()
Escape:: Excel_CellPreview_Close()
+r:: Excel_CellPreview_Close()
#HotIf