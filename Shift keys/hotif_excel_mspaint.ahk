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
#HotIf WinActive("ahk_exe EXCEL.EXE") && WinGetClass("A") != "#32770"

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

; Shift + V : Quickly paste and extract CSV (Paste, CSV to columns)
+v:: {
    Excel_AddMultipleRows()    ; Add multiple rows first
    Sleep 200
    Send "^{Up}"
    Sleep 300
    Send "{Down}"
    Sleep 200
    Send "^v"           ; Ctrl+V (Paste action)
    Sleep 200
    Excel_CSVToColumns(true)    ; Auto-select semicolon, bypass dialog
    Sleep 200
    Send "^{Down}"
    Send "{Down}"
    Send "{Shift down}"
    Send "^{Down}"
    Send "{Shift up}"
    Send "{Alt}"
    Send "3"
    Send "r"
    ; Send "{Down}"
    ; Excel_RemoveRows(10)
    ; Sleep 200
    ; Send "^{Up}"
}

; Shift + C : Turn CSV delimited by semicolon into columns (Alt, 0, 5, D, Enter, M, Enter, Enter)
+c:: {
    Excel_CSVToColumns()
}

; Helper function: Add multiple rows (repeat Alt, Alt, 0, 2 with delays)
Excel_AddMultipleRows(count := 15) {
    ShowSmallLoadingIndicator_ChatGPT("Adding " . count . " rows...")
    ; Extra initial delay so the first Alt+0,2 sequence isn't too fast
    Sleep 300
    loop count {
        Send "{Alt down}"
        Send "{Alt up}"
        Sleep 100
        Send "0"
        Sleep 50
        Send "2"
        Sleep 50
    }
    HideSmallLoadingIndicator_ChatGPT()
}

; Shift + A : Add multiple rows (repeat Alt, Alt, 0, 2 with delays)
+a:: {
    Excel_AddMultipleRows()    ; Call function directly
}

; Helper function: Row removal workflow (remove row, down arrow, repeat)
; Pre-condition: Place cursor in starting cell
; Step 1: Execute REMOVE ROW SHORTCUT (Alt, Alt, 3, R)
; Step 2: Press DOWN ARROW
; Step 3: Repeat Step 1 and Step 2 for specified iterations
; Purpose: Remove alternating sequences of empty and populated rows
Excel_RemoveRows(iterations := 8) {
    ShowSmallLoadingIndicator_ChatGPT("Removing rows...")
    loop iterations {
        ; Execute Remove Row shortcut (Alt, Alt, 3, R)
        Send "{Alt down}"
        Send "{Alt up}"
        Sleep 150
        Send "3"
        Sleep 100
        Send "r"
        Sleep 150
        ; Press Down Arrow
        Send "{Down}"
        Sleep 150
    }
    HideSmallLoadingIndicator_ChatGPT()
}

; Shift + R : Row removal workflow (remove row, down arrow, repeat 5-7 times)
+r:: {
    Excel_RemoveRows()    ; Call function directly
}

; Shift + P : Type previous day date
+p:: {
    ; Calculate yesterday's date
    ; Get current date/time and subtract exactly 24 hours (86400 seconds)
    currentTime := A_Now
    yesterdayTime := DateAdd(currentTime, -86400, "Seconds")
    ; Format as MM/dd/yyyy (MM/DD/YYYY format)
    dateStr := FormatTime(yesterdayTime, "MM/dd/yyyy")
    ; Small delay to ensure Excel is ready
    Sleep 50
    ; Type the date as a whole word at once
    SendText dateStr
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

; Helper: Autofit used columns, then cap any width greater than maxWidth to cappedWidth.
; Ends by selecting row 1 (A through last used column) and Alt, 0, 6 (Zoom to Selection).
Excel_NormalizeColumnWidths(maxWidth := 15, cappedWidth := 5) {
    ShowSmallLoadingIndicator_ChatGPT("Normalizing column widths...")
    try {
        xl := ComObjActive("Excel.Application")
        ws := xl.ActiveSheet
        ur := ws.UsedRange
        if (!ur) {
            HideSmallLoadingIndicator_ChatGPT()
            MsgBox("No used range on the active sheet.")
            return
        }
        ur.Columns.AutoFit()
        colCount := ur.Columns.Count
        startCol := ur.Column
        lastCol := startCol + colCount - 1
        loop colCount {
            col := ws.Columns(startCol + A_Index - 1)
            if (col.ColumnWidth > maxWidth)
                col.ColumnWidth := cappedWidth
        }
        ; First cell, then first row across used columns
        ws.Range("A1").Select()
        ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Select()
    } catch Error as err {
        HideSmallLoadingIndicator_ChatGPT()
        MsgBox("Could not normalize column widths:`n" err.Message)
        return
    }
    HideSmallLoadingIndicator_ChatGPT()
    ; QAT: Zoom to Selection (Alt, 0, 6)
    Sleep 100
    Send "{Alt}"
    Sleep 100
    Send "0"
    Sleep 100
    Send "6"
}

; Shift + N : Narrow oversized columns (autofit, then cap width >15 → 5)
+n:: {
    Excel_NormalizeColumnWidths()
}
