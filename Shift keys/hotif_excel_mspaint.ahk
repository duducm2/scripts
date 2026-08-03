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
