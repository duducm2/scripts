; =============================================================================
; Shift keys module: hotif_file_dialog.ahk
; File dialog hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

#HotIf IsFileDialogActive()

; Shift + F : Select first file - File
+f:: {
    if !IsFileDialogActive()
        return
    try {
        root := UIA.ElementFromHandle(WinExist("A"))

        ; First attempt: find and focus Items View list directly
        itemsList := root.FindFirst({ Type: "List", ClassName: "UIItemsView" })
        if !itemsList
            itemsList := root.FindFirst({ Type: "List", Name: "Items View" })
        if !itemsList
            itemsList := root.FindFirst({ Type: "List", AutomationId: "ItemsView" })

        if itemsList {
            itemsList.SetFocus()
            Sleep 120
            Send "{Home}"  ; Go to first item
            EnsureFocus()
            return
        }

        ; Second attempt: find header (Header control)
        hdr := root.FindFirst({ Type: "Header" })
        if !hdr {
            hdr := root.FindFirst({ Name: "Header", Type: "Header" })
            if !hdr
                hdr := root.FindFirst({ Name: "CabeÃ§alho", Type: "Header" })
        }
        if hdr {
            hdr.SetFocus()
            Sleep 120
            Send "+{Tab}"
            Send "{Home}"
            EnsureFocus()
            return
        }

        ; Third attempt: find file name ComboBox by AutomationId and Type
        ; This should work regardless of the name (File name: or Nome:)
        fileNameCombo := root.FindFirst({ Type: "ComboBox", AutomationId: "1148" })

        ; If not found, try by various possible names
        if !fileNameCombo {
            possibleNames := ["File name:", "Nome:", "Filename:", "File Name:"]
            for name in possibleNames {
                fileNameCombo := root.FindFirst({ Type: "ComboBox", Name: name })
                if fileNameCombo
                    break
            }
        }

        ; If ComboBox found, use it
        if fileNameCombo {
            fileNameCombo.SetFocus()
            Sleep 120
            Send "+{Tab}"
            Send "{Home}"
            EnsureFocus()
            return
        }
    } catch Error {
    }
    ; Last resort fallback: simple Shift+Tab then Home
    Send "+{Tab}"
    Sleep 120
    Send "{Home}"
    EnsureFocus()

}

; Shift + S : Focus search bar - Search bar
+s:: {
    if !IsFileDialogActive()
        return
    Send "^e"
}

; Shift + A : Focus address bar - Address bar
+a:: {
    if !IsFileDialogActive()
        return
    try {
        root := UIA.ElementFromHandle(WinExist("A"))
        ; Try to find address bar by common names
        addressBar := root.FindFirst({ Type: "Edit", Name: "Address:" })
        if !addressBar
            addressBar := root.FindFirst({ Type: "Edit", Name: "Endereço:" })
        if !addressBar
            addressBar := root.FindFirst({ Type: "ComboBox", AutomationId: "1001" })
        if !addressBar
            addressBar := root.FindFirst({ Type: "Edit", ClassName: "Edit" })

        if (addressBar) {
            addressBar.SetFocus()
            Sleep 50
            Send "^a"  ; Select all existing text
            return
        }
    } catch Error {
    }
    ; Fallback: Use Alt+D (common shortcut for address bar in file dialogs)
    Send "!d"
}

; Shift + N : New folder - New Folder
+n:: {
    if !IsFileDialogActive()
        return
    Send "^+n"
}

; Shift + P : Select first pinned item in sidebar - Pinned item
+p:: {
    if !IsFileDialogActive()
        return
    SelectExplorerSidebarFirstPinned()
}

; Shift + T : Select "This PC" / "Este computador" in sidebar - This PC
+t:: {
    if !IsFileDialogActive()
        return
    SelectExplorerSidebarFirstPinned()
    Sleep 200
    Send "{End}"
    Send "{Up}"
    Send "{Up}"
    Send "{Up}"
    Send "{Up}"
}

; Shift + M : Focus file name edit field - Name
+m:: {
    if !IsFileDialogActive()
        return
    try {
        root := UIA.ElementFromHandle(WinExist("A"))

        ; First attempt: Find by AutomationId and Type (most reliable)
        fileNameEdit := root.FindFirst({ Type: "Edit", AutomationId: "1148" })

        ; Second attempt: Try various possible names
        if !fileNameEdit {
            possibleNames := [
                "File name:",      ; English standard
                "Nome:",          ; Portuguese standard
                "Filename:",      ; Alternative English
                "File Name:",     ; Alternative capitalization
                "Name:",          ; Generic English
                "Nome do arquivo:", ; Full Portuguese
                "Save As:",       ; Save dialog English
                "Salvar como:"    ; Save dialog Portuguese
            ]
            for name in possibleNames {
                fileNameEdit := root.FindFirst({ Type: "Edit", Name: name })
                if fileNameEdit
                    break
            }
        }

        ; Third attempt: Try to find through parent ComboBox
        if !fileNameEdit {
            fileNameCombo := root.FindFirst({ Type: "ComboBox", AutomationId: "1148" })
            if fileNameCombo {
                fileNameEdit := fileNameCombo.FindFirst({ Type: "Edit" })
            }
        }

        if fileNameEdit {
            fileNameEdit.SetFocus()
            Sleep 50
            Send "^a"  ; Select all existing text
            return
        }
    } catch Error {
    }
    ; Fallback: Try to focus using keyboard navigation
    Send "!n"  ; Alt+N is a common shortcut for file name field
}

; Shift + O : Click Insert/Open/Save button - Open/Save
+o:: {
    if !IsFileDialogActive()
        return
    try {
        root := UIA.ElementFromHandle(WinExist("A"))

        ; First attempt: Find by AutomationId and Type (most reliable)
        actionBtn := root.FindFirst({ Type: "Button", AutomationId: "1" })

        ; Second attempt: Try various possible names
        if !actionBtn {
            possibleNames := [
                ; English variations
                "Insert",
                "Open",
                "Save",
                "Save As",
                "OK",
                ; Portuguese variations
                "Abrir",
                "Salvar",
                "Salvar como",
                "Inserir",
                ; Spanish variations (common in some systems)
                "Insertar",
                "Guardar",
                "Guardar como",
                ; French variations (common in some systems)
                "InsÃ©rer",
                "Ouvrir",
                "Enregistrer",
                "Enregistrer sous"
            ]
            for name in possibleNames {
                actionBtn := root.FindFirst({ Type: "Button", Name: name })
                if actionBtn
                    break
            }
        }

        ; Third attempt: Try SplitButton type (some dialogs use this instead)
        if !actionBtn {
            actionBtn := root.FindFirst({ Type: "SplitButton", AutomationId: "1" })
            if !actionBtn {
                for name in possibleNames {
                    actionBtn := root.FindFirst({ Type: "SplitButton", Name: name })
                    if actionBtn
                        break
                }
            }
        }

        if actionBtn {
            actionBtn.Click()
            return
        }
    } catch Error {
    }
    ; Fallback: Try common keyboard shortcuts
    Send "!s"  ; Alt+S (Save)
    Sleep 50
    Send "!o"  ; Alt+O (Open)
}

; Shift + C : Click Cancel button - Cancel
+c:: {
    if !IsFileDialogActive()
        return
    try {
        root := UIA.ElementFromHandle(WinExist("A"))

        ; First attempt: Find by AutomationId and Type (most reliable)
        cancelBtn := root.FindFirst({ Type: "Button", AutomationId: "2" })

        ; Second attempt: Try various possible names
        if !cancelBtn {
            possibleNames := [
                ; English variations
                "Cancel",
                "Close",
                "Exit",
                "Dismiss",
                ; Portuguese variations
                "Cancelar",
                "Fechar",
                "Sair",
                ; Spanish variations
                "Cancelar",
                "Cerrar",
                ; French variations
                "Annuler",
                "Fermer",
                ; German variations
                "Abbrechen",
                "SchlieÃŸen",
                ; Italian variations
                "Annulla",
                "Chiudi",
                ; Generic
                "No",
                "NÃ£o",
                "Ã—",  ; Sometimes used as close symbol
                "âœ•"   ; Alternative close symbol
            ]
            for name in possibleNames {
                cancelBtn := root.FindFirst({ Type: "Button", Name: name })
                if cancelBtn
                    break
            }
        }

        if cancelBtn {
            cancelBtn.Click()
            return
        }
    } catch Error {
    }
    ; Fallback: Try common keyboard shortcuts
    SendEscape()  ; Escape key is universal for cancel
}

; Shift + U : Save as CSV UTF-8, confirm overwrite + Excel multi-sheet warning
+u:: {
    if !IsFileDialogActive()
        return
    FileDialog_SaveAsCsvUtf8()
}

; Shift + I : Import CSV — Open → Load → close Queries pane
+i:: {
    if !IsFileDialogActive()
        return
    FileDialog_ImportCsvLoad()
}

#HotIf

; --- File Dialog helpers (CSV UTF-8 save flow) ---------------------------------

FileDialog_InvokeButton(btn) {
    if !btn
        return false
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
    return false
}

; Click Yes/OK/Sim/etc. without relying on Alt accelerators (EN vs PT differ).
FileDialog_ClickAffirmative(hwnd) {
    if !hwnd
        return false
    affirmativeNames := [
        "Yes", "Sim", "OK", "Confirmar", "Accept", "Aceitar",
        "Continue", "Continuar", "Proceed", "Prosseguir"
    ]
    affirmativeIds := ["CommandButton_6", "6", "CommandButton_1", "1"]
    title := ""
    try title := WinGetTitle("ahk_id " hwnd)
    catch {
    }
    try {
        root := UIA.ElementFromHandle(hwnd)
        btn := 0

        ; Prefer Name first: Excel Confirm Save As uses AutomationId "CommandButton_6",
        ; and FindFirst throws on missing AutomationId "6" (skipping later name match).
        for name in affirmativeNames {
            try {
                btn := root.FindFirst({ Type: "Button", Name: name })
            } catch {
                btn := 0
            }
            if btn
                break
        }
        if !btn {
            for id in affirmativeIds {
                try {
                    btn := root.FindFirst({ Type: "Button", AutomationId: id })
                } catch {
                    btn := 0
                }
                if btn
                    break
            }
        }
        ; Last UIA pass: scan all buttons for affirmative name (handles &Yes etc.)
        if !btn {
            try {
                for b in root.FindAll({ Type: "Button" }) {
                    bn := ""
                    try bn := b.Name
                    bnClean := StrReplace(bn, "&", "")
                    for name in affirmativeNames {
                        if (bnClean = name) {
                            btn := b
                            break 2
                        }
                    }
                }
            } catch {
            }
        }

        if btn && FileDialog_InvokeButton(btn)
            return true
    } catch {
    }
    try WinActivate("ahk_id " hwnd)
    catch {
    }
    ; Confirm Save As defaults to No — never send Enter. Prefer Alt+Y / Alt+S (EN/PT Yes/Sim).
    if FileDialog_IsConfirmSaveAsTitle(title) {
        Send "!y"
        Sleep 80
        if WinExist("ahk_id " hwnd)
            Send "!s"
        return true
    }
    Send "{Enter}"
    return true
}

FileDialog_IsConfirmSaveAsTitle(title) {
    return InStr(title, "Confirm Save As")
    || InStr(title, "Confirmar Salvar")
    || InStr(title, "Confirmar Guardar")
    || InStr(title, "Confirm Replace")
}

FileDialog_IsExcelMultiSheetDialog(hwnd) {
    if !hwnd
        return false
    title := ""
    try title := WinGetTitle("ahk_id " hwnd)
    catch {
        return false
    }
    if (title != "Microsoft Excel")
        return false
    ; Prefer message-text check when available; title alone is enough for Excel's modal.
    try {
        txt := WinGetText("ahk_id " hwnd)
        if txt {
            lower := StrLower(txt)
            if InStr(lower, "multiple sheets")
            || InStr(lower, "várias planilhas")
            || InStr(lower, "varias planilhas")
            || InStr(lower, "múltiplas planilhas")
            || InStr(lower, "multiplas planilhas")
            || InStr(lower, "does not support")
            || InStr(lower, "não oferece suporte")
            || InStr(lower, "nao oferece suporte")
                return true
            ; Title matched but text didn't: still treat as Excel warning after CSV save.
        }
    } catch {
    }
    return true
}

FileDialog_SelectCsvUtf8(root) {
    typeCombo := root.FindFirst({ Type: "ComboBox", AutomationId: "FileTypeControlHost" })
    if !typeCombo
        return false

    curVal := ""
    try curVal := typeCombo.Value
    catch {
    }
    if InStr(curVal, "CSV UTF-8", false)
        return true

    try {
        try typeCombo.ExpandCollapsePattern.Expand()
        catch {
            dropBtn := typeCombo.FindFirst({ Type: "Button", AutomationId: "DropDown" })
            if dropBtn
                FileDialog_InvokeButton(dropBtn)
        }
        Sleep 150
        item := root.FindFirst({ Type: "ListItem", Name: "CSV UTF-8", matchmode: "Substring" })
        if !item
            item := typeCombo.FindFirst({ Type: "ListItem", Name: "CSV UTF-8", matchmode: "Substring" })
        if item {
            try item.SelectionItemPattern.Select()
            catch {
                try item.Click()
                catch {
                }
            }
            Sleep 100
            try curVal := typeCombo.Value
            catch {
            }
            if InStr(curVal, "CSV UTF-8", false)
                return true
        }
    } catch {
    }

    ; Fallback: Alt+T focuses Save as type; type filter text.
    Send "!t"
    Sleep 80
    Send "csv utf-8"
    Sleep 80
    Send "{Enter}"
    Sleep 120
    try curVal := typeCombo.Value
    catch {
    }
    return InStr(curVal, "CSV UTF-8", false)
}

FileDialog_ClickSaveButton(root) {
    actionBtn := root.FindFirst({ Type: "Button", AutomationId: "1" })
    if !actionBtn {
        for name in ["Save", "Salvar", "Guardar", "OK"] {
            actionBtn := root.FindFirst({ Type: "Button", Name: name })
            if actionBtn
                break
        }
    }
    if actionBtn && FileDialog_InvokeButton(actionBtn)
        return true
    Send "!s"
    return true
}

FileDialog_HandlePostSaveDialogs() {
    handledConfirm := false
    handledExcel := false
    deadline := A_TickCount + 8000
    lastActionTick := A_TickCount

    prevMatchMode := A_TitleMatchMode
    SetTitleMatchMode 2
    try {
        while (A_TickCount < deadline) {
            if handledConfirm && handledExcel
                break

            hwnd := WinExist("ahk_class #32770")
            if hwnd {
                title := ""
                try title := WinGetTitle("ahk_id " hwnd)
                catch {
                }

                if !handledConfirm && FileDialog_IsConfirmSaveAsTitle(title) {
                    try WinActivate("ahk_id " hwnd)
                    catch {
                    }
                    Sleep 200
                    FileDialog_ClickAffirmative(hwnd)
                    handledConfirm := true
                    lastActionTick := A_TickCount
                    Sleep 250
                    continue
                }

                if !handledExcel && FileDialog_IsExcelMultiSheetDialog(hwnd) {
                    try WinActivate("ahk_id " hwnd)
                    catch {
                    }
                    Sleep 200
                    FileDialog_ClickAffirmative(hwnd)
                    handledExcel := true
                    lastActionTick := A_TickCount
                    Sleep 250
                    continue
                }
            }

            ; Also catch Excel warning by exact title if class differs.
            if !handledExcel {
                excelHwnd := WinExist("Microsoft Excel ahk_class #32770")
                if !excelHwnd
                    excelHwnd := WinExist("Microsoft Excel")
                if excelHwnd && FileDialog_IsExcelMultiSheetDialog(excelHwnd) {
                    try WinActivate("ahk_id " excelHwnd)
                    catch {
                    }
                    Sleep 200
                    FileDialog_ClickAffirmative(excelHwnd)
                    handledExcel := true
                    lastActionTick := A_TickCount
                    Sleep 250
                    continue
                }
            }

            ; Early exit once secondary dialogs are gone (or never appeared).
            ; Order is Confirm (optional) then Excel; do not bail before Excel has had time to show.
            quietMs := A_TickCount - lastActionTick
            stillDialog := WinExist("ahk_class #32770")

            if handledExcel {
                ; Excel warning is last in the chain.
                if (quietMs >= 400)
                    break
            } else if !stillDialog {
                if handledConfirm && (quietMs >= 2000)
                    break  ; overwrite confirmed; Excel never showed (single-sheet book)
                if !handledConfirm && (quietMs >= 3500)
                    break  ; no secondary dialogs at all
            } else if (quietMs >= 1500) && (handledConfirm || handledExcel) {
                titleLeft := ""
                try titleLeft := WinGetTitle("ahk_id " stillDialog)
                catch {
                }
                if !FileDialog_IsConfirmSaveAsTitle(titleLeft) && !FileDialog_IsExcelMultiSheetDialog(stillDialog)
                    break
            }

            Sleep 200
        }
    } finally {
        SetTitleMatchMode prevMatchMode
    }
}

FileDialog_SaveAsCsvUtf8() {
    hwnd := WinExist("A")
    if !hwnd
        return
    try {
        root := UIA.ElementFromHandle(hwnd)
        typeCombo := root.FindFirst({ Type: "ComboBox", AutomationId: "FileTypeControlHost" })
        if !typeCombo
            return  ; Open dialog / no Save as type — no-op

        if !FileDialog_SelectCsvUtf8(root)
            return

        Sleep 80
        FileDialog_ClickSaveButton(root)
        Sleep 200
        FileDialog_HandlePostSaveDialogs()
    } catch {
    }
}

; --- File Dialog helpers (Import CSV → Load → close Queries) -------------------

FileDialog_ClickOpenButton(root) {
    actionBtn := 0
    try actionBtn := root.FindFirst({ Type: "SplitButton", AutomationId: "1" })
    catch {
        actionBtn := 0
    }
    if !actionBtn {
        try actionBtn := root.FindFirst({ Type: "Button", AutomationId: "1" })
        catch {
            actionBtn := 0
        }
    }
    if !actionBtn {
        for name in ["Open", "Abrir"] {
            try actionBtn := root.FindFirst({ Type: "SplitButton", Name: name })
            catch {
                actionBtn := 0
            }
            if !actionBtn {
                try actionBtn := root.FindFirst({ Type: "Button", Name: name })
                catch {
                    actionBtn := 0
                }
            }
            if actionBtn
                break
        }
    }
    if actionBtn && FileDialog_InvokeButton(actionBtn)
        return true
    Send "!o"
    return true
}

FileDialog_FindLoadButton(root) {
    for name in ["Load", "Carregar"] {
        try {
            btn := root.FindFirst({ Type: "Button", Name: name })
            if btn
                return btn
        } catch {
        }
    }
    return 0
}

FileDialog_WaitAndClickLoad(timeoutMs := 15000) {
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        ; Prefer ImportFromTextDialog AutomationId (Excel text import preview).
        try {
            excelHwnd := WinExist("ahk_exe EXCEL.EXE")
            if excelHwnd {
                excelRoot := UIA.ElementFromHandle(excelHwnd)
                importDlg := 0
                try importDlg := excelRoot.FindFirst({ AutomationId: "ImportFromTextDialog" })
                catch {
                    importDlg := 0
                }
                if importDlg {
                    loadBtn := FileDialog_FindLoadButton(importDlg)
                    if loadBtn && FileDialog_InvokeButton(loadBtn)
                        return true
                }
                ; Fallback: Load/Carregar anywhere under Excel window.
                loadBtn := FileDialog_FindLoadButton(excelRoot)
                if loadBtn && FileDialog_InvokeButton(loadBtn)
                    return true
            }
        } catch {
        }

        ; Also try active window if focus moved to the import dialog.
        try {
            activeHwnd := WinExist("A")
            if activeHwnd {
                activeRoot := UIA.ElementFromHandle(activeHwnd)
                loadBtn := FileDialog_FindLoadButton(activeRoot)
                if loadBtn && FileDialog_InvokeButton(loadBtn)
                    return true
            }
        } catch {
        }

        Sleep 250
    }
    return false
}

FileDialog_CloseQueriesPaneViaCom() {
    try {
        xl := ComObjActive("Excel.Application")
    } catch {
        return false
    }
    for barName in ["Queries and Connections", "Consultas e Conexões"] {
        try {
            bar := xl.CommandBars(barName)
            if bar {
                bar.Visible := false
                return true
            }
        } catch {
        }
    }
    return false
}

FileDialog_CloseQueriesPaneViaUia() {
    excelHwnd := WinExist("ahk_exe EXCEL.EXE")
    if !excelHwnd
        return false
    try {
        root := UIA.ElementFromHandle(excelHwnd)
        optionsBtn := 0
        for name in ["Task Pane Options", "Opções do Painel de Tarefas", "Opcoes do Painel de Tarefas"] {
            try optionsBtn := root.FindFirst({ Type: "MenuItem", Name: name })
            catch {
                optionsBtn := 0
            }
            if !optionsBtn {
                try optionsBtn := root.FindFirst({ Type: "MenuItem", Name: name, matchmode: "Substring" })
                catch {
                    optionsBtn := 0
                }
            }
            if optionsBtn
                break
        }
        if !optionsBtn
            return false
        if !FileDialog_InvokeButton(optionsBtn)
            return false
        Sleep 200
        ; Menu may be under desktop / foreground after expand.
        menuRoot := root
        try {
            desktop := UIA.GetRootElement()
            if desktop
                menuRoot := desktop
        } catch {
        }
        for name in ["Close", "Fechar"] {
            closeItem := 0
            try closeItem := menuRoot.FindFirst({ Type: "MenuItem", Name: name })
            catch {
                closeItem := 0
            }
            if !closeItem {
                try closeItem := menuRoot.FindFirst({ Type: "Button", Name: name })
                catch {
                    closeItem := 0
                }
            }
            if closeItem && FileDialog_InvokeButton(closeItem)
                return true
        }
    } catch {
    }
    return false
}

FileDialog_QueriesPaneVisible(root) {
    try {
        pane := root.FindFirst({ Type: "Pane", Name: "Queries & Connections", matchmode: "Substring" })
        if pane
            return true
    } catch {
    }
    try {
        pane := root.FindFirst({ Type: "Pane", Name: "Consultas", matchmode: "Substring" })
        if pane
            return true
    } catch {
    }
    try {
        custom := root.FindFirst({ Name: "Queries & Connections", matchmode: "Substring" })
        if custom
            return true
    } catch {
    }
    try {
        custom := root.FindFirst({ Name: "Consultas e Conexões", matchmode: "Substring" })
        if custom
            return true
    } catch {
    }
    return false
}

FileDialog_CloseQueriesPane(timeoutMs := 10000) {
    deadline := A_TickCount + timeoutMs
    sawPane := false
    while (A_TickCount < deadline) {
        excelHwnd := WinExist("ahk_exe EXCEL.EXE")
        if excelHwnd {
            try {
                root := UIA.ElementFromHandle(excelHwnd)
                if FileDialog_QueriesPaneVisible(root) {
                    sawPane := true
                    if FileDialog_CloseQueriesPaneViaCom()
                        return true
                    if FileDialog_CloseQueriesPaneViaUia()
                        return true
                } else if sawPane {
                    return true  ; pane disappeared after a close attempt
                }
            } catch {
            }
        }
        Sleep 250
    }
    ; Final COM attempt even if UIA never saw the pane name.
    if FileDialog_CloseQueriesPaneViaCom()
        return true
    return false
}

FileDialog_ImportCsvLoad() {
    hwnd := WinExist("A")
    if !hwnd
        return
    try {
        root := UIA.ElementFromHandle(hwnd)
        ; Import dialog uses Files of type (1136), not Save-as FileTypeControlHost.
        ; Still allow Open on any file dialog with an Open/Save primary button.
        FileDialog_ClickOpenButton(root)
        Sleep 300
        if !FileDialog_WaitAndClickLoad()
            return
        Sleep 400
        FileDialog_CloseQueriesPane()
    } catch {
    }
}
