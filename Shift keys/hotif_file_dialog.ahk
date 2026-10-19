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
    FileDialog_FocusAddressBar()
}

; Shift + D : Navigate to Desktop - Desktop
+d:: {
    if !IsFileDialogActive()
        return
    FileDialog_NavigateToDesktop()
}

FileDialog_FocusAddressBar() {
    try {
        root := UIA.ElementFromHandle(WinExist("A"))
        addressBar := UiElements_TrySaved(root, "FileDialog", "Address")
        if !addressBar
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
            Send "^a"
            return true
        }
    } catch Error {
    }
    Send "!d"
    Sleep 50
    return false
}

FileDialog_FindFileNameEdit(root := 0) {
    if !root {
        try root := UIA.ElementFromHandle(WinExist("A"))
        catch {
            return 0
        }
    }
    saved := UiElements_TrySaved(root, "FileDialog", "FileName")
    if (saved)
        return saved
    try {
        fileNameEdit := root.FindFirst({ Type: "Edit", AutomationId: "1148" })
        if !fileNameEdit {
            possibleNames := [
                "File name:",
                "Nome:",
                "Filename:",
                "File Name:",
                "Name:",
                "Nome do arquivo:",
                "Save As:",
                "Salvar como:"
            ]
            for name in possibleNames {
                fileNameEdit := root.FindFirst({ Type: "Edit", Name: name })
                if fileNameEdit
                    break
            }
        }
        if !fileNameEdit {
            fileNameCombo := root.FindFirst({ Type: "ComboBox", AutomationId: "1148" })
            if fileNameCombo {
                fileNameEdit := fileNameCombo.FindFirst({ Type: "Edit" })
            }
        }
        return fileNameEdit ? fileNameEdit : 0
    } catch {
        return 0
    }
}

FileDialog_FocusFileNameField() {
    try {
        fileNameEdit := FileDialog_FindFileNameEdit()
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

FileDialog_NavigateToFolder(path) {
    if (path = "")
        return false
    FileDialog_FocusAddressBar()
    Sleep 50
    SendText(path)
    Send "{Enter}"
    Sleep 200
    FileDialog_FocusFileNameField()
    return true
}

FileDialog_NavigateToDesktop() {
    FileDialog_NavigateToFolder(A_Desktop)
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
    FileDialog_FocusFileNameField()
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

        ; First attempt: saved element, then AutomationId and Type
        cancelBtn := UiElements_TrySaved(root, "FileDialog", "Cancel")
        if !cancelBtn
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

; Shift + I : Import CSV — Open → Load (or PQ Open As CSV → Close & Load) → close Queries → format
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
        item := UiElements_TrySaved(root, "FileDialog", "CsvUtf8")
        if !item
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

; True only for real PDF save types (e.g. "PDF (_.pdf)"), not BMP / Bitmap.
FileDialog_IsPdfSaveType(val) {
    if (val = "")
        return false
    if InStr(val, "BMP", false) || InStr(val, "Bitmap", false)
    || InStr(val, "Mapa de bits", false)
        return false
    ; Require PDF label plus .pdf extension hint (Office shows PDF (_.pdf) / PDF (*.pdf)).
    return InStr(val, "PDF", false) && InStr(val, ".pdf", false)
}

FileDialog_FindPdfTypeListItem(scope) {
    if !scope
        return 0
    saved := UiElements_TrySaved(scope, "FileDialog", "PdfType")
    if (saved)
        return saved
    try {
        for item in scope.FindAll({ Type: "ListItem" }) {
            name := ""
            try name := item.Name
            catch {
                continue
            }
            if FileDialog_IsPdfSaveType(name) || RegExMatch(name, "i)^PDF\b")
                return item
        }
    } catch {
    }
    return 0
}

FileDialog_TrySelectPdfListItem(typeCombo, root) {
    try {
        try typeCombo.ExpandCollapsePattern.Expand()
        catch {
            dropBtn := typeCombo.FindFirst({ Type: "Button", AutomationId: "DropDown" })
            if dropBtn
                FileDialog_InvokeButton(dropBtn)
        }
        Sleep 150
        item := FileDialog_FindPdfTypeListItem(typeCombo)
        if !item
            item := FileDialog_FindPdfTypeListItem(root)
        if !item
            return false
        try item.SelectionItemPattern.Select()
        catch {
            try item.Click()
            catch {
                return false
            }
        }
        Sleep 100
        return true
    } catch {
    }
    return false
}

FileDialog_SelectPdf(root) {
    typeCombo := root.FindFirst({ Type: "ComboBox", AutomationId: "FileTypeControlHost" })
    if !typeCombo
        return false

    curVal := ""
    try curVal := typeCombo.Value
    catch {
    }
    if FileDialog_IsPdfSaveType(curVal)
        return true

    FileDialog_TrySelectPdfListItem(typeCombo, root)
    try curVal := typeCombo.Value
    catch {
    }
    if FileDialog_IsPdfSaveType(curVal)
        return true

    ; Fallback: Alt+T → type "PDF" prefix (not "pdf"); reject BMP via IsPdfSaveType.
    Send "!t"
    Sleep 80
    Send "^a"
    Sleep 40
    SendText("PDF")
    Sleep 80
    Send "{Enter}"
    Sleep 120
    try curVal := typeCombo.Value
    catch {
    }
    if FileDialog_IsPdfSaveType(curVal)
        return true

    ; Last try: expand and pick the PDF list item again after keyboard filter.
    FileDialog_TrySelectPdfListItem(typeCombo, root)
    try curVal := typeCombo.Value
    catch {
    }
    return FileDialog_IsPdfSaveType(curVal)
}

FileDialog_ClickSaveButton(root) {
    actionBtn := UiElements_TrySaved(root, "FileDialog", "Save")
    if !actionBtn
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

; Strip quotes and collapse ".csv.csv" so Save as type CSV cannot produce
; `"path\file.csv".csv` (Windows appends .csv when quotes hide the real extension).
FileDialog_NormalizeCsvSaveName(raw) {
    name := Trim(raw)
    if (name = "")
        return ""
    name := StrReplace(name, '"')
    name := Trim(name)
    while RegExMatch(name, "i)\.csv\.csv$")
        name := RegExReplace(name, "i)\.csv$", "")
    return name
}

FileDialog_ReadFileNameField(root) {
    el := FileDialog_FindFileNameEdit(root)
    if !el
        return ""
    val := ""
    try val := el.Value
    catch {
    }
    return Trim(val)
}

FileDialog_WriteFileNameField(root, name) {
    if (name = "")
        return false
    el := FileDialog_FindFileNameEdit(root)
    if el {
        try {
            el.Value := name
            return true
        } catch {
        }
    }
    FileDialog_FocusFileNameField()
    Sleep 50
    SendText(name)
    return true
}

; After CSV UTF-8 is selected, rewrite the file name if quoted or double-extended.
FileDialog_EnsureCsvSaveNameIsValid(root) {
    raw := FileDialog_ReadFileNameField(root)
    if (raw = "")
        return true
    clean := FileDialog_NormalizeCsvSaveName(raw)
    if (clean = "" || clean = raw)
        return true
    return FileDialog_WriteFileNameField(root, clean)
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
        FileDialog_EnsureCsvSaveNameIsValid(root)
        Sleep 50
        FileDialog_ClickSaveButton(root)
        Sleep 200
        FileDialog_HandlePostSaveDialogs()
    } catch {
    }
}

; --- File Dialog helpers (Import CSV → Load → close Queries → format → shade) -
; Some Excel builds skip ImportFromTextDialog and open Power Query Editor on
; binary File.Contents (Convert / Open As). Handle that path too.

FileDialog_ClickOpenButton(root) {
    actionBtn := UiElements_TrySaved(root, "FileDialog", "Open")
    if !actionBtn {
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
    }
    if actionBtn && FileDialog_InvokeButton(actionBtn)
        return true
    Send "!o"
    return true
}

FileDialog_FindLoadButton(root) {
    saved := UiElements_TrySaved(root, "FileDialog", "Load")
    if (saved)
        return saved
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

FileDialog_FindPowerQueryEditor(excelRoot := 0) {
    ; Prefer AutomationId EditorDialog (stable across locales).
    scopes := []
    if excelRoot
        scopes.Push(excelRoot)
    try {
        excelHwnd := WinExist("ahk_exe EXCEL.EXE")
        if excelHwnd {
            er := UIA.ElementFromHandle(excelHwnd)
            if er
                scopes.Push(er)
        }
    } catch {
    }
    try {
        activeHwnd := WinExist("A")
        if activeHwnd {
            ar := UIA.ElementFromHandle(activeHwnd)
            if ar
                scopes.Push(ar)
        }
    } catch {
    }
    for scope in scopes {
        if !scope
            continue
        try {
            pq := scope.FindFirst({ AutomationId: "EditorDialog" })
            if pq
                return pq
        } catch {
        }
        try {
            pq := scope.FindFirst({ Type: "Window", Name: "Power Query Editor", matchmode: "Substring" })
            if pq
                return pq
        } catch {
        }
        try {
            pq := scope.FindFirst({ Type: "Window", Name: "Editor do Power Query", matchmode: "Substring" })
            if pq
                return pq
        } catch {
        }
    }
    return 0
}

FileDialog_PqFindFormulaEdit(pqRoot) {
    if !pqRoot
        return 0
    ; Prefer the formula-bar edit (avoid FindAll over the IE/PQ web pane — it can hang).
    for needle in ["Editor content", "conteúdo do Editor", "conteudo do Editor"] {
        try {
            edit := pqRoot.FindFirst({ Type: "Edit", Name: needle, matchmode: "Substring" })
            if edit
                return edit
        } catch {
        }
    }
    return 0
}

FileDialog_PqGetFormulaValue(pqRoot) {
    edit := FileDialog_PqFindFormulaEdit(pqRoot)
    if !edit
        return ""
    try {
        return edit.Value
    } catch {
    }
    return ""
}

FileDialog_PqIsBinaryContentsView(pqRoot) {
    if !pqRoot
        return false
    val := FileDialog_PqGetFormulaValue(pqRoot)
    ; Still raw binary if File.Contents is present and not already wrapped in Csv.Document.
    if val != "" {
        if InStr(val, "File.Contents", false) && !InStr(val, "Csv.Document", false)
            return true
        if InStr(val, "Csv.Document", false)
            return false
    }
    ; Binary preview tile before Open As / double-click.
    for needle in ["Double click to open", "Clique duas vezes"] {
        try {
            dbl := pqRoot.FindFirst({ Name: needle, matchmode: "Substring" })
            if dbl
                return true
        } catch {
        }
    }
    return false
}

FileDialog_PqActivate(pqRoot) {
    if !pqRoot
        return
    try {
        winId := pqRoot.WinId
        if winId
            WinActivate("ahk_id " winId)
        else if WinExist("ahk_exe EXCEL.EXE")
            WinActivate("ahk_exe EXCEL.EXE")
    } catch {
        try {
            if WinExist("ahk_exe EXCEL.EXE")
                WinActivate("ahk_exe EXCEL.EXE")
        } catch {
        }
    }
}

; Most reliable for Convert/binary view: wrap File.Contents(...) as Csv.Document(...).
FileDialog_PqWrapFileContentsAsCsv(pqRoot) {
    if !pqRoot
        return false
    edit := FileDialog_PqFindFormulaEdit(pqRoot)
    if !edit
        return false
    val := ""
    try val := edit.Value
    catch {
        return false
    }
    if (val = "") || !InStr(val, "File.Contents", false)
        return false
    if InStr(val, "Csv.Document", false)
        return true

    ; "= File.Contents(...)" → "= Csv.Document(File.Contents(...))"
    newVal := val
    if RegExMatch(val, "i)^\s*=\s*File\.Contents\s*\(") {
        newVal := RegExReplace(val, "i)^\s*=\s*", "= Csv.Document(")
        if !RegExMatch(newVal, "\)\s*$")
            newVal .= ")"
        else
            newVal := RegExReplace(newVal, "\)\s*$", "))")
    } else {
        newVal := "= Csv.Document(" . Trim(val, " `t=") . ")"
        if !InStr(newVal, "File.Contents")
            return false
    }

    FileDialog_PqActivate(pqRoot)
    Sleep 100
    try edit.SetFocus()
    catch {
    }
    Sleep 80
    setOk := false
    try {
        edit.Value := newVal
        setOk := true
    } catch {
    }
    if !setOk {
        try {
            edit.ValuePattern.SetValue(newVal)
            setOk := true
        } catch {
        }
    }
    if !setOk {
        ; Keyboard: select all in formula bar and paste wrapped formula.
        clipSaved := ClipboardAll()
        try {
            try edit.Click("left")
            catch {
            }
            Sleep 80
            Send "^a"
            Sleep 40
            A_Clipboard := newVal
            Sleep 40
            Send "^v"
            setOk := true
        } finally {
            Sleep 80
            try A_Clipboard := clipSaved
            catch {
            }
        }
    }
    Sleep 80
    Send "{Enter}"
    Sleep 600
    ; Confirm wrap applied (or at least left binary-only formula).
    val2 := FileDialog_PqGetFormulaValue(pqRoot)
    return InStr(val2, "Csv.Document", false) || (val2 != "" && !InStr(val2, "File.Contents", false))
}

FileDialog_PqClickOpenAsCsv(pqRoot) {
    if !pqRoot
        return false

    ; 1) Formula wrap — works even when Convert ribbon clicks are ignored.
    if FileDialog_PqWrapFileContentsAsCsv(pqRoot)
        return true

    FileDialog_PqActivate(pqRoot)
    Sleep 80

    ; 2) Physical double-click on the binary preview tile / .csv caption.
    for needle in ["Double click to open", "Clique duas vezes"] {
        try {
            tile := pqRoot.FindFirst({ Name: needle, matchmode: "Substring" })
            if tile {
                try {
                    tile.Click("left", 2)
                    Sleep 500
                    if !FileDialog_PqIsBinaryContentsView(pqRoot)
                        return true
                } catch {
                }
            }
        } catch {
        }
    }

    ; 3) Convert ribbon: Open As dropdown → CSV
    openAs := 0
    for name in ["Open As", "Abrir como"] {
        try openAs := pqRoot.FindFirst({ Type: "SplitButton", Name: name, matchmode: "Substring" })
        catch {
            openAs := 0
        }
        if openAs
            break
    }
    if !openAs
        return false

    ; Prefer a real mouse click on the split button (Invoke often no-ops on NetUI/WinForms).
    try openAs.Click("left")
    catch {
        try openAs.Click()
        catch {
            return false
        }
    }
    Sleep 250

    menuRoot := pqRoot
    try {
        desktop := UIA.GetRootElement()
        if desktop
            menuRoot := desktop
    } catch {
    }
    for optName in ["CSV", "Text", "Texto"] {
        item := 0
        for typeName in ["MenuItem", "ListItem", "Button"] {
            try item := menuRoot.FindFirst({ Type: typeName, Name: optName })
            catch {
                item := 0
            }
            if item
                break
        }
        if item {
            try item.Click("left")
            catch {
                if !FileDialog_InvokeButton(item)
                    continue
            }
            Sleep 500
            return !FileDialog_PqIsBinaryContentsView(pqRoot)
        }
    }
    return false
}

FileDialog_FindCloseAndLoad(root) {
    if !root
        return 0
    ; Exact names first (avoid matching "Close & Load To..." when possible).
    for name in ["Close & Load", "Fechar e carregar", "Close and Load"] {
        for typeName in ["SplitButton", "Button", "MenuItem"] {
            try {
                btn := root.FindFirst({ Type: typeName, Name: name })
                if btn
                    return btn
            } catch {
            }
        }
    }
    for name in ["Close & Load", "Fechar e carregar", "Close and Load"] {
        for typeName in ["SplitButton", "Button", "MenuItem"] {
            try {
                btn := root.FindFirst({ Type: typeName, Name: name, matchmode: "Substring" })
                if btn
                    return btn
            } catch {
            }
        }
    }
    return 0
}

FileDialog_PqEnsureHomeTab(pqRoot) {
    if !pqRoot
        return
    try {
        homeTab := pqRoot.FindFirst({ Type: "TabItem", Name: "Home" })
        if !homeTab
            homeTab := pqRoot.FindFirst({ Type: "TabItem", Name: "Início" })
        if !homeTab
            homeTab := pqRoot.FindFirst({ Type: "TabItem", Name: "Inicio" })
        if homeTab
            FileDialog_InvokeButton(homeTab)
    } catch {
    }
}

FileDialog_PqCloseAndLoad(pqRoot, allowKeyboard := false) {
    if !pqRoot
        return false
    FileDialog_PqEnsureHomeTab(pqRoot)
    Sleep 150
    btn := FileDialog_FindCloseAndLoad(pqRoot)
    if btn && FileDialog_InvokeButton(btn)
        return true
    if !allowKeyboard
        return false
    ; Last resort: Home → Close & Load accelerators (locale-dependent).
    try {
        winId := pqRoot.WinId
        if winId
            WinActivate("ahk_id " winId)
        else if WinExist("ahk_exe EXCEL.EXE")
            WinActivate("ahk_exe EXCEL.EXE")
    } catch {
        try {
            if WinExist("ahk_exe EXCEL.EXE")
                WinActivate("ahk_exe EXCEL.EXE")
        } catch {
        }
    }
    Send "!h"
    Sleep 80
    Send "c"
    Sleep 80
    Send "l"
    Sleep 500
    ; Only treat as success if the editor actually closed.
    return !FileDialog_FindPowerQueryEditor()
}

FileDialog_TryClickImportLoad(root) {
    if !root
        return false
    importDlg := 0
    try importDlg := root.FindFirst({ AutomationId: "ImportFromTextDialog" })
    catch {
        importDlg := 0
    }
    if importDlg {
        loadBtn := FileDialog_FindLoadButton(importDlg)
        if loadBtn && FileDialog_InvokeButton(loadBtn)
            return true
    }
    loadBtn := FileDialog_FindLoadButton(root)
    if loadBtn && FileDialog_InvokeButton(loadBtn)
        return true
    return false
}

FileDialog_WaitAndClickLoad(timeoutMs := 25000) {
    deadline := A_TickCount + timeoutMs
    lastPqOpenAsTick := 0
    lastPqCloseKbTick := 0
    while (A_TickCount < deadline) {
        ; If Power Query Editor is up, skip scanning Excel for Load (avoids slow UIA walks).
        pq := 0
        try pq := FileDialog_FindPowerQueryEditor()
        catch {
            pq := 0
        }

        if pq {
            try {
                if FileDialog_PqIsBinaryContentsView(pq) {
                    if (A_TickCount - lastPqOpenAsTick > 2000) {
                        lastPqOpenAsTick := A_TickCount
                        try StandardLoadingBar_Update("⏳ Power Query: open as CSV…", BANNER_ACCENT_INTERMEDIATE)
                        catch {
                        }
                        FileDialog_PqClickOpenAsCsv(pq)
                        Sleep 400
                    }
                } else if FileDialog_PqCloseAndLoad(pq, false) {
                    return true
                } else if (A_TickCount - lastPqCloseKbTick > 2000) {
                    lastPqCloseKbTick := A_TickCount
                    try StandardLoadingBar_Update("⏳ Power Query: Close & Load…", BANNER_ACCENT_INTERMEDIATE)
                    catch {
                    }
                    if FileDialog_PqCloseAndLoad(pq, true)
                        return true
                }
            } catch {
            }
            Sleep 250
            continue
        }

        ; Classic path: ImportFromTextDialog Load / Carregar.
        try {
            excelHwnd := WinExist("ahk_exe EXCEL.EXE")
            if excelHwnd {
                excelRoot := UIA.ElementFromHandle(excelHwnd)
                if FileDialog_TryClickImportLoad(excelRoot)
                    return true
            }
        } catch {
        }
        try {
            activeHwnd := WinExist("A")
            if activeHwnd {
                activeRoot := UIA.ElementFromHandle(activeHwnd)
                if FileDialog_TryClickImportLoad(activeRoot)
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
        closeItem := 0
        try closeItem := UiElements_TrySaved(menuRoot, "EXCEL.EXE", "QueriesClose")
        if (closeItem && FileDialog_InvokeButton(closeItem))
            return true
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
        StandardLoadingBar_Show("⏳ Importing CSV…", BANNER_ACCENT_INTERMEDIATE, { passive: false,
            centerOnHwnd: hwnd, textWidth: 480 })
    } catch {
    }
    try {
        root := UIA.ElementFromHandle(hwnd)
        ; Import dialog uses Files of type (1136), not Save-as FileTypeControlHost.
        try StandardLoadingBar_Update("⏳ Opening file…", BANNER_ACCENT_INTERMEDIATE)
        catch {
        }
        FileDialog_ClickOpenButton(root)
        Sleep 300

        try StandardLoadingBar_Update("⏳ Waiting for Load / Power Query…", BANNER_ACCENT_INTERMEDIATE)
        catch {
        }
        if !FileDialog_WaitAndClickLoad() {
            try StandardLoadingBar_Update("❌ Load / Power Query Close & Load not found", BANNER_ACCENT_ERROR)
            catch {
            }
            try StandardLoadingBar_Hide(800)
            catch {
            }
            return
        }

        try StandardLoadingBar_Update("⏳ Closing Queries pane…", BANNER_ACCENT_INTERMEDIATE)
        catch {
        }
        Sleep 400
        FileDialog_CloseQueriesPane()

        try StandardLoadingBar_Update("✅ CSV imported — fixing headers & columns…", BANNER_ACCENT_SUCCESS)
        catch {
        }
        try StandardLoadingBar_Hide(400)
        catch {
        }
        Sleep 300
        ; Promote Column1/Column2/… → real header row (no-op if already real names)
        Excel_PromoteGenericHeaders()
        ; Autofit, cap widths >20 → 20, wrap text, center H/V
        Excel_FormatImportedTable()
        ; Dark surround: entire sheet gray, imported table fill cleared
        Excel_ShadeOutsideImportedTable()
        ; Same as one Shift+N: scale columns to ActiveWindow.UsableWidth (Balanced first)
        Excel_Layout_CyclePillar()
    } catch {
        try StandardLoadingBar_Update("❌ Import failed", BANNER_ACCENT_ERROR)
        catch {
        }
        try StandardLoadingBar_Hide(800)
        catch {
        }
    }
}
