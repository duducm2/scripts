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

#HotIf
