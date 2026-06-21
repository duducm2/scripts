; =============================================================================
; Shift keys module: hotif_powerbi.ahk
; Power BI hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

#HotIf (WinActive("ahk_exe PBIDesktop.exe") || InStr(WinGetTitle("A"), "powerbi", false)) && !IsFileDialogActive()

; Shift + C : Get data (Click Home tab, then Get data primary button)
+c:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        homeTab := root.FindFirst({ Type: "50019", Name: "Home", AutomationId: "home" })
        if !homeTab {
            homeTab := root.FindFirst({ Type: "50019", Name: "Home" })
        }
        if !homeTab {
            homeTab := root.FindFirst({ Type: "50019", AutomationId: "home" })
        }

        if homeTab {
            homeTab.Click()
            Sleep 200
        } else {
            MsgBox "Could not find the 'Home' tab.", "Power BI", "IconX"
            return
        }

        possibleNames := ["Get data", "Obter dados"]
        getDataBtn := ""

        for , name in possibleNames {
            getDataBtn := root.FindFirst({ Name: name, Type: "50000", ClassName: "splitPrimaryButton", matchmode: "Substring"
            })
            if getDataBtn
                break
            getDataBtn := root.FindFirst({ Name: name, Type: "50000", ClassName: "splitPrimaryButton root-332" })
            if getDataBtn
                break
            getDataBtn := root.FindFirst({ Name: name, Type: "50000", ClassName: "splitPrimaryButton root-320" })
            if getDataBtn
                break
        }

        if !getDataBtn {
            for , name in possibleNames {
                getDataBtn := root.FindFirst({ Name: name, Type: "50000" })
                if getDataBtn
                    break
            }
        }

        if getDataBtn {
            getDataBtn.Click()
        } else {
            MsgBox "Could not find the 'Get data' button.", "Power BI", "IconX"
        }
    } catch Error as e {
        MsgBox "Error triggering Get data: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + T : Transform data (Click Home tab, then T, then UIA click)
+t:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Click the Home tab
        homeTab := root.FindFirst({ Type: "50019", Name: "Home", AutomationId: "home" })
        if !homeTab {
            homeTab := root.FindFirst({ Type: "50019", Name: "Home" })
        }
        if !homeTab {
            homeTab := root.FindFirst({ Type: "50019", AutomationId: "home" })
        }

        if homeTab {
            homeTab.Click()
            Sleep 120
        } else {
            MsgBox "Could not find the 'Home' tab.", "Power BI", "IconX"
            return
        }

        Send "t"
        Sleep 250

        possibleNames := ["Transform data", "Transformar dados"]
        transformBtn := ""

        ; Try to find by Name and Type 50000 (Button)
        for , name in possibleNames {
            transformBtn := root.FindFirst({ Name: name, Type: "50000" })
            if transformBtn
                break
        }

        ; Fallback: try with ClassName
        if !transformBtn {
            transformBtn := root.FindFirst({ Type: "50000", ClassName: "splitPrimaryButton root-332" })
        }

        ; Fallback: try with partial ClassName match
        if !transformBtn {
            transformBtn := root.FindFirst({ Type: "50000", ClassName: "splitPrimaryButton", matchmode: "Substring" })
        }

        ; Fallback: try with partial Name match
        if !transformBtn {
            transformBtn := root.FindFirst({ Name: "Transform", Type: "50000", matchmode: "Substring" })
        }

        if transformBtn {
            transformBtn.Click()
        } else {
            MsgBox "Could not find the 'Transform data' menu item.", "Power BI", "IconX"
        }
    } catch Error as e {
        MsgBox "Error triggering Transform data: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + U : Close and apply (Alt, H, C, C)
+u:: {
    SendEscape()
    SendEscape()
    Send "{Alt down}"
    Send "{Alt down}"
    Sleep 200
    Sleep 200
    Send "{Alt up}"
    Send "h"
    Sleep 100
    Send "c"
    Sleep 100
    Send "c"
}

; Shift + I : Report view
+i:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Find the Report view tab by name only
        reportTab := root.FindFirst({ Name: "Report view" })
        if !reportTab {
            reportTab := root.FindFirst({ Name: "Report view", matchmode: "Substring" })
        }

        if reportTab {
            reportTab.Click()
        } else {
            MsgBox "Could not find the 'Report view' tab.", "Power BI", "IconX"
        }
    } catch Error as e {
        MsgBox "Error switching to Report view: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + O : Table view
+o:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Find the Table view tab by name only
        tableTab := root.FindFirst({ Name: "Table view" })
        if !tableTab {
            tableTab := root.FindFirst({ Name: "Table view", matchmode: "Substring" })
        }

        if tableTab {
            tableTab.Click()
        } else {
            MsgBox "Could not find the 'Table view' tab.", "Power BI", "IconX"
        }
    } catch Error as e {
        MsgBox "Error switching to Table view: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + P : Model view
+p:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Find the Model view tab by name only
        modelTab := root.FindFirst({ Name: "Model view" })
        if !modelTab {
            modelTab := root.FindFirst({ Name: "Model view", matchmode: "Substring" })
        }

        if modelTab {
            modelTab.Click()
        } else {
            MsgBox "Could not find the 'Model view' tab.", "Power BI", "IconX"
        }
    } catch Error as e {
        MsgBox "Error switching to Model view: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + H : Build visual
+h:: {
    try {
        win := WinExist("A")
        if !win
            return
        root := UIA.ElementFromHandle(win)

        possibleNames := [
            "Build visual",
            "Build visuals",
            "Build visualization",
            "Build pane",
            "Visualizar",
            "Criar visual",
            "Criar visualização",
            "Construir visual",
            "Construir visualização"
        ]

        buildTab := ""

        for name in possibleNames {
            buildTab := root.FindFirst({ Type: "50019", Name: name })
            if buildTab
                break
            buildTab := root.FindFirst({ Type: "TabItem", Name: name })
            if buildTab
                break
            buildTab := root.FindFirst({ Type: "50019", Name: name, matchmode: "Substring" })
            if buildTab
                break
            buildTab := root.FindFirst({ Type: "TabItem", Name: name, matchmode: "Substring" })
            if buildTab
                break
        }

        if !buildTab {
            tabCond := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.TabItem)
            tabs := ""
            try tabs := root.FindElements(tabCond, UIA.TreeScope.Descendants)
            if tabs {
                for tab in tabs {
                    if !tab
                        continue
                    tabName := tab.Name
                    for name in possibleNames {
                        if InStr(tabName, name) {
                            buildTab := tab
                            break
                        }
                    }
                    if buildTab
                        break
                }
            }
        }

        if buildTab {
            buildTab.Click()
        } else {
            MsgBox "Could not find the 'Build visual' tab.", "Power BI", "IconX"
        }
    } catch Error as e {
        MsgBox "Error switching to Build visual: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + J : Format visual
+j:: {
    try {
        win := WinExist("A")
        if !win
            return
        root := UIA.ElementFromHandle(win)

        formatTab := ""

        ; Try "Format page" first (this is the working solution and is fast)
        try {
            formatTab := root.FindFirst({ Type: "50019", Name: "Format page" })
        } catch {
            try {
                formatTab := root.FindFirst({ Type: "TabItem", Name: "Format page" })
            } catch {
                ; Continue to fallback searches
            }
        }

        ; If "Format page" not found, try original names (simplified - only most common)
        if !formatTab {
            possibleNames := ["Format visual", "Format visuals", "Formatting"]
            for name in possibleNames {
                try {
                    formatTab := root.FindFirst({ Type: "50019", Name: name })
                    if formatTab
                        break
                } catch {
                    try {
                        formatTab := root.FindFirst({ Type: "TabItem", Name: name })
                        if formatTab
                            break
                    } catch {
                        ; Continue to next name
                    }
                }
            }
        }

        ; Final fallback: use FindElements to search all tabs
        if !formatTab {
            try {
                tabCond := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.TabItem)
                tabs := root.FindElements(tabCond, UIA.TreeScope.Descendants)
                if tabs {
                    for tab in tabs {
                        if !tab
                            continue
                        tabName := tab.Name
                        if InStr(tabName, "Format page") || InStr(tabName, "Format visual") || InStr(tabName,
                            "Formatting") {
                            formatTab := tab
                            break
                        }
                    }
                }
            } catch {
                ; Ignore errors
            }
        }

        if formatTab {
            formatTab.Click()
        } else {
            MsgBox "Could not find the 'Format visual' tab.", "Power BI", "IconX"
        }
    } catch Error as e {
        MsgBox "Error switching to Format visual: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + S : Select the Power BI search edit field (Data anchor + Tab)
+s:: {
    try {
        win := WinExist("A")
        if !win
            return

        root := UIA.ElementFromHandle(win)

        dataBtn := ""
        try {
            for cfg in PowerBI_GetDrawerConfigs() {
                if (cfg.HasOwnProp("label") && cfg.label = "Data") {
                    dataBtn := PowerBI_FindDrawerButton(root, cfg)
                    if dataBtn
                        break
                }
            }
        }

        if !dataBtn {
            MsgBox "Could not locate the Data button anchor.", "Power BI", "IconX"
            return
        }

        focused := false
        try {
            dataBtn.SetFocus()
            focused := true
        } catch {
            try {
                dataBtn.Select()
                focused := true
            } catch {
            }
        }

        if !focused {
            MsgBox "Could not focus the Data button anchor.", "Power BI", "IconX"
            return
        }

        Sleep 120
        Send "{Tab}"
        Sleep 120
        Send "^a"
    } catch Error as e {
        MsgBox "Error selecting the Power BI search field: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + L : Click OK/Confirm button in Power BI modals
+l:: {
    try {
        root := UIA.ElementFromHandle(WinExist("A"))

        ; Try by name first (since we know it's "OK" with Type 50000)
        possibleNames := [
            ; English variations
            "OK",
            "Confirm",
            "Accept",
            "Apply",
            "Done"
            "Yes",
            "Continue",
            "Proceed",
            "Save",
            "Finish",
            ; Portuguese variations
            "Confirmar",
            "Aceitar",
            "Aplicar",
            "Sim",
            "Continuar",
            "Prosseguir",
            "Salvar",
            "Finalizar",
            ; Spanish variations
            "Aceptar",
            "Continuar",
            "Guardar",
            "Finalizar",
            ; French variations
            "Confirmer",
            "Accepter",
            "Continuer",
            "Enregistrer",
            ; German variations
            "Bestätigen",
            "Akzeptieren",
            "Fortfahren",
            "Speichern"
        ]

        ; First attempt: Try by name with Button type (numeric 50000 or string "Button")
        for name in possibleNames {
            confirmBtn := root.FindFirst({ Type: "Button", Name: name })
            if !confirmBtn {
                ; Try with numeric type code
                confirmBtn := root.FindFirst({ Type: 50000, Name: name })
            }
            if confirmBtn
                break
        }

        ; Second attempt: Find by AutomationId and Type
        if !confirmBtn {
            confirmBtn := root.FindFirst({ Type: "Button", AutomationId: "1" })
            if !confirmBtn {
                confirmBtn := root.FindFirst({ Type: 50000, AutomationId: "1" })
            }
        }

        ; Third attempt: Try SplitButton type (some dialogs use this instead)
        if !confirmBtn {
            for name in possibleNames {
                confirmBtn := root.FindFirst({ Type: "SplitButton", Name: name })
                if confirmBtn
                    break
            }
            if !confirmBtn {
                confirmBtn := root.FindFirst({ Type: "SplitButton", AutomationId: "1" })
            }
        }

        ; Fourth attempt: Search all buttons and find by name (more thorough)
        if !confirmBtn {
            allButtons := root.FindAll({ Type: "Button" })
            for btn in allButtons {
                btnName := btn.Name
                for name in possibleNames {
                    if (btnName = name) {
                        confirmBtn := btn
                        break
                    }
                }
                if confirmBtn
                    break
            }
        }

        if confirmBtn {
            confirmBtn.Click()
            return
        }
    } catch Error {
    }
    ; Fallback: Try common keyboard shortcuts
    Send "{Enter}"  ; Enter key is universal for OK/Confirm
}

; Shift + X : Click Cancel/Exit button in Power BI modals
+x:: {
    try {
        root := UIA.ElementFromHandle(WinExist("A"))

        ; First attempt: Find by AutomationId and Type (most reliable)
        cancelBtn := root.FindFirst({ Type: "Button", AutomationId: "2" })

        ; Second attempt: Try various possible names for Cancel/Exit
        if !cancelBtn {
            possibleNames := [
                ; English variations
                "Cancel",
                "Close",
                "Exit",
                "Dismiss",
                "No",
                "Abort",
                "Back",
                "Close",
                ; Portuguese variations
                "Cancelar",
                "Fechar",
                "Sair",
                "Descartar",
                "Não",
                "Voltar",
                ; Spanish variations
                "Cancelar",
                "Cerrar",
                "Salir",
                "Descartar",
                "No",
                ; French variations
                "Annuler",
                "Fermer",
                "Quitter",
                "Ignorer",
                "Non",
                ; German variations
                "Abbrechen",
                "Schließen",
                "Verlassen",
                "Abweisen",
                "Nein"
            ]
            for name in possibleNames {
                cancelBtn := root.FindFirst({ Type: "Button", Name: name })
                if cancelBtn
                    break
            }
        }

        ; Third attempt: Try SplitButton type (some dialogs use this instead)
        if !cancelBtn {
            cancelBtn := root.FindFirst({ Type: "SplitButton", AutomationId: "2" })
            if !cancelBtn {
                for name in possibleNames {
                    cancelBtn := root.FindFirst({ Type: "SplitButton", Name: name })
                    if cancelBtn
                        break
                }
            }
        }

        if cancelBtn {
            cancelBtn.Click()
            return
        }
    } catch Error {
    }
    ; Fallback: Try common keyboard shortcuts
    SendEscape()  ; Escape key is universal for cancels
}

; Shift + A : Right-click All pages button in Power BI
+a:: {
    try {
        root := UIA.ElementFromHandle(WinExist("A"))

        ; First attempt: Find by Name and Type
        prevPageBtn := root.FindFirst({ Type: "Button", Name: "Previous pages" })
        if !prevPageBtn {
            prevPageBtn := root.FindFirst({ Type: 50000, Name: "Previous pages" })
        }

        ; Second attempt: Find by ClassName
        if !prevPageBtn {
            prevPageBtn := root.FindFirst({ Type: "Button", ClassName: "carouselNavButton previousPage" })
            if !prevPageBtn {
                prevPageBtn := root.FindFirst({ Type: 50000, ClassName: "carouselNavButton previousPage" })
            }
        }

        ; Third attempt: Find by partial ClassName match
        if !prevPageBtn {
            allButtons := root.FindAll({ Type: "Button" })
            for btn in allButtons {
                btnClassName := btn.ClassName
                if InStr(btnClassName, "previousPage") {
                    prevPageBtn := btn
                    break
                }
            }
        }

        if prevPageBtn {
            ; Get button location and instantly move cursor to that position
            btnPos := prevPageBtn.Location
            x := btnPos.x + btnPos.w // 2
            y := btnPos.y + btnPos.h // 2

            ; Instantly set cursor position (no visible movement)
            DllCall("SetCursorPos", "Int", x, "Int", y)

            ; Perform right-click immediately
            saveCoordMode := A_CoordModeMouse
            CoordMode("Mouse", "Screen")
            Click(x " " y " Right")
            CoordMode("Mouse", saveCoordMode)
            return
        }
    } catch Error {
    }
}

; Shift + W : Click New page button
+w:: {
    try {
        root := UIA.ElementFromHandle(WinExist("A"))

        ; Find by Name
        newPageBtn := root.FindFirst({ Type: "Button", Name: "New page" })
        if !newPageBtn {
            newPageBtn := root.FindFirst({ Type: 50000, Name: "New page" })
        }

        ; Find by ClassName
        if !newPageBtn {
            newPageBtn := root.FindFirst({ Type: "Button", ClassName: "section static create" })
            if !newPageBtn {
                newPageBtn := root.FindFirst({ Type: 50000, ClassName: "section static create" })
            }
        }

        ; Find by partial ClassName match
        if !newPageBtn {
            allButtons := root.FindAll({ Type: "Button" })
            for btn in allButtons {
                if (btn.Name = "New page") {
                    newPageBtn := btn
                    break
                }
            }
        }

        if newPageBtn {
            newPageBtn.Click()
            return
        }
    } catch Error {
    }
}

; Shift + E : New measure (Click Home tab, then New measure button)
+e:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Click the Home tab
        homeTab := root.FindFirst({ Type: "50019", Name: "Home", AutomationId: "home" })
        if !homeTab {
            homeTab := root.FindFirst({ Type: "50019", Name: "Home" })
        }
        if !homeTab {
            homeTab := root.FindFirst({ Type: "50019", AutomationId: "home" })
        }

        if homeTab {
            homeTab.Click()
            Sleep 200
        } else {
            MsgBox "Could not find the 'Home' tab.", "Power BI", "IconX"
            return
        }

        ; Click the New measure button
        newMeasureBtn := root.FindFirst({ Type: "50000", Name: "New measure", AutomationId: "newMeasure" })
        if !newMeasureBtn {
            newMeasureBtn := root.FindFirst({ Type: "50000", Name: "New measure" })
        }
        if !newMeasureBtn {
            newMeasureBtn := root.FindFirst({ Type: "50000", AutomationId: "newMeasure" })
        }
        if !newMeasureBtn {
            newMeasureBtn := root.FindFirst({ Type: "50000", ClassName: "ms-Button root-336" })
        }

        if newMeasureBtn {
            newMeasureBtn.Click()
        } else {
            MsgBox "Could not find the 'New measure' button.", "Power BI", "IconX"
        }
    } catch Error as e {
        MsgBox "Error triggering New measure: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + Y : Refresh (Click Home tab, then Refresh button)
+y:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Click the Home tab
        homeTab := root.FindFirst({ Type: "50019", Name: "Home", AutomationId: "home" })
        if !homeTab {
            homeTab := root.FindFirst({ Type: "50019", Name: "Home" })
        }
        if !homeTab {
            homeTab := root.FindFirst({ Type: "50019", AutomationId: "home" })
        }

        if homeTab {
            homeTab.Click()
            Sleep 200
        } else {
            MsgBox "Could not find the 'Home' tab.", "Power BI", "IconX"
            return
        }

        ; Find the Refresh button
        refreshBtn := root.FindFirst({ Type: "50000", Name: "Refresh" })
        if !refreshBtn {
            refreshBtn := root.FindFirst({ Type: "50000", ClassName: "splitPrimaryButton root-332" })
        }
        if !refreshBtn {
            refreshBtn := root.FindFirst({ Type: "50000", ClassName: "splitPrimaryButton", matchmode: "Substring" })
        }

        if refreshBtn {
            refreshBtn.Click()
        } else {
            MsgBox "Could not find the 'Refresh' button.", "Power BI", "IconX"
        }
    } catch Error as e {
        MsgBox "Error triggering Refresh: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + B : Bring forward (Click Format tab, then click button 10 times)
+b:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Click the Format tab
        formatTab := root.FindFirst({ Type: "50019", Name: "Format", AutomationId: "visualFormatting" })
        if !formatTab {
            formatTab := root.FindFirst({ Type: "50019", Name: "Format" })
        }
        if !formatTab {
            formatTab := root.FindFirst({ Type: "50019", AutomationId: "visualFormatting" })
        }

        if formatTab {
            formatTab.Click()
            Sleep 200
        } else {
            MsgBox "Could not find the 'Format' tab.", "Power BI", "IconX"
            return
        }

        ; Find the Bring forward button
        bringForwardBtn := root.FindFirst({ Type: "50000", Name: "Bring forward" })
        if !bringForwardBtn {
            bringForwardBtn := root.FindFirst({ Type: "50000", ClassName: "splitPrimaryButton root-332", Name: "Bring forward" })
        }
        if !bringForwardBtn {
            ; Search all buttons for one named "Bring forward"
            allButtons := root.FindAll({ Type: "50000" })
            for btn in allButtons {
                if (btn.Name = "Bring forward") {
                    bringForwardBtn := btn
                    break
                }
            }
        }

        if !bringForwardBtn {
            MsgBox "Could not find the 'Bring forward' button.", "Power BI", "IconX"
            return
        }

        ; Show execution banner
        ShowSmallLoadingIndicator_ChatGPT("Bringing forward 10 times...")

        ; Click the button 10 times
        loop 10 {
            bringForwardBtn.Click()
            Sleep 50
        }

        ; Hide banner after completion
        Sleep 300
        HideSmallLoadingIndicator_ChatGPT()
    } catch Error as e {
        HideSmallLoadingIndicator_ChatGPT()
        MsgBox "Error triggering Bring forward: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + D : Send backward (Click Format tab, then click button 10 times)
+d:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Click the Format tab
        formatTab := root.FindFirst({ Type: "50019", Name: "Format", AutomationId: "visualFormatting" })
        if !formatTab {
            formatTab := root.FindFirst({ Type: "50019", Name: "Format" })
        }
        if !formatTab {
            formatTab := root.FindFirst({ Type: "50019", AutomationId: "visualFormatting" })
        }

        if formatTab {
            formatTab.Click()
            Sleep 200
        } else {
            MsgBox "Could not find the 'Format' tab.", "Power BI", "IconX"
            return
        }

        ; Find the Send backward button
        sendBackwardBtn := root.FindFirst({ Type: "50000", Name: "Send backward" })
        if !sendBackwardBtn {
            sendBackwardBtn := root.FindFirst({ Type: "50000", ClassName: "splitPrimaryButton root-332", Name: "Send backward" })
        }
        if !sendBackwardBtn {
            ; Search all buttons for one named "Send backward"
            allButtons := root.FindAll({ Type: "50000" })
            for btn in allButtons {
                if (btn.Name = "Send backward") {
                    sendBackwardBtn := btn
                    break
                }
            }
        }

        if !sendBackwardBtn {
            MsgBox "Could not find the 'Send backward' button.", "Power BI", "IconX"
            return
        }

        ; Show execution banner
        ShowSmallLoadingIndicator_ChatGPT("Sending backward 10 times...")

        ; Click the button 10 times
        loop 10 {
            sendBackwardBtn.Click()
            Sleep 50
        }

        ; Hide banner after completion
        Sleep 300
        HideSmallLoadingIndicator_ChatGPT()
    } catch Error as e {
        HideSmallLoadingIndicator_ChatGPT()
        MsgBox "Error triggering Send backward: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + K : Align (Click Format tab, then click Align button)
+k:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Click the Format tab
        formatTab := root.FindFirst({ Type: "50019", Name: "Format", AutomationId: "visualFormatting" })
        if !formatTab {
            formatTab := root.FindFirst({ Type: "50019", Name: "Format" })
        }
        if !formatTab {
            formatTab := root.FindFirst({ Type: "50019", AutomationId: "visualFormatting" })
        }

        if formatTab {
            formatTab.Click()
            Sleep 200
        } else {
            MsgBox "Could not find the 'Format' tab.", "Power BI", "IconX"
            return
        }

        ; Find the Align button
        alignBtn := root.FindFirst({ Type: "50000", Name: "Align", AutomationId: "alignFlyout" })
        if !alignBtn {
            alignBtn := root.FindFirst({ Type: "50000", Name: "Align" })
        }
        if !alignBtn {
            alignBtn := root.FindFirst({ Type: "50000", AutomationId: "alignFlyout" })
        }
        if !alignBtn {
            alignBtn := root.FindFirst({ Type: "50000", ClassName: "ms-Button ms-Button--hasMenu root-337" })
        }
        if !alignBtn {
            ; Search all buttons for one named "Align"
            allButtons := root.FindAll({ Type: "50000" })
            for btn in allButtons {
                if (btn.Name = "Align") {
                    alignBtn := btn
                    break
                }
            }
        }

        if !alignBtn {
            MsgBox "Could not find the 'Align' button.", "Power BI", "IconX"
            return
        }

        ; Click the Align button
        alignBtn.Click()
    } catch Error as e {
        MsgBox "Error triggering Align: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + V : Fit to page
+v:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Find the Fit to page button
        fitToPageBtn := root.FindFirst({ Type: "50000", Name: "Fit to page", AutomationId: "fitToPageButton" })
        if !fitToPageBtn {
            fitToPageBtn := root.FindFirst({ Type: "50000", Name: "Fit to page" })
        }
        if !fitToPageBtn {
            fitToPageBtn := root.FindFirst({ Type: "50000", AutomationId: "fitToPageButton" })
        }
        if !fitToPageBtn {
            fitToPageBtn := root.FindFirst({ Type: "50000", ClassName: "smallImageButton", AutomationId: "fitToPageButton" })
        }
        if !fitToPageBtn {
            ; Search all buttons for one named "Fit to page"
            allButtons := root.FindAll({ Type: "50000" })
            for btn in allButtons {
                if (btn.Name = "Fit to page") {
                    fitToPageBtn := btn
                    break
                }
            }
        }

        if !fitToPageBtn {
            MsgBox "Could not find the 'Fit to page' button.", "Power BI", "IconX"
            return
        }

        ; Click the Fit to page button
        fitToPageBtn.Click()
    } catch Error as e {
        MsgBox "Error triggering Fit to page: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + M : Format painter (Match format)
+m:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Find the Format painter button by AutomationId (primary method)
        formatPainterBtn := root.FindFirst({ Type: "50000", AutomationId: "formatPainter" })
        if !formatPainterBtn {
            formatPainterBtn := root.FindFirst({ Type: 50000, AutomationId: "formatPainter" })
        }

        ; Fallback: Try by Name and Type
        if !formatPainterBtn {
            formatPainterBtn := root.FindFirst({ Type: "50000", Name: "Format painter" })
        }
        if !formatPainterBtn {
            formatPainterBtn := root.FindFirst({ Type: 50000, Name: "Format painter" })
        }

        ; Fallback: Try by ClassName
        if !formatPainterBtn {
            formatPainterBtn := root.FindFirst({ Type: "50000", ClassName: "ms-Button root-345", AutomationId: "formatPainter" })
        }
        if !formatPainterBtn {
            formatPainterBtn := root.FindFirst({ Type: 50000, ClassName: "ms-Button root-345" })
        }

        ; Fallback: Try partial ClassName match
        if !formatPainterBtn {
            formatPainterBtn := root.FindFirst({ Type: "50000", ClassName: "ms-Button", matchmode: "Substring" })
            ; Verify it's the right button by checking AutomationId or Name
            if formatPainterBtn {
                try {
                    if (formatPainterBtn.AutomationId != "formatPainter" && formatPainterBtn.Name != "Format painter") {
                        formatPainterBtn := ""
                    }
                } catch {
                    formatPainterBtn := ""
                }
            }
        }

        ; Last resort: Search all buttons for Format painter
        if !formatPainterBtn {
            allButtons := root.FindAll({ Type: "50000" })
            for btn in allButtons {
                try {
                    if (btn.AutomationId = "formatPainter" || btn.Name = "Format painter") {
                        formatPainterBtn := btn
                        break
                    }
                } catch {
                }
            }
        }

        if !formatPainterBtn {
            MsgBox "Could not find the 'Format painter' button.", "Power BI", "IconX"
            return
        }

        ; Click the Format painter button
        formatPainterBtn.Click()
    } catch Error as e {
        MsgBox "Error triggering Format painter: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + N : Group visuals (Click Format tab, then click Group button)
+n:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Click the Format tab
        formatTab := root.FindFirst({ Type: "50019", Name: "Format", AutomationId: "visualFormatting" })
        if !formatTab {
            formatTab := root.FindFirst({ Type: "50019", Name: "Format" })
        }
        if !formatTab {
            formatTab := root.FindFirst({ Type: "50019", AutomationId: "visualFormatting" })
        }

        if formatTab {
            formatTab.Click()
            Sleep 200
        } else {
            MsgBox "Could not find the 'Format' tab.", "Power BI", "IconX"
            return
        }

        ; Find the Group button by AutomationId (primary method)
        groupBtn := root.FindFirst({ Type: "50000", AutomationId: "groupVisualsFlyout" })
        if !groupBtn {
            groupBtn := root.FindFirst({ Type: 50000, AutomationId: "groupVisualsFlyout" })
        }

        ; Fallback: Try by Name and Type
        if !groupBtn {
            groupBtn := root.FindFirst({ Type: "50000", Name: "Group" })
        }
        if !groupBtn {
            groupBtn := root.FindFirst({ Type: 50000, Name: "Group" })
        }

        ; Fallback: Try by ClassName
        if !groupBtn {
            groupBtn := root.FindFirst({ Type: "50000", ClassName: "ms-Button ms-Button--hasMenu root-337",
                AutomationId: "groupVisualsFlyout" })
        }
        if !groupBtn {
            groupBtn := root.FindFirst({ Type: 50000, ClassName: "ms-Button ms-Button--hasMenu root-337" })
        }

        ; Fallback: Try partial ClassName match
        if !groupBtn {
            groupBtn := root.FindFirst({ Type: "50000", ClassName: "ms-Button--hasMenu", matchmode: "Substring" })
            ; Verify it's the right button by checking AutomationId or Name
            if groupBtn {
                try {
                    if (groupBtn.AutomationId != "groupVisualsFlyout" && groupBtn.Name != "Group") {
                        groupBtn := ""
                    }
                } catch {
                    groupBtn := ""
                }
            }
        }

        ; Last resort: Search all buttons for Group
        if !groupBtn {
            allButtons := root.FindAll({ Type: "50000" })
            for btn in allButtons {
                try {
                    if (btn.AutomationId = "groupVisualsFlyout" || btn.Name = "Group") {
                        groupBtn := btn
                        break
                    }
                } catch {
                }
            }
        }

        if !groupBtn {
            MsgBox "Could not find the 'Group' button.", "Power BI", "IconX"
            return
        }

        ; Click the Group button
        groupBtn.Click()
    } catch Error as e {
        MsgBox "Error triggering Group: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + F : Close all Power BI drawers (Visualizations/Data/Properties/Filters)
+f:: {
    try {
        win := WinExist("A")
        if !win
            return

        root := UIA.ElementFromHandle(win)

        drawerConfigs := PowerBI_GetDrawerConfigs()

        closed := 0
        already := 0
        skipped := 0

        for , cfg in drawerConfigs {
            btn := PowerBI_FindDrawerButton(root, cfg)
            if !btn {
                skipped++
                continue
            }
            result := PowerBI_CollapseDrawerElement(btn)
            if result = 1 {
                closed++
            } else if result = 0 {
                already++
            } else {
                skipped++
            }
        }

        msg := closed
            ? Format("Closed {} drawer{}", closed, closed = 1 ? "" : "s")
                : "No drawers needed closing"

        if already
            msg .= Format(" | {} already closed", already)
        if skipped
            msg .= Format(" | {} skipped", skipped)

        ToolTip msg
        SetTimer(() => ToolTip(), -1500)
    } catch Error as e {
        MsgBox "Error closing Power BI drawers: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + G : Open all Power BI drawers (Visualizations/Data/Properties/Filters)
+g:: {
    try {
        win := WinExist("A")
        if !win
            return

        root := UIA.ElementFromHandle(win)
        drawerConfigs := PowerBI_GetDrawerConfigs()

        opened := 0
        already := 0
        skipped := 0

        for , cfg in drawerConfigs {
            btn := PowerBI_FindDrawerButton(root, cfg)
            if !btn {
                skipped++
                continue
            }
            result := PowerBI_ExpandDrawerElement(btn)
            if result = 1 {
                opened++
            } else if result = 0 {
                already++
            } else {
                skipped++
            }
        }

        msg := opened
            ? Format("Opened {} drawer{}", opened, opened = 1 ? "" : "s")
                : "No drawers needed opening"

        if already
            msg .= Format(" | {} already open", already)
        if skipped
            msg .= Format(" | {} skipped", skipped)

        ToolTip msg
        SetTimer(() => ToolTip(), -1500)
    } catch Error as e {
        MsgBox "Error opening Power BI drawers: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + R : Collapse Power BI table tree items
+r:: {
    try {
        win := WinExist("A")
        if !win
            return

        root := UIA.ElementFromHandle(win)

        treeItemCond := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.TreeItem)
        expandCond := UIA.CreatePropertyCondition(UIA.Property.IsExpandCollapsePatternAvailable, true)
        tableCond := UIA.CreatePropertyConditionEx(UIA.Property.Name, "Table ", UIA.PropertyConditionFlags.IgnoreCaseMatchSubstring
        )
        calcTableCond := UIA.CreatePropertyConditionEx(UIA.Property.Name, "Calculated Table", UIA.PropertyConditionFlags
            .IgnoreCaseMatchSubstring)
        nameCond := UIA.CreateOrCondition(tableCond, calcTableCond)
        targetCond := UIA.CreateAndCondition(treeItemCond, UIA.CreateAndCondition(expandCond, nameCond))

        items := ""
        try items := root.FindElements(targetCond, UIA.TreeScope.Descendants)

        if !items {
            MsgBox "Could not find any Power BI tables to collapse.", "Power BI", "IconX"
            return
        }

        collapsed := 0
        already := 0

        for item in items {
            if !item
                continue
            try {
                pat := item.ExpandCollapsePattern
                if pat.ExpandCollapseState != UIA.ExpandCollapseState.Collapsed {
                    pat.Collapse()
                    collapsed++
                    Sleep 35
                } else {
                    already++
                }
            } catch Error {
                try {
                    item.SetFocus()
                    Sleep 40
                    Send "{Left}"
                    collapsed++
                } catch {
                }
            }
        }

        if collapsed {
            ToolTip Format("Collapsed {} table{}", collapsed, collapsed = 1 ? "" : "s")
        } else if already {
            ToolTip "All tables already collapsed"
        } else {
            ToolTip "No tables collapsed"
        }
        SetTimer(() => ToolTip(), -1200)
    } catch Error as e {
        MsgBox "Error collapsing Power BI tables: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + Z : Copy table cell value (context menu → Enter ×2)
+z:: {
    Send "{AppsKey}"
    Sleep 200
    Send "{Enter}"
    Sleep 100
    Send "{Enter}"
}

