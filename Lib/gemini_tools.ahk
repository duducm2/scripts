; =============================================================================
; Lib: gemini_tools.ahk
; Personal Gemini Tools drawer (Create image / Deep research).
; Included from Utils.ahk so AppLaunchers and Shift keys share one implementation.
; Shift+I in Shift keys\gemini_chrome_02.ahk calls Gemini_ToggleCreateImage.
; =============================================================================

GEMINI_TOOLBOX_CHECKBOX_TYPE := 50002

; Active gemini.google.com document. Background Chrome tabs must not shadow the
; visible tab's Upload & tools button, menu, or Copy controls.
Gemini_ActivePageRoot(uia) {
    if (!IsObject(uia))
        return uia
    docs := []
    try docs := uia.FindAll({ Type: "Document" })
    catch
        docs := []
    geminiDocs := []
    for doc in docs {
        name := ""
        val := ""
        try name := doc.Name
        try val := doc.Value
        if (InStr(val, "gemini.google.com", false) || InStr(name, "Gemini", false))
            geminiDocs.Push(doc)
    }
    if (geminiDocs.Length = 0)
        return uia
    if (geminiDocs.Length = 1)
        return geminiDocs[1]
    title := ""
    try {
        hwnd := 0
        try hwnd := uia.BrowserId
        if (hwnd)
            title := WinGetTitle("ahk_id " hwnd)
    } catch {
        title := ""
    }
    if (!title) {
        try title := WinGetTitle("A")
        catch
            title := ""
    }
    best := 0
    bestLen := 0
    bestVisible := false
    for doc in geminiDocs {
        name := ""
        try name := doc.Name
        if (name && title && InStr(title, name, false) = 1) {
            off := true
            try off := doc.GetPropertyValue(UIA.Property.IsOffscreen)
            vis := !off
            nLen := StrLen(name)
            if (!best || (vis && !bestVisible) || (vis = bestVisible && nLen > bestLen)) {
                best := doc
                bestLen := nLen
                bestVisible := vis
            }
        }
    }
    if (best)
        return best
    for doc in geminiDocs {
        off := true
        try off := doc.GetPropertyValue(UIA.Property.IsOffscreen)
        if (!off)
            return doc
    }
    return geminiDocs[geminiDocs.Length]
}

Gemini_FindToolsButton(uia) {
    if !IsObject(uia)
        return 0
    root := Gemini_ActivePageRoot(uia)
    saved := UiElements_TrySaved(root, "Gemini", "Tools")
    if (saved)
        return saved
    toolsButton := 0
    try {
        toolsButton := root.FindFirst({ Name: "Upload & tools", Type: 50000 })
        if toolsButton
            return toolsButton

        allButtons := root.FindAll({ Type: 50000 })
        for button in allButtons {
            nm := button.Name
            if (InStr(nm, "tools", false) || InStr(nm, "ferramentas", false))
            && (InStr(nm, "Upload", false) || InStr(nm, "Enviar", false) || nm = "Tools")
                return button
        }

        for grp in root.FindAll({ Type: 50026 }) {
            if !InStr(grp.ClassName, "gem-menu-button", false)
                continue
            try {
                btn := grp.FindFirst({ Type: 50000 })
                if btn
                    return btn
            } catch {
            }
        }

        toolsButton := root.FindFirst({ Name: "Tools", Type: 50000 })
        if !toolsButton
            toolsButton := root.FindFirst({ Type: "Button", Name: "Tools" })
        if !toolsButton {
            for button in allButtons {
                if InStr(button.ClassName, "toolbox-drawer-button") && InStr(button.Name, "Tools") {
                    toolsButton := button
                    break
                }
            }
        }
        if !toolsButton {
            for button in allButtons {
                if (InStr(button.Name, "Tools") || InStr(button.Name, "Ferramentas")) && InStr(button.ClassName,
                    "toolbox-drawer-button") {
                    toolsButton := button
                    break
                }
            }
        }
    } catch {
        return 0
    }
    return toolsButton ? toolsButton : 0
}

Gemini_IsToolsMenuOpen(uia) {
    if !IsObject(uia)
        return false
    root := Gemini_ActivePageRoot(uia)
    try {
        if root.FindFirst({ Name: "Menu options", Type: 50009 })
            return true
        if root.FindFirst({ Name: "Menu options", Type: UIA.Type.Menu })
            return true
        if root.FindFirst({ Type: 50009 })
            return true
    } catch {
    }
    try {
        for el in root.FindAll({ Type: 50026 }) {
            cls := el.ClassName
            if InStr(cls, "card-container", false) || InStr(cls, "menu-list-container", false)
            || InStr(cls, "mat-mdc-menu-panel", false) || InStr(cls, "mat-menu-content", false)
            || InStr(cls, "cdk-overlay-pane", false)
                return true
        }
    } catch {
    }
    try {
        for cb in root.FindAll({ Type: GEMINI_TOOLBOX_CHECKBOX_TYPE }) {
            if InStr(cb.ClassName, "toolbox-drawer-item-list-button", false)
                return true
        }
    } catch {
    }
    try {
        for mi in root.FindAll({ Type: 50011 }) {
            nm := mi.Name
            if InStr(nm, "Create image", false) || InStr(nm, "Criar imagem", false)
            || InStr(nm, "Upload", false) || InStr(nm, "Deep research", false)
                return true
        }
    } catch {
    }
    if Gemini_FindToolboxCheckBoxInScope(root, ["Upload files", "Enviar arquivos", "More tools", "Mais ferramentas"])
        return true
    return false
}

Gemini_OpenToolsMenu(uia) {
    if Gemini_IsToolsMenuOpen(uia)
        return true
    btn := Gemini_FindToolsButton(uia)
    if !btn
        return false
    try {
        btn.Click("left")
        Sleep 250
        return true
    } catch {
    }
    try {
        btn.Click()
        Sleep 250
        return true
    } catch {
        return false
    }
}

; Open menu only when the toolbox drawer is not already in the tree (avoids toggling closed).
Gemini_EnsureToolsMenuOpen(&uia, hwnd := 0) {
    if (!hwnd) {
        try hwnd := uia.BrowserId
        catch
            hwnd := 0
        if (!hwnd)
            hwnd := WinExist("A")
    }
    if Gemini_IsToolsMenuOpen(uia)
        return true
    if !Gemini_OpenToolsMenu(uia)
        return false
    ; Refresh UIA pinned to hwnd — stale element trees are common right after the Tools click.
    try
        uia := hwnd ? UIA_Browser("ahk_id " hwnd) : UIA_Browser()
    catch
        return false
    deadline := A_TickCount + 3000
    while (A_TickCount < deadline) {
        if Gemini_IsToolsMenuOpen(uia)
            return true
        Sleep 80
        try
            uia := hwnd ? UIA_Browser("ahk_id " hwnd) : UIA_Browser()
        catch
            return false
    }
    return false
}

Gemini_FindToolboxMenuPane(uia) {
    if !IsObject(uia)
        return 0
    root := Gemini_ActivePageRoot(uia)
    try {
        m := root.FindFirst({ Name: "Menu options", Type: 50009 })
        if m
            return m
        m := root.FindFirst({ Name: "Menu options", Type: UIA.Type.Menu })
        if m
            return m
        m := root.FindFirst({ AutomationId: "toolbox-drawer-menu" })
        if m
            return m
        m := root.FindFirst({ Type: UIA.Type.Menu, AutomationId: "toolbox-drawer-menu" })
        if m
            return m
        for el in root.FindAll({ Type: 50026 }) {
            cls := el.ClassName
            if InStr(cls, "card-container", false) || InStr(cls, "mat-mdc-menu-panel", false)
            || InStr(cls, "menu-list-container", false) || InStr(cls, "mat-menu-content", false)
            || InStr(cls, "cdk-overlay-pane", false)
                return el
        }
        m := root.FindFirst({ Type: 50009 })
        if m
            return m
    } catch {
    }
    return 0
}

Gemini_FindMoreToolsButton(uia) {
    if !IsObject(uia)
        return 0
    root := Gemini_ActivePageRoot(uia)
    try {
        saved := UiElements_TrySaved(root, "Gemini", "MoreTools")
        if (saved)
            return saved
    } catch {
    }
    try {
        for name in ["More tools", "Mais ferramentas"] {
            btn := root.FindFirst({ Name: name, Type: 50000 })
            if btn
                return btn
        }
        for btn in root.FindAll({ Type: 50000 }) {
            cls := btn.ClassName
            nm := btn.Name
            if InStr(cls, "more-tools-button", false)
            || (InStr(cls, "toolbox-drawer-menu-item", false) && (InStr(nm, "More tools", false) || InStr(nm,
                "Mais ferramentas", false)))
                return btn
        }
    } catch {
    }
    return 0
}

; Expands the "More tools" submenu when the target checkbox is not in the top-level menu.
Gemini_EnsureMoreToolsExpanded(&uia, nameSubstrings, hwnd := 0) {
    if !IsObject(uia) || !IsObject(nameSubstrings) || nameSubstrings.Length = 0
        return false
    if (!hwnd) {
        try hwnd := uia.BrowserId
        catch
            hwnd := 0
        if (!hwnd)
            hwnd := WinExist("A")
    }
    if Gemini_FindToolboxCheckBox(uia, nameSubstrings)
        return true
    btn := Gemini_FindMoreToolsButton(uia)
    if !btn
        return false
    try {
        btn.Click()
    } catch {
        try btn.Click("left")
        catch
            return false
    }
    try
        uia := hwnd ? UIA_Browser("ahk_id " hwnd) : UIA_Browser()
    catch
        return false
    deadline := A_TickCount + 2000
    while (A_TickCount < deadline) {
        if Gemini_FindToolboxCheckBox(uia, nameSubstrings)
            return true
        Sleep 80
        try
            uia := hwnd ? UIA_Browser("ahk_id " hwnd) : UIA_Browser()
        catch
            return false
    }
    return false
}

; nameSubstrings: ordered list — first matching toolbox-drawer checkbox wins (EN/PT partial names).
; Searches inside toolbox-drawer-menu when present (fast + reliable); falls back to active document root.
Gemini_FindToolboxCheckBox(uia, nameSubstrings) {
    if !IsObject(uia) || !IsObject(nameSubstrings) || nameSubstrings.Length = 0
        return 0
    scope := Gemini_ActivePageRoot(uia)
    try {
        menu := Gemini_FindToolboxMenuPane(uia)
        if menu
            scope := menu
    } catch {
    }
    return Gemini_FindToolboxCheckBoxInScope(scope, nameSubstrings)
}

Gemini_FindToolboxCheckBoxInScope(scope, nameSubstrings) {
    if !IsObject(scope) || !IsObject(nameSubstrings) || nameSubstrings.Length = 0
        return 0
    try {
        allCb := scope.FindAll({ Type: GEMINI_TOOLBOX_CHECKBOX_TYPE })
        if !allCb || allCb.Length = 0
            allCb := scope.FindAll({ Type: "CheckBox" })
        if (allCb) {
            for sub in nameSubstrings {
                for cb in allCb {
                    try {
                        nm := cb.Name
                        cls := cb.ClassName
                    } catch {
                        continue
                    }
                    if !InStr(cls, "toolbox-drawer-item", false)
                        continue
                    if InStr(nm, sub, false)
                        return cb
                }
            }
        }
        ; Fallback for gem-menu-redesign (Upload & tools) where items may be MenuItem, Button, CheckBox, or ListItem
        ; without the legacy "toolbox-drawer-item" class.
        for ctrlType in [50011, GEMINI_TOOLBOX_CHECKBOX_TYPE, 50000, 50007] {
            items := 0
            try items := scope.FindAll({ Type: ctrlType })
            if (!items || items.Length = 0)
                continue
            for sub in nameSubstrings {
                for item in items {
                    try nm := item.Name
                    catch
                        continue
                    if (nm = "")
                        continue
                    if InStr(nm, "Deselect", false) || InStr(nm, "Remove ", false) || InStr(nm, "Remover ", false)
                        continue
                    if InStr(nm, sub, false)
                        return item
                }
            }
        }
    } catch {
        return 0
    }
    return 0
}

; Material toolbox items expose Invoke + Toggle; Invoke alone often dismisses the menu without toggling.
; Prefer Toggle only for CheckBox items; use physical center click (Click("left")) for MenuItem/Button.
Gemini_ActivateToolboxItem(el) {
    if !IsObject(el)
        return
    ctrlType := 0
    try ctrlType := el.Type
    catch
        ctrlType := 0
    if (ctrlType = GEMINI_TOOLBOX_CHECKBOX_TYPE) {
        try {
            if el.GetPropertyValue(UIA.Property.IsTogglePatternAvailable) {
                el.TogglePattern.Toggle()
                return
            }
        } catch {
        }
    }
    try {
        el.Click("left")
        return
    } catch {
    }
    try {
        el.Click()
    } catch {
    }
}

Gemini_WriteToggleImageLog(lines) {
    try {
        logPath := A_ScriptDir "\assets\data\gemini_toggle_image.log"
        try FileDelete(logPath)
        FileAppend(lines, logPath, "UTF-8")
    } catch {
    }
}

; Same action as Shift+I on personal Gemini: Tools menu, Create image.
Gemini_ToggleCreateImage() {
    try StandardLoadingBar_Show("⏳ Create image…", BANNER_ACCENT_INTERMEDIATE, { passive: true, centerOnHwnd: 0 })
    ok := false
    log := "ts=" A_Now "`n"
    try {
        hwnd := WinExist("A")
        title := ""
        try title := WinGetTitle("ahk_id " hwnd)
        log .= "hwnd=" hwnd " title=" title "`n"
        uia := hwnd ? UIA_Browser("ahk_id " hwnd) : UIA_Browser()
        if !IsObject(uia) {
            log .= "err=NO_UIA_BROWSER`n"
            return false
        }
        Sleep 150
        subs := ["Create image", "Create images", "Criar imagem", "Criar imagens"]
        root := Gemini_ActivePageRoot(uia)
        try log .= "root=" root.Name " type=" root.Type "`n"
        cb := 0
        try cb := UiElements_TrySaved(root, "Gemini", "CreateImage")
        if !cb
            cb := Gemini_FindToolboxCheckBox(uia, subs)
        log .= "preOpenCb=" (IsObject(cb) ? cb.Name : "0") "`n"
        if !cb {
            tBtn := Gemini_FindToolsButton(uia)
            try log .= "toolsBtn=" (IsObject(tBtn) ? (tBtn.Name " type=" tBtn.Type) : "0") "`n"
            menuOk := Gemini_EnsureToolsMenuOpen(&uia, hwnd)
            log .= "ensureMenuOpen=" menuOk "`n"
            if !menuOk {
                root := Gemini_ActivePageRoot(uia)
                cbFallback := Gemini_FindToolboxCheckBox(uia, subs)
                log .= "menuTimeoutCb=" (IsObject(cbFallback) ? cbFallback.Name : "0") "`n"
                if !cbFallback
                    return false
                cb := cbFallback
            } else {
                root := Gemini_ActivePageRoot(uia)
                try cb := UiElements_TrySaved(root, "Gemini", "CreateImage")
                if !cb
                    cb := Gemini_FindToolboxCheckBox(uia, subs)
            }
        }
        if !cb && Gemini_EnsureMoreToolsExpanded(&uia, subs, hwnd)
            cb := Gemini_FindToolboxCheckBox(uia, subs)
        log .= "finalCb=" (IsObject(cb) ? (cb.Name " type=" cb.Type " cls=" cb.ClassName) : "0") "`n"
        if !cb {
            try {
                for t in [50011, 50002, 50000, 50007] {
                    arr := root.FindAll({ Type: t })
                    if (arr && arr.Length > 0) {
                        for el in arr {
                            try {
                                if (el.Name != "")
                                    log .= "  cand[" t "] " el.Name " | " el.ClassName "`n"
                            } catch {
                            }
                        }
                    }
                }
            } catch {
            }
            return false
        }
        Gemini_ActivateToolboxItem(cb)
        ok := true
        log .= "activated=1`n"
        return true
    } catch Error as e {
        log .= "exception=" e.Message "`n"
        return false
    } finally {
        Gemini_WriteToggleImageLog(log)
        if (ok) {
            try StandardLoadingBar_Hide(0)
        } else {
            try StandardLoadingBar_Show("Create image failed", BANNER_ACCENT_ERROR, { passive: true, centerOnHwnd: 0 })
            try StandardLoadingBar_Hide(1800)
        }
    }
}

; Shift+I on whichever companion is active: toggle image generation, then return to the composer.
AiCompanion_ToggleImageGeneration(companion := "") {
    if (companion = "") {
        try companion := ResolveGlobalAICompanion()
        catch
            companion := ""
    }
    ok := false
    switch companion {
        case "enterprise":
            try ok := GeminiEnterprise_ClickCreateImages()
            catch
                ok := false
            try GeminiEnterprise_ReturnToComposer()
            catch {
            }
        case "copilot":
            try ok := CopilotWeb_ClickAddCapability(COPILOT_CAPABILITY_IMAGE_NAMES)
            catch
                ok := false
            try CopilotWeb_ReturnToComposer()
            catch {
            }
        default:
            try ok := Gemini_ToggleCreateImage()
            catch
                ok := false
            Sleep 80
            try Send("{Escape}")
            catch {
            }
            Sleep 60
            try {
                uia := UIA_Browser()
                Gemini_FocusPromptWithChime(uia, { playChime: false, useAnchorFallback: true })
            } catch {
            }
    }
    return ok
}
