; =============================================================================
; Lib: gemini_tools.ahk
; Personal Gemini Tools drawer (Create image / Deep research).
; Included from Utils.ahk so AppLaunchers and Shift keys share one implementation.
; Shift+I in Shift keys\gemini_chrome_02.ahk calls Gemini_ToggleCreateImage.
; =============================================================================

GEMINI_TOOLBOX_CHECKBOX_TYPE := 50002

Gemini_FindToolsButton(uia) {
    if !IsObject(uia)
        return 0
    saved := UiElements_TrySaved(uia, "Gemini", "Tools")
    if (saved)
        return saved
    toolsButton := 0
    try {
        toolsButton := uia.FindFirst({ Name: "Upload & tools", Type: 50000 })
        if toolsButton
            return toolsButton

        allButtons := uia.FindAll({ Type: 50000 })
        for button in allButtons {
            nm := button.Name
            if (InStr(nm, "tools", false) || InStr(nm, "ferramentas", false))
            && (InStr(nm, "Upload", false) || InStr(nm, "Enviar", false) || nm = "Tools")
                return button
        }

        for grp in uia.FindAll({ Type: 50026 }) {
            if !InStr(grp.ClassName, "gem-menu-button", false)
                continue
            try {
                btn := grp.FindFirst({ Type: 50000 })
                if btn
                    return btn
            } catch {
            }
        }

        toolsButton := uia.FindFirst({ Name: "Tools", Type: 50000 })
        if !toolsButton
            toolsButton := uia.FindFirst({ Type: "Button", Name: "Tools" })
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
    try {
        if uia.FindFirst({ Name: "Menu options", Type: 50009 })
            return true
        if uia.FindFirst({ Name: "Menu options", Type: UIA.Type.Menu })
            return true
    } catch {
    }
    try {
        for el in uia.FindAll({ Type: 50026 }) {
            cls := el.ClassName
            if InStr(cls, "card-container", false) || InStr(cls, "menu-list-container", false)
                return true
        }
    } catch {
    }
    try {
        for cb in uia.FindAll({ Type: GEMINI_TOOLBOX_CHECKBOX_TYPE }) {
            if InStr(cb.ClassName, "toolbox-drawer-item-list-button", false)
                return true
        }
    } catch {
    }
    return false
}

Gemini_OpenToolsMenu(uia) {
    if Gemini_IsToolsMenuOpen(uia)
        return true
    btn := Gemini_FindToolsButton(uia)
    if !btn
        return false
    try {
        btn.Click()
    } catch {
        return false
    }
    Sleep 250
    return true
}

; Open menu only when the toolbox drawer is not already in the tree (avoids toggling closed).
Gemini_EnsureToolsMenuOpen(&uia) {
    if Gemini_IsToolsMenuOpen(uia)
        return true
    if !Gemini_OpenToolsMenu(uia)
        return false
    ; Refresh UIA — stale element trees are common right after the Tools click.
    try
        uia := UIA_Browser()
    catch
        return false
    deadline := A_TickCount + 3000
    while (A_TickCount < deadline) {
        if Gemini_IsToolsMenuOpen(uia)
            return true
        Sleep 80
        try
            uia := UIA_Browser()
        catch
            return false
    }
    return false
}

Gemini_FindToolboxMenuPane(uia) {
    if !IsObject(uia)
        return 0
    try {
        m := uia.FindFirst({ Name: "Menu options", Type: 50009 })
        if m
            return m
        m := uia.FindFirst({ Name: "Menu options", Type: UIA.Type.Menu })
        if m
            return m
        m := uia.FindFirst({ AutomationId: "toolbox-drawer-menu" })
        if m
            return m
        m := uia.FindFirst({ Type: UIA.Type.Menu, AutomationId: "toolbox-drawer-menu" })
        if m
            return m
        for el in uia.FindAll({ Type: 50026 }) {
            if InStr(el.ClassName, "card-container", false)
                return el
        }
    } catch {
    }
    return 0
}

Gemini_FindMoreToolsButton(uia) {
    if !IsObject(uia)
        return 0
    try {
        saved := UiElements_TrySaved(uia, "Gemini", "MoreTools")
        if (saved)
            return saved
    } catch {
    }
    try {
        for name in ["More tools", "Mais ferramentas"] {
            btn := uia.FindFirst({ Name: name, Type: 50000 })
            if btn
                return btn
        }
        for btn in uia.FindAll({ Type: 50000 }) {
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
Gemini_EnsureMoreToolsExpanded(&uia, nameSubstrings) {
    if !IsObject(uia) || !IsObject(nameSubstrings) || nameSubstrings.Length = 0
        return false
    if Gemini_FindToolboxCheckBox(uia, nameSubstrings)
        return true
    btn := Gemini_FindMoreToolsButton(uia)
    if !btn
        return false
    try {
        btn.Click()
    } catch {
        return false
    }
    try
        uia := UIA_Browser()
    catch
        return false
    deadline := A_TickCount + 2000
    while (A_TickCount < deadline) {
        if Gemini_FindToolboxCheckBox(uia, nameSubstrings)
            return true
        Sleep 80
        try
            uia := UIA_Browser()
        catch
            return false
    }
    return false
}

; nameSubstrings: ordered list — first matching toolbox-drawer checkbox wins (EN/PT partial names).
; Searches inside toolbox-drawer-menu when present (fast + reliable); falls back to full tree.
Gemini_FindToolboxCheckBox(uia, nameSubstrings) {
    if !IsObject(uia) || !IsObject(nameSubstrings) || nameSubstrings.Length = 0
        return 0
    scope := uia
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
        if !allCb
            return 0
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
    } catch {
        return 0
    }
    return 0
}

; Material toolbox items expose Invoke + Toggle; Invoke alone often dismisses the menu without toggling.
; Prefer Toggle; then physical center click (Click("left")) — not bare Click() which tries Invoke first.
Gemini_ActivateToolboxItem(el) {
    if !IsObject(el)
        return
    try {
        if el.GetPropertyValue(UIA.Property.IsTogglePatternAvailable) {
            el.TogglePattern.Toggle()
            return
        }
    } catch {
    }
    try {
        el.Click("left")
    } catch {
    }
}

; Same action as Shift+I on personal Gemini: Tools menu, Create image.
Gemini_ToggleCreateImage() {
    try {
        uia := UIA_Browser()
        if !IsObject(uia)
            return false
        Sleep 150
        subs := ["Create image", "Criar imagem"]
        cb := 0
        try cb := UiElements_TrySaved(uia, "Gemini", "CreateImage")
        if !cb
            cb := Gemini_FindToolboxCheckBox(uia, subs)
        if !cb {
            if !Gemini_EnsureToolsMenuOpen(&uia)
                return false
            try cb := UiElements_TrySaved(uia, "Gemini", "CreateImage")
            if !cb
                cb := Gemini_FindToolboxCheckBox(uia, subs)
        }
        if !cb
            return false
        Gemini_ActivateToolboxItem(cb)
        return true
    } catch Error as e {
        return false
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
