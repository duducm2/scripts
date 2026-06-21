; =============================================================================
; Shift keys module: gemini_chrome_02.ahk
; Gemini Chrome tools drawer and hotkeys (part 2)
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

; --- Gemini Tools drawer (shared by +t, +i, +e) --------------------------------
GEMINI_TOOLBOX_CHECKBOX_TYPE := 50002

Gemini_FindToolsButton(uia) {
    if !IsObject(uia)
        return 0
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

; Shift + T : Click the Tools button - Tools
+t:: {
    try {
        uia := UIA_Browser()
        Sleep 300

        if (Gemini_OpenToolsMenu(uia)) {
            Sleep 100
            Send "{Tab}"
        }
    } catch Error as e {
        ; If all else fails, silently fail (no fallback action defined)
    }
}

; Shift + I : Tools menu — Create image (opens Tools if needed)
; $ = keyboard hook — reduces stray key passthrough to the prompt while this runs.
$+i:: {
    try {
        uia := UIA_Browser()
        if !IsObject(uia)
            return
        Sleep 150
        subs := ["Create image", "Criar imagem"]
        cb := Gemini_FindToolboxCheckBox(uia, subs)
        if !cb {
            if !Gemini_EnsureToolsMenuOpen(&uia)
                return
            cb := Gemini_FindToolboxCheckBox(uia, subs)
        }
        if (cb)
            Gemini_ActivateToolboxItem(cb)
    } catch Error as e {
    }
}

; Shift + E : Tools menu — Deep research (opens Tools if needed)
$+e:: {
    try {
        uia := UIA_Browser()
        if !IsObject(uia)
            return
        Sleep 150
        subs := ["Deep research", "Pesquisa aprofundada", "Investigação profunda", "Pesquisa profunda"]
        cb := Gemini_FindToolboxCheckBox(uia, subs)
        if !cb {
            if !Gemini_EnsureToolsMenuOpen(&uia)
                return
            cb := Gemini_FindToolboxCheckBox(uia, subs)
        }
        if !cb && Gemini_EnsureMoreToolsExpanded(&uia, subs)
            cb := Gemini_FindToolboxCheckBox(uia, subs)
        if (cb)
            Gemini_ActivateToolboxItem(cb)
    } catch Error as e {
    }
}

; Helper: Focus the Gemini prompt text field using UIA
FocusGeminiPromptField() {
    try {
        uia := UIA_Browser()
        Sleep 150  ; small settle per README (keep this snappy)

        ; Primary strategy: Find by Name (Gemini updated placeholder in 2025)
        try
            promptField := uia.FindFirst({ Name: "Enter a prompt for Gemini", Type: 50004 })
        catch
            promptField := ""

        ; Fallback 1: Legacy name "Enter a prompt here"
        if !promptField {
            try
                promptField := uia.FindFirst({ Name: "Enter a prompt here", Type: 50004 })
            catch
                promptField := ""
        }

        ; Shared scan for remaining fallbacks (single FindAll + scoring for efficiency)
        if !promptField {
            best := 0, bestScore := -1
            allEdits := uia.FindAll({ Type: 50004 })
            for edit in allEdits {
                cls := edit.ClassName
                name := edit.Name
                score := 0
                if InStr(cls, "ql-editor")
                    score += 3
                if InStr(cls, "new-input-ui")
                    score += 2
                if InStr(name, "Enter a prompt")
                    score += 3
                else if InStr(name, "prompt")
                    score += 2
                else if InStr(name, "Digite um prompt")
                    score += 2
                ; pick the highest scoring candidate
                if (score > bestScore) {
                    bestScore := score
                    best := edit
                }
            }
            if (bestScore >= 0) {
                promptField := best
            }
        }

        if (promptField) {
            promptField.SetFocus()
            Sleep 100
            ; Ensure focus was successful
            if (!promptField.HasKeyboardFocus) {
                ; Fallback: try clicking if SetFocus didn't work
                promptField.Click()
                Sleep 100
            }
            return true
        }
    } catch Error as e {
        ; If all else fails, silently fail (no fallback action defined)
    }
    return false
}

; Shift + P : Focus the prompt text field - Prompt
+p:: {
    FocusGeminiPromptField()
}

; Shift + C : Click the last Copy button (copies the preceding message) - Copy
+c:: {
    try {
        uia := UIA_Browser()
        Sleep 300

        ; Find all Copy buttons
        allCopyButtons := []

        ; Primary strategy: Find all buttons with Name "Copy"
        allButtons := uia.FindAll({ Type: 50000 })
        for button in allButtons {
            if (button.Name = "Copy" || InStr(button.Name, "Copy", false) = 1) {
                ; Additional check: ensure it has the Copy button className pattern
                if (InStr(button.ClassName, "icon-button") || InStr(button.ClassName, "mdc-button")) {
                    allCopyButtons.Push(button)
                }
            }
        }

        ; Fallback: Try by Type "Button" if the above didn't find enough
        if (allCopyButtons.Length = 0) {
            allButtons := uia.FindAll({ Type: "Button" })
            for button in allButtons {
                if (button.Name = "Copy" || InStr(button.Name, "Copy", false) = 1) {
                    allCopyButtons.Push(button)
                }
            }
        }

        if (allCopyButtons.Length = 0) {
            ; No Copy buttons found
            return
        }

        ; Find the last Copy button (the one with the highest Y position, meaning furthest down the page)
        lastCopyButton := 0
        highestY := -1

        for copyButton in allCopyButtons {
            try {
                btnPos := copyButton.Location
                btnBottomY := btnPos.y + btnPos.h

                ; The last button will be the one with the highest bottom Y coordinate
                if (btnBottomY > highestY) {
                    highestY := btnBottomY
                    lastCopyButton := copyButton
                }
            } catch {
                ; If getting location fails, skip this button
            }
        }

        ; If position-based approach didn't work, just use the last one in the array
        if (!lastCopyButton && allCopyButtons.Length > 0) {
            lastCopyButton := allCopyButtons[allCopyButtons.Length]
        }

        if (lastCopyButton) {
            lastCopyButton.Click()
        } else {
            ; Last resort: Could not find last Copy button
        }
    } catch Error as e {
        ; If all else fails, silently fail (no fallback action defined)
    }
}

; Shift + R : Read aloud the last message (click last "Show more options" then "Text to speech") - Read
+r:: {
    try {
        uia := UIA_Browser()
        Sleep 300

        ; Step 1: Find all "Show more options" buttons
        allMoreOptionsButtons := []

        ; Primary strategy: Find all buttons with Name "Show more options"
        allButtons := uia.FindAll({ Type: 50000 })
        for button in allButtons {
            if (button.Name = "Show more options" || InStr(button.Name, "Show more options", false) = 1) {
                ; Additional check: ensure it has the more-menu-button className pattern
                if (InStr(button.ClassName, "more-menu-button") || InStr(button.ClassName, "mdc-button")) {
                    allMoreOptionsButtons.Push(button)
                }
            }
        }

        ; Fallback: Try by Type "Button" if the above didn't find enough
        if (allMoreOptionsButtons.Length = 0) {
            allButtons := uia.FindAll({ Type: "Button" })
            for button in allButtons {
                if (button.Name = "Show more options" || InStr(button.Name, "Show more options", false) = 1) {
                    if (InStr(button.ClassName, "more-menu-button")) {
                        allMoreOptionsButtons.Push(button)
                    }
                }
            }
        }

        if (allMoreOptionsButtons.Length = 0) {
            ; No "Show more options" buttons found
            return
        }

        ; Find the last "Show more options" button (the one with the highest Y position, meaning furthest down the page)
        lastMoreOptionsButton := 0
        highestY := -1

        for moreOptionsButton in allMoreOptionsButtons {
            try {
                btnPos := moreOptionsButton.Location
                btnBottomY := btnPos.y + btnPos.h

                ; The last button will be the one with the highest bottom Y coordinate
                if (btnBottomY > highestY) {
                    highestY := btnBottomY
                    lastMoreOptionsButton := moreOptionsButton
                }
            } catch {
                ; If getting location fails, skip this button
            }
        }

        ; If position-based approach didn't work, just use the last one in the array
        if (!lastMoreOptionsButton && allMoreOptionsButtons.Length > 0) {
            lastMoreOptionsButton := allMoreOptionsButtons[allMoreOptionsButtons.Length]
        }

        if (!lastMoreOptionsButton) {
            ; Could not find last "Show more options" button
            return
        }

        ; Step 2: Click the last "Show more options" button
        lastMoreOptionsButton.Click()
        Sleep 400 ; Wait for menu to appear

        ; Step 3: Find and click the "Text to speech" menu item
        textToSpeechMenuItem := 0

        ; Primary strategy: Find by Name "Text to speech" with Type 50011 (MenuItem)
        textToSpeechMenuItem := uia.FindFirst({ Name: "Text to speech", Type: 50011 })

        ; Fallback 1: Try by Type "MenuItem" and Name "Text to speech"
        if !textToSpeechMenuItem {
            textToSpeechMenuItem := uia.FindFirst({ Type: "MenuItem", Name: "Text to speech" })
        }

        ; Fallback 2: Try by ClassName containing "mat-mdc-menu-item" (substring match)
        if !textToSpeechMenuItem {
            allMenuItems := uia.FindAll({ Type: 50011 })
            for menuItem in allMenuItems {
                if InStr(menuItem.Name, "Text to speech") || InStr(menuItem.Name, "speech") {
                    if InStr(menuItem.ClassName, "mat-mdc-menu-item") {
                        textToSpeechMenuItem := menuItem
                        break
                    }
                }
            }
        }

        ; Fallback 3: Try finding by Name with substring match (in case of localization variations)
        if !textToSpeechMenuItem {
            allMenuItems := uia.FindAll({ Type: 50011 })
            for menuItem in allMenuItems {
                if InStr(menuItem.Name, "Text to speech") || InStr(menuItem.Name, "Texto para fala") || InStr(
                    menuItem.Name,
                    "Ler em voz alta") {
                    if InStr(menuItem.ClassName, "mat-mdc-menu-item") {
                        textToSpeechMenuItem := menuItem
                        break
                    }
                }
            }
        }

        if (textToSpeechMenuItem) {
            textToSpeechMenuItem.Click()
        } else {
            ; Last resort: Could not find "Text to speech" menu item
        }
    } catch Error as e {
        ; If all else fails, silently fail (no fallback action defined)
    }
}

; Shift + G : Focus the prompt text field and send Gemini prompt text - Gemini
+g:: {
    try {
        uia := UIA_Browser()
        Sleep 300

        ; Primary strategy: Find by Name (Gemini updated placeholder in 2025)
        try
            promptField := uia.FindFirst({ Name: "Enter a prompt for Gemini", Type: 50004 })
        catch
            promptField := ""

        ; Fallback 1: Legacy name "Enter a prompt here"
        if !promptField {
            try
                promptField := uia.FindFirst({ Name: "Enter a prompt here", Type: 50004 })
            catch
                promptField := ""
        }

        ; Fallback 2: Try by ClassName containing "ql-editor" or "new-input-ui" (substring match)
        if !promptField {
            allEdits := uia.FindAll({ Type: 50004 })
            for edit in allEdits {
                if (InStr(edit.ClassName, "ql-editor") || InStr(edit.ClassName, "new-input-ui")) {
                    if InStr(edit.Name, "Enter a prompt") || InStr(edit.Name, "prompt") {
                        promptField := edit
                        break
                    }
                }
            }
        }

        ; Fallback 3: Try finding by ClassName containing "ql-editor" (most specific identifier)
        if !promptField {
            allEdits := uia.FindAll({ Type: 50004 })
            for edit in allEdits {
                if InStr(edit.ClassName, "ql-editor") {
                    promptField := edit
                    break
                }
            }
        }

        ; Fallback 4: Try finding by Name with substring match (in case of localization variations)
        if !promptField {
            allEdits := uia.FindAll({ Type: 50004 })
            for edit in allEdits {
                if InStr(edit.Name, "Enter a prompt") || InStr(edit.Name, "Digite um prompt") || InStr(edit.Name,
                    "prompt") {
                    ; Additional check to ensure it's the prompt field (has ql-editor in className)
                    if InStr(edit.ClassName, "ql-editor") {
                        promptField := edit
                        break
                    }
                }
            }
        }

        if (promptField) {
            promptField.SetFocus()
            Sleep 100
            ; Ensure focus was successful
            if (!promptField.HasKeyboardFocus) {
                ; Fallback: try clicking if SetFocus didn't work
                promptField.Click()
                Sleep 100
            }

            ; Read the Gemini_Prompt.txt file and paste its contents via clipboard
            promptFilePath := A_ScriptDir "\assets\data\Gemini_Prompt.txt"
            if FileExist(promptFilePath) {
                ; Save current clipboard
                oldClipboard := A_Clipboard
                try {
                    ; Read and set clipboard
                    promptText := FileRead(promptFilePath, "UTF-8")
                    if (promptText) {
                        A_Clipboard := promptText
                        ClipWait 1, 1  ; Wait for clipboard to be ready

                        ; Clear any existing text first (select all and delete)
                        Send "^a"
                        Sleep 50

                        ; Paste the text from clipboard
                        Send "^v"
                        Sleep 100

                        ; Restore original clipboard
                        A_Clipboard := oldClipboard

                        Sleep 400
                        Send "{Enter}"
                    }
                } catch Error as e {
                    ; If file reading fails, try to restore clipboard
                    try {
                        A_Clipboard := oldClipboard
                    }
                }
            } else {
                ; File not found - could show a message or just silently fail
            }
        } else {
            ; Last resort: Could not find prompt field
        }
    } catch Error as e {
        ; If all else fails, silently fail (no fallback action defined)
    }
}

; Shift + F : Click the Expand input to Fullscreen button
+f:: {
    try {
        uia := UIA_Browser()
        Sleep 300

        ; Primary strategy: Find by Name "Expand input to Fullscreen" with Type 50000 (Button)
        fullscreenButton := uia.FindFirst({ Name: "Expand input to Fullscreen", Type: 50000 })

        ; Fallback 1: Try by Type "Button" and Name "Expand input to Fullscreen"
        if !fullscreenButton {
            fullscreenButton := uia.FindFirst({ Type: "Button", Name: "Expand input to Fullscreen" })
        }

        ; Fallback 2: Try by ClassName containing "fullscreen-button" (substring match)
        if !fullscreenButton {
            allButtons := uia.FindAll({ Type: 50000 })
            for button in allButtons {
                if InStr(button.ClassName, "fullscreen-button") {
                    fullscreenButton := button
                    break
                }
            }
        }

        ; Fallback 3: Try finding by Name with substring match (in case of localization variations)
        if !fullscreenButton {
            allButtons := uia.FindAll({ Type: 50000 })
            for button in allButtons {
                if InStr(button.Name, "Expand input") || InStr(button.Name, "Fullscreen") || InStr(button.Name,
                    "Expandir") {
                    ; Additional check to ensure it's the fullscreen button (has fullscreen-button in className)
                    if InStr(button.ClassName, "fullscreen-button") {
                        fullscreenButton := button
                        break
                    }
                }
            }
        }

        if (fullscreenButton) {
            fullscreenButton.Click()
        } else {
            ; Last resort: Could not find fullscreen button
        }
    } catch Error as e {
        ; If all else fails, silently fail (no fallback action defined)
    }
}

; Enter : Send Enter and automatically monitor for response completion
; This automates the notification process so Ctrl+Enter is no longer needed
; Shift+Enter continues to work normally (for line breaks) as it's not intercepted
Enter:: {
    ; Only trigger if Shift or Ctrl are NOT pressed
    ; Shift+Enter = line break (should not trigger monitoring)
    ; Ctrl+Enter = handled by the ^Enter hotkey below
    if (GetKeyState("Shift", "P") || GetKeyState("Ctrl", "P")) {
        ; Pass through to let browser handle it normally
        Send "{Enter}"
        return
    }

    ; Send Enter key to submit the prompt
    Send "{Enter}"

    ; Phase 3: non-blocking daemon watch or legacy blocking monitor
    if (USE_DAEMON_MONITOR_GEMINI) {
        ShiftKeysIPC_StartGeminiWatch(300000, PlayCompletionChime_Gemini)
        return
    }
    WaitForStopResponseButton_Gemini()
}

; Control + Enter : Send Enter and monitor for response completion
^Enter:: {
    ; Send Enter key to submit the prompt
    Send "{Enter}"

    ; Phase 3: non-blocking daemon watch or legacy blocking monitor
    if (USE_DAEMON_MONITOR_GEMINI) {
        ShiftKeysIPC_StartGeminiWatch(300000, PlayCompletionChime_Gemini)
        return
    }
    WaitForStopResponseButton_Gemini()
}

; ---------------------------------------------------------------------------
; Monitor "Stop response" button and play chime when it disappears
; ---------------------------------------------------------------------------
WaitForStopResponseButton_Gemini(timeout := 300000) {
    ; Store Gemini window handle
    geminiHwnd := WinExist("A")
    if !geminiHwnd {
        return ; Gemini window not found
    }

    ; Obtain UIA context for Gemini window
    try {
        uia := UIA_Browser("ahk_id " geminiHwnd)
    } catch {
        return ; Failed to get UIA context
    }

    start := A_TickCount
    btn := ""
    buttonFound := false

    ; Wait for the "Stop response" button to appear
    deadline := (timeout > 0) ? (start + timeout) : 0
    while (timeout <= 0 || (A_TickCount < deadline)) {
        btn := ""

        ; Try to find the "Stop response" button
        try {
            btn := uia.FindFirst({ Type: "50000", Name: "Stop response" })
        } catch {
            btn := ""
        }

        if !btn {
            ; Fallback: Try by Type "Button" and Name "Stop response"
            try {
                btn := uia.FindFirst({ Type: "Button", Name: "Stop response" })
            } catch {
                btn := ""
            }
        }

        if !btn {
            ; Fallback: Try substring match for localization variations
            try {
                btn := uia.FindFirst({ Name: "Stop response", matchmode: "Substring" })
            } catch {
                btn := ""
            }
        }

        if btn {
            buttonFound := true
            ; Monitor the button until it disappears (with confirmation layer)
            while (timeout <= 0 || (A_TickCount < deadline)) {
                ; Monitor the button while it exists
                while btn && (timeout <= 0 || (A_TickCount < deadline)) {
                    Sleep 250
                    btn := ""

                    ; Check if button still exists
                    try {
                        btn := uia.FindFirst({ Type: "50000", Name: "Stop response" })
                    } catch {
                        btn := ""
                    }

                    if !btn {
                        try {
                            btn := uia.FindFirst({ Type: "Button", Name: "Stop response" })
                        } catch {
                            btn := ""
                        }
                    }

                    if !btn {
                        try {
                            btn := uia.FindFirst({ Name: "Stop response", matchmode: "Substring" })
                        } catch {
                            btn := ""
                        }
                    }
                }

                ; Button has disappeared - add confirmation layer
                ; Wait 1.5 seconds and check if it reappears (to avoid false positives)
                confirmationStart := A_TickCount
                confirmationPeriod := 1500  ; 1.5 seconds
                buttonReappeared := false

                ; Check multiple times during the confirmation period
                while ((A_TickCount - confirmationStart) < confirmationPeriod) && (timeout <= 0 || (A_TickCount <
                    deadline)) {
                    Sleep 300

                    ; Check if button reappeared
                    try {
                        btn := uia.FindFirst({ Type: "50000", Name: "Stop response" })
                    } catch {
                        btn := ""
                    }

                    if !btn {
                        try {
                            btn := uia.FindFirst({ Type: "Button", Name: "Stop response" })
                        } catch {
                            btn := ""
                        }
                    }

                    if !btn {
                        try {
                            btn := uia.FindFirst({ Name: "Stop response", matchmode: "Substring" })
                        } catch {
                            btn := ""
                        }
                    }

                    if btn {
                        ; Button reappeared - break out of confirmation loop and continue monitoring
                        buttonReappeared := true
                        break  ; Exit confirmation loop, will continue outer monitoring loop
                    }
                }

                ; If button didn't reappear during confirmation period, response is truly complete
                if !buttonReappeared {
                    break  ; Exit the outer monitoring loop
                }
                ; Otherwise, continue the outer loop to monitor the reappeared button
            }
            break
        }
        Sleep 250
    }

    ; Play chime when button disappears (only if we found it initially)
    if buttonFound {
        try {
            PlayCompletionChime_Gemini()
        } catch {
        }
    }
}

; ---------------------------------------------------------------------------
; Play completion chime for Gemini responses (debounced)
; ---------------------------------------------------------------------------
PlayCompletionChime_Gemini() {
    try {
        static lastTick := 0
        if (A_TickCount - lastTick < 1500)
            return
        lastTick := A_TickCount

        ScriptSoundPlay(A_ScriptDir . "\assets\sounds\gemini-completion.wav")
    } catch {
    }
}

#HotIf
