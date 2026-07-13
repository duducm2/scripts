; =============================================================================
; Shift keys module: gemini_chrome_01.ahk
; Gemini Chrome hotkeys (part 1)
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

#HotIf WinActive("ahk_exe chrome.exe") && InStr(WinGetTitle("A"), "gemini", false)

; Global state for Gemini drawer (main menu) – mirrors the state‑based toggle pattern
isGeminiDrawerOpen := false

; Global state for Gemini active model (3.1 Flash-Lite, 3.5 Flash, or 3.1 Pro)
isGeminiFastModel := "3.1 Flash-Lite"

; Global variables for Gemini model selector wizard menu
global g_GeminiModelSelectorGui := false
global g_GeminiModelSelectorActive := false
global g_GeminiModelHotkeyHandlers := []
global g_GeminiModelCharSequence := ["1", "2", "3", "4"]
global g_GeminiModels := [{ name: "3.1 Flash-Lite", description: "Fastest answers" }, { name: "3.5 Flash",
    description: "All-around help" }, { name: "3.1 Pro",
        description: "Advanced math and code" }, { name: "Thinking level", description: "Open thinking submenu (set level manually)" }
]

; Shift + D : Toggle the Main menu button (drawer) using fast state-based pattern
; $ + KeyWait — avoid leaking "d" into the prompt while UIA runs
$+d:: {
    KeyWait "d", "T1"
    ToggleGeminiDrawer()
}

; ---------------------------------------------------------------------------
ToggleGeminiDrawer() {
    global isGeminiDrawerOpen

    try {
        uia := UIA_Browser()
        if !IsObject(uia) {
            ; If we can't attach to Chrome, fall back to Escape like before
            SendEscape()
            return
        }

        ; Small settle time only – keep this snappy
        Sleep 100

        ; Exact-name regex (case-insensitive, anchored) with simple localization variant
        ; Same button toggles both open/close, we just keep state on our side
        mainMenuPattern := "i)^(Main menu|Menu principal)$"

        ; Use WaitForButton for robust, fast matching (similar to ToggleVoiceMessage)
        btn := WaitForButton(uia, mainMenuPattern, 1500)

        if (btn) {
            try {
                btn.Click()
            } catch Error as err {
                ; Try again in case of transient UIA glitch
                try {
                    btn.Click()
                } catch Error as err2 {
                    ; Swallow secondary failure – nothing else to do
                }
            }

            ; Flip our state after successful click
            isGeminiDrawerOpen := !isGeminiDrawerOpen

            ; Give Gemini a brief moment to redraw the drawer
            Sleep 200
        } else {
            ; If we couldn't find the button at all, keep behavior similar to previous version
            SendEscape()
        }
    } catch Error as e {
        ; If anything goes wrong, graceful fallback
        SendEscape()
    }
}

; Shift + N : New chat in Gemini (sends Ctrl-Shift-O)
$+n:: {
    KeyWait "n", "T1"
    Send "^+o"
}

; Shift + S : Click the Search button - Search
$+s:: {
    KeyWait "s", "T1"
    try {
        uia := UIA_Browser()
        Sleep 300

        ; Primary strategy: Find by Name "Search" with Type 50000 (Button)
        searchButton := uia.FindFirst({ Name: "Search", Type: 50000 })

        ; Fallback 1: Try by Type "Button" and Name "Search"
        if !searchButton {
            searchButton := uia.FindFirst({ Type: "Button", Name: "Search" })
        }

        ; Fallback 2: Try by ClassName containing "search-button" (substring match)
        if !searchButton {
            allButtons := uia.FindAll({ Type: 50000 })
            for button in allButtons {
                if InStr(button.ClassName, "search-button") && InStr(button.Name, "Search") {
                    searchButton := button
                    break
                }
            }
        }

        ; Fallback 3: Try finding by Name with substring match (in case of localization variations)
        if !searchButton {
            allButtons := uia.FindAll({ Type: 50000 })
            for button in allButtons {
                if InStr(button.Name, "Search") || InStr(button.Name, "Pesquisar") || InStr(button.Name,
                    "Buscar") {
                    ; Additional check to ensure it's the search button (has search-button in className)
                    if InStr(button.ClassName, "search-button") {
                        searchButton := button
                        break
                    }
                }
            }
        }

        if (searchButton) {
            searchButton.Click()
        } else {
            ; Last resort: Could try keyboard navigation if Gemini has a keyboard shortcut for search
            ; For now, we'll just not do anything if we can't find the button
        }
    } catch Error as e {
        ; If all else fails, silently fail (no fallback action defined)
    }
}

; Shift + M : Show model selector wizard menu (Fast, Thinking, Pro) - Model
$+m:: {
    KeyWait "m", "T1"
    ShowGeminiModelSelector()
}

; ---------------------------------------------------------------------------
; Cleanup function for Gemini model selector
CleanupGeminiModelSelector() {
    global g_GeminiModelSelectorActive, g_GeminiModelSelectorGui, g_GeminiModelHotkeyHandlers

    ; Disable active flag
    g_GeminiModelSelectorActive := false

    ; Disable all character hotkeys
    for handler in g_GeminiModelHotkeyHandlers {
        try {
            char := handler.char
            Hotkey(char, "Off")
            ; Also disable uppercase for numbers (though they're already uppercase)
            if (RegExMatch(char, "^[1-9]$")) {
                Hotkey(char, "Off")
            }
        } catch {
            ; Silently ignore errors
        }
    }

    ; Disable Escape hotkey
    try {
        Hotkey("Escape", "Off")
    } catch {
        ; Ignore
    }

    ; Clear handlers array
    g_GeminiModelHotkeyHandlers := []

    ; Close and destroy GUI
    if (IsObject(g_GeminiModelSelectorGui)) {
        try {
            g_GeminiModelSelectorGui.Destroy()
        } catch {
            ; Ignore
        }
        g_GeminiModelSelectorGui := false
    }
}

; ---------------------------------------------------------------------------
; Handler for Escape key in Gemini model selector
HandleGeminiModelSelectorEscape(*) {
    global g_GeminiModelSelectorActive
    if (g_GeminiModelSelectorActive) {
        CleanupGeminiModelSelector()
    }
}

; ---------------------------------------------------------------------------
; Factory function to create a handler that properly captures the model name
CreateGeminiModelCharHandler(char) {
    return (*) => HandleGeminiModelSelection(char)
}

; ---------------------------------------------------------------------------
; Read-only check: picker button name only (no second open/click of mode picker).
VerifyGeminiModelSelectedInOpenPicker(expectedModel, geminiHwnd := 0) {
    exp := GeminiNormalizeModelLabel(expectedModel)
    if (exp = "")
        return false
    uia := GeminiAttachBrowser(geminiHwnd, true)
    if !IsObject(uia)
        return false
    return GetGeminiActiveModelFromPickerOnly(uia) = exp
}

; ---------------------------------------------------------------------------
; Handler for model selection in Gemini model selector
HandleGeminiModelSelection(char) {
    global g_GeminiModelSelectorActive, g_GeminiModels, g_GeminiModelCharSequence
    global isGeminiFastModel

    ; Only process if selector is active
    if (!g_GeminiModelSelectorActive) {
        return
    }

    ; Get model index from character (1-based array index)
    modelIndex := -1
    for idx, ch in g_GeminiModelCharSequence {
        if (ch = char) {
            modelIndex := idx
            break
        }
    }

    if (modelIndex < 1 || modelIndex > g_GeminiModels.Length) {
        return
    }

    ; Get model info
    modelInfo := g_GeminiModels[modelIndex]
    if (!IsObject(modelInfo)) {
        return
    }
    try {
        modelName := modelInfo.name
    } catch {
        return
    }
    if (modelName = "") {
        return
    }

    ; Cleanup selector first (closes GUI, disables hotkeys)
    try {
        CleanupGeminiModelSelector()
    } catch {
        ; Ignore cleanup errors
    }

    try {
        SetTitleMatchMode(2)
        geminiHwnd := FindGeminiChromeHwnd()

        if (geminiHwnd) {
            if (!WinExist("ahk_id " geminiHwnd)) {
                ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
                return
            }
            WinActivate("ahk_id " geminiHwnd)
            if !WinWaitActive("ahk_id " geminiHwnd, , 1) {
                return
            }
        } else {
            ; Fallback: try to activate any Chrome window
            if (!WinExist("ahk_exe chrome.exe")) {
                ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
                return
            }
            WinActivate("ahk_exe chrome.exe")
            if !WinWaitActive("ahk_exe chrome.exe", , 2) {
                return
            }
        }

        StandardLoadingBar_Show("🔄 Switching Gemini model…", BANNER_ACCENT_INTERMEDIATE, { centerOnHwnd: geminiHwnd })
        try {
            if (modelName = "Thinking level") {
                StandardLoadingBar_Update("🔄 Opening Thinking level…", BANNER_ACCENT_INTERMEDIATE)
                opened := EnsureGeminiThinkingLevelMenuOpen(geminiHwnd)
                if (opened) {
                    StandardLoadingBar_Update("✅ Thinking level opened — set level manually",
                        BANNER_ACCENT_INTERMEDIATE)
                    StandardLoadingBar_Hide(700)
                } else {
                    StandardLoadingBar_Hide(0)
                    ShowCenteredOverlay_Utils("❌ Could not open Thinking level menu", 2800, BANNER_ACCENT_ERROR
                    )
                }
            } else {
                verified := EnsureGeminiModelViaMenu(modelName, geminiHwnd)
                if (!verified) {
                    StandardLoadingBar_Update("🔄 Model not confirmed. Retrying…",
                        BANNER_ACCENT_INTERMEDIATE)
                    verified := EnsureGeminiModelViaMenu(modelName, geminiHwnd)
                }

                if (verified) {
                    isGeminiFastModel := modelName
                    StandardLoadingBar_Update("✅ " . modelName . " model verified", BANNER_ACCENT_INTERMEDIATE)
                    try FocusGeminiPromptField()
                    StandardLoadingBar_Hide(700)
                } else {
                    StandardLoadingBar_Hide(0)
                    ShowCenteredOverlay_Utils("❌ Could not confirm Gemini model: " . modelName, 2800,
                        BANNER_ACCENT_ERROR)
                }
            }
        } finally {
            ; Ensure we never leave a stuck banner
            try StandardLoadingBar_Hide(0)
        }
    } catch Error as err {
        ; Silently fail if anything goes wrong
    }
}

; ---------------------------------------------------------------------------
; Show Gemini model selector wizard menu
ShowGeminiModelSelector() {
    global g_GeminiModelSelectorGui, g_GeminiModelSelectorActive, g_GeminiModelHotkeyHandlers
    global g_GeminiModels, g_GeminiModelCharSequence

    ; Verify global variables are initialized
    if (!IsObject(g_GeminiModels) || g_GeminiModels.Length = 0) {
        return
    }

    if (!IsObject(g_GeminiModelCharSequence) || g_GeminiModelCharSequence.Length = 0) {
        return
    }

    ; Close existing GUI if open
    if (g_GeminiModelSelectorActive && IsObject(g_GeminiModelSelectorGui)) {
        CleanupGeminiModelSelector()
        Sleep 50
    }

    ; Get monitor dimensions
    activeWin := 0
    try {
        activeWin := WinGetID("A")
    } catch {
        activeWin := 0
    }

    ; Default to primary monitor work area
    MonitorGetWorkArea(1, &monitorLeft, &monitorTop, &monitorRight, &monitorBottom)
    monitorWidth := monitorRight - monitorLeft
    monitorHeight := monitorBottom - monitorTop

    ; If we have an active window, find which monitor contains its center
    if (activeWin && activeWin != 0) {
        rect := Buffer(16, 0)
        if (DllCall("GetWindowRect", "ptr", activeWin, "ptr", rect)) {
            winLeft := NumGet(rect, 0, "int")
            winTop := NumGet(rect, 4, "int")
            winRight := NumGet(rect, 8, "int")
            winBottom := NumGet(rect, 12, "int")

            centerX := winLeft + (winRight - winLeft) // 2
            centerY := winTop + (winBottom - winTop) // 2

            monitorCount := MonitorGetCount()
            loop monitorCount {
                idx := A_Index
                MonitorGetWorkArea(idx, &l, &t, &r, &b)
                if (centerX >= l && centerX <= r && centerY >= t && centerY <= b) {
                    monitorLeft := l
                    monitorTop := t
                    monitorRight := r
                    monitorBottom := b
                    monitorWidth := r - l
                    monitorHeight := b - t
                    break
                }
            }
        }
    }

    ; Create GUI
    g_GeminiModelSelectorGui := Gui("+AlwaysOnTop +ToolWindow +E0x08000000", "Gemini Model Selector")
    fontSize := (monitorHeight < 800) ? 9 : 10
    g_GeminiModelSelectorGui.SetFont("s" . fontSize, "Segoe UI")
    g_GeminiModelSelectorGui.MarginX := 10
    g_GeminiModelSelectorGui.MarginY := 5

    ; Build display text
    displayText := "Gemini Model Selector`n`n"
    charIndex := 0
    for model in g_GeminiModels {
        if (charIndex < g_GeminiModelCharSequence.Length) {
            char := g_GeminiModelCharSequence[charIndex + 1]
            displayText .= "[" . char . "] " . model.name . " - " . model.description . "`n"
            charIndex++
        }
    }

    ; Calculate GUI size
    baseWidth := 400
    textControlHeight := Min(400, (g_GeminiModels.Length * 25) + 60)
    textControlWidth := baseWidth - 20

    ; Add text control with display
    g_GeminiModelSelectorGui.AddEdit("w" . textControlWidth . " h" . textControlHeight . " ReadOnly VScroll",
        displayText)

    ; Add Close button
    closeBtn := g_GeminiModelSelectorGui.AddButton("w100 Default Center", "Close")
    closeBtn.OnEvent("Click", (*) => CleanupGeminiModelSelector())

    ; Calculate total height
    totalHeight := 10 + textControlHeight + 40 + 10
    guiWidth := baseWidth

    ; Calculate center position
    marginX := 20
    marginY := 20
    guiX := monitorLeft + (monitorWidth - guiWidth) // 2
    guiY := monitorTop + (monitorHeight - totalHeight) // 2

    ; Ensure GUI stays within monitor bounds
    if (guiX < monitorLeft + marginX)
        guiX := monitorLeft + marginX
    if (guiY < monitorTop + marginY)
        guiY := monitorTop + marginY
    if (guiX + guiWidth > monitorLeft + monitorWidth - marginX)
        guiX := monitorLeft + monitorWidth - guiWidth - marginX
    if (guiY + totalHeight > monitorTop + monitorHeight - marginY)
        guiY := monitorTop + monitorHeight - totalHeight - marginY

    ; Show GUI centered on the active window's monitor
    try {
        g_GeminiModelSelectorGui.Show("NA w" . guiWidth . " h" . totalHeight . " x" . guiX . " y" . guiY)

        ; Small delay to ensure GUI is actually visible
        Sleep 50
    } catch Error as e {
        return
    }

    ; Set active flag
    g_GeminiModelSelectorActive := true

    ; Clear handlers array
    g_GeminiModelHotkeyHandlers := []

    ; Enable hotkeys for all assigned characters
    charIndex := 0
    for model in g_GeminiModels {
        if (charIndex < g_GeminiModelCharSequence.Length) {
            char := g_GeminiModelCharSequence[charIndex + 1]

            ; Create handler
            handler := CreateGeminiModelCharHandler(char)

            ; Store handler for cleanup
            g_GeminiModelHotkeyHandlers.Push({ char: char, handler: handler })

            ; Enable hotkey
            try {
                Hotkey(char, handler, "On")
            } catch {
                ; Silently ignore if we can't create hotkey
            }

            charIndex++
        }
    }

    ; Enable Escape hotkey
    Hotkey("Escape", HandleGeminiModelSelectorEscape, "On")
}
