; =============================================================================
; Shift keys module: gemini_chrome_01.ahk
; Gemini Chrome hotkeys (part 1)
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

#HotIf WinActive("ahk_exe chrome.exe") && IsConsumerGeminiChromeTitle(WinGetTitle("A"))

; Global state for Gemini drawer (main menu) – mirrors the state‑based toggle pattern
isGeminiDrawerOpen := false

; Last known Gemini model name (updated by AiCompanionModels_ApplyGemini)
isGeminiFastModel := "3.1 Flash-Lite"

; Shift + D : Toggle the Main menu button (drawer) using fast state-based pattern
+d:: {
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
+n:: {
    Send "^+o"
}

; Shift + S : Click the Search button - Search
+s:: {
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

; Shift + M : Select Deep model (INI)
+m:: {
    try AiCompanionModels_SelectRole(AI_COMPANION_GEMINI, "deep")
    catch {
    }
}

; Shift + Q : Select Fast / Quick model (INI)
+q:: {
    try AiCompanionModels_SelectRole(AI_COMPANION_GEMINI, "fast")
    catch {
    }
}

; Shift + L : Model list manager (1-9/letters, a add, r remove, f Fast, d Deep)
+l:: {
    try ShowAiCompanionModelSelector(AI_COMPANION_GEMINI)
    catch {
    }
}
