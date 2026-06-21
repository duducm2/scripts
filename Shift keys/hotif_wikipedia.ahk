; =============================================================================
; Shift keys module: hotif_wikipedia.ahk
; Wikipedia Chrome hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

#HotIf WinActive("ahk_exe chrome.exe") && InStr(WinGetTitle("A"), "Wikipedia", false)

; Shift + S: Focus the Wikipedia search field (prefer the field; if hidden, click the Search toggle first)
+s::
{
    try {
        ; Step 1: Always scroll to the beginning of the page first
        Send "^Home"
        Sleep 300 ; Give page time to scroll

        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 200 ; Give UIA time to attach

        ; Prefer document; fall back to browser root
        try {
            root := uia.GetCurrentDocumentElement()
        } catch {
            root := uia.BrowserElement
        }

        ; Step 2: Try to locate the "Search Wikipedia" field by name (combo box / edit)
        searchBox := 0

        ; First try: ComboBox with the expected name (Type 50003)
        try {
            searchBox := root.FindElement({ Type: 50003, Name: "Search Wikipedia", cs: false })
        } catch {
        }

        ; Second: Edit control with the same name (in case UI changes type)
        if (!searchBox) {
            try {
                searchBox := root.FindElement({ Type: 50004, Name: "Search Wikipedia", cs: false })
            } catch {
            }
        }

        ; Third: any element by name "Search Wikipedia"
        if (!searchBox) {
            try {
                searchBox := root.FindElement({ Name: "Search Wikipedia", cs: false })
            } catch {
            }
        }

        ; If we found the field, focus/click it and we're done.
        if (searchBox) {
            try {
                searchBox.SetFocus()
            } catch {
                searchBox.Click()
            }
            return
        }

        ; Step 3: If the field is not available, try clicking the "Search" toggle button/link first.
        searchToggle := 0

        ; Strategy 1: Try finding the search group first, then the button within it (most reliable)
        try {
            searchGroup := root.FindElement({ AutomationId: "p-search", cs: false })
            if (searchGroup) {
                try {
                    searchToggle := searchGroup.FindElement({ Type: 50005, Name: "Search", cs: false })
                } catch {
                }
                if (!searchToggle) {
                    try {
                        searchToggle := searchGroup.FindElement({ Type: 50005, Value: "https://en.wikipedia.org/wiki/Special:Search",
                            cs: false })
                    } catch {
                    }
                }
            }
        } catch {
        }

        ; Strategy 2: Search by Value (URL) directly from root
        if (!searchToggle) {
            try {
                searchToggle := root.FindElement({ Type: 50005, Value: "https://en.wikipedia.org/wiki/Special:Search",
                    cs: false })
            } catch {
            }
        }

        ; Strategy 3: Search by Type and Name (original method)
        if (!searchToggle) {
            try {
                searchToggle := root.FindElement({ Type: 50005, Name: "Search", cs: false })
            } catch {
            }
        }

        ; Strategy 4: Search by ControlType and Name
        if (!searchToggle) {
            try {
                searchToggle := root.FindElement({ ControlType: "Hyperlink", Name: "Search", cs: false })
            } catch {
            }
        }

        ; Strategy 5: Search for any link with the Special:Search URL
        if (!searchToggle) {
            try {
                searchToggle := root.FindElement({ Type: 50005, Value: "*Special:Search*", cs: false })
            } catch {
            }
        }

        ; Strategy 6: Try finding by LocalizedType "link" and Name
        if (!searchToggle) {
            try {
                searchToggle := root.FindElement({ LocalizedType: "link", Name: "Search", cs: false })
            } catch {
            }
        }

        if (searchToggle) {
            ; Validate element before clicking: check if it's offscreen or disabled
            isValid := true
            try {
                isOffscreen := searchToggle.GetPropertyValue(UIA.Property.IsOffscreen)
                if (isOffscreen) {
                    isValid := false
                }
            } catch {
                ; Continue even if property check fails
            }

            try {
                isEnabled := searchToggle.GetPropertyValue(UIA.Property.IsEnabled)
                if (!isEnabled) {
                    isValid := false
                }
            } catch {
                ; Continue even if property check fails
            }

            ; If element is invalid, try re-finding it
            if (!isValid) {
                Sleep 200
                searchToggle := 0
                try {
                    searchToggle := root.FindElement({ Type: 50005, Name: "Search", cs: false })
                } catch {
                }
                if (!searchToggle) {
                    try {
                        searchToggle := root.FindElement({ ControlType: "Hyperlink", Name: "Search", cs: false })
                    } catch {
                    }
                }
            }

            ; Try multiple click strategies if element is still valid
            if (searchToggle) {
                clicked := false

                ; Strategy 1: Try Invoke pattern (most reliable for links/buttons)
                try {
                    searchToggle.Invoke()
                    clicked := true
                } catch Error as invokeErr {
                }

                ; Strategy 2: Try SetFocus then Click
                if (!clicked) {
                    try {
                        searchToggle.SetFocus()
                        Sleep 50
                        searchToggle.Click()
                        clicked := true
                    } catch Error as focusClickErr {
                    }
                }

                ; Strategy 3: Try direct Click
                if (!clicked) {
                    try {
                        searchToggle.Click()
                        clicked := true
                    } catch Error as clickErr {
                    }
                }

                ; If all click strategies failed, fall back to accelerator
                if (!clicked) {
                    try {
                        uia.ControlSend("!f")
                    } catch {
                    }
                }

                ; Give the UI a moment to reveal the field, then try again to find it.
                Sleep 250

                searchBox := 0
                try {
                    searchBox := root.FindElement({ Type: 50003, Name: "Search Wikipedia", cs: false })
                } catch {
                }
                if (!searchBox) {
                    try {
                        searchBox := root.FindElement({ Type: 50004, Name: "Search Wikipedia", cs: false })
                    } catch {
                    }
                }
                if (!searchBox) {
                    try {
                        searchBox := root.FindElement({ Name: "Search Wikipedia", cs: false })
                    } catch {
                    }
                }

                if (searchBox) {
                    try {
                        searchBox.SetFocus()
                    } catch {
                        searchBox.Click()
                    }
                    return
                }
            }
        }

        ; Final fallback: use the accelerator key if all else fails.
        try {
            uia.ControlSend("!f")
            return
        } catch {
        }

        MsgBox "Could not find the 'Search Wikipedia' field."
    } catch Error as e {
        MsgBox "An error occurred: " e.Message
    }
}

; Shift + P: Save Wikipedia scroll position
+p::
{
    SaveWikipediaScrollPositionManually_ShiftKeys()
}

; Helper function to restore scroll position to a given percentage
; Returns true on success, false on failure
RestoreWikipediaScrollPosition(scrollPercentage, bannerText := "📜 Restoring scroll position... Please wait") {
    if (scrollPercentage <= 0.0 || scrollPercentage > 1.0) {
        return false
    }

    try {
        ; Create UIA_Browser once
        uia := UIA_Browser("ahk_exe chrome.exe")
        if (!uia) {
            return false
        }

        ; Show banner (foreground monitor)
        StandardLoadingBar_Show(bannerText, BANNER_ACCENT_INTERMEDIATE, { passive: true, centerOnHwnd: 0,
            textWidth: 500,
            fontSize: 17, passiveBgColor: BANNER_ACCENT_INTERMEDIATE })

        ; Block input during restoration (Phase 4: guaranteed cleanup in finally)
        BlockInput("On")
        try {
            ; Wait for page to be ready (condition-based, up to 500ms) instead of fixed Sleep(500)
            deadline := A_TickCount + 500
            docHeight := ""
            loop {
                docHeight := uia.JSReturnThroughClipboard("document.documentElement.scrollHeight")
                if (docHeight != "" && docHeight != "undefined" && docHeight != "null") {
                    try h := Float(docHeight)
                    catch {
                        h := 0
                    }
                    if (h > 0)
                        break
                }
                if (A_TickCount >= deadline)
                    break
                Sleep(50)
            }
            if (docHeight = "" || docHeight = "undefined" || docHeight = "null") {
                StandardLoadingBar_Hide(0)
                return false
            }
            docHeightFloat := Float(docHeight)
            if (docHeightFloat <= 0) {
                StandardLoadingBar_Hide(0)
                return false
            }

            ; Calculate and execute scroll
            targetScrollY := scrollPercentage * docHeightFloat
            uia.JSExecute("window.scrollTo(0, " . Round(targetScrollY) . ");")
            deadline2 := A_TickCount + 500
            while (A_TickCount < deadline2)
                Sleep(50)

            ; Update banner to show success
            try {
                StandardLoadingBar_Update("Scroll position restored!")
                Sleep(1000)
            } catch {
            }

            try {
                Sleep(500)
                StandardLoadingBar_Hide(0)
            } catch {
            }

            return true
        } catch Error as err {
            try StandardLoadingBar_Hide(0)
            catch {
            }
            return false
        } finally {
            BlockInput("Off")
        }
    } catch Error as err {
        BlockInput("Off")
        try StandardLoadingBar_Hide(0)
        catch {
        }
        return false
    }
}

; Helper function to normalize Wikipedia URLs
NormalizeWikipediaURL(url) {
    if (url = "" || !InStr(url, "wikipedia.org")) {
        return ""
    }
    ; Remove fragments and trailing slashes
    url := RegExReplace(url, "/#.*$", "")
    url := RegExReplace(url, "/+$", "")
    return url
}

; Helper function to get and normalize Wikipedia URL from active window
GetWikipediaURLNormalized() {
    try {
        if (!WinActive("ahk_exe chrome.exe") || !InStr(WinGetTitle("A"), "Wikipedia")) {
            return ""
        }
        uia := UIA_Browser("ahk_exe chrome.exe")
        url := uia.GetCurrentURL()
        return NormalizeWikipediaURL(url)
    } catch Error as err {
        return ""
    }
}

; Wikipedia scroll position save function (duplicated from AppLaunchers.ahk)
SaveWikipediaScrollPositionManually_ShiftKeys() {
    try {
        ; Check if Wikipedia window is currently active
        activeWindow := WinGetTitle("A")
        isChromeActive := WinActive("ahk_exe chrome.exe")
        hasWikipedia := InStr(activeWindow, "Wikipedia")
        if (!isChromeActive || !hasWikipedia) {
            return false
        }
    } catch Error as err {
        return false
    }

    ; Exit fullscreen before scroll position save (REQUIRED: UIA unreliable in fullscreen)
    Send("{F11}")
    Sleep(300)  ; Allow time for fullscreen exit (increased for reliability)

    ; Show banner to inform user that scroll position is being saved (foreground monitor)
    StandardLoadingBar_Show("💾 Saving scroll position... Please wait", BANNER_ACCENT_INTERMEDIATE, { passive: true,
        centerOnHwnd: 0,
        textWidth: 500, fontSize: 17, passiveBgColor: BANNER_ACCENT_INTERMEDIATE })
    fullscreenRestored := false  ; Track if we've re-entered fullscreen
    try {
        ; Get normalized Wikipedia URL
        url := GetWikipediaURLNormalized()
        if (url = "") {
            ; Re-enter fullscreen before returning
            Send("{F11}")
            Sleep(300)
            fullscreenRestored := true
            return false
        }

        ; Create UIA_Browser for getting scroll position
        uia := false
        try {
            uia := UIA_Browser("ahk_exe chrome.exe")
        } catch Error as uiaErr {
            try StandardLoadingBar_Hide(0)
            catch {
            }
            ; Re-enter fullscreen before returning
            Send("{F11}")
            Sleep(300)
            fullscreenRestored := true
            return false
        }

        if (!uia) {
            try StandardLoadingBar_Hide(0)
            catch {
            }
            ; Re-enter fullscreen before returning
            Send("{F11}")
            Sleep(300)
            fullscreenRestored := true
            return false
        }

        ; Wait for page to stabilize before measuring (critical for portrait orientation)
        ; Monitors 3 and 4 are portrait (1080x1920); layout may shift during measurement
        Sleep(500)  ; Brief stabilization wait

        ; Get scroll position with retry for stability
        scrollY := ""
        docHeight := ""
        scrollYRetries := 3
        lastScrollY := -1
        loop scrollYRetries {
            try {
                scrollY := uia.JSReturnThroughClipboard("window.pageYOffset")
                scrollYFloat := Float(scrollY)
                ; Check if scroll position is stable
                if (scrollYFloat = lastScrollY || lastScrollY = -1) {
                    if (scrollYFloat = lastScrollY) {
                        break  ; Stable, use this value
                    }
                    lastScrollY := scrollYFloat
                }
            } catch Error as scrollErr {
            }
            if (A_Index < scrollYRetries) {
                Sleep(200)  ; Wait between attempts
            }
        }

        ; Get document height to calculate percentage (with stability check)
        docHeightRetries := 3
        lastDocHeight := -1
        loop docHeightRetries {
            try {
                docHeight := uia.JSReturnThroughClipboard("document.documentElement.scrollHeight")
                docHeightFloat := Float(docHeight)
                ; Check if document height is stable
                if (docHeightFloat = lastDocHeight || lastDocHeight = -1) {
                    if (docHeightFloat = lastDocHeight) {
                        break  ; Stable, use this value
                    }
                    lastDocHeight := docHeightFloat
                }
            } catch Error as docErr {
            }
            if (A_Index < docHeightRetries) {
                Sleep(200)  ; Wait between attempts
            }
        }

        ; Convert to numbers and calculate percentage
        if (scrollY != "" && scrollY != "undefined" && scrollY != "null" && docHeight != "" && docHeight !=
            "undefined" && docHeight != "null") {
            scrollYFloat := Float(scrollY)
            docHeightFloat := Float(docHeight)
            if (scrollYFloat >= 0 && docHeightFloat > 0) {
                scrollPercentage := scrollYFloat / docHeightFloat
                ; Clamp to valid range
                if (scrollPercentage > 1.0) {
                    scrollPercentage := 1.0
                }

                ; Save to INI file
                scrollPositionsFile := A_ScriptDir "\assets\data\wikipedia_scroll_positions.ini"
                SplitPath(scrollPositionsFile, , &dir)
                if (dir != "" && !DirExist(dir)) {
                    DirCreate(dir)
                }
                ; Always add to history stack for "go back" functionality (independent of INI file save)
                global g_WikipediaScrollHistory
                ; Add current position to history (url, scrollPercentage)
                g_WikipediaScrollHistory.Push({ url: url, scrollPercentage: scrollPercentage })
                ; Limit history to last 10 positions to prevent memory issues
                if (g_WikipediaScrollHistory.Length > 10) {
                    g_WikipediaScrollHistory.RemoveAt(1)
                }

                ; Try to save to INI file (for persistence across sessions)
                saved := false
                try {
                    ; Read existing entries first (before deleting file) to preserve them
                    existingEntries := Map()
                    if (FileExist(scrollPositionsFile)) {
                        try {
                            ; Read all existing entries from the Positions section
                            fileContent := FileRead(scrollPositionsFile)
                            ; Parse INI format manually
                            inPositionsSection := false
                            loop parse fileContent, "`n", "`r" {
                                line := Trim(A_LoopField)
                                if (line = "[Positions]") {
                                    inPositionsSection := true
                                    continue
                                }
                                if (inPositionsSection && SubStr(line, 1, 1) = "[") {
                                    ; Hit another section, stop reading
                                    break
                                }
                                if (inPositionsSection && InStr(line, "=")) {
                                    pos := InStr(line, "=")
                                    key := Trim(SubStr(line, 1, pos - 1))
                                    value := Trim(SubStr(line, pos + 1))
                                    if (key != "" && value != "") {
                                        existingEntries[key] := value
                                    }
                                }
                            }
                        } catch {
                            ; If read fails, we'll just write the new entry
                        }
                    }

                    ; Update with new entry
                    existingEntries[url] := scrollPercentage

                    ; Delete file to recreate in UTF-8
                    if (FileExist(scrollPositionsFile)) {
                        try {
                            FileDelete(scrollPositionsFile)
                            Sleep(100)  ; Small delay to ensure file system updates
                        } catch {
                        }
                    }

                    ; Write all entries back in UTF-8 encoding
                    try {
                        ; Write UTF-8 BOM and section header
                        FileAppend("[Positions]`r`n", scrollPositionsFile, "UTF-8")
                        ; Write each entry
                        for key, value in existingEntries {
                            ; Escape special INI characters in key and value
                            escapedKey := StrReplace(key, "=", "`=")
                            escapedKey := StrReplace(escapedKey, ";", "`;")
                            escapedValue := StrReplace(value, "`n", "`;")
                            escapedValue := StrReplace(escapedValue, "`r", "")
                            FileAppend(escapedKey . "=" . escapedValue . "`r`n", scrollPositionsFile, "UTF-8")
                        }
                        saved := true
                    } catch {
                        ; Fallback to IniWrite if manual write fails
                        saved := IniWrite(scrollPercentage, scrollPositionsFile, "Positions", url)
                    }
                } catch Error as iniErr {
                    saved := false
                }

                if (saved) {
                    ; Update banner to show success
                    try {
                        StandardLoadingBar_Update("Scroll position saved!")
                        Sleep(1000)  ; Show success message for 1 second
                    } catch {
                    }
                    ; Re-enter fullscreen after successful save
                    Send("{F11}")
                    Sleep(300)
                    fullscreenRestored := true
                    return true
                } else {
                    ; INI save failed - show error message
                    try {
                        StandardLoadingBar_Update("Error: Save failed")
                        Sleep(2000)  ; Show error message
                    } catch {
                    }
                    ; Re-enter fullscreen even on failure
                    Send("{F11}")
                    Sleep(300)
                    fullscreenRestored := true
                    return false
                }
            }
        }
    } catch Error as err {
    } finally {
        ; Always hide the banner after save operation completes
        try {
            Sleep(500)  ; Brief delay before hiding
            StandardLoadingBar_Hide(0)
        } catch {
        }
        ; Re-enter fullscreen if we haven't already (e.g., if exception occurred or validation failed)
        if (!fullscreenRestored) {
            try {
                Send("{F11}")
                Sleep(300)
            } catch {
            }
        }
    }
    return false
}

; Restore previous scroll position from history
RestorePreviousWikipediaScrollPosition() {
    global g_WikipediaScrollHistory

    try {
        ; Check if Wikipedia window is currently active
        activeWindow := WinGetTitle("A")
        isChromeActive := WinActive("ahk_exe chrome.exe")
        hasWikipedia := InStr(activeWindow, "Wikipedia")
        if (!isChromeActive || !hasWikipedia) {
            return false
        }
    } catch Error as err {
        return false
    }

    ; If history is empty, try to fall back to INI file (for positions saved via activation restore or previous sessions)
    if (g_WikipediaScrollHistory.Length = 0) {
        ; Get current URL to load from INI
        try {
            url := GetWikipediaURLNormalized()
            if (url = "") {
                ; Show brief message that no history exists
                ShowCenteredOverlay_Utils("⚠ No previous scroll position found", 1500, BANNER_ACCENT_ERROR)
                return false
            }

            ; Load from INI file
            scrollPositionsFile := A_ScriptDir "\assets\data\wikipedia_scroll_positions.ini"
            savedPercentage := IniRead(scrollPositionsFile, "Positions", url, "0")
            savedPercentageFloat := Float(savedPercentage)

            if (savedPercentageFloat > 0.0) {
                ; Found a saved position in INI, restore it using helper function
                return RestoreWikipediaScrollPosition(savedPercentageFloat,
                    "Restoring previous scroll position... Please wait")
            } else {
                ; No saved position found in INI either
                ShowCenteredOverlay_Utils("⚠ No previous scroll position found", 1500, BANNER_ACCENT_ERROR)
                return false
            }
        } catch Error as err {
            ; Show brief message that no history exists
            ShowCenteredOverlay_Utils("⚠ No previous scroll position found", 1500, BANNER_ACCENT_ERROR)
            return false
        }
    }

    ; Get current URL to match with history
    try {
        url := GetWikipediaURLNormalized()
        ; #region agent log
        try {
            FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' . A_TickCount .
            ',"location":"Shift keys.ahk:3415","message":"Got normalized URL in restore","data":{"url":"' . url .
            '","urlLength":' . StrLen(url) . '},"sessionId":"debug-session","runId":"post-fix","hypothesisId":"F"}`n',
            DEBUG_LOG_PATH
        } catch {
        }
        ; #endregion
        if (url = "") {
            return false
        }

        ; Create UIA_Browser for checking scroll differences
        uia := UIA_Browser("ahk_exe chrome.exe")
    } catch Error as err {
        return false
    }

    ; Find the most recent previous position (not the current one)
    previousPosition := 0
    foundIndex := 0
    ; Search backwards through history to find a different position
    loop g_WikipediaScrollHistory.Length {
        idx := g_WikipediaScrollHistory.Length - A_Index + 1
        historyItem := g_WikipediaScrollHistory[idx]
        ; Check if this is a different position (different URL or different scroll percentage)
        if (historyItem.url != url) {
            ; Different article, use this one
            previousPosition := historyItem
            foundIndex := idx
            break
        } else {
            ; Same article, check if scroll position is different
            currentScroll := uia.JSReturnThroughClipboard("window.pageYOffset")
            docHeight := uia.JSReturnThroughClipboard("document.documentElement.scrollHeight")
            if (currentScroll != "" && currentScroll != "undefined" && currentScroll != "null" &&
                docHeight != "" && docHeight != "undefined" && docHeight != "null") {
                currentScrollFloat := Float(currentScroll)
                docHeightFloat := Float(docHeight)
                if (docHeightFloat > 0) {
                    currentPercentage := currentScrollFloat / docHeightFloat
                    diff := Abs(currentPercentage - historyItem.scrollPercentage)
                    ; If the saved percentage is different from current, use it
                    if (diff > 0.01) {
                        previousPosition := historyItem
                        foundIndex := idx
                        break
                    }
                }
            }
        }
    }

    if (!previousPosition) {
        ; No different position found in history
        ShowCenteredOverlay_Utils("⚠ No previous scroll position found", 1500, BANNER_ACCENT_ERROR)
        return false
    }

    ; Restore scroll position using helper function
    success := RestoreWikipediaScrollPosition(previousPosition.scrollPercentage,
        "Restoring previous scroll position... Please wait")

    ; Remove the restored position from history (since we just used it) and all positions after it
    if (success && foundIndex > 0 && foundIndex <= g_WikipediaScrollHistory.Length) {
        ; Remove from foundIndex to end
        loop (g_WikipediaScrollHistory.Length - foundIndex + 1) {
            g_WikipediaScrollHistory.Pop()
        }
    }
    return true
}

#HotIf
