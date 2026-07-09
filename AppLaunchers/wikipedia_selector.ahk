; =============================================================================
; AppLaunchers module: wikipedia_selector.ahk
; Wikipedia selector GUI and char handlers
; Extracted verbatim from AppLaunchers.ahk; loaded via #include into the
; AppLaunchers.ahk process, which remains the entry point / source of truth.
; =============================================================================

; Handler for character key press
HandleWikipediaChar(char) {
    global g_WikipediaSelectorActive, g_WikipediaItems

    ; Only process if selector is active
    if (!g_WikipediaSelectorActive) {
        return
    }

    ; Find the item for this character
    item := ""
    for i, itm in g_WikipediaItems {
        if (itm.char = char) {
            item := itm
            break
        }
    }

    if (item) {
        ; Cleanup first (closes GUI, disables hotkeys)
        CleanupWikipediaSelector()

        ; If item has a URL, open it in Chrome in a new window
        if (item.url != "") {
            Run "chrome.exe --new-window " item.url
            ; Wait for the window to appear and become active
            WinWait("ahk_exe chrome.exe", , 5)
            Sleep(500)  ; Give the page a moment to start loading

            ; Wait for the page to load (check for Wikipedia in title)
            SetTitleMatchMode 2
            if (!WinWait("Wikipedia", , 10)) {
                ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
                return
            }
            WinActivate("Wikipedia")
            WinWaitActive("Wikipedia", , 5)

            ; Check if page is ready by attempting to get URL
            ; This ensures the page has loaded before proceeding
            ; For new windows, we need more time for the page to fully load
            pageReady := false
            loadingRetries := 8  ; Increased retries for new windows
            loop loadingRetries {
                try {
                    url := GetWikipediaURL()
                    if (url != "" && InStr(url, "wikipedia.org")) {
                        ; Verify the page is actually interactive, not just loaded
                        ; Try to access UIA to ensure the page is ready for automation
                        try {
                            testUia := UIA_Browser("ahk_exe chrome.exe")
                            if (testUia) {
                                ; Try to get document height to verify page is fully interactive
                                ; This is the same operation we'll need for scroll restoration
                                testDocHeight := testUia.JSReturnThroughClipboard(
                                    "document.documentElement.scrollHeight")
                                if (testDocHeight != "" && testDocHeight != "undefined" && testDocHeight != "null") {
                                    testHeightFloat := Float(testDocHeight)
                                    if (testHeightFloat > 0) {
                                        ; Page is ready and UIA can access it
                                        pageReady := true
                                        break
                                    }
                                }
                            }
                        } catch {
                            ; UIA not ready yet, continue waiting
                        }
                    }
                } catch {
                    ; URL not accessible yet, page may still be loading
                }
                if (A_Index < loadingRetries) {
                    Sleep(500)  ; Longer wait for new windows
                }
            }

            ; If page wasn't ready after retries, wait a bit more for loading
            ; For new windows, we need extra time for all resources to load
            if (!pageReady) {
                Sleep(1000)  ; Additional delay for new window page loading
            } else {
                ; Even if page seems ready, give it a moment for layout to stabilize
                Sleep(800)  ; Additional stabilization time for new windows
            }

            ; Enter fullscreen mode once page is ready
            Send("{F11}")
            Sleep(300)  ; Allow time for fullscreen transition (increased for new windows)

            ; Enable focus mode to darken other monitors
            EnableFocusMode()

            ; Start monitoring Wikipedia focus for automatic blackout cancellation
            StartWikipediaFocusMonitor()

            ; Try to restore scroll position (only on configured portrait monitors: 3 and 4)
            restoreBanner := ""
            try {
                if (!IsWindowOnWikipediaScrollRestoreMonitor()) {
                    return
                }
                savedPercentage := LoadWikipediaScrollPosition(item.url)
                if (savedPercentage > 0.0) {
                    ; Exit fullscreen before scroll restoration (REQUIRED: UIA unreliable in fullscreen)
                    Send("{F11}")
                    Sleep(300)  ; Allow time for fullscreen exit

                    StandardLoadingBar_Show("📜 Restoring scroll position... Please wait", BANNER_ACCENT_INTERMEDIATE)
                    AL_InstallInputGuard()

                    ; Initialize UIA_Browser with retry logic
                    ; For new windows, UIA needs more time to initialize and attach to the browser
                    uia := false
                    uiaRetries := 5  ; Increased retries for new windows
                    loop uiaRetries {
                        try {
                            uia := UIA_Browser("ahk_exe chrome.exe")
                            if (uia) {
                                ; Verify UIA can actually access the page (not just initialized)
                                ; Try a simple operation to ensure the connection is ready
                                try {
                                    testUrl := uia.GetCurrentURL()
                                    if (testUrl != "" && InStr(testUrl, "wikipedia.org")) {
                                        ; UIA is ready and can access the page
                                        break
                                    }
                                } catch {
                                    ; UIA initialized but not ready yet, continue retrying
                                    uia := false
                                }
                            }
                        } catch Error as uiaErr {
                            ; UIA initialization failed, will retry
                        }
                        if (A_Index < uiaRetries) {
                            Sleep(800)  ; Longer wait for new windows (UIA initialization takes time)
                        }
                    }

                    if (!uia) {
                        AL_RemoveInputGuard()
                        StandardLoadingBar_Update("Error: Could not access browser")
                        StandardLoadingBar_Hide(2000)
                        Send("{F11}")
                        Sleep(300)
                        return
                    }

                    ; Wait longer for page to be ready and stabilize (critical for portrait orientation)
                    ; For new windows, the page needs more time to fully render and become interactive
                    ; Portrait monitors (3 and 4, 1080x1920) can cause layout shifts that affect document height
                    Sleep(2500)  ; Increased wait for new window page stabilization

                    ; Get current document height with retry logic and stabilization
                    ; Portrait monitors (3 and 4, 1080x1920): ensure layout is stable
                    ; For new windows, we need more retries and longer waits
                    docHeight := ""
                    docHeightRetries := 8  ; Increased retries for new windows
                    lastDocHeight := 0
                    stableCount := 0
                    loop docHeightRetries {
                        try {
                            docHeight := uia.JSReturnThroughClipboard("document.documentElement.scrollHeight")
                            if (docHeight != "" && docHeight != "undefined" && docHeight != "null") {
                                docHeightFloat := Float(docHeight)
                                ; For new windows, require 3 consecutive stable readings (more strict)
                                if (docHeightFloat = lastDocHeight) {
                                    stableCount++
                                    if (stableCount >= 3) {
                                        ; Document height is stable, use it
                                        break
                                    }
                                } else {
                                    stableCount := 0
                                    lastDocHeight := docHeightFloat
                                }
                            }
                        } catch Error as docErr {
                            if (A_Index < docHeightRetries) {
                                Sleep(600)  ; Longer wait for new windows
                            }
                        }
                        if (A_Index < docHeightRetries) {
                            Sleep(400)  ; Longer wait between measurements for new windows
                        }
                    }

                    if (docHeight = "" || docHeight = "undefined" || docHeight = "null") {
                        AL_RemoveInputGuard()
                        StandardLoadingBar_Update("Error: Page not ready")
                        StandardLoadingBar_Hide(2000)
                        Send("{F11}")
                        Sleep(300)
                        return
                    }

                    docHeightFloat := Float(docHeight)
                    if (docHeightFloat <= 0) {
                        AL_RemoveInputGuard()
                        StandardLoadingBar_Update("Error: Invalid page height")
                        StandardLoadingBar_Hide(2000)
                        Send("{F11}")
                        Sleep(300)
                        return
                    }

                    ; Execute scroll restoration
                    ; For portrait orientation (1080x1920 on monitors 3 and 4), use precise calculation
                    targetScrollY := savedPercentage * docHeightFloat
                    try {
                        ; Use precise scrolling for portrait orientation (requires precise pixel positioning)
                        scrollCommand := "window.scrollTo({top: " . targetScrollY . ", behavior: 'instant'});"
                        uia.JSExecute(scrollCommand)
                        Sleep(1000)  ; Increased wait for portrait orientation (layout may need more time)

                        ; Verify scroll position was applied correctly with retry
                        ; Portrait orientation may require multiple verification attempts
                        verificationRetries := 3
                        actualScrollYFloat := 0
                        scrollDiff := 999999
                        loop verificationRetries {
                            try {
                                actualScrollY := uia.JSReturnThroughClipboard("window.pageYOffset")
                                actualScrollYFloat := Float(actualScrollY)
                                scrollDiff := Abs(actualScrollYFloat - targetScrollY)
                                ; If difference is small enough (within 2 pixels for portrait), consider it successful
                                if (scrollDiff <= 2.0) {
                                    break
                                }
                            } catch {
                            }
                            if (A_Index < verificationRetries) {
                                Sleep(300)  ; Wait before retry
                            }
                        }

                        ; If scroll is significantly off, try to correct it
                        if (scrollDiff > 5.0) {
                            ; Re-scroll to correct position
                            uia.JSExecute(scrollCommand)
                            Sleep(500)
                            ; Verify again
                            try {
                                actualScrollY := uia.JSReturnThroughClipboard("window.pageYOffset")
                                actualScrollYFloat := Float(actualScrollY)
                                scrollDiff := Abs(actualScrollYFloat - targetScrollY)
                            } catch {
                            }
                        }

                        AL_RemoveInputGuard()
                        StandardLoadingBar_Update("Scroll position restored!")
                        StandardLoadingBar_Hide(1000)

                        ; Re-enter fullscreen after successful scroll restoration
                        Send("{F11}")
                        Sleep(300)  ; Allow time for fullscreen transition
                    } catch Error as scrollErr {
                        AL_RemoveInputGuard()
                        StandardLoadingBar_Update("Error: Scroll failed")
                        StandardLoadingBar_Hide(2000)
                        Send("{F11}")
                        Sleep(300)
                    }
                }
            } catch Error as err {
                AL_RemoveInputGuard()
                StandardLoadingBar_Update("Error: " . SubStr(err.Message, 1, 50))
                StandardLoadingBar_Hide(2000)
                ; Re-enter fullscreen after error
                try {
                    Send("{F11}")
                    Sleep(300)
                } catch {
                }
            }
        }
        ; Items 2-5 have no URL, so no action is taken
    }
}

; Factory function to create a handler that properly captures the character
CreateWikipediaCharHandler(char) {
    ; Return a function that captures the char value at creation time
    return (*) => HandleWikipediaChar(char)
}

; Used with g_OnEscapePressed so Utils_GlobalEscapeHandler (I10) closes the modal when *Escape (I0) never fires.
WikipediaSelector_GlobalEscapeCallback(*) {
    global g_WikipediaSelectorActive
    if (!g_WikipediaSelectorActive)
        return false
    WikipediaSelector_Cancel()
    return true
}

; GUI message-path Escape (helps when Esc is delivered via the GUI message pump)
WikipediaSelector_GuiEscape(*) {
    WikipediaSelector_Cancel()
}

; Cancel Wikipedia selector (same pattern as AiModelSelector_Cancel in Utils.ahk)
WikipediaSelector_Cancel(*) {
    CleanupWikipediaSelector()
}

; Load completed Wikipedia articles from CSV file
LoadCompletedArticles() {
    global g_WikipediaCompletedFile
    completedArticles := []

    try {
        if (!FileExist(g_WikipediaCompletedFile)) {
            return completedArticles
        }

        fileContent := FileRead(g_WikipediaCompletedFile)
        lines := StrSplit(fileContent, "`n")

        ; Skip header line and process each line
        loop lines.Length {
            if (A_Index = 1) {
                continue  ; Skip header
            }

            line := Trim(lines[A_Index])
            if (line != "") {
                completedArticles.Push(line)
            }
        }
    } catch Error as err {
        ; Return empty array on error
        return completedArticles
    }

    return completedArticles
}

; Cleanup Wikipedia selector (mirror AiModelSelector_Close in Utils.ahk)
CleanupWikipediaSelector() {
    global g_WikipediaSelectorActive, g_WikipediaSelectorGui, g_WikipediaSelectorHandlers

    if (!g_WikipediaSelectorActive)
        return

    g_WikipediaSelectorActive := false

    global g_OnEscapePressed
    g_OnEscapePressed := ""

    ; Disable hotkeys (same order as AiModelSelector_Close: digit keys, then Escape, then restore global Escape)
    for handler in g_WikipediaSelectorHandlers {
        try {
            Hotkey(handler.char, "Off")
        } catch {
        }
    }
    try {
        Hotkey("Escape", WikipediaSelector_Cancel, "Off")
    } catch {
    }
    try {
        Hotkey("*Escape", WikipediaSelector_Cancel, "Off")
    } catch {
    }
    Utils_EnsureGlobalEscapeHotkey()

    g_WikipediaSelectorHandlers := []

    ; Close and destroy GUI
    if (IsObject(g_WikipediaSelectorGui)) {
        try {
            g_WikipediaSelectorGui.Destroy()
        } catch {
        }
        g_WikipediaSelectorGui := false
    }
}

; Show Wikipedia selector GUI (mirror ShowAiModelSelector in Utils.ahk)
ShowWikipediaSelector() {
    global g_WikipediaSelectorGui, g_WikipediaSelectorActive, g_WikipediaSelectorHandlers
    global g_WikipediaItems

    ; Don't show if already active (same as ShowAiModelSelector)
    if (g_WikipediaSelectorActive)
        return

    ; LL keyboard hook may still be swallowing keys from a prior scroll-restore guard — Esc must reach hotkeys.
    AL_RemoveInputGuard()

    ; Get monitor dimensions early for responsive sizing
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
            ; Calculate window center
            winLeft := NumGet(rect, 0, "int")
            winTop := NumGet(rect, 4, "int")
            winRight := NumGet(rect, 8, "int")
            winBottom := NumGet(rect, 12, "int")

            centerX := winLeft + (winRight - winLeft) // 2
            centerY := winTop + (winBottom - winTop) // 2

            ; Find which monitor contains the window center
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

    ; Create GUI — same dark modal styling as ShowAiModelSelector (#!+C) in Utils.ahk
    ; +E0x08000000: non-activating so PowerToys Command Palette stays open
    g_WikipediaSelectorGui := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x08000000")
    g_WikipediaSelectorGui.BackColor := "1E1E2E"
    g_WikipediaSelectorGui.MarginX := 20
    g_WikipediaSelectorGui.MarginY := 15

    ; Load completed articles
    completedArticles := LoadCompletedArticles()

    ; Build display text (section headers: "— Label —" like Project Selector)
    displayText := ""
    displayText .= "— Available Articles —`n"
    for i, item in g_WikipediaItems {
        displayText .= "[" . item.char . "] > " . item.title . "`n"
    }

    ; Add History section if there are completed articles
    if (completedArticles.Length > 0) {
        displayText .= "`n"
        displayText .= "— History (Read) —`n"
        for i, article in completedArticles {
            displayText .= "  • " . article . "`n"
        }
    }

    ; Title + separator (match Utils ShowAiModelSelector / CursorTransfer selectors)
    baseWidth := (monitorWidth < 1200) ? 500 : 600
    wikiContentW := baseWidth - 40
    g_WikipediaSelectorGui.SetFont("s14 cCDD6F4 Bold", "Segoe UI")
    g_WikipediaSelectorGui.Add("Text", "w" . wikiContentW . " Center", "📖 Wikipedia Articles")
    g_WikipediaSelectorGui.Add("Text", "w" . wikiContentW . " h1 Background45475A")

    ; Calculate Edit height from line count (footer hint is separate Text controls)
    lineCount := 0
    loop parse, displayText, "`n" {
        lineCount++
    }
    lineHeight := 18
    textControlHeight := lineCount * lineHeight
    minHeight := 120
    maxHeightPercent := (monitorHeight < 800) ? 0.90 : 0.75
    maxHeight := Floor(monitorHeight * maxHeightPercent)
    if (textControlHeight < minHeight)
        textControlHeight := minHeight
    if (textControlHeight > maxHeight)
        textControlHeight := maxHeight

    g_WikipediaSelectorGui.SetFont("s12 cCDD6F4", "Segoe UI")
    g_WikipediaSelectorGui.AddEdit("w" . wikiContentW . " h" . textControlHeight . " ReadOnly VScroll Background313244",
        displayText)

    g_WikipediaSelectorGui.Add("Text", "w" . wikiContentW . " h1 Background45475A y+10")
    g_WikipediaSelectorGui.SetFont("s9 c6C7086", "Segoe UI")
    g_WikipediaSelectorGui.Add("Text", "w" . wikiContentW . " Center", "Press 1–5 | Esc to cancel")

    try {
        g_WikipediaSelectorGui.OnEvent("Escape", WikipediaSelector_GuiEscape)
    } catch {
    }

    ; Measure and center on the active window's monitor (same pattern as ShowAiModelSelector)
    g_WikipediaSelectorGui.Show("AutoSize Hide")
    g_WikipediaSelectorGui.GetPos(&gx, &gy, &gw, &gh)
    cx := monitorLeft + (monitorWidth - gw) // 2
    cy := monitorTop + (monitorHeight - gh) // 2
    g_WikipediaSelectorGui.Show("x" . cx . " y" . cy . " NA")

    g_WikipediaSelectorActive := true

    ; Enable hotkeys for 1–5 and Escape (same order as ShowAiModelSelector: digits first, then Escape)
    ; Clear handlers array
    g_WikipediaSelectorHandlers := []

    ; Enable hotkeys for characters 1-5
    for item in g_WikipediaItems {
        char := item.char
        handler := CreateWikipediaCharHandler(char)
        g_WikipediaSelectorHandlers.Push({ char: char, handler: handler })
        try {
            Hotkey(char, handler, "On")
        } catch {
        }
    }

    ; *Escape: not removed by Utils Hotkey("Escape","Off") (square selector, etc.); CursorTransfer uses the same pattern.
    Hotkey("*Escape", WikipediaSelector_Cancel, "On")

    ; If I0 *Escape never receives the key, I10 Utils_GlobalEscapeHandler still runs — same hook as g_OnEscapePressed (Utils.ahk).
    global g_OnEscapePressed
    g_OnEscapePressed := WikipediaSelector_GlobalEscapeCallback
    Utils_EnsureGlobalEscapeHotkey()
}
