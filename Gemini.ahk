#Requires AutoHotkey v2.0
#SingleInstance Force

; --- Includes ----------------------------------------------------------------
#include UIA-v2\Lib\UIA.ahk
#include UIA-v2\Lib\UIA_Browser.ahk
#include %A_ScriptDir%\env.ahk
#include %A_ScriptDir%\Utils.ahk

; --- Config ---------------------------------------------------------------
; Path to the file containing the initial prompt Gemini should receive.
PROMPT_FILE := A_ScriptDir "\data\Gemini_Prompt.txt"

; Copy response button names (EN/PT). Excludes "Copy prompt" / "Copiar prompt" which are different controls.
GEMINI_COPY_RESPONSE_NAMES := ["Copy", "Copiar"]

; --- Helper Functions --------------------------------------------------------
; FindGeminiPromptField and GEMINI_PROMPT_FIELD_NAMES are defined in Utils.ahk (included above).

; True if button name is the "Copy [last response]" button (EN or PT), not "Copy prompt".
IsGeminiCopyResponseButton(name) {
    if (!name || InStr(name, "prompt"))
        return false
    for n in GEMINI_COPY_RESPONSE_NAMES {
        if (name = n || InStr(name, n, false))
            return true
    }
    return false
}

; Return count of Gemini "Copy response" buttons (same logic as CopyLastGeminiMessageToClipboard). Caller must ensure tab is active and scrolled to bottom.
GetGeminiCopyButtonCount(uia) {
    allCopyButtons := []
    try {
        allButtons := uia.FindAll({ Type: 50000 })
        for button in allButtons {
            if (IsGeminiCopyResponseButton(button.Name)) {
                if (InStr(button.ClassName, "icon-button") || InStr(button.ClassName, "mdc-button"))
                    allCopyButtons.Push(button)
            }
        }
        if (allCopyButtons.Length = 0) {
            allButtons := uia.FindAll({ Type: "Button" })
            for button in allButtons {
                if (IsGeminiCopyResponseButton(button.Name))
                    allCopyButtons.Push(button)
            }
        }
    } catch {
        return 0
    }
    return allCopyButtons.Length
}

; True if at least one "Show more options" / "More options" exists (last response can then offer read aloud). Same labels as GeminiTriggerReadAloud.
GeminiHasMoreOptionsForResponse(uia) {
    allMoreOptionsButtons := []
    try {
        allMoreOptionsButtons := uia.FindAll({ Name: "Show more options" })
        try {
            moreOpt := uia.FindAll({ Name: "More options" })
            for btn in moreOpt
                allMoreOptionsButtons.Push(btn)
        } catch {
        }
    } catch {
    }
    if (allMoreOptionsButtons.Length = 0) {
        try {
            allMenuItems := uia.FindAll({ Type: 50011 })
            for menuItem in allMenuItems {
                name := menuItem.Name
                if (name = "Show more options" || name = "More options" || InStr(name, "Show more options", false) = 1 ||
                InStr(name, "More options", false) = 1)
                    allMoreOptionsButtons.Push(menuItem)
            }
        } catch {
        }
    }
    return allMoreOptionsButtons.Length > 0
}

; Find Gemini browser window (case-insensitive contains match for "gemini")
GetGeminiWindowHwnd() {
    try {
        for hwnd in WinGetList("ahk_exe chrome.exe") {
            try {
                if InStr(WinGetTitle("ahk_id " hwnd), "gemini", false)
                    return hwnd
            } catch {
                ; Silently skip invalid windows
            }
        }
    } catch {
        ; Silently handle WinGetList errors
    }
    return 0
}

; =============================================================================
; Get work area (left, top, right, bottom) of the monitor that contains the given window
; =============================================================================
GetWorkAreaForWindow(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return ""
    try {
        WinGetPos(&winX, &winY, &winW, &winH, "ahk_id " hwnd)
        centerX := winX + winW / 2
        centerY := winY + winH / 2
        n := MonitorGetCount()
        loop n {
            MonitorGet(A_Index, &L, &T, &R, &B)
            if (centerX >= L && centerX < R && centerY >= T && centerY < B) {
                MonitorGetWorkArea(A_Index, &wLeft, &wTop, &wRight, &wBottom)
                return { left: wLeft, top: wTop, right: wRight, bottom: wBottom }
            }
        }
    } catch {
    }
    return ""
}

; =============================================================================
; Get 1-based active tab index and tab count in Chrome via UIA (tab bar TabItem elements).
; Returns {index: n, count: c} on success; 0 if detection fails. Used when #!+i is triggered
; on an existing Gemini window (banner shown only when count >= 2).
;
; Implementation proposals for Chrome tab identification (for testing/verification):
;   A) UIA (current): Use UIA_Browser.GetTabs() and GetTab("") with SelectionItemIsSelected,
;      then match selected tab by RuntimeId in the tab list to get 1-based index. Reliable for
;      Chrome/Edge when the tab bar is exposed to UIA.
;   B) Window title: WinGetTitle() often includes the active tab's page title; does not give
;      tab index, but could be used to show "current tab title" in a text banner instead of a number.
;   C) Chrome DevTools Protocol (CDP): Requires Chrome started with --remote-debugging-port;
;      can query tabs via HTTP/json. More setup, not used here.
; =============================================================================
GetChromeActiveTabIndex(uia) {
    try {
        uia.GetCurrentMainPaneElement()
        tabs := uia.GetTabs()
        if (!tabs.Length)
            return 0
        current := uia.GetTab("")
        if (!current)
            return 0
        rid := current.RuntimeId
        for i, tab in tabs {
            try {
                if (tab.RuntimeId = rid)
                    return { index: i, count: tabs.Length }
            } catch {
                continue
            }
        }
    } catch {
    }
    return 0
}

; =============================================================================
; Unified banner builder – consistent shape/font/opacity for all banners here
; centerOnHwnd: optional; when set, banner is centered on that window's monitor
; textWidth: optional width of the text control (default 500)
; =============================================================================
CreateCenteredBanner(message, bgColor := "3772FF", fontColor := "FFFFFF", fontSize := 24, alpha := 178, centerOnHwnd :=
    0, textWidth := 500) {
    bGui := Gui()
    bGui.Opt("+AlwaysOnTop -Caption +ToolWindow")
    bGui.BackColor := bgColor
    bGui.SetFont("s" . fontSize . " c" . fontColor . " Bold", "Segoe UI")
    bGui.Add("Text", "w" . textWidth . " Center Wrap", message)

    workArea := (centerOnHwnd && GetWorkAreaForWindow(centerOnHwnd) != "") ? GetWorkAreaForWindow(centerOnHwnd) : ""
    if (workArea != "") {
        winX := workArea.left
        winY := workArea.top
        winW := workArea.right - workArea.left
        winH := workArea.bottom - workArea.top
    } else {
        activeWin := WinGetID("A")
        if (activeWin) {
            WinGetPos(&winX, &winY, &winW, &winH, activeWin)
        } else {
            MonitorGetWorkArea(, &wLeft, &wTop, &wRight, &wBottom)
            winX := wLeft, winY := wTop, winW := wRight - wLeft, winH := wBottom - wTop
        }
    }

    bGui.Show("AutoSize Hide")
    guiW := 0, guiH := 0
    bGui.GetPos(, , &guiW, &guiH)

    guiX := winX + (winW - guiW) / 2
    guiY := winY + (winH - guiH) / 2
    bGui.Show("x" . Round(guiX) . " y" . Round(guiY) . " NA")
    WinSetTransparent(alpha, bGui)
    return bGui
}

; =============================================================================
; Show tab indicator banner (1 = blue, 2 = yellow): square, center of active-window monitor.
; Delegates to Utils for identical behavior as #!+U tab-switching in Utils.ahk. Auto-hides after 700 ms.
; =============================================================================
ShowGeminiTabBanner(tabNumber, geminiHwnd := 0) {
    ShowSingleCharTabBanner_Utils(tabNumber)
}

; =============================================================================
; Helper function to show a notification on the active window
; =============================================================================
ShowNotification(message, durationMs := 500, bgColor := "FFFF00", fontColor := "000000", fontSize := 24) {
    notificationGui := CreateCenteredBanner(message, bgColor, fontColor, fontSize, 178)
    Sleep(durationMs)
    if IsObject(notificationGui) && notificationGui.Hwnd {
        notificationGui.Destroy()
    }
}

; =============================================================================
; Copy completed chime (single beep, debounced)
; =============================================================================
PlayCopyCompletedChime() {
    try {
        static lastTick := 0
        if (A_TickCount - lastTick < 1500)
            return
        lastTick := A_TickCount

        if (IsSoundEnabled()) {
            SoundPlay(A_ScriptDir . "\sounds\copy.wav")
        }
    } catch {
        ; Silently ignore errors
    }
}

; =============================================================================
; Small Loading Indicator Helpers
; =============================================================================
global smallLoadingGuis_Gemini := []

ShowSmallLoadingIndicator(state := "Loading…", bgColor := "3772FF", centerOnHwnd := 0, textWidth := 500, fontSize := 24
) {
    global smallLoadingGuis_Gemini

    ; If GUIs exist, just update the text of the topmost one (the message)
    if (smallLoadingGuis_Gemini.Length > 0) {
        try {
            ; The text control is expected to be in the first GUI of the stack
            if (smallLoadingGuis_Gemini[1].Controls.Length > 0)
                smallLoadingGuis_Gemini[1].Controls[1].Text := state
        } catch {
            ; Silently handle GUI/control errors and recreate
        }
        return
    }

    ; Create a single, high-contrast, centered banner (on given window's monitor if centerOnHwnd)
    textGui := CreateCenteredBanner(state, bgColor, "FFFFFF", fontSize, 178, centerOnHwnd, textWidth)
    smallLoadingGuis_Gemini.Push(textGui)
}

HideSmallLoadingIndicator() {
    StandardLoadingBar_Hide(0)
}

WaitForButtonAndShowSmallLoading(buttonNames, stateText := "Loading…", timeout := 15000) {
    try cUIA := UIA_Browser()
    catch {
        ; Silently ignore UIA browser errors
        return
    }
    start := A_TickCount
    deadline := (timeout > 0) ? (start + timeout) : 0
    btn := ""
    indicatorShown := false
    buttonEverFound := false
    buttonDisappeared := false
    while (timeout <= 0 || A_TickCount < deadline) {
        btn := ""
        for n in buttonNames {
            try {
                btn := cUIA.FindElement({ Name: n, Type: "Button" })
            } catch {
                btn := ""
            }
            if btn
                break
        }
        if btn {
            buttonEverFound := true
            if (!indicatorShown) {
                StandardLoadingBar_Show(stateText, "3772FF")
                indicatorShown := true
            }
            while btn && (timeout <= 0 || A_TickCount < deadline) {
                Sleep 250
                btn := ""
                for n in buttonNames {
                    try {
                        btn := cUIA.FindElement({ Name: n, Type: "Button" })
                    } catch {
                        btn := ""
                    }
                    if btn
                        break
                }
            }
            if !btn
                buttonDisappeared := true
            break
        }
        Sleep 250
    }
    ; Play completion sound only for actual AI responses when we saw the button and it disappeared
    try {
        if (buttonEverFound && buttonDisappeared && InStr(StrLower(stateText), "transcrib") = 0)
            PlayCopyCompletedChime()
    } catch {
        ; Silently ignore errors
    }
    StandardLoadingBar_Hide(0)
}

; =============================================================================
; Helper function to center mouse on the active window
; =============================================================================
CenterMouse() {
    Sleep 200
    Send("#!+q")
}

; --- Hotkeys ----------------------------------------------------------------

; Reusable: activate Gemini, handle Pause/Resume, then optionally copy last message and trigger "Text to speech".
; copyFirst: true = copy last response then read aloud (#!+o); false = only read aloud (#!+7).
; useTrashTab: when true, explicitly target the second Gemini tab (trash tab) instead of the main tab.
GeminiTriggerReadAloud(copyFirst := true, useTrashTab := false) {
    ; Step 1: Activate Gemini window globally
    SetTitleMatchMode(2)
    if hwnd := GetGeminiWindowHwnd() {
        try {
            WinActivate("ahk_id " hwnd)
        } catch {
            ShowCenteredOverlay_Utils("Error: Target window not found.", 2000)
            return
        }
    }
    if !WinWaitActive("ahk_exe chrome.exe", , 2)
        return
    Sleep 150

    ; When requested (#!+o trash tab), explicitly switch to the second Gemini tab.
    ; Chrome convention: Ctrl+2 selects the second tab in the window.
    if (useTrashTab) {
        Send("^2")
        Sleep 150
        ShowGeminiTabBanner(2, hwnd)
    }

    ; Step 2: Check if "Pause" button exists (if reading is active, pause it)
    uia := UIA_Browser()
    Sleep 120

    pauseButton := 0
    try {
        pauseButton := uia.FindFirst({ Name: "Pause", Type: 50000 })
        if !pauseButton
            pauseButton := uia.FindFirst({ Type: "Button", Name: "Pause" })
        if !pauseButton {
            allButtons := uia.FindAll({ Type: 50000 })
            for button in allButtons {
                if (button.Name = "Pause" || InStr(button.Name, "Pause", false) = 1) {
                    if (InStr(button.ClassName, "tts-button") || InStr(button.ClassName, "mdc-icon-button")) {
                        pauseButton := button
                        break
                    }
                }
            }
        }
    } catch {
    }

    if (pauseButton) {
        pauseButton.Click()
        ShowNotification("Paused", 800, "FFFF00", "000000", 24)
        Send "!{Tab}"
        return
    }

    resumeButton := 0
    try {
        resumeButton := uia.FindFirst({ Name: "Resume", Type: 50000 })
        if !resumeButton
            resumeButton := uia.FindFirst({ Type: "Button", Name: "Resume" })
        if !resumeButton {
            allButtons := uia.FindAll({ Type: 50000 })
            for button in allButtons {
                if (button.Name = "Resume" || InStr(button.Name, "Resume", false) = 1) {
                    if (InStr(button.ClassName, "tts-button") || InStr(button.ClassName, "mdc-icon-button")) {
                        resumeButton := button
                        break
                    }
                }
            }
        }
    } catch {
    }

    if (resumeButton) {
        resumeButton.Click()
        ShowNotification("Resumed", 800, "FFFF00", "000000", 24)
        Send "!{Tab}"
        return
    }

    ; Step 3: If copyFirst, find and click the last Copy button; else just scroll so last response is in view
    Send "^{End}"
    Sleep 350

    if (copyFirst) {
        allCopyButtons := []
        allButtons := uia.FindAll({ Type: 50000 })
        for button in allButtons {
            if (IsGeminiCopyResponseButton(button.Name)) {
                if (InStr(button.ClassName, "icon-button") || InStr(button.ClassName, "mdc-button"))
                    allCopyButtons.Push(button)
            }
        }
        if (allCopyButtons.Length = 0) {
            allButtons := uia.FindAll({ Type: "Button" })
            for button in allButtons {
                if (IsGeminiCopyResponseButton(button.Name))
                    allCopyButtons.Push(button)
            }
        }

        lastCopyButton := (allCopyButtons.Length > 0) ? allCopyButtons[allCopyButtons.Length] : 0
        if (lastCopyButton) {
            lastCopyButton.Click()
            PlayCopyCompletedChime()
        }
    }

    ; Step 4: Find the final "More options" / "Show more options" in the Gemini response tree (bottom-up).
    ; We target only the most recent Gemini response to avoid reading older messages. See gemini-tree.txt for tree structure.
    searchBanner := CreateCenteredBanner(copyFirst ? "Finding read aloud button and copying..." :
        "Finding read aloud button...", "3772FF", "FFFFFF", 24, 178)
    Sleep 250

    allMoreOptionsButtons := []
    try {
        ; Primary: "Show more options" (EN); include "More options" for alternate labels
        allMoreOptionsButtons := uia.FindAll({ Name: "Show more options" })
        try {
            moreOpt := uia.FindAll({ Name: "More options" })
            for btn in moreOpt
                allMoreOptionsButtons.Push(btn)
        } catch {
        }
    } catch {
    }
    if (allMoreOptionsButtons.Length = 0) {
        try {
            allMenuItems := uia.FindAll({ Type: 50011 })
            for menuItem in allMenuItems {
                name := menuItem.Name
                if (name = "Show more options" || name = "More options" || InStr(name, "Show more options", false) = 1 ||
                InStr(name, "More options", false) = 1)
                    allMoreOptionsButtons.Push(menuItem)
            }
        } catch {
        }
    }

    if (allMoreOptionsButtons.Length = 0) {
        if IsObject(searchBanner) && searchBanner.Hwnd
            searchBanner.Destroy()
        return
    }

    ; Bottom-up: select the last instance in the response tree = most recent response only.
    ; 1) Prefer button with the largest bottom Y (true bottom of page = final response).
    ; 2) Fallback: last element in FindAll order (document/tree order = last in tree).
    lastMoreOptionsButton := 0
    highestBottomY := -1
    for moreOptionsButton in allMoreOptionsButtons {
        try {
            btnPos := moreOptionsButton.Location
            bottomY := btnPos.y + btnPos.h
            if (bottomY > highestBottomY) {
                highestBottomY := bottomY
                lastMoreOptionsButton := moreOptionsButton
            }
        } catch {
            continue
        }
    }
    if (!lastMoreOptionsButton && allMoreOptionsButtons.Length > 0)
        lastMoreOptionsButton := allMoreOptionsButtons[allMoreOptionsButtons.Length]
    if (!lastMoreOptionsButton) {
        if IsObject(searchBanner) && searchBanner.Hwnd
            searchBanner.Destroy()
        return
    }

    try {
        lastMoreOptionsButton.Click()
        Sleep 200

        textToSpeechMenuItem := 0
        try {
            textToSpeechMenuItem := uia.FindFirst({ Name: "Text to speech", Type: 50011 })
        } catch {
        }
        if !textToSpeechMenuItem {
            try {
                textToSpeechMenuItem := uia.FindFirst({ Type: "MenuItem", Name: "Text to speech" })
            } catch {
            }
        }
        if !textToSpeechMenuItem {
            try {
                allMenuItems := uia.FindAll({ Type: 50011 })
                for menuItem in allMenuItems {
                    if (menuItem.Name = "Text to speech" || InStr(menuItem.Name, "Text to speech", false) = 1) {
                        if (InStr(menuItem.ClassName, "mat-mdc-menu-item")) {
                            textToSpeechMenuItem := menuItem
                            break
                        }
                    }
                }
            } catch {
            }
        }
        if !textToSpeechMenuItem {
            try {
                allMenuItems := uia.FindAll({ Type: 50011 })
                for menuItem in allMenuItems {
                    if (menuItem.Name = "Text to speech" || InStr(menuItem.Name, "Text to speech", false) = 1) {
                        textToSpeechMenuItem := menuItem
                        break
                    }
                }
            } catch {
            }
        }

        if (textToSpeechMenuItem) {
            textToSpeechMenuItem.Click()
            Sleep 200
        } else {
            Send "{Down}"
            Sleep 200
            Send "{Enter}"
        }
    } catch {
    }

    if IsObject(searchBanner) && searchBanner.Hwnd
        searchBanner.Destroy()

    Sleep 1500

    isReading := false
    try {
        if uia.FindFirst({ Name: "Pause", Type: 50000 })
            isReading := true
        else if uia.FindFirst({ Type: "Button", Name: "Pause" })
            isReading := true
        else {
            allButtons := uia.FindAll({ Type: 50000 })
            for button in allButtons {
                if (button.Name = "Pause" || InStr(button.Name, "Pause", false) = 1) {
                    if (InStr(button.ClassName, "tts-button") || InStr(button.ClassName, "mdc-icon-button")) {
                        isReading := true
                        break
                    }
                }
            }
        }
    } catch {
    }

    if (!isReading) {
        ShowNotification("Retrying read aloud...", 800, "FFFF00", "000000", 24)
        try {
            lastMoreOptionsButton.Click()
            Sleep 200

            textToSpeechMenuItem := 0
            try {
                textToSpeechMenuItem := uia.FindFirst({ Name: "Text to speech", Type: 50011 })
            } catch {
            }
            if !textToSpeechMenuItem {
                try {
                    textToSpeechMenuItem := uia.FindFirst({ Type: "MenuItem", Name: "Text to speech" })
                } catch {
                }
            }
            if !textToSpeechMenuItem {
                try {
                    allMenuItems := uia.FindAll({ Type: 50011 })
                    for menuItem in allMenuItems {
                        if (menuItem.Name = "Text to speech" || InStr(menuItem.Name, "Text to speech", false) = 1) {
                            textToSpeechMenuItem := menuItem
                            break
                        }
                    }
                } catch {
                }
            }

            if (textToSpeechMenuItem) {
                textToSpeechMenuItem.Click()
                Sleep 200
            } else {
                Send "{Down}"
                Sleep 200
                Send "{Enter}"
            }
        } catch {
        }
    }

    ShowNotification(copyFirst ? "Copied & Reading aloud" : "Reading aloud", 800, "FFFF00", "000000", 24)
    Send "!{Tab}"
}

; Win+Alt+Shift+O : Read aloud the last message in Gemini (or Pause/Resume if already reading)
#!+o:: {
    try {
        ; Standard behavior: operate on the currently active Gemini tab.
        GeminiTriggerReadAloud()
    } catch Error as e {
        ;
    }
}

; Copy last Gemini message to clipboard. Used by #!+p and by async pronunciation flow.
; options.restoreWindow (default true): send !{Tab} after copy. options.playChimeAndNotify (default true): play chime and show "Copied!".
; options.alreadyActive (default false): when true, skip activation; assume Gemini is already the active window (use UIA_Browser() with no arg).
; geminiHwnd: optional; if 0, uses GetGeminiWindowHwnd(). Returns true if copy succeeded, false otherwise.
CopyLastGeminiMessageToClipboard(options := "", geminiHwnd := 0) {
    restoreWindow := (options = "" || !options.HasProp("restoreWindow")) ? true : options.restoreWindow
    playChimeAndNotify := (options = "" || !options.HasProp("playChimeAndNotify")) ? true : options.playChimeAndNotify
    alreadyActive := (options != "" && options.HasProp("alreadyActive")) ? options.alreadyActive : false
    try {
        SetTitleMatchMode(2)
        if !geminiHwnd
            geminiHwnd := GetGeminiWindowHwnd()
        if !geminiHwnd
            return false
        if (!alreadyActive) {
            try {
                WinActivate("ahk_id " geminiHwnd)
            } catch {
                ShowCenteredOverlay_Utils("Error: Target window not found.", 2000)
                return false
            }
            if !WinWaitActive("ahk_exe chrome.exe", , 2)
                return false
            Sleep 150
        }

        ; Scroll to bottom *before* UIA so the last response is in the tree and we go down the chat.
        Send "^{End}"
        Sleep 350

        uia := alreadyActive ? UIA_Browser() : UIA_Browser("ahk_id " geminiHwnd)
        Sleep 120

        ; Bottom-up by tree order: FindAll returns elements in document order, so the *last* Copy button in the array is the last response.
        allCopyButtons := []
        allButtons := uia.FindAll({ Type: 50000 })
        for button in allButtons {
            if (IsGeminiCopyResponseButton(button.Name)) {
                if (InStr(button.ClassName, "icon-button") || InStr(button.ClassName, "mdc-button"))
                    allCopyButtons.Push(button)
            }
        }
        if (allCopyButtons.Length = 0) {
            allButtons := uia.FindAll({ Type: "Button" })
            for button in allButtons {
                if (IsGeminiCopyResponseButton(button.Name))
                    allCopyButtons.Push(button)
            }
        }

        ; Last in array = last in chat (tree order). Ignore all previous Copy buttons.
        lastCopyButton := (allCopyButtons.Length > 0) ? allCopyButtons[allCopyButtons.Length] : 0

        if (!lastCopyButton)
            return false
        lastCopyButton.Click()
        if !ClipWait(2)
            return false
        if (playChimeAndNotify) {
            PlayCopyCompletedChime()
            ShowNotification("Copied!", 800, "FFFF00", "000000", 24)
        }
        if (restoreWindow)
            Send "!{Tab}"
        return true
    } catch {
        return false
    }
}

; Win+Alt+Shift+P : Click the last Copy button in Gemini (activates Gemini, scrolls to bottom with Ctrl+End, then copies last response)
; Works in EN ("Copy") and PT ("Copiar") UI. Uses tree order: last Copy button in the UI tree = last response.
#!+p:: {
    try {
        if (!CopyLastGeminiMessageToClipboard())
            ShowNotification("Copy failed – ensure Gemini is open and has a response", 2500, "FF6666", "FFFFFF", 22)
    } catch as err {
        ShowNotification("Copy error: " (err.Message ? err.Message : "unknown"), 2500, "FF6666", "FFFFFF", 22)
    }
}

; Custom message so WindowManagement.ahk can trigger copy without Send (Send does not trigger hotkeys in another script).
WM_COPY_LAST_GEMINI := 0x8001
; Start background completion monitor for Ctrl+Alt+Win+L (wParam = originalHwnd, lParam = geminiHwnd). Sent from Utils.ahk.
WM_START_DELAYED_SUBMIT_MONITOR := 0x8002
; Path for bridge to verify that Copy Last Response (same as #!+p) actually succeeded
GEMINI_COPY_RESULT_PATH := A_ScriptDir "\.cursor\gemini_copy_result.txt"

OnMessage(WM_COPY_LAST_GEMINI, copyFromBridge)
OnMessage(WM_START_DELAYED_SUBMIT_MONITOR, handleStartDelayedSubmitMonitor)
handleStartDelayedSubmitMonitor(wParam, lParam, msg, hwnd) {
    GeminiDelayedSubmitMonitorStart(wParam, lParam)
}
copyFromBridge(*) {
    ; Guarantee layer: write result so bridge can confirm we copied Gemini's last response (same path as #!+p).
    try
        FileDelete(GEMINI_COPY_RESULT_PATH)
    try
        FileAppend("0", GEMINI_COPY_RESULT_PATH)
    r := CopyLastGeminiMessageToClipboard({ restoreWindow: false, playChimeAndNotify: false })
    try
        FileDelete(GEMINI_COPY_RESULT_PATH)
    try
        FileAppend(r ? "1" : "0", GEMINI_COPY_RESULT_PATH)
}

; =============================================================================
; TTS from selection – Win+Alt+Shift+7: copy selection, send "repeat exactly" to Gemini, then trigger read aloud
; =============================================================================
#!+7:: {
    (GeminiAsyncTTS()).Start()
}

; =============================================================================
; Get Pronunciation
; Hotkey: Win+Alt+Shift+8 — async: submit to Gemini, restore focus, show result in banner when ready
; =============================================================================
#!+8:: {
    (GeminiAsyncLookup()).Start()
}

; =============================================================================
; Initialize Gemini window on first-time opening
; =============================================================================
SendPromptToActiveGeminiTab(promptText) {
    try {
        if (StrLen(promptText) = 0)
            return false

        uia := UIA_Browser()
        Sleep 300

        promptField := FindGeminiPromptField(uia)
        if (!promptField)
            return false

        promptField.SetFocus()
        Sleep 100
        if (!promptField.HasKeyboardFocus) {
            try
                promptField.Click()
            Sleep 100
        }

        oldClip := A_Clipboard
        A_Clipboard := ""
        A_Clipboard := promptText
        ClipWait 1
        Send("^v")
        Sleep 100
        Send("{Enter}")
        Sleep 100
        A_Clipboard := oldClip
        return true
    } catch {
        return false
    }
}

InitializeGeminiFirstTime() {
    try {
        ; Show banner to inform user
        StandardLoadingBar_Show("Opening Gemini (2 tabs)...", "3772FF")

        ; Remember existing Chrome windows so we can find the one we're about to create
        existingChromeHwnds := []
        try {
            for hwnd in WinGetList("ahk_exe chrome.exe")
                existingChromeHwnds.Push(hwnd)
        } catch {
        }

        ; Run Chrome with new window and two Gemini tabs
        Run "chrome.exe --new-window https://gemini.google.com/ https://gemini.google.com/"
        Sleep 700   ; Give the system time to start Chrome before waiting for it

        ; Find the newly created Chrome window (not one that was already open)
        geminiHwnd := 0
        loop 35 {   ; 35 * 300ms ≈ 10.5s max wait for new window to appear
            for hwnd in WinGetList("ahk_exe chrome.exe") {
                isNew := true
                for existing in existingChromeHwnds {
                    if (existing = hwnd) {
                        isNew := false
                        break
                    }
                }
                if (isNew) {
                    geminiHwnd := hwnd
                    break 2
                }
            }
            Sleep 300
        }
        if !geminiHwnd {
            StandardLoadingBar_Hide(0)
            return
        }

        ; Activate the new Gemini window and wait until it is actually active
        try {
            WinActivate("ahk_id " geminiHwnd)
        } catch {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("Error: Target window not found.", 2000)
            return
        }
        if !WinWaitActive("ahk_id " geminiHwnd, , 4) {
            StandardLoadingBar_Hide(0)
            return
        }
        ; Wait for the first tab to load so the title contains "Gemini" before sending the prompt
        SetTitleMatchMode(2)
        start := A_TickCount
        while (A_TickCount - start < 6000) {
            try {
                if InStr(WinGetTitle("ahk_id " geminiHwnd), "Gemini", false)
                    break
            } catch {
            }
            Sleep 250
        }
        Sleep 550   ; Give window and tabs time to fully settle

        ; Read initial prompt from external file
        promptText := ""
        try promptText := FileRead(PROMPT_FILE, "UTF-8")
        if (StrLen(promptText) = 0)
            promptText := "hey, what's up?"

        ; Update banner status
        StandardLoadingBar_Show("Sending prompt to Gemini tabs...", "3772FF")

        geminiHwnd := GetGeminiWindowHwnd()
        ; Ensure first tab is active and send prompt (no tab banner during initial launch)
        Send("^1")
        Sleep 280
        SendPromptToActiveGeminiTab(promptText)

        ; Switch to second tab and send the same prompt
        Send("^2")
        Sleep 280
        SendPromptToActiveGeminiTab(promptText)

        ; Hide banner on success
        StandardLoadingBar_Hide(0)
    } catch Error as err {
        ; Hide banner on error
        StandardLoadingBar_Hide(0)
    }
}

; =============================================================================
; Open Gemini
; Hotkey: Win+Alt+Shift+I
; =============================================================================
#!+i:: {
    SetTitleMatchMode(2)
    if hwnd := GetGeminiWindowHwnd() {
        try {
            WinActivate("ahk_id " hwnd)
        } catch {
            ShowCenteredOverlay_Utils("Error: Target window not found.", 2000)
            return
        }
        if WinWaitActive("ahk_id " hwnd, , 2) {
            Sleep 120   ; Let the window and Chrome content settle before UIA attaches
            ; Bind UIA to this window so we never attach to a different Chrome window
            uia := UIA_Browser("ahk_id " hwnd)
            Sleep 120   ; UIA settle time (align with CopyLastGeminiMessageToClipboard)

            ; Show current active tab only when this window already has two Gemini tabs (not during initial launch).
            ; Brief extra delay so Chrome tab bar is ready for UIA; retry once if first attempt fails (timing).
            Sleep 80
            tabInfo := GetChromeActiveTabIndex(uia)
            if (!tabInfo) {
                Sleep 150
                tabInfo := GetChromeActiveTabIndex(uia)
            }
            ; Show tab-position banner: use UIA index when available, otherwise assume position 1 so banner always appears
            tabPosition := (tabInfo && tabInfo.count >= 2 && tabInfo.index) ? tabInfo.index : 1
            ShowSingleCharTabBanner_Utils(tabPosition)

            ; Find the anchor element: "Open upload file menu" button
            ; Combined search: Try exact match first, then case-insensitive (most efficient)
            anchorButton := 0
            try {
                anchorButton := uia.FindFirst({ Type: "50000", Name: "Open upload file menu", ControlType: "Button" })
                if (!anchorButton) {
                    anchorButton := uia.FindFirst({ Type: "50000", Name: "Open upload file menu", cs: false })
                }
            } catch {
            }

            ; Fallback: Only use expensive FindAll if first two strategies failed
            if (!anchorButton) {
                try {
                    allButtons := uia.FindAll({ Type: "50000" })
                    for button in allButtons {
                        try {
                            if (InStr(button.Name, "Open upload file menu", false)) {
                                anchorButton := button
                                break
                            }
                        } catch {
                            continue
                        }
                    }
                } catch {
                }
            }

            if (anchorButton) {
                ; Focus the anchor button (do NOT click) and navigate back
                try {
                    anchorButton.SetFocus()
                    Sleep 25   ; minimal wait for focus
                    SendInput "+{Tab}"  ; Use SendInput for faster keystroke
                    Sleep 15   ; minimal delay for navigation
                } catch {
                    ; Anchor strategy failed; will use direct prompt field below
                }
            }
            ; Ensure the prompt field actually has keyboard focus (same as SendPromptToActiveGeminiTab)
            promptField := FindGeminiPromptField(uia)
            if (promptField) {
                try {
                    promptField.SetFocus()
                    Sleep 100
                    if (!promptField.HasKeyboardFocus) {
                        try promptField.Click()
                        Sleep 100
                    }
                } catch {
                }
                if (IsSoundEnabled())
                    SoundPlay(A_ScriptDir . "\sounds\gemini-focused.wav")
            }
        }
    } else {
        InitializeGeminiFirstTime()
    }
}

; =============================================================================
; GeminiAsyncLookup – async pronunciation lookup (Win+Alt+Shift+8)
; User keeps focus; timer polls for completion; result shown in banner.
; =============================================================================
class GeminiAsyncLookup {
    static PronunciationPrompt :=
        "Below, you will find a word or phrase. I'd like you to answer in five sections: the 1st section you will repeat the word twice. For each time you repeat, use a point to finish the phrase. The 2nd section should have the definition of the word (You should also say each part of speech does the different definitions belong to). The 3d section should have the pronunciation of this word using the Internation Phonetic Alphabet characters (for American English).The 4th section should have the same word applied in a real sentence (put that in quotations, so I can identify that). In the 5th, Write down the translation of the word into Portuguese. Please, do not title any section. Thanks!"

    __New() {
        this.OriginalHwnd := 0
        this.GeminiHwnd := 0
        this.TimerCallback := ""
        this.RetryCount := 0
        this.MaxRetries := 60   ; 60 * 500ms = 30s timeout
        this.ButtonEverFound := false
    }

    Start() {
        this.OriginalHwnd := WinExist("A")
        if !this.OriginalHwnd
            return
        ; Show loading banner immediately, centered on the monitor where this window is (with warning)
        StandardLoadingBar_Show("Loading…", "3772FF")

        A_Clipboard := ""
        Send "^c"
        if !ClipWait(2)
            return
        SetTitleMatchMode(2)
        this.GeminiHwnd := GetGeminiWindowHwnd()
        if !this.GeminiHwnd {
            StandardLoadingBar_Hide(0)
            return
        }
        try {
            WinActivate("ahk_id " this.GeminiHwnd)
        } catch {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("Error: Target window not found.", 2000)
            return
        }
        if !WinWaitActive("ahk_exe chrome.exe", , 2) {
            StandardLoadingBar_Hide(0)
            return
        }
        ; For pronunciation lookup (#!+8), always use the trash tab (second Gemini tab).
        ; Chrome convention: Ctrl+2 selects the second tab in the window.
        Send("^2")
        Sleep 150
        ShowGeminiTabBanner(2, this.GeminiHwnd)
        uia := UIA_Browser()
        Sleep 300
        promptField := FindGeminiPromptField(uia)
        if (!promptField) {
            StandardLoadingBar_Hide(0)
            return
        }
        promptField.SetFocus()
        Sleep 100
        if (!promptField.HasKeyboardFocus) {
            try promptField.Click()
            Sleep 100
        }
        ; Switch to Fast model before the prompt (enough for this task)
        Send("@fast ")
        Sleep 200
        searchString := GeminiAsyncLookup.PronunciationPrompt
        A_Clipboard := searchString . "`n`nContent: " . A_Clipboard
        Sleep 100
        Send("^a")
        Sleep 500
        Send("^v")
        Sleep 500
        Send("{Enter}")
        Sleep 300
        ; Go back to the window where you triggered the hotkey so you can keep working (activate only; do not WinRestore or we lose maximized state)
        origHwnd := this.OriginalHwnd
        try {
            if WinExist("ahk_id " origHwnd) {
                WinActivate("ahk_id " origHwnd)
                WinWaitActive("ahk_id " origHwnd, , 1)
            }
        } catch {
            if (WinExist("ahk_id " origHwnd))
                WinActivate("ahk_id " origHwnd)
        }
        this.RetryCount := 0
        this.TimerCallback := this.CheckCompletion.Bind(this)
        SetTimer(this.TimerCallback, 500)
    }

    CheckCompletion() {
        this.RetryCount++
        if (this.RetryCount > this.MaxRetries) {
            SetTimer(this.TimerCallback, 0)
            StandardLoadingBar_Hide(0)
            return
        }
        ; Poll in background using raw UIA (no UIA_Browser) so the library never activates Gemini
        btn := ""
        buttonNames := ["Stop streaming", "Interromper transmissão", "Stop response"]
        try {
            root := UIA.ElementFromHandle(this.GeminiHwnd)
            for n in buttonNames {
                try {
                    btn := root.FindElement({ Name: n, Type: "Button" })
                } catch {
                    btn := ""
                }
                if btn
                    break
            }
        } catch {
            return
        }

        if btn {
            this.ButtonEverFound := true
            return   ; Still streaming
        }

        ; If button was found and now is gone, verify it's truly finished
        if (this.ButtonEverFound) {
            isTrulyGone := true
            loop 4 {
                Sleep 200
                try {
                    for n in buttonNames {
                        if root.ElementExist({ Name: n, Type: "Button" }) {
                            isTrulyGone := false
                            break
                        }
                    }
                } catch
                    isTrulyGone := true
                if !isTrulyGone
                    break
            }

            if isTrulyGone {
                SetTimer(this.TimerCallback, 0)
                ; Use the same sound as Shift keys.ahk for consistency
                try {
                    if (IsSoundEnabled()) {
                        SoundPlay(A_ScriptDir . "\sounds\gemini-completion.wav")
                    }
                } catch {
                    PlayCopyCompletedChime()
                }
                this.RetrieveResponse()
            }
        }
    }

    RetrieveResponse() {
        ; Activate Gemini once, then copy (and retries if needed) without switching back until done
        contentBefore := A_Clipboard
        try {
            WinActivate("ahk_id " this.GeminiHwnd)
        } catch {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("Error: Target window not found.", 2000)
            return
        }
        if !WinWaitActive("ahk_exe chrome.exe", , 2) {
            StandardLoadingBar_Hide(0)
            return
        }
        copyOpt := { restoreWindow: false, playChimeAndNotify: false, alreadyActive: true }
        if !CopyLastGeminiMessageToClipboard(copyOpt, this.GeminiHwnd) {
            StandardLoadingBar_Hide(0)
            return
        }
        Sleep 400
        if (A_Clipboard = "" || A_Clipboard = contentBefore) {
            CopyLastGeminiMessageToClipboard(copyOpt, this.GeminiHwnd)
            Sleep 400
        }
        if (A_Clipboard = "" || A_Clipboard = contentBefore) {
            CopyLastGeminiMessageToClipboard(copyOpt, this.GeminiHwnd)
            Sleep 400
        }
        WinActivate("ahk_id " this.OriginalHwnd)
        StandardLoadingBar_Hide(0)
        this.ShowResultBanner(A_Clipboard)
    }

    ShowResultBanner(text) {
        if (!text || StrLen(Trim(text)) = 0)
            return
        banner := Gui("+AlwaysOnTop -Caption +ToolWindow")
        banner.BackColor := "3772FF"
        banner.SetFont("s14 cFFFFFF", "Segoe UI")
        banner.Add("Text", "w600 Center Wrap", text)
        banner.SetFont("s11 cFFFFFF Bold", "Segoe UI")
        banner.Add("Text", "w600 Center", "Press Enter to close")
        ; Center on the same monitor as the window that triggered the hotkey
        workArea := GetWorkAreaForWindow(this.OriginalHwnd)
        if (workArea = "") {
            MonitorGetWorkArea(, &wLeft, &wTop, &wRight, &wBottom)
            w := wRight - wLeft
            h := wBottom - wTop
            banner.Show("AutoSize Hide")
            banner.GetPos(, , &gw, &gh)
            banner.Show("x" . Round(wLeft + (w - gw) / 2) . " y" . Round(wTop + (h - gh) / 2) . " NA")
        } else {
            w := workArea.right - workArea.left
            h := workArea.bottom - workArea.top
            banner.Show("AutoSize Hide")
            banner.GetPos(, , &gw, &gh)
            banner.Show("x" . Round(workArea.left + (w - gw) / 2) . " y" . Round(workArea.top + (h - gh) / 2) . " NA")
        }
        WinSetTransparent(220, banner)
        closeBanner(*) {
            SetTimer(closeBanner, 0)
            try Hotkey("Escape", closeBanner, "Off")
            try Hotkey("Enter", closeBanner, "Off")
            try banner.Destroy()
        }
        banner.OnEvent("Close", closeBanner)
        SetTimer(closeBanner, -50000)
        ; Press Escape or Enter to remove the banner
        Hotkey("Escape", closeBanner, "On")
        Hotkey("Enter", closeBanner, "On")
    }
}

; =============================================================================
; GeminiDelayedSubmitMonitor – background completion monitor for Ctrl+Alt+Win+L
; Reuses #!+8 completion detection; on completion shows "Copy? [N] [R]" with 4s timeout (N = no copy, R = copy + read aloud).
; =============================================================================
class GeminiDelayedSubmitMonitor {
    __New() {
        this.OriginalHwnd := 0
        this.GeminiHwnd := 0
        this.TimerCallback := ""
        this.RetryCount := 0
        this.MaxRetries := 300   ; 300 * 500ms = 150s timeout (covers Gemini Pro deep thinking)
        this.ButtonEverFound := false
        this.CopyBannerGui := ""
        this.CopyTimeoutTimer := ""
    }

    Start(originalHwnd, geminiHwnd) {
        if (!originalHwnd || !geminiHwnd)
            return
        this.OriginalHwnd := originalHwnd
        this.GeminiHwnd := geminiHwnd
        this.RetryCount := 0
        this.ButtonEverFound := false
        this.TimerCallback := this.CheckCompletion.Bind(this)
        SetTimer(this.TimerCallback, 500)
    }

    CheckCompletion() {
        this.RetryCount++
        if (this.RetryCount > this.MaxRetries) {
            SetTimer(this.TimerCallback, 0)
            return
        }
        btn := ""
        buttonNames := ["Stop streaming", "Interromper transmissão", "Stop response"]
        try {
            root := UIA.ElementFromHandle(this.GeminiHwnd)
            for n in buttonNames {
                try {
                    btn := root.FindElement({ Name: n, Type: "Button" })
                } catch {
                    btn := ""
                }
                if btn
                    break
            }
        } catch {
            return
        }

        if btn {
            this.ButtonEverFound := true
            return
        }

        if (this.ButtonEverFound) {
            isTrulyGone := true
            loop 4 {
                Sleep 200
                try {
                    for n in buttonNames {
                        if root.ElementExist({ Name: n, Type: "Button" }) {
                            isTrulyGone := false
                            break
                        }
                    }
                } catch
                    isTrulyGone := true
                if !isTrulyGone
                    break
            }

            if isTrulyGone {
                SetTimer(this.TimerCallback, 0)
                try {
                    if (IsSoundEnabled())
                        SoundPlay(A_ScriptDir . "\sounds\gemini-completion.wav")
                } catch {
                    PlayCopyCompletedChime()
                }
                this.ShowCopyDecisionBanner()
            }
        }
    }

    ShowCopyDecisionBanner() {
        banner := Gui("+AlwaysOnTop -Caption +ToolWindow")
        banner.BackColor := "3772FF"
        banner.SetFont("s8 cFFFFFF", "Segoe UI")
        banner.Add("Text", "w200 Center", "Copy? [N] [R]")
        workArea := GetWorkAreaForWindow(this.OriginalHwnd)
        if (workArea = "") {
            MonitorGetWorkArea(, &wLeft, &wTop, &wRight, &wBottom)
            w := wRight - wLeft
            h := wBottom - wTop
            banner.Show("AutoSize Hide")
            banner.GetPos(, , &gw, &gh)
            banner.Show("x" . Round(wLeft + (w - gw) / 2) . " y" . Round(wTop + (h - gh) / 2) . " NA")
        } else {
            w := workArea.right - workArea.left
            h := workArea.bottom - workArea.top
            banner.Show("AutoSize Hide")
            banner.GetPos(, , &gw, &gh)
            banner.Show("x" . Round(workArea.left + (w - gw) / 2) . " y" . Round(workArea.top + (h - gh) / 2) . " NA")
        }
        WinSetTransparent(204, banner)  ; 80% opacity, half-size banner
        this.CopyBannerGui := banner
        this.CopyTimeoutTimer := this.DoCopyOnTimeout.Bind(this)
        Hotkey("n", this.CancelCopy.Bind(this), "On")
        Hotkey("N", this.CancelCopy.Bind(this), "On")
        Hotkey("r", this.CopyAndReadAloud.Bind(this), "On")
        Hotkey("R", this.CopyAndReadAloud.Bind(this), "On")
        SetTimer(this.CopyTimeoutTimer, -4000)
    }

    ; Shared cleanup: stop timeout timer, disable N/R hotkeys, destroy banner. Used by CancelCopy, DoCopyOnTimeout, CopyAndReadAloud.
    CleanupCopyBanner() {
        try SetTimer(this.CopyTimeoutTimer, 0)
        try Hotkey("n", "Off")
        try Hotkey("N", "Off")
        try Hotkey("r", "Off")
        try Hotkey("R", "Off")
        if (IsObject(this.CopyBannerGui) && this.CopyBannerGui.Hwnd)
            try this.CopyBannerGui.Destroy()
        this.CopyBannerGui := ""
    }

    CancelCopy(*) {
        this.CleanupCopyBanner()
    }

    DoCopyOnTimeout(*) {
        this.CleanupCopyBanner()

        contentBefore := A_Clipboard
        try {
            WinActivate("ahk_id " this.GeminiHwnd)
        } catch {
            if (WinExist("ahk_id " this.OriginalHwnd))
                WinActivate("ahk_id " this.OriginalHwnd)
            return
        }
        if !WinWaitActive("ahk_exe chrome.exe", , 2) {
            if (WinExist("ahk_id " this.OriginalHwnd))
                WinActivate("ahk_id " this.OriginalHwnd)
            return
        }
        copyOpt := { restoreWindow: false, playChimeAndNotify: false, alreadyActive: true }
        if !CopyLastGeminiMessageToClipboard(copyOpt, this.GeminiHwnd) {
            if (WinExist("ahk_id " this.OriginalHwnd))
                WinActivate("ahk_id " this.OriginalHwnd)
            return
        }
        Sleep 400
        if (A_Clipboard = "" || A_Clipboard = contentBefore) {
            CopyLastGeminiMessageToClipboard(copyOpt, this.GeminiHwnd)
            Sleep 400
        }
        if (A_Clipboard = "" || A_Clipboard = contentBefore) {
            CopyLastGeminiMessageToClipboard(copyOpt, this.GeminiHwnd)
            Sleep 400
        }
        if (A_Clipboard != "" && A_Clipboard != contentBefore)
            PlayCopyCompletedChime()
        if (WinExist("ahk_id " this.OriginalHwnd))
            WinActivate("ahk_id " this.OriginalHwnd)
    }

    ; R key: copy last message and read it aloud, then restore focus (same tab as delayed submit).
    CopyAndReadAloud(*) {
        this.CleanupCopyBanner()

        contentBefore := A_Clipboard
        try {
            WinActivate("ahk_id " this.GeminiHwnd)
        } catch {
            if (WinExist("ahk_id " this.OriginalHwnd))
                WinActivate("ahk_id " this.OriginalHwnd)
            return
        }
        if !WinWaitActive("ahk_exe chrome.exe", , 2) {
            if (WinExist("ahk_id " this.OriginalHwnd))
                WinActivate("ahk_id " this.OriginalHwnd)
            return
        }
        copyOpt := { restoreWindow: false, playChimeAndNotify: false, alreadyActive: true }
        if !CopyLastGeminiMessageToClipboard(copyOpt, this.GeminiHwnd) {
            if (WinExist("ahk_id " this.OriginalHwnd))
                WinActivate("ahk_id " this.OriginalHwnd)
            return
        }
        Sleep 400
        if (A_Clipboard = "" || A_Clipboard = contentBefore) {
            CopyLastGeminiMessageToClipboard(copyOpt, this.GeminiHwnd)
            Sleep 400
        }
        if (A_Clipboard = "" || A_Clipboard = contentBefore) {
            CopyLastGeminiMessageToClipboard(copyOpt, this.GeminiHwnd)
            Sleep 400
        }
        if (A_Clipboard != "" && A_Clipboard != contentBefore)
            PlayCopyCompletedChime()
        GeminiTriggerReadAloud(false, false)   ; read aloud only (already copied)
        if (WinExist("ahk_id " this.OriginalHwnd))
            WinActivate("ahk_id " this.OriginalHwnd)
    }
}

; Callable from Utils.ahk after successful auto-send (Ctrl+Alt+Win+L).
GeminiDelayedSubmitMonitorStart(originalHwnd, geminiHwnd) {
    (GeminiDelayedSubmitMonitor()).Start(originalHwnd, geminiHwnd)
}

; =============================================================================
; GeminiAsyncTTS – copy selection, send "repeat exactly" to Gemini, then trigger read aloud (Win+Alt+Shift+7)
; =============================================================================
class GeminiAsyncTTS {
    static TTSPrompt :=
        "Repeat the following text exactly as it is. Do not add any introduction, explanation, or markdown formatting. Just output the text itself:`n`n"
    static PostStreamingDelayMs := 600

    __New() {
        this.OriginalHwnd := 0
        this.GeminiHwnd := 0
        this.TimerCallback := ""
        this.RetryCount := 0
        this.MaxRetries := 60   ; 60 * 500ms = 30s timeout
        this.ButtonEverFound := false
        this.CopyCountAtSubmit := 0
    }

    Start() {
        this.OriginalHwnd := WinExist("A")
        if !this.OriginalHwnd
            return
        StandardLoadingBar_Show("Loading…", "3772FF")

        A_Clipboard := ""
        Send "^c"
        if !ClipWait(2) {
            StandardLoadingBar_Hide(0)
            return
        }
        SetTitleMatchMode(2)
        this.GeminiHwnd := GetGeminiWindowHwnd()
        if !this.GeminiHwnd {
            StandardLoadingBar_Hide(0)
            return
        }
        try {
            WinActivate("ahk_id " this.GeminiHwnd)
        } catch {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("Error: Target window not found.", 2000)
            return
        }
        if !WinWaitActive("ahk_exe chrome.exe", , 2) {
            StandardLoadingBar_Hide(0)
            return
        }
        ; For TTS from selection (#!+7), always use the trash tab (second Gemini tab) when sending the prompt.
        ; Chrome convention: Ctrl+2 selects the second tab in the window.
        Send("^2")
        Sleep 150
        ShowGeminiTabBanner(2, this.GeminiHwnd)
        uia := UIA_Browser()
        Sleep 300
        promptField := FindGeminiPromptField(uia)
        if (!promptField) {
            StandardLoadingBar_Hide(0)
            return
        }
        promptField.SetFocus()
        Sleep 100
        if (!promptField.HasKeyboardFocus) {
            try promptField.Click()
            Sleep 100
        }
        ; Paste prompt + selected text and submit
        A_Clipboard := GeminiAsyncTTS.TTSPrompt . A_Clipboard
        Sleep 100
        Send("^a")
        Sleep 500
        Send("^v")
        Sleep 500
        ; Record Copy button count before submit (used when multiple-message validation is needed).
        Send("^{End}")
        Sleep 350
        this.CopyCountAtSubmit := GetGeminiCopyButtonCount(uia)
        Send("{Enter}")
        Sleep 300
        ; Return focus to original window (activate only; do not WinRestore or we lose maximized state)
        origHwnd := this.OriginalHwnd
        try {
            if WinExist("ahk_id " origHwnd) {
                WinActivate("ahk_id " origHwnd)
                WinWaitActive("ahk_id " origHwnd, , 1)
            }
        } catch {
            if (WinExist("ahk_id " origHwnd))
                WinActivate("ahk_id " origHwnd)
        }
        this.RetryCount := 0
        this.TimerCallback := this.CheckCompletion.Bind(this)
        SetTimer(this.TimerCallback, 500)
    }

    CheckCompletion() {
        this.RetryCount++
        if (this.RetryCount > this.MaxRetries) {
            SetTimer(this.TimerCallback, 0)
            StandardLoadingBar_Hide(0)
            return
        }
        btn := ""
        buttonNames := ["Stop streaming", "Interromper transmissão", "Stop response"]
        try {
            root := UIA.ElementFromHandle(this.GeminiHwnd)
            for n in buttonNames {
                try {
                    btn := root.FindElement({ Name: n, Type: "Button" })
                } catch {
                    btn := ""
                }
                if btn
                    break
            }
        } catch {
            return
        }

        if btn {
            this.ButtonEverFound := true
            return
        }

        ; Layer 1: streaming stopped
        if (this.ButtonEverFound) {
            isTrulyGone := true
            loop 4 {
                Sleep 200
                try {
                    for n in buttonNames {
                        if root.ElementExist({ Name: n, Type: "Button" }) {
                            isTrulyGone := false
                            break
                        }
                    }
                } catch
                    isTrulyGone := true
                if !isTrulyGone
                    break
            }

            if isTrulyGone {
                SetTimer(this.TimerCallback, 0)
                StandardLoadingBar_Hide(0)
                ; Completion detection matches GeminiAsyncLookup (#!+8): Layer 1 only (Stop button gone). No extra Layer 2 so we don't miss completion.
                try {
                    if (IsSoundEnabled())
                        SoundPlay(A_ScriptDir . "\sounds\gemini-completion.wav")
                } catch {
                    PlayCopyCompletedChime()
                }
                ; Allow DOM to finish rendering, then activate trash tab and trigger read aloud (same pattern as #!+8 RetrieveResponse).
                Sleep(GeminiAsyncTTS.PostStreamingDelayMs)
                try {
                    WinActivate("ahk_id " this.GeminiHwnd)
                } catch {
                    return
                }
                if !WinWaitActive("ahk_exe chrome.exe", , 2)
                    return
                Send("^2")
                Sleep 200
                ShowGeminiTabBanner(2, this.GeminiHwnd)
                ; After TTS from selection (#!+7), read aloud from the trash tab (second Gemini tab).
                GeminiTriggerReadAloud(false, true)   ; read aloud only, no copy (text was just sent via #!+7)
            }
        }
    }
}
