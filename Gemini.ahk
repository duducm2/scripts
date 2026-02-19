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
    global smallLoadingGuis_Gemini
    if (smallLoadingGuis_Gemini.Length > 0) {
        for gui in smallLoadingGuis_Gemini {
            try gui.Destroy()
            catch {
                ; Silently ignore GUI destroy errors
            }
        }
        smallLoadingGuis_Gemini := [] ; Reset the array
    }
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
                ShowSmallLoadingIndicator(stateText)
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
    HideSmallLoadingIndicator()
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
GeminiTriggerReadAloud(copyFirst := true) {
    ; Step 1: Activate Gemini window globally
    SetTitleMatchMode(2)
    if hwnd := GetGeminiWindowHwnd()
        WinActivate("ahk_id " hwnd)
    if !WinWaitActive("ahk_exe chrome.exe", , 2)
        return
    Sleep 150

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
            WinActivate("ahk_id " geminiHwnd)
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
; #region agent log
_DebugLog_Gemini(msg, data := "") {
    path := A_ScriptDir "\.cursor\debug.log"
    line := '{"location":"Gemini.ahk","message":"' . msg . '","data":' . (data = "" ? "{}" : data) . ',"timestamp":' .
    A_TickCount . '}' . "`n"
    FileAppend line, path
}
; #endregion
; Path for bridge to verify that Copy Last Response (same as #!+p) actually succeeded
GEMINI_COPY_RESULT_PATH := A_ScriptDir "\.cursor\gemini_copy_result.txt"

OnMessage(WM_COPY_LAST_GEMINI, copyFromBridge)
copyFromBridge(*) {
    _DebugLog_Gemini("WM_COPY_LAST_GEMINI received", "{}")
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
    _DebugLog_Gemini("CopyLastGeminiMessageToClipboard result", (r ? '{"ok":1}' : '{"ok":0}'))
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
InitializeGeminiFirstTime() {
    try {
        ; Show banner to inform user
        ShowSmallLoadingIndicator("Opening Gemini...")

        ; Run Chrome with new window
        Run "chrome.exe --new-window https://gemini.google.com/"
        if !WinWaitActive("ahk_exe chrome.exe", , 5) {
            HideSmallLoadingIndicator()
            return
        }

        ; Get the Gemini window handle
        geminiHwnd := WinExist("A")
        if !geminiHwnd {
            HideSmallLoadingIndicator()
            return
        }

        ; Activate the Gemini window
        WinActivate("ahk_id " geminiHwnd)
        WinWaitActive("ahk_id " geminiHwnd, , 2)
        Sleep 200 ; Give window time to fully activate

        ; Update banner status
        ShowSmallLoadingIndicator("Loading Gemini page...")

        ; Wait for page to load fully
        Sleep 300

        ; Find and focus the Gemini prompt field (EN/PT aware)
        uia := UIA_Browser()
        Sleep 300

        promptField := FindGeminiPromptField(uia)

        if (promptField) {
            ; Focus the prompt field
            promptField.SetFocus()
            Sleep 100
            ; Ensure focus was successful
            if (!promptField.HasKeyboardFocus) {
                ; Fallback: try clicking if SetFocus didn't work
                promptField.Click()
                Sleep 100
            }
        }

        ; Update banner status
        ShowSmallLoadingIndicator("Sending initial prompt...")

        ; Read initial prompt from external file & paste it
        promptText := ""
        try promptText := FileRead(PROMPT_FILE, "UTF-8")
        if (StrLen(promptText) = 0)
            promptText := "hey, what's up?"

        ; Copy–paste to handle Unicode & speed
        oldClip := A_Clipboard
        A_Clipboard := ""
        A_Clipboard := promptText
        ClipWait 1
        Send("^v")
        Sleep 100
        Send("{Enter}")
        Sleep 100
        A_Clipboard := oldClip

        ; Hide banner on success
        HideSmallLoadingIndicator()
    } catch Error as err {
        ; Hide banner on error
        HideSmallLoadingIndicator()
    }
}

; =============================================================================
; Open Gemini
; Hotkey: Win+Alt+Shift+I
; =============================================================================
#!+i:: {
    SetTitleMatchMode(2)
    if hwnd := GetGeminiWindowHwnd() {
        WinActivate("ahk_id " hwnd)
        if WinWaitActive("ahk_id " hwnd, , 2) {
            ; Focus the Gemini prompt field using Anchor & Backtrack strategy
            ; Strategy: Find "Open upload file menu" button (anchor), focus it, then Shift+Tab to prompt field
            uia := UIA_Browser()
            Sleep 80   ; combined settle time for UIA initialization

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

                    ; Play sound (non-blocking, no try-catch needed - SoundPlay is safe)
                    if (IsSoundEnabled()) {
                        SoundPlay(A_ScriptDir . "\sounds\gemini-focused.wav")
                    }
                } catch {
                    ; Fallback: direct prompt field search if anchor strategy fails (EN/PT aware)
                    promptField := FindGeminiPromptField(uia)
                    if (promptField) {
                        try promptField.SetFocus()
                        if (IsSoundEnabled())
                            SoundPlay(A_ScriptDir . "\sounds\gemini-focused.wav")
                    }
                }
            } else {
                ; Fallback: direct prompt field search if anchor not found (EN/PT aware)
                promptField := FindGeminiPromptField(uia)
                if (promptField) {
                    try promptField.SetFocus()
                    if (IsSoundEnabled())
                        SoundPlay(A_ScriptDir . "\sounds\gemini-focused.wav")
                }
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
        ShowSmallLoadingIndicator("Loading…`n`n⚠️ Please do not click or use the keyboard", "3772FF", this.OriginalHwnd,
            200, 16)

        A_Clipboard := ""
        Send "^c"
        if !ClipWait(2)
            return
        SetTitleMatchMode(2)
        this.GeminiHwnd := GetGeminiWindowHwnd()
        if !this.GeminiHwnd {
            HideSmallLoadingIndicator()
            return
        }
        WinActivate("ahk_id " this.GeminiHwnd)
        if !WinWaitActive("ahk_exe chrome.exe", , 2) {
            HideSmallLoadingIndicator()
            return
        }
        uia := UIA_Browser()
        Sleep 300
        promptField := FindGeminiPromptField(uia)
        if (!promptField) {
            HideSmallLoadingIndicator()
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
        ; Go back to the window where you triggered the hotkey so you can keep working
        origHwnd := this.OriginalHwnd
        try {
            if WinExist("ahk_id " origHwnd) {
                WinRestore("ahk_id " origHwnd)
                WinActivate("ahk_id " origHwnd)
                WinWaitActive("ahk_id " origHwnd, , 1)
            }
        } catch {
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
            HideSmallLoadingIndicator()
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
        WinActivate("ahk_id " this.GeminiHwnd)
        if !WinWaitActive("ahk_exe chrome.exe", , 2) {
            HideSmallLoadingIndicator()
            return
        }
        copyOpt := { restoreWindow: false, playChimeAndNotify: false, alreadyActive: true }
        if !CopyLastGeminiMessageToClipboard(copyOpt, this.GeminiHwnd) {
            HideSmallLoadingIndicator()
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
        HideSmallLoadingIndicator()
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
; GeminiAsyncTTS – copy selection, send "repeat exactly" to Gemini, then trigger read aloud (Win+Alt+Shift+7)
; =============================================================================
class GeminiAsyncTTS {
    static TTSPrompt :=
        "Repeat the following text exactly as it is. Do not add any introduction, explanation, or markdown formatting. Just output the text itself:`n`n"

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
        ShowSmallLoadingIndicator("Loading…`n`n⚠️ Please do not click or use the keyboard", "3772FF", this.OriginalHwnd,
            200, 16)

        A_Clipboard := ""
        Send "^c"
        if !ClipWait(2) {
            HideSmallLoadingIndicator()
            return
        }
        SetTitleMatchMode(2)
        this.GeminiHwnd := GetGeminiWindowHwnd()
        if !this.GeminiHwnd {
            HideSmallLoadingIndicator()
            return
        }
        WinActivate("ahk_id " this.GeminiHwnd)
        if !WinWaitActive("ahk_exe chrome.exe", , 2) {
            HideSmallLoadingIndicator()
            return
        }
        uia := UIA_Browser()
        Sleep 300
        promptField := FindGeminiPromptField(uia)
        if (!promptField) {
            HideSmallLoadingIndicator()
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
        Send("{Enter}")
        Sleep 300
        origHwnd := this.OriginalHwnd
        try {
            if WinExist("ahk_id " origHwnd) {
                WinRestore("ahk_id " origHwnd)
                WinActivate("ahk_id " origHwnd)
                WinWaitActive("ahk_id " origHwnd, , 1)
            }
        } catch {
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
            HideSmallLoadingIndicator()
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
                HideSmallLoadingIndicator()
                try {
                    if (IsSoundEnabled())
                        SoundPlay(A_ScriptDir . "\sounds\gemini-completion.wav")
                } catch {
                    PlayCopyCompletedChime()
                }
                GeminiTriggerReadAloud(false)   ; read aloud only, no copy (text was just sent via #!+7)
            }
        }
    }
}
