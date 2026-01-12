#Requires AutoHotkey v2.0
#SingleInstance Force

; --- Includes ----------------------------------------------------------------
#include UIA-v2\Lib\UIA.ahk
#include UIA-v2\Lib\UIA_Browser.ahk
#include %A_ScriptDir%\env.ahk

; --- Config ---------------------------------------------------------------
; Path to the file containing the initial prompt Gemini should receive.
PROMPT_FILE := A_ScriptDir "\data\Gemini_Prompt.txt"

; --- Helper Functions --------------------------------------------------------

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
; Unified banner builder – consistent shape/font/opacity for all banners here
; =============================================================================
CreateCenteredBanner(message, bgColor := "3772FF", fontColor := "FFFFFF", fontSize := 24, alpha := 178) {
    bGui := Gui()
    bGui.Opt("+AlwaysOnTop -Caption +ToolWindow")
    bGui.BackColor := bgColor
    bGui.SetFont("s" . fontSize . " c" . fontColor . " Bold", "Segoe UI")
    bGui.Add("Text", "w500 Center", message)

    activeWin := WinGetID("A")
    if (activeWin) {
        WinGetPos(&winX, &winY, &winW, &winH, activeWin)
    } else {
        workArea := SysGet.MonitorWorkArea(SysGet.MonitorPrimary)
        winX := workArea.Left, winY := workArea.Top, winW := workArea.Right - workArea.Left, winH := workArea.Bottom -
            workArea.Top
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

        played := false
        ; Prefer Windows MessageBeep (reliable through default output)
        try {
            rc := DllCall("User32\\MessageBeep", "UInt", 0xFFFFFFFF)
            if (rc)
                played := true
        } catch {
            ; Silently ignore errors
        }

        ; Fallback to system asterisk sound
        if !played {
            try {
                played := SoundPlay("*64", false)
            } catch {
                ; Silently ignore errors
            }
        }

        ; Last resort, attempt the classic beep
        if !played {
            try SoundBeep(1100, 130)
            catch {
            }
        }
    } catch {
        ; Silently ignore errors
    }
}

; =============================================================================
; Small Loading Indicator Helpers
; =============================================================================
global smallLoadingGuis_Gemini := []

ShowSmallLoadingIndicator(state := "Loading…", bgColor := "3772FF") {
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

    ; Create a single, high-contrast, centered banner using the unified builder
    textGui := CreateCenteredBanner(state, bgColor, "FFFFFF", 24, 178)
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

; Win+Alt+Shift+O : Read aloud and copy the last message in Gemini (activates Gemini window first, copies the message, then clicks last "Show more options" then "Text to speech")
; If reading is active, clicking this shortcut again will pause the reading
#!+o:: {

    try {
        ; Step 1: Activate Gemini window globally
        SetTitleMatchMode(2)
        if hwnd := GetGeminiWindowHwnd()
            WinActivate("ahk_id " hwnd)
        if !WinWaitActive("ahk_exe chrome.exe", , 2) {
            return
        }
        Sleep 150  ; small settle per README (keep this snappy)

        ; Step 2: Check if "Pause" button exists (if reading is active, pause it)
        uia := UIA_Browser()
        Sleep 120  ; minimal settle before querying UIA

        ; Try to find "Pause" button
        pauseButton := 0
        try {
            ; Primary strategy: Find by Name "Pause" with Type 50000 (Button)
            pauseButton := uia.FindFirst({ Name: "Pause", Type: 50000 })

            ; Fallback: Try by Type "Button" and Name "Pause"
            if !pauseButton {
                pauseButton := uia.FindFirst({ Type: "Button", Name: "Pause" })
            }

            ; Fallback: Search all buttons for one with "Pause" name and tts-button className
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
            ; Silently continue if pause button search fails
        }

        ; If "Pause" button found, click it and return to previous window
        if (pauseButton) {
            pauseButton.Click()
            ShowNotification("Paused", 800, "FFFF00", "000000", 24)
            Send "!{Tab}"
            return
        }

        ; Try to find "Resume" button (if reading is paused, resume it)
        resumeButton := 0
        try {
            ; Primary strategy: Find by Name "Resume" with Type 50000 (Button)
            resumeButton := uia.FindFirst({ Name: "Resume", Type: 50000 })

            ; Fallback: Try by Type "Button" and Name "Resume"
            if !resumeButton {
                resumeButton := uia.FindFirst({ Type: "Button", Name: "Resume" })
            }

            ; Fallback: Search all buttons for one with "Resume" name and tts-button className
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
            ; Silently continue if resume button search fails
        }

        ; If "Resume" button found, click it and return to previous window
        if (resumeButton) {
            resumeButton.Click()
            ShowNotification("Resumed", 800, "FFFF00", "000000", 24)
            Send "!{Tab}"
            return
        }

        ; Step 3: Find and click the last Copy button (copy the message we're about to read)
        allCopyButtons := []

        ; Primary strategy: Single pass - find all buttons and filter for any "Copy" variant
        allButtons := uia.FindAll({ Type: 50000 })
        for button in allButtons {
            if (button.Name = "Copy" || InStr(button.Name, "Copy", false)) {
                ; Additional check: ensure it has the Copy button className pattern
                if (InStr(button.ClassName, "icon-button") || InStr(button.ClassName, "mdc-button")) {
                    allCopyButtons.Push(button)
                }
            }
        }

        ; Fallback: broaden type only if none found on primary pass
        if (allCopyButtons.Length = 0) {
            allButtons := uia.FindAll({ Type: "Button" })
            for button in allButtons {
                if (button.Name = "Copy" || InStr(button.Name, "Copy", false)) {
                    allCopyButtons.Push(button)
                }
            }
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

        ; Click the copy button if found
        if (lastCopyButton) {
            lastCopyButton.Click()
            ; Play chime when copy completes
            PlayCopyCompletedChime()
        }

        ; Step 4: Find all "Show more options" elements (normal read-aloud flow)
        ; Show banner while searching
        searchBanner := CreateCenteredBanner("Finding read aloud button and copying...", "3772FF", "FFFFFF", 24, 178)
        Sleep 250  ; small delay to ensure UI is ready without feeling sluggish

        allMoreOptionsButtons := []

        ; Primary strategy: Search directly by name (most efficient - finds 8 elements vs searching 120+ buttons)
        try {
            allMoreOptionsButtons := uia.FindAll({ Name: "Show more options" })
        } catch {
            ; If direct name search fails, try Type 50011 (MenuItem) as fallback
            allMenuItems := uia.FindAll({ Type: 50011 })
            for menuItem in allMenuItems {
                if (menuItem.Name = "Show more options" || InStr(menuItem.Name, "Show more options", false) = 1) {
                    allMoreOptionsButtons.Push(menuItem)
                }
            }
        }

        if (allMoreOptionsButtons.Length = 0) {
            ; No "Show more options" buttons found
            if IsObject(searchBanner) && searchBanner.Hwnd {
                searchBanner.Destroy()
            }
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
            if IsObject(searchBanner) && searchBanner.Hwnd {
                searchBanner.Destroy()
            }
            return
        }

        ; Step 5 & 6: Click "Show more options" and navigate to "Text to speech"
        ; Note: For first-time read-aloud, Gemini requires doing this twice
        loop 2 {
            ; Click the last "Show more options" button
            lastMoreOptionsButton.Click()
            Sleep 200 ; Wait for menu to appear

            ; Hide search banner after first attempt
            if (A_Index = 1 && IsObject(searchBanner) && searchBanner.Hwnd) {
                searchBanner.Destroy()
            }

            ; Find and click "Text to speech" menu item using UIA
            textToSpeechMenuItem := 0
            try {
                ; Primary strategy: Find by Name "Text to speech" with Type 50011 (MenuItem)
                textToSpeechMenuItem := uia.FindFirst({ Name: "Text to speech", Type: 50011 })
            } catch {
                ; Silently continue to fallbacks
            }

            ; Fallback 1: Try by Type "MenuItem" and Name "Text to speech"
            if !textToSpeechMenuItem {
                try {
                    textToSpeechMenuItem := uia.FindFirst({ Type: "MenuItem", Name: "Text to speech" })
                } catch {
                    ; Silently continue
                }
            }

            ; Fallback 2: Search all MenuItems for one with "Text to speech" name and matching className
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
                    ; Silently continue
                }
            }

            ; Fallback 3: Search all MenuItems by name only (broader match)
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
                    ; Silently continue
                }
            }

            ; Click the menu item if found
            if (textToSpeechMenuItem) {
                textToSpeechMenuItem.Click()
                Sleep 200 ; Brief pause to ensure menu action completes
            } else {
                ; If UIA method fails, fallback to keyboard navigation
                Send "{Down}"
                Sleep 200
                Send "{Enter}"
            }

            ; Wait before next attempt (if needed) or before finishing
            if (A_Index = 1) {
                Sleep 200 ; Wait after first attempt before retrying
            }
        }

        ; Show notification that both copy and read-aloud actions completed
        ShowNotification("Copied & Reading aloud", 800, "FFFF00", "000000", 24)
        ; Return to previous window
        Send "!{Tab}"
    } catch Error as e {
        ; If all else fails, silently fail (no fallback action defined)
    }
}

; Win+Alt+Shift+P : Click the last Copy button in Gemini (activates Gemini window first, then copies the preceding message)
#!+p:: {
    try {
        ; Step 1: Activate Gemini window globally
        SetTitleMatchMode(2)
        if hwnd := GetGeminiWindowHwnd()
            WinActivate("ahk_id " hwnd)
        if !WinWaitActive("ahk_exe chrome.exe", , 2) {
            return
        }
        Sleep 150  ; keep snappy per README (short settle after activation)

        ; Step 2: Find all Copy buttons
        uia := UIA_Browser()
        Sleep 120  ; minimal settle before querying UIA

        allCopyButtons := []

        ; Primary pass: Find all buttons, filter for any "Copy" variant
        allButtons := uia.FindAll({ Type: 50000 })
        for button in allButtons {
            if (button.Name = "Copy" || InStr(button.Name, "Copy", false)) {
                ; Additional check: ensure it has the Copy button className pattern
                if (InStr(button.ClassName, "icon-button") || InStr(button.ClassName, "mdc-button")) {
                    allCopyButtons.Push(button)
                }
            }
        }

        ; Fallback: broaden type if none found on primary pass (still single filter loop)
        if (allCopyButtons.Length = 0) {
            allButtons := uia.FindAll({ Type: "Button" })
            for button in allButtons {
                if (button.Name = "Copy" || InStr(button.Name, "Copy", false)) {
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
            ; Play chime when copy completes
            PlayCopyCompletedChime()
            ; Show notification banner when copy button is clicked
            ShowNotification("Copied!", 800, "FFFF00", "000000", 24)
            ; Return to previous window
            Send "!{Tab}"
        } else {
            ; Last resort: Could not find last Copy button
        }
    } catch Error as e {
        ; If all else fails, silently fail (no fallback action defined)
    }
}

; =============================================================================
; Get Pronunciation
; Hotkey: Win+Alt+Shift+8
; =============================================================================
#!+8:: {
    A_Clipboard := ""
    Send "^c"
    ClipWait
    SetTitleMatchMode(2)
    if hwnd := GetGeminiWindowHwnd()
        WinActivate("ahk_id " hwnd)
    if !WinWaitActive("ahk_exe chrome.exe", , 2)
        return

    ; Find the Gemini prompt field
    uia := UIA_Browser()
    Sleep 300

    promptField := 0

    ; Primary strategy: Find by Name "Enter a prompt here" with Type 50004 (Edit)
    promptField := uia.FindFirst({ Name: "Enter a prompt here", Type: 50004 })

    ; Fallback 1: Try by Type "Edit" and Name "Enter a prompt here"
    if !promptField {
        promptField := uia.FindFirst({ Type: "Edit", Name: "Enter a prompt here" })
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
            if InStr(edit.Name, "Enter a prompt") || InStr(edit.Name, "Digite um prompt") || InStr(edit.Name, "prompt") {
                ; Additional check to ensure it's the prompt field (has ql-editor in className)
                if InStr(edit.ClassName, "ql-editor") {
                    promptField := edit
                    break
                }
            }
        }
    }

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

    searchString :=
        "Below, you will find a word or phrase. I'd like you to answer in five sections: the 1st section you will repeat the word twice. For each time you repeat, use a point to finish the phrase. The 2nd section should have the definition of the word (You should also say each part of speech does the different definitions belong to). The 3d section should have the pronunciation of this word using the Internation Phonetic Alphabet characters (for American English).The 4th section should have the same word applied in a real sentence (put that in quotations, so I can identify that). In the 5th, Write down the translation of the word into Portuguese. Please, do not title any section. Thanks!"
    A_Clipboard := searchString . "`n`nContent: " . A_Clipboard
    Sleep 100
    Send("^a")
    Sleep 500
    Send("^v")
    Sleep 500
    Send("{Enter}")
    Sleep 500
    ; After sending, show loading for Stop streaming
    Send "!{Tab}" ; Return to previous window
    buttonNames := ["Stop streaming", "Interromper transmissão"]
    WaitForButtonAndShowSmallLoading(buttonNames, "Waiting for response...")

    ; Go back to previous window
    Send "!{Tab}"
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

        ; Find and focus the Gemini prompt field
        uia := UIA_Browser()
        Sleep 300

        promptField := 0

        ; Primary strategy: Find by Name "Enter a prompt here" with Type 50004 (Edit)
        promptField := uia.FindFirst({ Name: "Enter a prompt here", Type: 50004 })

        ; Fallback 1: Try by Type "Edit" and Name "Enter a prompt here"
        if !promptField {
            promptField := uia.FindFirst({ Type: "Edit", Name: "Enter a prompt here" })
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
            CenterMouse()

            ; Focus the Gemini prompt field - optimized for speed
            Sleep 150  ; small settle per README (keep snappy)
            uia := UIA_Browser()
            Sleep 120  ; minimal settle before querying UIA

            promptField := 0

            ; Primary strategy: Find by Name "Enter a prompt here" with Type 50004 (Edit)
            try {
                promptField := uia.FindFirst({ Name: "Enter a prompt here", Type: 50004 })
            } catch {
                ; Silently continue to fallbacks
            }

            ; If primary strategy failed, use single-pass scoring over all edits
            if !promptField {
                try {
                    allEdits := uia.FindAll({ Type: 50004 })
                    best := 0, bestScore := -1
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
                        if (score > bestScore) {
                            bestScore := score
                            best := edit
                        }
                    }
                    if (bestScore >= 0) {
                        promptField := best
                    }
                } catch {
                    ; Silently continue if fallback fails
                }
            }

            if (promptField) {
                promptField.SetFocus()
                Sleep 50 ; Reduced from 100ms
                ; Ensure focus was successful
                if (!promptField.HasKeyboardFocus) {
                    ; Fallback: try clicking if SetFocus didn't work
                    promptField.Click()
                    Sleep 50 ; Reduced from 100ms
                }
                ; Play chime when field is successfully focused
                PlayCopyCompletedChime()
            }
        }
    } else {
        InitializeGeminiFirstTime()
    }
}
