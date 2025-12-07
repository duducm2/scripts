#Requires AutoHotkey v2.0
#SingleInstance Force

; --- Includes ----------------------------------------------------------------
#include UIA-v2\Lib\UIA.ahk
#include UIA-v2\Lib\UIA_Browser.ahk

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

; --- Hotkeys ----------------------------------------------------------------

; Win+Alt+Shift+7 : Read aloud the last message in Gemini (activates Gemini window first, then clicks last "Show more options" then "Text to speech")
; If reading is active, clicking this shortcut again will pause the reading
#!+7:: {
    try {
        ; Step 1: Activate Gemini window globally
        SetTitleMatchMode(2)
        if hwnd := GetGeminiWindowHwnd()
            WinActivate("ahk_id " hwnd)
        if !WinWaitActive("ahk_exe chrome.exe", , 2) {
            return
        }
        Sleep 300

        ; Step 2: Check if "Pause" button exists (if reading is active, pause it)
        uia := UIA_Browser()
        Sleep 300

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

        ; Step 3: Find all "Show more options" buttons (normal read-aloud flow)
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

        ; Step 4: Click the last "Show more options" button
        lastMoreOptionsButton.Click()
        Sleep 400 ; Wait for menu to appear

        ; Step 5: Find and click the "Text to speech" menu item
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
                if InStr(menuItem.Name, "Text to speech") || InStr(menuItem.Name, "Texto para fala") || InStr(menuItem.Name,
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

; Win+Alt+Shift+J : Click the last Copy button in Gemini (activates Gemini window first, then copies the preceding message)
#!+j:: {
    try {
        ; Step 1: Activate Gemini window globally
        SetTitleMatchMode(2)
        if hwnd := GetGeminiWindowHwnd()
            WinActivate("ahk_id " hwnd)
        if !WinWaitActive("ahk_exe chrome.exe", , 2) {
            return
        }
        Sleep 300

        ; Step 2: Find all Copy buttons
        uia := UIA_Browser()
        Sleep 300

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
