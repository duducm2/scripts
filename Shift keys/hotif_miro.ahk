; =============================================================================
; Shift keys module: hotif_miro.ahk
; Miro Chrome hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

; Miro Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe chrome.exe") && InStr(WinGetTitle("A"), "Miro", false)

; (removed) Shift + Y : Command palette (Ctrl+K)

; Shift + F : Frame List (Ctrl+Shift+F)
+f:: Send "^+f"

; Shift + G : Group (Ctrl+G)
+g:: Send "^g"

; Shift + U : Ungroup (Ctrl+Shift+G)
+u:: {
    ; #region agent log (6169a3)
    try {
        title := WinGetTitle("A")
        exe := WinGetProcessName("A")
        cls := WinGetClass("A")
        shiftP := GetKeyState("Shift", "P")
        titleEsc := StrReplace(StrReplace(title, "\", "\\"), '"', '\"')
        exeEsc := StrReplace(StrReplace(exe, "\", "\\"), '"', '\"')
        clsEsc := StrReplace(StrReplace(cls, "\", "\\"), '"', '\"')
        FileAppend(
            '{"sessionId":"6169a3","runId":"run1","hypothesisId":"H1-H4","location":"Shift keys.ahk:miro:+u","message":"Miro +u fired (pre-send)","data":{"title":"' titleEsc '","exe":"' exeEsc '","class":"' clsEsc '","shiftPhysical":' (
                shiftP ? "true" : "false") '},"timestamp":' A_TickCount '}'
            "`n",
            "debug-6169a3.log",
            "UTF-8"
        )
    } catch {
    }
    ; #endregion agent log (6169a3)

    ; Using {Blind} so the physical Shift from +u contributes to Ctrl+Shift+G.
    Send "{Blind}^g"

    ; #region agent log (6169a3)
    try {
        shiftP2 := GetKeyState("Shift", "P")
        FileAppend(
            '{"sessionId":"6169a3","runId":"run1","hypothesisId":"H2-H3","location":"Shift keys.ahk:miro:+u","message":"Miro +u sent {Blind}^g (post-send)","data":{"shiftPhysicalAfter":' (
                shiftP2 ? "true" : "false") '},"timestamp":' A_TickCount '}'
            "`n",
            "debug-6169a3.log",
            "UTF-8"
        )
    } catch {
    }
    ; #endregion agent log (6169a3)
}

; Shift + L : Lock/Unlock (Ctrl+Shift+L)
+l:: Send "^+l"

; Shift + K : Add/Edit Link (Alt+Ctrl+K)
+k:: Send "!^k"

; Shift + X : Close sidebar - Close sidebar
+x:: {
    try {
        uia := UIA_Browser()
        if !IsObject(uia) {
            ; Fallback: try keyboard shortcut if UIA fails
            Send "^+s"
            return
        }

        ; Retry logic: Try multiple times with delays to allow UI to load
        maxRetries := 3
        retryDelay := 300  ; milliseconds between retries
        closeButton := ""

        loop maxRetries {
            ; Strategy 1: Find by Name "Close sidebar" with Type Button (50000)
            try {
                closeButton := uia.FindFirst({ Type: "50000", Name: "Close sidebar", cs: false })
                if (closeButton) {
                    ; Verify button is valid
                    try {
                        btnName := closeButton.Name
                        if (btnName) {
                            break  ; Found valid button, exit retry loop
                        }
                    } catch {
                        closeButton := ""
                    }
                }
            } catch {
            }

            ; Strategy 2: Try case-sensitive search
            if (!closeButton) {
                try {
                    closeButton := uia.FindFirst({ Type: "50000", Name: "Close sidebar" })
                    if (closeButton) {
                        try {
                            btnName := closeButton.Name
                            if (btnName) {
                                break
                            }
                        } catch {
                            closeButton := ""
                        }
                    }
                } catch {
                }
            }

            ; Strategy 3: Try by ControlType "Button" and Name substring
            if (!closeButton) {
                try {
                    closeButton := uia.FindFirst({ ControlType: "Button", Name: "Close sidebar", matchmode: "Substring" })
                    if (closeButton) {
                        try {
                            btnName := closeButton.Name
                            if (btnName) {
                                break
                            }
                        } catch {
                            closeButton := ""
                        }
                    }
                } catch {
                }
            }

            ; Strategy 4: Search all buttons and find by name match
            if (!closeButton) {
                try {
                    allButtons := uia.FindAll({ Type: "50000" })
                    for button in allButtons {
                        try {
                            btnName := button.Name
                            if (InStr(btnName, "Close sidebar") || InStr(btnName, "close sidebar")) {
                                closeButton := button
                                break
                            }
                        } catch {
                            continue
                        }
                    }
                    if (closeButton) {
                        break
                    }
                } catch {
                }
            }

            Sleep retryDelay  ; Wait before next attempt
        }

        ; Confirmation layer: Verify button was found before clicking
        if (closeButton) {
            ; Additional verification: ensure button is still valid and clickable
            try {
                ; Check if button is enabled and visible
                isEnabled := closeButton.GetPropertyValue(UIA.Property.IsEnabled)
                isOffscreen := closeButton.GetPropertyValue(UIA.Property.IsOffscreen)

                if (!isEnabled || isOffscreen) {
                    ; Button found but not usable, try keyboard shortcut fallback
                    Send "^+s"
                    return
                }
            } catch {
                ; Property check failed, continue with click attempt
            }

            ; Try multiple click strategies in order of preference
            clicked := false

            ; Strategy 1: Try Invoke pattern (most reliable for buttons)
            try {
                closeButton.Invoke()
                clicked := true
            } catch {
            }

            ; Strategy 2: Try SetFocus then Click
            if (!clicked) {
                try {
                    closeButton.SetFocus()
                    Sleep 50
                    closeButton.Click()
                    clicked := true
                } catch {
                }
            }

            ; Strategy 3: Force coordinate-based click using "left" parameter
            if (!clicked) {
                try {
                    closeButton.Click("left")
                    clicked := true
                } catch {
                }
            }

            ; Strategy 4: Direct coordinate click using element Location
            if (!clicked) {
                try {
                    pos := closeButton.Location
                    if (pos && pos.w > 0 && pos.h > 0) {
                        ; Save current mouse position
                        MouseGetPos(&prevX, &prevY)

                        ; Click at center of element
                        CoordMode("Mouse", "Screen")
                        Click(pos.x + pos.w // 2, pos.y + pos.h // 2)
                        Sleep 50

                        ; Restore mouse position
                        MouseMove(prevX, prevY)
                        clicked := true
                    }
                } catch {
                }
            }

            ; If all click strategies failed, use keyboard shortcut fallback
            if (!clicked) {
                Send "^+s"
            }
        } else {
            ; Button not found after all retries, use keyboard shortcut fallback
            Send "^+s"
        }
    } catch Error as err {
        ; If any error occurs, use keyboard shortcut as fallback
        Send "^+s"
    }
}

#HotIf
