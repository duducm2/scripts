; =============================================================================
; Shift keys module: hotif_google_keep.ahk
; Google Keep hotkeys and reminder dismiss helpers
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

;-------------------------------------------------------------------
; Google Keep Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe chrome.exe") && (WinActive("Google Keep") || WinActive("keep.google.com") || InStr(
    WinGetTitle("A"), "Google Keep"))

; Shift + S : Search and select note
+s::
{
    ; Store the current active window handle
    currentWindow := WinExist("A")

    ; Show message box to get search text from user
    searchText := InputBox("Enter text to search for in your notes:", "Google Keep Search", "w300 h100")

    if (searchText.Result = "OK" && searchText.Value != "") {
        ; Explicitly activate the Google Keep window to ensure we're working with the right window
        WinActivate("ahk_id " currentWindow)
        WinWaitActive("ahk_id " currentWindow, , 2)

        ; Store the search text in clipboard
        oldClip := A_Clipboard
        A_Clipboard := searchText.Value

        ; Wait a moment for clipboard to be ready
        Sleep 200

        ; Send Escape to clear any current selection/focus
        SendEscape()
        Sleep 300

        ; Open search with Ctrl+F
        Send "^f"
        Sleep 200

        ; Paste the search text
        Send "^v"
        Sleep 900

        ; Press Escape to close search
        SendEscape()
        Sleep 300

        ; Press Enter to confirm selection
        Send "{Enter}"
        Sleep 300

        ; Restore original clipboard
        A_Clipboard := oldClip
    }
}

; Shift + M : Toggle main menu
+m::
{
    try {
        ; Store the current active window handle
        currentWindow := WinExist("A")

        ; Explicitly activate the Google Keep window to ensure we're working with the right window
        WinActivate("ahk_id " currentWindow)
        WinWaitActive("ahk_id " currentWindow, , 2)

        ; Use UIA to find and click the main menu button
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 300 ; Give UIA time to attach

        ; Find the main menu button by its properties
        mainMenuBtn := uia.FindElement({
            Name: "Main menu",
            Type: "Button",
            ClassName: "gb_Lc"
        })

        if (mainMenuBtn) {
            mainMenuBtn.Click()
        } else {
            ; Fallback: try to find by name only
            mainMenuBtn := uia.FindElement({ Name: "Main menu", Type: "Button" })
            if (mainMenuBtn) {
                mainMenuBtn.Click()
            } else {
                MsgBox "Could not find the Main menu button.", "Google Keep", "IconX"
            }
        }
    } catch Error as e {
        MsgBox "Error toggling main menu: " e.Message, "Google Keep Error", "IconX"
    }
}

#HotIf

ConfirmDismissAll() {
    if MsgBox("Dismiss all reminders?", "Confirm Dismiss", "YesNo Icon?") = "Yes"
        DismissAllReminders()
}

DismissAllReminders() {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)
        ; Try by AutomationId first
        btn := root.FindFirst({ AutomationId: "8345", ControlType: "Button" })
        ; Fallback: search by name
        if !btn
            btn := root.FindFirst({ Name: "Dismiss All", ControlType: "Button" })
        if !btn
            btn := root.FindFirst({ Name: "Dismiss all", ControlType: "Button" })
        if btn {
            btn.Click()
        } else {
            MsgBox("Could not find the 'Dismiss All' button.", "Dismiss All", "IconX")
        }
    } catch Error as e {
        MsgBox("UIA error:`n" e.Message, "Dismiss All Error", "IconX")
    }
}

; ---------------------------------------------------------------------------
; Helper for Mobills buttons â€" language-neutral search
; ---------------------------------------------------------------------------
GetMobillsButton(autoId, btnName) {
    try {
        root := UIA.ElementFromHandle(WinExist("A"))
        btn := root.FindFirst({ AutomationId: autoId, ControlType: "Button" })
        if !btn
            btn := root.FindFirst({ Name: btnName, ControlType: "Button" })
        return btn
    } catch Error {
        return ""
    }
}

;-------------------------------------------
; Helper functions
;-------------------------------------------
; When Mobills is visible but not focused, UIA_Browser attach can fail.
; Recovery: activate browser window, retry once. (No clicking.)
Mobills_ActivateBrowserForAttach(exe := "ahk_exe chrome.exe", titleNeedle := "Mobills") {
    try {
        bestHwnd := 0
        hwnds := WinGetList(exe)
        for hwnd in hwnds {
            try {
                t := WinGetTitle("ahk_id " hwnd)
                if (t != "" && InStr(t, titleNeedle)) {
                    bestHwnd := hwnd
                    break
                }
            } catch {
            }
        }
        if (!bestHwnd)
            return false

        WinActivate("ahk_id " bestHwnd)
        WinWaitActive("ahk_id " bestHwnd, , 1)
        return true
    } catch {
        return false
    }
}

TryAttachBrowser() {
    ; #region agent log
    try {
        FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' . A_TickCount .
        ',"location":"Shift keys.ahk:9806","message":"TryAttachBrowser entry","data":{"attempt":"chrome"},"sessionId":"debug-session","runId":"run1","hypothesisId":"A"}`n',
        DEBUG_LOG_PATH
    } catch {
    }
    ; #endregion
    ; Attach to the focused Chrome/Edge window only — never a background Mobills tab.
    try {
        result := ""
        try {
            hwnd := WinExist("A")
            if hwnd {
                exe := StrLower(WinGetProcessName("ahk_id " hwnd))
                if (exe = "chrome.exe" || exe = "msedge.exe")
                    result := UIA_Browser("ahk_id " hwnd)
            }
        }
        if (result) {
            ; #region agent log
            try {
                FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' .
                A_TickCount .
                ',"location":"Shift keys.ahk:9814","message":"TryAttachBrowser success","data":{"browser":"chrome","result":' .
                (result ? 1 : 0) . '},"sessionId":"debug-session","runId":"run1","hypothesisId":"A"}`n',
                DEBUG_LOG_PATH
            } catch {
            }
            ; #endregion
            return result
        }

        ; Recovery: activate + retry
        Mobills_ActivateBrowserForAttach("ahk_exe chrome.exe", "Mobills")
        try result := UIA_Browser("ahk_exe chrome.exe")
        ; #region agent log
        try {
            FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' . A_TickCount .
            ',"location":"Shift keys.ahk:9814","message":"TryAttachBrowser success","data":{"browser":"chrome","result":' .
            (result ? 1 : 0) . '},"sessionId":"debug-session","runId":"run1","hypothesisId":"A"}`n',
            DEBUG_LOG_PATH
        } catch {
        }
        ; #endregion
        return result
    }
    catch {
        ; #region agent log
        try {
            FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' . A_TickCount .
            ',"location":"Shift keys.ahk:9822","message":"TryAttachBrowser chrome failed, trying edge","data":{"error":"' .
            A_LastError . '"},"sessionId":"debug-session","runId":"run1","hypothesisId":"A"}`n', DEBUG_LOG_PATH
        } catch {
        }
        ; #endregion
        try {
            result := ""
            try result := UIA_Browser("ahk_exe msedge.exe")
            if (!result) {
                Mobills_ActivateBrowserForAttach("ahk_exe msedge.exe", "Mobills")
                try result := UIA_Browser("ahk_exe msedge.exe")
            }
            ; #region agent log
            try {
                FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' .
                A_TickCount .
                ',"location":"Shift keys.ahk:9829","message":"TryAttachBrowser success","data":{"browser":"edge","result":' .
                (result ? 1 : 0) . '},"sessionId":"debug-session","runId":"run1","hypothesisId":"A"}`n',
                DEBUG_LOG_PATH
            } catch {
            }
            ; #endregion
            return result
        }
        catch {
            ; #region agent log
            try {
                FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' .
                A_TickCount .
                ',"location":"Shift keys.ahk:9837","message":"TryAttachBrowser failed both browsers","data":{"error":"' .
                A_LastError . '"},"sessionId":"debug-session","runId":"run1","hypothesisId":"A"}`n',
                DEBUG_LOG_PATH
            } catch {
            }
            ; #endregion
            return 0
        }
    }
}

FindMonthGroup(uia) {
    ; #region agent log
    try {
        FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' . A_TickCount .
        ',"location":"Shift keys.ahk:9848","message":"FindMonthGroup entry","data":{"uia":' . (uia ? 1 : 0) .
        '},"sessionId":"debug-session","runId":"run1","hypothesisId":"B"}`n', DEBUG_LOG_PATH
    } catch {
    }
    ; #endregion
    ; Strategy 1 â€" look for known class name on the container
    try {
        ; #region agent log
        try {
            FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' . A_TickCount .
            ',"location":"Shift keys.ahk:9855","message":"FindMonthGroup Strategy 1 attempt","data":{"className":"sc-kAyceB"},"sessionId":"debug-session","runId":"run1","hypothesisId":"B"}`n',
            DEBUG_LOG_PATH
        } catch {
        }
        ; #endregion
        grp := uia.FindElement({ Type: "Group", ClassName: "sc-kAyceB", matchmode: "Substring" })
        if grp {
            ; #region agent log
            try {
                FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' .
                A_TickCount .
                ',"location":"Shift keys.ahk:9862","message":"FindMonthGroup Strategy 1 success","data":{"found":true},"sessionId":"debug-session","runId":"run1","hypothesisId":"B"}`n',
                DEBUG_LOG_PATH
            } catch {
            }
            ; #endregion
            return grp
        }
        ; #region agent log
        try {
            FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' . A_TickCount .
            ',"location":"Shift keys.ahk:9862","message":"FindMonthGroup Strategy 1 no result","data":{"found":false},"sessionId":"debug-session","runId":"run1","hypothesisId":"B"}`n',
            DEBUG_LOG_PATH
        } catch {
        }
        ; #endregion
    }
    catch Error as e {
        ; #region agent log
        try {
            FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' . A_TickCount .
            ',"location":"Shift keys.ahk:9876","message":"FindMonthGroup Strategy 1 exception","data":{"error":"' .
            e.Message .
            '"},"sessionId":"debug-session","runId":"run1","hypothesisId":"B"}`n', DEBUG_LOG_PATH
        } catch {
        }
        ; #endregion
    }
    ; Strategy 2 â€" locate by month text (any language)
    ; #region agent log
    try {
        FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' . A_TickCount .
        ',"location":"Shift keys.ahk:9883","message":"FindMonthGroup Strategy 2 attempt","data":{"monthCount":14},"sessionId":"debug-session","runId":"run1","hypothesisId":"C"}`n',
        DEBUG_LOG_PATH
    } catch {
    }
    ; #endregion
    months := ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
        "November", "December",
        "Janeiro", "Fevereiro", "MarÃ§o", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro",
        "Novembro",
        "Dezembro"]
    foundMonths := []
    for , m in months {
        try {
            el := uia.FindElement({ Name: m, Type: "Text", mm: 1, cs: false })
            if el {
                ; #region agent log
                try {
                    FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' .
                    A_TickCount .
                    ',"location":"Shift keys.ahk:9897","message":"FindMonthGroup found month text","data":{"month":"' .
                    m . '"},"sessionId":"debug-session","runId":"run1","hypothesisId":"C"}`n', DEBUG_LOG_PATH
                } catch {
                }
                ; #endregion
                grp := el.WalkTree("p", { Type: "Group" })
                if grp {
                    ; #region agent log
                    try {
                        FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' .
                        A_TickCount .
                        ',"location":"Shift keys.ahk:9904","message":"FindMonthGroup Strategy 2 success","data":{"month":"' .
                        m . '","found":true},"sessionId":"debug-session","runId":"run1","hypothesisId":"C"}`n',
                        DEBUG_LOG_PATH
                    } catch {
                    }
                    ; #endregion
                    return grp
                }
            }
        }
        catch {
        }
    }
    ; #region agent log
    try {
        FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' . A_TickCount .
        ',"location":"Shift keys.ahk:9912","message":"FindMonthGroup Strategy 2 failed","data":{"foundMonths":' .
        foundMonths.Length . '},"sessionId":"debug-session","runId":"run1","hypothesisId":"C"}`n',
        DEBUG_LOG_PATH
    } catch {
    }
    ; #endregion
    ; #region agent log
    try {
        ; Try to find what Groups exist
        try {
            allGroups := uia.FindAll({ Type: "Group" })
            groupCount := allGroups.Length
            FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' . A_TickCount .
            ',"location":"Shift keys.ahk:9918","message":"FindMonthGroup diagnostic","data":{"totalGroups":' .
            groupCount . '},"sessionId":"debug-session","runId":"run1","hypothesisId":"D"}`n', DEBUG_LOG_PATH
            ; Inspect first few Groups for className and name
            sampleCount := groupCount < 5 ? groupCount : 5
            loop sampleCount {
                try {
                    grp := allGroups[A_Index]
                    className := ""
                    name := ""
                    try className := grp.GetPropertyValue(UIA.Property.ClassName)
                    try name := grp.GetPropertyValue(UIA.Property.Name)
                    FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' .
                    A_TickCount .
                    ',"location":"Shift keys.ahk:9920","message":"FindMonthGroup Group sample","data":{"index":' .
                    A_Index . ',"className":"' . className . '","name":"' . name .
                    '"},"sessionId":"debug-session","runId":"run1","hypothesisId":"D"}`n', DEBUG_LOG_PATH
                } catch {
                }
            }
            ; Try to find pagination elements
            try {
                paginationBtns := uia.FindAll({ Name: "Go to next page", Type: 50000 })
                if !paginationBtns.Length {
                    paginationBtns := uia.FindAll({ Name: "Go to previous page", Type: 50000 })
                }
                FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' .
                A_TickCount .
                ',"location":"Shift keys.ahk:9930","message":"FindMonthGroup pagination check","data":{"foundPagination":' .
                (paginationBtns.Length > 0 ? 1 : 0) . ',"count":' . paginationBtns.Length .
                '},"sessionId":"debug-session","runId":"run1","hypothesisId":"E"}`n', DEBUG_LOG_PATH
            } catch {
            }
            ; Try to find any text elements that might contain dates/months
            try {
                allTexts := uia.FindAll({ Type: "Text" })
                textCount := allTexts.Length
                FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' .
                A_TickCount .
                ',"location":"Shift keys.ahk:9935","message":"FindMonthGroup text elements","data":{"totalTexts":' .
                textCount . '},"sessionId":"debug-session","runId":"run1","hypothesisId":"E"}`n',
                DEBUG_LOG_PATH
                ; Sample first 10 text elements
                sampleTextCount := textCount < 10 ? textCount : 10
                loop sampleTextCount {
                    try {
                        txt := allTexts[A_Index]
                        txtName := ""
                        try txtName := txt.GetPropertyValue(UIA.Property.Name)
                        if txtName && StrLen(txtName) > 0 {
                            FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) .
                            '","timestamp":' .
                            A_TickCount .
                            ',"location":"Shift keys.ahk:9940","message":"FindMonthGroup text sample","data":{"index":' .
                            A_Index . ',"text":"' . txtName .
                            '"},"sessionId":"debug-session","runId":"run1","hypothesisId":"E"}`n',
                            DEBUG_LOG_PATH
                        }
                    } catch {
                    }
                }
            } catch {
            }
        } catch Error as e2 {
            FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' . A_TickCount .
            ',"location":"Shift keys.ahk:9918","message":"FindMonthGroup diagnostic failed","data":{"error":"' .
            e2.Message .
            '"},"sessionId":"debug-session","runId":"run1","hypothesisId":"D"}`n', DEBUG_LOG_PATH
        }
    } catch {
    }
    ; #endregion
    return 0
}

;-------------------------------------------------------------------
; YouTube Shortcuts
;-------------------------------------------------------------------
