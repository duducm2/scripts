; =============================================================================
; Utils module: chrome_dictation_navigate.ahk
; Send dictation? [C] — new Chrome window, paste into address bar, Enter,
; then open the first Google result (same as Shift+U)
; =============================================================================

ChromeDictation_SnapshotHwnds() {
    ids := Map()
    for hwnd in WinGetList("ahk_exe chrome.exe") {
        if (hwnd is Integer) && hwnd > 0
            ids[hwnd] := true
    }
    return ids
}

ChromeDictation_FindNewMainHwnd(oldIds) {
    best := 0
    bestArea := 0
    for hwnd in WinGetList("ahk_exe chrome.exe") {
        if !(hwnd is Integer) || hwnd <= 0
            continue
        if oldIds.Has(hwnd)
            continue
        if !DllCall("IsWindowVisible", "Ptr", hwnd)
            continue
        try {
            WinGetPos(, , &w, &h, "ahk_id " hwnd)
            if (w < 200 || h < 200)
                continue
            area := w * h
        } catch {
            continue
        }
        if (area > bestArea) {
            bestArea := area
            best := hwnd
        }
    }
    return best
}

ChromeDictation_WaitForNewWindow(oldIds, timeoutMs := 20000) {
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        hwnd := ChromeDictation_FindNewMainHwnd(oldIds)
        if (hwnd > 0)
            return hwnd
        Sleep 150
    }
    return 0
}

; End-to-end: new Chrome window → Ctrl+L → paste dictation → Enter → first result.
ChromeDictation_NavigateFromClipboard(messageText) {
    query := Trim(messageText)
    if (query = "") {
        ShowCenteredOverlay_Utils("❌ No dictated text to open in Chrome.", 2200, BANNER_ACCENT_ERROR)
        return false
    }

    StandardLoadingBar_Show("⏳ Opening Chrome — please wait...", BANNER_ACCENT_INTERMEDIATE, {
        fontSize: 17,
        trackActiveMonitor: true
    })
    barOwned := true
    clipSaved := ClipboardAll()
    try {
        oldIds := ChromeDictation_SnapshotHwnds()
        StandardLoadingBar_Update("⏳ Opening new Chrome window...")
        try {
            Run("chrome.exe --new-window")
        } catch {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ Could not launch Chrome.", 2200, BANNER_ACCENT_ERROR)
            barOwned := false
            return false
        }

        StandardLoadingBar_Update("⏳ Waiting for Chrome window...")
        hwnd := ChromeDictation_WaitForNewWindow(oldIds, 20000)
        if (hwnd <= 0) {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ Chrome window did not appear in time.", 2500, BANNER_ACCENT_ERROR)
            barOwned := false
            return false
        }

        StandardLoadingBar_Update("⏳ Activating Chrome...")
        WinActivate("ahk_id " hwnd)
        if (!WinActive("ahk_id " hwnd))
            WinWaitActive("ahk_id " hwnd, , 3)
        if (!WinActive("ahk_id " hwnd)) {
            WinActivate("ahk_id " hwnd)
            if (!WinWaitActive("ahk_id " hwnd, , 2)) {
                StandardLoadingBar_Hide(0)
                ShowCenteredOverlay_Utils("❌ Could not activate Chrome.", 2200, BANNER_ACCENT_ERROR)
                barOwned := false
                return false
            }
        }

        A_Clipboard := ""
        A_Clipboard := query
        if (!ClipWait(2)) {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ CLIPBOARD ERROR - COULD NOT RESTORE DICTATION", 3000, BANNER_ACCENT_ERROR)
            barOwned := false
            return false
        }

        Send "{LWin Up}{RWin Up}{LAlt Up}{RAlt Up}{LShift Up}{RShift Up}"
        Sleep 80
        Send "^l"
        Sleep 120
        Send "^v"
        Sleep 350
        Send "{Enter}"

        StandardLoadingBar_Update("⏳ Waiting for Google results...")
        if (!GoogleSearch_WaitAndClickFirstResult(hwnd, 15000)) {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ First Google result not found.", 2200, BANNER_ACCENT_ERROR)
            barOwned := false
            return false
        }

        StandardLoadingBar_Hide(0)
        barOwned := false
        return true
    } finally {
        if (barOwned) {
            try StandardLoadingBar_Hide(0)
        }
        A_Clipboard := clipSaved
        if (ClipWait(1)) {
        }
    }
}
