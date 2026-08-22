; =============================================================================
; Utils module: whatsapp_jump_to_chat.ahk
; Shared Jump-to-Chat logic for WhatsApp (Web / Desktop)
; =============================================================================

; Chrome "WhatsApp Web" PWA app-id from the Start Menu .lnk (--app-id=...).
; Needed because an open chat often sets the window title to the contact name only
; (no "WhatsApp" substring), so WinExist("WhatsApp") misses the running app.
global WHATSAPP_CHROME_APP_ID := "hnpfjngllnobngcgfapefoaidbinmjnm"

; Cache-first find (rollback: WHATSAPP_JUMP_USE_HWND_CACHE := false).
global WHATSAPP_JUMP_USE_HWND_CACHE := true
global WHATSAPP_JUMP_CACHED_HWND := 0

; Fast path: after activate/Run+HWND, skip WaitUntilUiReady entirely (warm AND cold).
; Search spam in JumpToChat is the readiness gate. Rollback: false → legacy UI-ready waits.
global WHATSAPP_JUMP_WARM_FASTPATH := true
; Set by ActivateOrOpen so JumpToChat can use a longer search timeout after cold Run.
global WHATSAPP_JUMP_LAST_WAS_COLD := false

WhatsAppJump_ShowLoading(state) {
    StandardLoadingBar_Show(state, BANNER_ACCENT_INTERMEDIATE, {
        fontSize: 17,
        trackActiveMonitor: true
    })
}

WhatsAppJump_UpdateLoading(state) {
    global g_StandardLoadingBarGui
    if IsObject(g_StandardLoadingBarGui)
        StandardLoadingBar_Update(state)
    else
        WhatsAppJump_ShowLoading(state)
}

WhatsAppJump_HideLoading() {
    StandardLoadingBar_Hide(0)
}

; Chrome PWA AUMID is typically Chrome._crx_<app-id>.
WhatsAppJump_GetWindowAppUserModelId(hwnd) {
    if !(hwnd is Integer) || hwnd <= 0
        return ""
    iid := Buffer(16, 0)
    if DllCall("ole32\CLSIDFromString", "wstr", "{886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99}", "ptr", iid)
        return ""
    pstore := 0
    if DllCall("shell32\SHGetPropertyStoreForWindow", "ptr", hwnd, "ptr", iid, "ptr*", &pstore) || !pstore
        return ""
    pk := Buffer(20, 0)
    ; PKEY_AppUserModel_ID = {9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3}, pid 5
    if DllCall("ole32\CLSIDFromString", "wstr", "{9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3}", "ptr", pk) {
        ObjRelease(pstore)
        return ""
    }
    NumPut("uint", 5, pk, 16)
    prop := Buffer(A_PtrSize = 8 ? 24 : 16, 0)
    hr := ComCall(5, pstore, "ptr", pk, "ptr", prop) ; IPropertyStore::GetValue
    ObjRelease(pstore)
    if hr != 0
        return ""
    appId := ""
    try {
        ; VT_LPWSTR = 31; pointer payload starts at offset 8
        if (NumGet(prop, 0, "ushort") = 31) {
            pstr := NumGet(prop, 8, "ptr")
            if pstr
                appId := StrGet(pstr, "UTF-16")
        }
    } catch {
        appId := ""
    }
    DllCall("ole32\PropVariantClear", "ptr", prop)
    return appId
}

WhatsAppJump_IsWhatsAppChromeAppHwnd(hwnd) {
    global WHATSAPP_CHROME_APP_ID
    appId := WhatsAppJump_GetWindowAppUserModelId(hwnd)
    return (appId != "" && InStr(appId, WHATSAPP_CHROME_APP_ID))
}

WhatsAppJump_InvalidateHwndCache() {
    global WHATSAPP_JUMP_CACHED_HWND
    WHATSAPP_JUMP_CACHED_HWND := 0
}

; True if hwnd still looks like WhatsApp (title, native exe, or Chrome PWA AUMID).
WhatsAppJump_IsValidWhatsAppHwnd(hwnd) {
    if !(hwnd is Integer) || hwnd <= 0
        return false
    if !WinExist("ahk_id " hwnd)
        return false
    try {
        if (WinGetProcessName("ahk_id " hwnd) = "WhatsApp.exe")
            return true
    } catch {
    }
    prevTitleMode := A_TitleMatchMode
    try {
        SetTitleMatchMode(2)
        title := ""
        try title := WinGetTitle("ahk_id " hwnd)
        if (title != "" && InStr(title, "WhatsApp"))
            return true
    } finally {
        SetTitleMatchMode(prevTitleMode)
    }
    try {
        if WhatsAppJump_IsWhatsAppChromeAppHwnd(hwnd)
            return true
    } catch {
    }
    return false
}

WhatsAppJump_SetCachedHwnd(hwnd) {
    global WHATSAPP_JUMP_USE_HWND_CACHE, WHATSAPP_JUMP_CACHED_HWND
    if !WHATSAPP_JUMP_USE_HWND_CACHE
        return
    if (hwnd is Integer) && hwnd > 0
        WHATSAPP_JUMP_CACHED_HWND := hwnd
    else
        WHATSAPP_JUMP_CACHED_HWND := 0
}

; Returns HWND of an existing WhatsApp window, or 0.
; Matches: title "WhatsApp", native WhatsApp.exe, or Chrome PWA by AppUserModelID.
WhatsAppJump_FindHwnd() {
    global WHATSAPP_JUMP_USE_HWND_CACHE, WHATSAPP_JUMP_CACHED_HWND

    if (WHATSAPP_JUMP_USE_HWND_CACHE && WHATSAPP_JUMP_CACHED_HWND > 0) {
        if WhatsAppJump_IsValidWhatsAppHwnd(WHATSAPP_JUMP_CACHED_HWND)
            return WHATSAPP_JUMP_CACHED_HWND
        WhatsAppJump_InvalidateHwndCache()
    }

    hwnd := 0
    prevTitleMode := A_TitleMatchMode
    try {
        SetTitleMatchMode(2)
        hwnd := WinExist("WhatsApp")
    } finally {
        SetTitleMatchMode(prevTitleMode)
    }

    if !hwnd
        hwnd := WinExist("ahk_exe WhatsApp.exe")

    if !hwnd {
        for cand in WinGetList("ahk_exe chrome.exe") {
            try {
                if WhatsAppJump_IsWhatsAppChromeAppHwnd(cand) {
                    hwnd := cand
                    break
                }
            } catch {
            }
        }
    }
    ; Broader class scan — some PWA hosts are not reported as chrome.exe promptly.
    if !hwnd {
        for cand in WinGetList("ahk_class Chrome_WidgetWin_1") {
            try {
                if WhatsAppJump_IsWhatsAppChromeAppHwnd(cand) {
                    hwnd := cand
                    break
                }
            } catch {
            }
        }
    }


    if (hwnd)
        WhatsAppJump_SetCachedHwnd(hwnd)
    return hwnd
}

; Bounded wait until a WhatsApp window exists. Returns HWND or 0.
WhatsAppJump_WaitHwnd(timeoutSec := 30) {
    deadline := A_TickCount + (timeoutSec * 1000)
    prevTitleMode := A_TitleMatchMode
    try {
        SetTitleMatchMode(2)
        ; Fast path: title often becomes "WhatsApp" briefly on cold start.
        if WinWait("WhatsApp", , Min(timeoutSec, 3)) {
            if (hwnd := WhatsAppJump_FindHwnd())
                return hwnd
        }
        while (A_TickCount < deadline) {
            if (hwnd := WhatsAppJump_FindHwnd())
                return hwnd
            Sleep 150
        }
        return 0
    } finally {
        SetTitleMatchMode(prevTitleMode)
    }
}

WhatsAppJump_ActivateHwnd(hwnd, timeoutSec := 5) {
    if !(hwnd is Integer) || hwnd <= 0
        return false
    WinActivate("ahk_id " hwnd)
    if WinWaitActive("ahk_id " hwnd, , timeoutSec)
        return true
    WinActivate("ahk_id " hwnd)
    return !!WinWaitActive("ahk_id " hwnd, , 2)
}

; Shell gate: Unread/All filter tabs or Archived button (same anchors as hotif_whatsapp).
WhatsAppJump_IsUiReady(hwnd) {
    if !(hwnd is Integer) || hwnd <= 0
        return false
    try {
        uia := UIA_Browser("ahk_id " hwnd)
        if (!uia)
            return false
        try {
            if (uia.FindElement({ Name: "Unread", AutomationId: "unread-filter", Type: "TabItem" }))
                return true
        } catch {
        }
        try {
            if (uia.FindElement({ Name: "All", AutomationId: "all-filter", Type: "TabItem" }))
                return true
        } catch {
        }
        try {
            if (uia.FindElement({ Name: "Archived ", Type: "Button" }))
                return true
        } catch {
        }
    } catch {
    }
    return false
}

; Poll until shell chrome is stable. Cold start uses longer timeout/streak than warm.
; settleMs: fixed sleep after streak met (0 = none; legacy warm/cold used 600).
WhatsAppJump_WaitUntilUiReady(initialHwnd := 0, timeoutMs := 25000, neededStreak := 2, openStart := 0, settleMs := 600) {
    if (!openStart)
        openStart := A_TickCount
    deadline := A_TickCount + timeoutMs
    readyStreak := 0
    lastUpdate := 0
    while (A_TickCount < deadline) {
        hwnd := WhatsAppJump_FindHwnd()
        if (hwnd <= 0 && initialHwnd > 0 && WinExist("ahk_id " initialHwnd))
            hwnd := initialHwnd
        elapsed := Round((A_TickCount - openStart) / 1000)
        if ((A_TickCount - lastUpdate) >= 800) {
            WhatsAppJump_UpdateLoading("⏳ Waiting until WhatsApp is fully ready... (" elapsed "s)")
            lastUpdate := A_TickCount
        }
        if (hwnd > 0 && WhatsAppJump_IsUiReady(hwnd)) {
            readyStreak += 1
            if (readyStreak >= neededStreak) {
                WinActivate("ahk_id " hwnd)
                WinWaitActive("ahk_id " hwnd, , 2)
                WhatsAppJump_SetCachedHwnd(hwnd)
                if (settleMs > 0) {
                    WhatsAppJump_UpdateLoading("⏳ WhatsApp shell ready — finishing load...")
                    Sleep settleMs
                }
                return hwnd
            }
        } else {
            readyStreak := 0
        }
        Sleep 200
    }
    return 0
}

; True only when the WhatsApp *search* field is focused — not the chat message composer.
; Composer is also Type=Edit; requiring a search-like Name avoids false positives.
WhatsAppJump_IsSearchEditFocused() {
    try {
        fe := UIA.GetFocusedElement()
        if (!fe)
            return false
        nm := ""
        try nm := fe.Name
        if (nm = "") {
            try nm := fe.LegacyIAccessible.Name
        }
        if (nm = "")
            return false
        ; Exclude message composer / caption fields.
        if (RegExMatch(nm, "i)(type a message|digite uma mensagem|mensagem|message|caption|legenda)"))
            return false
        ; Search / start-chat field (EN + PT and close variants).
        if (RegExMatch(nm, "i)(search|buscar|pesquisar|encontrar|name or number|nome ou)"))
            return true
    } catch {
    }
    return false
}

; Poll until search Edit is focused (or deadline). Returns true if focused.
WhatsAppJump_WaitSearchEditFocused(timeoutMs := 450, pollMs := 40) {
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        if (WhatsAppJump_IsSearchEditFocused())
            return true
        Sleep pollMs
    }
    return WhatsAppJump_IsSearchEditFocused()
}

; After activate: keep firing search until the search field appears, then stop.
; Hotkey is Alt+K — same as Shift+S remap in hotif_whatsapp (+s → !k).
; Do not Send("+s") here: that types into WhatsApp and does not fire the other script's remap.
; One !k, then poll; only send again if search still not focused (avoids toggle open/close).
WhatsAppJump_OpenSearchUntilFocused(hwnd, timeoutMs := 10000, waitAfterSendMs := 280, pollMs := 40) {
    deadline := A_TickCount + timeoutMs
    lastUpdate := 0

    ; #!+r leaves Win/Alt/Shift down until released — wait so !k is not Alt+Shift+Win+K.
    KeyWait("LWin", "T0.4")
    KeyWait("RWin", "T0.4")
    KeyWait("LAlt", "T0.4")
    KeyWait("RAlt", "T0.4")
    KeyWait("LShift", "T0.4")
    KeyWait("RShift", "T0.4")
    Send "{LWin Up}{RWin Up}{LAlt Up}{RAlt Up}{LShift Up}{RShift Up}{LCtrl Up}{RCtrl Up}"

    while (A_TickCount < deadline) {
        if (hwnd > 0 && !WinActive("ahk_id " hwnd)) {
            WinActivate("ahk_id " hwnd)
            WinWaitActive("ahk_id " hwnd, , 1)
        }
        if (WhatsAppJump_IsSearchEditFocused())
            return true
        if ((A_TickCount - lastUpdate) >= 400) {
            WhatsAppJump_UpdateLoading("⏳ Opening search...")
            lastUpdate := A_TickCount
        }
        ; Same action as Shift+S in hotif_whatsapp
        Send "!k"
        if (WhatsAppJump_WaitSearchEditFocused(waitAfterSendMs, pollMs))
            return true
    }
    return WhatsAppJump_IsSearchEditFocused()
}

; True only when the chat *message composer* is focused (whatsapp.md: "Type a message to …").
; Rejects search Edit names so a stuck search field never counts as success.
WhatsAppJump_IsComposerEditFocused() {
    try {
        fe := UIA.GetFocusedElement()
        if (!fe)
            return false
        nm := ""
        try nm := fe.Name
        if (nm = "") {
            try nm := fe.LegacyIAccessible.Name
        }
        if (nm = "")
            return false
        if (RegExMatch(nm,
            "i)(search|buscar|pesquisar|encontrar|name or number|nome ou|start a new chat|nova conversa)"))
            return false
        if (RegExMatch(nm, "i)(type a message|digite uma mensagem|digite a mensagem)"))
            return true
    } catch {
    }
    return false
}

; Poll until message composer Edit is focused (or deadline). Returns true if focused.
WhatsAppJump_WaitComposerEditFocused(timeoutMs := 1200, pollMs := 40) {
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        if (WhatsAppJump_IsComposerEditFocused())
            return true
        Sleep pollMs
    }
    return WhatsAppJump_IsComposerEditFocused()
}

; Returns true after WhatsApp is focused (or launched). With WHATSAPP_JUMP_WARM_FASTPATH,
; does NOT wait for shell UIA readiness — JumpToChat spams search until the field appears.
WhatsAppJump_ActivateOrOpen() {
    global IS_WORK_ENVIRONMENT, WHATSAPP_JUMP_WARM_FASTPATH, WHATSAPP_JUMP_LAST_WAS_COLD
    WHATSAPP_JUMP_LAST_WAS_COLD := false
    prevTitleMode := A_TitleMatchMode
    try {
        SetTitleMatchMode(2)
        didColdStart := false
        openStart := A_TickCount
        hwnd := WhatsAppJump_FindHwnd()

        ; Missed HWND while PWA is open (contact-title window): Run .lnk focuses the existing app.
        if (!hwnd) {
            WhatsAppJump_UpdateLoading("⏳ Focusing WhatsApp...")
            if (IS_WORK_ENVIRONMENT) {
                Run "C:\Users\fie7ca\Documents\Shortcuts\WhatsApp.lnk"
            } else {
                Run "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\apps do Chrome\WhatsApp Web.lnk"
            }
            hwnd := WhatsAppJump_WaitHwnd(5)
            if (hwnd) {
                ; Existing instance came to front — treat as warm focus, not cold hydrate.
                WhatsAppJump_UpdateLoading("⏳ Activating WhatsApp...")
                if !WhatsAppJump_ActivateHwnd(hwnd, 3) {
                    WhatsAppJump_HideLoading()
                    ShowCenteredOverlay_Utils("❌ Could not activate WhatsApp.", 2000, BANNER_ACCENT_ERROR)
                    return false
                }
                WhatsAppJump_SetCachedHwnd(hwnd)
                if (WHATSAPP_JUMP_WARM_FASTPATH)
                    return true
            } else {
                didColdStart := true
                WHATSAPP_JUMP_LAST_WAS_COLD := true
                WhatsAppJump_UpdateLoading("⏳ Opening WhatsApp...")
                hwnd := WhatsAppJump_WaitHwnd(30)
                if !hwnd {
                    WhatsAppJump_HideLoading()
                    ShowCenteredOverlay_Utils("❌ WhatsApp did not start in time.", 2000, BANNER_ACCENT_ERROR)
                    return false
                }
                if !WhatsAppJump_ActivateHwnd(hwnd, 5) {
                    WhatsAppJump_HideLoading()
                    ShowCenteredOverlay_Utils("❌ Could not activate WhatsApp.", 2000, BANNER_ACCENT_ERROR)
                    return false
                }
                WhatsAppJump_SetCachedHwnd(hwnd)
                ; Fast path: do not WaitUntilUiReady — search spam is the gate (even on cold start).
                if (WHATSAPP_JUMP_WARM_FASTPATH)
                    return true
            }
        } else {
            WhatsAppJump_UpdateLoading("⏳ Activating WhatsApp...")
            if !WhatsAppJump_ActivateHwnd(hwnd, 3) {
                WhatsAppJump_HideLoading()
                ShowCenteredOverlay_Utils("❌ Could not activate WhatsApp.", 2000, BANNER_ACCENT_ERROR)
                return false
            }
            WhatsAppJump_SetCachedHwnd(hwnd)
            if (WHATSAPP_JUMP_WARM_FASTPATH)
                return true
        }

        ; Legacy path only (WHATSAPP_JUMP_WARM_FASTPATH := false).
        if (didColdStart) {
            timeoutMs := 45000
            neededStreak := 8
            settleMs := 600
        } else {
            timeoutMs := 8000
            neededStreak := 2
            settleMs := 600
        }
        hwnd := WhatsAppJump_WaitUntilUiReady(hwnd, timeoutMs, neededStreak, openStart, settleMs)
        if !hwnd {
            WhatsAppJump_HideLoading()
            ShowCenteredOverlay_Utils("❌ WhatsApp UI did not become ready in time.", 2500, BANNER_ACCENT_ERROR)
            return false
        }
        return true
    } finally {
        SetTitleMatchMode(prevTitleMode)
    }
}

; Opens search, pastes contact, Enter to select chat only. Does not paste the message.
; keepBarVisible: when true (D2C [Z]), leave the Loading Indication up for the caller to continue/hide.
; After Enter, requires composer focus ("Type a message to …"); retries search→paste→Enter up to 3 times.
WhatsAppJumpToChat(contact, keepBarVisible := false) {
    global WHATSAPP_JUMP_LAST_WAS_COLD
    WhatsAppJump_ShowLoading("⏳ Activating WhatsApp...")
    if (!WhatsAppJump_ActivateOrOpen())
        return false

    oldWinDelay := A_WinDelay
    oldKeyDelay := A_KeyDelay
    oldControlDelay := A_ControlDelay

    try {
        SetWinDelay 0
        SetKeyDelay 0, 0
        SetControlDelay 0

        Send "{LWin Up}{RWin Up}{LAlt Up}{RAlt Up}{LShift Up}{RShift Up}"

        hwnd := WhatsAppJump_FindHwnd()
        if (hwnd > 0 && !WinActive("ahk_id " hwnd)) {
            WhatsAppJump_UpdateLoading("⏳ Focusing WhatsApp...")
            WhatsAppJump_ActivateHwnd(hwnd, 2)
        }

        ; Cold hydrate needs a longer search spam window; warm can be shorter.
        searchTimeoutMs := WHATSAPP_JUMP_LAST_WAS_COLD ? 30000 : 12000

        maxAttempts := 3
        loop maxAttempts {
            attempt := A_Index
            if (attempt > 1) {
                WhatsAppJump_UpdateLoading("⏳ Retrying jump (" attempt "/" maxAttempts ")...")
                Send "{Escape}"
                Sleep 150
                hwnd := WhatsAppJump_FindHwnd()
                if (hwnd > 0 && !WinActive("ahk_id " hwnd)) {
                    WhatsAppJump_ActivateHwnd(hwnd, 2)
                }
            }

            ; Hammer search (!k = Shift+S remap) until the real search field is focused.
            WhatsAppJump_UpdateLoading("⏳ Opening search...")
            if (!WhatsAppJump_OpenSearchUntilFocused(hwnd, searchTimeoutMs)) {
                if (attempt = maxAttempts) {
                    WhatsAppJump_HideLoading()
                    ShowCenteredOverlay_Utils("❌ Could not focus WhatsApp search.", 2500, BANNER_ACCENT_ERROR)
                    return false
                }
                continue
            }

            WhatsAppJump_UpdateLoading("⏳ Searching contact...")
            clipOk := false
            loop 5 {
                A_Clipboard := ""
                A_Clipboard := contact
                if ClipWait(2) && (A_Clipboard = contact) {
                    clipOk := true
                    break
                }
                Sleep 100
            }
            if (!clipOk) {
                WhatsAppJump_HideLoading()
                ShowCenteredOverlay_Utils("❌ CLIPBOARD ERROR - TRY AGAIN", 3000, BANNER_ACCENT_ERROR)
                return false
            }

            Send "^v"
            Sleep 300
            WhatsAppJump_UpdateLoading("⏳ Opening chat...")
            Send "{Enter}"

            if (WhatsAppJump_WaitComposerEditFocused(1200, 40)) {
                if (!keepBarVisible)
                    WhatsAppJump_HideLoading()
                return true
            }
        }

        WhatsAppJump_HideLoading()
        ShowCenteredOverlay_Utils("❌ Could not open chat composer.", 2500, BANNER_ACCENT_ERROR)
        return false
    } finally {
        SetWinDelay oldWinDelay
        SetKeyDelay oldKeyDelay, 0
        SetControlDelay oldControlDelay
    }
}