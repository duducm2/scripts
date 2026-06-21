; =============================================================================
; Utils module: chrome_detach_03.ahk
; Chrome detach tab (part 3) and detach entry
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

Chrome_NormalizeFocusToPage(hwnd) {
    if !Chrome_EnsureBrowserForeground(hwnd)
        return false
    ClipAngel_ReleaseChordModifiersForSend()
    Send "{Escape}"
    Send "{Escape}"
    Send "{Escape}"
    Chrome_FocusPageContent(hwnd)
    return WM_EnsureForegroundForSend(hwnd, 1000)
}

; Detach hot path: caller already foregrounded — one Escape, skip page focus when UIA usable.
Chrome_NormalizeFocusToPageLight(hwnd, session := unset) {
    ClipAngel_ReleaseChordModifiersForSend()
    Send "{Escape}"
    if !(IsSet(session) && IsObject(session) && Chrome_SessionUiaUsable(session))
        Chrome_FocusPageContent(hwnd)
    if WinActive("ahk_id " hwnd)
        return true
    return WM_EnsureForegroundForSend(hwnd, 200)
}

Chrome_OpenActiveTabContextMenu(session) {
    global CHROME_DETACH_USE_LIGHT_NORMALIZE
    hwnd := session.hwnd
    if !WinActive("ahk_id " hwnd) && !Chrome_EnsureBrowserForeground(hwnd)
        return false
    if (session.menuPopupHwnd && Chrome_ContextMenuLooksLikeTabMenu(session)) {
        session.menuConfirmed := true
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_OpenActiveTabContextMenu", "reuse open tab menu", "E", "popup=" .
            session.menuPopupHwnd)
        ; #endregion
        return true
    }
    session.baselinePopups := Chrome_DetachListMenuPopups(hwnd)
    Chrome_DetachClearMenuPopup(session)
    if CHROME_DETACH_USE_LIGHT_NORMALIZE {
        if !Chrome_NormalizeFocusToPageLight(hwnd, session)
            return false
    } else if !Chrome_NormalizeFocusToPage(hwnd) {
        return false
    }
    if Chrome_TryF6KeyboardTabMenu(session) {
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_OpenActiveTabContextMenu", "opened via f6", "H", { result: 1 })
        ; #endregion
        return true
    }
    if Chrome_OpenActiveTabContextMenuViaTabFocus(session) {
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_OpenActiveTabContextMenu", "opened via tab focus", "H", { result: 1 })
        ; #endregion
        return true
    }
    if !CHROME_DETACH_DEEP_FALLBACK
        return false
    f6Ok := Chrome_FocusTabStripAndOpenContextMenu(session)
    ; #region agent log
    Chrome_DetachDebugLog("Chrome_OpenActiveTabContextMenu", "f6+AppsKey fallback result", "C", { path: "F6AppsKey",
        result: f6Ok ? 1 : 0 })
    ; #endregion
    return f6Ok
}

Chrome_DetachCountTabs(session) {
    if !IsObject(session.uia)
        return -1
    try {
        return session.uia.GetAllTabs().Length
    } catch {
        return -1
    }
}

; UIA Invoke first; PT keyboard m -> Enter -> n (Nova janela submenu). Never bare 'n' at top level.
Chrome_ActivateDetachMenuItem(session) {
    if !session.menuPopupHwnd
        Chrome_ContextMenuCapturePopup(session, session.baselinePopups)
    ; #region agent log
    Chrome_DetachDebugLog("Chrome_ActivateDetachMenuItem", "start", "E", "popup=" . session.menuPopupHwnd .
        " sample=" . Chrome_DetachDebugSampleMenuItems(session))
    ; #endregion
    if Chrome_ContextMenuPopupIsPageMenu(session) {
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_ActivateDetachMenuItem", "abort page menu", "E", "popup=" .
            session.menuPopupHwnd . ";sample=" . Chrome_DetachDebugSampleMenuItems(session))
        ; #endregion
        return false
    }
    focused := Chrome_ContextMenuFocusPopup(session)
    ClipAngel_ReleaseChordModifiersForSend()

    flat := Chrome_ContextMenuFindFirst(session, CHROME_DETACH_MENU_EN_NAMES, false)
    if flat && Chrome_ContextMenuActivateItem(flat) {
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_ActivateDetachMenuItem", "flat item invoked", "E", "path=flat")
        ; #endregion
        return true
    }

    parent := Chrome_ContextMenuFindFirst(session, CHROME_DETACH_MENU_PARENT_NAMES, true)
    if parent && Chrome_ContextMenuActivateItem(parent) {
        child := Chrome_ContextMenuWaitForSession(session, CHROME_DETACH_MENU_CHILD_NAMES, CHROME_DETACH_MENU_CHILD_MS,
            false)
        if child && Chrome_ContextMenuActivateItem(child) {
            ; #region agent log
            Chrome_DetachDebugLog("Chrome_ActivateDetachMenuItem", "parent+child invoked", "E", "path=submenu")
            ; #endregion
            return true
        }
    }

    if Chrome_ContextMenuPopupIsPageMenu(session) {
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_ActivateDetachMenuItem", "abort page menu", "E", "popup=" .
            session.menuPopupHwnd . ";sample=" . Chrome_DetachDebugSampleMenuItems(session))
        ; #endregion
        return false
    }
    ; #region agent log
    Chrome_DetachDebugLog("Chrome_ActivateDetachMenuItem", "UIA miss, keyboard PT", "E", "sample=" .
        Chrome_DetachDebugSampleMenuItems(session) . ";popupFocus=" . (focused ? 1 : 0))
    ; #endregion
    Chrome_ContextMenuFocusPopup(session)
    ClipAngel_ReleaseChordModifiersForSend()
    Chrome_ContextMenuSendKeys(session, "m")
    child := Chrome_ContextMenuWaitForSession(session, CHROME_DETACH_MENU_CHILD_NAMES, CHROME_DETACH_MENU_CHILD_MS,
        false)
    if child && Chrome_ContextMenuActivateItem(child) {
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_ActivateDetachMenuItem", "keyboard child invoked", "E", "path=keyboardChild")
        ; #endregion
        return true
    }
    ClipAngel_ReleaseChordModifiersForSend()
    Chrome_ContextMenuSendKeys(session, "{Enter}")
    child := Chrome_ContextMenuWaitForSession(session, CHROME_DETACH_MENU_CHILD_NAMES, CHROME_DETACH_MENU_CHILD_MS,
        false)
    if child && Chrome_ContextMenuActivateItem(child) {
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_ActivateDetachMenuItem", "keyboard child invoked", "E", "path=keyboardEnterChild"
        )
        ; #endregion
        return true
    }
    ; Submenu open: 'n' = Nova janela (safe here; top-level 'n' = Nova guia)
    Chrome_ContextMenuFocusPopup(session)
    ClipAngel_ReleaseChordModifiersForSend()
    Chrome_ContextMenuSendKeys(session, "n")
    child := Chrome_ContextMenuWaitForSession(session, CHROME_DETACH_MENU_CHILD_NAMES, CHROME_DETACH_MENU_CHILD_MS,
        false)
    if child && Chrome_ContextMenuActivateItem(child) {
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_ActivateDetachMenuItem", "keyboard n child invoked", "E", "path=keyboardN")
        ; #endregion
        return true
    }
    ClipAngel_ReleaseChordModifiersForSend()
    Chrome_ContextMenuSendKeys(session, "{Enter}")
    ; #region agent log
    Chrome_DetachDebugLog("Chrome_ActivateDetachMenuItem", "keyboard attempted unverified", "E", "path=keyboardMN")
    ; #endregion
    return false
}

Chrome_ActivateDetachViaPtKeyboard(session) {
    return Chrome_ActivateDetachMenuItem(session)
}

Chrome_ActivateDetachViaEnKeyboard(session) {
    ; Disabled: previously re-entered Chrome_OpenActiveTabContextMenu and doubled F6 stacks.
    return false
}

Chrome_DetachActiveTabToNewWindow_Legacy() {
    Send "{F6}"
    Sleep 100
    Send "{F6}"
    Sleep 100
    Send "{AppsKey}"
    Sleep 100
    Send "m"
    Sleep 100
    Send "{Enter}"
    Sleep 100
    Send "{Enter}"
}

Chrome_IsLikelyTopLevelBrowserWindow(hwnd) {
    if !(hwnd is Integer && hwnd > 0) || !WinExist("ahk_id " hwnd)
        return false
    try {
        if (WinGetProcessName("ahk_id " hwnd) != "chrome.exe")
            return false
        cls := WinGetClass("ahk_id " hwnd)
        if !InStr(cls, "Chrome_WidgetWin")
            return false
        title := WinGetTitle("ahk_id " hwnd)
        if (Trim(title) = "")
            return false
        WinGetPos(, , &w, &h, "ahk_id " hwnd)
        return (w >= 300 && h >= 200)
    } catch {
        return false
    }
}

Chrome_GetChromeTopLevelHwndSet() {
    existingSet := Map()
    try {
        for h in WinGetList("ahk_exe chrome.exe") {
            if Chrome_IsLikelyTopLevelBrowserWindow(h)
                existingSet[h] := true
        }
    } catch {
    }
    return existingSet
}

Chrome_DetachTitleCore(title) {
    title := RegExReplace(title, "i) - Google Chrome$", "")
    title := RegExReplace(title, "i) - Chromium$", "")
    return Trim(title)
}

Chrome_DetachTitleMatches(hwnd, expectedCore) {
    if (expectedCore = "")
        return true
    if !(hwnd is Integer && hwnd > 0)
        return false
    try actual := Chrome_DetachTitleCore(WinGetTitle("ahk_id " hwnd))
    catch {
        return false
    }
    if (actual = "")
        return false
    return (actual = expectedCore) || InStr(actual, expectedCore, false) || InStr(expectedCore, actual, false)
}

Chrome_DetachCaptureBaseline(hwnd) {
    return { existingSet: Chrome_GetChromeTopLevelHwndSet(), tabTitle: Chrome_DetachGetWindowTitleForMatch(hwnd) }
}

Chrome_DetachQuickTabCount(hwnd) {
    if !(hwnd is Integer && hwnd > 0)
        return -1
    try {
        winTitle := "ahk_id " hwnd
        UIA.ActivateChromiumAccessibility(winTitle, 200)
        uia := UIA_Browser(winTitle)
        return uia.GetAllTabs().Length
    } catch {
        return -1
    }
}

Chrome_FindNewTopLevelChromeCandidate(existingSet, tabTitle := "") {
    titleCore := Chrome_DetachTitleCore(tabTitle)
    if !(existingSet is Map)
        existingSet := Map()
    try {
        for hwnd in WinGetList("ahk_exe chrome.exe") {
            if existingSet.Has(hwnd)
                continue
            if !Chrome_IsLikelyTopLevelBrowserWindow(hwnd)
                continue
            if (titleCore != "" && !Chrome_DetachTitleMatches(hwnd, titleCore))
                continue
            return hwnd
        }
    } catch {
    }
    return 0
}

Chrome_DetachPickValidatedNewHwnd(candidates, existingSet, tabTitle := "") {
    titleCore := Chrome_DetachTitleCore(tabTitle)
    if !(existingSet is Map)
        existingSet := Map()
    for hwnd in candidates {
        if (existingSet.Has(hwnd))
            continue
        if !Chrome_IsLikelyTopLevelBrowserWindow(hwnd)
            continue
        if (titleCore != "" && !Chrome_DetachTitleMatches(hwnd, titleCore))
            continue
        return hwnd
    }
    return 0
}

Chrome_DetachWinEventProc(hWinEventHook, event, hwnd, idObject, idChild, idEventThread, dwmsEventTime) {
    global g_ChromeDetachCreatedHwnds, CHROME_DETACH_EVENT_OBJECT_CREATE, CHROME_DETACH_OBJID_WINDOW
    if (event = CHROME_DETACH_EVENT_OBJECT_CREATE && idObject = CHROME_DETACH_OBJID_WINDOW && hwnd)
        g_ChromeDetachCreatedHwnds.Push(hwnd)
}

Chrome_DetachArmWindowCreateHook(&hookState) {
    global g_ChromeDetachCreatedHwnds, CHROME_DETACH_EVENT_OBJECT_CREATE
    hookState := { hHook: 0, cb: 0 }
    g_ChromeDetachCreatedHwnds := []
    hookState.cb := CallbackCreate(Chrome_DetachWinEventProc, "F Fast", 7)
    hookState.hHook := DllCall("user32\SetWinEventHook", "UInt", CHROME_DETACH_EVENT_OBJECT_CREATE,
        "UInt", CHROME_DETACH_EVENT_OBJECT_CREATE, "Ptr", 0, "Ptr", hookState.cb, "UInt", 0, "UInt", 0, "UInt", 0,
        "Ptr")
}

Chrome_DetachDisarmWindowCreateHook(hookState) {
    if IsObject(hookState) && hookState.hHook {
        DllCall("user32\UnhookWinEvent", "Ptr", hookState.hHook)
        hookState.hHook := 0
    }
}

; Bounded wait: new top-level Chrome window not in baseline, optionally matching detached tab title.
Chrome_WaitForDetachNewWindow(existingSet, timeoutMs := 0, tabTitle := "", hookState := unset) {
    global CHROME_DETACH_EXTENSION_TIMEOUT_MS, CHROME_DETACH_VERIFY_TIMEOUT_MS, CHROME_DETACH_MENU_POLL_MS
    global CHROME_DETACH_USE_WIN_EVENT_HOOK, g_ChromeDetachCreatedHwnds
    waitMs := timeoutMs > 0 ? timeoutMs : CHROME_DETACH_EXTENSION_TIMEOUT_MS
    if (waitMs <= 0)
        waitMs := CHROME_DETACH_VERIFY_TIMEOUT_MS
    deadline := A_TickCount + waitMs
    pollMs := CHROME_DETACH_MENU_POLL_MS
    if !(existingSet is Map)
        existingSet := Map()

    if CHROME_DETACH_USE_WIN_EVENT_HOOK {
        ownsHook := !(IsSet(hookState) && IsObject(hookState) && hookState.hHook)
        if ownsHook {
            Chrome_DetachArmWindowCreateHook(&hookState)
        }
        try {
            while (A_TickCount < deadline) {
                hwnd := Chrome_DetachPickValidatedNewHwnd(g_ChromeDetachCreatedHwnds, existingSet, tabTitle)
                if !hwnd
                    hwnd := Chrome_FindNewTopLevelChromeCandidate(existingSet, tabTitle)
                g_ChromeDetachCreatedHwnds := []
                if hwnd {
                    try WinActivate("ahk_id " hwnd)
                    return hwnd
                }
                Sleep pollMs
            }
        } finally {
            if ownsHook
                Chrome_DetachDisarmWindowCreateHook(hookState)
        }
        if (Chrome_DetachTitleCore(tabTitle) != "") {
            hwnd := Chrome_FindNewTopLevelChromeCandidate(existingSet, "")
            if hwnd {
                try WinActivate("ahk_id " hwnd)
                return hwnd
            }
        }
        return 0
    }

    while (A_TickCount < deadline) {
        hwnd := Chrome_FindNewTopLevelChromeCandidate(existingSet, tabTitle)
        if hwnd {
            try WinActivate("ahk_id " hwnd)
            return hwnd
        }
        Sleep pollMs
    }
    ; HWND often appears before Chrome updates the window title — accept top-level without title gate.
    if (Chrome_DetachTitleCore(tabTitle) != "") {
        hwnd := Chrome_FindNewTopLevelChromeCandidate(existingSet, "")
        if hwnd {
            try WinActivate("ahk_id " hwnd)
            return hwnd
        }
    }
    return 0
}

Chrome_WaitForNewTopLevelChromeWindow(existingSet, timeoutMs := 0, tabTitle := "") {
    return Chrome_WaitForDetachNewWindow(existingSet, timeoutMs, tabTitle)
}

Chrome_DetachRemainingMs(opStart, totalMs := 0) {
    global CHROME_DETACH_TOTAL_TIMEOUT_MS
    budget := totalMs > 0 ? totalMs : CHROME_DETACH_TOTAL_TIMEOUT_MS
    return Max(0, budget - (A_TickCount - opStart))
}

Chrome_DetachExtensionBudgetMs(opStart) {
    global CHROME_DETACH_EXTENSION_TIMEOUT_MS
    return Max(0, Min(CHROME_DETACH_EXTENSION_TIMEOUT_MS, Chrome_DetachRemainingMs(opStart)))
}

Chrome_DetachActiveTabToNewWindow_ExtensionFast(baseline := unset, timeoutMs := 0) {
    global CHROME_DETACH_USE_EXTENSION, CHROME_DETACH_EXTENSION_TIMEOUT_MS
    if !CHROME_DETACH_USE_EXTENSION
        return 0
    waitMs := timeoutMs > 0 ? Min(timeoutMs, CHROME_DETACH_EXTENSION_TIMEOUT_MS) : CHROME_DETACH_EXTENSION_TIMEOUT_MS
    if !IsSet(baseline) || !IsObject(baseline)
        baseline := Chrome_DetachCaptureBaseline(WinExist("A"))
    existingSet := baseline.existingSet
    tabTitle := baseline.tabTitle
    Chrome_Detach_PreSendSanitizeModifiers()
    ClipAngel_ReleaseChordModifiersForSend()
    SendInput "{Ctrl down}{Shift down}y{Shift up}{Ctrl up}"
    result := Chrome_WaitForDetachNewWindow(existingSet, waitMs, tabTitle)
    if !result && waitMs > 400 {
        ClipAngel_ReleaseChordModifiersForSend()
        Send "^+y"
        retryMs := Min(CHROME_DETACH_EXTENSION_TIMEOUT_MS // 2, Max(400, waitMs // 2))
        result := Chrome_WaitForDetachNewWindow(existingSet, retryMs, tabTitle)
    }
    return result
}

Chrome_DetachActiveTabToNewWindow_ExtensionFallback(baseline) {
    return Chrome_DetachActiveTabToNewWindow_ExtensionFast(baseline)
}

; One F6 pass (max 3 steps) + menu invoke — phased state gates, no blind advance.
Chrome_DetachActiveTabToNewWindow_UiaSingleShot(hwnd, &wasF11, baseline := unset, verifyTimeoutMs := 0) {
    ; #region perf
    t0 := A_TickCount
    ; #endregion
    wasF11 := false
    if !Chrome_PrepareWindowForTabDetach(hwnd, &wasF11)
        return 0
    if (WM_WindowIsF11Fullscreen(hwnd))
        return 0
    if !IsSet(baseline) || !IsObject(baseline)
        baseline := Chrome_DetachCaptureBaseline(hwnd)
    session := Chrome_DetachSessionCreate(hwnd, baseline.existingSet)
    tabCount := Chrome_DetachCountTabs(session)
    if (tabCount = 1)
        return 0
    ; #region perf
    t1 := A_TickCount
    ; #endregion
    if !Chrome_OpenActiveTabContextMenu(session)
        return 0
    ; #region perf
    t2 := A_TickCount
    ; #endregion
    if !session.menuConfirmed && !Chrome_WaitForTabContextMenu(session, CHROME_DETACH_MENU_POPUP_MS) {
        Chrome_ContextMenuDismiss()
        return 0
    }
    ; #region perf
    t3 := A_TickCount
    ; #endregion
    hookState := {}
    if CHROME_DETACH_USE_WIN_EVENT_HOOK
        Chrome_DetachArmWindowCreateHook(&hookState)
    try {
        if !Chrome_ActivateDetachMenuItemConfirmed(session, session.menuConfirmed) {
            Chrome_ContextMenuDismiss()
            return 0
        }
        ; #region perf
        t4 := A_TickCount
        ; #endregion
        verifyMs := verifyTimeoutMs > 0 ? verifyTimeoutMs : CHROME_DETACH_VERIFY_TIMEOUT_MS
        result := Chrome_WaitForDetachNewWindow(baseline.existingSet, verifyMs, baseline.tabTitle, hookState)
        ; #region perf
        t5 := A_TickCount
        if (CHROME_DETACH_PERF_LOG_ENABLED) {
            try FileAppend Format("detach: prep={} menuOpen={} menuWait={} menuAct={} newWin={} total={}`n",
                t1 - t0, t2 - t1, t3 - t2, t4 - t3, t5 - t4, t5 - t0), A_ScriptDir "\.cursor\chrome_detach_perf.log"
        }
        ; #endregion
        return result
    } finally {
        Chrome_DetachDisarmWindowCreateHook(hookState)
    }
}

; Debug-only: legacy UIA menu stack (never used in normal Shift+W path).
Chrome_DetachActiveTabToNewWindow_UiaLegacyPath(hwnd, &newHwnd, &wasF11) {
    global CHROME_DETACH_USE_UIA, CHROME_DETACH_USE_EXTENSION, CHROME_DETACH_VERIFY_TIMEOUT_MS
    newHwnd := 0
    wasF11 := false
    if !CHROME_DETACH_USE_UIA
        return false
    if !Chrome_EnsureBrowserForeground(hwnd)
        return false
    if !Chrome_PrepareWindowForTabDetach(hwnd, &wasF11)
        return false
    if (WM_WindowIsF11Fullscreen(hwnd))
        return false
    session := Chrome_DetachSessionCreate(hwnd)
    session.tabsBeforeDetach := Chrome_DetachCountTabs(session)
    if (session.tabsBeforeDetach = 1)
        return false
    if Chrome_RunDetachMenuSequence(session) {
        newHwnd := session.newDetachedHwnd
        if newHwnd
            return true
    }
    if (session.menuPopupHwnd && Chrome_ContextMenuLooksLikeTabMenu(session))
        Chrome_ActivateDetachMenuItem(session)
    newHwnd := Chrome_WaitForNewWindow(session.existingSet, CHROME_DETACH_VERIFY_TIMEOUT_MS,
        Chrome_DetachGetWindowTitleForMatch(hwnd))
    if newHwnd
        return true
    if CHROME_DETACH_USE_EXTENSION
        newHwnd := Chrome_DetachActiveTabToNewWindow_ExtensionFast(Chrome_DetachCaptureBaseline(hwnd))
    return newHwnd ? true : false
}

Chrome_PrepareWindowForTabDetach(hwnd, &wasF11) {
    wasF11 := WM_WindowIsF11Fullscreen(hwnd)
    if (wasF11) {
        try StandardLoadingBar_Update("🔄 Exiting F11 fullscreen…", BANNER_ACCENT_INTERMEDIATE)
        if !Chrome_ExitF11ForDetach(hwnd)
            return false
        try StandardLoadingBar_Update("⏳ Detaching tab to new window…", BANNER_ACCENT_INTERMEDIATE)
    }
    if WinActive("ahk_id " hwnd) && !WM_WindowIsF11Fullscreen(hwnd)
        return true
    return Chrome_WaitUntilNotF11ForDetach(hwnd)
}

; Reliable Nova guia signal: baseline count must be known (>=1) and exactly +1 tab after menu keys.
Chrome_DetachNovaGuiaLikely(tabsBefore, tabsAfter) {
    return (tabsBefore >= 1 && tabsAfter = tabsBefore + 1)
}

; Close accidental Nova guia only when detach did NOT create a new window.
Chrome_DetachCloseSpuriousNovaGuia(originalHwnd, tabsBefore, tabsAfter, existingSet) {
    if !Chrome_DetachNovaGuiaLikely(tabsBefore, tabsAfter)
        return false
    if Chrome_WaitForNewWindow(existingSet, 200)
        return false
    if !Chrome_EnsureBrowserForeground(originalHwnd)
        return false
    ClipAngel_ReleaseChordModifiersForSend()
    Send "{Ctrl down}w{Ctrl up}"
    return true
}

Chrome_FinishTabDetach(originalHwnd, newHwnd, wasF11) {
    if (!wasF11) {
        Chrome_FocusDetachedWindow(newHwnd)
        return
    }

    if (!WinExist("ahk_id " originalHwnd))
        return

    if (newHwnd)
        Chrome_EnsureNewWindowIsWindowed(newHwnd)

    Chrome_RestoreF11OnOriginal(originalHwnd)

    Chrome_FocusDetachedWindow(newHwnd)
}

Chrome_RunDetachMenuSequence(session) {
    session.tabsBeforeDetach := Chrome_DetachCountTabs(session)
    try StandardLoadingBar_Update("⏳ Detaching tab…", BANNER_ACCENT_INTERMEDIATE)
    loop CHROME_DETACH_SEQUENCE_ATTEMPTS {
        if !(session.menuPopupHwnd && Chrome_ContextMenuLooksLikeTabMenu(session)) {
            if !session.menuPopupHwnd
                Chrome_ContextMenuCapturePopup(session, session.baselinePopups)
            if !(session.menuPopupHwnd && Chrome_ContextMenuLooksLikeTabMenu(session)) {
                Chrome_DetachClearMenuPopup(session)
                session.baselinePopups := Chrome_DetachListMenuPopups(session.hwnd)
                opened := Chrome_OpenActiveTabContextMenu(session)
            } else {
                opened := true
            }
            if !opened {
                ; #region agent log
                Chrome_DetachDebugLog("Chrome_RunDetachMenuSequence", "open menu failed", "B", "attempt=" . A_Index)
                ; #endregion
                Chrome_ContextMenuDismiss()
                continue
            }
            if !session.menuPopupHwnd
                Chrome_ContextMenuCapturePopup(session, session.baselinePopups)
            if !Chrome_ContextMenuLooksLikeTabMenu(session) {
                ; #region agent log
                Chrome_DetachDebugLog("Chrome_RunDetachMenuSequence", "reject non-tab menu", "D", "attempt=" . A_Index .
                    " sample=" . Chrome_DetachDebugSampleMenuItems(session))
                ; #endregion
                Chrome_ContextMenuDismiss()
                continue
            }
        }

        Chrome_ActivateDetachMenuItem(session)
        newHwnd := Chrome_WaitForDetachNewWindow(session.existingSet, CHROME_DETACH_VERIFY_TIMEOUT_MS,
            Chrome_DetachGetWindowTitleForMatch(session.hwnd))
        if newHwnd {
            session.newDetachedHwnd := newHwnd
            ; #region agent log
            Chrome_DetachDebugLog("Chrome_RunDetachMenuSequence", "detach verified", "E", "attempt=" . A_Index .
                ";newHwnd=" . newHwnd)
            ; #endregion
            return true
        }

        Chrome_ContextMenuDismiss()
    }
    ; #region agent log
    Chrome_DetachDebugLog("Chrome_RunDetachMenuSequence", "sequence failed", "E", { result: 0 })
    ; #endregion
    return false
}

Chrome_DetachActiveTabToNewWindow() {
    global g_ChromeDetachBusy, CHROME_DETACH_USE_EXTENSION, CHROME_DETACH_USE_UIA_SINGLE_SHOT
    global CHROME_DETACH_ALLOW_UIA_DEBUG_FALLBACK, CHROME_DETACH_USE_UIA, CHROME_DETACH_PREFLIGHT_TAB_COUNT
    global CHROME_DETACH_TOTAL_TIMEOUT_MS, CHROME_DETACH_PERF_LOG_ENABLED
    if (g_ChromeDetachBusy)
        return false
    g_ChromeDetachBusy := true
    opStart := A_TickCount

    hwnd := 0
    newHwnd := 0
    wasF11 := false
    success := false
    baseline := unset
    pathTaken := ""

    StandardLoadingBar_Show("⏳ Detaching tab to new window…", BANNER_ACCENT_INTERMEDIATE, { passive: true })
    try {
        Chrome_Detach_PreSendSanitizeModifiers()

        hwnd := WinExist("A")
        if !(hwnd is Integer && hwnd > 0)
            return false
        try {
            if (WinGetProcessName("ahk_id " hwnd) != "chrome.exe")
                return false
        } catch {
            return false
        }

        if !Chrome_EnsureBrowserForeground(hwnd)
            return false

        baseline := Chrome_DetachCaptureBaseline(hwnd)

        if CHROME_DETACH_PREFLIGHT_TAB_COUNT {
            tabCount := Chrome_DetachQuickTabCount(hwnd)
            if (tabCount = 1) {
                ShowCenteredOverlay_Utils(
                    "❌ Cannot detach: only one tab in this window.",
                    2800, BANNER_ACCENT_ERROR)
                return false
            }
        }

        if CHROME_DETACH_USE_EXTENSION {
            try StandardLoadingBar_Update("⏳ Detaching tab (extension)…", BANNER_ACCENT_INTERMEDIATE)
            extBudget := Chrome_DetachExtensionBudgetMs(opStart)
            if (extBudget >= 300)
                newHwnd := Chrome_DetachActiveTabToNewWindow_ExtensionFast(baseline, extBudget)
            if newHwnd {
                pathTaken := "extension"
                success := true
                return true
            }
        }

        remaining := Chrome_DetachRemainingMs(opStart)
        if CHROME_DETACH_USE_UIA_SINGLE_SHOT && remaining > 600 {
            try StandardLoadingBar_Update("⏳ Detaching tab (menu)…", BANNER_ACCENT_INTERMEDIATE)
            uiaVerifyMs := Min(remaining, CHROME_DETACH_VERIFY_TIMEOUT_MS)
            newHwnd := Chrome_DetachActiveTabToNewWindow_UiaSingleShot(hwnd, &wasF11, baseline, uiaVerifyMs)
            if newHwnd {
                pathTaken := "UIA single-shot"
                success := true
                return true
            }
        }

        if !CHROME_DETACH_ALLOW_UIA_DEBUG_FALLBACK {
            detachErr := CHROME_DETACH_USE_EXTENSION
                ? "❌ Could not detach tab. Install PopActiveTab (Ctrl+Shift+Y) or use 2+ tabs."
                    : "❌ Could not detach tab. Use 2+ tabs in this Chrome window."
            ShowCenteredOverlay_Utils(detachErr, 2800, BANNER_ACCENT_ERROR)
            return false
        }

        remaining := CHROME_DETACH_TOTAL_TIMEOUT_MS - (A_TickCount - opStart)
        if (remaining <= 800)
            return false

        try StandardLoadingBar_Update("⏳ Detaching tab (UIA debug)…", BANNER_ACCENT_INTERMEDIATE)
        if Chrome_DetachActiveTabToNewWindow_UiaLegacyPath(hwnd, &newHwnd, &wasF11) {
            pathTaken := "UIA legacy"
            success := true
            return true
        }
        return false
    } finally {
        ; #region perf
        if (CHROME_DETACH_PERF_LOG_ENABLED) {
            totalMs := A_TickCount - opStart
            try FileAppend Format("main: path={} success={} totalMs={} at={}`n",
                pathTaken, success ? 1 : 0, totalMs, A_Now), A_ScriptDir "\.cursor\chrome_detach_perf.log"
        }
        ; #endregion
        if (wasF11) {
            try StandardLoadingBar_Update("🔄 Restoring F11 fullscreen…", BANNER_ACCENT_INTERMEDIATE)
            Chrome_FinishTabDetach(hwnd, success ? newHwnd : 0, wasF11)
            if (!success && hwnd && WinExist("ahk_id " hwnd) && !WM_WindowIsF11Fullscreen(hwnd))
                Chrome_RestoreF11OnOriginal(hwnd)
        } else if (success && newHwnd) {
            Chrome_FocusDetachedWindow(newHwnd)
        }
        if (success && newHwnd)
            try StandardLoadingBar_Update("✅ Tab detached to new window", BANNER_ACCENT_SUCCESS)
        if (success)
            StandardLoadingBar_Hide(CHROME_DETACH_SUCCESS_HIDE_MS)
        else
            StandardLoadingBar_Hide(0)
        if (!success)
            try Send "{Escape}"
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_DetachActiveTabToNewWindow", "finish", "F", { success: success ? 1 : 0,
            newHwnd: newHwnd })
        ; #endregion
        g_ChromeDetachBusy := false
    }
}
