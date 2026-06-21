; =============================================================================
; Utils module: chrome_detach_02.ahk
; Chrome detach context menu phases (part 2)
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; --- State-gated context menu phases (detach tab) ---
Chrome_WaitForTabContextMenu(session, timeoutMs := 0) {
    waitMs := timeoutMs > 0 ? timeoutMs : CHROME_DETACH_MENU_POPUP_MS
    deadline := A_TickCount + waitMs
    while (A_TickCount < deadline) {
        if !session.menuPopupHwnd
            Chrome_ContextMenuCapturePopup(session, session.baselinePopups)
        if Chrome_ContextMenuLooksLikeTabMenu(session)
            return true
        if Chrome_ContextMenuFindTabMenuInBrowser(session) {
            session.menuPopupClassify := "tab"
            return true
        }
        Sleep CHROME_DETACH_MENU_POLL_MS
    }
    return false
}

Chrome_FindDetachMenuTarget(session) {
    flat := Chrome_ContextMenuFindFirst(session, CHROME_DETACH_MENU_EN_NAMES, false)
    if flat
        return { type: "flat", item: flat }
    parent := Chrome_ContextMenuFindFirst(session, CHROME_DETACH_MENU_PARENT_NAMES, true)
    if parent
        return { type: "parent", item: parent }
    return 0
}

Chrome_WaitForDetachMenuTarget(session, timeoutMs := 0) {
    waitMs := timeoutMs > 0 ? timeoutMs : CHROME_DETACH_MENU_CHILD_MS
    deadline := A_TickCount + waitMs
    while (A_TickCount < deadline) {
        target := Chrome_FindDetachMenuTarget(session)
        if target
            return target
        if !session.menuPopupHwnd
            Chrome_ContextMenuCapturePopup(session, session.baselinePopups)
        Sleep CHROME_DETACH_MENU_POLL_MS
    }
    return 0
}

Chrome_ContextMenuIsDismissed(session) {
    if (session.menuPopupHwnd && WinExist("ahk_id " session.menuPopupHwnd))
        return false
    try {
        if Chrome_ContextMenuFindTabMenuInBrowser(session)
            return false
    } catch {
    }
    try {
        if Chrome_ContextMenuFocusedLooksLikeTabMenu()
            return false
    } catch {
    }
    return true
}

Chrome_WaitForContextMenuDismissed(session, timeoutMs := 0) {
    waitMs := timeoutMs > 0 ? timeoutMs : CHROME_DETACH_MENU_POPUP_MS
    deadline := A_TickCount + waitMs
    while (A_TickCount < deadline) {
        if Chrome_ContextMenuIsDismissed(session) {
            Chrome_DetachClearMenuPopup(session)
            return true
        }
        Sleep CHROME_DETACH_MENU_POLL_MS
    }
    return false
}

; Invoke detach item after menu is ready. Window wait is caller's job (dismiss wait removed — hook detects HWND faster).
Chrome_ActivateDetachMenuItemConfirmed(session, menuReady := false) {
    if !menuReady && !Chrome_WaitForTabContextMenu(session, CHROME_DETACH_MENU_POPUP_MS)
        return false
    target := Chrome_FindDetachMenuTarget(session)
    if !target
        target := Chrome_WaitForDetachMenuTarget(session, CHROME_DETACH_MENU_CHILD_MS)
    if !target
        return false
    if (target.type = "flat")
        return Chrome_ContextMenuActivateItem(target.item)
    if !Chrome_ContextMenuActivateItem(target.item)
        return false
    child := Chrome_ContextMenuWaitForSession(session, CHROME_DETACH_MENU_CHILD_NAMES, CHROME_DETACH_MENU_CHILD_MS,
        false)
    if !child
        return false
    return Chrome_ContextMenuActivateItem(child)
}

Chrome_WindowHasCaption(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    try {
        style := DllCall("GetWindowLongPtr", "ptr", hwnd, "int", -16, "ptr")
        return !!(style & 0x00C00000)
    } catch {
        return false
    }
}

; Safe for F6 / tab context menu when NOT in F11 (F6 in F11 can open DevTools).
Chrome_WaitUntilNotF11ForDetach(hwnd, timeoutMs := 0) {
    settleMs := timeoutMs > 0 ? timeoutMs : CHROME_DETACH_F11_SETTLE_MS
    deadline := A_TickCount + settleMs
    loop {
        if (!hwnd || !WinExist("ahk_id " hwnd))
            return false
        isF11 := WM_WindowIsF11Fullscreen(hwnd)
        if (!isF11) {
            fgOk := WM_EnsureForegroundForSend(hwnd, Min(800, Max(0, deadline - A_TickCount)))
            return true
        }
        if (A_TickCount >= deadline)
            break
        Sleep 50
    }
    return false
}

; Legacy name used for new-window strip; only requires leaving F11, not caption.
Chrome_WaitUntilWindowedForDetach(hwnd, timeoutMs := 0) {
    return Chrome_WaitUntilNotF11ForDetach(hwnd, timeoutMs)
}

Chrome_WaitUntilF11ForHwnd(hwnd, timeoutMs := 0) {
    settleMs := timeoutMs > 0 ? timeoutMs : CHROME_DETACH_F11_SETTLE_MS
    return WM_WaitForF11State(hwnd, true, settleMs)
}

Chrome_ExitF11ForDetach(hwnd) {
    if (!WM_WindowIsF11Fullscreen(hwnd))
        return Chrome_WaitUntilWindowedForDetach(hwnd)
    loop 3 {
        if !WM_ExitF11FullscreenForHwnd(hwnd, CHROME_DETACH_F11_SETTLE_MS)
            continue
        if Chrome_WaitUntilNotF11ForDetach(hwnd)
            return true
    }
    return false
}

Chrome_EnsureNewWindowIsWindowed(newHwnd) {
    if (!newHwnd || !WinExist("ahk_id " newHwnd))
        return false
    if (!WM_WindowIsF11Fullscreen(newHwnd))
        return true
    WM_ExitF11FullscreenForHwnd(newHwnd, CHROME_DETACH_F11_SETTLE_MS)
    return !WM_WindowIsF11Fullscreen(newHwnd)
}

Chrome_FocusDetachedWindow(newHwnd) {
    if (!newHwnd || !WinExist("ahk_id " newHwnd))
        return false
    if WinActive("ahk_id " newHwnd)
        return true
    try WinActivate("ahk_id " newHwnd)
    catch {
        return false
    }
    return WinWaitActive("ahk_id " newHwnd, , 0.25) || WinActive("ahk_id " newHwnd)
}

Chrome_RestoreF11OnOriginal(originalHwnd) {
    if (!WinExist("ahk_id " originalHwnd))
        return false
    loop 3 {
        if !WM_EnterF11FullscreenForHwnd(originalHwnd, CHROME_DETACH_F11_SETTLE_MS)
            continue
        if Chrome_WaitUntilF11ForHwnd(originalHwnd)
            return true
    }
    return WM_WindowIsF11Fullscreen(originalHwnd)
}

Chrome_ActivateDetachedWindow(newHwnd, originalHwnd, wasF11) {
    Chrome_FocusDetachedWindow(newHwnd)
}

Chrome_WaitForNewWindow(existingSet, timeoutMs := 0, tabTitle := "") {
    return Chrome_WaitForDetachNewWindow(existingSet, timeoutMs, tabTitle)
}

Chrome_EnsureBrowserForeground(hwnd) {
    if !(hwnd is Integer && hwnd > 0)
        return false
    if WinActive("ahk_id " hwnd)
        return true
    try WinActivate("ahk_id " hwnd)
    catch {
        return false
    }
    return WinWaitActive("ahk_id " hwnd, , 1)
}

Chrome_IsValidTabElement(tab) {
    if (!IsObject(tab) || !tab)
        return false
    try return tab.Type = UIA.Type.TabItem
    catch {
        return false
    }
}

Chrome_DetachGetWindowTitleForMatch(hwnd) {
    try title := WinGetTitle("ahk_id " hwnd)
    catch {
        return ""
    }
    return RegExReplace(title, "i) - Google Chrome$", "")
}

Chrome_DetachGetActiveTab(session) {
    cached := session.activeTab
    if Chrome_IsValidTabElement(cached) {
        try {
            if (cached.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable)
            && cached.SelectionItemPattern.IsSelected)
                return cached
        } catch {
        }
    }
    uia := session.uia
    method := "none"
    tabCount := -1
    if !IsObject(uia) {
        try {
            uia := session.uia := UIA_Browser("ahk_id " session.hwnd)
        } catch {
            ; #region agent log
            Chrome_DetachDebugLog("Chrome_DetachGetActiveTab", "uia attach failed", "A", { tabCount: 0, method: "none" })
            ; #endregion
            return 0
        }
    }
    try {
        uia.GetCurrentMainPaneElement()
        try {
            tabCount := uia.GetAllTabs().Length
        } catch {
            tabCount := -1
        }
        try {
            tab := uia.GetTab("")
            if Chrome_IsValidTabElement(tab) {
                method := "getTab"
                session.activeTab := tab
                ; #region agent log
                Chrome_DetachDebugLog("Chrome_DetachGetActiveTab", "tab resolved", "A", { tabCount: tabCount,
                    method: method })
                ; #endregion
                return tab
            }
        } catch {
        }
        for t in uia.GetAllTabs() {
            try {
                if (t.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable)
                && t.SelectionItemPattern.IsSelected) {
                    if Chrome_IsValidTabElement(t) {
                        method := "selection"
                        session.activeTab := t
                        ; #region agent log
                        Chrome_DetachDebugLog("Chrome_DetachGetActiveTab", "tab resolved", "A", { tabCount: tabCount,
                            method: method })
                        ; #endregion
                        return t
                    }
                }
            } catch {
            }
        }
        chromeTitle := Chrome_DetachGetWindowTitleForMatch(session.hwnd)
        if (chromeTitle != "") {
            for t in uia.GetAllTabs() {
                try tabName := t.Name
                catch {
                    continue
                }
                if (tabName = "" || !Chrome_IsValidTabElement(t))
                    continue
                if (tabName = chromeTitle || InStr(chromeTitle, tabName, false) || InStr(tabName, chromeTitle, false)) {
                    method := "title"
                    session.activeTab := t
                    ; #region agent log
                    Chrome_DetachDebugLog("Chrome_DetachGetActiveTab", "tab resolved", "A", { tabCount: tabCount,
                        method: method, titleLen: StrLen(chromeTitle) })
                    ; #endregion
                    return t
                }
            }
        }
    } catch {
    }
    ; #region agent log
    Chrome_DetachDebugLog("Chrome_DetachGetActiveTab", "tab not found", "A", { tabCount: tabCount, method: method })
    ; #endregion
    return 0
}

Chrome_ContextMenuLooksLikeTabMenu(session) {
    if !session.menuPopupHwnd
        Chrome_ContextMenuCapturePopup(session, session.baselinePopups)
    if (session.menuPopupHwnd && session.menuPopupClassify = "tab") {
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_ContextMenuLooksLikeTabMenu", "tab menu cached", "D", { isTabMenu: true,
            popupHwnd: session.menuPopupHwnd })
        ; #endregion
        return true
    }
    if !session.menuPopupHwnd {
        if Chrome_ContextMenuFindTabMenuInBrowser(session) {
            ; #region agent log
            Chrome_DetachDebugLog("Chrome_ContextMenuLooksLikeTabMenu", "tab menu in browser tree", "G", { isTabMenu: true })
            ; #endregion
            return true
        }
        if Chrome_ContextMenuFocusedLooksLikeTabMenu() {
            ; #region agent log
            Chrome_DetachDebugLog("Chrome_ContextMenuLooksLikeTabMenu", "tab menu focused", "G", { isTabMenu: true })
            ; #endregion
            return true
        }
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_ContextMenuLooksLikeTabMenu", "no popup hwnd", "D", { isTabMenu: false,
            newWins: Chrome_DetachDebugListNewChromeWindows(session) })
        ; #endregion
        return false
    }
    inspected := Chrome_ContextMenuInspectPopupHwnd(session.menuPopupHwnd)
    if (inspected.classify = "page") {
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_ContextMenuLooksLikeTabMenu", "page menu rejected", "D", { isTabMenu: false,
            popupHwnd: session.menuPopupHwnd, menuSample: inspected.sample })
        ; #endregion
        return false
    }
    if (inspected.classify = "tab") {
        session.menuPopupClassify := "tab"
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_ContextMenuLooksLikeTabMenu", "tab menu confirmed", "D", { isTabMenu: true,
            popupHwnd: session.menuPopupHwnd })
        ; #endregion
        return true
    }
    deadline := A_TickCount + 150
    sample := inspected.sample
    while (A_TickCount < deadline) {
        inspected := Chrome_ContextMenuInspectPopupHwnd(session.menuPopupHwnd)
        sample := inspected.sample
        if (inspected.classify = "page") {
            ; #region agent log
            Chrome_DetachDebugLog("Chrome_ContextMenuLooksLikeTabMenu", "page menu rejected", "D", { isTabMenu: false,
                popupHwnd: session.menuPopupHwnd, menuSample: sample })
            ; #endregion
            return false
        }
        if (inspected.classify = "tab") {
            session.menuPopupClassify := "tab"
            ; #region agent log
            Chrome_DetachDebugLog("Chrome_ContextMenuLooksLikeTabMenu", "tab menu confirmed", "D", { isTabMenu: true,
                popupHwnd: session.menuPopupHwnd })
            ; #endregion
            return true
        }
        Sleep CHROME_DETACH_MENU_POLL_MS
    }
    ; #region agent log
    Chrome_DetachDebugLog("Chrome_ContextMenuLooksLikeTabMenu", "not tab menu", "D", { isTabMenu: false,
        popupHwnd: session.menuPopupHwnd, menuSample: sample })
    ; #endregion
    return false
}

Chrome_FocusedElementIsSelectedTab(session) {
    try {
        focused := UIA.GetFocusedElement()
        if !focused
            return false
        if (focused.Type = UIA.Type.TabItem) {
            try {
                if (focused.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable)
                && focused.SelectionItemPattern.IsSelected)
                    return true
            } catch {
            }
        }
        uia := session.uia
        if !IsObject(uia)
            return false
        try uia.GetCurrentMainPaneElement()
        tab := uia.GetTab("")
        if !Chrome_IsValidTabElement(tab)
            return false
        try {
            if UIA.CompareElements(focused, tab)
                return true
        } catch {
        }
    } catch {
    }
    return false
}

; Poll until selected tab keeps keyboard focus (avoids AppsKey on wrong F6 stop).
Chrome_WaitForSelectedTabFocus(session, timeoutMs := 0) {
    waitMs := timeoutMs > 0 ? timeoutMs : CHROME_DETACH_TAB_FOCUS_WAIT_MS
    stableNeed := CHROME_DETACH_TAB_FOCUS_STABLE_MS
    deadline := A_TickCount + waitMs
    stableSince := 0
    while (A_TickCount < deadline) {
        if Chrome_FocusedElementIsSelectedTab(session) {
            if !stableSince
                stableSince := A_TickCount
            if (A_TickCount - stableSince >= stableNeed) {
                ; #region agent log
                Chrome_DetachDebugLog("Chrome_WaitForSelectedTabFocus", "stable tab focus", "I", { waitedMs: A_TickCount -
                    (deadline - waitMs), stableMs: stableNeed })
                ; #endregion
                return true
            }
        } else {
            stableSince := 0
        }
        Sleep CHROME_DETACH_MENU_POLL_MS
    }
    ; #region agent log
    Chrome_DetachDebugLog("Chrome_WaitForSelectedTabFocus", "timeout", "I", { waitMs: waitMs })
    ; #endregion
    return false
}

; Hover tab (no click), settle, AppsKey — Chrome hit-tests cursor for context menu.
Chrome_TryHoverAppsKeyTabMenu(session, tab, settleMs := 0) {
    if !Chrome_HoverActiveTab(session, tab)
        return false
    waitMs := settleMs > 0 ? settleMs : CHROME_DETACH_HOVER_APPSKEY_SETTLE_MS
    Sleep waitMs
    ClipAngel_ReleaseChordModifiersForSend()
    session.baselinePopups := Chrome_DetachListMenuPopups(session.hwnd)
    Chrome_DetachClearMenuPopup(session)
    ; #region agent log
    Chrome_DetachDebugLog("Chrome_TryHoverAppsKeyTabMenu", "AppsKey after hover", "J", "focus=" .
        Chrome_DetachDebugFocusedElement(session) . ";settleMs=" . waitMs)
    ; #endregion
    SendInput "{AppsKey}"
    Sleep CHROME_DETACH_APPSKEY_AFTER_MS
    Chrome_ContextMenuCapturePopup(session, session.baselinePopups)
    if Chrome_ContextMenuLooksLikeTabMenu(session)
        return true
    if Chrome_ContextMenuFindTabMenuInBrowser(session) {
        Chrome_DetachClearMenuPopup(session)
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_TryHoverAppsKeyTabMenu", "tab menu in browser tree", "G", "")
        ; #endregion
        return true
    }
    ; #region agent log
    Chrome_DetachDebugLog("Chrome_TryHoverAppsKeyTabMenu", "not tab menu", "D", "popup=" .
        session.menuPopupHwnd . ";newWins=" . Chrome_DetachDebugListNewChromeWindows(session))
    ; #endregion
    Chrome_ContextMenuDismiss()
    return false
}

; F6/SetFocus path: keyboard focus on tab, then AppsKey via ControlSend.
Chrome_TryFocusAppsKeyTabMenu(session, tab) {
    if !Chrome_DetachFocusActiveTab(session)
        return false
    Chrome_HoverActiveTab(session, tab)
    Sleep CHROME_DETACH_HOVER_SETTLE_MS
    if !Chrome_WaitForSelectedTabFocus(session)
        return false
    Sleep CHROME_DETACH_APPSKEY_SETTLE_MS
    ClipAngel_ReleaseChordModifiersForSend()
    session.baselinePopups := Chrome_DetachListMenuPopups(session.hwnd)
    Chrome_DetachClearMenuPopup(session)
    ; #region agent log
    Chrome_DetachDebugLog("Chrome_TryFocusAppsKeyTabMenu", "AppsKey after focus", "J", "focus=" .
        Chrome_DetachDebugFocusedElement(session))
    ; #endregion
    try session.uia.ControlSend("{AppsKey}")
    catch {
        SendInput "{AppsKey}"
    }
    Sleep CHROME_DETACH_APPSKEY_AFTER_MS
    Chrome_ContextMenuCapturePopup(session, session.baselinePopups)
    if Chrome_ContextMenuLooksLikeTabMenu(session)
        return true
    Chrome_ContextMenuDismiss()
    return false
}

Chrome_UiaLooksLikeTabStripFocused(session) {
    if Chrome_FocusedElementIsSelectedTab(session)
        return true
    try {
        focused := UIA.GetFocusedElement()
        if focused && focused.Type = UIA.Type.TabItem
            return true
    } catch {
    }
    return false
}

Chrome_WaitForTabStripFocus(session, timeoutMs := 0) {
    global CHROME_DETACH_F6_FOCUS_POLL_MS, CHROME_DETACH_MENU_POLL_MS
    waitMs := timeoutMs > 0 ? timeoutMs : CHROME_DETACH_F6_FOCUS_POLL_MS
    deadline := A_TickCount + waitMs
    pollMs := CHROME_DETACH_MENU_POLL_MS
    while (A_TickCount < deadline) {
        if Chrome_UiaLooksLikeTabStripFocused(session)
            return true
        Sleep pollMs
    }
    return false
}

; Official keyboard path: F6 until tab strip focus, then AppsKey (no hover hit-test).
Chrome_TryF6KeyboardTabMenu(session) {
    ; Modifiers already released by Chrome_Detach_PreSendSanitizeModifiers — short safety check only.
    ClipAngel_ReleaseChordModifiersForSend()
    Send "{Escape}"
    deadline := A_TickCount + 30
    while (A_TickCount < deadline) {
        if !GetKeyState("Shift", "P") && !GetKeyState("Ctrl", "P") && !GetKeyState("Alt", "P")
            break
        Sleep 10
    }
    uiaUsable := Chrome_SessionUiaUsable(session)
    loop CHROME_DETACH_F6_FOCUS_MAX {
        if (session.menuPopupHwnd && Chrome_ContextMenuLooksLikeTabMenu(session)) {
            session.menuConfirmed := true
            return true
        }
        Send "{F6}"
        tabFocused := Chrome_WaitForTabStripFocus(session, CHROME_DETACH_F6_FOCUS_POLL_MS)
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_TryF6KeyboardTabMenu", "f6 step", "K", "i=" . A_Index .
            " focus=" . Chrome_DetachDebugFocusedElement(session) . ";tabFocused=" . (tabFocused ? 1 : 0))
        ; #endregion
        tryAppsKey := tabFocused
        if !tryAppsKey && !uiaUsable && A_Index >= 2
            tryAppsKey := true  ; UIA dead: AppsKey from 2nd F6 onward, confirm via menu wait
        if !tryAppsKey
            continue
        ClipAngel_ReleaseChordModifiersForSend()
        session.baselinePopups := Chrome_DetachListMenuPopups(session.hwnd)
        Chrome_DetachClearMenuPopup(session)
        try session.uia.ControlSend("{AppsKey}")
        catch {
            SendInput "{AppsKey}"
        }
        if Chrome_WaitForTabContextMenu(session, CHROME_DETACH_MENU_POPUP_MS) {
            session.menuConfirmed := true
            ; #region agent log
            Chrome_DetachDebugLog("Chrome_TryF6KeyboardTabMenu", "tab menu ready", "K", "i=" . A_Index)
            ; #endregion
            return true
        }
        Chrome_ContextMenuDismiss()
    }
    return false
}

Chrome_OpenActiveTabContextMenuViaTabFocus(session) {
    tab := Chrome_DetachGetActiveTab(session)
    if (session.menuPopupHwnd && Chrome_ContextMenuLooksLikeTabMenu(session))
        return true
    if !Chrome_IsValidTabElement(tab)
        return false
    if Chrome_TryF6KeyboardTabMenu(session)
        return true
    if Chrome_TryFocusAppsKeyTabMenu(session, tab)
        return true
    loop CHROME_DETACH_HOVER_ATTEMPTS {
        settle := CHROME_DETACH_HOVER_APPSKEY_SETTLE_MS + (A_Index - 1) * CHROME_DETACH_HOVER_APPSKEY_RETRY_EXTRA_MS
        if Chrome_TryHoverAppsKeyTabMenu(session, tab, settle)
            return true
    }
    if !CHROME_DETACH_F6_FALLBACK
        return false
    ClipAngel_ReleaseChordModifiersForSend()
    Send "{Escape}"
    Sleep 40
    loop CHROME_DETACH_F6_FOCUS_MAX {
        if (session.menuPopupHwnd && Chrome_ContextMenuLooksLikeTabMenu(session))
            return true
        Send "{F6}"
        Sleep CHROME_DETACH_F6_STEP_MS
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_OpenActiveTabContextMenuViaTabFocus", "f6+hover step", "C", "i=" . A_Index .
            " focus=" . Chrome_DetachDebugFocusedElement(session))
        ; #endregion
        if Chrome_TryHoverAppsKeyTabMenu(session, tab)
            return true
    }
    return false
}

; Move pointer over tab strip target only (no click).
Chrome_HoverActiveTab(session, tab) {
    if !Chrome_IsValidTabElement(tab)
        return false
    try {
        rect := tab.BoundingRectangle
        if !(rect.r > rect.l && rect.b > rect.t)
            return false
        x := rect.l + (rect.r - rect.l) // 2
        y := rect.t + (rect.b - rect.t) // 2
        uia := session.uia
        if IsObject(uia) {
            try {
                uia.GetCurrentMainPaneElement()
                tbRect := uia.TabBarElement.BoundingRectangle
                if (tbRect.b > tbRect.t)
                    y := tbRect.t + (tbRect.b - tbRect.t) // 2
            } catch {
            }
        }
        saveMode := A_CoordModeMouse
        CoordMode "Mouse", "Screen"
        MouseMove x, y, 0
        CoordMode "Mouse", saveMode
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_HoverActiveTab", "hovered tab", "H", { x: x, y: y })
        ; #endregion
        return true
    } catch {
        return false
    }
}

Chrome_DetachFocusActiveTab(session) {
    tab := Chrome_DetachGetActiveTab(session)
    if !Chrome_IsValidTabElement(tab)
        return false
    focused := false
    try {
        tab.SetFocus()
        focused := true
    } catch {
    }
    if !focused {
        try {
            if (tab.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable))
                tab.SelectionItemPattern.Select()
            focused := true
        } catch {
        }
    }
    ok := Chrome_WaitForSelectedTabFocus(session, 350)
    ; #region agent log
    Chrome_DetachDebugLog("Chrome_DetachFocusActiveTab", "focus tab", "H", { setFocus: focused ? 1 : 0,
        tabFocused: ok ? 1 : 0 })
    ; #endregion
    return ok
}

Chrome_FocusTabStripAndOpenContextMenu(session) {
    hwnd := session.hwnd
    if !Chrome_NormalizeFocusToPage(hwnd) {
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_FocusTabStripAndOpenContextMenu", "normalize page failed", "C", { result: 0 })
        ; #endregion
        return false
    }
    ClipAngel_ReleaseChordModifiersForSend()
    loop CHROME_DETACH_F6_FOCUS_MAX {
        if (session.menuPopupHwnd && Chrome_ContextMenuLooksLikeTabMenu(session))
            return true
        Send "{F6}"
        Sleep CHROME_DETACH_F6_STEP_MS
        tab := Chrome_DetachGetActiveTab(session)
        if !Chrome_IsValidTabElement(tab)
            continue
        try tab.SetFocus()
        catch {
        }
        Sleep CHROME_DETACH_F6_REFOCUS_MS
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_FocusTabStripAndOpenContextMenu", "f6 iteration", "C", { iteration: A_Index,
            stepMs: CHROME_DETACH_F6_STEP_MS, focus: Chrome_DetachDebugFocusedElement(session) })
        ; #endregion
        if Chrome_UiaLooksLikeTabStripFocused(session) {
            Sleep CHROME_DETACH_APPSKEY_SETTLE_MS
            ClipAngel_ReleaseChordModifiersForSend()
            session.baselinePopups := Chrome_DetachListMenuPopups(session.hwnd)
            Chrome_DetachClearMenuPopup(session)
            try session.uia.ControlSend("{AppsKey}")
            catch {
                SendInput "{AppsKey}"
            }
            Sleep CHROME_DETACH_APPSKEY_AFTER_MS
            Chrome_ContextMenuCapturePopup(session, session.baselinePopups)
            if Chrome_ContextMenuLooksLikeTabMenu(session) {
                ; #region agent log
                Chrome_DetachDebugLog("Chrome_FocusTabStripAndOpenContextMenu", "f6+AppsKey success", "C", { result: 1,
                    iteration: A_Index })
                ; #endregion
                return true
            }
            Chrome_ContextMenuDismiss()
        }
        if Chrome_TryHoverAppsKeyTabMenu(session, tab) {
            ; #region agent log
            Chrome_DetachDebugLog("Chrome_FocusTabStripAndOpenContextMenu", "f6+hover AppsKey success", "C", { result: 1,
                iteration: A_Index })
            ; #endregion
            return true
        }
        Chrome_ContextMenuDismiss()
    }
    ; #region agent log
    Chrome_DetachDebugLog("Chrome_FocusTabStripAndOpenContextMenu", "all paths failed", "C", { result: 0 })
    ; #endregion
    return false
}

Chrome_FocusPageContent(hwnd) {
    try {
        for ctrlName in ["Chrome_RenderWidgetHostHWND1", "Chrome_RenderWidgetHostHWND"] {
            try {
                rw := ControlGetHwnd(ctrlName, "ahk_id " hwnd)
                if (rw) {
                    ControlFocus ctrlName, "ahk_id " hwnd
                    return true
                }
            } catch {
            }
        }
    } catch {
    }
    return false
}
