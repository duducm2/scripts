; =============================================================================
; Utils module: clip_angel_activate.ahk
; Clip Angel activate with focus correction
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Clip Angel: Open/Activate with focus correction (Row 0)
; =============================================================================
; Show/restore always-running Clip Angel via AHK (no native Alt+P paste hotkey). UIA: dataGridView,
; Row 0 (first clip) per clip-angel.txt. One ElementFromHandle per flow; bounded polls; layout only when not foreground/hidden.
; skipRow0: true when caller will select Row 0 after filter change (e.g. OnSubmitO leaving favorites).
; forceLayout: true always move/maximize onto targetMon (toggle-open path).

; Post WM_HOTKEY (0x0312) to ClipAngel process windows — triggers native RegisterHotKey
; dispatch without SendInput / focus steal. modFlags: subset of "casw" (Ctrl/Alt/Shift/Win).
; Returns true if at least one PostMessage succeeded.
ClipAngel_PostHotkey(vkOrChar, modFlags := "") {
    static WM_HOTKEY := 0x0312
    static MOD_ALT := 1, MOD_CONTROL := 2, MOD_SHIFT := 4, MOD_WIN := 8
    if (vkOrChar = "")
        return false
    vk := 0
    if (vkOrChar is Integer)
        vk := Integer(vkOrChar)
    else {
        s := StrLower(Trim(vkOrChar))
        if (StrLen(s) = 1) {
            c := Ord(s)
            if (c >= Ord("a") && c <= Ord("z"))
                vk := c - Ord("a") + 0x41
            else if (c >= Ord("0") && c <= Ord("9"))
                vk := c
        }
    }
    if (vk <= 0)
        return false
    mods := 0
    mf := StrLower(modFlags)
    if InStr(mf, "a")
        mods |= MOD_ALT
    if InStr(mf, "c")
        mods |= MOD_CONTROL
    if InStr(mf, "s")
        mods |= MOD_SHIFT
    if InStr(mf, "w")
        mods |= MOD_WIN
    lParam := (vk << 16) | mods
    posted := false
    prevDetect := A_DetectHiddenWindows
    DetectHiddenWindows true
    try {
        for hwnd in WinGetList("ahk_exe ClipAngel.exe") {
            try {
                PostMessage(WM_HOTKEY, 0, lParam, , "ahk_id " hwnd)
                posted := true
            } catch {
            }
        }
    } finally {
        DetectHiddenWindows prevDetect
    }
    return posted
}

ActivateClipAngelWithFocusCorrection(silent := false, targetMon := 0, skipRow0 := false, forceLayout := false) {
    needBanner := false
    if !targetMon {
        try targetMon := GetAhkMonitorIndexFromHwnd(WinGetID("A"))
        catch
            targetMon := 0
    }
    hwnd := ClipAngel_MainHwnd()
    if !hwnd {
        if !silent
            ShowCenteredOverlay_Utils("❌ Clip Angel não está em execução.", 2000, BANNER_ACCENT_ERROR)
        return false
    }
    isActive := WinActive("ahk_id " hwnd)
    wasHidden := !ClipAngel_IsWindowShown(hwnd)
    wrongMonitor := false
    if (targetMon >= 1) {
        try wrongMonitor := (GetAhkMonitorIndexFromHwnd(hwnd) != targetMon)
        catch
            wrongMonitor := true
    }
    needsLayout := forceLayout || !isActive || wasHidden || wrongMonitor || ClipAngel_NeedsLayoutCorrection(hwnd)
    if isActive && !needsLayout {
        if !skipRow0
            ClipAngel_UiaEnsureRow0Selected(hwnd, false)
        return true
    }
    if wasHidden || !isActive || wrongMonitor || forceLayout {
        needBanner := !silent
        if needBanner
            StandardLoadingBar_Show("⏳ Opening Clip Angel...", BANNER_ACCENT_INTERMEDIATE, {
                passive: false,
                fontSize: 17
            })
    }
    if !ClipAngel_ShowWindow(hwnd) {
        if needBanner {
            try StandardLoadingBar_Hide(0)
            catch {
            }
        }
        if !silent
            ShowCenteredOverlay_Utils("❌ ClipAngel window not found.", 2000, BANNER_ACCENT_ERROR)
        return false
    }
    ; ApplyLayoutOnMonitor already activates on its fast/full paths — only activate when layout skipped.
    if needsLayout
        ClipAngel_ApplyLayoutOnMonitor(hwnd, targetMon)
    else
        ClipAngel_EnsureWindowActive(hwnd)
    if !skipRow0
        ClipAngel_UiaEnsureRow0Selected(hwnd, true)
    if needBanner {
        StandardLoadingBar_Update("✅ Clip Angel ready", BANNER_ACCENT_SUCCESS)
        StandardLoadingBar_Hide(350)
    }
    return true
}

; Shared open for paste / export / merge / suppress paths: show+layout, MarkFilter, Row 0.
; suppressVisual: park off-screen (favorite/paste/merge/export) — no foreground activate / loading bar.
; Returns true and sets &outHwnd / &outRoot when provided.
; Confirm Clip Angel startup hotkey registration dialogs (#32770 for Alt+P / Alt+B)
; in the background without stealing focus so the main window can finish opening.
ClipAngel_DismissStartupHotkeyDialogs(newPid := 0) {
    dismissed := 0
    prevDetect := A_DetectHiddenWindows
    DetectHiddenWindows true
    try {
        for dlgHwnd in WinGetList("ahk_exe ClipAngel.exe ahk_class #32770") {
            if (newPid) {
                try {
                    if (WinGetPID("ahk_id " dlgHwnd) != newPid)
                        continue
                } catch {
                    continue
                }
            }
            txt := ""
            try txt := StrLower(WinGetText("ahk_id " dlgHwnd))
            catch {
                txt := ""
            }
            isHotkeyDlg := (txt = ""
                || InStr(txt, "hotkey") || InStr(txt, "shortcut") || InStr(txt, "atalho")
                || InStr(txt, "alt") || InStr(txt, "register") || InStr(txt, "attribute")
                || InStr(txt, "could not") || InStr(txt, "failed") || !ClipAngel_MainHwnd())
            if (!isHotkeyDlg)
                continue
            try WinSetTransparent(0, "ahk_id " dlgHwnd)
            catch {
            }
            clicked := false
            try {
                ControlClick("Button1", "ahk_id " dlgHwnd, , "Left", 1, "NA")
                clicked := true
            } catch {
                clicked := false
            }
            if !clicked {
                try {
                    PostMessage(0x00F5, 0, 0, "Button1", "ahk_id " dlgHwnd) ; BM_CLICK
                    clicked := true
                } catch {
                    clicked := false
                }
            }
            if !clicked {
                try PostMessage(0x0112, 0xF060, 0, , "ahk_id " dlgHwnd) ; WM_CLOSE
                catch {
                }
            }
            dismissed += 1
        }
    } finally {
        DetectHiddenWindows prevDetect
    }
    return dismissed
}

ClipAngel_HasMainWindowChildren(hwnd) {
    if !hwnd
        return false
    try return WinGetControlsHwnd("ahk_id " hwnd).Length > 0
    catch
        return false
}

ClipAngel_OpenForAutomation(mode := "all", targetMon := 0, suppressVisual := false, &outHwnd := 0, &outRoot := 0, skipFilterAndRow0 := false, useNativeHotkey := true) {
    wantAll := (mode = "all" || mode = "")
    outHwnd := 0
    outRoot := 0
    priorFg := 0
    try priorFg := WinGetID("A")
    catch {
        priorFg := 0
    }
    ClipAngel_DismissStartupHotkeyDialogs()
    hwnd := ClipAngel_MainHwnd()
    if !hwnd && ProcessExist("ClipAngel.exe") {
        loop 2 {
            if !ClipAngel_PostHotkey("p", "a") {
                priorSendLevel := A_SendLevel
                SendLevel 0
                ClipAngel_ReleaseChordModifiersForSend()
                SendInput("!p")
                SendLevel priorSendLevel
            }
            hwnd := ClipAngel_WaitForMainHwnd(500)
            if hwnd
                break
        }
    }
    if !hwnd {
        if (ClipAngel_RecoverProcessForAutomation(suppressVisual))
            hwnd := ClipAngel_MainHwnd()
    }
    if !hwnd
        return false
    if suppressVisual {
        ClipAngel_BeginFavoriteSuppress(hwnd)
        if (priorFg && WinActive("ahk_exe ClipAngel.exe"))
            ClipAngel_RestorePriorFocus(priorFg, true)
    } else {
        ; Foreground open: prefer native Alt+P (reset all + Row 0) / Alt+B (favorites + Row 0)
        ; when not skipping filter/Row 0, avoiding slow UIA MarkFilter polls.
        if (useNativeHotkey && !skipFilterAndRow0 && ClipAngel_OpenNative(wantAll ? "p" : "b", targetMon, false)) {
            hwnd := ClipAngel_MainHwnd()
            if (hwnd) {
                outHwnd := hwnd
                outRoot := 0
                return true
            }
        }
        if !ActivateClipAngelWithFocusCorrection(true, targetMon, true, false)
            return false
        hwnd := ClipAngel_MainHwnd()
        if !hwnd
            return false
    }
    ; Wait briefly for WinForms child controls (MarkFilter / dataGridView) to initialize after wake.
    childDeadline := A_TickCount + 350
    while (A_TickCount < childDeadline) {
        if ClipAngel_HasMainWindowChildren(hwnd)
            break
        Sleep CLIPANGEL_UIA_POLL_MS
    }
    if (skipFilterAndRow0) {
        outHwnd := hwnd
        outRoot := 0
        return true
    }
    ; Suppressed background "all": native F9 (ClearFilter + ReloadList + select top) avoids slow UIA walks.
    if (suppressVisual && wantAll) {
        prevDetect := A_DetectHiddenWindows
        DetectHiddenWindows true
        try ControlSend("{F9}", , "ahk_id " hwnd)
        catch {
        }
        DetectHiddenWindows prevDetect
        Sleep 30
    }
    root := 0
    try root := UIA.ElementFromHandle(hwnd)
    catch
        root := 0
    if !ClipAngel_ApplyMarkFilterMode(wantAll, hwnd, root) {
        Sleep 100
        try root := UIA.ElementFromHandle(hwnd)
        catch
            root := 0
        if !ClipAngel_ApplyMarkFilterMode(wantAll, hwnd, root) {
            if !(wantAll && ClipAngel_UiaGetDataGrid(hwnd, root))
                return false
        }
    }
    ClipAngel_FastEnsureRow0(hwnd, root)
    outHwnd := hwnd
    outRoot := root
    return true
}

; Alt+P / Alt+B: Clip Angel's own native hotkeys (Open last with reset / Open favorites).
; Post WM_HOTKEY directly, wait briefly for unpark, then synchronously place/maximize on targetMon.
; Falls back to UIA ClipAngel_OpenWithMarkFilter only if the native post fails to show the window.
ClipAngel_OpenNative(letter, targetMon := 0, playSound := true) {
    global g_ClipAngelFavoriteSuppressActive
    mode := (letter = "b") ? "favorites" : "all"
    hwnd := ClipAngel_MainHwnd()
    if !hwnd {
        if (ProcessExist("ClipAngel.exe")) {
            ClipAngel_DismissStartupHotkeyDialogs()
            ClipAngel_PostHotkey(letter, "a")
            hwnd := ClipAngel_WaitForMainHwnd(400)
        }
        if !hwnd {
            ShowCenteredOverlay_Utils("❌ Clip Angel is not running.", 2000, BANNER_ACCENT_ERROR)
            return false
        }
    }
    if (!targetMon || targetMon < 1) {
        targetMon := ClipAngel_GetMonitorIndexFromCursor()
        if (!targetMon || targetMon < 1) {
            try targetMon := MonitorGetPrimary()
            catch
                targetMon := 1
        }
    }
    g_ClipAngelFavoriteSuppressActive := false
    prevDetect := A_DetectHiddenWindows
    DetectHiddenWindows true
    try {
        try WinSetExStyle("-0x08000000", "ahk_id " hwnd)
        catch {
        }
        try WinSetTransparent("Off", "ahk_id " hwnd)
        catch {
        }
    } finally {
        DetectHiddenWindows prevDetect
    }
    alreadyHere := WinActive("ahk_id " hwnd) && ClipAngel_IsWindowShown(hwnd) && !ClipAngel_NeedsLayoutCorrection(hwnd)
    ClipAngel_ReleaseChordModifiersForSend()
    if !ClipAngel_PostHotkey(letter, "a")
        return ClipAngel_OpenWithMarkFilter(mode)
    if alreadyHere
        return true
    ; Wait briefly for native ShowForPaste to unpark the window (typically 15–40ms).
    unparked := false
    deadline := A_TickCount + 180
    while (A_TickCount < deadline) {
        if ClipAngel_IsWindowShown(hwnd) {
            unparked := true
            break
        }
        Sleep 15
    }
    if !unparked
        unparked := ClipAngel_IsWindowShown(hwnd)
    if unparked {
        ClipAngel_ApplyLayoutOnMonitor(hwnd, targetMon)
        if playSound {
            try ScriptSoundPlay(A_ScriptDir "\assets\sounds\clipangel-open.wav")
            catch {
            }
        }
        return true
    }
    return ClipAngel_OpenWithMarkFilter(mode)
}

; AHK-owned Alt+P / Alt+B open: show Clip Angel, ControlSend Ctrl+1 (all marks) or Ctrl+2 (favorites), Row 0.
; Fallback when the native hotkey post does not show the window. No ribbon validation.
; Loading Indication via StandardLoadingBar (standard_information_display.md); silent activate so one bar owns UX.
CLIPANGEL_MARKFILTER_WAIT_MS := 300

ClipAngel_MarkFilterMatchesMode(wantAll, val) {
    val := StrLower(Trim(val))
    if wantAll
        return (val = "all marks")
    return InStr(val, "favor")
}

ClipAngel_OpenWithMarkFilter(mode) {
    global g_ClipAngelAutomationBusy
    wantAll := (mode = "all")
    loadMsg := wantAll ? "⏳ Opening Clip Angel (all marks)..." : "⏳ Opening Clip Angel (favorites)..."
    doneMsg := wantAll ? "✅ Clip Angel — all marks" : "✅ Clip Angel — favorites"
    priorHwnd := 0
    try priorHwnd := WinGetID("A")
    catch
        priorHwnd := 0
    hwnd := ClipAngel_MainHwnd()
    if !hwnd {
        ShowCenteredOverlay_Utils("❌ Clip Angel is not running.", 2000, BANNER_ACCENT_ERROR)
        return false
    }
    prevBusy := g_ClipAngelAutomationBusy
    g_ClipAngelAutomationBusy := true
    StandardLoadingBar_Show(loadMsg, BANNER_ACCENT_INTERMEDIATE, {
        passive: false,
        centerOnHwnd: priorHwnd,
        fontSize: 17
    })
    try {
        targetMon := ClipAngel_GetMonitorIndexFromCursor()
        if (!targetMon || targetMon < 1) {
            try targetMon := GetAhkMonitorIndexFromHwnd(priorHwnd)
            catch
                targetMon := 0
        }
        hwnd := 0
        root := 0
        if !ClipAngel_OpenForAutomation(wantAll ? "all" : "favorites", targetMon, false, &hwnd, &root, false, false) {
            try StandardLoadingBar_Hide(0)
            catch {
            }
            ShowCenteredOverlay_Utils("❌ Could not open Clip Angel / set filter.", 2000, BANNER_ACCENT_ERROR)
            return false
        }
        try StandardLoadingBar_Hide(0)
        catch {
        }
        ; Window is fully open. One Backspace on the clip list empties a leftover search
        ; from another filter; Clip Angel does not delete clips on Backspace.
        ClipAngel_ClearSearchWithBackspace(hwnd)
        ShowCenteredOverlay_Utils(doneMsg, 900, BANNER_ACCENT_SUCCESS)
        try ScriptSoundPlay(A_ScriptDir "\assets\sounds\clipangel-open.wav")
        catch {
        }
        return true
    } catch Error as e {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        ShowCenteredOverlay_Utils("❌ Clip Angel open failed: " . e.Message, 2500, BANNER_ACCENT_ERROR)
        return false
    } finally {
        g_ClipAngelAutomationBusy := prevBusy
    }
}

; After Alt+P / Alt+B, the main window is up and the clip list has focus.
; One Backspace on the list empties a leftover search (Clip Angel selects all text in the box).
; The key leaves the search box focused; SetFocus returns to the clip list.
; An empty box skips the key so focus never moves into the field.
ClipAngel_ClearSearchWithBackspace(hwnd) {
    if !hwnd
        return
    if !WinActive("ahk_id " hwnd)
        ClipAngel_EnsureWindowActive(hwnd, 400)
    dataGrid := 0
    try dataGrid := ClipAngel_UiaGetDataGrid(hwnd)
    catch
        dataGrid := 0
    found := false
    searchText := ClipAngel_SearchStringText(hwnd, &found)
    if (found && searchText = "") {
        if dataGrid {
            try dataGrid.SetFocus()
            catch {
            }
        }
        ClipAngel_ReturnFocusFromMainMenu(hwnd)
        return
    }
    if dataGrid {
        try dataGrid.SetFocus()
        catch {
        }
    }
    ClipAngel_ReleaseChordModifiersForSend()
    priorSendLevel := A_SendLevel
    SendLevel 0
    SendInput "{Backspace}"
    SendLevel priorSendLevel
    if dataGrid {
        try dataGrid.SetFocus()
        catch {
        }
    }
    ClipAngel_ReturnFocusFromMainMenu(hwnd)
}

; After native Alt+P / Alt+B open: short settle, one maximize gate, one retry if needed.
; Efficiency canon: hotkey returns immediately; same-monitor path skips MoveWindow+Sleep.
; Legacy: unused by AHK-owned !p/!b (kept for any external callers).
CLIPANGEL_NATIVE_OPEN_SETTLE_MS := 50
CLIPANGEL_NATIVE_OPEN_RETRY_MS := 100
global g_ClipAngelNativeOpenTargetMon := 0
global g_ClipAngelNativeOpenRetryArmed := false

; Monitor under the mouse — where the user is looking.
ClipAngel_GetMonitorIndexFromCursor() {
    CoordMode "Mouse", "Screen"
    MouseGetPos &mx, &my
    try {
        loop MonitorGetCount() {
            MonitorGet A_Index, &l, &t, &r, &b
            if (mx >= l && mx < r && my >= t && my < b)
                return A_Index
        }
    } catch {
    }
    try return MonitorGetPrimary()
    catch
        return 1
}

; True when shown, maximized, on targetMon, and foreground — sole quality gate.
ClipAngel_IsForegroundLayoutOk(hwnd, targetMon) {
    if !hwnd || !ClipAngel_IsWindowShown(hwnd)
        return false
    if !WinActive("ahk_id " hwnd)
        return false
    try {
        if (WinGetMinMax("ahk_id " hwnd) != 1)
            return false
    } catch {
        return false
    }
    if (targetMon >= 1) {
        try {
            if (GetAhkMonitorIndexFromHwnd(hwnd) != targetMon)
                return false
        } catch {
            return false
        }
    }
    return true
}

; Capture look-at monitor; replace any pending settle timer (second Alt+P/B does not stack layouts).
ClipAngel_EnsureForegroundAfterNativeOpen() {
    global g_ClipAngelNativeOpenTargetMon, g_ClipAngelNativeOpenRetryArmed
    targetMon := ClipAngel_GetMonitorIndexFromCursor()
    if (!targetMon || targetMon < 1) {
        try targetMon := MonitorGetPrimary()
        catch
            targetMon := 1
    }
    g_ClipAngelNativeOpenTargetMon := targetMon
    g_ClipAngelNativeOpenRetryArmed := false
    SetTimer(ClipAngel_ApplyForegroundMaximizeOnce, -CLIPANGEL_NATIVE_OPEN_SETTLE_MS)
}

; One layout+activate if gate fails; one short retry if still not ok.
ClipAngel_ApplyForegroundMaximizeOnce(*) {
    global g_ClipAngelNativeOpenTargetMon, g_ClipAngelNativeOpenRetryArmed
    targetMon := g_ClipAngelNativeOpenTargetMon
    hwnd := ClipAngel_MainHwnd()
    if !hwnd
        return
    if ClipAngel_IsForegroundLayoutOk(hwnd, targetMon)
        return
    try {
        if (WinGetMinMax("ahk_id " hwnd) = -1 || DllCall("IsIconic", "ptr", hwnd))
            WinRestore("ahk_id " hwnd)
    } catch {
    }
    try WinShow("ahk_id " hwnd)
    catch {
    }
    ClipAngel_ApplyLayoutOnMonitor(hwnd, targetMon)
    if ClipAngel_IsForegroundLayoutOk(hwnd, targetMon)
        return
    if (!g_ClipAngelNativeOpenRetryArmed) {
        g_ClipAngelNativeOpenRetryArmed := true
        SetTimer(ClipAngel_ApplyForegroundMaximizeOnce, -CLIPANGEL_NATIVE_OPEN_RETRY_MS)
    }
}

; =============================================================================
; Auto-minimize when Clip Angel loses focus to another process
; =============================================================================

CLIPANGEL_AUTO_MINIMIZE_DEBOUNCE_MS := 250
global g_ClipAngelFgHook := 0
global g_ClipAngelFgHookCb := 0
global g_ClipAngelLastFgHwnd := 0
; Default false here (Utils hosts); Shift keys hotif_clipangel toggles it for the filter selector.
global g_ClipAngelFilterSelectorActive := false

ClipAngel_InitAutoMinimizeOnDeactivate() {
    global g_ClipAngelFgHook, g_ClipAngelFgHookCb
    if (g_ClipAngelFgHook)
        return
    g_ClipAngelFgHookCb := CallbackCreate(ClipAngel_ForegroundHookProc, "F", 7)
    g_ClipAngelFgHook := DllCall("user32\SetWinEventHook",
        "UInt", 0x0003,  ; EVENT_SYSTEM_FOREGROUND
        "UInt", 0x0003,
        "Ptr", 0,
        "Ptr", g_ClipAngelFgHookCb,
        "UInt", 0,
        "UInt", 0,
        "UInt", 0,
        "Ptr")
    if (!g_ClipAngelFgHook) {
        try CallbackFree(g_ClipAngelFgHookCb)
        catch {
        }
        g_ClipAngelFgHookCb := 0
    }
}

; Hook thread: store hwnd and bounce to main thread.
ClipAngel_ForegroundHookProc(hHook, event, hwnd, idObject, idChild, idEventThread, dwmsEventTime) {
    global g_ClipAngelLastFgHwnd
    g_ClipAngelLastFgHwnd := hwnd
    ; -0 is treated as period 0 (disable timer). Must use -1 for next-tick once.
    SetTimer(ClipAngel_OnForegroundChanged, -1)
}

ClipAngel_ShouldSkipAutoMinimize() {
    global g_ClipAngelAutomationBusy, g_ClipAngelFilterSelectorActive, g_ClipAngelFavoriteSuppressActive
    if (g_ClipAngelAutomationBusy)
        return true
    if (g_ClipAngelFavoriteSuppressActive)
        return true
    if (g_ClipAngelFilterSelectorActive)
        return true
    ; Call by name so hosts that #include activate before constant_paste (or omit it) do not #Warn.
    try {
        fnName := "ClipAngel_ConstantPaste_IsDelimiterPromptActive"
        if %fnName%()
            return true
    } catch {
    }
    try {
        fnName := "ClipAngel_ConstantPaste_IsActive"
        if %fnName%()
            return true
    } catch {
    }
    return false
}

ClipAngel_FgIsClipAngelExe(hwnd) {
    if (!hwnd)
        return false
    try return StrLower(WinGetProcessName("ahk_id " hwnd)) = "clipangel.exe"
    catch
        return false
}

; Schedule or cancel debounced minimize when foreground changes.
ClipAngel_OnForegroundChanged(*) {
    if (ClipAngel_ShouldSkipAutoMinimize()) {
        SetTimer(ClipAngel_AutoMinimizeTick, 0)
        return
    }
    ca := ClipAngel_MainHwnd()
    if (!ca || !ClipAngel_IsWindowShown(ca)) {
        SetTimer(ClipAngel_AutoMinimizeTick, 0)
        return
    }
    global g_ClipAngelLastFgHwnd
    fg := g_ClipAngelLastFgHwnd
    if (!fg)
        fg := WinExist("A")
    if (fg = ca || ClipAngel_FgIsClipAngelExe(fg)) {
        SetTimer(ClipAngel_AutoMinimizeTick, 0)
        return
    }
    SetTimer(ClipAngel_AutoMinimizeTick, -CLIPANGEL_AUTO_MINIMIZE_DEBOUNCE_MS)
}

ClipAngel_AutoMinimizeTick(*) {
    if (ClipAngel_ShouldSkipAutoMinimize())
        return
    ca := ClipAngel_MainHwnd()
    if (!ca || !ClipAngel_IsWindowShown(ca))
        return
    fg := WinExist("A")
    if (fg = ca || ClipAngel_FgIsClipAngelExe(fg))
        return
    ClipAngel_HideWindow(ca)
}

; Escape while Clip Angel is focused: close filter overlay if any, then minimize.
; Cleanup lives in Shift keys\hotif_clipangel.ahk — call by name so Utils hosts do not #Warn.
ClipAngel_EscapeMinimize() {
    global g_ClipAngelFilterSelectorActive
    try {
        fnName := "ClipAngel_ConstantPaste_IsDelimiterPromptActive"
        if %fnName%()
            return
    } catch {
    }
    try {
        fnName := "ClipAngel_ConstantPaste_IsActive"
        if %fnName%()
            return
    } catch {
    }
    if (g_ClipAngelFilterSelectorActive) {
        fnName := "CleanupClipAngelFilterSelector"
        try %fnName%()
        catch {
        }
    }
    ClipAngel_CloseAndRestoreFocus(0)
}

; Resolve ClipAngel.exe for hard restart.
; Personal portable: C:\Users\eduev\Documents\ClipAngel\ClipAngel 2.13\…
; Work portable: OneDrive Bosch Documents\ClipAngel\ClipAngel.exe (also legacy
; Handy-style under fie7ca\Documents\ClipAngel 2.23 — docs/reference/clip-angel.txt).
; A_MyDocuments may be OneDrive — prefer known exact paths, then scan roots.
ClipAngel_ResolveExePath() {
    global IS_WORK_ENVIRONMENT
    try {
        hwnd := ClipAngel_MainHwnd()
        if (hwnd) {
            path := WinGetProcessPath("ahk_id " hwnd)
            if (path != "" && FileExist(path))
                return path
        }
    } catch {
    }
    if ProcessExist("ClipAngel.exe") {
        try {
            path := ProcessGetPath("ClipAngel.exe")
            if (path != "" && FileExist(path))
                return path
        } catch {
        }
    }

    ; Optional env.ahk helper when present — call by name so hosts without it do not #Warn.
    try {
        fnName := "GetClipAngelExePath"
        envPath := %fnName%()
        if (envPath != "" && FileExist(envPath))
            return envPath
    } catch {
    }

    personalExact := "C:\Users\eduev\Documents\ClipAngel\ClipAngel 2.13\ClipAngel.exe"
    workExact :=
        "C:\Users\fie7ca\OneDrive - Bosch Group\01 - Geral\16 - Others backups\Documents\ClipAngel\ClipAngel.exe"
    workLegacyExact := "C:\Users\fie7ca\Documents\ClipAngel\ClipAngel 2.23\ClipAngel.exe"
    isWork := IsSet(IS_WORK_ENVIRONMENT) && IS_WORK_ENVIRONMENT
    preferred := isWork ? workExact : personalExact
    other := isWork ? personalExact : workExact
    if FileExist(preferred)
        return preferred
    if FileExist(workLegacyExact)
        return workLegacyExact
    if FileExist(other)
        return other

    roots := []
    if (isWork) {
        roots.Push("C:\Users\fie7ca\OneDrive - Bosch Group\01 - Geral\16 - Others backups\Documents")
        roots.Push("C:\Users\fie7ca\Documents")
        userDocs := EnvGet("USERPROFILE") "\Documents"
        if (userDocs != "" && userDocs != "C:\Users\fie7ca\Documents")
            roots.Push(userDocs)
        if (A_MyDocuments != "" && A_MyDocuments != "C:\Users\fie7ca\Documents")
            roots.Push(A_MyDocuments)
    } else {
        roots.Push("C:\Users\eduev\Documents")
        userDocs := EnvGet("USERPROFILE") "\Documents"
        if (userDocs != "" && userDocs != "C:\Users\eduev\Documents")
            roots.Push(userDocs)
        if (A_MyDocuments != "" && A_MyDocuments != "C:\Users\eduev\Documents")
            roots.Push(A_MyDocuments)
    }
    for root in roots {
        for cand in [
            root "\ClipAngel\ClipAngel.exe",
            root "\ClipAngel\ClipAngel 2.23\ClipAngel.exe",
            root "\ClipAngel\ClipAngel 2.13\ClipAngel.exe"
        ] {
            if FileExist(cand)
                return cand
        }
        loop files root "\ClipAngel\*\ClipAngel.exe", "F" {
            return A_LoopFileFullPath
        }
    }
    return ""
}

; Force-kill ClipAngel.exe (not minimize). Returns true when no process remains.
ClipAngel_KillProcess(timeoutMs := 4000) {
    deadline := A_TickCount + timeoutMs
    while (ProcessExist("ClipAngel.exe") && A_TickCount < deadline) {
        try ProcessClose("ClipAngel.exe")
        catch {
        }
        Sleep 80
    }
    return !ProcessExist("ClipAngel.exe")
}

; SendMessageTimeout WM_NULL + IsHungAppWindow; false when unresponsive within timeout.
CLIPANGEL_RESPONSIVE_PROBE_MS := 300
CLIPANGEL_RESTART_HWND_WAIT_MS := 15000
CLIPANGEL_RESTART_RESPONSIVE_WAIT_MS := 20000
CLIPANGEL_RESTART_POST_ACTIVATE_RESPONSIVE_MS := 10000
CLIPANGEL_RESTART_PROCESS_WAIT_MS := 12000
CLIPANGEL_RESTART_LIST_WAIT_MS := 8000
CLIPANGEL_RESTART_QUALITY_POLL_MS := 100
CLIPANGEL_RESTART_MAX_ATTEMPTS := 2

ClipAngel_IsWindowResponsive(hwnd, timeoutMs := 0) {
    if !(hwnd is Integer) || hwnd <= 0 || !WinExist("ahk_id " hwnd)
        return false
    if (!timeoutMs)
        timeoutMs := CLIPANGEL_RESPONSIVE_PROBE_MS
    try {
        if DllCall("IsHungAppWindow", "Ptr", hwnd)
            return false
    } catch {
    }
    result := 0
    ; SMTO_ABORTIFHUNG = 0x0002
    ok := DllCall("SendMessageTimeout", "Ptr", hwnd, "UInt", 0, "Ptr", 0, "Ptr", 0, "UInt", 0x0002, "UInt",
        timeoutMs, "Ptr*", &result)
    return ok != 0
}

; Poll until Clip Angel's main window answers WM_NULL (or timeout).
ClipAngel_WaitUntilResponsive(hwnd, timeoutMs := 0) {
    if (!timeoutMs)
        timeoutMs := CLIPANGEL_RESTART_RESPONSIVE_WAIT_MS
    if !(hwnd is Integer) || hwnd <= 0
        return false
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        cur := ClipAngel_MainHwnd()
        if (cur)
            hwnd := cur
        if ClipAngel_IsWindowResponsive(hwnd)
            return true
        Sleep CLIPANGEL_RESTART_QUALITY_POLL_MS
    }
    cur := ClipAngel_MainHwnd()
    if (cur)
        hwnd := cur
    return ClipAngel_IsWindowResponsive(hwnd)
}

; True when hwnd is the ClipAngel WinForms main shell (not a random dialog).
ClipAngel_IsMainWindowIdentity(hwnd) {
    if !(hwnd is Integer) || hwnd <= 0 || !WinExist("ahk_id " hwnd)
        return false
    try {
        if (StrLower(WinGetProcessName("ahk_id " hwnd)) != "clipangel.exe")
            return false
    } catch {
        return false
    }
    try {
        title := WinGetTitle("ahk_id " hwnd)
        if !InStr(title, "ClipAngel")
            return false
    } catch {
        return false
    }
    try {
        cls := WinGetClass("ahk_id " hwnd)
        if !InStr(cls, "WindowsForms10")
            return false
    } catch {
        return false
    }
    return true
}

; New process after kill: ClipAngel.exe exists, PID != oldPid, optional path match.
ClipAngel_WaitForNewProcess(oldPid, expectedExe := "", timeoutMs := 0) {
    if (!timeoutMs)
        timeoutMs := CLIPANGEL_RESTART_PROCESS_WAIT_MS
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        pid := ProcessExist("ClipAngel.exe")
        if (pid && pid != oldPid) {
            if (expectedExe = "")
                return pid
            try path := ProcessGetPath("ClipAngel.exe")
            catch
                path := ""
            if (path = "" || StrLower(path) = StrLower(expectedExe))
                return pid
        }
        Sleep CLIPANGEL_RESTART_QUALITY_POLL_MS
    }
    return 0
}

; Wait for a main hwnd owned by newPid (when known) that passes identity checks.
ClipAngel_WaitForRestartHwnd(newPid := 0, timeoutMs := 0) {
    if (!timeoutMs)
        timeoutMs := CLIPANGEL_RESTART_HWND_WAIT_MS
    startTick := A_TickCount
    deadline := startTick + timeoutMs
    lastWakeTick := 0
    while (A_TickCount < deadline) {
        if (ClipAngel_DismissStartupHotkeyDialogs(newPid) > 0)
            Sleep 60
        hwnd := ClipAngel_MainHwnd()
        if (hwnd && ClipAngel_IsMainWindowIdentity(hwnd)) {
            if (!newPid)
                return hwnd
            try {
                if (WinGetPID("ahk_id " hwnd) = newPid)
                    return hwnd
            } catch {
            }
        }
        ; ClipAngel can start minimized to tray (HWND = 0); nudge with Alt+P after startup settle.
        elapsed := A_TickCount - startTick
        if (elapsed >= 700 && (lastWakeTick = 0 || (A_TickCount - lastWakeTick) >= 1200)) {
            if (!newPid || ProcessExist(newPid)) {
                lastWakeTick := A_TickCount
                if !ClipAngel_PostHotkey("p", "a") {
                    priorSendLevel := A_SendLevel
                    SendLevel 0
                    ClipAngel_ReleaseChordModifiersForSend()
                    SendInput("!p")
                    SendLevel priorSendLevel
                }
            }
        }
        Sleep CLIPANGEL_RESTART_QUALITY_POLL_MS
    }
    hwnd := ClipAngel_MainHwnd()
    if (hwnd && ClipAngel_IsMainWindowIdentity(hwnd)) {
        if (!newPid)
            return hwnd
        try {
            if (WinGetPID("ahk_id " hwnd) = newPid)
                return hwnd
        } catch {
        }
    }
    return 0
}

; Layout usable after restart: shown, not tiny bar, preferably maximized + foreground.
ClipAngel_IsRestartLayoutOk(hwnd) {
    if !hwnd || !ClipAngel_IsWindowShown(hwnd)
        return false
    if ClipAngel_NeedsLayoutCorrection(hwnd)
        return false
    if !WinActive("ahk_id " hwnd)
        return false
    try {
        if (WinGetMinMax("ahk_id " hwnd) != 1)
            return false
    } catch {
        return false
    }
    return true
}

; UIA gates: dataGridView present, then Row 0 / list ready. Never call if hung.
ClipAngel_WaitForRestartUiReady(hwnd, timeoutMs := 0) {
    if (!timeoutMs)
        timeoutMs := CLIPANGEL_RESTART_LIST_WAIT_MS
    if !(hwnd is Integer) || hwnd <= 0
        return false
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        cur := ClipAngel_MainHwnd()
        if (cur)
            hwnd := cur
        if !ClipAngel_IsWindowResponsive(hwnd)
            return false
        try {
            if ClipAngel_UiaGetDataGrid(hwnd) {
                if ClipAngel_IsListReady(&hwnd) {
                    try ClipAngel_UiaEnsureRow0Selected(hwnd, false)
                    catch {
                    }
                    return true
                }
            }
        } catch {
        }
        Sleep CLIPANGEL_UIA_POLL_MS
    }
    if !ClipAngel_IsWindowResponsive(hwnd)
        return false
    try return ClipAngel_IsListReady()
    catch
        return false
}

; Full post-restart quality gate. Sets failReason on false.
ClipAngel_RestartQualityOk(hwnd, expectedExe, oldPid, &failReason, backgroundVerify := false) {
    failReason := ""
    if !(hwnd is Integer) || hwnd <= 0 || !WinExist("ahk_id " hwnd) {
        failReason := "no window"
        return false
    }
    if !ClipAngel_IsMainWindowIdentity(hwnd) {
        failReason := "window identity mismatch"
        return false
    }
    try pid := WinGetPID("ahk_id " hwnd)
    catch
        pid := 0
    if (oldPid && pid && pid = oldPid) {
        failReason := "same PID as before kill"
        return false
    }
    if (expectedExe != "") {
        try path := WinGetProcessPath("ahk_id " hwnd)
        catch
            path := ""
        if (path != "" && StrLower(path) != StrLower(expectedExe)) {
            failReason := "exe path mismatch"
            return false
        }
    }
    if !ClipAngel_IsWindowResponsive(hwnd) {
        failReason := "not responding"
        return false
    }
    if (!backgroundVerify && !ClipAngel_IsRestartLayoutOk(hwnd)) {
        failReason := "layout/foreground not ready"
        return false
    }
    try {
        if !ClipAngel_UiaGetDataGrid(hwnd) {
            failReason := "clip list (dataGridView) missing"
            return false
        }
        if !ClipAngel_IsListReady(&hwnd) {
            failReason := "clip list Row 0 not ready"
            return false
        }
    } catch {
        failReason := "UIA list check failed"
        return false
    }
    return true
}

; One kill→launch→verify cycle. Returns true on success; failReason on false.
ClipAngel_RestartAttempt(exePath, oldPid, attempt, &failReason, backgroundVerify := false, keepSuppressed := false,
    hwndWaitMs := 0, listWaitMs := 0) {
    failReason := ""
    baselinePid := oldPid
    if (hwndWaitMs <= 0)
        hwndWaitMs := CLIPANGEL_RESTART_HWND_WAIT_MS
    if (listWaitMs <= 0)
        listWaitMs := CLIPANGEL_RESTART_LIST_WAIT_MS
    label := (CLIPANGEL_RESTART_MAX_ATTEMPTS > 1)
        ? " (attempt " attempt "/" CLIPANGEL_RESTART_MAX_ATTEMPTS ")"
        : ""

    if ProcessExist("ClipAngel.exe") {
        StandardLoadingBar_Update("⏳ Killing Clip Angel process…" label)
        if !ClipAngel_KillProcess(5000) {
            failReason := "could not kill ClipAngel.exe"
            return false
        }
        Sleep 250
    }

    StandardLoadingBar_Update("⏳ Opening Clip Angel…" label)
    SplitPath(exePath, , &exeDir)
    try Run('"' exePath '"', exeDir)
    catch as e {
        failReason := "failed to start: " e.Message
        return false
    }

    StandardLoadingBar_Update("⏳ Waiting for new Clip Angel process…" label)
    newPid := ClipAngel_WaitForNewProcess(baselinePid, exePath, CLIPANGEL_RESTART_PROCESS_WAIT_MS)
    if (!newPid) {
        failReason := "process did not start (or wrong exe path)"
        return false
    }

    StandardLoadingBar_Update("⏳ Waiting for Clip Angel window…" label)
    hwnd := ClipAngel_WaitForRestartHwnd(newPid, hwndWaitMs)
    if (!hwnd) {
        failReason := "main window did not appear"
        return false
    }

    StandardLoadingBar_Update("⏳ Waiting for Clip Angel to respond…" label)
    if !ClipAngel_WaitUntilResponsive(hwnd, CLIPANGEL_RESTART_RESPONSIVE_WAIT_MS) {
        failReason := "started but is not responding"
        return false
    }

    suppressStarted := false
    try {
        if (backgroundVerify) {
            StandardLoadingBar_Update("⏳ Preparing Clip Angel in background…" label)
            suppressStarted := ClipAngel_BeginFavoriteSuppress(hwnd)
            if (!suppressStarted) {
                failReason := "background suppress failed"
                return false
            }
        } else {
            StandardLoadingBar_Update("⏳ Activating Clip Angel…" label)
            ; silent + skipRow0: keep StandardLoadingBar as the only indicator; UIA list gate comes next.
            if !ActivateClipAngelWithFocusCorrection(true, 0, true, true) {
                failReason := "window not found after start"
                return false
            }
        }

        StandardLoadingBar_Update("⏳ Verifying Clip Angel is responsive…" label)
        hwnd := ClipAngel_MainHwnd()
        if !hwnd || !ClipAngel_WaitUntilResponsive(hwnd, CLIPANGEL_RESTART_POST_ACTIVATE_RESPONSIVE_MS) {
            failReason := "not responding after activate"
            return false
        }

        StandardLoadingBar_Update("⏳ Waiting for Clip Angel list…" label)
        if !ClipAngel_WaitForRestartUiReady(hwnd, listWaitMs) {
            failReason := "clip list not ready"
            return false
        }

        StandardLoadingBar_Update("⏳ Confirming restart quality…" label)
        hwnd := ClipAngel_MainHwnd()
        if !ClipAngel_RestartQualityOk(hwnd, exePath, baselinePid, &failReason, backgroundVerify) {
            if (failReason = "")
                failReason := "quality gate failed"
            return false
        }
        return true
    } finally {
        if (suppressStarted && !keepSuppressed)
            ClipAngel_EndFavoriteSuppress(hwnd)
    }
}

; Auto-recover a stuck/windowless ClipAngel.exe inside an already-locked automation flow
; (does not touch g_ClipAngelAutomationBusy so the caller's lock stays intact).
ClipAngel_RecoverProcessForAutomation(backgroundVerify := true) {
    oldPid := ProcessExist("ClipAngel.exe")
    exePath := ClipAngel_ResolveExePath()
    if (exePath = "" || !FileExist(exePath))
        return false
    StandardLoadingBar_Show("⏳ Recovering Clip Angel…", BANNER_ACCENT_INTERMEDIATE, { passive: true, noActivate: true })
    failReason := ""
    try {
        return ClipAngel_RestartAttempt(exePath, oldPid, 1, &failReason, backgroundVerify, backgroundVerify, 6000, 4000)
    } catch {
        return false
    } finally {
        try StandardLoadingBar_Hide(0)
        catch {
        }
    }
}

; Macros [R] — hard restart when Clip Angel is stuck.
; Loading Indication stays up for the whole path; up to 2 attempts with multi-gate verify.
; Returns true on success (also used by Act.ahk / ClipAngel_EnsureRunning).
ClipAngel_RestartHard(*) {
    global g_ClipAngelAutomationBusy
    StandardLoadingBar_Show("⏳ Restarting Clip Angel…", BANNER_ACCENT_INTERMEDIATE, { passive: true })
    g_ClipAngelAutomationBusy := true
    failReason := ""
    try {
        oldPid := ProcessExist("ClipAngel.exe")
        exePath := ClipAngel_ResolveExePath()
        ; Prefer path from still-running process before kill when resolver is empty.
        if ((exePath = "" || !FileExist(exePath)) && oldPid) {
            try {
                hwndLive := ClipAngel_MainHwnd()
                if (hwndLive)
                    exePath := WinGetProcessPath("ahk_id " hwndLive)
            } catch {
            }
        }
        if (exePath = "" || !FileExist(exePath)) {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ ClipAngel.exe not found", 2500, BANNER_ACCENT_ERROR)
            return false
        }

        loop CLIPANGEL_RESTART_MAX_ATTEMPTS {
            if ClipAngel_RestartAttempt(exePath, oldPid, A_Index, &failReason, true) {
                StandardLoadingBar_Hide(0)
                ShowCenteredOverlay_Utils("✅ Clip Angel restarted", 1500, BANNER_ACCENT_SUCCESS)
                try ScriptSoundPlay(A_ScriptDir "\assets\sounds\clipangel-restart-success.wav")
                catch {
                }
                return true
            }
            if (A_Index < CLIPANGEL_RESTART_MAX_ATTEMPTS) {
                StandardLoadingBar_Update("⏳ Restart incomplete — retrying…")
                Sleep 400
                oldPid := ProcessExist("ClipAngel.exe")
            }
        }

        StandardLoadingBar_Hide(0)
        msg := "❌ Clip Angel restart failed"
        if (failReason != "")
            msg .= ": " failReason
        ShowCenteredOverlay_Utils(msg, 3200, BANNER_ACCENT_ERROR)
        return false
    } catch as e {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        ShowCenteredOverlay_Utils("❌ Clip Angel restart failed: " . e.Message, 2800, BANNER_ACCENT_ERROR)
        return false
    } finally {
        g_ClipAngelAutomationBusy := false
    }
}

; Soft verify in background when already healthy; otherwise kill→launch→verify in background.
; Used by Act.ahk after launching hosts so Clip Angel is guaranteed up without stealing focus.
ClipAngel_EnsureRunning(*) {
    global g_ClipAngelAutomationBusy
    ClipAngel_DismissStartupHotkeyDialogs()
    hwnd := ClipAngel_MainHwnd()
    if (hwnd && ClipAngel_IsMainWindowIdentity(hwnd) && ClipAngel_IsWindowResponsive(hwnd)) {
        StandardLoadingBar_Show("⏳ Ensuring Clip Angel…", BANNER_ACCENT_INTERMEDIATE, { passive: true })
        g_ClipAngelAutomationBusy := true
        suppressStarted := false
        try {
            suppressStarted := ClipAngel_BeginFavoriteSuppress(hwnd)
            if (suppressStarted) {
                hwnd := ClipAngel_MainHwnd()
                if (hwnd && ClipAngel_WaitUntilResponsive(hwnd, CLIPANGEL_RESTART_POST_ACTIVATE_RESPONSIVE_MS)
                && ClipAngel_WaitForRestartUiReady(hwnd, CLIPANGEL_RESTART_LIST_WAIT_MS)) {
                    failReason := ""
                    exePath := ClipAngel_ResolveExePath()
                    if ClipAngel_RestartQualityOk(hwnd, exePath, 0, &failReason, true) {
                        StandardLoadingBar_Hide(0)
                        ShowCenteredOverlay_Utils("✅ Clip Angel ready", 1200, BANNER_ACCENT_SUCCESS)
                        return true
                    }
                }
            }
        } catch {
        } finally {
            if (suppressStarted)
                ClipAngel_EndFavoriteSuppress(hwnd)
            g_ClipAngelAutomationBusy := false
        }
        try StandardLoadingBar_Hide(0)
        catch {
        }
        ; Process present but not usable — fall through to hard restart.
    }
    return ClipAngel_RestartHard()
}

RegisterMacro(ClipAngel_RestartHard, "🔄 Restart Clip Angel (kill process)", "r")