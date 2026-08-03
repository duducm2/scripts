; =============================================================================
; WindowManagement module: window_tools.ahk
; Win+Alt+Shift+W window tools menu
; Extracted verbatim from WindowManagement.ahk; loaded via #include into the
; WindowManagement.ahk process, which remains the entry point / source of truth.
; =============================================================================

; =============================================================================
; Win+Alt+Shift+W — window tools menu (Interactive Input) + minimized list GUI
; =============================================================================

; F11 fullscreen helpers: WM_WindowIsF11Fullscreen*, WM_ExitF11FullscreenForHwnd, WM_EnterF11FullscreenForHwnd (Utils.ahk)

WM_ExitF11FullscreenAllWindows() {
    foreBefore := 0
    try foreBefore := WinGetID("A")
    WMAutomation_SuppressCursorCentering("exit_f11_fullscreen", 5000)
    seen := Map()
    candidates := []
    scanned := 0
    loop MonitorGetCount() {
        for win in GetVisibleWindowsOnMonitor(A_Index, true) {
            if (seen.Has(win.hwnd))
                continue
            seen[win.hwnd] := true
            scanned++
            if (WM_WindowIsF11Fullscreen(win.hwnd))
                candidates.Push(win.hwnd)
        }
    }
    if (candidates.Length > 0)
        StandardLoadingBar_Update("🔄 Exiting F11 fullscreen on " candidates.Length " window(s)...",
            BANNER_ACCENT_INTERMEDIATE)
    exited := 0
    for i, hwnd in candidates {
        if (candidates.Length > 1)
            StandardLoadingBar_Update("🔄 Exiting F11 fullscreen (" i "/" candidates.Length ")...",
                BANNER_ACCENT_INTERMEDIATE)
        if (WM_ExitF11FullscreenForHwnd(hwnd))
            exited++
    }
    if (foreBefore && WinExist("ahk_id " foreBefore)) {
        try WinActivate("ahk_id " foreBefore)
        catch {
        }
    }
    WMAutomation_ClearCursorSuppression("exit_f11_fullscreen")
    msg := (exited = 0)
        ? ("ℹ️ No F11 fullscreen windows found (" scanned " visible checked)")
        : ((exited = 1) ? "✅ Exited F11 fullscreen on 1 window" : "✅ Exited F11 fullscreen on " exited " windows")
    return { ok: true, exited: exited, scanned: scanned, message: msg }
}

WM_WindowTools_OnMaximizeLonely(*) {
    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Hide(0)
    Sleep 50
    WM_MaximizeLonelyVisibleOnAllMonitors()
}

WM_WindowTools_OnTileBackground(*) {
    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Hide(0)
    Sleep 50
    WM_BackgroundTitleExcludes_Init()
    foreBeforeScan := 0
    try foreBeforeScan := WinGetID("A")
    global g_WM_BackgroundScanBannerTick
    g_WM_BackgroundScanBannerTick := A_TickCount
    ; AutoSlot ON: fill free slots only — never move windows already in slots.
    autoslotOn := false
    try autoslotOn := !!AutoSlot_IsEnabled()
    catch
        autoslotOn := false
    if (autoslotOn) {
        ; AutoSlot rearrange mode: INFO accent throughout (not yellow/green).
        StandardLoadingBar_Show("Filling free AutoSlot slots...", BANNER_ACCENT_INFO, { passive: false,
            centerOnHwnd: 0 })
        try {
            result := AutoSlot_RunTileBackground()
            if (!result.ok) {
                if (result.HasProp("noBackground") && result.noBackground)
                    WM_PresentNoBackgroundWindowsEmpty(result.message)
                else {
                    StandardLoadingBar_Update(result.message, BANNER_ACCENT_INFO)
                    StandardLoadingBar_Hide(4500)
                }
                return
            }
            StandardLoadingBar_Update(result.message, BANNER_ACCENT_INFO)
            StandardLoadingBar_Hide(2000)
        } catch as err {
            StandardLoadingBar_Update("Slot fill failed: " err.Message, BANNER_ACCENT_ERROR)
            StandardLoadingBar_Hide(4000)
        }
        return
    }
    StandardLoadingBar_Show("Scanning hidden windows...", BANNER_ACCENT_INTERMEDIATE, { passive: false,
        centerOnHwnd: 0 })
    try {
        ; AutoSlot OFF: legacy tile, but skip already-slotted windows and cap 2/mon.
        result := WM_TileBackgroundWindowsPerMonitor(2, foreBeforeScan)
        if (!result.ok) {
            if (result.HasProp("noBackground") && result.noBackground)
                WM_PresentNoBackgroundWindowsEmpty(result.message)
            else {
                StandardLoadingBar_Update(result.message, BANNER_ACCENT_INFO)
                StandardLoadingBar_Hide(4500)
            }
            return
        }
        StandardLoadingBar_Update(result.message, BANNER_ACCENT_SUCCESS)
        StandardLoadingBar_Hide(2000)
    } catch as err {
        StandardLoadingBar_Update("Background tile failed: " err.Message, BANNER_ACCENT_ERROR)
        StandardLoadingBar_Hide(4000)
    }
}

WM_WindowTools_OnShowMinimizedList(*) {
    global g_WM_WindowToolsShowListLock, g_WM_WindowToolsShowListLastTick
    if (g_WM_WindowToolsShowListLock)
        return
    if (A_TickCount - g_WM_WindowToolsShowListLastTick < 400)
        return
    g_WM_WindowToolsShowListLock := true
    g_WM_WindowToolsShowListLastTick := A_TickCount
    try {
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        Sleep 50
        WM_BackgroundTitleExcludes_Init()
        foreBeforeScan := 0
        try foreBeforeScan := WinGetID("A")
        global g_WM_MinimizedListCollectForeHwnd
        g_WM_MinimizedListCollectForeHwnd := foreBeforeScan
        global g_WM_BackgroundScanBannerTick
        g_WM_BackgroundScanBannerTick := A_TickCount
        StandardLoadingBar_Show("⏳ Scanning hidden windows (z-order)...", BANNER_ACCENT_INTERMEDIATE, { passive: false,
            centerOnHwnd: 0 })
        try {
            rows := WM_CollectBackgroundWindows(foreBeforeScan)
            if (rows.Length = 0) {
                WM_PresentNoBackgroundWindowsEmpty()
                return
            }
            StandardLoadingBar_Update("✅ Found " . rows.Length . " hidden window(s)", BANNER_ACCENT_SUCCESS)
            StandardLoadingBar_Hide(900)
            WM_ShowMinimizedBackgroundList(rows)
        } catch as err {
            WM_PlayNoWindowSound()
            StandardLoadingBar_Update("❌ Background scan failed: " . err.Message, BANNER_ACCENT_ERROR)
            StandardLoadingBar_Hide(4000)
        }
    } finally {
        g_WM_WindowToolsShowListLock := false
    }
}

WM_WindowTools_OnExitF11Fullscreen(*) {
    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Hide(0)
    Sleep 50
    StandardLoadingBar_Show("⏳ Scanning for F11 fullscreen windows...", BANNER_ACCENT_INTERMEDIATE, { passive: false,
        centerOnHwnd: 0 })
    try {
        result := WM_ExitF11FullscreenAllWindows()
        accent := (result.exited > 0) ? BANNER_ACCENT_SUCCESS : BANNER_ACCENT_INFO
        StandardLoadingBar_Update(result.message, accent)
        StandardLoadingBar_Hide(result.exited > 0 ? 2000 : 4500)
    } catch as err {
        StandardLoadingBar_Update("❌ Exit F11 fullscreen failed: " err.Message, BANNER_ACCENT_ERROR)
        StandardLoadingBar_Hide(4000)
    }
}

WM_WindowTools_OnCancel(*) {
    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Hide(0)
}

WM_WindowTools_OnToggleAutoSlot(*) {
    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Hide(0)
    Sleep 50
    newState := false
    try newState := AutoSlot_SetEnabled(!AutoSlot_IsEnabled())
    catch {
        try ShowCenteredOverlay_Utils("❌ AutoSlot toggle unavailable", 2500, BANNER_ACCENT_ERROR)
        catch {
        }
        return
    }
    label := newState ? "ON" : "OFF"
    accent := newState ? BANNER_ACCENT_SUCCESS : BANNER_ACCENT_INFO
    try ShowCenteredOverlay_Utils("AutoSlot " label, 1400, accent)
    catch {
    }
    Sleep 200
    WM_WindowTools_ShowMenu()
}

WM_WindowTools_ShowMenu() {
    global g_WM_MinimizedListActive
    if (g_WM_MinimizedListActive)
        WM_MinimizedList_Cleanup()
    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Hide(0)
    Sleep 50
    autoSlotOn := true
    try autoSlotOn := AutoSlot_IsEnabled()
    catch
        autoSlotOn := true
    autoSlotLabel := autoSlotOn ? "ON" : "OFF"
    keyCallbacks := Map(
        "1", WM_WindowTools_OnMaximizeLonely,
        "2", WM_WindowTools_OnShowMinimizedList,
        "3", WM_WindowTools_OnTileBackground,
        "4", WM_WindowTools_OnExitF11Fullscreen,
        "5", WM_WindowTools_OnToggleAutoSlot,
        "Escape", WM_WindowTools_OnCancel)
    StandardLoadingBar_ShowWithKeys(
        "❓ Window tools — AutoSlot " autoSlotLabel " (8s)",
        keyCallbacks,
        8000,
        0,
        "",
        BANNER_ACCENT_INTERMEDIATE,
        480,
        17,
        "",
        false,
        "[1] Maximize lone (CAW+Z)  [2] Hidden list (CAW+U)  [3] Fill free slots (CAW+6; slotted stay)  [4] Exit F11 fullscreen (CAW+P)  [5] AutoSlot: " autoSlotLabel "  [Esc] Cancel",
        true,
        false,
        false)
}

WM_TruncateTitleForList(s, maxLen := 80) {
    if (StrLen(s) <= maxLen)
        return s
    return SubStr(s, 1, maxLen - 1) . "…"
}

WM_SortBackgroundRows(&rows) {
    n := rows.Length
    if (n < 2)
        return
    loop n - 1 {
        loop n - A_Index {
            j := A_Index
            a := rows[j]
            b := rows[j + 1]
            swap := false
            if (StrCompare(a.title, b.title, true) > 0)
                swap := true
            if (swap) {
                tmp := rows[j]
                rows[j] := rows[j + 1]
                rows[j + 1] := tmp
            }
        }
    }
}

WM_ArrJoin(arr, sep := "`n") {
    out := ""
    for item in arr
        out .= (out = "" ? "" : sep) . item
    return out
}

WM_BackgroundTitleExcludes_IniPath() {
    return A_ScriptDir "\assets\data\wm_background_excludes.ini"
}

WM_BackgroundTitleExcludes_Register(&list, &seen, needle) {
    n := Trim(needle)
    if (n = "")
        return
    key := StrLower(n)
    if (seen.Has(key))
        return
    seen[key] := true
    list.Push(n)
}

WM_BackgroundTitleExcludes_WriteList(list) {
    path := WM_BackgroundTitleExcludes_IniPath()
    try DirCreate(A_ScriptDir "\data")
    lines := ["[Excludes]", "; One entry per line: title substring, or exe|title"]
    for n in list
        lines.Push(n)
    try FileDelete(path)
    FileAppend(WM_ArrJoin(lines, "`n") "`n", path, "UTF-8")
}

WM_BackgroundTitleExcludes_ParseDiskEntries(raw) {
    entries := []
    seen := Map()
    for line in StrSplit(raw, "`n", "`r") {
        line := Trim(line)
        if (line = "" || SubStr(line, 1, 1) = ";" || line = "[Excludes]")
            continue
        if (SubStr(line, 1, 1) = "[")
            continue
        if RegExMatch(line, "i)^TitleContains=(.*)", &m)
            line := Trim(m[1])
        if (line = "")
            continue
        ; Legacy single-line pipe-separated values (no exe|title needles).
        if (InStr(line, "|") && !RegExMatch(line, "\.exe\|")) {
            for part in StrSplit(line, "|")
                WM_BackgroundTitleExcludes_Register(&entries, &seen, part)
            continue
        }
        WM_BackgroundTitleExcludes_Register(&entries, &seen, line)
    }
    return entries
}
