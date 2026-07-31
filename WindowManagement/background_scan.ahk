; =============================================================================
; WindowManagement module: background_scan.ahk
; Background window scan, excludes, and collection
; Extracted verbatim from WindowManagement.ahk; loaded via #include into the
; WindowManagement.ahk process, which remains the entry point / source of truth.
; =============================================================================

WM_BackgroundTitleExcludes_Init() {
    global g_WM_BackgroundTitleExcludes, g_WM_BackgroundTitleExcludesReady
    list := []
    seen := Map()
    for needle in ["IT Workplace", "Drafts Monitor", "Form1", "Screenpresso",
        "Sharing control bar |", "Meeting compact view",
        "barra de controle de compartilhamento", "modo de exibição compacto da reunião"]
        WM_BackgroundTitleExcludes_Register(&list, &seen, needle)
    path := WM_BackgroundTitleExcludes_IniPath()
    if (!FileExist(path)) {
        try {
            WM_BackgroundTitleExcludes_WriteList(list)
        } catch {
        }
    }
    try {
        raw := FileRead(path, "UTF-8")
        for entry in WM_BackgroundTitleExcludes_ParseDiskEntries(raw)
            WM_BackgroundTitleExcludes_Register(&list, &seen, entry)
    } catch {
    }
    g_WM_BackgroundTitleExcludes := list
    g_WM_BackgroundTitleExcludesReady := true
}

WM_BackgroundTitleExcludes_Ensure() {
    global g_WM_BackgroundTitleExcludesReady
    if (!g_WM_BackgroundTitleExcludesReady)
        WM_BackgroundTitleExcludes_Init()
}

WM_BackgroundFilterRowsByTitleExcludes(rows) {
    filtered := []
    for row in rows {
        exe := row.HasProp("exe") ? row.exe : ""
        if (!WM_BackgroundTitleIsExcluded(row.title, exe))
            filtered.Push(row)
    }
    return filtered
}

WM_BackgroundTitleIsExcluded(title, exe := "") {
    global g_WM_BackgroundTitleExcludes
    if (title = "" && exe = "")
        return false
    t := StrLower(title)
    e := StrLower(exe)
    for needle in g_WM_BackgroundTitleExcludes {
        n := Trim(needle)
        if (n = "")
            continue
        if (InStr(n, "|")) {
            parts := StrSplit(n, "|", , 2)
            exeNeedle := StrLower(Trim(parts[1]))
            titleNeedle := parts.Length > 1 ? StrLower(Trim(parts[2])) : ""
            if (exeNeedle != "" && e != "" && InStr(e, exeNeedle)) {
                if (titleNeedle = "" || (t != "" && InStr(t, titleNeedle)))
                    return true
            }
            continue
        }
        nLower := StrLower(n)
        if (t != "" && InStr(t, nLower))
            return true
        if (e != "" && InStr(e, nLower))
            return true
    }
    return false
}

WM_BackgroundTitleExcludes_FormatNeedle(title, exe := "") {
    title := Trim(title)
    exe := Trim(exe)
    if (title = "" && exe = "")
        return ""
    if (exe != "" && title != "")
        return exe . "|" . title
    return title != "" ? title : exe
}

WM_BackgroundNeedleToTitleExe(needle) {
    needle := Trim(needle)
    if (InStr(needle, "|")) {
        parts := StrSplit(needle, "|", , 2)
        return { title: Trim(parts[2]), exe: Trim(parts[1]) }
    }
    return { title: needle, exe: "" }
}

WM_BackgroundTitleExcludes_PersistAppend(needle, exe := "") {
    global g_WM_BackgroundTitleExcludes
    needle := WM_BackgroundTitleExcludes_FormatNeedle(needle, exe)
    if (needle = "")
        return false
    parsed := WM_BackgroundNeedleToTitleExe(needle)
    WM_BackgroundTitleExcludes_Init()
    if (WM_BackgroundTitleIsExcluded(parsed.title, parsed.exe)) {
        ShowCenteredOverlay_Utils("ℹ️ Already in exclude list", 2000, BANNER_ACCENT_INFO)
        return false
    }
    list := []
    seen := Map()
    for n in g_WM_BackgroundTitleExcludes
        WM_BackgroundTitleExcludes_Register(&list, &seen, n)
    WM_BackgroundTitleExcludes_Register(&list, &seen, needle)
    path := WM_BackgroundTitleExcludes_IniPath()
    try {
        WM_BackgroundTitleExcludes_WriteList(list)
    } catch as err {
        ShowCenteredOverlay_Utils("❌ Could not save exclude list: " . err.Message, 4000, BANNER_ACCENT_ERROR)
        return false
    }
    WM_BackgroundTitleExcludes_Init()
    if (!WM_BackgroundTitleIsExcluded(parsed.title, parsed.exe)) {
        ShowCenteredOverlay_Utils("❌ Exclude saved but could not be loaded — check " . path, 4500, BANNER_ACCENT_ERROR)
        return false
    }
    ShowCenteredOverlay_Utils("✅ Excluded: " . WM_TruncateTitleForList(needle, 50), 2200, BANNER_ACCENT_SUCCESS)
    return true
}

WM_DebugBackground_LogPath() {
    return A_ScriptDir "\.cursor\wm_background_scan.log"
}

WM_DebugBackgroundEnabled() {
    if (EnvGet("WM_DEBUG_BACKGROUND") = "1")
        return true
    try {
        return IniRead(A_ScriptDir "\assets\data\wm_debug.ini", "Debug", "BackgroundScan", "0") = "1"
    } catch {
        return false
    }
}

WM_WindowGetPlacementShowCmd(hwnd) {
    if (!hwnd)
        return 0
    wp := Buffer(44, 0)
    NumPut("UInt", 44, wp, 0)
    try {
        if !DllCall("GetWindowPlacement", "ptr", hwnd, "ptr", wp)
            return 0
        return NumGet(wp, 8, "UInt")
    } catch {
        return 0
    }
}

; Taskbar-minimized: WinGetMinMax, IsIconic, or GetWindowPlacement showCmd=SW_SHOWMINIMIZED (2).
WM_WindowIsTaskbarMinimized(hwnd) {
    if (!hwnd)
        return false
    try {
        if (WinGetMinMax(hwnd) = -1)
            return true
        if DllCall("IsIconic", "ptr", hwnd)
            return true
        if (WM_WindowGetPlacementShowCmd(hwnd) = 2)
            return true
    } catch {
    }
    return false
}

WM_BackgroundMinimizedPickScore(hwnd) {
    score := 0
    try {
        if (WinGetMinMax(hwnd) = -1)
            score += 100
        if DllCall("IsIconic", "ptr", hwnd)
            score += 50
        title := WinGetTitle(hwnd)
        score += Min(StrLen(title), 200)
        exStyle := DllCall("GetWindowLongPtr", "ptr", hwnd, "int", -20, "ptr")
        if !(exStyle & 0x00000080)
            score += 40
    } catch {
    }
    return score
}

WM_BackgroundMinimizedDedupeKey(hwnd) {
    try {
        return WinGetPID(hwnd) "|" WinGetTitle(hwnd)
    } catch {
        return ""
    }
}

WM_BackgroundBuildVisibleHwndSet() {
    visible := Map()
    loop MonitorGetCount() {
        try {
            for win in GetVisibleWindowsOnMonitor(A_Index, true)
                visible[win.hwnd] := true
        } catch {
        }
    }
    return visible
}

WM_BackgroundHwndOnAnyScriptMonitor(hwnd) {
    try {
        hMon := DllCall("MonitorFromWindow", "ptr", hwnd, "uint", 2, "ptr")
        loop MonitorGetCount() {
            MonitorGet A_Index, &ml, &mt, &mr, &mb
            cx := (ml + mr) // 2
            cy := (mt + mb) // 2
            point64 := (cy & 0xFFFFFFFF) << 32 | (cx & 0xFFFFFFFF)
            if (Integer(DllCall("MonitorFromPoint", "int64", point64, "uint", 2, "ptr")) = Integer(hMon))
                return true
        }
    } catch {
    }
    return false
}

; Same pre-checks as GetVisibleWindowsOnMonitor (before z-order covered test); minimized allowed.
WM_BackgroundPassesVisibleStackGates(hwnd) {
    if (!hwnd)
        return false
    try {
        isMinimized := (WinGetMinMax(hwnd) = -1)
        if (!isMinimized && !DllCall("IsWindowVisible", "ptr", hwnd))
            return false
        exStyle := DllCall("GetWindowLongPtr", "ptr", hwnd, "int", -20, "ptr")
        if (exStyle & 0x00000080)
            return false
        class := WinGetClass(hwnd)
        if (class = "Progman" || class = "WorkerW")
            return false
        if (WM_IsDesktopOrTaskbarClass(class))
            return false
        if (WinGetTitle(hwnd) = "")
            return false
        if (WM_IsExcludedIndicatorWindow(hwnd))
            return false
        if (!WM_BackgroundHwndOnAnyScriptMonitor(hwnd))
            return false
        if (!isMinimized) {
            rect := Buffer(16, 0)
            if !DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect)
                return false
            w := NumGet(rect, 8, "int") - NumGet(rect, 0, "int")
            h := NumGet(rect, 12, "int") - NumGet(rect, 4, "int")
            if (w < 120 || h < 80)
                return false
        }
    } catch {
        return false
    }
    return true
}

WM_BackgroundIsSystemNoiseTitle(title) {
    if (title = "")
        return false
    t := StrLower(title)
    for needle in [
        "bluetooth", "notification", "notifications", "windows input experience", "toast",
        "gdi+ hook", "msctfime", "cicero", "broadcastevent", "nvidia geforce", "widget",
        "program manager", "default ime", "systray", "action center", "quick settings"
    ] {
        if (InStr(t, needle))
            return true
    }
    return false
}

WM_BackgroundExplainReject(hwnd, foreHwnd) {
    if (!hwnd)
        return "no_hwnd"
    if (hwnd = foreHwnd)
        return "foreground"
    try {
        exStyle := DllCall("GetWindowLongPtr", "ptr", hwnd, "int", -20, "ptr")
        if (exStyle & 0x00000080)
            return "toolwindow"
        class := WinGetClass(hwnd)
        if (WM_IsDesktopOrTaskbarClass(class))
            return "desktop_class"
        title := WinGetTitle(hwnd)
        if (title = "")
            return "empty_title"
        exe := ""
        try exe := WinGetProcessName("ahk_id " hwnd)
        if (WM_BackgroundIsSystemNoiseTitle(title))
            return "system_noise"
        if (WM_BackgroundTitleIsExcluded(title, exe))
            return "title_excluded"
        if (WM_IsExcludedIndicatorWindow(hwnd))
            return "indicator"
    } catch {
        return "inspect_error"
    }
    return ""
}

; Open but not visible on any monitor: z-order covered and/or taskbar-minimized (inverse of GetVisibleWindowsOnMonitor per monitor).
WM_BackgroundEnumerateHiddenHwnds() {
    visibleAll := WM_BackgroundBuildVisibleHwndSet()
    bestByKey := Map()
    winListCount := 0
    hiddenCandidates := 0
    skippedVisible := 0
    skippedEmptyKey := 0
    prevDetect := A_DetectHiddenWindows
    DetectHiddenWindows true
    try {
        for hwnd in WinGetList() {
            winListCount++
            try {
                if (visibleAll.Has(hwnd)) {
                    skippedVisible++
                    continue
                }
                if (!WM_BackgroundPassesVisibleStackGates(hwnd))
                    continue
                hiddenCandidates++
                key := WM_BackgroundMinimizedDedupeKey(hwnd)
                if (key = "") {
                    skippedEmptyKey++
                    continue
                }
                score := WM_BackgroundMinimizedPickScore(hwnd)
                if (!bestByKey.Has(key) || score > bestByKey[key].score)
                    bestByKey[key] := { hwnd: hwnd, score: score }
            } catch {
            }
        }
    } finally {
        DetectHiddenWindows prevDetect
    }
    hwnds := []
    for , entry in bestByKey
        hwnds.Push(entry.hwnd)
    global g_WM_LastEnumerateStats
    g_WM_LastEnumerateStats := Map(
        "winList", winListCount,
        "visibleOnMonitors", visibleAll.Count,
        "hiddenCandidates", hiddenCandidates,
        "skippedVisible", skippedVisible,
        "deduped", hwnds.Length,
        "skippedEmptyKey", skippedEmptyKey)
    return hwnds
}

WM_PlayNoWindowSound() {
    try ScriptSoundPlay(A_ScriptDir . "\assets\sounds\no-window.wav", true)
}

; Fast scans can finish before the banner repaints; keep "Scanning…" visible briefly so the no-window chime is not raced by GUI setup.
WM_EnsureBackgroundScanMinimumDwell(minMs := 280) {
    global g_WM_BackgroundScanBannerTick
    if (!g_WM_BackgroundScanBannerTick)
        return
    elapsed := A_TickCount - g_WM_BackgroundScanBannerTick
    if (elapsed < minMs)
        Sleep(minMs - elapsed)
}

; Show empty-scan message, then play no-window chime (wait=true) so teardown cannot cut playback.
WM_PresentNoBackgroundWindowsEmpty(message := "", useExistingLoadingBar := true) {
    if (message = "")
        message := WM_FormatBackgroundCollectEmptyMessage()
    WM_EnsureBackgroundScanMinimumDwell()
    if (useExistingLoadingBar)
        StandardLoadingBar_Update(message, BANNER_ACCENT_INFO)
    else
        ShowCenteredOverlay_Utils(message, 4500, BANNER_ACCENT_INFO)
    Sleep 80
    WM_PlayNoWindowSound()
    if (useExistingLoadingBar)
        StandardLoadingBar_Hide(4500)
}

; Empty-scan message only (sound via WM_PresentNoBackgroundWindowsEmpty).
WM_NotifyNoBackgroundWindowsFound(foreHwnd := 0) {
    return WM_FormatBackgroundCollectEmptyMessage()
}

WM_FormatBackgroundCollectEmptyMessage() {
    global g_WM_LastBackgroundCollectStats
    st := g_WM_LastBackgroundCollectStats
    if (!IsObject(st) || st.Count = 0)
        return "ℹ️ No hidden windows found."
    hidden := st.Get("hiddenCandidates", 0)
    visible := st.Get("visibleOnMonitors", 0)
    deduped := st.Get("deduped", 0)
    if (hidden = 0)
        return "ℹ️ All open windows are visible on your monitors (" . visible . " unobstructed)."
    msg := "ℹ️ " . hidden . " hidden (" . visible . " visible on monitors), " . deduped . " unique — none listed."
    rejects := st.Get("rejects", Map())
    if (IsObject(rejects) && rejects.Count > 0) {
        parts := []
        for reason, count in rejects
            parts.Push(reason . ":" . count)
        msg .= " Excluded: " . WM_ArrJoin(parts, ", ") . "."
    } else
        msg .= " Check title excludes or foreground window."
    return msg
}

WM_DebugBackgroundWindowScan() {
    foreHwnd := 0
    try foreHwnd := WinGetID("A")
    collected := WM_CollectBackgroundWindows()
    inList := Map()
    for row in collected
        inList[row.hwnd] := true
    logPath := WM_DebugBackground_LogPath()
    try DirCreate(A_ScriptDir "\.cursor")
    catch {
    }
    lines := []
    visibleAll := WM_BackgroundBuildVisibleHwndSet()
    lines.Push("=== WM background scan " A_Now " foreHwnd=" foreHwnd " collected=" collected.Length
        " visibleOnMonitors=" visibleAll.Count " ===")
    prevDetect := A_DetectHiddenWindows
    DetectHiddenWindows true
    try {
        for hwnd in WinGetList() {
            line := ""
            try {
                exe := WinGetProcessName("ahk_id " hwnd)
                class := WinGetClass(hwnd)
                title := WM_TruncateTitleForList(WinGetTitle(hwnd), 60)
                minMax := WinGetMinMax(hwnd)
                iconic := DllCall("IsIconic", "ptr", hwnd) ? 1 : 0
                showCmd := WM_WindowGetPlacementShowCmd(hwnd)
                visible := DllCall("IsWindowVisible", "ptr", hwnd) ? 1 : 0
                exStyle := DllCall("GetWindowLongPtr", "ptr", hwnd, "int", -20, "ptr")
                toolWin := (exStyle & 0x00000080) ? 1 : 0
                monIdx := 0
                try {
                    hMon := DllCall("MonitorFromWindow", "ptr", hwnd, "uint", 2, "ptr")
                    loop MonitorGetCount() {
                        MonitorGet A_Index, &ml, &mt, &mr, &mb
                        cx := (ml + mr) // 2
                        cy := (mt + mb) // 2
                        point64 := (cy & 0xFFFFFFFF) << 32 | (cx & 0xFFFFFFFF)
                        if (Integer(DllCall("MonitorFromPoint", "int64", point64, "uint", 2, "ptr")) = Integer(hMon)) {
                            monIdx := A_Index
                            break
                        }
                    }
                }
                reject := WM_BackgroundExplainReject(hwnd, foreHwnd)
                inVisibleSet := visibleAll.Has(hwnd) ? 1 : 0
                line := Format(
                    "hwnd=0x{:X} exe={} class={} mon={} MinMax={} inVisibleSet={} visible={} toolWin={} inList={} reject={} title={}",
                    hwnd, exe, class, monIdx, minMax, inVisibleSet, visible, toolWin, inList.Has(hwnd) ? 1 : 0,
                    reject != "" ? reject : "ok", title)
            } catch as err {
                line := Format("hwnd=0x{:X} error={}", hwnd, err.Message)
            }
            lines.Push(line)
        }
    } finally {
        DetectHiddenWindows prevDetect
    }
    try {
        FileDelete(logPath)
    } catch {
    }
    try {
        FileAppend(WM_ArrJoin(lines, "`n") "`n", logPath, "UTF-8")
        ShowCenteredOverlay_Utils("📋 Background scan: " collected.Length " listed — see wm_background_scan.log", 3500,
            BANNER_ACCENT_INFO)
    } catch as err {
        ShowCenteredOverlay_Utils("❌ Background scan log failed: " err.Message, 3500, BANNER_ACCENT_ERROR)
    }
}
