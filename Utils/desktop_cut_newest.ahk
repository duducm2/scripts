; =============================================================================
; Utils module: desktop_cut_newest.ahk
; Cut, open, or copy-path newest Desktop item (file or folder).
; Trigger: Win+Alt+Shift+O (same tiering as #!+8 pronunciation):
;   1× = cut newest Desktop item, restore previous window
;   2× within 400 ms (AI_QD_DOUBLE_TAP_MS / ZMK tap-dance) = open with default app
;   hold 700 ms+ (PRONUNCIATION_HOLD_MS / Fast Copy / cheat sheet) = copy path as text
; =============================================================================

DESKTOP_CUT_NEWEST_HOLD_MS := 700

; Returns full path of newest item under desktopPath, or "" if none.
; Newest = later of Creation vs Modified; skips desktop.ini; includes files and folders.
DesktopCutNewest_ResolveNewestPath(desktopPath) {
    if (!desktopPath || !DirExist(desktopPath))
        return ""
    newestPath := ""
    newestStamp := ""
    loop files desktopPath "\*", "FD" {
        if (StrLower(A_LoopFileName) = "desktop.ini")
            continue
        try {
            tC := FileGetTime(A_LoopFileFullPath, "C")
            tM := FileGetTime(A_LoopFileFullPath, "M")
        } catch {
            continue
        }
        stamp := (tC >= tM) ? tC : tM
        if (newestStamp = "" || stamp > newestStamp) {
            newestStamp := stamp
            newestPath := A_LoopFileFullPath
        }
    }
    return newestPath
}

DesktopCutNewest_ResolveDesktopPath() {
    desktopPath := ""
    try desktopPath := GetDesktopToRecyclePath()
    catch
        desktopPath := A_Desktop
    if (!desktopPath || !DirExist(desktopPath))
        desktopPath := A_Desktop
    return DirExist(desktopPath) ? desktopPath : ""
}

DesktopCutNewest_Trigger() {
    desktopPath := DesktopCutNewest_ResolveDesktopPath()
    if (desktopPath = "") {
        ShowCenteredOverlay_Utils("❌ Desktop folder not found", 2500, BANNER_ACCENT_ERROR)
        return
    }

    newest := DesktopCutNewest_ResolveNewestPath(desktopPath)
    if (newest = "") {
        ShowCenteredOverlay_Utils("⚠ Desktop is empty", 2000, BANNER_ACCENT_INTERMEDIATE)
        return
    }
    DesktopCutNewest_CutPath(newest)
}

; Cut a specific Desktop file/folder path to the clipboard (CF_HDROP move).
DesktopCutNewest_CutPath(path) {
    origHwnd := WinExist("A")
    if (!path || !FileExist(path)) {
        ShowCenteredOverlay_Utils("❌ Desktop item not found", 2500, BANNER_ACCENT_ERROR)
        return false
    }

    if !Clipboard_CutFiles([path]) {
        ShowCenteredOverlay_Utils("❌ Failed to cut Desktop item", 2500, BANNER_ACCENT_ERROR)
        return false
    }
    if !Clipboard_ContainsFilePath(path) {
        ShowCenteredOverlay_Utils("❌ Cut verify failed", 2500, BANNER_ACCENT_ERROR)
        return false
    }

    SplitPath(path, &name)
    ShowCenteredOverlay_Utils("✂️ Cut: " name, 1800, BANNER_ACCENT_SUCCESS)

    if (origHwnd && WinExist("ahk_id " origHwnd)) {
        try WinActivate("ahk_id " origHwnd)
        WinWaitActive("ahk_id " origHwnd, , 1)
    }
    return true
}

DesktopCutNewest_OpenNewest() {
    desktopPath := DesktopCutNewest_ResolveDesktopPath()
    if (desktopPath = "") {
        ShowCenteredOverlay_Utils("❌ Desktop folder not found", 2500, BANNER_ACCENT_ERROR)
        return
    }

    newest := DesktopCutNewest_ResolveNewestPath(desktopPath)
    if (newest = "") {
        ShowCenteredOverlay_Utils("⚠ Desktop is empty", 2000, BANNER_ACCENT_INTERMEDIATE)
        return
    }

    SplitPath(newest, &name)
    browser := DesktopCutNewest_ResolveBrowserLaunch(newest)
    beforeMap := DesktopCutNewest_SnapshotWindowMap(browser ? browser.exeName : "")
    fgBefore := WinExist("A")

    ; Let the launched app take foreground (Windows focus-stealing guard).
    try DllCall("AllowSetForegroundWindow", "UInt", 0xFFFFFFFF)  ; ASFW_ANY
    catch {
    }

    try {
        if (browser)
            DesktopCutNewest_OpenInBrowserNewWindow(newest, browser)
        else
            Run('"' . newest . '"')
    } catch as err {
        ShowCenteredOverlay_Utils("❌ Failed to open Desktop item", 2500, BANNER_ACCENT_ERROR)
        return
    }

    newHwnd := DesktopCutNewest_WaitForOpenedWindow(beforeMap, fgBefore, browser ? 8000 : 5000, browser,
        name)
    if (!newHwnd)
        newHwnd := DesktopCutNewest_FindWindowByTitleHint(name)

    if (newHwnd)
        DesktopCutNewest_ActivateHwnd(newHwnd)

    ShowCenteredOverlay_Utils("📂 Open: " name, 1800, BANNER_ACCENT_SUCCESS)

    ; Overlay / AutoSlot can steal focus — re-activate after the banner clears.
    if (newHwnd)
        DesktopCutNewest_ScheduleActivate(newHwnd, 1900)
}

DesktopCutNewest_SnapshotWindowMap(exeName := "") {
    snap := Map()
    try {
        listSpec := exeName ? ("ahk_exe " exeName) : ""
        for hwnd in WinGetList(listSpec)
            snap[hwnd] := true
    } catch {
    }
    return snap
}

; First new visible top-level window after open, or FG if it changed (reuse case).
; When browser is set, only new windows for that exe count (avoids tab-in-existing-window).
DesktopCutNewest_WaitForOpenedWindow(beforeMap, fgBefore, timeoutMs := 5000, browser := "", titleHint := "") {
    if (!IsObject(beforeMap))
        beforeMap := Map()
    listSpec := (IsObject(browser) && browser.exeName) ? ("ahk_exe " browser.exeName) : ""
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        try {
            best := 0
            for hwnd in WinGetList(listSpec) {
                if beforeMap.Has(hwnd)
                    continue
                if !DesktopCutNewest_IsCandidateOpenWindow(hwnd)
                    continue
                if (titleHint != "" && DesktopCutNewest_TitleMatchesHint(hwnd, titleHint))
                    return hwnd
                if (!best)
                    best := hwnd
            }
            if (best)
                return best
            fg := WinExist("A")
            if (fg && fg != fgBefore) {
                if (listSpec = "" || DesktopCutNewest_HwndMatchesExe(fg, browser ? browser.exeName : ""))
                    return fg
            }
        } catch {
        }
        Sleep 50
    }
    fg := WinExist("A")
    if (fg && fg != fgBefore) {
        if (listSpec = "" || DesktopCutNewest_HwndMatchesExe(fg, browser ? browser.exeName : ""))
            return fg
    }
    return 0
}

DesktopCutNewest_IsCandidateOpenWindow(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    try {
        if !DllCall("IsWindowVisible", "Ptr", hwnd)
            return false
        if (DllCall("GetWindow", "Ptr", hwnd, "UInt", 4))  ; GW_OWNER
            return false
    } catch {
        return false
    }
    return true
}

DesktopCutNewest_HwndMatchesExe(hwnd, exeName) {
    if (exeName = "")
        return true
    try return (StrLower(WinGetProcessName("ahk_id " hwnd)) = StrLower(exeName))
    catch
        return false
}

DesktopCutNewest_TitleMatchesHint(hwnd, titleHint) {
    if (titleHint = "")
        return false
    try {
        title := WinGetTitle("ahk_id " hwnd)
        if (title = "")
            return false
        hint := StrLower(titleHint)
        titleLower := StrLower(title)
        if (InStr(titleLower, hint))
            return true
        SplitPath(titleHint, &hintName, , &hintExt)
        if (hintExt != "" && InStr(titleLower, StrLower(hintName)))
            return true
    } catch {
    }
    return false
}

DesktopCutNewest_FindWindowByTitleHint(titleHint) {
    if (titleHint = "")
        return 0
    try {
        for hwnd in WinGetList() {
            if !DesktopCutNewest_IsCandidateOpenWindow(hwnd)
                continue
            if DesktopCutNewest_TitleMatchesHint(hwnd, titleHint)
                return hwnd
        }
    } catch {
    }
    return 0
}

DesktopCutNewest_ScheduleActivate(hwnd, delayMs := 1500) {
    global g_DesktopCutNewest_ScheduleActivateHwnd, g_DesktopCutNewest_ScheduleActivateTimer
    if (!hwnd || delayMs < 1)
        return
    g_DesktopCutNewest_ScheduleActivateHwnd := hwnd
    if (!g_DesktopCutNewest_ScheduleActivateTimer)
        g_DesktopCutNewest_ScheduleActivateTimer := ObjBindMethod(DesktopCutNewest_ScheduleActivateObj, "OnTimer")
    SetTimer(g_DesktopCutNewest_ScheduleActivateTimer, -delayMs)
}

class DesktopCutNewest_ScheduleActivateObj {
    static OnTimer() {
        global g_DesktopCutNewest_ScheduleActivateHwnd
        hwnd := g_DesktopCutNewest_ScheduleActivateHwnd
        g_DesktopCutNewest_ScheduleActivateHwnd := 0
        DesktopCutNewest_ActivateHwnd(hwnd)
    }
}

DesktopCutNewest_ActivateHwnd(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    try {
        try {
            pid := WinGetPID("ahk_id " hwnd)
            if (pid)
                DllCall("AllowSetForegroundWindow", "UInt", pid)
        } catch {
        }
        if (WinGetMinMax("ahk_id " hwnd) = -1)
            WinRestore("ahk_id " hwnd)
        WinActivate("ahk_id " hwnd)
        if WinWaitActive("ahk_id " hwnd, , 0.8)
            return true
        DllCall("SwitchToThisWindow", "Ptr", hwnd, "Int", 1)
        DllCall("SetForegroundWindow", "Ptr", hwnd)
        DllCall("BringWindowToTop", "Ptr", hwnd)
        WinActivate("ahk_id " hwnd)
        if WinWaitActive("ahk_id " hwnd, , 0.5)
            return true
        ; Last-resort nudge used elsewhere when focus lock blocks WinActivate.
        WinSetAlwaysOnTop("On", "ahk_id " hwnd)
        Sleep 40
        WinSetAlwaysOnTop("Off", "ahk_id " hwnd)
        WinActivate("ahk_id " hwnd)
        return !!WinActive("ahk_id " hwnd)
    } catch {
        return false
    }
}

; When Windows opens this path with a browser, return {exe, exeName, flag} for a new window.
DesktopCutNewest_ResolveBrowserLaunch(path) {
    if (!path || !FileExist(path))
        return ""
    SplitPath(path, , , &ext)
    if (ext = "")
        return ""
    extDot := "." . StrLower(ext)
    browser := DesktopCutNewest_ParseBrowserFromAssocCommand(DesktopCutNewest_GetAssocOpenCommand(extDot))
    if (browser)
        return browser
    ; Web formats only — daily driver is chrome.exe elsewhere in this repo.
    static webExts := Map("html", 1, "htm", 1, "svg", 1, "mhtml", 1, "xhtml", 1)
    if (webExts.Has(StrLower(ext)))
        return { exe: "chrome.exe", exeName: "chrome.exe", flag: "--new-window" }
    return ""
}

DesktopCutNewest_ParseBrowserFromAssocCommand(cmd) {
    if (cmd = "")
        return ""
    exe := ""
    if RegExMatch(cmd, '"(?P<exe>[^"]+\.exe)"', &m)
        exe := m.exe
    else if RegExMatch(cmd, '(?i)([A-Z]:\\[^\s"]+\.exe)', &m)
        exe := m[1]
    if (exe = "")
        return ""
    exeName := StrLower(RegExReplace(exe, ".*\\", ""))
    static browserFlags := Map(
        "chrome.exe", "--new-window",
        "msedge.exe", "--new-window",
        "brave.exe", "--new-window",
        "chromium.exe", "--new-window",
        "vivaldi.exe", "--new-window",
        "opera.exe", "--new-window",
        "firefox.exe", "-new-window"
    )
    if !browserFlags.Has(exeName)
        return ""
    return { exe: exe, exeName: exeName, flag: browserFlags[exeName] }
}

; ASSOCSTR_COMMAND for ".ext" via AssocQueryStringW; "" on failure.
DesktopCutNewest_GetAssocOpenCommand(extWithDot) {
    if (!extWithDot)
        return ""
    ; ASSOCF_NONE = 0, ASSOCSTR_COMMAND = 1
    pcch := 0
    hr := DllCall("shlwapi\AssocQueryStringW", "UInt", 0, "UInt", 1, "WStr", extWithDot, "Ptr", 0, "Ptr", 0,
        "UInt*", &pcch, "UInt")
    if (pcch < 2)
        return ""
    buf := Buffer(pcch * 2, 0)
    hr := DllCall("shlwapi\AssocQueryStringW", "UInt", 0, "UInt", 1, "WStr", extWithDot, "Ptr", 0, "Ptr", buf,
        "UInt*", &pcch, "UInt")
    if (hr != 0)
        return ""
    return StrGet(buf, "UTF-16")
}

DesktopCutNewest_PathToFileUrl(path) {
    p := StrReplace(path, "\", "/")
    if (RegExMatch(p, "i)^[a-z]:"))
        p := "/" . p
    p := StrReplace(p, " ", "%20")
    return "file://" . p
}

DesktopCutNewest_OpenInBrowserNewWindow(path, browser) {
    if (!IsObject(browser) || browser.exe = "")
        throw Error("No browser launch info")
    fileUrl := DesktopCutNewest_PathToFileUrl(path)
    Run('"' . browser.exe . '" ' . browser.flag . ' "' . fileUrl . '"')
}

DesktopCutNewest_CopyPath() {
    desktopPath := DesktopCutNewest_ResolveDesktopPath()
    if (desktopPath = "") {
        ShowCenteredOverlay_Utils("❌ Desktop folder not found", 2500, BANNER_ACCENT_ERROR)
        return
    }

    newest := DesktopCutNewest_ResolveNewestPath(desktopPath)
    if (newest = "") {
        ShowCenteredOverlay_Utils("⚠ Desktop is empty", 2000, BANNER_ACCENT_INTERMEDIATE)
        return
    }

    try {
        A_Clipboard := newest
    } catch {
        ShowCenteredOverlay_Utils("❌ Failed to copy path", 2500, BANNER_ACCENT_ERROR)
        return
    }
    if !ClipWait(1) {
        ShowCenteredOverlay_Utils("❌ Clipboard did not update", 2500, BANNER_ACCENT_ERROR)
        return
    }

    SplitPath(newest, &name)
    ShowCenteredOverlay_Utils("📋 Path: " name, 1800, BANNER_ACCENT_SUCCESS)
}

; --- Win+Alt+Shift+O tap / double-tap / hold ---------------------------------
global g_DesktopCutNewest_DoubleTapArmed := false
global g_DesktopCutNewest_LastPressTick := 0
global g_DesktopCutNewest_DoubleTapTimer := 0
global g_DesktopCutNewest_ScheduleActivateHwnd := 0
global g_DesktopCutNewest_ScheduleActivateTimer := 0

class DesktopCutNewest_DoubleTapTimerObj {
    static OnSingleTapTimeout() {
        global g_DesktopCutNewest_DoubleTapArmed, g_DesktopCutNewest_DoubleTapTimer
        if (!g_DesktopCutNewest_DoubleTapArmed)
            return
        g_DesktopCutNewest_DoubleTapArmed := false
        g_DesktopCutNewest_DoubleTapTimer := 0
        DesktopCutNewest_Trigger()
    }
}

DesktopCutNewest_DisarmDoubleTap() {
    global g_DesktopCutNewest_DoubleTapArmed, g_DesktopCutNewest_DoubleTapTimer
    global g_DesktopCutNewest_LastPressTick
    g_DesktopCutNewest_DoubleTapArmed := false
    g_DesktopCutNewest_LastPressTick := 0
    if (g_DesktopCutNewest_DoubleTapTimer) {
        SetTimer(g_DesktopCutNewest_DoubleTapTimer, 0)
        g_DesktopCutNewest_DoubleTapTimer := 0
    }
}

DesktopCutNewest_OnHotkey() {
    global g_DesktopCutNewest_DoubleTapArmed, g_DesktopCutNewest_LastPressTick
    global g_DesktopCutNewest_DoubleTapTimer

    ; Hotkey fires on key-down. Drop queued auto-repeat ghosts that run after a hold
    ; released (those start with O already up and would otherwise arm single-tap cut).
    if !GetKeyState("o", "P")
        return

    thresholdMs := 400
    try thresholdMs := AI_QD_DOUBLE_TAP_MS
    catch {
        thresholdMs := 400
    }

    ; Detect double-tap on key-down (before KeyWait) so a slow release cannot miss the window.
    pressTime := A_TickCount
    elapsed := (g_DesktopCutNewest_LastPressTick > 0) ? (pressTime - g_DesktopCutNewest_LastPressTick) : 9999
    isSecondTap := g_DesktopCutNewest_DoubleTapArmed && elapsed >= 0 && elapsed < thresholdMs

    KeyWait "o", "T" . (DESKTOP_CUT_NEWEST_HOLD_MS / 1000)
    isHold := (A_TickCount - pressTime) >= DESKTOP_CUT_NEWEST_HOLD_MS

    if (isHold) {
        DesktopCutNewest_DisarmDoubleTap()
        DesktopCutNewest_CopyPath()
        ; Stay in this thread until physical release so a repeat cannot start mid-hold
        ; and arm single-tap after we return.
        KeyWait "o"
        return
    }

    if (isSecondTap) {
        DesktopCutNewest_DisarmDoubleTap()
        DesktopCutNewest_OpenNewest()
        return
    }

    ; Arm after quick release: window starts now (matches #!+8 / #!+9).
    if (g_DesktopCutNewest_DoubleTapTimer) {
        SetTimer(g_DesktopCutNewest_DoubleTapTimer, 0)
        g_DesktopCutNewest_DoubleTapTimer := 0
    }
    g_DesktopCutNewest_LastPressTick := A_TickCount
    g_DesktopCutNewest_DoubleTapArmed := true
    g_DesktopCutNewest_DoubleTapTimer := ObjBindMethod(DesktopCutNewest_DoubleTapTimerObj, "OnSingleTapTimeout")
    SetTimer(g_DesktopCutNewest_DoubleTapTimer, -thresholdMs)
}
