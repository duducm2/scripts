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

    beforeMap := Map()
    try {
        for hwnd in WinGetList()
            beforeMap[hwnd] := true
    } catch {
    }
    fgBefore := WinExist("A")

    ; Let the launched app take foreground (Windows focus-stealing guard).
    try DllCall("AllowSetForegroundWindow", "UInt", 0xFFFFFFFF)  ; ASFW_ANY
    catch {
    }

    try {
        if (DesktopCutNewest_ShouldOpenInNewChromeWindow(newest))
            DesktopCutNewest_OpenInNewChromeWindow(newest)
        else
            Run('"' . newest . '"')
    } catch as err {
        ShowCenteredOverlay_Utils("❌ Failed to open Desktop item", 2500, BANNER_ACCENT_ERROR)
        return
    }

    newHwnd := DesktopCutNewest_WaitForOpenedWindow(beforeMap, fgBefore, 5000)
    if (newHwnd)
        DesktopCutNewest_ActivateHwnd(newHwnd)

    SplitPath(newest, &name)
    ShowCenteredOverlay_Utils("📂 Open: " name, 1800, BANNER_ACCENT_SUCCESS)

    ; Overlay / AutoSlot can steal focus — finish on the opened window.
    if (newHwnd)
        DesktopCutNewest_ActivateHwnd(newHwnd)
}

; First new visible top-level window after open, or FG if it changed (reuse case).
DesktopCutNewest_WaitForOpenedWindow(beforeMap, fgBefore, timeoutMs := 5000) {
    if (!IsObject(beforeMap))
        beforeMap := Map()
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        try {
            for hwnd in WinGetList() {
                if beforeMap.Has(hwnd)
                    continue
                if !WinExist("ahk_id " hwnd)
                    continue
                try {
                    if !DllCall("IsWindowVisible", "Ptr", hwnd)
                        continue
                    ; Skip owned popups / tool windows without a normal title when possible.
                    if (DllCall("GetWindow", "Ptr", hwnd, "UInt", 4))  ; GW_OWNER
                        continue
                } catch {
                    continue
                }
                return hwnd
            }
            fg := WinExist("A")
            if (fg && fg != fgBefore)
                return fg
        } catch {
        }
        Sleep 50
    }
    fg := WinExist("A")
    return (fg && fg != fgBefore) ? fg : 0
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

; True when Windows would open this path with Chrome (so we can force --new-window).
DesktopCutNewest_ShouldOpenInNewChromeWindow(path) {
    if (!path || !FileExist(path))
        return false
    SplitPath(path, , , &ext)
    if (ext = "")
        return false
    extDot := "." . StrLower(ext)
    cmd := DesktopCutNewest_GetAssocOpenCommand(extDot)
    if (cmd != "" && InStr(cmd, "chrome", false))
        return true
    ; Common browser types when Chrome is the daily driver and assoc lookup is empty/odd.
    static browserExts := Map("html", 1, "htm", 1, "pdf", 1, "svg", 1, "mhtml", 1, "webp", 1)
    return browserExts.Has(StrLower(ext))
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

DesktopCutNewest_OpenInNewChromeWindow(path) {
    fileUrl := DesktopCutNewest_PathToFileUrl(path)
    Run('chrome.exe --new-window "' . fileUrl . '"')
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
