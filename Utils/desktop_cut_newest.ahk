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
    origHwnd := WinExist("A")
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

    if !Clipboard_CutFiles([newest]) {
        ShowCenteredOverlay_Utils("❌ Failed to cut Desktop item", 2500, BANNER_ACCENT_ERROR)
        return
    }
    if !Clipboard_ContainsFilePath(newest) {
        ShowCenteredOverlay_Utils("❌ Cut verify failed", 2500, BANNER_ACCENT_ERROR)
        return
    }

    SplitPath(newest, &name)
    ShowCenteredOverlay_Utils("✂️ Cut: " name, 1800, BANNER_ACCENT_SUCCESS)

    if (origHwnd && WinExist("ahk_id " origHwnd)) {
        try WinActivate("ahk_id " origHwnd)
        WinWaitActive("ahk_id " origHwnd, , 1)
    }
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

    try {
        Run('"' . newest . '"')
    } catch as err {
        ShowCenteredOverlay_Utils("❌ Failed to open Desktop item", 2500, BANNER_ACCENT_ERROR)
        return
    }

    SplitPath(newest, &name)
    ShowCenteredOverlay_Utils("📂 Open: " name, 1800, BANNER_ACCENT_SUCCESS)
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
