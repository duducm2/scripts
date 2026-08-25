; =============================================================================
; Utils module: desktop_cut_newest.ahk
; Cut or open newest Desktop item (file or folder).
; Trigger: Win+Alt+Shift+O tap-dance (400 ms = AI_QD_DOUBLE_TAP_MS / ZMK tap-dance):
;   1× = cut newest Desktop item, restore previous window
;   2× = open newest Desktop item with the default app
; =============================================================================

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

; --- Win+Alt+Shift+O tap-dance ------------------------------------------------
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

DesktopCutNewest_OnHotkey() {
    global g_DesktopCutNewest_DoubleTapArmed, g_DesktopCutNewest_LastPressTick
    global g_DesktopCutNewest_DoubleTapTimer

    thresholdMs := 400
    try thresholdMs := AI_QD_DOUBLE_TAP_MS
    catch {
        thresholdMs := 400
    }

    now := A_TickCount
    elapsed := (g_DesktopCutNewest_LastPressTick > 0) ? (now - g_DesktopCutNewest_LastPressTick) : 9999

    if (g_DesktopCutNewest_DoubleTapArmed && elapsed >= 0 && elapsed < thresholdMs) {
        g_DesktopCutNewest_DoubleTapArmed := false
        g_DesktopCutNewest_LastPressTick := 0
        if (g_DesktopCutNewest_DoubleTapTimer) {
            SetTimer(g_DesktopCutNewest_DoubleTapTimer, 0)
            g_DesktopCutNewest_DoubleTapTimer := 0
        }
        DesktopCutNewest_OpenNewest()
        return
    }

    g_DesktopCutNewest_LastPressTick := now
    g_DesktopCutNewest_DoubleTapArmed := true
    g_DesktopCutNewest_DoubleTapTimer := ObjBindMethod(DesktopCutNewest_DoubleTapTimerObj, "OnSingleTapTimeout")
    SetTimer(g_DesktopCutNewest_DoubleTapTimer, -thresholdMs)
}
