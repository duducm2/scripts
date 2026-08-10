; =============================================================================
; Utils module: desktop_cut_newest.ahk
; Cut newest Desktop item (file or folder) to clipboard, restore previous window.
; Trigger: Win+Alt+Shift+O
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

DesktopCutNewest_Trigger() {
    origHwnd := WinExist("A")
    desktopPath := ""
    try desktopPath := GetDesktopToRecyclePath()
    catch
        desktopPath := A_Desktop
    if (!desktopPath || !DirExist(desktopPath))
        desktopPath := A_Desktop
    if (!DirExist(desktopPath)) {
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
