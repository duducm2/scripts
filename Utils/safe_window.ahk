; =============================================================================
; Utils module: safe_window.ahk
; Safe WinGet* helpers for #HotIf predicates — never throw when no window is
; focused (focus flicker, desktop, Tasks/Chrome transitions).
; HotIf rule: on failure return "" / false silently. Do not show banners here.
; =============================================================================

SafeWinGetTitle(winTitle := "A") {
    try
        return WinGetTitle(winTitle)
    return ""
}

SafeWinGetClass(winTitle := "A") {
    try
        return WinGetClass(winTitle)
    return ""
}

; Power BI Desktop, or any window whose title mentions powerbi (web / other hosts).
; Pair with && !IsFileDialogActive() at the #HotIf site when that predicate exists.
IsPowerBIActive() {
    if WinActive("ahk_exe PBIDesktop.exe")
        return true
    return InStr(SafeWinGetTitle(), "powerbi", false)
}
