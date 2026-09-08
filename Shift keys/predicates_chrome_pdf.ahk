; =============================================================================
; Shift keys module: predicates_chrome_pdf.ahk
; IsChromePdfViewerActive predicate
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

; Chrome PDF Viewer Shortcuts
;-------------------------------------------------------------------
; Cache for IsChromePdfViewerActive (efficiency-canon: cheap #HotIf).
; Same Chrome HWND can switch PDF <-> non-PDF tabs, so cache is hwnd + TTL.
global g_ChromePdf_CacheHwnd := 0
global g_ChromePdf_CacheTick := 0
global g_ChromePdf_CacheResult := false
global g_ChromePdf_CacheTtlMs := 400

IsChromePdfViewerActive() {
    global g_ChromePdf_CacheHwnd, g_ChromePdf_CacheTick, g_ChromePdf_CacheResult, g_ChromePdf_CacheTtlMs

    ; Hard gate: avoid conflicts with non-Chrome apps
    if !WinActive("ahk_exe chrome.exe")
        return false

    hwnd := WinExist("A")
    if (!hwnd)
        return false

    now := A_TickCount
    if (hwnd = g_ChromePdf_CacheHwnd
        && g_ChromePdf_CacheTick
        && (now - g_ChromePdf_CacheTick) < g_ChromePdf_CacheTtlMs
        && WinExist("ahk_id " g_ChromePdf_CacheHwnd)) {
        return g_ChromePdf_CacheResult
    }

    result := false
    try {
        uia := UIA_Browser("ahk_id " hwnd)

        ; Strong fingerprint: Chrome's built-in PDF viewer extension web area
        ; From UIA tree: chrome-extension://mhjfbmdgcfjbbpaeojofohoefgiehjai/index.html
        if (uia.FindElement({ Type: 50030, Value: "chrome-extension://mhjfbmdgcfjbbpaeojofohoefgiehjai", matchmode: "Substring" })) {
            result := true
        } else if (uia.FindElement({ AutomationId: "pageSelector" }) && uia.FindElement({ AutomationId: "save" })) {
            ; Fallback: stable, non-localized PDF toolbar controls
            result := true
        }
    } catch {
        ; UIA failed; do not cache so next call retries
        return false
    }

    g_ChromePdf_CacheHwnd := hwnd
    g_ChromePdf_CacheTick := A_TickCount  ; stamp after UIA so TTL is from completion, not start
    g_ChromePdf_CacheResult := result
    return result
}
