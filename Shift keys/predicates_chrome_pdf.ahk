; =============================================================================
; Shift keys module: predicates_chrome_pdf.ahk
; IsChromePdfViewerActive predicate
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

; Chrome PDF Viewer Shortcuts
;-------------------------------------------------------------------
; Cache for IsChromePdfViewerActive (efficiency-canon: cheap #HotIf).
; Same Chrome HWND can switch PDF <-> non-PDF tabs, so the cache key is hwnd + title.
global g_ChromePdf_CacheHwnd := 0
global g_ChromePdf_CacheTitle := ""
global g_ChromePdf_CacheResult := false

ChromePdf_InvalidatePredicateCache() {
    global g_ChromePdf_CacheHwnd, g_ChromePdf_CacheTitle
    g_ChromePdf_CacheHwnd := 0
    g_ChromePdf_CacheTitle := ""
}

IsChromePdfViewerActive() {
    HookTiming_Begin("IsChromePdfViewerActive")
    try
        return IsChromePdfViewerActive_Run()
    finally
        HookTiming_End("IsChromePdfViewerActive")
}

IsChromePdfViewerActive_Run() {
    global g_ChromePdf_CacheHwnd, g_ChromePdf_CacheTitle, g_ChromePdf_CacheResult

    ; Hard gate: avoid conflicts with non-Chrome apps
    if !WinActive("ahk_exe chrome.exe")
        return false

    hwnd := WinExist("A")
    if (!hwnd)
        return false

    title := ""
    try title := WinGetTitle("ahk_id " hwnd)
    if (Chrome_IsAiCompanionTitle(title) || !InStr(title, ".pdf", false))
        return false

    if (hwnd = g_ChromePdf_CacheHwnd
        && title = g_ChromePdf_CacheTitle
        && WinExist("ahk_id " g_ChromePdf_CacheHwnd)) {
        if (!g_ChromePdf_CacheResult)
            ChromePdf_InvalidateSession()
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
        ; UIA failed; do not cache so next call retries. Drop any toolbar session too.
        ChromePdf_InvalidateSession()
        return false
    }

    g_ChromePdf_CacheHwnd := hwnd
    g_ChromePdf_CacheTitle := title
    g_ChromePdf_CacheResult := result
    if (!result)
        ChromePdf_InvalidateSession()
    return result
}
