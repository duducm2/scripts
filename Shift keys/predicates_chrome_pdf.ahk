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
; The window title is the PDF document title and often has no ".pdf" (see
; Utils/chrome-pdf-viwer.md). The address-bar path is the filename.
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
    if Chrome_IsAiCompanionTitle(title)
        return false

    if (hwnd = g_ChromePdf_CacheHwnd
        && title = g_ChromePdf_CacheTitle
        && WinExist("ahk_id " g_ChromePdf_CacheHwnd)) {
        if (!g_ChromePdf_CacheResult)
            ChromePdf_InvalidateSession()
        return g_ChromePdf_CacheResult
    }

    ; Omnibox only. A document search walks every PDF page and blocks the hook.
    url := ChromeHotIf_OmniboxUrl(hwnd)
    if (url = "") {
        ChromePdf_InvalidateSession()
        return false
    }
    result := ChromePdf_OmniboxIsViewer(url)

    g_ChromePdf_CacheHwnd := hwnd
    g_ChromePdf_CacheTitle := title
    g_ChromePdf_CacheResult := result
    if (!result)
        ChromePdf_InvalidateSession()
    return result
}

; Address-bar path ends in .pdf, or the bar is Chrome's built-in viewer extension.
; Query and fragment are ignored, so a search for "file.pdf" is not the viewer.
ChromePdf_OmniboxIsViewer(url) {
    if (url = "")
        return false
    u := StrLower(url)
    if InStr(u, "mhjfbmdgcfjbbpaeojofohoefgiehjai")
        return true
    path := RegExReplace(u, "[#?].*$", "")
    return RegExMatch(path, "\.pdf/?$") > 0
}
