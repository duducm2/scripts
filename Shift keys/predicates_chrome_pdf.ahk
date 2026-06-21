; =============================================================================
; Shift keys module: predicates_chrome_pdf.ahk
; IsChromePdfViewerActive predicate
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

; Chrome PDF Viewer Shortcuts
;-------------------------------------------------------------------
IsChromePdfViewerActive() {
    ; Hard gate: avoid conflicts with non-Chrome apps
    if !WinActive("ahk_exe chrome.exe")
        return false

    try {
        ; #region agent log
        AgentDebugLog("H1", "IsChromePdfViewerActive_entry")
        ; #endregion
        uia := UIA_Browser("ahk_exe chrome.exe")

        ; Strong fingerprint: Chrome's built-in PDF viewer extension web area
        ; From UIA tree: chrome-extension://mhjfbmdgcfjbbpaeojofohoefgiehjai/index.html
        if (uia.FindElement({ Type: 50030, Value: "chrome-extension://mhjfbmdgcfjbbpaeojofohoefgiehjai", matchmode: "Substring" })) {
            ; #region agent log
            AgentDebugLog("H2", "IsChromePdfViewerActive_extension_match")
            ; #endregion
            return true
        }

        ; Fallback: stable, non-localized PDF toolbar controls
        if (uia.FindElement({ AutomationId: "pageSelector" }) && uia.FindElement({ AutomationId: "save" })) {
            ; #region agent log
            AgentDebugLog("H3", "IsChromePdfViewerActive_toolbar_match")
            ; #endregion
            return true
        }
    } catch {
    }

    ; #region agent log
    AgentDebugLog("H4", "IsChromePdfViewerActive_return_false")
    ; #endregion
    return false
}
