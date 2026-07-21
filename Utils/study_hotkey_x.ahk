; =============================================================================
; Utils module: study_hotkey_x.ahk
; Peek PDF study hotkey #!+x
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

#!+x::
{
    hwnd := WinExist("ahk_exe QuickLook.exe")
    if hwnd {
        if (STUDY_TOPIC_QL_STRICT_LAYOUT) {
            QuickLook_ApplyStudyLayout(hwnd, true, 0)
        } else {
            try {
                WinShow("ahk_id " hwnd)
                WinActivate("ahk_id " hwnd)
                WinWaitActive("ahk_id " hwnd, , 1)
            } catch {
            }
            QuickLook_ClickWindowCenter(hwnd)
            try
                ControlSend("^End", "ahk_id " hwnd)
            catch {
                try {
                    WinActivate("ahk_id " hwnd)
                    WinWaitActive("ahk_id " hwnd, , 1)
                } catch {
                }
                try Send("^End")
            }
        }
        return
    }
    ShowStudyTopicSelector()
}
