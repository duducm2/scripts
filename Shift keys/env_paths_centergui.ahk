; =============================================================================
; Shift keys module: env_paths_centergui.ahk
; Env paths, ShowErr, CenterGuiOnActiveMonitor
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

;-------------------------------------------------------------------
; Environment paths (unchanged)
;-------------------------------------------------------------------
global WORK_SCRIPTS_PATH := "C:\Users\fie7ca\Documents\01 - Scripts"
global PERSONAL_SCRIPTS_PATH := "C:\Users\eduev\Meu Drive\17 - Projects\scripts"
; global IS_WORK_ENVIRONMENT   := true    ; set to false on personal rig // This will now be loaded from env.ahk

; ---------------------------------------------------------------------------
; ShowErr(msgOrErr)  â€" uniform MsgBox for any thrown value
; ---------------------------------------------------------------------------
ShowErr(err) {
    text := ""
    try {
        if (err is Error) {
            text := err.Message
            if (HasProp(err, "Extra") && err.Extra != "")
                text .= "`n`nExtra:`n" err.Extra
        } else if (err is String) {
            text := err
        } else {
            ; Covers TargetError and other thrown objects/values.
            text := Format("{}", err)
        }
    } catch {
        text := "<unprintable error value>"
    }
    MsgBox("Error:`n" text, "Error", "IconX")
}

; ---------------------------------------------------------------------------
; Centre cheat sheet on the same monitor as StandardLoadingBar / banners
; (GetActiveMonitorWorkArea_StandardBar in Utils.ahk). Do not clamp to (0,0):
; that pulls the window onto the primary display when wx/wy are negative and
; causes multi-monitor spanning / wrong placement.
; ---------------------------------------------------------------------------
CenterGuiOnActiveMonitor(guiObj) {
    guiObj.GetPos(, , &guiW, &guiH)
    GetActiveMonitorWorkArea_StandardBar(&wx, &wy, &wr, &wb)
    ww := wr - wx
    wh := wb - wy
    guiX := wx + (ww - guiW) / 2
    guiY := wy + (wh - guiH) / 2
    guiX := Max(wx, Min(guiX, wx + ww - guiW))
    guiY := Max(wy, Min(guiY, wy + wh - guiH))
    guiObj.Show("NoActivate x" Round(guiX) " y" Round(guiY))
}
