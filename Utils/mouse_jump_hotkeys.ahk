; =============================================================================
; Utils module: mouse_jump_hotkeys.ahk
; Win+Alt+Shift+Arrow five-step hotkeys
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; Win+Alt+Shift+Arrow: send that arrow key five times
#!+Right::
{
    Send("{Right 5}")
    return
}

#!+Left::
{
    Send("{Left 5}")
    return
}

#!+Down::
{
    Send("{Down 5}")
    return
}

#!+Up::
{
    Send("{Up 5}")
    return
}
