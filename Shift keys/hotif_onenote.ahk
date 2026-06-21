; =============================================================================
; Shift keys module: hotif_onenote.ahk
; OneNote hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

;-------------------------------------------------------------------
; OneNote Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe onenote.exe") && !IsFileDialogActive()

; Shift + P : Onenote: select line and children
+p:: Send("^+-") ; Remaps to Ctrl + Shift + -

; Shift + F : Advanced Searching with double quotes
+f:: {
    Send "^f"
    Sleep 50
    Send "^a"
    Sleep 20
    Send "{Del}"
    Sleep 20
    Send '""'
    Sleep 20
    Send "{Left}"
}

; Shift + D : Onenote: delete line and children
+d:: {
    Send("^+-") ; Remaps to Ctrl + Shift + -
    Send "{Del}"
}

; Shift + S : Onenote: delete only current line (keep children)
+s:: {
    Send("+{Right}")
    Send "{Del}"
}

; Shift + U : Onenote: collapse
+u:: Send("!+{+}")     ; Remaps to Alt + Shift + +

; Shift + Y : Onenote: expand
+y:: Send("!+{-}")     ; Remaps to Alt + Shift + -

; Shift + I : Onenote: collapse all
+i:: Send("!+1")     ; Remaps to Alt + Shift + 1

; Shift + O : Onenote: expand all
+o:: Send("!+0")     ; Remaps to Alt + Shift + 0

#HotIf
