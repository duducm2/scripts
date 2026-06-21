; =============================================================================
; Shift keys module: hotif_command_palette.ahk
; Command Palette hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

#HotIf WinActive("Command Palette")

; Ctrl + H : Trigger Ctrl+Shift+E
^h:: Send "^+e"

; Shift + C : Trigger Ctrl+Shift+C (Copy file path)
+c:: Send "^+c"

; Shift + B : Go Home (select all then delete 3 chars)
+b:: {
    Send "^a"
    Sleep 30
    Send "{Backspace 3}"
}

; Shift + U : Insert double quotes twice, then hit left arrow
+u:: Send '""{Left}'

; Shift + O : Focus on Folders Only
+o:: {
    Send "!+w"
    Sleep 120
    Send "{Tab}"
    Sleep 30
    Send "{Enter}"
    Sleep 30
    Send "{Up}"
    Sleep 30
    Send "{Up}"
    Sleep 30
    Send "{Down}"
    Sleep 30
    Send "{Enter}"
    Sleep 50
    Send "{Tab}"
}

; Shift + P : Focus on Files Only
+p:: {
    Send "!+w"
    Sleep 120
    Send "{Tab}"
    Sleep 30
    Send "{Enter}"
    Sleep 30
    Send "{Up}"
    Sleep 30
    Send "{Up}"
    Sleep 30
    Send "{Down}"
    Sleep 30
    Send "{Down}"
    Sleep 30
    Send "{Enter}"
    Sleep 50
    Send "{Tab}"
}

; Shift + I : Trigger Command Palette Bookmark "add new bookmark" shortcut
+i:: {
    Send "^!#m"
}

; Shift + D : Exclude current bookmark (Ctrl+Shift+Delete, Tab, Enter)
+d:: {
    Send "^+{Delete}"
    Sleep 50
    Send "{Tab}"
    Sleep 30
    Send "{Enter}"
}

; Ctrl + 1 : Trigger Enter
^1:: Send "{Enter}"

; Ctrl + 2 : Trigger Down then Enter
^2:: {
    Send "{Down}"
    Send "{Enter}"
}

; Ctrl + 3 : Trigger Down twice then Enter
^3:: {
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Ctrl + 4 : Trigger Down three times then Enter
^4:: {
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Ctrl + 5 : Trigger Down four times then Enter
^5:: {
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Ctrl + 6 : Trigger Down five times then Enter
^6:: {
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Easy Selection
; Alt + 1 : Easy Selection - 1st item
!1::
{
    Send "{Enter}"
}

; Alt + 2 : Easy Selection - 2nd item
!2::
{
    Send "{Down}"
    Send "{Enter}"
}

; Alt + 3 : Easy Selection - 3rd item
!3::
{
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Alt + 4 : Easy Selection - 4th item
!4::
{
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Alt + 5 : Easy Selection - 5th item
!5::
{
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

#HotIf
