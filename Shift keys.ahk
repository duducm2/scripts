/********************************************************************
 *   Win+Alt+Shift symbol layer shortcuts (AHK v2)
 *   • Provides system-wide symbol shortcuts
 ********************************************************************/

#Requires AutoHotkey v2.0+

; Function to send symbol characters
SendSymbol(sym) {
    SendText(sym)
}

; Symbol shortcuts using Win+Alt+Shift combinations
+y:: SendSymbol("?")   ; Win+Alt+Shift+Y → ?

;-------------------------------------------------------------------
; OneNote Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe onenote.exe")

; Shift + U : Onenote: select line and children
+u:: Send("^+-") ; Remaps to Ctrl + Shift + -

; Shift + I : Onenote: expand all
+i:: Send("!+-")     ; Remaps to Alt + Shift + 0

; Shift + O : Onenote: collapse all
+o:: Send("!+{+}")     ; Remaps to Alt + Shift + 1

; Shift + I : Onenote: expand all
+k:: Send("!+1")     ; Remaps to Alt + Shift + 0

; Shift + O : Onenote: collapse all
+l:: Send("!+0")     ; Remaps to Alt + Shift + 1

#HotIf

;-------------------------------------------------------------------
; WhatsApp Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("WhatsApp")

; Shift + h: Trigger 10 tabs backwards
+h::
{
    loop 10 {
        Send "+{Tab}"
    }
}

#HotIf