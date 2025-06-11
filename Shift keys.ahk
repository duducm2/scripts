/********************************************************************
 *   Win+Alt+Shift symbol layer shortcuts (AHK v2)
 *   • Provides system-wide symbol shortcuts
 ********************************************************************/

#Requires AutoHotkey v2.0+

#SingleInstance Force

#include UIA-v2\Lib\UIA.ahk
#include UIA-v2\Lib\UIA_Browser.ahk

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

; Shift + h: Focus the current conversation
+h::
{
    try
    {
        ; WhatsApp desktop is Chromium-based, so we can use UIA_Browser.
        ; It should attach to the active window, which is WhatsApp thanks to #HotIf.
        uia := UIA_Browser()
        Sleep 300 ; Give UIA time to attach to the browser. A similar delay is in the reference script.

        ; Find the "Archived" button to use as an anchor.
        ; The user provided: Name:"Archived "
        archivedButton := uia.FindElement({ Name: "Archived ", Type: "Button" })

        if (archivedButton) {
            ; Focus the button without clicking it.
            archivedButton.SetFocus()
            ; Send Tab to move to the main conversation list.
            ; From there, the focus should be on the selected chat.
            SendInput "{Tab}"
        }
        else {
            MsgBox "Could not find the 'Archived' button."
        }
    }
    catch Error as e {
        MsgBox "An error occurred while trying to focus WhatsApp conversation: " e.Message
    }
}

#HotIf

;-------------------------------------------------------------------
; Microsoft Teams Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe teams.exe")

; Shift + P : Pin (🌟)
+p::
{
    Sleep "500"
    Send "^1"
    Sleep "300"
    Send("{AppsKey}")
    Sleep "300"
    Send "{Down}"
    Send "{Down}"
    Send "{Right}"
    Send "{Enter}"
}

; Shift + V : Like (🌟)
+v::
{
    Send "{Enter}"
    Send "{Enter}"
    Send "{Esc}"
}

; Shift + B : Heart (🌟)
+b::
{
    Send "{Enter}"
    Send "{Down}"
    Send "{Enter}"
    Send "{Esc}"
}

; Shift + C : Laugh (🌟)
+c::
{
    Send "{Enter}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
    Send "{Esc}"
}

; Shift + Ç : Remove pin (🌟)
+ç::
{
    Sleep "500"
    Send "^1"
    Sleep "300"
    Send("{AppsKey}")
    Sleep "300"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Shift + R : Reply Message (🌟)
+r::
{
    Send "{Enter}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Shift + E : Edit message (🌟)
+e::
{
    Send "{Enter}"
    Sleep 100
    Send "{Left}"
    Sleep 100
    Send "{Left}"
    Sleep 100
    Send "{Enter}"
}

; Shift + J : Mark as unread (🌟)
+j::
{
    Send "^1"
    Sleep "400"
    Send("{AppsKey}")
    Sleep "400"
    Send "{Down}"
    Send "{Enter}"
}

#HotIf

;-------------------------------------------------------------------
; Outlook Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe OUTLOOK.EXE")

; Shift + 8 : Send to general (🌟)
+8::
{
    Send "!5"
    Send "O"
    Send "{Home}"
    Send "00"
    Send "{Enter}"
}

; Shift + 9 : Send to newsletter (🌟)
+9::
{
    Send "!5"
    Send "O"
    Send "{Home}"
    Send "01"
    Send "{Enter}"
}

#HotIf

;-------------------------------------------------------------------
; Google Chrome Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe chrome.exe")

; Shift + G : Pop up the current tab (🌟)
+g::
{
    Send "{F6}"                        ; Focus address bar (omnibox)
    Sleep 100
    Send "{F6}"                        ; Focus the tab strip (current tab)
    Sleep 100
    Send "{AppsKey}"                   ; Open the tab's context menu (AppsKey or Shift+F10)
    Sleep 100                          ; Wait a moment for menu to open
    Send "m"                           ; Select "Move tab to new window" (press 'm')
    Sleep 100
    Send "{Enter}"                     ; Confirm the action (detach tab)
    Sleep 100
    Send "{Enter}"                     ; Confirm the action (detach tab)
}

#HotIf