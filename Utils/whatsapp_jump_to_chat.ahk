; =============================================================================
; Utils module: whatsapp_jump_to_chat.ahk
; Shared Jump-to-Chat logic for WhatsApp (Web / Desktop)
; =============================================================================

; Returns true if WhatsApp is active and ready for keyboard shortcuts.
; Cold-starts settle ~2s after the window appears so the SPA can accept Alt+K.
WhatsAppJump_ActivateOrOpen() {
    global IS_WORK_ENVIRONMENT
    prevTitleMode := A_TitleMatchMode
    try {
        SetTitleMatchMode(2)
        wasOpen := WinExist("WhatsApp")
        if (wasOpen) {
            WinActivate("WhatsApp")
            if !WinWaitActive("WhatsApp", , 3) {
                WinActivate("WhatsApp")
                if !WinWaitActive("WhatsApp", , 2) {
                    ShowCenteredOverlay_Utils("❌ Could not activate WhatsApp.", 2000, BANNER_ACCENT_ERROR)
                    return false
                }
            }
            return true
        }

        if (IS_WORK_ENVIRONMENT) {
            Run "C:\Users\fie7ca\Documents\Shortcuts\WhatsApp.lnk"
        } else {
            Run "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\apps do Chrome\WhatsApp Web.lnk"
        }

        if !WinWait("WhatsApp", , 30) {
            ShowCenteredOverlay_Utils("❌ WhatsApp did not start in time.", 2000, BANNER_ACCENT_ERROR)
            return false
        }
        WinActivate("WhatsApp")
        if !WinWaitActive("WhatsApp", , 5) {
            WinActivate("WhatsApp")
            if !WinWaitActive("WhatsApp", , 3) {
                ShowCenteredOverlay_Utils("❌ Could not activate WhatsApp.", 2000, BANNER_ACCENT_ERROR)
                return false
            }
        }
        ; SPA / Chrome App needs time before Alt+K and paste work.
        Sleep 2000
        WinActivate("WhatsApp")
        Sleep 150
        return true
    } finally {
        SetTitleMatchMode(prevTitleMode)
    }
}

; Opens search, pastes contact, Enter to select chat only. Does not paste the message.
WhatsAppJumpToChat(contact) {
    if (!WhatsAppJump_ActivateOrOpen())
        return false

    oldWinDelay := A_WinDelay
    oldKeyDelay := A_KeyDelay
    oldControlDelay := A_ControlDelay

    try {
        SetWinDelay 0
        SetKeyDelay 0, 0
        SetControlDelay 0

        Send "{LWin Up}{RWin Up}{LAlt Up}{RAlt Up}{LShift Up}{RShift Up}"
        Sleep 80

        ; Ensure WhatsApp still has focus before search (InputBox may have stolen it).
        prevTitleMode := A_TitleMatchMode
        try {
            SetTitleMatchMode(2)
            if WinExist("WhatsApp") {
                WinActivate("WhatsApp")
                WinWaitActive("WhatsApp", , 2)
            }
        } finally {
            SetTitleMatchMode(prevTitleMode)
        }
        Sleep 100

        ; Alt+K is the WhatsApp search shortcut (Shift keys maps Shift+S to this)
        Send "!k"
        Sleep 250

        loop 5 {
            A_Clipboard := ""
            A_Clipboard := contact
            if ClipWait(2) && (A_Clipboard = contact)
                break
            if A_Index = 5 {
                ShowCenteredOverlay_Utils("❌ CLIPBOARD ERROR - TRY AGAIN", 3000, BANNER_ACCENT_ERROR)
                return false
            }
            Sleep 100
        }

        Send "^v"
        Sleep 300
        Send "{Enter}"
        ; Let the chat composer take focus before caller pastes the message.
        Sleep 700

        return true
    } finally {
        SetWinDelay oldWinDelay
        SetKeyDelay oldKeyDelay, 0
        SetControlDelay oldControlDelay
    }
}
