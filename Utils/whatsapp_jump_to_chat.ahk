; =============================================================================
; Utils module: whatsapp_jump_to_chat.ahk
; Shared Jump-to-Chat logic for WhatsApp (Web / Desktop)
; =============================================================================

WhatsAppJump_ActivateOrOpen() {
    global IS_WORK_ENVIRONMENT
    prevTitleMode := A_TitleMatchMode
    try {
        SetTitleMatchMode(2)
        if WinExist("WhatsApp") {
            WinActivate("WhatsApp")
            return true
        } else {
            if (IS_WORK_ENVIRONMENT) {
                Run "C:\Users\fie7ca\Documents\Shortcuts\WhatsApp.lnk"
            } else {
                Run "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\apps do Chrome\WhatsApp Web.lnk"
            }
            if WinWaitActive("WhatsApp", , 10) {
                return true
            } else {
                ShowCenteredOverlay_Utils("❌ WhatsApp did not start in time.", 2000, BANNER_ACCENT_ERROR)
                return false
            }
        }
    } finally {
        SetTitleMatchMode(prevTitleMode)
    }
}

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

        ; Alt+K is the WhatsApp search shortcut (Shift keys maps Shift+S to this)
        Send "!k"
        Sleep 150

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

        return true
    } finally {
        SetWinDelay oldWinDelay
        SetKeyDelay oldKeyDelay, 0
        SetControlDelay oldControlDelay
    }
}
