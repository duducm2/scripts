; =============================================================================
; Utils module: whatsapp_jump_to_chat.ahk
; Shared Jump-to-Chat logic for WhatsApp (Web / Desktop)
; =============================================================================

; Chrome "WhatsApp Web" PWA app-id from the Start Menu .lnk (--app-id=...).
; Needed because an open chat often sets the window title to the contact name only
; (no "WhatsApp" substring), so WinExist("WhatsApp") misses the running app.
global WHATSAPP_CHROME_APP_ID := "hnpfjngllnobngcgfapefoaidbinmjnm"

WhatsAppJump_ShowLoading(state) {
    StandardLoadingBar_Show(state, BANNER_ACCENT_INTERMEDIATE, {
        fontSize: 17,
        trackActiveMonitor: true
    })
}

WhatsAppJump_UpdateLoading(state) {
    global g_StandardLoadingBarGui
    if IsObject(g_StandardLoadingBarGui)
        StandardLoadingBar_Update(state)
    else
        WhatsAppJump_ShowLoading(state)
}

WhatsAppJump_HideLoading() {
    StandardLoadingBar_Hide(0)
}

; Chrome PWA AUMID is typically Chrome._crx_<app-id>.
WhatsAppJump_GetWindowAppUserModelId(hwnd) {
    if !(hwnd is Integer) || hwnd <= 0
        return ""
    iid := Buffer(16, 0)
    if DllCall("ole32\CLSIDFromString", "wstr", "{886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99}", "ptr", iid)
        return ""
    pstore := 0
    if DllCall("shell32\SHGetPropertyStoreForWindow", "ptr", hwnd, "ptr", iid, "ptr*", &pstore) || !pstore
        return ""
    pk := Buffer(20, 0)
    ; PKEY_AppUserModel_ID = {9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3}, pid 5
    if DllCall("ole32\CLSIDFromString", "wstr", "{9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3}", "ptr", pk) {
        ObjRelease(pstore)
        return ""
    }
    NumPut("uint", 5, pk, 16)
    prop := Buffer(A_PtrSize = 8 ? 24 : 16, 0)
    hr := ComCall(5, pstore, "ptr", pk, "ptr", prop) ; IPropertyStore::GetValue
    ObjRelease(pstore)
    if hr != 0
        return ""
    appId := ""
    try {
        ; VT_LPWSTR = 31; pointer payload starts at offset 8
        if (NumGet(prop, 0, "ushort") = 31) {
            pstr := NumGet(prop, 8, "ptr")
            if pstr
                appId := StrGet(pstr, "UTF-16")
        }
    } catch {
        appId := ""
    }
    DllCall("ole32\PropVariantClear", "ptr", prop)
    return appId
}

WhatsAppJump_IsWhatsAppChromeAppHwnd(hwnd) {
    global WHATSAPP_CHROME_APP_ID
    appId := WhatsAppJump_GetWindowAppUserModelId(hwnd)
    return (appId != "" && InStr(appId, WHATSAPP_CHROME_APP_ID))
}

; Returns HWND of an existing WhatsApp window, or 0.
; Matches: title "WhatsApp", native WhatsApp.exe, or Chrome PWA by AppUserModelID.
WhatsAppJump_FindHwnd() {
    prevTitleMode := A_TitleMatchMode
    try {
        SetTitleMatchMode(2)
        if (hwnd := WinExist("WhatsApp"))
            return hwnd
    } finally {
        SetTitleMatchMode(prevTitleMode)
    }

    if (hwnd := WinExist("ahk_exe WhatsApp.exe"))
        return hwnd

    for hwnd in WinGetList("ahk_exe chrome.exe") {
        try {
            if WhatsAppJump_IsWhatsAppChromeAppHwnd(hwnd)
                return hwnd
        } catch {
        }
    }
    return 0
}

; Bounded wait until a WhatsApp window exists. Returns HWND or 0.
WhatsAppJump_WaitHwnd(timeoutSec := 30) {
    deadline := A_TickCount + (timeoutSec * 1000)
    prevTitleMode := A_TitleMatchMode
    try {
        SetTitleMatchMode(2)
        ; Fast path: title often becomes "WhatsApp" briefly on cold start.
        if WinWait("WhatsApp", , Min(timeoutSec, 3)) {
            if (hwnd := WhatsAppJump_FindHwnd())
                return hwnd
        }
        while (A_TickCount < deadline) {
            if (hwnd := WhatsAppJump_FindHwnd())
                return hwnd
            Sleep 150
        }
        return 0
    } finally {
        SetTitleMatchMode(prevTitleMode)
    }
}

WhatsAppJump_ActivateHwnd(hwnd, timeoutSec := 5) {
    if !(hwnd is Integer) || hwnd <= 0
        return false
    WinActivate("ahk_id " hwnd)
    if WinWaitActive("ahk_id " hwnd, , timeoutSec)
        return true
    WinActivate("ahk_id " hwnd)
    return !!WinWaitActive("ahk_id " hwnd, , 2)
}

; Returns true if WhatsApp is active and ready for keyboard shortcuts.
; Cold-starts settle ~2s after the window appears so the SPA can accept Alt+K.
WhatsAppJump_ActivateOrOpen() {
    global IS_WORK_ENVIRONMENT
    prevTitleMode := A_TitleMatchMode
    try {
        SetTitleMatchMode(2)
        hwnd := WhatsAppJump_FindHwnd()
        if (hwnd) {
            WhatsAppJump_UpdateLoading("⏳ Activating WhatsApp...")
            if !WhatsAppJump_ActivateHwnd(hwnd, 3) {
                WhatsAppJump_HideLoading()
                ShowCenteredOverlay_Utils("❌ Could not activate WhatsApp.", 2000, BANNER_ACCENT_ERROR)
                return false
            }
            WhatsAppJump_HideLoading()
            return true
        }

        WhatsAppJump_UpdateLoading("⏳ Opening WhatsApp...")
        if (IS_WORK_ENVIRONMENT) {
            Run "C:\Users\fie7ca\Documents\Shortcuts\WhatsApp.lnk"
        } else {
            Run "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\apps do Chrome\WhatsApp Web.lnk"
        }

        WhatsAppJump_UpdateLoading("⏳ Waiting for WhatsApp...")
        hwnd := WhatsAppJump_WaitHwnd(30)
        if !hwnd {
            WhatsAppJump_HideLoading()
            ShowCenteredOverlay_Utils("❌ WhatsApp did not start in time.", 2000, BANNER_ACCENT_ERROR)
            return false
        }
        if !WhatsAppJump_ActivateHwnd(hwnd, 5) {
            WhatsAppJump_HideLoading()
            ShowCenteredOverlay_Utils("❌ Could not activate WhatsApp.", 2000, BANNER_ACCENT_ERROR)
            return false
        }
        ; SPA / Chrome App needs time before Alt+K and paste work.
        WhatsAppJump_UpdateLoading("⏳ WhatsApp loading...")
        Sleep 2000
        WinActivate("ahk_id " hwnd)
        Sleep 150
        WhatsAppJump_HideLoading()
        return true
    } finally {
        SetTitleMatchMode(prevTitleMode)
    }
}

; Opens search, pastes contact, Enter to select chat only. Does not paste the message.
; keepBarVisible: when true (D2C [Z]), leave the Loading Indication up for the caller to continue/hide.
WhatsAppJumpToChat(contact, keepBarVisible := false) {
    WhatsAppJump_ShowLoading("⏳ Opening WhatsApp...")
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

        ; ActivateOrOpen hides the bar on success; UpdateLoading recreates it for jump steps.
        WhatsAppJump_UpdateLoading("⏳ Focusing WhatsApp...")
        if (hwnd := WhatsAppJump_FindHwnd()) {
            WhatsAppJump_ActivateHwnd(hwnd, 2)
        }
        Sleep 100

        ; Alt+K is the WhatsApp search shortcut (Shift keys maps Shift+S to this)
        WhatsAppJump_UpdateLoading("⏳ Searching contact...")
        Send "!k"
        Sleep 250

        loop 5 {
            A_Clipboard := ""
            A_Clipboard := contact
            if ClipWait(2) && (A_Clipboard = contact)
                break
            if A_Index = 5 {
                WhatsAppJump_HideLoading()
                ShowCenteredOverlay_Utils("❌ CLIPBOARD ERROR - TRY AGAIN", 3000, BANNER_ACCENT_ERROR)
                return false
            }
            Sleep 100
        }

        Send "^v"
        Sleep 300
        WhatsAppJump_UpdateLoading("⏳ Opening chat...")
        Send "{Enter}"
        ; Let the chat composer take focus before caller pastes the message.
        Sleep 700

        if (!keepBarVisible)
            WhatsAppJump_HideLoading()
        return true
    } finally {
        SetWinDelay oldWinDelay
        SetKeyDelay oldKeyDelay, 0
        SetControlDelay oldControlDelay
    }
}
