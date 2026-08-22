; =============================================================================
; AppLaunchers module: launch_hotkeys.ahk
; Chrome, WhatsApp, YouTube, Cursor launch hotkeys
; Extracted verbatim from AppLaunchers.ahk; loaded via #include into the
; AppLaunchers.ahk process, which remains the entry point / source of truth.
; =============================================================================

; =============================================================================
; Open/Activate Google Chrome
; Hotkey: Win+Alt+Shift+F
; Original File: Open Google.ahk
; =============================================================================
#!+f::
{
    Run "chrome.exe"
    WinWait("ahk_exe chrome.exe", , 10)  ; Wait for window to exist (up to 10 seconds)
    Sleep(300)
    if (!WinExist("ahk_exe chrome.exe")) {
        ShowCenteredOverlay_Utils("⚠ Chrome did not start in time.", 2000, BANNER_ACCENT_ERROR)
        return
    }
    WinActivate("ahk_exe chrome.exe")    ; Explicitly activate the window
    WinWaitActive("ahk_exe chrome.exe", , 2)  ; Wait for activation to complete
    ClipAngelBanner_Show("🔍 Checking search bar...", BANNER_ACCENT_INTERMEDIATE)
    try {
        hwnd := WinGetID("ahk_exe chrome.exe")
        root := UIA.ElementFromHandle(hwnd)
        addrBar := root.FindFirst({ Type: 50004, AutomationId: "view_1012" })
        if (addrBar) {
            focused := UIA.GetFocusedElement()
            if (!UIA.CompareElements(addrBar, focused)) {
                try
                    addrBar.SetFocus()
                catch
                    Send "^l"
            }
        }
    } catch {
        ; UIA failed; banner still shows Done below
    }
    ClipAngelBanner_Show("✅ Done", BANNER_ACCENT_SUCCESS)
    SetTimer(ClipAngelBanner_Hide, -500)
    CenterMouse()
}

; =============================================================================
; Open/Activate WhatsApp
; Hotkey: Win+Alt+Shift+Z
; Original File: Open WhatsApp.ahk
; Fast path: Run the .lnk only (focuses existing PWA or cold-starts). No HWND
; scan / UI-ready wait / loading banner — JumpToChat still uses ActivateOrOpen.
; =============================================================================
#!+z::
{
    if (IS_WORK_ENVIRONMENT) {
        Run "C:\Users\fie7ca\Documents\Shortcuts\WhatsApp.lnk"
    } else {
        Run "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\apps do Chrome\WhatsApp Web.lnk"
    }
    CenterMouse()
}

; =============================================================================
; Open/Activate YouTube
; Hotkey: Win+Alt+Shift+H
; Original File: Youtube - Activate.ahk
; =============================================================================
; Session start: **assumes** the watch-page video is paused/stopped. One Send("k") toggles play (YouTube shortcut).
; Trade-off: if the video was already playing, k pauses — use when entering focus with a paused video, or press #!+h again to exit.
; Fast path: one UIA_Browser bound to ytHwnd + GetCurrentURL only (no FindFirst / tree scans). See docs/efficiency-canon.md §11.
YouTube_PlayWhenOpened(ytHwnd := 0) {
    if !(ytHwnd is Integer) || ytHwnd <= 0
        ytHwnd := WinExist("YouTube ahk_exe chrome.exe")
    if !ytHwnd
        return
    if !WinActive("ahk_id " ytHwnd) {
        WinActivate("ahk_id " ytHwnd)
        WinWaitActive("ahk_id " ytHwnd, , 2)
    }
    try {
        uia := UIA_Browser("ahk_id " ytHwnd)
        if !InStr(uia.GetCurrentURL(), "youtube.com/watch")
            return
        Send("k")
    } catch {
        ; UIA/URL unavailable — do not Send(k) blind (wrong-focus risk).
    }
}

#!+h::
{
    ; Preserve and restore title match mode (efficiency-canon: no leaked global state)
    prevTitleMode := A_TitleMatchMode
    global g_YoutubeFocusSessionActive
    try {
        SetTitleMatchMode 2
        if (g_YoutubeFocusSessionActive) {
            YouTube_EndFocusSession()
            return
        }

        YouTube_PauseSpotifyBeforeYoutube()

        ; Prefer focusing an existing YouTube Chrome window to avoid duplicates.
        hwnd := WinExist("YouTube ahk_exe chrome.exe")
        if hwnd {
            WinActivate("ahk_id " hwnd)
            WinWaitActive("ahk_id " hwnd, , 2)
            CenterMouse()
            YouTube_PlayWhenOpened(hwnd)
            StartYoutubeFocusMonitor(hwnd)
            g_YoutubeFocusSessionActive := true
            return
        }
        ; No YouTube window detected: open History URL in a new Chrome window
        ; so it doesn't attach as a tab to an existing instance.
        Run 'chrome.exe --new-window "https://www.youtube.com/feed/history"'
        if WinWaitActive("YouTube ahk_exe chrome.exe", , 10) {
            CenterMouse()
            YouTube_PlayWhenOpened(WinExist("A"))
            hwnd := WinExist("A")
            if hwnd {
                StartYoutubeFocusMonitor(hwnd)
                g_YoutubeFocusSessionActive := true
            }
        }
    } finally {
        try SetTitleMatchMode prevTitleMode
    }
}

; =============================================================================
; Open/Activate Cursor
; Hotkey: Win+Alt+Shift+,
; =============================================================================
#!+,::
{
    SetTitleMatchMode 2
    if WinExist("ahk_exe Cursor.exe") {
        WinActivate
        CenterMouse()
    } else {
        target := IS_WORK_ENVIRONMENT ? "C:\\Users\\fie7ca\\AppData\\Local\\Programs\\cursor\\Cursor.exe" :
            "C:\\Users\\eduev\\AppData\\Local\\Programs\\cursor\\Cursor.exe"
        Run target
        if (!WinWaitActive("ahk_exe Cursor.exe", , 10)) {
            ShowCenteredOverlay_Utils("⚠ Cursor did not start.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        CenterMouse()
    }
}
