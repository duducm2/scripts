; =============================================================================
; AppLaunchers module: launch_hotkeys.ahk
; Chrome (#!+F tap) / Import Management (#!+F double-tap), WhatsApp, Cursor;
; Win+Alt+Shift+H → Utility Shortcuts Prompts
; Extracted verbatim from AppLaunchers.ahk; loaded via #include into the
; AppLaunchers.ahk process, which remains the entry point / source of truth.
; =============================================================================

; =============================================================================
; Open/Activate Google Chrome  /  Import Management
; Hotkey: Win+Alt+Shift+F
;   1× tap  = Chrome (address bar focus)
;   2× tap  = Import Management (same as Utility Shortcuts [J]; both entry points kept)
; Original File: Open Google.ahk
; Double-tap window: AI_QD_DOUBLE_TAP_MS (400) — matches #!+D / ZMK tap-dance.
; =============================================================================
global g_LaunchF_DoubleTapArmed := false
global g_LaunchF_LastPressTick := 0
global g_LaunchF_DoubleTapTimer := 0

class LaunchF_DoubleTapTimerObj {
    static OnSingleTapTimeout() {
        global g_LaunchF_DoubleTapArmed, g_LaunchF_DoubleTapTimer
        if (!g_LaunchF_DoubleTapArmed)
            return
        g_LaunchF_DoubleTapArmed := false
        g_LaunchF_DoubleTapTimer := 0
        LaunchF_OpenChrome()
    }
}

LaunchF_DisarmDoubleTap() {
    global g_LaunchF_DoubleTapArmed, g_LaunchF_DoubleTapTimer, g_LaunchF_LastPressTick
    g_LaunchF_DoubleTapArmed := false
    g_LaunchF_LastPressTick := 0
    if (g_LaunchF_DoubleTapTimer) {
        SetTimer(g_LaunchF_DoubleTapTimer, 0)
        g_LaunchF_DoubleTapTimer := 0
    }
}

LaunchF_OpenChrome() {
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

#!+f:: {
    global g_LaunchF_DoubleTapArmed, g_LaunchF_LastPressTick, g_LaunchF_DoubleTapTimer

    if !GetKeyState("f", "P")
        return

    thresholdMs := 400
    try thresholdMs := AI_QD_DOUBLE_TAP_MS
    catch {
        thresholdMs := 400
    }

    pressTime := A_TickCount
    elapsed := (g_LaunchF_LastPressTick > 0) ? (pressTime - g_LaunchF_LastPressTick) : 9999
    isSecondTap := g_LaunchF_DoubleTapArmed && elapsed >= 0 && elapsed < thresholdMs

    KeyWait "f"

    if (isSecondTap) {
        LaunchF_DisarmDoubleTap()
        try ImportMgmt_LaunchApp()
        catch {
            ShowCenteredOverlay_Utils("⚠ Import Management unavailable", 2000, BANNER_ACCENT_ERROR)
        }
        return
    }

    if (g_LaunchF_DoubleTapTimer) {
        SetTimer(g_LaunchF_DoubleTapTimer, 0)
        g_LaunchF_DoubleTapTimer := 0
    }
    g_LaunchF_LastPressTick := A_TickCount
    g_LaunchF_DoubleTapArmed := true
    g_LaunchF_DoubleTapTimer := ObjBindMethod(LaunchF_DoubleTapTimerObj, "OnSingleTapTimeout")
    SetTimer(g_LaunchF_DoubleTapTimer, -thresholdMs)
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
; Utility Shortcuts → Prompts (prompt manager)
; Hotkey: Win+Alt+Shift+H
; Same UI as #!+U then [R]; toggles closed if Prompts is already open.
; =============================================================================
#!+h::
{
    ShowHotstringSelector("Prompts")
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
