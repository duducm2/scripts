#Requires AutoHotkey v2.0+
#SingleInstance Force

; -----------------------------------------------------------------------------
; This script consolidates all Application/Website launcher hotkeys.
; -----------------------------------------------------------------------------

; --- AppLauncher polyglot IPC feature flags (default off until phases verified) ---
global AL_USE_DAEMON := false
global AL_USE_MMF_IPC := false
global AL_USE_EVENT_HOOKS := false
global AL_USE_WIKI_FSM := false

; --- Includes ----------------------------------------------------------------
#include %A_ScriptDir%\env.ahk
#include %A_ScriptDir%\aux\AppLauncherIPC.ahk
OnExit(AL_AppLaunchersExit, 1)
AL_AppLaunchersExit(*) {
    AL_RemoveInputGuard()
    AL_UnregisterForegroundHook()
    AL_IPC_Teardown()
}
#include UIA-v2\Lib\UIA.ahk
#include UIA-v2\Lib\UIA_Browser.ahk
#include %A_ScriptDir%\Utils.ahk

; --- Global Variables ---
global DEBUG_LOG_PATH := A_ScriptDir "\.cursor\debug.log"

; Phase 3: WinEvent hook for foreground (replaces 200ms Wikipedia focus polling)
global g_AL_WinEventHookHandle := 0
global g_AL_LastForegroundHwnd := 0

; Phase 4: Safe input guard (replaces BlockInput; escape = Ctrl+Shift+Escape)
global g_AL_InputGuardEscaped := false
global g_AL_hHookKbd := 0
global g_AL_hHookMouse := 0
global g_AL_InputGuardCallbackKbd := 0
global g_AL_InputGuardCallbackMouse := 0
; Phase 4: Wikipedia FSM state (Idle, LaunchRequested, AwaitWindow, AwaitPageReady, AwaitUIAReady, RestoreScroll, Verify, Completed, Failed)
global AL_WikiState := "Idle"

; --- Hotkeys & Functions -----------------------------------------------------

; =============================================================================
; Open/Activate Cursor with specific window requirements
; Hotkey: Win+Alt+Shift+N
; Original File: OneNote - Open.ahk (Updated for Cursor)
; =============================================================================
#!+n::
{
    targetWindow := ""
    fallbackWindow := ""

    ; Phase 2: use daemon ResolveCursorTargets when IPC enabled
    if (AL_USE_DAEMON && AL_USE_MMF_IPC) {
        resp := AL_IPC_Call("ResolveCursorTargets", Map(), 3000)
        if (resp.Has("ok") && resp["ok"]) {
            pHwnd := AL_IPC_GetResultInt(resp, "primaryHwnd")
            fHwnd := AL_IPC_GetResultInt(resp, "fallbackHwnd")
            if (pHwnd)
                targetWindow := "ahk_id " pHwnd
            if (fHwnd)
                fallbackWindow := "ahk_id " fHwnd
        }
    }

    ; Legacy: WinGetList enumeration when daemon off or unavailable
    if (targetWindow = "" && fallbackWindow = "") {
        for proc in ["ahk_exe Cursor.exe", "ahk_exe Code.exe"] {
            for hwnd in WinGetList(proc) {
                try {
                    winTitle := WinGetTitle("ahk_id " hwnd)
                    winTitleLower := StrLower(winTitle)
                    if InStr(winTitleLower, "preview")
                        continue
                    if (!fallbackWindow)
                        fallbackWindow := "ahk_id " hwnd
                    if (InStr(winTitleLower, "habits") || InStr(winTitleLower, "home") || InStr(winTitleLower,
                        "punctual") || InStr(winTitleLower, "work")) {
                        targetWindow := "ahk_id " hwnd
                        break
                    }
                } catch {
                    continue
                }
            }
            if (targetWindow)
                break
        }
    }

    if (targetWindow) {
        if (!WinExist(targetWindow)) {
            ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        WinActivate(targetWindow)
        if WinWaitActive(targetWindow, , 2) {
            CenterMouse()
            Sleep(100)
            Send("^t")
        }
    } else if (fallbackWindow) {
        if (!WinExist(fallbackWindow)) {
            ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        WinActivate(fallbackWindow)
        if WinWaitActive(fallbackWindow, , 2) {
            CenterMouse()
            Sleep(100)
            Send("^t")
        }
    } else {
        ShowCursorFallbackPanel()
    }
}

; =============================================================================
; Show fallback panel when no matching Cursor window is found
; =============================================================================
ShowCursorFallbackPanel() {
    ; Create a notification panel similar to ChatGPT.ahk style
    fallbackGui := Gui()
    fallbackGui.Opt("+AlwaysOnTop -Caption +ToolWindow")
    fallbackGui.BackColor := "3772FF"  ; Blue background like ChatGPT notifications
    fallbackGui.SetFont("s20 cFFFFFF Bold", "Segoe UI")

    ; Create the message
    message := "No Cursor window found with names: habits, home, punctual, or work"
    fallbackGui.Add("Text", "w600 Center", message)

    ; Center on active monitor
    activeWin := WinGetID("A")
    if (activeWin) {
        WinGetPos(&winX, &winY, &winW, &winH, activeWin)
    } else {
        MonitorGetWorkArea(1, &l, &t, &r, &b)
        winX := l
        winY := t
        winW := r - l
        winH := b - t
    }

    fallbackGui.Show("AutoSize Hide")
    guiW := 0, guiH := 0
    fallbackGui.GetPos(, , &guiW, &guiH)

    guiX := winX + (winW - guiW) / 2
    guiY := winY + (winH - guiH) / 2
    fallbackGui.Show("x" . Round(guiX) . " y" . Round(guiY) . " NA")
    WinSetTransparent(178, fallbackGui)

    ; Auto-hide after 3 seconds
    SetTimer(() => fallbackGui.Destroy(), -3000)
}

; =============================================================================
; Open/Activate Desktop in Explorer
; Hotkey: Shift+Win+E
; Original File: Open Desktop.ahk
; =============================================================================
+#e::
{
    SetTitleMatchMode 2
    targetHwnd := 0

    ; Check for existing window (PT or EN)
    if WinExist("Área de Trabalho ahk_class CabinetWClass")
        targetHwnd := WinExist("Área de Trabalho ahk_class CabinetWClass")
    else if WinExist("Desktop ahk_class CabinetWClass")
        targetHwnd := WinExist("Desktop ahk_class CabinetWClass")

    if (!targetHwnd) {
        target := IS_WORK_ENVIRONMENT ? "C:\Users\fie7ca\Desktop" : "C:\Users\eduev\OneDrive\Desktop"
        Run 'explorer.exe "' target '"'

        ; Wait for window to appear
        loop 40 { ; Wait up to 2 seconds
            if WinExist("Área de Trabalho ahk_class CabinetWClass") {
                targetHwnd := WinExist("Área de Trabalho ahk_class CabinetWClass")
                break
            }
            if WinExist("Desktop ahk_class CabinetWClass") {
                targetHwnd := WinExist("Desktop ahk_class CabinetWClass")
                break
            }
            Sleep 50
        }
    }

    if (targetHwnd) {
        if (!WinExist("ahk_id " targetHwnd)) {
            ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        ; Layer 1: Restore if minimized
        if (WinGetMinMax("ahk_id " targetHwnd) = -1) {
            WinRestore("ahk_id " targetHwnd)
        }

        ; Layer 2: Standard Activation
        WinActivate("ahk_id " targetHwnd)

        ; Layer 3: Aggressive Activation if not active immediately
        if !WinWaitActive("ahk_id " targetHwnd, , 0.2) {
            DllCall("SwitchToThisWindow", "Ptr", targetHwnd, "Int", 1)
            DllCall("SetForegroundWindow", "Ptr", targetHwnd)
            WinActivate("ahk_id " targetHwnd)
        }

        WinMaximize("ahk_id " targetHwnd)

        Sleep 350
        Send "^{Up}"
        Sleep 100
        Send "{F5}"

        CenterMouse()
    }
}

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
; =============================================================================
#!+z::
{
    SetTitleMatchMode(2)
    if WinExist("WhatsApp") {
        WinActivate("WhatsApp")
        CenterMouse()
    } else {
        if (IS_WORK_ENVIRONMENT) {
            Run "C:\Users\fie7ca\Documents\Shortcuts\WhatsApp.lnk"
        } else {
            Run "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\apps do Chrome\WhatsApp Web.lnk"
        }
        WinWaitActive("WhatsApp")
        CenterMouse()
    }
}

; =============================================================================
; Open/Activate YouTube
; Hotkey: Win+Alt+Shift+H
; Original File: Youtube - Activate.ahk
; =============================================================================
#!+h::
{
    SetTitleMatchMode 2
    if WinExist("YouTube") {
        WinActivate
        CenterMouse()
    } else {
        Run "chrome.exe --new-window https://www.youtube.com/feed/playlists"
        WinWaitActive("YouTube")
        CenterMouse()
    }
}

; =============================================================================
; Open/Activate Gmail
; Hotkey: Win+Alt+Shift+W
; =============================================================================
#!+w::
{
    SetTitleMatchMode 2
    if WinExist("Gmail ahk_exe chrome.exe") {
        WinActivate
        CenterMouse()
    } else {
        target := IS_WORK_ENVIRONMENT ?
            "C:\Users\fie7ca\Documents\Shortcuts\Gmail.lnk" :
                "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Chrome Apps\Gmail.lnk"
        Run target
        WinWaitActive("Gmail ahk_exe chrome.exe")
        CenterMouse()
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

; =============================================================================
; Wikipedia Selector with Character Shortcuts
; Hotkey: Win+Alt+Shift+K
; Shows a GUI with Wikipedia article options. Pressing a character (1-5)
; immediately opens the corresponding article or performs the action.
; =============================================================================

; Global variables for Wikipedia selector
global g_WikipediaSelectorGui := false
global g_WikipediaSelectorActive := false
global g_WikipediaSelectorHandlers := []  ; Store hotkey handlers for cleanup

; Global variables for Wikipedia scroll position save/restore
global g_WikipediaScrollPositionsFile := A_ScriptDir "\data\wikipedia_scroll_positions.ini"

; Global variable for Wikipedia focus monitoring (automatic blackout cancellation)
global g_WikipediaFocusMonitorTimer := false

; Global variable for Wikipedia completed articles CSV file
global g_WikipediaCompletedFile := A_ScriptDir "\data\wikipedia_completed.csv"

; Wikipedia article items configuration (Taoist philosophy completed and removed)
; Item 1: Claude Debussy
; Item 2: Daoshi
; Item 3: Self-cultivation
; Item 4: Cognitive therapy
; Item 5: Key
global g_WikipediaItems := [{ char: "1", title: "Claude Debussy", url: "https://en.wikipedia.org/wiki/Claude_Debussy" }, { char: "2",
    title: "Daoshi", url: "https://en.wikipedia.org/wiki/Daoshi" }, { char: "3", title: "Self-cultivation",
        url: "https://en.wikipedia.org/wiki/Self-cultivation" }, { char: "4", title: "Cognitive therapy", url: "https://en.wikipedia.org/wiki/Cognitive_therapy" }, { char: "5",
            title: "Key", url: "https://en.wikipedia.org/wiki/Key_(music)" }
]

; =============================================================================
; Wikipedia Focus Monitoring for Automatic Blackout Cancellation
; =============================================================================

; Monitor Wikipedia window focus and automatically disable focus mode when Wikipedia loses focus
MonitorWikipediaFocus() {
    global g_WikipediaFocusMonitorTimer

    ; Check if Wikipedia is still the active window
    SetTitleMatchMode 2
    if (!WinActive("Wikipedia")) {
        ; Wikipedia is no longer active - exit fullscreen and disable focus mode
        ; Send F11 to exit fullscreen mode before disabling focus
        SetTitleMatchMode 2
        if (WinExist("Wikipedia")) {
            ; Get the window handle
            wikipediaHwnd := WinExist("Wikipedia")
            if (wikipediaHwnd) {
                ; Store current active window to restore focus after
                currentActiveHwnd := WinExist("A")

                ; Briefly activate Wikipedia window to send F11
                try {
                    WinActivate("ahk_id " . wikipediaHwnd)
                } catch {
                    ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
                    return
                }
                Sleep(50)  ; Brief delay to ensure window is active
                Send("{F11}")
                Sleep(100)  ; Allow time for fullscreen exit

                ; Restore focus to the window that was previously active
                if (currentActiveHwnd && WinExist("ahk_id " . currentActiveHwnd)) {
                    try {
                        WinActivate("ahk_id " . currentActiveHwnd)
                    } catch {
                        ; Window closed, skip restore
                    }
                }
            }
        }
        ; Disable focus mode and stop monitoring
        DisableFocusMode()
        StopWikipediaFocusMonitor()
    }
}

; Start monitoring Wikipedia window focus
StartWikipediaFocusMonitor() {
    global g_WikipediaFocusMonitorTimer

    StopWikipediaFocusMonitor()

    ; Phase 3: event-driven foreground hook when enabled; else 200ms polling
    if (AL_USE_EVENT_HOOKS) {
        AL_RegisterForegroundHook()
        g_WikipediaFocusMonitorTimer := true  ; flag only; no timer
    } else {
        g_WikipediaFocusMonitorTimer := MonitorWikipediaFocus
        SetTimer(g_WikipediaFocusMonitorTimer, 200)
    }
}

; Stop monitoring Wikipedia window focus
StopWikipediaFocusMonitor() {
    global g_WikipediaFocusMonitorTimer

    if (g_WikipediaFocusMonitorTimer) {
        if (Type(g_WikipediaFocusMonitorTimer) = "Func")
            SetTimer(g_WikipediaFocusMonitorTimer, 0)
        g_WikipediaFocusMonitorTimer := false
    }
}

; Phase 3: Foreground hook callback (runs in hook thread; only schedule main-thread work)
AL_ForegroundHookProc(hHook, event, hwnd, idObject, idChild, idEventThread, dwmsEventTime) {
    global g_AL_LastForegroundHwnd
    g_AL_LastForegroundHwnd := hwnd
    SetTimer(AL_OnWikipediaForegroundChanged, -0)
}

; Phase 3: Main-thread handler when foreground changed (Wikipedia lost focus?)
AL_OnWikipediaForegroundChanged() {
    global g_WikipediaFocusMonitorTimer, g_AL_LastForegroundHwnd
    if (!g_WikipediaFocusMonitorTimer)
        return
    try {
        title := WinGetTitle("ahk_id " g_AL_LastForegroundHwnd)
        if (InStr(title, "Wikipedia"))
            return
    } catch {
        return
    }
    ; Wikipedia lost focus: same logic as MonitorWikipediaFocus()
    SetTitleMatchMode 2
    if (WinExist("Wikipedia")) {
        wikipediaHwnd := WinExist("Wikipedia")
        if (wikipediaHwnd) {
            currentActiveHwnd := WinExist("A")
            try {
                WinActivate("ahk_id " wikipediaHwnd)
            } catch {
                DisableFocusMode()
                StopWikipediaFocusMonitor()
                return
            }
            Sleep(50)
            Send("{F11}")
            Sleep(100)
            if (currentActiveHwnd && WinExist("ahk_id " currentActiveHwnd)) {
                try
                    WinActivate("ahk_id " currentActiveHwnd)
                catch Any {
                    ; window closed, skip restore
                }
            }
        }
    }
    DisableFocusMode()
    StopWikipediaFocusMonitor()
}

AL_RegisterForegroundHook() {
    global g_AL_WinEventHookHandle
    if (g_AL_WinEventHookHandle)
        return
    ; EVENT_SYSTEM_FOREGROUND = 0x0003, WINEVENT_OUTOFCONTEXT = 0
    cb := CallbackCreate(AL_ForegroundHookProc, "F", 7)
    h := DllCall("user32\SetWinEventHook", "UInt", 0x0003, "UInt", 0x0003, "Ptr", 0, "Ptr", cb, "UInt", 0, "UInt", 0,
        "UInt", 0, "Ptr")
    if (h) {
        g_AL_WinEventHookHandle := h
        ; Keep callback alive (store in global so not freed)
        global g_AL_ForegroundHookCallback := cb
    }
}

AL_UnregisterForegroundHook() {
    global g_AL_WinEventHookHandle
    if (g_AL_WinEventHookHandle) {
        DllCall("user32\UnhookWinEvent", "Ptr", g_AL_WinEventHookHandle)
        g_AL_WinEventHookHandle := 0
    }
}

; Phase 4: Low-level input guard (replaces BlockInput); Ctrl+Shift+Escape = emergency escape
AL_InstallInputGuard() {
    global g_AL_InputGuardEscaped, g_AL_hHookKbd, g_AL_hHookMouse, g_AL_InputGuardCallbackKbd,
        g_AL_InputGuardCallbackMouse
    g_AL_InputGuardEscaped := false
    if (g_AL_hHookKbd)
        return
    g_AL_InputGuardCallbackKbd := CallbackCreate(AL_InputGuardKeyboardProc, "F", 4)
    g_AL_InputGuardCallbackMouse := CallbackCreate(AL_InputGuardMouseProc, "F", 4)
    g_AL_hHookKbd := DllCall("user32\SetWindowsHookEx", "Int", 13, "Ptr", g_AL_InputGuardCallbackKbd, "Ptr", 0, "UInt",
        0, "Ptr")
    g_AL_hHookMouse := DllCall("user32\SetWindowsHookEx", "Int", 14, "Ptr", g_AL_InputGuardCallbackMouse, "Ptr", 0,
        "UInt", 0, "Ptr")
}

AL_RemoveInputGuard() {
    global g_AL_hHookKbd, g_AL_hHookMouse
    if (g_AL_hHookKbd) {
        DllCall("user32\UnhookWindowsHookEx", "Ptr", g_AL_hHookKbd)
        g_AL_hHookKbd := 0
    }
    if (g_AL_hHookMouse) {
        DllCall("user32\UnhookWindowsHookEx", "Ptr", g_AL_hHookMouse)
        g_AL_hHookMouse := 0
    }
}

AL_InputGuardKeyboardProc(nCode, wParam, lParam) {
    global g_AL_InputGuardEscaped, g_AL_hHookKbd
    if (nCode >= 0 && wParam = 0x100) {
        vkCode := NumGet(lParam, 0, "UInt")
        if (vkCode = 0x1B) {
            if (DllCall("user32\GetAsyncKeyState", "Int", 0x11) & 0x8000 && DllCall("user32\GetAsyncKeyState", "Int",
                0x10) & 0x8000) {
                g_AL_InputGuardEscaped := true
                AL_RemoveInputGuard()
                return DllCall("user32\CallNextHookEx", "Ptr", 0, "Int", nCode, "Ptr", wParam, "Ptr", lParam, "Ptr")
            }
        }
        return 1
    }
    return DllCall("user32\CallNextHookEx", "Ptr", 0, "Int", nCode, "Ptr", wParam, "Ptr", lParam, "Ptr")
}

AL_InputGuardMouseProc(nCode, wParam, lParam) {
    if (nCode >= 0)
        return 1
    return DllCall("user32\CallNextHookEx", "Ptr", 0, "Int", nCode, "Ptr", wParam, "Ptr", lParam, "Ptr")
}

; =============================================================================
; Wikipedia Scroll Position Storage Functions
; =============================================================================

; Check if the active window is on Monitor 3
IsWindowOnMonitor3() {
    hwnd := WinExist("A")

    if (!hwnd) {
        return false
    }

    rect := Buffer(16, 0)
    if (!DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect)) {
        return false
    }

    left := NumGet(rect, 0, "int")
    top := NumGet(rect, 4, "int")
    right := NumGet(rect, 8, "int")
    bottom := NumGet(rect, 12, "int")

    centerX := left + (right - left) // 2
    centerY := top + (bottom - top) // 2

    monitorCount := MonitorGetCount()
    loop monitorCount {
        MonitorGet(A_Index, &ml, &mt, &mr, &mb)
        if (centerX >= ml && centerX <= mr && centerY >= mt && centerY <= mb) {
            isMonitor3 := (A_Index = 3)
            return isMonitor3
        }
    }

    return false
}

; Helper function to normalize Wikipedia URLs
NormalizeWikipediaURL(url) {
    if (url = "" || !InStr(url, "wikipedia.org")) {
        return ""
    }
    ; Remove fragments and trailing slashes
    url := RegExReplace(url, "/#.*$", "")
    url := RegExReplace(url, "/+$", "")
    return url
}

; Get current Wikipedia article URL from the active Chrome window
GetWikipediaURL() {
    try {
        if (!WinActive("ahk_exe chrome.exe")) {
            return ""
        }

        winTitle := WinGetTitle("A")
        if (!InStr(winTitle, "Wikipedia")) {
            return ""
        }

        uia := UIA_Browser("ahk_exe chrome.exe")
        if (!uia) {
            return ""
        }

        url := uia.GetCurrentURL()
        normalizedUrl := NormalizeWikipediaURL(url)
        return normalizedUrl
    } catch Error as err {
        return ""
    }
}

; Helper function to restore scroll position to a given percentage
; Returns true on success, false on failure
RestoreWikipediaScrollPosition(scrollPercentage, bannerText := "Restoring scroll position... Please wait") {
    if (scrollPercentage <= 0.0 || scrollPercentage > 1.0) {
        return false
    }

    try {
        StandardLoadingBar_Show(bannerText, BANNER_ACCENT_INTERMEDIATE)
        Sleep(10)

        uia := UIA_Browser("ahk_exe chrome.exe")
        if (!uia) {
            StandardLoadingBar_Hide(0)
            return false
        }

        ; Block input during restoration
        AL_InstallInputGuard()

        ; Wait for page to be ready
        Sleep(500)

        ; Get document height
        docHeight := uia.JSReturnThroughClipboard("document.documentElement.scrollHeight")
        if (docHeight = "" || docHeight = "undefined" || docHeight = "null") {
            AL_RemoveInputGuard()
            StandardLoadingBar_Hide(0)
            return false
        }

        docHeightFloat := Float(docHeight)
        if (docHeightFloat <= 0) {
            AL_RemoveInputGuard()
            StandardLoadingBar_Hide(0)
            return false
        }

        ; Calculate and execute scroll
        targetScrollY := scrollPercentage * docHeightFloat
        uia.JSExecute("window.scrollTo(0, " . Round(targetScrollY) . ");")
        Sleep(500)

        AL_RemoveInputGuard()
        StandardLoadingBar_Update("Scroll position restored!")
        StandardLoadingBar_Hide(500)
        return true
    } catch Error as err {
        AL_RemoveInputGuard()
        StandardLoadingBar_Hide(0)
        return false
    }
}

; Save scroll position for a Wikipedia article URL
; Now saves as percentage (0.0 to 1.0) instead of absolute pixels
SaveWikipediaScrollPosition(url, scrollPercentage) {
    global g_WikipediaScrollPositionsFile
    try {
        if (url = "" || scrollPercentage = "" || scrollPercentage < 0 || scrollPercentage > 1) {
            return false
        }
        ; Normalize URL to match load format - remove trailing slashes and fragments
        normalizedUrl := RegExReplace(url, "/#.*$", "")
        normalizedUrl := RegExReplace(normalizedUrl, "/+$", "")
        ; Ensure directory exists
        SplitPath(g_WikipediaScrollPositionsFile, , &dir)
        if (dir != "" && !DirExist(dir)) {
            DirCreate(dir)
        }
        ; Read existing entries first (before deleting file) to preserve them
        existingEntries := Map()
        if (FileExist(g_WikipediaScrollPositionsFile)) {
            try {
                ; Read all existing entries from the Positions section
                ; We'll read the file manually to handle UTF-16 encoding issues
                fileContent := FileRead(g_WikipediaScrollPositionsFile)
                ; Parse INI format manually
                inPositionsSection := false
                loop parse fileContent, "`n", "`r" {
                    line := Trim(A_LoopField)
                    if (line = "[Positions]") {
                        inPositionsSection := true
                        continue
                    }
                    if (inPositionsSection && SubStr(line, 1, 1) = "[") {
                        ; Hit another section, stop reading
                        break
                    }
                    if (inPositionsSection && InStr(line, "=")) {
                        pos := InStr(line, "=")
                        key := Trim(SubStr(line, 1, pos - 1))
                        value := Trim(SubStr(line, pos + 1))
                        if (key != "" && value != "") {
                            existingEntries[key] := value
                        }
                    }
                }
            } catch {
                ; If read fails, try IniRead as fallback
                try {
                    ; Get all keys in Positions section (this is a workaround)
                    ; We'll just update the one we need
                } catch {
                }
            }
        }

        ; Update with new entry
        existingEntries[normalizedUrl] := scrollPercentage

        ; Delete file to recreate in UTF-8
        if (FileExist(g_WikipediaScrollPositionsFile)) {
            try {
                FileDelete(g_WikipediaScrollPositionsFile)
                Sleep(100)  ; Small delay to ensure file system updates
            } catch {
            }
        }

        ; Write all entries back in UTF-8 encoding
        try {
            ; Write UTF-8 BOM and section header
            FileAppend("[Positions]`r`n", g_WikipediaScrollPositionsFile, "UTF-8")
            ; Write each entry
            for key, value in existingEntries {
                ; Escape special INI characters in key and value
                escapedKey := StrReplace(key, "=", "`=")
                escapedKey := StrReplace(escapedKey, ";", "`;")
                escapedValue := StrReplace(value, "`n", "`;")
                escapedValue := StrReplace(escapedValue, "`r", "")
                FileAppend(escapedKey . "=" . escapedValue . "`r`n", g_WikipediaScrollPositionsFile, "UTF-8")
            }
        } catch {
            ; Fallback to IniWrite if manual write fails
            IniWrite(scrollPercentage, g_WikipediaScrollPositionsFile, "Positions", normalizedUrl)
        }
        return true
    } catch Error as err {
        return false
    }
}

; Load saved scroll position for a Wikipedia article URL
; Returns percentage (0.0 to 1.0) instead of absolute pixels
LoadWikipediaScrollPosition(url) {
    global g_WikipediaScrollPositionsFile
    try {
        if (url = "") {
            return 0.0
        }

        ; Verify INI file path is set
        if (!g_WikipediaScrollPositionsFile) {
            return 0.0
        }

        ; Normalize URL to match save format - remove trailing slashes and fragments
        normalizedUrl := RegExReplace(url, "/#.*$", "")
        normalizedUrl := RegExReplace(normalizedUrl, "/+$", "")

        ; Ensure directory exists (in case it was deleted)
        SplitPath(g_WikipediaScrollPositionsFile, , &dir)
        if (dir != "" && !DirExist(dir)) {
            DirCreate(dir)
        }

        ; Read from INI file
        ; Try manual parsing first (handles UTF-8 BOM and encoding issues better)
        scrollPos := "0"
        try {
            if (FileExist(g_WikipediaScrollPositionsFile)) {
                fileContent := FileRead(g_WikipediaScrollPositionsFile)
                ; Parse INI format manually
                inPositionsSection := false
                loop parse fileContent, "`n", "`r" {
                    line := Trim(A_LoopField)
                    ; Skip empty lines and comments
                    if (line = "" || SubStr(line, 1, 1) = ";") {
                        continue
                    }
                    if (line = "[Positions]") {
                        inPositionsSection := true
                        continue
                    }
                    if (inPositionsSection && SubStr(line, 1, 1) = "[") {
                        ; Hit another section, stop reading
                        break
                    }
                    if (inPositionsSection && InStr(line, "=")) {
                        pos := InStr(line, "=")
                        key := Trim(SubStr(line, 1, pos - 1))
                        value := Trim(SubStr(line, pos + 1))
                        ; Unescape special characters that were escaped during save
                        key := StrReplace(key, "`=", "=")
                        key := StrReplace(key, "`;", ";")
                        ; Compare normalized URLs (case-sensitive for Wikipedia URLs)
                        if (key = normalizedUrl && value != "" && value != "0") {
                            scrollPos := value
                            break
                        }
                    }
                }
            }
        } catch {
            ; If manual parsing fails, fall back to IniRead
            try {
                scrollPos := IniRead(g_WikipediaScrollPositionsFile, "Positions", normalizedUrl, "0")
            } catch {
                scrollPos := "0"
            }
        }

        scrollPercentage := Float(scrollPos)
        return scrollPercentage
    } catch Error as err {
        return 0.0
    }
}

; Handler for character key press
HandleWikipediaChar(char) {
    global g_WikipediaSelectorActive, g_WikipediaItems

    ; Only process if selector is active
    if (!g_WikipediaSelectorActive) {
        return
    }

    ; Find the item for this character
    item := ""
    for i, itm in g_WikipediaItems {
        if (itm.char = char) {
            item := itm
            break
        }
    }

    if (item) {
        ; Cleanup first (closes GUI, disables hotkeys)
        CleanupWikipediaSelector()

        ; If item has a URL, open it in Chrome in a new window
        if (item.url != "") {
            Run "chrome.exe --new-window " item.url
            ; Wait for the window to appear and become active
            WinWait("ahk_exe chrome.exe", , 5)
            Sleep(500)  ; Give the page a moment to start loading

            ; Wait for the page to load (check for Wikipedia in title)
            SetTitleMatchMode 2
            if (!WinWait("Wikipedia", , 10)) {
                ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
                return
            }
            WinActivate("Wikipedia")
            WinWaitActive("Wikipedia", , 5)

            ; Check if page is ready by attempting to get URL
            ; This ensures the page has loaded before proceeding
            ; For new windows, we need more time for the page to fully load
            pageReady := false
            loadingRetries := 8  ; Increased retries for new windows
            loop loadingRetries {
                try {
                    url := GetWikipediaURL()
                    if (url != "" && InStr(url, "wikipedia.org")) {
                        ; Verify the page is actually interactive, not just loaded
                        ; Try to access UIA to ensure the page is ready for automation
                        try {
                            testUia := UIA_Browser("ahk_exe chrome.exe")
                            if (testUia) {
                                ; Try to get document height to verify page is fully interactive
                                ; This is the same operation we'll need for scroll restoration
                                testDocHeight := testUia.JSReturnThroughClipboard(
                                    "document.documentElement.scrollHeight")
                                if (testDocHeight != "" && testDocHeight != "undefined" && testDocHeight != "null") {
                                    testHeightFloat := Float(testDocHeight)
                                    if (testHeightFloat > 0) {
                                        ; Page is ready and UIA can access it
                                        pageReady := true
                                        break
                                    }
                                }
                            }
                        } catch {
                            ; UIA not ready yet, continue waiting
                        }
                    }
                } catch {
                    ; URL not accessible yet, page may still be loading
                }
                if (A_Index < loadingRetries) {
                    Sleep(500)  ; Longer wait for new windows
                }
            }

            ; If page wasn't ready after retries, wait a bit more for loading
            ; For new windows, we need extra time for all resources to load
            if (!pageReady) {
                Sleep(1000)  ; Additional delay for new window page loading
            } else {
                ; Even if page seems ready, give it a moment for layout to stabilize
                Sleep(800)  ; Additional stabilization time for new windows
            }

            ; Enter fullscreen mode once page is ready
            Send("{F11}")
            Sleep(300)  ; Allow time for fullscreen transition (increased for new windows)

            ; Enable focus mode to darken other monitors
            EnableFocusMode()

            ; Start monitoring Wikipedia focus for automatic blackout cancellation
            StartWikipediaFocusMonitor()

            ; Try to restore scroll position (only if on Monitor 3)
            restoreBanner := ""
            try {
                if (!IsWindowOnMonitor3()) {
                    return
                }
                savedPercentage := LoadWikipediaScrollPosition(item.url)
                if (savedPercentage > 0.0) {
                    ; Exit fullscreen before scroll restoration (REQUIRED: UIA unreliable in fullscreen)
                    Send("{F11}")
                    Sleep(300)  ; Allow time for fullscreen exit

                    StandardLoadingBar_Show("📜 Restoring scroll position... Please wait", BANNER_ACCENT_INTERMEDIATE)
                    AL_InstallInputGuard()

                    ; Initialize UIA_Browser with retry logic
                    ; For new windows, UIA needs more time to initialize and attach to the browser
                    uia := false
                    uiaRetries := 5  ; Increased retries for new windows
                    loop uiaRetries {
                        try {
                            uia := UIA_Browser("ahk_exe chrome.exe")
                            if (uia) {
                                ; Verify UIA can actually access the page (not just initialized)
                                ; Try a simple operation to ensure the connection is ready
                                try {
                                    testUrl := uia.GetCurrentURL()
                                    if (testUrl != "" && InStr(testUrl, "wikipedia.org")) {
                                        ; UIA is ready and can access the page
                                        break
                                    }
                                } catch {
                                    ; UIA initialized but not ready yet, continue retrying
                                    uia := false
                                }
                            }
                        } catch Error as uiaErr {
                            ; UIA initialization failed, will retry
                        }
                        if (A_Index < uiaRetries) {
                            Sleep(800)  ; Longer wait for new windows (UIA initialization takes time)
                        }
                    }

                    if (!uia) {
                        AL_RemoveInputGuard()
                        StandardLoadingBar_Update("Error: Could not access browser")
                        StandardLoadingBar_Hide(2000)
                        Send("{F11}")
                        Sleep(300)
                        return
                    }

                    ; Wait longer for page to be ready and stabilize (critical for portrait orientation)
                    ; For new windows, the page needs more time to fully render and become interactive
                    ; Portrait orientation (1080x1920) can cause layout shifts that affect document height
                    Sleep(2500)  ; Increased wait for new window page stabilization

                    ; Get current document height with retry logic and stabilization
                    ; Monitor 3 is portrait (1080x1920), so we need to ensure layout is stable
                    ; For new windows, we need more retries and longer waits
                    docHeight := ""
                    docHeightRetries := 8  ; Increased retries for new windows
                    lastDocHeight := 0
                    stableCount := 0
                    loop docHeightRetries {
                        try {
                            docHeight := uia.JSReturnThroughClipboard("document.documentElement.scrollHeight")
                            if (docHeight != "" && docHeight != "undefined" && docHeight != "null") {
                                docHeightFloat := Float(docHeight)
                                ; For new windows, require 3 consecutive stable readings (more strict)
                                if (docHeightFloat = lastDocHeight) {
                                    stableCount++
                                    if (stableCount >= 3) {
                                        ; Document height is stable, use it
                                        break
                                    }
                                } else {
                                    stableCount := 0
                                    lastDocHeight := docHeightFloat
                                }
                            }
                        } catch Error as docErr {
                            if (A_Index < docHeightRetries) {
                                Sleep(600)  ; Longer wait for new windows
                            }
                        }
                        if (A_Index < docHeightRetries) {
                            Sleep(400)  ; Longer wait between measurements for new windows
                        }
                    }

                    if (docHeight = "" || docHeight = "undefined" || docHeight = "null") {
                        AL_RemoveInputGuard()
                        StandardLoadingBar_Update("Error: Page not ready")
                        StandardLoadingBar_Hide(2000)
                        Send("{F11}")
                        Sleep(300)
                        return
                    }

                    docHeightFloat := Float(docHeight)
                    if (docHeightFloat <= 0) {
                        AL_RemoveInputGuard()
                        StandardLoadingBar_Update("Error: Invalid page height")
                        StandardLoadingBar_Hide(2000)
                        Send("{F11}")
                        Sleep(300)
                        return
                    }

                    ; Execute scroll restoration
                    ; For portrait orientation (1080x1920 on Monitor 3), use precise calculation
                    targetScrollY := savedPercentage * docHeightFloat
                    try {
                        ; Use precise scrolling for portrait orientation (requires precise pixel positioning)
                        scrollCommand := "window.scrollTo({top: " . targetScrollY . ", behavior: 'instant'});"
                        uia.JSExecute(scrollCommand)
                        Sleep(1000)  ; Increased wait for portrait orientation (layout may need more time)

                        ; Verify scroll position was applied correctly with retry
                        ; Portrait orientation may require multiple verification attempts
                        verificationRetries := 3
                        actualScrollYFloat := 0
                        scrollDiff := 999999
                        loop verificationRetries {
                            try {
                                actualScrollY := uia.JSReturnThroughClipboard("window.pageYOffset")
                                actualScrollYFloat := Float(actualScrollY)
                                scrollDiff := Abs(actualScrollYFloat - targetScrollY)
                                ; If difference is small enough (within 2 pixels for portrait), consider it successful
                                if (scrollDiff <= 2.0) {
                                    break
                                }
                            } catch {
                            }
                            if (A_Index < verificationRetries) {
                                Sleep(300)  ; Wait before retry
                            }
                        }

                        ; If scroll is significantly off, try to correct it
                        if (scrollDiff > 5.0) {
                            ; Re-scroll to correct position
                            uia.JSExecute(scrollCommand)
                            Sleep(500)
                            ; Verify again
                            try {
                                actualScrollY := uia.JSReturnThroughClipboard("window.pageYOffset")
                                actualScrollYFloat := Float(actualScrollY)
                                scrollDiff := Abs(actualScrollYFloat - targetScrollY)
                            } catch {
                            }
                        }

                        AL_RemoveInputGuard()
                        StandardLoadingBar_Update("Scroll position restored!")
                        StandardLoadingBar_Hide(1000)

                        ; Re-enter fullscreen after successful scroll restoration
                        Send("{F11}")
                        Sleep(300)  ; Allow time for fullscreen transition
                    } catch Error as scrollErr {
                        AL_RemoveInputGuard()
                        StandardLoadingBar_Update("Error: Scroll failed")
                        StandardLoadingBar_Hide(2000)
                        Send("{F11}")
                        Sleep(300)
                    }
                }
            } catch Error as err {
                AL_RemoveInputGuard()
                StandardLoadingBar_Update("Error: " . SubStr(err.Message, 1, 50))
                StandardLoadingBar_Hide(2000)
                ; Re-enter fullscreen after error
                try {
                    Send("{F11}")
                    Sleep(300)
                } catch {
                }
            }
        }
        ; Items 2-5 have no URL, so no action is taken
    }
}

; Factory function to create a handler that properly captures the character
CreateWikipediaCharHandler(char) {
    ; Return a function that captures the char value at creation time
    return (*) => HandleWikipediaChar(char)
}

; Handler for Escape key
HandleWikipediaEscape(*) {
    global g_WikipediaSelectorActive
    if (g_WikipediaSelectorActive) {
        CleanupWikipediaSelector()
    }
}

; Load completed Wikipedia articles from CSV file
LoadCompletedArticles() {
    global g_WikipediaCompletedFile
    completedArticles := []

    try {
        if (!FileExist(g_WikipediaCompletedFile)) {
            return completedArticles
        }

        fileContent := FileRead(g_WikipediaCompletedFile)
        lines := StrSplit(fileContent, "`n")

        ; Skip header line and process each line
        loop lines.Length {
            if (A_Index = 1) {
                continue  ; Skip header
            }

            line := Trim(lines[A_Index])
            if (line != "") {
                completedArticles.Push(line)
            }
        }
    } catch Error as err {
        ; Return empty array on error
        return completedArticles
    }

    return completedArticles
}

; Cleanup Wikipedia selector
CleanupWikipediaSelector() {
    global g_WikipediaSelectorActive, g_WikipediaSelectorGui, g_WikipediaSelectorHandlers

    ; Disable active flag
    g_WikipediaSelectorActive := false

    ; Disable all character hotkeys
    for handler in g_WikipediaSelectorHandlers {
        try {
            char := handler.char
            Hotkey(char, "Off")
        } catch {
            ; Silently ignore errors
        }
    }

    ; Disable Escape hotkey
    try {
        Hotkey("Escape", "Off")
    } catch {
        ; Ignore
    }

    ; Clear handlers array
    g_WikipediaSelectorHandlers := []

    ; Close and destroy GUI
    if (IsObject(g_WikipediaSelectorGui)) {
        try {
            g_WikipediaSelectorGui.Destroy()
        } catch {
            ; Ignore
        }
        g_WikipediaSelectorGui := false
    }
}

; Show Wikipedia selector GUI
ShowWikipediaSelector() {
    global g_WikipediaSelectorGui, g_WikipediaSelectorActive, g_WikipediaSelectorHandlers
    global g_WikipediaItems

    ; Close existing GUI if open
    if (g_WikipediaSelectorActive && IsObject(g_WikipediaSelectorGui)) {
        CleanupWikipediaSelector()
        Sleep 50
    }

    ; Get monitor dimensions early for responsive sizing
    activeWin := 0
    try {
        activeWin := WinGetID("A")
    } catch {
        activeWin := 0
    }

    ; Default to primary monitor work area
    MonitorGetWorkArea(1, &monitorLeft, &monitorTop, &monitorRight, &monitorBottom)
    monitorWidth := monitorRight - monitorLeft
    monitorHeight := monitorBottom - monitorTop

    ; If we have an active window, find which monitor contains its center
    if (activeWin && activeWin != 0) {
        rect := Buffer(16, 0)
        if (DllCall("GetWindowRect", "ptr", activeWin, "ptr", rect)) {
            ; Calculate window center
            winLeft := NumGet(rect, 0, "int")
            winTop := NumGet(rect, 4, "int")
            winRight := NumGet(rect, 8, "int")
            winBottom := NumGet(rect, 12, "int")

            centerX := winLeft + (winRight - winLeft) // 2
            centerY := winTop + (winBottom - winTop) // 2

            ; Find which monitor contains the window center
            monitorCount := MonitorGetCount()
            loop monitorCount {
                idx := A_Index
                MonitorGetWorkArea(idx, &l, &t, &r, &b)
                if (centerX >= l && centerX <= r && centerY >= t && centerY <= b) {
                    monitorLeft := l
                    monitorTop := t
                    monitorRight := r
                    monitorBottom := b
                    monitorWidth := r - l
                    monitorHeight := b - t
                    break
                }
            }
        }
    }

    ; Create GUI
    ; Create non-activating GUI so PowerToys Command Palette stays open
    g_WikipediaSelectorGui := Gui("+AlwaysOnTop +ToolWindow +E0x08000000", "Wikipedia Articles")
    ; Use slightly smaller font for better fit on small monitors
    fontSize := (monitorHeight < 800) ? 9 : 10
    g_WikipediaSelectorGui.SetFont("s" . fontSize, "Segoe UI")
    g_WikipediaSelectorGui.MarginX := 10
    g_WikipediaSelectorGui.MarginY := 5

    ; Load completed articles
    completedArticles := LoadCompletedArticles()

    ; Build display text
    displayText := ""
    displayText .= "Available Articles:`n"
    for i, item in g_WikipediaItems {
        displayText .= "[" . item.char . "] > " . item.title . "`n"
    }

    ; Add History section if there are completed articles
    if (completedArticles.Length > 0) {
        displayText .= "`n─────────────────────────`n"
        displayText .= "History (Read):`n"
        for i, article in completedArticles {
            displayText .= "  • " . article . "`n"
        }
    }

    displayText .= "`nPress Escape to cancel."

    ; Calculate text control height based on actual content (number of lines)
    lineCount := 1  ; Start at 1 (first line doesn't have a newline before it)
    loop parse, displayText, "`n" {
        lineCount++
    }
    ; Calculate height: ~16 pixels per line
    lineHeight := 16
    textControlHeight := lineCount * lineHeight
    ; Ensure minimum and maximum bounds
    minHeight := 150
    maxHeightPercent := (monitorHeight < 800) ? 0.90 : 0.75
    maxHeight := Floor(monitorHeight * maxHeightPercent)
    if (textControlHeight < minHeight)
        textControlHeight := minHeight
    if (textControlHeight > maxHeight)
        textControlHeight := maxHeight

    ; Make width responsive to monitor size
    baseWidth := (monitorWidth < 1200) ? 500 : 600
    textControlWidth := baseWidth - 20  ; Account for margins

    ; Enable vertical scrolling for long content
    g_WikipediaSelectorGui.AddEdit("w" . textControlWidth . " h" . textControlHeight . " ReadOnly VScroll", displayText
    )

    ; Add Close button (set as default so it gets focus, not the Edit control)
    closeBtn := g_WikipediaSelectorGui.AddButton("w100 Default Center", "Close")
    closeBtn.OnEvent("Click", (*) => CleanupWikipediaSelector())

    ; Calculate total height: margins + text control + button + spacing
    totalHeight := 10 + textControlHeight + 40 + 10  ; margins + content + button + spacing
    guiWidth := baseWidth

    ; Calculate center position for the GUI with margins
    marginX := 20  ; Horizontal margin from screen edges
    marginY := 20  ; Vertical margin from screen edges
    guiX := monitorLeft + (monitorWidth - guiWidth) // 2
    guiY := monitorTop + (monitorHeight - totalHeight) // 2

    ; Ensure the GUI stays within monitor bounds with margins
    if (guiX < monitorLeft + marginX)
        guiX := monitorLeft + marginX
    if (guiY < monitorTop + marginY)
        guiY := monitorTop + marginY
    if (guiX + guiWidth > monitorLeft + monitorWidth - marginX)
        guiX := monitorLeft + monitorWidth - guiWidth - marginX
    if (guiY + totalHeight > monitorTop + monitorHeight - marginY)
        guiY := monitorTop + monitorHeight - totalHeight - marginY

    ; Show GUI centered on the active window's monitor
    g_WikipediaSelectorGui.Show("NA w" . guiWidth . " h" . totalHeight . " x" . guiX . " y" . guiY)

    ; Set active flag
    g_WikipediaSelectorActive := true

    ; Clear handlers array
    g_WikipediaSelectorHandlers := []

    ; Enable hotkeys for characters 1-5
    for item in g_WikipediaItems {
        char := item.char
        ; Use factory function to create handler with properly captured char value
        handler := CreateWikipediaCharHandler(char)

        ; Store handler for cleanup
        g_WikipediaSelectorHandlers.Push({ char: char, handler: handler })

        ; Enable hotkey
        try {
            Hotkey(char, handler, "On")
        } catch {
            ; Silently ignore if we can't create hotkey for this character
        }
    }

    ; Enable Escape hotkey
    Hotkey("Escape", HandleWikipediaEscape, "On")
}

; =============================================================================
; Open/Activate Wikipedia
; Hotkey: Win+Alt+Shift+K
; =============================================================================
#!+k::
{
    SetTitleMatchMode 2
    if WinExist("Wikipedia") {
        ; Window already exists - use reduced delay for activation
        windowAlreadyOpen := true
        WinActivate
        WinWaitActive("Wikipedia", , 2)
        ; Ensure Chrome is active (Wikipedia windows are Chrome windows)
        WinWaitActive("ahk_exe chrome.exe", , 2)
        Sleep(100)  ; Reduced delay since window is already open
        CenterMouse()
    } else {
        ; Window doesn't exist yet - it will be created by selector
        ; This path is handled by the selector logic below
        windowAlreadyOpen := false
    }

    ; If window was activated (already existed), proceed with fullscreen setup
    if (windowAlreadyOpen) {
        ; Check for saved scroll position immediately to prioritize feedback
        url := GetWikipediaURL()
        savedPercentage := 0.0
        if (url != "") {
            savedPercentage := LoadWikipediaScrollPosition(url)
        }

        if (savedPercentage > 0.0) {
            StandardLoadingBar_Show("📜 Restoring scroll position... Please wait", BANNER_ACCENT_INTERMEDIATE)
            AL_InstallInputGuard()

            ; Initialize UIA_Browser with retry logic
            uia := false
            uiaRetries := 3
            loop uiaRetries {
                try {
                    uia := UIA_Browser("ahk_exe chrome.exe")
                    if (uia) {
                        break
                    }
                } catch Error as uiaErr {
                    if (A_Index < uiaRetries) {
                        Sleep(200)
                    }
                }
            }

            if (!uia) {
                AL_RemoveInputGuard()
                StandardLoadingBar_Update("Error: Could not access browser")
                StandardLoadingBar_Hide(1000)
                Send("{F11}")
                Sleep(300)
            } else {
                ; Wait for page stability
                Sleep(500)

                ; Get current document height with retry logic and stabilization
                docHeight := ""
                docHeightRetries := 5
                lastDocHeight := 0
                stableCount := 0
                loop docHeightRetries {
                    try {
                        docHeight := uia.JSReturnThroughClipboard("document.documentElement.scrollHeight")
                        if (docHeight != "" && docHeight != "undefined" && docHeight != "null") {
                            docHeightFloat := Float(docHeight)
                            if (docHeightFloat = lastDocHeight) {
                                stableCount++
                                if (stableCount >= 2) {
                                    break
                                }
                            } else {
                                stableCount := 0
                                lastDocHeight := docHeightFloat
                            }
                        }
                    } catch Error as docErr {
                        if (A_Index < docHeightRetries) {
                            Sleep(200)
                        }
                    }
                    if (A_Index < docHeightRetries) {
                        Sleep(200)
                    }
                }

                if (docHeight != "" && docHeight != "undefined" && docHeight != "null") {
                    docHeightFloat := Float(docHeight)
                    if (docHeightFloat > 0) {
                        targetScrollY := savedPercentage * docHeightFloat
                        try {
                            scrollCommand := "window.scrollTo({top: " . targetScrollY . ", behavior: 'instant'});"
                            uia.JSExecute(scrollCommand)
                            Sleep(500)

                            ; Verify scroll position
                            verificationRetries := 3
                            loop verificationRetries {
                                try {
                                    actualScrollY := uia.JSReturnThroughClipboard("window.pageYOffset")
                                    actualScrollYFloat := Float(actualScrollY)
                                    if (Abs(actualScrollYFloat - targetScrollY) <= 5.0) {
                                        break
                                    }
                                } catch {
                                }
                                if (A_Index < verificationRetries) {
                                    Sleep(200)
                                }
                            }

                            StandardLoadingBar_Update("Scroll position restored!")
                        } catch Error as scrollErr {
                            StandardLoadingBar_Update("Error: Scroll failed")
                        }
                    }
                } else {
                    StandardLoadingBar_Update("Error: Page not ready")
                }

                AL_RemoveInputGuard()
                Send("{F11}")
                Sleep(300)
                StandardLoadingBar_Hide(1000)
            }
        } else {
            ; No saved position, just enter fullscreen
            Send("{F11}")
            Sleep(200)
        }

        ; Enable focus mode to darken other monitors
        EnableFocusMode()

        ; Start monitoring Wikipedia focus for automatic blackout cancellation
        StartWikipediaFocusMonitor()
    } else {
        ShowWikipediaSelector()
    }
}

; =============================================================================
; Pomodoro Timer System - Local Timer with CSV Logging
; Hotkey: Win+Alt+Shift+9
; =============================================================================

; Global variables for Pomodoro timer management
global g_PomodoroTimer := false
global g_ChimeTimer := false
global g_ChimeStopTimer := false
global g_PomodoroOverlay := false
global g_PomodoroTinyIndicator := false
global g_PomodoroLogFile := A_ScriptDir "\data\pomodoro_log.csv"
global g_PomodoroCount := 0  ; Track Pomodoro count in work environment

; Show water bottle image overlay as hydration reminder
ShowWaterBottleOverlay() {
    imagePath := ""
    for name in ["water-bottle.jpg"] {
        candidate := A_ScriptDir "\pictures\" name
        if FileExist(candidate) {
            imagePath := candidate
            break
        }
    }

    overlay := Gui()
    overlay.Opt("+AlwaysOnTop -Caption +ToolWindow +E0x20")
    if (imagePath != "") {
        overlay.Add("Picture", "w240 h240 Center", imagePath)
    } else {
        overlay.SetFont("s36", "Segoe UI")
        overlay.Add("Text", "Center cRed", "Pomodoro")
        overlay.BackColor := "FFFFFF"
    }
    overlay.Show("AutoSize Center")
    return overlay
}

; Show periodic TrayTip notification during pomodoro (more reliable than overlay)
ShowTinyWaterBottleIndicator() {
    ; No initial notification - only periodic reminders
    ; The water bottle overlay is shown separately

    ; Return a dummy object to maintain compatibility
    ; The periodic notifications will be handled by a timer
    return { Hwnd: 0, Destroy: () => {} }
}

; Log Pomodoro session to CSV file
LogPomodoroSession() {
    global g_PomodoroLogFile, IS_WORK_ENVIRONMENT

    ; Suppress CSV logging in work environment
    if (IS_WORK_ENVIRONMENT) {
        return
    }

    ; Ensure data directory exists
    SplitPath(g_PomodoroLogFile, , &dir)
    if (dir != "" && !DirExist(dir)) {
        DirCreate(dir)
    }

    ; Check if file exists, if not create with headers
    if (!FileExist(g_PomodoroLogFile)) {
        FileAppend("Date,Time`n", g_PomodoroLogFile)
    }

    ; Get current date and time
    currentDate := FormatTime(, "yyyy/MM/dd")
    currentTime := FormatTime(, "HH:mm")

    ; Append entry to CSV
    FileAppend(currentDate . "," . currentTime . "`n", g_PomodoroLogFile)
}

; Check pomodoro status from last CSV entry
CheckPomodoroStatus() {
    global g_PomodoroLogFile

    ; Check if log file exists
    if (!FileExist(g_PomodoroLogFile)) {
        result := MsgBox("No pomodoro records found.`n`nWould you like to start a new Pomodoro?",
            "Pomodoro Status", "YesNo Icon?")
        if (result = "Yes") {
            StartPomodoroTimer()
        }
        return
    }

    ; Read the CSV file
    try {
        fileContent := FileRead(g_PomodoroLogFile)
        lines := StrSplit(fileContent, "`n")

        ; Find the last non-empty line (skip header and empty lines)
        lastLine := ""
        loop lines.Length {
            idx := lines.Length - A_Index + 1
            line := Trim(lines[idx])
            if (line != "" && line != "Date,Time" && InStr(line, ",")) {
                lastLine := line
                break
            }
        }

        if (lastLine = "") {
            result := MsgBox("No pomodoro records found.`n`nWould you like to start a new Pomodoro?",
                "Pomodoro Status", "YesNo Icon?")
            if (result = "Yes") {
                StartPomodoroTimer()
            }
            return
        }

        ; Parse the last entry
        parts := StrSplit(lastLine, ",")
        if (parts.Length < 2) {
            result := MsgBox("Invalid pomodoro record format.`n`nWould you like to start a new Pomodoro?",
                "Pomodoro Status", "YesNo Icon?")
            if (result = "Yes") {
                StartPomodoroTimer()
            }
            return
        }

        lastDate := Trim(parts[1])
        lastTime := Trim(parts[2])

        ; Parse date and time
        ; Format: yyyy/MM/dd and HH:mm
        dateTimeStr := lastDate . " " . lastTime
        currentDateTimeStr := FormatTime(, "yyyy/MM/dd HH:mm")

        ; Calculate time difference in minutes
        timeDiffMinutes := CalculateMinutesDifference(dateTimeStr, currentDateTimeStr)

        ; Check if calculation failed
        ; timeDiffMinutes = 0 means same minute (just started), which is valid
        ; Negative means calculation error or future date (shouldn't happen)
        if (timeDiffMinutes < 0) {
            result := MsgBox("Could not calculate time difference.`n`nWould you like to start a new Pomodoro?",
                "Pomodoro Status", "YesNo Icon?")
            if (result = "Yes") {
                StartPomodoroTimer()
            }
            return
        }

        ; timeDiffMinutes = 0 means pomodoro was just started (within same minute) - this is valid

        ; Check if probably in pomodoro (within 25 minutes)
        probablyInPomodoro := (timeDiffMinutes >= 0 && timeDiffMinutes <= 25)

        ; Build message
        statusMsg := "Last Pomodoro:`n"
        statusMsg .= "Date: " . lastDate . "`n"
        statusMsg .= "Time: " . lastTime . "`n"
        statusMsg .= "Time ago: " . Round(timeDiffMinutes) . " minutes`n`n"

        if (probablyInPomodoro) {
            statusMsg .= "✅ You are PROBABLY in a Pomodoro session."
        } else {
            statusMsg .= "❌ You are PROBABLY NOT in a Pomodoro session.`n`n"
            statusMsg .= "Would you like to start a new Pomodoro?"
        }

        if (probablyInPomodoro) {
            MsgBox(statusMsg, "Pomodoro Status", "Iconi")
        } else {
            result := MsgBox(statusMsg, "Pomodoro Status", "YesNo Icon?")
            if (result = "Yes") {
                StartPomodoroTimer()
            }
        }

    } catch Error as err {
        result := MsgBox("Error reading pomodoro log: " . err.Message . "`n`nWould you like to start a new Pomodoro?",
            "Pomodoro Status", "YesNo Icon?")
        if (result = "Yes") {
            StartPomodoroTimer()
        }
    }
}

; Helper function to calculate minutes difference between two date/time strings
CalculateMinutesDifference(dateTimeStr1, dateTimeStr2) {
    ; Parse format: "yyyy/MM/dd HH:mm"
    ; Calculate difference in minutes
    try {
        ; Parse both date/time strings
        time1 := ParseDateTimeToMinutes(dateTimeStr1)
        time2 := ParseDateTimeToMinutes(dateTimeStr2)

        if (time1 = 0 || time2 = 0) {
            return 0
        }

        return time2 - time1
    } catch Error as err {
        return 0
    }
}

; Helper to convert date/time string to total minutes since a reference point
ParseDateTimeToMinutes(dateTimeStr) {
    try {
        parts := StrSplit(dateTimeStr, " ")
        if (parts.Length < 2) {
            return 0
        }

        datePart := parts[1]  ; "yyyy/MM/dd"
        timePart := parts[2]  ; "HH:mm"

        ; Split date components
        dateComponents := StrSplit(datePart, "/")
        if (dateComponents.Length < 3) {
            return 0
        }

        year := Integer(dateComponents[1])
        month := Integer(dateComponents[2])
        day := Integer(dateComponents[3])

        ; Split time components
        timeComponents := StrSplit(timePart, ":")
        if (timeComponents.Length < 2) {
            return 0
        }

        hour := Integer(timeComponents[1])
        minute := Integer(timeComponents[2])

        ; More accurate: use days since year 2000
        daysSince2000 := CalculateDaysSince2000(year, month, day)
        totalMinutes := daysSince2000 * 1440 + hour * 60 + minute

        return totalMinutes
    } catch Error as err {
        return 0
    }
}

; Calculate days since January 1, 2000
CalculateDaysSince2000(year, month, day) {
    ; Simple calculation: approximate days
    ; More accurate would require handling leap years, but for our use case (25 minute window) this is sufficient
    days := 0

    ; Days from 2000 to year-1
    if (year > 2000) {
        loop (year - 2000) {
            yearNum := 2000 + A_Index - 1
            days += IsLeapYear(yearNum) ? 366 : 365
        }
    } else if (year < 2000) {
        ; Handle years before 2000 (shouldn't happen for pomodoro logs, but handle gracefully)
        loop (2000 - year) {
            yearNum := 2000 - A_Index
            days -= IsLeapYear(yearNum) ? 366 : 365
        }
    }

    ; Days from Jan 1 to month-1 in current year
    monthDays := [0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    if (IsLeapYear(year)) {
        monthDays[3] := 29  ; February has 29 days in leap year
    }

    loop (month - 1) {
        days += monthDays[A_Index]
    }

    ; Add days in current month
    days += day - 1

    return days
}

; Check if year is a leap year
IsLeapYear(year) {
    return (Mod(year, 4) = 0 && Mod(year, 100) != 0) || (Mod(year, 400) = 0)
}

; Play chime callback - plays sound every 1 second with multiple methods for maximum audibility
PomodoroChimeCallback(*) {
    ; Play multiple sounds simultaneously for maximum audibility (if enabled)
    if (!IsSoundEnabled()) {
        return
    }

    ; Method 1: Primary - SoundBeep with high frequency and longer duration (most reliable and audible)
    try {
        SoundBeep(2000, 300)  ; High frequency (2000 Hz) and longer duration (300 ms) for better audibility
    } catch {
    }

    ; Method 2: Also try MessageBeep as additional sound
    try {
        DllCall("User32\MessageBeep", "UInt", 0xFFFFFFFF)
    } catch {
    }

    ; Method 3: Also try system sound as additional alert
    try {
        SoundPlay("*16")  ; System asterisk sound
    } catch {
    }
}

; Stop chime callback - stops both chime timers
PomodoroStopChimeCallback(*) {
    global g_ChimeTimer, g_ChimeStopTimer
    if (g_ChimeTimer) {
        SetTimer(g_ChimeTimer, 0)
        g_ChimeTimer := false
    }
    if (g_ChimeStopTimer) {
        SetTimer(g_ChimeStopTimer, 0)
        g_ChimeStopTimer := false
    }
}

; Auto-hide Pomodoro overlay after 5 seconds
PomodoroHideOverlayCallback(*) {
    global g_PomodoroOverlay
    if (g_PomodoroOverlay && IsObject(g_PomodoroOverlay) && g_PomodoroOverlay.Hwnd) {
        try {
            g_PomodoroOverlay.Destroy()
        } catch {
        }
        g_PomodoroOverlay := false
    }
}

; Play completion chime for specified duration
PlayCompletionChime(durationMs) {
    global g_ChimeTimer, g_ChimeStopTimer

    ; Cancel any existing chime timers
    if (g_ChimeTimer) {
        SetTimer(g_ChimeTimer, 0)
        g_ChimeTimer := false
    }
    if (g_ChimeStopTimer) {
        SetTimer(g_ChimeStopTimer, 0)
        g_ChimeStopTimer := false
    }

    ; Play immediate sound when timer completes (before starting periodic chime)
    PomodoroChimeCallback()

    ; Start chime timer (every 1 second for better audibility)
    g_ChimeTimer := PomodoroChimeCallback
    SetTimer(g_ChimeTimer, 1000)

    ; Set timer to stop chime after duration
    g_ChimeStopTimer := PomodoroStopChimeCallback
    SetTimer(g_ChimeStopTimer, -durationMs)
}

; Handler when Pomodoro timer completes (25 minutes)
OnPomodoroComplete() {
    global g_PomodoroTimer, g_PomodoroOverlay, g_PomodoroTinyIndicator, g_ChimeTimer, g_ChimeStopTimer

    ; Cancel the main timer
    if (g_PomodoroTimer) {
        SetTimer(g_PomodoroTimer, 0)
        g_PomodoroTimer := false
    }

    ; Hide tiny water bottle indicator when timer completes
    if (g_PomodoroTinyIndicator && IsObject(g_PomodoroTinyIndicator)) {
        try {
            if (g_PomodoroTinyIndicator.Hwnd) {
                g_PomodoroTinyIndicator.Destroy()
            }
        } catch {
        }
        ; Clear any pending tray notifications
        TrayTip()  ; Clear tray tip
        g_PomodoroTinyIndicator := false
    }

    ; Play 30-second completion chime (plays immediate sound, then every 1 second for 30 seconds)
    ; The chime will continue playing even while the message box is shown
    PlayCompletionChime(30000)

    ; Show completion message box immediately (this is blocking, but chime continues in background)
    result := MsgBox("Pomodoro session complete!`n`nTrigger another Pomodoro?", "Pomodoro Complete",
        "YesNo Icon?")

    ; Stop chime when message box is dismissed (works for both Yes and No)
    if (g_ChimeTimer) {
        SetTimer(g_ChimeTimer, 0)
        g_ChimeTimer := false
    }
    if (g_ChimeStopTimer) {
        SetTimer(g_ChimeStopTimer, 0)
        g_ChimeStopTimer := false
    }

    ; If user wants to trigger another Pomodoro, start it
    if (result = "Yes") {
        StartPomodoroTimer()
    }
}

; Start a new Pomodoro timer session
StartPomodoroTimer() {
    global g_PomodoroTimer, g_PomodoroOverlay, g_PomodoroTinyIndicator, g_PomodoroCount, IS_WORK_ENVIRONMENT
    ; Cancel any existing timer
    if (g_PomodoroTimer) {
        SetTimer(g_PomodoroTimer, 0)
        g_PomodoroTimer := false
    }

    ; Increment Pomodoro count in work environment, otherwise log to CSV
    if (IS_WORK_ENVIRONMENT) {
        g_PomodoroCount++
    } else {
        LogPomodoroSession()
    }

    ; Show water bottle image overlay (large, auto-hides after 5 seconds)
    if (g_PomodoroOverlay && IsObject(g_PomodoroOverlay) && g_PomodoroOverlay.Hwnd) {
        try {
            g_PomodoroOverlay.Destroy()
        } catch {
        }
    }
    g_PomodoroOverlay := ShowWaterBottleOverlay()

    ; Auto-hide large overlay after 5 seconds
    SetTimer(PomodoroHideOverlayCallback, -5000)

    ; Show tiny water bottle indicator (periodic TrayTip notifications)
    if (g_PomodoroTinyIndicator && IsObject(g_PomodoroTinyIndicator)) {
        try {
            if (g_PomodoroTinyIndicator.Hwnd) {
                g_PomodoroTinyIndicator.Destroy()
            }
        } catch {
        }
    }
    g_PomodoroTinyIndicator := ShowTinyWaterBottleIndicator()

    ; Play start sound (if enabled)
    try {
        if (IsSoundEnabled()) {
            SoundPlay(A_ScriptDir . "\sounds\pomodo-start.wav")
        }
    } catch {
    }

    ; Set up 25-minute completion timer (1,500,000 ms = 25 minutes)
    g_PomodoroTimer := OnPomodoroComplete
    SetTimer(g_PomodoroTimer, -1500000)
}

; =============================================================================
; Pomodoro Timer - Hotkey: Win+Alt+Shift+9
; Quick press: Start pomodoro timer
; Long press (2 seconds): Check pomodoro status
; =============================================================================
#!+9::
{
    ; Record press time
    static pressTime := 0
    pressTime := A_TickCount

    ; Wait for key release or timeout (1 second for long press)
    KeyWait("9", "T1")

    holdTime := A_TickCount - pressTime

    if (holdTime >= 1000) {
        ; Long press (1+ seconds) - check pomodoro status
        CheckPomodoroStatus()
    } else {
        ; Quick press - start pomodoro timer
        StartPomodoroTimer()
    }
}

; =============================================================================
; Helper function to center mouse on the active window
; =============================================================================
CenterMouse() {
    Sleep(200)
    Send("#!+q")
}

; =============================================================================
; Send specific key combinations
; Hotkey: Win+Alt+Shift+.
; =============================================================================
#!+.::
{
    Sleep(100)
    Send("^c")
    Sleep(200)
    Send("!v")
    Sleep(700)
    Send("!q")
    Sleep(200)
    SendEscape()
}

; =============================================================================
; Initialize Wikipedia scroll position auto-save timer - REMOVED
; =============================================================================
; Auto-save timer removed - now using manual save via Shift keys.ahk shortcut
