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
; false = UIA gate before warm exit (rollback); true = keyboard-only when Desktop Explorer exists
global AL_DESKTOP_WARM_KEYBOARD_ONLY := true

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
; Slow-path UIA cache for Desktop Explorer (Shift+Win+E).
AL_DESKTOP_CACHE := UIA.CreateCacheRequest(["Name", "AutomationId", "BoundingRectangle"], ["Selection", "SelectionItem"])
#include %A_ScriptDir%\Utils.ahk

; Focus mode (#!+Y) and Study Topic (#!+X) need the same process as EnableFocusMode; unregister duplicate Utils hotkeys here.
try Hotkey("#!+Y", "Off")
try Hotkey("#!+X", "Off")

; Quick Update relaunch: volume is scheduled from Utils.ahk /Updated block after success overlay + chime.
if !(A_Args.Length > 0 && A_Args[1] = "/Updated")
    ScheduleApplyScriptMasterVolumeTargetWithRetries()

; -----------------------------------------------------------------------------
; MODULE MAP - AppLaunchers.ahk stays the runnable entry point and #includes each
; module below. Early preamble (feature flags, includes, OnExit, volume schedule)
; stays here. See AppLaunchers/MODULARIZATION_PROGRESS.md for the full module list.
; -----------------------------------------------------------------------------

; --- Global Variables ---

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
; Win+Alt+Shift+N — Context file browser (paste path into active input)
; =============================================================================
#!+n:: {
    ShowContextBrowser()
}

; =============================================================================
; Open/Activate Desktop in Explorer
; Hotkey: Shift+Win+E
; Original File: Open Desktop.ahk
; =============================================================================
+#e::
{
    prevTitleMatchMode := A_TitleMatchMode
    try {
        SetTitleMatchMode 2
        targetHwnd := AL_FindDesktopExplorerWindow()
        hadDesktopHwnd := targetHwnd

        if (targetHwnd) {
            if (!WinExist("ahk_id " targetHwnd)) {
                ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
                return
            }

            AL_DesktopActivateMinimal(targetHwnd)

            if (AL_DESKTOP_WARM_KEYBOARD_ONLY) {
                ; Desktop Explorer already open: ^{Home} is enough; skip UIA/F5 (efficiency-canon §11).
                Send "^{Home}"
                AL_CenterMouseOnHwnd(targetHwnd)
                return
            }

            if (AL_IsFirstDesktopItemAlreadySelected(targetHwnd)) {
                Send "^{Home}"
                AL_CenterMouseOnHwnd(targetHwnd)
                return
            }
        }

        StandardLoadingBar_Show("⏳ Opening Desktop and selecting first file...", BANNER_ACCENT_INTERMEDIATE)

        if (!targetHwnd) {
            target := IS_WORK_ENVIRONMENT ? "C:\Users\fie7ca\Desktop" : "C:\Users\eduev\OneDrive\Desktop"
            Run 'explorer.exe "' target '"'

            ; Wait for window to appear (bounded by deadline; no unbounded waits).
            deadline := A_TickCount + 2000
            while (A_TickCount < deadline) {
                targetHwnd := AL_FindDesktopExplorerWindow()
                if (targetHwnd)
                    break
                Sleep 50
            }
        }

        if (targetHwnd) {
            if (!WinExist("ahk_id " targetHwnd)) {
                ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
                return
            }

            if (hadDesktopHwnd && WinActive("ahk_id " targetHwnd) && WinGetMinMax("ahk_id " targetHwnd) != -1) {
                ; Already activated on warm-path probe; skip duplicate activation.
            } else {
                AL_DesktopActivateAggressive(targetHwnd)
            }

            AL_DesktopWaitForItemsView(targetHwnd, 350)
            Send "^{Up}"
            Sleep 100
            Send "{F5}"

            AL_SelectFirstDesktopItem(targetHwnd)
            AL_CenterMouseOnHwnd(targetHwnd)
        }
    } finally {
        SetTitleMatchMode prevTitleMatchMode
        StandardLoadingBar_Hide(0)
    }
}

AL_FindDesktopExplorerWindow() {
    hwnd := WinExist("Área de Trabalho ahk_class CabinetWClass")
    if (hwnd)
        return hwnd
    return WinExist("Desktop ahk_class CabinetWClass")
}

AL_DesktopActivateMinimal(targetHwnd) {
    if !(targetHwnd is Integer) || targetHwnd <= 0
        return false
    if (WinGetMinMax("ahk_id " targetHwnd) = -1)
        WinRestore("ahk_id " targetHwnd)
    if !WinActive("ahk_id " targetHwnd) {
        WinActivate("ahk_id " targetHwnd)
        WinWaitActive("ahk_id " targetHwnd, , 0.15)
    }
    return true
}

AL_DesktopActivateAggressive(targetHwnd) {
    if !(targetHwnd is Integer) || targetHwnd <= 0
        return false
    if (WinGetMinMax("ahk_id " targetHwnd) = -1)
        WinRestore("ahk_id " targetHwnd)
    WinActivate("ahk_id " targetHwnd)
    if !WinWaitActive("ahk_id " targetHwnd, , 0.2) {
        DllCall("SwitchToThisWindow", "Ptr", targetHwnd, "Int", 1)
        DllCall("SetForegroundWindow", "Ptr", targetHwnd)
        WinActivate("ahk_id " targetHwnd)
    }
    return true
}

AL_DesktopWaitForItemsView(targetHwnd, timeoutMs := 350) {
    if !(targetHwnd is Integer) || targetHwnd <= 0
        return false
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        try {
            root := UIA.ElementFromHandle(targetHwnd)
            if AL_FindExplorerItemsView(root)
                return true
        } catch {
        }
        Sleep 50
    }
    return false
}

AL_CenterMouseOnHwnd(hwnd) {
    if !(hwnd is Integer) || hwnd <= 0
        return false
    rect := Buffer(16, 0)
    if !DllCall("GetWindowRect", "Ptr", hwnd, "Ptr", rect)
        return false
    left := NumGet(rect, 0, "Int")
    top := NumGet(rect, 4, "Int")
    right := NumGet(rect, 8, "Int")
    bottom := NumGet(rect, 12, "Int")
    centerX := left + (right - left) // 2
    centerY := top + (bottom - top) // 2
    DllCall("SetCursorPos", "Int", centerX, "Int", centerY)
    return true
}

AL_GetFirstDesktopListItem(listRoot) {
    if !listRoot
        return 0

    try {
        firstItem := listRoot.FindFirst({ Type: "ListItem", Name: "bill.pdf" })
        if firstItem
            return firstItem
    } catch {
    }

    try {
        firstItem := listRoot.FindFirst({ Type: "ListItem", AutomationId: "0" })
        if firstItem
            return firstItem
    } catch {
    }

    try {
        return listRoot.FindFirst({ Type: "ListItem" })
    } catch {
    }

    return 0
}

AL_GetFirstDesktopListItemBuildCache(listRoot, cacheRequest) {
    if !listRoot || !cacheRequest
        return 0

    try {
        firstItem := listRoot.FindFirstBuildCache(cacheRequest, { Type: "ListItem", Name: "bill.pdf" })
        if firstItem
            return firstItem
    } catch {
    }

    try {
        firstItem := listRoot.FindFirstBuildCache(cacheRequest, { Type: "ListItem", AutomationId: "0" })
        if firstItem
            return firstItem
    } catch {
    }

    try {
        return listRoot.FindFirstBuildCache(cacheRequest, { Type: "ListItem" })
    } catch {
    }

    return 0
}

AL_FindExplorerItemsViewBuildCache(root, cacheRequest) {
    if !root || !cacheRequest
        return 0

    try {
        itemsView := root.FindFirstBuildCache(cacheRequest, { AutomationId: "ItemsView", Type: "List" })
        if itemsView
            return itemsView
    } catch {
    }

    try {
        itemsView := root.FindFirstBuildCache(cacheRequest, { ClassName: "UIItemsView", Type: "List" })
        if itemsView
            return itemsView
    } catch {
    }

    try {
        itemsView := root.FindFirstBuildCache(cacheRequest, { Name: "Items View", Type: "List", matchmode: "Substring" })
        if itemsView
            return itemsView
    } catch {
    }

    return 0
}

AL_DesktopFirstItemIsSelected(firstItem) {
    if !firstItem
        return false
    try {
        if firstItem.CachedSelectionItemPattern.IsSelected
            return true
    } catch {
    }
    try {
        if firstItem.SelectionItemPattern.IsSelected
            return true
    } catch {
    }
    return false
}

AL_UIAElementsMatch(a, b) {
    if !a || !b
        return false
    try {
        if (a.RuntimeId = b.RuntimeId)
            return true
    } catch {
    }
    try {
        if (a.Name != "" && a.Name = b.Name)
            return true
    } catch {
    }
    return false
}

AL_IsFirstDesktopItemAlreadySelected(targetHwnd) {
    if !(targetHwnd is Integer) || targetHwnd <= 0
        return false

    loop 2 {
        try {
            root := UIA.ElementFromHandle(targetHwnd)
            listRoot := AL_FindExplorerItemsView(root)
            if !listRoot
                throw Error("ItemsView not found")

            firstItem := AL_GetFirstDesktopListItem(listRoot)
            if !firstItem
                return false

            try {
                if firstItem.SelectionItemPattern.IsSelected
                    return true
            } catch {
            }

            try {
                if listRoot.GetPropertyValue(UIA.Property.IsSelectionPatternAvailable) {
                    selected := listRoot.SelectionPattern.GetSelection()
                    if (selected.Length = 1 && AL_UIAElementsMatch(selected[1], firstItem))
                        return true
                }
            } catch {
            }
        } catch {
        }

        if (A_Index = 1)
            Sleep 50
    }

    return false
}

AL_SelectFirstDesktopItem(targetHwnd) {
    if !(targetHwnd is Integer) || targetHwnd <= 0
        return false

    Send "^{Home}"

    loop 4 {
        try {
            root := UIA.ElementFromHandleBuildCache(AL_DESKTOP_CACHE, targetHwnd)
            listRoot := AL_FindExplorerItemsViewBuildCache(root, AL_DESKTOP_CACHE)
            if !listRoot {
                Sleep 80
                continue
            }

            try listRoot.SetFocus()

            firstItem := AL_GetFirstDesktopListItemBuildCache(listRoot, AL_DESKTOP_CACHE)
            if (firstItem) {
                if AL_DesktopFirstItemIsSelected(firstItem)
                    return true
                try firstItem.ScrollIntoView()
                try firstItem.Select()
                try firstItem.SetFocus()
                return true
            }
        } catch {
        }

        Sleep 80
    }

    Send "{Home}"
    return false
}

AL_FindExplorerItemsView(root) {
    if !root
        return 0

    try {
        itemsView := root.FindFirst({ AutomationId: "ItemsView", Type: "List" })
        if itemsView
            return itemsView
    }

    try {
        itemsView := root.FindFirst({ ClassName: "UIItemsView", Type: "List" })
        if itemsView
            return itemsView
    }

    try {
        itemsView := root.FindFirst({ Name: "Items View", Type: "List", matchmode: "Substring" })
        if itemsView
            return itemsView
    }

    return 0
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
    global g_AL_InputGuardEscaped, g_AL_hHookKbd, g_WikipediaSelectorActive
    if (nCode >= 0 && wParam = 0x100) {
        vkCode := NumGet(lParam, 0, "UInt")
        if (vkCode = 0x1B) {
            ; Hook returns 1 below without CallNextHookEx — AHK hotkeys never see Esc. Wikipedia modal needs Esc.
            if (g_WikipediaSelectorActive) {
                return DllCall("user32\CallNextHookEx", "Ptr", 0, "Int", nCode, "Ptr", wParam, "Ptr", lParam, "Ptr")
            }
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

; True when the active window is on an AHK monitor where Wikipedia fullscreen scroll restore runs.
; Current layout: monitors 3 and 4 are portrait (1080x1920); both use the same restore path as M3.
IsWindowOnWikipediaScrollRestoreMonitor() {
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
            idx := A_Index
            return (idx = 3 || idx = 4)
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

            ; Try to restore scroll position (only on configured portrait monitors: 3 and 4)
            restoreBanner := ""
            try {
                if (!IsWindowOnWikipediaScrollRestoreMonitor()) {
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
                    ; Portrait monitors (3 and 4, 1080x1920) can cause layout shifts that affect document height
                    Sleep(2500)  ; Increased wait for new window page stabilization

                    ; Get current document height with retry logic and stabilization
                    ; Portrait monitors (3 and 4, 1080x1920): ensure layout is stable
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
                    ; For portrait orientation (1080x1920 on monitors 3 and 4), use precise calculation
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

; Used with g_OnEscapePressed so Utils_GlobalEscapeHandler (I10) closes the modal when *Escape (I0) never fires.
WikipediaSelector_GlobalEscapeCallback(*) {
    WikipediaSelector_Cancel()
}

; GUI message-path Escape (helps when Esc is delivered via the GUI message pump)
WikipediaSelector_GuiEscape(*) {
    WikipediaSelector_Cancel()
}

; Cancel Wikipedia selector (same pattern as AiModelSelector_Cancel in Utils.ahk)
WikipediaSelector_Cancel(*) {
    CleanupWikipediaSelector()
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

; Cleanup Wikipedia selector (mirror AiModelSelector_Close in Utils.ahk)
CleanupWikipediaSelector() {
    global g_WikipediaSelectorActive, g_WikipediaSelectorGui, g_WikipediaSelectorHandlers

    if (!g_WikipediaSelectorActive)
        return

    g_WikipediaSelectorActive := false

    global g_OnEscapePressed
    g_OnEscapePressed := ""

    ; Disable hotkeys (same order as AiModelSelector_Close: digit keys, then Escape, then restore global Escape)
    for handler in g_WikipediaSelectorHandlers {
        try {
            Hotkey(handler.char, "Off")
        } catch {
        }
    }
    try {
        Hotkey("Escape", WikipediaSelector_Cancel, "Off")
    } catch {
    }
    try {
        Hotkey("*Escape", WikipediaSelector_Cancel, "Off")
    } catch {
    }
    Utils_EnsureGlobalEscapeHotkey()

    g_WikipediaSelectorHandlers := []

    ; Close and destroy GUI
    if (IsObject(g_WikipediaSelectorGui)) {
        try {
            g_WikipediaSelectorGui.Destroy()
        } catch {
        }
        g_WikipediaSelectorGui := false
    }
}

; Show Wikipedia selector GUI (mirror ShowAiModelSelector in Utils.ahk)
ShowWikipediaSelector() {
    global g_WikipediaSelectorGui, g_WikipediaSelectorActive, g_WikipediaSelectorHandlers
    global g_WikipediaItems

    ; Don't show if already active (same as ShowAiModelSelector)
    if (g_WikipediaSelectorActive)
        return

    ; LL keyboard hook may still be swallowing keys from a prior scroll-restore guard — Esc must reach hotkeys.
    AL_RemoveInputGuard()

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

    ; Create GUI — same dark modal styling as ShowAiModelSelector (#!+C) in Utils.ahk
    ; +E0x08000000: non-activating so PowerToys Command Palette stays open
    g_WikipediaSelectorGui := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x08000000")
    g_WikipediaSelectorGui.BackColor := "1E1E2E"
    g_WikipediaSelectorGui.MarginX := 20
    g_WikipediaSelectorGui.MarginY := 15

    ; Load completed articles
    completedArticles := LoadCompletedArticles()

    ; Build display text (section headers: "— Label —" like Project Selector)
    displayText := ""
    displayText .= "— Available Articles —`n"
    for i, item in g_WikipediaItems {
        displayText .= "[" . item.char . "] > " . item.title . "`n"
    }

    ; Add History section if there are completed articles
    if (completedArticles.Length > 0) {
        displayText .= "`n"
        displayText .= "— History (Read) —`n"
        for i, article in completedArticles {
            displayText .= "  • " . article . "`n"
        }
    }

    ; Title + separator (match Utils ShowAiModelSelector / CursorTransfer selectors)
    baseWidth := (monitorWidth < 1200) ? 500 : 600
    wikiContentW := baseWidth - 40
    g_WikipediaSelectorGui.SetFont("s14 cCDD6F4 Bold", "Segoe UI")
    g_WikipediaSelectorGui.Add("Text", "w" . wikiContentW . " Center", "📖 Wikipedia Articles")
    g_WikipediaSelectorGui.Add("Text", "w" . wikiContentW . " h1 Background45475A")

    ; Calculate Edit height from line count (footer hint is separate Text controls)
    lineCount := 0
    loop parse, displayText, "`n" {
        lineCount++
    }
    lineHeight := 18
    textControlHeight := lineCount * lineHeight
    minHeight := 120
    maxHeightPercent := (monitorHeight < 800) ? 0.90 : 0.75
    maxHeight := Floor(monitorHeight * maxHeightPercent)
    if (textControlHeight < minHeight)
        textControlHeight := minHeight
    if (textControlHeight > maxHeight)
        textControlHeight := maxHeight

    g_WikipediaSelectorGui.SetFont("s12 cCDD6F4", "Segoe UI")
    g_WikipediaSelectorGui.AddEdit("w" . wikiContentW . " h" . textControlHeight . " ReadOnly VScroll Background313244",
        displayText)

    g_WikipediaSelectorGui.Add("Text", "w" . wikiContentW . " h1 Background45475A y+10")
    g_WikipediaSelectorGui.SetFont("s9 c6C7086", "Segoe UI")
    g_WikipediaSelectorGui.Add("Text", "w" . wikiContentW . " Center", "Press 1–5 | Esc to cancel")

    try {
        g_WikipediaSelectorGui.OnEvent("Escape", WikipediaSelector_GuiEscape)
    } catch {
    }

    ; Measure and center on the active window's monitor (same pattern as ShowAiModelSelector)
    g_WikipediaSelectorGui.Show("AutoSize Hide")
    g_WikipediaSelectorGui.GetPos(&gx, &gy, &gw, &gh)
    cx := monitorLeft + (monitorWidth - gw) // 2
    cy := monitorTop + (monitorHeight - gh) // 2
    g_WikipediaSelectorGui.Show("x" . cx . " y" . cy . " NA")

    g_WikipediaSelectorActive := true

    ; Enable hotkeys for 1–5 and Escape (same order as ShowAiModelSelector: digits first, then Escape)
    ; Clear handlers array
    g_WikipediaSelectorHandlers := []

    ; Enable hotkeys for characters 1-5
    for item in g_WikipediaItems {
        char := item.char
        handler := CreateWikipediaCharHandler(char)
        g_WikipediaSelectorHandlers.Push({ char: char, handler: handler })
        try {
            Hotkey(char, handler, "On")
        } catch {
        }
    }

    ; *Escape: not removed by Utils Hotkey("Escape","Off") (square selector, etc.); CursorTransfer uses the same pattern.
    Hotkey("*Escape", WikipediaSelector_Cancel, "On")

    ; If I0 *Escape never receives the key, I10 Utils_GlobalEscapeHandler still runs — same hook as g_OnEscapePressed (Utils.ahk).
    global g_OnEscapePressed
    g_OnEscapePressed := WikipediaSelector_GlobalEscapeCallback
    Utils_EnsureGlobalEscapeHotkey()
}

; =============================================================================
; SelectWikipediaInHandy() — Opens/closes selector or activates Wikipedia (same pattern as SelectAiModelInHandy in Utils.ahk)
; =============================================================================
SelectWikipediaInHandy() {
    global g_WikipediaSelectorActive
    if (g_WikipediaSelectorActive) {
        WikipediaSelector_Cancel()
        return
    }

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
; Open/Activate Wikipedia
; Hotkey: Win+Alt+Shift+K
; =============================================================================
#!+k::
{
    SelectWikipediaInHandy()
}

; [AppLaunchers module] Pomodoro timer system with CSV logging -> AppLaunchers\pomodoro_timer.ahk
#include %A_ScriptDir%\AppLaunchers\pomodoro_timer.ahk

; [AppLaunchers module] CenterMouse helper on active window -> AppLaunchers\center_mouse.ahk
#include %A_ScriptDir%\AppLaunchers\center_mouse.ahk

; [AppLaunchers module] #!+. Clip Angel paste and favorite flow -> AppLaunchers\hotkey_clipangel.ahk
#include %A_ScriptDir%\AppLaunchers\hotkey_clipangel.ahk
