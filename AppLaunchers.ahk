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

; [AppLaunchers module] Wikipedia scroll position save/load/restore -> AppLaunchers\wikipedia_scroll.ahk
#include %A_ScriptDir%\AppLaunchers\wikipedia_scroll.ahk

; [AppLaunchers module] Wikipedia selector GUI and char handlers -> AppLaunchers\wikipedia_selector.ahk
#include %A_ScriptDir%\AppLaunchers\wikipedia_selector.ahk

; [AppLaunchers module] SelectWikipediaInHandy and #!+k hotkey -> AppLaunchers\wikipedia_entry.ahk
#include %A_ScriptDir%\AppLaunchers\wikipedia_entry.ahk

; [AppLaunchers module] Pomodoro timer system with CSV logging -> AppLaunchers\pomodoro_timer.ahk
#include %A_ScriptDir%\AppLaunchers\pomodoro_timer.ahk

; [AppLaunchers module] CenterMouse helper on active window -> AppLaunchers\center_mouse.ahk
#include %A_ScriptDir%\AppLaunchers\center_mouse.ahk

; [AppLaunchers module] #!+. Clip Angel paste and favorite flow -> AppLaunchers\hotkey_clipangel.ahk
#include %A_ScriptDir%\AppLaunchers\hotkey_clipangel.ahk
