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
    ScriptSoundBeep(2000, 300)

    ; Method 2: Also try MessageBeep as additional sound
    ScriptMessageBeep(0xFFFFFFFF)

    ; Method 3: Also try system sound as additional alert
    ScriptSoundPlaySystem("*16")
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
    try ScriptSoundPlay(A_ScriptDir . "\sounds\pomodo-start.wav")
    catch {
    }

    ; Set up 25-minute completion timer (1,500,000 ms = 25 minutes)
    g_PomodoroTimer := OnPomodoroComplete
    SetTimer(g_PomodoroTimer, -1500000)
}

; NOTE: Win+Alt+Shift+8 is reserved for Gemini pronunciation/translation.
; Do not bind #!+8 in this file.

; =============================================================================
; AIB Allow watcher (multi-instance)
; Trigger sources:
; 1) VSCode_SubmitChat in Utils.ahk calls AIB_StartAllowWatcher("chat_submit", hwnd)
; 2) Start implementation probe detects active-window transition: present -> absent
; =============================================================================
global g_AIB_AllowWatcherActive := false
global g_AIB_AllowWatcherTimer := ""
global g_AIB_AllowWatcherStateByHwnd := Map()
global g_AIB_AllowWatcherPromptLock := false
global g_AIB_AllowWatcherDecision := ""
global g_AIB_AllowWatcherSource := ""
global g_AIB_AllowWatcherWindowScope := "ide"
global g_AIB_AllowWatcherPinnedHwnd := 0
global g_AIB_AllowWatcherStartedTick := 0
global g_AIB_FlowBannerHwnd := 0
global g_AIB_FlowBannerLastMonitorIdx := 0
global g_AIB_AllowWatcherTimeoutMs := 8 * 60 * 1000
global g_AIB_AllowWatcherRoundRobinOffset := 0
global g_AIB_AllowWatcherDebugTraceEnabled := false
global g_AIB_AllowWatcherDebugLogFile := A_ScriptDir . "\\docs\\aib-allow-runtime-debug.log"
global g_AIB_AllowWatcherRestrictToVSCode := true
global g_AIB_AllowWatcherPersistentMode := false
global g_AIB_AllowWatcherEnableOcrFallback := false
global g_AIB_AllowWatcherOcrCooldownMs := 1800
global g_AIB_AllowWatcherLastOcrTickByHwnd := Map()
global g_AIB_AllowWatcherOcrScriptPath := A_ScriptDir . "\\tools\\Get-AllowPromptOcr.ps1"
global g_AIB_StartImplProbeTimer := ""
global g_AIB_StartImplPrevPresentByHwnd := Map()
global g_AIB_RapidFireTestSourceDir := ""
global g_AIB_RapidFireTestTargetDir := ""
global g_AIB_StartBannerCooldownMs := 900
global g_AIB_LastStartBannerTickByKey := Map()

; Keep this lightweight timer always on so manual "Start implementation" clicks can arm the watcher.
g_AIB_StartImplProbeTimer := AIB_StartImplementationProbeTick
SetTimer(g_AIB_StartImplProbeTimer, 900)

if (AIB_HasLaunchArg("/StartPersistentAllowWatcher"))
    AIB_ArmAllowWatcher("persistent_startup", 0, "vscode", false, false)

AIB_StartAllowWatcher(triggerSource := "manual", targetHwnd := 0, windowScope := "ide", pinToTarget := false) {
    global g_AIB_AllowWatcherActive, g_AIB_AllowWatcherStartedTick, g_AIB_AllowWatcherSource
    global g_AIB_AllowWatcherStateByHwnd, g_AIB_AllowWatcherTimer, g_AIB_AllowWatcherWindowScope
    global g_AIB_AllowWatcherPinnedHwnd, g_AIB_AllowWatcherPersistentMode
    global g_AIB_AllowWatcherPromptLock, g_AIB_AllowWatcherRoundRobinOffset

    if (windowScope = "")
        windowScope := "ide"
    windowScope := AIB_NormalizeAllowWatcherScope(windowScope)
    if (AIB_IsPersistentAllowWatcherTrigger(triggerSource))
        g_AIB_AllowWatcherPersistentMode := true

    wasActive := g_AIB_AllowWatcherActive
    if (!g_AIB_AllowWatcherActive) {
        g_AIB_AllowWatcherActive := true
        g_AIB_AllowWatcherStartedTick := A_TickCount
        g_AIB_AllowWatcherSource := triggerSource
        g_AIB_AllowWatcherWindowScope := windowScope
        g_AIB_AllowWatcherPinnedHwnd := 0
        g_AIB_AllowWatcherPromptLock := false
        g_AIB_AllowWatcherRoundRobinOffset := 0
        if (g_AIB_AllowWatcherTimer)
            SetTimer(g_AIB_AllowWatcherTimer, 0)
        g_AIB_AllowWatcherTimer := AIB_AllowWatcherTick
        SetTimer(g_AIB_AllowWatcherTimer, 350)

        AIB_AllowDebug_StartSession(triggerSource, windowScope, 0)
    }

    windows := []
    if (targetHwnd && WinExist("ahk_id " targetHwnd)) {
        windows.Push(targetHwnd)
    } else {
        windows := AIB_GetWatcherWindowHwnds(windowScope)
    }

    for hwnd in windows
        AIB_AllowWatcherEnsureState(hwnd, triggerSource, windowScope)

    AIB_AllowDebug_Write("rearm source=" triggerSource " scope=" windowScope " target=" targetHwnd " sessions=" g_AIB_AllowWatcherStateByHwnd
        .Count)

    if (g_AIB_AllowWatcherStateByHwnd.Count = 0 && wasActive && !AIB_IsPersistentAllowWatcherMode())
        AIB_StopAllowWatcher("", false)
}

; Callable by name from Utils.ahk (avoids #Warn UseUnsetLocal for AIB_StartAllowWatcher).
AIB_StartAllowWatcher_Bridge(triggerSource := "manual", targetHwnd := 0) {
    return AIB_ArmAllowWatcher(triggerSource, targetHwnd, "vscode", true, true)
}

AIB_NormalizeAllowWatcherScope(windowScope := "ide") {
    global g_AIB_AllowWatcherRestrictToVSCode

    if (windowScope = "")
        windowScope := "ide"

    if (g_AIB_AllowWatcherRestrictToVSCode) {
        if (windowScope = "cursor")
            return "vscode"
        if (windowScope = "ide")
            return "vscode"
    }

    return windowScope
}

AIB_HasLaunchArg(argNeedle) {
    for arg in A_Args {
        if (arg = argNeedle)
            return true
    }
    return false
}

AIB_IsPersistentAllowWatcherTrigger(triggerSource) {
    return (triggerSource = "persistent_startup" || triggerSource = "persistent_hotkey")
}

AIB_IsPersistentAllowWatcherMode() {
    global g_AIB_AllowWatcherPersistentMode
    return !!g_AIB_AllowWatcherPersistentMode
}

AIB_IsCursorComposerFocused(hwnd) {
    global UIA
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false

    ; Fast path: current focused element is already the Cursor composer input.
    try {
        fe := UIA.GetFocusedElement()
        if (fe) {
            cls := ""
            ctype := 0
            try cls := fe.ClassName
            try ctype := fe.ControlType
            if (ctype = 50004 && InStr(cls, "aislash-editor-input"))
                return true
        }
    } catch {
    }

    ; Fallback path: scan visible edit controls in the active Cursor window.
    try {
        root := UIA.ElementFromHandle(hwnd)
        if (!root)
            return false
        edits := root.FindAll({ Type: 50004 })
        for editEl in edits {
            cls := ""
            try cls := editEl.ClassName
            if (!InStr(cls, "aislash-editor-input"))
                continue
            try {
                if (editEl.HasKeyboardFocus)
                    return true
            } catch {
            }
        }
    } catch {
    }
    return false
}

#HotIf WinActive("ahk_exe Code.exe") || WinActive("ahk_exe Cursor.exe")
$Enter::
{
    global g_AIB_AllowWatcherActive

    ; Preserve modified-enter semantics (Shift/Ctrl/Alt/Win variants).
    if (GetKeyState("Shift", "P") || GetKeyState("Ctrl", "P") || GetKeyState("Alt", "P") || GetKeyState("LWin", "P") ||
    GetKeyState("RWin", "P")) {
        SendInput "{Enter}"
        return
    }

    hwnd := 0
    proc := ""
    try hwnd := WinGetID("A")
    catch {
        hwnd := 0
    }
    if (hwnd) {
        try proc := StrLower(WinGetProcessName("ahk_id " hwnd))
    }

    if (proc = "code.exe") {
        preSendReady := false
        try preSendReady := VSCode_IsChatSendReady(hwnd)

        if (VSCode_IsChatInputFocused(hwnd)) {
            AIB_AllowDebug_Write("enter-trigger detected proc=code hwnd=" hwnd)
            if (VSCode_SubmitChat(hwnd)) {
                AIB_AllowDebug_Write("enter-trigger fallback-arm proc=code hwnd=" hwnd)
                AIB_ArmAllowWatcher("chat_submit", hwnd, "vscode", false, true)
                return
            }
            AIB_AllowDebug_Write("enter-trigger submit-failed proc=code hwnd=" hwnd)
        }

        ; Workspace-safe fallback: if send transitioned ready->not-ready after Enter,
        ; infer chat submit and arm watcher even when focus gate misses.
        SendInput "{Enter}"
        Sleep(80)
        postSendReady := true
        try postSendReady := VSCode_IsChatSendReady(hwnd)
        if (preSendReady && !postSendReady) {
            AIB_AllowDebug_Write("enter-trigger inferred-submit proc=code hwnd=" hwnd)
            AIB_ArmAllowWatcher("chat_submit", hwnd, "vscode", false, true)
            return
        }
        return
    }

    if (proc = "cursor.exe") {
        if (AIB_IsCursorComposerFocused(hwnd)) {
            AIB_AllowDebug_Write("enter-trigger detected proc=cursor hwnd=" hwnd)
            SendInput "{Enter}"
            return
        }
        SendInput "{Enter}"
        return
    }

    SendInput "{Enter}"
}
#HotIf

AIB_BuildStartBannerText(triggerSource := "manual", windowScope := "ide") {
    if (triggerSource = "persistent_startup")
        return "✅ Persistent Allow watcher enabled"
    if (triggerSource = "persistent_hotkey")
        return "✅ Persistent Allow watcher enabled"
    if (triggerSource = "chat_submit")
        return "⏳ Allow flow started from Enter send"
    if (triggerSource = "chat_submit_cursor")
        return "⏳ Allow flow started from Enter send (Cursor)"
    if (triggerSource = "start_implementation")
        return "⏳ Allow flow started from Start Implementation"
    if (triggerSource = "rapid_fire_hotkey")
        return "✅ Allow flow started from #!+9"

    if (windowScope = "cursor")
        return "⏳ Allow flow started (Cursor)"
    if (windowScope = "vscode")
        return "⏳ Allow flow started (VS Code)"
    return "⏳ Allow flow started"
}

AIB_ShouldShowStartBanner(triggerSource, targetHwnd := 0) {
    global g_AIB_StartBannerCooldownMs, g_AIB_LastStartBannerTickByKey

    key := triggerSource "|" String(targetHwnd ? targetHwnd : 0)
    now := A_TickCount
    lastTick := g_AIB_LastStartBannerTickByKey.Has(key) ? g_AIB_LastStartBannerTickByKey[key] : 0
    if (now - lastTick < g_AIB_StartBannerCooldownMs)
        return false

    g_AIB_LastStartBannerTickByKey[key] := now
    if (g_AIB_LastStartBannerTickByKey.Count > 200)
        g_AIB_LastStartBannerTickByKey := Map()
    return true
}

AIB_ArmAllowWatcher(triggerSource := "manual", targetHwnd := 0, windowScope := "ide", pinToTarget := false,
    showStartBanner := true, bannerText := "", bannerAccent := BANNER_ACCENT_INTERMEDIATE) {
    windowScope := AIB_NormalizeAllowWatcherScope(windowScope)
    AIB_StartAllowWatcher(triggerSource, targetHwnd, windowScope, pinToTarget)
    AIB_AllowDebug_Write("trigger-accepted source=" triggerSource " scope=" windowScope " target=" targetHwnd " pinned=" pinToTarget
    )

    ; Keep the persistent indicator visible even when the startup toast is suppressed.
    AIB_PersistentBannerCreate()

    if (!showStartBanner)
        return true
    if (!AIB_ShouldShowStartBanner(triggerSource, targetHwnd))
        return true

    if (bannerText = "")
        bannerText := AIB_BuildStartBannerText(triggerSource, windowScope)
    try ShowCenteredOverlay_Utils(bannerText, 1300, bannerAccent)
    catch {
    }
    AIB_AllowDebug_Write("trigger-banner source=" triggerSource " text=" bannerText)
    return true
}

AIB_StopAllowWatcher(reason := "", showInfo := false) {
    global g_AIB_AllowWatcherActive, g_AIB_AllowWatcherTimer, g_AIB_AllowWatcherStateByHwnd
    global g_AIB_AllowWatcherPromptLock, g_AIB_AllowWatcherDecision, g_AIB_AllowWatcherSource
    global g_AIB_AllowWatcherStartedTick, g_AIB_AllowWatcherRoundRobinOffset, g_AIB_AllowWatcherWindowScope
    global g_AIB_AllowWatcherPinnedHwnd, g_AIB_AllowWatcherPersistentMode

    if (g_AIB_AllowWatcherTimer)
        SetTimer(g_AIB_AllowWatcherTimer, 0)
    g_AIB_AllowWatcherTimer := ""
    g_AIB_AllowWatcherActive := false
    g_AIB_AllowWatcherPromptLock := false
    g_AIB_AllowWatcherDecision := ""
    g_AIB_AllowWatcherSource := ""
    g_AIB_AllowWatcherWindowScope := "ide"
    g_AIB_AllowWatcherPinnedHwnd := 0
    g_AIB_AllowWatcherPersistentMode := false
    g_AIB_AllowWatcherStartedTick := 0
    g_AIB_AllowWatcherRoundRobinOffset := 0
    g_AIB_AllowWatcherStateByHwnd := Map()

    ; Destroy persistent banner
    AIB_PersistentBannerDestroy()

    try StandardLoadingBar_CloseKeysOverlay()
    catch {
    }

    if (showInfo && reason != "")
        ShowCenteredOverlay_Utils(reason, 1800, BANNER_ACCENT_INFO)

    AIB_AllowDebug_Write("stop reason=" reason)
}

AIB_AllowWatcherEnsureState(hwnd, source := "manual", scope := "ide") {
    global g_AIB_AllowWatcherStateByHwnd, g_AIB_AllowWatcherTimeoutMs
    if (!hwnd)
        return
    key := String(hwnd)
    if (!g_AIB_AllowWatcherStateByHwnd.Has(key)) {
        g_AIB_AllowWatcherStateByHwnd[key] := {
            seenAllow: false,
            skipCurrentAllow: false,
            absenceStreak: 0,
            lastSeenTick: A_TickCount,
            lastClickTick: 0,
            ignoreUntilTick: 0,
            startedTick: A_TickCount,
            source: source,
            scope: scope,
            timeoutMs: g_AIB_AllowWatcherTimeoutMs
        }
    } else {
        st := g_AIB_AllowWatcherStateByHwnd[key]
        st.source := source
        st.scope := scope
        st.startedTick := A_TickCount
        st.lastClickTick := 0
        st.ignoreUntilTick := 0
        st.timeoutMs := g_AIB_AllowWatcherTimeoutMs
        g_AIB_AllowWatcherStateByHwnd[key] := st
    }
}

AIB_AllowWatcherTick(*) {
    global g_AIB_AllowWatcherActive, g_AIB_AllowWatcherTimeoutMs
    global g_AIB_AllowWatcherStateByHwnd, g_AIB_AllowWatcherPromptLock, g_AIB_AllowWatcherRoundRobinOffset
    global g_AIB_AllowWatcherSource

    if (!g_AIB_AllowWatcherActive)
        return

    if (g_AIB_AllowWatcherStateByHwnd.Count = 0 && !AIB_IsPersistentAllowWatcherMode()) {
        AIB_StopAllowWatcher("", false)
        return
    }

    ; Auto-enroll any new IDE windows that opened after arming
    try {
        global g_AIB_AllowWatcherWindowScope
        scope := g_AIB_AllowWatcherWindowScope ? g_AIB_AllowWatcherWindowScope : "ide"
        for hwnd in AIB_GetWatcherWindowHwnds(scope) {
            key := String(hwnd)
            if (!g_AIB_AllowWatcherStateByHwnd.Has(key))
                AIB_AllowWatcherEnsureState(hwnd, g_AIB_AllowWatcherSource, scope)
        }
    } catch {
    }

    keys := []
    for key, _ in g_AIB_AllowWatcherStateByHwnd
        keys.Push(key)
    if (keys.Length = 0)
        return

    ; Fairness: rotate start index every tick.
    g_AIB_AllowWatcherRoundRobinOffset := Mod(g_AIB_AllowWatcherRoundRobinOffset + 1, keys.Length)

    loop keys.Length {
        idx := Mod((A_Index - 1) + g_AIB_AllowWatcherRoundRobinOffset, keys.Length) + 1
        key := keys[idx]
        if (!g_AIB_AllowWatcherStateByHwnd.Has(key))
            continue
        hwnd := Integer(key)
        if (!WinExist("ahk_id " hwnd)) {
            g_AIB_AllowWatcherStateByHwnd.Delete(key)
            continue
        }

        st := g_AIB_AllowWatcherStateByHwnd[key]
        if (A_TickCount - st.startedTick >= st.timeoutMs) {
            AIB_AllowDebug_Write("tick session-timeout hwnd=" hwnd " source=" st.source)
            if (AIB_IsPersistentAllowWatcherMode()) {
                st.seenAllow := false
                st.absenceStreak := 0
                st.startedTick := A_TickCount
                st.lastSeenTick := A_TickCount
            } else {
                g_AIB_AllowWatcherStateByHwnd.Delete(key)
                continue
            }
        }
        st.lastSeenTick := A_TickCount

        ; Skip checking if we recently clicked (1 second cooldown)
        if (st.lastClickTick && A_TickCount - st.lastClickTick < 1000) {
            continue
        }

        ; Ignore retries for 30 seconds when user chooses Ignore.
        if (st.ignoreUntilTick && A_TickCount < st.ignoreUntilTick) {
            continue
        }
        if (st.ignoreUntilTick && A_TickCount >= st.ignoreUntilTick)
            st.ignoreUntilTick := 0

        failReason := ""
        hasAllow := AIB_WindowHasAllowButton(hwnd, &failReason)
        if (hasAllow) {
            AIB_AllowDebug_Write("tick allow-detected hwnd=" hwnd " title=" AIB_GetSafeWindowTitle(hwnd))
            st.seenAllow := true
            st.absenceStreak := 0

            if (st.skipCurrentAllow) {
                AIB_AllowDebug_Write("tick allow-skipped hwnd=" hwnd)
                g_AIB_AllowWatcherStateByHwnd[key] := st
                continue
            }

            ; One decision overlay at a time; when approved, click once and resume monitoring.
            if (!g_AIB_AllowWatcherPromptLock) {
                g_AIB_AllowWatcherPromptLock := true
                decision := ""
                try {
                    decision := AIB_RunAllowDecisionFlow(hwnd)
                } catch as e {
                    AIB_AllowDebug_Write("decision-flow-error hwnd=" hwnd " err=" e.Message)
                    ; Fail open so a transient overlay error does not deadlock the watcher.
                    decision := "Y"
                } finally {
                    g_AIB_AllowWatcherPromptLock := false
                }

                if (decision = "I") {
                    st.ignoreUntilTick := A_TickCount + 30000
                    st.seenAllow := false
                    st.absenceStreak := 0
                    st.startedTick := A_TickCount
                    AIB_AllowDebug_Write("tick allow-ignore-armed hwnd=" hwnd " ms=30000")
                    g_AIB_AllowWatcherStateByHwnd[key] := st
                    continue
                }

                if (decision = "N") {
                    if (AIB_IsPersistentAllowWatcherMode()) {
                        st.skipCurrentAllow := true
                        st.seenAllow := true
                        st.absenceStreak := 0
                        st.startedTick := A_TickCount
                        AIB_AllowDebug_Write("tick allow-skip-armed hwnd=" hwnd)
                        g_AIB_AllowWatcherStateByHwnd[key] := st
                    } else {
                        AIB_AllowDebug_Write("tick session-cancel hwnd=" hwnd)
                        g_AIB_AllowWatcherStateByHwnd.Delete(key)
                    }
                    continue
                }

                clickFail := ""
                prevForegroundHwnd := AIB_GetForegroundWindowForRestore()
                AIB_PrepareWindowForAllowClick(hwnd)
                AIB_ClickAllowButtonInWindow(hwnd, &clickFail)
                AIB_RestoreForegroundWindow(prevForegroundHwnd, hwnd)
                if (clickFail != "") {
                    AIB_AllowDebug_Write("tick click-fail hwnd=" hwnd " reason=" clickFail)
                } else {
                    ; Successful click: set cooldown so we don't re-click immediately
                    st.lastClickTick := A_TickCount
                    g_AIB_AllowWatcherStateByHwnd[key] := st
                    AIB_AllowDebug_Write("tick click-success hwnd=" hwnd)
                }
            } else {
                AIB_AllowDebug_Write("tick allow-detected prompt-locked hwnd=" hwnd)
            }
        } else {
            if (st.skipCurrentAllow) {
                st.skipCurrentAllow := false
                st.seenAllow := false
                st.absenceStreak := 0
                st.startedTick := A_TickCount
                AIB_AllowDebug_Write("tick allow-skip-cleared hwnd=" hwnd)
            }

            if (st.seenAllow) {
                st.absenceStreak += 1

                ; For non-rapid-fire sources, completion is when Send is ready again.
                if (st.source != "rapid_fire_hotkey" && AIB_WindowHasSendReady(hwnd)) {
                    AIB_AllowDebug_Write("tick send-ready-complete hwnd=" hwnd " source=" st.source)
                    if (AIB_IsPersistentAllowWatcherMode()) {
                        st.seenAllow := false
                        st.absenceStreak := 0
                        st.startedTick := A_TickCount
                    } else {
                        g_AIB_AllowWatcherStateByHwnd.Delete(key)
                    }
                    continue
                }

                ; Re-arm for next occurrence instead of permanently completing this window.
                if (st.absenceStreak >= 2) {
                    st.seenAllow := false
                    st.absenceStreak := 0
                }
            }
        }

        g_AIB_AllowWatcherStateByHwnd[key] := st
    }

    ; Update persistent banner position if monitor changed
    AIB_PersistentBannerMonitorUpdate()

    if (g_AIB_AllowWatcherStateByHwnd.Count = 0 && !AIB_IsPersistentAllowWatcherMode())
        AIB_StopAllowWatcher("✓ Allow flow completed (all sessions)", true)
}

AIB_AllowDebug_StartSession(triggerSource, scope, pinnedHwnd) {
    global g_AIB_AllowWatcherDebugTraceEnabled, g_AIB_AllowWatcherDebugLogFile
    if (!g_AIB_AllowWatcherDebugTraceEnabled)
        return

    ; Append a separator instead of clearing so probe entries before arm are preserved
    AIB_AllowDebug_Write("--- session-start source=" triggerSource " scope=" scope " pinned=" pinnedHwnd " ---")
}

AIB_AllowDebug_Write(line) {
    global g_AIB_AllowWatcherDebugTraceEnabled, g_AIB_AllowWatcherDebugLogFile
    if (!g_AIB_AllowWatcherDebugTraceEnabled)
        return

    stamp := FormatTime(, "yyyy-MM-dd HH:mm:ss")
    try FileAppend(stamp " | " line "`n", g_AIB_AllowWatcherDebugLogFile, "UTF-8")
}

AIB_AllowDebug_RectText(rect) {
    if (!rect)
        return "none"
    return "(" rect.l "," rect.t ")-(" rect.r "," rect.b ")"
}

AIB_PrepareWindowForAllowClick(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return
    try {
        ; Ensure the target IDE is foreground before clicking Allow.
        WinActivate("ahk_id " hwnd)
        WinWaitActive("ahk_id " hwnd, , 0.35)
    } catch {
    }
}

AIB_GetForegroundWindowForRestore() {
    hwnd := 0
    try hwnd := WinGetID("A")
    catch {
        hwnd := 0
    }
    return hwnd
}

AIB_RestoreForegroundWindow(hwnd, skippedHwnd := 0) {
    if (!hwnd || hwnd = skippedHwnd)
        return false
    if (!WinExist("ahk_id " hwnd))
        return false
    try {
        WinActivate("ahk_id " hwnd)
        return true
    } catch {
        return false
    }
}

AIB_WindowHasSendReady(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false

    exeName := ""
    try exeName := StrLower(WinGetProcessName("ahk_id " hwnd))
    catch {
    }

    ; Only apply send-ready completion for IDE windows with chat send controls.
    if (exeName != "code.exe" && exeName != "cursor.exe")
        return false

    try return VSCode_IsChatSendReady(hwnd)
    catch {
        return false
    }
}

AIB_RunAllowDecisionFlow(hwnd) {
    global g_AIB_AllowWatcherDecision

    title := AIB_GetSafeWindowTitle(hwnd)
    g_AIB_AllowWatcherDecision := ""
    anchorHwnd := WinActive("A")
    if (!anchorHwnd)
        anchorHwnd := hwnd
    AIB_AllowDebug_Write("decision-flow hwnd=" hwnd " anchor=" anchorHwnd)

    cmdPreview := AIB_GetAllowCommandPreview(hwnd)

    ; Play sound at START to alert user
    SoundPlay(A_ScriptDir . "\\sounds\\clicking-allow.wav")

    ; Banner 1: exactly 2 seconds with progressive loading + command display.
    ; Display command prominently with newlines for readability
    banner1Text := "⏳ Allow detected in " . title . "`n`nCommand to execute:`n" . cmdPreview
    StandardLoadingBar_Show(
        banner1Text,
        BANNER_ACCENT_INTERMEDIATE, { textWidth: 700, noBorder: true, trackActiveMonitor: true, manualProgress: true,
            centerOnHwnd: anchorHwnd, fontSize: 15 }
    )
    StandardLoadingBar_StartTimedProgress(2000)
    Sleep(2000)
    StandardLoadingBar_Hide(0)

    keyCallbacks := Map(
        "Y", AIB_AllowWatcherDecisionYes,
        "N", AIB_AllowWatcherDecisionNo,
        "I", AIB_AllowWatcherDecisionIgnore,
        "Escape", AIB_AllowWatcherDecisionNo
    )

    ; Banner 2: exactly 7 seconds with progressive loading + decision input + command.
    ; Display command with emphasis for user approval, separated by big space
    banner2Text := "❓ Click Allow now? (7s)"
    commandDisplay := "📋 " . cmdPreview
    StandardLoadingBar_ShowWithKeys(
        banner2Text,
        keyCallbacks,
        7000,
        anchorHwnd,
        AIB_AllowWatcherDecisionTimeout,
        BANNER_ACCENT_INTERMEDIATE,
        700,
        17,
        BANNER_ACCENT_INFO,
        true,
        "`n`n`n`n" . commandDisplay . "`n`n`n`n[Y] Click Allow  [N] Stop watcher  [I] Ignore 30s",
        true,
        true
    )

    waitDeadline := A_TickCount + 7400
    while (g_AIB_AllowWatcherDecision = "" && A_TickCount < waitDeadline)
        Sleep(30)

    if (g_AIB_AllowWatcherDecision = "")
        g_AIB_AllowWatcherDecision := "Y"

    return g_AIB_AllowWatcherDecision
}

AIB_AllowWatcherDecisionYes(*) {
    global g_AIB_AllowWatcherDecision
    g_AIB_AllowWatcherDecision := "Y"
}

AIB_AllowWatcherDecisionNo(*) {
    global g_AIB_AllowWatcherDecision
    g_AIB_AllowWatcherDecision := "N"
}

AIB_AllowWatcherDecisionIgnore(*) {
    global g_AIB_AllowWatcherDecision
    g_AIB_AllowWatcherDecision := "I"
}

AIB_AllowWatcherDecisionTimeout(*) {
    ; Requirement: timeout on Banner 2 is treated as Y.
    global g_AIB_AllowWatcherDecision
    g_AIB_AllowWatcherDecision := "Y"
}

AIB_GetAllowCommandPreview(hwnd) {
    global UIA
    placeholder := "<command not found>"

    if (!hwnd || !WinExist("ahk_id " hwnd))
        return placeholder

    try root := UIA.ElementFromHandle(hwnd)
    catch {
        return placeholder
    }
    if (!root)
        return placeholder

    cmd := AIB_ExtractAllowCommandFromRoot(root)
    cmd := AIB_NormalizeAllowCommandPreview(cmd)
    if (cmd = "")
        return placeholder
    return cmd
}

AIB_ExtractAllowCommandFromRoot(root) {
    if (!root)
        return ""

    ; 1) Preferred source: confirmation group/container context.
    try {
        groups := root.FindAll({ Type: 50026 })
        for grp in groups {
            cls := ""
            nm := ""
            try cls := StrLower(grp.ClassName)
            try nm := grp.Name
            if (!InStr(cls, "chat-confirmation-widget-container") && !InStr(StrLower(nm), "chat confirmation dialog"))
                continue

            parsed := AIB_ParseAllowCommandText(nm)
            if (parsed != "")
                return parsed

            try {
                texts := grp.FindAll({ Type: 50020 })
                for txt in texts {
                    t := ""
                    try t := txt.Name
                    parsed := AIB_ParseAllowCommandText(t)
                    if (parsed != "")
                        return parsed
                }
            } catch {
            }
        }
    } catch {
    }

    ; 2) Fallback source: chat list item body (layout can vary).
    try {
        items := root.FindAll({ Type: 50007 })
        for item in items {
            nm := ""
            try nm := item.Name
            if (!AIB_LooksLikeAllowCommandCarrier(nm))
                continue
            parsed := AIB_ParseAllowCommandText(nm)
            if (parsed != "")
                return parsed
        }
    } catch {
    }

    return ""
}

AIB_LooksLikeAllowCommandCarrier(text) {
    t := StrLower(Trim(text))
    if (t = "")
        return false

    if (InStr(t, "chat confirmation required"))
        return true
    if (InStr(t, "press control+enter to accept") || InStr(t, "press ctrl+enter to accept"))
        return true
    if (InStr(t, " run ") && InStr(t, " command within ") && InStr(t, "?"))
        return true
    return false
}

AIB_ParseAllowCommandText(rawText) {
    txt := AIB_NormalizeAllowCommandPreview(rawText, 1200)
    if (txt = "")
        return ""

    ; Support prompt variants such as:
    ; "... Run `pwsh` command within `<path>`?: <cmd>. Press Control+Enter ..."
    ; "... command?: <cmd>. Press Control+Enter ..."
    if (RegExMatch(txt, "i)run\s+`?\w+`?\s+command(?:\s+within)?[^?]*\?\s*:?\s*(.+)$", &m))
        txt := m[1]
    else if (RegExMatch(txt, "i)command(?:\s+within)?[^?]*\?\s*:?\s*(.+)$", &m))
        txt := m[1]
    else if (InStr(txt, "?:"))
        txt := Trim(SubStr(txt, InStr(txt, "?:") + 2))
    else
        return ""

    lower := StrLower(txt)
    cutPos := 0
    for marker in [" press control+enter", " press ctrl+enter", " alt+backspace", " to accept", " to cancel",
        " code blocks:"] {
        p := InStr(lower, marker)
        if (p && (!cutPos || p < cutPos))
            cutPos := p
    }

    if (cutPos)
        txt := SubStr(txt, 1, cutPos - 1)

    return Trim(txt)
}

AIB_NormalizeAllowCommandPreview(text, maxLen := 150) {
    if (text = "")
        return ""

    out := StrReplace(text, Chr(160), " ")
    out := StrReplace(out, "`r", " ")
    out := StrReplace(out, "`n", " ")
    out := RegExReplace(out, "\s+", " ")
    out := Trim(out)

    if (out = "")
        return ""

    if (maxLen > 3 && StrLen(out) > maxLen)
        out := SubStr(out, 1, maxLen - 3) . "..."

    return out
}

AIB_WindowHasAllowButton(hwnd, &failReason := "") {
    global UIA
    failReason := ""
    if (!hwnd || !WinExist("ahk_id " hwnd)) {
        failReason := "window missing"
        return false
    }

    try root := UIA.ElementFromHandle(hwnd)
    catch {
        failReason := "UIA failed"
        return false
    }
    if (!root) {
        failReason := "root not found"
        return false
    }

    try {
        btn := AIB_FindAllowButtonInChatConfirmation(root)
        if (btn)
            return true
        if (AIB_HasChatConfirmationAcceptHint(root))
            return true
        frame := AIB_GetPreferredAllowSearchFrame(root, hwnd)
        if (frame) {
            if (AIB_GetBestAllowButton(root, frame))
                return true
            if (AIB_FindAllowButtonInDocumentSubtree(root, frame))
                return true
            failReason := "Allow not found in confirmation frame"
            return false
        }

        if (AIB_FindAllowButtonInDocumentSubtree(root))
            return true
        failReason := "Allow not found in AI chat confirmation context"
        return false
    } catch {
        failReason := "button search failed"
        return false
    }

    failReason := "Allow button not found"
    return false
}

AIB_WindowHasStartImplementationButton(hwnd) {
    global UIA
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false

    try root := UIA.ElementFromHandle(hwnd)
    catch
        return false
    if (!root)
        return false

    ; VS Code/Electron web content is inside a RootWebArea Document element
    ; FindAll from window handle doesn't traverse into it, so navigate manually
    try {
        docs := root.FindAll({ Type: 50030 })  ; Type 50030 = Document
        for doc in docs {
            if (!doc)
                continue
            ; Search for buttons inside the document
            try {
                btns := doc.FindAll({ Type: 50000 })
                AIB_AllowDebug_Write("start-impl-btn doc-scan hwnd=" hwnd " doc-btns=" btns.Length)
                for btn in btns {
                    nm := ""
                    cls := ""
                    try nm := btn.Name
                    try cls := btn.ClassName
                    if (InStr(cls, "chat-suggest") || InStr(cls, "chat-welcome"))
                        AIB_AllowDebug_Write("start-impl-btn candidate hwnd=" hwnd " name='" nm "' class='" cls "'")
                    if (InStr(StrLower(Trim(nm)), "start implementation")) {
                        AIB_AllowDebug_Write("start-impl-btn found hwnd=" hwnd " name='" nm "'")
                        return true
                    }
                }
            } catch {
            }
        }
    } catch {
    }

    ; Fallback: also try direct window scan (non-web windows)
    try {
        allBtns := root.FindAll({ Type: 50000 })
        AIB_AllowDebug_Write("start-impl-btn window-scan hwnd=" hwnd " total-btns=" allBtns.Length)
        for btn in allBtns {
            nm := ""
            try nm := btn.Name
            if (InStr(StrLower(Trim(nm)), "start implementation")) {
                AIB_AllowDebug_Write("start-impl-btn found-window hwnd=" hwnd " name='" nm "'")
                return true
            }
        }
    } catch {
    }
    return false
}

; =============================================================================
; Persistent Flow Banner Functions
; =============================================================================

AIB_PersistentBannerCreate() {
    global g_AIB_FlowBannerHwnd, g_AIB_FlowBannerLastMonitorIdx

    ; Destroy any existing banner
    AIB_PersistentBannerDestroy()

    ; Get active monitor work area
    GetActiveMonitorWorkArea_StandardBar(&ml, &mt, &mr, &mb)
    g_AIB_FlowBannerLastMonitorIdx := GetMonitorIndexForForeground_StandardBar()

    ; Create tight emoji-only indicator — no caption, no margins, no fixed dimensions
    bannerGui := Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale")
    bannerGui.BackColor := "1F1F1F"
    bannerGui.MarginX := 4
    bannerGui.MarginY := 4
    bannerGui.SetFont("s14 c90EE90", "Segoe UI")
    bannerGui.Add("Text", "Center", "⏳")

    ; Auto-size first, then anchor to bottom-left
    bannerGui.Show("AutoSize Hide")
    bannerGui.GetPos(, , &bw, &bh)

    marginX := 20
    marginY := 20
    bannerX := ml + marginX
    bannerY := mb - bh - marginY

    bannerGui.Show("x" . bannerX . " y" . bannerY . " NA")
    WinSetTransparent(210, bannerGui)

    g_AIB_FlowBannerHwnd := bannerGui.Hwnd
    AIB_AllowDebug_Write("banner-created hwnd=" g_AIB_FlowBannerHwnd " x=" bannerX " y=" bannerY " w=" bw " h=" bh)
    return g_AIB_FlowBannerHwnd
}

AIB_PersistentBannerMonitorUpdate() {
    global g_AIB_FlowBannerHwnd, g_AIB_FlowBannerLastMonitorIdx

    if (!g_AIB_FlowBannerHwnd || !WinExist("ahk_id " g_AIB_FlowBannerHwnd))
        return

    ; Check if monitor changed
    currentMonitorIdx := GetMonitorIndexForForeground_StandardBar()
    if (currentMonitorIdx = g_AIB_FlowBannerLastMonitorIdx)
        return  ; No change, skip repositioning

    ; Monitor changed, reposition banner
    GetActiveMonitorWorkArea_StandardBar(&ml, &mt, &mr, &mb)
    g_AIB_FlowBannerLastMonitorIdx := currentMonitorIdx

    marginX := 20
    marginY := 20
    bh := 0
    try WinGetPos(, , , &bh, "ahk_id " g_AIB_FlowBannerHwnd)
    if (bh <= 0)
        bh := 30
    bannerX := ml + marginX
    bannerY := mb - bh - marginY

    WinMove(bannerX, bannerY, , , "ahk_id " g_AIB_FlowBannerHwnd)
    AIB_AllowDebug_Write("banner-repositioned hwnd=" g_AIB_FlowBannerHwnd " monitor=" currentMonitorIdx " x=" bannerX " y=" bannerY
    )
}

AIB_PersistentBannerDestroy() {
    global g_AIB_FlowBannerHwnd

    if (!g_AIB_FlowBannerHwnd)
        return

    try {
        if (WinExist("ahk_id " g_AIB_FlowBannerHwnd)) {
            WinClose("ahk_id " g_AIB_FlowBannerHwnd)
        }
    } catch {
    }

    AIB_AllowDebug_Write("banner-destroyed hwnd=" g_AIB_FlowBannerHwnd)
    g_AIB_FlowBannerHwnd := 0
}

AIB_StartImplementationProbeTick(*) {
    global g_AIB_AllowWatcherActive, g_AIB_StartImplPrevPresentByHwnd

    if (g_AIB_AllowWatcherActive) {
        AIB_AllowDebug_Write("start-impl-probe blocked watcher-active=1")
        return
    }

    windows := AIB_GetWatcherWindowHwnds("ide")
    AIB_AllowDebug_Write("start-impl-probe ide-windows=" windows.Length)
    alive := Map()

    for hwnd in windows {
        key := String(hwnd)

        alive[key] := true
        presentNow := AIB_WindowHasStartImplementationButton(hwnd)
        wasPresent := g_AIB_StartImplPrevPresentByHwnd.Has(key) ? g_AIB_StartImplPrevPresentByHwnd[key] : false

        if (wasPresent != presentNow)
            AIB_AllowDebug_Write("start-impl-probe hwnd=" hwnd " was=" wasPresent " now=" presentNow " active=" (
                WinActive("ahk_id " hwnd) ? 1 : 0))

        ; Transition in foreground window usually indicates the click just happened.
        if (wasPresent && !presentNow && WinActive("ahk_id " hwnd)) {
            proc := ""
            try proc := StrLower(WinGetProcessName("ahk_id " hwnd))
            scope := (proc = "code.exe") ? "vscode" : "cursor"
            AIB_AllowDebug_Write("start-impl-probe arm scope=" scope " hwnd=" hwnd)
            AIB_ArmAllowWatcher("start_implementation", 0, "vscode", false, true)
        }

        g_AIB_StartImplPrevPresentByHwnd[key] := presentNow
    }

    for key, _ in g_AIB_StartImplPrevPresentByHwnd {
        if (!alive.Has(key))
            g_AIB_StartImplPrevPresentByHwnd.Delete(key)
    }
}

AIB_GetWatcherWindowHwnds(scope := "ide", pinnedHwnd := 0) {
    scope := AIB_NormalizeAllowWatcherScope(scope)

    if (pinnedHwnd && WinExist("ahk_id " pinnedHwnd)) {
        proc := ""
        try proc := StrLower(WinGetProcessName("ahk_id " pinnedHwnd))
        if (scope = "vscode" && proc != "code.exe")
            pinnedHwnd := 0
        else if (scope = "cursor" && proc != "cursor.exe")
            pinnedHwnd := 0
        else if (scope = "ide" && proc != "code.exe" && proc != "cursor.exe")
            pinnedHwnd := 0
    }
    if (pinnedHwnd && WinExist("ahk_id " pinnedHwnd)) {
        out := []
        out.Push(pinnedHwnd)
        return out
    }

    windows := []
    for hwnd in WinGetList() {
        try {
            procName := WinGetProcessName("ahk_id " hwnd)
            procLower := StrLower(procName)
            isCursor := (procLower = "cursor.exe")
            isVsCode := (procLower = "code.exe")

            if (scope = "vscode") {
                if (!isVsCode)
                    continue
            } else if (scope = "cursor") {
                if (!isCursor)
                    continue
            } else {
                if (!isCursor && !isVsCode)
                    continue
            }

            winTitle := WinGetTitle("ahk_id " hwnd)
            if (!winTitle)
                continue

            windows.Push(hwnd)
        } catch {
            continue
        }
    }
    return windows
}

AIB_ClickAllowButtonInAllIDEWindows() {
    global UIA
    sourceHwnd := 0
    try sourceHwnd := WinGetID("A")

    windows := AIB_GetAllIDEWindowHwnds()
    if (windows.Length = 0) {
        ShowCenteredOverlay_Utils("ℹ No open Cursor/VS Code windows", 1800, BANNER_ACCENT_INFO)
        return 0
    }

    clickedCount := 0
    totalCount := windows.Length
    centerHwnd := sourceHwnd
    debugReasons := []

    try {
        StandardLoadingBar_Show(
            "⏳ Scanning IDE windows for Allow button...",
            BANNER_ACCENT_INTERMEDIATE, { centerOnHwnd: centerHwnd, textWidth: 640 }
        )

        for hwnd in windows {
            StandardLoadingBar_Update(
                "⏳ Checking " . A_Index . "/" . totalCount . ": " . AIB_GetSafeWindowTitle(hwnd),
                BANNER_ACCENT_INTERMEDIATE
            )
            failReason := ""
            SoundPlay(A_ScriptDir . "\sounds\clicking-allow.wav")
            prevForegroundHwnd := sourceHwnd ? sourceHwnd : AIB_GetForegroundWindowForRestore()
            if (AIB_ClickAllowButtonInWindow(hwnd, &failReason)) {
                AIB_RestoreForegroundWindow(prevForegroundHwnd, hwnd)
                clickedCount += 1
                break
            }
            AIB_RestoreForegroundWindow(prevForegroundHwnd, hwnd)
            if (failReason != "" && debugReasons.Length < 3)
                debugReasons.Push(AIB_GetSafeWindowTitle(hwnd) . " => " . failReason)
        }
    } finally {
        StandardLoadingBar_Hide(0)
    }

    if (clickedCount > 0)
        ShowCenteredOverlay_Utils("✅ Clicked Allow and stopped search", 1800, BANNER_ACCENT_SUCCESS)
    else {
        msg := "ℹ No Allow button found in chat confirmation"
        if (debugReasons.Length > 0) {
            for reasonLine in debugReasons
                msg .= "`n" . reasonLine
        }
        ShowCenteredOverlay_Utils(msg, 3200, BANNER_ACCENT_INFO)
    }

    return clickedCount
}

AIB_ClickAllowButtonInWindow(hwnd, &failReason := "") {
    global UIA
    failReason := ""

    ; Find Allow button
    try root := UIA.ElementFromHandle(hwnd)
    catch {
        failReason := "UIA failed"
        return false
    }
    if (!root) {
        failReason := "root not found"
        return false
    }

    btn := 0
    frame := 0
    frameSource := "none"
    try {
        ; Prefer exact confirmation widget target first.
        btn := AIB_FindAllowButtonInChatConfirmation(root)
        if (!btn) {
            frame := AIB_GetPreferredAllowSearchFrame(root, hwnd, &frameSource)
            if (frame)
                btn := AIB_GetBestAllowButton(root, frame)
        }
    } catch {
        failReason := "button search failed"
        AIB_AllowDebug_Write("click-search-failed hwnd=" hwnd)
        return false
    }

    if (!btn) {
        AIB_AllowDebug_Write("click-no-btn hwnd=" hwnd " frameSource=" frameSource " frame=" AIB_AllowDebug_RectText(
            frame))
        ; Fallback: confirmation is present but button object may be inaccessible in this cycle.
        if (AIB_HasChatConfirmationAcceptHint(root)) {
            route := ""
            if (AIB_SendAcceptShortcutToWindow(hwnd, &route)) {
                verifyReason := ""
                if (!AIB_WindowHasAllowButton(hwnd, &verifyReason))
                    return true
            }

            ocrRoute := ""
            if (AIB_TryOcrAllowClick(hwnd, root, &ocrRoute)) {
                verifyReason := ""
                if (!AIB_WindowHasAllowButton(hwnd, &verifyReason))
                    return true
            }

            AIB_AllowDebug_Write("click-hint-shortcut-failed hwnd=" hwnd " route=" route)
            failReason := "accept hint found but shortcut did not clear prompt"
            return false
        }

        failReason := "Allow button not found"
        return false
    }

    actionBranch := ""
    if (AIB_ClickAllowButtonElement(btn, hwnd, &actionBranch)) {
        btnName := ""
        btnClass := ""
        btnRect := 0
        try btnName := btn.Name
        try btnClass := btn.ClassName
        try btnRect := btn.BoundingRectangle
        AIB_AllowDebug_Write(
            "click-attempt hwnd=" hwnd
            " frameSource=" frameSource
            " frame=" AIB_AllowDebug_RectText(frame)
            " branch=" actionBranch
            " btn='" btnName "'"
            " class='" btnClass "'"
            " rect=" AIB_AllowDebug_RectText(btnRect)
        )

        ; Verify the prompt progressed; if not, use the explicit accelerator fallback.
        Sleep(180)
        verifyReason := ""
        if (!AIB_WindowHasAllowButton(hwnd, &verifyReason))
            return true

        ; Safety: only use shortcut fallback when confirmation context is explicit.
        if (!AIB_HasChatConfirmationAcceptHint(root)) {
            AIB_AllowDebug_Write("click-post-shortcut-skipped hwnd=" hwnd " reason=no-confirmation-hint")
            failReason := "allow still present; shortcut skipped (no confirmation hint)"
            return false
        }

        route := ""
        AIB_SendAcceptShortcutToWindow(hwnd, &route)
        AIB_AllowDebug_Write("click-post-shortcut hwnd=" hwnd " route=" route)

        verifyReason := ""
        if (!AIB_WindowHasAllowButton(hwnd, &verifyReason))
            return true

        AIB_AllowDebug_Write("click-still-present hwnd=" hwnd " branch=" actionBranch)
        failReason := "allow still present after click"
        return false
    }

    AIB_AllowDebug_Write("click-action-failed hwnd=" hwnd)

    ocrRoute := ""
    if (AIB_TryOcrAllowClick(hwnd, root, &ocrRoute)) {
        verifyReason := ""
        if (!AIB_WindowHasAllowButton(hwnd, &verifyReason))
            return true
        AIB_AllowDebug_Write("click-ocr-still-present hwnd=" hwnd " route=" ocrRoute)
    }

    failReason := "invoke/click failed"
    return false
}

AIB_SendAcceptShortcutToWindow(hwnd, &usedRoute := "") {
    usedRoute := "none"
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false

    try {
        ControlSend("^{Enter}", , "ahk_id " hwnd)
        Sleep(220)
        usedRoute := "window-controlsend"
        return true
    } catch {
    }

    ; Chromium-based IDE windows sometimes need the keystroke sent to the render host HWND.
    try {
        ctrlHwnds := WinGetControlsHwnd("ahk_id " hwnd)
        for ctrlHwnd in ctrlHwnds {
            cls := ""
            try cls := StrLower(WinGetClass("ahk_id " ctrlHwnd))
            if (!InStr(cls, "chrome_renderwidgethosthwnd"))
                continue
            try {
                ControlSend("^{Enter}", , "ahk_id " ctrlHwnd)
                Sleep(220)
                usedRoute := "renderhost-controlsend"
                return true
            } catch {
                continue
            }
        }
    } catch {
    }

    return false
}

AIB_FindAllowButtonInChatConfirmation(root) {
    if (!root)
        return 0

    try {
        groups := root.FindAll({ Type: 50026 })
        AIB_AllowDebug_Write("find-chat-confirm root-scan groups=" groups.Length)
        for grp in groups {
            cls := ""
            try cls := StrLower(grp.ClassName)
            AIB_AllowDebug_Write("find-chat-confirm group-check class='" cls "'")
            if (!InStr(cls, "chat-confirmation-widget-container"))
                continue

            AIB_AllowDebug_Write("find-chat-confirm found-container")
            ; In this widget, prefer explicit action button style used by "Allow Once (Ctrl+Enter)".
            buttons := grp.FindAll({ Type: 50000 })
            AIB_AllowDebug_Write("find-chat-confirm buttons=" buttons.Length)
            best := 0
            bestScore := -2147483647
            buttonSnapshot := ""
            for btn in buttons {
                nm := ""
                bcls := ""
                offscreen := true
                try nm := btn.Name
                try bcls := StrLower(btn.ClassName)
                try offscreen := !!btn.GetPropertyValue(UIA.Property.IsOffscreen)

                if (buttonSnapshot != "")
                    buttonSnapshot .= " | "
                buttonSnapshot .= "name='" nm "' class='" bcls "' offscreen=" offscreen

                if (!AIB_IsAllowButtonName(nm)) {
                    AIB_AllowDebug_Write("find-chat-confirm btn-reject name='" nm "' (not allow)")
                    continue
                }

                if (!AIB_IsInAiChatAllowContext(btn)) {
                    AIB_AllowDebug_Write("find-chat-confirm btn-reject name='" nm "' (not in context)")
                    continue
                }

                AIB_AllowDebug_Write("find-chat-confirm btn-candidate name='" nm "'")
                score := 0
                nml := StrLower(Trim(nm))
                if (InStr(nml, "allow once"))
                    score += 150
                if (InStr(nml, "ctrl+enter"))
                    score += 80
                if (InStr(bcls, "monaco-text-button"))
                    score += 160
                if (InStr(bcls, "monaco-button"))
                    score += 70
                if (!offscreen)
                    score += 35

                if (score > bestScore) {
                    best := btn
                    bestScore := score
                }
            }

            if (best) {
                AIB_AllowDebug_Write("find-chat-confirm result=found score=" bestScore)
                return best
            }

            AIB_AllowDebug_Write("confirm-container-no-candidate buttons=" buttonSnapshot)
        }
    } catch error {
        AIB_AllowDebug_Write("find-chat-confirm error=" error.What)
        return 0
    }

    return 0
}

AIB_HasChatConfirmationAcceptHint(root) {
    if (!root)
        return false

    try {
        ; Strong signal from the chat list item text in VS Code:
        ; "... Press Control+Enter to accept or Alt+Backspace to cancel"
        items := root.FindAll({ Type: 50007 })
        for item in items {
            nm := ""
            try nm := StrLower(item.Name)
            if (nm = "")
                continue
            if (InStr(nm, "ctrl+enter to accept") || InStr(nm, "control+enter to accept"))
                return true
            if (InStr(nm, "chat confirmation required"))
                return true
        }
    } catch {
    }

    ; Secondary signal: confirmation container itself exists.
    try {
        groups := root.FindAll({ Type: 50026 })
        for grp in groups {
            cls := ""
            try cls := StrLower(grp.ClassName)
            if (InStr(cls, "chat-confirmation-widget-container"))
                return true
        }
    } catch {
    }

    return false
}

AIB_ClickAllowButtonElement(btn, ownerHwnd := 0, &usedBranch := "") {
    global UIA
    usedBranch := "none"
    if (!btn)
        return false

    ; Ensure the target IDE window and UIA element have focus before trying click routes.
    if (ownerHwnd && WinExist("ahk_id " ownerHwnd)) {
        try {
            if (!WinActive("ahk_id " ownerHwnd)) {
                WinActivate("ahk_id " ownerHwnd)
                WinWaitActive("ahk_id " ownerHwnd, , 0.35)
            }
        } catch {
        }
    }
    try {
        btn.SetFocus()
        Sleep(70)
    } catch {
    }

    try {
        if (btn.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)) {
            btn.InvokePattern.Invoke()
            Sleep(160)
            usedBranch := "uia-invoke"
            return true
        }
    } catch {
    }

    try {
        btn.Click()
        Sleep(160)
        usedBranch := "uia-click"
        return true
    } catch {
    }

    ; Final fallback: click the element bounds without activating the window.
    if (ownerHwnd && WinExist("ahk_id " ownerHwnd)) {
        try {
            br := btn.BoundingRectangle
            x := Floor((br.l + br.r) / 2)
            y := Floor((br.t + br.b) / 2)
            if (x > 0 && y > 0) {
                ControlClick("x" x " y" y, "ahk_id " ownerHwnd, , "Left", 1, "NA Pos")
                Sleep(180)
                usedBranch := "controlclick-pos"
                return true
            }
        } catch {
        }
    }

    return false
}

AIB_GetBestAllowButton(root, searchFrame := 0) {
    if (!root)
        return 0

    best := 0
    bestScore := -2147483647
    bestTop := -2147483647

    try {
        allBtns := root.FindAll({ Type: 50000 })
        for btn in allBtns {
            nm := ""
            try nm := btn.Name
            if (!AIB_IsAllowButtonName(nm))
                continue

            if (!AIB_IsInAiChatAllowContext(btn))
                continue

            if (searchFrame && !AIB_IsElementCenterInsideFrame(btn, searchFrame))
                continue

            score := 0
            top := -2147483647
            cls := ""
            offscreen := true

            try cls := StrLower(btn.ClassName)
            try offscreen := !!btn.GetPropertyValue(UIA.Property.IsOffscreen)
            try {
                br := btn.BoundingRectangle
                top := br.t
            }

            nl := StrLower(Trim(nm))
            if (InStr(nl, "allow once"))
                score += 80
            if (InStr(nl, "ctrl+enter"))
                score += 40
            if (!offscreen)
                score += 25
            if (InStr(cls, "monaco-text-button"))
                score += 90
            if (InStr(cls, "monaco-button"))
                score += 40
            if (AIB_IsInChatConfirmationDialog(btn))
                score += 220

            if (score > bestScore || (score = bestScore && top > bestTop)) {
                best := btn
                bestScore := score
                bestTop := top
            }
        }
    } catch {
        return 0
    }

    return best
}

AIB_GetPreferredAllowSearchFrame(root, hwnd := 0, &frameSource := "") {
    frameSource := "none"
    frame := AIB_GetChatConfirmationContainerFrame(root)
    if (frame) {
        frameSource := "chat-confirmation-container"
        return frame
    }

    return 0
}

AIB_GetChatConfirmationContainerFrame(root) {
    if (!root)
        return 0

    try {
        groups := root.FindAll({ Type: 50026 })
        for grp in groups {
            cls := ""
            try cls := StrLower(grp.ClassName)
            if (!InStr(cls, "chat-confirmation-widget-container"))
                continue

            try {
                br := grp.BoundingRectangle
                if (br.r <= br.l || br.b <= br.t)
                    continue

                ; Expand a bit so edge-aligned button centers are still included.
                return {
                    l: br.l - 6,
                    t: br.t - 6,
                    r: br.r + 6,
                    b: br.b + 6
                }
            } catch {
                continue
            }
        }
    } catch {
    }

    return 0
}

AIB_IsElementCenterInsideFrame(el, frame) {
    if (!el || !frame)
        return false

    try {
        br := el.BoundingRectangle
        x := Floor((br.l + br.r) / 2)
        y := Floor((br.t + br.b) / 2)
        return (x >= frame.l && x <= frame.r && y >= frame.t && y <= frame.b)
    } catch {
        return false
    }
}

; =============================================================================
; AIB Persistent Allow Watcher - Hotkey: Win+Alt+Shift+9
; Toggle persistent watcher loop for VS Code Allow monitoring.
; =============================================================================
global g_AIB_Hotkey9ToggleCooldownTick := 0

#!+9::
{
    global g_AIB_AllowWatcherActive, g_AIB_Hotkey9ToggleCooldownTick

    now := A_TickCount
    if (now - g_AIB_Hotkey9ToggleCooldownTick < 300)
        return
    g_AIB_Hotkey9ToggleCooldownTick := now

    if (g_AIB_AllowWatcherActive) {
        AIB_StopAllowWatcher("ℹ Persistent Allow watcher stopped", true)
        return
    }

    targetHwnd := 0
    try targetHwnd := WinGetID("A")
    catch {
        targetHwnd := 0
    }

    if (targetHwnd && WinExist("ahk_id " targetHwnd)) {
        proc := ""
        try proc := StrLower(WinGetProcessName("ahk_id " targetHwnd))
        if (proc != "code.exe")
            targetHwnd := 0
    }

    ; Optional environment update for rapid-fire testing.
    AIB_RapidFireTrySwapTestFolderContent()
    AIB_ArmAllowWatcher("persistent_hotkey", targetHwnd, "vscode", true, true, "✅ Persistent Allow watcher enabled",
        BANNER_ACCENT_SUCCESS)
}

AIB_RapidFireTrySwapTestFolderContent() {
    global g_AIB_RapidFireTestSourceDir, g_AIB_RapidFireTestTargetDir
    if (Trim(g_AIB_RapidFireTestSourceDir) = "" || Trim(g_AIB_RapidFireTestTargetDir) = "")
        return false
    if (!DirExist(g_AIB_RapidFireTestSourceDir) || !DirExist(g_AIB_RapidFireTestTargetDir))
        return false
    return AIB_SwapFolderContentNoErase(g_AIB_RapidFireTestSourceDir, g_AIB_RapidFireTestTargetDir)
}

AIB_SwapFolderContentNoErase(sourceDir, targetDir) {
    try {
        archiveRoot := targetDir . "\\_aib_swap_archive"
        if (!DirExist(archiveRoot))
            DirCreate(archiveRoot)
        stamp := FormatTime(, "yyyyMMdd_HHmmss")
        archiveDir := archiveRoot . "\\" . stamp
        DirCreate(archiveDir)

        ; Move current target content into archive instead of deleting.
        loop files, targetDir . "\\*", "FD" {
            name := A_LoopFileName
            if (name = "_aib_swap_archive")
                continue
            fromPath := A_LoopFilePath
            toPath := archiveDir . "\\" . name
            try {
                if (InStr(FileExist(fromPath), "D"))
                    DirMove(fromPath, toPath)
                else
                    FileMove(fromPath, toPath)
            } catch {
            }
        }

        ; Copy source content into target.
        loop files, sourceDir . "\\*", "FD" {
            name := A_LoopFileName
            fromPath := A_LoopFilePath
            toPath := targetDir . "\\" . name
            if (InStr(FileExist(fromPath), "D"))
                DirCopy(fromPath, toPath, true)
            else
                FileCopy(fromPath, toPath, true)
        }
        return true
    } catch {
        return false
    }
}

AIB_ClickMoreActionsInAllIDEWindows() {
    sourceHwnd := 0
    try sourceHwnd := WinGetID("A")

    windows := AIB_GetAllIDEWindowHwnds()
    if (windows.Length = 0) {
        ShowCenteredOverlay_Utils("ℹ No open Cursor/VS Code windows", 1800, BANNER_ACCENT_INFO)
        return 0
    }

    clickedCount := 0
    totalCount := windows.Length
    centerHwnd := 0
    centerHwnd := sourceHwnd
    debugReasons := []

    try {
        StandardLoadingBar_Show(
            "⏳ Scanning IDE windows for More actions button...",
            BANNER_ACCENT_INTERMEDIATE, { centerOnHwnd: centerHwnd, textWidth: 640 }
        )

        for hwnd in windows {
            StandardLoadingBar_Update(
                "⏳ Checking " . A_Index . "/" . totalCount . ": " . AIB_GetSafeWindowTitle(hwnd),
                BANNER_ACCENT_INTERMEDIATE
            )
            failReason := ""
            if (AIB_ClickMoreActionsButtonInWindow(hwnd, &failReason)) {
                clickedCount += 1
                break
            }
            if (failReason != "" && debugReasons.Length < 3)
                debugReasons.Push(AIB_GetSafeWindowTitle(hwnd) . " => " . failReason)
        }
    } finally {
        StandardLoadingBar_Hide(0)
    }

    if (clickedCount > 0)
        ShowCenteredOverlay_Utils("✅ Clicked More actions and stopped search", 1800, BANNER_ACCENT_SUCCESS)
    else {
        msg := "ℹ No More actions button found in chat confirmation"
        if (debugReasons.Length > 0) {
            for reasonLine in debugReasons
                msg .= "`n" . reasonLine
        }
        ShowCenteredOverlay_Utils(msg, 3200, BANNER_ACCENT_INFO)
    }

    return clickedCount
}

AIB_RestoreWindowFocus(hwnd) {
    if (!hwnd)
        return false
    if !WinExist("ahk_id " hwnd)
        return false
    try {
        WinActivate("ahk_id " hwnd)
        WinWaitActive("ahk_id " hwnd, , 1.5)
        return true
    } catch {
        return false
    }
}

AIB_GetAllIDEWindowHwnds() {
    ; WinGetList() returns windows in z-order, so this keeps most-recent first.
    windows := []
    for hwnd in WinGetList() {
        try {
            procName := WinGetProcessName("ahk_id " hwnd)
            if (procName != "Cursor.exe" && procName != "Code.exe")
                continue

            winTitle := WinGetTitle("ahk_id " hwnd)
            if (!winTitle)
                continue

            windows.Push(hwnd)
        } catch {
            continue
        }
    }
    return windows
}

AIB_ClickMoreActionsButtonInWindow(hwnd, &failReason := "") {
    global UIA
    failReason := ""

    try {
        WinActivate("ahk_id " hwnd)
        WinWaitActive("ahk_id " hwnd, , 1.2)
    } catch {
        failReason := "activate failed"
        return false
    }
    Sleep(120)

    moreBtn := AIB_FindMoreActionsButtonInWindow(hwnd)
    if (!moreBtn) {
        failReason := "button not found"
        return false
    }

    try {
        if (moreBtn.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)) {
            moreBtn.InvokePattern.Invoke()
            Sleep(400)  ; Let UI process the invoke
            Send("{Down}")  ; Send down arrow to navigate menu
            return true
        }
    } catch {
    }

    try {
        moreBtn.Click()
        Sleep(400)  ; Let UI process the click
        Send("{Down}")  ; Send down arrow to navigate menu
        return true
    } catch {
        failReason := "invoke/click failed"
        return false
    }
}

AIB_GetSafeWindowTitle(hwnd) {
    try {
        t := WinGetTitle("ahk_id " hwnd)
        t := Trim(t)
        if (t = "")
            return "Untitled"
        if (StrLen(t) > 52)
            return SubStr(t, 1, 49) . "..."
        return t
    } catch {
        return "Unknown Window"
    }
}

AIB_FindMoreActionsButtonInWindow(hwnd) {
    global UIA

    try root := UIA.ElementFromHandle(hwnd)
    catch
        return 0
    if (!root)
        return 0

    ; Hard constraint: More Actions must be adjacent to the Allow button.
    ; No global/top-bar fallback to avoid wrong clicks.
    return AIB_FindMoreActionsNearAllow(root)
}

AIB_FindMoreActionsNearAllow(root) {
    if (!root)
        return 0

    try {
        allBtns := root.FindAll({ Type: 50000 })
        for allowBtn in allBtns {
            allowName := ""
            try allowName := allowBtn.Name
            if (!AIB_IsAllowButtonName(allowName))
                continue

            ; 1) Same immediate container as Allow.
            try p1 := allowBtn.GetParentElement()
            catch {
                p1 := 0
            }
            cand := AIB_FindMoreActionsInContainer(p1)
            if (cand)
                return cand

            ; 2) One level above (some builds wrap each button inside its own group).
            try p2 := p1 ? p1.GetParentElement() : 0
            catch {
                p2 := 0
            }
            cand := AIB_FindMoreActionsInContainer(p2)
            if (cand && AIB_ButtonHasAllowInNearbyContext(cand, 2))
                return cand
        }

        ; 3) Fallback: search directly in chat-confirmation-widget-container
        try {
            dlg := AIB_FindChatConfirmationDialog(root)
            if (dlg) {
                moreBtn := AIB_FindMoreActionsInContainer(dlg)
                if (moreBtn)
                    return moreBtn
            }
        } catch {
        }
    } catch {
    }

    return 0
}

AIB_FindMoreActionsInContainer(containerEl) {
    if (!containerEl)
        return 0
    try {
        buttons := containerEl.FindAll({ Type: 50000 })
        for btn in buttons {
            nm := ""
            try nm := btn.Name
            if (!AIB_IsMoreActionsName(nm))
                continue
            cn := ""
            try cn := btn.ClassName
            ; Accept button if name matches "More Actions" (class check is secondary)
            if (cn && AIB_IsPreferredDialogMoreActionsClass(cn))
                return btn
        }
        ; Fallback: just return first "More Actions" button even if class doesn't match perfectly
        buttons := containerEl.FindAll({ Type: 50000 })
        for btn in buttons {
            nm := ""
            try nm := btn.Name
            if (AIB_IsMoreActionsName(nm))
                return btn
        }
    } catch {
    }
    return 0
}

AIB_FindChatConfirmationDialog(root) {
    if (!root)
        return 0
    try {
        ; Search for chat-confirmation-widget-container
        allGroups := root.FindAll({ Type: 50026 })
        for grp in allGroups {
            cn := ""
            try cn := grp.ClassName
            if (InStr(cn, "chat-confirmation-widget-container"))
                return grp
        }
    } catch {
    }
    return 0
}

AIB_FindMoreActionsInDialog(dlg) {
    if (!dlg)
        return 0
    try {
        ; Most reliable path: locate Allow, then search in the same local container.
        allowBtn := dlg.FindFirst({
            Type: 50000,
            Name: "Allow",
            matchmode: "Substring"
        })
        if (allowBtn) {
            try localGroup := allowBtn.GetParentElement()
            catch {
                localGroup := 0
            }
            if (localGroup) {
                try {
                    btn := localGroup.FindFirst({
                        Type: 50000,
                        ClassName: "monaco-dropdown-button",
                        matchmode: "Substring"
                    })
                    if (btn)
                        return btn
                } catch {
                }
            }
        }

        ; Primary path: exact class used by the chat confirmation dropdown button.
        btn := dlg.FindFirst({
            Type: 50000,
            ClassName: "monaco-button small monaco-dropdown-button codicon codicon-drop-down-button",
            matchmode: "Substring"
        })
        if (btn && AIB_ButtonHasAllowInNearbyContext(btn))
            return btn

        buttons := dlg.FindAll({ Type: 50000 })
        for btn in buttons {
            btnName := ""
            try btnName := btn.Name
            if (!AIB_IsMoreActionsName(btnName))
                continue
            cn := ""
            try cn := btn.ClassName
            if (!AIB_IsPreferredDialogMoreActionsClass(cn))
                continue
            if (AIB_ButtonHasAllowInNearbyContext(btn))
                return btn
        }
    } catch {
    }
    return 0
}

AIB_IsMoreActionsName(btnName) {
    n := Trim(btnName)
    if (n = "More Actions" || n = "More Actions..." || InStr(n, "More Actions"))
        return true
    return false
}

AIB_IsAllowButtonName(btnName) {
    n := StrLower(Trim(btnName))
    if (n = "")
        return false
    ; Accept "Allow", "Allow Once", "Allow (Ctrl+Enter)", etc.
    ; Check if button name starts with "allow" or contains " allow" as word.
    if (!InStr(n, "allow"))
        return false
    if (AIB_IsAllowNegativeButtonName(n))
        return false
    return true
}

AIB_IsAllowNegativeButtonName(btnName) {
    n := StrLower(Trim(btnName))
    if (n = "")
        return false

    ; Explicitly reject known non-target actions that contain "allow" token.
    if (InStr(n, "proceed without executing"))
        return true
    if (InStr(n, "do not allow"))
        return true
    if (InStr(n, "don't allow"))
        return true
    if (InStr(n, "deny") || InStr(n, "decline"))
        return true
    return false
}

AIB_IsPreferredDialogMoreActionsClass(className) {
    cn := Trim(className)
    if (cn = "")
        return false
    if (InStr(cn, "monaco-dropdown-button") && InStr(cn, "drop-down-button"))
        return true
    return false
}

AIB_ButtonHasAllowInNearbyContext(btn, maxDepth := 6) {
    cur := btn
    loop maxDepth {
        if (!cur)
            return false
        try {
            allowBtn := cur.FindFirst({
                Type: 50000,
                Name: "Allow",
                matchmode: "Substring"
            })
            if (allowBtn)
                return true
        } catch {
        }

        try cur := cur.GetParentElement()
        catch {
            return false
        }
    }
    return false
}

AIB_IsInChatConfirmationDialog(el, maxDepth := 12) {
    ; In the chat-confirmation-widget-container search path, buttons are already in context.
    ; This is a fast validation that we're not in a false positive scenario.
    if (!el)
        return false

    cur := el
    loop maxDepth {
        if (!cur)
            break
        cls := ""
        try cls := cur.ClassName
        if (InStr(cls, "chat-confirmation-widget-container"))
            return true
        try cur := cur.GetParentElement()
        catch
            break
    }

    ; Fallback: if we can't validate parent, assume it's OK if reached from
    ; AIB_FindAllowButtonInChatConfirmation (which already confirmed the container)
    return true
}

AIB_IsInAiChatAllowContext(btn) {
    if (!btn)
        return false

    if (!AIB_IsInChatConfirmationDialog(btn))
        return false

    nm := ""
    try nm := btn.Name
    if (AIB_IsAllowNegativeButtonName(nm))
        return false

    return true
}

AIB_FindAllowButtonInDocumentSubtree(root, searchFrame := 0) {
    if (!root)
        return 0

    try {
        docs := root.FindAll({ Type: 50030 })
        for doc in docs {
            if (!doc)
                continue
            btn := AIB_GetBestAllowButton(doc, searchFrame)
            if (btn)
                return btn
        }
    } catch {
    }

    return 0
}

AIB_TryOcrAllowClick(hwnd, root := 0, &usedRoute := "") {
    global g_AIB_AllowWatcherEnableOcrFallback, g_AIB_AllowWatcherOcrCooldownMs, g_AIB_AllowWatcherLastOcrTickByHwnd
    global UIA
    usedRoute := "none"

    if (!g_AIB_AllowWatcherEnableOcrFallback)
        return false
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false

    key := String(hwnd)
    now := A_TickCount
    lastTick := g_AIB_AllowWatcherLastOcrTickByHwnd.Has(key) ? g_AIB_AllowWatcherLastOcrTickByHwnd[key] : 0
    if (now - lastTick < g_AIB_AllowWatcherOcrCooldownMs)
        return false

    if (!root) {
        try root := UIA.ElementFromHandle(hwnd)
        catch {
            root := 0
        }
    }
    if (!root)
        return false

    if (!AIB_HasChatConfirmationAcceptHint(root))
        return false

    frame := AIB_GetPreferredAllowSearchFrame(root, hwnd)
    if (!frame)
        return false

    ocrText := ""
    x := 0
    y := 0
    if (!AIB_RunAllowOcrProbe(frame, &ocrText, &x, &y))
        return false

    n := StrLower(ocrText)
    if (RegExMatch(n, "(^|\\W)allow(\\W|$)") <= 0)
        return false
    if (InStr(n, "proceed without executing") || InStr(n, "do not allow") || InStr(n, "don't allow"))
        return false
    if (!InStr(n, "ctrl+enter") && !InStr(n, "control+enter") && !InStr(n, "chat confirmation required"))
        return false

    SoundPlay(A_ScriptDir . "\sounds\clicking-allow.wav")
    prevForegroundHwnd := AIB_GetForegroundWindowForRestore()
    AIB_PrepareWindowForAllowClick(hwnd)
    try {
        ControlClick("x" x " y" y, "ahk_id " hwnd, , "Left", 1, "NA Pos")
        g_AIB_AllowWatcherLastOcrTickByHwnd[key] := now
        usedRoute := "ocr-controlclick"
        AIB_AllowDebug_Write("click-ocr hwnd=" hwnd " x=" x " y=" y " text='" ocrText "'")
        Sleep(220)
        AIB_RestoreForegroundWindow(prevForegroundHwnd, hwnd)
        return true
    } catch {
        AIB_AllowDebug_Write("click-ocr-failed hwnd=" hwnd)
    }

    AIB_RestoreForegroundWindow(prevForegroundHwnd, hwnd)

    return false
}

AIB_RunAllowOcrProbe(frame, &ocrText := "", &clickX := 0, &clickY := 0) {
    global g_AIB_AllowWatcherOcrScriptPath
    ocrText := ""
    clickX := 0
    clickY := 0

    if (!FileExist(g_AIB_AllowWatcherOcrScriptPath))
        return false

    w := Max(1, frame.r - frame.l)
    h := Max(1, frame.b - frame.t)
    tmpOut := A_Temp . "\\aib_allow_ocr_" . A_TickCount . ".txt"

    cmd := "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"" g_AIB_AllowWatcherOcrScriptPath "`""
    cmd .= " -X " frame.l " -Y " frame.t " -Width " w " -Height " h
    cmd .= " -OutFile `"" tmpOut "`""

    ok := false
    try {
        exitCode := RunWait(cmd, "", "Hide")
        if (exitCode = 0 && FileExist(tmpOut)) {
            content := Trim(FileRead(tmpOut, "UTF-8"))
            if (content != "") {
                parts := StrSplit(content, "|")
                if (parts.Length >= 4 && parts[1] = "FOUND") {
                    clickX := Integer(parts[2])
                    clickY := Integer(parts[3])
                    ocrText := parts[4]
                    ok := true
                }
            }
        }
    } catch {
        ok := false
    }

    try {
        if (FileExist(tmpOut))
            FileDelete(tmpOut)
    } catch {
    }

    return ok
}

; =============================================================================
; Helper function to center mouse on the active window
; =============================================================================
CenterMouse() {
    hwnd := WinExist("A")
    if hwnd
        AL_CenterMouseOnHwnd(hwnd)
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
