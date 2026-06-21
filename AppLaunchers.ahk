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

; [AppLaunchers module] Chrome, WhatsApp, YouTube, Cursor launch hotkeys -> AppLaunchers\launch_hotkeys.ahk
#include %A_ScriptDir%\AppLaunchers\launch_hotkeys.ahk

; [AppLaunchers module] Wikipedia selector globals and article list -> AppLaunchers\wikipedia_globals.ahk
#include %A_ScriptDir%\AppLaunchers\wikipedia_globals.ahk

; [AppLaunchers module] Wikipedia focus monitor and input guard -> AppLaunchers\wikipedia_focus_guard.ahk
#include %A_ScriptDir%\AppLaunchers\wikipedia_focus_guard.ahk

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
