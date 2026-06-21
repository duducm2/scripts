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

; [AppLaunchers module] Phase 3/4 global state for hooks and wiki FSM -> AppLaunchers\config_globals.ahk
#include %A_ScriptDir%\AppLaunchers\config_globals.ahk

; --- Hotkeys & Functions -----------------------------------------------------

; [AppLaunchers module] #!+n context file browser hotkey -> AppLaunchers\hotkey_context_browser.ahk
#include %A_ScriptDir%\AppLaunchers\hotkey_context_browser.ahk

; [AppLaunchers module] Shift+Win+E desktop explorer and UIA helpers -> AppLaunchers\desktop_explorer.ahk
#include %A_ScriptDir%\AppLaunchers\desktop_explorer.ahk

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
