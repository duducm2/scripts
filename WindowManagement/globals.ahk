; =============================================================================
; WindowManagement module: globals.ahk
; Global variable declarations + startup timers (MonitorActiveWindow, background
; excludes init) and the tray test item. Contains top-level auto-execute code, so
; its #include MUST stay at the original position in WindowManagement.ahk.
; Extracted verbatim from WindowManagement.ahk (the entry point / source of truth).
; =============================================================================

; --- Globals & Timers --------------------------------------------------------
global g_LastActiveHwnd := 0
global g_LastMouseClickTick := 0   ; Timestamp of the most recent mouse click (A_TickCount)
global g_WindowCycleIndices := Map()  ; Keeps per-monitor cycling position
global g_WMAutomationSuppressUntil := 0
global g_WMAutomationSuppressReason := ""
global g_WM_MinimizedListGui := false
global g_WM_MinimizedListActive := false
global g_WM_MinimizedListRows := []
global g_WM_MinimizedListEscPollPrev := false
global g_WM_MinimizedCharSequence := ["1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "a", "b", "c", "d", "e", "f",
    "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"]
global g_WM_MinimizedKeyMap := Map()
global g_WM_MinimizedListOpenFile := A_ScriptDir "\.cursor\wm_minimized_list_open"
global g_WM_MinimizedListCloseRequestFile := A_ScriptDir "\.cursor\wm_minimized_list_close_request"
global g_WM_MinimizedListCloseCheckTimer := ""
global g_WM_MinimizedHotkeyHandlers := []
global g_WM_MinimizedKeysPollTimer := ""
global g_WM_MinimizedKeysPollPrev := Map()
global g_WM_MinimizedKeysPollCallbacks := Map()
global g_WM_MinimizedListRefreshing := false
global g_WM_MinimizedListTrackTimer := ""
global g_WM_MinimizedListLastForegroundMonitorIdx := 0
global g_WM_BackgroundTitleExcludes := []
global g_WM_BackgroundTitleExcludesReady := false
global g_WM_MinimizedListCollectForeHwnd := 0
global g_WM_MinimizedListExcludePickerActive := false
global g_WM_MinimizedListExcludePickerRows := []
global g_WM_MinimizedListExcludePickerMap := Map()
global g_WM_MinimizedListExcludePickerDigitSequence := ["1", "2", "3", "4", "5", "6", "7", "8", "9", "0"]
global g_WM_MinimizedListCloseModeArmed := false
global g_WM_MinimizedListCharActionLock := false
global g_WM_MinimizedListLastCharActionTick := 0
global g_WM_MinimizedListLastCharActionKey := ""
global g_WM_MinimizedListLastCloseArmTick := 0
global g_WM_WindowToolsShowListLock := false
global g_WM_WindowToolsShowListLastTick := 0
global g_WM_LastEnumerateStats := Map()
global g_WM_LastBackgroundCollectStats := Map()
global g_WM_BackgroundScanBannerTick := 0
; When daemon is used, foreground is driven by daemon cache (lower-frequency check); else legacy 100ms polling
if (WM_UsesAutomationDaemon())
    SetTimer MonitorActiveWindow, 250
else
    SetTimer MonitorActiveWindow, 100
SetTimer(WM_BackgroundTitleExcludes_Init, -1)

; Tray: verify cycle logic without keyboard hooks (compare to ^!#q failures).
A_TrayMenu.Add("Test Cycle M1", (*) => CycleWindowsOnMonitor(1))
