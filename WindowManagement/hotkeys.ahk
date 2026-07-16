; =============================================================================
; WindowManagement module: hotkeys.ahk
; Global hotkey bindings (minimize/maximize, window tools menu, move/close/cycle/
; minimize on monitor). Definitions only; handlers live in WindowManagement.ahk.
; Cheat-sheet descriptions: Shift keys/cheat_sheet_registry.ahk GLOBAL_CHEAT_SHEET_RAW
; Extracted verbatim from WindowManagement.ahk; loaded via #include into the
; WindowManagement.ahk process, which remains the entry point / source of truth.
; =============================================================================

; =============================================================================
; Minimize Active Window
; Hotkey: Win+Alt+Shift+6
; Original File: Minimize.ahk
; =============================================================================
#!+6::
{
    WinMinimize "A"
}

; =============================================================================
; Maximize Active Window
; Hotkey: Win+Alt+Shift+M
; Hotkey: Ctrl+Alt+Win+V (ZMK / hardware — same handler)
; Original File: Maximize window.ahk
; =============================================================================
#!+M::
{
    WM_MaximizeActiveWindow()
}

^!#v::
{
    WM_MaximizeActiveWindow()
}

; =============================================================================
; Window tools (maximize lone / hidden list / tile background / exit F11 fullscreen)
; Menu: Win+Alt+Shift+W — direct CAW: Z=[1], 6=[2], Y=[3], P=[4]; X=Snap 50/50 (not in menu)
; =============================================================================
^!#z:: WM_WindowTools_OnMaximizeLonely()
^!#6:: WM_WindowTools_OnShowMinimizedList()
^!#y:: WM_WindowTools_OnTileBackground()
^!#p:: WM_WindowTools_OnExitF11Fullscreen()
^!#x:: WM_SnapHalfPairActiveWindow()
#!+w:: WM_WindowTools_ShowMenu()

; Dev: log taskbar-minimized background scan (Ctrl+Alt+Win+Shift+B)
^!+#b:: WM_DebugBackgroundWindowScan()

; Clip Angel Shift+P / Shift+B: pass-through to native open; AHK only maximizes + foreground afterward.
; Always assist when Clip Angel is running — skipping when it was already "active" left it
; maximized on another monitor / behind other windows with no cleanup pass.
; AutoSlot freeze lives here (WindowManagement includes AutoSlot; Utils does not).
#HotIf WinExist("ahk_exe ClipAngel.exe")
~+p:: {
    try AutoSlot_BeginPlaceFreeze()
    catch {
    }
    try AutoSlot_BeginSwapQuiet(AutoSlot_RECENT_MS)
    catch {
    }
    ClipAngel_EnsureForegroundAfterNativeOpen()
}
~+b:: {
    try AutoSlot_BeginPlaceFreeze()
    catch {
    }
    try AutoSlot_BeginSwapQuiet(AutoSlot_RECENT_MS)
    catch {
    }
    ClipAngel_EnsureForegroundAfterNativeOpen()
}
#HotIf

; =============================================================================
; Move Active Window to Monitor by POSITION (left-to-right order)
; Hotkeys: Ctrl+Alt+Win+A/S/D/F move active window to 1st–4th monitors (left-to-right)
; =============================================================================
; ^!#a vs ^!+#a are distinct hotkeys (Shift in the latter); no #HotIf on physical Shift — avoids desync with Shift keys.ahk.
^!#a:: MoveWinToOrderedMonitor(1)  ; Left-most
^!#s:: MoveWinToOrderedMonitor(2)  ; 2nd from the left
^!#d:: MoveWinToOrderedMonitor(3)  ; 3rd from the left
^!#f:: MoveWinToOrderedMonitor(4)  ; 4th from the left

; Shift variants: close the active window on the specified monitor
; Note: ^!+#a shares the letter with ^!#a (move to M1); AHK treats them as separate chords.
; Top-row 1: *^!+#1 + *^!+#SC002 at end of script (wildcard); ^!+#g/^!+#z here if IDE on ordinal 2 ignores digit 1.
; Numpad1 / Z / G: fallbacks when the top-row 1 chord fails on the IDE / ordinal-2 monitor (Ctrl+Alt+Win+Shift+G).
^!+#a:: CloseWindowOnMonitor(1)  ; Close window on monitor 1
^!+#Numpad1:: CloseWindowOnMonitor(1)
^!+#z:: CloseWindowOnMonitor(1)  ; Alternate close-M1 (no digit 1 — avoids IDE / selector conflicts)
^!+#g:: CloseWindowOnMonitor(1)  ; Reliable close-M1 from center/IDE display when ^!+#1 is eaten
; No-Win close-M1 when Win+ is swallowed (Electron/IDE). $ forces kbd hook. Comment out if another app uses this chord.
$^!+1:: CloseWindowOnMonitor(1)
$^!+g:: CloseWindowOnMonitor(1)
^!+#s:: CloseWindowOnMonitor(2)  ; Close window on monitor 2
^!+#d:: CloseWindowOnMonitor(3)  ; Close window on monitor 3
^!+#f:: CloseWindowOnMonitor(4)  ; Close window on monitor 4

^!#q:: CycleWindowsOnMonitor(1)  ; Cycle windows on monitor 1
^!#w:: CycleWindowsOnMonitor(2)  ; Cycle windows on monitor 2
^!#e:: CycleWindowsOnMonitor(3)  ; Cycle windows on monitor 3
^!#r:: CycleWindowsOnMonitor(4)  ; Cycle windows on monitor 4

; Shift variants: minimize the active window on the specified monitor
^!+#q:: MinimizeWindowOnMonitor(1)  ; Minimize window on monitor 1
^!+#w:: MinimizeWindowOnMonitor(2)  ; Minimize window on monitor 2
^!+#e:: MinimizeWindowOnMonitor(3)  ; Minimize window on monitor 3
^!+#r:: MinimizeWindowOnMonitor(4)  ; Minimize window on monitor 4
