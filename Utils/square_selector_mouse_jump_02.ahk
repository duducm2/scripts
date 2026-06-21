; =============================================================================
; Utils module: square_selector_mouse_jump_02.ahk
; Square selector mouse jump (part 2)
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

DestroyGuiArray(guis) {
    if (!guis || guis.Length = 0) {
        return
    }
    for gui in guis {
        try {
            if (IsObject(gui) && gui.Hwnd) {
                gui.Destroy()
            }
        } catch {
            ; Silently ignore errors
        }
    }
    guis.Length := 0  ; Clear array efficiently
}

; Helper function to cleanup direction indicator squares
CleanupDirectionIndicators() {
    global g_DirectionIndicatorGuis
    DestroyGuiArray(g_DirectionIndicatorGuis)
}

; Factory function to create a handler that properly captures the index
; This ensures each handler gets its own copy of the index value
CreateSquareSelectorHandler(index) {
    ; Return a function that captures the index value at creation time
    return (*) => SelectSquareByIndex(index)
}

; Function to setup hotkey listeners for letter keys
; Uses individual hotkeys that are only active when square selector is shown
SetupLetterKeyListener() {
    global g_SquareSelectorLetters, g_SquareSelectorHotkeyHandlers

    ; Clear any existing handlers
    g_SquareSelectorHotkeyHandlers := []

    ; Create a handler for each letter using factory function
    ; This ensures proper closure capture - each handler gets its own index value
    for i, letter in g_SquareSelectorLetters {
        ; Use factory function to create handler with properly captured index
        handler := CreateSquareSelectorHandler(i)

        ; Store handler reference for cleanup (optional, but good practice)
        g_SquareSelectorHotkeyHandlers.Push({ letter: letter, handler: handler })

        ; Enable both uppercase and lowercase versions
        Hotkey(letter, handler, "On")
        Hotkey(StrLower(letter), handler, "On")
    }
}

; Helper function to disable letter hotkeys (used when entering loop mode)
DisableLetterHotkeys() {
    global g_SquareSelectorLetters

    ; Disable all letter hotkeys
    for letter in g_SquareSelectorLetters {
        try {
            Hotkey(letter, "Off")
            Hotkey(StrLower(letter), "Off")
        } catch {
            ; Silently ignore errors if hotkey doesn't exist
        }
    }
}

; Function to toggle click mode and update square colors
ToggleClickMode() {
    global g_SquareSelectorClickMode, g_SquareSelectorActive, g_SquareSelectorGuis

    ; Only toggle if squares are visible
    if (!g_SquareSelectorActive) {
        return
    }

    ; Toggle click mode flag
    g_SquareSelectorClickMode := !g_SquareSelectorClickMode

    ; Update all square colors based on click mode
    newColor := g_SquareSelectorClickMode ? "0000FF" : "FF0000"  ; Blue or Red
    for gui in g_SquareSelectorGuis {
        try {
            if (IsObject(gui) && gui.Hwnd) {
                gui.BackColor := newColor
                ; Force redraw by hiding and showing
                gui.Show("Hide")
                gui.Show("NA")
            }
        } catch {
            ; Silently ignore errors
        }
    }
}

; Handler for CTRL key to toggle click mode
HandleCtrlToggle() {
    global g_SquareSelectorActive
    ; Only toggle if squares are active
    if (g_SquareSelectorActive) {
        ToggleClickMode()
    }
}

; Handler for letter key press - uses index directly to avoid matching issues
SelectSquareByIndex(index) {
    global g_SquareSelectorActive, g_SquareSelectorPositions, g_ActiveDirection
    global g_SquareSelectorLoopMode, g_SquareSelectorLock

    ; Double-check that selector is active (safety check)
    if (!g_SquareSelectorActive) {
        return
    }

    ; Verify positions array is valid
    if (!g_SquareSelectorPositions || g_SquareSelectorPositions.Length = 0) {
        ; Positions array is empty, cleanup and abort
        CleanupSquareSelector()
        return
    }

    ; Validate index
    if (index < 1 || index > g_SquareSelectorPositions.Length) {
        CleanupSquareSelector()
        return
    }

    ; Get the position for this square (index is 1-based)
    targetPos := g_SquareSelectorPositions[index]

    ; Move mouse to the center of the selected square
    DllCall("SetCursorPos", "int", targetPos.x, "int", targetPos.y)

    ; Check if click mode is active
    global g_SquareSelectorClickMode
    if (g_SquareSelectorClickMode) {
        ; Click mode: perform a click and exit completely
        ; STEP 1: Store target position before cleanup (targetPos is already stored)

        ; STEP 2: Immediately disable all hotkeys and cancel ALL timers
        DisableLetterHotkeys()
        global g_SquareSelectorTimer, g_SquareSelectorBackupTimer, g_OldSquaresCleanupTimer
        if (g_SquareSelectorTimer) {
            SetTimer(g_SquareSelectorTimer, 0)
            g_SquareSelectorTimer := false
        }
        if (g_SquareSelectorBackupTimer) {
            SetTimer(g_SquareSelectorBackupTimer, 0)
            g_SquareSelectorBackupTimer := false
        }
        ; Cancel the old squares cleanup timer from ShowSquareSelector
        if (g_OldSquaresCleanupTimer) {
            SetTimer(g_OldSquaresCleanupTimer, 0)
            g_OldSquaresCleanupTimer := false
        }
        ; Clear start time
        global g_SquareSelectorStartTime
        g_SquareSelectorStartTime := 0

        ; Disable other hotkeys immediately
        try {
            Hotkey("Ctrl", "Off")
        } catch {
            ; Ignore
        }
        DisableDirectionSwitchHotkeys()
        try {
            Hotkey("Escape", "Off")
        } catch {
            ; Ignore
        }
        Utils_EnsureGlobalEscapeHotkey()

        ; STEP 3: Destroy all square GUIs immediately and aggressively
        ; This must happen BEFORE the click so squares don't block it
        global g_SquareSelectorGuis
        ; Destroy all squares in the array
        for gui in g_SquareSelectorGuis {
            try {
                if (IsObject(gui) && gui.Hwnd) {
                    ; Force immediate destruction - no hiding, just destroy
                    gui.Destroy()
                }
            } catch {
                ; Silently ignore errors
            }
        }
        ; Clear arrays immediately
        g_SquareSelectorGuis := []
        g_SquareSelectorPositions := []

        ; Also destroy direction indicators immediately
        CleanupDirectionIndicators()

        ; Brief delay to ensure GUI destruction is complete
        Sleep 15

        ; STEP 4: Wait briefly for GUI cleanup to complete
        Sleep 25

        ; STEP 5: Find window at target position (now that squares are gone)
        targetHwnd := DllCall("WindowFromPoint", "Int64", (targetPos.y << 32) | (targetPos.x & 0xFFFFFFFF), "Ptr")
        if (targetHwnd) {
            ; Get the root window (in case we got a child window)
            rootHwnd := DllCall("GetAncestor", "Ptr", targetHwnd, "UInt", 2, "Ptr")  ; GA_ROOT = 2
            if (rootHwnd) {
                targetHwnd := rootHwnd
            }
            ; Activate the window
            try {
                WinActivate("ahk_id " . targetHwnd)
                WinWaitActive("ahk_id " . targetHwnd, , 0.35)
            } catch {
                ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
            }
        }

        ; STEP 6: Move mouse and click
        DllCall("SetCursorPos", "int", targetPos.x, "int", targetPos.y)
        Sleep 40

        ; Update last mouse click tick to prevent MonitorActiveWindow interference
        try {
            g_LastMouseClickTick := A_TickCount
        } catch {
            ; Ignore if variable doesn't exist
        }

        ; Perform the click
        Click

        ; STEP 8: Final cleanup - ensure everything is reset
        ; Double-check that all squares are destroyed (defensive cleanup)
        global g_SquareSelectorGuis
        if (g_SquareSelectorGuis.Length > 0) {
            for gui in g_SquareSelectorGuis {
                try {
                    if (IsObject(gui) && gui.Hwnd) {
                        gui.Destroy()
                    }
                } catch {
                    ; Ignore
                }
            }
            g_SquareSelectorGuis := []
        }

        ; Reset all state flags
        g_SquareSelectorLock := false
        g_ActiveDirection := ""
        g_SquareSelectorClickMode := false
        g_SquareSelectorActive := false
        g_SquareSelectorLoopMode := false

        ; Final cleanup of direction indicators (defensive)
        CleanupDirectionIndicators()

        return
    }

    ; Normal mode: Keep letter/number squares visible - don't destroy them
    ; Store the current direction before cleanup (for predicting next direction)
    currentDirection := g_ActiveDirection

    ; Cancel timeout timer since we're entering loop mode
    global g_SquareSelectorTimer, g_SquareSelectorBackupTimer
    if (g_SquareSelectorTimer) {
        SetTimer(g_SquareSelectorTimer, 0)
        g_SquareSelectorTimer := false
    }
    if (g_SquareSelectorBackupTimer) {
        SetTimer(g_SquareSelectorBackupTimer, 0)
        g_SquareSelectorBackupTimer := false
    }
    ; Clear start time since we're entering loop mode
    global g_SquareSelectorStartTime
    g_SquareSelectorStartTime := 0

    ; Predict user wants to continue in same direction - show new squares immediately
    ; This speeds up the workflow (user doesn't need to press arrow key)
    ; Keep old squares visible - don't destroy them, just show new ones
    if (currentDirection) {
        ; Small delay to ensure mouse position is stable
        Sleep 50

        ; Store old squares temporarily so we can clean them up after showing new ones
        global g_SquareSelectorGuis
        oldSquares := g_SquareSelectorGuis.Clone()

        ; Automatically show new squares in the same direction
        ; The mouse is now at the selected square position, so new squares will continue from there
        ; ShowSquareSelector will try to clean up, but we'll preserve old squares
        ShowSquareSelector(currentDirection)

        ; Clean up old squares after a brief delay to allow new squares to appear
        SetTimer(() => CleanupOldSquareGuis(oldSquares), -100)

        ; Cancel the timeout timer that ShowSquareSelector set up - we're in loop mode, no timeout
        global g_SquareSelectorTimer, g_SquareSelectorBackupTimer
        if (g_SquareSelectorTimer) {
            SetTimer(g_SquareSelectorTimer, 0)
            g_SquareSelectorTimer := false
        }
        if (g_SquareSelectorBackupTimer) {
            SetTimer(g_SquareSelectorBackupTimer, 0)
            g_SquareSelectorBackupTimer := false
        }
        ; Clear start time since we're in loop mode
        global g_SquareSelectorStartTime
        g_SquareSelectorStartTime := 0
    } else {
        ; No direction - keep squares visible (don't destroy them)
        ; The algorithm finishes but letters remain displayed
        ; Just disable hotkeys but keep squares visible
        DisableLetterHotkeys()
        try {
            Hotkey("Ctrl", "Off")
        } catch {
            ; Ignore
        }
        DisableDirectionSwitchHotkeys()
        try {
            Hotkey("Escape", "Off")
        } catch {
            ; Ignore
        }
        Utils_EnsureGlobalEscapeHotkey()
        ; Don't destroy squares - keep them visible
        g_SquareSelectorActive := false
        g_SquareSelectorLock := false
    }

    ; Show direction indicator squares AFTER new squares are shown
    ; (ShowSquareSelector calls CleanupSquareSelector which removes direction indicators,
    ;  so we need to show them after to prevent blinking/vanishing)
    ShowDirectionIndicators()

    ; Enter loop mode
    ; Letter hotkeys are now re-enabled by ShowSquareSelector
    ; Disable letter/number hotkeys will be handled by loop mode handlers

    ; Disable direction switch hotkeys before enabling loop mode hotkeys
    DisableDirectionSwitchHotkeys()

    ; Set loop mode flag
    g_SquareSelectorLoopMode := true

    ; Enable loop mode hotkeys (Escape and arrow keys)
    EnableLoopModeHotkeys()

    ; DO NOT clear g_ActiveDirection - needed for context
    ; DO NOT release lock - maintained during loop mode
}

; Simplified helper function to handle direction hotkey
HandleDirectionHotkey(direction) {
    ; TEST: Uncomment next line to verify hotkey is firing
    ; MsgBox "Hotkey triggered: " . direction, "Debug"

    global g_SquareSelectorActive, g_ActiveDirection, g_SquareSelectorTimer
    global g_SquareSelectorLock, g_SquareSelectorClickMode

    ; STEP 0: Preserve click mode state BEFORE cleanup (so blue squares stay blue when changing direction)
    preservedClickMode := g_SquareSelectorClickMode

    ; STEP 1: IMMEDIATELY disable active flag and clear positions
    ; This prevents letter hotkeys from using old positions
    g_SquareSelectorActive := false
    global g_SquareSelectorPositions
    g_SquareSelectorPositions := []

    ; STEP 2: Increment session ID to invalidate any old timers
    global g_SquareSelectorSessionID
    g_SquareSelectorSessionID++

    ; STEP 3: Cancel any existing timers FIRST (prevents old timers from cleaning up new squares)
    if (g_SquareSelectorTimer) {
        SetTimer(g_SquareSelectorTimer, 0)
        g_SquareSelectorTimer := false
    }
    global g_SquareSelectorBackupTimer
    if (g_SquareSelectorBackupTimer) {
        SetTimer(g_SquareSelectorBackupTimer, 0)
        g_SquareSelectorBackupTimer := false
    }

    ; STEP 4: Clean up old selector completely (GUIs, hotkeys, etc.)
    CleanupSquareSelector()

    ; STEP 5: Reset lock to ensure clean state
    g_SquareSelectorLock := false

    ; STEP 6: Wait a bit for cleanup to complete and brief delay before showing new squares
    Sleep 80

    ; STEP 7: Set new active direction
    g_ActiveDirection := StrLower(direction)

    ; STEP 8: Show new squares (delay already included above)

    ; STEP 9: Disable loop mode if it was active (transitioning from loop mode)
    global g_SquareSelectorLoopMode
    if (g_SquareSelectorLoopMode) {
        DisableLoopModeHotkeys()
        g_SquareSelectorLoopMode := false
    }

    ; STEP 10: Show the new squares (with session ID)
    ShowSquareSelector(g_ActiveDirection)

    ; STEP 11: Restore click mode state if it was active (so blue squares remain blue)
    if (preservedClickMode) {
        g_SquareSelectorClickMode := true
        ; Update all square colors to blue to reflect click mode
        global g_SquareSelectorGuis
        for gui in g_SquareSelectorGuis {
            try {
                if (IsObject(gui) && gui.Hwnd) {
                    gui.BackColor := "0000FF"  ; Blue
                    ; Force redraw by hiding and showing
                    gui.Show("Hide")
                    gui.Show("NA")
                }
            } catch {
                ; Silently ignore errors
            }
        }
    }
}

; Helper function to cancel squares (works in both initial mode and loop mode)
CancelSquareSelector() {
    global g_SquareSelectorLoopMode, g_SquareSelectorLock, g_ActiveDirection
    global g_SquareSelectorActive

    ; Only handle if squares are active
    if (!g_SquareSelectorActive && !g_SquareSelectorLoopMode) {
        return
    }

    ; ALWAYS disable loop mode hotkeys (including mouse button hotkeys) to prevent blocking clicks
    try {
        DisableLoopModeHotkeys()
    } catch {
        ; Silently ignore if there's an error
    }

    ; Cleanup completely
    CleanupSquareSelector()

    ; Reset all state
    g_SquareSelectorLoopMode := false
    g_SquareSelectorLock := false
    g_ActiveDirection := ""
}

; Helper function to exit loop mode (shared by Escape and mouse handlers)
ExitLoopMode() {
    global g_SquareSelectorLoopMode, g_SquareSelectorLock, g_ActiveDirection

    ; Only handle if in loop mode
    if (!g_SquareSelectorLoopMode) {
        return
    }

    ; Disable loop mode hotkeys
    DisableLoopModeHotkeys()

    ; Cleanup completely
    CleanupSquareSelector()

    ; Reset all state
    g_SquareSelectorLoopMode := false
    g_SquareSelectorLock := false
    g_ActiveDirection := ""
}

; Mouse click handlers for loop mode (exit and forward the click)
HandleLoopModeLButton() {
    global g_SquareSelectorLoopMode
    if (g_SquareSelectorLoopMode) {
        ExitLoopMode()
        ; Send the click after exiting loop mode
        Click("Left")
    }
}

HandleLoopModeRButton() {
    global g_SquareSelectorLoopMode
    if (g_SquareSelectorLoopMode) {
        ExitLoopMode()
        ; Send the click after exiting loop mode
        Click("Right")
    }
}

HandleLoopModeMButton() {
    global g_SquareSelectorLoopMode
    if (g_SquareSelectorLoopMode) {
        ExitLoopMode()
        ; Send the click after exiting loop mode
        Click("Middle")
    }
}

; Helper function to enable direction switch hotkeys (arrow keys for switching directions immediately)
EnableDirectionSwitchHotkeys() {
    ; Enable arrow key hotkeys for immediate direction switching (without modifiers)
    Hotkey("Right", (*) => HandleDirectionHotkey("Right"), "On")
    Hotkey("Left", (*) => HandleDirectionHotkey("Left"), "On")
    Hotkey("Down", (*) => HandleDirectionHotkey("Down"), "On")
    Hotkey("Up", (*) => HandleDirectionHotkey("Up"), "On")
}

; Helper function to disable direction switch hotkeys
DisableDirectionSwitchHotkeys() {
    ; Disable arrow key hotkeys
    try {
        Hotkey("Right", "Off")
        Hotkey("Left", "Off")
        Hotkey("Down", "Off")
        Hotkey("Up", "Off")
    } catch {
        ; Silently ignore if hotkeys don't exist
    }
}

; Helper function to enable loop mode hotkeys (Escape, arrow keys, and mouse clicks)
EnableLoopModeHotkeys() {
    ; Enable Escape hotkey for loop mode (uses CancelSquareSelector which works for both modes)
    Hotkey("Escape", (*) => CancelSquareSelector(), "On")

    ; Enable arrow key hotkeys for loop mode (without modifiers)
    Hotkey("Right", (*) => HandleLoopModeRight(), "On")
    Hotkey("Left", (*) => HandleLoopModeLeft(), "On")
    Hotkey("Down", (*) => HandleLoopModeDown(), "On")
    Hotkey("Up", (*) => HandleLoopModeUp(), "On")

    ; Enable mouse click hotkeys to exit loop mode (forward click after exit)
    Hotkey("LButton", (*) => HandleLoopModeLButton(), "On")
    Hotkey("RButton", (*) => HandleLoopModeRButton(), "On")
    Hotkey("MButton", (*) => HandleLoopModeMButton(), "On")
}

; Helper function to disable loop mode hotkeys
DisableLoopModeHotkeys() {
    ; Disable Escape hotkey
    try {
        Hotkey("Escape", "Off")
    } catch {
        ; Silently ignore if hotkey doesn't exist
    }
    Utils_EnsureGlobalEscapeHotkey()

    ; Disable arrow key hotkeys
    try {
        Hotkey("Right", "Off")
        Hotkey("Left", "Off")
        Hotkey("Down", "Off")
        Hotkey("Up", "Off")
    } catch {
        ; Silently ignore if hotkeys don't exist
    }

    ; Disable mouse click hotkeys
    try {
        Hotkey("LButton", "Off")
        Hotkey("RButton", "Off")
        Hotkey("MButton", "Off")
    } catch {
        ; Silently ignore if hotkeys don't exist
    }
}

; Escape key handler for loop mode
HandleEscapeKey() {
    ExitLoopMode()
}

; Loop mode arrow key handlers (only active when in loop mode)
HandleLoopModeRight() {
    global g_SquareSelectorLoopMode
    if (g_SquareSelectorLoopMode) {
        HandleDirectionHotkey("Right")
    }
}

HandleLoopModeLeft() {
    global g_SquareSelectorLoopMode
    if (g_SquareSelectorLoopMode) {
        HandleDirectionHotkey("Left")
    }
}

HandleLoopModeDown() {
    global g_SquareSelectorLoopMode
    if (g_SquareSelectorLoopMode) {
        HandleDirectionHotkey("Down")
    }
}

HandleLoopModeUp() {
    global g_SquareSelectorLoopMode
    if (g_SquareSelectorLoopMode) {
        HandleDirectionHotkey("Up")
    }
}
