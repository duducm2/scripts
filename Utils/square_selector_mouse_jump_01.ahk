; =============================================================================
; Utils module: square_selector_mouse_jump_01.ahk
; Square selector mouse jump (part 1)
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Square Selector System for Mouse Jump
; Shows 15 red squares with letters in chosen direction, waits for letter selection
; =============================================================================

; Global variables for square selector system
global g_SquareSelectorActive := false
global g_SquareSelectorGuis := []
global g_SquareSelectorPositions := []  ; Array of {x, y} positions for each square
global g_SquareSelectorLetters := ["1", "2", "3", "4", "5", "Q", "W", "E", "R", "T", "A", "S", "D", "F", "G", "Z", "X",
    "C", "V", "B", "6", "7", "8", "9", "0", "y", "u", "i", "o", "p", "h", "j", "k", "l", "n", "m", ",", "."]
global g_SquareSelectorTimer := false
global g_SquareSelectorLetterMap := Map()  ; Map to store letter to index mapping
global g_SquareSelectorSessionID := 0  ; Unique session ID to prevent timer conflicts

; Global array to store hotkey handlers for cleanup
global g_SquareSelectorHotkeyHandlers := []

; Lock flag to prevent multiple square selectors from running simultaneously
global g_SquareSelectorLock := false

; Active direction flag - prevents old selectors from interfering
global g_ActiveDirection := ""

; Loop mode flag - indicates waiting for Escape or arrow key after selection
global g_SquareSelectorLoopMode := false

; Click mode flag - when true, squares are blue and selection will click and exit
global g_SquareSelectorClickMode := false

; Direction indicator GUIs (4 squares around mouse pointer in loop mode)
global g_DirectionIndicatorGuis := []

; Timestamp when squares were last shown (for guaranteed cleanup)
global g_SquareSelectorStartTime := 0

; Backup cleanup timer (guaranteed to fire after 10 seconds)
global g_SquareSelectorBackupTimer := false

; Timer for cleaning up old squares when showing new ones
global g_OldSquaresCleanupTimer := false

; Timer handler for square selector timeout
SquareSelectorTimerHandler(sessionID) {
    global g_SquareSelectorLock, g_ActiveDirection, g_SquareSelectorTimer
    global g_SquareSelectorActive, g_SquareSelectorSessionID

    ; CRITICAL: Check if this timer is for the current session
    ; If session ID doesn't match, this timer is stale and should be ignored
    if (sessionID != g_SquareSelectorSessionID) {
        ; This timer is for an old session, ignore it
        return
    }

    ; Check if selector is still active (might have been cleaned up by new direction)
    if (!g_SquareSelectorActive) {
        ; Already cleaned up, just clear timer reference
        g_SquareSelectorTimer := false
        return
    }

    ; Only cleanup if selector is still active and session matches
    CleanupSquareSelector()
    g_SquareSelectorLock := false
    g_ActiveDirection := ""  ; Clear active direction on timeout
    g_SquareSelectorTimer := false  ; Clear timer reference
}

; Force cleanup function - aggressively destroys all squares regardless of state
; This is a backup mechanism to ensure squares never persist forever
ForceCleanupAllSquares() {
    global g_SquareSelectorGuis, g_DirectionIndicatorGuis
    global g_SquareSelectorActive, g_SquareSelectorLock, g_ActiveDirection
    global g_SquareSelectorLoopMode, g_SquareSelectorClickMode
    global g_SquareSelectorTimer, g_SquareSelectorBackupTimer
    global g_SquareSelectorStartTime

    ; Force disable active flag
    g_SquareSelectorActive := false

    ; Aggressively destroy all square GUIs
    for gui in g_SquareSelectorGuis {
        try {
            if (IsObject(gui)) {
                try {
                    if (gui.Hwnd) {
                        gui.Hide()
                        gui.Destroy()
                    }
                } catch {
                    ; Ignore hide/destroy errors
                }
            }
        } catch {
            ; Silently ignore all errors
        }
    }
    g_SquareSelectorGuis := []

    ; Aggressively destroy all direction indicator GUIs
    DestroyGuiArray(g_DirectionIndicatorGuis)

    ; Cancel all timers
    if (g_SquareSelectorTimer) {
        try {
            SetTimer(g_SquareSelectorTimer, 0)
        } catch {
            ; Ignore
        }
        g_SquareSelectorTimer := false
    }

    if (g_SquareSelectorBackupTimer) {
        try {
            SetTimer(g_SquareSelectorBackupTimer, 0)
        } catch {
            ; Ignore
        }
        g_SquareSelectorBackupTimer := false
    }

    ; Reset all state
    g_SquareSelectorLock := false
    g_ActiveDirection := ""
    g_SquareSelectorLoopMode := false
    g_SquareSelectorClickMode := false
    g_SquareSelectorStartTime := 0

    ; Disable all hotkeys (best effort) to prevent bugs
    try {
        DisableLetterHotkeys()
    } catch {
        ; Ignore
    }
    try {
        DisableDirectionSwitchHotkeys()
    } catch {
        ; Ignore
    }
    try {
        DisableLoopModeHotkeys()
    } catch {
        ; Ignore
    }
    try {
        Hotkey("Ctrl", "Off")
    } catch {
        ; Ignore
    }
    try {
        Hotkey("Escape", "Off")
    } catch {
        ; Ignore
    }
    Utils_EnsureGlobalEscapeHotkey()
}

; Backup timer handler - guaranteed to fire after 7 seconds
BackupCleanupTimer() {
    global g_SquareSelectorStartTime, g_SquareSelectorGuis, g_SquareSelectorBackupTimer
    global g_SquareSelectorActive

    ; If start time is 0, squares have been cleaned up, stop the timer
    if (g_SquareSelectorStartTime == 0) {
        if (g_SquareSelectorBackupTimer) {
            SetTimer(g_SquareSelectorBackupTimer, 0)
            g_SquareSelectorBackupTimer := false
        }
        return
    }

    ; Check if squares have been visible for more than 7 seconds
    elapsed := (A_TickCount - g_SquareSelectorStartTime) / 1000  ; Convert to seconds
    if (elapsed >= 7) {
        ; Force cleanup if squares have been visible for 7+ seconds
        ForceCleanupAllSquares()
        return
    }

    ; If there are no GUIs and not active, cleanup is done, stop timer
    if (g_SquareSelectorGuis.Length = 0 && !g_SquareSelectorActive) {
        ; No GUIs and not active - cleanup is done, stop timer
        if (g_SquareSelectorBackupTimer) {
            SetTimer(g_SquareSelectorBackupTimer, 0)
            g_SquareSelectorBackupTimer := false
        }
        g_SquareSelectorStartTime := 0
    }
}

; Helper to create a timer handler bound to a specific session ID
CreateTimerHandler(sessionID) {
    return () => SquareSelectorTimerHandler(sessionID)
}

; Helper function to cleanup old square GUIs (used by ShowSquareSelector)
CleanupOldSquareGuis(oldGuis) {
    for gui in oldGuis {
        try {
            if (IsObject(gui) && gui.Hwnd) {
                gui.Destroy()
            }
        } catch {
            ; Silently ignore errors
        }
    }
}

; Function to cleanup square selector system
CleanupSquareSelector() {
    global g_SquareSelectorActive, g_SquareSelectorGuis, g_SquareSelectorTimer
    global g_SquareSelectorLetters, g_SquareSelectorHotkeyHandlers
    global g_SquareSelectorLock, g_ActiveDirection, g_SquareSelectorLoopMode

    ; Disable active flag immediately
    g_SquareSelectorActive := false

    ; Disable all letter hotkeys immediately using stored handlers
    for letter in g_SquareSelectorLetters {
        try {
            Hotkey(letter, "Off")
            Hotkey(StrLower(letter), "Off")
        } catch {
            ; Silently ignore errors if hotkey doesn't exist
        }
    }

    ; ALWAYS disable loop mode hotkeys (including mouse button hotkeys) to prevent blocking clicks
    ; This ensures mouse clicks work even if hotkeys were enabled through a race condition
    try {
        DisableLoopModeHotkeys()
    } catch {
        ; Silently ignore if there's an error
    }
    g_SquareSelectorLoopMode := false

    ; Disable CTRL hotkey (click mode toggle)
    try {
        Hotkey("Ctrl", "Off")
    } catch {
        ; Silently ignore if hotkey doesn't exist
    }

    ; Disable direction switch hotkeys
    DisableDirectionSwitchHotkeys()

    ; Disable Escape hotkey
    try {
        Hotkey("Escape", "Off")
    } catch {
        ; Silently ignore if hotkey doesn't exist
    }
    Utils_EnsureGlobalEscapeHotkey()

    ; Reset click mode flag
    global g_SquareSelectorClickMode
    g_SquareSelectorClickMode := false

    ; Clear hotkey handlers array
    g_SquareSelectorHotkeyHandlers := []

    ; Destroy all square GUIs
    DestroyGuiArray(g_SquareSelectorGuis)
    g_SquareSelectorPositions := []

    ; Clean up direction indicator squares
    CleanupDirectionIndicators()

    ; Cancel timer if active
    if (g_SquareSelectorTimer) {
        SetTimer(g_SquareSelectorTimer, 0)
        g_SquareSelectorTimer := false
    }

    ; Cancel backup timer if active
    global g_SquareSelectorBackupTimer
    if (g_SquareSelectorBackupTimer) {
        SetTimer(g_SquareSelectorBackupTimer, 0)
        g_SquareSelectorBackupTimer := false
    }

    ; Cancel old squares cleanup timer if active
    global g_OldSquaresCleanupTimer
    if (g_OldSquaresCleanupTimer) {
        SetTimer(g_OldSquaresCleanupTimer, 0)
        g_OldSquaresCleanupTimer := false
    }

    ; Clear start time
    global g_SquareSelectorStartTime
    g_SquareSelectorStartTime := 0

    ; Release lock and clear active direction to prevent bugs
    ; This ensures the hotkeys can be used again after cleanup
    g_SquareSelectorLock := false
    g_ActiveDirection := ""
}

; Function to show 15 squares with letters in a line in the chosen direction
ShowSquareSelector(direction) {
    global g_SquareSelectorActive, g_SquareSelectorGuis, g_SquareSelectorPositions
    global g_SquareSelectorLetters, g_SquareSelectorLock

    ; Don't clear arrays immediately - preserve old squares
    ; We'll clean them up after showing new ones if needed
    oldGuis := g_SquareSelectorGuis.Clone()
    oldPositions := g_SquareSelectorPositions.Clone()

    ; Clear arrays for new squares
    g_SquareSelectorGuis := []
    g_SquareSelectorPositions := []

    ; Don't call CleanupSquareSelector here - it destroys squares
    ; Instead, just disable hotkeys temporarily
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

    ; Clean up old squares after a brief delay (allows new squares to appear first)
    ; Cancel any existing old squares cleanup timer first
    global g_OldSquaresCleanupTimer
    if (g_OldSquaresCleanupTimer) {
        SetTimer(g_OldSquaresCleanupTimer, 0)
        g_OldSquaresCleanupTimer := false
    }

    if (oldGuis.Length > 0) {
        g_OldSquaresCleanupTimer := () => CleanupOldSquareGuis(oldGuis)
        SetTimer(g_OldSquaresCleanupTimer, -50)
    }

    Sleep 10

    ; Get current mouse position
    pos := GetMousePos()
    startX := pos.x
    startY := pos.y

    ; Configuration
    squareSize := 24  ; Reduced for more precision
    spacing := 20  ; Reduced for more precision
    numSquares := 38  ; Updated to match total characters in g_SquareSelectorLetters

    ; Normalize direction
    directionLower := StrLower(direction)

    ; STEP 1: Calculate all center positions first
    ; First square (1) starts AFTER mouse position, not centered on it
    ; Initial offset: half square size (12px) + spacing (20px) = 32px from mouse position
    ; This ensures the first square's left edge starts after the mouse cursor
    initialOffset := (squareSize / 2.0) + spacing  ; 12 + 20 = 32 pixels

    calculatedPositions := []
    if (directionLower = "right" || directionLower = "left") {
        ; Horizontal line
        directionMultiplier := directionLower = "right" ? 1 : -1
        loop numSquares {
            i := A_Index
            ; Calculate offset for square i
            ; First square (i=1): initialOffset (32px) - starts after mouse
            ; Subsequent squares: initialOffset + (i-1) * (squareSize + spacing)
            ; For i=1: 32px, for i=2: 32 + 44 = 76px, for i=3: 32 + 88 = 120px, etc.
            offset := (initialOffset + (i - 1) * (squareSize + spacing)) * directionMultiplier
            squareCenterX := Round(startX + offset)
            squareCenterY := startY
            calculatedPositions.Push({ x: squareCenterX, y: squareCenterY })
        }
    } else {
        ; Vertical line (up or down)
        directionMultiplier := directionLower = "down" ? 1 : -1
        loop numSquares {
            i := A_Index
            ; Same calculation for vertical: first square starts after mouse
            offset := (initialOffset + (i - 1) * (squareSize + spacing)) * directionMultiplier
            squareCenterX := startX
            squareCenterY := Round(startY + offset)
            calculatedPositions.Push({ x: squareCenterX, y: squareCenterY })
        }
    }

    ; STEP 2: Create all GUIs at once (don't show yet)
    guiArray := []
    loop numSquares {
        i := A_Index
        pos := calculatedPositions[i]

        ; Create square GUI with letter
        squareGui := Gui("+AlwaysOnTop -Caption +ToolWindow")
        ; Color depends on click mode: blue if click mode active, red otherwise
        global g_SquareSelectorClickMode
        squareGui.BackColor := g_SquareSelectorClickMode ? "0000FF" : "FF0000"  ; Blue or Red
        squareGui.SetFont("s8 Bold cFFFFFF", "Segoe UI")  ; White text, bold, smaller for precision

        ; Set GUI margins to 0 to eliminate any padding that could affect centering
        squareGui.MarginX := 0
        squareGui.MarginY := 0

        ; Create text control that perfectly centers the letter
        ; Center = 0x1 (SS_CENTER) for horizontal centering
        ; 0x200 = SS_CENTERIMAGE for vertical centering
        ; 0x201 combines both (SS_CENTER | SS_CENTERIMAGE) for perfect centering
        ; Text control fills entire square (40x40) to ensure proper centering
        letterText := squareGui.AddText("w" . squareSize . " h" . squareSize . " Center 0x201",
            g_SquareSelectorLetters[i])

        ; Calculate top-left position for this square
        guiX := Round(pos.x - squareSize / 2.0)
        guiY := Round(pos.y - squareSize / 2.0)

        ; Store GUI and position info (not shown yet)
        guiArray.Push({ gui: squareGui, x: guiX, y: guiY, calculatedCenter: pos })
    }

    ; STEP 3: Prepare all GUIs (position while hidden for instant showing)
    for guiInfo in guiArray {
        ; Position while hidden (no rendering delay)
        guiInfo.gui.Show("x" . guiInfo.x . " y" . guiInfo.y . " w" . squareSize . " h" . squareSize . " NA Hide")
        ; Set 80% opacity (204 = 80% opacity, 255 = fully opaque, 0 = fully transparent)
        WinSetTransparent(204, guiInfo.gui)
    }

    ; STEP 4: Show all GUIs simultaneously (batch show for instant appearance)
    ; Use Show() instead of SetWindowPos to ensure windows actually appear
    ; Show all windows using Show() - this is more reliable than SetWindowPos
    for guiInfo in guiArray {
        try {
            ; Show window using Show() - ensure it actually appears
            guiInfo.gui.Show("NA")  ; Show without activating
        } catch {
            ; If Show() fails, try using the position again
            try {
                guiInfo.gui.Show("x" . guiInfo.x . " y" . guiInfo.y . " w" . squareSize . " h" . squareSize . " NA")
            }
        }
    }

    ; STEP 5: Brief delay to ensure all GUIs are fully rendered
    Sleep 20  ; Increased delay to ensure windows are fully rendered before querying positions

    ; STEP 6: Query actual GUI positions and store actual centers for mouse jump
    ; Query actual window positions using GetWindowRect to get exact centers
    ; This accounts for any window borders, padding, or DPI adjustments
    for i, guiInfo in guiArray {
        squareGuiObj := guiInfo.gui  ; Use different variable name to avoid conflict
        g_SquareSelectorGuis.Push(squareGuiObj)

        ; Query actual window rectangle using GetWindowRect
        ; This gives us the actual physical pixel coordinates after DPI adjustments
        rect := Buffer(16, 0)  ; RECT structure: left, top, right, bottom (4 ints)
        if (DllCall("GetWindowRect", "ptr", squareGuiObj.Hwnd, "ptr", rect)) {
            ; Extract rectangle coordinates (physical pixels with DPI awareness)
            winLeft := NumGet(rect, 0, "int")
            winTop := NumGet(rect, 4, "int")
            winRight := NumGet(rect, 8, "int")
            winBottom := NumGet(rect, 12, "int")

            ; Calculate actual center from window rectangle
            actualCenterX := winLeft + (winRight - winLeft) / 2
            actualCenterY := winTop + (winBottom - winTop) / 2

            ; Store actual center position (rounded to nearest pixel)
            g_SquareSelectorPositions.Push({ x: Round(actualCenterX), y: Round(actualCenterY) })
        } else {
            ; Fallback to calculated position if GetWindowRect fails
            g_SquareSelectorPositions.Push({ x: guiInfo.calculatedCenter.x, y: guiInfo.calculatedCenter.y })
        }
    }

    ; Activate letter selection mode
    g_SquareSelectorActive := true
    g_SquareSelectorClickMode := false  ; Reset click mode when showing new squares
    SetupLetterKeyListener()

    ; Enable CTRL hotkey to toggle click mode
    Hotkey("Ctrl", (*) => HandleCtrlToggle(), "On")

    ; Enable arrow keys for immediate direction switching
    EnableDirectionSwitchHotkeys()

    ; Enable Escape key to cancel squares (works in initial mode)
    Hotkey("Escape", (*) => CancelSquareSelector(), "On")

    ; Record start time for guaranteed cleanup
    global g_SquareSelectorStartTime
    g_SquareSelectorStartTime := A_TickCount

    ; Set timer to cleanup after 7 seconds if nothing is pressed
    ; Create cleanup function bound to this session ID (prevents old timers from cleaning up new squares)
    currentSessionID := g_SquareSelectorSessionID
    g_SquareSelectorTimer := CreateTimerHandler(currentSessionID)
    SetTimer(g_SquareSelectorTimer, -7000)  ; 7 second timeout

    ; Set up backup cleanup timer that checks every 2 seconds (guaranteed cleanup after 7 seconds)
    global g_SquareSelectorBackupTimer
    if (g_SquareSelectorBackupTimer) {
        SetTimer(g_SquareSelectorBackupTimer, 0)  ; Cancel old backup timer
    }
    g_SquareSelectorBackupTimer := () => BackupCleanupTimer()
    SetTimer(g_SquareSelectorBackupTimer, 2000)  ; Check every 2 seconds

    ; Lock will be released when timer fires, user selects a letter (enters loop mode), or presses Escape
}

; Function to show 4 direction indicator squares around the mouse pointer
ShowDirectionIndicators() {
    global g_DirectionIndicatorGuis

    ; Clean up any existing direction indicators
    CleanupDirectionIndicators()

    ; Get current mouse position
    pos := GetMousePos()
    mouseX := pos.x
    mouseY := pos.y

    ; Configuration
    squareSize := 24  ; Reduced for more precision
    offset := 35  ; Reduced for more precision

    ; Arrow symbols for each direction
    arrowUp := "↑"
    arrowRight := "→"
    arrowDown := "↓"
    arrowLeft := "←"
    arrows := [arrowUp, arrowRight, arrowDown, arrowLeft]

    ; Positions relative to mouse: Up, Right, Down, Left
    positions := []
    positions.Push({ x: mouseX, y: mouseY - offset })           ; Up
    positions.Push({ x: mouseX + offset, y: mouseY })           ; Right
    positions.Push({ x: mouseX, y: mouseY + offset })           ; Down
    positions.Push({ x: mouseX - offset, y: mouseY })            ; Left

    ; Create all 4 indicator squares
    guiArray := []
    for i, arrow in arrows {
        pos := positions[i]

        ; Create square GUI with arrow
        indicatorGui := Gui("+AlwaysOnTop -Caption +ToolWindow")
        indicatorGui.BackColor := "FF0000"  ; Red
        indicatorGui.SetFont("s10 Bold cFFFFFF", "Segoe UI")  ; White text, bold, smaller for precision

        ; Set GUI margins to 0
        indicatorGui.MarginX := 0
        indicatorGui.MarginY := 0

        ; Create text control that perfectly centers the arrow
        arrowText := indicatorGui.AddText("w" . squareSize . " h" . squareSize . " Center 0x201", arrow)

        ; Calculate top-left position for this square
        guiX := Round(pos.x - squareSize / 2.0)
        guiY := Round(pos.y - squareSize / 2.0)

        ; Store GUI and position info
        guiArray.Push({ gui: indicatorGui, x: guiX, y: guiY })
    }

    ; Position all GUIs while hidden, then show simultaneously
    for guiInfo in guiArray {
        guiInfo.gui.Show("x" . guiInfo.x . " y" . guiInfo.y . " w" . squareSize . " h" . squareSize . " NA Hide")
        ; Set less opaque (same as letter squares)
        WinSetTransparent(80, guiInfo.gui)
    }

    ; Show all GUIs simultaneously
    for guiInfo in guiArray {
        try {
            guiInfo.gui.Show("NA")
        } catch {
            try {
                guiInfo.gui.Show("x" . guiInfo.x . " y" . guiInfo.y . " w" . squareSize . " h" . squareSize . " NA")
            }
        }
        g_DirectionIndicatorGuis.Push(guiInfo.gui)
    }
}

; Helper function to destroy GUI objects in an array (reusable)
