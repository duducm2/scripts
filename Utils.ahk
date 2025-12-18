#Requires AutoHotkey v2.0+
#SingleInstance Force

; -----------------------------------------------------------------------------
; This script consolidates various utility hotkeys.
; -----------------------------------------------------------------------------

; =============================================================================
; Hotstring Core Functions
; =============================================================================

; ----------------------
; Safer hotstrings core
; ----------------------
global g_hotstrings := []
global g_lastExpansion := 0

; Trigger only on Space or Tab, not Enter or punctuation
Hotstring("EndChars", " `t")

RegisterHotstring(trigger, expansion, category := "", title := "") {
    global g_hotstrings
    g_hotstrings.Push({ trigger: trigger, expansion: expansion, category: category, title: title })
}

GetHotstringsCheatSheetText() {
    global g_hotstrings
    if (!IsSet(g_hotstrings) || g_hotstrings.Length = 0)
        return ""
    txt := ""
    for hs in g_hotstrings {
        line := "[" hs.trigger "] > " hs.expansion
        if (txt = "")
            txt := line
        else
            txt := txt . "`n" . line
    }
    return txt
}

; Safe paste insertion to avoid app shortcuts and re-triggers
InsertText(text) {
    global g_lastExpansion
    ; Debounce to prevent rapid duplicate expansions (e.g., double Space)
    if (A_TickCount - g_lastExpansion) < 250
        return
    g_lastExpansion := A_TickCount

    saved := ClipboardAll()
    try {
        A_Clipboard := text
        ClipWait(0.3)
        Sleep 50  ; Give time for clipboard to fully update
        Send "^v"
    } finally {
        Sleep 150  ; Wait longer for paste to complete before restoring clipboard
        A_Clipboard := saved
    }
}

; ----------------------
; Define hotstrings below
; (same triggers, now using InsertText)
; ----------------------

:o:myl::
{
    InsertText("My Links")
}

:o:gintegra::
{
    InsertText("GS_UX core team_UX and CIP Integration")
}

:o:gdash::
{
    InsertText("GS_E&S_CIP Dashboard research and design")
}

:o:gb2c::
{
    InsertText("GS_B2C_Credit_Management_Strategy_UI_Mentoring")
}

:o:gug::
{
    InsertText("GS_UX Core Team_Monitoring for B2C in Brazil")
}

:o:gpm::
{
    InsertText("GS_UX_Project_Management_Activities_LA")
}

:o:guxcip::
{
    InsertText("GS_UX_and_CIP")
}

:o:gtrain::
{
    InsertText("GS_UX core team_Trainings Management")
}

:o:gbp::
{
    InsertText("GS_B2C_Portals and Key Accounts Process POC")
}

:o:cgrammar::
{
    InsertText(
        "Correct grammar and spelling. Remove any dashes from the text. The text should be plain with no styles. Give back only the text. Use lininebreaks and a space betetween paragraphs and look like a human."
    )
}

:o:ebosch::
{
    InsertText("eduardo.figueiredo@br.bosch.com")
}

:o:egoogle::
{
    InsertText("edu.evangelista.figueiredo@gmail.com")
}

:o:mtask::
{
    InsertText(
        "This is a message, summary, text or any textual information that translates into a task for me to do. Translate this way, into a task, make informative and start with the emoji 🔲. Make it clear and consise."
    )
}

:o:flog::
{
    InsertText(
        "( LTrim`nFood_Log dictation → Excel CSV`n`nROLE`nYou transcribe my meal dictation (PT/EN) into rows for my Excel Food_Log.`n`nHOW IT WORKS`n- I will dictate one or more meals in free speech.`n- Process immediately without asking questions.`n`nOUTPUT FORMAT (STRICT)`n- Output ONLY the final CSV block. NO bullet points, NO analysis, NO extra text.`n- Plain text CSV. NO markdown, NO header, NO commentary.`n- Output a single block of text. NO empty lines between rows.`n- Separator: semicolon (;)`n`nEXAMPLE OUTPUT`n2025-10-25;Breakfast;08:30;coffee;caffeine;0;`n2025-10-25;Lunch;12:15;rice, beans;protein;0;`n`nOUTPUT RULES`n- STRICTLY CONTINUOUS LINES. Do not insert empty lines between rows.`n- Do not group by meal. List all sequentially in one block.`n- Use single newlines (\n) only. No paragraph breaks.`n- Trim whitespace from each row.`n- Each row format: Date;Meal;Time;Main_Items;Tags;Satisfaction_with_Speech;Notes`n- Sort by Date then Time.`n`nFIELD RULES`n- Date: YYYY‑MM‑DD. Use " "today" " for current date in America/Sao_Paulo timezone.`n  (IMPORTANT: For the date always consider one day before the current one, unless specified otherwise.)`n- Time: HH:MM in 24h; pad leading zeros (e.g., 08:05).`n- Meal: Breakfast | Lunch | Dinner | Snack.`n  PT mapping: café da manhã→Breakfast; almoço→Lunch; jantar/janta→Dinner; lanche→Snack.`n- Main_Items: comma‑separated simple item names (e.g., coffee, bread, butter).`n- Tags: comma‑separated, from this set when present or inferable:`n  caffeine, sugar, alcohol, dairy, gluten, fried, spicy, high-carb, low-carb, processed, protein, fiber, late-night, home-cooked, fast-food.`n  Add " "late-night" " automatically if Time ≥ 22:00.`n- Satisfaction_with_Speech: integer 0–3 (0 = liked a lot; 3 = disliked a lot). If not stated, leave empty.`n- Notes: short free text when I provide context.`n`nMISSING INFO`n- If Date or Time is missing, use " "today" " and infer time from meal type (Breakfast=08:00, Lunch=12:00, Dinner=19:00, Snack=15:00).`n`nACK`n- Process the dictation immediately and output CSV rows only.)"
    )
}

:o:aiopt::
{
    InsertText(
        "( LTrim`nTask: Rewrite the input text so it becomes AI-oriented.`n`nGoal: Produce a version that is concise, unambiguous, free of redundancy, and easy for an AI to parse.`n`nContext: The input text contains instructions and requests intended for a second AI (AIB). You must preserve ALL important information, especially any instructions, requests, or requirements meant for AIB. Do not remove information in the name of clarity or conciseness.`n`nInstructions:`n`n1. Preserve all essential information, especially instructions and requests for the second AI.`n2. Use positive, direct instructions.`n3. Maintain consistent terminology and simple syntax.`n4. Resolve ambiguity and clarify references.`n5. Output in a clean, structured format with no extra commentary.`n6. Do NOT omit any important information, requests, or instructions that the user provided for the second AI.`n7. Do NOT bias the prompt with your own concerns, interpretations, or modifications. Process the text as-is without adding your own perspective or concerns.`n`nInput: <insert text here>`n`nOutput:`nCRITICAL: The output must contain ONLY the processed result with no additional text, commentary, explanations, or formatting. The output must be ready for direct copy and paste without any modification.`n`nA rewritten version of the input text that is optimized for AI interpretation and contains:`n`n* Clear meaning`n* No repeated ideas`n* No filler wording`n* No contradictions`n* Stable terminology`n* Straightforward sentence structure`n* ALL important information preserved, especially instructions for the second AI)"
    )
}

; ----------------------
; Register hotstrings for cheat sheet display
; ----------------------
InitHotstringsCheatSheet() {
    ; Prompts (4 items) - First category
    RegisterHotstring(":o:cgrammar",
        "Correct grammar and spelling. Remove any dashes from the text. The text should be plain with no styles. Give back only the text. Use linebreaks and a space between paragraphs and look like a human.",
        "Prompts",
        "✏️ Grammar & Spelling Corrector"
    )
    RegisterHotstring(":o:mtask",
        "This is a message, summary, text or any textual information that translates into a task for me to do. Translate this way, into a task, make informative and start with the emoji 🔲. Make it clear and consise.",
        "Prompts",
        "🔲 Convert to Task"
    )
    RegisterHotstring(":o:flog",
        "( LTrim`nFood_Log dictation → Excel CSV`n`nROLE`nYou transcribe my meal dictation (PT/EN) into rows for my Excel Food_Log.`n`nHOW IT WORKS`n- I will dictate one or more meals in free speech.`n- Process immediately without asking questions.`n`nOUTPUT FORMAT (STRICT)`n- Output ONLY the final CSV block. NO bullet points, NO analysis, NO extra text.`n- Plain text CSV. NO markdown, NO header, NO commentary.`n- Output a single block of text. NO empty lines between rows.`n- Separator: semicolon (;)`n`nEXAMPLE OUTPUT`n2025-10-25;Breakfast;08:30;coffee;caffeine;0;`n2025-10-25;Lunch;12:15;rice, beans;protein;0;`n`nOUTPUT RULES`n- STRICTLY CONTINUOUS LINES. Do not insert empty lines between rows.`n- Do not group by meal. List all sequentially in one block.`n- Use single newlines (\n) only. No paragraph breaks.`n- Trim whitespace from each row.`n- Each row format: Date;Meal;Time;Main_Items;Tags;Satisfaction_with_Speech;Notes`n- Sort by Date then Time.`n`nFIELD RULES`n- Date: YYYY‑MM‑DD. Use " "today" " for current date in America/Sao_Paulo timezone.`n  (IMPORTANT: For the date always consider one day before the current one, unless specified otherwise.)`n- Time: HH:MM in 24h; pad leading zeros (e.g., 08:05).`n- Meal: Breakfast | Lunch | Dinner | Snack.`n  PT mapping: café da manhã→Breakfast; almoço→Lunch; jantar/janta→Dinner; lanche→Snack.`n- Main_Items: comma‑separated simple item names (e.g., coffee, bread, butter).`n- Tags: comma‑separated, from this set when present or inferable:`n  caffeine, sugar, alcohol, dairy, gluten, fried, spicy, high-carb, low-carb, processed, protein, fiber, late-night, home-cooked, fast-food.`n  Add " "late-night" " automatically if Time ≥ 22:00.`n- Satisfaction_with_Speech: integer 0–3 (0 = liked a lot; 3 = disliked a lot). If not stated, leave empty.`n- Notes: short free text when I provide context.`n`nMISSING INFO`n- If Date or Time is missing, use " "today" " and infer time from meal type (Breakfast=08:00, Lunch=12:00, Dinner=19:00, Snack=15:00).`n`nACK`n- Process the dictation immediately and output CSV rows only.)",
        "Prompts",
        "🍽️ Food Log Dictation"
    )
    RegisterHotstring(":o:aiopt",
        "( LTrim`nTask: Rewrite the input text so it becomes AI-oriented.`n`nGoal: Produce a version that is concise, unambiguous, free of redundancy, and easy for an AI to parse.`n`nContext: The input text contains instructions and requests intended for a second AI (AIB). You must preserve ALL important information, especially any instructions, requests, or requirements meant for AIB. Do not remove information in the name of clarity or conciseness.`n`nInstructions:`n`n1. Preserve all essential information, especially instructions and requests for the second AI.`n2. Use positive, direct instructions.`n3. Maintain consistent terminology and simple syntax.`n4. Resolve ambiguity and clarify references.`n5. Output in a clean, structured format with no extra commentary.`n6. Do NOT omit any important information, requests, or instructions that the user provided for the second AI.`n7. Do NOT bias the prompt with your own concerns, interpretations, or modifications. Process the text as-is without adding your own perspective or concerns.`n`nInput: <insert text here>`n`nOutput:`nCRITICAL: The output must contain ONLY the processed result with no additional text, commentary, explanations, or formatting. The output must be ready for direct copy and paste without any modification.`n`nA rewritten version of the input text that is optimized for AI interpretation and contains:`n`n* Clear meaning`n* No repeated ideas`n* No filler wording`n* No contradictions`n* Stable terminology`n* Straightforward sentence structure`n* ALL important information preserved, especially instructions for the second AI)",
        "Prompts",
        "🤖 AI Text Optimizer"
    )

    ; Placeholder prompts (7 slots reserved for future prompts)
    RegisterHotstring("", "", "Prompts", "Reserved 1")
    RegisterHotstring("", "", "Prompts", "Reserved 2")
    RegisterHotstring("", "", "Prompts", "Reserved 3")
    RegisterHotstring("", "", "Prompts", "Reserved 4")
    RegisterHotstring("", "", "Prompts", "Reserved 5")
    RegisterHotstring("", "", "Prompts", "Reserved 6")
    RegisterHotstring("", "", "Prompts", "Reserved 7")

    ; General Information (2 items) - Second category
    RegisterHotstring(":o:ebosch", "eduardo.figueiredo@br.bosch.com", "General", "💼 Bosch Email")
    RegisterHotstring(":o:egoogle", "edu.evangelista.figueiredo@gmail.com", "General", "📧 Gmail")

    ; Project Names (Projects category)
    ; Order here defines character assignment within the Projects segment of g_HotstringCharSequence.
    ; First three keep their current slots, then we introduce "14-my-notes" at the former X slot,
    ; and drop the previous X/C/7 entries so later projects shift left by one as needed.
    RegisterHotstring(":o:myl", "My Links", "Projects", "🔗 My Links")
    RegisterHotstring(":o:gintegra", "GS_UX core team_UX and CIP Integration", "Projects", "🔄 UX and CIP Integration")
    RegisterHotstring(":o:gdash", "GS_E&S_CIP Dashboard research and design", "Projects", "📊 CIP Dashboard")
    ; New project for the X slot in the selector
    RegisterHotstring(":o:14notes", "14-my-notes", "Projects", "📝 14-my-notes")
    ; Remaining projects (previously after X/C/7) shift left one character each
    RegisterHotstring(":o:gpm", "GS_UX_Project_Management_Activities_LA", "Projects", "📋 Project Management LA")
    RegisterHotstring(":o:guxcip", "GS_UX_and_CIP", "Projects", "🔗 UX and CIP")
    RegisterHotstring(":o:gtrain", "GS_UX core team_Trainings Management", "Projects", "🎓 Trainings Management")
    RegisterHotstring(":o:26ai", "26-ai-experiment", "Projects", "🤖 26-ai-experiment")
}
InitHotstringsCheatSheet()

; ------------
; Optional: scope Explorer-only hotstrings used for renaming
; Uncomment to restrict selected triggers to File Explorer or Save dialogs
;------------
;#HotIf WinActive("ahk_exe explorer.exe") || WinActive("ahk_class #32770")
;:o:gdash::
;    InsertText("GS_E&S_CIP Dashboard research and design")
;return
;#HotIf

; --- Hotkeys & Functions -----------------------------------------------------

; Ensure per-monitor DPI awareness so coordinates are physical pixels across mixed scaling
InitDpiAwareness() {
    static PER_MONITOR_AWARE_V2 := -4 ; DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2
    try DllCall("SetProcessDpiAwarenessContext", "ptr", PER_MONITOR_AWARE_V2, "ptr")
}
InitDpiAwareness()

; =============================================================================
; Jump Mouse to Middle of Active Window
; Hotkey: Win+Alt+Shift+3
; Original File: Jump mouse on the middle.ahk
; =============================================================================
#!+Q::
{
    hwnd := WinExist("A")
    if !hwnd {
        return ; silently abort if no active window
    }
    rect := Buffer(16, 0)
    if !DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect) {
        MsgBox "GetWindowRect failed"
        return
    }
    left := NumGet(rect, 0, "int")
    top := NumGet(rect, 4, "int")
    right := NumGet(rect, 8, "int")
    bottom := NumGet(rect, 12, "int")
    centerX := left + (right - left) // 2
    centerY := top + (bottom - top) // 2
    DllCall("SetCursorPos", "int", centerX, "int", centerY)
}

; =============================================================================
; Activate Cursor and Send Key Sequence with Options
; Hotkey: Win+Alt+Shift+C
; =============================================================================

; Global variable to remember target window for Cursor action
global gCursorActionTargetWin := 0

; Auto-submit handler for Cursor action modal
AutoSubmitCursorAction(ctrl, *) {
    currentValue := ctrl.Text
    if (currentValue != "" && IsInteger(currentValue)) {
        choice := Integer(currentValue)
        if (choice >= 1 && choice <= 3) {
            ctrl.Gui.Destroy()
            ExecuteCursorAction(choice)
        }
    }
}

; Manual submit handler (OK button)
SubmitCursorAction(ctrl, *) {
    currentValue := ctrl.Gui["CursorActionInput"].Text
    if (currentValue != "" && IsInteger(currentValue)) {
        choice := Integer(currentValue)
        if (choice >= 1 && choice <= 3) {
            ctrl.Gui.Destroy()
            ExecuteCursorAction(choice)
        } else {
            MsgBox "Invalid selection. Please choose 1-3.", "Cursor Action Selection", "IconX"
        }
    }
}

; Cancel handler
CancelCursorAction(ctrl, *) {
    ctrl.Gui.Destroy()
}

; Execute the Cursor key sequence based on numeric choice
ExecuteCursorAction(choice) {
    try {
        ; First activate Cursor
        SetTitleMatchMode 2
        if WinExist("ahk_exe Cursor.exe") {
            WinActivate
            WinWaitActive("ahk_exe Cursor.exe", "", 2)
        } else {
            target := IS_WORK_ENVIRONMENT ? "C:\\Users\\fie7ca\\AppData\\Local\\Programs\\cursor\\Cursor.exe" :
                "C:\\Users\\eduev\\AppData\\Local\\Programs\\cursor\\Cursor.exe"
            Run target
            WinWaitActive("ahk_exe Cursor.exe", "", 10)
        }

        Sleep 200

        switch choice {
            case 1:
            {
                Send "^+e"
                Sleep 100
                Send "^i"
                Sleep 100
                Send "+{Backspace}"
            }
            case 2:
            {
                Send "^+e"
                Sleep 100
                Send "^i"
                Sleep 100
                Send "{Enter}"
            }
            case 3:
            {
                Send "^+e"
                Sleep 100
                Send "^i"
                Sleep 100
                Send "{Up}"
                Sleep 100
                Send "{Enter}"
            }
        }
    } catch Error as e {
        MsgBox "Error executing Cursor action: " e.Message, "Cursor Action Error", "IconX"
    }
}

; Function to show auto-submit modal and then run Cursor action
CursorKeySequence() {
    try {
        gCursorActionTargetWin := WinExist("A")

        cursorGui := Gui("+AlwaysOnTop +ToolWindow", "Cursor Action Selection")
        cursorGui.SetFont("s10", "Segoe UI")

        cursorGui.AddText("w360 Center",
            "Choose Action:`n`n1. Shift backspace`n2. Enter`n3. Allow`n`nType a number (1-3):")
        cursorGui.AddEdit("w50 Center vCursorActionInput Limit1 Number")
        cursorGui.AddButton("w80 xp-40 y+10", "OK").OnEvent("Click", SubmitCursorAction)
        cursorGui.AddButton("w80 xp+90", "Cancel").OnEvent("Click", CancelCursorAction)
        cursorGui["CursorActionInput"].OnEvent("Change", AutoSubmitCursorAction)

        cursorGui.Show("w360 h200")
        cursorGui["CursorActionInput"].Focus()
    } catch Error as e {
        MsgBox "Error opening Cursor action selector: " e.Message, "Cursor Action Error", "IconX"
    }
}

#!+C::
{
    CursorKeySequence()
}

; =============================================================================
; Mouse Jump Shortcuts
; Hotkeys: Win+Alt+Shift+Arrow Keys
; Jump mouse cursor by fixed pixel distance in each direction with multi-monitor support
; =============================================================================

; Set coordinate mode to screen for proper multi-monitor support
CoordMode "Mouse", "Screen"

; Define the movement distance in pixels (increased from 200 to 300)
global MOUSE_JUMP_DISTANCE := 300

; Helper function to get current mouse position using physical screen coordinates
GetMousePos() {
    pt := Buffer(8, 0)
    DllCall("GetCursorPos", "ptr", pt)
    return { x: NumGet(pt, 0, "int"), y: NumGet(pt, 4, "int") }
}

; Helper function to get all monitor information
GetMonitorInfo() {
    ; Get the number of monitors
    monitorCount := SysGet(80)  ; SM_CMONITORS
    monitors := []

    loop monitorCount {
        ; Get work area for each monitor (excludes taskbar)
        MonitorGetWorkArea(A_Index, &left, &top, &right, &bottom)
        monitors.Push({
            left: left,
            top: top,
            right: right,
            bottom: bottom,
            width: right - left,
            height: bottom - top
        })
    }

    return monitors
}

; Helper function: get virtual desktop bounds (supports negative X/Y)
GetVirtualBounds() {
    left := DllCall("GetSystemMetrics", "int", 76)    ; SM_XVIRTUALSCREEN
    top := DllCall("GetSystemMetrics", "int", 77)     ; SM_YVIRTUALSCREEN
    width := DllCall("GetSystemMetrics", "int", 78)   ; SM_CXVIRTUALSCREEN
    height := DllCall("GetSystemMetrics", "int", 79)  ; SM_CYVIRTUALSCREEN
    return { left: left, top: top, right: left + width - 1, bottom: top + height - 1 }
}

Clamp(n, lo, hi) {
    return n < lo ? lo : n > hi ? hi : n
}

; Helper function to find which monitor contains the given coordinates
FindMonitorForCoords(x, y, monitors) {
    for monitor in monitors {
        if (x >= monitor.left && x <= monitor.right && y >= monitor.top && y <= monitor.bottom) {
            return monitor
        }
    }
    return false  ; Not found in any monitor
}

; Helper function to safely move mouse with proper multi-monitor boundary checking
; Always shows both prediction squares (blue for short, red for long) in the direction of movement
SafeMouseMove(deltaX, deltaY) {
    pos := GetMousePos()
    v := GetVirtualBounds()
    ; Calculate target position where mouse will jump to (current + jump distance)
    targetX := Clamp(pos.x + deltaX, v.left, v.right)
    targetY := Clamp(pos.y + deltaY, v.top, v.bottom)

    ; Move the mouse to the target position first
    DllCall("SetCursorPos", "int", targetX, "int", targetY)

    ; After moving, show both prediction squares in the direction of movement
    ; Blue square: shows where mouse will land with short jump (without Control)
    ; Red square: shows where mouse will land with long jump (with Control)
    ShowBothPredictionSquares(targetX, targetY, deltaX, deltaY)
}

; Global array to track all feedback GUI windows
global g_MouseMoveFeedbackGuis := []

; Helper function to close all feedback GUIs
CloseMouseMoveFeedback() {
    global g_MouseMoveFeedbackGuis
    try {
        for gui in g_MouseMoveFeedbackGuis {
            try {
                if (IsObject(gui) && gui.Hwnd) {
                    gui.Destroy()
                }
            } catch {
                ; Silently ignore errors for individual GUIs
            }
        }
        g_MouseMoveFeedbackGuis := []
    } catch {
        ; Silently ignore errors during cleanup
    }
}

; Helper function to show both prediction squares (blue and red) in the direction of movement
; Shows where the mouse will land if user presses short (blue) or long (red) jump in the same direction
ShowBothPredictionSquares(currentX, currentY, deltaX, deltaY) {
    global g_MouseMoveFeedbackGuis
    global MOUSE_JUMP_DISTANCE
    v := GetVirtualBounds()

    ; Close any existing feedback GUIs first
    CloseMouseMoveFeedback()

    ; Determine the direction of movement from the sign of deltaX/deltaY
    ; The squares always use the base MOUSE_JUMP_DISTANCE, regardless of current jump distance

    if (deltaX != 0) {
        ; Horizontal movement - determine direction from sign of deltaX
        directionX := deltaX > 0 ? 1 : -1  ; 1 for right, -1 for left

        ; Blue square: shows where mouse will land with short jump (MOUSE_JUMP_DISTANCE in this direction)
        shortPredictionX := Clamp(currentX + MOUSE_JUMP_DISTANCE * directionX, v.left, v.right)
        ; Red square: shows where mouse will land with long jump (MOUSE_JUMP_DISTANCE * 2 in this direction)
        longPredictionX := Clamp(currentX + MOUSE_JUMP_DISTANCE * 2 * directionX, v.left, v.right)

        ; Show blue square (short distance) in the direction of movement
        ShowPredictionSquare(shortPredictionX, currentY, "0000FF")
        ; Show red square (long distance) in the direction of movement
        ShowPredictionSquare(longPredictionX, currentY, "FF0000")
    } else if (deltaY != 0) {
        ; Vertical movement - determine direction from sign of deltaY
        directionY := deltaY > 0 ? 1 : -1  ; 1 for down, -1 for up

        ; Blue square: shows where mouse will land with short jump (MOUSE_JUMP_DISTANCE in this direction)
        shortPredictionY := Clamp(currentY + MOUSE_JUMP_DISTANCE * directionY, v.top, v.bottom)
        ; Red square: shows where mouse will land with long jump (MOUSE_JUMP_DISTANCE * 2 in this direction)
        longPredictionY := Clamp(currentY + MOUSE_JUMP_DISTANCE * 2 * directionY, v.top, v.bottom)

        ; Show blue square (short distance) in the direction of movement
        ShowPredictionSquare(currentX, shortPredictionY, "0000FF")
        ; Show red square (long distance) in the direction of movement
        ShowPredictionSquare(currentX, longPredictionY, "FF0000")
    }

    ; Auto-hide after 1300ms (1 second longer than before)
    SetTimer(CloseMouseMoveFeedback, -1300)
}

; Helper function to show a single prediction square
ShowPredictionSquare(x, y, color) {
    global g_MouseMoveFeedbackGuis
    squareSize := 40

    ; Create a simple GUI window with specified color background
    squareGui := Gui("+AlwaysOnTop -Caption +ToolWindow")
    squareGui.BackColor := color

    ; Position the square centered at the target position
    guiX := x - (squareSize // 2)
    guiY := y - (squareSize // 2)

    ; Show the square
    squareGui.Show("x" . guiX . " y" . guiY . " w" . squareSize . " h" . squareSize . " NA")
    WinSetTransparent(100, squareGui)  ; Less opaque for better visibility

    ; Store reference for cleanup
    g_MouseMoveFeedbackGuis.Push(squareGui)
}

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
                ; Ignore if activation fails
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

; Jump mouse right (short distance) - now shows square selector
#!+Right::
{
    HandleDirectionHotkey("Right")
    return
}

; Jump mouse left (short distance) - now shows square selector
#!+Left::
{
    HandleDirectionHotkey("Left")
}

; Jump mouse down (short distance) - now shows square selector
#!+Down::
{
    HandleDirectionHotkey("Down")
}

; Jump mouse up (short distance) - now shows square selector
#!+Up::
{
    HandleDirectionHotkey("Up")
}

; =============================================================================
; Activate Hunt and Peck
; Hotkey: Win+Alt+Shift+X
; Original File: Hunt and Peck.ahk
; Note: A short press activates Hunt and Peck. A long press (>1s) will activate
;       loop mode, which will continuously reactivate Hunt and Peck after each
;       selection until either long pressed again or max iterations reached.
; =============================================================================

; Global variables for Hunt and Peck loop mode
; Path to Hunt and Peck executable – adjust if you install it elsewhere
#Include %A_ScriptDir%\env.ahk
global g_HnPExePath := GetHnPExePath()
global g_HnPLoopActive := false
global g_HnPLoopGui := false
global g_HnPMaxIterations := 25
global g_HnPCurrentIteration := 0
global g_HnPTargetWindow := 0  ; Store the window handle
global g_HnPRetryCount := 0    ; Track retry attempts for Program Manager recovery

; -----------------------------------------------------------------------------
; Helper: Force-close any running Hunt-and-Peck (hap.exe) processes
; -----------------------------------------------------------------------------
CloseHuntAndPeckProcess() {
    ; Attempt to terminate every instance of hap.exe. Ignoring any errors keeps
    ; the call simple and side-effect-free if the process isn’t running.
    try ProcessClose("hap.exe")
}

; Helper: returns true if active window is Program Manager (desktop) or taskbar – i.e. Hunt-and-Peck anchored incorrectly
IsBadHnPAnchor() {
    return WinActive("Program Manager") || WinActive("ahk_class Shell_TrayWnd")
}

; Safely activate the stored target window if it still exists
SafeActivateTarget(hwnd) {
    if (hwnd && WinExist("ahk_id " hwnd))
        WinActivate("ahk_id " hwnd)
}

; Send a quick right-click to the centred mouse position – this shifts focus to the window's
; main area without selecting items.  Any context menu will be dismissed automatically by
; Hunt-and-Peck's overlay / Esc logic.
RightClickFocus() {
    Click "Right"
    Sleep 10  ; Brief pause for focus shift
}

; Shows or hides the loop mode indicator
ShowLoopIndicator(show := true) {
    global g_HnPLoopGui

    if (show && !g_HnPLoopGui) {
        g_HnPLoopGui := Gui()
        g_HnPLoopGui.Opt("+AlwaysOnTop -Caption +ToolWindow")
        g_HnPLoopGui.BackColor := "33AA33"
        g_HnPLoopGui.Add("Text", , "Hunt && Peck Loop Mode Active")
        g_HnPLoopGui.Show("NoActivate y0")
    } else if (!show && g_HnPLoopGui) {
        g_HnPLoopGui.Destroy()
        g_HnPLoopGui := false
    }
}

; Activates Hunt and Peck and handles potential Program Manager activation
ActivateHuntAndPeck(isLoopMode := false) {
    global g_HnPTargetWindow, g_HnPRetryCount, g_HnPExePath

    ; NOTE: Do NOT force-kill hap.exe here – it interferes with subsequent launches.

    ; For single activation (not loop mode), always use current window
    if (!isLoopMode) {
        g_HnPTargetWindow := WinExist("A")
    }

    ; If we don't have a valid window, we can't proceed
    if (!g_HnPTargetWindow || !WinExist("ahk_id " g_HnPTargetWindow)) {
        return false
    }

    ; Ensure our target window is active
    SafeActivateTarget(g_HnPTargetWindow)
    Sleep 40

    ; Shift keyboard focus with a harmless right-click
    RightClickFocus()

    ; Wait up to 200 ms for the window to actually become active (covers fast Alt-Tab cases)
    if !WinWaitActive("ahk_id " g_HnPTargetWindow, "", 0.2) {
        ; If it still isn't active, give up on this attempt
        return false
    }

    ; ----------------------------------------------------------------------
    ; Prefer the far more reliable CLI (hap.exe /hint). After running, ensure
    ; we did NOT end up focused on Program Manager **or** the taskbar.
    ; ----------------------------------------------------------------------
    boolSuccess := false
    if (FileExist(g_HnPExePath)) {
        try {
            ; Focus already fixed; just launch
            Run g_HnPExePath " /hint", , "Hide"
            Sleep 120
            boolSuccess := !IsBadHnPAnchor()
        }
    }

    ; If the CLI call failed (or exe not found) fall back to the legacy hotkey
    if (!boolSuccess) {
        ; Legacy hotkey path
        RightClickFocus()
        Send "!ç"
        Sleep 40
        ; Re-activate target window to pull overlay back
        SafeActivateTarget(g_HnPTargetWindow)
        Sleep 40

        ; Still anchored to Program Manager / taskbar? => one retry only.
        if (IsBadHnPAnchor()) {
            if (g_HnPRetryCount < 1) {
                g_HnPRetryCount++

                SafeActivateTarget(g_HnPTargetWindow)
                Sleep 40
                RightClickFocus()
                Sleep 80

                if (IsBadHnPAnchor()) {
                    g_HnPRetryCount := 0
                    return false
                }
            } else {
                g_HnPRetryCount := 0
                return false
            }
        }
    }

    ; Success – reset retry counter
    g_HnPRetryCount := 0
    return true
}

; Handles the Hunt and Peck loop mode
HnPLoopMode() {
    global g_HnPLoopActive, g_HnPCurrentIteration, g_HnPTargetWindow

    if (g_HnPLoopActive) {
        ; Stop the loop mode
        g_HnPLoopActive := false
        g_HnPCurrentIteration := 0
        ShowLoopIndicator(false)

        ; Stop the timer immediately
        SetTimer ActivateHnP, 0

        ; Brief delay to ensure any pending Hunt and Peck activation is complete
        Sleep 100

        ; Ensure we're in the target window and clear any Hunt and Peck state
        if (g_HnPTargetWindow && WinExist("ahk_id " g_HnPTargetWindow)) {
            SafeActivateTarget(g_HnPTargetWindow)
            Sleep 30
            Send "{Esc}"

            ; Safeguard: after 1000 ms send Esc to dismiss any late overlay.
            SetTimer(() => Send("{Esc}"), -1000)

            ; Ensure the Hunt-and-Peck process is fully terminated
            CloseHuntAndPeckProcess()
        }

        ; Enhanced cleanup: Check for remaining HAP.EXE instances multiple times within 2 seconds
        ; and close them if found. This ensures all instances are properly terminated.
        SetTimer(() => CloseHuntAndPeckProcess(), -500)   ; First check at 500ms
        SetTimer(() => CloseHuntAndPeckProcess(), -1000)  ; Second check at 1000ms
        SetTimer(() => CloseHuntAndPeckProcess(), -1500)  ; Third check at 1500ms

        g_HnPTargetWindow := 0
        return
    }

    ; Store the current active window
    g_HnPTargetWindow := WinExist("A")
    if (!g_HnPTargetWindow) {
        return  ; Don't show error, just fail silently
    }

    ; Start the loop mode
    g_HnPLoopActive := true
    g_HnPCurrentIteration := 0
    ShowLoopIndicator(true)

    ; Activate Hunt and Peck immediately for the first time
    ActivateHuntAndPeck()

    ; Define customizable interval (ms) for Hunt and Peck loop
    static loopIntervalMs := 3000   ; was 2000 – extended for more selection time

    ; Start the loop with the new interval
    SetTimer ActivateHnP, loopIntervalMs
}

; Timer function to activate Hunt and Peck
ActivateHnP() {
    global g_HnPLoopActive, g_HnPCurrentIteration, g_HnPMaxIterations, g_HnPTargetWindow

    if (!g_HnPLoopActive || g_HnPCurrentIteration >= g_HnPMaxIterations) {
        SetTimer ActivateHnP, 0  ; Stop the timer
        g_HnPLoopActive := false
        g_HnPCurrentIteration := 0
        g_HnPTargetWindow := 0
        ShowLoopIndicator(false)

        ; Ensure any residual Hunt-and-Peck overlay is cleared
        Send "{Esc}"
        SetTimer(() => Send("{Esc}"), -1000)

        ; Terminate any lingering hap.exe process
        CloseHuntAndPeckProcess()

        ; Enhanced cleanup: Check for remaining HAP.EXE instances multiple times within 2 seconds
        ; and close them if found. This ensures all instances are properly terminated.
        SetTimer(() => CloseHuntAndPeckProcess(), -500)   ; First check at 500ms
        SetTimer(() => CloseHuntAndPeckProcess(), -1000)  ; Second check at 1000ms
        SetTimer(() => CloseHuntAndPeckProcess(), -1500)  ; Third check at 1500ms

        return
    }

    g_HnPCurrentIteration++
    if (!ActivateHuntAndPeck(true)) {  ; Pass true to indicate loop mode
        ; If Hunt and Peck activation failed after retry, stop the loop
        SetTimer ActivateHnP, 0
        g_HnPLoopActive := false
        g_HnPCurrentIteration := 0
        g_HnPTargetWindow := 0
        ShowLoopIndicator(false)

        ; Enhanced cleanup: Check for remaining HAP.EXE instances multiple times within 2 seconds
        ; and close them if found. This ensures all instances are properly terminated.
        SetTimer(() => CloseHuntAndPeckProcess(), -500)   ; First check at 500ms
        SetTimer(() => CloseHuntAndPeckProcess(), -1000)  ; Second check at 1000ms
        SetTimer(() => CloseHuntAndPeckProcess(), -1500)  ; Third check at 1500ms
    }
}

; -----------------------------------------------------------------------------
; Helper: Schedule multiple attempts to close Hunt-and-Peck after single activation
;           Provides redundancy in case the first attempt fails or HnP hangs.
; -----------------------------------------------------------------------------
ScheduleHnPCleanup() {
    global g_HnPLoopActive
    ; Execute CloseHuntAndPeckProcess() after progressive delays, but only if we’re not in loop mode
    static delays := [3000, 7000, 12000]  ; milliseconds
    for d in delays
        SetTimer(() => (!g_HnPLoopActive ? CloseHuntAndPeckProcess() : ""), -d)
}

#!+x::
{
    ; Treat press-and-hold >400 ms as loop-mode trigger (keeps parity with keyboard firmware)
    if KeyWait("x", "T0.4") {
        ; Released within 400 ms → single activation
        ActivateHuntAndPeck(false)  ; Pass false to indicate single activation
        ; Schedule redundant cleanup attempts to ensure hap.exe is terminated
        ScheduleHnPCleanup()
    }
    else {
        ; Key was held down ≥400 ms (long press)
        KeyWait("x")  ; Wait for the key to be released
        HnPLoopMode()
    }
}

; =============================================================================
; Hotstring Selector with Character Shortcuts
; Hotkey: Win+Alt+Shift+U
; Shows a GUI with all hotstrings and their assigned characters. Pressing
; a character immediately pastes the corresponding text.
; =============================================================================

; Global variables for hotstring selector
global g_HotstringSelectorGui := false
global g_HotstringSelectorActive := false
global g_HotstringCharMap := Map()  ; Maps character to expansion text
global g_HotstringHotkeyHandlers := []  ; Store hotkey handlers for cleanup

; Character sequence for assignment: 1 2 3 4 5 q w e r t a s d f g z x c v b 6 7 8 9 0 y u i o p h j k l n m , .
global g_HotstringCharSequence := ["1", "2", "3", "4", "5", "q", "w", "e", "r", "t", "a", "s", "d", "f", "g", "z", "x",
    "c", "v", "b", "6", "7", "8", "9", "0", "y", "u", "i", "o", "p", "h", "j", "k", "l", "n", "m", ",", "."]

; Category display order (Prompts first, General second, Projects last)
global g_HotstringCategories := ["Prompts", "General", "Projects"]

; Build character-to-expansion mapping grouped by category
BuildHotstringCharMap() {
    global g_hotstrings
    charMap := Map()

    if (!IsSet(g_hotstrings) || g_hotstrings.Length = 0) {
        return charMap
    }

    ; Group hotstrings by category
    categorized := Map()
    categorized["Projects"] := []
    categorized["Prompts"] := []
    categorized["General"] := []

    for hs in g_hotstrings {
        category := hs.category
        if (category = "Projects" || category = "Prompts" || category = "General") {
            categorized[category].Push(hs)
        }
    }

    ; Assign characters sequentially within each category
    charIndex := 1
    for category in g_HotstringCategories {
        for hs in categorized[category] {
            if (charIndex <= g_HotstringCharSequence.Length) {
                char := g_HotstringCharSequence[charIndex]
                ; Only assign characters to hotstrings that have an expansion (skip empty placeholders)
                if (hs.expansion != "") {
                    charMap[char] := hs.expansion
                }
                ; Always increment charIndex to reserve slots for placeholders
                charIndex++
            }
        }
    }

    return charMap
}

; Get categorized hotstrings for display
GetCategorizedHotstrings() {
    global g_hotstrings
    categorized := Map()
    categorized["Projects"] := []
    categorized["Prompts"] := []
    categorized["General"] := []

    if (!IsSet(g_hotstrings) || g_hotstrings.Length = 0) {
        return categorized
    }

    for hs in g_hotstrings {
        category := hs.category
        if (category = "Projects" || category = "Prompts" || category = "General") {
            categorized[category].Push(hs)
        }
    }

    return categorized
}

; Get preview text (truncate long text for display, replace newlines with spaces)
GetPreviewText(text, maxLength := 60) {
    ; Replace newlines and multiple spaces with single space for cleaner preview
    preview := RegExReplace(text, "`r?`n", " ")
    preview := RegExReplace(preview, "\s+", " ")
    preview := Trim(preview)

    if (StrLen(preview) <= maxLength) {
        return preview
    }
    return SubStr(preview, 1, maxLength) . "..."
}

; Cleanup hotstring selector
CleanupHotstringSelector() {
    global g_HotstringSelectorActive, g_HotstringSelectorGui, g_HotstringHotkeyHandlers
    global g_HotstringCharMap

    ; Disable active flag
    g_HotstringSelectorActive := false

    ; Disable all character hotkeys
    for handler in g_HotstringHotkeyHandlers {
        try {
            char := handler.char
            ; Handle special VK codes
            if (char = ",") {
                Hotkey("vkBC", "Off")
            } else if (char = ".") {
                Hotkey("vkBE", "Off")
            } else {
                Hotkey(char, "Off")
                ; Also disable uppercase for lowercase letters
                if (RegExMatch(char, "^[a-z]$")) {
                    Hotkey(StrUpper(char), "Off")
                }
            }
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
    g_HotstringHotkeyHandlers := []

    ; Close and destroy GUI
    if (IsObject(g_HotstringSelectorGui)) {
        try {
            g_HotstringSelectorGui.Destroy()
        } catch {
            ; Ignore
        }
        g_HotstringSelectorGui := false
    }

    ; Clear char map
    g_HotstringCharMap := Map()
}

; Handler for character key press
HandleHotstringChar(char) {
    global g_HotstringSelectorActive, g_HotstringCharMap

    ; Only process if selector is active
    if (!g_HotstringSelectorActive) {
        return
    }

    ; Get expansion text for this character
    expansion := g_HotstringCharMap.Get(char, "")
    if (expansion = "") {
        ; Try lowercase if uppercase
        expansion := g_HotstringCharMap.Get(StrLower(char), "")
    }

    if (expansion != "") {
        ; Cleanup first (closes GUI, disables hotkeys)
        CleanupHotstringSelector()

        ; Small delay to ensure target window has focus before pasting
        Sleep 150

        ; Paste the text
        InsertText(expansion)
    }
}

; Factory function to create a handler that properly captures the character
; This ensures each handler gets its own copy of the char value
CreateHotstringCharHandler(char) {
    ; Return a function that captures the char value at creation time
    return (*) => HandleHotstringChar(char)
}

; Handler for Escape key
HandleHotstringEscape(*) {
    global g_HotstringSelectorActive
    if (g_HotstringSelectorActive) {
        CleanupHotstringSelector()
    }
}

; Show hotstring selector GUI
ShowHotstringSelector() {
    global g_HotstringSelectorGui, g_HotstringSelectorActive, g_HotstringCharMap
    global g_HotstringHotkeyHandlers, g_HotstringCategories

    ; Close existing GUI if open
    if (g_HotstringSelectorActive && IsObject(g_HotstringSelectorGui)) {
        CleanupHotstringSelector()
        Sleep 50
    }

    ; Build character mapping
    g_HotstringCharMap := BuildHotstringCharMap()

    if (g_HotstringCharMap.Count = 0) {
        ; Use tray notification to avoid stealing focus/closing other palettes
        TrayTip("Hotstring Selector", "No hotstrings found.", "IconX")
        SetTimer(() => TrayTip(), -5000)  ; Auto-hide after ~5s
        return
    }

    ; Get categorized hotstrings
    categorized := GetCategorizedHotstrings()

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
    g_HotstringSelectorGui := Gui("+AlwaysOnTop +ToolWindow +E0x08000000", "Hotstring Shortcuts")
    ; Use slightly smaller font for better fit on small monitors
    fontSize := (monitorHeight < 800) ? 9 : 10
    g_HotstringSelectorGui.SetFont("s" . fontSize, "Segoe UI")
    g_HotstringSelectorGui.MarginX := 10
    g_HotstringSelectorGui.MarginY := 5

    ; Build reverse map: expansion -> character
    expansionToChar := Map()
    for char, expansion in g_HotstringCharMap {
        expansionToChar[expansion] := char
    }

    ; Build display text grouped by category
    ; Show all characters in sequence, including empty slots as placeholders
    displayText := ""
    hotstringCount := 0

    ; Build a map of character index to hotstring info
    charIndexToHotstring := Map()
    charIndex := 1
    for category in g_HotstringCategories {
        for hs in categorized[category] {
            if (charIndex <= g_HotstringCharSequence.Length) {
                charIndexToHotstring[charIndex] := { hotstring: hs, category: category }
            }
            charIndex++
        }
    }

    ; Display each category with header, showing all character slots
    currentCharIndex := 1
    for category in g_HotstringCategories {
        ; Calculate how many character slots belong to this category
        categorySlotCount := categorized[category].Length

        if (categorySlotCount > 0 || currentCharIndex <= g_HotstringCharSequence.Length) {
            ; Add category header
            displayText .= "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n"
            displayText .= category . "`n"
            displayText .= "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n"

            ; Display all character slots for this category (including empty ones)
            loop categorySlotCount {
                if (currentCharIndex <= g_HotstringCharSequence.Length) {
                    char := g_HotstringCharSequence[currentCharIndex]

                    ; Check if this character has a hotstring assigned
                    if (charIndexToHotstring.Has(currentCharIndex)) {
                        hsInfo := charIndexToHotstring[currentCharIndex]
                        hs := hsInfo.hotstring

                        ; Use title if available (for all categories), otherwise use preview text
                        if (hs.title != "") {
                            displayText .= "[" . char . "] > " . hs.title . "`n"
                            hotstringCount++
                        } else if (hs.expansion != "") {
                            preview := GetPreviewText(hs.expansion)
                            displayText .= "[" . char . "] > " . preview . "`n"
                            hotstringCount++
                        } else {
                            ; Empty placeholder slot
                            displayText .= "[" . char . "] > (empty)`n"
                        }
                    } else {
                        ; Character slot exists but no hotstring assigned
                        displayText .= "[" . char . "] > (empty)`n"
                    }

                    currentCharIndex++
                }
            }

            displayText .= "`n"  ; Space between categories
        }
    }

    ; Show any remaining unassigned character slots
    if (currentCharIndex <= g_HotstringCharSequence.Length) {
        displayText .= "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n"
        displayText .= "Unassigned`n"
        displayText .= "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n"
        while (currentCharIndex <= g_HotstringCharSequence.Length) {
            char := g_HotstringCharSequence[currentCharIndex]
            displayText .= "[" . char . "] > (empty)`n"
            currentCharIndex++
        }
        displayText .= "`n"
    }

    displayText .= "Press Escape to cancel."

    ; Calculate text control height based on actual content (number of lines)
    ; Count actual lines in displayText (including empty lines for spacing)
    lineCount := 1  ; Start at 1 (first line doesn't have a newline before it)
    loop parse, displayText, "`n" {
        lineCount++
    }
    ; Calculate height: ~16 pixels per line (reduced for more compact display)
    lineHeight := 16
    textControlHeight := lineCount * lineHeight
    ; Ensure minimum and maximum bounds - use smaller percentage on small monitors
    minHeight := 150
    ; Use adaptive max height: 85% for large monitors, 90% for small monitors
    maxHeightPercent := (monitorHeight < 800) ? 0.90 : 0.75
    maxHeight := Floor(monitorHeight * maxHeightPercent)
    if (textControlHeight < minHeight)
        textControlHeight := minHeight
    if (textControlHeight > maxHeight)
        textControlHeight := maxHeight

    ; Make width responsive to monitor size (smaller on small monitors)
    ; Minimum width 500px, maximum 650px, scale with monitor width
    baseWidth := (monitorWidth < 1200) ? 500 : 600
    textControlWidth := baseWidth - 20  ; Account for margins

    ; Enable vertical scrolling for long content
    g_HotstringSelectorGui.AddEdit("w" . textControlWidth . " h" . textControlHeight . " ReadOnly VScroll", displayText
    )

    ; Add Close button (set as default so it gets focus, not the Edit control)
    closeBtn := g_HotstringSelectorGui.AddButton("w100 Default Center", "Close")
    closeBtn.OnEvent("Click", (*) => CleanupHotstringSelector())

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
    g_HotstringSelectorGui.Show("NA w" . guiWidth . " h" . totalHeight . " x" . guiX . " y" . guiY)

    ; Set active flag
    g_HotstringSelectorActive := true

    ; Clear handlers array
    g_HotstringHotkeyHandlers := []

    ; Enable hotkeys for all assigned characters
    for char in g_HotstringCharSequence {
        if (g_HotstringCharMap.Has(char)) {
            ; Use factory function to create handler with properly captured char value
            handler := CreateHotstringCharHandler(char)

            ; Store handler for cleanup
            g_HotstringHotkeyHandlers.Push({ char: char, handler: handler })

            ; Enable hotkey (both uppercase and lowercase)
            try {
                ; Handle special characters that need VK codes
                if (char = ",") {
                    Hotkey("vkBC", handler, "On")  ; VK code for comma
                } else if (char = ".") {
                    Hotkey("vkBE", handler, "On")  ; VK code for period
                } else {
                    Hotkey(char, handler, "On")
                    ; Also enable uppercase for lowercase letters
                    if (RegExMatch(char, "^[a-z]$")) {
                        Hotkey(StrUpper(char), handler, "On")
                    }
                }
            } catch {
                ; Silently ignore if we can't create hotkey for this character
            }
        }
    }

    ; Enable Escape hotkey
    Hotkey("Escape", HandleHotstringEscape, "On")
}

; AI NOTE: keep the hotstring selector message consistent. The sequence of
; characters is fixed and should list every slot in order, even when a slot
; is empty, so downstream AI/debugging always sees the full set.
; Win+Alt+Shift+U hotkey
#!+U::
{
    ShowHotstringSelector()
}

; =============================================================================
; Alt+Shift+W Shortcut
; Hotkey: Alt+Shift+W
; Sends Alt+Shift+W again, then shows a message box
; =============================================================================
!+W::
{
    ; Send Alt+Shift+W again
    ; Small delay to ensure previous key release is complete
    Sleep 50

    ; Send Win+Ctrl+Alt+Y using SendInput for better reliability
    ; SendInput is more reliable for complex modifier combinations
    SendInput "#^!y"

    ; Show message box

    Sleep 50

    Send '""'

    Sleep 50

    Send "{Left}"

}

; =============================================================================
; Focus Mode (multi-monitor blackout)
; Hotkey: Win+Alt+Shift+Y
; =============================================================================

global g_FocusModeOn := false
global g_FocusModeActiveMonitor := 0
global g_FocusModeOverlays := []  ; array of GUI overlays (one per covered monitor)

FocusDbg_EscapeJson(s) {
    s := StrReplace(s, "\", "\\")
    s := StrReplace(s, '"', '\"')
    s := StrReplace(s, "`r", "")
    s := StrReplace(s, "`n", "\n")
    return s
}

FocusDbg_JsonVal(v) {
    if (IsNumber(v)) {
        return v
    }
    return '"' FocusDbg_EscapeJson(String(v)) '"'
}

FocusDbgLog(hypothesisId, location, message, data := 0, runId := "pre-fix") {
    ; Writes NDJSON lines to: <scriptdir>\.cursor\debug.log
    logPath := A_ScriptDir "\.cursor\debug.log"
    ts := DllCall("GetTickCount64", "uint64")

    dataJson := "{}"
    if (IsObject(data)) {
        parts := []
        for k, v in data {
            parts.Push('"' FocusDbg_EscapeJson(String(k)) '":' FocusDbg_JsonVal(v))
        }
        dataJson := "{" StrJoin(parts, ",") "}"
    }

    line := '{'
        . '"sessionId":"debug-session",'
        . '"runId":"' FocusDbg_EscapeJson(runId) '",'
        . '"hypothesisId":"' FocusDbg_EscapeJson(hypothesisId) '",'
        . '"location":"' FocusDbg_EscapeJson(location) '",'
        . '"message":"' FocusDbg_EscapeJson(message) '",'
        . '"timestamp":' ts ','
        . '"data":' dataJson
        . '}'

    try FileAppend(line "`n", logPath, "UTF-8")
}

StrJoin(arr, delim := ",") {
    if (!IsObject(arr) || arr.Length = 0) {
        return ""
    }
    out := ""
    for i, s in arr {
        out .= (i = 1 ? s : delim s)
    }
    return out
}

FocusDbg_TryGetMonitorDpiFromRect(l, t, r, b, &dpiX, &dpiY) {
    dpiX := ""
    dpiY := ""
    try {
        cx := l + (r - l) // 2
        cy := t + (b - t) // 2
        pt := (cy << 32) | (cx & 0xFFFFFFFF)
        hMon := DllCall("MonitorFromPoint", "int64", pt, "uint", 2, "ptr") ; MONITOR_DEFAULTTONEAREST=2
        if (!hMon) {
            return false
        }
        x := 0, y := 0
        ; MDT_EFFECTIVE_DPI = 0
        hr := DllCall("Shcore\GetDpiForMonitor", "ptr", hMon, "int", 0, "uint*", &x, "uint*", &y, "uint")
        if (hr != 0) {
            return false
        }
        dpiX := x
        dpiY := y
        return true
    } catch {
        return false
    }
}

GetActiveMonitorIndex() {
    hwnd := WinExist("A")
    if (!hwnd) {
        ;#region agent log
        FocusDbgLog("H4", "Utils.ahk:GetActiveMonitorIndex", "No active hwnd", Map("hwnd", 0))
        ;#endregion agent log
        return 0
    }

    rect := Buffer(16, 0)
    if (!DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect)) {
        ;#region agent log
        FocusDbgLog("H4", "Utils.ahk:GetActiveMonitorIndex", "GetWindowRect failed", Map("hwnd", hwnd))
        ;#endregion agent log
        return 0
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
            ;#region agent log
            FocusDbgLog("H2", "Utils.ahk:GetActiveMonitorIndex", "Matched monitor by center point", Map(
                "hwnd", hwnd,
                "centerX", centerX, "centerY", centerY,
                "matchedMonitor", A_Index,
                "ml", ml, "mt", mt, "mr", mr, "mb", mb
            ))
            ;#endregion agent log
            return A_Index
        }
    }

    ;#region agent log
    FocusDbgLog("H4", "Utils.ahk:GetActiveMonitorIndex", "No monitor matched center point", Map(
        "hwnd", hwnd,
        "centerX", centerX, "centerY", centerY,
        "monitorCount", monitorCount
    ))
    ;#endregion agent log
    return 0
}

EnableFocusMode() {
    global g_FocusModeOn, g_FocusModeActiveMonitor, g_FocusModeOverlays

    if (g_FocusModeOn) {
        ;#region agent log
        FocusDbgLog("H5", "Utils.ahk:EnableFocusMode", "Already enabled; no-op", Map("g_FocusModeOn", g_FocusModeOn))
        ;#endregion agent log
        return
    }

    activeMon := GetActiveMonitorIndex()
    if (!activeMon) {
        ;#region agent log
        FocusDbgLog("H4", "Utils.ahk:EnableFocusMode", "Active monitor detection failed; abort", Map("activeMon",
            activeMon))
        ;#endregion agent log
        return
    }

    g_FocusModeActiveMonitor := activeMon
    g_FocusModeOverlays := []

    monitorCount := MonitorGetCount()

    ;#region agent log
    ctx := 0, awareness := ""
    try {
        ctx := DllCall("GetProcessDpiAwarenessContext", "ptr")
        awareness := DllCall("GetAwarenessFromDpiAwarenessContext", "ptr", ctx, "int")
    }
    FocusDbgLog("H1", "Utils.ahk:EnableFocusMode", "DPI + monitor count snapshot", Map(
        "activeMon", activeMon,
        "monitorCount", monitorCount,
        "dpiContextPtr", ctx,
        "dpiAwareness", awareness
    ))
    ;#endregion agent log

    loop monitorCount {
        MonitorGet(A_Index, &ml, &mt, &mr, &mb)
        name := ""
        dx := "", dy := ""
        try name := MonitorGetName(A_Index)
        try FocusDbg_TryGetMonitorDpiFromRect(ml, mt, mr, mb, &dx, &dy)
        ;#region agent log
        FocusDbgLog("H2", "Utils.ahk:EnableFocusMode", "MonitorGet bounds", Map(
            "i", A_Index,
            "name", name,
            "dpiX", dx, "dpiY", dy,
            "ml", ml, "mt", mt, "mr", mr, "mb", mb,
            "w", mr - ml, "h", mb - mt
        ))
        ;#endregion agent log
    }

    loop monitorCount {
        i := A_Index
        if (i = activeMon) {
            continue
        }

        MonitorGet(i, &l, &t, &r, &b)
        w := r - l
        h := b - t
        if (w <= 0 || h <= 0) {
            continue
        }

        overlay := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x20") ; WS_EX_TRANSPARENT => click-through
        ; Critical: disable GUI DPI scaling so x/y/w/h are interpreted in raw screen coords
        ; Without this, AHK scales the overlay (e.g. 150%) and it spills into other monitors.
        overlay.Opt("-DPIScale")
        overlay.BackColor := "000000"
        overlay.Show("NA x" l " y" t " w" w " h" h)
        g_FocusModeOverlays.Push(overlay)

        ; Query actual overlay window rect (physical pixels) to detect scaling/border issues
        rect := Buffer(16, 0)
        ok := 0
        try ok := DllCall("GetWindowRect", "ptr", overlay.Hwnd, "ptr", rect)
        if (ok) {
            wl := NumGet(rect, 0, "int"), wt := NumGet(rect, 4, "int")
            wr := NumGet(rect, 8, "int"), wb := NumGet(rect, 12, "int")
            ;#region agent log
            FocusDbgLog("H3", "Utils.ahk:EnableFocusMode", "Overlay shown; actual rect", Map(
                "monitorIndex", i,
                "targetL", l, "targetT", t, "targetW", w, "targetH", h,
                "actualL", wl, "actualT", wt, "actualR", wr, "actualB", wb,
                "actualW", wr - wl, "actualH", wb - wt
            ))
            ;#endregion agent log
        } else {
            ;#region agent log
            FocusDbgLog("H3", "Utils.ahk:EnableFocusMode", "Overlay shown; GetWindowRect failed", Map(
                "monitorIndex", i,
                "targetL", l, "targetT", t, "targetW", w, "targetH", h
            ))
            ;#endregion agent log
        }
    }

    g_FocusModeOn := true
}

DisableFocusMode() {
    global g_FocusModeOn, g_FocusModeActiveMonitor, g_FocusModeOverlays

    for overlay in g_FocusModeOverlays {
        try {
            if (IsObject(overlay) && overlay.Hwnd) {
                overlay.Destroy()
            }
        } catch {
            ; Ignore
        }
    }

    g_FocusModeOverlays := []
    g_FocusModeActiveMonitor := 0
    g_FocusModeOn := false
}

ToggleFocusMode() {
    global g_FocusModeOn
    ;#region agent log
    FocusDbgLog("H5", "Utils.ahk:ToggleFocusMode", "Toggle invoked", Map(
        "g_FocusModeOn_before", g_FocusModeOn,
        "hwndActive", WinExist("A")
    ))
    ;#endregion agent log
    if (g_FocusModeOn) {
        DisableFocusMode()
    } else {
        EnableFocusMode()
    }
}

#!+Y::
{
    ToggleFocusMode()
}

; =============================================================================
; Print Screen with Chime
; Hotkey: Alt+PrintScreen
; Lets the native Alt+PrintScreen pass through (captures active window)
; and plays a chime sound to confirm the action
; =============================================================================
~!PrintScreen::
{
    ; Let the key combination pass through to Windows (captures active window)
    ; The ~ prefix allows the original key to work normally

    ; Brief delay to ensure screenshot is taken, then play chime
    ; This is safe - one-shot timer automatically cleans up after execution
    Sleep 10
    SoundBeep(800, 200)
}

; =============================================================================
; Block Escape Key from handy.exe
; Prevents Escape key from closing handy.exe while still allowing
; Escape to work normally in other applications
; =============================================================================

; Force hook-based hotkey to intercept Escape at a lower level
; This should catch it before handy.exe's hook processes it
#UseHook
#InputLevel 10
Escape::
{
    ;#region agent log
    FocusDbgLog("H1", "Utils.ahk:Escape", "Escape hotkey triggered", Map("timestamp", A_TickCount))
    ;#endregion agent log

    ; Get active window info before checking
    activeWin := WinGetID("A")
    activeWinTitle := WinGetTitle("A")
    activeWinProcess := WinGetProcessName("A")
    activeWinProcessPath := WinGetProcessPath("A")
    isHandyActive := WinActive("ahk_exe handy.exe")

    ; Try alternative detection methods
    isHandyByTitle := InStr(activeWinTitle, "handy", false) > 0
    isHandyByProcess := InStr(activeWinProcess, "handy", false) > 0
    isHandyByPath := InStr(activeWinProcessPath, "handy", false) > 0

    ; Check if the "Recording" window from handy.exe exists
    ; Only block Escape when actually recording, not just when handy.exe is running
    recordingWindowExists := false
    try {
        recordingWindowExists := WinExist("Recording ahk_exe handy.exe")
    } catch {
        recordingWindowExists := false
    }

    ;#region agent log
    FocusDbgLog("H2", "Utils.ahk:Escape", "Process and window check", Map(
        "activeWin", activeWin,
        "activeWinTitle", activeWinTitle,
        "activeWinProcess", activeWinProcess,
        "activeWinProcessPath", activeWinProcessPath,
        "isHandyActive", isHandyActive,
        "recordingWindowExists", recordingWindowExists,
        "note", "Only blocking when Recording window exists"
    ))
    ;#endregion agent log

    ; Block Escape only if the "Recording" window exists
    ; This allows Escape to work normally when handy.exe is running but not recording
    shouldBlock := (recordingWindowExists > 0)

    ;#region agent log
    FocusDbgLog("H4", "Utils.ahk:Escape", "Blocking decision", Map(
        "recordingWindowExists", recordingWindowExists,
        "shouldBlock", shouldBlock,
        "reason", "Blocking because Recording window exists (dictation active)"
    ))
    ;#endregion agent log

    ; Block Escape if handy.exe is running
    if (shouldBlock) {
        ;#region agent log
        FocusDbgLog("H3", "Utils.ahk:Escape", "Blocking Escape for handy.exe", Map(
            "blocking", true,
            "returning", true
        ))
        ;#endregion agent log

        ; Block Escape from reaching handy.exe - do nothing
        ; This prevents the dictation software from closing
        return
    }

    ;#region agent log
    FocusDbgLog("H3", "Utils.ahk:Escape", "Forwarding Escape to system", Map(
        "blocking", false,
        "forwarding", true
    ))
    ;#endregion agent log

    ; Otherwise, forward Escape to the system
    ; Use SendInput for more reliable key forwarding
    SendInput "{Escape}"

    ;#region agent log
    FocusDbgLog("H3", "Utils.ahk:Escape", "Escape forwarded", Map("afterSend", true))
    ;#endregion agent log
}
#InputLevel 0

;#region agent log
; Log that Escape hotkey section was loaded
FocusDbgLog("H1", "Utils.ahk", "Escape hotkey registered", Map("inputLevel", 10, "useHook", true))
;#endregion agent log

; =============================================================================
; Dictation Indicator - Red Pulsing Square
; Shows a red pulsing square at the top center of the active monitor when
; dictation is active with handy.exe. Toggles with Win+Alt+Shift+0.
; =============================================================================

; Global variables for dictation indicator
global g_DictationActive := false
global g_DictationIndicatorGui := false
global g_DictationPulseTimer := false
global g_DictationCheckTimer := false  ; Timer to check if Recording window still exists
global g_DictationPulseDirection := 1  ; 1 = fading in, -1 = fading out
global g_DictationPulseOpacity := 128  ; Current opacity (50-255)
global g_LastActiveMonitor := 0  ; Track last monitor to detect changes

; Constants for dictation indicator
global DICTATION_SQUARE_SIZE := 50
global DICTATION_PULSE_MIN := 50      ; Minimum opacity (~20%)
global DICTATION_PULSE_MAX := 255     ; Maximum opacity (100%)
global DICTATION_PULSE_STEP := 15     ; Opacity change per tick
global DICTATION_PULSE_INTERVAL := 50 ; Timer interval in ms (smooth animation)

; Get the monitor that contains the active window
; Returns monitor index (1-based) or 0 if not found
GetDictationActiveMonitor() {
    hwnd := WinExist("A")
    if (!hwnd) {
        return 1  ; Default to primary monitor
    }

    rect := Buffer(16, 0)
    if (!DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect)) {
        return 1  ; Default to primary monitor
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
            return A_Index
        }
    }

    return 1  ; Default to primary monitor
}

; Show or update the dictation indicator at the top center of the active monitor
ShowDictationIndicator() {
    global g_DictationIndicatorGui, g_DictationPulseOpacity, g_LastActiveMonitor
    global DICTATION_SQUARE_SIZE

    ; Get active monitor bounds
    monitorIndex := GetDictationActiveMonitor()
    MonitorGet(monitorIndex, &monitorLeft, &monitorTop, &monitorRight, &monitorBottom)
    monitorWidth := monitorRight - monitorLeft

    ; Calculate position: top center of monitor with small offset from top
    squareX := monitorLeft + (monitorWidth - DICTATION_SQUARE_SIZE) // 2
    squareY := monitorTop + 20  ; 20 pixels from top

    ; Check if indicator already exists
    if (IsObject(g_DictationIndicatorGui) && g_DictationIndicatorGui.Hwnd) {
        ; Indicator exists - just reposition it if monitor changed
        if (monitorIndex != g_LastActiveMonitor) {
            g_DictationIndicatorGui.Show("NA x" . squareX . " y" . squareY . " w" . DICTATION_SQUARE_SIZE . " h" .
                DICTATION_SQUARE_SIZE)
            g_LastActiveMonitor := monitorIndex
        }
        return
    }

    ; Create new indicator
    ; +AlwaysOnTop: stays on top of all windows
    ; -Caption: no title bar
    ; +ToolWindow: doesn't appear in taskbar
    ; +E0x20: click-through (WS_EX_TRANSPARENT)
    ; -DPIScale: use raw screen coordinates
    g_DictationIndicatorGui := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x20")
    g_DictationIndicatorGui.Opt("-DPIScale")
    g_DictationIndicatorGui.BackColor := "FF0000"  ; Red

    ; Reset pulse opacity
    g_DictationPulseOpacity := 128
    g_LastActiveMonitor := monitorIndex

    ; Show the indicator without activating it
    g_DictationIndicatorGui.Show("NA x" . squareX . " y" . squareY . " w" . DICTATION_SQUARE_SIZE . " h" .
        DICTATION_SQUARE_SIZE)
    WinSetTransparent(g_DictationPulseOpacity, g_DictationIndicatorGui)
}

; Hide and destroy the dictation indicator
HideDictationIndicator() {
    global g_DictationIndicatorGui

    if (IsObject(g_DictationIndicatorGui)) {
        try {
            if (g_DictationIndicatorGui.Hwnd) {
                g_DictationIndicatorGui.Destroy()
            }
        } catch {
            ; Ignore errors
        }
        g_DictationIndicatorGui := false
    }
}

; Update the pulse animation (called by timer)
UpdateDictationIndicatorPulse() {
    global g_DictationIndicatorGui, g_DictationPulseOpacity, g_DictationPulseDirection
    global DICTATION_PULSE_MIN, DICTATION_PULSE_MAX, DICTATION_PULSE_STEP

    ; Check if indicator GUI is still valid
    if (!IsObject(g_DictationIndicatorGui) || !g_DictationIndicatorGui.Hwnd) {
        return
    }

    ; Update opacity based on direction
    g_DictationPulseOpacity += DICTATION_PULSE_STEP * g_DictationPulseDirection

    ; Reverse direction at bounds
    if (g_DictationPulseOpacity >= DICTATION_PULSE_MAX) {
        g_DictationPulseOpacity := DICTATION_PULSE_MAX
        g_DictationPulseDirection := -1  ; Start fading out
    } else if (g_DictationPulseOpacity <= DICTATION_PULSE_MIN) {
        g_DictationPulseOpacity := DICTATION_PULSE_MIN
        g_DictationPulseDirection := 1   ; Start fading in
    }

    ; Apply new transparency
    try {
        WinSetTransparent(g_DictationPulseOpacity, g_DictationIndicatorGui)
    } catch {
        ; Ignore errors (window might have been destroyed)
    }
}

; Start the pulse animation timer
StartDictationPulseTimer() {
    global g_DictationPulseTimer, DICTATION_PULSE_INTERVAL, g_DictationPulseDirection

    ; Reset pulse direction to fade in
    g_DictationPulseDirection := 1

    ; Stop any existing timer
    StopDictationPulseTimer()

    ; Create and start new timer
    g_DictationPulseTimer := UpdateDictationIndicatorPulse
    SetTimer(g_DictationPulseTimer, DICTATION_PULSE_INTERVAL)
}

; Stop the pulse animation timer
StopDictationPulseTimer() {
    global g_DictationPulseTimer

    if (g_DictationPulseTimer) {
        try {
            SetTimer(g_DictationPulseTimer, 0)
        } catch {
            ; Ignore errors
        }
        g_DictationPulseTimer := false
    }
}

; Check if Recording window still exists and update indicator accordingly
; Also updates indicator position to follow active window
CheckDictationRecordingWindow() {
    global g_DictationActive

    ; Check if the "Recording" window exists
    recordingWindowExists := false
    try {
        recordingWindowExists := WinExist("Recording ahk_exe handy.exe")
    } catch {
        recordingWindowExists := false
    }

    ; Update state based on whether Recording window exists
    newState := (recordingWindowExists > 0)

    ; Update if state changed
    if (newState != g_DictationActive) {
        g_DictationActive := newState

        if (g_DictationActive) {
            ; Recording window appeared - show indicator
            ShowDictationIndicator()
            StartDictationPulseTimer()
        } else {
            ; Recording window disappeared - hide indicator
            StopDictationPulseTimer()
            HideDictationIndicator()
        }
    } else if (g_DictationActive) {
        ; Recording is active - update indicator position to follow active window
        ShowDictationIndicator()  ; This will reposition if monitor changed
    }
}

; Start timer to periodically check Recording window state
StartDictationCheckTimer() {
    global g_DictationCheckTimer

    ; Stop any existing timer
    StopDictationCheckTimer()

    ; Check every 500ms
    g_DictationCheckTimer := CheckDictationRecordingWindow
    SetTimer(g_DictationCheckTimer, 500)
}

; Stop the check timer
StopDictationCheckTimer() {
    global g_DictationCheckTimer

    if (g_DictationCheckTimer) {
        try {
            SetTimer(g_DictationCheckTimer, 0)
        } catch {
            ; Ignore errors
        }
        g_DictationCheckTimer := false
    }
}

; Toggle dictation mode on/off
; The check timer handles everything automatically, this just triggers an immediate check
ToggleDictationMode() {
    ; Trigger immediate check (the timer will handle showing/hiding)
    CheckDictationRecordingWindow()
}

; Cleanup dictation indicator resources
CleanupDictationIndicator(*) {
    StopDictationPulseTimer()
    StopDictationCheckTimer()
    HideDictationIndicator()
}

; Register cleanup on script exit
OnExit(CleanupDictationIndicator)

; Toggle dictation mode with Win+Alt+Shift+0
; The ~ prefix allows the key combination to pass through to handy.exe
~#!+0::
{
    ToggleDictationMode()
}

; Start the check timer automatically when script loads
; This continuously monitors for the Recording window and updates indicator position
StartDictationCheckTimer()