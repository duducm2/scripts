; =============================================================================
; Shift keys module: outlook_appointment_palette.ahk
; Outlook appointment configuration palette
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

; =============================================================================
; Outlook Appointment Configuration Palette
; Shows a grid of 24 letter-labeled squares for selecting appointment configurations
; Shift + . → Show palette (display only, no actions triggered yet)
; =============================================================================

; Global variables for Outlook Appointment palette
global g_OutlookPaletteActive := false
global g_OutlookPaletteGuis := []
global g_OutlookPaletteTimer := false
global g_OutlookPaletteSessionID := 0

; Letter mapping for Outlook Appointment palette (24 combinations)
; Format: [Letter, Status, All-day, Private, Reminder]
; Status: 1=Free, 2=Busy, 3=Out of office
; All-day: 1=Yes, 2=No
; Private: 1=Off, 2=On
; Reminder: 1=15min, 2=2days
global g_OutlookPaletteMapping := Map(
    "Q", { Status: 1, AllDay: 1, Private: 1, Reminder: 1 },  ; Free, All-day Yes, Private Off, 15min
    "W", { Status: 1, AllDay: 1, Private: 1, Reminder: 2 },  ; Free, All-day Yes, Private Off, 2days
    "E", { Status: 1, AllDay: 1, Private: 2, Reminder: 1 },  ; Free, All-day Yes, Private On, 15min
    "R", { Status: 1, AllDay: 1, Private: 2, Reminder: 2 },  ; Free, All-day Yes, Private On, 2days
    "A", { Status: 1, AllDay: 2, Private: 1, Reminder: 1 },  ; Free, All-day No, Private Off, 15min
    "S", { Status: 1, AllDay: 2, Private: 1, Reminder: 2 },  ; Free, All-day No, Private Off, 2days
    "D", { Status: 1, AllDay: 2, Private: 2, Reminder: 1 },  ; Free, All-day No, Private On, 15min
    "F", { Status: 1, AllDay: 2, Private: 2, Reminder: 2 },  ; Free, All-day No, Private On, 2days
    "Z", { Status: 2, AllDay: 1, Private: 1, Reminder: 1 },  ; Busy, All-day Yes, Private Off, 15min
    "X", { Status: 2, AllDay: 1, Private: 1, Reminder: 2 },  ; Busy, All-day Yes, Private Off, 2days
    "C", { Status: 2, AllDay: 1, Private: 2, Reminder: 1 },  ; Busy, All-day Yes, Private On, 15min
    "V", { Status: 2, AllDay: 1, Private: 2, Reminder: 2 },  ; Busy, All-day Yes, Private On, 2days
    "B", { Status: 2, AllDay: 2, Private: 1, Reminder: 1 },  ; Busy, All-day No, Private Off, 15min
    "N", { Status: 2, AllDay: 2, Private: 1, Reminder: 2 },  ; Busy, All-day No, Private Off, 2days
    "M", { Status: 2, AllDay: 2, Private: 2, Reminder: 1 },  ; Busy, All-day No, Private On, 15min
    ",", { Status: 2, AllDay: 2, Private: 2, Reminder: 2 },  ; Busy, All-day No, Private On, 2days
    "U", { Status: 3, AllDay: 1, Private: 1, Reminder: 1 },  ; Out of office, All-day Yes, Private Off, 15min
    "I", { Status: 3, AllDay: 1, Private: 1, Reminder: 2 },  ; Out of office, All-day Yes, Private Off, 2days
    "O", { Status: 3, AllDay: 1, Private: 2, Reminder: 1 },  ; Out of office, All-day Yes, Private On, 15min
    "P", { Status: 3, AllDay: 1, Private: 2, Reminder: 2 },  ; Out of office, All-day Yes, Private On, 2days
    "J", { Status: 3, AllDay: 2, Private: 1, Reminder: 1 },  ; Out of office, All-day No, Private Off, 15min
    "K", { Status: 3, AllDay: 2, Private: 1, Reminder: 2 },  ; Out of office, All-day No, Private Off, 2days
    "L", { Status: 3, AllDay: 2, Private: 2, Reminder: 1 },  ; Out of office, All-day No, Private On, 15min
    ";", { Status: 3, AllDay: 2, Private: 2, Reminder: 2 }   ; Out of office, All-day No, Private On, 2days
)

; Letters in display order (3 status groups, each with 2 rows × 4 columns)
global g_OutlookPaletteLetters := [
    ; Free status (row 1: All-day Yes, row 2: All-day No)
    "Q", "W", "E", "R",
    "A", "S", "D", "F",
    ; Busy status (row 1: All-day Yes, row 2: All-day No)
    "Z", "X", "C", "V",
    "B", "N", "M", ",",
    ; Out of office status (row 1: All-day Yes, row 2: All-day No)
    "U", "I", "O", "P",
    "J", "K", "L", ";"
]

; Timer handler for palette timeout
OutlookPaletteTimerHandler(sessionID) {
    global g_OutlookPaletteActive, g_OutlookPaletteSessionID, g_OutlookPaletteTimer
    if (sessionID != g_OutlookPaletteSessionID) {
        return
    }
    if (!g_OutlookPaletteActive) {
        g_OutlookPaletteTimer := false
        return
    }
    CleanupOutlookPalette()
    g_OutlookPaletteTimer := false
}

; Cleanup function for Outlook palette
CleanupOutlookPalette() {
    global g_OutlookPaletteGuis, g_OutlookPaletteActive
    g_OutlookPaletteActive := false
    for gui in g_OutlookPaletteGuis {
        try {
            if (IsObject(gui) && gui.Hwnd) {
                gui.Hide()
                gui.Destroy()
            }
        } catch {
            ; Silently ignore errors
        }
    }
    g_OutlookPaletteGuis := []
}

; Show Outlook Appointment palette
ShowOutlookAppointmentPalette() {
    global g_OutlookPaletteActive, g_OutlookPaletteGuis, g_OutlookPaletteLetters
    global g_OutlookPaletteTimer, g_OutlookPaletteSessionID

    ; Cleanup any existing palette
    if (g_OutlookPaletteActive) {
        CleanupOutlookPalette()
    }

    ; Increment session ID
    g_OutlookPaletteSessionID++
    g_OutlookPaletteActive := true

    ; Configuration
    squareSize := 40
    spacing := 8
    statusGroupSpacing := 20  ; Space between status groups
    rowsPerStatus := 2
    colsPerStatus := 4

    ; Get mouse position for palette placement
    MouseGetPos(&startX, &startY)

    ; Calculate positions for 3 status groups (each 2×4 grid)
    ; Layout: 3 groups side by side, each group is 2 rows × 4 columns
    guiArray := []
    statusGroupWidth := (squareSize * colsPerStatus) + (spacing * (colsPerStatus - 1))

    loop 3 {  ; 3 status groups
        statusIndex := A_Index
        groupOffsetX := (statusIndex - 1) * (statusGroupWidth + statusGroupSpacing)

        loop rowsPerStatus {  ; 2 rows per status
            rowIndex := A_Index
            loop colsPerStatus {  ; 4 columns per row
                colIndex := A_Index

                ; Calculate letter index in the flat array
                letterIndex := ((statusIndex - 1) * rowsPerStatus * colsPerStatus) +
                ((rowIndex - 1) * colsPerStatus) + colIndex

                if (letterIndex > g_OutlookPaletteLetters.Length) {
                    continue
                }

                letter := g_OutlookPaletteLetters[letterIndex]

                ; Calculate position
                squareX := startX + groupOffsetX + ((colIndex - 1) * (squareSize + spacing))
                squareY := startY + ((rowIndex - 1) * (squareSize + spacing))

                ; Create square GUI
                squareGui := Gui("+AlwaysOnTop -Caption +ToolWindow")
                squareGui.BackColor := "333333"  ; Dark gray background
                squareGui.SetFont("s10 Bold cFFFFFF", "Segoe UI")
                squareGui.MarginX := 0
                squareGui.MarginY := 0

                ; Add letter text
                letterText := squareGui.AddText("w" . squareSize . " h" . squareSize . " Center 0x201", letter)

                ; Calculate top-left position
                guiX := Round(squareX - squareSize / 2.0)
                guiY := Round(squareY - squareSize / 2.0)

                guiArray.Push({ gui: squareGui, x: guiX, y: guiY })
            }
        }
    }

    ; Position all GUIs while hidden
    for guiInfo in guiArray {
        guiInfo.gui.Show("x" . guiInfo.x . " y" . guiInfo.y . " w" . squareSize . " h" . squareSize . " NA Hide")
        WinSetTransparent(220, guiInfo.gui)  ; ~86% opacity
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
        g_OutlookPaletteGuis.Push(guiInfo.gui)
    }

    ; Set timeout timer (5 seconds)
    timerHandler := () => OutlookPaletteTimerHandler(g_OutlookPaletteSessionID)
    g_OutlookPaletteTimer := timerHandler
    SetTimer(timerHandler, -5000)
}

; -----------------------------------------------------------------------------
; Outlook Appointment – Cascaded selection via dialogs (good UX, text-focused)
; Uses a 3-step flow:
;   1) Pick Private × All-day (4 options)
;   2) Pick Status (3 options)
;   3) Pick Reminder (2 options)
; Final result is shown as a clear text summary (no fields changed yet).
; -----------------------------------------------------------------------------

; Global variable to store Outlook Appointment selection choice
global g_OutlookAppointmentChoice := ""

; Cancel handler for Outlook Appointment selection dialogs
CancelOutlookOption(optionGui, *) {
    global g_OutlookAppointmentChoice
    g_OutlookAppointmentChoice := ""
    optionGui.Destroy()
}

; Auto-submit handler for Outlook Appointment selection dialogs
Outlook_OptionAutoSubmit(ctrl, optionsMap) {
    global g_OutlookAppointmentChoice
    currentValue := ctrl.Text
    if (currentValue != "" && IsInteger(currentValue)) {
        choice := Integer(currentValue)
        choiceStr := String(choice)
        if (optionsMap.Has(choiceStr)) {
            ; Store the choice in global variable
            g_OutlookAppointmentChoice := choiceStr
            ctrl.Gui.Destroy()
        }
    }
}

; Factory function to create auto-submit handler with captured optionsMap
CreateOutlookOptionHandler(optionsMap) {
    return (ctrl, *) => Outlook_OptionAutoSubmit(ctrl, optionsMap)
}

; Show selection dialog with immediate auto-submit on number entry (no Enter needed)
Outlook_SelectOptionByInputBox(title, basePrompt, optionsMap) {
    ; Build prompt text with better spacing and grouping
    prompt := basePrompt . "`n`n"
    validList := ""
    lastGroup := 0
    for key, opt in optionsMap {
        ; Add extra spacing between groups if this option has a Group property
        if (opt.HasProp("Group") && opt.Group != lastGroup && lastGroup > 0) {
            prompt .= "`n"  ; Add blank line between groups
        }
        if (opt.HasProp("Group")) {
            lastGroup := opt.Group
        }

        prompt .= key . ") " . opt.Label . "`n"
        if (validList != "")
            validList .= ", "
        validList .= key
    }
    prompt .= "`n`nType a number (" . validList . "):"

    ; Create GUI dialog
    try {
        optionGui := Gui("+AlwaysOnTop +ToolWindow", title)
        optionGui.SetFont("s10", "Segoe UI")
        optionGui.AddText("w480 Center", prompt)
        optionGui.AddEdit("w60 Center vOptionInput", "")

        ; Set up auto-submit handler using factory function to capture optionsMap
        handler := CreateOutlookOptionHandler(optionsMap)
        optionGui["OptionInput"].OnEvent("Change", handler)

        ; Add Cancel button
        cancelBtn := optionGui.AddButton("w80", "Cancel")
        cancelBtn.OnEvent("Click", CancelOutlookOption.Bind(optionGui))

        optionGui.Show("w500 h250")
        optionGui["OptionInput"].Focus()

        ; Wait for dialog to close
        WinWaitClose("ahk_id " optionGui.Hwnd)

        ; Retrieve the selected choice from global variable
        global g_OutlookAppointmentChoice
        choice := g_OutlookAppointmentChoice
        g_OutlookAppointmentChoice := ""  ; Clear for next use
        return choice
    } catch Error as e {
        MsgBox "Error in selection dialog: " . e.Message, title . " Error", "IconX"
        return ""
    }
}
