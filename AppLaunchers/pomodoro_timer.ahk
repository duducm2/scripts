; =============================================================================
; AppLaunchers module: pomodoro_timer.ahk
; Pomodoro timer system with CSV logging
; Extracted verbatim from AppLaunchers.ahk; loaded via #include into the
; AppLaunchers.ahk process, which remains the entry point / source of truth.
; =============================================================================

; =============================================================================
; Pomodoro Timer System - Local Timer with CSV Logging
; (No global hotkey assigned; Win+Alt+Shift+9 is free.)
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
