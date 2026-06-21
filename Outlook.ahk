#Requires AutoHotkey v2.0+
#SingleInstance Force

; -----------------------------------------------------------------------------
; This script consolidates all Outlook related hotkeys and functions.
; -----------------------------------------------------------------------------

; --- Includes ----------------------------------------------------------------
#include vendor\UIA-v2\Lib\UIA.ahk
#include %A_ScriptDir%\Utils.ahk
try Hotkey("#!+Y", "Off")
try Hotkey("#!+X", "Off")

; --- Configuration -----------------------------------------------------------
; Classic: rctrl_renwnd32. New Outlook (Monarch): Outlook Host.
global OUTLOOK_EXE := "OUTLOOK.EXE"
; New Outlook (Microsoft Store / Monarch) runs as olk.exe — same windows as "Outlook Host".
global OUTLOOK_OLK_EXE := "olk.exe"
global OUTLOOK_CLASS := "rctrl_renwnd32"
global OUTLOOK_CLASS_HOST := "Outlook Host"
global OUTLOOK_BASE_CRIT := "ahk_class " OUTLOOK_CLASS " ahk_exe " OUTLOOK_EXE

; Mailbox: classic titles may contain this (and not OUTLOOK_MAILBOX_TITLE_EXCLUDE). New Outlook uses
; "Mail - … - Outlook" / "Inbox - … - Outlook" (see OutlookTitleIsMailModule).
global OUTLOOK_MAILBOX_TITLE_CONTAINS := "Eduardo.Figueiredo@br.bosch.com"
global OUTLOOK_MAILBOX_TITLE_EXCLUDE := "Calendar"

; Calendar: legacy partial needles plus "Calendar - … - Outlook" handled in OutlookTitleIsCalendarModule.
global OUTLOOK_CALENDAR_TITLES := ["Calendar - Eduardo"]

; Reminders: substring match (classic "Reminder(s)"; new "Reminders - …")
global OUTLOOK_REMINDER_TITLE := "Reminder"

OutlookIsMainFrameClass(cls) {
    return cls = OUTLOOK_CLASS || cls = OUTLOOK_CLASS_HOST
}

OutlookTitleIsExcludedMainSurface(title) {
    if RegExMatch(title, "i)^(Calendar|Reminders|Copilot)\s-")
        return true
    if RegExMatch(title, "i)^(New event)\b")
        return true
    return false
}

; Classic mailbox window title, or new Outlook (Mail / Inbox / …) top-level window.
OutlookTitleIsMailModule(title) {
    if InStr(title, OUTLOOK_MAILBOX_TITLE_CONTAINS) && !InStr(title, OUTLOOK_MAILBOX_TITLE_EXCLUDE)
        return true
    if !InStr(title, " - Outlook")
        return false
    if OutlookTitleIsExcludedMainSurface(title)
        return false
    return RegExMatch(title, "i)^(Mail|Inbox|Drafts|Sent Items|Deleted Items|Junk Email|Outbox|Archive)\s-\s")
}

OutlookTitleIsCalendarModule(title) {
    if InStr(title, "Calendar -") && InStr(title, " - Outlook")
        return true
    for _, needle in OUTLOOK_CALENDAR_TITLES {
        if InStr(title, needle)
            return true
    }
    return false
}

OutlookProcessNameIsOutlook(procName) {
    switch StrLower(procName) {
        case "outlook.exe", "olk.exe":
            return true
        default:
            return false
    }
}

OutlookWinActive() {
    return WinActive("ahk_exe " OUTLOOK_EXE) || WinActive("ahk_exe " OUTLOOK_OLK_EXE)
}

; Timeouts (ms)
global OUTLOOK_ACTIVATE_WAIT_MS := 2000
global OUTLOOK_VOICE_WAIT_MS := 1500
global OUTLOOK_BANNER_MS := 2000
global OUTLOOK_BANNER_FAIL_MS := 3000
global OUTLOOK_ACTIVATE_PROMPT_TIMEOUT_MS := 6000
global OUTLOOK_SWITCH_VERIFY_TIMEOUT_MS := 700

global g_OutlookActivatePromptPending := false

; Voice Aloud option range
global VOICE_ALOUD_OPTION_MIN := 1
global VOICE_ALOUD_OPTION_MAX := 2

; Feature flags. Rollout order: 1) HWND cache + state waits, 2) COM core, 3) WM_COMMAND/UIA Read Aloud.
; Parity: #!+b #!+g #!+v #!+d behave as before. Safety: no blind key injection; no leaked SetTitleMatchMode.
global OUTLOOK_USE_HWND_CACHE := true
global OUTLOOK_USE_STATE_WAITS := true
global OUTLOOK_USE_COM_CORE := false
global OUTLOOK_USE_WMCOMMAND_READALOUD := false
global OUTLOOK_USE_UIA_READALOUD := true

; --- Read Aloud UI-bound abstraction (Phase II) ------------------------------
; Invokes "Read Aloud" via UIA or WM_COMMAND when flags set; else synthetic Alt+1.
; Returns true if invocation was performed (UIA/keystroke), false on failure.
InvokeReadAloudStart(hwnd) {
    if (hwnd <= 0)
        return false
    if OUTLOOK_USE_UIA_READALOUD {
        try {
            root := UIA.ElementFromHandle(hwnd)
            for _, name in ["Read Aloud", "Ler em voz alta", "Read aloud"] {
                btn := root.FindFirst({ Type: 50000, Name: name })
                if btn {
                    btn.Invoke()
                    return true
                }
            }
        } catch {
        }
    }
    if OUTLOOK_USE_WMCOMMAND_READALOUD {
        ; WM_COMMAND 0x0111; wParam high word = 0, low word = control ID (if known)
        try {
            if DllCall("PostMessage", "Ptr", hwnd, "UInt", 0x0111, "UPtr", 0, "Ptr", 0)
                return true
        } catch {
        }
    }
    ; Fallback: synthetic keystroke (Outlook must be foreground).
    Send "{Alt down}1{Alt up}"
    return true
}

; --- Singleton HWND cache ----------------------------------------------------
class OutlookHwndCache {
    static MailboxHwnd := 0
    static CalendarHwnd := 0
    static ReminderHwnd := 0

    static _Valid(hwnd) {
        if (!(hwnd is Integer) || hwnd <= 0)
            return false
        if !WinExist("ahk_id " hwnd)
            return false
        try {
            return OutlookProcessNameIsOutlook(WinGetProcessName("ahk_id " hwnd))
        } catch {
            return false
        }
    }

    static GetMailboxHwnd() {
        if OUTLOOK_USE_HWND_CACHE && OutlookHwndCache._Valid(OutlookHwndCache.MailboxHwnd)
            return OutlookHwndCache.MailboxHwnd
        hwnd := OutlookHwndCache._ResolveMailbox()
        if (hwnd > 0)
            OutlookHwndCache.MailboxHwnd := hwnd
        else
            OutlookHwndCache.MailboxHwnd := 0
        return hwnd
    }

    static GetCalendarHwnd() {
        if OUTLOOK_USE_HWND_CACHE && OutlookHwndCache._Valid(OutlookHwndCache.CalendarHwnd)
            return OutlookHwndCache.CalendarHwnd
        hwnd := OutlookHwndCache._ResolveCalendar()
        if (hwnd > 0)
            OutlookHwndCache.CalendarHwnd := hwnd
        else
            OutlookHwndCache.CalendarHwnd := 0
        return hwnd
    }

    static GetReminderHwnd() {
        if OUTLOOK_USE_HWND_CACHE && OutlookHwndCache._Valid(OutlookHwndCache.ReminderHwnd)
            return OutlookHwndCache.ReminderHwnd
        hwnd := OutlookHwndCache._ResolveReminder()
        if (hwnd > 0)
            OutlookHwndCache.ReminderHwnd := hwnd
        else
            OutlookHwndCache.ReminderHwnd := 0
        return hwnd
    }

    static _ResolveMailbox() {
        for _, exe in [OUTLOOK_EXE, OUTLOOK_OLK_EXE] {
            for hwnd in WinGetList("ahk_exe " exe) {
                try {
                    if !OutlookIsMainFrameClass(WinGetClass("ahk_id " hwnd))
                        continue
                    title := WinGetTitle("ahk_id " hwnd)
                    if OutlookTitleIsMailModule(title)
                        return (hwnd is Integer) && (hwnd > 0) ? hwnd : 0
                } catch {
                }
            }
        }
        return 0
    }

    static _ResolveCalendar() {
        for _, exe in [OUTLOOK_EXE, OUTLOOK_OLK_EXE] {
            for hwnd in WinGetList("ahk_exe " exe) {
                try {
                    if !OutlookIsMainFrameClass(WinGetClass("ahk_id " hwnd))
                        continue
                    title := WinGetTitle("ahk_id " hwnd)
                    if OutlookTitleIsCalendarModule(title)
                        return (hwnd is Integer) && (hwnd > 0) ? hwnd : 0
                } catch {
                }
            }
        }
        return 0
    }

    static _ResolveReminder() {
        ; Classic: dialog #32770. New Outlook: top-level "Outlook Host" (e.g. "Reminders - … - Outlook").
        for _, exe in [OUTLOOK_EXE, OUTLOOK_OLK_EXE] {
            for hwnd in WinGetList("ahk_exe " exe) {
                try {
                    cls := WinGetClass("ahk_id " hwnd)
                    if (cls != "#32770" && cls != OUTLOOK_CLASS_HOST)
                        continue
                    title := WinGetTitle("ahk_id " hwnd)
                    if InStr(title, OUTLOOK_REMINDER_TITLE)
                        return (hwnd is Integer) && (hwnd > 0) ? hwnd : 0
                } catch {
                }
            }
        }
        return 0
    }

    static InvalidateMailbox() {
        OutlookHwndCache.MailboxHwnd := 0
    }
    static InvalidateCalendar() {
        OutlookHwndCache.CalendarHwnd := 0
    }
    static InvalidateReminder() {
        OutlookHwndCache.ReminderHwnd := 0
    }
}

; --- Optional WinEvent: invalidate Outlook HWND cache when cached window is destroyed ---
; Default off: enables hook-driven invalidation without scanning WinExist on every cache hit path.
; Enable after validating EVENT_OBJECT_DESTROY callback overhead on your machine.
global OUTLOOK_USE_WINEVENT_INVALIDATE := false
global OUTLOOK_WINEVENT_HOOK_HANDLE := 0
global OUTLOOK_WINEVENT_CB := 0

Outlook_WinEventDestroyProc(hWinEventHook, event, hwnd, idObject, idChild, idEventThread, dwmsEventTime) {
    if (event != 0x8001) ; EVENT_OBJECT_DESTROY
        return
    if (idObject != 0) ; OBJID_WINDOW — ignore non-window objects
        return
    if !hwnd
        return
    try {
        if (hwnd = OutlookHwndCache.MailboxHwnd)
            OutlookHwndCache.InvalidateMailbox()
        if (hwnd = OutlookHwndCache.CalendarHwnd)
            OutlookHwndCache.InvalidateCalendar()
        if (hwnd = OutlookHwndCache.ReminderHwnd)
            OutlookHwndCache.InvalidateReminder()
    } catch {
    }
}

Outlook_RegisterDestroyWinEvent() {
    global OUTLOOK_WINEVENT_HOOK_HANDLE, OUTLOOK_WINEVENT_CB
    if OUTLOOK_WINEVENT_HOOK_HANDLE
        return
    OUTLOOK_WINEVENT_CB := CallbackCreate(Outlook_WinEventDestroyProc, "F", 7)
    OUTLOOK_WINEVENT_HOOK_HANDLE := DllCall("user32\SetWinEventHook", "UInt", 0x8001, "UInt", 0x8001, "Ptr", 0,
        "Ptr", OUTLOOK_WINEVENT_CB, "UInt", 0, "UInt", 0, "UInt", 0, "Ptr")
}

Outlook_UnregisterDestroyWinEvent() {
    global OUTLOOK_WINEVENT_HOOK_HANDLE
    if OUTLOOK_WINEVENT_HOOK_HANDLE {
        DllCall("user32\UnhookWinEvent", "Ptr", OUTLOOK_WINEVENT_HOOK_HANDLE)
        OUTLOOK_WINEVENT_HOOK_HANDLE := 0
    }
}

Outlook_OnExitUnregisterWinEvent(*) {
    Outlook_UnregisterDestroyWinEvent()
}

if OUTLOOK_USE_WINEVENT_INVALIDATE
    Outlook_RegisterDestroyWinEvent()
OnExit Outlook_OnExitUnregisterWinEvent, -1

; =============================================================================
; Outlook window resolution (strict HWND contract: success = positive HWND, failure = 0)
; =============================================================================
ResolveOutlookMailboxHwnd() {
    if OUTLOOK_USE_HWND_CACHE
        return OutlookHwndCache.GetMailboxHwnd()
    return OutlookHwndCache._ResolveMailbox()
}

ResolveOutlookCalendarHwnd() {
    if OUTLOOK_USE_HWND_CACHE
        return OutlookHwndCache.GetCalendarHwnd()
    return OutlookHwndCache._ResolveCalendar()
}

ResolveOutlookReminderHwnd() {
    if OUTLOOK_USE_HWND_CACHE
        return OutlookHwndCache.GetReminderHwnd()
    return OutlookHwndCache._ResolveReminder()
}

; =============================================================================
; Outlook window activation helpers (return true if activated, false otherwise)
; =============================================================================
ActivateOutlookMailbox() {
    hwnd := ResolveOutlookMailboxHwnd()
    if (hwnd > 0) {
        try {
            WinActivate("ahk_id " hwnd)
            return true
        } catch {
            OutlookHwndCache.InvalidateMailbox()
            ShowCenteredOverlay_Utils("❌ Error: Target window not found.", OUTLOOK_BANNER_MS, BANNER_ACCENT_ERROR)
            return false
        }
    }
    return false
}

ActivateOutlookCalendar() {
    hwnd := ResolveOutlookCalendarHwnd()
    if (hwnd > 0) {
        try {
            WinActivate("ahk_id " hwnd)
            return true
        } catch {
            OutlookHwndCache.InvalidateCalendar()
            ShowCenteredOverlay_Utils("❌ Error: Target window not found.", OUTLOOK_BANNER_MS, BANNER_ACCENT_ERROR)
            return false
        }
    }
    return false
}

; Ensures Outlook mailbox is active; returns true if Outlook is foreground within timeout, false otherwise.
EnsureOutlookMailActive() {
    hwnd := ResolveOutlookMailboxHwnd()
    if (hwnd <= 0)
        return false
    try {
        WinActivate("ahk_id " hwnd)
    } catch {
        OutlookHwndCache.InvalidateMailbox()
        return false
    }
    if !OUTLOOK_USE_STATE_WAITS
        return true
    timeoutSec := OUTLOOK_VOICE_WAIT_MS / 1000.0
    return WinWaitActive("ahk_id " hwnd, "", timeoutSec)
}

; Shared fallback executor for Mail/Calendar hotkeys (preserves message and priority order).
ActivateOutlookWithFallback(primaryFn, secondaryFn, failureMsg) {
    if primaryFn()
        return
    if secondaryFn()
        return
    PromptActivateOutlookBanner(failureMsg)
}

GetOutlookMainModuleState() {
    try {
        if !OutlookWinActive()
            return ""
        root := UIA.ElementFromHandle(WinExist("A"))
        if !root
            return ""

        mailItem := root.FindFirst({ Name: "Mail", Type: 50000 })
        if !mailItem
            mailItem := root.FindFirst({ Name: "Mail", Type: "50007" })
        if !mailItem
            mailItem := root.FindFirst({ Name: "Mail", ClassName: "NetUIListViewItem" })

        calendarItem := root.FindFirst({ Name: "Calendar", Type: 50000 })
        if !calendarItem
            calendarItem := root.FindFirst({ Name: "Calendar", Type: "50007" })
        if !calendarItem
            calendarItem := root.FindFirst({ Name: "Calendar", ClassName: "NetUIListViewItem" })

        if mailItem {
            if mailItem.IsTogglePatternAvailable {
                if mailItem.TogglePattern.CurrentToggleState = UIA.ToggleState.On
                    return "mail"
            } else if mailItem.IsSelected
                return "mail"
        }
        if calendarItem {
            if calendarItem.IsTogglePatternAvailable {
                if calendarItem.TogglePattern.CurrentToggleState = UIA.ToggleState.On
                    return "calendar"
            } else if calendarItem.IsSelected
                return "calendar"
        }
    } catch {
    }
    return ""
}

ClickOutlookModuleNavItem(targetModule) {
    try {
        if !OutlookWinActive()
            return false
        root := UIA.ElementFromHandle(WinExist("A"))
        if !root
            return false

        targetName := (targetModule = "mail") ? "Mail" : "Calendar"
        targetItem := root.FindFirst({ Name: targetName, Type: 50000 })
        if !targetItem
            targetItem := root.FindFirst({ Name: targetName, Type: "50007" })
        if !targetItem
            targetItem := root.FindFirst({ Name: targetName, ClassName: "NetUIListViewItem" })
        if !targetItem
            return false

        targetItem.SetFocus()
        Sleep 50
        try {
            targetItem.Click()
        } catch {
            try targetItem.Invoke()
            catch {
                return false
            }
        }
        return true
    } catch {
        return false
    }
}

IsOutlookTitleMatchingModule(targetModule) {
    try {
        if !OutlookWinActive()
            return false
        title := WinGetTitle("A")
        if (targetModule = "calendar")
            return OutlookTitleIsCalendarModule(title)
        if (targetModule = "mail")
            return OutlookTitleIsMailModule(title)
    } catch {
    }
    return false
}

VerifyOutlookModuleReached(targetModule, timeoutMs := 700) {
    deadline := A_TickCount + timeoutMs
    loop {
        state := GetOutlookMainModuleState()
        if (state = targetModule) {
            return true
        }
        if (state = "" && IsOutlookTitleMatchingModule(targetModule)) {
            return true
        }
        if (A_TickCount >= deadline)
            break
        Sleep 50
    }
    return false
}

EnsureOutlookMainModule(targetModule) {
    currentModule := GetOutlookMainModuleState()
    if (currentModule = targetModule)
        return true

    ; Reuse Shift+M semantics without key injection: click the target nav item directly.
    if ((currentModule = "mail" && targetModule = "calendar") || (currentModule = "calendar" && targetModule = "mail")) {
        switched := ClickOutlookModuleNavItem(targetModule)
        verified := VerifyOutlookModuleReached(targetModule, OUTLOOK_SWITCH_VERIFY_TIMEOUT_MS)
        return switched && verified
    }

    ; If selection state cannot be read, try direct UIA click on target module and verify.
    if ClickOutlookModuleNavItem(targetModule) {
        afterClickVerified := VerifyOutlookModuleReached(targetModule, OUTLOOK_SWITCH_VERIFY_TIMEOUT_MS)
        return afterClickVerified
    }
    return false
}

NavigateOutlookToModule(targetModule, failureMsg) {
    targetLabel := (targetModule = "mail") ? "Mailbox" : "Calendar"
    StandardLoadingBar_Show("⏳ Outlook: opening " targetLabel "...", BANNER_ACCENT_INTERMEDIATE, { passive: false,
        centerOnHwnd: 0, textWidth: 460, fontSize: 17 })

    ; If Outlook is already the foreground app (including Copilot surface), we can switch modules directly.
    ; This avoids false "activation failed" when the active Outlook title starts with "Copilot - ...".
    try {
        if OutlookWinActive() {
            activeHwnd := WinGetID("A")
            if (activeHwnd > 0) && OutlookIsMainFrameClass(WinGetClass("ahk_id " activeHwnd)) {
                WinActivate("ahk_id " activeHwnd)
                StandardLoadingBar_Update("🔄 Outlook: switching to " targetLabel "...", BANNER_ACCENT_INTERMEDIATE)
                if EnsureOutlookMainModule(targetModule) {
                    StandardLoadingBar_Update("✅ Outlook: " targetLabel " ready", BANNER_ACCENT_SUCCESS)
                    StandardLoadingBar_Hide(220)
                    return
                }
            }
        }
    } catch {
    }

    mailboxActivated := ActivateOutlookMailbox()
    calendarActivated := false
    if !mailboxActivated
        calendarActivated := ActivateOutlookCalendar()
    if !(mailboxActivated || calendarActivated) {
        StandardLoadingBar_Hide(0)
        PromptActivateOutlookBanner(failureMsg)
        return
    }

    StandardLoadingBar_Update("🔄 Outlook: switching to " targetLabel "...", BANNER_ACCENT_INTERMEDIATE)
    switched := EnsureOutlookMainModule(targetModule)
    if switched {
        StandardLoadingBar_Update("✅ Outlook: " targetLabel " ready", BANNER_ACCENT_SUCCESS)
        StandardLoadingBar_Hide(220)
        return
    }

    StandardLoadingBar_Hide(0)
    moduleLabel := (targetModule = "mail") ? "mailbox" : "calendar"
    ShowCenteredOverlay_Utils("❌ Outlook: Could not switch to " moduleLabel ".", OUTLOOK_BANNER_FAIL_MS,
        BANNER_ACCENT_ERROR)
}

GetOutlookLaunchPath() {
    outlookPath := ""
    if (IS_WORK_ENVIRONMENT) {
        outlookPath := "C:\Users\fie7ca\Documents\Atalhos\Microsoft Outlook.lnk"
        if (!FileExist(outlookPath))
            outlookPath := ""
    } else {
        outlookPath :=
            "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Microsoft Outlook.lnk"
        if (!FileExist(outlookPath))
            outlookPath := ""
    }
    return outlookPath
}

ActivateOrLaunchOutlook() {
    if OutlookProcessRunning() {
        try {
            for _, exe in [OUTLOOK_EXE, OUTLOOK_OLK_EXE] {
                if !ProcessExist(exe)
                    continue
                WinActivate("ahk_exe " exe)
                if WinWaitActive("ahk_exe " exe, "", 2) {
                    ShowCenteredOverlay_Utils("✅ Outlook activated.", OUTLOOK_BANNER_MS, BANNER_ACCENT_SUCCESS)
                    return
                }
            }
        } catch {
        }
    }

    try {
        outlookPath := GetOutlookLaunchPath()
        if (outlookPath != "")
            Run outlookPath
        else {
            olkPath := OutlookGetOlkExePath()
            if (olkPath != "")
                Run olkPath
            else
                Run OUTLOOK_EXE
        }
        ShowCenteredOverlay_Utils("⏳ Activating Outlook...", OUTLOOK_BANNER_MS, BANNER_ACCENT_INTERMEDIATE)
    } catch {
        ShowCenteredOverlay_Utils("❌ Failed to activate Outlook.", OUTLOOK_BANNER_FAIL_MS, BANNER_ACCENT_ERROR)
    }
}

PromptActivateOutlookBanner(failureMsg) {
    global g_OutlookActivatePromptPending
    if g_OutlookActivatePromptPending
        return
    g_OutlookActivatePromptPending := true

    message := failureMsg "`n❓ Would you like to activate Outlook?"
    keyCallbacks := Map(
        "Y", OutlookActivatePrompt_OnYes,
        "N", OutlookActivatePrompt_OnNo,
        "Escape", OutlookActivatePrompt_OnNo
    )
    StandardLoadingBar_ShowWithKeys(message, keyCallbacks, OUTLOOK_ACTIVATE_PROMPT_TIMEOUT_MS, 0,
        OutlookActivatePrompt_OnTimeout, "1E1E2E", 620, 17, "", true, "[Y] Activate Outlook  [N] Cancel")
}

OutlookActivatePrompt_OnYes(*) {
    global g_OutlookActivatePromptPending
    g_OutlookActivatePromptPending := false
    ActivateOrLaunchOutlook()
}

OutlookActivatePrompt_OnNo(*) {
    global g_OutlookActivatePromptPending
    g_OutlookActivatePromptPending := false
    ShowCenteredOverlay_Utils("ℹ Outlook activation cancelled.", OUTLOOK_BANNER_MS, BANNER_ACCENT_INTERMEDIATE)
}

OutlookActivatePrompt_OnTimeout(*) {
    global g_OutlookActivatePromptPending
    g_OutlookActivatePromptPending := false
    ShowCenteredOverlay_Utils("ℹ Outlook activation prompt timed out.", OUTLOOK_BANNER_MS, BANNER_ACCENT_INTERMEDIATE)
}

; =============================================================================
; Open Outlook Mail
; Hotkey: Win+Alt+Shift+B
; Original File: Outlook - Open mail.ahk
; =============================================================================
#!+b::
{
    NavigateOutlookToModule("mail",
        "⚠ Outlook: Mailbox and Calendar are not open (activation failed)")
}

; =============================================================================
; Open Outlook Calendar
; Hotkey: Win+Alt+Shift+G
; Original File: Outlook - Open calendar.ahk
; =============================================================================
#!+g::
{
    NavigateOutlookToModule("calendar",
        "⚠ Outlook: Calendar and Mailbox are not open (activation failed)")
}

; =============================================================================
; Open Outlook Reminders
; Hotkey: Win+Alt+Shift+V
; Original File: Outlook - Open Reminder.ahk
; =============================================================================
#!+v::
{
    ; Check if Outlook is closed and prompt to open if needed
    if (!CheckAndOpenOutlookTeams(true, false)) {
        return  ; User cancelled opening Outlook
    }

    hwnd := ResolveOutlookReminderHwnd()
    if (hwnd > 0) {
        try
            WinActivate("ahk_id " hwnd)
        catch {
            OutlookHwndCache.InvalidateReminder()
            ShowCenteredOverlay_Utils("❌ Error: Target window not found.", OUTLOOK_BANNER_MS, BANNER_ACCENT_ERROR)
        }
    } else
        ShowCenteredOverlay_Utils("❌ Error: Target window not found.", OUTLOOK_BANNER_MS, BANNER_ACCENT_ERROR)
}

; =============================================================================
; Voice Aloud Email
; Hotkey: Win+Alt+Shift+D
; =============================================================================
#!+d::
{
    try {
        ; Create GUI for voice aloud selection with auto-submit
        voiceAloudGui := Gui("+AlwaysOnTop +ToolWindow", "Voice Aloud Email")
        voiceAloudGui.SetFont("s10", "Segoe UI")

        ; Add instruction text
        voiceAloudGui.AddText("w350 Center",
            "Select voice aloud option:`n`n1) Voice aloud the email (from cursor)`n2) Voice aloud the email from the beginning`n`nType 1 or 2:"
        )

        ; Add input field with auto-submit functionality
        voiceAloudGui.AddEdit("w50 Center vVoiceAloudInput Limit1 Number")

        ; Add OK and Cancel buttons (as backup)
        voiceAloudGui.AddButton("w80 xp-40 y+10", "OK").OnEvent("Click", SubmitVoiceAloud)
        voiceAloudGui.AddButton("w80 xp+90", "Cancel").OnEvent("Click", CancelVoiceAloud)

        ; Set up auto-submit on text change
        voiceAloudGui["VoiceAloudInput"].OnEvent("Change", AutoSubmitVoiceAloud)

        ; Show GUI and focus input
        voiceAloudGui.Show("w350 h200")
        voiceAloudGui["VoiceAloudInput"].Focus()

    } catch Error as e {
        MsgBox "Error in voice aloud selector: " e.Message, "Voice Aloud Error", "IconX"
    }
}

; Single validation-and-execute helper for Voice Aloud (used by AutoSubmit and Submit).
TryExecuteVoiceChoice(value, gui, showInvalidMsg := false) {
    if (value = "" || !IsInteger(value))
        return
    choice := Integer(value)
    if (choice >= VOICE_ALOUD_OPTION_MIN && choice <= VOICE_ALOUD_OPTION_MAX) {
        gui.Destroy()
        ExecuteVoiceAloudOption(GetVoiceAloudOptionByNumber(value))
    } else if showInvalidMsg
        MsgBox "Invalid selection. Please choose " VOICE_ALOUD_OPTION_MIN "-" VOICE_ALOUD_OPTION_MAX ".",
            "Voice Aloud Selection", "IconX"
}

; Function to get voice aloud option by number
GetVoiceAloudOptionByNumber(numberText) {
    try number := Integer(numberText)
    catch {
        return ""
    }
    optionMap := Map()
    optionMap[1] := "from_cursor"
    optionMap[2] := "from_beginning"
    if (number >= VOICE_ALOUD_OPTION_MIN && number <= VOICE_ALOUD_OPTION_MAX)
        return optionMap[number]
    return ""
}

; Function to execute voice aloud option (state-driven: no blind key injection)
ExecuteVoiceAloudOption(option) {
    if (option = "")
        return

    Send "{Media_Stop}"
    if OUTLOOK_USE_STATE_WAITS
        Sleep(80)
    else
        Sleep(200)

    ; Activate mailbox and wait for foreground with timeout; no synthetic hotkey.
    if (!EnsureOutlookMailActive()) {
        ShowCenteredOverlay_Utils("⚠ Outlook not active; aborting.", OUTLOOK_BANNER_MS, BANNER_ACCENT_ERROR)
        return
    }

    if OUTLOOK_USE_STATE_WAITS
        Sleep(80)
    else
        Sleep(200)

    hwnd := WinGetID("A")
    if (option = "from_cursor") {
        InvokeReadAloudStart(hwnd)
        if OUTLOOK_USE_STATE_WAITS
            Sleep(80)
        else
            Sleep(200)
        SendEscape()
    } else if (option = "from_beginning") {
        Send "{Right}"
        if OUTLOOK_USE_STATE_WAITS
            Sleep(80)
        else
            Sleep(300)
        Send "^{Home}"
        if OUTLOOK_USE_STATE_WAITS
            Sleep(80)
        else
            Sleep(200)
        InvokeReadAloudStart(hwnd)
        if OUTLOOK_USE_STATE_WAITS
            Sleep(80)
        else
            Sleep(200)
        SendEscape()
    }
}

; Auto-submit function for voice aloud
AutoSubmitVoiceAloud(ctrl, *) {
    TryExecuteVoiceChoice(ctrl.Text, ctrl.Gui, false)
}

; Manual submit function for voice aloud (backup)
SubmitVoiceAloud(ctrl, *) {
    TryExecuteVoiceChoice(ctrl.Gui["VoiceAloudInput"].Text, ctrl.Gui, true)
}

; Cancel function for voice aloud
CancelVoiceAloud(ctrl, *) {
    ctrl.Gui.Destroy()
}
