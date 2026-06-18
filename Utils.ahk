#Requires AutoHotkey v2.0+
#SingleInstance Force
#include %A_ScriptDir%\env.ahk
#include %A_ScriptDir%\StudyLinkHelpers.ahk

global g_StudyLinkSubmenuGui := ""

#include UIA-v2\Lib\UIA.ahk
#include UIA-v2\Lib\UIA_Browser.ahk
#include %A_ScriptDir%\lib\Media.ahk
#include %A_ScriptDir%\SpotifyWASAPI.ahk

; UIA ControlType constants (Button=50000). Shared with Gemini.ahk focus helpers.
global UIA_ControlType_Button := 50000

; -----------------------------------------------------------------------------
; Timing and window utilities for performance measurement and lightweight checks
; -----------------------------------------------------------------------------
; StartTimer() -> returns a timer object { start: A_TickCount }
StartTimer() {
    return { start: A_TickCount }
}

; EndTimer(timer, message := "") -> returns elapsed ms; optional message is appended to chrome_detach.log
EndTimer(timer, message := "") {
    elapsed := A_TickCount - timer.start
    if (message != "") {
        try {
            FileAppend(message . ": " . elapsed . " ms`n", A_ScriptDir . "\chrome_detach.log")
        } catch {
            ; swallow logging errors
        }
    }
    return elapsed
}

; WindowExists(windowTitle) - Checks if window exists without waiting
; Returns: True if window exists, false otherwise
WindowExists(windowTitle) {
    prev := A_TitleMatchMode
    SetTitleMatchMode(2)
    exists := !!WinExist(windowTitle)
    SetTitleMatchMode(prev)
    return exists
}

; =============================================================================
; Semantic banner accents (must be defined early)
; Some startup/update helpers call ShowCenteredOverlay_Utils / StandardLoadingBar_Show
; before the later globals block is reached.
; =============================================================================
global BANNER_ACCENT_SUCCESS := "27AE60"      ; Dark green: positive / success
global BANNER_ACCENT_ERROR := "C0392B"        ; Red: negative / error
global BANNER_ACCENT_INTERMEDIATE := "F1C40F" ; Yellow: loading, actionable, neutral
global BANNER_ACCENT_INFO := "2980B9"         ; Blue: info / alternate mode

; Possible Gemini prompt field names (EN and PT) for work/personal env. Used by FindGeminiPromptField.
global GEMINI_PROMPT_FIELD_NAMES := ["Enter a prompt for Gemini", "Enter a prompt here",
    "Digite um prompt para o Gemini", "Digite um prompt aqui"]
global GEMINI_PROMPT_FOCUS_POLL_MS := 25
global GEMINI_PROMPT_FOCUS_TIMEOUT_MS := 300
global GEMINI_OPEN_FAST_SETTLE_MS := 0

; Find the Gemini prompt field via UIA (returns element or 0). Supports EN and PT labels. Used by Gemini.ahk and Utils.ahk.
; Happy path: FindFirst per name only. FindAll({ Type: 50004 }) runs only when those fail (failure path).
FindGeminiPromptField(uia) {
    promptField := 0
    for name in GEMINI_PROMPT_FIELD_NAMES {
        try {
            promptField := uia.FindFirst({ Name: name, Type: 50004 })
            if (promptField) {
                return promptField
            }
        } catch
            continue
    }
    try {
        promptField := uia.FindFirst({ Type: "Edit", Name: GEMINI_PROMPT_FIELD_NAMES[1] })
        if (promptField)
            return promptField
    } catch {
    }
    try {
        allEdits := uia.FindAll({ Type: 50004 })
        for edit in allEdits {
            if (InStr(edit.ClassName, "ql-editor") || InStr(edit.ClassName, "new-input-ui")) {
                for name in GEMINI_PROMPT_FIELD_NAMES {
                    if InStr(edit.Name, name) || InStr(edit.Name, "prompt") {
                        return edit
                    }
                }
            }
        }
    } catch {
    }
    try {
        allEdits := uia.FindAll({ Type: 50004 })
        for edit in allEdits {
            if InStr(edit.ClassName, "ql-editor") {
                return edit
            }
        }
    } catch {
    }
    try {
        allEdits := uia.FindAll({ Type: 50004 })
        for edit in allEdits {
            if InStr(edit.ClassName, "ql-editor") {
                for name in GEMINI_PROMPT_FIELD_NAMES {
                    if InStr(edit.Name, name) || InStr(edit.Name, "prompt")
                        return edit
                }
            }
        }
    } catch {
    }
    return 0
}

; Move keyboard focus to the main Ask Gemini field when that Chrome window is already foreground.
; Does not WinActivate (avoids stealing focus from another app). Optional chime matches Gemini.ahk open-hotkey UX.
Utils_PlayGeminiFocusedChime(minIntervalMs := 400) {
    static lastChimeTick := 0
    if (!IsSoundEnabled())
        return false
    now := A_TickCount
    if (lastChimeTick && (now - lastChimeTick) < minIntervalMs)
        return false
    lastChimeTick := now
    ScriptSoundPlay(A_ScriptDir . "\sounds\gemini-focused.wav")
    return true
}

Gemini_PollPromptKeyboardFocus(promptField, timeoutMs := 0, pollMs := 0) {
    if (timeoutMs <= 0)
        timeoutMs := GEMINI_PROMPT_FOCUS_TIMEOUT_MS
    if (pollMs <= 0)
        pollMs := GEMINI_PROMPT_FOCUS_POLL_MS
    start := A_TickCount
    while (A_TickCount - start < timeoutMs) {
        try {
            if (promptField.HasKeyboardFocus)
                return true
        } catch {
        }
        Sleep pollMs
    }
    try
        return promptField.HasKeyboardFocus
    catch
        return false
}

Gemini_FindUploadAnchorButton(uia) {
    anchorButton := 0
    try {
        anchorButton := uia.FindFirst({ Type: UIA_ControlType_Button, Name: "Open upload file menu", ControlType: "Button" })
        if (!anchorButton)
            anchorButton := uia.FindFirst({ Type: UIA_ControlType_Button, Name: "Open upload file menu", cs: false })
    } catch {
    }
    if (!anchorButton) {
        try {
            allButtons := uia.FindAll({ Type: UIA_ControlType_Button })
            for button in allButtons {
                try {
                    if (InStr(button.Name, "Open upload file menu", false)) {
                        anchorButton := button
                        break
                    }
                } catch {
                    continue
                }
            }
        } catch {
        }
    }
    return anchorButton
}

Gemini_ApplyAnchorBacktrack(anchorButton) {
    try {
        anchorButton.SetFocus()
        Sleep GEMINI_PROMPT_FOCUS_POLL_MS
        SendInput "+{Tab}"
        Sleep GEMINI_PROMPT_FOCUS_POLL_MS
        return true
    } catch {
        return false
    }
}

; Focus Gemini prompt: direct FindFirst → bounded poll → anchor Shift+Tab fallback. Optional ready chime.
Gemini_FocusPromptWithChime(uia, options := "", &outPhase := "") {
    outPhase := "focus_failed"
    if (!IsObject(uia))
        return false

    playChime := true
    useAnchorFallback := true
    if (IsObject(options)) {
        if (options.HasProp("playChime"))
            playChime := options.playChime
        if (options.HasProp("useAnchorFallback"))
            useAnchorFallback := options.useAnchorFallback
    }

    promptField := 0
    try
        promptField := FindGeminiPromptField(uia)
    catch {
    }
    if (!promptField)
        return false

    try {
        if (promptField.HasKeyboardFocus) {
            outPhase := "fast_already_focused"
            if (playChime)
                Utils_PlayGeminiFocusedChime()
            return promptField
        }
    } catch {
    }

    try
        promptField.SetFocus()
    catch {
    }
    if (Gemini_PollPromptKeyboardFocus(promptField)) {
        outPhase := "direct_focus"
        if (playChime)
            Utils_PlayGeminiFocusedChime()
        return promptField
    }

    try
        promptField.Click()
    catch {
    }
    if (Gemini_PollPromptKeyboardFocus(promptField, GEMINI_PROMPT_FOCUS_TIMEOUT_MS // 2)) {
        outPhase := "direct_focus"
        if (playChime)
            Utils_PlayGeminiFocusedChime()
        return promptField
    }

    if (useAnchorFallback) {
        anchorButton := Gemini_FindUploadAnchorButton(uia)
        if (anchorButton && Gemini_ApplyAnchorBacktrack(anchorButton)) {
            try
                promptField := FindGeminiPromptField(uia)
            catch {
            }
            if (promptField) {
                try
                    promptField.SetFocus()
                catch {
                }
                if (!Gemini_PollPromptKeyboardFocus(promptField)) {
                    try
                        promptField.Click()
                    catch {
                    }
                    Gemini_PollPromptKeyboardFocus(promptField, GEMINI_PROMPT_FOCUS_TIMEOUT_MS // 2)
                }
                try {
                    if (promptField.HasKeyboardFocus) {
                        outPhase := "anchor_fallback"
                        if (playChime)
                            Utils_PlayGeminiFocusedChime()
                        return promptField
                    }
                } catch {
                }
            }
        }
    }

    return false
}

; Shared by Gemini.ahk #!+i, Shift keys Fast Copy, and async flows. Resolves UIA when omitted.
Gemini_FocusPromptSameAsOpenHotkey(uia, playChime := true) {
    if (!IsObject(uia)) {
        try
            uia := UIA_Browser()
        catch
            return false
    }
    return Gemini_FocusPromptWithChime(uia, { playChime: playChime, useAnchorFallback: true })
}

Gemini_WaitForPromptFieldDiscoverable(uia, timeoutMs := 0) {
    if (timeoutMs <= 0)
        timeoutMs := GEMINI_PROMPT_FOCUS_TIMEOUT_MS + 200
    start := A_TickCount
    while (A_TickCount - start < timeoutMs) {
        try {
            if (FindGeminiPromptField(uia))
                return true
        } catch {
        }
        Sleep GEMINI_PROMPT_FOCUS_POLL_MS
    }
    return false
}

Gemini_ShowDeferredTabBanner(uia) {
    try {
        tabInfo := GetChromeActiveTabIndex(uia)
        if (!tabInfo) {
            Sleep 80
            tabInfo := GetChromeActiveTabIndex(uia)
        }
        if (tabInfo && tabInfo.count >= 2 && tabInfo.index)
            ShowSingleCharTabBanner_Utils(tabInfo.index)
    } catch {
    }
}

FocusGeminiAskFieldForHwnd(geminiHwnd, playChime := false) {
    if (!geminiHwnd)
        return false
    try {
        if (!WinActive("ahk_id " geminiHwnd))
            return false
    } catch {
        return false
    }
    try {
        uia := UIA_Browser("ahk_id " geminiHwnd)
        result := Gemini_FocusPromptWithChime(uia, { playChime: playChime, useAnchorFallback: true })
        return !!result
    } catch {
    }
    return false
}

; 1-based active Chrome tab index via UIA (tab bar). Returns { index, count } or 0. Shared with Gemini.ahk.
GetChromeActiveTabIndex(uia) {
    try {
        uia.GetCurrentMainPaneElement()
        tabs := uia.GetTabs()
        if (!tabs.Length)
            return 0
        current := uia.GetTab("")
        if (!current)
            return 0
        rid := current.RuntimeId
        for i, tab in tabs {
            try {
                if (tab.RuntimeId = rid)
                    return { index: i, count: tabs.Length }
            } catch {
                continue
            }
        }
    } catch {
    }
    return 0
}

; =============================================================================
; F11 fullscreen detection / toggle (shared by WindowManagement + Chrome detach)
; =============================================================================

_WMF11_GetWindowRectHwnd(hwnd, &left, &top, &right, &bottom) {
    rect := Buffer(16, 0)
    if !DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect)
        return false
    left := NumGet(rect, 0, "int")
    top := NumGet(rect, 4, "int")
    right := NumGet(rect, 8, "int")
    bottom := NumGet(rect, 12, "int")
    return true
}

_WMF11_GetHwndMonitorIndex(hwnd) {
    if (!hwnd)
        return 0
    try {
        hMon := DllCall("MonitorFromWindow", "ptr", hwnd, "uint", 2, "ptr")
        loop MonitorGetCount() {
            MonitorGet A_Index, &ml, &mt, &mr, &mb
            cx := (ml + mr) // 2
            cy := (mt + mb) // 2
            point64 := (cy & 0xFFFFFFFF) << 32 | (cx & 0xFFFFFFFF)
            if (Integer(DllCall("MonitorFromPoint", "int64", point64, "uint", 2, "ptr")) = Integer(hMon))
                return A_Index
        }
    } catch {
    }
    return 0
}

_WMF11_IsDesktopOrTaskbarClass(cls) {
    return cls = "Progman" || cls = "WorkerW" || cls = "Shell_TrayWnd" || cls = "Shell_SecondaryTrayWnd"
}

_WMF11_IsExcludedIndicatorWindow(hwnd) {
    if (!hwnd)
        return false
    try {
        exe := StrLower(WinGetProcessName("ahk_id " hwnd))
    } catch {
        return false
    }
    if (exe = "handy.exe")
        return true
    try {
        title := WinGetTitle(hwnd)
    } catch {
        return false
    }
    if (InStr(StrLower(title), "windowmanagement.ahk"))
        return true
    return false
}

_WMF11_BackgroundHwndOnAnyScriptMonitor(hwnd) {
    try {
        hMon := DllCall("MonitorFromWindow", "ptr", hwnd, "uint", 2, "ptr")
        loop MonitorGetCount() {
            MonitorGet A_Index, &ml, &mt, &mr, &mb
            cx := (ml + mr) // 2
            cy := (mt + mb) // 2
            point64 := (cy & 0xFFFFFFFF) << 32 | (cx & 0xFFFFFFFF)
            if (Integer(DllCall("MonitorFromPoint", "int64", point64, "uint", 2, "ptr")) = Integer(hMon))
                return true
        }
    } catch {
    }
    return false
}

; F11 fullscreen: fills monitor (often past work area) or work area with no caption; not ordinary Win-maximize.
WM_WindowIsF11FullscreenRejectReason(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return "no_hwnd"
    try {
        minMax := WinGetMinMax(hwnd)
        if (minMax = -1)
            return "minimized"
        if !DllCall("IsWindowVisible", "ptr", hwnd)
            return "not_visible"
        exStyle := DllCall("GetWindowLongPtr", "ptr", hwnd, "int", -20, "ptr")
        if (exStyle & 0x00000080)
            return "toolwindow"
        class := WinGetClass(hwnd)
        if (class = "Progman" || class = "WorkerW")
            return "desktop_class"
        if (_WMF11_IsDesktopOrTaskbarClass(class))
            return "taskbar_class"
        if (WinGetTitle(hwnd) = "")
            return "empty_title"
        if (_WMF11_IsExcludedIndicatorWindow(hwnd))
            return "excluded_indicator"
        if (!_WMF11_BackgroundHwndOnAnyScriptMonitor(hwnd))
            return "not_script_monitor"
        mon := _WMF11_GetHwndMonitorIndex(hwnd)
        if (!mon)
            return "no_monitor"
        MonitorGet mon, &ml, &mt, &mr, &mb
        MonitorGetWorkArea mon, &wl, &wt, &wr, &wb
        if (!_WMF11_GetWindowRectHwnd(hwnd, &left, &top, &right, &bottom))
            return "no_rect"
        tol := 24
        fillsMonitor := (Abs(left - ml) <= tol && Abs(top - mt) <= tol && Abs(right - mr) <= tol && Abs(bottom - mb) <=
            tol)
        fillsWorkArea := (Abs(left - wl) <= tol && Abs(top - wt) <= tol && Abs(right - wr) <= tol && Abs(bottom - wb) <=
            tol)
        extendsPastWorkArea := (bottom > wb + tol || right > wr + tol || left < wl - tol || top < wt - tol)
        style := DllCall("GetWindowLongPtr", "ptr", hwnd, "int", -16, "ptr")
        hasCaption := !!(style & 0x00C00000)
        if (minMax = 1 && fillsWorkArea && !extendsPastWorkArea && hasCaption)
            return "win_maximized"
        if (fillsMonitor && extendsPastWorkArea)
            return ""
        if (fillsWorkArea && !hasCaption)
            return ""
        if (!fillsMonitor && !fillsWorkArea)
            return "not_monitor_fill"
        if (fillsMonitor && !extendsPastWorkArea)
            return "within_work_area"
        if (fillsWorkArea && hasCaption)
            return "work_area_with_caption"
        return "no_match"
    } catch as err {
        return "exception:" . err.Message
    }
}

WM_WindowIsF11Fullscreen(hwnd) {
    return WM_WindowIsF11FullscreenRejectReason(hwnd) = ""
}

WM_WaitForF11State(hwnd, wantFullscreen, timeoutMs := 500) {
    deadline := A_TickCount + timeoutMs
    loop {
        if (WM_WindowIsF11Fullscreen(hwnd) = wantFullscreen)
            return true
        if (A_TickCount >= deadline)
            break
        Sleep 50
    }
    return WM_WindowIsF11Fullscreen(hwnd) = wantFullscreen
}

WM_EnsureForegroundForSend(hwnd, timeoutMs := 2000) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        try WinActivate "ahk_id " hwnd
        catch {
            return false
        }
        if WinActive("ahk_id " hwnd)
            return true
        Sleep 50
    }
    return WinActive("ahk_id " hwnd)
}

WM_ExitF11FullscreenForHwnd(hwnd, settleMs := 1200) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    if (!WM_WindowIsF11Fullscreen(hwnd))
        return true
    try {
        if !WM_EnsureForegroundForSend(hwnd)
            return false
        Sleep 80
        ClipAngel_ReleaseChordModifiersForSend()
        SendInput "{F11}"
        return WM_WaitForF11State(hwnd, false, settleMs)
    } catch {
        return false
    }
}

WM_EnterF11FullscreenForHwnd(hwnd, settleMs := 1200) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    if (WM_WindowIsF11Fullscreen(hwnd))
        return true
    try {
        if !WM_EnsureForegroundForSend(hwnd)
            return false
        Sleep 80
        ClipAngel_ReleaseChordModifiersForSend()
        SendInput "{F11}"
        return WM_WaitForF11State(hwnd, true, settleMs)
    } catch {
        return false
    }
}

; --- Chrome: detach active tab to new window (Shift+W) ---
; Primary UIA menu path (F6 + tab context menu). Optional: MV3 PopActiveTab (Ctrl+Shift+Y)
; in chrome\PopActiveTab — set CHROME_DETACH_USE_EXTENSION := true when loaded in Chrome.
; Success gates: top-level HWND + title match + SetWinEventHook (not fixed sleeps).
global CHROME_DETACH_USE_EXTENSION := false
global CHROME_DETACH_USE_WIN_EVENT_HOOK := true
global CHROME_DETACH_USE_LIGHT_NORMALIZE := true
global CHROME_DETACH_PREFLIGHT_TAB_COUNT := false
global CHROME_DETACH_EVENT_OBJECT_CREATE := 0x8000
global CHROME_DETACH_OBJID_WINDOW := 0
global g_ChromeDetachCreatedHwnds := []
global CHROME_DETACH_USE_UIA := false
global CHROME_DETACH_USE_UIA_SINGLE_SHOT := true
global CHROME_DETACH_ALLOW_UIA_DEBUG_FALLBACK := false
global CHROME_DETACH_LEGACY_KEYS := false
global CHROME_DETACH_DEBUG_LOG_ENABLED := false
global CHROME_DETACH_PERF_LOG_ENABLED := false
global CHROME_DETACH_DEEP_FALLBACK := false
global CHROME_DETACH_F6_FALLBACK := true
global CHROME_DETACH_EXTENSION_TIMEOUT_MS := 1400
global CHROME_DETACH_TOTAL_TIMEOUT_MS := 4500
global CHROME_DETACH_VERIFY_TIMEOUT_MS := 800
global CHROME_DETACH_EXTENSION_VERIFY_MS := 1400
global CHROME_DETACH_F11_SETTLE_MS := 1500
global CHROME_DETACH_MENU_POPUP_MS := 280          ; ↓ from 350 — Chrome renders menu quickly; UIA detection reliable
global CHROME_DETACH_MENU_CHILD_MS := 40            ; ↓ from 50
global CHROME_DETACH_MENU_DISMISS_MS := 30          ; ↓ from 50
global CHROME_DETACH_A11Y_MS := 40                  ; ↓ from 50
global CHROME_DETACH_A11Y_MS_FOREGROUND := 40       ; ↓ from 50
global CHROME_DETACH_MENU_POLL_MS := 10             ; ↓ from 20 — faster polling loop response
global CHROME_DETACH_SUCCESS_HIDE_MS := 400
global CHROME_DETACH_SEQUENCE_ATTEMPTS := 1
global CHROME_DETACH_HOVER_ATTEMPTS := 2
global CHROME_DETACH_F6_FOCUS_MAX := 3
global CHROME_DETACH_F6_STEP_MS := 80               ; ↓ from 160 — F6 is instant in Chrome, half the delay is safe
global CHROME_DETACH_F6_FOCUS_POLL_MS := 120        ; ↓ from 200 — poll faster for UIA focus confirmation
global CHROME_DETACH_F6_REFOCUS_MS := 50            ; ↓ from 80
global CHROME_DETACH_TAB_FOCUS_WAIT_MS := 300       ; ↓ from 500 — poll is fast, no need for long initial wait
global CHROME_DETACH_TAB_FOCUS_STABLE_MS := 80      ; ↓ from 120
global CHROME_DETACH_HOVER_SETTLE_MS := 80          ; ↓ from 120
global CHROME_DETACH_APPSKEY_AFTER_MS := 40         ; ↓ from 50
; Hover settle before AppsKey — Chrome hit-tests cursor; too low opens page menu.
global CHROME_DETACH_HOVER_APPSKEY_SETTLE_MS := 340 ; ↓ from 420 — reduced settle; banner appears within ~200ms
global CHROME_DETACH_HOVER_APPSKEY_RETRY_EXTRA_MS := 100 ; ↓ from 140
global CHROME_DETACH_APPSKEY_SETTLE_MS := 60        ; ↓ from 80
global g_ChromeDetachBusy := false
global g_ChromeDetachDebugLogPath := ""

global CHROME_DETACH_MENU_PARENT_NAMES := ["Mover guia para outra janela", "Mover guia para uma nova janela",
    "Move tab to another window"]
global CHROME_DETACH_MENU_PARENT_SUBSTR := ["Mover guia", "Move tab to another", "Move tab to a new"]
global CHROME_DETACH_MENU_CHILD_NAMES := ["Nova janela", "New window"]
global CHROME_DETACH_MENU_EN_NAMES := ["Move tab to new window", "Mover guia para nova janela"]
global CHROME_DETACH_MENU_TAB_MARKER_NAMES := ["Nova guia", "New tab"]
global CHROME_DETACH_MENU_TAB_MARKER_SUBSTR := ["Mover guia", "Move tab", "Fechar guia", "Close tab", "Duplicar guia",
    "Duplicate"]
global CHROME_DETACH_MENU_PAGE_MARKER_NAMES := ["Voltar", "Back", "Avançar", "Forward"]
global CHROME_DETACH_MENU_PAGE_MARKER_SUBSTR := ["Salvar como", "Save as", "Imprimir", "Print", "Ver código",
    "View page source", "Inspecionar", "Inspect"]

; Shift+W hotkey: wait for physical Shift release before menu mnemonics / synthetic keys.
Chrome_Detach_PreSendSanitizeModifiers(timeoutMs := 220) {
    tw := "T" (timeoutMs / 1000)
    KeyWait "LShift", tw
    KeyWait "RShift", tw
    ClipAngel_ReleaseChordModifiersForSend()
    deadline := A_TickCount + 120
    while (A_TickCount < deadline) {
        if !GetKeyState("Shift", "P") && !GetKeyState("Ctrl", "P") && !GetKeyState("Alt", "P")
            break
        Sleep 10
    }
}

Chrome_DetachGetDebugLogPath() {
    global g_ChromeDetachDebugLogPath
    if !g_ChromeDetachDebugLogPath
        g_ChromeDetachDebugLogPath := RegExReplace(A_LineFile, "i)\\[^\\]+$", "") . "\debug-79854f.log"
    return g_ChromeDetachDebugLogPath
}

Chrome_DetachDebugFocusedElement(session) {
    info := "none"
    try {
        el := UIA.GetFocusedElement()
        if el {
            n := "", t := ""
            try n := el.Name
            try t := el.Type
            info := "type=" t " name=" SubStr(n, 1, 32)
        }
    } catch {
    }
    return info
}

Chrome_DetachDebugLog(location, message, hypothesisId := "", data := unset) {
    global CHROME_DETACH_DEBUG_LOG_ENABLED
    if !CHROME_DETACH_DEBUG_LOG_ENABLED
        return
    try {
        extra := ""
        if IsSet(data) {
            if (data is String)
                extra := data
            else if IsObject(data) {
                for k, v in data
                    extra .= (extra = "" ? "" : ";") . k . "=" . v
            }
        }
        line := A_TickCount . "|" . hypothesisId . "|" . location . "|" . message . "|" . extra . "`n"
        for logPath in [A_Temp . "\debug-79854f.log", Chrome_DetachGetDebugLogPath()] {
            try {
                FileAppend line, logPath, "UTF-8"
                return
            } catch {
            }
        }
        OutputDebug line
    } catch {
    }
}

Chrome_DetachSampleMenuItemsForHwnd(hwnd, maxItems := 6) {
    sample := ""
    if !hwnd
        return sample
    try {
        root := UIA.ElementFromHandle(hwnd)
        n := 0
        for el in root.FindAll({ Type: UIA.Type.MenuItem }, UIA.TreeScope.Subtree) {
            try {
                name := el.Name
            } catch {
                continue
            }
            if (name = "")
                continue
            sample .= (sample = "" ? "" : "|") . SubStr(name, 1, 36)
            if (++n >= maxItems)
                break
        }
    } catch {
    }
    return sample
}

Chrome_DetachDebugSampleMenuItems(session, maxItems := 6) {
    return Chrome_DetachSampleMenuItemsForHwnd(session.menuPopupHwnd, maxItems)
}

Chrome_DetachWindowLooksLikeContextPopup(hwnd) {
    try {
        WinGetPos(, , &w, &h, "ahk_id " hwnd)
        if !(w > 40 && w < 750 && h > 40 && h < 950)
            return false
        class := WinGetClass("ahk_id " hwnd)
        if (class = "#32768")
            return true
        if RegExMatch(class, "i)^Chrome_WidgetWin")
            return true
    } catch {
    }
    return false
}

Chrome_DetachListMenuPopups(chromeHwnd := 0) {
    popups := Map()
    try {
        for h in WinGetList("ahk_class #32768")
            popups[h] := true
    } catch {
    }
    if chromeHwnd {
        try {
            for h in WinGetList("ahk_exe chrome.exe") {
                if (h = chromeHwnd)
                    continue
                if Chrome_DetachWindowLooksLikeContextPopup(h)
                    popups[h] := true
            }
        } catch {
        }
    }
    return popups
}

Chrome_DetachDebugListNewChromeWindows(session) {
    info := ""
    base := session.baselinePopups
    try {
        for h in WinGetList("ahk_exe chrome.exe") {
            if (h = session.hwnd)
                continue
            if (IsObject(base) && base.Has(h))
                continue
            try {
                WinGetPos(, , &w, &hgt, "ahk_id " h)
                info .= (info = "" ? "" : "|") . h . ":" . WinGetClass("ahk_id " h) . "@" . w . "x" . hgt
            } catch {
            }
        }
    } catch {
    }
    return info
}

Chrome_DetachSessionCreate(hwnd, existingSet := unset) {
    if !(existingSet is Map) {
        existingSet := Map()
        try {
            for h in WinGetList("ahk_exe chrome.exe")
                existingSet[h] := true
        } catch {
        }
    }
    uia := 0
    winTitle := "ahk_id " hwnd
    a11yMs := WinActive("ahk_id " hwnd) ? CHROME_DETACH_A11Y_MS_FOREGROUND : CHROME_DETACH_A11Y_MS
    try UIA.ActivateChromiumAccessibility(winTitle, a11yMs)
    catch {
    }
    try {
        uia := UIA_Browser(winTitle)
        uia.GetCurrentMainPaneElement()
    } catch {
    }
    return { hwnd: hwnd, uia: uia, existingSet: existingSet, menuPopupHwnd: 0, menuPopupClassify: "",
        activeTab: 0, newDetachedHwnd: 0, menuConfirmed: false, baselinePopups: Chrome_DetachListMenuPopups(hwnd) }
}

Chrome_SessionUiaUsable(session) {
    if !IsObject(session.uia)
        return false
    try {
        session.uia.GetCurrentMainPaneElement()
        return true
    } catch {
        return false
    }
}

Chrome_ContextMenuNameMatches(el, names, substrs := "") {
    try name := el.Name
    catch {
        return false
    }
    if (name = "")
        return false
    for candidate in names {
        if (name = candidate || InStr(name, candidate, false))
            return true
    }
    if (substrs) {
        for sub in substrs {
            if InStr(name, sub, false)
                return true
        }
    }
    return false
}

Chrome_ContextMenuFindInRoot(root, names, substrs := "") {
    if !IsObject(root)
        return 0
    for name in names {
        try {
            el := root.FindElement({ Type: UIA.Type.MenuItem, Name: name, matchMode: 2 }, UIA.TreeScope.Subtree)
            if el
                return el
        } catch {
        }
    }
    if (substrs) {
        try {
            for el in root.FindAll({ Type: UIA.Type.MenuItem }, UIA.TreeScope.Subtree) {
                if Chrome_ContextMenuNameMatches(el, [], substrs)
                    return el
            }
        } catch {
        }
    }
    return 0
}

Chrome_ContextMenuFindFirst(session, names, useParentSubstr := true) {
    substrs := useParentSubstr ? CHROME_DETACH_MENU_PARENT_SUBSTR : ""
    if (session.menuPopupHwnd) {
        try {
            el := Chrome_ContextMenuFindInRoot(UIA.ElementFromHandle(session.menuPopupHwnd), names, substrs)
            if el
                return el
        } catch {
        }
    }
    try {
        return Chrome_ContextMenuFindInRoot(UIA.ElementFromHandle(session.hwnd), names, substrs)
    } catch {
    }
    return 0
}

Chrome_ContextMenuWaitForSession(session, names, timeoutMs := 0, useParentSubstr := true) {
    waitMs := timeoutMs > 0 ? timeoutMs : CHROME_DETACH_MENU_CHILD_MS
    deadline := A_TickCount + waitMs
    while (A_TickCount < deadline) {
        el := Chrome_ContextMenuFindFirst(session, names, useParentSubstr)
        if el
            return el
        Sleep CHROME_DETACH_MENU_POLL_MS
    }
    return 0
}

Chrome_ContextMenuListNewPopupCandidates(session, baselinePopups := "") {
    seen := Map()
    list := []
    base := IsObject(baselinePopups) ? baselinePopups : session.baselinePopups
    try {
        for h in WinGetList("ahk_class #32768") {
            if (IsObject(base) && base.Has(h))
                continue
            if !seen.Has(h) {
                seen[h] := true
                list.Push(h)
            }
        }
        for h in WinGetList("ahk_exe chrome.exe") {
            if (h = session.hwnd)
                continue
            if (IsObject(base) && base.Has(h))
                continue
            if !Chrome_DetachWindowLooksLikeContextPopup(h)
                continue
            if !seen.Has(h) {
                seen[h] := true
                list.Push(h)
            }
        }
    } catch {
    }
    return list
}

Chrome_ContextMenuSampleLooksLikePageMenu(sample) {
    if (sample = "")
        return false
    for marker in CHROME_DETACH_MENU_PAGE_MARKER_SUBSTR {
        if InStr(sample, marker, false)
            return true
    }
    for marker in CHROME_DETACH_MENU_PAGE_MARKER_NAMES {
        if InStr(sample, marker, false)
            return true
    }
    return false
}

Chrome_ContextMenuInspectPopupHwnd(hwnd) {
    info := { classify: "unknown", sample: "" }
    if !hwnd
        return info
    try {
        info.sample := Chrome_DetachSampleMenuItemsForHwnd(hwnd)
        if Chrome_ContextMenuSampleLooksLikePageMenu(info.sample) {
            info.classify := "page"
            return info
        }
        if Chrome_ContextMenuSampleLooksLikeTabMenu(info.sample) {
            info.classify := "tab"
            return info
        }
        root := UIA.ElementFromHandle(hwnd)
        if Chrome_ContextMenuFindInRoot(root, CHROME_DETACH_MENU_PAGE_MARKER_NAMES,
            CHROME_DETACH_MENU_PAGE_MARKER_SUBSTR
        ) {
            info.classify := "page"
            return info
        }
        if Chrome_ContextMenuFindInRoot(root, CHROME_DETACH_MENU_TAB_MARKER_NAMES,
            CHROME_DETACH_MENU_TAB_MARKER_SUBSTR
        )
            info.classify := "tab"
    } catch {
    }
    return info
}

Chrome_ContextMenuClassifyPopupHwnd(hwnd) {
    return Chrome_ContextMenuInspectPopupHwnd(hwnd).classify
}

Chrome_DetachClearMenuPopup(session) {
    session.menuPopupHwnd := 0
    session.menuPopupClassify := ""
}

Chrome_ContextMenuCapturePopup(session, baselinePopups := "", timeoutMs := 0) {
    waitMs := timeoutMs > 0 ? timeoutMs : CHROME_DETACH_MENU_POPUP_MS
    base := IsObject(baselinePopups) ? baselinePopups : session.baselinePopups
    deadline := A_TickCount + waitMs
    while (A_TickCount < deadline) {
        for h in Chrome_ContextMenuListNewPopupCandidates(session, base) {
            inspected := Chrome_ContextMenuInspectPopupHwnd(h)
            ; #region agent log
            Chrome_DetachDebugLog("Chrome_ContextMenuCapturePopup", "popup candidate", "G", "hwnd=" . h .
                ";classify=" . inspected.classify . ";sample=" . inspected.sample)
            ; #endregion
            if (inspected.classify = "page")
                continue
            if (inspected.classify = "tab") {
                session.menuPopupHwnd := h
                session.menuPopupClassify := "tab"
                ; #region agent log
                Chrome_DetachDebugLog("Chrome_ContextMenuCapturePopup", "tab popup selected", "G", "hwnd=" . h)
                ; #endregion
                return h
            }
        }
        Sleep CHROME_DETACH_MENU_POLL_MS
    }
    if Chrome_ContextMenuFindTabMenuInBrowser(session) {
        session.menuPopupClassify := "tab"
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_ContextMenuCapturePopup", "tab menu in browser tree", "G", "hwnd=" .
            session.hwnd)
        ; #endregion
        return session.hwnd
    }
    Chrome_DetachClearMenuPopup(session)
    return 0
}

Chrome_ContextMenuPopupIsPageMenu(session) {
    if !session.menuPopupHwnd
        return false
    if (session.menuPopupClassify = "page")
        return true
    if (session.menuPopupClassify = "tab")
        return false
    return Chrome_ContextMenuInspectPopupHwnd(session.menuPopupHwnd).classify = "page"
}

Chrome_ContextMenuFindTabMenuInBrowser(session) {
    try {
        root := UIA.ElementFromHandle(session.hwnd)
        if Chrome_ContextMenuFindInRoot(root, CHROME_DETACH_MENU_PAGE_MARKER_NAMES,
            CHROME_DETACH_MENU_PAGE_MARKER_SUBSTR
        )
            return false
        if Chrome_ContextMenuFindInRoot(root, CHROME_DETACH_MENU_TAB_MARKER_NAMES,
            CHROME_DETACH_MENU_TAB_MARKER_SUBSTR
        )
            return true
        for el in root.FindAll({ Type: UIA.Type.MenuItem }, UIA.TreeScope.Subtree) {
            if Chrome_ContextMenuNameMatches(el, CHROME_DETACH_MENU_TAB_MARKER_NAMES,
                CHROME_DETACH_MENU_TAB_MARKER_SUBSTR
            )
                return true
        }
    } catch {
    }
    return false
}

Chrome_ContextMenuFocusedLooksLikeTabMenu() {
    try {
        el := UIA.GetFocusedElement()
        if !el
            return false
        if (el.Type = UIA.Type.MenuItem || el.Type = UIA.Type.Menu) {
            if Chrome_ContextMenuNameMatches(el, CHROME_DETACH_MENU_TAB_MARKER_NAMES,
                CHROME_DETACH_MENU_TAB_MARKER_SUBSTR
            )
                return true
            if Chrome_ContextMenuNameMatches(el, CHROME_DETACH_MENU_TAB_MARKER_NAMES, [])
                return true
        }
    } catch {
    }
    return false
}

Chrome_ContextMenuDismiss() {
    ClipAngel_ReleaseChordModifiersForSend()
    Send "{Escape}"
}

Chrome_ContextMenuSampleLooksLikeTabMenu(sample) {
    if (sample = "")
        return false
    for marker in CHROME_DETACH_MENU_TAB_MARKER_SUBSTR {
        if InStr(sample, marker, false)
            return true
    }
    for marker in CHROME_DETACH_MENU_TAB_MARKER_NAMES {
        if InStr(sample, marker, false)
            return true
    }
    return false
}

Chrome_ContextMenuFocusPopup(session) {
    if !session.menuPopupHwnd
        return false
    try {
        UIA.ElementFromHandle(session.menuPopupHwnd).SetFocus()
        return true
    } catch {
    }
    return false
}

Chrome_ContextMenuSendKeys(session, keys) {
    ClipAngel_ReleaseChordModifiersForSend()
    if session.menuPopupHwnd {
        try {
            ControlSend keys, , "ahk_id " session.menuPopupHwnd
            return true
        } catch {
        }
    }
    SendInput keys
    return true
}

Chrome_ContextMenuActivateItem(item) {
    if !item
        return false
    try item.Invoke()
    catch {
        try item.Click()
        catch {
            return false
        }
    }
    return true
}

; --- State-gated context menu phases (detach tab) ---
Chrome_WaitForTabContextMenu(session, timeoutMs := 0) {
    waitMs := timeoutMs > 0 ? timeoutMs : CHROME_DETACH_MENU_POPUP_MS
    deadline := A_TickCount + waitMs
    while (A_TickCount < deadline) {
        if !session.menuPopupHwnd
            Chrome_ContextMenuCapturePopup(session, session.baselinePopups)
        if Chrome_ContextMenuLooksLikeTabMenu(session)
            return true
        if Chrome_ContextMenuFindTabMenuInBrowser(session) {
            session.menuPopupClassify := "tab"
            return true
        }
        Sleep CHROME_DETACH_MENU_POLL_MS
    }
    return false
}

Chrome_FindDetachMenuTarget(session) {
    flat := Chrome_ContextMenuFindFirst(session, CHROME_DETACH_MENU_EN_NAMES, false)
    if flat
        return { type: "flat", item: flat }
    parent := Chrome_ContextMenuFindFirst(session, CHROME_DETACH_MENU_PARENT_NAMES, true)
    if parent
        return { type: "parent", item: parent }
    return 0
}

Chrome_WaitForDetachMenuTarget(session, timeoutMs := 0) {
    waitMs := timeoutMs > 0 ? timeoutMs : CHROME_DETACH_MENU_CHILD_MS
    deadline := A_TickCount + waitMs
    while (A_TickCount < deadline) {
        target := Chrome_FindDetachMenuTarget(session)
        if target
            return target
        if !session.menuPopupHwnd
            Chrome_ContextMenuCapturePopup(session, session.baselinePopups)
        Sleep CHROME_DETACH_MENU_POLL_MS
    }
    return 0
}

Chrome_ContextMenuIsDismissed(session) {
    if (session.menuPopupHwnd && WinExist("ahk_id " session.menuPopupHwnd))
        return false
    try {
        if Chrome_ContextMenuFindTabMenuInBrowser(session)
            return false
    } catch {
    }
    try {
        if Chrome_ContextMenuFocusedLooksLikeTabMenu()
            return false
    } catch {
    }
    return true
}

Chrome_WaitForContextMenuDismissed(session, timeoutMs := 0) {
    waitMs := timeoutMs > 0 ? timeoutMs : CHROME_DETACH_MENU_POPUP_MS
    deadline := A_TickCount + waitMs
    while (A_TickCount < deadline) {
        if Chrome_ContextMenuIsDismissed(session) {
            Chrome_DetachClearMenuPopup(session)
            return true
        }
        Sleep CHROME_DETACH_MENU_POLL_MS
    }
    return false
}

; Invoke detach item after menu is ready. Window wait is caller's job (dismiss wait removed — hook detects HWND faster).
Chrome_ActivateDetachMenuItemConfirmed(session, menuReady := false) {
    if !menuReady && !Chrome_WaitForTabContextMenu(session, CHROME_DETACH_MENU_POPUP_MS)
        return false
    target := Chrome_FindDetachMenuTarget(session)
    if !target
        target := Chrome_WaitForDetachMenuTarget(session, CHROME_DETACH_MENU_CHILD_MS)
    if !target
        return false
    if (target.type = "flat")
        return Chrome_ContextMenuActivateItem(target.item)
    if !Chrome_ContextMenuActivateItem(target.item)
        return false
    child := Chrome_ContextMenuWaitForSession(session, CHROME_DETACH_MENU_CHILD_NAMES, CHROME_DETACH_MENU_CHILD_MS,
        false)
    if !child
        return false
    return Chrome_ContextMenuActivateItem(child)
}

Chrome_WindowHasCaption(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    try {
        style := DllCall("GetWindowLongPtr", "ptr", hwnd, "int", -16, "ptr")
        return !!(style & 0x00C00000)
    } catch {
        return false
    }
}

; Safe for F6 / tab context menu when NOT in F11 (F6 in F11 can open DevTools).
Chrome_WaitUntilNotF11ForDetach(hwnd, timeoutMs := 0) {
    settleMs := timeoutMs > 0 ? timeoutMs : CHROME_DETACH_F11_SETTLE_MS
    deadline := A_TickCount + settleMs
    loop {
        if (!hwnd || !WinExist("ahk_id " hwnd))
            return false
        isF11 := WM_WindowIsF11Fullscreen(hwnd)
        if (!isF11) {
            fgOk := WM_EnsureForegroundForSend(hwnd, Min(800, Max(0, deadline - A_TickCount)))
            return true
        }
        if (A_TickCount >= deadline)
            break
        Sleep 50
    }
    return false
}

; Legacy name used for new-window strip; only requires leaving F11, not caption.
Chrome_WaitUntilWindowedForDetach(hwnd, timeoutMs := 0) {
    return Chrome_WaitUntilNotF11ForDetach(hwnd, timeoutMs)
}

Chrome_WaitUntilF11ForHwnd(hwnd, timeoutMs := 0) {
    settleMs := timeoutMs > 0 ? timeoutMs : CHROME_DETACH_F11_SETTLE_MS
    return WM_WaitForF11State(hwnd, true, settleMs)
}

Chrome_ExitF11ForDetach(hwnd) {
    if (!WM_WindowIsF11Fullscreen(hwnd))
        return Chrome_WaitUntilWindowedForDetach(hwnd)
    loop 3 {
        if !WM_ExitF11FullscreenForHwnd(hwnd, CHROME_DETACH_F11_SETTLE_MS)
            continue
        if Chrome_WaitUntilNotF11ForDetach(hwnd)
            return true
    }
    return false
}

Chrome_EnsureNewWindowIsWindowed(newHwnd) {
    if (!newHwnd || !WinExist("ahk_id " newHwnd))
        return false
    if (!WM_WindowIsF11Fullscreen(newHwnd))
        return true
    WM_ExitF11FullscreenForHwnd(newHwnd, CHROME_DETACH_F11_SETTLE_MS)
    return !WM_WindowIsF11Fullscreen(newHwnd)
}

Chrome_FocusDetachedWindow(newHwnd) {
    if (!newHwnd || !WinExist("ahk_id " newHwnd))
        return false
    if WinActive("ahk_id " newHwnd)
        return true
    try WinActivate("ahk_id " newHwnd)
    catch {
        return false
    }
    return WinWaitActive("ahk_id " newHwnd, , 0.25) || WinActive("ahk_id " newHwnd)
}

Chrome_RestoreF11OnOriginal(originalHwnd) {
    if (!WinExist("ahk_id " originalHwnd))
        return false
    loop 3 {
        if !WM_EnterF11FullscreenForHwnd(originalHwnd, CHROME_DETACH_F11_SETTLE_MS)
            continue
        if Chrome_WaitUntilF11ForHwnd(originalHwnd)
            return true
    }
    return WM_WindowIsF11Fullscreen(originalHwnd)
}

Chrome_ActivateDetachedWindow(newHwnd, originalHwnd, wasF11) {
    Chrome_FocusDetachedWindow(newHwnd)
}

Chrome_WaitForNewWindow(existingSet, timeoutMs := 0, tabTitle := "") {
    return Chrome_WaitForDetachNewWindow(existingSet, timeoutMs, tabTitle)
}

Chrome_EnsureBrowserForeground(hwnd) {
    if !(hwnd is Integer && hwnd > 0)
        return false
    if WinActive("ahk_id " hwnd)
        return true
    try WinActivate("ahk_id " hwnd)
    catch {
        return false
    }
    return WinWaitActive("ahk_id " hwnd, , 1)
}

Chrome_IsValidTabElement(tab) {
    if (!IsObject(tab) || !tab)
        return false
    try return tab.Type = UIA.Type.TabItem
    catch {
        return false
    }
}

Chrome_DetachGetWindowTitleForMatch(hwnd) {
    try title := WinGetTitle("ahk_id " hwnd)
    catch {
        return ""
    }
    return RegExReplace(title, "i) - Google Chrome$", "")
}

Chrome_DetachGetActiveTab(session) {
    cached := session.activeTab
    if Chrome_IsValidTabElement(cached) {
        try {
            if (cached.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable)
                && cached.SelectionItemPattern.IsSelected)
                return cached
        } catch {
        }
    }
    uia := session.uia
    method := "none"
    tabCount := -1
    if !IsObject(uia) {
        try {
            uia := session.uia := UIA_Browser("ahk_id " session.hwnd)
        } catch {
            ; #region agent log
            Chrome_DetachDebugLog("Chrome_DetachGetActiveTab", "uia attach failed", "A", { tabCount: 0, method: "none" })
            ; #endregion
            return 0
        }
    }
    try {
        uia.GetCurrentMainPaneElement()
        try {
            tabCount := uia.GetAllTabs().Length
        } catch {
            tabCount := -1
        }
        try {
            tab := uia.GetTab("")
            if Chrome_IsValidTabElement(tab) {
                method := "getTab"
                session.activeTab := tab
                ; #region agent log
                Chrome_DetachDebugLog("Chrome_DetachGetActiveTab", "tab resolved", "A", { tabCount: tabCount,
                    method: method })
                ; #endregion
                return tab
            }
        } catch {
        }
        for t in uia.GetAllTabs() {
            try {
                if (t.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable)
                    && t.SelectionItemPattern.IsSelected) {
                    if Chrome_IsValidTabElement(t) {
                        method := "selection"
                        session.activeTab := t
                        ; #region agent log
                        Chrome_DetachDebugLog("Chrome_DetachGetActiveTab", "tab resolved", "A", { tabCount: tabCount,
                            method: method })
                        ; #endregion
                        return t
                    }
                }
            } catch {
            }
        }
        chromeTitle := Chrome_DetachGetWindowTitleForMatch(session.hwnd)
        if (chromeTitle != "") {
            for t in uia.GetAllTabs() {
                try tabName := t.Name
                catch {
                    continue
                }
                if (tabName = "" || !Chrome_IsValidTabElement(t))
                    continue
                if (tabName = chromeTitle || InStr(chromeTitle, tabName, false) || InStr(tabName, chromeTitle, false)) {
                    method := "title"
                    session.activeTab := t
                    ; #region agent log
                    Chrome_DetachDebugLog("Chrome_DetachGetActiveTab", "tab resolved", "A", { tabCount: tabCount,
                        method: method, titleLen: StrLen(chromeTitle) })
                    ; #endregion
                    return t
                }
            }
        }
    } catch {
    }
    ; #region agent log
    Chrome_DetachDebugLog("Chrome_DetachGetActiveTab", "tab not found", "A", { tabCount: tabCount, method: method })
    ; #endregion
    return 0
}

Chrome_ContextMenuLooksLikeTabMenu(session) {
    if !session.menuPopupHwnd
        Chrome_ContextMenuCapturePopup(session, session.baselinePopups)
    if (session.menuPopupHwnd && session.menuPopupClassify = "tab") {
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_ContextMenuLooksLikeTabMenu", "tab menu cached", "D", { isTabMenu: true,
            popupHwnd: session.menuPopupHwnd })
        ; #endregion
        return true
    }
    if !session.menuPopupHwnd {
        if Chrome_ContextMenuFindTabMenuInBrowser(session) {
            ; #region agent log
            Chrome_DetachDebugLog("Chrome_ContextMenuLooksLikeTabMenu", "tab menu in browser tree", "G", { isTabMenu: true })
            ; #endregion
            return true
        }
        if Chrome_ContextMenuFocusedLooksLikeTabMenu() {
            ; #region agent log
            Chrome_DetachDebugLog("Chrome_ContextMenuLooksLikeTabMenu", "tab menu focused", "G", { isTabMenu: true })
            ; #endregion
            return true
        }
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_ContextMenuLooksLikeTabMenu", "no popup hwnd", "D", { isTabMenu: false,
            newWins: Chrome_DetachDebugListNewChromeWindows(session) })
        ; #endregion
        return false
    }
    inspected := Chrome_ContextMenuInspectPopupHwnd(session.menuPopupHwnd)
    if (inspected.classify = "page") {
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_ContextMenuLooksLikeTabMenu", "page menu rejected", "D", { isTabMenu: false,
            popupHwnd: session.menuPopupHwnd, menuSample: inspected.sample })
        ; #endregion
        return false
    }
    if (inspected.classify = "tab") {
        session.menuPopupClassify := "tab"
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_ContextMenuLooksLikeTabMenu", "tab menu confirmed", "D", { isTabMenu: true,
            popupHwnd: session.menuPopupHwnd })
        ; #endregion
        return true
    }
    deadline := A_TickCount + 150
    sample := inspected.sample
    while (A_TickCount < deadline) {
        inspected := Chrome_ContextMenuInspectPopupHwnd(session.menuPopupHwnd)
        sample := inspected.sample
        if (inspected.classify = "page") {
            ; #region agent log
            Chrome_DetachDebugLog("Chrome_ContextMenuLooksLikeTabMenu", "page menu rejected", "D", { isTabMenu: false,
                popupHwnd: session.menuPopupHwnd, menuSample: sample })
            ; #endregion
            return false
        }
        if (inspected.classify = "tab") {
            session.menuPopupClassify := "tab"
            ; #region agent log
            Chrome_DetachDebugLog("Chrome_ContextMenuLooksLikeTabMenu", "tab menu confirmed", "D", { isTabMenu: true,
                popupHwnd: session.menuPopupHwnd })
            ; #endregion
            return true
        }
        Sleep CHROME_DETACH_MENU_POLL_MS
    }
    ; #region agent log
    Chrome_DetachDebugLog("Chrome_ContextMenuLooksLikeTabMenu", "not tab menu", "D", { isTabMenu: false,
        popupHwnd: session.menuPopupHwnd, menuSample: sample })
    ; #endregion
    return false
}

Chrome_FocusedElementIsSelectedTab(session) {
    try {
        focused := UIA.GetFocusedElement()
        if !focused
            return false
        if (focused.Type = UIA.Type.TabItem) {
            try {
                if (focused.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable)
                    && focused.SelectionItemPattern.IsSelected)
                    return true
            } catch {
            }
        }
        uia := session.uia
        if !IsObject(uia)
            return false
        try uia.GetCurrentMainPaneElement()
        tab := uia.GetTab("")
        if !Chrome_IsValidTabElement(tab)
            return false
        try {
            if UIA.CompareElements(focused, tab)
                return true
        } catch {
        }
    } catch {
    }
    return false
}

; Poll until selected tab keeps keyboard focus (avoids AppsKey on wrong F6 stop).
Chrome_WaitForSelectedTabFocus(session, timeoutMs := 0) {
    waitMs := timeoutMs > 0 ? timeoutMs : CHROME_DETACH_TAB_FOCUS_WAIT_MS
    stableNeed := CHROME_DETACH_TAB_FOCUS_STABLE_MS
    deadline := A_TickCount + waitMs
    stableSince := 0
    while (A_TickCount < deadline) {
        if Chrome_FocusedElementIsSelectedTab(session) {
            if !stableSince
                stableSince := A_TickCount
            if (A_TickCount - stableSince >= stableNeed) {
                ; #region agent log
                Chrome_DetachDebugLog("Chrome_WaitForSelectedTabFocus", "stable tab focus", "I", { waitedMs: A_TickCount -
                    (deadline - waitMs), stableMs: stableNeed })
                ; #endregion
                return true
            }
        } else {
            stableSince := 0
        }
        Sleep CHROME_DETACH_MENU_POLL_MS
    }
    ; #region agent log
    Chrome_DetachDebugLog("Chrome_WaitForSelectedTabFocus", "timeout", "I", { waitMs: waitMs })
    ; #endregion
    return false
}

; Hover tab (no click), settle, AppsKey — Chrome hit-tests cursor for context menu.
Chrome_TryHoverAppsKeyTabMenu(session, tab, settleMs := 0) {
    if !Chrome_HoverActiveTab(session, tab)
        return false
    waitMs := settleMs > 0 ? settleMs : CHROME_DETACH_HOVER_APPSKEY_SETTLE_MS
    Sleep waitMs
    ClipAngel_ReleaseChordModifiersForSend()
    session.baselinePopups := Chrome_DetachListMenuPopups(session.hwnd)
    Chrome_DetachClearMenuPopup(session)
    ; #region agent log
    Chrome_DetachDebugLog("Chrome_TryHoverAppsKeyTabMenu", "AppsKey after hover", "J", "focus=" .
        Chrome_DetachDebugFocusedElement(session) . ";settleMs=" . waitMs)
    ; #endregion
    SendInput "{AppsKey}"
    Sleep CHROME_DETACH_APPSKEY_AFTER_MS
    Chrome_ContextMenuCapturePopup(session, session.baselinePopups)
    if Chrome_ContextMenuLooksLikeTabMenu(session)
        return true
    if Chrome_ContextMenuFindTabMenuInBrowser(session) {
        Chrome_DetachClearMenuPopup(session)
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_TryHoverAppsKeyTabMenu", "tab menu in browser tree", "G", "")
        ; #endregion
        return true
    }
    ; #region agent log
    Chrome_DetachDebugLog("Chrome_TryHoverAppsKeyTabMenu", "not tab menu", "D", "popup=" .
        session.menuPopupHwnd . ";newWins=" . Chrome_DetachDebugListNewChromeWindows(session))
    ; #endregion
    Chrome_ContextMenuDismiss()
    return false
}

; F6/SetFocus path: keyboard focus on tab, then AppsKey via ControlSend.
Chrome_TryFocusAppsKeyTabMenu(session, tab) {
    if !Chrome_DetachFocusActiveTab(session)
        return false
    Chrome_HoverActiveTab(session, tab)
    Sleep CHROME_DETACH_HOVER_SETTLE_MS
    if !Chrome_WaitForSelectedTabFocus(session)
        return false
    Sleep CHROME_DETACH_APPSKEY_SETTLE_MS
    ClipAngel_ReleaseChordModifiersForSend()
    session.baselinePopups := Chrome_DetachListMenuPopups(session.hwnd)
    Chrome_DetachClearMenuPopup(session)
    ; #region agent log
    Chrome_DetachDebugLog("Chrome_TryFocusAppsKeyTabMenu", "AppsKey after focus", "J", "focus=" .
        Chrome_DetachDebugFocusedElement(session))
    ; #endregion
    try session.uia.ControlSend("{AppsKey}")
    catch {
        SendInput "{AppsKey}"
    }
    Sleep CHROME_DETACH_APPSKEY_AFTER_MS
    Chrome_ContextMenuCapturePopup(session, session.baselinePopups)
    if Chrome_ContextMenuLooksLikeTabMenu(session)
        return true
    Chrome_ContextMenuDismiss()
    return false
}

Chrome_UiaLooksLikeTabStripFocused(session) {
    if Chrome_FocusedElementIsSelectedTab(session)
        return true
    try {
        focused := UIA.GetFocusedElement()
        if focused && focused.Type = UIA.Type.TabItem
            return true
    } catch {
    }
    return false
}

Chrome_WaitForTabStripFocus(session, timeoutMs := 0) {
    global CHROME_DETACH_F6_FOCUS_POLL_MS, CHROME_DETACH_MENU_POLL_MS
    waitMs := timeoutMs > 0 ? timeoutMs : CHROME_DETACH_F6_FOCUS_POLL_MS
    deadline := A_TickCount + waitMs
    pollMs := CHROME_DETACH_MENU_POLL_MS
    while (A_TickCount < deadline) {
        if Chrome_UiaLooksLikeTabStripFocused(session)
            return true
        Sleep pollMs
    }
    return false
}

; Official keyboard path: F6 until tab strip focus, then AppsKey (no hover hit-test).
Chrome_TryF6KeyboardTabMenu(session) {
    ; Modifiers already released by Chrome_Detach_PreSendSanitizeModifiers — short safety check only.
    ClipAngel_ReleaseChordModifiersForSend()
    Send "{Escape}"
    deadline := A_TickCount + 30
    while (A_TickCount < deadline) {
        if !GetKeyState("Shift", "P") && !GetKeyState("Ctrl", "P") && !GetKeyState("Alt", "P")
            break
        Sleep 10
    }
    uiaUsable := Chrome_SessionUiaUsable(session)
    loop CHROME_DETACH_F6_FOCUS_MAX {
        if (session.menuPopupHwnd && Chrome_ContextMenuLooksLikeTabMenu(session)) {
            session.menuConfirmed := true
            return true
        }
        Send "{F6}"
        tabFocused := Chrome_WaitForTabStripFocus(session, CHROME_DETACH_F6_FOCUS_POLL_MS)
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_TryF6KeyboardTabMenu", "f6 step", "K", "i=" . A_Index .
            " focus=" . Chrome_DetachDebugFocusedElement(session) . ";tabFocused=" . (tabFocused ? 1 : 0))
        ; #endregion
        tryAppsKey := tabFocused
        if !tryAppsKey && !uiaUsable && A_Index >= 2
            tryAppsKey := true  ; UIA dead: AppsKey from 2nd F6 onward, confirm via menu wait
        if !tryAppsKey
            continue
        ClipAngel_ReleaseChordModifiersForSend()
        session.baselinePopups := Chrome_DetachListMenuPopups(session.hwnd)
        Chrome_DetachClearMenuPopup(session)
        try session.uia.ControlSend("{AppsKey}")
        catch {
            SendInput "{AppsKey}"
        }
        if Chrome_WaitForTabContextMenu(session, CHROME_DETACH_MENU_POPUP_MS) {
            session.menuConfirmed := true
            ; #region agent log
            Chrome_DetachDebugLog("Chrome_TryF6KeyboardTabMenu", "tab menu ready", "K", "i=" . A_Index)
            ; #endregion
            return true
        }
        Chrome_ContextMenuDismiss()
    }
    return false
}

Chrome_OpenActiveTabContextMenuViaTabFocus(session) {
    tab := Chrome_DetachGetActiveTab(session)
    if (session.menuPopupHwnd && Chrome_ContextMenuLooksLikeTabMenu(session))
        return true
    if !Chrome_IsValidTabElement(tab)
        return false
    if Chrome_TryF6KeyboardTabMenu(session)
        return true
    if Chrome_TryFocusAppsKeyTabMenu(session, tab)
        return true
    loop CHROME_DETACH_HOVER_ATTEMPTS {
        settle := CHROME_DETACH_HOVER_APPSKEY_SETTLE_MS + (A_Index - 1) * CHROME_DETACH_HOVER_APPSKEY_RETRY_EXTRA_MS
        if Chrome_TryHoverAppsKeyTabMenu(session, tab, settle)
            return true
    }
    if !CHROME_DETACH_F6_FALLBACK
        return false
    ClipAngel_ReleaseChordModifiersForSend()
    Send "{Escape}"
    Sleep 40
    loop CHROME_DETACH_F6_FOCUS_MAX {
        if (session.menuPopupHwnd && Chrome_ContextMenuLooksLikeTabMenu(session))
            return true
        Send "{F6}"
        Sleep CHROME_DETACH_F6_STEP_MS
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_OpenActiveTabContextMenuViaTabFocus", "f6+hover step", "C", "i=" . A_Index .
            " focus=" . Chrome_DetachDebugFocusedElement(session))
        ; #endregion
        if Chrome_TryHoverAppsKeyTabMenu(session, tab)
            return true
    }
    return false
}

; Move pointer over tab strip target only (no click).
Chrome_HoverActiveTab(session, tab) {
    if !Chrome_IsValidTabElement(tab)
        return false
    try {
        rect := tab.BoundingRectangle
        if !(rect.r > rect.l && rect.b > rect.t)
            return false
        x := rect.l + (rect.r - rect.l) // 2
        y := rect.t + (rect.b - rect.t) // 2
        uia := session.uia
        if IsObject(uia) {
            try {
                uia.GetCurrentMainPaneElement()
                tbRect := uia.TabBarElement.BoundingRectangle
                if (tbRect.b > tbRect.t)
                    y := tbRect.t + (tbRect.b - tbRect.t) // 2
            } catch {
            }
        }
        saveMode := A_CoordModeMouse
        CoordMode "Mouse", "Screen"
        MouseMove x, y, 0
        CoordMode "Mouse", saveMode
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_HoverActiveTab", "hovered tab", "H", { x: x, y: y })
        ; #endregion
        return true
    } catch {
        return false
    }
}

Chrome_DetachFocusActiveTab(session) {
    tab := Chrome_DetachGetActiveTab(session)
    if !Chrome_IsValidTabElement(tab)
        return false
    focused := false
    try {
        tab.SetFocus()
        focused := true
    } catch {
    }
    if !focused {
        try {
            if (tab.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable))
                tab.SelectionItemPattern.Select()
            focused := true
        } catch {
        }
    }
    ok := Chrome_WaitForSelectedTabFocus(session, 350)
    ; #region agent log
    Chrome_DetachDebugLog("Chrome_DetachFocusActiveTab", "focus tab", "H", { setFocus: focused ? 1 : 0,
        tabFocused: ok ? 1 : 0 })
    ; #endregion
    return ok
}

Chrome_FocusTabStripAndOpenContextMenu(session) {
    hwnd := session.hwnd
    if !Chrome_NormalizeFocusToPage(hwnd) {
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_FocusTabStripAndOpenContextMenu", "normalize page failed", "C", { result: 0 })
        ; #endregion
        return false
    }
    ClipAngel_ReleaseChordModifiersForSend()
    loop CHROME_DETACH_F6_FOCUS_MAX {
        if (session.menuPopupHwnd && Chrome_ContextMenuLooksLikeTabMenu(session))
            return true
        Send "{F6}"
        Sleep CHROME_DETACH_F6_STEP_MS
        tab := Chrome_DetachGetActiveTab(session)
        if !Chrome_IsValidTabElement(tab)
            continue
        try tab.SetFocus()
        catch {
        }
        Sleep CHROME_DETACH_F6_REFOCUS_MS
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_FocusTabStripAndOpenContextMenu", "f6 iteration", "C", { iteration: A_Index,
            stepMs: CHROME_DETACH_F6_STEP_MS, focus: Chrome_DetachDebugFocusedElement(session) })
        ; #endregion
        if Chrome_UiaLooksLikeTabStripFocused(session) {
            Sleep CHROME_DETACH_APPSKEY_SETTLE_MS
            ClipAngel_ReleaseChordModifiersForSend()
            session.baselinePopups := Chrome_DetachListMenuPopups(session.hwnd)
            Chrome_DetachClearMenuPopup(session)
            try session.uia.ControlSend("{AppsKey}")
            catch {
                SendInput "{AppsKey}"
            }
            Sleep CHROME_DETACH_APPSKEY_AFTER_MS
            Chrome_ContextMenuCapturePopup(session, session.baselinePopups)
            if Chrome_ContextMenuLooksLikeTabMenu(session) {
                ; #region agent log
                Chrome_DetachDebugLog("Chrome_FocusTabStripAndOpenContextMenu", "f6+AppsKey success", "C", { result: 1,
                    iteration: A_Index })
                ; #endregion
                return true
            }
            Chrome_ContextMenuDismiss()
        }
        if Chrome_TryHoverAppsKeyTabMenu(session, tab) {
            ; #region agent log
            Chrome_DetachDebugLog("Chrome_FocusTabStripAndOpenContextMenu", "f6+hover AppsKey success", "C", { result: 1,
                iteration: A_Index })
            ; #endregion
            return true
        }
        Chrome_ContextMenuDismiss()
    }
    ; #region agent log
    Chrome_DetachDebugLog("Chrome_FocusTabStripAndOpenContextMenu", "all paths failed", "C", { result: 0 })
    ; #endregion
    return false
}

Chrome_FocusPageContent(hwnd) {
    try {
        for ctrlName in ["Chrome_RenderWidgetHostHWND1", "Chrome_RenderWidgetHostHWND"] {
            try {
                rw := ControlGetHwnd(ctrlName, "ahk_id " hwnd)
                if (rw) {
                    ControlFocus ctrlName, "ahk_id " hwnd
                    return true
                }
            } catch {
            }
        }
    } catch {
    }
    return false
}

; Dismiss overlays, focus page content, then F6×2 opens tab context menu (page → toolbar → tab strip).
Chrome_NormalizeFocusToPage(hwnd) {
    if !Chrome_EnsureBrowserForeground(hwnd)
        return false
    ClipAngel_ReleaseChordModifiersForSend()
    Send "{Escape}"
    Send "{Escape}"
    Send "{Escape}"
    Chrome_FocusPageContent(hwnd)
    return WM_EnsureForegroundForSend(hwnd, 1000)
}

; Detach hot path: caller already foregrounded — one Escape, skip page focus when UIA usable.
Chrome_NormalizeFocusToPageLight(hwnd, session := unset) {
    ClipAngel_ReleaseChordModifiersForSend()
    Send "{Escape}"
    if !(IsSet(session) && IsObject(session) && Chrome_SessionUiaUsable(session))
        Chrome_FocusPageContent(hwnd)
    if WinActive("ahk_id " hwnd)
        return true
    return WM_EnsureForegroundForSend(hwnd, 200)
}

Chrome_OpenActiveTabContextMenu(session) {
    global CHROME_DETACH_USE_LIGHT_NORMALIZE
    hwnd := session.hwnd
    if !WinActive("ahk_id " hwnd) && !Chrome_EnsureBrowserForeground(hwnd)
        return false
    if (session.menuPopupHwnd && Chrome_ContextMenuLooksLikeTabMenu(session)) {
        session.menuConfirmed := true
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_OpenActiveTabContextMenu", "reuse open tab menu", "E", "popup=" .
            session.menuPopupHwnd)
        ; #endregion
        return true
    }
    session.baselinePopups := Chrome_DetachListMenuPopups(hwnd)
    Chrome_DetachClearMenuPopup(session)
    if CHROME_DETACH_USE_LIGHT_NORMALIZE {
        if !Chrome_NormalizeFocusToPageLight(hwnd, session)
            return false
    } else if !Chrome_NormalizeFocusToPage(hwnd) {
        return false
    }
    if Chrome_TryF6KeyboardTabMenu(session) {
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_OpenActiveTabContextMenu", "opened via f6", "H", { result: 1 })
        ; #endregion
        return true
    }
    if Chrome_OpenActiveTabContextMenuViaTabFocus(session) {
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_OpenActiveTabContextMenu", "opened via tab focus", "H", { result: 1 })
        ; #endregion
        return true
    }
    if !CHROME_DETACH_DEEP_FALLBACK
        return false
    f6Ok := Chrome_FocusTabStripAndOpenContextMenu(session)
    ; #region agent log
    Chrome_DetachDebugLog("Chrome_OpenActiveTabContextMenu", "f6+AppsKey fallback result", "C", { path: "F6AppsKey",
        result: f6Ok ? 1 : 0 })
    ; #endregion
    return f6Ok
}

Chrome_DetachCountTabs(session) {
    if !IsObject(session.uia)
        return -1
    try {
        return session.uia.GetAllTabs().Length
    } catch {
        return -1
    }
}

; UIA Invoke first; PT keyboard m -> Enter -> n (Nova janela submenu). Never bare 'n' at top level.
Chrome_ActivateDetachMenuItem(session) {
    if !session.menuPopupHwnd
        Chrome_ContextMenuCapturePopup(session, session.baselinePopups)
    ; #region agent log
    Chrome_DetachDebugLog("Chrome_ActivateDetachMenuItem", "start", "E", "popup=" . session.menuPopupHwnd .
        " sample=" . Chrome_DetachDebugSampleMenuItems(session))
    ; #endregion
    if Chrome_ContextMenuPopupIsPageMenu(session) {
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_ActivateDetachMenuItem", "abort page menu", "E", "popup=" .
            session.menuPopupHwnd . ";sample=" . Chrome_DetachDebugSampleMenuItems(session))
        ; #endregion
        return false
    }
    focused := Chrome_ContextMenuFocusPopup(session)
    ClipAngel_ReleaseChordModifiersForSend()

    flat := Chrome_ContextMenuFindFirst(session, CHROME_DETACH_MENU_EN_NAMES, false)
    if flat && Chrome_ContextMenuActivateItem(flat) {
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_ActivateDetachMenuItem", "flat item invoked", "E", "path=flat")
        ; #endregion
        return true
    }

    parent := Chrome_ContextMenuFindFirst(session, CHROME_DETACH_MENU_PARENT_NAMES, true)
    if parent && Chrome_ContextMenuActivateItem(parent) {
        child := Chrome_ContextMenuWaitForSession(session, CHROME_DETACH_MENU_CHILD_NAMES, CHROME_DETACH_MENU_CHILD_MS,
            false)
        if child && Chrome_ContextMenuActivateItem(child) {
            ; #region agent log
            Chrome_DetachDebugLog("Chrome_ActivateDetachMenuItem", "parent+child invoked", "E", "path=submenu")
            ; #endregion
            return true
        }
    }

    if Chrome_ContextMenuPopupIsPageMenu(session) {
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_ActivateDetachMenuItem", "abort page menu", "E", "popup=" .
            session.menuPopupHwnd . ";sample=" . Chrome_DetachDebugSampleMenuItems(session))
        ; #endregion
        return false
    }
    ; #region agent log
    Chrome_DetachDebugLog("Chrome_ActivateDetachMenuItem", "UIA miss, keyboard PT", "E", "sample=" .
        Chrome_DetachDebugSampleMenuItems(session) . ";popupFocus=" . (focused ? 1 : 0))
    ; #endregion
    Chrome_ContextMenuFocusPopup(session)
    ClipAngel_ReleaseChordModifiersForSend()
    Chrome_ContextMenuSendKeys(session, "m")
    child := Chrome_ContextMenuWaitForSession(session, CHROME_DETACH_MENU_CHILD_NAMES, CHROME_DETACH_MENU_CHILD_MS,
        false)
    if child && Chrome_ContextMenuActivateItem(child) {
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_ActivateDetachMenuItem", "keyboard child invoked", "E", "path=keyboardChild")
        ; #endregion
        return true
    }
    ClipAngel_ReleaseChordModifiersForSend()
    Chrome_ContextMenuSendKeys(session, "{Enter}")
    child := Chrome_ContextMenuWaitForSession(session, CHROME_DETACH_MENU_CHILD_NAMES, CHROME_DETACH_MENU_CHILD_MS,
        false)
    if child && Chrome_ContextMenuActivateItem(child) {
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_ActivateDetachMenuItem", "keyboard child invoked", "E", "path=keyboardEnterChild")
        ; #endregion
        return true
    }
    ; Submenu open: 'n' = Nova janela (safe here; top-level 'n' = Nova guia)
    Chrome_ContextMenuFocusPopup(session)
    ClipAngel_ReleaseChordModifiersForSend()
    Chrome_ContextMenuSendKeys(session, "n")
    child := Chrome_ContextMenuWaitForSession(session, CHROME_DETACH_MENU_CHILD_NAMES, CHROME_DETACH_MENU_CHILD_MS,
        false)
    if child && Chrome_ContextMenuActivateItem(child) {
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_ActivateDetachMenuItem", "keyboard n child invoked", "E", "path=keyboardN")
        ; #endregion
        return true
    }
    ClipAngel_ReleaseChordModifiersForSend()
    Chrome_ContextMenuSendKeys(session, "{Enter}")
    ; #region agent log
    Chrome_DetachDebugLog("Chrome_ActivateDetachMenuItem", "keyboard attempted unverified", "E", "path=keyboardMN")
    ; #endregion
    return false
}

Chrome_ActivateDetachViaPtKeyboard(session) {
    return Chrome_ActivateDetachMenuItem(session)
}

Chrome_ActivateDetachViaEnKeyboard(session) {
    ; Disabled: previously re-entered Chrome_OpenActiveTabContextMenu and doubled F6 stacks.
    return false
}

Chrome_DetachActiveTabToNewWindow_Legacy() {
    Send "{F6}"
    Sleep 100
    Send "{F6}"
    Sleep 100
    Send "{AppsKey}"
    Sleep 100
    Send "m"
    Sleep 100
    Send "{Enter}"
    Sleep 100
    Send "{Enter}"
}

Chrome_IsLikelyTopLevelBrowserWindow(hwnd) {
    if !(hwnd is Integer && hwnd > 0) || !WinExist("ahk_id " hwnd)
        return false
    try {
        if (WinGetProcessName("ahk_id " hwnd) != "chrome.exe")
            return false
        cls := WinGetClass("ahk_id " hwnd)
        if !InStr(cls, "Chrome_WidgetWin")
            return false
        title := WinGetTitle("ahk_id " hwnd)
        if (Trim(title) = "")
            return false
        WinGetPos(, , &w, &h, "ahk_id " hwnd)
        return (w >= 300 && h >= 200)
    } catch {
        return false
    }
}

Chrome_GetChromeTopLevelHwndSet() {
    existingSet := Map()
    try {
        for h in WinGetList("ahk_exe chrome.exe") {
            if Chrome_IsLikelyTopLevelBrowserWindow(h)
                existingSet[h] := true
        }
    } catch {
    }
    return existingSet
}

Chrome_DetachTitleCore(title) {
    title := RegExReplace(title, "i) - Google Chrome$", "")
    title := RegExReplace(title, "i) - Chromium$", "")
    return Trim(title)
}

Chrome_DetachTitleMatches(hwnd, expectedCore) {
    if (expectedCore = "")
        return true
    if !(hwnd is Integer && hwnd > 0)
        return false
    try actual := Chrome_DetachTitleCore(WinGetTitle("ahk_id " hwnd))
    catch {
        return false
    }
    if (actual = "")
        return false
    return (actual = expectedCore) || InStr(actual, expectedCore, false) || InStr(expectedCore, actual, false)
}

Chrome_DetachCaptureBaseline(hwnd) {
    return { existingSet: Chrome_GetChromeTopLevelHwndSet(), tabTitle: Chrome_DetachGetWindowTitleForMatch(hwnd) }
}

Chrome_DetachQuickTabCount(hwnd) {
    if !(hwnd is Integer && hwnd > 0)
        return -1
    try {
        winTitle := "ahk_id " hwnd
        UIA.ActivateChromiumAccessibility(winTitle, 200)
        uia := UIA_Browser(winTitle)
        return uia.GetAllTabs().Length
    } catch {
        return -1
    }
}

Chrome_FindNewTopLevelChromeCandidate(existingSet, tabTitle := "") {
    titleCore := Chrome_DetachTitleCore(tabTitle)
    if !(existingSet is Map)
        existingSet := Map()
    try {
        for hwnd in WinGetList("ahk_exe chrome.exe") {
            if existingSet.Has(hwnd)
                continue
            if !Chrome_IsLikelyTopLevelBrowserWindow(hwnd)
                continue
            if (titleCore != "" && !Chrome_DetachTitleMatches(hwnd, titleCore))
                continue
            return hwnd
        }
    } catch {
    }
    return 0
}

Chrome_DetachPickValidatedNewHwnd(candidates, existingSet, tabTitle := "") {
    titleCore := Chrome_DetachTitleCore(tabTitle)
    if !(existingSet is Map)
        existingSet := Map()
    for hwnd in candidates {
        if (existingSet.Has(hwnd))
            continue
        if !Chrome_IsLikelyTopLevelBrowserWindow(hwnd)
            continue
        if (titleCore != "" && !Chrome_DetachTitleMatches(hwnd, titleCore))
            continue
        return hwnd
    }
    return 0
}

Chrome_DetachWinEventProc(hWinEventHook, event, hwnd, idObject, idChild, idEventThread, dwmsEventTime) {
    global g_ChromeDetachCreatedHwnds, CHROME_DETACH_EVENT_OBJECT_CREATE, CHROME_DETACH_OBJID_WINDOW
    if (event = CHROME_DETACH_EVENT_OBJECT_CREATE && idObject = CHROME_DETACH_OBJID_WINDOW && hwnd)
        g_ChromeDetachCreatedHwnds.Push(hwnd)
}

Chrome_DetachArmWindowCreateHook(&hookState) {
    global g_ChromeDetachCreatedHwnds, CHROME_DETACH_EVENT_OBJECT_CREATE
    hookState := { hHook: 0, cb: 0 }
    g_ChromeDetachCreatedHwnds := []
    hookState.cb := CallbackCreate(Chrome_DetachWinEventProc, "F Fast", 7)
    hookState.hHook := DllCall("user32\SetWinEventHook", "UInt", CHROME_DETACH_EVENT_OBJECT_CREATE,
        "UInt", CHROME_DETACH_EVENT_OBJECT_CREATE, "Ptr", 0, "Ptr", hookState.cb, "UInt", 0, "UInt", 0, "UInt", 0,
        "Ptr")
}

Chrome_DetachDisarmWindowCreateHook(hookState) {
    if IsObject(hookState) && hookState.hHook {
        DllCall("user32\UnhookWinEvent", "Ptr", hookState.hHook)
        hookState.hHook := 0
    }
}

; Bounded wait: new top-level Chrome window not in baseline, optionally matching detached tab title.
Chrome_WaitForDetachNewWindow(existingSet, timeoutMs := 0, tabTitle := "", hookState := unset) {
    global CHROME_DETACH_EXTENSION_TIMEOUT_MS, CHROME_DETACH_VERIFY_TIMEOUT_MS, CHROME_DETACH_MENU_POLL_MS
    global CHROME_DETACH_USE_WIN_EVENT_HOOK, g_ChromeDetachCreatedHwnds
    waitMs := timeoutMs > 0 ? timeoutMs : CHROME_DETACH_EXTENSION_TIMEOUT_MS
    if (waitMs <= 0)
        waitMs := CHROME_DETACH_VERIFY_TIMEOUT_MS
    deadline := A_TickCount + waitMs
    pollMs := CHROME_DETACH_MENU_POLL_MS
    if !(existingSet is Map)
        existingSet := Map()

    if CHROME_DETACH_USE_WIN_EVENT_HOOK {
        ownsHook := !(IsSet(hookState) && IsObject(hookState) && hookState.hHook)
        if ownsHook {
            Chrome_DetachArmWindowCreateHook(&hookState)
        }
        try {
            while (A_TickCount < deadline) {
                hwnd := Chrome_DetachPickValidatedNewHwnd(g_ChromeDetachCreatedHwnds, existingSet, tabTitle)
                if !hwnd
                    hwnd := Chrome_FindNewTopLevelChromeCandidate(existingSet, tabTitle)
                g_ChromeDetachCreatedHwnds := []
                if hwnd {
                    try WinActivate("ahk_id " hwnd)
                    return hwnd
                }
                Sleep pollMs
            }
        } finally {
            if ownsHook
                Chrome_DetachDisarmWindowCreateHook(hookState)
        }
        if (Chrome_DetachTitleCore(tabTitle) != "") {
            hwnd := Chrome_FindNewTopLevelChromeCandidate(existingSet, "")
            if hwnd {
                try WinActivate("ahk_id " hwnd)
                return hwnd
            }
        }
        return 0
    }

    while (A_TickCount < deadline) {
        hwnd := Chrome_FindNewTopLevelChromeCandidate(existingSet, tabTitle)
        if hwnd {
            try WinActivate("ahk_id " hwnd)
            return hwnd
        }
        Sleep pollMs
    }
    ; HWND often appears before Chrome updates the window title — accept top-level without title gate.
    if (Chrome_DetachTitleCore(tabTitle) != "") {
        hwnd := Chrome_FindNewTopLevelChromeCandidate(existingSet, "")
        if hwnd {
            try WinActivate("ahk_id " hwnd)
            return hwnd
        }
    }
    return 0
}

Chrome_WaitForNewTopLevelChromeWindow(existingSet, timeoutMs := 0, tabTitle := "") {
    return Chrome_WaitForDetachNewWindow(existingSet, timeoutMs, tabTitle)
}

Chrome_DetachRemainingMs(opStart, totalMs := 0) {
    global CHROME_DETACH_TOTAL_TIMEOUT_MS
    budget := totalMs > 0 ? totalMs : CHROME_DETACH_TOTAL_TIMEOUT_MS
    return Max(0, budget - (A_TickCount - opStart))
}

Chrome_DetachExtensionBudgetMs(opStart) {
    global CHROME_DETACH_EXTENSION_TIMEOUT_MS
    return Max(0, Min(CHROME_DETACH_EXTENSION_TIMEOUT_MS, Chrome_DetachRemainingMs(opStart)))
}

Chrome_DetachActiveTabToNewWindow_ExtensionFast(baseline := unset, timeoutMs := 0) {
    global CHROME_DETACH_USE_EXTENSION, CHROME_DETACH_EXTENSION_TIMEOUT_MS
    if !CHROME_DETACH_USE_EXTENSION
        return 0
    waitMs := timeoutMs > 0 ? Min(timeoutMs, CHROME_DETACH_EXTENSION_TIMEOUT_MS) : CHROME_DETACH_EXTENSION_TIMEOUT_MS
    if !IsSet(baseline) || !IsObject(baseline)
        baseline := Chrome_DetachCaptureBaseline(WinExist("A"))
    existingSet := baseline.existingSet
    tabTitle := baseline.tabTitle
    Chrome_Detach_PreSendSanitizeModifiers()
    ClipAngel_ReleaseChordModifiersForSend()
    SendInput "{Ctrl down}{Shift down}y{Shift up}{Ctrl up}"
    result := Chrome_WaitForDetachNewWindow(existingSet, waitMs, tabTitle)
    if !result && waitMs > 400 {
        ClipAngel_ReleaseChordModifiersForSend()
        Send "^+y"
        retryMs := Min(CHROME_DETACH_EXTENSION_TIMEOUT_MS // 2, Max(400, waitMs // 2))
        result := Chrome_WaitForDetachNewWindow(existingSet, retryMs, tabTitle)
    }
    return result
}

Chrome_DetachActiveTabToNewWindow_ExtensionFallback(baseline) {
    return Chrome_DetachActiveTabToNewWindow_ExtensionFast(baseline)
}

; One F6 pass (max 3 steps) + menu invoke — phased state gates, no blind advance.
Chrome_DetachActiveTabToNewWindow_UiaSingleShot(hwnd, &wasF11, baseline := unset, verifyTimeoutMs := 0) {
    ; #region perf
    t0 := A_TickCount
    ; #endregion
    wasF11 := false
    if !Chrome_PrepareWindowForTabDetach(hwnd, &wasF11)
        return 0
    if (WM_WindowIsF11Fullscreen(hwnd))
        return 0
    if !IsSet(baseline) || !IsObject(baseline)
        baseline := Chrome_DetachCaptureBaseline(hwnd)
    session := Chrome_DetachSessionCreate(hwnd, baseline.existingSet)
    tabCount := Chrome_DetachCountTabs(session)
    if (tabCount = 1)
        return 0
    ; #region perf
    t1 := A_TickCount
    ; #endregion
    if !Chrome_OpenActiveTabContextMenu(session)
        return 0
    ; #region perf
    t2 := A_TickCount
    ; #endregion
    if !session.menuConfirmed && !Chrome_WaitForTabContextMenu(session, CHROME_DETACH_MENU_POPUP_MS) {
        Chrome_ContextMenuDismiss()
        return 0
    }
    ; #region perf
    t3 := A_TickCount
    ; #endregion
    hookState := {}
    if CHROME_DETACH_USE_WIN_EVENT_HOOK
        Chrome_DetachArmWindowCreateHook(&hookState)
    try {
        if !Chrome_ActivateDetachMenuItemConfirmed(session, session.menuConfirmed) {
            Chrome_ContextMenuDismiss()
            return 0
        }
        ; #region perf
        t4 := A_TickCount
        ; #endregion
        verifyMs := verifyTimeoutMs > 0 ? verifyTimeoutMs : CHROME_DETACH_VERIFY_TIMEOUT_MS
        result := Chrome_WaitForDetachNewWindow(baseline.existingSet, verifyMs, baseline.tabTitle, hookState)
        ; #region perf
        t5 := A_TickCount
        if (CHROME_DETACH_PERF_LOG_ENABLED) {
            try FileAppend Format("detach: prep={} menuOpen={} menuWait={} menuAct={} newWin={} total={}`n",
                t1 - t0, t2 - t1, t3 - t2, t4 - t3, t5 - t4, t5 - t0), A_ScriptDir "\.cursor\chrome_detach_perf.log"
        }
        ; #endregion
        return result
    } finally {
        Chrome_DetachDisarmWindowCreateHook(hookState)
    }
}

; Debug-only: legacy UIA menu stack (never used in normal Shift+W path).
Chrome_DetachActiveTabToNewWindow_UiaLegacyPath(hwnd, &newHwnd, &wasF11) {
    global CHROME_DETACH_USE_UIA, CHROME_DETACH_USE_EXTENSION, CHROME_DETACH_VERIFY_TIMEOUT_MS
    newHwnd := 0
    wasF11 := false
    if !CHROME_DETACH_USE_UIA
        return false
    if !Chrome_EnsureBrowserForeground(hwnd)
        return false
    if !Chrome_PrepareWindowForTabDetach(hwnd, &wasF11)
        return false
    if (WM_WindowIsF11Fullscreen(hwnd))
        return false
    session := Chrome_DetachSessionCreate(hwnd)
    session.tabsBeforeDetach := Chrome_DetachCountTabs(session)
    if (session.tabsBeforeDetach = 1)
        return false
    if Chrome_RunDetachMenuSequence(session) {
        newHwnd := session.newDetachedHwnd
        if newHwnd
            return true
    }
    if (session.menuPopupHwnd && Chrome_ContextMenuLooksLikeTabMenu(session))
        Chrome_ActivateDetachMenuItem(session)
    newHwnd := Chrome_WaitForNewWindow(session.existingSet, CHROME_DETACH_VERIFY_TIMEOUT_MS,
        Chrome_DetachGetWindowTitleForMatch(hwnd))
    if newHwnd
        return true
    if CHROME_DETACH_USE_EXTENSION
        newHwnd := Chrome_DetachActiveTabToNewWindow_ExtensionFast(Chrome_DetachCaptureBaseline(hwnd))
    return newHwnd ? true : false
}

Chrome_PrepareWindowForTabDetach(hwnd, &wasF11) {
    wasF11 := WM_WindowIsF11Fullscreen(hwnd)
    if (wasF11) {
        try StandardLoadingBar_Update("🔄 Exiting F11 fullscreen…", BANNER_ACCENT_INTERMEDIATE)
        if !Chrome_ExitF11ForDetach(hwnd)
            return false
        try StandardLoadingBar_Update("⏳ Detaching tab to new window…", BANNER_ACCENT_INTERMEDIATE)
    }
    if WinActive("ahk_id " hwnd) && !WM_WindowIsF11Fullscreen(hwnd)
        return true
    return Chrome_WaitUntilNotF11ForDetach(hwnd)
}

; Reliable Nova guia signal: baseline count must be known (>=1) and exactly +1 tab after menu keys.
Chrome_DetachNovaGuiaLikely(tabsBefore, tabsAfter) {
    return (tabsBefore >= 1 && tabsAfter = tabsBefore + 1)
}

; Close accidental Nova guia only when detach did NOT create a new window.
Chrome_DetachCloseSpuriousNovaGuia(originalHwnd, tabsBefore, tabsAfter, existingSet) {
    if !Chrome_DetachNovaGuiaLikely(tabsBefore, tabsAfter)
        return false
    if Chrome_WaitForNewWindow(existingSet, 200)
        return false
    if !Chrome_EnsureBrowserForeground(originalHwnd)
        return false
    ClipAngel_ReleaseChordModifiersForSend()
    Send "{Ctrl down}w{Ctrl up}"
    return true
}

Chrome_FinishTabDetach(originalHwnd, newHwnd, wasF11) {
    if (!wasF11) {
        Chrome_FocusDetachedWindow(newHwnd)
        return
    }

    if (!WinExist("ahk_id " originalHwnd))
        return

    if (newHwnd)
        Chrome_EnsureNewWindowIsWindowed(newHwnd)

    Chrome_RestoreF11OnOriginal(originalHwnd)

    Chrome_FocusDetachedWindow(newHwnd)
}

Chrome_RunDetachMenuSequence(session) {
    session.tabsBeforeDetach := Chrome_DetachCountTabs(session)
    try StandardLoadingBar_Update("⏳ Detaching tab…", BANNER_ACCENT_INTERMEDIATE)
    loop CHROME_DETACH_SEQUENCE_ATTEMPTS {
        if !(session.menuPopupHwnd && Chrome_ContextMenuLooksLikeTabMenu(session)) {
            if !session.menuPopupHwnd
                Chrome_ContextMenuCapturePopup(session, session.baselinePopups)
            if !(session.menuPopupHwnd && Chrome_ContextMenuLooksLikeTabMenu(session)) {
                Chrome_DetachClearMenuPopup(session)
                session.baselinePopups := Chrome_DetachListMenuPopups(session.hwnd)
                opened := Chrome_OpenActiveTabContextMenu(session)
            } else {
                opened := true
            }
            if !opened {
                ; #region agent log
                Chrome_DetachDebugLog("Chrome_RunDetachMenuSequence", "open menu failed", "B", "attempt=" . A_Index)
                ; #endregion
                Chrome_ContextMenuDismiss()
                continue
            }
            if !session.menuPopupHwnd
                Chrome_ContextMenuCapturePopup(session, session.baselinePopups)
            if !Chrome_ContextMenuLooksLikeTabMenu(session) {
                ; #region agent log
                Chrome_DetachDebugLog("Chrome_RunDetachMenuSequence", "reject non-tab menu", "D", "attempt=" . A_Index .
                    " sample=" . Chrome_DetachDebugSampleMenuItems(session))
                ; #endregion
                Chrome_ContextMenuDismiss()
                continue
            }
        }

        Chrome_ActivateDetachMenuItem(session)
        newHwnd := Chrome_WaitForDetachNewWindow(session.existingSet, CHROME_DETACH_VERIFY_TIMEOUT_MS,
            Chrome_DetachGetWindowTitleForMatch(session.hwnd))
        if newHwnd {
            session.newDetachedHwnd := newHwnd
            ; #region agent log
            Chrome_DetachDebugLog("Chrome_RunDetachMenuSequence", "detach verified", "E", "attempt=" . A_Index .
                ";newHwnd=" . newHwnd)
            ; #endregion
            return true
        }

        Chrome_ContextMenuDismiss()
    }
    ; #region agent log
    Chrome_DetachDebugLog("Chrome_RunDetachMenuSequence", "sequence failed", "E", { result: 0 })
    ; #endregion
    return false
}

Chrome_DetachActiveTabToNewWindow() {
    global g_ChromeDetachBusy, CHROME_DETACH_USE_EXTENSION, CHROME_DETACH_USE_UIA_SINGLE_SHOT
    global CHROME_DETACH_ALLOW_UIA_DEBUG_FALLBACK, CHROME_DETACH_USE_UIA, CHROME_DETACH_PREFLIGHT_TAB_COUNT
    global CHROME_DETACH_TOTAL_TIMEOUT_MS, CHROME_DETACH_PERF_LOG_ENABLED
    if (g_ChromeDetachBusy)
        return false
    g_ChromeDetachBusy := true
    opStart := A_TickCount

    hwnd := 0
    newHwnd := 0
    wasF11 := false
    success := false
    baseline := unset
    pathTaken := ""

    StandardLoadingBar_Show("⏳ Detaching tab to new window…", BANNER_ACCENT_INTERMEDIATE, { passive: true })
    try {
        Chrome_Detach_PreSendSanitizeModifiers()

        hwnd := WinExist("A")
        if !(hwnd is Integer && hwnd > 0)
            return false
        try {
            if (WinGetProcessName("ahk_id " hwnd) != "chrome.exe")
                return false
        } catch {
            return false
        }

        if !Chrome_EnsureBrowserForeground(hwnd)
            return false

        baseline := Chrome_DetachCaptureBaseline(hwnd)

        if CHROME_DETACH_PREFLIGHT_TAB_COUNT {
            tabCount := Chrome_DetachQuickTabCount(hwnd)
            if (tabCount = 1) {
                ShowCenteredOverlay_Utils(
                    "❌ Cannot detach: only one tab in this window.",
                    2800, BANNER_ACCENT_ERROR)
                return false
            }
        }

        if CHROME_DETACH_USE_EXTENSION {
            try StandardLoadingBar_Update("⏳ Detaching tab (extension)…", BANNER_ACCENT_INTERMEDIATE)
            extBudget := Chrome_DetachExtensionBudgetMs(opStart)
            if (extBudget >= 300)
                newHwnd := Chrome_DetachActiveTabToNewWindow_ExtensionFast(baseline, extBudget)
            if newHwnd {
                pathTaken := "extension"
                success := true
                return true
            }
        }

        remaining := Chrome_DetachRemainingMs(opStart)
        if CHROME_DETACH_USE_UIA_SINGLE_SHOT && remaining > 600 {
            try StandardLoadingBar_Update("⏳ Detaching tab (menu)…", BANNER_ACCENT_INTERMEDIATE)
            uiaVerifyMs := Min(remaining, CHROME_DETACH_VERIFY_TIMEOUT_MS)
            newHwnd := Chrome_DetachActiveTabToNewWindow_UiaSingleShot(hwnd, &wasF11, baseline, uiaVerifyMs)
            if newHwnd {
                pathTaken := "UIA single-shot"
                success := true
                return true
            }
        }

        if !CHROME_DETACH_ALLOW_UIA_DEBUG_FALLBACK {
            detachErr := CHROME_DETACH_USE_EXTENSION
                ? "❌ Could not detach tab. Install PopActiveTab (Ctrl+Shift+Y) or use 2+ tabs."
                : "❌ Could not detach tab. Use 2+ tabs in this Chrome window."
            ShowCenteredOverlay_Utils(detachErr, 2800, BANNER_ACCENT_ERROR)
            return false
        }

        remaining := CHROME_DETACH_TOTAL_TIMEOUT_MS - (A_TickCount - opStart)
        if (remaining <= 800)
            return false

        try StandardLoadingBar_Update("⏳ Detaching tab (UIA debug)…", BANNER_ACCENT_INTERMEDIATE)
        if Chrome_DetachActiveTabToNewWindow_UiaLegacyPath(hwnd, &newHwnd, &wasF11) {
            pathTaken := "UIA legacy"
            success := true
            return true
        }
        return false
    } finally {
        ; #region perf
        if (CHROME_DETACH_PERF_LOG_ENABLED) {
            totalMs := A_TickCount - opStart
            try FileAppend Format("main: path={} success={} totalMs={} at={}`n",
                pathTaken, success ? 1 : 0, totalMs, A_Now), A_ScriptDir "\.cursor\chrome_detach_perf.log"
        }
        ; #endregion
        if (wasF11) {
            try StandardLoadingBar_Update("🔄 Restoring F11 fullscreen…", BANNER_ACCENT_INTERMEDIATE)
            Chrome_FinishTabDetach(hwnd, success ? newHwnd : 0, wasF11)
            if (!success && hwnd && WinExist("ahk_id " hwnd) && !WM_WindowIsF11Fullscreen(hwnd))
                Chrome_RestoreF11OnOriginal(hwnd)
        } else if (success && newHwnd) {
            Chrome_FocusDetachedWindow(newHwnd)
        }
        if (success && newHwnd)
            try StandardLoadingBar_Update("✅ Tab detached to new window", BANNER_ACCENT_SUCCESS)
        if (success)
            StandardLoadingBar_Hide(CHROME_DETACH_SUCCESS_HIDE_MS)
        else
            StandardLoadingBar_Hide(0)
        if (!success)
            try Send "{Escape}"
        ; #region agent log
        Chrome_DetachDebugLog("Chrome_DetachActiveTabToNewWindow", "finish", "F", { success: success ? 1 : 0,
            newHwnd: newHwnd })
        ; #endregion
        g_ChromeDetachBusy := false
    }
}

; --- Gemini mode picker (3.1 Flash-Lite / 3.5 Flash / 3.1 Pro), mouse + UIA ----
; UI tree reference: gemini-no-context-menu.md

global GEMINI_MODEL_CANONICAL_NAMES := ["3.1 Flash-Lite", "3.5 Flash", "3.1 Pro"]
global GEMINI_MODE_PICKER_NAME_SUBSTR := "Open mode picker"
global GEMINI_MODE_MENU_WAIT_MS := 450
global GEMINI_MODE_MENU_POLL_MS := 40
global GEMINI_MODE_PICKER_LABEL_WAIT_MS := 500

FindGeminiChromeHwnd() {
    for hwnd in WinGetList("ahk_exe chrome.exe") {
        try {
            if InStr(WinGetTitle("ahk_id " hwnd), "gemini", false)
                return hwnd
        } catch {
        }
    }
    return 0
}

; One UIA_Browser bind per flow. lightAttach: skip activate/a11y when caller already focused Gemini.
GeminiAttachBrowser(geminiHwnd := 0, lightAttach := false) {
    if (!geminiHwnd)
        geminiHwnd := FindGeminiChromeHwnd()
    if (!geminiHwnd || !WinExist("ahk_id " geminiHwnd))
        return 0
    winTitle := "ahk_id " geminiHwnd
    if (!lightAttach) {
        if (!WinActive(winTitle)) {
            try WinActivate(winTitle)
            try WinWaitActive(winTitle, , 1)
        }
        try UIA.ActivateChromiumAccessibility(winTitle, 300)
        catch {
        }
    }
    try {
        return UIA_Browser(winTitle)
    } catch {
        return 0
    }
}

GeminiFindRenderWidgetControl(browserHwnd) {
    for ctrlName in ["Chrome_RenderWidgetHostHWND1", "Chrome_RenderWidgetHostHWND"] {
        try {
            if ControlGetHwnd(ctrlName, "ahk_id " browserHwnd)
                return ctrlName
        } catch {
        }
    }
    try {
        for ctrlName in WinGetControls("ahk_id " browserHwnd) {
            if InStr(ctrlName, "Chrome_RenderWidgetHostHWND")
                return ctrlName
        }
    } catch {
    }
    return ""
}

GeminiGetElementWindowClickCoords(el, browserHwnd, uia := 0) {
    if !IsObject(el) || !browserHwnd
        return 0
    elPos := 0
    try {
        elPos := el.GetPos("window", browserHwnd)
    } catch {
        return 0
    }
    if (!IsObject(elPos) || elPos.w <= 0 || elPos.h <= 0)
        return 0
    cx := elPos.x + elPos.w // 2
    cy := elPos.y + elPos.h // 2
    docOffset := "none"
    if IsObject(uia) {
        try {
            doc := uia.GetCurrentDocumentElement()
            if IsObject(doc) {
                docPos := doc.GetPos("window", browserHwnd)
                if (IsObject(docPos) && docPos.w > 0 && docPos.h > 0) {
                    cx := elPos.x - docPos.x + elPos.w // 2
                    cy := elPos.y - docPos.y + elPos.h // 2
                    docOffset := "yes"
                }
            }
        } catch {
        }
    }
    return { cx: cx, cy: cy, docOffset: docOffset, elPosX: elPos.x, elPosY: elPos.y }
}

; skipActivate: caller already activated Gemini (hot path for mode picker).
GeminiMouseClickElement(el, browserHwnd := 0, uia := 0, skipActivate := false) {
    if !IsObject(el)
        return false
    winTitle := browserHwnd ? "ahk_id " browserHwnd : ""
    if (browserHwnd && !skipActivate) {
        if (!WinExist(winTitle))
            return false
        if (!WinActive(winTitle)) {
            WinActivate(winTitle)
            WinWaitActive(winTitle, , 1)
        }
    }
    if (browserHwnd) {
        try {
            el.ControlClick("left", 1, "", winTitle)
            return true
        } catch {
        }
        try {
            if (el.Click())
                return true
        } catch {
        }
        try {
            pt := el.GetClickablePoint()
            prevCoordMode := A_CoordModeMouse
            CoordMode("Mouse", "Screen")
            Click(pt.x, pt.y)
            CoordMode("Mouse", prevCoordMode)
            return true
        } catch {
        }
        coords := GeminiGetElementWindowClickCoords(el, browserHwnd, uia)
        if IsObject(coords) {
            renderCtrl := GeminiFindRenderWidgetControl(browserHwnd)
            if (renderCtrl != "") {
                try {
                    ControlGetPos(&rwx, &rwy, , , renderCtrl, winTitle)
                    ControlClick("X" . (coords.cx - rwx) . " Y" . (coords.cy - rwy), winTitle, renderCtrl)
                    return true
                } catch {
                }
            }
        }
    }
    return false
}

GeminiGetBrowserHwndFromUia(uia) {
    if !IsObject(uia)
        return 0
    try {
        hwnd := uia.BrowserId
        if (hwnd && WinExist("ahk_id " hwnd))
            return hwnd
    } catch {
    }
    return WinExist("ahk_exe chrome.exe") ? WinExist("ahk_exe chrome.exe") : 0
}

GeminiDismissModePickerMenu(uia, browserHwnd := 0) {
    if !IsObject(uia)
        return false
    if (!browserHwnd)
        browserHwnd := GeminiGetBrowserHwndFromUia(uia)
    promptField := FindGeminiPromptField(uia)
    if !promptField
        return false
    return GeminiMouseClickElement(promptField, browserHwnd, uia, true)
}

FindGeminiModePickerButton(uia) {
    if !IsObject(uia)
        return 0
    for typeSpec in [50000, "Button"] {
        try {
            el := uia.FindFirst({ Type: typeSpec, Name: GEMINI_MODE_PICKER_NAME_SUBSTR, mm: 2, cs: 0 })
            if el
                return el
        } catch {
        }
    }
    return 0
}

GeminiGetActiveModelFromPickerElement(picker) {
    if !IsObject(picker)
        return ""
    try {
        if RegExMatch(picker.Name, "i)currently\s+(.+)$", &m) {
            norm := GeminiNormalizeModelLabel(Trim(m[1]))
            if (norm != "")
                return norm
        }
    } catch {
    }
    return ""
}

FindGeminiModelMenuItem(uia, modelName) {
    if !IsObject(uia)
        return 0
    exp := GeminiNormalizeModelLabel(modelName)
    if (exp = "")
        return 0
    for typeSpec in [50011, "MenuItem", 50008, "RadioButton", 50003, "ListItem"] {
        try {
            el := uia.FindFirst({ Type: typeSpec, Name: exp, mm: 3, cs: 0 })
            if el
                return el
        } catch {
        }
        try {
            el := uia.FindFirst({ Type: typeSpec, Name: exp, mm: 2, cs: 0 })
            if el
                return el
        } catch {
        }
    }
    return 0
}

GeminiWaitForModelMenuItem(uia, modelName, timeoutMs := 0) {
    if (timeoutMs <= 0)
        timeoutMs := GEMINI_MODE_MENU_WAIT_MS
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        el := FindGeminiModelMenuItem(uia, modelName)
        if el
            return el
        Sleep GEMINI_MODE_MENU_POLL_MS
    }
    return 0
}

GeminiWaitForPickerShowsModel(uia, expected, timeoutMs := 0) {
    if (timeoutMs <= 0)
        timeoutMs := GEMINI_MODE_PICKER_LABEL_WAIT_MS
    exp := GeminiNormalizeModelLabel(expected)
    if (exp = "")
        return false
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        picker := FindGeminiModePickerButton(uia)
        if (picker && GeminiGetActiveModelFromPickerElement(picker) = exp)
            return true
        Sleep GEMINI_MODE_MENU_POLL_MS
    }
    return false
}

GeminiNormalizeModelLabel(name) {
    if (name = "")
        return ""
    for canonical in GEMINI_MODEL_CANONICAL_NAMES {
        if (name = canonical || RegExMatch(name, "i)^" . RegExReplace(canonical, "\.", "\.") . "(\s|$)"))
            return canonical
    }
    ; Picker button suffix after "currently" (e.g. Flash-Lite, Flash, Pro)
    if RegExMatch(name, "i)Flash-Lite")
        return "3.1 Flash-Lite"
    if RegExMatch(name, "i)\bFlash\b") && !RegExMatch(name, "i)Flash-Lite")
        return "3.5 Flash"
    if RegExMatch(name, "i)\bPro\b")
        return "3.1 Pro"
    return ""
}

GetGeminiActiveModelFromPickerOnly(uia) {
    return GeminiGetActiveModelFromPickerElement(FindGeminiModePickerButton(uia))
}

GeminiCollectModelMenuItemState(mi) {
    className := ""
    try {
        className := mi.ClassName
    } catch {
        className := ""
    }
    isDisabled := false
    try {
        if (InStr(className, "disabled") || InStr(className, "mat-mdc-button-disabled"))
            isDisabled := true
        try {
            if (!mi.GetPropertyValue(UIA.Property.IsEnabled))
                isDisabled := true
        } catch {
        }
    } catch {
    }
    isSelected := false
    try {
        if (mi.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable))
            isSelected := mi.SelectionItemPattern.IsSelected
    } catch {
    }
    hasLegacy := false
    legacyState := -1
    try {
        hasLegacy := mi.GetPropertyValue(UIA.Property.IsLegacyIAccessiblePatternAvailable)
    } catch {
        hasLegacy := false
    }
    if (hasLegacy) {
        try {
            legacyState := mi.LegacyIAccessiblePattern.State
        } catch {
            try {
                legacyState := mi.LegacyIAccessiblePattern.CurrentState
            } catch {
                legacyState := -1
            }
        }
    }
    if (!isSelected) {
        try {
            if (InStr(className, "is-selected") || InStr(className, "selected") || InStr(className, "active") ||
                InStr(className, "mdc-selected"))
                isSelected := true
        } catch {
        }
    }
    if (!isSelected) {
        try {
            if (hasLegacy && legacyState != -1 && (legacyState & 0x4))
                isSelected := true
        } catch {
        }
    }
    return { isDisabled: isDisabled, isSelected: isSelected, className: className }
}

GeminiCollectModelMenuItems(uia) {
    modelButtons := []
    menuItems := []
    for typeSpec in [50011, "MenuItem", 50008, "RadioButton", 50003, "ListItem"] {
        try {
            found := uia.FindAll({ Type: typeSpec })
        } catch {
            continue
        }
        if (IsObject(found) && found.Length) {
            menuItems := found
            break
        }
    }
    for mi in menuItems {
        try {
            fullName := mi.Name
            shortName := GeminiNormalizeModelLabel(fullName)
            if (shortName = "")
                continue
            st := GeminiCollectModelMenuItemState(mi)
            modelButtons.Push({ btn: mi, name: shortName, isSelected: st.isSelected, isDisabled: st.isDisabled,
                className: st.className })
        } catch {
        }
    }
    return modelButtons
}

; Back-compat alias for callers not yet updated
GeminiCollectModelOptionButtons(uia) {
    return GeminiCollectModelMenuItems(uia)
}

GeminiOpenModePickerMenu(uia, picker := "", browserHwnd := 0) {
    if !IsObject(uia)
        return false
    if (!browserHwnd)
        browserHwnd := GeminiGetBrowserHwndFromUia(uia)
    if !IsObject(picker)
        picker := FindGeminiModePickerButton(uia)
    if !picker
        return false
    return GeminiMouseClickElement(picker, browserHwnd, uia, true)
}

EnsureGeminiModelViaMenu(expected, geminiHwnd := 0) {
    exp := GeminiNormalizeModelLabel(expected)
    if (exp = "")
        return false
    uia := GeminiAttachBrowser(geminiHwnd)
    if !IsObject(uia)
        return false
    browserHwnd := geminiHwnd ? geminiHwnd : GeminiGetBrowserHwndFromUia(uia)
    picker := FindGeminiModePickerButton(uia)
    if !picker
        return false
    if (GeminiGetActiveModelFromPickerElement(picker) = exp)
        return true
    if (!GeminiOpenModePickerMenu(uia, picker, browserHwnd)) {
        GeminiDismissModePickerMenu(uia, browserHwnd)
        return false
    }
    targetBtn := GeminiWaitForModelMenuItem(uia, exp)
    if !targetBtn {
        GeminiDismissModePickerMenu(uia, browserHwnd)
        return false
    }
    if (!GeminiMouseClickElement(targetBtn, browserHwnd, uia, true)) {
        GeminiDismissModePickerMenu(uia, browserHwnd)
        return false
    }
    return GeminiWaitForPickerShowsModel(uia, exp)
}

EnsureGeminiThinkingLevelMenuOpen(geminiHwnd := 0) {
    uia := GeminiAttachBrowser(geminiHwnd)
    if !IsObject(uia)
        return false
    browserHwnd := geminiHwnd ? geminiHwnd : GeminiGetBrowserHwndFromUia(uia)
    picker := FindGeminiModePickerButton(uia)
    if (!picker || !GeminiOpenModePickerMenu(uia, picker, browserHwnd)) {
        GeminiDismissModePickerMenu(uia, browserHwnd)
        return false
    }
    deadline := A_TickCount + GEMINI_MODE_MENU_WAIT_MS
    thinkingItem := 0
    while (A_TickCount < deadline) {
        try {
            thinkingItem := uia.FindFirst({ Name: "Thinking level", mm: 2, cs: 0, Type: 50011 })
        } catch {
            thinkingItem := 0
        }
        if thinkingItem
            break
        Sleep GEMINI_MODE_MENU_POLL_MS
    }
    if !thinkingItem {
        GeminiDismissModePickerMenu(uia, browserHwnd)
        return false
    }
    if (!GeminiMouseClickElement(thinkingItem, browserHwnd, uia, true)) {
        GeminiDismissModePickerMenu(uia, browserHwnd)
        return false
    }
    return true
}

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

; Cross-process IPC for Hotstring Selector (symmetric with WindowManagement project selector IPC)
global g_HS_SelectorOpenFile := A_ScriptDir "\.cursor\hs_selector_open"
global g_HS_SelectorCloseRequestFile := A_ScriptDir "\.cursor\hs_selector_close_request"
global g_HS_SelectorCloseCheckTimer := ""

Utils_CheckHotstringSelectorCloseRequest() {
    global g_HotstringSelectorActive, g_HS_SelectorCloseRequestFile
    if (!g_HotstringSelectorActive)
        return
    if (FileExist(g_HS_SelectorCloseRequestFile)) {
        try FileDelete(g_HS_SelectorCloseRequestFile)
        catch {
        }
        CleanupHotstringSelector()
    }
}

; Trigger only on Space or Tab, not Enter or punctuation
Hotstring("EndChars", " `t")

RegisterHotstring(trigger, expansion, category := "", title := "", char := "") {
    global g_hotstrings
    g_hotstrings.Push({ trigger: trigger, expansion: expansion, category: category, title: title, char: char })
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
    InsertText("my links")
}

:o:gintegra::
{
    InsertText("GS_UX core team_UX and CIP Integration")
}

:o:gdash::
{
    InsertText("GS_E&S_CIP Dashboard research and design")
}

:o:opex-cim-journey-mapping::
{
    InsertText("opex-cim-journey-mapping")
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
    InsertText("project management LA")
}

:o:guxcip::
{
    InsertText("UX and CIP")
}

:o:gtrain::
{
    InsertText("GS_UX core team_Trainings Management")
}

:o:gbp::
{
    InsertText("GS_B2C_Portals and Key Accounts Process POC")
}

GetPromptText(key) {
    promptFile := A_ScriptDir "\prompt\" key ".txt"
    try {
        return FileRead(promptFile)
    } catch {
        return "[PROMPT FILE MISSING: " key "]"
    }
}

:o:cgrammar::
{
    InsertText(GetPromptText("grammar"))
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
    InsertText(GetPromptText("mtask"))
}

:o:flog::
{
    InsertText(GetPromptText("flog"))
}

:o:aiopt::
{
    InsertText(GetPromptText("aiopt"))
}

:o:cplant::
{
    InsertText(GetPromptText("cplant"))
}

:o:aibrapid::
{
    InsertText(GetPromptText("aib-rapid-fire-template"))
}

:o:pptslide::
{
    InsertText(GetPromptText("slide-creation"))
}

:o:pptslideref::
{
    InsertText(GetPromptText("slide-creation-with-ref"))
}

:o:csvfill::
{
    InsertText(GetPromptText("unstructured-to-csv"))
}

:o:mdunesc::
{
    UnescapeMarkdownClipboard()
}

; MyNotes technique prompts: live repo path first, then mirror under prompt\technique (synced by aux\Sync-MyNotesTechniquePrompts.ps1).
GetTechniquePromptFilePath(fileName) {
    repo := GetNotesRepoPath()
    dir := (repo != "") ? repo "\studies\technique\prompts" : ""
    if (dir != "" && FileExist(dir "\" fileName))
        return dir "\" fileName
    mirror := A_ScriptDir "\prompt\technique\" fileName
    if FileExist(mirror)
        return mirror
    if (dir != "")
        return dir "\" fileName
    return mirror
}

InitTechniquePromptHotstrings() {
    ; Five files live in MyNotes: studies\technique\prompts (resolved via GetNotesRepoPath in env.ahk).
    defs := [
        ["story-prompt.txt", ":o:mnemonic", "📖 Creating mnemonic stories", "", "Reserved 3"],
        ["video-transcription-prompt.txt", ":o:ytranscript", "🎬 Transcript Youtube Video", "", "Reserved 4"],
        ["story-reduction-prompt.txt", ":o:storyreduction", "📝 Story reduction", "a", "Reserved 5"],
        ["punctual-beast-append-prompt.txt", ":o:punctualbeast", "🧩 Punctual beast append", "p", "Reserved 6"],
        ["image-background-preservation-prompt.txt", ":o:imgpreserve", "🛡️ Preserve background for image generation",
            "g", "Reserved 7"],
    ]
    for row in defs {
        fileName := row[1]
        trigger := row[2]
        title := row[3]
        exChar := row[4]
        reserved := row[5]
        try {
            body := FileRead(GetTechniquePromptFilePath(fileName))
            if (exChar != "")
                RegisterHotstring(trigger, body, "Prompts", title, exChar)
            else
                RegisterHotstring(trigger, body, "Prompts", title)
        } catch {
            RegisterHotstring("", "", "Prompts", reserved)
        }
    }
}

; ----------------------
; Register hotstrings for cheat sheet display
; ----------------------
InitHotstringsCheatSheet() {
    promptDir := A_ScriptDir "\prompt"

    ; Prompts (4 items) - First category
    try {
        RegisterHotstring(":o:cgrammar", FileRead(promptDir "\grammar.txt"), "Prompts",
            "✏️ Grammar & Spelling Corrector")
    } catch {
        RegisterHotstring(":o:cgrammar",
            "Correct grammar, spelling, punctuation, and casing. Give back only the text.`n", "Prompts",
            "✏️ Grammar & Spelling Corrector")
    }
    try {
        RegisterHotstring(":o:mtask", FileRead(promptDir "\mtask.txt"), "Prompts", "🔲 Convert to Task")
    } catch {
        RegisterHotstring(":o:mtask", "Translate this into a task. Output ONLY the task. Start with 🔲.`n", "Prompts",
            "🔲 Convert to Task")
    }
    try {
        RegisterHotstring(":o:flog", FileRead(promptDir "\flog.txt"), "Prompts", "🍽️ Food Log Dictation")
    } catch {
        RegisterHotstring(":o:flog", "Food_Log dictation → Excel CSV. Output ONLY the final CSV block.`n", "Prompts",
            "🍽️ Food Log Dictation")
    }
    try {
        RegisterHotstring(":o:aiopt", FileRead(promptDir "\aiopt.txt"), "Prompts", "🤖 AI Text Optimizer")
    } catch {
        RegisterHotstring(":o:aiopt",
            "Rewrite the input text so it becomes AI-oriented. Preserve all important information.`n", "Prompts",
            "🤖 AI Text Optimizer")
    }

    ; Placeholder prompts (6 slots reserved for future prompts)
    ; Technical Architect & Code Planner (content from prompt/markdown-plan.txt)
    try {
        planPrompt := FileRead(promptDir "\markdown-plan.txt")
        RegisterHotstring(":o:cplan", planPrompt, "Prompts", "📋 Technical Architect & Code Planner")
    } catch {
        RegisterHotstring(":o:cplan",
            "You are an expert Technical Architect and Code Planner. Generate a .plan.md file. **Input Task:**`n",
            "Prompts", "📋 Technical Architect & Code Planner")
    }
    try {
        RegisterHotstring(":o:cplant", FileRead(promptDir "\cplant.txt"), "Prompts", "📝 Plan File Template")
    } catch {
        RegisterHotstring(":o:cplant",
            "---`nname: [Title]`noverview: [Summary]`ntodos:`n  - id: x`n    content: [Step]`n    status: pending`n---`n",
            "Prompts", "📝 Plan File Template")
    }
    InitTechniquePromptHotstrings()
    try {
        aibRapidFireTpl := FileRead(promptDir "\aib-rapid-fire-template.txt")
        RegisterHotstring(":o:aibrapid", aibRapidFireTpl, "Prompts", "📜 Junior AI: ⚡ rapid-fire template")
    } catch {
        RegisterHotstring(":o:aibrapid",
            "Junior AI (AIB): planning doc with ⚡ - conceptual above, execution steps below.`n", "Prompts",
            "📜 Junior AI: ⚡ rapid-fire template")
    }
    try {
        RegisterHotstring(":o:pptslide", FileRead(promptDir "\slide-creation.txt"), "Prompts",
            "📊 Create PowerPoint slide")
    } catch {
        RegisterHotstring(":o:pptslide", "Create one PowerPoint slide as an image.`n", "Prompts",
            "📊 Create PowerPoint slide")
    }
    try {
        RegisterHotstring(":o:pptslideref", FileRead(promptDir "\slide-creation-with-ref.txt"), "Prompts",
            "📊 Create PowerPoint slide (reference)")
    } catch {
        RegisterHotstring(":o:pptslideref",
            "Create one PowerPoint slide as an image using the attached reference as the main visual guide.`n", "Prompts",
            "📊 Create PowerPoint slide (reference)")
    }
    try {
        RegisterHotstring(":o:csvfill", FileRead(promptDir "\unstructured-to-csv.txt"), "Prompts",
            "📋 Fill CSV from unstructured text")
    } catch {
        RegisterHotstring(":o:csvfill",
            "Extract information from unstructured text and fill/update CSV rows using the provided column schema.`n",
            "Prompts", "📋 Fill CSV from unstructured text")
    }

    ; Hotstrings: emails
    RegisterHotstring(":o:ebosch", "eduardo.figueiredo@br.bosch.com", "Hotstrings", "💼 Bosch Email")
    RegisterHotstring(":o:egoogle", "edu.evangelista.figueiredo@gmail.com", "Hotstrings", "📧 Gmail")

    ; Projects (Cursor workspaces) - keys align with Project Selector 2
    RegisterHotstring(":o:gintegra", "GS_UX core team_UX and CIP Integration", "Projects", "🔄 UX and CIP Integration",
        "u")
    RegisterHotstring(":o:gdash", "GS_E&S_CIP Dashboard research and design", "Projects", "📊 CIP Dashboard", "d")
    RegisterHotstring(":o:boiler-plate", "boiler-plate", "Projects", "🧱 boiler-plate", "0")
    RegisterHotstring(":o:astra", "astra", "Projects", "⭐ astrA", "a")
    RegisterHotstring(":o:opex-cim-journey-mapping", "opex-cim-journey-mapping", "Projects",
        "E&S Opex CIM Journey Mapping",
        "o")
    RegisterHotstring(":o:gpilotb2b", "Piloto PT B2B", "Projects", "🧪 Piloto PT B2B", "b")
    RegisterHotstring(":o:gpython", "17 - Python Scripts", "Projects", "🐍 Python Scripts", "t")

    ; Hotstrings (non-workspace "project-like" names)
    RegisterHotstring(":o:myl", "my links", "Hotstrings", "🔗 my links", "m")
    RegisterHotstring(":o:gpm", "project management LA", "Hotstrings", "📋 project management LA", "p")
    RegisterHotstring(":o:guxcip", "UX and CIP", "Hotstrings", "🔗 UX and CIP", "x")
    RegisterHotstring(":o:gtrain", "GS_UX core team_Trainings Management", "Hotstrings", "🎓 Trainings Management", "t"
    )
}
InitHotstringsCheatSheet()

; =============================================================================
; Files & Links System
; =============================================================================

; Global variables for quick open files
global g_QuickOpenFiles := []
global g_QuickOpenFileCharMap := Map()  ; Maps character to file path

; Register a file for quick opening
RegisterQuickOpenFile(filePath, title) {
    global g_QuickOpenFiles
    g_QuickOpenFiles.Push({ filePath: filePath, title: title, category: "Files & Links" })
}

; Initialize quick open files
InitQuickOpenFiles() {
    ; Register dissertation Power BI file with character 'y'
    RegisterQuickOpenFile(
        "C:\Users\eduev\Meu Drive\04 - Pós-graduação\01 - Mestrado\26-ai-experiment\infoVis\Dissertation InfoVis  - PowerBI - Charts.pbix",
        "📊 Dissertation InfoVis"
    )

    ; Register radio-tiso exercises YouTube link
    RegisterQuickOpenFile(
        "https://www.youtube.com/watch?v=I6ZRH9Mraqw&t=2s",
        "📻 Radio-Tiso Exercises"
    )

    ; Register GS_UX core team_UX and CIP Integration Miro
    RegisterQuickOpenFile(
        "https://miro.com/app/board/uXjVJdbNFkA=/",
        "🎨 GS_UX core team_UX and CIP Integration Miro"
    )

    ; Register GS_E&S_CIP Dashboard research and design Miro
    RegisterQuickOpenFile(
        "https://miro.com/app/board/uXjVJVZSXvk=/",
        "📊 GS_E&S_CIP Dashboard research and design Miro"
    )
}
InitQuickOpenFiles()

; =============================================================================
; Macros System
; =============================================================================

; Global variables for macros
global g_Macros := []
global g_MacroCharMap := Map()  ; Maps character to macro function
global g_ProgrammaticDictationStop := false  ; Skip ~#!+0 when script sends #!+0 programmatically
global g_GeminiToggleTab := 1  ; Last Gemini tab chosen by ^!#4 (UIA-synced); other code may still assume 1/2 toggle state

; Register a macro
RegisterMacro(func, title, char := "") {
    global g_Macros
    g_Macros.Push({ func: func, title: title, category: "Macros", char: char })
}

; Get scripts directory path based on environment
GetScriptsDirectory() {
    global IS_WORK_ENVIRONMENT
    if (IS_WORK_ENVIRONMENT) {
        return "C:\Users\fie7ca\Documents\scripts"
    } else {
        return "C:\Users\eduev\Meu Drive\17 - Projects\scripts"
    }
}

; Get list of script files to update.
; AppLaunchers.ahk must be last so QuickUpdateScripts can relaunch it with /Updated (success overlay; includes Utils.ahk).
GetScriptFiles() {
    scriptsDir := GetScriptsDirectory()
    return [
        scriptsDir "\WindowManagement.ahk",
        scriptsDir "\Spotify.ahk",
        scriptsDir "\Shift keys.ahk",
        scriptsDir "\Outlook.ahk",
        scriptsDir "\Microsoft Teams.ahk",
        scriptsDir "\Gemini.ahk",
        scriptsDir "\mousemaster.ahk",
        scriptsDir "\AppLaunchers.ahk"
    ]
}

; Quick Update Scripts macro: PowerShell handoff restart (local-only, no git).
QuickUpdateScripts() {
    static s_isQuickUpdateRunning := false
    if (s_isQuickUpdateRunning) {
        return
    }
    s_isQuickUpdateRunning := true

    try {
        ; Do not call ApplyScriptMasterVolumeTarget here - it only affects sessions that are about to be killed,
        ; then ExitApp cancels timers; new processes need volume applied after relaunch (see /Updated in Utils).
        scripts := GetScriptFiles()
        scriptsNames := ""
        for scriptPath in scripts {
            parts := StrSplit(scriptPath, "\")
            scriptsNames .= parts[parts.Length] . ";"
        }

        StandardLoadingBar_Show("⏳ Restarting scripts...", BANNER_ACCENT_INTERMEDIATE, { passive: false })

        anchorPath := ""
        psPaths := []

        for scriptPath in scripts {
            if (InStr(scriptPath, "\AppLaunchers.ahk"))
                anchorPath := scriptPath
            psPaths.Push("'" . StrReplace(scriptPath, "'", "''") . "'")
        }

        if (!anchorPath || anchorPath = "") {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ AppLaunchers.ahk not found in script list", 2500, BANNER_ACCENT_ERROR)
            return
        }

        psAnchor := StrReplace(anchorPath, "'", "''")

        ; PowerShell contract (deterministic): sleep 500ms -> stop AutoHotkey* -> sleep 500ms -> Start-Process scripts.
        ps := ""
        ps .= "Start-Sleep -Milliseconds 500; "
        ps .= "Stop-Process -Name 'AutoHotkey*' -Force -ErrorAction SilentlyContinue; "
        ps .= "Start-Sleep -Milliseconds 500; "
        ps .= "$anchorPath = '" . psAnchor . "'; "
        ; AHK Array.Join() isn't available in this environment; build a CSV manually.
        psPathsJoined := ""
        for i, p in psPaths {
            if (i > 1)
                psPathsJoined .= ","
            psPathsJoined .= p
        }
        ps .= "$scripts = @(" . psPathsJoined . "); "
        ; Stagger launches so earlier scripts begin loading before the next; AppLaunchers stays last with /Updated.
        ; Short pause after all Start-Process so audio sessions can register before the relaunched AppLaunchers runs its volume schedule.
        ps .=
            "foreach ($s in $scripts) { if ($s -eq $anchorPath) { Start-Process -FilePath $s -ArgumentList '/Updated' | Out-Null } else { Start-Process -FilePath $s | Out-Null }; Start-Sleep -Milliseconds 450 }; "
        ps .= "Start-Sleep -Seconds 2; "

        ; Execute asynchronously, then terminate this AHK instance immediately.
        pid := 0
        try {
            pid := Run('powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "' . ps . '"', , "Hide")
        } catch as e {
            throw
        }
        ExitApp
    } catch as e {
        try StandardLoadingBar_Hide(0)
        ShowCenteredOverlay_Utils("❌ QuickUpdateScripts error: " . e.Message, 3500, BANNER_ACCENT_ERROR)
    } finally {
        ; If we didn't ExitApp (error path), reset the guard.
        s_isQuickUpdateRunning := false
    }
}

; Update Gemini script specifically (local-only): restart Gemini.ahk.
UpdateGeminiScript() {
    scriptsDir := GetScriptsDirectory()
    geminiPath := scriptsDir "\Gemini.ahk"

    if (!FileExist(geminiPath)) {
        ShowCenteredOverlay_Utils("❌ Gemini.ahk not found at: " geminiPath, 3000, BANNER_ACCENT_ERROR)
        return
    }

    StandardLoadingBar_Show("⏳ Reloading Gemini...", BANNER_ACCENT_INTERMEDIATE, { passive: false })
    try {
        Run(geminiPath)
        StandardLoadingBar_Hide(0)
        ShowCenteredOverlay_Utils("✅ Gemini script updated!", 1500, BANNER_ACCENT_SUCCESS)
    } catch {
        StandardLoadingBar_Hide(0)
        ShowCenteredOverlay_Utils("❌ Failed to reload Gemini", 2000, BANNER_ACCENT_ERROR)
    }
}

; Add specific word to Handy macro function
AddWordToHandy() {
    targetPath := GetHandyShortcutPath()

    try {
        if WinExist("Handy ahk_class Tauri Window") {
            WinActivate
        } else {
            if (targetPath = "" || !FileExist(targetPath)) {
                MsgBox "Failed to launch Handy.`n`nShortcut not found.", "Utils.ahk", "IconX"
                return
            }
            Run targetPath
            if !WinWait("Handy ahk_class Tauri Window", , 5) {
                MsgBox "Failed to launch Handy."
                return
            }
        }

        if (!WinWaitActive("Handy ahk_class Tauri Window", , 2)) {
            ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        hwnd := WinExist("Handy ahk_class Tauri Window")

        ; Initialize UIA
        el := UIA.ElementFromHandle(hwnd)
        if !el
            return

        ; Locate and click "Advanced" - find Group element directly by Type and ClassName pattern
        ; The clickable target is the Group (50026) with ClassName containing "cursor-pointer" and "flex gap-2 items-center"
        advancedBtn := ""
        try {
            allGroups := el.FindAll({ Type: 50026 })
            if allGroups {
                for group in allGroups {
                    try {
                        groupClassName := group.ClassName
                        ; Look for Group with cursor-pointer and flex gap-2 items-center (Advanced button pattern)
                        if (InStr(groupClassName, "cursor-pointer") && InStr(groupClassName, "flex gap-2 items-center")) {
                            ; Verify it contains "Advanced" text by checking children
                            try {
                                advancedText := group.FindFirst({ Type: 50020, Name: "Advanced" })
                                if advancedText {
                                    advancedBtn := group
                                    break
                                }
                            } catch {
                            }
                        }
                    } catch {
                    }
                }
            }
        } catch {
        }

        ; Fallback: Find by Name "Advanced" and verify parent has correct ClassName
        if !advancedBtn {
            advancedElement := el.FindFirst({ Name: "Advanced" })
            if advancedElement {
                try {
                    parentGroup := advancedElement.GetParentElement()
                    ; Verify parent is a Group with cursor-pointer class (clickable)
                    if (parentGroup && parentGroup.Type = 50026 && InStr(parentGroup.ClassName, "cursor-pointer")) {
                        advancedBtn := parentGroup
                    }
                } catch {
                }
            }
        }

        if advancedBtn {
            try {
                advancedBtn.Click()
            } catch Error as clickErr {
                try {
                    advancedBtn.Invoke()
                } catch {
                }
            }
        }

        Sleep 200

        ; Locate and focus "Add a word" text field
        addWordEdit := el.FindFirst({ Type: 50004, Name: "Add a word" })
        if addWordEdit {
            addWordEdit.SetFocus()
        }
    } catch Error as e {
        MsgBox "Error in AddWordToHandy macro: " e.Message
    }
}

; Resolve Handy shortcut/executable path (environment-aware: work vs home)
; Work: uses Documents\Handy\handy.exe; Home: uses Start Menu shortcuts.
GetHandyShortcutPath() {
    global IS_WORK_ENVIRONMENT

    if (IS_WORK_ENVIRONMENT) {
        ; Work: direct exe path first, then work shortcut fallback
        workExe := "C:\Users\fie7ca\Documents\Handy\handy.exe"
        if (FileExist(workExe))
            return workExe
        for , p in ["C:\Users\fie7ca\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Handy\Handy.lnk",
            "C:\Users\fie7ca\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Handy.lnk"] {
            if (p != "" && FileExist(p))
                return p
        }
        return ""
    }

    ; Home/personal: Start Menu shortcuts only
    candidates := [
        "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Handy\Handy.lnk",
        "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Handy\Handy.lnk",
        "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Handy.lnk"
    ]
    for , p in candidates {
        try {
            if (p != "" && FileExist(p))
                return p
        } catch {
        }
    }
    return ""
}

; Expected Handy process exe path for current environment (for targeting correct instance).
; Work: work exe path; Home: "" (any Handy window).
GetHandyProcessPath() {
    global IS_WORK_ENVIRONMENT
    if (IS_WORK_ENVIRONMENT) {
        workExe := "C:\Users\fie7ca\Documents\Handy\handy.exe"
        return FileExist(workExe) ? workExe : ""
    }
    return ""
}

; =============================================================================
; Clip Angel: Merge Non-Favorite Clips
; =============================================================================
; Ensure Clip Angel window is closed (Alt+V then WinClose fallback)
EnsureClipAngelClosed() {
    if !WinExist("ClipAngel")
        return
    Send "!v"
    Sleep 400
    if WinExist("ClipAngel") {
        Send "!v"
        Sleep 300
    }
    if WinExist("ClipAngel") {
        try WinClose("ClipAngel")
    }
}

; Extract title from first non-favorite clip in ClipAngel
MergeNonFavoriteClips() {
    try {
        ; Show persistent banner for the duration of the algorithm
        AiModelBanner_Show("📋 Merging non-favorite clips...", "FFCC00")

        ; Step 1: Send Alt+B to activate ClipAngel (this opens the window if not visible)
        Send "!b"
        Sleep 500  ; Wait for ClipAngel window to appear

        ; Step 2: Check if ClipAngel window exists now
        if !WinExist("ClipAngel") {
            AiModelBanner_Hide()
            MsgBox "ClipAngel window did not appear. Make sure ClipAngel is running.", "Merge Clips", "IconX"
            return
        }
        try {
            WinActivate("ClipAngel")
        } catch {
            AiModelBanner_Hide()
            ShowCenteredOverlay_Utils("❌ ClipAngel window not found.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        WinWaitActive("ClipAngel", , 2)

        ; Step 3: Initialize UIA on ClipAngel window
        hwnd := WinExist("A")
        el := UIA.ElementFromHandle(hwnd)
        if !el {
            AiModelBanner_Hide()
            EnsureClipAngelClosed()
            MsgBox "Failed to initialize UIA for ClipAngel.", "Merge Clips", "IconX"
            return
        }

        ; Step 4: Find DataGridView by AutomationId
        dataGrid := 0
        try dataGrid := el.FindFirst({ Type: 50036, AutomationId: "dataGridView" })
        if !dataGrid {
            AiModelBanner_Hide()
            EnsureClipAngelClosed()
            MsgBox "Could not find DataGridView in ClipAngel.", "Merge Clips", "IconX"
            return
        }

        ; Step 5: Find Row 0
        row0 := 0
        try row0 := dataGrid.FindFirst({ Type: 50025, Name: "Row 0" })
        if !row0 {
            AiModelBanner_Hide()
            EnsureClipAngelClosed()
            MsgBox "No clips found in Row 0.", "Merge Clips", "IconX"
            return
        }

        ; Step 6: Find "Title Row 0" element
        titleElement := 0
        try titleElement := row0.FindFirst({ Type: 50006, Name: "Title Row 0" })
        if !titleElement {
            AiModelBanner_Hide()
            EnsureClipAngelClosed()
            MsgBox "Could not find Title element in Row 0.", "Merge Clips", "IconX"
            return
        }

        ; Step 7: Extract RTF value
        rtfValue := ""
        try rtfValue := titleElement.Value
        if (rtfValue = "" || rtfValue = "System.Drawing.Bitmap") {
            AiModelBanner_Hide()
            EnsureClipAngelClosed()
            MsgBox "Title Row 0 contains no text data.", "Merge Clips", "IconX"
            return
        }

        ; Step 8: Parse RTF to extract plain text (this is our target favorite clip title)
        favoriteClipTitle := ParseRTFToPlainText(rtfValue)

        ; Step 9: Switch to "All Clips" view to search for this favorite clip
        Send "!b"  ; Close current view
        Sleep 600
        Send "!v"  ; Open "All Clips" view (non-favorites first, favorites second)
        Sleep 500  ; Wait for view to update

        ; Step 10: Re-initialize UIA for the updated view
        hwnd := WinExist("A")
        el := UIA.ElementFromHandle(hwnd)
        if !el {
            AiModelBanner_Hide()
            EnsureClipAngelClosed()
            MsgBox "Failed to re-initialize UIA after switching views.", "Merge Clips", "IconX"
            return
        }

        ; Step 11: Find DataGridView again
        dataGrid := 0
        try dataGrid := el.FindFirst({ Type: 50036, AutomationId: "dataGridView" })
        if !dataGrid {
            AiModelBanner_Hide()
            EnsureClipAngelClosed()
            MsgBox "Could not find DataGridView in All Clips view.", "Merge Clips", "IconX"
            return
        }

        ; Step 12: Focus on Row 0 to start the search
        try {
            row0 := dataGrid.FindFirst({ Type: 50025, Name: "Row 0" })
            if row0 {
                row0.SetFocus()
                Sleep 100
            }
        }

        ; Step 13: Iterative search through rows (max 40 iterations)
        maxIterations := 40
        foundMatch := false
        currentRow := 0

        loop maxIterations {
            currentRow := A_Index - 1  ; 0-based row index

            ; Find current row
            currentRowElement := 0
            try currentRowElement := dataGrid.FindFirst({ Type: 50025, Name: "Row " . currentRow })

            if !currentRowElement {
                ; No more rows, stop searching
                break
            }

            ; Find title element in current row
            currentTitleElement := 0
            try currentTitleElement := currentRowElement.FindFirst({ Type: 50006, Name: "Title Row " . currentRow })

            if currentTitleElement {
                ; Extract and parse the title
                currentRtfValue := ""
                try currentRtfValue := currentTitleElement.Value

                if (currentRtfValue != "" && currentRtfValue != "System.Drawing.Bitmap") {
                    currentTitle := ParseRTFToPlainText(currentRtfValue)

                    ; Compare with the favorite clip title
                    if (currentTitle = favoriteClipTitle) {
                        foundMatch := true

                        ; Step 14: Select and merge non-favorite clips (cursor is on first favorite)
                        ; Move up once to last non-favorite clip
                        Send "{Up}"
                        Sleep 150
                        ; Select from current position to top of list (all non-favorites)
                        Send "^+{Home}"
                        Sleep 150
                        ; Merge the selected clips
                        Send "^!j"
                        Sleep 300  ; Wait for merge to complete

                        ; Step 15: Copy merged clip to clipboard
                        Send "{Tab}"   ; Focus merged content area
                        Sleep 150
                        Send "^a"     ; Select all
                        Sleep 100
                        Send "^c"     ; Copy

                        AiModelBanner_Hide()
                        ShowCenteredOverlay_Utils("✅ Merged non-favorite clips (copied)", 2000, BANNER_ACCENT_SUCCESS)
                        break
                    }
                }
            }

            ; Move to next row
            Send "{Down}"
            Sleep 100  ; Small delay between iterations
        }

        if !foundMatch {
            AiModelBanner_Hide()
            ShowCenteredOverlay_Utils("⚠ Favorite clip not found in first " . maxIterations . " rows", 2000,
                BANNER_ACCENT_INTERMEDIATE)
        }

        ; Guarantee Clip Angel is closed when macro finishes (success or not found)
        EnsureClipAngelClosed()

    } catch Error as e {
        AiModelBanner_Hide()
        EnsureClipAngelClosed()
        MsgBox "Error in MergeNonFavoriteClips: " . e.Message, "Merge Clips", "IconX"
    }
}

; Helper function to parse RTF and extract plain text
ParseRTFToPlainText(rtf) {
    ; Remove RTF header and formatting
    ; Pattern: extract text between last formatting and \par
    plainText := rtf

    ; Remove RTF control sequences (backslash followed by letters/numbers)
    plainText := RegExReplace(plainText, "\\[a-z]+[0-9]*\s?", "")
    ; Remove braces
    plainText := RegExReplace(plainText, "[{}]", "")
    ; Remove everything after \par
    plainText := RegExReplace(plainText, "\\par.*$", "")

    ; Trim whitespace
    plainText := Trim(plainText)

    return plainText
}

; =============================================================================
; Clip Angel: Open/Activate with focus correction (Row 0)
; =============================================================================
; Alt+V: Activate Clip Angel and ensure focus is on "Row 0" (fixes bug where focus
; defaults to upper tabs). Uses UIA: Type 50025, Name "Row 0" per clipangel-tree.txt.
ActivateClipAngelWithFocusCorrection() {
    needBanner := false
    if WinExist("ClipAngel") {
        try {
            WinActivate("ClipAngel")
        } catch {
            ShowCenteredOverlay_Utils("❌ ClipAngel window not found.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        WinWaitActive("ClipAngel", , 2)
    } else {
        needBanner := true
        ClipAngelBanner_Show("📂 Opening Clip Angel...", BANNER_ACCENT_INTERMEDIATE)
        Send "!v"
        if !WinWait("ClipAngel", , 10) {
            ClipAngelBanner_Hide()
            return
        }
        try {
            WinActivate("ClipAngel")
        } catch {
            ClipAngelBanner_Hide()
            ShowCenteredOverlay_Utils("❌ ClipAngel window not found.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        WinWaitActive("ClipAngel", , 2)
    }
    Sleep 50
    ; Must use Clip Angel's HWND, not WinExist("A") - another app can be foreground and UIA targets the wrong tree.
    hwnd := WinExist("ClipAngel")
    if !hwnd {
        if needBanner
            ClipAngelBanner_Hide()
        return
    }
    el := UIA.ElementFromHandle(hwnd)
    if !el {
        if needBanner
            ClipAngelBanner_Hide()
        return
    }
    try {
        dataGrid := ClipAngel_UiaFindFirst(el, { Type: 50036, AutomationId: "dataGridView" })
        if !dataGrid {
            if needBanner
                ClipAngelBanner_Hide()
            return
        }
        row0 := ClipAngel_UiaFindFirst(dataGrid, { Type: 50025, Name: "Row 0" })
        if !row0 {
            if needBanner
                ClipAngelBanner_Hide()
            return
        }
        hasSel := row0.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable)
        isSelected := hasSel && row0.SelectionItemPattern.IsSelected
        if (!isSelected) {
            if !needBanner
                ClipAngelBanner_Show("🎯 Focusing Row 0...", BANNER_ACCENT_INTERMEDIATE)
            needBanner := true
            try {
                if hasSel
                    row0.SelectionItemPattern.Select()
                else
                    row0.SetFocus()
            } catch {
                try row0.SetFocus()
            }
        }
    } catch {
        if needBanner
            ClipAngelBanner_Hide()
        return
    }
    if needBanner {
        ClipAngelBanner_Show("✅ Done", BANNER_ACCENT_SUCCESS)
        SetTimer(ClipAngelBanner_Hide, -500)
    }
}

; =============================================================================
; Clip Angel: Mark Last Clip as Favorite
; =============================================================================
; Wait after clipboard change before favoriting newest clip (copy / dictation ingest).
CLIPANGEL_PRE_FAVORITE_INGEST_DELAY_MS := 400
; Settle after row focus, before Alt+Q (all favorite paths).
CLIPANGEL_FAVORITE_UI_SETTLE_MS := 50
; Shortcut flow (matches app): open Clip Angel, ensure list focus (not Window tab),
; select first or last grid row, Send Alt+Q. Optional: target "last" for bottom row.
; UIA-v2 FindFirst throws TargetError when nothing matches - never chain with if !c without try.
ClipAngel_UiaFindFirst(root, conditions) {
    if !root
        return 0
    try return root.FindFirst(conditions)
    catch
        return 0
}

ClipAngel_FindFavoriteCell(row) {
    if !row
        return 0
    rn := ""
    try rn := row.Name
    catch {
        rn := ""
    }
    suffix := "0"
    if RegExMatch(rn, "i)(?:Row|Linha)\s*(\d+)", &m)
        suffix := m[1]
    else if RegExMatch(rn, "(\d+)\s*$", &m)
        suffix := m[1]
    ; EN + PT-BR column headers seen in Clip Angel / localized WinForms.
    for cand in [
        "Favorite Row " . suffix, "Favourite Row " . suffix, "Favorito Row " . suffix,
        "Favorite Linha " . suffix, "Favorito Linha " . suffix
    ] {
        c := ClipAngel_UiaFindFirst(row, { Type: UIA.Type.CheckBox, Name: cand })
        if c
            return c
        c := ClipAngel_UiaFindFirst(row, { Type: 50002, Name: cand })
        if c
            return c
    }
    try {
        for c in row.FindAll({ Type: 50002 }) {
            try n := c.Name
            catch
                continue
            if RegExMatch(n, "i)favorite|favourite|favorito")
                return c
        }
        boxes := row.FindAll({ Type: 50002 })
        if boxes.Length >= 2
            return boxes[boxes.Length]
    } catch {
    }
    return 0
}

ClipAngel_FavoriteCellIsOn(cell) {
    if !cell
        return false
    try {
        if cell.GetPropertyValue(UIA.Property.IsTogglePatternAvailable)
            return cell.TogglePattern.ToggleState = UIA.ToggleState.On
        ts := cell.GetPropertyValue(UIA.Property.ToggleToggleState)
        if ts != ""
            return ts = UIA.ToggleState.On
    } catch {
    }
    ; Value only for read-only grid cells - Legacy CHECKED (0x10) often false-positives on DataGrid cells.
    try {
        v := cell.Value
        if (v = "true" || v = "True" || v = "1")
            return true
    } catch {
    }
    return false
}

ClipAngel_MainHwnd() {
    h := WinExist("ClipAngel")
    if h
        return h
    return WinExist("ahk_exe ClipAngel.exe")
}

; Macro hotkeys use Ctrl+Alt+Win - if those keys are still down, Send "!q" is not plain Alt+Q (Win+Alt+... hijacks it).
ClipAngel_ReleaseChordModifiersForSend() {
    SendInput "{LWin up}{RWin up}{LControl up}{RControl up}{LAlt up}{RAlt up}{LShift up}{RShift up}"
}

; Wait for physical release (KeyWait) then synthetic up - chord hotkeys often leave keys logically down.
ClipAngel_WaitChordModifiersReleased() {
    tw := "T0.45"
    KeyWait "Ctrl", tw
    KeyWait "Alt", tw
    KeyWait "Shift", tw
    KeyWait "LWin", tw
    KeyWait "RWin", tw
}

; target: "first" = top grid row (Row 0 / newest), "last" = last row returned by UIA FindAll
; (virtualized lists may only expose visible rows - use "first" for reliable top-clip behavior).
MarkLastClipAsFavorite(target := "first", waitForIngest := false) {
    if waitForIngest
        Sleep(CLIPANGEL_PRE_FAVORITE_INGEST_DELAY_MS)
    ActivateClipAngelWithFocusCorrection()
    hwnd := ClipAngel_MainHwnd()
    if !hwnd {
        ShowCenteredOverlay_Utils("❌ Clip Angel did not open.", 2000, BANNER_ACCENT_ERROR)
        return
    }
    try {
        try WinActivate("ahk_id " hwnd)
        catch {
            ShowCenteredOverlay_Utils("❌ Clip Angel window not found.", 2000, BANNER_ACCENT_ERROR)
            EnsureClipAngelClosed()
            return
        }
        if !WinWaitActive("ahk_id " hwnd, , 2) {
            ShowCenteredOverlay_Utils("❌ Clip Angel did not become active.", 2000, BANNER_ACCENT_ERROR)
            EnsureClipAngelClosed()
            return
        }
        el := UIA.ElementFromHandle(hwnd)
        if !el {
            ShowCenteredOverlay_Utils("❌ Clip Angel UI not available.", 2000, BANNER_ACCENT_ERROR)
            EnsureClipAngelClosed()
            return
        }
        dataGrid := ClipAngel_UiaFindFirst(el, { Type: 50036, AutomationId: "dataGridView" })
        if !dataGrid {
            ShowCenteredOverlay_Utils("❌ Clip list not found (Window tab may still have focus).", 2500,
                BANNER_ACCENT_ERROR)
            EnsureClipAngelClosed()
            return
        }
        rows := 0
        try rows := dataGrid.FindAll({ Type: 50025 })
        catch {
            rows := 0
        }
        if !rows || rows.Length < 1 {
            ShowCenteredOverlay_Utils("❌ No clips in list.", 2000, BANNER_ACCENT_ERROR)
            EnsureClipAngelClosed()
            return
        }
        rowTarget := rows[1]
        if (target = "last")
            rowTarget := rows[rows.Length]
        hasSel := rowTarget.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable)
        try {
            if hasSel
                rowTarget.SelectionItemPattern.Select()
            else
                rowTarget.SetFocus()
        } catch {
            try rowTarget.SetFocus()
        }
        try {
            if rowTarget.GetPropertyValue(UIA.Property.IsScrollItemPatternAvailable)
                rowTarget.ScrollItemPattern.ScrollIntoView()
        } catch {
        }
        favCell := ClipAngel_FindFavoriteCell(rowTarget)
        if favCell {
            if ClipAngel_FavoriteCellIsOn(favCell) {
                ShowCenteredOverlay_Utils("✅ Selected clip is already a favorite.", 1500, BANNER_ACCENT_SUCCESS)
                EnsureClipAngelClosed()
                return
            }
        }
        if !WinActive("ahk_id " hwnd) {
            try WinActivate("ahk_id " hwnd)
            if !WinWaitActive("ahk_id " hwnd, , 2) {
                ShowCenteredOverlay_Utils("❌ Clip Angel lost focus before Alt+Q.", 2000, BANNER_ACCENT_ERROR)
                EnsureClipAngelClosed()
                return
            }
        }
        Sleep(CLIPANGEL_FAVORITE_UI_SETTLE_MS)
        ClipAngel_WaitChordModifiersReleased()
        ClipAngel_ReleaseChordModifiersForSend()
        SendInput "!q"
        ScriptSoundPlay(A_ScriptDir "\sounds\favorite-set.wav")
        ShowCenteredOverlay_Utils("✅ Sent Alt+Q - marked focused clip as favorite.", 1500, BANNER_ACCENT_SUCCESS)
        EnsureClipAngelClosed()
    } catch Error as e {
        ShowCenteredOverlay_Utils("❌ Mark favorite failed: " . e.Message, 2500, BANNER_ACCENT_ERROR)
        EnsureClipAngelClosed()
    }
}

; =============================================================================
; AI Model Selection System for Handy
; =============================================================================
; Configuration: Maps selection numbers (1-4) to AI model names.
; These are partial name prefixes used to find buttons in the UIA tree (Type 50000, botão).
; Descriptions match Handy Transcription Models UI for quick verification.
; Slots 3-4: set Cohere Language on General tab before selecting the Cohere model (modelClickName).
global g_HandyAiModels := Map(
    1, { name: "Parakeet V2", desc: "English only. Best model for English speakers." },
    2, { name: "Parakeet V3", desc: "Fast and accurate. Multi-language." },
    3, { name: "Cohere English", desc: "Sets Cohere language to English (General), then activates Cohere.",
        cohereLanguage: "English", modelClickName: "Cohere" },
    4, { name: "Cohere Portuguese", desc: "Sets Cohere language to Portuguese (General), then activates Cohere.",
        cohereLanguage: "Portuguese", modelClickName: "Cohere" }
)

; Picker indices for ^!#9 / ^!#b; update g_HandyAiModels names if Handy renames models.
global HANDY_AI_SLOT_COHERE_PORTUGUESE := 4
global HANDY_AI_SLOT_COHERE_ENGLISH := 3

; GUI state for AI model selector
global g_AiModelSelectorGui := false
global g_AiModelSelectorActive := false
global g_AiModelBannerGui := false

; Persistent language flag indicator (slot 3 = UK, slot 4 = Brazil); see docs/standard_information_display.md "Persistent Indicators".
global g_LanguageFlagGuis := []
global g_LanguageFlagSlot := 0
global LANGUAGE_FLAG_WIDTH := 45                ; px (~30% smaller than 64); aspect kept via Picture h:-1
global LANGUAGE_FLAG_MARGIN := 20               ; px from work-area right/bottom

; Restore the persistent flag on script load (Reload-safe). Deferred so the GUI
; subsystem is ready and any concurrent auto-execute side-effects settle first.
SetTimer(LanguageFlag_InitFromPersistedSlot, -250)

Handy_GetHandyAiModelIniPath() {
    return A_ScriptDir "\data\handy_ai_model.ini"
}

; Returns persisted slot 1-4, or 0 if missing / invalid / not in g_HandyAiModels.
Handy_GetPersistedAiModelSlot() {
    global g_HandyAiModels
    path := Handy_GetHandyAiModelIniPath()
    s := ""
    try s := IniRead(path, "Handy", "Slot", "")
    if (s = "")
        return 0
    if !IsInteger(s)
        return 0
    n := Integer(s)
    if !g_HandyAiModels.Has(n)
        return 0
    return n
}

Handy_SetPersistedAiModelSlot(slot) {
    global g_HandyAiModels
    if !g_HandyAiModels.Has(slot)
        return
    path := Handy_GetHandyAiModelIniPath()
    try {
        dataDir := A_ScriptDir "\data"
        if !DirExist(dataDir)
            DirCreate(dataDir)
        IniWrite(String(slot), path, "Handy", "Slot")
    } catch {
    }
}

; =============================================================================
; ShowAiModelSelector() - Display selection GUI with immediate key capture
; =============================================================================
ShowAiModelSelector() {
    global g_AiModelSelectorGui, g_AiModelSelectorActive, g_HandyAiModels

    ; Don't show if already active
    if (g_AiModelSelectorActive)
        return

    currentSlot := Handy_GetPersistedAiModelSlot()

    ; Create selection GUI
    g_AiModelSelectorGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Owner")
    g_AiModelSelectorGui.BackColor := "1E1E2E"
    g_AiModelSelectorGui.MarginX := 20
    g_AiModelSelectorGui.MarginY := 15

    ; Title
    g_AiModelSelectorGui.SetFont("s14 cCDD6F4 Bold", "Segoe UI")
    g_AiModelSelectorGui.Add("Text", "w280 Center", "🎙️ Select AI Model")
    g_AiModelSelectorGui.Add("Text", "w280 h1 Background45475A")  ; separator

    ; Model options (green row = last saved slot from data\handy_ai_model.ini)
    g_AiModelSelectorGui.SetFont("s12 cCDD6F4", "Segoe UI")
    for num, model in g_HandyAiModels {
        if (num = currentSlot) {
            g_AiModelSelectorGui.SetFont("s12 cA6E3A1 Bold", "Segoe UI")
            g_AiModelSelectorGui.Add("Text", "w280 Background313244", "[" . num . "] " . model.name)
            g_AiModelSelectorGui.SetFont("s9 cA6E3A1", "Segoe UI")
            g_AiModelSelectorGui.Add("Text", "w280 y+2 Background313244", "    " . model.desc)
        } else {
            g_AiModelSelectorGui.SetFont("s12 cCDD6F4", "Segoe UI")
            g_AiModelSelectorGui.Add("Text", "w280", "[" . num . "] " . model.name)
            g_AiModelSelectorGui.SetFont("s9 c6C7086", "Segoe UI")
            g_AiModelSelectorGui.Add("Text", "w280 y+2", "    " . model.desc)
        }
        g_AiModelSelectorGui.SetFont("s12 cCDD6F4", "Segoe UI")
    }

    ; Footer
    g_AiModelSelectorGui.Add("Text", "w280 h1 Background45475A y+10")
    g_AiModelSelectorGui.SetFont("s9 c6C7086", "Segoe UI")
    g_AiModelSelectorGui.Add("Text", "w280 Center", "Green row = last saved model")
    g_AiModelSelectorGui.Add("Text", "w280 Center y+4", "Press 1-4 | Esc to cancel")

    ; Get active window to determine which monitor to center on
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

    ; Measure GUI size and center on the active monitor
    g_AiModelSelectorGui.Show("AutoSize Hide")
    g_AiModelSelectorGui.GetPos(&gx, &gy, &gw, &gh)
    cx := monitorLeft + (monitorWidth - gw) // 2
    cy := monitorTop + (monitorHeight - gh) // 2
    g_AiModelSelectorGui.Show("x" . cx . " y" . cy . " NA")

    g_AiModelSelectorActive := true

    ; Enable hotkeys for 1-4 and Escape
    Hotkey("1", AiModelSelector_HandleKey, "On")
    Hotkey("2", AiModelSelector_HandleKey, "On")
    Hotkey("3", AiModelSelector_HandleKey, "On")
    Hotkey("4", AiModelSelector_HandleKey, "On")
    Hotkey("Escape", AiModelSelector_Cancel, "On")
}

; Handle key press in AI model selector
AiModelSelector_HandleKey(key) {
    global g_AiModelSelectorGui, g_AiModelSelectorActive, g_HandyAiModels

    if (!g_AiModelSelectorActive)
        return

    ; Get the selection number
    selection := Integer(key)

    ; Close selector GUI
    AiModelSelector_Close()

    ; Execute the selection
    if (g_HandyAiModels.Has(selection)) {
        ExecuteHandyAiModelSelection(selection)
    }
}

; Cancel AI model selector
AiModelSelector_Cancel(*) {
    AiModelSelector_Close()
}

; Close the selector GUI and disable hotkeys
AiModelSelector_Close() {
    global g_AiModelSelectorGui, g_AiModelSelectorActive

    if (!g_AiModelSelectorActive)
        return

    g_AiModelSelectorActive := false

    ; Disable hotkeys
    try Hotkey("1", "Off")
    try Hotkey("2", "Off")
    try Hotkey("3", "Off")
    try Hotkey("4", "Off")
    try Hotkey("Escape", AiModelSelector_Cancel, "Off")
    Utils_EnsureGlobalEscapeHotkey()

    ; Destroy GUI
    if (IsObject(g_AiModelSelectorGui) && g_AiModelSelectorGui.Hwnd) {
        try g_AiModelSelectorGui.Destroy()
    }
    g_AiModelSelectorGui := false
}

; =============================================================================
; Gemini-to-Cursor transfer: numeric Cursor window selector (1-9) and activate/focus/paste
; Used when user presses [C] Transfer in Gemini copy-decision banner.
; =============================================================================
global g_CursorTransferSelectorGui := false
global g_CursorTransferSelectorActive := false
global g_CursorTransferSelectorResult := ""   ; "" = waiting, 0 = cancel, integer = selected hwnd
global g_CursorTransferWindowList := []      ; up to 9 { hwnd, title }
global g_CursorTransferHotkeyHandlers := []
global g_CursorTransferPidCmdCache := Map()
global g_CursorTransferLastHandledIndex := 0 ; Diagnostic: track last selected index

; === Environment-aware transfer target resolution ===

; Return transfer target app executable based on IS_WORK_ENVIRONMENT: "Cursor.exe" or "Code.exe"
CursorTransfer_GetTargetAppExecutable() {
    if (IsSet(IS_WORK_ENVIRONMENT) && IS_WORK_ENVIRONMENT)
        return "Code.exe"
    return "Cursor.exe"
}

; Return display name for transfer target app: "Cursor" or "VS Code"
CursorTransfer_GetTargetAppName() {
    if (IsSet(IS_WORK_ENVIRONMENT) && IS_WORK_ENVIRONMENT)
        return "VS Code"
    return "Cursor"
}

CursorTransfer_SelectorClose() {
    global g_CursorTransferSelectorActive, g_CursorTransferSelectorGui, g_CursorTransferHotkeyHandlers
    if (!g_CursorTransferSelectorActive)
        return
    g_CursorTransferSelectorActive := false
    for h in g_CursorTransferHotkeyHandlers {
        if (IsObject(h) && h.HasProp("key") && h.HasProp("callback")) {
            try Hotkey(h.key, h.callback, "Off")
        } else {
            try Hotkey(h, "Off")
        }
    }
    g_CursorTransferHotkeyHandlers := []
    if (IsObject(g_CursorTransferSelectorGui) && g_CursorTransferSelectorGui.Hwnd) {
        try g_CursorTransferSelectorGui.Destroy()
    }
    g_CursorTransferSelectorGui := false
}

CursorTransfer_SelectorHandleKey(index, *) {
    global g_CursorTransferSelectorResult, g_CursorTransferWindowList, g_CursorTransferLastHandledIndex

    ; Validate index bounds
    if (!IsInteger(index) || index < 1 || index > g_CursorTransferWindowList.Length) {
        ; Invalid index: keep selector open, log for diagnostics
        g_CursorTransferLastHandledIndex := index
        return
    }

    ; Retrieve target hwnd
    targetItem := g_CursorTransferWindowList[index]
    if (!targetItem || !targetItem.HasProp("hwnd")) {
        g_CursorTransferLastHandledIndex := index
        return
    }

    targetHwnd := targetItem.hwnd

    ; Verify hwnd still exists before accepting selection
    if (!WinExist("ahk_id " targetHwnd)) {
        g_CursorTransferLastHandledIndex := index
        return
    }

    ; Valid selection: set result and close
    g_CursorTransferSelectorResult := targetHwnd
    g_CursorTransferLastHandledIndex := index
    CursorTransfer_SelectorClose()
}

CursorTransfer_SelectorEscape(*) {
    global g_CursorTransferSelectorResult
    g_CursorTransferSelectorResult := 0
    CursorTransfer_SelectorClose()
}

; Return project order index from g_Projects for a window title; 0 = no match.
CursorTransfer_GetProjectOrderForTitle(winTitle) {
    idx := CursorTransfer_GetMatchingProjectIndexForTitle(winTitle)
    return idx
}

; Build canonical project-index -> character mapping (same logic as standard project selector).
CursorTransfer_BuildProjectIndexToChar() {
    global g_Projects, g_ProjectCategories, g_ProjectCharSequence
    projectIndexToChar := Map()
    projectIndexToCategory := Map()
    loop g_Projects.Length {
        projectIndex := A_Index
        project := g_Projects[projectIndex]
        category := project.HasProp("category") ? project.category : "Personal"
        projectIndexToCategory[projectIndex] := category
    }
    charIndex := 1
    for category in g_ProjectCategories {
        categoryProjectIndices := []
        for projectIndex, cat in projectIndexToCategory {
            if (cat = category)
                categoryProjectIndices.Push(projectIndex)
        }
        for projectIndex in categoryProjectIndices {
            project := g_Projects[projectIndex]
            if (project.name = "" && project.path = "" && project.workPath = "") {
                charIndex++
                continue
            }
            if (charIndex > g_ProjectCharSequence.Length)
                break
            char := g_ProjectCharSequence[charIndex]
            ; Keep parity with standard selector where 3 is reserved.
            if (char = "3") {
                charIndex++
                if (charIndex > g_ProjectCharSequence.Length)
                    break
                char := g_ProjectCharSequence[charIndex]
            }
            projectIndexToChar[projectIndex] := char
            charIndex++
        }
    }
    return projectIndexToChar
}

; Return matching project index from g_Projects for a window title; 0 = no match.
; Uses longest matching path segment so "user-scripts" wins over "scripts" when both match.
CursorTransfer_GetMatchingProjectIndexForTitle(winTitle) {
    global g_Projects
    if (!winTitle || !IsObject(g_Projects))
        return 0
    try {
        isWork := (IsSet(IS_WORK_ENVIRONMENT) && IS_WORK_ENVIRONMENT)
        winLow := StrLower(winTitle)
        bestIdx := 0
        bestScore := 0
        loop g_Projects.Length {
            project := g_Projects[A_Index]
            if (project.name = "" && project.path = "" && project.workPath = "")
                continue
            projectPath := isWork ? project.workPath : project.path
            if (isWork && (projectPath = ""))
                projectPath := project.path
            if (projectPath = "")
                continue
            matchSegments := ExtractProjectMatchSegments(projectPath)
            for segment in matchSegments {
                if (segment = "")
                    continue
                if (InStr(winLow, StrLower(segment))) {
                    len := StrLen(segment)
                    if (len > bestScore) {
                        bestScore := len
                        bestIdx := A_Index
                    }
                }
            }
        }
        return bestIdx
    } catch {
    }
    return 0
}

; Return project path according to current environment for a given project object.
CursorTransfer_GetEffectiveProjectPath(project) {
    isWork := (IsSet(IS_WORK_ENVIRONMENT) && IS_WORK_ENVIRONMENT)
    projectPath := isWork ? project.workPath : project.path
    if (isWork && (projectPath = ""))
        projectPath := project.path
    return projectPath
}

; True when the OS title has no workspace/file segment (e.g. bare "Cursor" welcome screen).
CursorTransfer_IsUninformativeCursorTitle(winTitle) {
    t := Trim(winTitle)
    if (t = "")
        return true
    tl := StrLower(t)
    if (tl = "cursor")
        return true
    ; Strip trailing " - Cursor"; if nothing remains, title had no file/workspace name.
    rest := RegExReplace(tl, "\s*[---]\s*cursor\s*$", "")
    if (Trim(rest) = "")
        return true
    return false
}

; Longest project path wins when multiple g_Projects paths appear in the same command line.
CursorTransfer_GetMatchingProjectIndexByCmdLine(cmdLine) {
    global g_Projects
    if (!cmdLine || !IsObject(g_Projects))
        return 0
    cmdLow := StrLower(cmdLine)
    bestIdx := 0
    bestLen := 0
    try {
        loop g_Projects.Length {
            project := g_Projects[A_Index]
            if (project.name = "" && project.path = "" && project.workPath = "")
                continue
            projectPath := CursorTransfer_GetEffectiveProjectPath(project)
            if (projectPath = "")
                continue
            projLow := StrLower(RTrim(projectPath, "\"))
            if (projLow != "" && InStr(cmdLow, projLow)) {
                len := StrLen(projLow)
                if (len > bestLen) {
                    bestLen := len
                    bestIdx := A_Index
                }
            }
        }
    } catch {
    }
    return bestIdx
}

; Get process command line by PID (cached). Returns "" on failure.
CursorTransfer_GetProcessCommandLine(pid) {
    global g_CursorTransferPidCmdCache
    if (!pid)
        return ""
    if (g_CursorTransferPidCmdCache.Has(pid))
        return g_CursorTransferPidCmdCache[pid]
    cmd := ""
    try {
        locator := ComObject("WbemScripting.SWbemLocator")
        svc := locator.ConnectServer(".", "root\cimv2")
        for proc in svc.ExecQuery("SELECT CommandLine FROM Win32_Process WHERE ProcessId = " pid) {
            try cmd := proc.CommandLine
            break
        }
    } catch {
        cmd := ""
    }
    g_CursorTransferPidCmdCache[pid] := cmd ? cmd : ""
    return g_CursorTransferPidCmdCache[pid]
}

; Title match when the title lists workspace/file; cmd-line when title is bare "Cursor" etc.
; (Same PID often shares one command line - use longest path in cmd for disambiguation.)
CursorTransfer_GetMatchingProjectIndex(hwnd, winTitle := "") {
    global g_Projects
    if (!IsObject(g_Projects))
        return 0
    pid := 0
    try pid := WinGetPID("ahk_id " hwnd)
    cmdLine := CursorTransfer_GetProcessCommandLine(pid)
    idxCmd := CursorTransfer_GetMatchingProjectIndexByCmdLine(cmdLine)
    idxTitle := (winTitle != "") ? CursorTransfer_GetMatchingProjectIndexForTitle(winTitle) : 0
    if (CursorTransfer_IsUninformativeCursorTitle(winTitle)) {
        if (idxCmd > 0)
            return idxCmd
        return idxTitle
    }
    if (idxTitle > 0)
        return idxTitle
    return idxCmd
}

; Stable insertion sort for small arrays by project order, then by title.
CursorTransfer_SortWindowsByProjectOrder(&arr) {
    n := arr.Length
    if (n <= 1)
        return
    loop n - 1 {
        i := A_Index + 1
        key := arr[i]
        j := i - 1
        while (j >= 1) {
            left := arr[j]
            shouldShift := false
            ; Coerce to integer so comparison never sees a string (projectOrder can be string from g_Projects).
            try
                leftOrd := Integer(left.projectOrder)
            catch
                leftOrd := 0
            try
                keyOrd := Integer(key.projectOrder)
            catch
                keyOrd := 0
            if (leftOrd > keyOrd) {
                shouldShift := true
            } else if (leftOrd = keyOrd) {
                leftName := ""
                keyName := ""
                try
                    leftName := String(left.displayName)
                catch
                    leftName := ""
                try
                    keyName := String(key.displayName)
                catch
                    keyName := ""
                if (StrCompare(leftName, keyName) > 0)
                    shouldShift := true
            }
            if (!shouldShift)
                break
            arr[j + 1] := left
            j--
        }
        arr[j + 1] := key
    }
}

; Return project name from g_Projects if window title matches a project path; otherwise "".
CursorTransfer_GetProjectNameForTitle(winTitle) {
    global g_Projects
    idx := CursorTransfer_GetMatchingProjectIndexForTitle(winTitle)
    if (!idx || !IsObject(g_Projects))
        return ""
    try {
        project := g_Projects[idx]
        if (project.name != "")
            return project.name
    } catch {
    }
    return ""
}

CursorTransfer_StripStaticScriptTokenForDisplay(projectName) {
    ; Removes redundant static token(s) like "Script"/"Scripts" from the project label.
    ; If stripping would erase the whole name (e.g. folder is literally "Scripts"), keep the original.
    if (!projectName)
        return ""
    orig := Trim(projectName)
    cleaned := projectName
    ; Match whole-word "Script" or "Scripts" (case-insensitive).
    cleaned := RegExReplace(cleaned, "(?i)\bscript(s)?\b", "")
    cleaned := RegExReplace(cleaned, "\s{2,}", " ")
    cleaned := Trim(cleaned)
    if (StrLen(cleaned) < 2)
        return orig
    return cleaned
}

; Collapse "file.ext (Label) (file.ext)" in Cursor titles to a single filename + label.
CursorTransfer_StripDuplicateFilenameInParens(title) {
    if (!title)
        return ""
    return RegExReplace(title, "(\S+\.\w+)\s+(\([^)]+\))\s+\(\1\)", "$1 $2")
}

; If the window title starts with the project name (already shown in brackets), drop that prefix.
CursorTransfer_StripLeadingProjectFromTitle(title, cleanProjName, projName) {
    if (!title)
        return ""
    ; Longer label first so a shorter prefix cannot steal a match from a longer project name.
    names := []
    if (StrLen(cleanProjName) > StrLen(projName)) {
        if (cleanProjName != "")
            names.Push(cleanProjName)
        if (projName != "")
            names.Push(projName)
    } else {
        if (projName != "")
            names.Push(projName)
        if (cleanProjName != "" && cleanProjName != projName)
            names.Push(cleanProjName)
    }
    for name in names {
        if (StrLen(name) < 2)
            continue
        tl := StrLower(title)
        nl := StrLower(name)
        if (SubStr(tl, 1, StrLen(nl)) != nl)
            continue
        rest := SubStr(title, StrLen(name) + 1)
        rest := Trim(rest)
        if (rest = "")
            return ""
        if (SubStr(rest, 1, 1) = "-" || SubStr(rest, 1, 1) = "|" || SubStr(rest, 1, 1) = ":" || SubStr(rest, 1, 1) =
            "-")
            rest := Trim(SubStr(rest, 2))
        rest := Trim(LTrim(rest, "- "))
        return (rest != "") ? rest : title
    }
    return title
}

; Remove repetitive app suffix (e.g., " - Cursor", " - Visual Studio Code") from list labels; all windows are already from target app.
CursorTransfer_StripTrailingCursorAppSuffix(s) {
    if (!s)
        return ""
    t := Trim(s)
    targetApp := CursorTransfer_GetTargetAppName()
    if (StrLower(t) = StrLower(targetApp))
        return ""
    ; Strip " - Cursor", " - VS Code", " - Visual Studio Code", and similar variants
    result := Trim(RegExReplace(t, "i)\s*[---]\s*(?:Cursor|VS Code|Visual Studio Code)\s*$", ""))
    return result
}

Clipboard_GetSequenceNumber() {
    ; WinAPI: https://learn.microsoft.com/windows/win32/api/winuser/nf-winuser-getclipboardsequencenumber
    try {
        return DllCall("GetClipboardSequenceNumber", "uint")
    } catch {
        return 0
    }
}

Clipboard_WaitForSequenceChange(seqBefore, totalTimeoutMs := 2000, fastPhaseMs := 850) {
    ; Tight, bounded wait: returns true as soon as the clipboard sequence changes.
    ; Uses a fast polling phase first, then a slower phase for the remainder.
    start := A_TickCount
    deadline := start + totalTimeoutMs
    fastDeadline := start + fastPhaseMs
    while (A_TickCount < deadline) {
        seqNow := Clipboard_GetSequenceNumber()
        if (seqNow && seqNow != seqBefore)
            return true
        Sleep((A_TickCount < fastDeadline) ? 20 : 50)
    }
    return false
}

GetGeminiScriptMsgTargetHwnd() {
    ; Cache-first resolver for Gemini.ahk AutoHotkey script window.
    static cached := 0
    if (cached && WinExist("ahk_id " cached)) {
        try {
            if (InStr(WinGetTitle("ahk_id " cached), "Gemini.ahk"))
                return cached
        } catch {
        }
    }

    prevMatch := A_TitleMatchMode
    DetectHiddenWindows(true)
    SetTitleMatchMode(2)
    found := 0
    try {
        for hwnd in WinGetList("ahk_exe AutoHotkey64.exe") {
            try {
                if (InStr(WinGetTitle("ahk_id " hwnd), "Gemini.ahk")) {
                    found := hwnd
                    break
                }
            } catch {
                continue
            }
        }
        if (!found) {
            for hwnd in WinGetList("ahk_exe AutoHotkey32.exe") {
                try {
                    if (InStr(WinGetTitle("ahk_id " hwnd), "Gemini.ahk")) {
                        found := hwnd
                        break
                    }
                } catch {
                    continue
                }
            }
        }
    } finally {
        SetTitleMatchMode(prevMatch)
        DetectHiddenWindows(false)
    }

    if (found)
        cached := found
    return found
}

; Returns selected Cursor window HWND or 0 on cancel/timeout/no windows. Blocking with timeout.
; centerOnHwnd: optional window to center the modal on (uses that window's monitor); 0 = foreground monitor.
CursorTransfer_ShowWindowSelector(centerOnHwnd := 0) {
    global g_CursorTransferSelectorGui, g_CursorTransferSelectorActive, g_CursorTransferSelectorResult
    global g_CursorTransferWindowList, g_CursorTransferHotkeyHandlers
    global g_Projects, g_ProjectCharSequence
    CursorTransfer_SelectorClose()
    list := []

    ; Resolve target app executable based on environment
    targetAppExe := CursorTransfer_GetTargetAppExecutable()
    appDisplayName := CursorTransfer_GetTargetAppName()

    ; Only enumerate visible windows (DetectHiddenWindows off = default).
    ; Hidden Cursor/VS Code renderer/worker processes can have non-empty titles, pass the
    ; title filter, and end up at position 1 in the sorted list. Their hwnd cannot be found
    ; by WinExist in the hotkey handler thread (which also runs with DetectHiddenWindows off),
    ; causing the first-item selection to silently fail while all other items work fine.
    try {
        for hwnd in WinGetList("ahk_exe " targetAppExe) {
            try {
                title := WinGetTitle("ahk_id " hwnd)
                if (title = "" || InStr(StrLower(title), "preview"))
                    continue
                if (title = "MSCTFIME UI" || title = "Default IME")
                    continue
                list.Push({ hwnd: hwnd, title: title })
                if (list.Length >= 9)
                    break
            } catch {
                continue
            }
        }
    } catch {
    }
    if (list.Length = 0) {
        msg := "❌ No " appDisplayName " windows found"
        ShowCenteredOverlay_Utils(msg, 2000, BANNER_ACCENT_ERROR)
        return 0
    }
    ; Enrich list using path-first project identification, then sort by canonical project order.
    enriched := []
    for w in list {
        winTitle := w.title ? w.title : ""
        projectIndex := CursorTransfer_GetMatchingProjectIndex(w.hwnd, winTitle)
        projectOrder := Integer(projectIndex > 0 ? projectIndex : 10000 + enriched.Length)
        projName := ""
        if (projectIndex > 0) {
            try {
                project := g_Projects[projectIndex]
                projName := project.name
            }
        }
        displayName := ""
        shortAfterParens := ""
        shortAfterLeadStrip := ""
        cleanProjName := ""
        if (projName != "") {
            shortTitle := winTitle ? winTitle : ""
            cleanProjName := CursorTransfer_StripStaticScriptTokenForDisplay(projName)
            if (shortTitle != "") {
                shortTitle := CursorTransfer_StripDuplicateFilenameInParens(shortTitle)
                shortAfterParens := shortTitle
                shortTitle := CursorTransfer_StripLeadingProjectFromTitle(shortTitle, cleanProjName, projName)
                shortAfterLeadStrip := shortTitle
            }
            ; Label = window title only (no g_Projects name prefix). If stripping the duplicate workspace
            ; segment would leave only "Cursor", keep the fuller line (e.g. "scripts - Cursor").
            if (shortAfterParens != "") {
                cand := shortAfterLeadStrip
                if (cand = "" || CursorTransfer_IsUninformativeCursorTitle(cand))
                    displayName := shortAfterParens
                else
                    displayName := cand
                if (CursorTransfer_IsUninformativeCursorTitle(displayName))
                    displayName := displayName . " · #" . w.hwnd
            } else {
                displayName := winTitle ? winTitle : (appDisplayName . " Window " . w.hwnd)
            }
        } else {
            displayName := winTitle ? winTitle : (appDisplayName . " Window " . w.hwnd)
        }
        displayName := CursorTransfer_StripTrailingCursorAppSuffix(displayName)
        if (displayName = "")
            displayName := "#" . w.hwnd
        enriched.Push({
            hwnd: w.hwnd,
            title: winTitle,
            displayName: displayName,
            projectOrder: projectOrder,
            projectIndex: projectIndex,
            hotkeyChar: ""
        })
    }
    if (enriched.Length = 0) {
        ShowCenteredOverlay_Utils("❌ No mapped Cursor projects found", 2000, BANNER_ACCENT_ERROR)
        return 0
    }
    try {
        CursorTransfer_SortWindowsByProjectOrder(&enriched)
        ; Most important first: put active Cursor window at position 1 if it's in the list.
        try {
            activeHwnd := WinGetID("A")
            if (activeHwnd) {
                loop enriched.Length {
                    if (enriched[A_Index].hwnd = activeHwnd) {
                        if (A_Index > 1) {
                            swap := enriched[1]
                            enriched[1] := enriched[A_Index]
                            enriched[A_Index] := swap
                        }
                        break
                    }
                }
            }
        } catch {
        }
        if (enriched.Length > 9) {
            trimmed := []
            loop 9
                trimmed.Push(enriched[A_Index])
            list := trimmed
        } else {
            list := enriched
        }
        ; Number keys 1-9 (like #!+C / SelectAiModelInHandy modal).
        loop list.Length {
            list[A_Index].hotkeyChar := String(A_Index)
        }
        g_CursorTransferWindowList := list
        g_CursorTransferSelectorResult := ""
        g_CursorTransferSelectorActive := true
        g_CursorTransferSelectorGui := Gui("+AlwaysOnTop -Caption +ToolWindow")
        g_CursorTransferSelectorGui.BackColor := "1E1E2E"
        g_CursorTransferSelectorGui.MarginX := 20
        g_CursorTransferSelectorGui.MarginY := 15
        g_CursorTransferSelectorGui.OnEvent("Escape", CursorTransfer_SelectorEscape)
        g_CursorTransferSelectorGui.SetFont("s14 cCDD6F4 Bold", "Segoe UI")
        transferSelGuiW := 720
        g_CursorTransferSelectorGui.Add("Text", "w" . transferSelGuiW . " Center", "📋 Transfer to " . appDisplayName)
        g_CursorTransferSelectorGui.Add("Text", "w" . transferSelGuiW . " h1 Background45475A")
        g_CursorTransferSelectorGui.SetFont("s12 cCDD6F4", "Segoe UI")
        for w in list {
            g_CursorTransferSelectorGui.Add("Text", "w" . transferSelGuiW, "[" . w.hotkeyChar . "] " . w.displayName)
        }
        g_CursorTransferSelectorGui.Add("Text", "w" . transferSelGuiW . " h1 Background45475A y+10")
        g_CursorTransferSelectorGui.SetFont("s9 c6C7086", "Segoe UI")
        g_CursorTransferSelectorGui.Add("Text", "w" . transferSelGuiW . " Center", "Press 1-9 | N or Esc to cancel")
        g_CursorTransferSelectorGui.Show("AutoSize Hide")
        g_CursorTransferSelectorGui.GetPos(&gx, &gy, &gw, &gh)
        ; Center on centerOnHwnd's monitor when provided; otherwise foreground window's monitor (dictation / transfer flow).
        monitorIndex := 1
        if (centerOnHwnd && WinExist("ahk_id " centerOnHwnd)) {
            rect := Buffer(16, 0)
            if (DllCall("GetWindowRect", "ptr", centerOnHwnd, "ptr", rect)) {
                cx := NumGet(rect, 0, "int") + (NumGet(rect, 8, "int") - NumGet(rect, 0, "int")) // 2
                cy := NumGet(rect, 4, "int") + (NumGet(rect, 12, "int") - NumGet(rect, 4, "int")) // 2
                monitorCount := MonitorGetCount()
                loop monitorCount {
                    idx := A_Index
                    MonitorGet(idx, &ml, &mt, &mr, &mb)
                    if (cx >= ml && cx <= mr && cy >= mt && cy <= mb) {
                        monitorIndex := idx
                        break
                    }
                }
            }
        } else {
            monitorIndex := GetMonitorIndexForForeground_StandardBar()
        }
        MonitorGetWorkArea(monitorIndex, &ml, &mt, &mr, &mb)
        mw := mr - ml
        mh := mb - mt
        cx := ml + (mw - gw) // 2
        cy := mt + (mh - gh) // 2
        g_CursorTransferSelectorGui.Show("x" . cx . " y" . cy . " NA")
    } catch as err {
        ShowCenteredOverlay_Utils("❌ Selector error", 2000, BANNER_ACCENT_ERROR)
        return 0
    }
    ; Wildcard prefix so hotkeys fire even when modifier held (e.g. C still down when modal opens).
    g_CursorTransferHotkeyHandlers := []
    loop list.Length {
        i := A_Index
        keyChar := list[i].hotkeyChar
        if (keyChar = "")
            continue
        try {
            fn := CursorTransfer_SelectorHandleKey.Bind(i)
            hotkeyKey := "*" . keyChar
            Hotkey(hotkeyKey, fn, "On")
            g_CursorTransferHotkeyHandlers.Push({ key: hotkeyKey, callback: fn })
        } catch {
        }
    }
    try {
        Hotkey("*Escape", CursorTransfer_SelectorEscape, "On")
        g_CursorTransferHotkeyHandlers.Push({ key: "*Escape", callback: CursorTransfer_SelectorEscape })
        Hotkey("*N", CursorTransfer_SelectorEscape, "On")
        g_CursorTransferHotkeyHandlers.Push({ key: "*N", callback: CursorTransfer_SelectorEscape })
        Hotkey("*n", CursorTransfer_SelectorEscape, "On")
        g_CursorTransferHotkeyHandlers.Push({ key: "*n", callback: CursorTransfer_SelectorEscape })
    } catch {
    }
    start := A_TickCount
    timeoutMs := 30000
    lastCursorTransferMonitorIdx := monitorIndex
    keyWasDownByIndex := []
    loop list.Length
        keyWasDownByIndex.Push(false)
    cancelWasDown := false
    try {
        while (g_CursorTransferSelectorResult = "") {
            if ((A_TickCount - start) >= timeoutMs)
                break
            Sleep 50

            ; Fallback key polling (edge-triggered): survives global Hotkey("1") conflicts from other scripts.
            loop list.Length {
                if (g_CursorTransferSelectorResult != "")
                    break
                keyChar := list[A_Index].hotkeyChar
                if (keyChar = "")
                    continue
                isDown := false
                try isDown := GetKeyState(keyChar, "P") || GetKeyState("Numpad" . keyChar, "P")
                if (isDown && !keyWasDownByIndex[A_Index])
                    CursorTransfer_SelectorHandleKey(A_Index)
                keyWasDownByIndex[A_Index] := isDown
            }

            isCancelDown := false
            try isCancelDown := GetKeyState("Escape", "P") || GetKeyState("n", "P") || GetKeyState("N", "P")
            if (isCancelDown && !cancelWasDown && g_CursorTransferSelectorResult = "")
                CursorTransfer_SelectorEscape()
            cancelWasDown := isCancelDown

            if (IsObject(g_CursorTransferSelectorGui)) {
                curIdx := GetMonitorIndexForForeground_StandardBar()
                if (curIdx != lastCursorTransferMonitorIdx) {
                    lastCursorTransferMonitorIdx := curIdx
                    MonitorGetWorkArea(curIdx, &ml, &mt, &mr, &mb)
                    mw := mr - ml
                    mh := mb - mt
                    try {
                        g_CursorTransferSelectorGui.GetPos(&gxOld, &gyOld, &gw, &gh)
                        cx := ml + (mw - gw) // 2
                        cy := mt + (mh - gh) // 2
                        g_CursorTransferSelectorGui.Show("x" . cx . " y" . cy . " NA")
                    } catch {
                    }
                }
            }
        }
    } catch as loopErr {
        ; ignore loop exceptions
    }
    durationMs := A_TickCount - start
    result := (g_CursorTransferSelectorResult = "") ? 0 : Integer(g_CursorTransferSelectorResult)
    CursorTransfer_SelectorClose()
    return result
}

; =============================================================================
; Cursor AI text field focus (shared logic for Gemini transfer and WindowManagement)
; =============================================================================

; VS Code exposes both the main editor and chat composer as native-edit-context Edit controls.
; To avoid false positives, pick the lowest compact edit near the Send button region.
VSCode_FindChatInputField(root) {
    bestEdit := ""
    bestTop := -2147483647
    sendBr := ""
    if (!root)
        return ""
    try {
        sendBtn := VSCode_FindChatSendButton(root)
        if (sendBtn)
            sendBr := sendBtn.BoundingRectangle
    } catch {
        sendBr := ""
    }
    try {
        allEdits := root.FindAll({ Type: UIA.Type.Edit })
        for editEl in allEdits {
            try {
                className := editEl.ClassName
                if (!InStr(className, "native-edit-context"))
                    continue
                if (editEl.GetPropertyValue(UIA.Property.IsOffscreen))
                    continue
                br := editEl.BoundingRectangle
                h := br.b - br.t
                w := br.r - br.l
                if (h <= 0 || w <= 0)
                    continue

                ; Skip large editor surfaces; chat composer is typically compact.
                if (h > 260)
                    continue

                ; When send button exists, require geometric proximity to chat footer area.
                if (sendBr != "") {
                    cx := (br.l + br.r) / 2
                    cy := (br.t + br.b) / 2
                    sendCx := (sendBr.l + sendBr.r) / 2
                    sendCy := (sendBr.t + sendBr.b) / 2
                    if (Abs(cx - sendCx) > 900)
                        continue
                    if (Abs(cy - sendCy) > 420)
                        continue
                }

                ; Prefer the visually lowest eligible edit field.
                if (br.t >= bestTop) {
                    bestTop := br.t
                    bestEdit := editEl
                }
            } catch {
                continue
            }
        }
    } catch {
    }
    return bestEdit
}

VSCode_EnsureChatInputHasFocus(editEl) {
    if (!editEl)
        return false
    try {
        editEl.SetFocus()
    } catch {
    }
    loop 3 {
        try {
            if (editEl.HasKeyboardFocus)
                return true
        } catch {
        }
        Sleep 40
    }
    try {
        editEl.ScrollIntoView()
    } catch {
    }
    try {
        editEl.Click()
    } catch {
    }
    Sleep 80
    try {
        return editEl.HasKeyboardFocus
    } catch {
        return false
    }
}

VSCode_FindChatSendButton(root) {
    if (!root)
        return ""
    lastMatchingButton := ""
    try {
        buttons := root.FindAll({ Type: UIA.Type.Button })
        for buttonEl in buttons {
            try {
                buttonName := buttonEl.Name
                buttonClass := buttonEl.ClassName
                if (!(InStr(buttonName, "Send ") || InStr(buttonName, "Send to New Chat")))
                    continue
                if (InStr(buttonClass, "arrow-up"))
                    return buttonEl
                lastMatchingButton := buttonEl
            } catch {
                continue
            }
        }
    } catch {
    }
    return lastMatchingButton
}

VSCode_IsChatSendReady(targetHwnd) {
    if (!IsSet(UIA))
        return true
    try {
        root := UIA.ElementFromHandle(targetHwnd)
        sendButton := VSCode_FindChatSendButton(root)
        if (!sendButton)
            return false
        try {
            if (sendButton.GetPropertyValue(UIA.Property.IsEnabled))
                return true
        } catch {
        }
        try {
            return !InStr(sendButton.ClassName, "disabled")
        } catch {
        }
    } catch {
    }
    return false
}

VSCode_IsChatInputFocused(targetHwnd) {
    if (!IsSet(UIA))
        return false
    try {
        root := UIA.ElementFromHandle(targetHwnd)
        ; Fast path: check which element currently has focus.
        ; The main code editor uses native-edit-context but is far from the Send button.
        try {
            focusedEl := UIA.GetFocusedElement()
            if (focusedEl) {
                focusedClass := ""
                try focusedClass := focusedEl.ClassName
                if (InStr(focusedClass, "native-edit-context")) {
                    sendBtn := VSCode_FindChatSendButton(root)
                    if (!sendBtn)
                        return false
                    try {
                        focusedBr := focusedEl.BoundingRectangle
                        sendBr := sendBtn.BoundingRectangle
                        if (Abs(((focusedBr.l + focusedBr.r) / 2) - ((sendBr.l + sendBr.r) / 2)) > 900)
                            return false
                        if (Abs(((focusedBr.t + focusedBr.b) / 2) - ((sendBr.t + sendBr.b) / 2)) > 420)
                            return false
                        return true
                    } catch {
                        return false
                    }
                }
            }
        } catch {
        }
        ; Fallback: scan for chat input and verify keyboard focus.
        editEl := VSCode_FindChatInputField(root)
        if (!editEl)
            return false
        try {
            return editEl.HasKeyboardFocus
        } catch {
        }
    } catch {
    }
    return false
}

VSCode_IsSafeChatPasteTarget(targetHwnd) {
    ; Conservative gate: if uncertain, do not paste.
    if (!IsSet(UIA))
        return false
    if (!targetHwnd || !WinExist("ahk_id " targetHwnd))
        return false
    try {
        root := UIA.ElementFromHandle(targetHwnd)
        if (!root)
            return false
        editEl := VSCode_FindChatInputField(root)
        sendBtn := VSCode_FindChatSendButton(root)
        if (!editEl || !sendBtn)
            return false

        ; Must be keyboard-focused to avoid pasting into editor/other controls.
        try {
            if (!editEl.HasKeyboardFocus)
                return false
        } catch {
            return false
        }

        editBr := editEl.BoundingRectangle
        sendBr := sendBtn.BoundingRectangle
        ew := editBr.r - editBr.l
        eh := editBr.b - editBr.t
        if (ew < 260 || eh <= 0 || eh > 260)
            return false

        ; Composer and send button must be in the same footer region.
        editCx := (editBr.l + editBr.r) / 2
        editCy := (editBr.t + editBr.b) / 2
        sendCx := (sendBr.l + sendBr.r) / 2
        sendCy := (sendBr.t + sendBr.b) / 2
        if (Abs(editCx - sendCx) > 900)
            return false
        if (Abs(editCy - sendCy) > 260)
            return false

        ; Require right-side chat pane placement (avoids inline "Get comment" editors).
        rect := Buffer(16, 0)
        if (!DllCall("GetWindowRect", "ptr", targetHwnd, "ptr", rect))
            return false
        winLeft := NumGet(rect, 0, "int")
        winRight := NumGet(rect, 8, "int")
        winWidth := winRight - winLeft
        if (winWidth <= 0)
            return false
        thresholdX := winLeft + (winWidth * 0.52)
        if (editCx < thresholdX || sendCx < thresholdX)
            return false

        return true
    } catch {
    }
    return false
}

VSCode_SubmitChat(targetHwnd) {
    if (IsSet(UIA)) {
        try {
            root := UIA.ElementFromHandle(targetHwnd)
            sendButton := VSCode_FindChatSendButton(root)
            if (sendButton && VSCode_IsChatSendReady(targetHwnd)) {
                try {
                    sendButton.Click()
                    return true
                } catch {
                }
                try {
                    sendButton.SetFocus()
                    Sleep 40
                    SendInput "{Enter}"
                    return true
                } catch {
                }
            }
        } catch {
        }
    }
    SendInput "{Enter}"
    return true
}

; Activate VS Code window and focus chat input. Returns true on success.
VSCode_FocusChatInput(targetHwnd := 0) {
    try {
        if (targetHwnd) {
            WinActivate("ahk_id " targetHwnd)
            if (!WinWaitActive("ahk_id " targetHwnd, , 2))
                return false
        } else {
            targetHwnd := WinExist("ahk_exe Code.exe")
            if (!targetHwnd)
                return false
            WinActivate("ahk_id " targetHwnd)
            if (!WinWaitActive("ahk_id " targetHwnd, , 2))
                return false
        }
        Sleep 180

        ; Ensure the chat surface exists first.
        SendInput "^!i"
        Sleep 350

        if (!IsSet(UIA)) {
            SendInput "{Tab}"
            Sleep 120
            return true
        }

        loop 2 {
            try {
                root := UIA.ElementFromHandle(targetHwnd)
                editEl := VSCode_FindChatInputField(root)
                if (editEl && VSCode_EnsureChatInputHasFocus(editEl)) {
                    return true
                }
            } catch {
            }

            ; Fallback nudge inside the chat view without clicking toolbar buttons.
            SendInput "{Tab}"
            Sleep 140
        }

        return false
    } catch {
        return false
    }
}

Cursor_EnsureComposerHasFocus(editEl) {
    if (!editEl)
        return false
    try {
        editEl.SetFocus()
    } catch {
    }
    loop 3 {
        try {
            if (editEl.HasKeyboardFocus)
                return true
        } catch {
        }
        Sleep 40
    }
    try {
        editEl.ScrollIntoView()
    } catch {
    }
    try {
        editEl.Click()
    } catch {
    }
    Sleep 60
    try {
        return editEl.HasKeyboardFocus
    } catch {
        return false
    }
}

Cursor_FindComposerInput(root) {
    try {
        allEdits := root.FindAll({ Type: UIA.Type.Edit })
        for editEl in allEdits {
            cn := editEl.ClassName
            if (InStr(cn, "aislash-editor-input") && !InStr(cn, "readonly"))
                return editEl
        }
    } catch {
    }
    return ""
}

; Activate Cursor window and focus AI composer input. Returns true on success.
Cursor_FocusAITextField(targetHwnd := 0) {
    try {
        if (targetHwnd) {
            WinActivate("ahk_id " targetHwnd)
            WinWaitActive("ahk_id " targetHwnd, , 2)
        } else {
            targetHwnd := WinExist("ahk_exe Cursor.exe")
            if (!targetHwnd)
                return false
            WinWaitActive("ahk_id " targetHwnd, , 2)
        }
        Sleep 200
        paneWasOpen := false
        focusDone := false
        if (IsSet(UIA)) {
            try {
                root := UIA.ElementFromHandle(targetHwnd)
                if (root) {
                    toggleEl := root.FindFirst({ Type: UIA.Type.CheckBox, Name: "Toggle AI Pane", matchmode: 2 })
                    paneOpen := toggleEl && InStr(toggleEl.ClassName, "checked")
                    paneWasOpen := paneOpen
                    if (paneOpen) {
                        editEl := Cursor_FindComposerInput(root)
                        if (editEl) {
                            if (Cursor_EnsureComposerHasFocus(editEl))
                                focusDone := true
                        }
                    } else {
                        editEl := Cursor_FindComposerInput(root)
                        if (editEl) {
                            if (Cursor_EnsureComposerHasFocus(editEl))
                                focusDone := true
                        } else {
                            Send "^i"
                            loop 15 {
                                Sleep 200
                                root := UIA.ElementFromHandle(targetHwnd)
                                editEl := Cursor_FindComposerInput(root)
                                if (editEl) {
                                    if (Cursor_EnsureComposerHasFocus(editEl))
                                        focusDone := true
                                    break
                                }
                            }
                        }
                    }
                }
            } catch {
            }
        }
        if (!focusDone) {
            if (IsSet(UIA)) {
                try {
                    root := UIA.ElementFromHandle(targetHwnd)
                    if (root) {
                        editEl := Cursor_FindComposerInput(root)
                        if (editEl) {
                            if (Cursor_EnsureComposerHasFocus(editEl))
                                focusDone := true
                        }
                    }
                } catch {
                }
            }
            if (!focusDone) {
                if (!paneWasOpen) {
                    Send "^i"
                    Sleep 1200
                }
                return false
            }
        }
        return true
    } catch {
        return false
    }
}

; Minimum clipboard length for transfer (match Gemini/bridge validation)
CURSOR_TRANSFER_MIN_CLIPBOARD_LENGTH := 10
; Electron/Cursor needs time after paste before Enter; after Enter before foreground changes or submit can drop.
CURSOR_TRANSFER_POST_PASTE_BEFORE_ENTER_MS := 80
CURSOR_TRANSFER_POST_ENTER_BEFORE_RESTORE_MS := 400
VSCODE_TRANSFER_POST_PASTE_SETTLE_MS := 180
VSCODE_TRANSFER_PASTE_RETRY_COUNT := 2

; Activate Cursor/VS Code window, focus AI field, paste clipboard, send Enter. Non-blocking feedback on failure.
; restoreFocusHwnd: if set, WinActivate this window after Enter is processed (before success overlay) so focus does not stay on the target window.
CursorTransfer_ActivateFocusPaste(targetHwnd, restoreFocusHwnd := 0) {
    appDisplayName := CursorTransfer_GetTargetAppName()
    targetIsVSCode := (CursorTransfer_GetTargetAppExecutable() = "Code.exe")
    if (!targetHwnd || !WinExist("ahk_id " targetHwnd)) {
        ShowCenteredOverlay_Utils("❌ " appDisplayName " window not found", 2000, BANNER_ACCENT_ERROR)
        return
    }
    clip := Trim(A_Clipboard)
    if (clip = "" || StrLen(clip) < CURSOR_TRANSFER_MIN_CLIPBOARD_LENGTH) {
        ShowCenteredOverlay_Utils("❌ Clipboard empty or too short", 2000, BANNER_ACCENT_ERROR)
        return
    }
    try {
        WinActivate("ahk_id " targetHwnd)
        if (!WinWaitActive("ahk_id " targetHwnd, , 2)) {
            ShowCenteredOverlay_Utils("❌ Could not activate " appDisplayName, 2000, BANNER_ACCENT_ERROR)
            return
        }
        Sleep 100
        focusSucceeded := false
        if (targetIsVSCode) {
            focusSucceeded := VSCode_FocusChatInput(targetHwnd)
        } else {
            focusSucceeded := Cursor_FocusAITextField(targetHwnd)
        }
        if (!focusSucceeded) {
            ShowCenteredOverlay_Utils("❌ Could not focus AI field", 2000, BANNER_ACCENT_ERROR)
            return
        }
        Sleep(targetIsVSCode ? 220 : 100)
        try {
            if (WinGetID("A") != targetHwnd) {
                WinActivate("ahk_id " targetHwnd)
                WinWaitActive("ahk_id " targetHwnd, , 2)
            }
        } catch {
        }
        if (Trim(A_Clipboard) = "" || StrLen(Trim(A_Clipboard)) < CURSOR_TRANSFER_MIN_CLIPBOARD_LENGTH) {
            ShowCenteredOverlay_Utils("❌ Clipboard lost before paste", 2000, BANNER_ACCENT_ERROR)
            return
        }
        pasteDetected := true
        pasteAttempts := targetIsVSCode ? VSCODE_TRANSFER_PASTE_RETRY_COUNT : 1
        loop pasteAttempts {
            if (targetIsVSCode) {
                ; Hard gate: never paste until strict chat-pane targeting is verified.
                if (!VSCode_IsSafeChatPasteTarget(targetHwnd)) {
                    if (!VSCode_FocusChatInput(targetHwnd))
                        break
                    Sleep 160
                    ; Require stability across two checks to avoid transient focus races.
                    if (!VSCode_IsSafeChatPasteTarget(targetHwnd)) {
                        Sleep 90
                    }
                    if (!VSCode_IsSafeChatPasteTarget(targetHwnd)) {
                        if (A_Index < pasteAttempts)
                            continue
                        break
                    }
                }
            }
            SendInput "^v"
            Sleep CURSOR_TRANSFER_POST_PASTE_BEFORE_ENTER_MS + (targetIsVSCode ? VSCODE_TRANSFER_POST_PASTE_SETTLE_MS :
                0)
            if (!targetIsVSCode) {
                pasteDetected := true
                break
            }
            pasteDetected := VSCode_IsChatSendReady(targetHwnd)
            if (pasteDetected)
                break
            if (A_Index < pasteAttempts) {
                if (!VSCode_FocusChatInput(targetHwnd))
                    break
                Sleep 150
            }
        }
        if (!pasteDetected) {
            ShowCenteredOverlay_Utils("❌ Paste blocked: AI text field not confidently focused", 2600,
                BANNER_ACCENT_ERROR)
            return
        }
        if (targetIsVSCode) {
            if (!VSCode_SubmitChat(targetHwnd)) {
                ShowCenteredOverlay_Utils("❌ Could not submit VS Code chat", 2200, BANNER_ACCENT_ERROR)
                return
            }
        } else {
            SendInput "{Enter}"
        }
        if (restoreFocusHwnd && WinExist("ahk_id " restoreFocusHwnd)) {
            ; Wait so target editor keeps foreground until paste + Enter are processed; restoring sooner drops Enter.
            Sleep CURSOR_TRANSFER_POST_ENTER_BEFORE_RESTORE_MS
            try {
                WinActivate("ahk_id " restoreFocusHwnd)
                if (!WinActive("ahk_id " restoreFocusHwnd))
                    WinWaitActive("ahk_id " restoreFocusHwnd, , 0.5)
            } catch {
            }
        }
        appDisplayName := CursorTransfer_GetTargetAppName()
        ShowCenteredOverlay_Utils("✅ Sent to " . appDisplayName, 1500, BANNER_ACCENT_SUCCESS)
    } catch as err {
        ShowCenteredOverlay_Utils("❌ Transfer failed", 2000, BANNER_ACCENT_ERROR)
    }
}

; =============================================================================
; Status Banner Functions (non-blocking; use standard loading indicator)
; =============================================================================
AiModelBanner_Show(text, bgColor := BANNER_ACCENT_INTERMEDIATE) {
    StandardLoadingBar_Show(text, bgColor, { passive: true, centerOnHwnd: 0, textWidth: 450, fontSize: 17,
        passiveBgColor: bgColor, alpha: 200 })
}

AiModelBanner_Hide() {
    StandardLoadingBar_Hide(0)
}

; =============================================================================
; Persistent Language Flag Indicator (slot 3 = UK, slot 4 = Brazil)
; =============================================================================
; Opaque, always-on-top flag chips pinned to the bottom-right of every monitor.
; Slot 3 shows English (UK flag), slot 4 shows Portuguese (Brazil flag).
; Use the visible flag as the spoken-language indicator across all screens.
; =============================================================================

LanguageFlag_GetImagePath(slot) {
    rel := (slot = 3) ? "\images\flags\united-kingdom.png" : (slot = 4) ? "\images\flags\brazil.png" : ""
    if (rel = "")
        return ""
    ; Prefer the running script's folder, then the folder that contains Utils.ahk (covers odd layouts).
    candidates := [A_ScriptDir . rel]
    SplitPath(A_LineFile, , &utilsDir)
    if (utilsDir != "" && utilsDir != A_ScriptDir)
        candidates.Push(utilsDir . rel)
    for p in candidates {
        if FileExist(p)
            return p
    }
    return ""
}

LanguageFlag_CreateGui(slot, imagePath) {
    global LANGUAGE_FLAG_WIDTH

    ; Borderless, zero-margin window so the GUI sizes exactly to the bitmap.
    ; Do NOT use WS_EX_TRANSPARENT (part of +E0x80020): that style suppresses
    ; window painting and the flag can disappear entirely.
    flagGui := Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale")
    flagGui.BackColor := "313244"
    flagGui.MarginX := 0
    flagGui.MarginY := 0

    usedPicture := false
    if (imagePath != "") {
        try {
            flagGui.Add("Picture", "w" . LANGUAGE_FLAG_WIDTH . " h-1", imagePath)
            usedPicture := true
        } catch {
            usedPicture := false
        }
    }
    if !usedPicture {
        flagGui.SetFont("s13 cFFFFFF Bold", "Segoe UI")
        label := (slot = 3) ? "EN" : "PT"
        flagGui.Add("Text", "Center w" . LANGUAGE_FLAG_WIDTH . " h31 Background45475A", label)
    }
    flagGui.Show("AutoSize Hide")
    return flagGui
}

LanguageFlag_Show(slot) {
    global g_LanguageFlagGuis, g_LanguageFlagSlot

    if (slot != 3 && slot != 4) {
        LanguageFlag_Hide()
        return
    }

    LanguageFlag_Hide()
    g_LanguageFlagSlot := slot

    imagePath := LanguageFlag_GetImagePath(slot)
    monitorCount := MonitorGetCount()
    if (monitorCount < 1)
        return

    loop monitorCount {
        idx := A_Index
        flagGui := LanguageFlag_CreateGui(slot, imagePath)
        g_LanguageFlagGuis.Push({ monitor: idx, gui: flagGui })
    }

    LanguageFlag_RepositionAllMonitors()
}

LanguageFlag_Hide() {
    global g_LanguageFlagGuis, g_LanguageFlagSlot
    for item in g_LanguageFlagGuis {
        try {
            if IsObject(item.gui)
                item.gui.Destroy()
        } catch {
        }
    }
    g_LanguageFlagGuis := []
    g_LanguageFlagSlot := 0
}

LanguageFlag_RepositionAllMonitors() {
    global g_LanguageFlagGuis, LANGUAGE_FLAG_MARGIN
    if (!g_LanguageFlagGuis.Length)
        return

    for item in g_LanguageFlagGuis {
        monitorIdx := item.monitor
        flagGui := item.gui
        if !IsObject(flagGui)
            continue

        try {
            MonitorGetWorkArea(monitorIdx, &ml, &mt, &mr, &mb)
            flagGui.GetPos(, , &gw, &gh)
        } catch {
            continue
        }

        guiX := mr - gw - LANGUAGE_FLAG_MARGIN
        guiY := mb - gh - LANGUAGE_FLAG_MARGIN
        if (guiX < ml)
            guiX := ml
        if (guiY < mt)
            guiY := mt

        try {
            flagGui.Move(guiX, guiY)
            hwnd := flagGui.Hwnd
            if (hwnd) {
                ; SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE = 0x0001 | 0x0004 | 0x0010 = 0x0015
                DllCall("SetWindowPos", "Ptr", hwnd, "Ptr", 0, "Int", guiX, "Int", guiY, "Int", 0, "Int", 0,
                    "UInt", 0x0015)
            }
            flagGui.Show("NA")
        } catch {
        }
    }
}

LanguageFlag_InitFromPersistedSlot() {
    slot := Handy_GetPersistedAiModelSlot()
    if (slot = 3 || slot = 4)
        LanguageFlag_Show(slot)
}

; Small banner for Clip Angel (uses standard loading indicator).
ClipAngelBanner_Show(text, bgColor := BANNER_ACCENT_INTERMEDIATE) {
    StandardLoadingBar_Show(text, bgColor, { passive: true, centerOnHwnd: 0, textWidth: 200, fontSize: 17,
        passiveBgColor: bgColor, alpha: 220 })
}

ClipAngelBanner_Hide() {
    StandardLoadingBar_Hide(0)
}

; Fast Copy Mode (Shift keys - Clip Angel sequential paste): persistent banner with live copy count.
FastCopyModeBanner_Show() {
    StandardLoadingBar_Show("📋 Fast Copy Mode - copies: 0", BANNER_ACCENT_INFO, { passive: true, centerOnHwnd: 0,
        textWidth: 480, fontSize: 17, passiveBgColor: BANNER_ACCENT_INFO, alpha: 220,
        promptKeys: "[Win+Alt+Shift+J] Finish and paste", trackActiveMonitor: true })
}

FastCopyModeBanner_Update(copyCount) {
    StandardLoadingBar_Update("📋 Fast Copy Mode - copies: " copyCount)
}

FastCopyModeBanner_Hide() {
    StandardLoadingBar_Hide(0)
}

; =============================================================================
; Single-character tab banner (uses standard loading indicator). tabNumber 1 = blue, 2 = yellow. Auto-hides after 700 ms.
; =============================================================================
ShowSingleCharTabBanner_Utils(tabNumber) {
    msg := String(tabNumber)
    bgColor := (tabNumber = 1) ? "0000FF" : "FFFF00"
    StandardLoadingBar_Show(msg, bgColor, { passive: true, centerOnHwnd: 0, textWidth: 120, fontSize: 72,
        passiveBgColor: bgColor, alpha: 178 })
    StandardLoadingBar_Hide(700)
}

; =============================================================================
; ExecuteHandyAiModelSelection() - Main automation logic for Handy
; =============================================================================
ExecuteHandyAiModelSelection(selection) {
    global g_HandyAiModels

    modelInfo := g_HandyAiModels[selection]
    modelDisplayName := modelInfo.name
    modelClickName := modelInfo.HasProp("modelClickName") ? modelInfo.modelClickName : modelInfo.name

    try {
        ; Step 1: Launch/activate Handy
        AiModelBanner_Show("🚀 Launching Handy...")
        handyHwnd := Handy_ActivateOrLaunch()
        if (!handyHwnd) {
            AiModelBanner_Show("❌ Failed to launch Handy", "E74C3C")
            Sleep 2000
            AiModelBanner_Hide()
            return
        }

        ; Step 1b: Optional - set Cohere language on General (explicit English / Portuguese)
        if (modelInfo.HasProp("cohereLanguage") && modelInfo.cohereLanguage != "") {
            AiModelBanner_Show("🌐 Setting Cohere language: " . modelInfo.cohereLanguage . "...")
            if (!Handy_SetCohereLanguage(handyHwnd, modelInfo.cohereLanguage)) {
                AiModelBanner_Show("❌ Could not set Cohere language", "E74C3C")
                Sleep 2000
                AiModelBanner_Hide()
                return
            }
            Sleep 120
        }

        ; Step 2: Open AI model menu
        AiModelBanner_Show("📋 Opening AI model menu...")
        if (!Handy_OpenAiModelMenu(handyHwnd)) {
            AiModelBanner_Show("❌ Could not open AI menu", "E74C3C")
            Sleep 2000
            AiModelBanner_Hide()
            return
        }

        ; Step 3: Select the model
        AiModelBanner_Show("🎯 Selecting " . modelDisplayName . "...")
        if (!Handy_ClickAiModel(handyHwnd, modelClickName)) {
            AiModelBanner_Show("❌ Model not found: " . modelClickName, "E74C3C")
            Sleep 2000
            AiModelBanner_Hide()
            return
        }
        Handy_SetPersistedAiModelSlot(selection)

        ; Update persistent language flag indicator (slot 3 = UK, slot 4 = BR; hidden for slots 1/2).
        if (selection = 3 || selection = 4)
            LanguageFlag_Show(selection)
        else
            LanguageFlag_Hide()

        ; Step 4: Wait for model to finish loading (poll button name until "loading" disappears)
        AiModelBanner_Show("⏳ Waiting for model...", BANNER_ACCENT_INTERMEDIATE)
        Handy_WaitForModelReady(handyHwnd, 20000)

        ; Step 4.5: Play confirmation sound when model is ready
        soundPath := A_ScriptDir . "\sounds\handy-model-chosen.mp3"
        if (FileExist(soundPath))
            ScriptSoundPlay(soundPath)

        ; Step 5: Close Handy window
        AiModelBanner_Show("✅ Done! Closing Handy...", BANNER_ACCENT_SUCCESS)
        try WinClose("ahk_id " . handyHwnd)
        Sleep 150

        AiModelBanner_Hide()

    } catch Error as e {
        AiModelBanner_Show("❌ Error: " . e.Message, "E74C3C")
        Sleep 2000
        AiModelBanner_Hide()
    }
}

; =============================================================================
; Handy UIA Helper Functions
; =============================================================================

; Activate existing Handy window or launch it; returns hwnd or 0
Handy_ActivateOrLaunch() {
    targetPath := GetHandyShortcutPath()
    expectedExePath := GetHandyProcessPath()

    ; Find existing Handy window
    matchingHwnd := 0
    for hwnd in WinGetList("Handy ahk_class Tauri Window") {
        try {
            procPath := WinGetProcessPath(hwnd)
            if (expectedExePath = "" || StrCompare(procPath, expectedExePath, false) = 0) {
                matchingHwnd := hwnd
                break
            }
        } catch {
            if (expectedExePath = "") {
                matchingHwnd := hwnd
                break
            }
        }
    }

    if (matchingHwnd) {
        WinActivate("ahk_id " . matchingHwnd)
        WinWaitActive("ahk_id " . matchingHwnd, , 2)
        Handy_WaitForMainUiReady(matchingHwnd, 2000)
        return matchingHwnd
    }

    ; Launch Handy
    if (targetPath = "" || !FileExist(targetPath))
        return 0

    Run targetPath
    if !WinWait("Handy ahk_class Tauri Window", , 8)
        return 0

    ; Find the window we just launched
    for h in WinGetList("Handy ahk_class Tauri Window") {
        try {
            procPath := WinGetProcessPath(h)
            if (expectedExePath = "" || StrCompare(procPath, expectedExePath, false) = 0) {
                WinActivate("ahk_id " . h)
                WinWaitActive("ahk_id " . h, , 2)
                if (!Handy_WaitForMainUiReady(h, 9000))
                    return 0
                return h
            }
        } catch {
            if (expectedExePath = "") {
                WinActivate("ahk_id " . h)
                WinWaitActive("ahk_id " . h, , 2)
                if (!Handy_WaitForMainUiReady(h, 9000))
                    return 0
                return h
            }
        }
    }
    return 0
}

; Wait for Handy main UI to be interactive (needed most on cold launch).
Handy_WaitForMainUiReady(hwnd, maxWaitMs := 9000) {
    global UIA
    start := A_TickCount
    pollMs := 120
    loop {
        if ((A_TickCount - start) >= maxWaitMs)
            return false
        el := UIA.ElementFromHandle(hwnd)
        if (el) {
            try {
                if (el.FindFirst({ Type: 50000, Name: "Check for updates" }))
                    return true
            }
            try {
                if (el.FindFirst({ Type: 50000, Name: "Verificar atualizações" }))
                    return true
            }
            try {
                if (el.FindFirst({ Type: 50000, Name: "Update available" }))
                    return true
            }
        }
        Sleep pollMs
    }
}

; True when General tab content (COHERE SETTINGS) is visible.
Handy_GeneralTabVisible(el) {
    if !el
        return false
    try {
        return el.FindFirst({ Type: 50020, Name: "COHERE SETTINGS" }) != 0
    } catch {
        return false
    }
}

; Click sidebar "General" so COHERE SETTINGS is shown (needed from Models/About/etc.).
Handy_EnsureGeneralTab(hwnd) {
    el := UIA.ElementFromHandle(hwnd)
    if !el
        return false
    if (Handy_GeneralTabVisible(el))
        return true
    try {
        gen := el.FindFirst({ Type: 50020, Name: "General" })
        if gen {
            try gen.Click()
            catch {
                try gen.Invoke()
            }
            Sleep 220
            el2 := UIA.ElementFromHandle(hwnd)
            return Handy_GeneralTabVisible(el2)
        }
    } catch {
    }
    return false
}

; Language dropdown under COHERE SETTINGS: class uses "rounded min-w-[200px]" (Microphone uses rounded-md).
Handy_FindHandyLanguageButton(el) {
    if !el
        return 0
    try {
        buttons := el.FindAll({ Type: 50000 })
        for btn in buttons {
            cn := ""
            try cn := btn.ClassName
            if (cn != "" && InStr(cn, "rounded min-w-[200px]"))
                return btn
        }
    } catch {
    }
    return 0
}

; Current Cohere language label on General tab ("" if unknown).
Handy_ReadCohereLanguage(hwnd) {
    el := UIA.ElementFromHandle(hwnd)
    if !el
        return ""
    langBtn := Handy_FindHandyLanguageButton(el)
    if !langBtn
        return ""
    try return langBtn.Name
    return ""
}

; Poll until language button shows langName (short window; no-op if already correct).
Handy_WaitCohereLanguage(hwnd, langName, maxWaitMs := 450) {
    if (langName = "")
        return false
    pollMs := 50
    start := A_TickCount
    loop {
        if (Handy_ReadCohereLanguage(hwnd) = langName)
            return true
        if ((A_TickCount - start) >= maxWaitMs)
            break
        Sleep pollMs
    }
    return Handy_ReadCohereLanguage(hwnd) = langName
}

; Open the COHERE language dropdown on General tab.
Handy_OpenCohereLanguageDropdown(hwnd) {
    el := UIA.ElementFromHandle(hwnd)
    if !el
        return false
    langBtn := Handy_FindHandyLanguageButton(el)
    if !langBtn
        return false
    try langBtn.Click()
    catch {
        try langBtn.Invoke()
    }
    Sleep 200
    return true
}

; With language dropdown open: focus search, type langName, choose row or Enter.
Handy_SetCohereLanguage_PickFromOpenDropdown(hwnd, langName) {
    el := UIA.ElementFromHandle(hwnd)
    if !el
        return false
    searchEl := 0
    try {
        for ed in el.FindAll({ Type: UIA.Type.Edit }) {
            searchEl := ed
            break
        }
    } catch {
    }
    if (searchEl) {
        try {
            searchEl.SetFocus()
        } catch {
            try searchEl.Click()
        }
        Sleep 50
    }
    Send "^a"
    SendText langName
    Sleep 120
    picked := false
    try {
        for btn in el.FindAll({ Type: 50000 }) {
            n := ""
            try n := btn.Name
            if (n != langName)
                continue
            cn := ""
            try cn := btn.ClassName
            if (InStr(cn, "w-full px-3 py-2 text-left") || InStr(cn, "w-full px-3 py-2 text-start")) {
                try btn.Click()
                picked := true
                break
            }
        }
    } catch {
    }
    if !picked
        Send "{Enter}"
    Sleep 80
    return true
}

; Set Cohere transcription language on General tab (explicit list pick, not Auto Detect).
; Retries with verify-after-pick until correct or max attempts (slots 3–4 / English & Portuguese).
Handy_SetCohereLanguage(hwnd, langName) {
    if !hwnd || langName = ""
        return false
    if !Handy_EnsureGeneralTab(hwnd)
        return false
    if (Handy_ReadCohereLanguage(hwnd) = langName)
        return true

    maxAttempts := 3
    loop maxAttempts {
        if (A_Index > 1) {
            Send "{Escape}"
            Sleep 80
            if !Handy_EnsureGeneralTab(hwnd)
                continue
        }
        if !Handy_OpenCohereLanguageDropdown(hwnd)
            continue
        Handy_SetCohereLanguage_PickFromOpenDropdown(hwnd, langName)
        if (Handy_WaitCohereLanguage(hwnd, langName))
            return true
        Send "{Escape}"
        Sleep 80
    }
    return false
}

; Open the AI model dropdown menu using keyboard navigation
Handy_OpenAiModelMenu(hwnd) {
    el := UIA.ElementFromHandle(hwnd)
    if !el {
        return false
    }

    ; Find anchor: primary "Check for updates" button
    anchor := 0
    try anchor := el.FindFirst({
        Type: 50000,
        ClassName: "transition-colors disabled:opacity-50 tabular-nums text-text/60 hover:text-text/80"
    })
    if (!anchor) {
        try anchor := el.FindFirst({ Type: 50000, Name: "Check for updates" })
    }
    if (!anchor) {
        try anchor := el.FindFirst({ Type: 50000, Name: "Verificar atualizações" })
    }

    ; Fallback: "Update available" anchor when a system update banner is shown
    if (!anchor) {
        try anchor := el.FindFirst({
            Type: 50000,
            ClassName: "transition-colors disabled:opacity-50 tabular-nums text-logo-primary hover:text-logo-primary/80 font-medium"
        })
    }
    if (!anchor) {
        try anchor := el.FindFirst({ Type: 50000, Name: "Update available" })
    }
    if (!anchor) {
        ; Last-resort: use technical condition path to reach the "Update available" button
        try anchor := el.ElementFromPath({ T: 33 }, { T: 33 }, { T: 33 }, { T: 33, CN: "BrowserRootView" }, { T: 33 }, { T: 33,
            CN: "EmbeddedBrowserFrameView" }, { T: 33, CN: "BrowserView" }, { T: 33, CN: "SidebarContentsSplitView" }, { T: 33 }, { T: 33 }, { T: 33 }, { T: 30 }, { T: 26 }, { T: 0,
                CN: "transition-colors disabled:opacity-50 tabular-nums text-logo-primary hover:text-logo-primary/80 font-medium" }
        )
    }

    if (!anchor) {
        return false
    }

    ; Focus anchor, Shift+Tab to model button, Enter to open menu
    try anchor.SetFocus()
    catch {
        try anchor.Click()
    }
    Sleep 60
    Send "+{Tab}"
    Sleep 60
    Send "{Enter}"

    ; Context menu can open slowly in Handy; wait for menu row(s) to actually exist.
    return Handy_WaitForAiModelMenuOpen(hwnd, 2500)
}

; Wait for AI model context menu rows to appear after opening the menu.
Handy_WaitForAiModelMenuOpen(hwnd, maxWaitMs := 2500) {
    global UIA
    start := A_TickCount
    pollMs := 100
    loop {
        if ((A_TickCount - start) >= maxWaitMs)
            return false
        el := UIA.ElementFromHandle(hwnd)
        if (el) {
            try {
                for btn in el.FindAll({ Type: 50000 }) {
                    cn := ""
                    try cn := btn.ClassName
                    if (InStr(cn, "w-full px-3 py-2 text-left") || InStr(cn, "w-full px-3 py-2 text-start"))
                        return true
                }
            }
        }
        Sleep pollMs
    }
}

; Find and click the AI model button by partial name match
Handy_ClickAiModel(hwnd, modelName) {
    el := UIA.ElementFromHandle(hwnd)
    if !el {
        return false
    }

    ; Model buttons have class containing "w-full px-3 py-2 text-left"
    ; and names starting with the model name (e.g., "Whisper Large Good accuracy...")
    ; Try to find by partial name match
    modelBtn := 0
    buttonCount := 0
    nameMatchNoClass := ""

    ; Strategy 1: Find button whose Name starts with modelName
    try {
        buttons := el.FindAll({ Type: 50000 })
        for btn in buttons {
            buttonCount++
            btnName := ""
            try btnName := btn.Name
            if (btnName != "" && InStr(btnName, modelName) = 1) {
                btnClass := ""
                try btnClass := btn.ClassName
                ; Menu items: w-full px-3 py-2 text-left (legacy) or text-start (new Handy UI); header: flex items-center gap-2
                if (InStr(btnClass, "w-full px-3 py-2 text-left") || InStr(btnClass, "w-full px-3 py-2 text-start") ||
                    InStr(btnClass, "flex items-center gap-2")) {
                    modelBtn := btn
                    break
                }
                if (nameMatchNoClass = "")
                    nameMatchNoClass := btnClass
            }
        }
    }

    if (!modelBtn)
        return false

    ; Click the model button
    try {
        modelBtn.Click()
        return true
    } catch as e {
        return false
    }
}

; Poll the AI model selection button until Name no longer contains "loading", or maxWaitMs elapses.
; Button: Type 50000, ClassName "flex items-center gap-2 hover:text-text/80 transition-colors "
; Returns true when loading text disappeared, false on timeout or if button not found.
Handy_WaitForModelReady(hwnd, maxWaitMs) {
    global UIA
    pollInterval := 250
    start := A_TickCount
    firstLog := true
    loop {
        if ((A_TickCount - start) >= maxWaitMs) {
            return false
        }
        el := UIA.ElementFromHandle(hwnd)
        if !el {
            Sleep pollInterval
            continue
        }
        btn := 0
        try btn := el.FindFirst({ Type: 50000, ClassName: "flex items-center gap-2 hover:text-text/80 transition-colors " })
        if (!btn) {
            if (firstLog) {
                firstLog := false
            }
            Sleep pollInterval
            continue
        }
        btnName := ""
        try btnName := btn.Name
        if (InStr(btnName, "loading") = 0) {
            return true
        }
        Sleep pollInterval
    }
}

; =============================================================================
; SelectAiModelInHandy() - Opens or closes the selector GUI (hotkey entry point)
; =============================================================================
; Select AI model in Handy via interactive GUI selector.
; Win+Alt+Shift+C toggles: open when closed, close when open.
; Targets the correct Handy instance by environment: work = Documents\Handy\handy.exe, home = any.
SelectAiModelInHandy() {
    global g_AiModelSelectorActive
    if (g_AiModelSelectorActive)
        AiModelSelector_Close()
    else
        ShowAiModelSelector()
}

; =============================================================================
; Helper: Show centered overlay banner (uses standard loading indicator; non-blocking).
; =============================================================================
ShowCenteredOverlay_Utils(text, duration := 1500, bgColor := BANNER_ACCENT_INTERMEDIATE) {
    ; centerOnHwnd 0 = foreground monitor (GetActiveMonitorWorkArea_StandardBar); same intent as prior WinGetID("A") path.
    StandardLoadingBar_Show(text, bgColor, { passive: true, centerOnHwnd: 0, textWidth: 500, fontSize: 17,
        passiveBgColor: bgColor })
    StandardLoadingBar_Hide(duration)
}

; =============================================================================
; Helper: Pre-movement warning (sound + 2s delay) before automated window changes.
; =============================================================================
PlayPreMovementWarning(targetName) {
    ScriptSoundPlay(A_ScriptDir . "\sounds\pre-movement.wav")
    ShowCenteredOverlay_Utils("✋ Hands off! Moving to " . targetName . "...", 2000, BANNER_ACCENT_INTERMEDIATE)
    Sleep 2000
}

; =============================================================================
; Standard loading bar (monitor-aware, show/update/hide lifecycle)
; Use for long-running shortcuts; replace ad-hoc banners/overlays with this.
; Supports passive (text-only) mode and ShowWithKeys for letter-keystroke commands.
; Semantic accent colors (colorblind accessibility): border only; background stays dark.
; =============================================================================
global BANNER_ACCENT_SUCCESS := "27AE60"      ; Dark green: positive / success
global BANNER_ACCENT_ERROR := "C0392B"        ; Red: negative / error
global BANNER_ACCENT_INTERMEDIATE := "F1C40F" ; Yellow: loading, actionable, neutral
global BANNER_ACCENT_INFO := "2980B9"         ; Blue: distinct from green/yellow for color vision (info / alternate mode)
; Monitor blackout countdown (Study Topic / focus dwell): unmistakable vs generic loading banners
global BANNER_BLACKOUT_PANEL := "4A148C"       ; Purple panel fill
global BANNER_BLACKOUT_BORDER := "FF9800"    ; Orange border + progress strip
global g_StandardLoadingBarGui := 0
global g_StandardLoadingBarValue := 0
global g_StandardLoadingBarIsKeysOverlay := false
global g_StandardLoadingBarKeysHotkeys := []
global g_StandardLoadingBarKeysTimeoutTimer := ""
global g_StandardLoadingBarBorderGui := 0
global g_StandardLoadingBarTrackTimer := ""
global g_StandardLoadingBarLastForegroundMonitorIdx := 0
global g_StandardLoadingBarTimedProgressTimer := ""
global g_StandardLoadingBarTimedProgressStartTick := 0
global g_StandardLoadingBarTimedProgressDurationMs := 0
global D2C_SUBMIT_MENU_TIMEOUT_MS := 5000
; Keys overlay: same escape stack as Outlook Copilot selector (#!+l) — *Escape alone misses when the bar is NA/no focus.
global g_StandardLoadingBarKeysEscapeUserCb := ""
global g_StandardLoadingBarKeysEscapeActive := false
global g_StandardLoadingBarEscPollPrev := false
global g_StandardLoadingBarKeysPollTimer := ""
global g_StandardLoadingBarKeysPollPrev := Map()
global g_StandardLoadingBarKeysPollCallbacks := Map()

; Return work area { left, top, right, bottom } for the monitor containing hwnd, or "" on failure.
GetWorkAreaForWindow_StandardBar(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return ""
    try {
        WinGetPos(&winX, &winY, &winW, &winH, "ahk_id " hwnd)
        centerX := winX + winW / 2
        centerY := winY + winH / 2
        n := MonitorGetCount()
        loop n {
            MonitorGet(A_Index, &L, &T, &R, &B)
            if (centerX >= L && centerX < R && centerY >= T && centerY < B) {
                MonitorGetWorkArea(A_Index, &wLeft, &wTop, &wRight, &wBottom)
                return { left: wLeft, top: wTop, right: wRight, bottom: wBottom }
            }
        }
    } catch {
    }
    return ""
}

GetActiveMonitorWorkArea_StandardBar(&left, &top, &right, &bottom) {
    left := top := 0
    right := A_ScreenWidth
    bottom := A_ScreenHeight
    activeWin := 0
    try activeWin := WinGetID("A")
    catch
        activeWin := 0
    MonitorGetWorkArea(1, &monitorLeft, &monitorTop, &monitorRight, &monitorBottom)
    mLeft := monitorLeft
    mTop := monitorTop
    mRight := monitorRight
    mBottom := monitorBottom
    if (activeWin && activeWin != 0) {
        rect := Buffer(16, 0)
        if (DllCall("GetWindowRect", "ptr", activeWin, "ptr", rect)) {
            winLeft := NumGet(rect, 0, "int")
            winTop := NumGet(rect, 4, "int")
            winRight := NumGet(rect, 8, "int")
            winBottom := NumGet(rect, 12, "int")
            centerX := winLeft + (winRight - winLeft) // 2
            centerY := winTop + (winBottom - winTop) // 2
            monitorCount := MonitorGetCount()
            loop monitorCount {
                idx := A_Index
                MonitorGetWorkArea(idx, &l, &t, &r, &b)
                if (centerX >= l && centerX <= r && centerY >= t && centerY <= b) {
                    mLeft := l
                    mTop := t
                    mRight := r
                    mBottom := b
                    break
                }
            }
        }
    }
    left := mLeft
    top := mTop
    right := mRight
    bottom := mBottom
}

; 1-based monitor index for the monitor containing the center of the foreground window; 1 if unknown.
GetMonitorIndexForForeground_StandardBar() {
    activeWin := 0
    try activeWin := WinGetID("A")
    catch
        activeWin := 0
    if (!activeWin)
        return 1
    rect := Buffer(16, 0)
    if (!DllCall("GetWindowRect", "ptr", activeWin, "ptr", rect))
        return 1
    winLeft := NumGet(rect, 0, "int")
    winTop := NumGet(rect, 4, "int")
    winRight := NumGet(rect, 8, "int")
    winBottom := NumGet(rect, 12, "int")
    centerX := winLeft + (winRight - winLeft) // 2
    centerY := winTop + (winBottom - winTop) // 2
    monitorCount := MonitorGetCount()
    loop monitorCount {
        idx := A_Index
        MonitorGetWorkArea(idx, &l, &t, &r, &b)
        if (centerX >= l && centerX <= r && centerY >= t && centerY <= b)
            return idx
    }
    return 1
}

StandardLoadingBar_StopActiveMonitorTracking() {
    global g_StandardLoadingBarTrackTimer, g_StandardLoadingBarLastForegroundMonitorIdx
    try SetTimer(g_StandardLoadingBarTrackTimer, 0)
    catch {
    }
    g_StandardLoadingBarTrackTimer := ""
    g_StandardLoadingBarLastForegroundMonitorIdx := 0
}

StandardLoadingBar_RepositionToActiveMonitor() {
    global g_StandardLoadingBarGui, g_StandardLoadingBarBorderGui
    if !IsObject(g_StandardLoadingBarGui)
        return
    GetActiveMonitorWorkArea_StandardBar(&ml, &mt, &mr, &mb)
    monitorWidth := mr - ml
    try {
        g_StandardLoadingBarGui.GetPos(, , &gw, &gh)
    } catch {
        return
    }
    guiX := Round(ml + (monitorWidth - gw) / 2)
    if (guiX < ml)
        guiX := ml
    if (guiX + gw > mr)
        guiX := mr - gw
    guiY := mt + 40
    if (IsObject(g_StandardLoadingBarBorderGui)) {
        borderWidth := 6
        try {
            g_StandardLoadingBarBorderGui.Move(guiX - borderWidth, guiY - borderWidth, gw + 2 * borderWidth, gh + 2 *
                borderWidth)
        } catch {
        }
    }
    try {
        g_StandardLoadingBarGui.Move(guiX, guiY)
        hwnd := g_StandardLoadingBarGui.Hwnd
        if (hwnd)
            DllCall("SetWindowPos", "Ptr", hwnd, "Ptr", 0, "Int", guiX, "Int", guiY, "Int", 0, "Int", 0, "UInt", 0x0015
            )
    } catch {
    }
}

StandardLoadingBar_TrackActiveMonitorTick() {
    global g_StandardLoadingBarGui, g_StandardLoadingBarLastForegroundMonitorIdx
    if !IsObject(g_StandardLoadingBarGui) {
        StandardLoadingBar_StopActiveMonitorTracking()
        return
    }
    newIdx := GetMonitorIndexForForeground_StandardBar()
    if (newIdx != g_StandardLoadingBarLastForegroundMonitorIdx) {
        g_StandardLoadingBarLastForegroundMonitorIdx := newIdx
        StandardLoadingBar_RepositionToActiveMonitor()
    }
}

StandardLoadingBar_Show(state := "Working...", barColor := BANNER_ACCENT_INTERMEDIATE, options := "") {
    global g_StandardLoadingBarGui, g_StandardLoadingBarValue, g_StandardLoadingBarIsKeysOverlay,
        g_StandardLoadingBarBorderGui
    try StandardLoadingBar_CloseKeysOverlay()
    try StandardLoadingBar_Hide(0)
    passive := options && options.HasProp("passive") && options.passive
    centerOnHwnd := options && options.HasProp("centerOnHwnd") ? options.centerOnHwnd : 0
    textWidth := options && options.HasProp("textWidth") ? options.textWidth : 0
    fontSize := options && options.HasProp("fontSize") ? options.fontSize : 17
    alpha := options && options.HasProp("alpha") ? options.alpha : 235
    passiveBgColor := options && options.HasProp("passiveBgColor") ? options.passiveBgColor : ""
    noBorder := options && options.HasProp("noBorder") ? options.noBorder : false
    promptKeys := options && options.HasProp("promptKeys") ? options.promptKeys : ""
    trackActiveMonitor := options && options.HasProp("trackActiveMonitor") && options.trackActiveMonitor
    manualProgress := options && options.HasProp("manualProgress") && options.manualProgress
    overlayBgColor := options && options.HasProp("overlayBgColor") && options.overlayBgColor != "" ? options.overlayBgColor :
        "1E1E2E"

    if (centerOnHwnd) {
        workArea := GetWorkAreaForWindow_StandardBar(centerOnHwnd)
        if (workArea != "") {
            ml := workArea.left
            mt := workArea.top
            mr := workArea.right
            mb := workArea.bottom
        } else
            GetActiveMonitorWorkArea_StandardBar(&ml, &mt, &mr, &mb)
    } else
        GetActiveMonitorWorkArea_StandardBar(&ml, &mt, &mr, &mb)
    monitorWidth := mr - ml
    monitorHeight := mb - mt
    barWidth := textWidth > 0 ? textWidth : Min(900, Max(360, Floor(monitorWidth * 0.6)))
    overlayGui := Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale")
    overlayGui.BackColor := overlayBgColor
    overlayGui.MarginX := 16
    overlayGui.MarginY := 10
    overlayGui.SetFont("s" . fontSize . " cFFFFFF", "Segoe UI")
    overlayGui.Add("Text", "w" . barWidth . (passive ? " Wrap Center" : " Center"), state)
    if (promptKeys != "") {
        overlayGui.SetFont("s" . fontSize . " cFFFFFF", "Segoe UI")
        ; Wrap so long key strips (e.g. Gemini Copy response? + [F]) are not clipped at fixed width.
        overlayGui.Add("Text", "xm w" . barWidth . " Center Wrap", promptKeys)
    }
    if (!passive) {
        progressOpts := "w" . barWidth . " h10 c" . barColor . " Background45475A Smooth vOverlayProg"
        overlayGui.Add("Progress", progressOpts, 0)
    }
    overlayGui.Show("AutoSize Hide")
    overlayGui.GetPos(, , &gw, &gh)
    guiX := Round(ml + (monitorWidth - gw) / 2)
    if (guiX < ml)
        guiX := ml
    if (guiX + gw > mr)
        guiX := mr - gw
    guiY := mt + 40

    ; Create border frame behind the overlay for visibility (optional; skip when noBorder to show a single banner). Accent color when passiveBgColor set, else yellow.
    if (!noBorder) {
        borderWidth := 6
        try {
            if IsObject(g_StandardLoadingBarBorderGui)
                g_StandardLoadingBarBorderGui.Destroy()
        } catch {
        }
        borderGui := Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale")
        borderGui.BackColor := (passiveBgColor != "") ? passiveBgColor : BANNER_ACCENT_INTERMEDIATE
        borderGui.Show("NA x" . (guiX - borderWidth) . " y" . (guiY - borderWidth) . " w" . (gw + 2 * borderWidth) .
            " h" .
            (gh + 2 * borderWidth))
        g_StandardLoadingBarBorderGui := borderGui
    } else {
        try {
            if IsObject(g_StandardLoadingBarBorderGui)
                g_StandardLoadingBarBorderGui.Destroy()
        } catch {
        }
        g_StandardLoadingBarBorderGui := 0
    }
    overlayGui.Show("x" . guiX . " y" . guiY . " NA")
    try {
        hwnd := overlayGui.Hwnd
        if (hwnd)
            DllCall("SetWindowPos", "Ptr", hwnd, "Ptr", 0, "Int", guiX, "Int", guiY, "Int", 0, "Int", 0, "UInt", 0x0015
            )
    }
    WinSetTransparent(alpha, overlayGui)
    g_StandardLoadingBarGui := overlayGui
    g_StandardLoadingBarValue := 0
    if (!passive && !manualProgress)
        SetTimer(StandardLoadingBar_Tick, 40)
    if (trackActiveMonitor) {
        StandardLoadingBar_StopActiveMonitorTracking()
        g_StandardLoadingBarLastForegroundMonitorIdx := GetMonitorIndexForForeground_StandardBar()
        g_StandardLoadingBarTrackTimer := SetTimer(StandardLoadingBar_TrackActiveMonitorTick, 115)
    }
}

StandardLoadingBar_Tick() {
    global g_StandardLoadingBarGui, g_StandardLoadingBarValue
    if !IsObject(g_StandardLoadingBarGui) {
        SetTimer(StandardLoadingBar_Tick, 0)
        return
    }
    try {
        g_StandardLoadingBarValue += 4
        if (g_StandardLoadingBarValue > 100)
            g_StandardLoadingBarValue := 0
        g_StandardLoadingBarGui["OverlayProg"].Value := g_StandardLoadingBarValue
    } catch {
        SetTimer(StandardLoadingBar_Tick, 0)
    }
}

StandardLoadingBar_StopTimedProgress() {
    global g_StandardLoadingBarTimedProgressTimer, g_StandardLoadingBarTimedProgressStartTick,
        g_StandardLoadingBarTimedProgressDurationMs
    try SetTimer(g_StandardLoadingBarTimedProgressTimer, 0)
    catch {
    }
    g_StandardLoadingBarTimedProgressTimer := ""
    g_StandardLoadingBarTimedProgressStartTick := 0
    g_StandardLoadingBarTimedProgressDurationMs := 0
}

StandardLoadingBar_SetProgressValue(value) {
    global g_StandardLoadingBarGui, g_StandardLoadingBarValue
    if !IsObject(g_StandardLoadingBarGui)
        return
    clamped := Max(0, Min(100, Round(value)))
    g_StandardLoadingBarValue := clamped
    try g_StandardLoadingBarGui["OverlayProg"].Value := clamped
    catch {
    }
}

StandardLoadingBar_StartTimedProgress(durationMs) {
    global g_StandardLoadingBarTimedProgressTimer, g_StandardLoadingBarTimedProgressStartTick,
        g_StandardLoadingBarTimedProgressDurationMs
    StandardLoadingBar_StopTimedProgress()
    SetTimer(StandardLoadingBar_Tick, 0)
    if (durationMs <= 0) {
        StandardLoadingBar_SetProgressValue(100)
        return
    }
    g_StandardLoadingBarTimedProgressStartTick := A_TickCount
    g_StandardLoadingBarTimedProgressDurationMs := durationMs
    StandardLoadingBar_SetProgressValue(0)
    g_StandardLoadingBarTimedProgressTimer := SetTimer(StandardLoadingBar_TimedProgressTick, 40)
}

StandardLoadingBar_TimedProgressTick() {
    global g_StandardLoadingBarGui, g_StandardLoadingBarTimedProgressStartTick,
        g_StandardLoadingBarTimedProgressDurationMs
    if !IsObject(g_StandardLoadingBarGui) {
        StandardLoadingBar_StopTimedProgress()
        return
    }
    durationMs := g_StandardLoadingBarTimedProgressDurationMs
    if (durationMs <= 0) {
        StandardLoadingBar_SetProgressValue(100)
        StandardLoadingBar_StopTimedProgress()
        return
    }
    elapsedMs := A_TickCount - g_StandardLoadingBarTimedProgressStartTick
    if (elapsedMs >= durationMs) {
        StandardLoadingBar_SetProgressValue(100)
        StandardLoadingBar_StopTimedProgress()
        return
    }
    StandardLoadingBar_SetProgressValue((elapsedMs * 100.0) / durationMs)
}

StandardLoadingBar_Update(state := "", barColor := "") {
    global g_StandardLoadingBarGui
    if !IsObject(g_StandardLoadingBarGui)
        return
    try {
        if (state != "" && g_StandardLoadingBarGui.Controls.Length > 0)
            g_StandardLoadingBarGui.Controls[1].Text := state
    } catch {
    }
    if (barColor != "") {
        try
            g_StandardLoadingBarGui["OverlayProg"].Opt("c" . barColor)
        catch {
        }
    }
}

StandardLoadingBar_Hide(delayMs := 0) {
    global g_StandardLoadingBarGui, g_StandardLoadingBarValue, g_StandardLoadingBarIsKeysOverlay,
        g_StandardLoadingBarBorderGui
    if (delayMs > 0) {
        SetTimer(() => StandardLoadingBar_Hide(0), -delayMs)
        return
    }
    StandardLoadingBar_StopActiveMonitorTracking()
    StandardLoadingBar_StopTimedProgress()
    if (g_StandardLoadingBarIsKeysOverlay) {
        StandardLoadingBar_CloseKeysOverlay()
        return
    }
    SetTimer(StandardLoadingBar_Tick, 0)
    try {
        if IsObject(g_StandardLoadingBarGui)
            g_StandardLoadingBarGui.Destroy()
    } catch {
    }
    g_StandardLoadingBarGui := 0
    g_StandardLoadingBarValue := 0
    try {
        if IsObject(g_StandardLoadingBarBorderGui)
            g_StandardLoadingBarBorderGui.Destroy()
    } catch {
    }
    g_StandardLoadingBarBorderGui := 0
}

; Unregister keys and timeout timer for the "ShowWithKeys" overlay, then hide. Idempotent.
StandardLoadingBar_CloseKeysOverlay() {
    global g_StandardLoadingBarKeysHotkeys, g_StandardLoadingBarKeysTimeoutTimer
    global g_StandardLoadingBarGui, g_StandardLoadingBarValue, g_StandardLoadingBarIsKeysOverlay,
        g_StandardLoadingBarBorderGui
    global g_StandardLoadingBarKeysEscapeActive, g_OnEscapePressed, g_StandardLoadingBarKeysEscapeUserCb
    g_StandardLoadingBarIsKeysOverlay := false
    StandardLoadingBar_StopKeysSelectionPoll()
    hadEscStack := g_StandardLoadingBarKeysEscapeActive
    g_StandardLoadingBarKeysEscapeActive := false
    g_StandardLoadingBarKeysEscapeUserCb := ""
    g_StandardLoadingBarEscPollPrev := false
    try SetTimer(StandardLoadingBar_KeysEscapePoll, 0)
    catch {
    }
    try SetTimer(g_StandardLoadingBarKeysTimeoutTimer, 0)
    catch {
    }
    g_StandardLoadingBarKeysTimeoutTimer := ""
    StandardLoadingBar_StopActiveMonitorTracking()
    StandardLoadingBar_StopTimedProgress()
    for key in g_StandardLoadingBarKeysHotkeys {
        try Hotkey(key, "Off")
        catch {
        }
    }
    g_StandardLoadingBarKeysHotkeys := []
    if (hadEscStack) {
        g_OnEscapePressed := ""
        Utils_EnsureGlobalEscapeHotkey()
    }
    SetTimer(StandardLoadingBar_Tick, 0)
    try {
        if IsObject(g_StandardLoadingBarGui)
            g_StandardLoadingBarGui.Destroy()
    } catch {
    }
    g_StandardLoadingBarGui := 0
    g_StandardLoadingBarValue := 0
    try {
        if IsObject(g_StandardLoadingBarBorderGui)
            g_StandardLoadingBarBorderGui.Destroy()
    } catch {
    }
    g_StandardLoadingBarBorderGui := 0
}

; Return the Escape cancel callback from keyCallbacks, if any (used for robust Esc handling).
StandardLoadingBar_EscapeCallbackFromKeyCallbacks(keyCallbacks) {
    if (keyCallbacks is Map) {
        if keyCallbacks.Has("*Escape")
            return keyCallbacks["*Escape"]
        if keyCallbacks.Has("Escape")
            return keyCallbacks["Escape"]
        return ""
    }
    try {
        for kn, cb in keyCallbacks {
            if (!cb)
                continue
            k := StrLower(Trim(kn))
            if (k = "*escape" || k = "escape")
                return cb
        }
    } catch {
    }
    return ""
}

StandardLoadingBar_KeysSelectionModifiersDown() {
    try {
        ; Chord modifiers only (not Shift): WaitForTriggerKeyRelease already waits for Shift; including Shift here
        ; blocked poll edges while Shift was still down after #!+… chords.
        return GetKeyState("LWin", "P") || GetKeyState("RWin", "P") || GetKeyState("Ctrl", "P") || GetKeyState("Alt",
            "P")
    } catch {
        return false
    }
}

; AHK v2: do not use keyName >= "0" on letters — throws "Expected a Number but got a String".
StandardLoadingBar_IsDigitKey(keyName) {
    if (StrLen(keyName) != 1)
        return false
    o := Ord(keyName)
    return o >= Ord("0") && o <= Ord("9")
}

StandardLoadingBar_KeysSelectionKeyDown(keyName) {
    try {
        if (StandardLoadingBar_IsDigitKey(keyName))
            return GetKeyState(keyName, "P") || GetKeyState("Numpad" . keyName, "P")
        return GetKeyState(keyName, "P")
    } catch {
        return false
    }
}

StandardLoadingBar_StopKeysSelectionPoll() {
    global g_StandardLoadingBarKeysPollTimer, g_StandardLoadingBarKeysPollPrev, g_StandardLoadingBarKeysPollCallbacks
    try SetTimer(g_StandardLoadingBarKeysPollTimer, 0)
    catch {
    }
    g_StandardLoadingBarKeysPollTimer := ""
    g_StandardLoadingBarKeysPollPrev := Map()
    g_StandardLoadingBarKeysPollCallbacks := Map()
}

StandardLoadingBar_StartKeysSelectionPoll(keyCallbacks) {
    global g_StandardLoadingBarKeysPollTimer, g_StandardLoadingBarKeysPollPrev, g_StandardLoadingBarKeysPollCallbacks
    StandardLoadingBar_StopKeysSelectionPoll()
    g_StandardLoadingBarKeysPollCallbacks := Map()
    g_StandardLoadingBarKeysPollPrev := Map()
    try {
        for keyName, cb in keyCallbacks {
            if (!cb)
                continue
            knL := StrLower(Trim(keyName))
            if (knL = "escape" || knL = "*escape")
                continue
            g_StandardLoadingBarKeysPollCallbacks[keyName] := cb
            g_StandardLoadingBarKeysPollPrev[keyName] := StandardLoadingBar_KeysSelectionKeyDown(keyName)
        }
    } catch {
    }
    if (g_StandardLoadingBarKeysPollCallbacks.Count = 0)
        return
    g_StandardLoadingBarKeysPollTimer := SetTimer(StandardLoadingBar_KeysSelectionPoll, 50)
}

; Edge-triggered poll: survives Hotkey("1") conflicts from other scripts (see CursorTransfer selector loop).
StandardLoadingBar_KeysSelectionPoll() {
    global g_StandardLoadingBarIsKeysOverlay, g_StandardLoadingBarKeysPollPrev, g_StandardLoadingBarKeysPollCallbacks
    if (!g_StandardLoadingBarIsKeysOverlay) {
        StandardLoadingBar_StopKeysSelectionPoll()
        return
    }
    if (StandardLoadingBar_KeysSelectionModifiersDown()) {
        for keyName, cb in g_StandardLoadingBarKeysPollCallbacks {
            if (!cb)
                continue
            g_StandardLoadingBarKeysPollPrev[keyName] := StandardLoadingBar_KeysSelectionKeyDown(keyName)
        }
        return
    }
    for keyName, cb in g_StandardLoadingBarKeysPollCallbacks {
        if (!cb)
            continue
        isDown := StandardLoadingBar_KeysSelectionKeyDown(keyName)
        wasDown := g_StandardLoadingBarKeysPollPrev.Has(keyName) ? g_StandardLoadingBarKeysPollPrev[keyName] : false
        g_StandardLoadingBarKeysPollPrev[keyName] := isDown
        if (isDown && !wasDown)
            StandardLoadingBar_KeyWrapper(keyName, cb)
    }
}

; After chord release, wait until digit selection keys are up so poll arming does not treat a held key as wasDown.
StandardLoadingBar_WaitForSelectionKeysRelease(keyCallbacks) {
    try {
        for keyName, cb in keyCallbacks {
            if (!cb)
                continue
            knL := StrLower(Trim(keyName))
            if (knL = "escape" || knL = "*escape")
                continue
            if (StandardLoadingBar_IsDigitKey(keyName)) {
                while StandardLoadingBar_KeysSelectionKeyDown(keyName)
                    KeyWait keyName
            }
        }
    } catch {
    }
}

; After a chord hotkey (e.g. #!+w), wait until Win/Ctrl/Alt/Shift and the trigger key are released
; so *1 / *2 selection hotkeys are not swallowed while modifiers are still held.
StandardLoadingBar_WaitForTriggerKeyRelease() {
    try {
        if (A_ThisHotkey = "")
            return
        th := A_ThisHotkey
        if InStr(th, "#") {
            try KeyWait "LWin"
            try KeyWait "RWin"
        }
        if InStr(th, "^")
            try KeyWait "Ctrl"
        if InStr(th, "!")
            try KeyWait "Alt"
        if InStr(th, "+")
            try KeyWait "Shift"
        hk := th
        hk := StrReplace(hk, "+", "")
        hk := StrReplace(hk, "^", "")
        hk := StrReplace(hk, "!", "")
        hk := StrReplace(hk, "#", "")
        if (StrLen(hk) = 1)
            KeyWait hk
    } catch {
    }
}

; Poll Esc — fallback when $*Escape / g_OnEscapePressed miss (same idea as OutlookCopilotSelector_EscapePoll).
StandardLoadingBar_KeysEscapePoll() {
    global g_StandardLoadingBarKeysEscapeActive, g_StandardLoadingBarIsKeysOverlay, g_StandardLoadingBarEscPollPrev
    if (!g_StandardLoadingBarKeysEscapeActive || !g_StandardLoadingBarIsKeysOverlay) {
        try SetTimer(StandardLoadingBar_KeysEscapePoll, 0)
        catch {
        }
        return
    }
    escDown := GetKeyState("Escape", "P") || (DllCall("user32\GetAsyncKeyState", "int", 0x1B) & 0x8000)
    if escDown {
        if !g_StandardLoadingBarEscPollPrev {
            g_StandardLoadingBarEscPollPrev := true
            StandardLoadingBar_KeysEscapeDismiss()
        }
    } else
        g_StandardLoadingBarEscPollPrev := false
}

; $*Escape, I10 g_OnEscapePressed, Gui Escape, and poll all route here (align with #!+l Outlook Copilot selector).
StandardLoadingBar_KeysEscapeDismiss(*) {
    global g_StandardLoadingBarIsKeysOverlay, g_StandardLoadingBarKeysEscapeUserCb
    if !g_StandardLoadingBarIsKeysOverlay
        return
    cb := g_StandardLoadingBarKeysEscapeUserCb
    if cb {
        try cb.Call()
        catch {
        }
    }
    StandardLoadingBar_CloseKeysOverlay()
}

; Show overlay and register hotkeys; optional timeout. keyCallbacks: Map/object key -> callback (e.g. "N" -> fn, "R" -> fn).
; timeoutCallback: called when timeout fires (can be empty). Registers both upper and lower case for letter keys.
; passiveBgColor: optional; when set, used as border color. Prefer BANNER_ACCENT_SUCCESS / BANNER_ACCENT_ERROR / BANNER_ACCENT_INTERMEDIATE. Overlay background stays dark.
; noBorder: when true, do not create the yellow border (single banner only).
; promptKeys: optional; fixed bottom strip text (e.g. "[Y] Confirm  [N] Cancel"). Shown in uniform position below main message.
; trackActiveMonitor: when true, reposition the bar to follow the foreground window's monitor while visible (dictation/Gemini flows).
; showProgress: when true, show a single timed 0-100 progress fill while waiting for keys.
; preserveUserFocus: when true, keep the current active window focused (do not activate overlay GUI).
; overlayBgColor: optional main banner panel color (default dark 1E1E2E); use for themed banners e.g. blackout countdown.
; skipEscapeDismiss: when true, do not register $*Escape / poll (fragile UIs e.g. Command Palette bookmark prompt).
StandardLoadingBar_ShowWithKeys(state, keyCallbacks, timeoutMs := 0, centerOnHwnd := 0, timeoutCallback := "", barColor :=
    BANNER_ACCENT_INTERMEDIATE, textWidth := 500, fontSize := 17, passiveBgColor := "", noBorder := false, promptKeys :=
    "", trackActiveMonitor := false, showProgress := false, preserveUserFocus := false, overlayBgColor := "",
    skipEscapeDismiss := false) {
    global g_StandardLoadingBarIsKeysOverlay, g_StandardLoadingBarKeysHotkeys, g_StandardLoadingBarKeysTimeoutTimer
    global g_StandardLoadingBarGui, g_StandardLoadingBarKeysEscapeUserCb, g_StandardLoadingBarKeysEscapeActive,
        g_StandardLoadingBarEscPollPrev, g_OnEscapePressed
    opts := { passive: !showProgress, centerOnHwnd: centerOnHwnd, textWidth: textWidth, fontSize: fontSize }
    if (showProgress)
        opts.manualProgress := true
    if (overlayBgColor != "")
        opts.overlayBgColor := overlayBgColor
    if (passiveBgColor != "")
        opts.passiveBgColor := passiveBgColor
    if (noBorder)
        opts.noBorder := true
    if (promptKeys != "")
        opts.promptKeys := promptKeys
    if (trackActiveMonitor)
        opts.trackActiveMonitor := true
    StandardLoadingBar_Show(state, barColor, opts)
    if (showProgress)
        StandardLoadingBar_StartTimedProgress(timeoutMs)
    g_StandardLoadingBarIsKeysOverlay := true
    g_StandardLoadingBarKeysHotkeys := []
    escCb := StandardLoadingBar_EscapeCallbackFromKeyCallbacks(keyCallbacks)
    g_StandardLoadingBarKeysEscapeActive := false

    StandardLoadingBar_WaitForTriggerKeyRelease()
    StandardLoadingBar_WaitForSelectionKeysRelease(keyCallbacks)

    ; Register selection hotkeys as GLOBAL while the overlay is open.
    ; Critical: the overlay may fail to activate immediately, and we still need the keys (e.g. "N") to be captured
    ; instead of falling through to the underlying app. Also, clear any caller #HotIf context before registering.
    try HotIf()
    catch {
    }

    ; Register primary and case-variant keys.
    for keyName, cb in keyCallbacks {
        if (!cb)
            continue
        if (!skipEscapeDismiss) {
            knL := StrLower(Trim(keyName))
            if (knL = "*escape" || knL = "escape")
                continue
        }
        StandardLoadingBar_RegisterKeyHandler(keyName, cb)
        if (StrLen(keyName) = 1) {
            o := Ord(keyName)
            alt := ""
            if (o >= Ord("a") && o <= Ord("z"))
                alt := StrUpper(keyName)
            else if (o >= Ord("A") && o <= Ord("Z"))
                alt := StrLower(keyName)
            if (alt != "" && alt != keyName)
                StandardLoadingBar_RegisterKeyHandler(alt, cb)
            if (StandardLoadingBar_IsDigitKey(keyName))
                StandardLoadingBar_RegisterKeyHandler("Numpad" . keyName, cb)
        }
    }

    if (!skipEscapeDismiss) {
        g_StandardLoadingBarKeysEscapeUserCb := escCb ? escCb : ""
        try {
            Hotkey("$*Escape", StandardLoadingBar_KeysEscapeDismiss, "On")
            g_StandardLoadingBarKeysHotkeys.Push("$*Escape")
        } catch {
        }
        g_OnEscapePressed := StandardLoadingBar_KeysEscapeDismiss
        Utils_EnsureGlobalEscapeHotkey()
        try {
            if IsObject(g_StandardLoadingBarGui)
                g_StandardLoadingBarGui.OnEvent("Escape", StandardLoadingBar_KeysEscapeDismiss)
        } catch {
        }
        g_StandardLoadingBarEscPollPrev := GetKeyState("Escape", "P") || (DllCall("user32\GetAsyncKeyState", "int",
            0x1B) &
            0x8000)
        SetTimer(StandardLoadingBar_KeysEscapePoll, 50)
        g_StandardLoadingBarKeysEscapeActive := true
    }

    ; Default behavior keeps key capture reliable by activating the overlay.
    ; Some flows (e.g. dictation E/V paste target) must preserve the user's current text-field focus.
    if (!preserveUserFocus) {
        try {
            if IsObject(g_StandardLoadingBarGui) && g_StandardLoadingBarGui.Hwnd
                WinActivate(g_StandardLoadingBarGui.Hwnd)
        } catch {
        }
    }

    ; Reset any HotIf context so we don't leak it to unrelated hotkeys.
    try HotIf()
    catch {
    }

    if (timeoutMs > 0) {
        g_StandardLoadingBarKeysTimeoutTimer := SetTimer(StandardLoadingBar_KeysTimeoutFired.Bind(timeoutCallback), -
            timeoutMs)
    }

    StandardLoadingBar_StartKeysSelectionPoll(keyCallbacks)
}

StandardLoadingBar_RegisterKeyHandler(key, cb) {
    global g_StandardLoadingBarKeysHotkeys
    if (!cb)
        return
    ; $* prefix: hook hotkey, ignore sent keys; * allows extra modifiers (Shift) still held.
    keyToReg := key
    if (StrLen(key) = 1 || RegExMatch(key, "i)^Numpad[0-9]$"))
        keyToReg := "$*" . key
    fn := StandardLoadingBar_KeyWrapper.Bind(key, cb)
    try {
        #InputLevel 10
        Hotkey(keyToReg, fn, "On")
        #InputLevel 0
        g_StandardLoadingBarKeysHotkeys.Push(keyToReg)
    } catch as err {
    }
}

StandardLoadingBar_KeyWrapper(key, cb, *) {
    global g_StandardLoadingBarIsKeysOverlay
    if (!g_StandardLoadingBarIsKeysOverlay)
        return
    ; Run callback first so it can close the overlay (avoids destroying GUI from hotkey context before callback runs).
    if (cb) {
        try {
            cb.Call()
        }
        catch {
        }
    }
    ; Callback may have closed the keys overlay and started a loading bar — do not destroy the replacement GUI.
    if (g_StandardLoadingBarIsKeysOverlay)
        StandardLoadingBar_CloseKeysOverlay()
}

StandardLoadingBar_KeysTimeoutFired(timeoutCallback) {
    global g_StandardLoadingBarIsKeysOverlay
    ; Nested Func refs from hotkey closures can fail a plain `if (timeoutCallback)` truth test in v2; use HasMethod.
    cbCallable := false
    try {
        if (IsObject(timeoutCallback))
            cbCallable := HasMethod(timeoutCallback, "Call")
    } catch {
        cbCallable := false
    }
    ; Only run timeout callback if overlay was not already dismissed (e.g. user pressed N); avoids copy when timer fires after cancel.
    if (g_StandardLoadingBarIsKeysOverlay && cbCallable) {
        StandardLoadingBar_SetProgressValue(100)
        try timeoutCallback.Call()
        catch {
        }
    }
    StandardLoadingBar_CloseKeysOverlay()
}

; =============================================================================
; Hotstring Selector: Gemini Redirect Banner (non-blocking; uses standard loading indicator)
; =============================================================================
HotstringGeminiBanner_Show(text := "📤 Gemini: inserting prompt...") {
    StandardLoadingBar_CloseKeysOverlay()
    Sleep 50
    StandardLoadingBar_Show(text, BANNER_ACCENT_INTERMEDIATE, { passive: true, centerOnHwnd: 0, textWidth: 280,
        fontSize: 17,
        alpha: 204 })
}

HotstringGeminiBanner_Hide(*) {
    StandardLoadingBar_Hide(0)
}

; Dictation → Gemini: join preset prompt and dictated text for InsertText into Gemini prompt field.
D2C_CombinePresetWithDictation(presetText, dictationText) {
    p := Trim(presetText)
    d := Trim(dictationText)
    if (p = "")
        return d
    if (d = "")
        return p
    return p . "`n`n" . d
}

; Buffer after dictation ends before "Send to Gemini?" (linear loading bar + accidental key buffer).
D2C_SUBMIT_MENU_DELAY_MS := 2000
D2C_SUBMIT_MENU_PROGRESS_TICK_MS := 50

; Linear 0–100% loading bar for D2C_SUBMIT_MENU_DELAY_MS, then hide (no keys). Stops StandardLoadingBar_Tick spinner.
D2C_RunSubmitMenuDelayBar() {
    global g_StandardLoadingBarGui
    delayMs := D2C_SUBMIT_MENU_DELAY_MS
    tickMs := D2C_SUBMIT_MENU_PROGRESS_TICK_MS
    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Hide(0)
    StandardLoadingBar_Show("⏳ Preparing menu…", BANNER_ACCENT_INTERMEDIATE, { textWidth: 520, fontSize: 17,
        noBorder: true, trackActiveMonitor: true })
    SetTimer(StandardLoadingBar_Tick, 0)
    try {
        if IsObject(g_StandardLoadingBarGui)
            g_StandardLoadingBarGui["OverlayProg"].Value := 0
    } catch {
    }
    steps := delayMs // tickMs
    if (steps < 1)
        steps := 1
    loop steps {
        Sleep(tickMs)
        pct := Min(100, Round(100 * A_Index / steps))
        try {
            if IsObject(g_StandardLoadingBarGui)
                g_StandardLoadingBarGui["OverlayProg"].Value := pct
        } catch {
            break
        }
    }
    try {
        if IsObject(g_StandardLoadingBarGui)
            g_StandardLoadingBarGui["OverlayProg"].Value := 100
    } catch {
    }
    StandardLoadingBar_Hide(0)
}

TryGetSelectedTextViaUIA_QuickLook() {
    hwnd := WinExist("A")
    focused := 0
    try {
        focused := UIA.GetFocusedElement()
    } catch {
        focused := 0
    }

    if (focused && focused.IsTextPatternAvailable) {
        try {
            ranges := focused.TextPattern.GetSelection()
            if (IsObject(ranges) && ranges.Length >= 1) {
                txt := ranges[1].GetText(512)
                return Trim(txt)
            }
        } catch {
        }
    }

    try {
        root := UIA.ElementFromHandle(hwnd)
        doc := root.FindFirst({ Type: UIA.ControlType.Document })
        if (doc && doc.IsTextPatternAvailable) {
            try {
                ranges := doc.TextPattern.GetSelection()
                if (IsObject(ranges) && ranges.Length >= 1) {
                    txt := ranges[1].GetText(512)
                    if (Trim(txt) != "")
                        return Trim(txt)
                }
            } catch {
            }
            try {
                txt := doc.TextPattern.DocumentRange.GetText(512)
                return Trim(txt)
            } catch {
            }
        }
    } catch {
    }
    return ""
}

TryCopySelectionToClipboard_QuickLookAware() {
    proc := ""
    try {
        proc := WinGetProcessName("A")
    } catch {
        proc := ""
    }

    A_Clipboard := ""
    Send "^c"
    if ClipWait(0.7)
        return true

    A_Clipboard := ""
    Send "^{Insert}"
    if ClipWait(0.7)
        return true

    if (proc = "QuickLook.exe") {
        A_Clipboard := ""
        Send "{AppsKey}"
        Sleep 60
        Send "c"
        if ClipWait(0.9)
            return true

        txt := TryGetSelectedTextViaUIA_QuickLook()
        if (txt != "" && StrLen(Trim(txt)) > 0) {
            A_Clipboard := txt
            return true
        }
    }

    return false
}

; Heuristic pt/en/de when Python daemon or lingua is unavailable (#!+8 auto-detect timeout path).
DetectLang_AhkFallback(text) {
    static inited := false
    static ptMap := Map()
    static enMap := Map()
    static deMap := Map()
    if (!inited) {
        inited := true
        for w in StrSplit("que nao com para uma dos das por mas sao esta ser tem foi", " ") {
            if (w != "")
                ptMap[w] := true
        }
        for w in StrSplit("the and of to in is that it for on with as by this are", " ") {
            if (w != "")
                enMap[w] := true
        }
        for w in StrSplit("der die das und ist nicht ein eine mit auf fur sich von zu als", " ") {
            if (w != "")
                deMap[w] := true
        }
    }
    t := Trim(text)
    if (t = "")
        return "en"
    tl := StrLower(t)
    if RegExMatch(tl, "[ãõçáàâéêíóôú]")
        return "pt"
    if RegExMatch(tl, "[äöüß]")
        return "de"
    clean := RegExReplace(tl, "[^a-zà-ÿ]+", " ")
    scorePt := 0
    scoreEn := 0
    scoreDe := 0
    for word in StrSplit(RegExReplace(clean, " +", " "), " ") {
        if (word = "")
            continue
        if ptMap.Has(word)
            scorePt++
        if enMap.Has(word)
            scoreEn++
        if deMap.Has(word)
            scoreDe++
    }
    m := Max(scorePt, scoreEn, scoreDe)
    if (m = 0)
        return "en"
    wins := []
    if (scorePt = m)
        wins.Push("pt")
    if (scoreEn = m)
        wins.Push("en")
    if (scoreDe = m)
        wins.Push("de")
    if (wins.Length = 1)
        return wins[1]
    return "en"
}

#include %A_ScriptDir%\lib\CopilotWeb.ahk

; =============================================================================
; =============================================================================
; D2C_FlowManager: Unified state machine for Dictation → Gemini → Cursor flow.
; Replaces legacy fragmented functions with a central authority to prevent race conditions.
; =============================================================================
class D2C_FlowManager {
    static _instance := 0

    static GetInstance() {
        if (!D2C_FlowManager._instance)
            D2C_FlowManager._instance := D2C_FlowManager()
        return D2C_FlowManager._instance
    }

    __New() {
        this.Reset()
    }

    Reset() {
        this.CurrentPhase := "Idle"
        this.OriginHwnd := 0
        this.GeminiHwnd := 0
        this.CursorHwnd := 0
        this.MonitorTimer := ""
        this.MonitorRetryCount := 0
        this.MonitorMaxRetries := 300 ; 150s
        this.MonitorButtonEverFound := false
        this.MonitorLastCheckTick := 0
        this.HasCopiedForThisResponse := false
    }

    ; --- Entry Points ---

    StartFromDictation() {
        global g_D2C_DictationSubmitMenuCycleFinished
        if (g_D2C_DictationSubmitMenuCycleFinished)
            return
        if (this.CurrentPhase = "PromptingSubmit")
            return
        if (this.CurrentPhase != "Idle")
            return
        this.Reset()
        D2C_RunSubmitMenuDelayBar()
        this.OriginHwnd := 0
        try this.OriginHwnd := WinGetID("A")
        this.PromptForGeminiSubmit()
    }

    StartFromHotstring() {
        if (this.CurrentPhase != "Idle") {
            return
        }
        this.Reset()
        this.OriginHwnd := WinActive("A")
        this.ExecuteGeminiSubmit(true, "", true)
    }

    ; --- Phase 1: Submit Prompt ---

    PromptForGeminiSubmit() {
        this.CurrentPhase := "PromptingSubmit"
        keyCallbacks := Map(
            "G", this.OnSubmitG.Bind(this),
            "A", this.OnSubmitA.Bind(this),
            "Y", this.OnSubmitY.Bind(this),
            "S", this.OnSubmitS.Bind(this),
            "V", this.OnSubmitV.Bind(this),
            "E", this.OnSubmitE.Bind(this),
            "F", this.OnSubmitF.Bind(this),
            "O", this.OnSubmitO.Bind(this),
            "N", this.OnSubmitN.Bind(this)
        )
        aiLabel := GetGlobalAIProviderLabel()
        StandardLoadingBar_ShowWithKeys(
            "❓ Send to " . aiLabel . "? (5s)",
            keyCallbacks,
            D2C_SUBMIT_MENU_TIMEOUT_MS,
            0,
            this.OnSubmitTimeout.Bind(this),
            BANNER_ACCENT_INTERMEDIATE, 560, 17, "", true,
            "[G] Grammar  [A] AI opt  [Y] Send  [S] Paste only  [V] Paste dictated  [E] Paste & send  [F] Favorite  [O] Clip Angel  [N] Cancel",
            true,
            true,
            true
        )
    }

    ; OriginHwnd is set after the submit delay bar ends, before the keys overlay activates.
    ActivateOriginForPaste() {
        if (!this.OriginHwnd || !WinExist("ahk_id " this.OriginHwnd))
            return
        WinActivate("ahk_id " this.OriginHwnd)
        if (!WinActive("ahk_id " this.OriginHwnd))
            WinWaitActive("ahk_id " this.OriginHwnd, , 0.3)
    }

    OnSubmitG(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return
        this.ExecuteGeminiSubmit(true, "grammar")
    }

    OnSubmitA(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return
        this.ExecuteGeminiSubmit(true, "aiopt")
    }

    OnSubmitY(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return
        this.ExecuteGeminiSubmit(true)
    }

    OnSubmitS(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return
        this.ExecuteGeminiSubmit(false)
    }

    OnSubmitV(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return

        ; Lock target at keypress time: last text field the user selected before pressing V.
        targetHwnd := 0
        try targetHwnd := WinGetID("A")

        this.CurrentPhase := "PastingDictation"
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        HideDictationIndicator()

        if (targetHwnd && WinExist("ahk_id " targetHwnd)) {
            try WinActivate("ahk_id " targetHwnd)
            if (!WinActive("ahk_id " targetHwnd))
                WinWaitActive("ahk_id " targetHwnd, , 0.2)
        }
        Sleep 60
        Send("^v")

        global g_D2C_DictationSubmitMenuCycleFinished
        g_D2C_DictationSubmitMenuCycleFinished := true
        this.Reset()
    }

    ; Paste at the field that was focused when E was pressed, then Enter (no Gemini).
    OnSubmitE(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return

        ; Lock target at keypress time: last text field the user selected before pressing E.
        targetHwnd := 0
        try targetHwnd := WinGetID("A")

        this.CurrentPhase := "PastingSendHere"
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        HideDictationIndicator()

        if (targetHwnd && WinExist("ahk_id " targetHwnd)) {
            try WinActivate("ahk_id " targetHwnd)
            if (!WinActive("ahk_id " targetHwnd))
                WinWaitActive("ahk_id " targetHwnd, , 0.2)
        }
        Sleep 60
        Send("^v")
        Sleep 150
        Send("{Enter}")

        global g_D2C_DictationSubmitMenuCycleFinished
        g_D2C_DictationSubmitMenuCycleFinished := true
        this.Reset()
    }

    ; [F] Mark newest Clip Angel clip (dictation) as favorite; no Gemini.
    OnSubmitF(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        HideDictationIndicator()
        try {
            clipRaw := A_Clipboard
            clip := Trim(clipRaw)
            if (clip = "" || StrLen(clip) < 10) {
                ShowCenteredOverlay_Utils("❌ Nothing to favorite - clipboard empty or too short", 2000,
                    BANNER_ACCENT_ERROR)
                if (this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd))
                    WinActivate("ahk_id " this.OriginHwnd)
                return
            }
            MarkLastClipAsFavorite("first", true)
        } finally {
            global g_D2C_DictationSubmitMenuCycleFinished
            g_D2C_DictationSubmitMenuCycleFinished := true
            this.Reset()
        }
    }

    ; [O] Open Clip Angel (Row 0), then Edit text (F4 — same as Shift+E in Shift keys.ahk for ClipAngel). O avoids C = Transfer on Copy response? banner.
    OnSubmitO(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return
        StandardLoadingBar_CloseKeysOverlay()
        HideDictationIndicator()

        StandardLoadingBar_Show("⏳ Clip Angel: opening...", BANNER_ACCENT_INTERMEDIATE)
        try {
            ; Use the origin window (what the user was looking at) to decide the target monitor for ClipAngel.
            originHwnd := this.OriginHwnd
            if (!originHwnd)
                try originHwnd := WinGetID("A")
            originMon := GetAhkMonitorIndexFromHwnd(originHwnd)

            ActivateClipAngelWithFocusCorrection()
            clipHwnd := WinExist("ClipAngel")
            if (!clipHwnd) {
                StandardLoadingBar_Update("❌ Clip Angel: window not found", BANNER_ACCENT_ERROR)
                return
            }

            ; Activation is already handled inside ActivateClipAngelWithFocusCorrection(), but keep a short bounded wait
            ; so we never inject keys into the wrong window.
            if (!WinWaitActive("ahk_id " clipHwnd, , 0.6)) {
                StandardLoadingBar_Update("❌ Clip Angel: failed to activate", BANNER_ACCENT_ERROR)
                return
            }

            if (originMon) {
                StandardLoadingBar_Update("⏳ Clip Angel: moving to your monitor...", BANNER_ACCENT_INTERMEDIATE)
                MoveWindowToMonitor(clipHwnd, originMon)
                ; If a move caused focus loss, reacquire quickly (bounded).
                if (!WinActive("ahk_id " clipHwnd))
                    WinWaitActive("ahk_id " clipHwnd, , 0.4)
            }

            StandardLoadingBar_Update("⏳ Clip Angel: opening editor...", BANNER_ACCENT_INTERMEDIATE)
            Send "{F4}"

            StandardLoadingBar_Update("⏳ Clip Angel: maximizing...", BANNER_ACCENT_INTERMEDIATE)
            TryMaximizeWindow(clipHwnd)
            StandardLoadingBar_Update("✅ Clip Angel: ready", BANNER_ACCENT_SUCCESS)
        } finally {
            StandardLoadingBar_Hide(350)
            global g_D2C_DictationSubmitMenuCycleFinished
            g_D2C_DictationSubmitMenuCycleFinished := true
            this.Reset()
        }
    }

    OnSubmitN(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return
        this.CancelFlow(GetGlobalAIProviderLabel() . " submission cancelled")
    }

    OnSubmitTimeout(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return
        this.CancelFlow(GetGlobalAIProviderLabel() . " submission cancelled")
    }

    ; --- Phase 2: Submit Execute ---

    ; presetMode: "" = Clip Angel first snippet; "grammar" | "aiopt" = preset from prompt/*.txt + clipboard dictation via InsertText.
    ; showPreMovementWarning: true only for non-banner-triggered submits (e.g., hotstring path).
    ExecuteGeminiSubmit(autoSubmit := true, presetMode := "", showPreMovementWarning := false) {
        this.CurrentPhase := "Submitting"
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        HideDictationIndicator()

        aiLabel := GetGlobalAIProviderLabel()
        ; For explicit first-banner choices (Y/G/A/S), skip the handoff cue: user intentionally chose the AI target.
        if (showPreMovementWarning)
            PlayPreMovementWarning(aiLabel)

        optionalSnippet := ""
        if (presetMode = "grammar" || presetMode = "aiopt") {
            dictation := ""
            try dictation := A_Clipboard
            preset := presetMode = "grammar" ? GetGrammarPromptText() : GetAioptPromptText()
            optionalSnippet := D2C_CombinePresetWithDictation(preset, dictation)
        }

        if (UseCopilotWebForGlobalAI()) {
            this.GeminiHwnd := CopilotWeb_NavigateFocusAndPaste(optionalSnippet, false)
            if (!this.GeminiHwnd)
                this.GeminiHwnd := GetCopilotWebWindowHwnd()
        } else {
            if (optionalSnippet != "")
                GeminiNavigateFocusAndPasteFirstSnippet(optionalSnippet, false)
            else
                GeminiNavigateFocusAndPasteFirstSnippet("", false)
            this.GeminiHwnd := WinExist("A")
        }

        if (autoSubmit) {
            Sleep 1000 ; Pre-enter delay
            endTick := A_TickCount + 5000
            while (A_TickCount < endTick) {
                hasContent := UseCopilotWebForGlobalAI()
                    ? (CopilotWeb_ComposerGetText(this.GeminiHwnd) != "")
                    : (GeminiPromptFieldGetText() != "")
                if (hasContent)
                    break
                Sleep 200
            }
            if (UseCopilotWebForGlobalAI()) {
                try {
                    uia := UIA_Browser("ahk_id " this.GeminiHwnd)
                    CopilotWeb_TrySubmit(uia)
                } catch {
                    Send("{Enter}")
                }
            } else
                Send("{Enter}")
            this.StartGeminiMonitor()
        }

        if (this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd))
            WinActivate("ahk_id " this.OriginHwnd)

        if (!autoSubmit)
            this.Reset()
    }

    ; --- Phase 3: Monitor ---

    StartGeminiMonitor() {
        this.CurrentPhase := "Monitoring"
        this.MonitorRetryCount := 0
        this.MonitorButtonEverFound := false
        this.MonitorTimer := this.CheckGeminiCompletion.Bind(this)
        SetTimer(this.MonitorTimer, 500)
    }

    CheckGeminiCompletion() {
        delta := this.MonitorLastCheckTick ? (A_TickCount - this.MonitorLastCheckTick) : -1
        this.MonitorLastCheckTick := A_TickCount
        if (this.CurrentPhase != "Monitoring") {
            SetTimer(this.MonitorTimer, 0)
            return
        }

        this.MonitorRetryCount++
        if (this.MonitorRetryCount > this.MonitorMaxRetries) {
            SetTimer(this.MonitorTimer, 0)
            this.Reset()
            return
        }

        btn := ""
        useCopilot := UseCopilotWebForGlobalAI()
        buttonNames := useCopilot ? ["Stop generating"] : ["Stop streaming", "Interromper transmissão", "Stop response"]
        root := 0
        try {
            if (useCopilot) {
                root := CopilotWeb_ReadRootFromHwnd(this.GeminiHwnd)
                if (root)
                    btn := CopilotWeb_FindStopGenerating(root)
            } else {
                root := UIA.ElementFromHandle(this.GeminiHwnd)
                for n in buttonNames {
                    try {
                        btn := root.FindElement({ Name: n, Type: "Button" })
                    } catch {
                        btn := ""
                    }
                    if (btn)
                        break
                }
            }
        } catch {
            return
        }

        if (btn) {
            this.MonitorButtonEverFound := true
            return
        }

        if (this.MonitorButtonEverFound) {
            ; Suspend timer to prevent re-entrancy during the 800ms Sleep block
            SetTimer(this.MonitorTimer, 0)

            isTrulyGone := true
            loop 4 {
                Sleep 200
                try {
                    if (useCopilot) {
                        copRoot := CopilotWeb_ReadRootFromHwnd(this.GeminiHwnd)
                        if (copRoot && CopilotWeb_FindStopGenerating(copRoot))
                            isTrulyGone := false
                    } else {
                        for n in buttonNames {
                            if root.ElementExist({ Name: n, Type: "Button" }) {
                                isTrulyGone := false
                                break
                            }
                        }
                    }
                } catch {
                    isTrulyGone := true
                }
                if (!isTrulyGone)
                    break
            }

            if (isTrulyGone) {
                ; Timer is already stopped, proceed to next phase
                try ScriptSoundPlay(A_ScriptDir . "\sounds\gemini-completion.wav")
                catch {
                    ; Ignore chime failures
                }
                this.PromptForResponseAction()
            } else {
                ; False alarm, the button is still there. Resume polling.
                SetTimer(this.MonitorTimer, 500)
            }
        }
    }

    ; --- Phase 4: Action Prompt ---

    PromptForResponseAction() {
        ; Prevent duplicate banner spawns from timer re-entrancy
        if (this.CurrentPhase = "PromptingAction") {
            return
        }
        this.CurrentPhase := "PromptingAction"
        keyCallbacks := Map(
            "Y", this.OnActionY.Bind(this),
            "C", this.OnActionC.Bind(this),
            "R", this.OnActionR.Bind(this),
            "N", this.OnActionN.Bind(this),
            "F", this.OnActionF.Bind(this)
        )
        pk := "[Y] Copy  [N] No  [R] Copy+Read  [C] Transfer  [F] Copy+Favorite"
        StandardLoadingBar_ShowWithKeys(
            "❓ Copy response? (5s)",
            keyCallbacks,
            D2C_SUBMIT_MENU_TIMEOUT_MS,
            0,
            this.OnActionTimeout.Bind(this),
            BANNER_ACCENT_INTERMEDIATE, 520, 17, "", true,
            pk,
            true,
            true
        )
    }

    OnActionY(*) {
        if (this.CurrentPhase != "PromptingAction")
            return
        this.ExecuteAction(false, false)
    }

    OnActionC(*) {
        if (this.CurrentPhase != "PromptingAction")
            return
        this.PromptForCursorTransfer()
    }

    OnActionR(*) {
        if (this.CurrentPhase != "PromptingAction")
            return
        this.ExecuteAction(true, false)
    }

    ; F: copy last Gemini reply, then mark newest Clip Angel clip as favorite (same as Gemini.ahk CopyAndFavorite).
    OnActionF(*) {
        if (this.CurrentPhase != "PromptingAction")
            return
        this.CleanupActionPrompt()
        try {
            this.DoCopyCore(false, false)
            clipRaw := A_Clipboard
            clip := Trim(clipRaw)
            if (clip = "" || StrLen(clip) < 10) {
                ShowCenteredOverlay_Utils("❌ Copy failed or empty - try again", 2000, BANNER_ACCENT_ERROR)
                if (this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd))
                    WinActivate("ahk_id " this.OriginHwnd)
                return
            }
            MarkLastClipAsFavorite("first", true)
        } finally {
            this.Reset()
        }
    }

    OnActionN(*) {
        if (this.CurrentPhase != "PromptingAction")
            return
        this.CleanupActionPrompt()
        this.Reset()
    }

    OnActionTimeout(*) {
        if (this.CurrentPhase != "PromptingAction")
            return
        this.CleanupActionPrompt()
        this.Reset()
    }

    CleanupActionPrompt() {
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
    }

    ExecuteAction(readAloud := false, skipRestoreFocus := false) {
        this.CleanupActionPrompt()
        try {
            this.DoCopyCore(readAloud, skipRestoreFocus)
        } finally {
            this.Reset()
        }
    }

    DoCopyCore(readAloud := false, skipRestoreFocus := false) {
        if (this.HasCopiedForThisResponse) {
            return
        }
        this.HasCopiedForThisResponse := true

        aiLabel := GetGlobalAIProviderLabel()
        useCopilot := UseCopilotWebForGlobalAI()
        PlayPreMovementWarning(aiLabel)

        if (!WinExist("ahk_id " this.GeminiHwnd)) {
            if (this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd))
                WinActivate("ahk_id " this.OriginHwnd)
            return
        }

        ; By the time DoCopyCore runs, Gemini should already be active. If it is not,
        ; just activate it without a pre-movement warning (source is no longer Original).
        if (!WinActive("ahk_id " this.GeminiHwnd)) {
            try WinActivate("ahk_id " this.GeminiHwnd)
            catch {
                if (this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd))
                    WinActivate("ahk_id " this.OriginHwnd)
                return
            }
            if (!WinWaitActive("ahk_exe chrome.exe", , 0.5)) {
                if (this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd))
                    WinActivate("ahk_id " this.OriginHwnd)
                return
            }
        }

        ; Y / R / C / timeout: same synchronous copy first. R then blocks on read-aloud IPC (wParam=1 skips duplicate Copy in Gemini).
        clipBefore := A_Clipboard
        seqBefore := Clipboard_GetSequenceNumber()
        WM_COPY_LAST_GEMINI := 0x8001
        WM_COPY_LAST_COPILOT := 0x8005
        WM_TRIGGER_READ_ALOUD := 0x8004
        WM_TRIGGER_COPILOT_READ_ALOUD := 0x8006
        wmCopy := useCopilot ? WM_COPY_LAST_COPILOT : WM_COPY_LAST_GEMINI
        wmRead := useCopilot ? WM_TRIGGER_COPILOT_READ_ALOUD : WM_TRIGGER_READ_ALOUD
        targetHwnd := GetGeminiScriptMsgTargetHwnd()
        sendOk := false
        clipOk := false

        if (targetHwnd) {
            prevDH := A_DetectHiddenWindows
            DetectHiddenWindows true
            try {
                SendMessage(wmCopy, 0, this.GeminiHwnd, , "ahk_id " targetHwnd, , , , 20000)
                sendOk := true
            } catch {
            } finally {
                DetectHiddenWindows prevDH
            }
            changed := Clipboard_WaitForSequenceChange(seqBefore, 2000, 850)
            clipOk := sendOk && changed && A_Clipboard != clipBefore && Trim(A_Clipboard) != ""
            if (!sendOk)
                ShowCenteredOverlay_Utils("❌ " . aiLabel . " copy timed out or IPC failed", 3500, BANNER_ACCENT_ERROR)
            else if (!clipOk)
                ShowCenteredOverlay_Utils("❌ Copy failed or clipboard empty - try again", 3000, BANNER_ACCENT_ERROR)
            else if (readAloud) {
                DetectHiddenWindows true
                try {
                    SendMessage(wmRead, 1, this.OriginHwnd, , "ahk_id " targetHwnd, , , , 120000)
                } catch {
                    ShowCenteredOverlay_Utils("❌ Read aloud failed or timed out", 4000, BANNER_ACCENT_ERROR)
                } finally {
                    DetectHiddenWindows prevDH
                }
            } else
                try ScriptSoundPlay(A_ScriptDir . "\sounds\copy.wav")
        } else {
            ShowCenteredOverlay_Utils("❌ Gemini.ahk not running", 2000, BANNER_ACCENT_ERROR)
        }

        ; Gemini/Clipboard → Original: return transitions are immediate (no warning).
        if (!skipRestoreFocus && this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd) && !WinActive("ahk_id " this.OriginHwnd
        )) {
            WinActivate("ahk_id " this.OriginHwnd)
            ; Fast-path: avoid WinWaitActive if we are already active.
            if (!WinActive("ahk_id " this.OriginHwnd))
                WinWaitActive("ahk_id " this.OriginHwnd, , 0.5)
        }
    }

    ; --- Phase 5: Cursor Transfer ---

    PromptForCursorTransfer() {
        this.CurrentPhase := "Transferring"
        this.CleanupActionPrompt()
        try {
            ; Skip restoring focus so clipboard is not overwritten
            this.DoCopyCore(false, true)

            clipRaw := A_Clipboard
            clip := Trim(clipRaw)
            if (clip = "" || StrLen(clip) < 10) {
                ShowCenteredOverlay_Utils("❌ Copy failed or empty - try again", 2000, BANNER_ACCENT_ERROR)
                if (this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd))
                    WinActivate("ahk_id " this.OriginHwnd)
                return
            }

            ; Restore the pre-handoff anchored window so the user sees the selector/paste
            ; happening in the exact app they were monitoring before the Gemini handoff.
            if (this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd)) {
                try {
                    WinActivate("ahk_id " this.OriginHwnd)
                    ; Fast-path: avoid WinWaitActive if we are already active.
                    if (!WinActive("ahk_id " this.OriginHwnd))
                        WinWaitActive("ahk_id " this.OriginHwnd, , 0.5)
                } catch {
                }
            }
            try A_Clipboard := clipRaw

            tSelectorStart := A_TickCount
            this.CursorHwnd := CursorTransfer_ShowWindowSelector(0)
            tSelectorMs := A_TickCount - tSelectorStart
            if (!this.CursorHwnd) {
                if (this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd))
                    WinActivate("ahk_id " this.OriginHwnd)
                try A_Clipboard := clipRaw
                return
            }

            ; Gemini → Cursor: no pre-movement warning (source is not Original).
            try A_Clipboard := clipRaw
            CursorTransfer_ActivateFocusPaste(this.CursorHwnd, this.OriginHwnd)
        } finally {
            this.Reset()
        }
    }

    ; --- Helpers ---

    CancelFlow(message) {
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        ShowCenteredOverlay_Utils("⚠ " . message, 1500, BANNER_ACCENT_INTERMEDIATE)
        global g_D2C_DictationSubmitMenuCycleFinished
        g_D2C_DictationSubmitMenuCycleFinished := true
        this.Reset()
    }
}

; Dictation: "Send to Gemini?" confirmation banner (5s, Y to confirm; uses standard loading indicator)
; =============================================================================
; DEPRECATED: Use D2C_FlowManager
DEPRECATED_DictationGeminiConfirm_Show() {
    ; No-op; use DictationGeminiConfirm_ShowAndWait() which uses StandardLoadingBar_ShowWithKeys.
}

DEPRECATED_DictationGeminiConfirm_Hide(*) {
    StandardLoadingBar_CloseKeysOverlay()
}

; submitToGemini=false (N or timeout): terminal. submitToGemini=true: delayed-submit (paste+Enter). pasteOnly=true: paste to Gemini only, no Enter.
DEPRECATED_DictationGeminiConfirm_CleanupAndMaybeSubmit(submitToGemini, pasteOnly := false) {
    global g_DictationGeminiConfirmBannerVisible
    ; Only clear banner-visible when proceeding (Y/S/timeout); leave true on N so a stray ShowAndWait does not re-show and register a second 5s timer (logs showed second timeout firing after N).
    if (submitToGemini || pasteOnly)
        g_DictationGeminiConfirmBannerVisible := false
    ; Unregister 5s banner keys (same * prefix as StandardLoadingBar_RegisterKeyHandler uses)
    try Hotkey("*y", "Off")
    try Hotkey("*Y", "Off")
    try Hotkey("*s", "Off")
    try Hotkey("*S", "Off")
    try Hotkey("*n", "Off")
    try Hotkey("*N", "Off")
    SetTimer(DEPRECATED_DictationGeminiConfirm_OnTimeout, 0)
    DEPRECATED_DictationGeminiConfirm_Hide()
    ; S or N at 5s: no submit flow, so stop any running "Copy response?" monitor so it never shows.
    if (!submitToGemini)
        GeminiDelayedSubmitMonitorStopFromUtils()
    if (pasteOnly) {
        Sleep 350
        DEPRECATED_GeminiDictationPasteOnlyFlow()
    } else if (submitToGemini) {
        Sleep 350
        DEPRECATED_GeminiDelayedSubmitFlow()
    }
}

DEPRECATED_DictationGeminiConfirm_OnY(*) {
    DEPRECATED_DictationGeminiConfirm_CleanupAndMaybeSubmit(true)
}

; S = paste to Gemini only (no Enter, no 4s banner).
DEPRECATED_DictationGeminiConfirm_OnS(*) {
    DEPRECATED_DictationGeminiConfirm_CleanupAndMaybeSubmit(false, true)
}

; Default action on 5s timeout: proceed as Yes (DelayedSubmitFlow), same as user pressing Y.
DEPRECATED_DictationGeminiConfirm_OnTimeout(*) {
    DEPRECATED_DictationGeminiConfirm_CleanupAndMaybeSubmit(true)
}

; N = terminate flow: no paste, no Enter, no 4s, no copy; only cleanup and cancel overlay.
DEPRECATED_DictationGeminiConfirm_OnCancel(*) {
    DEPRECATED_DictationGeminiConfirm_CleanupAndMaybeSubmit(false)  ; submitToGemini=false, pasteOnly=false => no flow runs
    ShowCenteredOverlay_Utils("⚠ Gemini submission cancelled", 1500, BANNER_ACCENT_INTERMEDIATE)
}

DEPRECATED_DictationGeminiConfirm_ShowAndWait() {
    global g_DictationGeminiConfirmBannerVisible
    ; Only one banner: atomic check-and-set so only one invocation can pass (prevents duplicate from multiple PlayDictationCompletionChime runs).
    Critical "On"
    if (g_DictationGeminiConfirmBannerVisible) {
        Critical "Off"
        return
    }
    g_DictationGeminiConfirmBannerVisible := true
    Critical "Off"
    ; Only the official loading bar (standard loading indicator) may show this content. Hide any other bar/overlay first.
    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Hide(0)
    HideDictationIndicator()
    Sleep 50
    keyCallbacks := Map("Y", DEPRECATED_DictationGeminiConfirm_OnY, "S", DEPRECATED_DictationGeminiConfirm_OnS, "N",
        DEPRECATED_DictationGeminiConfirm_OnCancel)
    ; Official loading bar only; no blue; single banner (no border); fixed bottom strip for input.
    StandardLoadingBar_ShowWithKeys("❓ Send to " . GetGlobalAIProviderLabel() . "? (5s)", keyCallbacks,
        D2C_SUBMIT_MENU_TIMEOUT_MS,
        0,
        DEPRECATED_DictationGeminiConfirm_OnTimeout, BANNER_ACCENT_INTERMEDIATE, 380, 17, "", true,
        "[Y] Send  [S] Paste only  [N] Cancel",
        true,
        true)
}

; =============================================================================
; Global Sound Toggle System
; File-backed state management for muting/unmuting sounds across all scripts
; =============================================================================

; Check if sound is enabled (reads from INI file for cross-process persistence)
IsSoundEnabled() {
    settingsFile := A_ScriptDir . "\data\settings.ini"
    ; Default to enabled (1) if file doesn't exist or key is missing
    soundEnabled := IniRead(settingsFile, "Settings", "SoundEnabled", "1")
    return (soundEnabled = "1")
}

; -----------------------------------------------------------------------------
; Central gate for the global sound toggle: all script-triggered audio must use
; these helpers (SoundPlay, WMP, SoundBeep, MessageBeep, system scheme sounds).
; When IsSoundEnabled() is false, these are no-ops.
; -----------------------------------------------------------------------------
ScriptSoundPlay(path, wait := false) {
    if (!IsSoundEnabled())
        return false
    try {
        SoundPlay(path, wait)
        return true
    } catch {
        return false
    }
}

; System scheme sounds, e.g. *64 (asterisk), *16 (exclamation) - see SoundPlay docs.
ScriptSoundPlaySystem(scheme) {
    if (!IsSoundEnabled())
        return false
    try {
        SoundPlay(scheme, false)
        return true
    } catch {
        return false
    }
}

; Study subtopic / article link API save success (Manage Study Link flows).
StudyLink_PlayApiSuccessSound() {
    try ScriptSoundPlay(A_ScriptDir . "\sounds\api-success.mp3")
}

ScriptSoundBeep(freq, duration) {
    if (!IsSoundEnabled())
        return false
    try {
        SoundBeep(freq, duration)
        return true
    } catch {
        return false
    }
}

ScriptMessageBeep(type := 0xFFFFFFFF) {
    if (!IsSoundEnabled())
        return false
    try {
        return DllCall("User32\MessageBeep", "UInt", type)
    } catch {
        return false
    }
}

; Toggle sound state and show visual feedback
ToggleSoundState() {
    settingsFile := A_ScriptDir . "\data\settings.ini"
    currentState := IsSoundEnabled()
    newState := currentState ? "0" : "1"

    ; Update INI file
    IniWrite(newState, settingsFile, "Settings", "SoundEnabled")

    ; Show visual feedback
    if (newState = "1") {
        ShowCenteredOverlay_Utils("🔊 Sound: ON", 2000, BANNER_ACCENT_INTERMEDIATE)
    } else {
        ShowCenteredOverlay_Utils("🔇 Sound: OFF", 2000, BANNER_ACCENT_INTERMEDIATE)
    }
}

; =============================================================================
; Centralized audio levels (AHK playback app volume vs mic capture - not Windows master)
; =============================================================================
global SCRIPT_MASTER_VOLUME_PERCENT := 40
global SCRIPT_MIC_CAPTURE_VOLUME_PERCENT := 100
global SCRIPT_MICROPHONE_INPUT_SLIDER_PERCENT := 100

; Per-process AutoHotkey playback volume via WASAPI (see ApplyAutoHotkeyAudioSessionsVolumePercent).
; Does not call SoundSetVolume - leaves the default device master volume unchanged.
ApplyScriptMasterVolumeTarget() {
    return 0 ; Only SetAutoHotkeyVolume.ps1 should set volume now.
}

; Apply now and again after delays - audio sessions for new AutoHotkey processes often do not exist for hundreds of ms after Start-Process (Quick Update / multi-script startup).
; Use distinct timer callbacks (lambdas): SetTimer with the *same* function reference replaces the previous timer - two SetTimer(ApplyScriptMasterVolumeTarget, ...) would only keep the last delay.
ScheduleApplyScriptMasterVolumeTargetWithRetries() {
    ApplyScriptMasterVolumeTarget()
    SetTimer(() => ApplyScriptMasterVolumeTarget(), -2500)
    SetTimer(() => ApplyScriptMasterVolumeTarget(), -6000)
}

; Run after Quick Update relaunch only: AppLaunchers starts last with /Updated; other scripts need time to spawn audio sessions.
ScheduleApplyScriptMasterVolumeTargetAfterQuickUpdate() {
    ApplyScriptMasterVolumeTarget()
    SetTimer(() => ApplyScriptMasterVolumeTarget(), -2000)
    SetTimer(() => ApplyScriptMasterVolumeTarget(), -5000)
    SetTimer(() => ApplyScriptMasterVolumeTarget(), -10000)
    SetTimer(() => ApplyScriptMasterVolumeTarget(), -15000)
    SetTimer(() => ApplyScriptMasterVolumeTarget(), -20000)
    SetTimer(() => ApplyScriptMasterVolumeTarget(), -25000)
}

RunSetMicVolumeScript() {
    micVolumeScript := A_ScriptDir "\scripts\Set-MicVolume.ps1"
    if (!FileExist(micVolumeScript))
        return
    try {
        Run("powershell.exe -ExecutionPolicy Bypass -File `"" micVolumeScript "`" -Level " SCRIPT_MIC_CAPTURE_VOLUME_PERCENT, ,
            "Hide")
    } catch {
    }
}

; =============================================================================
; Outlook: classic OUTLOOK.EXE and Microsoft Store "new" Outlook (olk.exe)
; =============================================================================
OutlookGetOlkExePath() {
    candidate :=
        "C:\Program Files\WindowsApps\Microsoft.OutlookForWindows_1.2026.317.100_x64__8wekyb3d8bbwe\olk.exe"
    if FileExist(candidate)
        return candidate
    try {
        loop files "C:\Program Files\WindowsApps\Microsoft.OutlookForWindows_*_x64__8wekyb3d8bbwe\olk.exe", "F" {
            return A_LoopFileFullPath
        }
    } catch {
    }
    return ""
}

OutlookProcessRunning() {
    return ProcessExist("OUTLOOK.EXE") || ProcessExist("olk.exe")
}

; =============================================================================
; Toggle Outlook and Teams
; Toggles Outlook and Teams applications to manage RAM usage.
; If both are open: Closes Outlook and minimizes Teams to system tray.
; If one or both are closed: Launches both applications.
; =============================================================================
ToggleOutlookAndTeams() {
    loadingShown := false
    try {
        ; Check if both applications are running
        outlookRunning := OutlookProcessRunning()
        teamsRunning := ProcessExist("ms-teams.exe")
        isOpeningFlow := !(outlookRunning && teamsRunning)
        hadError := false
        firstError := ""

        ; Show start banner
        if (!isOpeningFlow) {
            ShowCenteredOverlay_Utils("📤 Closing Outlook and Teams...", 1500, BANNER_ACCENT_INTERMEDIATE)
        } else {
            StandardLoadingBar_Show("⏳ Opening Outlook and Teams...", BANNER_ACCENT_INTERMEDIATE, {
                passive: false,
                centerOnHwnd: 0,
                textWidth: 560,
                fontSize: 17,
                passiveBgColor: BANNER_ACCENT_INTERMEDIATE
            })
            loadingShown := true
        }

        if (outlookRunning && teamsRunning) {
            ; Both are open: Close Outlook and minimize Teams to system tray
            ; Close Outlook process(es) - classic and/or Store (olk.exe)
            try {
                if ProcessExist("OUTLOOK.EXE")
                    ProcessClose("OUTLOOK.EXE")
                if ProcessExist("olk.exe")
                    ProcessClose("olk.exe")
            } catch Error as e {
                MsgBox "Error closing Outlook: " e.Message
            }

            ; Close all Teams windows (this keeps Teams in system tray)
            try {
                ; Teams can have multiple process names, check all
                for hwnd in WinGetList("ahk_exe ms-teams.exe") {
                    WinClose(hwnd)
                }
                ; Also check for Teams.exe and MSTeams.exe variants
                for hwnd in WinGetList("ahk_exe Teams.exe") {
                    WinClose(hwnd)
                }
                for hwnd in WinGetList("ahk_exe MSTeams.exe") {
                    WinClose(hwnd)
                }
            } catch Error as e {
                MsgBox "Error closing Teams windows: " e.Message
            }
        } else {
            ; One or both are closed: Launch both applications
            ; Launch Outlook
            if (!outlookRunning) {
                StandardLoadingBar_Update("⏳ Opening Outlook...")
                try {
                    outlookPath := ""
                    if (IS_WORK_ENVIRONMENT) {
                        ; Try work environment shortcut path
                        outlookPath := "C:\Users\fie7ca\Documents\Atalhos\Microsoft Outlook.lnk"
                        if (!FileExist(outlookPath)) {
                            outlookPath := ""
                        }
                    } else {
                        ; Try personal environment shortcut path
                        outlookPath :=
                            "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Microsoft Outlook.lnk"
                        if (!FileExist(outlookPath)) {
                            outlookPath := ""
                        }
                    }

                    ; Launch using shortcut if available, otherwise olk.exe or OUTLOOK.EXE
                    if (outlookPath != "") {
                        Run outlookPath
                    } else {
                        olkPath := OutlookGetOlkExePath()
                        if (olkPath != "")
                            Run olkPath
                        else
                            Run "OUTLOOK.EXE"
                    }
                } catch Error as e {
                    hadError := true
                    if (firstError = "")
                        firstError := "Outlook: " . e.Message
                }
            }

            ; Launch/Activate Teams
            ; Simplified approach: Just run the executable. This handles both launching and bringing to front.
            StandardLoadingBar_Update("⏳ Opening Teams...")
            try {
                if (IS_WORK_ENVIRONMENT) {
                    teamsExePath :=
                        "C:\Program Files\WindowsApps\MSTeams_25332.1210.4188.1171_x64__8wekyb3d8bbwe\ms-teams.exe"
                    if (FileExist(teamsExePath)) {
                        Run teamsExePath
                    } else {
                        Run "ms-teams.exe"
                    }
                } else {
                    ; Personal environment
                    teamsPath :=
                        "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Microsoft Teams.lnk"
                    if (FileExist(teamsPath)) {
                        Run teamsPath
                    } else {
                        Run "ms-teams.exe"
                    }
                }

                ; Wait for window to appear and become active
                if (WinWaitActive("ahk_exe ms-teams.exe", , 10)) {
                } else {
                    hadError := true
                    if (firstError = "")
                        firstError := "Teams window not found"
                }
            } catch Error as e {
                hadError := true
                if (firstError = "")
                    firstError := "Teams: " . e.Message
            }

            ; Second: Activate Outlook last (so it gets final focus)
            StandardLoadingBar_Update("⏳ Activating Outlook...")
            try {
                if (OutlookProcessRunning()) {
                    ex := ProcessExist("OUTLOOK.EXE") ? "OUTLOOK.EXE" : "olk.exe"
                    WinWait("ahk_exe " ex, , 5)
                    if (!WinExist("ahk_exe " ex)) {
                        hadError := true
                        if (firstError = "")
                            firstError := "Outlook not running"
                    } else {
                        WinActivate("ahk_exe " ex)
                        WinWaitActive("ahk_exe " ex, , 2)
                    }
                } else {
                    hadError := true
                    if (firstError = "")
                        firstError := "Outlook process not detected"
                }
            } catch Error as e {
                hadError := true
                if (firstError = "")
                    firstError := "Outlook activation: " . e.Message
            }

            if (loadingShown) {
                StandardLoadingBar_Hide(0)
                loadingShown := false
            }

            if (hadError) {
                ShowCenteredOverlay_Utils("❌ Open completed with issues: " . firstError, 2500, BANNER_ACCENT_ERROR)
            } else {
                ShowCenteredOverlay_Utils("✅ Outlook and Teams opened", 1500, BANNER_ACCENT_SUCCESS)
            }

            return
        }

        ; Show finish banner
        ShowCenteredOverlay_Utils("✅ Done", 1500, BANNER_ACCENT_SUCCESS)
    } catch Error as e {
        if (loadingShown)
            StandardLoadingBar_Hide(0)
        MsgBox "Error in ToggleOutlookAndTeams macro: " e.Message
    }
}

; =============================================================================
; Check and Prompt to Open Outlook/Teams
; Checks if Outlook or Teams are closed and prompts user to open them if needed
; Parameters:
;   - checkOutlook: true to check Outlook, false otherwise
;   - checkTeams: true to check Teams, false otherwise
; Returns: true if applications are running (or were opened), false if user cancelled
; =============================================================================
CheckAndOpenOutlookTeams(checkOutlook := false, checkTeams := false) {
    outlookClosed := false
    teamsClosed := false

    ; Check Outlook status (classic OUTLOOK.EXE or Store new Outlook olk.exe)
    if (checkOutlook) {
        outlookRunning := OutlookProcessRunning()
        if (!outlookRunning) {
            outlookClosed := true
        }
    }

    ; Check Teams status
    if (checkTeams) {
        teamsRunning := ProcessExist("ms-teams.exe")
        if (!teamsRunning) {
            teamsClosed := true
        }
    }

    ; If both are open, no action needed
    if (!outlookClosed && !teamsClosed) {
        return true
    }

    ; Build message based on what's closed
    message := ""
    if (outlookClosed && teamsClosed) {
        message := "Outlook and Teams are closed. Do you want to open them?"
    } else if (outlookClosed) {
        message := "Outlook is closed. Do you want to open it?"
    } else if (teamsClosed) {
        message := "Teams is closed. Do you want to open it?"
    }

    ; Show message box
    response := MsgBox(message, "Open Applications?", "YesNo Icon?")

    ; If user confirms, open the applications (only open, don't toggle)
    if (response = "Yes") {
        ; Only open the closed applications, don't toggle
        try {
            ; Launch Outlook if closed
            if (outlookClosed) {
                outlookPath := ""
                if (IS_WORK_ENVIRONMENT) {
                    outlookPath := "C:\Users\fie7ca\Documents\Atalhos\Microsoft Outlook.lnk"
                    if (!FileExist(outlookPath)) {
                        outlookPath := ""
                    }
                } else {
                    outlookPath :=
                        "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Microsoft Outlook.lnk"
                    if (!FileExist(outlookPath)) {
                        outlookPath := ""
                    }
                }

                if (outlookPath != "") {
                    Run outlookPath
                } else {
                    olkPath := OutlookGetOlkExePath()
                    if (olkPath != "")
                        Run olkPath
                    else
                        Run "OUTLOOK.EXE"
                }
            }

            ; Launch Teams if closed
            if (teamsClosed) {
                if (IS_WORK_ENVIRONMENT) {
                    teamsExePath :=
                        "C:\Program Files\WindowsApps\MSTeams_25332.1210.4188.1171_x64__8wekyb3d8bbwe\ms-teams.exe"
                    if (FileExist(teamsExePath)) {
                        Run teamsExePath
                    } else {
                        Run "ms-teams.exe"
                    }
                } else {
                    teamsPath :=
                        "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Microsoft Teams.lnk"
                    if (FileExist(teamsPath)) {
                        Run teamsPath
                    } else {
                        Run "ms-teams.exe"
                    }
                }
            }

            ; Wait a bit for applications to start
            Sleep 2000
            return true
        } catch Error as e {
            MsgBox "Error opening applications: " e.Message, "Error", "IconX"
            return false
        }
    }

    ; User cancelled
    return false
}

; Chime for "clean now" confirmations (desktop recycle Y, clean clipboard Y). Not used on auto-timeout.
; Quiet: WASAPI attenuation + synchronous SoundPlay - no WMPlayer.OCX (its volume pins same-PID mixer ~10% despite later WASAPI). try/finally restores SCRIPT_MASTER_VOLUME_PERCENT deterministically.
; One-shot timer re-applies target: a new session can appear right after SoundPlay returns; first enumeration may miss it (mixer stuck ~10%).
PlayCleaningDesktopSound() {
    if (!IsSoundEnabled())
        return
    soundPath := A_ScriptDir "\sounds\cleaning-desktop.wav"
    if (!FileExist(soundPath))
        return
    ScriptSoundPlay(soundPath, false)
}

; Clean Clipboard macro: session/cancel guards (prevents double-run and ignored N during automation)
global g_CleanClipboardCanceled := false
global g_CleanClipboardInProgress := false
global g_CleanClipboardSessionId := 0
global g_CleanClipboardProceedClaimed := false

CleanClipboard_BeginSession() {
    global g_CleanClipboardCanceled, g_CleanClipboardSessionId, g_CleanClipboardProceedClaimed
    g_CleanClipboardSessionId += 1
    g_CleanClipboardCanceled := false
    g_CleanClipboardProceedClaimed := false
    return g_CleanClipboardSessionId
}

CleanClipboard_EndSession() {
    global g_CleanClipboardInProgress
    g_CleanClipboardInProgress := false
}

CleanClipboard_ShouldAbort(sessionId := 0) {
    global g_CleanClipboardCanceled, g_CleanClipboardSessionId
    if (g_CleanClipboardCanceled)
        return true
    if (sessionId && sessionId != g_CleanClipboardSessionId)
        return true
    return false
}

CleanClipboard_UnwindClipAngel() {
    try {
        if WinExist("ClipAngel")
            SendInput "!v"
    } catch {
    }
    Sleep 100
}

; N/Esc while automation runs (overlay already closed; StandardLoadingBar keys are inactive)
CleanClipboard_SetAbortHotkeys(enable := true) {
    if (enable) {
        Hotkey("*n", CleanClipboard_OnCancel, "On")
        Hotkey("*Escape", CleanClipboard_OnCancel, "On")
    } else {
        try Hotkey("*n", "Off")
        catch {
        }
        try Hotkey("*Escape", "Off")
        catch {
        }
    }
}

; Internal helper: Performs clipboard cleanup without showing prompt
; sessionId: macro session from CleanClipboard_ShowCountdown; 0 = dictation legacy path (no session guard)
CleanClipboardInternal(sessionId := 0) {
    Sleep 200
    if (CleanClipboard_ShouldAbort(sessionId)) {
        CleanClipboard_UnwindClipAngel()
        return
    }

    SendInput "!v"
    Sleep 600
    if (CleanClipboard_ShouldAbort(sessionId)) {
        CleanClipboard_UnwindClipAngel()
        return
    }

    SendInput "^!k"
    Sleep 600
    if (CleanClipboard_ShouldAbort(sessionId)) {
        CleanClipboard_UnwindClipAngel()
        return
    }

    SendInput "{Enter}"
    Sleep 800
    if (CleanClipboard_ShouldAbort(sessionId)) {
        CleanClipboard_UnwindClipAngel()
        return
    }

    SendInput "!v"
    Sleep 400
}

; =============================================================================
; Dictation: Non-modal clipboard cleanup countdown (used by Win+Alt+Shift+7)
; =============================================================================
global g_DictationCleanupGui := 0
global g_DictationCleanupBorderGui := 0
global g_DictationCleanupTextCtrl := 0
global g_DictationCleanupRemaining := 0
global g_DictationCleanupCanceled := false

DictationCleanup_SetCancelHotkeys(enable := true) {
    ; Removed ~ prefix to prevent key leakage into active applications
    if (enable) {
        Hotkey("*n", DictationCleanup_Cancel, "On")
        Hotkey("*y", DictationCleanup_Proceed, "On")
        Hotkey("*End", DictationCleanup_Cancel, "On")
    } else {
        try Hotkey("*n", "Off")
        catch {
        }
        try Hotkey("*y", "Off")
        catch {
        }
        try Hotkey("*End", "Off")
        catch {
        }
    }
}

DictationCleanup_ShowBanner() {
    global g_DictationCleanupGui, g_DictationCleanupTextCtrl, g_DictationCleanupRemaining

    ; Destroy any previous banner instance
    try {
        if IsObject(g_DictationCleanupGui)
            g_DictationCleanupGui.Destroy()
    } catch {
    }
    g_DictationCleanupGui := 0
    g_DictationCleanupTextCtrl := 0

    target := WinGetID("A")
    hasWindow := false
    if target && WinExist("ahk_id " target) {
        try {
            WinGetPos(&wx, &wy, &ww, &wh, target)
            hasWindow := (ww > 0 && wh > 0)
        } catch {
            hasWindow := false
        }
    }

    ov := Gui("+AlwaysOnTop -Caption +ToolWindow")
    ov.BackColor := "1E1E2E"
    ov.SetFont("s24 cFFFFFF Bold", "Segoe UI")
    g_DictationCleanupTextCtrl := ov.Add("Text", "w650 Center", "Clearing clipboard in " g_DictationCleanupRemaining "... (press Y to proceed, N or End to cancel)"
    )
    ov.Show("AutoSize Hide")
    ov.GetPos(&gx, &gy, &gw, &gh)

    if hasWindow {
        cx := wx + (ww - gw) // 2
        cy := wy + (wh - gh) // 2
    } else {
        vx := SysGet(76)  ; SM_XVIRTUALSCREEN
        vy := SysGet(77)  ; SM_YVIRTUALSCREEN
        vw := SysGet(78)  ; SM_CXVIRTUALSCREEN
        vh := SysGet(79)  ; SM_CYVIRTUALSCREEN
        cx := vx + (vw - gw) // 2
        cy := vy + (vh - gh) // 2
    }

    borderWidth := 6
    try {
        if IsObject(g_DictationCleanupBorderGui)
            g_DictationCleanupBorderGui.Destroy()
    } catch {
    }
    borderGui := Gui("+AlwaysOnTop -Caption +ToolWindow")
    borderGui.BackColor := BANNER_ACCENT_INTERMEDIATE
    borderGui.Show("NA x" . (cx - borderWidth) . " y" . (cy - borderWidth) . " w" . (gw + 2 * borderWidth) . " h" . (gh +
        2 * borderWidth))
    g_DictationCleanupBorderGui := borderGui

    ov.Show("x" . cx . " y" . cy . " NA")
    WinSetTransparent(178, ov)
    g_DictationCleanupGui := ov
}

DictationCleanup_HideBanner() {
    global g_DictationCleanupGui, g_DictationCleanupBorderGui, g_DictationCleanupTextCtrl
    try {
        if IsObject(g_DictationCleanupBorderGui)
            g_DictationCleanupBorderGui.Destroy()
    } catch {
    }
    g_DictationCleanupBorderGui := 0
    try {
        if IsObject(g_DictationCleanupGui)
            g_DictationCleanupGui.Destroy()
    } catch {
    }
    g_DictationCleanupGui := 0
    g_DictationCleanupTextCtrl := 0
}

DictationCleanup_UpdateBannerText() {
    global g_DictationCleanupTextCtrl, g_DictationCleanupRemaining
    try {
        if IsObject(g_DictationCleanupTextCtrl) {
            g_DictationCleanupTextCtrl.Text := "Clearing clipboard in " g_DictationCleanupRemaining "... (press Y to proceed, N or End to cancel)"
        }
    } catch {
    }
}

DictationCleanup_StartCountdown(seconds := 5) {
    global g_DictationCleanupRemaining, g_DictationCleanupCanceled

    ; Reset state
    g_DictationCleanupCanceled := false
    g_DictationCleanupRemaining := seconds

    ; Show initial banner + enable cancel keys
    DictationCleanup_ShowBanner()
    DictationCleanup_SetCancelHotkeys(true)

    ; Start ticking immediately (1s cadence)
    SetTimer(DictationCleanup_Tick, 0)
    SetTimer(DictationCleanup_Tick, 1000)
}

DictationCleanup_StopCountdown(showCancelledBanner := false) {
    global g_DictationCleanupCanceled

    ; Stop timer + disable cancel keys
    SetTimer(DictationCleanup_Tick, 0)
    DictationCleanup_SetCancelHotkeys(false)

    if (showCancelledBanner) {
        ; Reuse the same banner GUI for a short "cancelled" message (non-blocking)
        global g_DictationCleanupTextCtrl
        try {
            if IsObject(g_DictationCleanupTextCtrl) {
                g_DictationCleanupTextCtrl.Text := "Clipboard cleanup cancelled"
            }
        } catch {
        }
        SetTimer(DictationCleanup_HideBanner, -900)
    } else {
        DictationCleanup_HideBanner()
    }
}

DictationCleanup_Cancel(*) {
    global g_DictationCleanupCanceled
    g_DictationCleanupCanceled := true
    DictationCleanup_StopCountdown(true)
}

DictationCleanup_Proceed(*) {
    ; Immediately proceed with clipboard cleanup, skipping countdown
    DictationCleanup_StopCountdown(false)
    CleanClipboardInternal()
}

DictationCleanup_Tick() {
    global g_DictationCleanupRemaining, g_DictationCleanupCanceled

    ; If already cancelled, ensure everything is stopped
    if (g_DictationCleanupCanceled) {
        DictationCleanup_StopCountdown(true)
        return
    }

    ; Decrement remaining time
    g_DictationCleanupRemaining -= 1

    if (g_DictationCleanupRemaining <= 0) {
        ; Countdown finished -> hide banner and clear clipboard using the existing workflow (Alt+V, Ctrl+Alt+K, etc.)
        DictationCleanup_StopCountdown(false)
        CleanClipboardInternal()
        return
    }

    DictationCleanup_UpdateBannerText()
}

; =============================================================================
; Dictation: Merge non-favorite clips countdown (at end of loop)
; Same UI pattern as clipboard cleanup: 5s banner, N or End to cancel.
; =============================================================================
global g_DictationMergeGui := 0
global g_DictationMergeBorderGui := 0
global g_DictationMergeTextCtrl := 0
global g_DictationMergeRemaining := 0
global g_DictationMergeCanceled := false

DictationMerge_SetCancelHotkeys(enable := true) {
    ; Removed ~ prefix to prevent key leakage into active applications
    if (enable) {
        Hotkey("*n", DictationMerge_Cancel, "On")
        Hotkey("*y", DictationMerge_Proceed, "On")
        Hotkey("*End", DictationMerge_Cancel, "On")
    } else {
        try Hotkey("*n", "Off")
        catch {
        }
        try Hotkey("*y", "Off")
        catch {
        }
        try Hotkey("*End", "Off")
        catch {
        }
    }
}

DictationMerge_ShowBanner() {
    global g_DictationMergeGui, g_DictationMergeTextCtrl, g_DictationMergeRemaining

    try {
        if IsObject(g_DictationMergeGui)
            g_DictationMergeGui.Destroy()
    } catch {
    }
    g_DictationMergeGui := 0
    g_DictationMergeTextCtrl := 0

    target := WinGetID("A")
    hasWindow := false
    if target && WinExist("ahk_id " target) {
        try {
            WinGetPos(&wx, &wy, &ww, &wh, target)
            hasWindow := (ww > 0 && wh > 0)
        } catch {
            hasWindow := false
        }
    }

    ov := Gui("+AlwaysOnTop -Caption +ToolWindow")
    ov.BackColor := "1E1E2E"
    ov.SetFont("s24 cFFFFFF Bold", "Segoe UI")
    g_DictationMergeTextCtrl := ov.Add("Text", "w650 Center", "Merging non-favorite clips in " g_DictationMergeRemaining "... (press Y to proceed, N or End to cancel)"
    )
    ov.Show("AutoSize Hide")
    ov.GetPos(&gx, &gy, &gw, &gh)
    if (hasWindow) {
        cx := wx + (ww - gw) // 2
        cy := wy + (wh - gh) // 2
    } else {
        vx := SysGet(76)
        vy := SysGet(77)
        vw := SysGet(78)
        vh := SysGet(79)
        cx := vx + (vw - gw) // 2
        cy := vy + (vh - gh) // 2
    }
    borderWidth := 6
    try {
        if IsObject(g_DictationMergeBorderGui)
            g_DictationMergeBorderGui.Destroy()
    } catch {
    }
    borderGui := Gui("+AlwaysOnTop -Caption +ToolWindow")
    borderGui.BackColor := BANNER_ACCENT_INTERMEDIATE
    borderGui.Show("NA x" . (cx - borderWidth) . " y" . (cy - borderWidth) . " w" . (gw + 2 * borderWidth) . " h" . (gh +
        2 * borderWidth))
    g_DictationMergeBorderGui := borderGui
    ov.Show("x" . cx . " y" . cy . " NA")
    g_DictationMergeGui := ov
}

DictationMerge_HideBanner() {
    global g_DictationMergeGui, g_DictationMergeBorderGui, g_DictationMergeTextCtrl
    try {
        if IsObject(g_DictationMergeBorderGui)
            g_DictationMergeBorderGui.Destroy()
    } catch {
    }
    g_DictationMergeBorderGui := 0
    try {
        if IsObject(g_DictationMergeGui)
            g_DictationMergeGui.Destroy()
    } catch {
    }
    g_DictationMergeGui := 0
    g_DictationMergeTextCtrl := 0
}

DictationMerge_UpdateBannerText() {
    global g_DictationMergeTextCtrl, g_DictationMergeRemaining
    try {
        if IsObject(g_DictationMergeTextCtrl) {
            g_DictationMergeTextCtrl.Text := "Merging non-favorite clips in " g_DictationMergeRemaining "... (press Y to proceed, N or End to cancel)"
        }
    } catch {
    }
}

DictationMerge_StartCountdown(seconds := 5) {
    global g_DictationMergeRemaining, g_DictationMergeCanceled

    g_DictationMergeCanceled := false
    g_DictationMergeRemaining := seconds

    DictationMerge_ShowBanner()
    DictationMerge_SetCancelHotkeys(true)

    SetTimer(DictationMerge_Tick, 0)
    SetTimer(DictationMerge_Tick, 1000)
}

DictationMerge_StopCountdown(showCancelledBanner := false) {
    SetTimer(DictationMerge_Tick, 0)
    DictationMerge_SetCancelHotkeys(false)

    if (showCancelledBanner) {
        global g_DictationMergeTextCtrl
        try {
            if IsObject(g_DictationMergeTextCtrl) {
                g_DictationMergeTextCtrl.Text := "Merge cancelled"
            }
        } catch {
        }
        SetTimer(DictationMerge_HideBanner, -900)
    } else {
        DictationMerge_HideBanner()
    }
}

DictationMerge_Cancel(*) {
    global g_DictationMergeCanceled
    g_DictationMergeCanceled := true
    DictationMerge_StopCountdown(true)
}

DictationMerge_Proceed(*) {
    ; Immediately proceed with merge, skipping countdown
    DictationMerge_StopCountdown(false)
    MergeNonFavoriteClips()
}

DictationMerge_Tick() {
    global g_DictationMergeRemaining, g_DictationMergeCanceled

    if (g_DictationMergeCanceled) {
        DictationMerge_StopCountdown(true)
        return
    }

    g_DictationMergeRemaining -= 1

    if (g_DictationMergeRemaining <= 0) {
        DictationMerge_StopCountdown(false)
        MergeNonFavoriteClips()
        return
    }

    DictationMerge_UpdateBannerText()
}

; Clean the Clipboard macro function
; Shows a non-modal 4s countdown; auto-continues unless user cancels with N; Y proceeds immediately.
CleanClipboard() {
    CleanClipboard_ShowCountdown()
}

; Strip markdown escape backslashes chat UIs add on copy (e.g. meeting\_id → meeting_id).
UnescapeMarkdownFromText(text) {
    if (text = "")
        return { text: "", count: 0 }
    if RegExMatch(text, '^\s*```[^\R]*\R(.*?)\R```\s*$', &fence)
        text := fence[1]
    count := 0
    for ch in ["*", "_", "{", "}", "[", "]", "(", ")", "#", "+", ".", "!", "-", Chr(96), Chr(34)] {
        replaced := 0
        text := StrReplace(text, '\' . ch, ch, , &replaced)
        count += replaced
    }
    replaced := 0
    text := StrReplace(text, '\\', '\', , &replaced)
    count += replaced
    return { text: text, count: count }
}

UnescapeMarkdownClipboard() {
    clip := A_Clipboard
    if (Trim(clip) = "") {
        ShowCenteredOverlay_Utils("Clipboard empty", 1500, BANNER_ACCENT_INTERMEDIATE)
        return
    }
    result := UnescapeMarkdownFromText(clip)
    if (result.text = clip) {
        ShowCenteredOverlay_Utils("No markdown escapes found", 1500, BANNER_ACCENT_INTERMEDIATE)
        return
    }
    saved := ClipboardAll()
    try {
        A_Clipboard := result.text
        ClipWait(0.3)
    } finally {
        if (A_Clipboard != result.text)
            A_Clipboard := saved
    }
    msg := (result.count = 1) ? "Removed 1 markdown escape" : "Removed " result.count " markdown escapes"
    ShowCenteredOverlay_Utils("📋 " msg, 1500, BANNER_ACCENT_SUCCESS)
    try ScriptSoundPlay(A_ScriptDir . "\sounds\copy.wav")
}

CleanClipboard_ShowCountdown() {
    global g_CleanClipboardInProgress

    if (g_CleanClipboardInProgress) {
        ShowCenteredOverlay_Utils("Clipboard cleanup already running", 1500, BANNER_ACCENT_INTERMEDIATE)
        return
    }

    CleanClipboard_BeginSession()

    ; Ensure any previous keys overlay is closed before showing a new one
    try StandardLoadingBar_CloseKeysOverlay()
    catch {
    }
    StandardLoadingBar_Hide(0)
    Sleep 50

    state := "❓ Clean the clipboard? (removes stored clips, 4s)`nPress [Y] to clean now, or [N] within 4s to cancel."
    keyCallbacks := Map("N", CleanClipboard_OnCancel, "Y", CleanClipboard_OnYConfirm, "*Escape",
        CleanClipboard_OnCancel)

    ; Center on active monitor (centerOnHwnd := 0), use red accent for destructive action.
    StandardLoadingBar_ShowWithKeys(
        state,
        keyCallbacks,
        4000,
        0,
        CleanClipboard_OnTimeout.Bind(g_CleanClipboardSessionId),
        BANNER_ACCENT_ERROR,
        0,
        17,
        "",
        false,
        "[Y] Clean now  [N] Cancel (auto-continue in 4s)",
        true,
        true,
        true)
}

CleanClipboard_OnCancel(*) {
    global g_CleanClipboardCanceled
    g_CleanClipboardCanceled := true
    try StandardLoadingBar_CloseKeysOverlay()
    catch {
    }
    StandardLoadingBar_Hide(0)
    ShowCenteredOverlay_Utils("Clipboard cleanup cancelled", 1500, BANNER_ACCENT_INTERMEDIATE)
}

CleanClipboard_Proceed(sessionId := 0) {
    global g_CleanClipboardCanceled, g_CleanClipboardInProgress, g_CleanClipboardSessionId

    if (g_CleanClipboardCanceled)
        return
    if (g_CleanClipboardInProgress)
        return
    if (sessionId && sessionId != g_CleanClipboardSessionId)
        return

    g_CleanClipboardInProgress := true
    try {
        if (g_CleanClipboardCanceled)
            return
        try StandardLoadingBar_CloseKeysOverlay()
        catch {
        }
        StandardLoadingBar_Hide(0)
        CleanClipboard_SetAbortHotkeys(true)
        CleanClipboardInternal(sessionId)
    } finally {
        CleanClipboard_SetAbortHotkeys(false)
        CleanClipboard_EndSession()
    }
}

CleanClipboard_OnYConfirm(*) {
    global g_CleanClipboardCanceled, g_CleanClipboardSessionId, g_CleanClipboardProceedClaimed
    if (g_CleanClipboardCanceled)
        return
    ; Hotkey + poll can both fire Y before Proceed sets inProgress
    if (g_CleanClipboardProceedClaimed)
        return
    g_CleanClipboardProceedClaimed := true
    PlayCleaningDesktopSound()
    CleanClipboard_Proceed(g_CleanClipboardSessionId)
}

; Auto-continue when countdown ends (no chime; only Y plays the sound)
CleanClipboard_OnTimeout(sessionId, *) {
    global g_CleanClipboardCanceled, g_CleanClipboardProceedClaimed
    if (g_CleanClipboardCanceled)
        return
    if (g_CleanClipboardProceedClaimed)
        return
    g_CleanClipboardProceedClaimed := true
    CleanClipboard_Proceed(sessionId)
}

; =============================================================================
; Project Data (for Cursor Window Focus Selector)
; Ported from WindowManagement.ahk to ensure consistent key mapping
; =============================================================================

; Character sequence for assignment: 1 2 3 4 5 q w e r t a s d f g z x c v b 6 7 8 9 0 y u i o p h j k l n m , .
global g_ProjectCharSequence := ["1", "2", "3", "4", "5", "q", "w", "e", "r", "t", "a", "s", "d", "f", "g", "z", "x",
    "c", "v", "b", "6", "7", "8", "9", "0", "y", "u", "i", "o", "p", "h", "j", "k", "l", "n", "m", ",", "."]

; Category display order (General first, Personal second, Work last)
global g_ProjectCategories := ["General", "Personal", "Work"]

; Global project list - must match WindowManagement.ahk for consistent key mapping
; Each project should have: name, path, workPath, and category ("General", "Personal", or "Work")
global g_Projects := [
    ; General category
    { name: "Scripts", path: "C:\Users\eduev\Meu Drive\17 - Projects\scripts", workPath: "C:\Users\fie7ca\Documents\scripts",
        category: "General" }, { name: "14-my-notes", path: "C:\Users\eduev\Meu Drive\17 - Projects\notes", workPath: "C:\Users\fie7ca\OneDrive - Bosch Group\14-my-notes",
            category: "General" }, { name: "", path: "", workPath: "", category: "General" }, { name: "", path: "",
                workPath: "", category: "General" }, { name: "", path: "", workPath: "", category: "General" },
                ; Personal category
                { name: "ZMK Sofle", path: "C:\Users\eduev\Documents\ZMK\zmk-sofle", workPath: "", category: "Personal" }, { name: "AI Experiment",
                    path: "C:\Users\eduev\Meu Drive\04 - Pós-graduação\01 - Mestrado\26-ai-experiment", workPath: "",
                    category: "Personal" }, { name: "my-personal-repo", path: "C:\Users\eduev\Meu Drive\17 - Projects\my-personal-repo",
                        workPath: "C:\Users\fie7ca\OneDrive - Bosch Group\13 - General workspace\my-personal-repo",
                        category: "Personal" }, { name: "",
                            path: "", workPath: "", category: "Personal" }, { name: "", path: "", workPath: "",
                                category: "Personal" },
                            ; Work category
                            { name: "dashboard-model-research", path: "", workPath: "C:\Users\fie7ca\OneDrive - Bosch Group\13 - General workspace\GS_E&S_CIP Dashboard research and design workspace folder\dashboard-model-research",
                                category: "Work" }, { name: "GS_UX core team_UX and CIP Integration", path: "",
                                    workPath: "C:\Users\fie7ca\OneDrive - Bosch Group\13 - General workspace\GS_UX core team_UX and CIP Integration",
                                    category: "Work" }, { name: "🪂 Avante", path: "", workPath: "C:\Users\fie7ca\OneDrive - Bosch Group\General - GS_BDU_Team\00_UX_GS_Team\AM_Planning\Avante",
                                        category: "Work" }, { name: "Piloto PT B2B", path: "", workPath: "C:\Users\fie7ca\OneDrive - Bosch Group\SO UX - LA (Internal) - Data Insights SO - Piloto PT B2B",
                                            category: "Work" }, { name: "Python ScripTs", path: "", workPath: "C:\Users\fie7ca\OneDrive - Bosch Group\17 - Python Scripts",
                                                category: "Work", char: "t" }, { name: "",
                                                    path: "", workPath: "", category: "Work" }
]

; Extract matching segments from project path for window title matching
; Cursor window titles have format: "filename - folder-name - Cursor" or "filename - path-segment - Cursor"
ExtractProjectMatchSegments(projectPath) {
    ; Normalize the project path (remove trailing backslashes)
    normalizedPath := RTrim(projectPath, "\")

    ; Split path into segments
    pathSegments := StrSplit(normalizedPath, "\")

    ; Extract the last folder name (e.g., "zmk-sofle", "26-ai-experiment", "scripts")
    lastSegment := pathSegments[pathSegments.Length]

    ; Build list of potential match strings
    matchSegments := [lastSegment]

    ; If we have at least 2 segments, also try the combination
    if (pathSegments.Length >= 2) {
        ; Try last two segments joined with " - " (for cases like "17 - Projects")
        lastTwoJoined := pathSegments[pathSegments.Length - 1] . " - " . pathSegments[pathSegments.Length]
        if (lastTwoJoined != lastSegment) {  ; Only add if different
            matchSegments.Push(lastTwoJoined)
        }
    }

    return matchSegments
}

; =============================================================================
; Global AI generation state: Cursor + Gemini stop-button detectors (Efficiency Canon)
; =============================================================================
; Cursor: Type 50026 (Group), ClassName contains "stop-button".
; Gemini: Chrome window title contains "gemini"; Type 50000, Name "Stop response", ClassName match.
; =============================================================================
Cursor_HasGeneratingStopButton() {
    global UIA
    try {
        cursorHwnds := WinGetList("ahk_exe Cursor.exe")
        if (!cursorHwnds.Length)
            return false
        cr := UIA.CreateCacheRequest(["Type", "ClassName"], , 5)
        for hwnd in cursorHwnds {
            if (!hwnd || !WinExist("ahk_id " hwnd))
                continue
            try {
                root := UIA.ElementFromHandleBuildCache(cr, hwnd)
            } catch {
                try root := UIA.ElementFromHandle(hwnd)
                catch
                    continue
            }
            if (!root)
                continue
            try {
                el := root.FindFirstBuildCache(cr, { Type: 50026, ClassName: "stop-button", matchmode: "Substring" }, 4
                )
                if (el)
                    return true
            } catch {
            }
        }
    } catch {
    }
    return false
}

Gemini_HasGeneratingStopButton() {
    global UIA
    try {
        for hwnd in WinGetList("ahk_exe chrome.exe") {
            if (!hwnd || !WinExist("ahk_id " hwnd))
                continue
            try {
                if (InStr(WinGetTitle("ahk_id " hwnd), "gemini", false) = 0)
                    continue
            } catch {
                continue
            }
            try {
                cr := UIA.CreateCacheRequest(["Type", "ClassName", "Name"], , 5)
                root := UIA.ElementFromHandleBuildCache(cr, hwnd)
            } catch {
                try root := UIA.ElementFromHandle(hwnd)
                catch
                    continue
            }
            if (!root)
                continue
            ; Stop response: Type 50000, Name "Stop response", ClassName contains "send-button" and "stop"
            try {
                el := root.FindFirstBuildCache(cr, { Type: 50000, Name: "Stop response", ClassName: "send-button",
                    matchmode: "Substring" }, 4)
                if (el)
                    return true
            } catch {
            }
        }
    } catch {
    }
    return false
}

IsAnyAiGenerating() {
    return Cursor_HasGeneratingStopButton() || Gemini_HasGeneratingStopButton()
}

PlayAiWorkingStateSound(isWorking) {
    try {
        if (isWorking)
            ScriptSoundPlay(A_ScriptDir . "\sounds\robots-are-working.wav")
        else
            ScriptSoundPlay(A_ScriptDir . "\sounds\no-robot-working.wav")
    } catch {
    }
}

; =============================================================================
; U macro: Global AI generation state (Cursor + Gemini) with sound and banner
; =============================================================================
; Runs Cursor + Gemini stop-button checks, plays robots-are-working / no-robot-working,
; shows red banner when any AI is working, green when none.
; =============================================================================
Cursor_FindComposerIconAcrossInstances() {
    global BANNER_ACCENT_SUCCESS, BANNER_ACCENT_ERROR
    try {
        isWorking := IsAnyAiGenerating()
        PlayAiWorkingStateSound(isWorking)
        if (isWorking)
            ShowCenteredOverlay_Utils("AI is working (stop button found)", 2000, BANNER_ACCENT_ERROR)
        else
            ShowCenteredOverlay_Utils("No AI is working", 2000, BANNER_ACCENT_SUCCESS)
    } catch Error as e {
        PlayAiWorkingStateSound(false)
        ShowCenteredOverlay_Utils("No AI is working", 2000, BANNER_ACCENT_SUCCESS)
    }
}

; Initialize macros
InitMacros() {
    ; Quick Update to Your Scripts macro
    RegisterMacro(QuickUpdateScripts, "⚡ Quick Update to Your Scripts")
    ; Add specific word to Handy macroh
    RegisterMacro(AddWordToHandy, "➕ Add specific word to Handy")
    ; Toggle Outlook and Teams macro
    RegisterMacro(ToggleOutlookAndTeams, "🔄 Toggle Outlook & Teams")
    ; Clean the Clipboard macro (assigned to "P")
    RegisterMacro(CleanClipboard, "🧹 Clean the Clipboard", "p")
    RegisterMacro(UnescapeMarkdownClipboard, "📋 Unescape markdown clipboard", "e")
    ; Toggle Sound macro
    RegisterMacro(ToggleSoundState, "🔊 Toggle Sound (Mute/Unmute)")
    ; Global AI generation state: Cursor + Gemini (assigned to "U")
    RegisterMacro(Cursor_FindComposerIconAcrossInstances, "🔍 AI working? (Cursor + Gemini)", "u")
    ; Mark Last Clip as Favorite macro (assigned to "J")
    RegisterMacro(MarkLastClipAsFavorite, "⭐ Mark Last Clip as Favorite", "j")
    ; Move Desktop to Recycle Bin (assigned to "N") - red banner, Y/N confirm
    RegisterMacro(DesktopToRecycle_Trigger, "🗑️ Move Desktop to Recycle Bin", "n")
}

InitMacros()

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

; Auto-execute: show success after QuickUpdateScripts relaunches AppLaunchers.ahk (includes this file) with "/Updated".
if (A_Args.Length > 0 && A_Args[1] = "/Updated") {
    try {
        ShowCenteredOverlay_Utils("✅ Scripts updated and relaunched", 6500, BANNER_ACCENT_SUCCESS)
        soundPath := A_ScriptDir "\sounds\quick-update-success.wav"
        ; Play success chime to completion before scheduling volume: async SoundPlay can register a new session after
        ; the first Apply pass, leaving that session at a low default (~10% in the mixer) until something re-enumerates.
        try {
            if (FileExist(soundPath))
                ScriptSoundPlay(soundPath, true)
        } catch {
        }
        ; After all scripts have been started (this block runs last in include order for AppLaunchers /Updated), apply AHK volume - not at Quick Update start (old sessions / dead timers).
        ScheduleApplyScriptMasterVolumeTargetAfterQuickUpdate()
    } catch {
    }
}

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
; Select AI Model in Handy (Win+Alt+Shift+C)
; =============================================================================
#!+C::
{
    SelectAiModelInHandy()
}

; =============================================================================
; Move all Desktop items to Recycle Bin (recoverable)
; Trigger: Win+Alt+Shift+U selector → letter N
; Target path: OneDrive Desktop. Standard banner with 4s timeout (N = cancel, Y or timeout = run); then success/error banner.
; =============================================================================
global g_DesktopToRecyclePath := ""  ; Set from GetDesktopToRecyclePath() when macro runs
global g_DesktopToRecycleCloseHwnd := 0

DesktopToRecycle_OnConfirm(*) {
    PlayCleaningDesktopSound()
    DesktopToRecycle_Run()
}

DesktopToRecycle_OnCancel(*) {
    ShowCenteredOverlay_Utils("⚠ Desktop cleanup cancelled", 1500, BANNER_ACCENT_INTERMEDIATE)
}

DesktopToRecycle_OnTimeout(*) {
    DesktopToRecycle_Run()
}

; Normalize folder path for comparison (trim trailing backslash, lowercase on Windows)
DesktopToRecycle_NormalizePath(p) {
    p := RTrim(p, "\")
    try return StrLower(p)
    return p
}

; Close any Explorer window(s) showing the given folder path (via Shell.Application)
DesktopToRecycle_CloseDesktopExplorer(targetPath) {
    if (!targetPath || targetPath = "")
        return
    normTarget := DesktopToRecycle_NormalizePath(targetPath)
    try {
        shell := ComObject("Shell.Application")
        for window in shell.Windows {
            try {
                if (!window || !window.hwnd)
                    continue
                path := window.Document.Folder.Self.Path
                if (DesktopToRecycle_NormalizePath(path) = normTarget) {
                    window.Quit()
                    return
                }
            } catch
                continue
        }
    } catch {
    }
    ; Fallback: close by hwnd if we had stored it at trigger time
    global g_DesktopToRecycleCloseHwnd
    if (g_DesktopToRecycleCloseHwnd && WinExist("ahk_id " g_DesktopToRecycleCloseHwnd)) {
        try WinClose("ahk_id " g_DesktopToRecycleCloseHwnd)
    }
    g_DesktopToRecycleCloseHwnd := 0
}

DesktopToRecycle_Run() {
    global g_DesktopToRecyclePath, g_DesktopToRecycleCloseHwnd
    ; Resolve path: use configured path; if empty or missing, fall back to A_Desktop (works on any PC)
    path := g_DesktopToRecyclePath
    if (!path || path = "" || !DirExist(path))
        path := A_Desktop
    ; Use .NET FileIO.FileSystem SendToRecycleBin (no Shell verbs); process dirs last so parent exists
    ui := "[Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs"
    rec := "[Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin"
    ps := "Add-Type -AssemblyName Microsoft.VisualBasic;$d='" . path .
        "';if(-not(Test-Path -LiteralPath $d)){exit 1};$files=@(Get-ChildItem -LiteralPath $d -Force|Where-Object{-not $_.PSIsContainer});$dirs=@(Get-ChildItem -LiteralPath $d -Force|Where-Object{$_.PSIsContainer});foreach($f in $files){try{[Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($f.FullName," .
        ui . "," . rec .
        ")}catch{}};foreach($dir in $dirs){try{[Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($dir.FullName," .
        ui . "," . rec . ")}catch{}};exit 0"
    try {
        exitCode := RunWait('powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "' . ps . '"', "", "Hide")
        if (exitCode = 0) {
            ShowCenteredOverlay_Utils("✅ Desktop items moved to Recycle Bin", 2000, BANNER_ACCENT_SUCCESS)
            DesktopToRecycle_CloseDesktopExplorer(path)
        } else {
            ShowCenteredOverlay_Utils("❌ Desktop path not found or error: " path, 3500, BANNER_ACCENT_ERROR)
            DesktopToRecycle_CloseDesktopExplorer(path)
        }
    } catch as err {
        ShowCenteredOverlay_Utils("❌ Error moving to Recycle Bin", 2500, BANNER_ACCENT_ERROR)
    }
    g_DesktopToRecycleCloseHwnd := 0
}

; Entry point when "N" is pressed in Win+Alt+Shift+U selector
DesktopToRecycle_Trigger() {
    global g_DesktopToRecycleCloseHwnd, g_DesktopToRecyclePath
    g_DesktopToRecyclePath := GetDesktopToRecyclePath()
    ; Remember active window if it's Explorer showing Desktop - close it after cleaning
    hwnd := WinExist("A")
    g_DesktopToRecycleCloseHwnd := 0
    if (hwnd && WinGetProcessName("ahk_id " hwnd) = "explorer.exe") {
        try {
            if (InStr(WinGetTitle("ahk_id " hwnd), "Desktop", false))
                g_DesktopToRecycleCloseHwnd := hwnd
        } catch {
        }
    }
    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Hide(0)
    Sleep 50
    state := "🗑️ Move all items from:`n" . g_DesktopToRecyclePath . "`nto Recycle Bin? (4s)"
    keyCallbacks := Map("Y", DesktopToRecycle_OnConfirm, "N", DesktopToRecycle_OnCancel)
    ; Center on active monitor (centerOnHwnd := 0), use standard intermediate accent with border.
    StandardLoadingBar_ShowWithKeys(
        state,
        keyCallbacks,
        4000,
        0,
        DesktopToRecycle_OnTimeout,
        BANNER_ACCENT_INTERMEDIATE,
        0,
        17,
        "",
        false,
        "[Y] Yes  [N] Cancel",
        true,
        true,
        true)
}

; =============================================================================
; Clip Angel: Open/Activate with focus correction (Row 0)
; Hotkey: Alt+V - when closed: open + focus Row 0; when open: pass Alt+V to close (toggle).
; =============================================================================
!v::
{
    if WinExist("ClipAngel") {
        Send "!v"   ; Already open: close it (Clip Angel toggle)
        return
    }
    ActivateClipAngelWithFocusCorrection()
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

; Win+Alt+Shift+Arrow: send that arrow key five times
#!+Right::
{
    Send("{Right 5}")
    return
}

#!+Left::
{
    Send("{Left 5}")
    return
}

#!+Down::
{
    Send("{Down 5}")
    return
}

#!+Up::
{
    Send("{Up 5}")
    return
}

; =============================================================================
; Peek PDF - Win+Alt+Shift+X
; If Peek is open: activate it. Otherwise: show study-topic selector (same aesthetic as Win+Alt+Shift+C).
; =============================================================================

; Study topics for Win+Alt+Shift+X selector. Paths are relative to notes repo (GetNotesRepoPath()).
; plansPath values match filenames in the notes repo (see studies/*/ *-plan.md, plan-english.md, learning-techniques.md).
global g_StudyTopics := Map(
    0, { name: "Technique (how to create studies)", mnemonicsPath: "\studies\technique\README.md",
        plansPath: "\studies\technique\plans.md" },
    1, { name: "Mnemonics", mnemonicsPath: "\studies\skills\mnemonics-skills.md",
        plansPath: "\studies\skills\learning-techniques.md" },
    2, { name: "Science", mnemonicsPath: "\studies\science\mnemonics-science.md",
        plansPath: "\studies\science\science-plan.md" },
    3, { name: "Piano", mnemonicsPath: "\studies\piano\mnemonics-piano.md",
        plansPath: "\studies\piano\piano-plan.md" },
    4, { name: "English", mnemonicsPath: "\studies\english\mnemonics-english.md",
        plansPath: "\studies\english\plan-english.md" },
    5, { name: "Communication", mnemonicsPath: "\studies\communication\mnemonics-communication.md",
        plansPath: "\studies\communication\communication-plan.md" },
    6, { name: "German", mnemonicsPath: "\studies\german\mnemonics-german.md",
        plansPath: "\studies\german\german-plan.md" }
)
#include %A_ScriptDir%\StudyArticleLink.ahk
#include %A_ScriptDir%\StudyFavoriteLink.ahk

global g_StudyTopicSelectorGui := false
global g_StudyTopicSelectorActive := false
global g_StudyTopicSelectorPhase := ""           ; "category" | "topic"
global g_StudyTopicSelectorCategory := ""        ; "mnemonics" | "plans"
global g_StudyTopicSelectorLastForegroundMonitorIdx := 0   ; for trackActiveMonitor-style follow (standard_information_display.md)
global g_StudyTopicEscPollPrev := false   ; edge-detect Esc for StudyTopicSelector_EscapePoll (parity with OutlookCopilotSelector)
global STUDY_TOPIC_BLACKOUT_DELAY_MS := 3000
; Study Topic QuickLook: strict bounded waits + shared layout (false = legacy 2s WinWait + inline scroll).
global STUDY_TOPIC_QL_STRICT_LAYOUT := true
global g_QuickLookDeferredLayoutScroll := true
global g_QuickLookDeferredLayoutPath := ""

; PDF focus monitoring for automatic blackout cancellation (Win+Alt+Shift+X)
global g_PdfFocusMonitorTimer := false
global g_PdfFocusTrackedHwnd := 0
global g_PdfFocusLossMode := "Immediate"      ; "Immediate" or "Debounced"
global g_PdfFocusDebounceMs := 1200            ; Allow transient focus loss without un-blackouting
global g_PdfFocusLostSinceTick := 0

; Monitor foreground vs keep-clear display: relocate only when anchor HWND moves; else disable.
MonitorPdfFocus() {
    global g_FocusModeOn, g_FocusModeActiveMonitor, g_PdfFocusTrackedHwnd, g_FocusModeTrackedWindow
    global g_FocusModeAnchorHwnd, g_PdfFocusLossMode, g_PdfFocusDebounceMs, g_PdfFocusLostSinceTick

    if (!g_FocusModeOn) {
        if (g_PdfFocusTrackedHwnd)
            StopPdfFocusMonitor()
        return
    }

    if (FocusMode_CheckCrossProcessRequests())
        return

    fg := WinExist("A")
    if (!fg)
        return

    fgMon := GetActiveMonitorIndex()
    if (!fgMon)
        return

    keepMon := g_FocusModeActiveMonitor
    if (keepMon && fgMon != keepMon) {
        anchor := g_FocusModeAnchorHwnd
        sameAnchor := anchor && fg = anchor && WinExist("ahk_id " . anchor)
        if (sameAnchor) {
            if (g_PdfFocusLossMode = "Debounced") {
                if (!g_PdfFocusLostSinceTick)
                    g_PdfFocusLostSinceTick := A_TickCount
                if ((A_TickCount - g_PdfFocusLostSinceTick) >= g_PdfFocusDebounceMs) {
                    FocusMode_SetKeepMonitor(fgMon)
                    g_PdfFocusLostSinceTick := 0
                }
            } else {
                FocusMode_SetKeepMonitor(fgMon)
                g_PdfFocusLostSinceTick := 0
            }
        } else {
            DisableFocusMode()
            return
        }
    } else {
        g_PdfFocusLostSinceTick := 0
    }

    g_FocusModeTrackedWindow := fg
    g_PdfFocusTrackedHwnd := fg
}

; Start monitoring PDF window focus
StartPdfFocusMonitor(hwnd := 0, focusLossMode := "Immediate") {
    global g_PdfFocusMonitorTimer, g_PdfFocusTrackedHwnd, g_PdfFocusLossMode, g_PdfFocusLostSinceTick
    global g_FocusModeAnchorHwnd

    StopPdfFocusMonitor()

    g_PdfFocusTrackedHwnd := hwnd ? hwnd : WinExist("A")
    if (!g_PdfFocusTrackedHwnd)
        return

    if (hwnd)
        g_FocusModeAnchorHwnd := hwnd
    else if (!g_FocusModeAnchorHwnd || !WinExist("ahk_id " . g_FocusModeAnchorHwnd))
        g_FocusModeAnchorHwnd := g_PdfFocusTrackedHwnd

    g_PdfFocusLossMode := focusLossMode
    g_PdfFocusLostSinceTick := 0

    g_PdfFocusMonitorTimer := MonitorPdfFocus
    SetTimer(g_PdfFocusMonitorTimer, 200)
}

; Stop monitoring PDF window focus
StopPdfFocusMonitor() {
    global g_PdfFocusMonitorTimer, g_PdfFocusTrackedHwnd

    ; First-call safety: global may be unset
    if (!IsSet(g_PdfFocusMonitorTimer))
        g_PdfFocusMonitorTimer := false

    if (g_PdfFocusMonitorTimer) {
        SetTimer(g_PdfFocusMonitorTimer, 0)
        g_PdfFocusMonitorTimer := false
    }
    g_PdfFocusTrackedHwnd := 0
}

StudyTopic_GetBlackoutKeepMonitorIndex() {
    try return MonitorGetPrimary()
    catch
        return 1
}

; Shared countdown flag for FocusBlackoutWatcher (Study Topic path clears without setting).
BlackoutCountdown_Begin() {
    global g_FocusBlackoutWatcherCountdownActive
    g_FocusBlackoutWatcherCountdownActive := true
}

BlackoutCountdown_End() {
    global g_FocusBlackoutWatcherCountdownActive
    g_FocusBlackoutWatcherCountdownActive := false
}

StudyTopic_CancelBlackoutCountdown(targetHwnd := 0, *) {
    global g_FocusBlackoutWatcherDeniedHwnd
    BlackoutCountdown_End()
    if (targetHwnd && WinExist("ahk_id " . targetHwnd))
        g_FocusBlackoutWatcherDeniedHwnd := targetHwnd
    else {
        fg := WinExist("A")
        if (fg)
            g_FocusBlackoutWatcherDeniedHwnd := fg
    }
    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Hide(0)
}

; Apply blackout using foreground at timeout (user may juggle monitors during the 3s banner).
StudyTopic_ApplyBlackoutCountdownTimeout(targetHwnd := 0, pdfFocusLossMode := "Debounced") {
    global g_BlackoutSuppressedUntil, g_FocusModeOn

    BlackoutCountdown_End()

    fg := WinExist("A")
    if (!fg && targetHwnd && WinExist("ahk_id " . targetHwnd))
        fg := targetHwnd
    if (!fg)
        return

    if (Blackout_IsSuppressed())
        return

    keepIdx := GetActiveMonitorIndex()
    if (!keepIdx)
        keepIdx := StudyTopic_GetBlackoutKeepMonitorIndex()
    if (g_FocusModeOn)
        FocusMode_SetKeepMonitor(keepIdx)
    else
        EnableFocusMode(keepIdx)
    StartPdfFocusMonitor(fg, pdfFocusLossMode)
}

; --- Blackout suppression logic ---
global g_BlackoutSuppressedUntil := 0
global BLACKOUT_SUPPRESS_MS := 7 * 60 * 1000

Blackout_IsSuppressed() {
    global g_BlackoutSuppressedUntil
    if (!g_BlackoutSuppressedUntil)
        return false
    return (g_BlackoutSuppressedUntil - A_TickCount) > 0
}

; D key handler: suppress all blackout banners for BLACKOUT_SUPPRESS_MS; reset dwell for post-suppress window.
Blackout_Disable7Min(*) {
    global g_BlackoutSuppressedUntil, g_FocusBlackoutWatcherDwellStartTick
    g_BlackoutSuppressedUntil := A_TickCount + BLACKOUT_SUPPRESS_MS
    g_FocusBlackoutWatcherDwellStartTick := A_TickCount
    BlackoutCountdown_End()
    FocusBlackoutWatcher_DebugLog("Blackout_Disable7Min until=" . g_BlackoutSuppressedUntil . " tick=" . A_TickCount .
        " remaining=" . (g_BlackoutSuppressedUntil - A_TickCount))
    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Hide(0)
}

StudyTopic_StartBlackoutCountdown(targetHwnd) {
    if (!targetHwnd || !WinExist("ahk_id " . targetHwnd))
        return
    if (Blackout_IsSuppressed()) {
        BlackoutCountdown_End()
        return
    }
    BlackoutCountdown_Begin()
    global g_FocusBlackoutWatcherDwellStartTick
    g_FocusBlackoutWatcherDwellStartTick := A_TickCount
    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Hide(0)
    Sleep 50
    cancelCb := StudyTopic_CancelBlackoutCountdown.Bind(targetHwnd)
    keyCallbacks := Map("N", cancelCb, "*Escape", cancelCb)
    keyCallbacks["D"] := Blackout_Disable7Min
    timeoutCb := StudyTopic_ApplyBlackoutCountdownTimeout.Bind(, "Immediate")
    StandardLoadingBar_ShowWithKeys(
        "⏳ Blacking out secondary monitors in 3s",
        keyCallbacks,
        STUDY_TOPIC_BLACKOUT_DELAY_MS,
        0,
        timeoutCb,
        BANNER_BLACKOUT_BORDER,
        420,
        17,
        BANNER_BLACKOUT_BORDER,
        false,
        "[N] Cancel blackout    [D] Disable for 7 min",
        true,
        true,
        true,
        BANNER_BLACKOUT_PANEL)
}

; Focus dwell watcher: after continuous foreground on one window, offer same blackout banner as Study Topic (#!+X).
global g_FocusBlackoutWatcherStarted := false
global g_FocusBlackoutWatcherLastHwnd := 0
global g_FocusBlackoutWatcherDwellStartTick := 0
global g_FocusBlackoutWatcherDeniedHwnd := 0
global g_FocusBlackoutWatcherCountdownActive := false
global FOCUS_BLACKOUT_DWELL_MS := 20000
global FOCUS_BLACKOUT_DEBUG_LOG := false

FocusBlackoutWatcher_DebugLog(message) {
    global FOCUS_BLACKOUT_DEBUG_LOG
    if (!FOCUS_BLACKOUT_DEBUG_LOG)
        return
    try {
        FileAppend "[FBW] " . message . "`n", A_ScriptDir "\focus_blackout_debug.log", "UTF-8"
    } catch {
    }
}

FocusBlackoutWatcher_OnCancel(hwnd, *) {
    StudyTopic_CancelBlackoutCountdown(hwnd)
}

FocusBlackoutWatcher_OnBlackoutTimeout(hwnd, *) {
    StudyTopic_ApplyBlackoutCountdownTimeout(, "Immediate")
}

FocusBlackoutWatcher_StartCountdown(hwnd) {
    global g_FocusBlackoutWatcherDwellStartTick
    if (!hwnd || !WinExist("ahk_id " . hwnd))
        return
    if (Blackout_IsSuppressed()) {
        BlackoutCountdown_End()
        FocusBlackoutWatcher_DebugLog("StartCountdown skipped (suppressed)")
        return
    }
    BlackoutCountdown_Begin()
    g_FocusBlackoutWatcherDwellStartTick := A_TickCount
    FocusBlackoutWatcher_DebugLog("StartCountdown for hwnd " . hwnd)
    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Hide(0)
    Sleep 50
    cancelCb := FocusBlackoutWatcher_OnCancel.Bind(hwnd)
    keyCallbacks := Map("N", cancelCb, "*Escape", cancelCb)
    keyCallbacks["D"] := Blackout_Disable7Min
    timeoutCb := FocusBlackoutWatcher_OnBlackoutTimeout.Bind(hwnd)  ; hwnd unused at apply; foreground at timeout wins
    StandardLoadingBar_ShowWithKeys(
        "⏳ Blacking out secondary monitors in 3s",
        keyCallbacks,
        STUDY_TOPIC_BLACKOUT_DELAY_MS,
        0,
        timeoutCb,
        BANNER_BLACKOUT_BORDER,
        420,
        17,
        BANNER_BLACKOUT_BORDER,
        false,
        "[N] Cancel blackout    [D] Disable for 7 min",
        true,
        true,
        true,
        BANNER_BLACKOUT_PANEL)
}

FocusBlackoutWatcher_Tick() {
    global g_FocusBlackoutWatcherLastHwnd, g_FocusBlackoutWatcherDwellStartTick
    global g_FocusBlackoutWatcherDeniedHwnd, g_FocusBlackoutWatcherCountdownActive, FOCUS_BLACKOUT_DWELL_MS
    global g_FocusModeOn, g_FocusModeActiveMonitor, g_FocusModeTrackedWindow

    try {
        if (MonitorGetCount() <= 1)
            return
    } catch {
        return
    }

    try {
        hwnd := WinExist("A")
        if (!hwnd) {
            g_FocusBlackoutWatcherLastHwnd := 0
            FocusBlackoutWatcher_DebugLog("No active window")
            return
        }

        if (g_FocusBlackoutWatcherDeniedHwnd && hwnd != g_FocusBlackoutWatcherDeniedHwnd) {
            FocusBlackoutWatcher_DebugLog("Reset denied hwnd (user switched window)")
            g_FocusBlackoutWatcherDeniedHwnd := 0
        }

        if (hwnd != g_FocusBlackoutWatcherLastHwnd) {
            FocusBlackoutWatcher_DebugLog("Window changed. Reset dwell timer.")
            g_FocusBlackoutWatcherLastHwnd := hwnd
            g_FocusBlackoutWatcherDwellStartTick := A_TickCount
            return
        }

        ; Suppression before countdown-active so dwell resets even if the flag was left set (e.g. Esc).
        if (Blackout_IsSuppressed()) {
            BlackoutCountdown_End()
            FocusBlackoutWatcher_DebugLog("Blackout suppressed. Reset dwell timer.")
            g_FocusBlackoutWatcherDwellStartTick := A_TickCount
            return
        }

        if (g_FocusBlackoutWatcherCountdownActive) {
            FocusBlackoutWatcher_DebugLog("Countdown already active")
            return
        }

        if (g_FocusBlackoutWatcherDeniedHwnd && hwnd = g_FocusBlackoutWatcherDeniedHwnd) {
            FocusBlackoutWatcher_DebugLog("Blackout denied for this hwnd")
            return
        }

        if (g_FocusModeOn && g_FocusModeActiveMonitor) {
            curMon := GetActiveMonitorIndex()
            if (curMon && curMon = g_FocusModeActiveMonitor) {
                FocusBlackoutWatcher_DebugLog("Focus mode already on for this monitor")
                return
            }
        }

        elapsed := (A_TickCount - g_FocusBlackoutWatcherDwellStartTick)
        if (elapsed >= FOCUS_BLACKOUT_DWELL_MS) {
            FocusBlackoutWatcher_DebugLog("Dwell met (" . elapsed . " ms). Starting blackout countdown for hwnd " .
                hwnd)
            FocusBlackoutWatcher_StartCountdown(hwnd)
            return
        }
        FocusBlackoutWatcher_DebugLog("Dwell not yet met: " . elapsed . " ms")
    } catch as e {
        FocusBlackoutWatcher_DebugLog("tick error: " . e.Message)
    }
}

FocusBlackoutWatcher_Start() {
    global g_FocusBlackoutWatcherStarted
    if (g_FocusBlackoutWatcherStarted)
        return
    g_FocusBlackoutWatcherStarted := true
    SetTimer(FocusBlackoutWatcher_Tick, 200)
}

FocusBlackoutWatcher_Stop() {
    global g_FocusBlackoutWatcherStarted
    if (!g_FocusBlackoutWatcherStarted)
        return
    SetTimer(FocusBlackoutWatcher_Tick, 0)
    g_FocusBlackoutWatcherStarted := false
}

; YouTube focus session (Win+Alt+Shift+H): toggle on/off; SMTC for Spotify play/pause (not toggle).
global g_YoutubeFocusMonitorTimer := false
global g_YoutubeFocusTrackedHwnd := 0
global g_YoutubeSpotifyPausePending := false
global g_YoutubeFocusSessionActive := false

; Find Spotify's Windows.Media.Control session (SourceAppUserModelId contains "Spotify").
YouTube_FindSpotifyMediaSession() {
    try {
        for session in Media.GetSessions() {
            try {
                id := session.SourceAppUserModelId
                if InStr(id, "Spotify")
                    return session
            } catch {
                continue
            }
        }
    } catch {
        return 0
    }
    return 0
}

; Pause Spotify via SMTC only when status is Playing (avoids starting playback when already paused).
YouTube_PauseSpotifyBeforeYoutube() {
    global g_YoutubeSpotifyPausePending
    if (g_YoutubeSpotifyPausePending)
        return
    try {
        session := YouTube_FindSpotifyMediaSession()
        if !session
            return
        if (session.PlaybackStatus != Media.PlaybackStatus.Playing)
            return
        session.Pause()
        g_YoutubeSpotifyPausePending := true
    } catch {
        ; WinRT/SMTC unavailable - do not fall back to Media_Play_Pause (toggle bug).
    }
}

; Resume Spotify with SMTC Play() only when we paused it and session reports Paused.
YouTube_ResumeSpotifyAfterYoutubeIfPending(restoreHwnd := 0) {
    global g_YoutubeSpotifyPausePending
    if (!g_YoutubeSpotifyPausePending)
        return
    g_YoutubeSpotifyPausePending := false
    try {
        session := YouTube_FindSpotifyMediaSession()
        if (session && session.PlaybackStatus == Media.PlaybackStatus.Paused)
            session.Play()
    } catch {
        ;
    }
    if (restoreHwnd && WinExist("ahk_id " restoreHwnd)) {
        try
            WinActivate("ahk_id " restoreHwnd)
    }
}

; End YouTube focus session: pause YouTube (k), resume Spotify if pending, stop window monitor.
YouTube_EndFocusSession() {
    global g_YoutubeFocusTrackedHwnd, g_YoutubeFocusSessionActive
    restoreHwnd := WinExist("A")
    if (g_YoutubeFocusTrackedHwnd && WinExist("ahk_id " . g_YoutubeFocusTrackedHwnd)) {
        try {
            WinActivate("ahk_id " . g_YoutubeFocusTrackedHwnd)
            Sleep(50)
            Send("k")
            Sleep(100)
        }
    }
    YouTube_ResumeSpotifyAfterYoutubeIfPending(restoreHwnd)
    StopYoutubeFocusMonitor()
    g_YoutubeFocusSessionActive := false
}

; Monitor tracked YouTube window: only when it is destroyed, run full session teardown (same as second hotkey).
MonitorYoutubeFocus() {
    global g_YoutubeFocusTrackedHwnd

    if (g_YoutubeFocusTrackedHwnd && !WinExist("ahk_id " . g_YoutubeFocusTrackedHwnd))
        YouTube_EndFocusSession()
}

; Start monitoring YouTube window focus
StartYoutubeFocusMonitor(hwnd := 0) {
    global g_YoutubeFocusMonitorTimer, g_YoutubeFocusTrackedHwnd

    StopYoutubeFocusMonitor()

    g_YoutubeFocusTrackedHwnd := hwnd ? hwnd : WinExist("A")
    if (!g_YoutubeFocusTrackedHwnd)
        return

    g_YoutubeFocusMonitorTimer := MonitorYoutubeFocus
    SetTimer(g_YoutubeFocusMonitorTimer, 200)
}

; Stop monitoring YouTube window focus
StopYoutubeFocusMonitor() {
    global g_YoutubeFocusMonitorTimer, g_YoutubeFocusTrackedHwnd

    if (g_YoutubeFocusMonitorTimer) {
        SetTimer(g_YoutubeFocusMonitorTimer, 0)
        g_YoutubeFocusMonitorTimer := false
    }
    g_YoutubeFocusTrackedHwnd := 0
}

StudyTopic_GetRelPath(topic, category) {
    if (category = "plans")
        return topic.plansPath
    return topic.mnemonicsPath
}

; Opens notes-repo-relative path in QuickLook (PDF sibling → .md). Returns false on failure.
; scrollToEnd: mnemonics jump to bottom of long docs; plans stay at top.
StudyTopic_OpenRepoRelativeMarkdown(relPath, scrollToEnd := true) {
    basePath := GetNotesRepoPath()
    if (basePath = "") {
        try ShowCenteredOverlay_Utils("⚠ Notes repo path not set (env.ahk).", 3000, BANNER_ACCENT_INTERMEDIATE)
        return false
    }
    fullPath := RTrim(basePath, "\") . relPath
    if (StrLower(SubStr(fullPath, -3)) = "pdf") {
        mdPath := SubStr(fullPath, 1, StrLen(fullPath) - 3) . "md"
    } else {
        mdPath := fullPath
    }
    if (!FileExist(mdPath)) {
        try ShowCenteredOverlay_Utils("❌ Markdown not found: " mdPath, 3500, BANNER_ACCENT_ERROR)
        return false
    }
    QuickLook_OpenPath(mdPath, scrollToEnd)
    return true
}

; Center in work-area rect using Round + clamp (horizontal matches StandardLoadingBar_RepositionToActiveMonitor; vertical added for true center).
StudyTopicSelector_ComputeCenterTopLeftInWorkArea(ml, mt, mr, mb, gw, gh, &cx, &cy) {
    monitorWidth := mr - ml
    monitorHeight := mb - mt
    cx := Round(ml + (monitorWidth - gw) / 2)
    if (cx < ml)
        cx := ml
    if (cx + gw > mr)
        cx := mr - gw
    cy := Round(mt + (monitorHeight - gh) / 2)
    if (cy < mt)
        cy := mt
    if (cy + gh > mb)
        cy := mb - gh
}

; Initial placement: GetActiveMonitorWorkArea_StandardBar (same source as StandardLoadingBar / standard_information_display.md); Outlook Copilot uses an equivalent MonitorGetWorkArea loop in Shift keys.ahk.
StudyTopicSelector_PositionGuiLikeOutlook(gui) {
    GetActiveMonitorWorkArea_StandardBar(&ml, &mt, &mr, &mb)
    gui.Show("AutoSize Hide")
    gui.GetPos(, , &gw, &gh)
    StudyTopicSelector_ComputeCenterTopLeftInWorkArea(ml, mt, mr, mb, gw, gh, &cx, &cy)
    ; Avoid "NA": if focus stays in QuickLook, Esc is consumed there first (ShowOutlookCopilotSelector comment).
    gui.Show("x" . cx . " y" . cy)
    try WinActivate(gui.Hwnd)
}

StudyTopicSelector_StopActiveMonitorTracking() {
    try SetTimer(StudyTopicSelector_TrackActiveMonitorTick, 0)
}

; forMonitorIdx: 1-based index from StudyTopicSelector_TrackActiveMonitorTick (same as GetMonitorIndexForForeground_StandardBar).
; Use Show("AutoSize Hide") then Show("x y") like StudyTopicSelector_PositionGuiLikeOutlook — Move() alone can mis-center across mixed-DPI monitors.
StudyTopicSelector_RepositionToActiveMonitor(forMonitorIdx := 0) {
    global g_StudyTopicSelectorGui, g_StudyTopicSelectorActive
    if (!IsObject(g_StudyTopicSelectorGui) || !g_StudyTopicSelectorGui.Hwnd)
        return
    idx := forMonitorIdx
    if (idx < 1 || idx > MonitorGetCount())
        idx := GetMonitorIndexForForeground_StandardBar()
    MonitorGetWorkArea(idx, &ml, &mt, &mr, &mb)
    try {
        g_StudyTopicSelectorGui.Show("AutoSize Hide")
        g_StudyTopicSelectorGui.GetPos(, , &gw, &gh)
    } catch {
        return
    }
    StudyTopicSelector_ComputeCenterTopLeftInWorkArea(ml, mt, mr, mb, gw, gh, &cx, &cy)
    try {
        g_StudyTopicSelectorGui.Show("x" . cx . " y" . cy)
        if (g_StudyTopicSelectorActive)
            try WinActivate(g_StudyTopicSelectorGui.Hwnd)
    } catch {
    }
}

; Follow foreground window's monitor while the selector is open (parity with StandardLoadingBar trackActiveMonitor).
StudyTopicSelector_TrackActiveMonitorTick() {
    global g_StudyTopicSelectorGui, g_StudyTopicSelectorActive, g_StudyTopicSelectorLastForegroundMonitorIdx
    if (!g_StudyTopicSelectorActive || !IsObject(g_StudyTopicSelectorGui) || !g_StudyTopicSelectorGui.Hwnd) {
        StudyTopicSelector_StopActiveMonitorTracking()
        return
    }
    newIdx := GetMonitorIndexForForeground_StandardBar()
    if (newIdx != g_StudyTopicSelectorLastForegroundMonitorIdx) {
        StudyTopicSelector_RepositionToActiveMonitor(newIdx)
        g_StudyTopicSelectorLastForegroundMonitorIdx := newIdx
    }
}

StudyTopicSelector_UnbindCategoryHotkeys() {
    try Hotkey("0", "Off")
    try Hotkey("1", "Off")
    try Hotkey("2", "Off")
    try Hotkey("3", "Off")
    try Hotkey("4", "Off")
    try Hotkey("5", "Off")
    try Hotkey("6", "Off")
}

StudyTopicSelector_UnbindDigitHotkeys() {
    loop 7 {
        try Hotkey(String(A_Index - 1), "Off")
    }
}

; Poll Esc — fallback when $*Escape / g_OnEscapePressed miss (OutlookCopilotSelector_EscapePoll).
StudyTopicSelector_EscapePoll() {
    global g_StudyTopicSelectorActive, g_StudyTopicEscPollPrev
    if (!g_StudyTopicSelectorActive) {
        SetTimer(StudyTopicSelector_EscapePoll, 0)
        return
    }
    escSync := GetKeyState("Escape", "P")
    escAsync := (DllCall("user32\GetAsyncKeyState", "int", 0x1B) & 0x8000) != 0
    escDown := escSync || escAsync
    if (escDown) {
        if (!g_StudyTopicEscPollPrev) {
            g_StudyTopicEscPollPrev := true
            StudyTopicSelector_Cancel()
        }
    } else {
        g_StudyTopicEscPollPrev := false
    }
}

StudyTopicSelector_EscapeFromHotkey(*) {
    StudyTopicSelector_Cancel()
}

StudyTopicSelector_GlobalEscapeCallback(*) {
    StudyTopicSelector_Cancel()
}

StudyTopicSelector_GuiEscape(*) {
    StudyTopicSelector_Cancel()
}

; Same escape registration order as ShowOutlookCopilotSelector (after Gui.OnEvent registered at build time).
StudyTopicSelector_BindRobustEscape() {
    global g_StudyTopicSelectorGui, g_OnEscapePressed, g_StudyTopicEscPollPrev
    SetTimer(StudyTopicSelector_EscapePoll, 0)
    if (!IsObject(g_StudyTopicSelectorGui) || !g_StudyTopicSelectorGui.Hwnd)
        return
    Hotkey("$*Escape", StudyTopicSelector_EscapeFromHotkey, "On")
    global g_OnEscapePressed
    g_OnEscapePressed := StudyTopicSelector_GlobalEscapeCallback
    Utils_EnsureGlobalEscapeHotkey()
    g_StudyTopicEscPollPrev := false
    SetTimer(StudyTopicSelector_EscapePoll, 50)
}

StudyTopicSelector_UnbindRobustEscape() {
    global g_OnEscapePressed, g_StudyTopicEscPollPrev
    SetTimer(StudyTopicSelector_EscapePoll, 0)
    g_StudyTopicEscPollPrev := false
    try Hotkey("Escape", StudyTopicSelector_Cancel, "Off")
    catch {
    }
    try Hotkey("*Escape", StudyTopicSelector_Cancel, "Off")
    catch {
    }
    try Hotkey("$*Escape", StudyTopicSelector_EscapeFromHotkey, "Off")
    catch {
    }
    global g_OnEscapePressed
    g_OnEscapePressed := ""
    Utils_EnsureGlobalEscapeHotkey()
}

; Study material (`#!+x`): Gui.Destroy() throws if the window never existed or was already destroyed — swallow and continue.
StudyTopicSelector_SafeDestroyGui(gui) {
    if (!IsObject(gui))
        return
    try gui.Destroy()
    catch {
    }
}

; Category menu (Technique README / Mnemonics / Plans). Bind Escape + Backspace to cancel; Backspace on topic menu goes back via StudyTopicSelector_BackFromTopic.
StudyTopicSelector_ShowCategoryPhase() {
    global g_StudyTopicSelectorGui, g_StudyTopics

    StudyTopicSelector_SafeDestroyGui(g_StudyTopicSelectorGui)
    g_StudyTopicSelectorGui := false

    g_StudyTopicSelectorGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Owner -DPIScale")
    g_StudyTopicSelectorGui.BackColor := "1E1E2E"
    g_StudyTopicSelectorGui.MarginX := 20
    g_StudyTopicSelectorGui.MarginY := 15

    g_StudyTopicSelectorGui.SetFont("s14 cCDD6F4 Bold", "Segoe UI")
    g_StudyTopicSelectorGui.Add("Text", "w300 Center", "📚 Study material (QuickLook)")
    g_StudyTopicSelectorGui.Add("Text", "w300 h1 Background45475A")

    g_StudyTopicSelectorGui.SetFont("s12 cCDD6F4", "Segoe UI")
    g_StudyTopicSelectorGui.Add("Text", "w300", "[1] Mnemonics")
    g_StudyTopicSelectorGui.Add("Text", "w300", "[2] Plans")
    g_StudyTopicSelectorGui.Add("Text", "w300", "[3] Manage Study Subtopic Link")
    g_StudyTopicSelectorGui.Add("Text", "w300", "[4] Manage Study Article Link")
    g_StudyTopicSelectorGui.Add("Text", "w300", "[5] Manage Study Favorite Link")
    g_StudyTopicSelectorGui.Add("Text", "w300", "[6] Technique")
    g_StudyTopicSelectorGui.Add("Text", "w300 h1 Background45475A y+10")
    g_StudyTopicSelectorGui.SetFont("s9 c6C7086", "Segoe UI")
    g_StudyTopicSelectorGui.Add("Text", "w300 Center", "Press 1-6 | Backspace/Esc to cancel")

    try {
        g_StudyTopicSelectorGui.OnEvent("Escape", StudyTopicSelector_GuiEscape)
    } catch {
    }
    StudyTopicSelector_PositionGuiLikeOutlook(g_StudyTopicSelectorGui)
    global g_StudyTopicSelectorActive
    g_StudyTopicSelectorActive := true

    Hotkey("1", StudyTopicSelector_SelectMnemonics, "On")
    Hotkey("2", StudyTopicSelector_SelectPlans, "On")
    Hotkey("3", StudyTopicSelector_ManageLinks, "On")
    Hotkey("4", StudyTopicSelector_ManageArticleLinks, "On")
    Hotkey("5", StudyTopicSelector_ManageFavoriteLinks, "On")
    Hotkey("6", StudyTopicSelector_SelectTechnique, "On")
    Hotkey("Backspace", StudyTopicSelector_Cancel, "On")
    StudyTopicSelector_BindRobustEscape()

}

; After closing Manage Links GUI (Esc) or finishing Open Link from submenu — main category Gui still exists.
StudyTopicSelector_ResumeSelectorEscapeAfterLinks(*) {
    global g_StudyTopicSelectorActive, g_StudyTopicSelectorGui
    if (g_StudyTopicSelectorActive && IsObject(g_StudyTopicSelectorGui) && g_StudyTopicSelectorGui.Hwnd)
        StudyTopicSelector_BindRobustEscape()
}

StudyTopicSelector_ManageLinksEsc(*) {
    global g_StudyLinksGui
    StudyTopicSelector_SafeDestroyGui(g_StudyLinksGui)
    g_StudyLinksGui := false
    StudyTopicSelector_ResumeSelectorEscapeAfterLinks()
}

; Persistent global GUI for link management submenu
global g_StudyLinksGui := false
StudyTopicSelector_ManageLinks(*) {
    global g_StudyLinksGui
    StudyLink_EnsureManageSubtopicSentinel()
    StudyTopicSelector_UnbindRobustEscape()
    StudyTopicSelector_SafeDestroyGui(g_StudyLinksGui)
    g_StudyLinksGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Owner -DPIScale")
    ; Unbind previous hotkeys to avoid conflicts
    loop 9 {
        try Hotkey(String(A_Index), "Off")
    }
    try Hotkey("Escape", "Off")
    g_StudyLinksGui.BackColor := "1E1E2E"
    g_StudyLinksGui.MarginX := 20
    g_StudyLinksGui.MarginY := 15
    g_StudyLinksGui.SetFont("s14 cCDD6F4 Bold", "Segoe UI")
    g_StudyLinksGui.Add("Text", "w400 Center", "🔗 Manage Study Subtopic Link")
    g_StudyLinksGui.Add("Text", "w400 h1 Background45475A")
    g_StudyLinksGui.SetFont("s11 cCDD6F4", "Segoe UI")
    ytResult := StudyLink_GetResult(STUDYLINK_KEY_YOUTUBE)
    if (ytResult["ok"])
        StudyLink_PlayApiSuccessSound()
    g_StudyLinksGui.Add("Text", "w400", "Current YouTube link: " . StudyLink_FormatLinkLabel(ytResult))
    g_StudyLinksGui.Add("Text", "w400", "[1] Open YouTube link")
    g_StudyLinksGui.Add("Text", "w400", "[2] Set YouTube link")
    g_StudyLinksGui.Add("Text", "w400 h1 Background45475A y+10")
    g_StudyLinksGui.SetFont("s9 c6C7086", "Segoe UI")
    g_StudyLinksGui.Add("Text", "w400 Center", "Press 1-2 | Esc to cancel")
    StudyTopicSelector_PositionGuiLikeOutlook(g_StudyLinksGui)
    Hotkey("1", StudyTopicSelector_ManageLinks_Open, "On")
    Hotkey("2", StudyTopicSelector_ManageLinks_Set, "On")
    Hotkey("Escape", StudyTopicSelector_ManageLinksEsc, "On")
    g_StudyLinksGui.Show()
}

; [1] Open the saved YouTube subtopic link in Google Chrome
StudyTopicSelector_ManageLinks_Open(*) {
    StudyTopicSelector_Close()
    linkResult := StudyLink_GetResult(STUDYLINK_KEY_YOUTUBE)
    if (!linkResult["ok"]) {
        ShowCenteredOverlay_Utils("❌ Could not load link from API: " . linkResult["err"], 3500, BANNER_ACCENT_ERROR)
        return
    }
    StudyLink_PlayApiSuccessSound()
    url := linkResult["url"]
    if (url != "") {
        StudyLink_OpenUrlInChrome(url, true)
        ShowCenteredOverlay_Utils("✅ Opening YouTube link in a new Chrome window...", 2000, BANNER_ACCENT_SUCCESS)
    } else {
        ShowCenteredOverlay_Utils("⚠ No YouTube link stored. Use [2] Set YouTube link first.", 2500,
            BANNER_ACCENT_INTERMEDIATE)
    }
}

StudyLink_UiaInvokeOrClick(el, preferClick := false) {
    if !IsObject(el)
        return false
    global STUDYLINK_YT_MS_BEFORE_CLICK, STUDYLINK_YT_MS_AFTER_CLICK
    Sleep STUDYLINK_YT_MS_BEFORE_CLICK
    try el.SetFocus()
    catch {
    }
    if (preferClick) {
        try {
            el.Click()
            Sleep STUDYLINK_YT_MS_AFTER_CLICK
            return true
        } catch {
        }
    }
    try {
        if el.GetPropertyValue(UIA.Property.IsInvokePatternAvailable) {
            el.Invoke()
            Sleep STUDYLINK_YT_MS_AFTER_CLICK
            return true
        }
    } catch {
    }
    try {
        el.Click()
        Sleep STUDYLINK_YT_MS_AFTER_CLICK
        return true
    } catch {
    }
    return false
}

StudyLink_UiaWaitFor(root, conditions, timeoutMs := 4000) {
    if !IsObject(root)
        return 0
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        el := ClipAngel_UiaFindFirst(root, conditions)
        if el
            return el
        Sleep 80
    }
    return 0
}

; EN + PT-BR YouTube action-bar / share-panel button labels (Name property).
global STUDYLINK_YT_BTN_SHARE := ["Share", "Compartilhar"]
global STUDYLINK_YT_BTN_COPY := ["Copy", "Copiar"]
global STUDYLINK_YT_BTN_CANCEL := ["Cancel", "Cancelar", "Close", "Fechar"]

; UI pacing (ms) — modest gaps so focus/invoke/click registers on YouTube Web UI.
global STUDYLINK_YT_MS_BEFORE_CLICK := 120
global STUDYLINK_YT_MS_AFTER_CLICK := 80
global STUDYLINK_YT_MS_AFTER_SHARE_CLICK := 350
global STUDYLINK_YT_MS_AFTER_START_AT := 250
global STUDYLINK_YT_MS_AFTER_COPY := 450
global STUDYLINK_YT_MS_BEFORE_CLOSE_PANEL := 280

StudyLink_UiaFindButtonByNames(root, nameList, timeoutMs := 0) {
    if !IsObject(root)
        return 0
    deadline := timeoutMs ? A_TickCount + timeoutMs : 0
    loop {
        for name in nameList {
            el := ClipAngel_UiaFindFirst(root, { Type: 50000, Name: name })
            if el
                return el
        }
        if (!timeoutMs || A_TickCount >= deadline)
            break
        Sleep 80
    }
    return 0
}

StudyLink_IsYoutubeVideoPageUrl(url) {
    return InStr(url, "youtube.com/watch") || InStr(url, "youtube.com/live") || InStr(url, "youtube.com/shorts")
}

StudyLink_CaptureYoutubeTimestampUrl(uia, &errMsg := "") {
    errMsg := ""
    global STUDYLINK_YT_MS_AFTER_SHARE_CLICK, STUDYLINK_YT_MS_AFTER_START_AT
    if !IsObject(uia) {
        errMsg := "Could not attach to Chrome."
        return ""
    }
    try currentUrl := uia.GetCurrentURL()
    catch {
        errMsg := "Could not read the current page URL."
        return ""
    }
    if !StudyLink_IsYoutubeVideoPageUrl(currentUrl) {
        errMsg := "Open a YouTube video page and try again."
        return ""
    }
    shareBtn := StudyLink_UiaFindButtonByNames(uia, STUDYLINK_YT_BTN_SHARE, 3000)
    if !shareBtn {
        errMsg := "Share button not found."
        return ""
    }
    if !StudyLink_UiaInvokeOrClick(shareBtn) {
        errMsg := "Could not open the Share panel."
        return ""
    }
    Sleep STUDYLINK_YT_MS_AFTER_SHARE_CLICK
    startAt := StudyLink_UiaWaitFor(uia, { AutomationId: "start-at-checkbox", Type: 50002 })
    if !startAt {
        errMsg := "Share panel did not open in time."
        return ""
    }
    startAtOn := ClipAngel_FavoriteCellIsOn(startAt)
    if !startAtOn {
        toggled := false
        try {
            if startAt.GetPropertyValue(UIA.Property.IsTogglePatternAvailable) {
                startAt.TogglePattern.Toggle()
                toggled := true
            }
        } catch {
        }
        if (!toggled)
            toggled := StudyLink_UiaInvokeOrClick(startAt)
        if (!toggled) {
            errMsg := "Could not enable Start at."
            return ""
        }
        Sleep STUDYLINK_YT_MS_AFTER_START_AT
    }
    deadline := A_TickCount + 2000
    while (A_TickCount < deadline) {
        shareUrlEl := ClipAngel_UiaFindFirst(uia, { AutomationId: "share-url", Type: 50004 })
        if shareUrlEl {
            try shareVal := Trim(shareUrlEl.Value)
            catch
                shareVal := ""
            if (shareVal != "" && InStr(shareVal, "t="))
                break
        }
        Sleep 80
    }
    shareUrlEl := StudyLink_UiaWaitFor(uia, { AutomationId: "share-url", Type: 50004 }, 2000)
    if !shareUrlEl {
        errMsg := "Share link field not found."
        return ""
    }
    url := ""
    try url := Trim(shareUrlEl.Value)
    catch
        url := ""
    copyBtn := 0
    copyGroup := ClipAngel_UiaFindFirst(uia, { AutomationId: "copy-button" })
    if copyGroup {
        for name in STUDYLINK_YT_BTN_COPY {
            copyBtn := ClipAngel_UiaFindFirst(copyGroup, { Type: 50000, Name: name })
            if copyBtn
                break
        }
    }
    if !copyBtn
        copyBtn := StudyLink_UiaFindButtonByNames(uia, STUDYLINK_YT_BTN_COPY)
    if !copyBtn {
        errMsg := "Copy button not found."
        return ""
    }
    savedClip := A_Clipboard
    if !StudyLink_UiaInvokeOrClick(copyBtn) {
        errMsg := "Could not click Copy."
        return ""
    }
    clipDeadline := A_TickCount + 2000
    clipUrl := ""
    while (A_TickCount < clipDeadline) {
        clipUrl := Trim(A_Clipboard)
        if (clipUrl != "" && clipUrl != savedClip && InStr(clipUrl, "t="))
            break
        Sleep 80
    }
    if (url = "" || !InStr(url, "t=")) {
        if (clipUrl != "" && InStr(clipUrl, "t="))
            url := clipUrl
    }
    if (url = "" || !InStr(url, "youtu") || !InStr(url, "t=")) {
        errMsg := "Could not read the timestamped share link."
        return ""
    }
    global STUDYLINK_YT_MS_AFTER_COPY
    Sleep STUDYLINK_YT_MS_AFTER_COPY
    return url
}

StudyLink_CleanupYoutubeSharePanel(uia, chromeHwnd := 0) {
    global STUDYLINK_YT_MS_BEFORE_CLOSE_PANEL, STUDYLINK_YT_BTN_CANCEL
    if !IsObject(uia)
        return
    if !ClipAngel_UiaFindFirst(uia, { AutomationId: "start-at-checkbox", Type: 50002 })
        && !ClipAngel_UiaFindFirst(uia, { AutomationId: "share-url", Type: 50004 })
        return
    if chromeHwnd {
        try {
            WinActivate("ahk_id " chromeHwnd)
            WinWaitActive("ahk_id " chromeHwnd, , 1)
        } catch {
        }
    }
    Sleep STUDYLINK_YT_MS_BEFORE_CLOSE_PANEL
    cancelBtn := 0
    closeGroup := ClipAngel_UiaFindFirst(uia, { AutomationId: "close-button" })
    if closeGroup {
        for name in STUDYLINK_YT_BTN_CANCEL {
            cancelBtn := ClipAngel_UiaFindFirst(closeGroup, { Type: 50000, Name: name })
            if cancelBtn
                break
        }
        if !cancelBtn {
            try cancelBtn := closeGroup.FindFirst({ Type: 50000, AutomationId: "button" })
            catch {
                cancelBtn := 0
            }
        }
    }
    if !cancelBtn
        cancelBtn := StudyLink_UiaFindButtonByNames(uia, STUDYLINK_YT_BTN_CANCEL)
    if cancelBtn
        StudyLink_UiaInvokeOrClick(cancelBtn, true)
    Sleep 220
    panelStillOpen := !!ClipAngel_UiaFindFirst(uia, { AutomationId: "share-url", Type: 50004 })
    if (!panelStillOpen)
        return
    if chromeHwnd {
        try {
            WinActivate("ahk_id " chromeHwnd)
            WinWaitActive("ahk_id " chromeHwnd, , 1)
        } catch {
        }
    }
    try Send "{Escape}"
    catch {
    }
    Sleep 250
}

; [2] Set the link: YouTube Share + Start at + Copy, save via API
StudyTopicSelector_ManageLinks_Set(*) {
    StudyTopicSelector_Close()
    try {
        chromeHwnd := WinExist("ahk_class Chrome_WidgetWin_1")
        if chromeHwnd {
            WinActivate("ahk_id " chromeHwnd)
            if !WinWaitActive("ahk_id " chromeHwnd, , 2) {
                ShowCenteredOverlay_Utils("❌ Chrome window did not become active.", 2500, BANNER_ACCENT_ERROR)
                return
            }
        } else {
            ShowCenteredOverlay_Utils("❌ Chrome window not found. Switch to Chrome and try again.", 3000,
                BANNER_ACCENT_ERROR)
            return
        }
        Sleep 280
        try UIA.ActivateChromiumAccessibility("ahk_id " chromeHwnd, 300)
        catch {
        }
        uia := UIA_Browser("ahk_id " chromeHwnd)
        errMsg := ""
        url := StudyLink_CaptureYoutubeTimestampUrl(uia, &errMsg)
        if (url != "") {
            ; Close share panel while Chrome still has focus (before loading overlay steals it).
            StudyLink_CleanupYoutubeSharePanel(uia, chromeHwnd)
            StandardLoadingBar_Show("Saving link…", BANNER_ACCENT_INTERMEDIATE, { passive: false })
            setOk := StudyLink_Set(STUDYLINK_KEY_YOUTUBE, url)
            try StandardLoadingBar_Hide(0)
            if setOk
                ShowCenteredOverlay_Utils("✅ Link saved to study notes.", 3000, BANNER_ACCENT_SUCCESS)
            else
                ShowCenteredOverlay_Utils("❌ Could not save the link (API failed).", 3500, BANNER_ACCENT_ERROR)
        } else {
            StudyLink_CleanupYoutubeSharePanel(uia, chromeHwnd)
            ShowCenteredOverlay_Utils(errMsg != "" ? "❌ " errMsg : "❌ Could not capture the link.", 2500,
                BANNER_ACCENT_ERROR)
        }
    } catch as e {
        ShowCenteredOverlay_Utils("❌ Error: " . e.Message, 3000, BANNER_ACCENT_ERROR)
    }
}

ShowStudyTopicSelector() {
    global g_StudyTopicSelectorGui, g_StudyTopicSelectorActive, g_StudyTopics, g_StudyTopicSelectorPhase,
        g_StudyTopicSelectorCategory, g_StudyTopicSelectorLastForegroundMonitorIdx

    if (g_StudyTopicSelectorActive)
        return

    g_StudyTopicSelectorCategory := ""
    g_StudyTopicSelectorPhase := "category"

    StudyTopicSelector_ShowCategoryPhase()

    StudyTopicSelector_StopActiveMonitorTracking()
    g_StudyTopicSelectorLastForegroundMonitorIdx := GetMonitorIndexForForeground_StandardBar()
    SetTimer(StudyTopicSelector_TrackActiveMonitorTick, 115)
}

StudyTopicSelector_ShowTopicPhase() {
    global g_StudyTopicSelectorGui, g_StudyTopicSelectorActive, g_StudyTopics, g_StudyTopicSelectorPhase,
        g_StudyTopicSelectorCategory, g_StudyTopicSelectorLastForegroundMonitorIdx

    if (!g_StudyTopicSelectorActive || g_StudyTopicSelectorPhase != "category")
        return

    StudyTopicSelector_UnbindCategoryHotkeys()
    try Hotkey("Backspace", StudyTopicSelector_Cancel, "Off")

    catLabel := (g_StudyTopicSelectorCategory = "plans") ? "Plans" : "Mnemonics"
    StudyTopicSelector_SafeDestroyGui(g_StudyTopicSelectorGui)
    g_StudyTopicSelectorGui := false

    g_StudyTopicSelectorGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Owner -DPIScale")
    g_StudyTopicSelectorGui.BackColor := "1E1E2E"
    g_StudyTopicSelectorGui.MarginX := 20
    g_StudyTopicSelectorGui.MarginY := 15

    g_StudyTopicSelectorGui.SetFont("s14 cCDD6F4 Bold", "Segoe UI")
    g_StudyTopicSelectorGui.Add("Text", "w300 Center", "📚 " . catLabel . " — topic")
    g_StudyTopicSelectorGui.Add("Text", "w300 h1 Background45475A")

    g_StudyTopicSelectorGui.SetFont("s12 cCDD6F4", "Segoe UI")
    for num, topic in g_StudyTopics {
        n := Integer(num)
        if (g_StudyTopicSelectorCategory = "mnemonics" && n = 0)
            continue
        g_StudyTopicSelectorGui.Add("Text", "w300", "[" . num . "] " . topic.name)
    }

    g_StudyTopicSelectorGui.Add("Text", "w300 h1 Background45475A y+10")
    g_StudyTopicSelectorGui.SetFont("s9 c6C7086", "Segoe UI")
    footerHint := (g_StudyTopicSelectorCategory = "mnemonics") ? "Press 1-6 | Backspace back | Esc cancel" :
        "Press 0-6 | Backspace back | Esc cancel"
    g_StudyTopicSelectorGui.Add("Text", "w300 Center", footerHint)

    try {
        g_StudyTopicSelectorGui.OnEvent("Escape", StudyTopicSelector_GuiEscape)
    } catch {
    }
    StudyTopicSelector_PositionGuiLikeOutlook(g_StudyTopicSelectorGui)
    g_StudyTopicSelectorLastForegroundMonitorIdx := GetMonitorIndexForForeground_StandardBar()

    g_StudyTopicSelectorPhase := "topic"
    if (g_StudyTopicSelectorCategory = "plans") {
        Hotkey("0", StudyTopicSelector_HandleKey, "On")
        Hotkey("1", StudyTopicSelector_HandleKey, "On")
        Hotkey("2", StudyTopicSelector_HandleKey, "On")
        Hotkey("3", StudyTopicSelector_HandleKey, "On")
        Hotkey("4", StudyTopicSelector_HandleKey, "On")
        Hotkey("5", StudyTopicSelector_HandleKey, "On")
        Hotkey("6", StudyTopicSelector_HandleKey, "On")
    } else {
        Hotkey("1", StudyTopicSelector_HandleKey, "On")
        Hotkey("2", StudyTopicSelector_HandleKey, "On")
        Hotkey("3", StudyTopicSelector_HandleKey, "On")
        Hotkey("4", StudyTopicSelector_HandleKey, "On")
        Hotkey("5", StudyTopicSelector_HandleKey, "On")
        Hotkey("6", StudyTopicSelector_HandleKey, "On")
    }
    Hotkey("Backspace", StudyTopicSelector_BackFromTopic, "On")
    StudyTopicSelector_BindRobustEscape()
}

StudyTopicSelector_BackFromTopic(*) {
    global g_StudyTopicSelectorGui, g_StudyTopicSelectorActive, g_StudyTopicSelectorPhase, g_StudyTopicSelectorCategory,
        g_StudyTopicSelectorLastForegroundMonitorIdx

    if (!g_StudyTopicSelectorActive || g_StudyTopicSelectorPhase != "topic")
        return
    StudyTopicSelector_UnbindDigitHotkeys()
    try Hotkey("Backspace", StudyTopicSelector_BackFromTopic, "Off")
    g_StudyTopicSelectorCategory := ""
    g_StudyTopicSelectorPhase := "category"
    StudyTopicSelector_SafeDestroyGui(g_StudyTopicSelectorGui)
    g_StudyTopicSelectorGui := false
    StudyTopicSelector_ShowCategoryPhase()
    g_StudyTopicSelectorLastForegroundMonitorIdx := GetMonitorIndexForForeground_StandardBar()
}

; [1] Technique: open studies/technique/README.md in QuickLook (single file). Technique plans: [3] Plans then [0].
StudyTopicSelector_SelectTechnique(*) {
    global g_StudyTopicSelectorActive, g_StudyTopicSelectorPhase, g_StudyTopics
    if (!g_StudyTopicSelectorActive || g_StudyTopicSelectorPhase != "category")
        return
    if (!g_StudyTopics.Has(0))
        return
    StudyTopicSelector_Close()
    topic := g_StudyTopics[0]
    relPath := StudyTopic_GetRelPath(topic, "mnemonics")
    StudyTopic_OpenRepoRelativeMarkdown(relPath, true)
}

StudyTopicSelector_SelectMnemonics(*) {
    global g_StudyTopicSelectorActive, g_StudyTopicSelectorPhase, g_StudyTopicSelectorCategory
    if (!g_StudyTopicSelectorActive || g_StudyTopicSelectorPhase != "category")
        return
    g_StudyTopicSelectorCategory := "mnemonics"
    StudyTopicSelector_ShowTopicPhase()
}

StudyTopicSelector_SelectPlans(*) {
    global g_StudyTopicSelectorActive, g_StudyTopicSelectorPhase, g_StudyTopicSelectorCategory
    if (!g_StudyTopicSelectorActive || g_StudyTopicSelectorPhase != "category")
        return
    g_StudyTopicSelectorCategory := "plans"
    StudyTopicSelector_ShowTopicPhase()
}

StudyTopicSelector_HandleKey(key) {
    global g_StudyTopicSelectorActive, g_StudyTopics, g_StudyTopicSelectorPhase, g_StudyTopicSelectorCategory

    if (!g_StudyTopicSelectorActive || g_StudyTopicSelectorPhase != "topic")
        return
    selection := Integer(key)
    category := g_StudyTopicSelectorCategory
    StudyTopicSelector_Close()
    if (!g_StudyTopics.Has(selection))
        return

    topic := g_StudyTopics[selection]
    relPath := StudyTopic_GetRelPath(topic, category)
    StudyTopic_OpenRepoRelativeMarkdown(relPath, category != "plans")
}

StudyTopicSelector_Cancel(*) {
    StudyTopicSelector_Close()
}

; Tear-down order aligned with OutlookCopilotSelector_Close (Shift keys.ahk).
StudyTopicSelector_Close() {
    global g_StudyTopicSelectorGui, g_StudyTopicSelectorActive, g_StudyTopicSelectorPhase, g_StudyTopicSelectorCategory,
        g_StudyTopicSelectorLastForegroundMonitorIdx, g_StudyLinksGui, g_StudyArticleLinksGui, g_StudyFavoriteLinksGui,
        g_StudyLinkSubmenuGui

    if (!g_StudyTopicSelectorActive)
        return
    StudyTopicSelector_UnbindRobustEscape()
    g_StudyTopicSelectorActive := false
    g_StudyTopicSelectorPhase := ""
    g_StudyTopicSelectorCategory := ""

    StudyTopicSelector_StopActiveMonitorTracking()
    g_StudyTopicSelectorLastForegroundMonitorIdx := 0

    StudyTopicSelector_UnbindCategoryHotkeys()
    StudyTopicSelector_UnbindDigitHotkeys()
    try Hotkey("Backspace", "Off")
    catch {
    }
    Utils_EnsureGlobalEscapeHotkey()
    StudyTopicSelector_SafeDestroyGui(g_StudyTopicSelectorGui)
    g_StudyTopicSelectorGui := false
    StudyTopicSelector_SafeDestroyGui(g_StudyLinksGui)
    g_StudyLinksGui := false
    StudyTopicSelector_SafeDestroyGui(g_StudyArticleLinksGui)
    g_StudyArticleLinksGui := false
    StudyTopicSelector_SafeDestroyGui(g_StudyFavoriteLinksGui)
    g_StudyFavoriteLinksGui := false
    StudyTopicSelector_SafeDestroyGui(g_StudyLinkSubmenuGui)
    g_StudyLinkSubmenuGui := ""
    loop 9 {
        try Hotkey(String(A_Index), "Off")
    }
    try Hotkey("Escape", "Off")
}

PeekPdf_GetIniPath() {
    return A_ScriptDir "\data\peek_pdf.ini"
}

PeekPdf_NormalizePath(path) {
    path := Trim(path)
    q := Chr(34)
    if (SubStr(path, 1, 1) = q && SubStr(path, -1) = q)
        path := SubStr(path, 2, StrLen(path) - 2)
    return Trim(path)
}

; Resolve the Peek executable path.
; Priority: 1) INI [Peek] ExePath  2) Environment-specific path (GetPeekExePath)  3) "peek.exe" (PATH)
PeekPdf_ResolvePeekExePath() {
    iniPath := PeekPdf_GetIniPath()
    exePath := ""
    try exePath := IniRead(iniPath, "Peek", "ExePath", "")
    exePath := PeekPdf_NormalizePath(exePath)
    if (exePath != "" && FileExist(exePath))
        return exePath
    ; Legacy: this previously used GetPeekExePath() and PowerToys Peek.
    ; QuickLook is now the primary study viewer; for compatibility, fall back to QuickLook.
    envExe := GetQuickLookExePath()
    if (FileExist(envExe))
        return envExe
    return "QuickLook.exe"
}

QuickLook_GetIniPath() {
    return A_ScriptDir "\data\quicklook.ini"
}

QuickLook_NormalizePath(path) {
    path := Trim(path)
    q := Chr(34)
    if (SubStr(path, 1, 1) = q && SubStr(path, -1) = q)
        path := SubStr(path, 2, StrLen(path) - 2)
    return Trim(path)
}

; Note: `GetQuickLookExePath()` is defined in `env.ahk` (environment-specific paths).

QuickLook_ResolveExePath() {
    iniPath := QuickLook_GetIniPath()
    exePath := ""
    try exePath := IniRead(iniPath, "QuickLook", "ExePath", "")
    exePath := QuickLook_NormalizePath(exePath)
    if (exePath != "" && FileExist(exePath))
        return exePath
    envExe := GetQuickLookExePath()
    if (FileExist(envExe))
        return envExe
    return "QuickLook.exe"
}

; Detect whether a given monitor index is currently connected and has a usable work area.
; Uses AHK built-ins so we don't depend on custom geometry assumptions.
IsMonitorConnected(monitorIndex) {
    try monitorCount := MonitorGetCount()
    catch
        return false
    if (monitorIndex < 1 || monitorIndex > monitorCount)
        return false
    try {
        MonitorGetWorkArea(monitorIndex, &l, &t, &r, &b)
        if ((r - l) <= 0 || (b - t) <= 0)
            return false
    } catch {
        return false
    }
    return true
}

; Map a Windows display device name (e.g. "\\.\DISPLAY2") to the current AHK monitor index.
; Returns 0 when not found/connected.
GetMonitorIndexByDeviceName(deviceName) {
    if (deviceName = "")
        return 0
    try monitorCount := MonitorGetCount()
    catch
        return 0
    loop monitorCount {
        idx := A_Index
        nm := ""
        try nm := MonitorGetName(idx)
        if (nm = deviceName)
            return idx
    }
    return 0
}

; Preferred target for "main monitor (2)" in this setup: Windows DISPLAY2.
; Falls back to the primary monitor if DISPLAY2 isn't present.
GetQuickLookTargetMonitorIndex() {
    idx := GetMonitorIndexByDeviceName("\\.\DISPLAY2")
    if (idx && IsMonitorConnected(idx))
        return idx
    try return MonitorGetPrimary()
    catch
        return 1
}

; Poll until QuickLook.exe has a window (cold start can exceed 2s). Returns hwnd or 0.
QuickLook_WaitForHwnd(timeoutMs := 10000) {
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        hwnd := WinExist("ahk_exe QuickLook.exe")
        if (hwnd)
            return hwnd
        Sleep 75
    }
    return 0
}

QuickLook_WindowCenterOnMonitor(hwnd, targetMon) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    try {
        rect := Buffer(16, 0)
        if (!DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect))
            return false
        wl := NumGet(rect, 0, "int"), wt := NumGet(rect, 4, "int")
        wr := NumGet(rect, 8, "int"), wb := NumGet(rect, 12, "int")
        cx := wl + (wr - wl) // 2
        cy := wt + (wb - wt) // 2
        MonitorGetWorkArea(targetMon, &ml, &mt, &mr, &mb)
        return (cx >= ml && cx <= mr && cy >= mt && cy <= mb)
    } catch {
        return false
    }
}

QuickLook_WindowLooksMaximized(hwnd, targetMon) {
    if (!hwnd)
        return false
    try {
        if (WinGetMinMax("ahk_id " hwnd) = 2)
            return true
    } catch {
    }
    try {
        MonitorGetWorkArea(targetMon, &ml, &mt, &mr, &mb)
        rect := Buffer(16, 0)
        if (DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect)) {
            wl := NumGet(rect, 0, "int"), wt := NumGet(rect, 4, "int")
            wr := NumGet(rect, 8, "int"), wb := NumGet(rect, 12, "int")
            wArea := (mr - ml) * (mb - mt)
            wWin := (wr - wl) * (wb - wt)
            if (wArea > 0 && wWin >= 0.9 * wArea)
                return true
        }
    } catch {
    }
    return false
}

; Bounded wait: center on target monitor and maximized (or ~90% work area).
QuickLook_WaitForLayoutReady(hwnd, targetMon, timeoutMs := 2500) {
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        if (!WinExist("ahk_id " hwnd))
            return false
        if (QuickLook_WindowCenterOnMonitor(hwnd, targetMon) && QuickLook_WindowLooksMaximized(hwnd, targetMon))
            return true
        Sleep 50
    }
    return false
}

QuickLook_FindDocumentElement(hwnd) {
    if (!hwnd)
        return 0
    try {
        root := UIA.ElementFromHandle(hwnd)
        doc := root.FindFirst({ Type: UIA.ControlType.Document })
        if (doc)
            return doc
    } catch {
    }
    return 0
}

; UIA Document present and enabled on two consecutive polls (WebView settled).
QuickLook_WaitForViewerReady(hwnd, timeoutMs := 3000) {
    deadline := A_TickCount + timeoutMs
    stableCount := 0
    while (A_TickCount < deadline) {
        if (!WinExist("ahk_id " hwnd))
            return false
        doc := QuickLook_FindDocumentElement(hwnd)
        ok := false
        if (doc) {
            try {
                ok := doc.GetPropertyValue(UIA.Property.IsEnabled)
            } catch {
                ok := true
            }
        }
        if (ok) {
            stableCount++
            if (stableCount >= 2)
                return true
            Sleep 100
        } else {
            stableCount := 0
            Sleep 75
        }
    }
    return false
}

QuickLook_ClickWindowCenter(hwnd) {
    if (!hwnd)
        return
    try {
        rect := Buffer(16, 0)
        if (DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect)) {
            wl := NumGet(rect, 0, "int"), wt := NumGet(rect, 4, "int")
            wr := NumGet(rect, 8, "int"), wb := NumGet(rect, 12, "int")
            CoordMode("Mouse", "Screen")
            Click wl + (wr - wl) // 2, wt + (wb - wt) // 2
        }
    } catch {
    }
}

QuickLook_ScrollViaUIA(hwnd) {
    doc := QuickLook_FindDocumentElement(hwnd)
    if (!doc)
        return false
    try {
        if (doc.GetPropertyValue(UIA.Property.IsScrollPatternAvailable)) {
            doc.ScrollPattern.SetScrollPercent(-1, 100)
            return true
        }
    } catch {
    }
    return false
}

QuickLook_GetScrollVerticalPercent(hwnd) {
    doc := QuickLook_FindDocumentElement(hwnd)
    if (!doc)
        return -1.0
    try {
        if (doc.GetPropertyValue(UIA.Property.IsScrollPatternAvailable)) {
            p := doc.GetPropertyValue(UIA.Property.ScrollVerticalScrollPercent)
            if (p >= 0)
                return p
        }
    } catch {
    }
    return -1.0
}

QuickLook_ScrollViaKeystroke(hwnd) {
    try {
        WinActivate("ahk_id " hwnd)
        WinWaitActive("ahk_id " hwnd, , 1)
        SendInput("^{End}")
        return true
    } catch {
        try {
            WinActivate("ahk_id " hwnd)
            WinWaitActive("ahk_id " hwnd, , 1)
            Send("{Ctrl down}{End}{Ctrl up}")
            return true
        } catch {
            try {
                ControlSend("^End", "ahk_id " hwnd)
                return true
            } catch {
                return false
            }
        }
    }
}

; UIA scroll first, then Ctrl+End; verify vertical % when available (extraRetries for slow title gate).
QuickLook_ScrollToEnd(hwnd, extraRetries := 0) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return
    maxAttempts := 3 + extraRetries
    deadline := A_TickCount + 2000
    loop maxAttempts {
        if (A_TickCount >= deadline)
            break
        QuickLook_ScrollViaUIA(hwnd)
        pct := QuickLook_GetScrollVerticalPercent(hwnd)
        if (pct >= 95.0)
            return
        QuickLook_ScrollViaKeystroke(hwnd)
        pct := QuickLook_GetScrollVerticalPercent(hwnd)
        if (pct >= 95.0)
            return
        if (A_Index < maxAttempts)
            Sleep 150
    }
}

; Single authority: DISPLAY2 (or fallback), maximize, focus, optional scroll-to-end.
QuickLook_ApplyStudyLayout(hwnd, scrollToEnd := true, extraScrollRetries := 0) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    targetMon := GetQuickLookTargetMonitorIndex()
    MoveWindowToMonitor(hwnd, targetMon)
    TryMaximizeWindow(hwnd)
    QuickLook_WaitForLayoutReady(hwnd, targetMon, 2500)
    try {
        WinShow("ahk_id " hwnd)
        WinActivate("ahk_id " hwnd)
        WinWaitActive("ahk_id " hwnd, , 1)
    } catch {
    }
    QuickLook_ClickWindowCenter(hwnd)
    if (scrollToEnd)
        QuickLook_ScrollToEnd(hwnd, extraScrollRetries)
    return true
}

QuickLook_DeferredLayoutAfterStart(*) {
    global g_QuickLookDeferredLayoutScroll, g_QuickLookDeferredLayoutPath
    scrollToEnd := g_QuickLookDeferredLayoutScroll
    path := g_QuickLookDeferredLayoutPath
    g_QuickLookDeferredLayoutPath := ""
    hwnd := WinExist("ahk_exe QuickLook.exe")
    if (!hwnd)
        return
    extra := 0
    if (path != "") {
        gate := QuickLook_WaitForOpenReady(hwnd, path, 6000)
        if (!gate["ok"])
            return
        if (STUDY_TOPIC_QL_STRICT_LAYOUT)
            QuickLook_WaitForViewerReady(hwnd, 4000)
        extra := gate["fallback"] ? 2 : 0
    }
    QuickLook_ApplyStudyLayout(hwnd, scrollToEnd, extra)
    StudyTopic_StartBlackoutCountdown(hwnd)
}

; Legacy open path (STUDY_TOPIC_QL_STRICT_LAYOUT := false).
QuickLook_OpenPath_Legacy(path, scrollToEnd := true) {
    if WinWait("ahk_exe QuickLook.exe", , 2) {
        hwnd := WinExist("ahk_exe QuickLook.exe")
        if (hwnd) {
            gate := QuickLook_WaitForOpenReady(hwnd, path)
            if (!gate["ok"]) {
                try ShowCenteredOverlay_Utils("⚠ QuickLook closed before the file finished loading.", 3200,
                    BANNER_ACCENT_INTERMEDIATE)
                return
            }
            if (gate["matched"])
                Sleep 75
            else
                Sleep 50
            gateUsedFallback := gate["fallback"]
            targetMon := GetQuickLookTargetMonitorIndex()
            MoveWindowToMonitor(hwnd, targetMon)
            WinMaximize("ahk_id " hwnd)
            deadline := A_TickCount + 1500
            while (A_TickCount < deadline) {
                try {
                    rect := Buffer(16, 0)
                    if (DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect)) {
                        wl := NumGet(rect, 0, "int"), wt := NumGet(rect, 4, "int")
                        wr := NumGet(rect, 8, "int"), wb := NumGet(rect, 12, "int")
                        cx := wl + (wr - wl) // 2
                        cy := wt + (wb - wt) // 2
                        MonitorGetWorkArea(targetMon, &ml, &mt, &mr, &mb)
                        if (cx >= ml && cx <= mr && cy >= mt && cy <= mb)
                            break
                    }
                } catch {
                }
                Sleep 50
            }
            try {
                WinShow("ahk_id " hwnd)
                WinActivate("ahk_id " hwnd)
                WinWaitActive("ahk_id " hwnd, , 1)
            } catch {
            }
            QuickLook_ClickWindowCenter(hwnd)
            if (scrollToEnd) {
                scrollAttempts := gateUsedFallback ? 2 : 1
                loop scrollAttempts {
                    if (A_Index > 1)
                        Sleep 175
                    try {
                        WinActivate("ahk_id " hwnd)
                        WinWaitActive("ahk_id " hwnd, , 1)
                        SendInput("^{End}")
                    } catch {
                        try {
                            WinActivate("ahk_id " hwnd)
                            WinWaitActive("ahk_id " hwnd, , 1)
                            Send("{Ctrl down}{End}{Ctrl up}")
                        } catch {
                            try ControlSend("^End", "ahk_id " hwnd)
                        }
                    }
                }
            }
            StudyTopic_StartBlackoutCountdown(hwnd)
        }
    }
}

; Wait until QuickLook's title reflects the opened file (QL-Win shows basename) or timeout with graceful fallback.
; Returns Map: ok (hwnd still valid), matched (title contained basename), fallback (timed out; caller may retry scroll).
QuickLook_WaitForOpenReady(hwnd, path, timeoutMs := 8000) {
    SplitPath path, &baseName
    if (baseName = "")
        return Map("ok", false, "matched", false, "fallback", false)
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        if (!WinExist("ahk_id " hwnd))
            return Map("ok", false, "matched", false, "fallback", false)
        try {
            title := WinGetTitle("ahk_id " hwnd)
        } catch {
            Sleep 75
            continue
        }
        tL := StrLower(title)
        if InStr(tL, StrLower(baseName))
            return Map("ok", true, "matched", true, "fallback", false)
        ; Windows sometimes truncates long titles — try first 31 chars of basename
        if (StrLen(baseName) > 31) {
            shortNeedle := StrLower(SubStr(baseName, 1, 31))
            if InStr(tL, shortNeedle)
                return Map("ok", true, "matched", true, "fallback", false)
        }
        Sleep 75
    }
    ; Option A: proceed after extra delay when title never matched (atypical QuickLook build)
    Sleep 300
    if (!WinExist("ahk_id " hwnd))
        return Map("ok", false, "matched", false, "fallback", false)
    return Map("ok", true, "matched", false, "fallback", true)
}

; scrollToEnd: after open, focus viewer and send Ctrl+End (mnemonics); false leaves viewport at top (plans).
QuickLook_OpenPath(path, scrollToEnd := true) {
    quickLookExe := QuickLook_ResolveExePath()
    if (!FileExist(quickLookExe)) {
        try ShowCenteredOverlay_Utils("❌ QuickLook executable not found: " quickLookExe, 2500, BANNER_ACCENT_ERROR)
        return
    }
    if (!FileExist(path)) {
        try ShowCenteredOverlay_Utils("❌ Markdown not found: " path, 3500, BANNER_ACCENT_ERROR)
        return
    }
    try {
        Run('"' quickLookExe '" "' path '"')
    } catch as e {
        try ShowCenteredOverlay_Utils("❌ Failed to open QuickLook: " e.Message, 3000, BANNER_ACCENT_ERROR)
        return
    }
    if (!STUDY_TOPIC_QL_STRICT_LAYOUT) {
        QuickLook_OpenPath_Legacy(path, scrollToEnd)
        return
    }
    hwnd := QuickLook_WaitForHwnd(10000)
    if (!hwnd) {
        try ShowCenteredOverlay_Utils("⚠ QuickLook did not start in time.", 3200, BANNER_ACCENT_INTERMEDIATE)
        global g_QuickLookDeferredLayoutScroll, g_QuickLookDeferredLayoutPath
        g_QuickLookDeferredLayoutScroll := scrollToEnd
        g_QuickLookDeferredLayoutPath := path
        SetTimer(QuickLook_DeferredLayoutAfterStart, -800)
        return
    }
    gate := QuickLook_WaitForOpenReady(hwnd, path)
    if (!gate["ok"]) {
        try ShowCenteredOverlay_Utils("⚠ QuickLook closed before the file finished loading.", 3200,
            BANNER_ACCENT_INTERMEDIATE)
        return
    }
    if (gate["matched"])
        Sleep 75
    else
        Sleep 50
    QuickLook_WaitForViewerReady(hwnd, 3000)
    extraScroll := gate["fallback"] ? 2 : 0
    QuickLook_ApplyStudyLayout(hwnd, scrollToEnd, extraScroll)
    StudyTopic_StartBlackoutCountdown(hwnd)
}

; Open a specific PDF in PowerToys Peek and run WaitAndConfigure. Caller must validate pdfPath and exe exist.
; skipGoToLastPage: if true, do not navigate to the last page (e.g. for short docs like technique README).
PeekPdf_OpenPath(pdfPath, skipGoToLastPage := false) {
    peekExe := PeekPdf_ResolvePeekExePath()
    peekEsc := StrReplace(peekExe, "'", "''")
    pdfEsc := StrReplace(pdfPath, "'", "''")
    psArg := "& " . Chr(39) . peekEsc . Chr(39) . " " . Chr(39) . pdfEsc . Chr(39)
    cmd := "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command " . Chr(34) . psArg . Chr(34)
    try Run cmd, "", "Hide"
    catch as e {
        try ShowCenteredOverlay_Utils("❌ Failed to open Peek: " e.Message, 3000, BANNER_ACCENT_ERROR)
        return
    }
    if WinWait("Peek", "", 5) {
        hwnd := WinExist("ahk_exe PowerToys.Peek.UI.exe")
        if (!hwnd)
            hwnd := WinExist("Peek")
        MoveWindowToMonitor(hwnd, 2)
        WinMaximize("ahk_id " hwnd)
        PeekPdf_WaitAndConfigure(skipGoToLastPage)
        try {
            WinActivate("ahk_id " hwnd)
            WinWaitActive("ahk_id " hwnd, , 1)
        } catch {
        }
        StudyTopic_StartBlackoutCountdown(hwnd)
    }
}

PeekPdf_OpenStored() {
    iniPath := PeekPdf_GetIniPath()

    ; Resolve PDF path per environment, with legacy fallback
    pdfPath := ""
    try {
        if (IS_WORK_ENVIRONMENT) {
            pdfPath := IniRead(iniPath, "Peek", "PdfPathWork", "")
            if (pdfPath = "")
                pdfPath := IniRead(iniPath, "Peek", "PdfPath", "")
        } else {
            pdfPath := IniRead(iniPath, "Peek", "PdfPathPersonal", "")
            if (pdfPath = "")
                pdfPath := IniRead(iniPath, "Peek", "PdfPath", "")
        }
    }

    pdfPath := PeekPdf_NormalizePath(pdfPath)
    if (pdfPath = "") {
        try ShowCenteredOverlay_Utils("⚠ No PDF path set. Hold Win+Alt+Shift+X to set.", 3000,
            BANNER_ACCENT_INTERMEDIATE)
        return
    }
    peekExe := PeekPdf_ResolvePeekExePath()
    if (!FileExist(peekExe)) {
        try ShowCenteredOverlay_Utils("❌ Peek executable not found.", 2500, BANNER_ACCENT_ERROR)
        return
    }
    if (!FileExist(pdfPath)) {
        try ShowCenteredOverlay_Utils("❌ PDF file not found: " pdfPath, 3500, BANNER_ACCENT_ERROR)
        return
    }
    peekEsc := StrReplace(peekExe, "'", "''")
    pdfEsc := StrReplace(pdfPath, "'", "''")
    psArg := "& " . Chr(39) . peekEsc . Chr(39) . " " . Chr(39) . pdfEsc . Chr(39)
    cmd := "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command " . Chr(34) . psArg . Chr(34)
    try Run cmd, "", "Hide"
    catch as e {
        try ShowCenteredOverlay_Utils("❌ Failed to open Peek: " e.Message, 3000, BANNER_ACCENT_ERROR)
        return
    }
    if WinWait("Peek", "", 5) {
        hwnd := WinExist("ahk_exe PowerToys.Peek.UI.exe")
        if (!hwnd)
            hwnd := WinExist("Peek")
        MoveWindowToMonitor(hwnd, 2)
        WinMaximize("ahk_id " hwnd)
        PeekPdf_WaitAndConfigure()
    }
}

MoveWindowToMonitor(hwnd, monitorIndex := 2) {
    if (!hwnd)
        return
    try {
        MonitorGetWorkArea(monitorIndex, &l, &t, &r, &b)
    } catch {
        return
    }
    w := r - l
    h := b - t
    ; Restore before moving, otherwise some apps "teleport" as a 1px/tiny bar.
    try {
        mm := WinGetMinMax("ahk_id " hwnd) ; 1=min,2=max,0=normal
        if (mm != 0) {
            WinRestore("ahk_id " hwnd)
            Sleep 80
        }
    } catch {
    }
    try WinMove(l, t, w, h, "ahk_id " hwnd)
}

; Maximize by hwnd with WinAPI fallback (more reliable than keystrokes).
TryMaximizeWindow(hwnd) {
    if (!hwnd)
        return false
    try {
        WinMaximize("ahk_id " hwnd)
        return true
    } catch {
        try {
            PostMessage 0x0112, 0xF030, , , "ahk_id " hwnd  ; WM_SYSCOMMAND, SC_MAXIMIZE
            return true
        } catch {
            return false
        }
    }
}

; Map a window handle to an AutoHotkey monitor index (1..MonitorGetCount()).
; Needed because MonitorFromWindow returns an HMONITOR handle, not an AHK index.
GetAhkMonitorIndexFromHwnd(hwnd) {
    if (!hwnd)
        return 0
    hMon := 0
    try hMon := DllCall("user32\MonitorFromWindow", "ptr", hwnd, "uint", 2, "ptr") ; MONITOR_DEFAULTTONEAREST
    catch
        return 0
    if (!hMon)
        return 0

    count := 0
    try count := MonitorGetCount()
    catch
        return 0
    if (count < 1)
        return 0
    loop count {
        i := A_Index
        try MonitorGet i, &l, &t, &r, &b
        catch
            continue
        cx := (l + r) // 2
        cy := (t + b) // 2
        ; MonitorFromPoint takes a POINT passed by value (two 32-bit signed ints packed into an int64).
        pt64 := ((cy & 0xFFFFFFFF) << 32) | (cx & 0xFFFFFFFF)
        cur := 0
        try cur := DllCall("user32\MonitorFromPoint", "int64", pt64, "uint", 2, "ptr") ; MONITOR_DEFAULTTONEAREST
        catch
            cur := 0
        if (cur = hMon)
            return i
    }
    return 0
}

; Wait for Peek PDF toolbar to load (Page view button), click it, two-page view, focus; optionally go to last page.
; Current state: PDF opening and window maximization are working correctly.
; Execution order: 1) Get Peek hwnd  2) UIA root from hwnd  3) Poll for "Page view" anchor
;  4) Wait for anchor visible + short delay before click  5) Click Page view  6) Right Arrow
;  7) Click window center  8) If not skipGoToLastPage: go to last page (UIA or Ctrl+End). Fallback: Sleep 400 + Click if UIA or anchor fails.
PeekPdf_WaitAndConfigure(skipGoToLastPage := false) {
    global UIA
    ; Standard loading bar: show for the whole process so user knows when we started and when we finished
    StandardLoadingBar_Show("⏳ Peek PDF: configuring...", BANNER_ACCENT_INTERMEDIATE)
    ; 1) Get Peek window hwnd
    hwnd := WinExist("Peek")
    if (!hwnd)
        hwnd := WinExist("ahk_exe PowerToys.Peek.UI.exe")
    if (!hwnd) {
        StandardLoadingBar_Update("❌ Peek PDF: window not found", BANNER_ACCENT_ERROR)
        StandardLoadingBar_Hide(2000)
        Sleep 300
        Click "Left"
        return
    }
    try {
        ; 2) UIA root
        el := UIA.ElementFromHandle(hwnd)
        ; 3) Poll for Page view (layouts) anchor (up to 20s to accommodate Peek load > 5s)
        ; FindFirst throws when no element matches; catch so we keep polling instead of exiting.
        pageViewBtn := ""
        pollIter := 0
        loop 80 {
            pollIter := A_Index
            try
                pageViewBtn := el.FindFirst({ Type: 50000, Name: "Page view", AutomationId: "layouts" })
            catch
                pageViewBtn := ""
            if (pageViewBtn)
                break
            Sleep 150
        }
        if (pageViewBtn) {
            ; 4) Ensure anchor is visible, then short delay so toolbar is ready before click
            visIter := 0
            loop 20 {
                visIter := A_Index
                try {
                    if (!pageViewBtn.GetPropertyValue(UIA.Property.IsOffscreen)) {
                        br := pageViewBtn.BoundingRectangle
                        if (IsObject(br) && (br.r - br.l) > 0 && (br.b - br.t) > 0)
                            break
                    }
                } catch {
                }
                Sleep 30
            }
            ; Short delay so toolbar is ready before clicking Page view
            Sleep 1000
            ; 5) Click Page view button
            invokeOk := false
            clickOk := false
            try {
                pageViewBtn.Invoke()
                invokeOk := true
            } catch as invErr {
                try {
                    pageViewBtn.Click()
                    clickOk := true
                } catch as clickErr {
                }
            }
            Sleep 600
            ; 6) Select "Two page" from the open Page view menu (main window + foreground popup; else ControlSend Right to Peek)
            fgHwnd := 0
            try fgHwnd := WinGetID("A")

            twoPageEl := ""
            twoPageClicked := false

            twoPageScope := "none"
            menuRect := ""
            try {
                abr := pageViewBtn.BoundingRectangle
                if (IsObject(abr))
                    menuRect := { l: abr.l - 700, t: abr.b, r: abr.r + 700, b: abr.b + 650 }
            } catch {
                menuRect := ""
            }

            ; Search in active window region (menu is visible on screen but may not be exposed as Buttons).
            try {
                elActive := UIA.ElementFromHandle(fgHwnd ? fgHwnd : hwnd)
                if (IsObject(menuRect)) {
                    for cand in elActive.FindAll({ IsOffscreen: 0 }) {
                        try {
                            br := cand.BoundingRectangle
                            if (!IsObject(br))
                                continue
                            inRegion := (br.l < menuRect.r && br.r > menuRect.l && br.t < menuRect.b && br.b > menuRect
                                .t)
                            if (!inRegion)
                                continue
                            nm := ""
                            try nm := cand.Name
                            if (nm != "" && InStr(nm, "Two page")) {
                                twoPageEl := cand
                                twoPageScope := (fgHwnd ? "active_region" : "peek_region")
                                break
                            }
                        } catch {
                        }
                    }
                }
            } catch {
            }

            if (twoPageEl) {
                try {
                    ; Click by coordinates for maximum compatibility (works even if Invoke/Click patterns are absent).
                    br := twoPageEl.BoundingRectangle
                    if (IsObject(br)) {
                        cx := br.l + (br.r - br.l) // 2
                        cy := br.t + (br.b - br.t) // 2
                        Click cx, cy
                        twoPageClicked := true
                    }
                } catch {
                    twoPageClicked := false
                }
            } else {
                ; Last-resort: keystroke fallback (kept for robustness)
                try ControlSend "{Right}", "ahk_id " hwnd
                catch
                    Send "{Right}"
            }

            Sleep 400
            ; 7) Click center of Peek window (focus)
            WinGetPos(&wx, &wy, &ww, &wh, "ahk_id " hwnd)
            if (ww > 0 && wh > 0) {
                cx := wx + ww // 2
                cy := wy + wh // 2
                Click cx, cy
            }
            Sleep 150
            ; 8) Go to final page via UIA (keystrokes do not reach the embedded Edge PDF viewer); skip if skipGoToLastPage
            if (!skipGoToLastPage) {
                lastPageSet := false
                try {
                    for doc in el.FindAll({ Type: 50030 }) {
                        try {
                            nm := doc.Name
                            ; Match "containing N pages" (EN) or "N pages"/"N páginas" (avoid "Page 1")
                            if (RegExMatch(nm, "containing\s+(\d+)\s+pages", &m) || RegExMatch(nm,
                                "document.*?(\d+)\s*(?:pages|páginas)", &m)) {
                                totalPages := Integer(m[1])
                                if (totalPages > 0) {
                                    pageSel := el.FindFirst({ Type: 50004, AutomationId: "pageselector" })
                                    if (pageSel) {
                                        try {
                                            pageSel.SetFocus()
                                            Sleep(200)
                                            WinActivate("ahk_id " hwnd)
                                            Sleep(120)
                                            Send("^a")
                                            Sleep(50)
                                            Send(String(totalPages))
                                            Sleep(50)
                                            Send("{Enter}")
                                            lastPageSet := true
                                        } catch {
                                            ; ignore focus/send errors
                                        }
                                    }
                                    break
                                }
                            }
                        } catch {
                            ; ignore per-doc errors
                        }
                    }
                } catch {
                    ; ignore UIA errors for last-page navigation
                }
                if (!lastPageSet) {
                    try
                        ControlSend("^End", "ahk_id " hwnd)
                    catch
                        Send("^End")
                }
            }
            Sleep 100
            StandardLoadingBar_Update("✅ Peek PDF: done", BANNER_ACCENT_SUCCESS)
            StandardLoadingBar_Hide(2000)
        } else {
            StandardLoadingBar_Update("✅ Peek PDF: finished (fallback)", BANNER_ACCENT_SUCCESS)
            StandardLoadingBar_Hide(2000)
            Sleep 400
            Click "Left"
        }
    } catch {
        StandardLoadingBar_Update("✅ Peek PDF: finished (fallback)", BANNER_ACCENT_SUCCESS)
        StandardLoadingBar_Hide(2000)
        Sleep 400
        Click "Left"
    }
}

#!+x::
{
    hwnd := WinExist("ahk_exe QuickLook.exe")
    if hwnd {
        if (STUDY_TOPIC_QL_STRICT_LAYOUT) {
            QuickLook_ApplyStudyLayout(hwnd, true, 0)
        } else {
            try {
                WinShow("ahk_id " hwnd)
                WinActivate("ahk_id " hwnd)
                WinWaitActive("ahk_id " hwnd, , 1)
            } catch {
            }
            QuickLook_ClickWindowCenter(hwnd)
            try
                ControlSend("^End", "ahk_id " hwnd)
            catch {
                try {
                    WinActivate("ahk_id " hwnd)
                    WinWaitActive("ahk_id " hwnd, , 1)
                } catch {
                }
                try Send("^End")
            }
        }
        StudyTopic_StartBlackoutCountdown(hwnd)
        return
    }
    ShowStudyTopicSelector()
}

; =============================================================================
; Hotstring Selector System
; =============================================================================
; PURPOSE: Provides a modal GUI-based interface for accessing hotstrings, quick-open files,
;          and executable macros via single-character keyboard shortcuts.
;
; HOTKEY: Windows + Alt + Shift + U (#!+U)
;
; FUNCTIONALITY:
;   - Displays categorized list of available actions (Prompts, General, Projects, Files & Links, Macros)
;   - Each action is assigned a unique character from g_HotstringCharSequence
;   - User presses assigned character to execute corresponding action
;   - Actions include: text expansion (hotstrings), file opening, macro execution
;
; ARCHITECTURE:
;   - Character-to-action mapping built dynamically via BuildHotstringCharMap()
;   - Character assignments follow sequential order within each category
;   - Explicit character assignments (via RegisterMacro/RegisterHotstring char parameter) take precedence
;   - GUI adapts to monitor configuration (landscape/portrait, resolution, scaling)
;
; =============================================================================

; Global state variables for hotstring selector system
global g_HotstringSelectorGui := false          ; GUI object reference (false when not initialized)
global g_HotstringSelectorActive := false       ; Boolean flag indicating selector is currently displayed
global g_HotstringCharMap := Map()              ; Character-to-text-expansion mapping for hotstrings
global g_UtilityHotstringCharMapByCategory := Map() ; Category -> Map(char -> expansion) used by Utility Shortcuts
global g_HotstringHotkeyHandlers := []          ; Array of hotkey handler objects for cleanup on close
global g_HotstringPromptCharMap := Map()        ; Map of prompt-assigned chars => true (rebuilt on each ShowHotstringSelector)
global g_HotstringGeminiArmed := false          ; When true, next Prompts selection is redirected to Gemini
global g_HotstringGeminiAutoSubmit := true      ; During delayed flow: true = send Enter after paste; false = paste only
global g_HotstringGeminiSubmitTimer := false   ; Timer reference for 4s delayed submit (for cleanup if needed)
global g_HotstringGeminiRestoreHwnd := 0        ; Window to restore focus to after paste (set at start of GeminiDelayedSubmitFlow)

; Utility selector hierarchy state
global g_UtilitySelectorMode := "top"           ; "top" | "category"
global g_UtilitySelectorCategory := ""          ; One of g_UtilityTopCategories

; Top-level categories (numbers 1-6 select these)
global g_UtilityTopCategories := ["Prompts", "Projects", "Macros", "General", "Links", "Hotstrings"]
; Top-level trigger keys (lowercase so UtilitySelector_RebindHotkeys auto-binds uppercase too)
global g_UtilityTopCategoryById := Map("r", "Prompts", "p", "Projects", "m", "Macros", "g", "General", "l", "Links",
    "h", "Hotstrings", "c", "Context")

; Utility selector cached UI data (rebuilt each time ShowHotstringSelector() runs)
global g_UtilitySelectorAllItems := []          ; Array of {category, char, text, isEmpty, [explicitIndex]}
global g_UtilitySelectorIsPortrait := false
global g_UtilitySelectorMonitor := Map()        ; {left, top, width, height}
global g_UtilitySelectorTitleCtrl := false
global g_UtilitySelectorEditCtrl := false
global g_UtilitySelectorFontSize := 9           ; Base point size for RichEdit rendering (set on ShowHotstringSelector)
global g_UtilitySelectorFooterCtrl := false     ; Footer ctrl ref; repositioned on each content refresh

; Character assignment sequence: defines order in which characters are assigned to actions
; Format: ["1", "2", "3", "4", "5", "q", "w", "e", "r", "t", "a", "s", "d", "f", "g", "z", "x",
;          "c", "v", "b", "6", "7", "8", "9", "0", "y", "u", "i", "o", "p", "h", "j", "k", "l", "n", "m", ",", "."]
global g_HotstringCharSequence := ["1", "2", "3", "4", "5", "q", "w", "e", "r", "t", "a", "s", "d", "f", "g", "z", "x",
    "c", "v", "b", "6", "7", "8", "9", "0", "y", "u", "i", "o", "p", "h", "j", "k", "l", "n", "m", ",", "."]

; Category display order: defines the sequence in which action categories appear in the GUI
; Order: Prompts → General → Projects → Links → Macros → Hotstrings
; Note: Utility-only views are not included here.
; "Hotstrings" must be present so BuildHotstringCharMap() populates g_UtilityHotstringCharMapByCategory["Hotstrings"].
global g_HotstringCategories := ["Prompts", "General", "Projects", "Files & Links", "Macros", "Hotstrings"]

; Reserved empty character: never assigned to any action; always shows as (empty) in selector
; Set to "" to disable reservation
global g_ReservedEmptyChar := ""

; =============================================================================
; RichEdit helpers (mnemonic emphasis for selectors)
; =============================================================================
global g_MnemonicRichDll := 0

MnemonicRich_EnsureDll() {
    global g_MnemonicRichDll
    ; msftedit.dll must be loaded before creating RichEdit50W controls (ClassRichEdit50W).
    if (!g_MnemonicRichDll)
        g_MnemonicRichDll := DllCall("LoadLibrary", "str", "msftedit.dll", "ptr")
}

; UTF-16 code unit count for RichEdit character indices (BMP = 1, supplementary = 2).
MnemonicRich_Utf16Units(s) {
    n := 0
    for c in StrSplit(s, "") {
        o := Ord(c)
        n += (o > 0xFFFF) ? 2 : 1
    }
    return n
}

; EM_SETTEXTEX = 0x461, ST_UNICODE = 8 - RichEdit's native UTF-16 path.
MnemonicRich_SetPlainUtf16(ctrl, plain) {
    hwnd := ctrl.Hwnd
    flags := 8 ; ST_UNICODE
    cp := 1200
    settextex := Buffer(8, 0)
    NumPut("uint", flags, settextex, 0)
    NumPut("uint", cp, settextex, 4)
    if (plain = "") {
        emptyBuf := Buffer(2, 0)
        SendMessage(0x461, settextex.Ptr, emptyBuf.Ptr, hwnd)
        return
    }
    textBuf := Buffer((StrLen(plain) + 1) * 2)
    StrPut(plain, textBuf, "UTF-16")
    SendMessage(0x461, settextex.Ptr, textBuf.Ptr, hwnd)
}

MnemonicRich_ThemingOff(ctrl) {
    hwnd := ctrl.Hwnd
    DllCall("uxtheme\SetWindowTheme", "ptr", hwnd, "wstr", "", "wstr", "")
    parent := DllCall("GetParent", "ptr", hwnd, "ptr")
    if (parent)
        DllCall("uxtheme\SetWindowTheme", "ptr", parent, "wstr", "", "wstr", "")
}

; CHARFORMAT2W buffer (116 bytes). textColor is COLORREF (0x00BBGGRR).
MnemonicRich_CharFormat2(faceName, pt, textColor, bold := false) {
    yh := Round(pt * 20)
    cf := Buffer(116, 0)
    NumPut("uint", 116, cf, 0) ; cbSize
    mask := 0x40000000 | 0x80000000 | 0x20000000 | 0x1 ; CFM_COLOR | CFM_SIZE | CFM_FACE | CFM_BOLD
    NumPut("uint", mask, cf, 4) ; dwMask
    NumPut("uint", bold ? 0x1 : 0, cf, 8) ; dwEffects
    NumPut("int", yh, cf, 12) ; yHeight (twips)
    NumPut("int", 0, cf, 16) ; yOffset
    NumPut("uint", textColor, cf, 20) ; crTextColor
    NumPut("uchar", 1, cf, 24) ; bCharSet DEFAULT_CHARSET
    NumPut("uchar", 0, cf, 25) ; bPitchAndFamily
    StrPut(faceName, cf.Ptr + 26, 64, "UTF-16")
    return cf
}

MnemonicRich_ApplyCharFormat(ctrl, scopeAll, cfBuf) {
    w := scopeAll ? 4 : 1 ; SCF_ALL = 4, SCF_SELECTION = 1
    return SendMessage(0x444, w, cfBuf.Ptr, ctrl.Hwnd) ; EM_SETCHARFORMAT
}

MnemonicRich_SetSel(hwnd, cpMin, cpMax) {
    return SendMessage(0xB1, cpMin, cpMax, hwnd) ; EM_SETSEL
}

; Render lines (joined by CR only) and emphasize mnemonic letter (bumpPx) in [key] and first title occurrence.
MnemonicRich_Render(ctrl, lines, basePt, bumpPx := 6, faceName := "Segoe UI", rgbHex := "CDD6F4", bgHex := "1E1E2E") {
    MnemonicRich_EnsureDll()
    MnemonicRich_ThemingOff(ctrl)

    ; Convert RGB hex (RRGGBB) to COLORREF (0x00BBGGRR).
    rr := Integer("0x" . SubStr(rgbHex, 1, 2))
    gg := Integer("0x" . SubStr(rgbHex, 3, 2))
    bb := Integer("0x" . SubStr(rgbHex, 5, 2))
    textColor := (bb << 16) | (gg << 8) | rr

    br := Integer("0x" . SubStr(bgHex, 1, 2))
    bg := Integer("0x" . SubStr(bgHex, 3, 2))
    bb2 := Integer("0x" . SubStr(bgHex, 5, 2))
    bgColor := (bb2 << 16) | (bg << 8) | br

    bumpPt := bumpPx * 72 / 96
    bigPt := basePt + bumpPt

    plain := ""
    spans := [] ; {start,len} in UTF-16 units
    subsectionSpans := [] ; {start,len} for mnemonic subsection lines (distinct color)
    u16Pos := 0
    first := true

    RenderTitleKey(lineText, key, baseU16) {
        if (key = "")
            return
        rb := InStr(lineText, "]")
        if (!rb)
            return
        after := SubStr(lineText, rb + 1)
        tpos := InStr(after, key, false)
        if (!tpos)
            tpos := InStr(after, StrUpper(key), false)
        if (!tpos)
            return
        preNoLast := SubStr(lineText, 1, rb + tpos - 1)
        spans.Push({ start: baseU16 + MnemonicRich_Utf16Units(preNoLast), len: 1 })
    }

    for ln in lines {
        if (!first) {
            plain .= "`r"
            u16Pos += 1
        }
        first := false

        lineText := ln.text
        lineStartU16 := u16Pos
        key := ln.HasProp("key") ? ln.key : ""
        RenderTitleKey(lineText, key, u16Pos)

        ; Optional right-side key emphasis for two-column layouts.
        if (ln.HasProp("keyRight") && ln.keyRight != "" && ln.HasProp("rightStartCharPos") && ln.rightStartCharPos > 1) {
            rightStart := ln.rightStartCharPos
            prefix := SubStr(lineText, 1, rightStart - 1)
            rightText := SubStr(lineText, rightStart)
            baseRightU16 := u16Pos + MnemonicRich_Utf16Units(prefix)
            RenderTitleKey(rightText, ln.keyRight, baseRightU16)
        }

        plain .= lineText
        u16Pos += MnemonicRich_Utf16Units(lineText)
        if (ln.HasProp("isMnemonicSubsection") && ln.isMnemonicSubsection && lineText != "") {
            subsectionSpans.Push({ start: lineStartU16, len: MnemonicRich_Utf16Units(lineText) })
        }
    }

    MnemonicRich_SetPlainUtf16(ctrl, plain)

    hwnd := ctrl.Hwnd
    SendMessage(0x4CF, 0, 0, hwnd) ; EM_SETREADONLY FALSE while formatting
    SendMessage(0x443, 0, bgColor, hwnd) ; EM_SETBKGNDCOLOR

    baseCf := MnemonicRich_CharFormat2(faceName, basePt, textColor, false)
    MnemonicRich_SetSel(hwnd, 0, -1)
    MnemonicRich_ApplyCharFormat(ctrl, false, baseCf)

    ; Subsection headers inside Prompts (mnemonic technique): softer accent, bold
    if (subsectionSpans.Length > 0) {
        subRgb := "CBA6F7" ; mauve, distinct from body
        srr := Integer("0x" . SubStr(subRgb, 1, 2))
        sgg := Integer("0x" . SubStr(subRgb, 3, 2))
        sbb := Integer("0x" . SubStr(subRgb, 5, 2))
        subColor := (sbb << 16) | (sgg << 8) | srr
        subCf := MnemonicRich_CharFormat2(faceName, basePt + 1, subColor, true)
        for ss in subsectionSpans {
            if (ss.len <= 0)
                continue
            MnemonicRich_SetSel(hwnd, ss.start, ss.start + ss.len)
            MnemonicRich_ApplyCharFormat(ctrl, false, subCf)
        }
    }

    bigCf := MnemonicRich_CharFormat2(faceName, bigPt, textColor, false)
    for sp in spans {
        if (sp.len <= 0)
            continue
        MnemonicRich_SetSel(hwnd, sp.start, sp.start + sp.len)
        MnemonicRich_ApplyCharFormat(ctrl, false, bigCf)
    }
    MnemonicRich_SetSel(hwnd, 0, 0)
    SendMessage(0xB7, 0, 0, hwnd) ; EM_SCROLLCARET
    SendMessage(0x4CF, 1, 0, hwnd) ; EM_SETREADONLY TRUE
    SendMessage(0x443, 0, bgColor, hwnd) ; RichEdit can reset bg after readonly
}

; =============================================================================
; BuildHotstringCharMap()
; =============================================================================
; PURPOSE: Constructs character-to-action mappings for all registered items (hotstrings, files, macros)
;          and assigns characters sequentially within each category according to g_HotstringCharSequence.
;
; PROCESS:
;   1. Groups hotstrings by category (Projects, Prompts, General)
;   2. Processes each category in g_HotstringCategories order:
;      - Files & Links: Maps characters to file paths for quick-open functionality
;      - Macros: Maps characters to executable macro functions (explicit assignments first, then sequential)
;      - Other categories: Maps characters to hotstring expansion text
;   3. Returns Map of character → expansion text for hotstrings
;
; RETURNS: Map object where keys are characters and values are expansion text strings
; SIDE EFFECTS: Populates global maps g_QuickOpenFileCharMap and g_MacroCharMap
; =============================================================================
BuildHotstringCharMap() {
    global g_hotstrings, g_QuickOpenFiles, g_HotstringCategories, g_Macros
    charMap := Map()
    global g_QuickOpenFileCharMap := Map()
    global g_MacroCharMap := Map()
    global g_UtilityHotstringCharMapByCategory

    ; Category-scoped hotstring maps used by Utility Shortcuts selector.
    ; This allows the same char to exist in multiple categories (e.g. Prompts 'a' and Projects 'a').
    g_UtilityHotstringCharMapByCategory := Map()
    g_UtilityHotstringCharMapByCategory["Prompts"] := Map()
    g_UtilityHotstringCharMapByCategory["Projects"] := Map()
    g_UtilityHotstringCharMapByCategory["General"] := Map()
    g_UtilityHotstringCharMapByCategory["Hotstrings"] := Map()

    ; Group hotstrings by category
    categorized := Map()
    categorized["Projects"] := []
    categorized["Prompts"] := []
    categorized["General"] := []
    categorized["Hotstrings"] := []

    if (IsSet(g_hotstrings) && g_hotstrings.Length > 0) {
        for hs in g_hotstrings {
            category := hs.category
            if (category = "Projects" || category = "Prompts" || category = "General" || category = "Hotstrings") {
                categorized[category].Push(hs)
            } else {
                categorized["General"].Push(hs)
            }
        }
    }

    ; Assign characters sequentially within each category
    charIndex := 1
    for category in g_HotstringCategories {
        if (category = "Files & Links") {
            ; Handle quick open files
            if (IsSet(g_QuickOpenFiles) && g_QuickOpenFiles.Length > 0) {
                for fileEntry in g_QuickOpenFiles {
                    while (charIndex <= g_HotstringCharSequence.Length && g_ReservedEmptyChar != "" &&
                        g_HotstringCharSequence[charIndex] = g_ReservedEmptyChar)
                        charIndex++
                    if (charIndex <= g_HotstringCharSequence.Length) {
                        char := g_HotstringCharSequence[charIndex]
                        g_QuickOpenFileCharMap[char] := fileEntry.filePath
                        charIndex++
                    }
                }
            }
        } else if (category = "Macros") {
            ; Handle macros
            if (IsSet(g_Macros) && g_Macros.Length > 0) {
                ; First pass: assign macros with explicit character assignments
                for macroEntry in g_Macros {
                    if (macroEntry.HasProp("char") && macroEntry.char != "" && (g_ReservedEmptyChar = "" || macroEntry.char !=
                        g_ReservedEmptyChar)) {
                        ; Check if character is in the sequence and not already assigned
                        charIndexInSequence := 0
                        for idx, seqChar in g_HotstringCharSequence {
                            if (seqChar = macroEntry.char) {
                                charIndexInSequence := idx
                                break
                            }
                        }
                        if (charIndexInSequence > 0) {
                            ; Check if this character is already assigned
                            if (!g_MacroCharMap.Has(macroEntry.char)) {
                                g_MacroCharMap[macroEntry.char] := macroEntry.func
                            }
                        }
                    }
                }
                ; Second pass: assign remaining macros sequentially, skipping already assigned characters
                for macroEntry in g_Macros {
                    ; Skip if this macro already has a character assigned
                    alreadyAssigned := false
                    for assignedChar, assignedFunc in g_MacroCharMap {
                        if (assignedFunc = macroEntry.func) {
                            alreadyAssigned := true
                            break
                        }
                    }
                    if (alreadyAssigned) {
                        continue
                    }

                    ; Find next available character (skip reserved empty char if set)
                    while (charIndex <= g_HotstringCharSequence.Length) {
                        char := g_HotstringCharSequence[charIndex]
                        if (g_ReservedEmptyChar != "" && char = g_ReservedEmptyChar) {
                            charIndex++
                            continue
                        }
                        ; Check if this character is already assigned to a macro
                        if (!g_MacroCharMap.Has(char)) {
                            g_MacroCharMap[char] := macroEntry.func
                            charIndex++
                            break
                        }
                        charIndex++
                    }
                }
            }
        } else {
            ; Handle hotstring categories
            if (categorized.Has(category)) {
                ; Utility Shortcuts: assign within-category (independent) to avoid cross-category collisions.
                utilCharIndex := 1
                utilTaken := Map()

                ; Explicit assignments first
                for hs in categorized[category] {
                    if (hs.HasProp("char") && hs.char != "" && (g_ReservedEmptyChar = "" || hs.char !=
                        g_ReservedEmptyChar)) {
                        if (hs.expansion != "" && !utilTaken.Has(hs.char)) {
                            g_UtilityHotstringCharMapByCategory[category][hs.char] := hs.expansion
                            utilTaken[hs.char] := true
                        }
                    }
                }

                ; Sequential assignments for remaining hotstrings in this category
                for hs in categorized[category] {
                    alreadyAssigned := false
                    for assignedChar, assignedExpansion in g_UtilityHotstringCharMapByCategory[category] {
                        if (assignedExpansion = hs.expansion) {
                            alreadyAssigned := true
                            break
                        }
                    }
                    if (alreadyAssigned)
                        continue

                    while (utilCharIndex <= g_HotstringCharSequence.Length) {
                        ch := g_HotstringCharSequence[utilCharIndex]
                        utilCharIndex++
                        if (g_ReservedEmptyChar != "" && ch = g_ReservedEmptyChar)
                            continue
                        if (!utilTaken.Has(ch)) {
                            if (hs.expansion != "") {
                                g_UtilityHotstringCharMapByCategory[category][ch] := hs.expansion
                                utilTaken[ch] := true
                            }
                            break
                        }
                    }
                }

                ; First pass: assign hotstrings with explicit character assignments
                for hs in categorized[category] {
                    if (hs.HasProp("char") && hs.char != "" && (g_ReservedEmptyChar = "" || hs.char !=
                        g_ReservedEmptyChar)) {
                        ; Check if character is in the sequence and not already assigned
                        charIndexInSequence := 0
                        for idx, seqChar in g_HotstringCharSequence {
                            if (seqChar = hs.char) {
                                charIndexInSequence := idx
                                break
                            }
                        }
                        if (charIndexInSequence > 0) {
                            ; Check if this character is already assigned
                            if (!charMap.Has(hs.char) && hs.expansion != "") {
                                charMap[hs.char] := hs.expansion
                            }
                        }
                    }
                }
                ; Second pass: assign remaining hotstrings sequentially, skipping already assigned characters
                for hs in categorized[category] {
                    ; Skip if this hotstring already has a character assigned
                    alreadyAssigned := false
                    for assignedChar, assignedExpansion in charMap {
                        if (assignedExpansion = hs.expansion) {
                            alreadyAssigned := true
                            break
                        }
                    }
                    if (alreadyAssigned) {
                        continue
                    }

                    ; Find next available character (skip reserved empty char if set)
                    while (charIndex <= g_HotstringCharSequence.Length) {
                        char := g_HotstringCharSequence[charIndex]
                        if (g_ReservedEmptyChar != "" && char = g_ReservedEmptyChar) {
                            charIndex++
                            continue
                        }
                        ; Check if this character is already assigned
                        if (!charMap.Has(char)) {
                            ; Only assign characters to hotstrings that have an expansion (skip empty placeholders)
                            if (hs.expansion != "") {
                                charMap[char] := hs.expansion
                            }
                            charIndex++
                            break
                        }
                        charIndex++
                    }
                }
            }
        }
    }

    return charMap
}

; =============================================================================
; GetCategorizedHotstrings()
; =============================================================================
; PURPOSE: Organizes all registered items (hotstrings, quick-open files, macros) into category-based
;          data structure for GUI display purposes.
;
; PROCESS:
;   1. Initializes empty arrays for each category: Projects, Prompts, General, Files & Links, Macros
;   2. Groups hotstrings by their category property
;   3. Adds quick-open file entries to "Files & Links" category
;   4. Adds macro entries to "Macros" category
;
; RETURNS: Map object where keys are category names and values are arrays of item objects
;          Each item object contains: trigger, expansion, title, category, and optionally char
; =============================================================================
GetCategorizedHotstrings() {
    global g_hotstrings, g_QuickOpenFiles, g_Macros
    categorized := Map()
    categorized["Projects"] := []
    categorized["Prompts"] := []
    categorized["General"] := []
    categorized["Hotstrings"] := []
    categorized["Links"] := []
    categorized["Files & Links"] := []
    categorized["Macros"] := []

    ; Add hotstrings
    if (IsSet(g_hotstrings) && g_hotstrings.Length > 0) {
        for hs in g_hotstrings {
            category := hs.category
            if (category = "Projects" || category = "Prompts" || category = "General" || category = "Hotstrings") {
                categorized[category].Push(hs)
            } else {
                categorized["General"].Push(hs)
            }
        }
    }

    ; Add quick open files (rendered under Links in Utility selector)
    if (IsSet(g_QuickOpenFiles) && g_QuickOpenFiles.Length > 0) {
        for fileEntry in g_QuickOpenFiles {
            categorized["Files & Links"].Push(fileEntry)
        }
    }

    ; Add macros
    if (IsSet(g_Macros) && g_Macros.Length > 0) {
        for macroEntry in g_Macros {
            categorized["Macros"].Push(macroEntry)
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

; Find and activate Power BI file, or open it if not already open
FindAndActivatePowerBIFile(filePath) {
    ; Check if file exists
    if (!FileExist(filePath)) {
        return false
    }

    ; Extract filename from path (Power BI window titles don't include .pbix extension)
    SplitPath(filePath, , , , &fileNameNoExt)

    ; Normalize the filename for comparison (trim whitespace, case-insensitive)
    fileNameNoExt := Trim(fileNameNoExt)
    fileNameLower := StrLower(fileNameNoExt)

    ; Search for Power BI Desktop windows with matching filename
    ; Power BI window titles are like "Dissertation InfoVis  - PowerBI - Charts" (no extension)
    try {
        for hwnd in WinGetList("ahk_exe PBIDesktop.exe") {
            try {
                winTitle := WinGetTitle("ahk_id " hwnd)
                winTitleLower := StrLower(Trim(winTitle))

                ; Check if window title matches the filename (case-insensitive)
                ; Power BI window title should be exactly the filename or start with it
                ; Check for exact match first, then check if title starts with filename
                if (winTitleLower = fileNameLower || InStr(winTitleLower, fileNameLower) = 1) {
                    ; Found matching window, activate it
                    WinActivate("ahk_id " hwnd)
                    WinWaitActive("ahk_id " hwnd, , 2)
                    return true
                }
            } catch {
                ; Skip windows we can't access
                continue
            }
        }
    } catch {
        ; No Power BI windows found or error accessing them
    }

    ; No matching window found, open the file
    try {
        Run(filePath)
        return true
    } catch Error as e {
        ; Failed to open file
        return false
    }
}

; Find and activate Miro window by title keywords and URL
; Returns true if window was found and activated, or if URL was opened successfully
FindAndActivateMiroWindow(url, titleKeywords) {
    ; Normalize title keywords for matching (case-insensitive)
    keywordsLower := StrLower(titleKeywords)

    ; Search for Chrome windows with Miro in title
    try {
        for hwnd in WinGetList("ahk_exe chrome.exe") {
            try {
                winTitle := WinGetTitle("ahk_id " hwnd)
                winTitleLower := StrLower(Trim(winTitle))

                ; Check if window is a Miro window and contains the keywords
                if (InStr(winTitleLower, "miro") && InStr(winTitleLower, keywordsLower)) {
                    ; Found matching window, activate it and bring to front
                    ; Use a separate try-catch for activation to ensure we return even if activation fails
                    try {
                        ; Ensure window is not minimized first
                        if (WinGetMinMax("ahk_id " hwnd) = -1) {
                            WinRestore("ahk_id " hwnd)
                        }

                        ; Activate the window
                        WinActivate("ahk_id " hwnd)
                        WinWaitActive("ahk_id " hwnd, , 2)

                        ; Bring to front using AlwaysOnTop trick to ensure it's not hidden
                        WinSetAlwaysOnTop("On", "ahk_id " hwnd)
                        Sleep 50
                        WinSetAlwaysOnTop("Off", "ahk_id " hwnd)
                    } catch Error as activateErr {
                        ; Even if activation fails, we found the window, so return to prevent opening a new one
                    }

                    ; Return immediately after activating - don't continue searching or open new window
                    return true
                }
            } catch Error as e {
                ; Skip windows we can't access
                continue
            }
        }
    } catch {
        ; No Chrome windows found or error accessing them
    }

    ; No matching window found, open the URL
    try {
        ; Open URL in Chrome
        Run("chrome.exe --new-window " . url)

        ; Wait for window to appear and become active
        ; Wait up to 10 seconds for the window to appear
        WinWait("ahk_exe chrome.exe", , 10)

        ; Find the newly opened window by checking for Miro in title
        ; Give it a moment to load
        Sleep(1000)

        ; Try to find the window with Miro in title
        loop 10 {
            for hwnd in WinGetList("ahk_exe chrome.exe") {
                try {
                    winTitle := WinGetTitle("ahk_id " hwnd)
                    winTitleLower := StrLower(Trim(winTitle))

                    if (InStr(winTitleLower, "miro") && InStr(winTitleLower, keywordsLower)) {
                        ; Found the window, activate it
                        ; Ensure window is not minimized first
                        if (WinGetMinMax("ahk_id " hwnd) = -1) {
                            WinRestore("ahk_id " hwnd)
                        }
                        ; Activate the window
                        WinActivate("ahk_id " hwnd)
                        WinWaitActive("ahk_id " hwnd, , 2)
                        ; Bring to front using AlwaysOnTop trick to ensure it's not hidden
                        WinSetAlwaysOnTop("On", "ahk_id " hwnd)
                        Sleep 50
                        WinSetAlwaysOnTop("Off", "ahk_id " hwnd)
                        ; Additional check: ensure window is actually active
                        if (WinActive("ahk_id " hwnd)) {
                            return true
                        }
                    }
                } catch {
                    continue
                }
            }
            Sleep(500)  ; Wait before next attempt
        }

        ; If we couldn't find by title, just activate the most recent Chrome window
        ; This is a fallback in case the title hasn't updated yet
        try {
            chromeWindows := WinGetList("ahk_exe chrome.exe")
            if (chromeWindows.Length > 0) {
                ; Get the first (most recent) Chrome window
                hwnd := chromeWindows[1]

                ; Ensure window is not minimized
                if (WinGetMinMax("ahk_id " hwnd) = -1) {
                    WinRestore("ahk_id " hwnd)
                }
                ; Activate the window
                WinActivate("ahk_id " hwnd)
                WinWaitActive("ahk_id " hwnd, , 2)
                ; Bring to front
                WinSetAlwaysOnTop("On", "ahk_id " hwnd)
                Sleep 50
                WinSetAlwaysOnTop("Off", "ahk_id " hwnd)
                return true
            }
        } catch {
        }

        return true  ; Assume success if we got this far
    } catch Error as e {
        ; Failed to open URL
        return false
    }
}

; =============================================================================
; Context file browser — browse context/ and paste full local paths (Win+Alt+Shift+N)
; =============================================================================
ContextBrowser_EnsureGlobals() {
    global CONTEXT_ROOT, g_ContextBrowserActive, g_ContextBrowserGui, g_ContextBrowserCurrentDir
    global g_ContextBrowserEntries, g_ContextBrowserListView, g_ContextBrowserPathCtrl
    if !IsSet(CONTEXT_ROOT)
        CONTEXT_ROOT := A_ScriptDir "\context"
    if !IsSet(g_ContextBrowserActive)
        g_ContextBrowserActive := false
    if !IsSet(g_ContextBrowserGui)
        g_ContextBrowserGui := false
    if !IsSet(g_ContextBrowserCurrentDir)
        g_ContextBrowserCurrentDir := ""
    if !IsSet(g_ContextBrowserEntries)
        g_ContextBrowserEntries := []
    if !IsSet(g_ContextBrowserListView)
        g_ContextBrowserListView := false
    if !IsSet(g_ContextBrowserPathCtrl)
        g_ContextBrowserPathCtrl := false
}
ContextBrowser_EnsureGlobals()

Context_GetRoot() {
    ContextBrowser_EnsureGlobals()
    global CONTEXT_ROOT
    return CONTEXT_ROOT
}

Context_IsAtRoot(dir) {
    root := Context_GetRoot()
    return (StrLower(RTrim(dir, "\")) = StrLower(RTrim(root, "\")))
}

Context_IsExistingFile(path) {
    if (path = "")
        return false
    attr := FileExist(path)
    return (attr && !InStr(attr, "D"))
}

Context_ReadFirstNonemptyLine(path) {
    try {
        content := FileRead(path)
    } catch {
        return ""
    }
    for line in StrSplit(content, "`n", "`r") {
        trimmed := Trim(line)
        if (trimmed != "")
            return trimmed
    }
    return ""
}

Context_IsAbsoluteFilePath(line) {
    return (line != "" && RegExMatch(line, "^[A-Za-z]:\\"))
}

Context_ProbeReference(path) {
    result := { isRef: false, targetPath: "", targetBasename: "" }
    if !Context_IsExistingFile(path)
        return result
    SplitPath path, , , &ext
    ext := StrLower(ext)
    if (ext = "lnk") {
        try {
            target := ComObject("WScript.Shell").CreateShortcut(path).TargetPath
            if (target != "" && Context_IsExistingFile(target)) {
                result.isRef := true
                result.targetPath := target
                SplitPath target, &base
                result.targetBasename := base
            }
        } catch {
        }
        return result
    }
    line := Context_ReadFirstNonemptyLine(path)
    if !Context_IsAbsoluteFilePath(line)
        return result
    if Context_IsExistingFile(line) {
        result.isRef := true
        result.targetPath := line
        SplitPath line, &base
        result.targetBasename := base
    }
    return result
}

; Returns path to paste, or "" if a pointer/shortcut reference is broken (caller shows error).
Context_ResolvePastePath(path) {
    if (path = "")
        return ""
    SplitPath path, , , &ext
    ext := StrLower(ext)
    if (ext = "lnk") {
        try {
            target := ComObject("WScript.Shell").CreateShortcut(path).TargetPath
            if (target != "" && Context_IsExistingFile(target))
                return target
        } catch {
        }
        return ""
    }
    line := Context_ReadFirstNonemptyLine(path)
    if Context_IsAbsoluteFilePath(line) {
        if Context_IsExistingFile(line)
            return line
        return ""
    }
    return path
}

Context_SortNames(names) {
    if (names.Length < 2)
        return names
    list := ""
    for n in names
        list .= n "`n"
    list := Sort(RTrim(list, "`n"), "P`n")
    sorted := []
    if (list != "") {
        for line in StrSplit(list, "`n")
            sorted.Push(line)
    }
    return sorted
}

Context_ListDirEntries(dir) {
    folders := []
    files := []
    try {
        Loop Files, dir "\*", "D" {
            if (A_LoopFileAttrib ~= "[HS]")
                continue
            folders.Push(A_LoopFileName)
        }
        Loop Files, dir "\*", "F" {
            if (A_LoopFileAttrib ~= "[HS]")
                continue
            files.Push(A_LoopFileName)
        }
    } catch {
    }
    folders := Context_SortNames(folders)
    files := Context_SortNames(files)
    entries := []
    for name in folders
        entries.Push({ type: "folder", name: name, path: dir "\" name })
    for name in files
        entries.Push({ type: "file", name: name, path: dir "\" name })
    return entries
}

ContextBrowser_BuildViewEntries(dir) {
    entries := []
    if !Context_IsAtRoot(dir)
        entries.Push({ type: "parent", name: "..", path: "" })
    for entry in Context_ListDirEntries(dir)
        entries.Push(entry)
    return entries
}

Context_GetRelativeSubtitle(dir) {
    root := Context_GetRoot()
    rootNorm := RTrim(root, "\")
    dirNorm := RTrim(dir, "\")
    if (StrLower(dirNorm) = StrLower(rootNorm))
        return rootNorm
    rel := SubStr(dirNorm, StrLen(rootNorm) + 2)
    return rootNorm "  »  " StrReplace(rel, "\", " » ")
}

ContextBrowser_FormatEntryLabel(entry) {
    if (entry.type = "folder")
        return entry.name
    probe := Context_ProbeReference(entry.path)
    if (probe.isRef && probe.targetBasename != "")
        return entry.name "  →  " probe.targetBasename
    return entry.name
}

ContextBrowser_EntryKindLabel(entry) {
    if (entry.type = "parent")
        return "Up"
    if (entry.type = "folder")
        return "Folder"
    return "File"
}

ContextBrowser_GetActiveMonitorWorkArea(&left, &top, &right, &bottom) {
    MonitorGetWorkArea(1, &left, &top, &right, &bottom)
    activeWin := 0
    try activeWin := WinGetID("A")
    catch {
        return
    }
    if !activeWin
        return
    rect := Buffer(16, 0)
    if !DllCall("GetWindowRect", "ptr", activeWin, "ptr", rect)
        return
    centerX := NumGet(rect, 0, "int") + (NumGet(rect, 8, "int") - NumGet(rect, 0, "int")) // 2
    centerY := NumGet(rect, 4, "int") + (NumGet(rect, 12, "int") - NumGet(rect, 4, "int")) // 2
    loop MonitorGetCount() {
        MonitorGetWorkArea(A_Index, &l, &t, &r, &b)
        if (centerX >= l && centerX <= r && centerY >= t && centerY <= b) {
            left := l
            top := t
            right := r
            bottom := b
            return
        }
    }
}

CleanupContextBrowser() {
    ContextBrowser_EnsureGlobals()
    global g_ContextBrowserActive, g_ContextBrowserGui, g_ContextBrowserCurrentDir
    global g_ContextBrowserEntries, g_ContextBrowserListView, g_ContextBrowserPathCtrl

    g_ContextBrowserActive := false
    try Hotkey("Backspace", "Off")
    try Hotkey("Enter", "Off")
    catch {
    }
    g_ContextBrowserEntries := []
    g_ContextBrowserCurrentDir := ""
    g_ContextBrowserListView := false
    g_ContextBrowserPathCtrl := false
    if (IsObject(g_ContextBrowserGui)) {
        try g_ContextBrowserGui.Destroy()
        catch {
        }
        g_ContextBrowserGui := false
    }
}

HandleContextBrowserEscape(*) {
    ContextBrowser_EnsureGlobals()
    global g_ContextBrowserActive
    if (g_ContextBrowserActive)
        CleanupContextBrowser()
}

ContextBrowser_HandleBack() {
    ContextBrowser_EnsureGlobals()
    global g_ContextBrowserActive, g_ContextBrowserCurrentDir
    if (!g_ContextBrowserActive)
        return
    if Context_IsAtRoot(g_ContextBrowserCurrentDir)
        return
    SplitPath g_ContextBrowserCurrentDir, , &parentDir
    if (parentDir = "" || !DirExist(parentDir))
        return
    g_ContextBrowserCurrentDir := parentDir
    ContextBrowser_RefreshView()
}

ContextBrowser_ActivateEntry(entry) {
    global g_ContextBrowserCurrentDir
    if (!IsObject(entry))
        return
    if (entry.type = "parent") {
        ContextBrowser_HandleBack()
        return
    }
    if (entry.type = "folder") {
        g_ContextBrowserCurrentDir := entry.path
        ContextBrowser_RefreshView()
        return
    }
    pastePath := Context_ResolvePastePath(entry.path)
    if (pastePath = "") {
        ShowCenteredOverlay_Utils("❌ Reference target not found for: " entry.name, 2500, BANNER_ACCENT_ERROR)
        return
    }
    CleanupContextBrowser()
    Sleep 50
    InsertText(pastePath)
}

ContextBrowser_OnSelectRow(rowNum) {
    ContextBrowser_EnsureGlobals()
    global g_ContextBrowserActive, g_ContextBrowserEntries
    if (!g_ContextBrowserActive)
        return
    if (rowNum < 1 || rowNum > g_ContextBrowserEntries.Length)
        return
    ContextBrowser_ActivateEntry(g_ContextBrowserEntries[rowNum])
}

ContextBrowser_OnListDoubleClick(lv, guiEvent, *) {
    rowNum := 0
    try rowNum := guiEvent.EventInfo
    if !rowNum
        rowNum := lv.GetNext(0, "F")
    ContextBrowser_OnSelectRow(rowNum)
}

ContextBrowser_OnEnter(*) {
    ContextBrowser_EnsureGlobals()
    global g_ContextBrowserListView
    if (!IsObject(g_ContextBrowserListView))
        return
    rowNum := g_ContextBrowserListView.GetNext(0, "F")
    ContextBrowser_OnSelectRow(rowNum)
}

ContextBrowser_RefreshView() {
    global g_ContextBrowserActive, g_ContextBrowserCurrentDir, g_ContextBrowserEntries
    global g_ContextBrowserGui, g_ContextBrowserListView, g_ContextBrowserPathCtrl

    dir := g_ContextBrowserCurrentDir
    if (dir = "" || !DirExist(dir)) {
        TrayTip("Context", "Folder not found.", "IconX")
        SetTimer(() => TrayTip(), -5000)
        CleanupContextBrowser()
        return
    }

    g_ContextBrowserEntries := ContextBrowser_BuildViewEntries(dir)
    if (IsObject(g_ContextBrowserPathCtrl))
        g_ContextBrowserPathCtrl.Value := Context_GetRelativeSubtitle(dir)

    if (!IsObject(g_ContextBrowserListView))
        return

    lv := g_ContextBrowserListView
    lv.Opt("-Redraw")
    lv.Delete()
    for entry in g_ContextBrowserEntries
        lv.Add("", ContextBrowser_FormatEntryLabel(entry), ContextBrowser_EntryKindLabel(entry))
    lv.Opt("+Redraw")
    if (g_ContextBrowserEntries.Length) {
        lv.Modify(1, "Select Focus Vis")
        try lv.Focus()
        catch {
        }
    }
    g_ContextBrowserActive := true
}

ContextBrowser_CreateGui() {
    global g_ContextBrowserGui, g_ContextBrowserListView, g_ContextBrowserPathCtrl

    g_ContextBrowserGui := Gui("+AlwaysOnTop +Resize +MinSize420x320", "Context")
    g_ContextBrowserGui.SetFont("s10", "Segoe UI")
    g_ContextBrowserPathCtrl := g_ContextBrowserGui.Add("Text", "xm w520", "")
    g_ContextBrowserListView := g_ContextBrowserGui.Add("ListView", "xm w520 h380 Grid -Multi", ["Name", "Kind"])
    g_ContextBrowserListView.OnEvent("DoubleClick", ContextBrowser_OnListDoubleClick)
    g_ContextBrowserGui.Add("Text", "xm", "↑↓ move · Enter select · Backspace up · Esc close")
    g_ContextBrowserGui.OnEvent("Escape", HandleContextBrowserEscape)
    g_ContextBrowserGui.OnEvent("Close", (*) => CleanupContextBrowser())
}

ContextBrowser_ShowGui() {
    global g_ContextBrowserGui

    MonitorGetWorkArea(1, &monitorLeft, &monitorTop, &monitorRight, &monitorBottom)
    ContextBrowser_GetActiveMonitorWorkArea(&monitorLeft, &monitorTop, &monitorRight, &monitorBottom)
    monitorWidth := monitorRight - monitorLeft
    monitorHeight := monitorBottom - monitorTop

    g_ContextBrowserGui.Show("w540 h480")
    g_ContextBrowserGui.GetPos(, , &gw, &gh)
    guiX := monitorLeft + (monitorWidth - gw) // 2
    guiY := monitorTop + (monitorHeight - gh) // 2
    margin := 16
    guiX := Max(monitorLeft + margin, Min(guiX, monitorRight - gw - margin))
    guiY := Max(monitorTop + margin, Min(guiY, monitorBottom - gh - margin))
    g_ContextBrowserGui.Show("x" guiX " y" guiY " w540 h480")

    try Hotkey("Backspace", (*) => ContextBrowser_HandleBack(), "On")
    try Hotkey("Enter", ContextBrowser_OnEnter, "On")
    catch {
    }
    try g_ContextBrowserListView.Focus()
    catch {
    }
}

ContextBrowser_OpenAtCurrentDir() {
    ContextBrowser_EnsureGlobals()
    global g_ContextBrowserGui, g_ContextBrowserActive, g_ContextBrowserCurrentDir

    dir := g_ContextBrowserCurrentDir
    if (dir = "" || !DirExist(dir)) {
        TrayTip("Context", "Folder not found.", "IconX")
        SetTimer(() => TrayTip(), -5000)
        CleanupContextBrowser()
        return
    }

    if (!IsObject(g_ContextBrowserGui)) {
        ContextBrowser_CreateGui()
        ContextBrowser_ShowGui()
    }
    ContextBrowser_RefreshView()
    g_ContextBrowserActive := true
}

ShowContextBrowser() {
    ContextBrowser_EnsureGlobals()
    global g_ContextBrowserActive, g_ContextBrowserCurrentDir

    if (g_ContextBrowserActive) {
        CleanupContextBrowser()
        return
    }

    try {
        if (IsSet(g_HotstringSelectorActive) && g_HotstringSelectorActive)
            CleanupHotstringSelector()
    } catch {
    }
    try {
        if (IsSet(g_ProjectSelectorActive) && g_ProjectSelectorActive && IsSet(CleanupProjectSelector))
            CleanupProjectSelector()
    } catch {
    }

    root := Context_GetRoot()
    if !DirExist(root) {
        TrayTip("Context", "context folder not found at:`n" root, "IconX")
        SetTimer(() => TrayTip(), -5000)
        return
    }

    g_ContextBrowserCurrentDir := root
    ContextBrowser_OpenAtCurrentDir()
}

; One-shot: close Utility Shortcuts if still open (no expansion/macro chosen in time)
HotstringSelector_AutoCloseIfIdle() {
    global g_HotstringSelectorActive
    if (g_HotstringSelectorActive)
        CleanupHotstringSelector()
}

; =============================================================================
; CleanupHotstringSelector()
; =============================================================================
; PURPOSE: Closes hotstring selector GUI and disables all associated hotkeys.
;          Called when selector is closed via Escape key, character selection, or toggle.
;
; PROCESS:
;   1. Sets g_HotstringSelectorActive = false to prevent further character processing
;   2. Disables all character hotkeys (including uppercase variants for lowercase letters)
;   3. Handles special VK codes for comma (vkBC) and period (vkBE)
;   4. Disables Escape key handler
;   5. Destroys GUI object if it exists
;   6. Clears hotkey handlers array
;   7. Clears character mapping maps
;
; RETURNS: None (void function)
; SIDE EFFECTS: Resets all global state variables to initial values
; =============================================================================
CleanupHotstringSelector() {
    global g_HotstringSelectorActive, g_HotstringSelectorGui, g_HotstringHotkeyHandlers
    global g_HotstringPromptCharMap, g_HotstringGeminiArmed
    global g_HotstringCharMap
    global g_HS_SelectorOpenFile, g_HS_SelectorCloseRequestFile, g_HS_SelectorCloseCheckTimer
    global g_UtilitySelectorMode, g_UtilitySelectorCategory

    SetTimer(HotstringSelector_AutoCloseIfIdle, 0)
    ; Disable active flag
    g_HotstringSelectorActive := false

    ; Disable all character hotkeys
    for handler in g_HotstringHotkeyHandlers {
        try {
            key := handler.HasProp("key") ? handler.key : handler.char
            char := handler.HasProp("char") ? handler.char : key
            ; Handle special VK codes
            if (key = "vkBC" || char = ",") {
                Hotkey("vkBC", "Off")
            } else if (key = "vkBE" || char = ".") {
                Hotkey("vkBE", "Off")
            } else {
                Hotkey(key, "Off")
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
    Utils_EnsureGlobalEscapeHotkey()

    ; Disable Backspace hotkey (menu back)
    try {
        Hotkey("Backspace", "Off")
    } catch {
        ; Ignore
    }

    ; Stop IPC timer and clear sentinel files
    try SetTimer(Utils_CheckHotstringSelectorCloseRequest, 0)
    g_HS_SelectorCloseCheckTimer := ""
    try FileDelete(g_HS_SelectorOpenFile)
    catch {
    }
    try FileDelete(g_HS_SelectorCloseRequestFile)
    catch {
    }

    ; Clear handlers array
    g_HotstringHotkeyHandlers := []
    g_HotstringPromptCharMap := Map()
    g_HotstringGeminiArmed := false
    g_UtilitySelectorMode := "top"
    g_UtilitySelectorCategory := ""

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

; =============================================================================
; HandleHotstringChar(char)
; =============================================================================
; PURPOSE: Processes character key press events when hotstring selector is active.
;          Executes the action (text expansion, file open, or macro) associated with the character.
;
; PARAMETERS:
;   char: String - Single character that was pressed (e.g., "a", "1", ",")
;
; EXECUTION ORDER:
;   1. Special cases: Characters "9" and "0" trigger Miro board activation (hardcoded)
;   2. Quick-open files: Check g_QuickOpenFileCharMap for file path mappings
;   3. Macros: Check g_MacroCharMap for executable macro functions
;   4. Hotstrings: Check g_HotstringCharMap for text expansion mappings
;
; BEHAVIOR:
;   - Performs case-insensitive lookup (tries both original and lowercase)
;   - Closes selector GUI before executing action
;   - For hotstrings: Inserts text using InsertText() after 150ms delay
;   - For files: Opens file based on extension (Power BI files use special handler)
;   - For macros: Executes macro function directly
;
; RETURNS: None (void function)
; =============================================================================

; Navigate to Gemini, focus the prompt field, then paste first clipboard snippet (same as Win+Alt+Shift+1).
; Reference: "order called snippets" - Clip Angel top item sent via !v then ^!b.
; If optionalPromptText is non-empty, inserts that text into the prompt field instead (same as Win+Alt+Shift+U then L, prompt char).
; switchToFirstTab: when true (default), send Ctrl+1 and show tab-1 banner (AI Text Optimizer / ^!#4). When false, use currently active Gemini tab if any, else first Gemini window, without changing tab (delay-submit flow).
GeminiNavigateFocusAndPasteFirstSnippet(optionalPromptText := "", switchToFirstTab := true) {
    SetTitleMatchMode(2)
    geminiHwnd := 0
    if (!switchToFirstTab) {
        ; Prefer the currently active window if it is already a Gemini tab (do not switch tabs).
        try {
            activeHwnd := WinExist("A")
            if (activeHwnd && WinGetProcessName("ahk_id " activeHwnd) = "chrome.exe" && InStr(WinGetTitle("ahk_id " activeHwnd
            ), "gemini", false))
                geminiHwnd := activeHwnd
        } catch {
        }
    }
    if (!geminiHwnd) {
        try {
            for hwnd in WinGetList("ahk_exe chrome.exe") {
                try {
                    if InStr(WinGetTitle("ahk_id " hwnd), "gemini", false) {
                        geminiHwnd := hwnd
                        break
                    }
                } catch {
                }
            }
        } catch {
        }
    }

    if (!geminiHwnd) {
        ; Navigate to Gemini website
        Run "chrome.exe --new-window https://gemini.google.com/"
        if !WinWaitActive("ahk_exe chrome.exe", , 5)
            return
        Sleep 2500  ; Allow page to load
        geminiHwnd := WinExist("A")
    }

    if (geminiHwnd) {
        WinActivate("ahk_id " geminiHwnd)
        WinWaitActive("ahk_id " geminiHwnd, , 2)
        Sleep 350  ; Let UI finish processing the window transition before focus/paste
    } else {
        WinActivate("ahk_exe chrome.exe")
        WinWaitActive("ahk_exe chrome.exe", , 2)
        Sleep 350
    }

    if (switchToFirstTab) {
        ; Switch to first Gemini tab (Ctrl+1) and show number-one banner (consistent with Gemini.ahk)
        Send("^1")
        Sleep 280
        ShowSingleCharTabBanner_Utils(1)
    }

    ; Focus the Gemini prompt field (Anchor & Backtrack strategy)
    try {
        uia := UIA_Browser()
        Sleep 80
        anchorButton := 0
        try {
            anchorButton := uia.FindFirst({ Type: "50000", Name: "Open upload file menu", ControlType: "Button" })
            if (!anchorButton)
                anchorButton := uia.FindFirst({ Type: "50000", Name: "Open upload file menu", cs: false })
        } catch {
        }
        if (!anchorButton) {
            try {
                allButtons := uia.FindAll({ Type: "50000" })
                for button in allButtons {
                    try {
                        if (InStr(button.Name, "Open upload file menu", false)) {
                            anchorButton := button
                            break
                        }
                    } catch {
                        continue
                    }
                }
            } catch {
            }
        }
        if (anchorButton) {
            try {
                anchorButton.SetFocus()
                Sleep 25
                SendInput "+{Tab}"
                Sleep 15
            } catch {
                try {
                    promptField := FindGeminiPromptField(uia)
                    if (promptField)
                        promptField.SetFocus()
                } catch as e {
                }
            }
        } else {
            try {
                promptField := FindGeminiPromptField(uia)
                if (promptField)
                    promptField.SetFocus()
            } catch as e {
            }
        }
    } catch {
    }

    ; Explicitly target Gemini window again before paste so paste goes to Gemini, not the trigger window
    if (geminiHwnd && WinExist("ahk_id " geminiHwnd)) {
        WinActivate("ahk_id " geminiHwnd)
        WinWaitActive("ahk_id " geminiHwnd, , 2)
        Sleep 150
    }

    if (optionalPromptText != "") {
        InsertText(optionalPromptText)
    } else {
        ; Paste first clipboard snippet (same as Win+Alt+Shift+1: order called snippets)
        Send "!v"
        Sleep 50
        Send "^!b"
    }
    ; Brief delay so paste is received and UI/character limits register before any submit or focus change
    Sleep 250
    ; Same sound as when opening Gemini (focus/paste feedback)
    ScriptSoundPlay(A_ScriptDir . "\sounds\gemini-focused.wav")
}

; Returns grammar preset (from prompt/grammar.txt or fallback). Matches InitHotstringsCheatSheet catch for :o:cgrammar.
GetGrammarPromptText() {
    promptDir := A_ScriptDir "\prompt"
    try {
        return FileRead(promptDir "\grammar.txt")
    } catch {
        return "Correct grammar, spelling, punctuation, and casing. Give back only the text.`n"
    }
}

; Returns the AI Text Optimizer prompt text (from prompt/aiopt.txt or fallback). Used by Ctrl+Alt+Win+4 and L+4 flow.
GetAioptPromptText() {
    promptDir := A_ScriptDir "\prompt"
    try {
        return FileRead(promptDir "\aiopt.txt")
    } catch {
        return "Rewrite the input text so it becomes AI-oriented. Preserve all important information.`n"
    }
}

; Delayed submit flow: show 4s banner, allow N to cancel auto-submit; then navigate+paste and optionally send Enter.
; Tell Gemini.ahk to start background completion monitor (must match WM_START_DELAYED_SUBMIT_MONITOR in Gemini.ahk).
GeminiDelayedSubmitMonitorStartFromUtils(originalHwnd, geminiChromeHwnd) {
    WM_START_DELAYED_SUBMIT_MONITOR := 0x8002
    targetHwnd := GetGeminiScriptMsgTargetHwnd()
    if (targetHwnd) {
        try SendMessage(WM_START_DELAYED_SUBMIT_MONITOR, originalHwnd, geminiChromeHwnd, , "ahk_id " targetHwnd)
    }
}

; Tell Gemini.ahk to stop the delayed-submit monitor so "Copy response?" is not shown (when user chose S or N at 5s).
GeminiDelayedSubmitMonitorStopFromUtils() {
    WM_STOP_DELAYED_SUBMIT_MONITOR := 0x8003
    targetHwnd := GetGeminiScriptMsgTargetHwnd()
    if (targetHwnd) {
        try SendMessage(WM_STOP_DELAYED_SUBMIT_MONITOR, 0, 0, , "ahk_id " targetHwnd)
    }
}

; Paste transcription to Gemini prompt only (no Enter, no 4s banner). Used when user presses S at 5s dictation confirm.
DEPRECATED_GeminiDictationPasteOnlyFlow() {
    restoreHwnd := WinExist("A")
    GeminiNavigateFocusAndPasteFirstSnippet("", false)
    if (restoreHwnd && WinExist("ahk_id " restoreHwnd))
        WinActivate("ahk_id " restoreHwnd)
}

DEPRECATED_GeminiDelayedSubmitFlow() {
    global g_HotstringGeminiAutoSubmit, g_HotstringGeminiRestoreHwnd
    g_HotstringGeminiRestoreHwnd := WinExist("A")
    g_HotstringGeminiAutoSubmit := true
    GeminiFinalizeSubmit()
}

; Delay (ms) after paste and before Send Enter in Gemini delayed-submit flow. Prevents premature send and lets the UI register paste + character limits.
global g_GeminiDelayedSubmit_PreEnterDelayMs := 1000
; Max ms to wait for prompt field to show content before sending Enter (guarantee layer). Poll interval = 200 ms.
global g_GeminiDelayedSubmit_WaitContentMaxMs := 5000

; Returns non-empty trimmed text if Gemini prompt field has content (Value or TextPattern); "" on failure or empty. Used to guarantee message is present before submit.
GeminiPromptFieldGetText() {
    try {
        uia := UIA_Browser()
        pf := FindGeminiPromptField(uia)
        if (!pf)
            return ""
        try {
            text := Trim(pf.Value)
            if (text != "")
                return text
        } catch {
        }
        try {
            text := Trim(pf.TextPattern.DocumentRange.GetText(-1))
            if (text != "")
                return text
        } catch {
        }
    } catch {
    }
    return ""
}

GeminiFinalizeSubmit() {
    global g_HotstringGeminiAutoSubmit, g_HotstringGeminiRestoreHwnd, g_GeminiDelayedSubmit_PreEnterDelayMs,
        g_GeminiDelayedSubmit_WaitContentMaxMs

    try Hotkey("n", "Off")
    try Hotkey("N", "Off")
    try Hotkey("y", "Off")
    try Hotkey("Y", "Off")
    HotstringGeminiBanner_Hide()

    ; Delay-submit flow: do not switch tabs; paste to currently active Gemini tab
    GeminiNavigateFocusAndPasteFirstSnippet("", false)

    didAutoSubmit := false
    geminiChromeHwnd := 0
    if (g_HotstringGeminiAutoSubmit) {
        ; Execution delay so paste is fully received and UI/character limits register before submit
        Sleep (g_GeminiDelayedSubmit_PreEnterDelayMs)
        ; Guarantee layer: wait until prompt field has content (or timeout) so we don't send Enter prematurely
        pollIntervalMs := 200
        endTick := A_TickCount + g_GeminiDelayedSubmit_WaitContentMaxMs
        contentFound := false
        while (A_TickCount < endTick) {
            if (GeminiPromptFieldGetText() != "") {
                contentFound := true
                break
            }
            Sleep pollIntervalMs
        }
        Send("{Enter}")
        geminiChromeHwnd := WinExist("A")
        didAutoSubmit := true
    }

    g_HotstringGeminiAutoSubmit := true

    ; Return focus to the window the user had before paste (whether Enter was sent or not)
    if (g_HotstringGeminiRestoreHwnd && WinExist("ahk_id " g_HotstringGeminiRestoreHwnd)) {
        WinActivate("ahk_id " g_HotstringGeminiRestoreHwnd)
    }

    ; If we auto-submitted (user did not cancel), ask Gemini.ahk to monitor for completion and show "Copy? [N] [R]" when done
    if (didAutoSubmit && geminiChromeHwnd)
        GeminiDelayedSubmitMonitorStartFromUtils(g_HotstringGeminiRestoreHwnd, geminiChromeHwnd)
}

HandleHotstringChar(char) {
    global g_HotstringSelectorActive, g_HotstringCharMap, g_QuickOpenFileCharMap, g_MacroCharMap
    global g_HotstringPromptCharMap, g_HotstringGeminiArmed
    global g_UtilitySelectorMode, g_UtilityTopCategoryById

    ; Only process if selector is active
    if (!g_HotstringSelectorActive) {
        return
    }

    ; Top-level category selection (R/P/M/G/L/H) or Context browser (C)
    if (g_UtilitySelectorMode = "top") {
        ch := StrLower(char)
        if (ch = "c") {
            CleanupHotstringSelector()
            ShowContextBrowser()
            return
        }
        if (g_UtilityTopCategoryById.Has(ch)) {
            category := g_UtilityTopCategoryById[ch]
            if (category = "Context")
                return
            UtilitySelector_SwitchToCategory(category)
        }
        return
    }

    ; L key: first press = arm Gemini mode (show banner); second press (double-tap) = navigate to Gemini, focus field, paste first snippet.
    if (char = "l" || char = "L") {
        if (g_HotstringGeminiArmed) {
            ; Double-tap L: delayed submit flow (paste + Enter to Gemini).
            CleanupHotstringSelector()
            D2C_FlowManager.GetInstance().StartFromHotstring()
            g_HotstringGeminiArmed := false
            return
        }
        g_HotstringGeminiArmed := true
        ; Show banner when entering Gemini mode (same pattern as Project Selector "Entering Selection Mode").
        HotstringGeminiBanner_Show("⌨ Entering Gemini Mode - Select prompt")
        SetTimer(HotstringGeminiBanner_Hide, -1500)  ; Hide banner after 1.5 s
        SetTimer(DisarmHotstringGeminiMode, -4000)
        return
    }

    ; Consume the armed state on the next selection (any selection), but only redirect Prompts.
    useGemini := false
    if (g_HotstringGeminiArmed) {
        useGemini := g_HotstringPromptCharMap.Has(StrLower(char)) || g_HotstringPromptCharMap.Has(char)
        g_HotstringGeminiArmed := false
    }

    ; Special handling for Miro boards (characters "9" and "0")
    ; Gate to Links category so other views can't trigger it.
    global g_UtilitySelectorCategory
    if (g_UtilitySelectorCategory = "Links") {
        if (char = "9") {
            CleanupHotstringSelector()
            FindAndActivateMiroWindow("https://miro.com/app/board/uXjVJdbNFkA=/", "CIP & UX Integration")
            return
        } else if (char = "0") {
            CleanupHotstringSelector()
            FindAndActivateMiroWindow("https://miro.com/app/board/uXjVJVZSXvk=/", "CIP Dashboard")
            return
        }
    }

    ; Category-scoped dispatch (prevents cross-menu execution when chars overlap)
    global g_UtilityHotstringCharMapByCategory, g_UtilitySelectorCategory

    ResolveExpansion() {
        exp := ""
        try {
            if (IsObject(g_UtilityHotstringCharMapByCategory) && g_UtilityHotstringCharMapByCategory.Has(
                g_UtilitySelectorCategory)) {
                exp := g_UtilityHotstringCharMapByCategory[g_UtilitySelectorCategory].Get(char, "")
                if (exp = "")
                    exp := g_UtilityHotstringCharMapByCategory[g_UtilitySelectorCategory].Get(StrLower(char), "")
            }
        } catch {
            exp := ""
        }
        if (exp = "") {
            exp := g_HotstringCharMap.Get(char, "")
            if (exp = "")
                exp := g_HotstringCharMap.Get(StrLower(char), "")
        }
        return exp
    }

    ResolveFilePath() {
        fp := g_QuickOpenFileCharMap.Get(char, "")
        if (fp = "")
            fp := g_QuickOpenFileCharMap.Get(StrLower(char), "")
        return fp
    }

    ResolveMacro() {
        fn := g_MacroCharMap.Get(char, "")
        if (fn = "")
            fn := g_MacroCharMap.Get(StrLower(char), "")
        return fn
    }

    TryRunFile(fp) {
        if (fp = "")
            return false
        CleanupHotstringSelector()
        SplitPath(fp, , , &ext)
        ext := StrLower(ext)
        if (ext = "pbix") {
            FindAndActivatePowerBIFile(fp)
        } else {
            fpTrim := Trim(fp)
            if (SubStr(fpTrim, 1, 8) = "https://" || SubStr(fpTrim, 1, 7) = "http://") {
                StudyLink_OpenUrlInChrome(fpTrim, true)
            } else {
                try Run(fp)
                catch {
                }
            }
        }
        return true
    }

    TryRunMacro(fn) {
        if (fn = "")
            return false
        CleanupHotstringSelector()
        try fn()
        catch {
        }
        return true
    }

    expansion := ""
    filePath := ""
    macroFunc := ""

    if (g_UtilitySelectorCategory = "Links") {
        filePath := ResolveFilePath()
        if (TryRunFile(filePath))
            return
        ; fallback for unexpected collisions
        expansion := ResolveExpansion()
    } else if (g_UtilitySelectorCategory = "Macros") {
        macroFunc := ResolveMacro()
        if (TryRunMacro(macroFunc))
            return
        expansion := ResolveExpansion()
    } else {
        ; Projects / Prompts / Hotstrings / General: hotstring-first
        expansion := ResolveExpansion()
        if (expansion = "") {
            macroFunc := ResolveMacro()
            if (TryRunMacro(macroFunc))
                return
            filePath := ResolveFilePath()
            if (TryRunFile(filePath))
                return
        }
    }

    if (expansion != "") {
        ; Cleanup first (closes GUI, disables hotkeys)
        CleanupHotstringSelector()

        if (useGemini) {
            ; L+Prompt selection: redirect to Gemini (focus prompt field, paste, do NOT submit).
            HotstringGeminiBanner_Show("📤 Gemini: inserting prompt...")
            try {
                SetTitleMatchMode(2)
                geminiHwnd := 0
                try {
                    for hwnd in WinGetList("ahk_exe chrome.exe") {
                        try {
                            if InStr(WinGetTitle("ahk_id " hwnd), "gemini", false) {
                                geminiHwnd := hwnd
                                break
                            }
                        } catch {
                            ; Skip invalid windows
                        }
                    }
                } catch {
                    ; Ignore WinGetList errors
                }

                if (geminiHwnd) {
                    WinActivate("ahk_id " geminiHwnd)
                    WinWaitActive("ahk_id " geminiHwnd, , 2)
                } else {
                    ; Per your preference: fallback to any Chrome window if Gemini isn't found.
                    WinActivate("ahk_exe chrome.exe")
                    WinWaitActive("ahk_exe chrome.exe", , 2)
                }

                ; Explicitly target Gemini tabs based on character:
                ; - L+1/2/3  -> Tab 2 (right Gemini tab, temporary prompts)
                ; - L+4/5/Q/W/E/R/T/A -> Tab 1 (left Gemini tab, main workflow)
                if (char = "1" || char = "2" || char = "3") {
                    ; Chrome convention: Ctrl+2 selects the second tab in the window.
                    Send("^2")
                    Sleep 120
                    ShowSingleCharTabBanner_Utils(2)
                } else if (char = "4" || char = "5"
                    || char = "q" || char = "Q"
                    || char = "w" || char = "W"
                    || char = "e" || char = "E"
                    || char = "r" || char = "R"
                    || char = "t" || char = "T"
                    || char = "a" || char = "A") {
                    ; Ensure Tab 1 is active before inserting the prompt.
                    Send("^1")
                    Sleep 120
                    ShowSingleCharTabBanner_Utils(1)
                }

                ; Focus the Gemini prompt field (shared helper; no chime — paste path plays its own sound).
                try {
                    uia := UIA_Browser()
                    Gemini_FocusPromptWithChime(uia, { playChime: false, useAnchorFallback: true })
                } catch {
                    ; If focus fails, we still attempt to paste (user can click manually).
                }

                ; Paste the text (do NOT send Enter)
                InsertText(expansion)
                ; Same sound as when opening Gemini (focus/paste feedback)
                ScriptSoundPlay(A_ScriptDir . "\sounds\gemini-focused.wav")
            } finally {
                HotstringGeminiBanner_Hide()
            }
            return
        }

        ; Standard behavior: paste into current active text field.
        ; Small delay to ensure target window has focus before pasting
        Sleep 150
        InsertText(expansion)
    }
}

; =============================================================================
; CreateHotstringCharHandler(char)
; =============================================================================
; PURPOSE: Factory function that creates a hotkey handler function with proper closure over character value.
;          Required because AutoHotkey hotkey handlers need unique function instances per character.
;
; PARAMETERS:
;   char: String - Character to create handler for
;
; RETURNS: Function object that calls HandleHotstringChar(char) when invoked
; =============================================================================
DisarmHotstringGeminiMode(*) {
    global g_HotstringGeminiArmed
    g_HotstringGeminiArmed := false
}

CreateHotstringCharHandler(char) {
    ; Return a function that captures the char value at creation time via closure
    return (*) => HandleHotstringChar(char)
}

; =============================================================================
; HandleHotstringEscape(*)
; =============================================================================
; PURPOSE: Handles Escape key press to close hotstring selector without executing any action.
;
; PARAMETERS: None (varargs function signature for hotkey compatibility)
; RETURNS: None (void function)
; =============================================================================
HandleHotstringEscape(*) {
    global g_HotstringSelectorActive
    if (g_HotstringSelectorActive) {
        CleanupHotstringSelector()
    }
}

HandleUtilitySelectorBack(*) {
    global g_HotstringSelectorActive, g_UtilitySelectorMode
    if (!g_HotstringSelectorActive)
        return
    if (g_UtilitySelectorMode = "category") {
        UtilitySelector_SwitchToTop()
    }
}

UtilitySelector_SwitchToTop() {
    global g_UtilitySelectorMode, g_UtilitySelectorCategory
    g_UtilitySelectorMode := "top"
    g_UtilitySelectorCategory := ""
    try UtilitySelector_RefreshUiAndHotkeys()
    catch {
    }
}

UtilitySelector_SwitchToCategory(category) {
    global g_UtilitySelectorMode, g_UtilitySelectorCategory
    g_UtilitySelectorMode := "category"
    g_UtilitySelectorCategory := category
    try UtilitySelector_RefreshUiAndHotkeys()
    catch {
    }
}

UtilitySelector_MapInternalCategoryToTop(internalCategory) {
    if (internalCategory = "Files & Links")
        return "Links"
    if (internalCategory = "General")
        return "General"
    if (internalCategory = "Links")
        return "Links"
    if (internalCategory = "Hotstrings")
        return "Hotstrings"
    ; Unknown/legacy categories are folded into General now that top-level "Hot Strings" is removed.
    if (internalCategory != "Prompts" && internalCategory != "Projects" && internalCategory != "Macros")
        return "General"
    return internalCategory
}

UtilitySelector_GetAllowedCharsForCurrentView() {
    global g_UtilitySelectorMode, g_UtilitySelectorCategory
    global g_UtilityTopCategoryById, g_UtilitySelectorAllItems
    allowed := Map()

    if (g_UtilitySelectorMode = "top") {
        for id, category in g_UtilityTopCategoryById {
            allowed[id] := true
        }
        return allowed
    }

    ; Category view: enable only actionable items in the selected category.
    ; (Empty placeholders are displayed but not bound.)
    for item in g_UtilitySelectorAllItems {
        if (item.category = g_UtilitySelectorCategory && !item.isEmpty) {
            allowed[item.char] := true
        }
    }

    ; Prompts view: enable Gemini modifier key 'L' workflow
    if (g_UtilitySelectorCategory = "Prompts") {
        allowed["l"] := true
        allowed["L"] := true
    }

    return allowed
}

UtilitySelector_RebindHotkeys() {
    global g_HotstringHotkeyHandlers, g_UtilitySelectorMode
    allowed := UtilitySelector_GetAllowedCharsForCurrentView()

    ; Disable previously-bound hotkeys
    for handler in g_HotstringHotkeyHandlers {
        try {
            key := handler.key
            if (key = "vkBC") {
                Hotkey("vkBC", "Off")
            } else if (key = "vkBE") {
                Hotkey("vkBE", "Off")
            } else {
                Hotkey(key, "Off")
            }
        } catch {
        }
    }
    g_HotstringHotkeyHandlers := []

    ; Bind allowed character hotkeys
    for char, _ in allowed {
        handler := CreateHotstringCharHandler(char)
        try {
            if (char = ",") {
                Hotkey("vkBC", handler, "On")
                g_HotstringHotkeyHandlers.Push({ key: "vkBC", char: char, handler: handler })
            } else if (char = ".") {
                Hotkey("vkBE", handler, "On")
                g_HotstringHotkeyHandlers.Push({ key: "vkBE", char: char, handler: handler })
            } else {
                Hotkey(char, handler, "On")
                g_HotstringHotkeyHandlers.Push({ key: char, char: char, handler: handler })
                if (RegExMatch(char, "^[a-z]$")) {
                    Hotkey(StrUpper(char), handler, "On")
                    g_HotstringHotkeyHandlers.Push({ key: StrUpper(char), char: char, handler: handler })
                }
            }
        } catch {
        }
    }

    ; Back navigation
    if (g_UtilitySelectorMode = "category") {
        try Hotkey("Backspace", HandleUtilitySelectorBack, "On")
    } else {
        try Hotkey("Backspace", "Off")
    }

    ; Escape always closes
    Hotkey("Escape", HandleHotstringEscape, "On")
}

UtilitySelector_BuildTopLevelText() {
    global g_UtilityTopCategories, g_UtilitySelectorAllItems
    ; Count actionable items per category
    counts := Map()
    for cat in g_UtilityTopCategories
        counts[cat] := 0
    for item in g_UtilitySelectorAllItems {
        if (!item.isEmpty && counts.Has(item.category)) {
            counts[item.category] := counts[item.category] + 1
        }
    }

    text := ""
    text .= "[R] Prompts (" . counts["Prompts"] . ")`n"
    text .= "[P] Projects (" . counts["Projects"] . ")`n"
    text .= "[M] Macros (" . counts["Macros"] . ")`n"
    text .= "[G] General (" . counts["General"] . ")`n"
    text .= "[L] Links (" . counts["Links"] . ")`n"
    text .= "[H] Hotstrings (" . counts["Hotstrings"] . ")`n"
    text .= "[C] Context — paste file path`n"
    text .= "`nPress R/P/M/G/L/H/C to open.`n"
    return text
}

UtilitySelector_BuildTopLevelRich() {
    global g_UtilityTopCategories, g_UtilitySelectorAllItems
    counts := Map()
    for cat in g_UtilityTopCategories
        counts[cat] := 0
    for item in g_UtilitySelectorAllItems {
        if (!item.isEmpty && counts.Has(item.category)) {
            counts[item.category] := counts[item.category] + 1
        }
    }

    lines := []
    lines.Push({ text: "[R] Prompts (" . counts["Prompts"] . ")", key: "r" })
    lines.Push({ text: "[P] Projects (" . counts["Projects"] . ")", key: "p" })
    lines.Push({ text: "[M] Macros (" . counts["Macros"] . ")", key: "m" })
    lines.Push({ text: "[G] General (" . counts["General"] . ")", key: "g" })
    lines.Push({ text: "[L] Links (" . counts["Links"] . ")", key: "l" })
    lines.Push({ text: "[H] Hotstrings (" . counts["Hotstrings"] . ")", key: "h" })
    lines.Push({ text: "[C] Context — paste file path", key: "c" })
    lines.Push({ text: "" })
    lines.Push({ text: "Press R/P/M/G/L/H/C to open." })
    return lines
}

UtilitySelector_BuildCategoryText(isPortrait) {
    global g_UtilitySelectorCategory, g_UtilitySelectorAllItems
    ; Filter items for this category
    items := []
    for item in g_UtilitySelectorAllItems {
        if (item.category = g_UtilitySelectorCategory)
            items.Push(item)
    }

    header := "- " . g_UtilitySelectorCategory . " -`n"
    if (items.Length = 0) {
        return header . "(no items)`n`nBackspace = back | Esc = close"
    }

    if (isPortrait) {
        text := header
        for item in items {
            if (item.HasProp("isSectionSpacer") && item.isSectionSpacer) {
                text .= "`n"
                continue
            }
            text .= item.text . "`n"
        }
        text .= "`nBackspace = back | Esc = close"
        return text
    }

    ; Landscape: two columns; full-width for mnemonic subsection rows (same as Rich path)
    PadString(str, width) {
        len := StrLen(str)
        if (len >= width)
            return str
        padding := width - len
        spaces := ""
        loop padding
            spaces .= " "
        return str . spaces
    }

    maxItemLength := 0
    for item in items {
        if (item.HasProp("isSectionHeader") && item.isSectionHeader)
            continue
        if (item.HasProp("isSectionSpacer") && item.isSectionSpacer)
            continue
        if (StrLen(item.text) > maxItemLength)
            maxItemLength := StrLen(item.text)
    }
    columnWidth := maxItemLength + 2
    if (columnWidth < 36)
        columnWidth := 36
    columnSpacing := "  "
    fullWidth := columnWidth * 2 + StrLen(columnSpacing)

    CenterStringInWidth(str, width) {
        len := StrLen(str)
        if (len >= width)
            return SubStr(str, 1, width)
        pad := width - len
        left := pad // 2
        right := pad - left
        ls := ""
        rs := ""
        loop left
            ls .= " "
        loop right
            rs .= " "
        return ls . str . rs
    }

    text := header
    i := 1
    while (i <= items.Length) {
        it := items[i]
        if (it.HasProp("isSectionSpacer") && it.isSectionSpacer) {
            text .= "`n"
            i++
            continue
        }
        if (it.HasProp("isSectionHeader") && it.isSectionHeader) {
            text .= CenterStringInWidth(it.text, fullWidth) . "`n"
            i++
            continue
        }
        leftItem := it
        i++
        rightItem := ""
        if (i <= items.Length) {
            rit := items[i]
            if (!(rit.HasProp("isSectionHeader") && rit.isSectionHeader) && !(rit.HasProp("isSectionSpacer") && rit.isSectionSpacer
            ))
                rightItem := rit, i++
        }
        leftText := PadString(leftItem.text, columnWidth)
        rightText := (IsObject(rightItem) && rightItem.HasProp("text")) ? rightItem.text : ""
        text .= leftText . columnSpacing . rightText . "`n"
    }
    text .= "`nBackspace = back | Esc = close"
    return text
}

UtilitySelector_BuildCategoryRich(isPortrait) {
    global g_UtilitySelectorCategory, g_UtilitySelectorAllItems
    items := []
    for item in g_UtilitySelectorAllItems {
        if (item.category = g_UtilitySelectorCategory)
            items.Push(item)
    }

    lines := []
    lines.Push({ text: "- " . g_UtilitySelectorCategory . " -" })
    if (items.Length = 0) {
        lines.Push({ text: "(no items)" })
        lines.Push({ text: "" })
        lines.Push({ text: "Backspace = back | Esc = close" })
        return lines
    }

    if (isPortrait) {
        for item in items {
            if (item.HasProp("isSectionSpacer") && item.isSectionSpacer) {
                lines.Push({ text: "", key: "" })
                continue
            }
            isSub := item.HasProp("isSectionHeader") && item.isSectionHeader
            lines.Push({ text: item.text, key: item.isEmpty ? "" : item.char, isMnemonicSubsection: isSub })
        }
        lines.Push({ text: "" })
        lines.Push({ text: "Backspace = back | Esc = close" })
        return lines
    }

    ; Landscape: two columns; full-width rows for mnemonic subsection spacer/header inside Prompts
    PadString(str, width) {
        len := StrLen(str)
        if (len >= width)
            return str
        padding := width - len
        spaces := ""
        loop padding
            spaces .= " "
        return str . spaces
    }

    CenterStringInWidth(str, width) {
        len := StrLen(str)
        if (len >= width)
            return SubStr(str, 1, width)
        pad := width - len
        left := pad // 2
        right := pad - left
        ls := ""
        rs := ""
        loop left
            ls .= " "
        loop right
            rs .= " "
        return ls . str . rs
    }

    maxItemLength := 0
    for item in items {
        if (item.HasProp("isSectionHeader") && item.isSectionHeader)
            continue
        if (item.HasProp("isSectionSpacer") && item.isSectionSpacer)
            continue
        if (StrLen(item.text) > maxItemLength)
            maxItemLength := StrLen(item.text)
    }
    columnWidth := maxItemLength + 2
    if (columnWidth < 36)
        columnWidth := 36
    columnSpacing := "  "
    fullWidth := columnWidth * 2 + StrLen(columnSpacing)

    i := 1
    while (i <= items.Length) {
        it := items[i]
        if (it.HasProp("isSectionSpacer") && it.isSectionSpacer) {
            lines.Push({ text: "", key: "", keyRight: "", rightStartCharPos: 0 })
            i++
            continue
        }
        if (it.HasProp("isSectionHeader") && it.isSectionHeader) {
            lines.Push({ text: CenterStringInWidth(it.text, fullWidth), key: "", keyRight: "", rightStartCharPos: 0,
                isMnemonicSubsection: true })
            i++
            continue
        }
        leftItem := it
        i++
        rightItem := ""
        if (i <= items.Length) {
            rit := items[i]
            if (!(rit.HasProp("isSectionHeader") && rit.isSectionHeader) && !(rit.HasProp("isSectionSpacer") && rit.isSectionSpacer
            ))
                rightItem := rit, i++
        }
        leftText := PadString(leftItem.text, columnWidth)
        rightText := rightItem ? rightItem.text : ""
        leftKey := leftItem.isEmpty ? "" : leftItem.char
        rightKey := (IsObject(rightItem) && rightItem != "") ? (rightItem.isEmpty ? "" : rightItem.char) : ""
        lineText := leftText . columnSpacing . rightText
        rightStartCharPos := StrLen(leftText . columnSpacing) + 1
        lines.Push({ text: lineText, key: leftKey, keyRight: rightKey, rightStartCharPos: rightStartCharPos })
    }
    lines.Push({ text: "" })
    lines.Push({ text: "Backspace = back | Esc = close" })
    return lines
}

UtilitySelector_BuildDisplayText(isPortrait) {
    global g_UtilitySelectorMode
    if (g_UtilitySelectorMode = "top")
        return UtilitySelector_BuildTopLevelText()
    return UtilitySelector_BuildCategoryText(isPortrait)
}

UtilitySelector_BuildDisplayRich(isPortrait) {
    global g_UtilitySelectorMode
    if (g_UtilitySelectorMode = "top")
        return UtilitySelector_BuildTopLevelRich()
    return UtilitySelector_BuildCategoryRich(isPortrait)
}

UtilitySelector_RefreshUiAndHotkeys() {
    global g_HotstringSelectorGui, g_UtilitySelectorTitleCtrl, g_UtilitySelectorEditCtrl
    global g_UtilitySelectorIsPortrait, g_UtilitySelectorMonitor
    global g_UtilitySelectorMode, g_UtilitySelectorCategory

    if (!IsObject(g_HotstringSelectorGui) || !IsObject(g_UtilitySelectorEditCtrl))
        return

    title := "Utility Shortcuts"
    if (g_UtilitySelectorMode = "category" && g_UtilitySelectorCategory != "")
        title := title . " - " . g_UtilitySelectorCategory

    if (IsObject(g_UtilitySelectorTitleCtrl))
        try g_UtilitySelectorTitleCtrl.Text := title

    global g_UtilitySelectorFontSize
    displayText := UtilitySelector_BuildDisplayText(g_UtilitySelectorIsPortrait)
    displayLines := UtilitySelector_BuildDisplayRich(g_UtilitySelectorIsPortrait)
    try MnemonicRich_Render(g_UtilitySelectorEditCtrl, displayLines, g_UtilitySelectorFontSize, 6, "Consolas", "CDD6F4",
        "1E1E2E")

    ; Resize based on new content (reuse existing sizing rules)
    lineCount := 1
    loop parse, displayText, "`n"
        lineCount++
    lineHeight := 18
    textControlHeight := lineCount * lineHeight
    minHeight := 150

    monitorWidth := g_UtilitySelectorMonitor["width"]
    monitorHeight := g_UtilitySelectorMonitor["height"]
    monitorLeft := g_UtilitySelectorMonitor["left"]
    monitorTop := g_UtilitySelectorMonitor["top"]

    if (g_UtilitySelectorIsPortrait) {
        maxHeightPercent := 0.85
        maxHeight := Floor(monitorHeight * maxHeightPercent)
        if (textControlHeight < minHeight)
            textControlHeight := minHeight
        if (textControlHeight > maxHeight)
            textControlHeight := maxHeight
        baseWidth := (monitorWidth < 800) ? 400 : (monitorWidth < 1200) ? 500 : 500
        if (baseWidth > monitorWidth - 40)
            baseWidth := monitorWidth - 40
    } else {
        maxHeightPercent := (monitorHeight < 800) ? 0.90 : 0.85
        maxHeight := Floor(monitorHeight * maxHeightPercent)
        if (textControlHeight < minHeight)
            textControlHeight := minHeight
        if (textControlHeight > maxHeight)
            textControlHeight := maxHeight
        baseWidth := (monitorWidth < 1200) ? 650 : (monitorWidth < 1920) ? 800 : 1000
        if (baseWidth > monitorWidth - 40)
            baseWidth := monitorWidth - 40
    }

    textControlWidth := baseWidth - 20
    try {
        g_UtilitySelectorTitleCtrl.Move(, , textControlWidth)
        g_UtilitySelectorEditCtrl.Move(, , textControlWidth, textControlHeight)
    } catch {
    }
    global g_UtilitySelectorFooterCtrl
    if (IsObject(g_UtilitySelectorFooterCtrl)) {
        try {
            g_UtilitySelectorEditCtrl.GetPos(, &ftEditY)
            if (ftEditY > 0)
                g_UtilitySelectorFooterCtrl.Move(, ftEditY + textControlHeight + 10, textControlWidth)
        } catch {
        }
    }

    ; totalHeight: top-margin + title(s11) + gap + separator + gap + edit + gap + footer + bottom-margin
    totalHeight := 10 + 24 + 10 + 1 + 10 + textControlHeight + 10 + 18 + 10
    guiWidth := baseWidth

    marginX := 20
    marginY := 20
    guiX := monitorLeft + (monitorWidth - guiWidth) // 2
    guiY := monitorTop + (monitorHeight - totalHeight) // 2
    if (guiX < monitorLeft + marginX)
        guiX := monitorLeft + marginX
    if (guiY < monitorTop + marginY)
        guiY := monitorTop + marginY
    if (guiX + guiWidth > monitorLeft + monitorWidth - marginX)
        guiX := monitorLeft + monitorWidth - guiWidth - marginX
    if (guiY + totalHeight > monitorTop + monitorHeight - marginY)
        guiY := monitorTop + monitorHeight - totalHeight - marginY

    try g_HotstringSelectorGui.Show("NA w" . guiWidth . " h" . totalHeight . " x" . guiX . " y" . guiY)

    UtilitySelector_RebindHotkeys()
    SetTimer(HotstringSelector_AutoCloseIfIdle, -3000)
}

; Triggers for InitTechniquePromptHotstrings - used to group Utility Shortcuts Prompts submenu.
UtilitySelector_IsMnemonicTechniquePrompt(trigger) {
    if (trigger = "")
        return false
    static mnemonic := Map(
        ":o:mnemonic", true,
        ":o:ytranscript", true,
        ":o:readaloud", true,
        ":o:revision", true,
        ":o:storyreduction", true,
        ":o:punctualbeast", true,
        ":o:imgpreserve", true,
    )
    return mnemonic.Has(trigger)
}

; After global char sort: keep non-Prompts order; replace Prompts subsequence with general, subsection row, mnemonic technique.
UtilitySelector_ReorderPromptsMnemonicsSection(&rebuilt) {
    promptItems := []
    for it in rebuilt {
        if (IsObject(it) && it.HasProp("category") && it.category = "Prompts")
            promptItems.Push(it)
    }
    if (promptItems.Length = 0)
        return

    general := []
    tech := []
    for it in promptItems {
        tr := it.HasProp("trigger") ? it.trigger : ""
        if (UtilitySelector_IsMnemonicTechniquePrompt(tr))
            tech.Push(it)
        else
            general.Push(it)
    }

    ; Spacer + full-width header keep mnemonic prompts visually separate inside Prompts (same category).
    mnemonicBanner := "  ━━━  Mnemonic technique (MyNotes)  ━━━"
    newPromptSlice := []
    if (tech.Length = 0) {
        for it in promptItems
            newPromptSlice.Push(it)
    } else if (general.Length = 0) {
        newPromptSlice.Push({ category: "Prompts", char: "", text: " ", isEmpty: true, isSectionSpacer: true })
        newPromptSlice.Push({ category: "Prompts", char: "", text: mnemonicBanner, isEmpty: true, isSectionHeader: true })
        for it in tech
            newPromptSlice.Push(it)
    } else {
        for it in general
            newPromptSlice.Push(it)
        newPromptSlice.Push({ category: "Prompts", char: "", text: " ", isEmpty: true, isSectionSpacer: true })
        newPromptSlice.Push({ category: "Prompts", char: "", text: mnemonicBanner, isEmpty: true, isSectionHeader: true })
        for it in tech
            newPromptSlice.Push(it)
    }

    newRebuilt := []
    inserted := false
    for it in rebuilt {
        if (!IsObject(it) || !it.HasProp("category") || it.category != "Prompts") {
            newRebuilt.Push(it)
            continue
        }
        if (!inserted) {
            inserted := true
            for np in newPromptSlice
                newRebuilt.Push(np)
        }
    }
    rebuilt.Length := 0
    for x in newRebuilt
        rebuilt.Push(x)
}

; =============================================================================
; ShowHotstringSelector()
; =============================================================================
; PURPOSE: Displays the hotstring selector modal GUI and enables character-based hotkeys.
;          GUI shows categorized list of available actions with their assigned characters.
;
; PROCESS:
;   1. Closes existing selector if already open
;   2. Builds character-to-action mappings via BuildHotstringCharMap()
;   3. Validates that at least one action is available (shows tray tip if none)
;   4. Gets categorized hotstring data via GetCategorizedHotstrings()
;   5. Calculates optimal GUI size based on monitor configuration
;   6. Creates and displays GUI with categorized action list
;   7. Enables hotkeys for all assigned characters
;   8. Enables Escape key handler for cancellation
;
; GUI FEATURES:
;   - Responsive layout: Adapts to monitor orientation (landscape/portrait)
;   - Dual-column layout for landscape monitors
;   - Single-column layout for portrait monitors
;   - Categories displayed in order: Prompts → General → Projects → Files & Links → Macros
;
; RETURNS: None (void function)
; SIDE EFFECTS: Sets g_HotstringSelectorActive = true, creates GUI object, enables hotkeys
; =============================================================================
ShowHotstringSelector() {
    global g_HotstringSelectorGui, g_HotstringSelectorActive, g_HotstringCharMap
    global g_HotstringHotkeyHandlers, g_HotstringCategories
    global g_HS_SelectorOpenFile, g_HS_SelectorCloseRequestFile, g_HS_SelectorCloseCheckTimer
    ; Close existing GUI if open
    if (g_HotstringSelectorActive && IsObject(g_HotstringSelectorGui)) {
        CleanupHotstringSelector()
        Sleep 50
    }

    ; In-process mutual exclusion: if project selector is active in this process, close it first
    try {
        if (IsSet(g_ProjectSelectorActive) && g_ProjectSelectorActive && IsSet(CleanupProjectSelector)) {
            CleanupProjectSelector()
            Sleep 50
        }
    } catch {
        ; Ignore failures - hotstring selector should still open
    }

    ; Cross-process safety: if WindowManagement project selector is open in another process,
    ; request it to close via the existing sentinel file mechanism.
    try {
        sentinel := A_ScriptDir "\.cursor\wm_selector_open"
        if (FileExist(sentinel)) {
            closeReq := A_ScriptDir "\.cursor\wm_selector_close_request"
            try FileAppend "", closeReq
            catch {
            }
            Sleep 50
        }
    } catch {
        ; Ignore IPC failures - hotstring selector should still open
    }

    ; Build character mapping
    g_HotstringCharMap := BuildHotstringCharMap()

    ; Check if we have any items to display (hotstrings, quick open files, or macros)
    global g_QuickOpenFileCharMap, g_MacroCharMap
    hasItems := (g_HotstringCharMap.Count > 0) || (g_QuickOpenFileCharMap.Count > 0) || (g_MacroCharMap.Count > 0)
    if (!hasItems) {
        ; Use tray notification to avoid stealing focus/closing other palettes
        TrayTip("Utility Selector", "No items found.", "IconX")
        SetTimer(() => TrayTip(), -5000)  ; Auto-hide after ~5s
        return
    }

    ; Get categorized hotstrings
    categorized := GetCategorizedHotstrings()

    ; =============================================================================
    ; Dynamic Modal UI Adaptation Based on Monitor Configuration
    ;
    ; Monitor dataset (reference only; UI uses live work area from the active window):
    ; {
    ;   "monitor_dataset": [
    ;     {
    ;       "id": 1,
    ;       "resolution": "1920x1080",
    ;       "orientation": "landscape",
    ;       "scale": "125%",
    ;       "ui_strategy": "dual_column_wide"
    ;     },
    ;     {
    ;       "id": 2,
    ;       "resolution": "3840x2160",
    ;       "orientation": "landscape",
    ;       "scale": "150%",
    ;       "ui_strategy": "dual_column_max_width_constrained"
    ;     },
    ;     {
    ;       "id": 3,
    ;       "resolution": "1080x1920",
    ;       "orientation": "portrait",
    ;       "scale": "100%",
    ;       "ui_strategy": "single_column_vertical_stretch"
    ;     },
    ;     {
    ;       "id": 4,
    ;       "resolution": "1080x1920",
    ;       "orientation": "portrait",
    ;       "scale": "100%",
    ;       "ui_strategy": "single_column_vertical_stretch"
    ;     }
    ;   ],
    ;   "instruction_logic": {
    ;     "landscape_rule": "Apply two-column layout; prioritize width expansion.",
    ;     "portrait_rule": "Apply single-column layout; prioritize height expansion."
    ;   }
    ; }
    ; =============================================================================

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

    ; Detect monitor orientation: portrait (height > width) vs landscape (width >= height)
    isPortrait := (monitorHeight > monitorWidth)

    ; Create GUI (match Win+Alt+Shift+C AI Model Selector visual style)
    ; Create non-activating GUI so PowerToys Command Palette stays open
    g_HotstringSelectorGui := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x08000000", "Utility Shortcuts")
    g_HotstringSelectorGui.BackColor := "1E1E2E"
    g_HotstringSelectorGui.MarginX := 14
    g_HotstringSelectorGui.MarginY := 10
    ; Use slightly smaller font for compact display; Segoe UI to match C menu
    fontSize := (monitorHeight < 800) ? 9 : 9
    g_HotstringSelectorGui.SetFont("s" . fontSize . " cCDD6F4", "Segoe UI")

    ; Build reverse map: expansion -> character (legacy global mapping; still used elsewhere)
    expansionToChar := Map()
    for char, expansion in g_HotstringCharMap {
        expansionToChar[expansion] := char
    }

    ; Track which selector characters belong to the Prompts category (non-empty expansions only).
    ; Used to restrict the L-modifier redirect behavior to Prompts only.
    global g_HotstringPromptCharMap
    g_HotstringPromptCharMap := Map()
    try {
        global g_UtilityHotstringCharMapByCategory
        if (IsObject(g_UtilityHotstringCharMapByCategory) && g_UtilityHotstringCharMapByCategory.Has("Prompts")) {
            for ch, exp in g_UtilityHotstringCharMapByCategory["Prompts"] {
                try {
                    if (exp != "")
                        g_HotstringPromptCharMap[ch] := true
                } catch {
                }
            }
        }
    } catch {
        ; Ignore prompt-char tracking failures (selector still works normally)
    }

    ; Build items list grouped by category for two-column layout
    ; Collect all items first, then format in two columns
    hotstringCount := 0
    allItems := []  ; Array of {category, char, text, isEmpty}

    ; Build a map of character index to hotstring/file info
    charIndexToHotstring := Map()
    charIndex := 1
    for category in g_HotstringCategories {
        for item in categorized[category] {
            if (charIndex <= g_HotstringCharSequence.Length) {
                charIndexToHotstring[charIndex] := { hotstring: item, category: category }
            }
            charIndex++
        }
    }

    ; First, add explicitly assigned macros at their character positions
    global g_MacroCharMap
    for char, macroFunc in g_MacroCharMap {
        ; Find the macro entry for this function
        macroEntry := ""
        for macro in g_Macros {
            if (macro.func = macroFunc) {
                macroEntry := macro
                break
            }
        }
        if (macroEntry != "") {
            ; Find the index of this character in the sequence
            charIndexInSeq := 0
            for idx, seqChar in g_HotstringCharSequence {
                if (seqChar = char) {
                    charIndexInSeq := idx
                    break
                }
            }
            if (charIndexInSeq > 0) {
                itemText := "[" . char . "] > " . macroEntry.title
                hotstringCount++
                allItems.Push({ category: "Macros", char: char, text: itemText, isEmpty: false, explicitIndex: charIndexInSeq })
            }
        }
    }

    ; Collect all items with their categories (sequential assignment for non-explicit items)
    currentCharIndex := 1
    for category in g_HotstringCategories {
        ; Calculate how many character slots belong to this category
        categorySlotCount := categorized[category].Length

        if (categorySlotCount > 0 || currentCharIndex <= g_HotstringCharSequence.Length) {
            ; Collect all character slots for this category (including empty ones)
            loop categorySlotCount {
                if (currentCharIndex <= g_HotstringCharSequence.Length) {
                    ; Skip reserved empty char if set so it always shows as (empty)
                    while (currentCharIndex <= g_HotstringCharSequence.Length && g_ReservedEmptyChar != "" &&
                        g_HotstringCharSequence[currentCharIndex] = g_ReservedEmptyChar)
                        currentCharIndex++
                    if (currentCharIndex > g_HotstringCharSequence.Length)
                        break
                    char := g_HotstringCharSequence[currentCharIndex]

                    ; Skip if this character is already explicitly assigned to a macro
                    if (g_MacroCharMap.Has(char)) {
                        currentCharIndex++
                        continue
                    }

                    itemText := ""
                    isEmpty := false

                    ; Check if this character has a hotstring assigned
                    if (charIndexToHotstring.Has(currentCharIndex)) {
                        hsInfo := charIndexToHotstring[currentCharIndex]
                        hs := hsInfo.hotstring

                        ; Skip macros that have explicit char assignments
                        if (hsInfo.category = "Macros" && hs.HasProp("char") && hs.char != "") {
                            ; This macro has an explicit assignment, skip it here
                            currentCharIndex++
                            continue
                        }

                        ; Use title if available (for all categories including quick open files), otherwise use preview text
                        if (hs.HasProp("title") && hs.title != "") {
                            itemText := "[" . char . "] > " . hs.title
                            hotstringCount++
                        } else if (hs.HasProp("expansion") && hs.expansion != "") {
                            preview := GetPreviewText(hs.expansion)
                            itemText := "[" . char . "] > " . preview
                            hotstringCount++
                        } else {
                            ; Empty placeholder slot
                            itemText := "[" . char . "] > (empty)"
                            isEmpty := true
                        }
                    } else {
                        ; Character slot exists but no hotstring assigned
                        if (char = "l") {
                            itemText :=
                                "[L] > Gemini: L = arm; L+L = open Gemini + paste first snippet (or Ctrl+Alt+Win+L)"
                            isEmpty := false
                        } else {
                            itemText := "[" . char . "] > (empty)"
                            isEmpty := true
                        }
                    }

                    topCategory := UtilitySelector_MapInternalCategoryToTop(category)
                    allItems.Push({ category: topCategory, char: char, text: itemText, isEmpty: isEmpty })
                    currentCharIndex++
                }
            }
        }
    }

    ; Sort allItems by explicitIndex (if exists) or sequential position, then by category order
    ; Items with explicitIndex should be at their explicit position
    sortedItems := []
    for idx, seqChar in g_HotstringCharSequence {
        ; First check for explicitly assigned items at this position
        found := false
        for item in allItems {
            if (item.HasProp("explicitIndex") && item.explicitIndex = idx) {
                sortedItems.Push(item)
                found := true
                break
            }
        }
        ; If not found as explicit, check for sequential items
        if (!found) {
            for item in allItems {
                if (!item.HasProp("explicitIndex") && item.char = seqChar) {
                    ; Check if this item was already added
                    alreadyAdded := false
                    for added in sortedItems {
                        if (added.char = item.char && added.text = item.text) {
                            alreadyAdded := true
                            break
                        }
                    }
                    if (!alreadyAdded) {
                        sortedItems.Push(item)
                        break
                    }
                }
            }
        }
    }

    allItems := sortedItems

    ; Remove empty placeholder slots (no Unassigned category in the revised hierarchy)
    filtered := []
    for item in allItems {
        if (!item.isEmpty)
            filtered.Push(item)
    }
    allItems := filtered

    ; -------------------------------------------------------------------------
    ; Utility Shortcuts rendering: build items using explicit/per-category maps
    ; -------------------------------------------------------------------------
    ; The legacy block above assigns display chars by sequential slot, which can
    ; differ from mnemonic explicit chars (e.g., Projects 'a' / '0'). For the
    ; hierarchical selector, rebuild the item list from the category-scoped
    ; mapping so display + hotkeys match the selected category.
    try {
        global g_UtilityHotstringCharMapByCategory, g_QuickOpenFileCharMap, g_MacroCharMap, g_HotstringCharSequence

        charOrder := Map()
        for idx, c in g_HotstringCharSequence
            charOrder[c] := idx

        rebuilt := []
        seen := Map() ; key = category "|" char

        BuildExpansionToChar(catMap) {
            m := Map()
            try {
                for ch, exp in catMap
                    m[exp] := ch
            } catch {
            }
            return m
        }

        AddItem(cat, ch, titleText, seenRef, rebuiltRef, trigger := "") {
            if (ch = "" || titleText = "")
                return
            key := cat . "|" . ch
            if (seenRef.Has(key))
                return
            seenRef[key] := true
            row := { category: cat, char: ch, text: "[" . ch . "] > " . titleText, isEmpty: false, trigger: trigger }
            rebuiltRef.Push(row)
        }

        ; Hotstrings (text expansions) by category using category-scoped maps
        for cat in ["Prompts", "Projects", "General", "Hotstrings"] {
            if (!categorized.Has(cat))
                continue
            catMap := (IsObject(g_UtilityHotstringCharMapByCategory) && g_UtilityHotstringCharMapByCategory.Has(cat)) ?
                g_UtilityHotstringCharMapByCategory[cat] : Map()
            expToChar := BuildExpansionToChar(catMap)

            for hs in categorized[cat] {
                try {
                    if (!hs.HasProp("expansion") || hs.expansion = "")
                        continue

                    ch := (hs.HasProp("char") && hs.char != "") ? hs.char : expToChar.Get(hs.expansion, "")
                    if (ch = "")
                        continue

                    titleText := ""
                    if (hs.HasProp("title") && hs.title != "")
                        titleText := hs.title
                    else
                        titleText := GetPreviewText(hs.expansion)

                    tr := hs.HasProp("trigger") ? hs.trigger : ""
                    AddItem(cat, ch, titleText, seen, rebuilt, tr)
                } catch {
                }
            }
        }

        ; Files & Links show under Links in Utility menu
        filePathToChar := Map()
        try {
            for ch, fp in g_QuickOpenFileCharMap
                filePathToChar[fp] := ch
        } catch {
        }
        if (categorized.Has("Files & Links")) {
            for fileEntry in categorized["Files & Links"] {
                try {
                    ch := filePathToChar.Get(fileEntry.filePath, "")
                    if (ch = "")
                        continue
                    titleText := fileEntry.HasProp("title") ? fileEntry.title : ""
                    if (titleText = "")
                        titleText := fileEntry.filePath
                    AddItem("Links", ch, titleText, seen, rebuilt)
                } catch {
                }
            }
        }

        ; Macros
        funcToChar := Map()
        try {
            for ch, fn in g_MacroCharMap
                funcToChar[fn] := ch
        } catch {
        }
        if (categorized.Has("Macros")) {
            for macroEntry in categorized["Macros"] {
                try {
                    ch := (macroEntry.HasProp("char") && macroEntry.char != "") ? macroEntry.char : funcToChar.Get(
                        macroEntry.func, "")
                    if (ch = "")
                        continue
                    titleText := macroEntry.HasProp("title") ? macroEntry.title : ""
                    if (titleText = "")
                        titleText := "(macro)"
                    AddItem("Macros", ch, titleText, seen, rebuilt)
                } catch {
                }
            }
        }

        ; Sort by character order for a consistent layout
        try {
            rebuilt.Sort((a, b) => (charOrder.Get(a.char, 9999) = charOrder.Get(b.char, 9999)) ?
                (a.category < b.category ? -1 : 1) :
                (charOrder.Get(a.char, 9999) < charOrder.Get(b.char, 9999) ? -1 : 1))
        } catch {
        }

        try {
            UtilitySelector_ReorderPromptsMnemonicsSection(&rebuilt)
        } catch {
        }

        allItems := rebuilt
    } catch {
        ; Fallback to legacy list if rebuild fails
    }

    ; Helper function to pad string to specified width
    PadString(str, width) {
        len := StrLen(str)
        if (len >= width)
            return str
        padding := width - len
        spaces := ""
        loop padding {
            spaces .= " "
        }
        return str . spaces
    }

    ; Helper function to center string in specified width
    CenterString(str, width) {
        len := StrLen(str)
        if (len >= width)
            return str
        padding := (width - len) // 2
        leftSpaces := ""
        rightSpaces := ""
        loop padding {
            leftSpaces .= " "
        }
        loop (width - len - padding) {
            rightSpaces .= " "
        }
        return leftSpaces . str . rightSpaces
    }

    ; Helper function to create separator line
    CreateSeparator(width) {
        separator := ""
        loop width {
            separator .= "─"
        }
        return separator
    }

    ; Cache UI data for hierarchical selector refresh
    global g_UtilitySelectorAllItems, g_UtilitySelectorIsPortrait, g_UtilitySelectorMonitor
    g_UtilitySelectorAllItems := allItems
    g_UtilitySelectorIsPortrait := isPortrait
    g_UtilitySelectorMonitor := Map("left", monitorLeft, "top", monitorTop, "width", monitorWidth, "height",
        monitorHeight)

    ; Always open in top-level category screen
    global g_UtilitySelectorMode, g_UtilitySelectorCategory
    g_UtilitySelectorMode := "top"
    g_UtilitySelectorCategory := ""

    displayText := UtilitySelector_BuildDisplayText(isPortrait)
    ; Calculate text control height based on actual content (number of lines)
    ; Count actual lines in displayText (including empty lines for spacing)
    lineCount := 1  ; Start at 1 (first line doesn't have a newline before it)
    loop parse, displayText, "`n" {
        lineCount++
    }
    ; Calculate height: ~18 pixels per line (Consolas 9pt in RichEdit with internal line padding)
    lineHeight := 18
    textControlHeight := lineCount * lineHeight
    ; Ensure minimum and maximum bounds
    minHeight := 150

    ; Adjust sizing based on orientation
    if (isPortrait) {
        ; PORTRAIT: Prioritize height expansion, use narrower width
        ; Use more vertical space for portrait monitors (up to 85% of height)
        maxHeightPercent := 0.85
        maxHeight := Floor(monitorHeight * maxHeightPercent)
        if (textControlHeight < minHeight)
            textControlHeight := minHeight
        if (textControlHeight > maxHeight)
            textControlHeight := maxHeight

        ; Narrower width for portrait (optimized for vertical scrolling)
        baseWidth := (monitorWidth < 800) ? 400 : (monitorWidth < 1200) ? 500 : 500
        ; Ensure we don't exceed monitor width with margins
        if (baseWidth > monitorWidth - 40)
            baseWidth := monitorWidth - 40
    } else {
        ; LANDSCAPE: Prioritize width expansion, use two-column layout
        ; Use adaptive max height: 85% for large monitors, 90% for small monitors
        maxHeightPercent := (monitorHeight < 800) ? 0.90 : 0.85
        maxHeight := Floor(monitorHeight * maxHeightPercent)
        if (textControlHeight < minHeight)
            textControlHeight := minHeight
        if (textControlHeight > maxHeight)
            textControlHeight := maxHeight

        ; Wide width for landscape (two-column layout)
        ; Further reduced width to eliminate empty space: 650px minimum, scale up to 1000px based on monitor width
        baseWidth := (monitorWidth < 1200) ? 650 : (monitorWidth < 1920) ? 800 : 1000
        ; Ensure we don't exceed monitor width with margins
        if (baseWidth > monitorWidth - 40)
            baseWidth := monitorWidth - 40
    }
    textControlWidth := baseWidth - 20  ; Account for margins

    ; Title and separator (compact)
    g_HotstringSelectorGui.SetFont("s11 cCDD6F4 Bold", "Segoe UI")
    global g_UtilitySelectorTitleCtrl
    g_UtilitySelectorTitleCtrl := g_HotstringSelectorGui.Add("Text", "w" . textControlWidth . " Center",
        "Utility Shortcuts")
    g_HotstringSelectorGui.Add("Text", "w" . textControlWidth . " h1 Background45475A")
    g_HotstringSelectorGui.SetFont("s" . fontSize . " cCDD6F4", "Segoe UI")

    ; Cache base font size for RichEdit rendering in refresh
    global g_UtilitySelectorFontSize
    g_UtilitySelectorFontSize := fontSize

    ; Enable vertical scrolling for long content (RichEdit so we can style mnemonic letters)
    global g_UtilitySelectorEditCtrl
    MnemonicRich_EnsureDll()
    g_UtilitySelectorEditCtrl := g_HotstringSelectorGui.Add("Custom",
        "ClassRichEdit50W w" . textControlWidth . " h" . textControlHeight
        . " +0x44 -E0x200 +VScroll -HScroll -Border Background1E1E2E")
    try MnemonicRich_Render(g_UtilitySelectorEditCtrl, UtilitySelector_BuildDisplayRich(isPortrait), fontSize, 6,
        "Consolas",
        "CDD6F4", "1E1E2E")
    g_HotstringSelectorGui.SetFont("s9 c89B4FA", "Segoe UI")
    global g_UtilitySelectorFooterCtrl
    g_UtilitySelectorFooterCtrl := g_HotstringSelectorGui.Add("Text", "w" . textControlWidth . " Center",
        "Press Escape to close.")

    ; Total height: top-margin + title(s11) + gap + separator + gap + edit + gap + footer + bottom-margin
    totalHeight := 10 + 24 + 10 + 1 + 10 + textControlHeight + 10 + 18 + 10
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

    ; Cross-process IPC: mark Hotstring Selector as open and start close-request timer
    try {
        DirCreate(A_ScriptDir "\.cursor")
        FileAppend("", g_HS_SelectorOpenFile)
    } catch {
    }
    g_HS_SelectorCloseCheckTimer := SetTimer(Utils_CheckHotstringSelectorCloseRequest, 120)

    ; Bind top-level hotkeys (1-6) + Escape; category view binds are applied when user selects a category.
    UtilitySelector_RebindHotkeys()
    SetTimer(HotstringSelector_AutoCloseIfIdle, -3000)
}

; =============================================================================
; Hotkey Handler: Windows + Alt + Shift + U (#!+U)
; =============================================================================
; PURPOSE: Toggles the hotstring selector GUI on/off.
;
; BEHAVIOR:
;   - If selector is currently open: Closes selector via CleanupHotstringSelector()
;   - If selector is closed: Opens selector via ShowHotstringSelector()
;
; TECHNICAL NOTE: The character sequence displayed in the GUI must remain consistent
;                  and list every slot in order, even when empty, to ensure downstream
;                  AI systems and debugging tools can reliably parse the full character set.
; =============================================================================
#!+U::
{
    global g_HotstringSelectorActive, g_HotstringSelectorGui

    ; Toggle behavior: Close if open, open if closed
    if (g_HotstringSelectorActive && IsObject(g_HotstringSelectorGui)) {
        CleanupHotstringSelector()
    } else {
        ShowHotstringSelector()
    }
}

; Ctrl+Alt+Win+L - direct D2C submit path (paste + Enter, then monitor)
^!#L:: D2C_FlowManager.GetInstance().StartFromHotstring()

; Ctrl+Alt+Win+4 - Gemini tab 1/2 toggle + banner
^!#4::
{
    global g_GeminiToggleTab

    geminiHwnd := 0
    try {
        for hwnd in WinGetList("ahk_exe chrome.exe") {
            try {
                if InStr(WinGetTitle("ahk_id " hwnd), "gemini", false) {
                    geminiHwnd := hwnd
                    break
                }
            } catch {
            }
        }
    } catch {
    }

    if (!geminiHwnd)
        return

    WinActivate("ahk_id " geminiHwnd)
    if (!WinWaitActive("ahk_id " geminiHwnd, , 2))
        return

    Sleep(120)
    uia := UIA_Browser("ahk_id " geminiHwnd)
    tabInfo := GetChromeActiveTabIndex(uia)
    if (!tabInfo) {
        Sleep(150)
        tabInfo := GetChromeActiveTabIndex(uia)
    }
    targetTab := (tabInfo && tabInfo.index == 1) ? 2 : 1
    g_GeminiToggleTab := targetTab
    Send("^" . targetTab)
    ShowSingleCharTabBanner_Utils(targetTab)
    Sleep(200)

    tabInfoAfter := GetChromeActiveTabIndex(uia)
    if (!tabInfoAfter) {
        Sleep(100)
        tabInfoAfter := GetChromeActiveTabIndex(uia)
    }
    tabOk := tabInfoAfter && tabInfoAfter.index == targetTab
    if (!tabOk)
        ShowCenteredOverlay_Utils("❌ Shortcut execution failed", 2000, BANNER_ACCENT_ERROR)
}

; Ctrl+Alt+Win+2..8 - same macros as HotStrings panel (Win+Alt+Shift+U); secondary triggers only
^!#2:: QuickUpdateScripts()
^!#3:: ToggleOutlookAndTeams()
^!#5:: CleanClipboard()
; Same macro; InputLevel 10 + hook so chord wins over other low-level handlers (optional ghosting fallback: ^!#j).
#InputLevel 10
#UseHook
^!#7:: MarkLastClipAsFavorite()
^!#j:: MarkLastClipAsFavorite()
#UseHook False
#InputLevel 0
^!#8:: DesktopToRecycle_Trigger()
; Ctrl+Alt+Win+9 / +B - Handy Cohere Portuguese / English (g_HandyAiModels slots 4 and 3)
^!#9:: ExecuteHandyAiModelSelection(HANDY_AI_SLOT_COHERE_PORTUGUESE)
^!#b:: ExecuteHandyAiModelSelection(HANDY_AI_SLOT_COHERE_ENGLISH)

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
global g_FocusModeTrackedWindow := 0  ; window handle that was active when focus mode was enabled
global g_FocusModeAnchorHwnd := 0  ; window that may relocate blackout when it moves to another monitor
global g_FocusModeMonitorTimer := false  ; timer for monitoring window focus changes
global g_FocusModeKeepMonitorFile := A_ScriptDir "\.cursor\focus_mode_keep_monitor"
global g_FocusModeDisableRequestFile := A_ScriptDir "\.cursor\focus_mode_disable_request"

FocusMode_WriteKeepMonitorState(mon) {
    global g_FocusModeKeepMonitorFile
    if (!mon)
        return
    try {
        cursorDir := A_ScriptDir "\.cursor"
        if !DirExist(cursorDir)
            DirCreate(cursorDir)
        try FileDelete(g_FocusModeKeepMonitorFile)
        FileAppend(String(mon), g_FocusModeKeepMonitorFile, "UTF-8")
    } catch {
    }
}

FocusMode_ClearKeepMonitorState() {
    global g_FocusModeKeepMonitorFile, g_FocusModeDisableRequestFile
    try FileDelete(g_FocusModeKeepMonitorFile)
    catch {
    }
    try FileDelete(g_FocusModeDisableRequestFile)
    catch {
    }
}

FocusMode_ReadKeepMonitorFromFile() {
    global g_FocusModeKeepMonitorFile
    try {
        if !FileExist(g_FocusModeKeepMonitorFile)
            return 0
        t := Trim(FileRead(g_FocusModeKeepMonitorFile, "UTF-8"))
        if (t != "" && t is Integer)
            return Integer(t)
    } catch {
    }
    return 0
}

FocusMode_RequestDisableCrossProcess() {
    global g_FocusModeDisableRequestFile
    try {
        cursorDir := A_ScriptDir "\.cursor"
        if !DirExist(cursorDir)
            DirCreate(cursorDir)
        FileAppend("", g_FocusModeDisableRequestFile, "UTF-8")
    } catch {
    }
}

FocusMode_CheckCrossProcessRequests() {
    global g_FocusModeDisableRequestFile
    try {
        if !FileExist(g_FocusModeDisableRequestFile)
            return false
        FileDelete(g_FocusModeDisableRequestFile)
    } catch {
        return false
    }
    DisableFocusMode()
    return true
}

GetActiveMonitorIndex() {
    hwnd := WinExist("A")
    if (!hwnd) {
        return 0
    }

    rect := Buffer(16, 0)
    if (!DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect)) {
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
            return A_Index
        }
    }

    return 0
}

FocusMode_DestroyOverlays() {
    global g_FocusModeOverlays

    if (!IsSet(g_FocusModeOverlays))
        g_FocusModeOverlays := []
    for overlay in g_FocusModeOverlays {
        try {
            if (IsObject(overlay) && overlay.Hwnd)
                overlay.Destroy()
        } catch {
        }
    }
    g_FocusModeOverlays := []
}

FocusMode_BuildOverlays(keepMonitorIndex) {
    global g_FocusModeOverlays

    if (!keepMonitorIndex)
        return
    FocusMode_DestroyOverlays()

    monitorCount := MonitorGetCount()
    loop monitorCount {
        i := A_Index
        if (i = keepMonitorIndex)
            continue

        MonitorGet(i, &l, &t, &r, &b)
        w := r - l
        h := b - t
        if (w <= 0 || h <= 0)
            continue

        overlay := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x20") ; WS_EX_TRANSPARENT => click-through
        overlay.Opt("-DPIScale")
        overlay.BackColor := "000000"
        overlay.Show("NA x" l " y" t " w" w " h" h)
        g_FocusModeOverlays.Push(overlay)
    }
}

FocusMode_SetKeepMonitor(keepMonitorIndex) {
    global g_FocusModeOn, g_FocusModeActiveMonitor, g_FocusModeOverlays, g_FocusModeTrackedWindow

    if (!g_FocusModeOn || !keepMonitorIndex)
        return
    try {
        if (keepMonitorIndex < 1 || keepMonitorIndex > MonitorGetCount())
            return
    } catch {
        return
    }

    if (g_FocusModeActiveMonitor = keepMonitorIndex) {
        hasLiveOverlay := false
        for overlay in g_FocusModeOverlays {
            if (IsObject(overlay) && overlay.Hwnd && WinExist("ahk_id " . overlay.Hwnd)) {
                hasLiveOverlay := true
                break
            }
        }
        if (hasLiveOverlay)
            return
    }

    g_FocusModeActiveMonitor := keepMonitorIndex
    FocusMode_BuildOverlays(keepMonitorIndex)
    FocusMode_WriteKeepMonitorState(keepMonitorIndex)
    g_FocusModeTrackedWindow := WinExist("A")
}

EnableFocusMode(keepMonitorIndex := 0) {
    global g_FocusModeOn, g_FocusModeActiveMonitor, g_FocusModeOverlays, g_FocusModeTrackedWindow,
        g_FocusModeAnchorHwnd

    ; Ensure globals are initialized (avoid "variable has not been assigned" on first call)
    if (!IsSet(g_FocusModeOverlays))
        g_FocusModeOverlays := []
    if (!IsSet(g_FocusModeOn))
        g_FocusModeOn := false

    ; Flag plus overlay refs / WinExist: trust non-empty array like ToggleFocusMode (WinExist can miss fullscreen Gui overlays).
    hasOverlayRefs := IsObject(g_FocusModeOverlays) && g_FocusModeOverlays.Length > 0
    hasLiveOverlay := false
    if (hasOverlayRefs) {
        for overlay in g_FocusModeOverlays {
            if (IsObject(overlay) && overlay.Hwnd && WinExist("ahk_id " . overlay.Hwnd)) {
                hasLiveOverlay := true
                break
            }
        }
    }

    if (g_FocusModeOn || hasOverlayRefs) {
        if (hasOverlayRefs && !g_FocusModeOn) {
            DisableFocusMode()
        } else if (g_FocusModeOn && hasLiveOverlay) {
            return
        } else if (g_FocusModeOn && hasOverlayRefs && !hasLiveOverlay) {
            DisableFocusMode()
        }
        ; g_FocusModeOn && !hasOverlayRefs: inconsistent — fall through and recreate overlays
    }

    activeMon := keepMonitorIndex
    if (!activeMon)
        activeMon := GetActiveMonitorIndex()
    if (!activeMon) {
        return
    }

    g_FocusModeActiveMonitor := activeMon
    g_FocusModeTrackedWindow := WinExist("A")
    g_FocusModeAnchorHwnd := g_FocusModeTrackedWindow
    FocusMode_WriteKeepMonitorState(activeMon)
    StartFocusModeWindowMonitor()
    FocusMode_BuildOverlays(activeMon)
    g_FocusModeOn := true
}

DisableFocusMode() {
    global g_FocusModeOn, g_FocusModeActiveMonitor, g_FocusModeOverlays, g_FocusModeTrackedWindow,
        g_FocusModeMonitorTimer, g_FocusModeAnchorHwnd

    ; Stop monitoring window focus changes
    StopFocusModeWindowMonitor()
    StopPdfFocusMonitor()
    FocusMode_ClearKeepMonitorState()

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
    g_FocusModeTrackedWindow := 0
    g_FocusModeAnchorHwnd := 0
    g_FocusModeOn := false
}

ToggleFocusMode() {
    global g_FocusModeOn, g_FocusModeOverlays

    ; Treat non-empty overlay array as active even if WinExist fails on some setups (Dpi/multi-monitor),
    ; so #!+Y always tears down auto-blackout state instead of calling EnableFocusMode() by mistake.
    hasOverlayRefs := IsObject(g_FocusModeOverlays) && g_FocusModeOverlays.Length > 0
    actualState := g_FocusModeOn || hasOverlayRefs

    if (actualState) {
        DisableFocusMode()
    } else {
        EnableFocusMode()
        try {
            if (MonitorGetCount() > 1) {
                hwnd := WinExist("A")
                if (hwnd)
                    StartPdfFocusMonitor(hwnd, "Immediate")
            }
        } catch {
        }
    }
}

; Keep tracked HWND in sync while foreground stays on the keep-clear monitor.
FocusModeWindowMonitor(*) {
    global g_FocusModeOn, g_FocusModeActiveMonitor, g_FocusModeTrackedWindow

    if (!g_FocusModeOn)
        return

    fg := WinExist("A")
    if (!fg)
        return

    fgMon := GetActiveMonitorIndex()
    if (fgMon && g_FocusModeActiveMonitor && fgMon = g_FocusModeActiveMonitor)
        g_FocusModeTrackedWindow := fg
}

; Start monitoring window focus changes
StartFocusModeWindowMonitor() {
    global g_FocusModeMonitorTimer

    ; Stop any existing timer first
    StopFocusModeWindowMonitor()

    ; Start timer to check window focus every 200ms
    g_FocusModeMonitorTimer := SetTimer(FocusModeWindowMonitor, 200)
}

; Stop monitoring window focus changes
StopFocusModeWindowMonitor() {
    global g_FocusModeMonitorTimer

    ; First-call safety: global may be unset
    if (!IsSet(g_FocusModeMonitorTimer))
        g_FocusModeMonitorTimer := false

    if (g_FocusModeMonitorTimer) {
        try {
            SetTimer(g_FocusModeMonitorTimer, 0)  ; Disable timer
        } catch {
            ; Ignore errors
        }
        g_FocusModeMonitorTimer := false
    }
}

#!+Y::
{
    ToggleFocusMode()
}

; =============================================================================
; Print Screen with Chime
; Hotkey: Alt+PrintScreen
; Intercepts the hotkey to prevent other apps from handling it,
; manually triggers the screenshot, and plays a single chime
; =============================================================================
global g_LastPrintScreenSound := 0  ; Track last sound time for debouncing
global g_PrintScreenInProgress := false  ; Prevents recursion from Send

; Audio firewall for PrintScreen: Throttle sounds to prevent duplicates
; Uses Critical section to ensure atomic check-and-update (like dictation mode)
SafePlayPrintScreenSound() {
    Critical  ; Prevents thread interruption - ensures atomic check-and-update sequence
    global g_LastPrintScreenSound

    ; If less than 1000ms has passed since last sound, ignore this call
    if (A_TickCount - g_LastPrintScreenSound < 1000) {
        return
    }

    ; Update timestamp and play sound (if enabled)
    g_LastPrintScreenSound := A_TickCount
    ScriptSoundPlay(A_ScriptDir . "\sounds\print-screen.wav")
}

; Set higher InputLevel to ensure our handler processes before others
#InputLevel 10
!PrintScreen::  ; Removed ~ prefix to CONSUME the hotkey (prevents other apps from receiving it)
{
    global g_PrintScreenInProgress

    ; Prevent recursion: if we're already processing, skip (this handles Send retriggering)
    if (g_PrintScreenInProgress) {
        return
    }

    ; Set flag to prevent recursion from the Send below
    g_PrintScreenInProgress := true

    ; Manually send Alt+PrintScreen to Windows to trigger the screenshot
    ; SendInput is more reliable and won't retrigger our hotkey due to SendLevel
    SendInput("!{PrintScreen}")

    ; Brief delay to ensure screenshot is captured, then play single chime
    ; Uses Critical section to prevent duplicate sounds from concurrent handlers
    Sleep 10
    SafePlayPrintScreenSound()

    ; Reset flag after a brief delay to allow normal operation
    Sleep 100
    g_PrintScreenInProgress := false
}
#InputLevel 0

; True when SendEscape() must not inject Escape (Handy dictation/recording UI). Physical Escape uses Escape:: below.
IsHandyDictationEscapeSuppressed() {
    global g_DictationActive
    return g_DictationActive || WinActive("Recording ahk_exe handy.exe") || WinActive(
        "Recording Overlay ahk_exe handy.exe") || WinExist("Recording ahk_exe handy.exe") || WinExist(
            "Recording Overlay ahk_exe handy.exe")
}

; Helper to send Escape while respecting dictation state (no-op when suppressed; see IsHandyDictationEscapeSuppressed).
SendEscape(count := 1) {
    if (IsHandyDictationEscapeSuppressed()) {
        return
    }
    if (count > 1)
        Send("{Escape " . count . "}")
    else
        Send("{Escape}")
}

; =============================================================================
; Global Escape hotkey (registered via Hotkey(), not Escape:: label)
; Modals use Hotkey("Escape", modalFn, "On") which replaces this binding; Hotkey("Escape", "Off") leaves Escape
; unhandled until Utils_EnsureGlobalEscapeHotkey() runs (see AiModel cleanup, square selector, hotstring cleanup, etc.).
; =============================================================================

; Optional escape callback: when set (e.g. by WindowManagement for project selector), Utils runs it and consumes Escape.
global g_OnEscapePressed := ""
Utils_GlobalEscapeHandler(*) {
    global g_OnEscapePressed

    if (g_OnEscapePressed) {
        try {
            g_OnEscapePressed.Call()
        } catch {
        }
        return
    }

    try {
        sentinel := A_ScriptDir "\.cursor\wm_selector_open"
        if (FileExist(sentinel)) {
            closeReq := A_ScriptDir "\.cursor\wm_selector_close_request"
            try FileAppend "", closeReq
            catch {
            }
            return
        }
        minSentinel := A_ScriptDir "\.cursor\wm_minimized_list_open"
        if (FileExist(minSentinel)) {
            minCloseReq := A_ScriptDir "\.cursor\wm_minimized_list_close_request"
            try FileAppend "", minCloseReq
            catch {
            }
            return
        }
    } catch {
    }

    ; I10 hotkey: forward at send level 0 so this handler is not re-triggered by SendInput (SendLevel / #InputLevel).
    SendLevel 0
    SendInput "{Escape}"
}

Utils_EnsureGlobalEscapeHotkey() {
    ; I10: distinct from modals that use Hotkey("Escape", fn, "On") at default level; avoids replace ambiguity.
    ; Forward path uses SendLevel 0 so synthetic Escape does not re-enter this handler.
    try {
        Hotkey("Escape", Utils_GlobalEscapeHandler, "I10 On")
    } catch {
    }
}

; Initial registration (replaces legacy Escape:: label)
Utils_EnsureGlobalEscapeHotkey()

; =============================================================================
; Dictation Indicator - Red pulsing inner square with yellow border
; Anchored to the top-center of the active window (clamped inside); falls back to
; active monitor work area. Follows focus/window moves via pulse timer. Toggles with Win+Alt+Shift+0.
; =============================================================================

; Global variables for dictation indicator
global g_DictationActive := false
global g_DictationIndicatorGui := false
global g_DictationIndicatorText := false  ; Text control for status messages
global g_DictationPulseTimer := false
global g_DictationCheckTimer := false  ; Timer to check if Recording window still exists
global g_DictationPulseDirection := 1  ; 1 = fading in, -1 = fading out
global g_DictationPulseOpacity := 128  ; Current opacity (50-255)
global g_DictationFollowCache := ""  ; "x,y,w,h" when unchanged skip redundant Move (follow active window)
global g_DictationCompletionChimeScheduled := false  ; Flag to prevent multiple completion chimes
global g_LastDictationSoundTick := 0  ; Timestamp of last dictation sound to throttle audio output
global g_DictationStartSound := A_ScriptDir . "\sounds\speach-start.wav"
global g_DictationStopSound := A_ScriptDir . "\sounds\speach-finished.wav"
global g_PendingDictationAction := ""  ; Action to execute after transcription: "Paste" (reserved for future)
global g_PendingGeminiPromptAfterDictation := false  ; When set by ~#!+0 stop, show "Send to Gemini? Y (4s)" after completion
global g_D2C_DictationSubmitMenuCycleFinished := false  ; After V/E/N/timeout/F/O: block stray second StartFromDictation for this wave
global g_DictationGeminiConfirmBannerVisible := false  ; Guard: only one "Send to Gemini?" banner at a time
global g_KeepIndicatorVisible := false  ; Flag to keep indicator visible until paste action completes
global g_LastStateTransitionTick := 0  ; Timestamp of last state transition to prevent rapid re-detection
global g_DictationSoundPlayed := false  ; Atomic test-and-set: one start chime per session
global g_DictationStartClipboardText := "" ; Track clipboard content at start to detect changes
global g_DictationHotkeyOwnerHandle := 0 ; Named mutex handle for cross-process single-owner dictation hotkey
global g_DictationHotkeyIsOwner := false ; True only in the single process that owns dictation hotkey handling

; Ensure only one script process handles the dictation hotkey logic.
InitializeDictationHotkeyOwnership() {
    global g_DictationHotkeyOwnerHandle, g_DictationHotkeyIsOwner
    mutexName := "Local\D2C_Dictation_Hotkey_Owner"
    hMutex := DllCall("CreateMutex", "Ptr", 0, "Int", 0, "Str", mutexName, "Ptr")
    if (!hMutex) {
        g_DictationHotkeyIsOwner := false
        return
    }
    err := DllCall("GetLastError", "UInt")
    if (err = 183) { ; ERROR_ALREADY_EXISTS
        g_DictationHotkeyIsOwner := false
        DllCall("CloseHandle", "Ptr", hMutex)
        return
    }
    g_DictationHotkeyOwnerHandle := hMutex
    g_DictationHotkeyIsOwner := true
}

ReleaseDictationHotkeyOwnership(*) {
    global g_DictationHotkeyOwnerHandle
    if (g_DictationHotkeyOwnerHandle) {
        try DllCall("CloseHandle", "Ptr", g_DictationHotkeyOwnerHandle)
        g_DictationHotkeyOwnerHandle := 0
    }
}

InitializeDictationHotkeyOwnership()
OnExit(ReleaseDictationHotkeyOwnership)

; Debug logging helper for dictation workflow
LogDebug(sessionId, runId, hypothesisId, location, message, data := "") {
    logPath := A_ScriptDir "\.cursor\debug.log"
    timestamp := A_Now "." Format("{:03}", A_MSec)
    logEntry := Format(
        '{{"sessionId":"{}","runId":"{}","hypothesisId":"{}","location":"{}","message":"{}","timestamp":"{}","data":{}}}',
        sessionId, runId, hypothesisId, location, message, timestamp, data ? '"' . data . '"' : '""')
    try {
        FileAppend(logEntry . "`n", logPath)
    } catch {
        ; Silently ignore logging errors
    }
}

; Constants for dictation indicator
global DICTATION_SQUARE_SIZE := 150  ; Inner red area (was 50 base scaled up)
global DICTATION_BORDER_PX := 2      ; Yellow border for colorblind visibility (outside red)
global DICTATION_YELLOW_BORDER := "F1C40F"  ; Same hue family as BANNER_ACCENT_INTERMEDIATE
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

; Outer indicator size: yellow border + inner red square.
GetDictationIndicatorOuterSize() {
    global DICTATION_SQUARE_SIZE, DICTATION_BORDER_PX
    return DICTATION_SQUARE_SIZE + 2 * DICTATION_BORDER_PX
}

; Screen rect for the dictation indicator: prefer top-center inside the active window; else top-center of active monitor work area.
GetDictationIndicatorScreenRect(&outX, &outY, &outW, &outH) {
    global DICTATION_SQUARE_SIZE, DICTATION_BORDER_PX
    outW := GetDictationIndicatorOuterSize()
    outH := outW
    marginTop := 8
    hwnd := WinExist("A")
    if (hwnd) {
        try {
            if (WinGetMinMax("ahk_id " hwnd) = -1)
                hwnd := 0
        } catch {
            hwnd := 0
        }
    }
    if (hwnd) {
        rect := Buffer(16, 0)
        if (DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect)) {
            wl := NumGet(rect, 0, "int")
            wt := NumGet(rect, 4, "int")
            wr := NumGet(rect, 8, "int")
            wb := NumGet(rect, 12, "int")
            winW := wr - wl
            winH := wb - wt
            if (winW >= 8 && winH >= 8) {
                outX := wl + (winW - outW) // 2
                outY := wt + marginTop
                if (outX < wl + 2)
                    outX := wl + 2
                if (outY < wt + 2)
                    outY := wt + 2
                if (outX + outW > wr - 2)
                    outX := wr - outW - 2
                if (outY + outH > wb - 2)
                    outY := wb - outH - 2
                if (outX < wl)
                    outX := wl
                if (outY < wt)
                    outY := wt
                return true
            }
        }
    }
    mon := GetDictationActiveMonitor()
    MonitorGetWorkArea(mon, &ml, &mt, &mr, &mb)
    mw := mr - ml
    outX := ml + (mw - outW) // 2
    outY := mt + 12
    if (outX < ml)
        outX := ml
    if (outY < mt)
        outY := mt
    if (outX + outW > mr)
        outX := mr - outW
    if (outY + outH > mb)
        outY := mb - outH
    return true
}

; Reposition indicator when the active window moves or focus changes (called from pulse timer).
DictationIndicator_SyncPosition() {
    global g_DictationIndicatorGui, g_DictationFollowCache
    if (!IsObject(g_DictationIndicatorGui) || !g_DictationIndicatorGui.Hwnd)
        return
    GetDictationIndicatorScreenRect(&sx, &sy, &sw, &sh)
    key := sx . "," . sy . "," . sw . "," . sh
    if (key = g_DictationFollowCache)
        return
    g_DictationFollowCache := key
    try {
        g_DictationIndicatorGui.Show("NA x" . sx . " y" . sy . " w" . sw . " h" . sh)
    } catch {
    }
}

; Show or update the dictation indicator: yellow border + red inner, anchored to active window (or work area fallback).
ShowDictationIndicator() {
    global g_DictationIndicatorGui, g_DictationPulseOpacity, g_DictationFollowCache
    global DICTATION_SQUARE_SIZE, DICTATION_BORDER_PX, DICTATION_YELLOW_BORDER

    GetDictationIndicatorScreenRect(&squareX, &squareY, &outerW, &outerH)

    ; Check if indicator already exists
    if (IsObject(g_DictationIndicatorGui) && g_DictationIndicatorGui.Hwnd) {
        g_DictationFollowCache := ""
        DictationIndicator_SyncPosition()
        UpdateDictationIndicatorText("")
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
    g_DictationIndicatorGui.BackColor := DICTATION_YELLOW_BORDER
    g_DictationIndicatorGui.MarginX := 0
    g_DictationIndicatorGui.MarginY := 0

    ; Inner red fill, then status text on top (transparent over red)
    g_DictationIndicatorGui.Add("Text",
        "x" . DICTATION_BORDER_PX . " y" . DICTATION_BORDER_PX . " w" . DICTATION_SQUARE_SIZE . " h" .
        DICTATION_SQUARE_SIZE .
        " BackgroundFF0000", "")
    g_DictationIndicatorGui.SetFont("s14 cFFFFFF Bold", "Segoe UI")
    g_DictationIndicatorText := g_DictationIndicatorGui.Add("Text",
        "x" . DICTATION_BORDER_PX . " y" . DICTATION_BORDER_PX . " w" . DICTATION_SQUARE_SIZE . " h" .
        DICTATION_SQUARE_SIZE .
        " Center +BackgroundTrans", "")

    ; Reset pulse opacity
    g_DictationPulseOpacity := 128
    g_DictationFollowCache := ""

    ; Show the indicator without activating it
    g_DictationIndicatorGui.Show("NA x" . squareX . " y" . squareY . " w" . outerW . " h" . outerH)
    DictationIndicator_SyncPosition()
    ; Apply initial transparency defensively (GUI may have been destroyed concurrently)
    try {
        if (g_DictationIndicatorGui.Hwnd)
            WinSetTransparent(g_DictationPulseOpacity, g_DictationIndicatorGui)
    } catch {
        ; Ignore "target window not found" or similar errors
    }
}

; Update indicator text (for status messages)
UpdateDictationIndicatorText(message := "") {
    global g_DictationIndicatorGui, g_DictationIndicatorText

    if (IsObject(g_DictationIndicatorGui) && g_DictationIndicatorGui.Hwnd && IsObject(g_DictationIndicatorText)) {
        try {
            g_DictationIndicatorText.Value := message
        } catch {
            ; Ignore errors
        }
    }
}

; Hide and destroy the dictation indicator
HideDictationIndicator() {
    global g_DictationIndicatorGui, g_DictationIndicatorText, g_DictationFollowCache

    g_DictationFollowCache := ""
    if (IsObject(g_DictationIndicatorGui)) {
        try {
            if (g_DictationIndicatorGui.Hwnd) {
                g_DictationIndicatorGui.Destroy()
            }
        } catch {
            ; Ignore errors
        }
        g_DictationIndicatorGui := false
        g_DictationIndicatorText := false
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

    DictationIndicator_SyncPosition()

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

; Audio firewall: Throttle dictation sounds to prevent duplicates
; Enforces a minimum 1000ms gap between sounds regardless of how many times logic fires
SafePlayDictationSound(filePath) {
    Critical  ; Prevents thread interruption - ensures atomic check-and-update sequence
    global g_LastDictationSoundTick, g_DictationStartSound
    static lastStartSoundTick := 0

    ; Special handling for start sound: 7 second cooldown to prevent duplicates
    if (InStr(filePath, "speach-start.wav")) {
        if (A_TickCount - lastStartSoundTick < 7000) {
            return
        }
        lastStartSoundTick := A_TickCount
    } else {
        ; Standard 1 second cooldown for other sounds
        if (A_TickCount - g_LastDictationSoundTick < 1000) {
            return
        }
    }

    ; Update timestamp and play sound (if enabled)
    g_LastDictationSoundTick := A_TickCount
    if (FileExist(filePath)) {
        try {
            ScriptSoundPlay(filePath)
        } catch {
            ; Silently ignore playback failures (missing file, sync placeholder, format, etc.)
        }
    }
}

; Handler for clipboard changes during dictation completion
DictationClipboardHandler(DataType) {
    ; Remove handler immediately to prevent multiple triggers
    OnClipboardChange(DictationClipboardHandler, 0)

    ; Trigger completion logic immediately
    PlayDictationCompletionChime()
}

; Play completion chime after transcription finishes
PlayDictationCompletionChime(*) {
    global g_DictationCompletionChimeScheduled, g_PendingDictationAction,
        g_KeepIndicatorVisible, g_PendingGeminiPromptAfterDictation, g_D2C_DictationSubmitMenuCycleFinished

    ; Ensure clipboard handler is removed (safe to call even if already removed)
    try {
        OnClipboardChange(DictationClipboardHandler, 0)
    }

    ; Cancel fallback timer to prevent redundant calls
    SetTimer(PlayDictationCompletionChime, 0)

    ; CRITICAL: Test-and-set pattern - clear flag IMMEDIATELY to prevent duplicates
    ; Use Critical to ensure atomicity
    Critical "On"
    chimeShouldPlay := g_DictationCompletionChimeScheduled
    g_DictationCompletionChimeScheduled := false  ; Clear IMMEDIATELY to prevent other calls
    Critical "Off"

    ; Only play if flag was set (prevent duplicate execution)
    if (chimeShouldPlay) {
        g_D2C_DictationSubmitMenuCycleFinished := false
        SafePlayDictationSound(g_DictationStopSound)

        ; Execute pending action if one was set (reserved for future use).
        pendingAction := g_PendingDictationAction
        g_PendingDictationAction := ""  ; Clear immediately after reading

        if (pendingAction = "Paste") {
            ; Update indicator text to show status
            UpdateDictationIndicatorText("Pasting...")
            ; Execute paste command
            Send "^v"
            ; Wait for paste to complete before hiding indicator
            Sleep 100  ; Small delay to ensure paste completes
            ; Hide indicator only after paste completes
            HideDictationIndicator()
            g_KeepIndicatorVisible := false
        }

        ; If user stopped dictation with Win+Alt+Shift+0 (no pending action), show Gemini confirm banner (once only).
        Critical "On"
        pendingGemini := g_PendingGeminiPromptAfterDictation
        g_PendingGeminiPromptAfterDictation := false  ; Claim atomically so only one invocation shows the banner
        Critical "Off"
        if (pendingGemini && pendingAction = "") {
            try ScriptSoundPlay(A_ScriptDir . "\sounds\dictation-selection-menu.wav")
            D2C_FlowManager.GetInstance().StartFromDictation()
        }
    }
}

; Called when dictation stop detected: play chime now if clipboard already changed, else wait for change
DictationCompletionChimeOrWaitForClipboard() {
    global g_DictationStartClipboardText
    currentClip := ""
    try {
        currentClip := A_Clipboard
    }
    if (currentClip != g_DictationStartClipboardText) {
        PlayDictationCompletionChime()
    } else {
        OnClipboardChange(DictationClipboardHandler)
        SetTimer(PlayDictationCompletionChime, -1500)
    }
}

CheckDictationRecordingWindow() {
    global g_DictationActive, g_LastStateTransitionTick, g_DictationStartClipboardText
    global g_DictationSoundPlayed, g_DictationCompletionChimeScheduled, g_DictationPulseTimer, g_KeepIndicatorVisible
    ; Check if the "Recording" window exists
    windowExists := false
    try {
        windowExists := WinExist("Recording ahk_exe handy.exe")
    } catch {
        windowExists := false
    }

    ; Handle Start: window exists
    if (windowExists) {
        if (!g_DictationActive) {
            g_DictationActive := true
            g_LastStateTransitionTick := A_TickCount

            ; Capture current clipboard content to detect changes later
            try {
                g_DictationStartClipboardText := A_Clipboard
            } catch {
                g_DictationStartClipboardText := ""
            }

            try {
                RunSetMicVolumeScript()
            } catch Error as e {
                ; Silently handle errors - don't interrupt dictation if script fails
            }

            ShowDictationIndicator()
            StartDictationPulseTimer()
        }

        ; Atomic test-and-set: one sound per session when window first detected
        Critical "On"
        if (!g_DictationSoundPlayed) {
            g_DictationSoundPlayed := true
            Critical "Off"
            SafePlayDictationSound(g_DictationStartSound)
        } else {
            Critical "Off"
        }
    }
    ; Handle Stop: window gone and was active
    else if (!windowExists && g_DictationActive) {
        Critical "On"
        if (!g_DictationActive || g_DictationCompletionChimeScheduled) {
            Critical "Off"
            return
        }

        if (g_LastStateTransitionTick && (A_TickCount - g_LastStateTransitionTick < 500)) {
            Critical "Off"
            return
        }

        g_DictationCompletionChimeScheduled := true
        g_LastStateTransitionTick := A_TickCount
        g_DictationActive := false
        Critical "Off"
        g_DictationSoundPlayed := false

        StopDictationPulseTimer()
        HideDictationIndicator()
        DictationCompletionChimeOrWaitForClipboard()
    } else if (g_DictationActive && windowExists) {
        ShowDictationIndicator()
        if (!g_DictationPulseTimer) {
            StartDictationPulseTimer()
        }
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
    ; This provides instant detection if window already exists
    CheckDictationRecordingWindow()

    ; OPTIMIZED: Ultra-fast polling for instant window detection and audio feedback
    ; Start with 25ms polling (4x faster than normal) for ultra-responsive detection
    ; This ensures zero-delay audio feedback when handy.exe launches
    SetTimer(CheckDictationRecordingWindow, 25)
    ; Revert to normal 500ms polling after 3 seconds (window should be detected by then)
    SetTimer(RevertDictationPolling, -3000)
}

RevertDictationPolling() {
    SetTimer(CheckDictationRecordingWindow, 500)
}

; Force end dictation immediately (e.g., when Ask action is triggered)
; This immediately removes Esc restriction and hides the indicator
EndDictation() {
    global g_DictationActive, g_DictationSoundPlayed

    g_DictationActive := false
    g_DictationSoundPlayed := false

    StopDictationPulseTimer()
    HideDictationIndicator()
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
; ~ prefix: key passes through to handy.exe. First press starts dictation, second stops and copies.
; Uses KeyWait + state machine + recursion guard to prevent duplicate triggers (typematic repeats).
~#!+0::
{
    global g_DictationActive, g_LastStateTransitionTick, g_DictationStartSound
    global g_ProgrammaticDictationStop, g_PendingGeminiPromptAfterDictation, g_D2C_DictationSubmitMenuCycleFinished
    global g_DictationHotkeyIsOwner
    static lastHotkeyTick := 0
    static isProcessing := false

    ; Defensive: if Utils is included after a script-level auto-execute return, globals may be uninitialized.
    ; Default to "not owner" to avoid double-handling dictation across processes.
    if (!IsSet(g_DictationHotkeyIsOwner))
        g_DictationHotkeyIsOwner := false

    if (!g_DictationHotkeyIsOwner) {
        return
    }

    ; Skip when script sends #!+0 programmatically
    if (g_ProgrammaticDictationStop) {
        g_ProgrammaticDictationStop := false
        return
    }

    if (isProcessing)
        return

    currentTick := A_TickCount
    if (currentTick - lastHotkeyTick < 200)
        return
    lastHotkeyTick := currentTick
    isProcessing := true
    ; Capture before KeyWait: check timer may clear g_DictationActive when Recording window closes,
    ; so by the time we reach if/else it can be false even when user intended to stop.
    dictationWasActiveOnKeyPress := g_DictationActive

    keyWaitStart := A_TickCount
    KeyWait("0", "L")

    if (!g_DictationActive) {
        g_DictationActive := true
        g_LastStateTransitionTick := A_TickCount
        ShowDictationIndicator()
        StartDictationPulseTimer()
        ; Sound: monitoring loop plays when window detected (zero latency)

        try {
            RunSetMicVolumeScript()
        } catch {
        }
    }

    ; User was stopping dictation (had been active when they pressed key) -> show Gemini confirm after completion
    if (dictationWasActiveOnKeyPress) {
        g_PendingGeminiPromptAfterDictation := true
        g_D2C_DictationSubmitMenuCycleFinished := false
        g_DictationGeminiConfirmBannerVisible := false  ; Allow 5s banner to show for this cycle (reset from previous N cancel)
    } else {
    }

    ToggleDictationMode()
    isProcessing := false
}

; Win+Alt+Shift+7 is defined in Gemini.ahk (TTS from selection: repeat exactly + read aloud).
