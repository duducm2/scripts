#Requires AutoHotkey v2.0+
#SingleInstance Force
#include %A_ScriptDir%\env.ahk
#include %A_ScriptDir%\StudyLinkHelpers.ahk

global g_StudyLinkSubmenuGui := ""

#include UIA-v2\Lib\UIA.ahk
#include UIA-v2\Lib\UIA_Browser.ahk
#include %A_ScriptDir%\lib\Media.ahk
#include %A_ScriptDir%\SpotifyWASAPI.ahk
#include %A_ScriptDir%\aux\ClipboardFiles.ahk

; UIA ControlType constants (Button=50000). Shared with Gemini.ahk focus helpers.
global UIA_ControlType_Button := 50000

; -----------------------------------------------------------------------------
; MODULE MAP - Utils.ahk stays the runnable entry point / shared library and
; #includes each module below. For a given feature, open just its small module.
; Early BANNER_ACCENT_* and GEMINI_PROMPT_* globals stay here (must load first).
; lib\CopilotWeb.ahk #include stays inline before d2c_flow_manager module.
; See Utils/MODULARIZATION_PROGRESS.md for the full module list.
; -----------------------------------------------------------------------------

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
        Chrome_DetachDebugLog("Chrome_ActivateDetachMenuItem", "keyboard child invoked", "E", "path=keyboardEnterChild"
        )
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

; Paste file(s) via CF_HDROP (file attachment), not path text. Returns true on success.
InsertFiles(paths) {
    global g_lastExpansion
    if (A_TickCount - g_lastExpansion) < 250
        return false
    if (!paths || paths.Length = 0)
        return false
    for path in paths {
        if !Clipboard_PathIsExistingFile(path)
            return false
    }

    g_lastExpansion := A_TickCount
    saved := ClipboardAll()
    ok := false
    try {
        if !Clipboard_SetFiles(paths)
            return false
        if !Clipboard_WaitForFileDrop(800)
            return false
        Sleep 50
        Send "^v"
        ok := true
        ; Brief settle so browser upload handlers receive the paste before clipboard restore.
        Sleep 400
    } finally {
        try A_Clipboard := saved
    }
    return ok
}

InsertFiles_IsAiChatForeground() {
    try {
        title := WinGetTitle("A")
        if (title = "")
            return false
        lower := StrLower(title)
        return InStr(lower, "gemini") || InStr(lower, "copilot")
    } catch {
        return false
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

:o:boschimg::
{
    InsertText(GetPromptText("bosch-brand-image"))
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
            "Create one PowerPoint slide as an image using the attached reference as the main visual guide.`n",
            "Prompts",
            "📊 Create PowerPoint slide (reference)")
    }
    try {
        RegisterHotstring(":o:boschimg", FileRead(promptDir "\bosch-brand-image.txt"), "Prompts",
        "🎨 Bosch brand-compliant image")
    } catch {
        RegisterHotstring(":o:boschimg",
            "Generate one Bosch Brand Guide and BDDS compliant image.`n", "Prompts",
            "🎨 Bosch brand-compliant image")
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
; Minimize Clip Angel after automation (process stays running; no native Alt+P / WinClose).
EnsureClipAngelClosed() {
    hwnd := ClipAngel_MainHwnd()
    if !hwnd
        return
    ClipAngel_HideWindow(hwnd)
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
        ; TODO/follow-up: replace native Alt+V with UIA (All Clips view) — Alt+V breaks with always-open mode.
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
; Show/restore always-running Clip Angel via AHK (no native Alt+P paste hotkey). UIA: dataGridView,
; Row 0 (first clip) per clip-angel.txt. One ElementFromHandle per flow; bounded polls; layout only when not foreground/hidden.
ActivateClipAngelWithFocusCorrection(silent := false, targetMon := 0) {
    needBanner := false
    if !targetMon {
        try targetMon := GetAhkMonitorIndexFromHwnd(WinGetID("A"))
        catch
            targetMon := 0
    }
    hwnd := ClipAngel_MainHwnd()
    if !hwnd {
        if !silent
            ShowCenteredOverlay_Utils("❌ Clip Angel não está em execução.", 2000, BANNER_ACCENT_ERROR)
        return false
    }
    isActive := WinActive("ahk_id " hwnd)
    wasHidden := !ClipAngel_IsWindowShown(hwnd)
    needsLayout := !isActive || wasHidden
    if isActive && !needsLayout {
        ClipAngel_UiaEnsureRow0Selected(hwnd, false)
        return true
    }
    if wasHidden || !isActive {
        needBanner := !silent
        if needBanner
            ClipAngelBanner_Show("📂 Opening Clip Angel...", BANNER_ACCENT_INTERMEDIATE)
    }
    if !ClipAngel_ShowWindow(hwnd) {
        if needBanner
            ClipAngelBanner_Hide()
        if !silent
            ShowCenteredOverlay_Utils("❌ ClipAngel window not found.", 2000, BANNER_ACCENT_ERROR)
        return false
    }
    if needsLayout
        ClipAngel_ApplyLayoutOnMonitor(hwnd, targetMon)
    ClipAngel_EnsureWindowActive(hwnd)
    ClipAngel_UiaEnsureRow0Selected(hwnd, true)
    if needBanner {
        ClipAngelBanner_Show("✅ Done", BANNER_ACCENT_SUCCESS)
        SetTimer(ClipAngelBanner_Hide, -500)
    }
    return true
}

; =============================================================================
; Clip Angel: Mark Last Clip as Favorite
; =============================================================================
; Wait after clipboard change before favoriting newest clip (copy / dictation ingest).
CLIPANGEL_PRE_FAVORITE_INGEST_DELAY_MS := 400
; Settle after row focus, before Alt+Q (all favorite paths).
CLIPANGEL_FAVORITE_UI_SETTLE_MS := 50
; Bounded poll after native Alt+P open before favoriting (cold start can exceed fixed sleeps).
CLIPANGEL_FAVORITE_OPEN_READY_MS := 1200
CLIPANGEL_INCREMENTAL_PASTE_SETTLE_MS := 250
CLIPANGEL_SEQUENTIAL_PASTE_GAP_MS := 300
CLIPANGEL_UIA_POLL_MS := 30
CLIPANGEL_GRID_WAIT_MS := 400
CLIPANGEL_ROW0_WAIT_MS := 300
CLIPANGEL_ROW0_SELECT_WAIT_MS := 250
global g_ClipAngelAutomationBusy := false

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

ClipAngel_IsWindowShown(hwnd) {
    if !hwnd || !WinExist("ahk_id " hwnd)
        return false
    try {
        if WinGetMinMax("ahk_id " hwnd) = 1
            return false
    } catch {
        return false
    }
    return DllCall("IsWindowVisible", "ptr", hwnd)
}

ClipAngel_ShowWindow(hwnd) {
    if !hwnd
        return false
    if WinActive("ahk_id " hwnd) && ClipAngel_IsWindowShown(hwnd)
        return true
    try {
        if WinGetMinMax("ahk_id " hwnd) = 1
            WinRestore("ahk_id " hwnd)
    } catch {
    }
    try WinShow("ahk_id " hwnd)
    catch {
    }
    return ClipAngel_EnsureWindowActive(hwnd)
}

ClipAngel_HideWindow(hwnd) {
    if !hwnd
        return false
    try {
        WinMinimize("ahk_id " hwnd)
        return true
    } catch {
        return false
    }
}

; dataGridView (Type 50036, AutomationId dataGridView) — clip-angel.txt.
; Pass root when caller already has UIA.ElementFromHandle(hwnd) to avoid duplicate COM round-trips.
ClipAngel_UiaGetDataGrid(hwnd, root := 0) {
    if !hwnd
        return 0
    if !root {
        root := UIA.ElementFromHandle(hwnd)
        if !root
            return 0
    }
    return ClipAngel_UiaFindFirst(root, { Type: 50036, AutomationId: "dataGridView" })
}

ClipAngel_WaitForDataGrid(hwnd, timeoutMs := CLIPANGEL_GRID_WAIT_MS, root := 0) {
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        dataGrid := ClipAngel_UiaGetDataGrid(hwnd, root)
        if dataGrid
            return dataGrid
        if !root {
            root := UIA.ElementFromHandle(hwnd)
            if !root
                return 0
        }
        Sleep CLIPANGEL_UIA_POLL_MS
    }
    return 0
}

ClipAngel_WaitForRow0(dataGrid, timeoutMs := CLIPANGEL_ROW0_WAIT_MS) {
    if !dataGrid
        return 0
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        row0 := ClipAngel_UiaFindFirst(dataGrid, { Type: 50025, Name: "Row 0" })
        if row0
            return row0
        Sleep CLIPANGEL_UIA_POLL_MS
    }
    return 0
}

; Invoke List menu when Window tab has focus and dataGridView is missing.
ClipAngel_EnsureListView(hwnd, root := 0) {
    if !hwnd
        return false
    if !root {
        root := UIA.ElementFromHandle(hwnd)
        if !root
            return false
    }
    listItem := ClipAngel_UiaFindFirst(root, { Type: 50011, Name: "List" })
    if !listItem
        return false
    try {
        if listItem.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)
            listItem.InvokePattern.Invoke()
        else
            listItem.SetFocus()
        return true
    } catch {
        return false
    }
}

ClipAngel_UiaResolveRow0(dataGrid) {
    if !dataGrid
        return 0
    row0 := ClipAngel_UiaFindFirst(dataGrid, { Type: 50025, Name: "Row 0" })
    if row0
        return row0
    try {
        rows := dataGrid.FindAll({ Type: 50025 })
        if rows && rows.Length >= 1
            return rows[1]
    } catch {
    }
    return ClipAngel_WaitForRow0(dataGrid)
}

ClipAngel_UiaGridHasSelectionPattern(dataGrid) {
    if !dataGrid
        return false
    try return dataGrid.GetPropertyValue(UIA.Property.IsSelectionPatternAvailable)
    catch
        return false
}

ClipAngel_UiaWaitRow0Selected(row0, dataGrid, timeoutMs := CLIPANGEL_ROW0_SELECT_WAIT_MS) {
    if !row0
        return false
    gridHasSel := ClipAngel_UiaGridHasSelectionPattern(dataGrid)
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        if ClipAngel_UiaRow0IsSelected(row0, dataGrid, gridHasSel)
            return true
        Sleep CLIPANGEL_UIA_POLL_MS
    }
    return false
}

ClipAngel_UiaGridSelectionIncludesRow0(dataGrid) {
    if !dataGrid
        return false
    try {
        if dataGrid.GetPropertyValue(UIA.Property.IsSelectionPatternAvailable) {
            for item in dataGrid.SelectionPattern.GetSelection() {
                try n := item.Name
                catch
                    continue
                if RegExMatch(n, "i)^Row\s*0")
                    return true
            }
        }
    } catch {
    }
    return false
}

ClipAngel_UiaRowLegacySelected(row) {
    if !row
        return false
    try {
        state := row.GetPropertyValue(UIA.Property.LegacyIAccessibleState)
        if (state & 0x2)
            return true
    } catch {
    }
    return false
}

; WinForms DataGridView: cheap property checks first; grid SelectionPattern only when available.
ClipAngel_UiaRow0IsSelected(row0, dataGrid := 0, gridHasSelectionPattern := false) {
    if row0 {
        try {
            if row0.GetPropertyValue(UIA.Property.SelectionItemIsSelected)
                return true
        } catch {
        }
        if ClipAngel_UiaRowLegacySelected(row0)
            return true
    }
    if gridHasSelectionPattern && dataGrid && ClipAngel_UiaGridSelectionIncludesRow0(dataGrid)
        return true
    return false
}

ClipAngel_UiaTryLegacySelectRow(row) {
    if !row
        return false
    try {
        if !row.GetPropertyValue(UIA.Property.IsLegacyIAccessiblePatternAvailable)
            return false
        row.LegacyIAccessiblePattern.Select(3)
        return true
    } catch {
        return false
    }
}

; F10 toggles list vs preview — only send when preview pane has focus, not when grid already focused.
ClipAngel_UiaEnsureGridListFocus(dataGrid, hwnd, root := 0) {
    if !dataGrid
        return false
    try dataGrid.SetFocus()
    catch {
    }
    deadline := A_TickCount + 200
    while (A_TickCount < deadline) {
        try {
            if dataGrid.HasKeyboardFocus
                return true
        } catch {
        }
        Sleep CLIPANGEL_UIA_POLL_MS
    }
    if !hwnd
        return true
    if !root {
        root := UIA.ElementFromHandle(hwnd)
        if !root
            return true
    }
    preview := ClipAngel_UiaFindFirst(root, { AutomationId: "richTextBox" })
    previewFocused := false
    try previewFocused := preview && preview.HasKeyboardFocus
    catch {
    }
    if previewFocused {
        ClipAngel_ReleaseChordModifiersForSend()
        Send "{F10}"
        deadline := A_TickCount + 150
        while (A_TickCount < deadline) {
            try {
                if dataGrid.HasKeyboardFocus
                    return true
            } catch {
            }
            Sleep CLIPANGEL_UIA_POLL_MS
        }
    } else {
        try dataGrid.Click()
        catch {
        }
        try dataGrid.SetFocus()
        catch {
        }
    }
    return true
}

; First list row (Row 0 / rows[1]). force=false skips work when Row 0 is already selected.
ClipAngel_UiaEnsureRow0Selected(hwnd, force := false) {
    if !hwnd
        return false
    listInvoked := false
    try {
        root := UIA.ElementFromHandle(hwnd)
        if !root
            return false
        dataGrid := ClipAngel_UiaGetDataGrid(hwnd, root)
        if !dataGrid {
            listInvoked := ClipAngel_EnsureListView(hwnd, root)
            dataGrid := ClipAngel_WaitForDataGrid(hwnd, CLIPANGEL_GRID_WAIT_MS, root)
            if !dataGrid
                return false
        }
        row0 := ClipAngel_UiaResolveRow0(dataGrid)
        if !row0 && !listInvoked {
            ClipAngel_EnsureListView(hwnd, root)
            dataGrid := ClipAngel_WaitForDataGrid(hwnd, CLIPANGEL_GRID_WAIT_MS, root)
            if dataGrid
                row0 := ClipAngel_UiaResolveRow0(dataGrid)
        }
        if !row0
            return false
        gridHasSel := ClipAngel_UiaGridHasSelectionPattern(dataGrid)
        if !force && ClipAngel_UiaRow0IsSelected(row0, dataGrid, gridHasSel)
            return true
        try {
            if row0.GetPropertyValue(UIA.Property.IsScrollItemPatternAvailable)
                row0.ScrollItemPattern.ScrollIntoView()
        } catch {
        }
        if ClipAngel_UiaTryLegacySelectRow(row0) && ClipAngel_UiaWaitRow0Selected(row0, dataGrid,
            CLIPANGEL_ROW0_SELECT_WAIT_MS)
            return true
        try {
            if row0.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable)
                row0.SelectionItemPattern.Select()
        } catch {
            try row0.SetFocus()
        }
        if ClipAngel_UiaWaitRow0Selected(row0, dataGrid, CLIPANGEL_ROW0_SELECT_WAIT_MS)
            return true
        ClipAngel_UiaEnsureGridListFocus(dataGrid, hwnd, root)
        ClipAngel_ReleaseChordModifiersForSend()
        Send "^{Home}"
        if ClipAngel_UiaWaitRow0Selected(row0, dataGrid, CLIPANGEL_ROW0_SELECT_WAIT_MS)
            return true
        ClipAngel_ReleaseChordModifiersForSend()
        Send "{Home}"
        if ClipAngel_UiaWaitRow0Selected(row0, dataGrid, CLIPANGEL_ROW0_SELECT_WAIT_MS)
            return true
        rn := ""
        try rn := row0.Name
        catch {
            rn := ""
        }
        titleCell := ClipAngel_UiaFindFirst(row0, { Type: 50006, Name: "Title " rn })
        if !titleCell && rn != "Row 0"
            titleCell := ClipAngel_UiaFindFirst(row0, { Type: 50006, Name: "Title Row 0" })
        if titleCell {
            try {
                if titleCell.GetPropertyValue(UIA.Property.IsLegacyIAccessiblePatternAvailable)
                    titleCell.LegacyIAccessiblePattern.Select(3)
                else
                    titleCell.Click()
                if ClipAngel_UiaWaitRow0Selected(row0, dataGrid, CLIPANGEL_ROW0_SELECT_WAIT_MS)
                    return true
            } catch {
            }
        }
        try row0.SetFocus()
        catch {
        }
        return ClipAngel_UiaWaitRow0Selected(row0, dataGrid, CLIPANGEL_ROW0_SELECT_WAIT_MS)
    } catch {
        return false
    }
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

; Legacy native Alt+V send — All Clips view in MergeNonFavoriteClips only.
; Open/close and paste flows use ClipAngel_ShowWindow/HideWindow or Alt+P (+ Enter) instead.
ClipAngel_SendToggleHotkey() {
    ClipAngel_WaitChordModifiersReleased()
    ClipAngel_ReleaseChordModifiersForSend()
    Send "!v"
}

ClipAngel_EnsureWindowActive(hwnd, timeoutMs := 800) {
    if !hwnd
        return false
    if WinActive("ahk_id " hwnd)
        return true
    endTick := A_TickCount + timeoutMs
    loop 3 {
        try WinActivate("ahk_id " hwnd)
        catch
            return false
        remaining := endTick - A_TickCount
        if (remaining <= 0)
            break
        if WinWaitActive("ahk_id " hwnd, , Max(0.05, remaining / 1000.0))
            return true
        Sleep 50
    }
    return WinActive("ahk_id " hwnd)
}

; Move to monitor work area and maximize; re-activate if layout steals focus.
ClipAngel_ApplyLayoutOnMonitor(hwnd, targetMon := 0) {
    if !hwnd
        return false
    if (!targetMon || targetMon < 1 || targetMon > MonitorGetCount()) {
        try targetMon := GetAhkMonitorIndexFromHwnd(WinGetID("A"))
        catch
            targetMon := 0
    }
    if (!targetMon) {
        try targetMon := MonitorGetPrimary()
        catch
            targetMon := 1
    }
    MoveWindowToMonitor(hwnd, targetMon)
    if !WinActive("ahk_id " hwnd)
        ClipAngel_EnsureWindowActive(hwnd)
    TryMaximizeWindow(hwnd)
    return ClipAngel_EnsureWindowActive(hwnd)
}

ClipAngel_IsListReady(&outHwnd := 0) {
    outHwnd := ClipAngel_MainHwnd()
    if !outHwnd
        return false
    dataGrid := ClipAngel_UiaGetDataGrid(outHwnd)
    if !dataGrid
        return false
    return ClipAngel_UiaFindFirst(dataGrid, { Type: 50025, Name: "Row 0" }) ? true : false
}

; Poll until dataGridView + Row 0 exist (e.g. after native Alt+P open). Retries row-0 selection at end.
ClipAngel_WaitForListReady(timeoutMs := CLIPANGEL_FAVORITE_OPEN_READY_MS) {
    deadline := A_TickCount + timeoutMs
    hwnd := 0
    while (A_TickCount < deadline) {
        if ClipAngel_IsListReady(&hwnd) {
            ClipAngel_UiaEnsureRow0Selected(hwnd, false)
            return true
        }
        if hwnd := ClipAngel_MainHwnd()
            ClipAngel_ShowWindow(hwnd)
        Sleep CLIPANGEL_UIA_POLL_MS
    }
    if hwnd := ClipAngel_MainHwnd() {
        ClipAngel_ShowWindow(hwnd)
        ClipAngel_UiaEnsureRow0Selected(hwnd, true)
        return ClipAngel_IsListReady()
    }
    return false
}

ClipAngel_ResolvePriorHwnd(priorHwnd := 0) {
    if (priorHwnd && WinExist("ahk_id " priorHwnd))
        return priorHwnd
    try {
        activeHwnd := WinGetID("A")
        if (activeHwnd && !WinActive("ahk_exe ClipAngel.exe"))
            return activeHwnd
    } catch {
    }
    return 0
}

ClipAngel_RestorePriorFocus(priorHwnd) {
    if (!priorHwnd || !WinExist("ahk_id " priorHwnd))
        return
    if WinActive("ahk_exe ClipAngel.exe")
        return
    try {
        WinActivate("ahk_id " priorHwnd)
        WinWaitActive("ahk_id " priorHwnd, , 2)
    } catch {
    }
}

ClipAngel_TryAcquireAutomationLock() {
    global g_ClipAngelAutomationBusy
    if (g_ClipAngelAutomationBusy) {
        ShowCenteredOverlay_Utils("⏳ Clip Angel busy.", 1200, BANNER_ACCENT_INTERMEDIATE)
        return false
    }
    g_ClipAngelAutomationBusy := true
    return true
}

ClipAngel_ReleaseAutomationLock() {
    global g_ClipAngelAutomationBusy
    g_ClipAngelAutomationBusy := false
}

ClipAngel_EnsureOpenAndReady(silent := true) {
    if !ActivateClipAngelWithFocusCorrection(silent)
        return false
    return ClipAngel_IsListReady()
}

ClipAngel_SendIncrementalPaste() {
    ClipAngel_WaitChordModifiersReleased()
    ClipAngel_ReleaseChordModifiersForSend()
    SendInput "^!b"
    Sleep(CLIPANGEL_INCREMENTAL_PASTE_SETTLE_MS)
}

ClipAngel_CloseAndRestoreFocus(priorHwnd := 0) {
    EnsureClipAngelClosed()
    ClipAngel_RestorePriorFocus(priorHwnd)
}

; Native open + row 0: Alt+P + ShowWindow + ^Home (shared by paste and favorite flows).
; Trade-off: brief Clip Angel visibility vs full UIA open/row-select (efficiency-canon §11).
ClipAngel_ActivateNativeFirstClip(priorHwnd := 0) {
    ClipAngel_WaitChordModifiersReleased()
    ClipAngel_ReleaseChordModifiersForSend()
    if (priorHwnd)
        ClipAngel_EnsureWindowActive(priorHwnd)
    SendInput "{Alt up}{Shift up}{Win up}{Ctrl up}"
    Sleep 200
    SendInput "!p"
    Sleep 200
    if hwnd := ClipAngel_MainHwnd()
        ClipAngel_ShowWindow(hwnd)
    SendInput "^{Home}"
    Sleep 200
    SendInput "{Alt up}{Shift up}{Win up}{Ctrl up}"
}

; Native top-item paste: open via ActivateNativeFirstClip, then incremental paste (^!b).
ClipAngel_SendNativeTopItemKeys(priorHwnd := 0) {
    ClipAngel_ActivateNativeFirstClip(priorHwnd)
    SendInput "^!b"
}

; Send top list item via Clip Angel native keys. Closes after paste and restores prior focus.
ClipAngel_SendTopListItem(priorHwnd := 0) {
    if !ClipAngel_TryAcquireAutomationLock()
        return false
    priorHwnd := ClipAngel_ResolvePriorHwnd(priorHwnd)
    ok := false
    try {
        ClipAngel_SendNativeTopItemKeys(priorHwnd)
        ok := true
    } catch Error as e {
        ShowCenteredOverlay_Utils("❌ Clip Angel paste failed: " . e.Message, 2500, BANNER_ACCENT_ERROR)
        ok := false
    } finally {
        EnsureClipAngelClosed()
        ClipAngel_RestorePriorFocus(priorHwnd)
        ClipAngel_ReleaseAutomationLock()
    }
    return ok
}

; Paste N top-list items in order. Opens once, incremental paste between items, closes at end.
ClipAngel_SendTopListItemSequential(count, priorHwnd := 0) {
    if (!IsInteger(count))
        return false
    n := Integer(count)
    if (n < 1)
        return false
    if !ClipAngel_TryAcquireAutomationLock()
        return false
    priorHwnd := ClipAngel_ResolvePriorHwnd(priorHwnd)
    ok := false
    try {
        ClipAngel_WaitChordModifiersReleased()
        ClipAngel_ReleaseChordModifiersForSend()
        loop n {
            if (A_Index = 1) {
                ClipAngel_SendNativeTopItemKeys(priorHwnd)
            } else {
                Sleep(CLIPANGEL_SEQUENTIAL_PASTE_GAP_MS)
                ClipAngel_ReleaseChordModifiersForSend()
                SendInput "^!b"
            }
            Sleep(CLIPANGEL_INCREMENTAL_PASTE_SETTLE_MS)
        }
        ClipAngel_CloseAndRestoreFocus(priorHwnd)
        ok := true
    } catch Error as e {
        ShowCenteredOverlay_Utils("❌ Clip Angel sequential paste failed: " . e.Message, 2500, BANNER_ACCENT_ERROR)
        ok := false
    } finally {
        EnsureClipAngelClosed()
        ClipAngel_RestorePriorFocus(priorHwnd)
        ClipAngel_ReleaseAutomationLock()
    }
    return ok
}

; target: "first" = top grid row (Row 0 / newest), "last" = last row returned by UIA FindAll
; (virtualized lists may only expose visible rows - use "first" for reliable top-clip behavior).
MarkLastClipAsFavorite(target := "first", waitForIngest := false) {
    if waitForIngest
        Sleep(CLIPANGEL_PRE_FAVORITE_INGEST_DELAY_MS)
    if !ClipAngel_TryAcquireAutomationLock()
        return
    priorHwnd := ClipAngel_ResolvePriorHwnd(0)
    try {
        if (target = "last") {
            MarkLastClipAsFavorite_UiaLastRow()
            return
        }
        ClipAngel_ActivateNativeFirstClip()
        if !ClipAngel_WaitForListReady(CLIPANGEL_FAVORITE_OPEN_READY_MS) {
            ShowCenteredOverlay_Utils("❌ Clip Angel did not open.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        Sleep(CLIPANGEL_FAVORITE_UI_SETTLE_MS)
        ClipAngel_WaitChordModifiersReleased()
        ClipAngel_ReleaseChordModifiersForSend()
        SendInput "!q"
        ScriptSoundPlay(A_ScriptDir "\sounds\favorite-set.wav")
        ShowCenteredOverlay_Utils("✅ Sent Alt+Q - marked focused clip as favorite.", 1500, BANNER_ACCENT_SUCCESS)
    } catch Error as e {
        ShowCenteredOverlay_Utils("❌ Mark favorite failed: " . e.Message, 2500, BANNER_ACCENT_ERROR)
    } finally {
        EnsureClipAngelClosed()
        ClipAngel_RestorePriorFocus(priorHwnd)
        ClipAngel_ReleaseAutomationLock()
    }
}

; Legacy UIA path for target="last" only (no callers today; API preserved).
MarkLastClipAsFavorite_UiaLastRow() {
    ActivateClipAngelWithFocusCorrection()
    hwnd := ClipAngel_MainHwnd()
    if !hwnd {
        ShowCenteredOverlay_Utils("❌ Clip Angel did not open.", 2000, BANNER_ACCENT_ERROR)
        return
    }
    try WinActivate("ahk_id " hwnd)
    catch {
        ShowCenteredOverlay_Utils("❌ Clip Angel window not found.", 2000, BANNER_ACCENT_ERROR)
        return
    }
    if !WinWaitActive("ahk_id " hwnd, , 2) {
        ShowCenteredOverlay_Utils("❌ Clip Angel did not become active.", 2000, BANNER_ACCENT_ERROR)
        return
    }
    el := UIA.ElementFromHandle(hwnd)
    if !el {
        ShowCenteredOverlay_Utils("❌ Clip Angel UI not available.", 2000, BANNER_ACCENT_ERROR)
        return
    }
    dataGrid := ClipAngel_UiaFindFirst(el, { Type: 50036, AutomationId: "dataGridView" })
    if !dataGrid {
        ShowCenteredOverlay_Utils("❌ Clip list not found (Window tab may still have focus).", 2500,
            BANNER_ACCENT_ERROR)
        return
    }
    rows := 0
    try rows := dataGrid.FindAll({ Type: 50025 })
    catch {
        rows := 0
    }
    if !rows || rows.Length < 1 {
        ShowCenteredOverlay_Utils("❌ No clips in list.", 2000, BANNER_ACCENT_ERROR)
        return
    }
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
    if favCell && ClipAngel_FavoriteCellIsOn(favCell) {
        ShowCenteredOverlay_Utils("✅ Selected clip is already a favorite.", 1500, BANNER_ACCENT_SUCCESS)
        return
    }
    if !WinActive("ahk_id " hwnd) {
        try WinActivate("ahk_id " hwnd)
        if !WinWaitActive("ahk_id " hwnd, , 2) {
            ShowCenteredOverlay_Utils("❌ Clip Angel lost focus before Alt+Q.", 2000, BANNER_ACCENT_ERROR)
            return
        }
    }
    Sleep(CLIPANGEL_FAVORITE_UI_SETTLE_MS)
    ClipAngel_WaitChordModifiersReleased()
    ClipAngel_ReleaseChordModifiersForSend()
    SendInput "!q"
    ScriptSoundPlay(A_ScriptDir "\sounds\favorite-set.wav")
    ShowCenteredOverlay_Utils("✅ Sent Alt+Q - marked focused clip as favorite.", 1500, BANNER_ACCENT_SUCCESS)
}

; [Utils module] Handy AI model configuration map and persistence -> Utils\handy_ai_model_config.ahk
#include %A_ScriptDir%\Utils\handy_ai_model_config.ahk
; [Utils module] ShowAiModelSelector GUI -> Utils\handy_ai_model_gui.ahk
#include %A_ScriptDir%\Utils\handy_ai_model_gui.ahk

; [Utils module] Gemini-to-Cursor transfer numeric window selector -> Utils\gemini_cursor_transfer.ahk
#include %A_ScriptDir%\Utils\gemini_cursor_transfer.ahk

; [Utils module] Cursor AI text field focus and VS Code chat helpers -> Utils\cursor_composer_focus.ahk
#include %A_ScriptDir%\Utils\cursor_composer_focus.ahk

; [Utils module] Language flag indicator and status banners -> Utils\language_flag_indicator.ahk
#include %A_ScriptDir%\Utils\language_flag_indicator.ahk

; [Utils module] Handy UIA helper functions and ShowAiModelSelector support -> Utils\handy_uia_helpers.ahk
#include %A_ScriptDir%\Utils\handy_uia_helpers.ahk

; [Utils module] SelectAiModelInHandy entry and pre-movement warning -> Utils\handy_selector_entry.ahk
#include %A_ScriptDir%\Utils\handy_selector_entry.ahk

; [Utils module] Standard loading bar show/update/hide lifecycle -> Utils\standard_loading_bar.ahk
#include %A_ScriptDir%\Utils\standard_loading_bar.ahk

; [Utils module] Hotstring Gemini banner and D2C preset helpers -> Utils\hotstring_gemini_banner.ahk
#include %A_ScriptDir%\Utils\hotstring_gemini_banner.ahk

#include %A_ScriptDir%\lib\CopilotWeb.ahk

; [Utils module] D2C_FlowManager dictation-Gemini-Cursor state machine -> Utils\d2c_flow_manager.ahk
#include %A_ScriptDir%\Utils\d2c_flow_manager.ahk

; [Utils module] Deprecated dictation Gemini confirm banner -> Utils\dictation_legacy.ahk
#include %A_ScriptDir%\Utils\dictation_legacy.ahk
; [Utils module] Global sound toggle and script audio helpers -> Utils\global_sound_audio.ahk
#include %A_ScriptDir%\Utils\global_sound_audio.ahk

; [Utils module] ToggleOutlookAndTeams macro -> Utils\toggle_outlook_teams.ahk
#include %A_ScriptDir%\Utils\toggle_outlook_teams.ahk

; [Utils module] CheckAndOpenOutlookTeams prompt helper -> Utils\outlook_teams_check.ahk
#include %A_ScriptDir%\Utils\outlook_teams_check.ahk

; [Utils module] Dictation clipboard cleanup countdown -> Utils\dictation_cleanup.ahk
#include %A_ScriptDir%\Utils\dictation_cleanup.ahk

; [Utils module] Dictation merge non-favorite clips countdown -> Utils\dictation_merge.ahk
#include %A_ScriptDir%\Utils\dictation_merge.ahk

; [Utils module] Clean clipboard countdown macro -> Utils\cleanclipboard.ahk
#include %A_ScriptDir%\Utils\cleanclipboard.ahk

; [Utils module] Project data for Cursor window focus selector -> Utils\project_data_cursor.ahk
#include %A_ScriptDir%\Utils\project_data_cursor.ahk

; [Utils module] Global AI generation state U macro -> Utils\ai_generation_state.ahk
#include %A_ScriptDir%\Utils\ai_generation_state.ahk

; [Utils module] Jump mouse to window center #!+Q -> Utils\jump_mouse_middle.ahk
#include %A_ScriptDir%\Utils\jump_mouse_middle.ahk

; [Utils module] Handy AI model selector hotkey #!+C -> Utils\handy_selector_hotkey.ahk
#include %A_ScriptDir%\Utils\handy_selector_hotkey.ahk

; [Utils module] Desktop to Recycle Bin macro -> Utils\desktop_recycle.ahk
#include %A_ScriptDir%\Utils\desktop_recycle.ahk

; [Utils module] Mouse jump helpers and prediction squares -> Utils\mouse_jump_arrows.ahk
#include %A_ScriptDir%\Utils\mouse_jump_arrows.ahk

; [Utils module] Square selector mouse jump (part 1) -> Utils\square_selector_mouse_jump_01.ahk
#include %A_ScriptDir%\Utils\square_selector_mouse_jump_01.ahk
; [Utils module] Square selector mouse jump (part 2) -> Utils\square_selector_mouse_jump_02.ahk
#include %A_ScriptDir%\Utils\square_selector_mouse_jump_02.ahk

; [Utils module] Win+Alt+Shift+Arrow five-step hotkeys -> Utils\mouse_jump_hotkeys.ahk
#include %A_ScriptDir%\Utils\mouse_jump_hotkeys.ahk

; [Utils module] Peek PDF / QuickLook study helpers (part 1) -> Utils\peek_pdf_study_01.ahk
#include %A_ScriptDir%\Utils\peek_pdf_study_01.ahk
; [Utils module] Peek PDF / QuickLook study helpers (part 2) -> Utils\peek_pdf_study_02.ahk
#include %A_ScriptDir%\Utils\peek_pdf_study_02.ahk

; [Utils module] Peek PDF / QuickLook study helpers (part 3) -> Utils\peek_pdf_study_03.ahk
#include %A_ScriptDir%\Utils\peek_pdf_study_03.ahk

; [Utils module] Peek PDF study hotkey #!+x -> Utils\study_hotkey_x.ahk
#include %A_ScriptDir%\Utils\study_hotkey_x.ahk

; [Utils module] Hotstring selector system core and BuildHotstringCharMap -> Utils\hotstring_selector_core.ahk
#include %A_ScriptDir%\Utils\hotstring_selector_core.ahk

; [Utils module] Context file browser (Win+Alt+Shift+N) -> Utils\context_file_browser.ahk
#include %A_ScriptDir%\Utils\context_file_browser.ahk

; One-shot: close Utility Shortcuts if still open (no expansion/macro chosen in time)
; [Utils module] CleanupHotstringSelector and auto-close idle -> Utils\hotstring_selector_cleanup.ahk
#include %A_ScriptDir%\Utils\hotstring_selector_cleanup.ahk

; [Utils module] HandleHotstringChar and Gemini paste helpers -> Utils\hotstring_selector_handlers_01.ahk
#include %A_ScriptDir%\Utils\hotstring_selector_handlers_01.ahk

; [Utils module] Hotstring selector utility category handlers -> Utils\hotstring_selector_handlers_02.ahk
#include %A_ScriptDir%\Utils\hotstring_selector_handlers_02.ahk

; [Utils module] ShowHotstringSelector GUI and category view -> Utils\hotstring_selector_gui.ahk
#include %A_ScriptDir%\Utils\hotstring_selector_gui.ahk

; [Utils module] Utility shortcuts #!+U and ^!# secondary triggers -> Utils\utility_shortcuts.ahk
#include %A_ScriptDir%\Utils\utility_shortcuts.ahk

; [Utils module] Focus mode multi-monitor blackout (#!+Y) -> Utils\focus_mode.ahk
#include %A_ScriptDir%\Utils\focus_mode.ahk

; [Utils module] Print Screen chime, global Escape hotkey -> Utils\print_screen_escape.ahk
#include %A_ScriptDir%\Utils\print_screen_escape.ahk

; [Utils module] Dictation indicator, ~#!+0 hotkey, ToggleDictationMode -> Utils\dictation_toggle.ahk
#include %A_ScriptDir%\Utils\dictation_toggle.ahk

; Win+Alt+Shift+7 is defined in Gemini.ahk (TTS from selection: repeat exactly + read aloud).
