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

            ActivateClipAngelWithFocusCorrection(true, originMon)
            clipHwnd := WinExist("ClipAngel")
            if (!clipHwnd) {
                StandardLoadingBar_Update("❌ Clip Angel: window not found", BANNER_ACCENT_ERROR)
                return
            }

            ; Activation/layout handled inside ActivateClipAngelWithFocusCorrection(); bounded wait before F4.
            if (!WinWaitActive("ahk_id " clipHwnd, , 0.6)) {
                StandardLoadingBar_Update("❌ Clip Angel: failed to activate", BANNER_ACCENT_ERROR)
                return
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
    try EnsureClipAngelClosed()
    catch {
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

    if hwnd := ClipAngel_MainHwnd()
        ClipAngel_ShowWindow(hwnd)
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

    EnsureClipAngelClosed()
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
        ; Countdown finished -> hide banner and clear clipboard using the existing workflow (show Clip Angel, Ctrl+Alt+K, etc.)
        DictationCleanup_StopCountdown(false)
        CleanClipboardInternal()
        return
    }

    DictationCleanup_UpdateBannerText()
}

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
