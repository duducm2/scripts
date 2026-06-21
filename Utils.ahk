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
; [Utils module] Chrome detach tab (part 3) and detach entry -> Utils\chrome_detach_03.ahk
#include %A_ScriptDir%\Utils\chrome_detach_03.ahk

; [Utils module] Gemini mode picker mouse + UIA -> Utils\gemini_mode_picker.ahk
#include %A_ScriptDir%\Utils\gemini_mode_picker.ahk

; -----------------------------------------------------------------------------
; This script consolidates various utility hotkeys.
; -----------------------------------------------------------------------------

; [Utils module] Hotstrings core (InitHotstringsCheatSheet) -> Utils\hotstrings_core.ahk
#include %A_ScriptDir%\Utils\hotstrings_core.ahk
; [Utils module] Files and links quick-open (InitQuickOpenFiles) -> Utils\files_links.ahk
#include %A_ScriptDir%\Utils\files_links.ahk
; [Utils module] Macros system RegisterMacro and assignments -> Utils\macros_system.ahk
#include %A_ScriptDir%\Utils\macros_system.ahk
; [Utils module] Clip Angel merge non-favorite clips -> Utils\clip_angel_merge.ahk
#include %A_ScriptDir%\Utils\clip_angel_merge.ahk
; [Utils module] Clip Angel activate with focus correction -> Utils\clip_angel_activate.ahk
#include %A_ScriptDir%\Utils\clip_angel_activate.ahk
; [Utils module] Clip Angel mark favorite and related flows -> Utils\clip_angel_favorite.ahk
#include %A_ScriptDir%\Utils\clip_angel_favorite.ahk

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
