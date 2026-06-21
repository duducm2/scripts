/* ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** **
    * Win + Alt + Shift symbol layer shortcuts (AHK v2)
    * â€¢ Provides system - wide symbol shortcuts
    ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** /
    /********************************************************************
     *   AVAILABLE WIN+ALT+SHIFT COMBINATIONS
     *   The following combinations are not currently in use:
     *
     *   Letters still free: P, U
     *   Win+Alt+Shift+L: Outlook Copilot shortcut modal (global; actions target New Outlook)
     *
     *   Ctrl+Alt+Win+V: RESERVED — maximize active window (handled in WindowManagement.ahk;
     *   used by ZMK hold on minimize/close key). Do not bind another global ^!#v action here.
     *   Ctrl+Alt+Win+N: TEMPORARY — M365 Copilot auto-continue loop (toggle; remove block at file end)
     *   Numbers: 9 is free; 0-8 are used
     *   Symbols: ; ' [ ] \ | ` ~ @ # $ % ^ & * ( ) - _ = + { } : " < > ? /
     *
     *   Note: Some combinations use Ctrl+Alt+Shift+Arrow keys for extended mouse movement
********************************************************************/
#Requires AutoHotkey v2.0+

#SingleInstance Force

SetTitleMatchMode 2

; -----------------------------------------------------------------------------
; MODULE MAP - this file stays the runnable entry point / source of truth and
; #includes each module below. For a given feature, open just its small module
; (handy for low-context AI). Anything not listed still lives inline in this file.
;   Shift keys\helpers.ahk                  - early globals, SafeDebugLog, GetChatGPTWindowHwnd
;   Shift keys\config.ahk                     - config, cheat-sheet string utils, ShiftKeysIPC_Bootstrap
;   Shift keys\cheat_sheet_data.ahk           - cheatSheets map population
;   Shift keys\app_hotkeys.ahk                - per-app Win+Alt+Shift hotkey definitions
;   Shift keys\cheat_sheet_gui.ahk            - cheat sheet GUI, search, hold detection
;   Shift keys\fast_copy_clipangel.ahk        - Clip Angel fast copy mode + #!+1/#!+J
;   Shift keys\hotif_*.ahk                    - #HotIf app handlers (OneNote, Outlook, Teams, ...)
;   Shift keys\teams_chat_uia_*.ahk           - Teams chat UIA automation
;   Shift keys\outlook_appointment_palette.ahk - Outlook appointment configuration palette
;   Shift keys\outlook_appointment_uia_*.ahk  - Outlook appointment UIA state checking
;   Shift keys\mobills_*.ahk                  - Mobills pagination, banner, URL hotkeys
;   Shift keys\m365_copilot_temp.ahk          - TEMPORARY M365 Copilot auto-continue (^!#n)
; -----------------------------------------------------------------------------

#include %A_ScriptDir%\env.ahk
#include UIA-v2\Lib\UIA.ahk
#include UIA-v2\Lib\UIA_Browser.ahk
#include %A_ScriptDir%\Utils.ahk
; Focus dwell watcher + #!+Y (Utils) must share this process so ToggleFocusMode sees the same globals as EnableFocusMode.
FocusBlackoutWatcher_Start()
; Volume: AppLaunchers also schedules retries; this catches Shift keys process when sessions register slightly later.
SetTimer(() => ApplyScriptMasterVolumeTarget(), -3500)
#include %A_ScriptDir%\aux\ShiftKeysIPC.ahk
#include %A_ScriptDir%\CheatSheetRich.ahk

; --- Global Variables ---
global DEBUG_LOG_PATH := A_ScriptDir "\.cursor\debug.log"
; Phase 5: Gate debug I/O; set to true only when diagnosing (avoids file I/O in hot paths).
global DEBUG_SHIFTKEYS := false
global g_BlackoutSuppressedUntil

; --- Blackout Banner Suppression Integration (implementation in Utils.ahk) ---
IsBlackoutSuppressed() {
    return Blackout_IsSuppressed()
}

DisableBlackout7Min(*) {
    Blackout_Disable7Min()
}

; Debug mode agent logging (runtime evidence for this session only)
; (disabled) agent log debug-31b036

; Helper function for safe debug logging with retry on file lock
; Handles file locking gracefully by retrying with exponential backoff
; No-op when DEBUG_SHIFTKEYS is false (production).
SafeDebugLog(text) {
    if (!DEBUG_SHIFTKEYS)
        return true
    maxRetries := 3
    retryDelay := 10
    loop maxRetries {
        try {
            FileAppend text, DEBUG_LOG_PATH
            return true
        } catch Error as err {
            ; Check if error has Number property before accessing it
            ; File lock error is typically error code 32
            hasNumber := false
            errNumber := 0
            try {
                errNumber := err.Number
                hasNumber := true
            } catch {
                ; Error doesn't have Number property, treat as non-retryable
                hasNumber := false
            }

            ; If it's a file lock error (32) and we have retries left, wait and retry
            if (hasNumber && errNumber = 32 && A_Index < maxRetries) {
                Sleep retryDelay * A_Index  ; Exponential backoff
            } else {
                ; For other errors or final retry, silently fail to not interrupt script execution
                return false
            }
        }
    }
    return false
}

; Helper: write NDJSON log line for debug agent (no-op on failure)
AgentDebugLog(hypothesisId, message, runId := "initial") {
    ; Intentionally no-op.
    return
}

; Helper: find ChatGPT chrome window by case-insensitive contains match
GetChatGPTWindowHwnd() {
    for hwnd in WinGetList("ahk_exe chrome.exe") {
        if InStr(WinGetTitle("ahk_id " hwnd), "chatgpt", false)
            return hwnd
    }
    return 0
}

; --- Config ---------------------------------------------------------------
PROMPT_FILE := A_ScriptDir "\\ChatGPT_Prompt.txt"
global USE_CHROME_PDF_PRESENT_FALLBACK := true

; ShiftKeys daemon IPC: bootstrap connection on load (non-blocking)
ShiftKeysIPC_Bootstrap()

; Pad first bracket group to a stable display width (inner text centered with spaces).
CheatSheet_StrRepeat(s, count) {
    r := ""
    loop count
        r .= s
    return r
}

PadShortcut(shortcut, targetWidth := 18) {
    if !RegExMatch(shortcut, "^\[(.+)\]$", &m)
        return shortcut
    inner := m[1]
    bracketLen := 2 + StrLen(inner)
    if (bracketLen >= targetWidth)
        return shortcut
    pad := targetWidth - bracketLen
    left := Floor(pad / 2)
    right := pad - left
    return "[" . CheatSheet_StrRepeat(" ", left) . inner . CheatSheet_StrRepeat(" ", right) . "]"
}

JoinLines(lines, sep := "`n") {
    out := ""
    for i, line in lines {
        out .= (i = 1 ? "" : sep) . line
    }
    return out
}

; Strip >>>/---; unwrap each [...] to inner text so search matches visible words (same as RichEdit plain text).
CheatSheet_LineSearchHaystack(line) {
    s := RegExReplace(line, "^>>>\s*|^---\s*", "")
    loop 40 {
        s2 := RegExReplace(s, "\[([^\]]*)\]", "$1")
        if (s2 = s)
            break
        s := s2
    }
    return StrLower(Trim(s))
}

; Multi-word AND on description haystack only (see CheatSheet_LineSearchHaystack).
CheatSheet_LineMatchesQuery(line, query) {
    q := Trim(query)
    if (q = "")
        return true
    hay := CheatSheet_LineSearchHaystack(line)
    terms := StrSplit(StrLower(q), " ", " `t")
    for term in terms {
        if (term = "")
            continue
        if !InStr(hay, term)
            return false
    }
    return true
}

; Filtered overlay body: keep AND/haystack matching on raw lines; prefix each hit with the last === section === label.
CheatSheet_BuildFilteredBodyWithSections(body, query) {
    q := Trim(query)
    if (q = "")
        return body
    lines := StrSplit(body, "`n")
    currentSection := ""
    filtered := []
    for line in lines {
        if RegExMatch(line, "^\s*===\s*(.+?)\s*===\s*$", &m) {
            currentSection := Trim(m[1])
            if CheatSheet_LineMatchesQuery(line, q)
                filtered.Push(line)
            continue
        }
        if CheatSheet_LineMatchesQuery(line, q) {
            if (currentSection != "")
                filtered.Push("[" currentSection "] " line)
            else
                filtered.Push(line)
        }
    }
    return filtered.Length ? JoinLines(filtered) : "(no matches)"
}

CheatSheet_FocusSearchEdit(ctrl) {
    if (!IsObject(ctrl))
        return
    try {
        gw := ctrl.Gui.Hwnd
        fcn := ControlGetFocus("ahk_id " gw)
        fh := ControlGetHwnd(fcn, "ahk_id " gw)
        if (fh = ctrl.Hwnd)
            return
    } catch {
    }
    ctrl.Focus()
    len := StrLen(ctrl.Value)
    ; EM_SETSEL: caret at end without selection (programmatic Focus() often selects all; next key would replace).
    SendMessage(0xB1, len, len, ctrl)
}

; Updating the multiline body can steal focus from the filter Edit; refocus after resize.
; Do not use SetTimer(() => ..., -50): each keystroke created a new timer object and stacked Focus() calls
; (select-all + next key replaced text). Skip redundant Focus via HWND match; EM_SETSEL after real refocus.
CheatSheet_DeferFocusSearch(ctrl) {
    if (IsObject(ctrl))
        CheatSheet_FocusSearchEdit(ctrl)
}

CheatSheet_OnEscapeApp(*) {
    global g_helpGui, g_helpShown
    if (IsObject(g_helpGui) && g_helpShown) {
        g_helpGui.Hide()
        g_helpShown := false
    }
}

CheatSheet_OnEscapeGlobal(*) {
    global g_globalGui, g_globalShown
    if (IsObject(g_globalGui) && g_globalShown) {
        g_globalGui.Hide()
        g_globalShown := false
    }
}

CheatSheet_OnEscapeSearchAll(*) {
    global g_searchAllGui
    if IsObject(g_searchAllGui)
        g_searchAllGui.Hide()
}

; Map "=== section label ===" to a modifier prefix for implied chords (first [key] only).
CheatSheet_SectionLabelToModifierPrefix(label) {
    s := Trim(StrLower(label))
    if (s = "")
        return ""
    if InStr(s, "function") && InStr(s, "misc")
        return ""
    if InStr(s, "ctrl+alt+win")
        return ""
    if InStr(s, "ctrl+shift")
        return "Ctrl+Shift+"
    if InStr(s, "ctrl+alt")
        return "Ctrl+Alt+"
    if InStr(s, "alt+shift")
        return "Alt+Shift+"
    if RegExMatch(s, "^alt \(other")
        return "Alt+"
    if RegExMatch(s, "^alt \(ahk")
        return "Alt+"
    if RegExMatch(s, "^ctrl \(no other")
        return "Ctrl+"
    if (s = "shift")
        return "Shift+"
    if (s = "alt")
        return "Alt+"
    if (s = "ctrl")
        return "Ctrl+"
    if (s = "win")
        return "Win+"
    return ""
}

; First line like "Explorer (Shift)" — single parenthetical modifier.
CheatSheet_ContextParensToModifierPrefix(inner) {
    w := Trim(StrLower(inner))
    wCompact := RegExReplace(w, "\s+", "")
    if (w = "shift")
        return "Shift+"
    if (wCompact = "ctrl+alt")
        return "Ctrl+Alt+"
    if (w = "ctrl")
        return "Ctrl+"
    if (w = "alt")
        return "Alt+"
    if (w = "win")
        return "Win+"
    return ""
}

; True if the bracket inner already lists a full chord or should not be prefixed.
CheatSheet_BracketInnerHasExplicitChord(inner) {
    if (inner = "")
        return true
    if InStr(inner, "+")
        return true
    if InStr(inner, "/") || InStr(inner, "...")
        return true
    if InStr(inner, "(") ; e.g. Esc (bulk)
        return true
    return false
}

; Combine section/cluster prefix with first [inner], e.g. Shift+ + M -> [Shift+M].
CheatSheet_MergeClusterIntoBracket(bracket, clusterPrefix) {
    if (clusterPrefix = "")
        return bracket
    if !RegExMatch(bracket, "^\[(.+)\]$", &m)
        return bracket
    inner := m[1]
    if CheatSheet_BracketInnerHasExplicitChord(inner)
        return bracket
    return "[" . clusterPrefix . inner . "]"
}

; Function to process cheat sheet text and pad all shortcuts
ProcessCheatSheetText(text) {
    lines := StrSplit(text, "`n")
    processedLines := []
    clusterPrefix := ""

    for line in lines {
        if RegExMatch(line, "^\s*===\s*(.+?)\s*===\s*$", &hm) {
            clusterPrefix := CheatSheet_SectionLabelToModifierPrefix(Trim(hm[1]))
            processedLines.Push(line)
            continue
        }
        ; Subsections like "Outlook (Shift)" then "Outlook (Ctrl+Alt)" must each update clusterPrefix (do not stop after first).
        if (!InStr(line, "[") && RegExMatch(line, "^\s*(.+)\s*\(([^)]+)\)\s*$", &cm)) {
            pfx := CheatSheet_ContextParensToModifierPrefix(cm[2])
            if (pfx != "")
                clusterPrefix := pfx
        }

        if RegExMatch(line, "^([^\[\]]*?)(\[.*?\])(.*)$", &match) {
            emoji := match[1]
            bracket := match[2]
            restOfLine := match[3]
            displayBracket := CheatSheet_MergeClusterIntoBracket(bracket, clusterPrefix)
            paddedShortcut := PadShortcut(displayBracket)

            if (emoji != "") {
                processedLine := emoji . " " . paddedShortcut . " " . restOfLine
            } else {
                processedLine := paddedShortcut . " " . restOfLine
            }

            if (IsBuiltInShortcut(bracket)) {
                processedLine := "--- " . processedLine
            } else {
                processedLine := ">>> " . processedLine
            }

            processedLines.Push(processedLine)
        } else {
            processedLines.Push(line)
        }
    }

    return JoinLines(processedLines)
}

; Single alternation (same semantics as former per-pattern loop).
IsBuiltInShortcut(shortcut) {
    static builtinRe := ""
    if (builtinRe = "") {
        sub := [
            "Ctrl \+ [A-Z]", "Ctrl \+ [0-9]", "Ctrl \+ [F1-F12]", "Alt \+ [A-Z]", "Alt \+ [0-9]", "Alt \+ [F1-F12]",
            "Alt \+ [↑↓←→]", "Ctrl \+ Shift \+ [A-Z]", "Ctrl \+ Shift \+ [0-9]", "Ctrl \+ Enter", "Ctrl \+ Space",
            "Ctrl \+ Tab", "Ctrl \+ Esc", "Ctrl \+ Home", "Ctrl \+ End", "Ctrl \+ PageUp", "Ctrl \+ PageDown",
            "Ctrl \+ Insert", "Ctrl \+ Delete", "Ctrl \+ Backspace", "Shift \+ [A-Z]", "Shift \+ [0-9]",
            "Shift \+ [F1-F12]", "Shift \+ [↑↓←→]", "Shift \+ Enter", "Shift \+ Delete", "Shift \+ Tab", "Shift \+ Esc",
            "F[1-9]|F1[0-2]", "Esc", "Enter", "Space", "Tab", "Backspace", "Delete", "Insert", "Home", "End",
            "PageUp", "PageDown", "↑|↓|←|→"
        ]
        builtinRe := "^(?i)(?:"
        for i, p in sub {
            if (i > 1)
                builtinRe .= "|"
            builtinRe .= "(?:" . p . ")"
        }
        builtinRe .= ")$"
    }
    content := RegExReplace(shortcut, "\[|\]", "")
    return RegExMatch(content, builtinRe)
}

; Helper: normalize common UTF-8→CP1252 mojibake so arrows and punctuation display correctly
NormalizeMojibake(str) {
    if (str = "")
        return str
    reps := Map(
        "â†’", "→",   ; right arrow
        "â†", "←",   ; left arrow
        "â†‘", "↑",   ; up arrow
        "â†“", "↓",   ; down arrow
        "â€¢", "•",
        "â€“", "–",
        "â€”", "—",
        "â€¦", "…",
        "â€˜", "'",
        "â€™", "'",
        "â€œ", Chr(34),
        "â€", Chr(34),
        "Ã—", "×"
    )
    for k, v in reps
        str := StrReplace(str, k, v)
    return str
}

;-------------------------------------------------------------------
; Cheat-sheet overlay (Win + Alt + Shift + A) â€" shows remapped shortcuts
;-------------------------------------------------------------------

; [SK module] cheatSheets map population (Mercado Livre, Shopee) -> Shift keys\cheat_sheet_data.ahk
#include %A_ScriptDir%\Shift keys\cheat_sheet_data.ahk

; [SK module] Per-app Win+Alt+Shift hotkey cheat sheet definitions -> Shift keys\app_hotkeys.ahk
#include %A_ScriptDir%\Shift keys\app_hotkeys.ahk

; [SK module] Cheat sheet GUI, search, hold detection -> Shift keys\cheat_sheet_gui.ahk
#include %A_ScriptDir%\Shift keys\cheat_sheet_gui.ahk

; [SK module] Clip Angel fast copy mode and #!+1/#!+J -> Shift keys\fast_copy_clipangel.ahk
#include %A_ScriptDir%\Shift keys\fast_copy_clipangel.ahk

; [SK module] Env paths, ShowErr, CenterGuiOnActiveMonitor -> Shift keys\env_paths_centergui.ahk
#include %A_ScriptDir%\Shift keys\env_paths_centergui.ahk

; [SK module] OneNote hotkeys -> Shift keys\hotif_onenote.ahk
#include %A_ScriptDir%\Shift keys\hotif_onenote.ahk

; [SK module] ClipAngel hotkeys and filter selector -> Shift keys\hotif_clipangel.ahk
#include %A_ScriptDir%\Shift keys\hotif_clipangel.ahk

; [SK module] WhatsApp desktop hotkeys -> Shift keys\hotif_whatsapp.ahk
#include %A_ScriptDir%\Shift keys\hotif_whatsapp.ahk

;-------------------------------------------------------------------
; Outlook Reminder Window Shortcuts
;-------------------------------------------------------------------
; [SK module] Outlook reminder/appointment hotkeys -> Shift keys\hotif_outlook_reminder.ahk
#include %A_ScriptDir%\Shift keys\hotif_outlook_reminder.ahk

; [SK module] Teams meeting/chat predicate helpers -> Shift keys\teams_predicates.ahk
#include %A_ScriptDir%\Shift keys\teams_predicates.ahk
; [SK module] Teams meeting window hotkeys -> Shift keys\hotif_teams_meeting.ahk
#include %A_ScriptDir%\Shift keys\hotif_teams_meeting.ahk

;-------------------------------------------------------------------
; Wikipedia Shortcuts
;-------------------------------------------------------------------
; Global variable to track scroll position history (stack: most recent last)
global g_WikipediaScrollHistory := []

; [SK module] Wikipedia Chrome hotkeys -> Shift keys\hotif_wikipedia.ahk
#include %A_ScriptDir%\Shift keys\hotif_wikipedia.ahk

;-------------------------------------------------------------------
; Chrome PDF Viewer Shortcuts
;-------------------------------------------------------------------
IsChromePdfViewerActive() {
    ; Hard gate: avoid conflicts with non-Chrome apps
    if !WinActive("ahk_exe chrome.exe")
        return false

    try {
        ; #region agent log
        AgentDebugLog("H1", "IsChromePdfViewerActive_entry")
        ; #endregion
        uia := UIA_Browser("ahk_exe chrome.exe")

        ; Strong fingerprint: Chrome's built-in PDF viewer extension web area
        ; From UIA tree: chrome-extension://mhjfbmdgcfjbbpaeojofohoefgiehjai/index.html
        if (uia.FindElement({ Type: 50030, Value: "chrome-extension://mhjfbmdgcfjbbpaeojofohoefgiehjai", matchmode: "Substring" })) {
            ; #region agent log
            AgentDebugLog("H2", "IsChromePdfViewerActive_extension_match")
            ; #endregion
            return true
        }

        ; Fallback: stable, non-localized PDF toolbar controls
        if (uia.FindElement({ AutomationId: "pageSelector" }) && uia.FindElement({ AutomationId: "save" })) {
            ; #region agent log
            AgentDebugLog("H3", "IsChromePdfViewerActive_toolbar_match")
            ; #endregion
            return true
        }
    } catch {
    }

    ; #region agent log
    AgentDebugLog("H4", "IsChromePdfViewerActive_return_false")
    ; #endregion
    return false
}

; [SK module] Chrome PDF viewer hotkeys -> Shift keys\hotif_chrome_pdf.ahk
#include %A_ScriptDir%\Shift keys\hotif_chrome_pdf.ahk

;-------------------------------------------------------------------
; Mercado Livre (Brazil) Shortcuts
;-------------------------------------------------------------------
; Cache for IsMercadoLivreActive (per efficiency-canon: cache-first with validation).
; Invalidated when foreground HWND changes so we only run UIA once per window/tab focus.
global g_ML_CacheHwnd := 0
global g_ML_CacheResult := false
; Cache for initial-page-load workaround: right-click + close context menu to make hotkeys work (once per window).
global g_ML_ReceptivityHwnd := 0

; Workaround for ML: hotkeys fail on initial page load until page is "initialized". Right-click on empty area then close context menu.
; force=true: run even if cache says we already did (error-driven retry when hotkey action failed).
ML_EnsureHotkeyReceptivity(force := false) {
    global g_ML_ReceptivityHwnd
    if !WinActive("ahk_exe chrome.exe")
        return
    hwnd := WinExist("A")
    if (!hwnd)
        return
    if (!force && hwnd = g_ML_ReceptivityHwnd && WinExist("ahk_id " g_ML_ReceptivityHwnd))
        return
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        try
            root := uia.GetCurrentDocumentElement()
        catch
            root := uia.BrowserElement
        if (!root)
            return
        br := root.BoundingRectangle
        if (!br || (br.r <= br.l) || (br.b <= br.t))
            return
        x := br.l + (br.r - br.l) * 0.15
        y := br.t + (br.b - br.t) * 0.20
        prevMode := A_CoordModeMouse
        CoordMode("Mouse", "Screen")
        MouseClick("Right", x, y)
        SendEscape()
        CoordMode("Mouse", prevMode)
        g_ML_ReceptivityHwnd := hwnd
    } catch {
    }
}

IsMercadoLivreActive() {
    global g_ML_CacheHwnd, g_ML_CacheResult
    if !WinActive("ahk_exe chrome.exe")
        return false
    hwnd := WinExist("A")
    if (!hwnd)
        return false
    ; Cache hit: same window as last check (avoids UIA on every keystroke / cheat sheet open)
    if (hwnd = g_ML_CacheHwnd && WinExist("ahk_id " g_ML_CacheHwnd))
        return g_ML_CacheResult
    ; Platform identification by URL only (do not use window title; it changes to product name). See shopping uia3.md.
    ; URL check via UIA (address bar: Chrome exposes AcceleratorKey "Ctrl+L", not AccessKey). Bounded to this window only.
    try {
        root := UIA.ElementFromHandle(hwnd)
        addressBar := root.FindFirst({ Type: 50004, AcceleratorKey: "Ctrl+L" })
        if (addressBar) {
            url := addressBar.Value
            if InStr(url, "mercadolivre.com") || InStr(url, "mercadolibre.com") {
                g_ML_CacheHwnd := hwnd
                g_ML_CacheResult := true
                return true
            }
        }
    } catch {
        ; UIA failed; do not cache so next call retries
    }
    g_ML_CacheHwnd := hwnd
    g_ML_CacheResult := false
    return false
}

; Mercado Livre UIA helpers: get document root and find/invoke elements (bounded, no global state).
; When UIA_Browser init fails (document not ready), fallback: get Document from foreground window tree (Type 50030 = Document).
ML_GetDocRoot() {
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        try
            return uia.GetCurrentDocumentElement()
        catch
            return uia.BrowserElement
    } catch {
        ; Fallback: document not ready for UIA_Browser init (e.g. on first load). Get Document from active window.
        try {
            hwnd := WinExist("A")
            if (hwnd && WinActive("ahk_exe chrome.exe")) {
                root := UIA.ElementFromHandle(hwnd)
                doc := root.FindFirst({ Type: 50030 })
                if (doc)
                    return doc
            }
        } catch {
        }
        return 0
    }
}

; #region agent log - debug helper
ML_DebugLog(hypothesisId, message, detail, runId := "pre-fix") {
    ; Intentionally no-op.
    return
}
; #endregion agent log

; Locate the Mercado Livre Preço filter container, scoped to the block that
; contains the predefined price links (Até R$ 40 / R$ 40 a R$ 95 / Mais de R$ 95)
; and the Mínimo/Máximo edits. Returns the container element or 0.
ML_GetPriceFilter(root) {
    if !root
        return 0
    priceLinkNames := ["Até R$ 40", "R$ 40 a R$ 95", "Mais de R$ 95"]
    for _, name in priceLinkNames {
        link := ML_Find(root, { Type: 50005, Name: name, cs: false })
        if !link
            continue
        parent := link
        ; Walk up a few levels to find a container that has the Mínimo edit as a descendant.
        loop 6 {
            try parent := parent.Parent
            catch
                break
            if !parent
                break
            minimo := ML_Find(parent, { Type: 50004, Name: "Mínimo", cs: false })
            if (minimo) {
                return parent
            }
        }
    }
    return 0
}

; Best-effort helper to set text of an edit element. Tries ValuePattern first,
; then falls back to focus + clear + send keys.
ML_SetEditText(el, text) {
    if (!el)
        return false
    if (text = "")
        return false
    ok := false
    try {
        if el.GetPropertyValue(UIA.Property.IsValuePatternAvailable) {
            vp := el.ValuePattern
            vp.SetValue(text)
            ok := true
        }
    } catch {
        ML_DebugLog("C", "ValuePattern attempt failed", "exception in SetValue", "run1")
    }
    if (!ok) {
        try el.ScrollIntoView()
        catch {
        }
        try {
            el.SetFocus()
        } catch {
            try el.Click()
            catch {
            }
        }
        Sleep 50
        ML_DebugLog("C", "Focus acquired for edit", text, "run1")
        Send "^a{Del}"
        Sleep 30
        SendText(text)
        ok := true
    }
    return ok
}

; Try conditions in order; invoke or click first match. Returns true if invoked/clicked, false otherwise.
ML_FindAndInvoke(conditionList) {
    root := ML_GetDocRoot()
    if (!root)
        return false
    for cond in conditionList {
        try {
            el := root.FindElement(cond, UIA.TreeScope.Descendants)
            if (el) {
                try el.Invoke()
                catch {
                    try el.Click()
                    catch
                        return false
                }
                return true
            }
        } catch
            continue
    }
    return false
}

; Find single element by condition (Descendants). Returns element or 0.
ML_Find(root, condition) {
    try
        return root.FindElement(condition, UIA.TreeScope.Descendants)
    catch
        return 0
}

;-------------------------------------------------------------------
; Shopee (Brazil) detection and UIA helpers
;-------------------------------------------------------------------
; Cache for IsShopeeActive (same pattern as IsMercadoLivreActive).
global g_Shopee_CacheHwnd := 0
global g_Shopee_CacheResult := false

IsShopeeActive() {
    global g_Shopee_CacheHwnd, g_Shopee_CacheResult
    if !WinActive("ahk_exe chrome.exe")
        return false
    hwnd := WinExist("A")
    if (!hwnd)
        return false
    ; Cache hit: same window as last check (avoids UIA on every keystroke / cheat sheet open)
    if (hwnd = g_Shopee_CacheHwnd && WinExist("ahk_id " g_Shopee_CacheHwnd))
        return g_Shopee_CacheResult
    ; Platform identification by URL only (do not use window title; it changes to product name). See shopping uia3.md.
    ; URL check via UIA (Chrome address bar: AcceleratorKey "Ctrl+L", not AccessKey)
    try {
        root := UIA.ElementFromHandle(hwnd)
        addressBar := root.FindFirst({ Type: 50004, AcceleratorKey: "Ctrl+L" })
        if (addressBar) {
            url := addressBar.Value
            if InStr(url, "shopee.com", false) {
                g_Shopee_CacheHwnd := hwnd
                g_Shopee_CacheResult := true
                return true
            }
        }
    } catch {
        ; UIA failed; do not cache so next call retries
    }
    g_Shopee_CacheHwnd := hwnd
    g_Shopee_CacheResult := false
    return false
}

Shopee_GetDocRoot() {
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        try
            return uia.GetCurrentDocumentElement()
        catch
            return uia.BrowserElement
    } catch
        return 0
}

Shopee_Find(root, condition) {
    try
        return root.FindElement(condition, UIA.TreeScope.Descendants)
    catch
        return 0
}

Shopee_FindAndInvoke(conditionList) {
    root := Shopee_GetDocRoot()
    if (!root)
        return false
    for cond in conditionList {
        try {
            el := root.FindElement(cond, UIA.TreeScope.Descendants)
            if (el) {
                try el.Invoke()
                catch {
                    try el.Click()
                    catch
                        return false
                }
                return true
            }
        } catch
            continue
    }
    return false
}

; Navigate Shopee search results pagination by relative offset (e.g. +1 next, -1 previous).
Shopee_NavMove(offset) {
    if (offset = 0)
        return false
    hwnd := WinExist("A")
    if (!hwnd)
        return false

    ; Get current page index from Chrome address bar (?page=N), defaulting to 0 when absent.
    currentPage := 0
    try {
        rootWin := UIA.ElementFromHandle(hwnd)
        addressBar := rootWin.FindFirst({ Type: 50004, AcceleratorKey: "Ctrl+L" })
        if (addressBar) {
            url := addressBar.Value
            m := ""
            if RegExMatch(url, "i)[?&]page=(\d+)", &m) {
                try currentPage := Integer(m[1])
            }
        }
    } catch {
    }

    root := Shopee_GetDocRoot()
    if (!root)
        return false

    navGroup := Shopee_Find(root, { Type: 50026, Name: "Navegação entre páginas", cs: false })
    if (!navGroup)
        return false

    links := 0
    try links := navGroup.FindAll({ Type: 50005 })
    catch {
        return false
    }
    if (!links)
        return false

    bestEl := 0
    bestDelta := 0x7FFFFFFF

    for link in links {
        value := ""
        try value := link.Value
        catch {
            continue
        }
        if (value = "")
            continue
        m2 := ""
        if !RegExMatch(value, "i)[?&]page=(\d+)", &m2)
            continue
        targetPage := 0
        try targetPage := Integer(m2[1])
        catch {
            continue
        }

        ; Choose the nearest page ahead (offset>0) or behind (offset<0) relative to currentPage.
        if (offset > 0 && targetPage > currentPage) {
            delta := targetPage - currentPage
            if (delta < bestDelta) {
                bestDelta := delta
                bestEl := link
            }
        } else if (offset < 0 && targetPage < currentPage) {
            delta := currentPage - targetPage
            if (delta < bestDelta) {
                bestDelta := delta
                bestEl := link
            }
        }
    }

    if (!bestEl)
        return false

    try {
        bestEl.Invoke()
        return true
    } catch {
        try {
            bestEl.Click()
            return true
        } catch {
            return false
        }
    }
}

; [SK module] Mercado Livre hotkeys -> Shift keys\hotif_mercado_livre.ahk
#include %A_ScriptDir%\Shift keys\hotif_mercado_livre.ahk
; [SK module] Shopee hotkeys -> Shift keys\hotif_shopee.ahk
#include %A_ScriptDir%\Shift keys\hotif_shopee.ahk

;-------------------------------------------------------------------
; Microsoft Teams Shortcuts (chat)
;-------------------------------------------------------------------
; [SK module] Teams chat window hotkeys and UIA -> Shift keys\hotif_teams_chat.ahk
#include %A_ScriptDir%\Shift keys\hotif_teams_chat.ahk

; [SK module] Outlook helper functions (part 1) -> Shift keys\outlook_helpers_01.ahk
#include %A_ScriptDir%\Shift keys\outlook_helpers_01.ahk
; [SK module] Outlook helper functions (part 2) -> Shift keys\outlook_helpers_02.ahk
#include %A_ScriptDir%\Shift keys\outlook_helpers_02.ahk

; [SK module] Outlook main window hotkeys -> Shift keys\hotif_outlook_main.ahk
#include %A_ScriptDir%\Shift keys\hotif_outlook_main.ahk

; Appointment/Meeting inspector-specific hotkeys
; [SK module] Outlook appointment inspector hotkeys (part 1) -> Shift keys\outlook_appointment_hotif_01.ahk
#include %A_ScriptDir%\Shift keys\outlook_appointment_hotif_01.ahk

; [SK module] Outlook appointment date/time helpers and hotkeys -> Shift keys\outlook_appointment_hotif_02.ahk
#include %A_ScriptDir%\Shift keys\outlook_appointment_hotif_02.ahk

; [SK module] Outlook appointment configuration palette -> Shift keys\outlook_appointment_palette.ahk
#include %A_ScriptDir%\Shift keys\outlook_appointment_palette.ahk

; [SK module] Outlook appointment UIA state checking -> Shift keys\outlook_appointment_uia.ahk
#include %A_ScriptDir%\Shift keys\outlook_appointment_uia.ahk

#HotIf

; [SK module] Google Chrome general hotkeys -> Shift keys\hotif_chrome_general.ahk
#include %A_ScriptDir%\Shift keys\hotif_chrome_general.ahk

; [SK module] ChatGPT hotkeys -> Shift keys\hotif_chatgpt.ahk
#include %A_ScriptDir%\Shift keys\hotif_chatgpt.ahk

;-------------------------------------------------------------------
; Settings Window Shortcuts
;-------------------------------------------------------------------
; [SK module] Windows Settings hotkeys -> Shift keys\hotif_settings.ahk
#include %A_ScriptDir%\Shift keys\hotif_settings.ahk
#HotIf

;-------------------------------------------------------------------
; Windows Explorer Shortcuts
;-------------------------------------------------------------------
; [SK module] Windows Explorer hotkeys -> Shift keys\hotif_explorer.ahk
#include %A_ScriptDir%\Shift keys\hotif_explorer.ahk
#HotIf

;-------------------------------------------------------------------
; Microsoft Paint Shortcuts
;-------------------------------------------------------------------
; [SK module] Excel and Paint hotkeys -> Shift keys\hotif_excel_mspaint.ahk
#include %A_ScriptDir%\Shift keys\hotif_excel_mspaint.ahk
#HotIf

;-------------------------------------------------------------------
; Power BI Shortcuts
;-------------------------------------------------------------------
; [SK module] Power BI hotkeys -> Shift keys\hotif_powerbi.ahk
#include %A_ScriptDir%\Shift keys\hotif_powerbi.ahk
#HotIf

; [SK module] Power BI drawer config helpers -> Shift keys\powerbi_helpers.ahk
#include %A_ScriptDir%\Shift keys\powerbi_helpers.ahk

;-------------------------------------------------------------------
; Gmail Shortcuts
;-------------------------------------------------------------------
; [SK module] Gmail hotkeys -> Shift keys\hotif_gmail.ahk
#include %A_ScriptDir%\Shift keys\hotif_gmail.ahk
#HotIf

; [SK module] Cursor/VS Code editor detection and UIA helpers -> Shift keys\cursor_predicates.ahk
#include %A_ScriptDir%\Shift keys\cursor_predicates.ahk
; [SK module] Cursor IDE hotkeys -> Shift keys\hotif_cursor.ahk
#include %A_ScriptDir%\Shift keys\hotif_cursor.ahk
#HotIf

; Shared Editor Shortcuts (Cursor + VS Code)
;-------------------------------------------------------------------
; [SK module] Cursor/VS Code editor hotkeys (part 1) -> Shift keys\hotif_editor_01.ahk
#include %A_ScriptDir%\Shift keys\hotif_editor_01.ahk
; [SK module] Cursor/VS Code editor hotkeys (part 2) -> Shift keys\hotif_editor_02.ahk
#include %A_ScriptDir%\Shift keys\hotif_editor_02.ahk
#HotIf

VSCode_TriggerGenerateCommitMessage(hwnd := 0) {
    if (!hwnd)
        hwnd := WinExist("ahk_exe Code.exe")
    if (!hwnd)
        return false

    ; Ensure Source Control view is focused so the Generate button is rendered.
    FocusSourceControlViewForCommitGeneration()

    try root := UIA.ElementFromHandle(hwnd)
    catch
        return false
    if (!root)
        return false

    genBtn := 0
    for name in ["Generate Commit Message (Ctrl+Alt+.)", "Generate Commit Message"] {
        try {
            genBtn := root.FindFirst({ Type: 50000, Name: name })
        } catch {
            genBtn := 0
        }
        if (genBtn)
            break
    }

    if (!genBtn) {
        try {
            allButtons := root.FindAll({ Type: 50000 })
            if (allButtons) {
                for btn in allButtons {
                    try {
                        nm := btn.Name
                        if (InStr(nm, "Generate Commit Message")) {
                            genBtn := btn
                            break
                        }
                    } catch {
                        continue
                    }
                }
            }
        } catch {
            genBtn := 0
        }
    }

    if (!genBtn)
        return false

    try {
        if (genBtn.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)) {
            genBtn.InvokePattern.Invoke()
            return true
        }
    } catch {
    }

    try {
        genBtn.Click()
        return true
    } catch {
        return false
    }
}

; VS Code IDE — VS Code-specific Shortcuts
;-------------------------------------------------------------------
; [SK module] VS Code hotkeys -> Shift keys\hotif_code.ahk
#include %A_ScriptDir%\Shift keys\hotif_code.ahk
#HotIf

; [SK module] Global Alt+U scroll AI feed and related -> Shift keys\hotif_scroll_ai.ahk
#include %A_ScriptDir%\Shift keys\hotif_scroll_ai.ahk

;-------------------------------------------------------------------
; Spotify Shortcuts
;-------------------------------------------------------------------
; [SK module] Spotify hotkeys -> Shift keys\hotif_spotify.ahk
#include %A_ScriptDir%\Shift keys\hotif_spotify.ahk
#HotIf

;-------------------------------------------------------------------
; Figma Shortcuts
;-------------------------------------------------------------------
; [SK module] Figma hotkeys -> Shift keys\hotif_figma.ahk
#include %A_ScriptDir%\Shift keys\hotif_figma.ahk
#HotIf

; [SK module] Mobills title WinActive hotkeys -> Shift keys\hotif_mobills.ahk
#include %A_ScriptDir%\Shift keys\hotif_mobills.ahk

; K/L month navigation: see #HotIf Mobills_ShouldHandleMonthNavKeys() below (single definition; skips text fields).

; [SK module] Mobills pagination unified -> Shift keys\mobills_pagination.ahk
#include %A_ScriptDir%\Shift keys\mobills_pagination.ahk

; [SK module] Mobills running overlay banner -> Shift keys\mobills_running_banner.ahk
#include %A_ScriptDir%\Shift keys\mobills_running_banner.ahk

; [SK module] Mobills URL-scoped month nav hotkeys -> Shift keys\mobills_hotkeys_fallback.ahk
#include %A_ScriptDir%\Shift keys\mobills_hotkeys_fallback.ahk

; [SK module] Google Keep hotkeys and reminder dismiss helpers -> Shift keys\hotif_google_keep.ahk
#include %A_ScriptDir%\Shift keys\hotif_google_keep.ahk
; [SK module] YouTube Chrome hotkeys -> Shift keys\hotif_youtube.ahk
#include %A_ScriptDir%\Shift keys\hotif_youtube.ahk

;-------------------------------------------------------------------
; Gemini Website Shortcuts
;-------------------------------------------------------------------
; [SK module] Gemini Chrome hotkeys (part 1) -> Shift keys\gemini_chrome_01.ahk
#include %A_ScriptDir%\Shift keys\gemini_chrome_01.ahk
; [SK module] Gemini Chrome tools drawer and hotkeys (part 2) -> Shift keys\gemini_chrome_02.ahk
#include %A_ScriptDir%\Shift keys\gemini_chrome_02.ahk

;-------------------------------------------------------------------
; M365 Copilot web (Chrome) — same Shift shortcuts as Gemini
;-------------------------------------------------------------------
; [SK module] M365 Copilot web Chrome hotkeys -> Shift keys\hotif_copilot_web.ahk
#include %A_ScriptDir%\Shift keys\hotif_copilot_web.ahk

;-------------------------------------------------------------------
; Google Maps Shortcuts (Chrome)
;-------------------------------------------------------------------
; [SK module] Google Maps Chrome hotkeys -> Shift keys\hotif_google_maps.ahk
#include %A_ScriptDir%\Shift keys\hotif_google_maps.ahk

;-------------------------------------------------------------------
; Google Search Shortcuts
;-------------------------------------------------------------------
; [SK module] Google Search Chrome hotkeys -> Shift keys\hotif_google_search.ahk
#include %A_ScriptDir%\Shift keys\hotif_google_search.ahk

;-------------------------------------------------------------------
; File Dialog (Namespace Tree Control) Shortcuts
;-------------------------------------------------------------------
; [SK module] File dialog hotkeys -> Shift keys\hotif_file_dialog.ahk
#include %A_ScriptDir%\Shift keys\hotif_file_dialog.ahk

IsFileDialogActive() {
    hwnd := WinActive("A")
    if !hwnd {
        return false
    }

    winClass := WinGetClass("ahk_id " hwnd)
    winTitle := WinGetTitle("ahk_id " hwnd)
    winExe := WinGetProcessName("ahk_id " hwnd)
    if winClass != "#32770" {
        return false
    }

    ; Exclude Outlook Reminders which can share dialog-like traits (classic + new Outlook / olk.exe)
    try {
        if (winExe = "OUTLOOK.EXE" || StrLower(winExe) = "olk.exe") {
            if RegExMatch(winTitle, "i)Reminder") {
                return false
            }
        }
    } catch Error {
    }

    try {
        root := UIA.ElementFromHandle(hwnd)

        ; Check UIA properties: Type = Window (50032) and LocalizedType = "dialog"
        rootType := ""
        rootLocalizedType := ""
        rootName := ""
        try rootType := root.Type
        try rootLocalizedType := root.LocalizedType
        try rootName := root.Name

        ; Primary check: Type must be Window and LocalizedType must be "dialog"
        if (rootType = UIA.Type.Window && rootLocalizedType = "dialog") {
            return true
        }

        ; Fallback 1: Check if window name matches common file dialog names
        ; This handles cases where LocalizedType might vary by language/application
        if (rootType = UIA.Type.Window) {
            fileDialogNames := ["Save As", "Open", "Browse", "Select File", "Choose File",
                "Salvar como", "Abrir", "Procurar", "Selecionar arquivo",
                "Guardar como", "Guardar", "Explorar"]
            for dialogName in fileDialogNames {
                if (InStr(rootName, dialogName, false)) {
                    return true
                }
            }
        }

        ; Fallback 2: Check for file dialog characteristic elements (Namespace Tree Control, file lists, etc.)
        ; This is a last resort if UIA properties don't match but it's still a file dialog
        try {
            ; Check for Namespace Tree Control in window text (original method as backup)
            txt := WinGetText("ahk_id " hwnd)
            if InStr(txt, "Namespace Tree Control") || InStr(txt, "Controle da Ãrvore de Namespace") {
                return true
            }
        } catch {
            ; Window text check failed, continue
        }

        ; If we get here, it's not a file dialog
        return false
    } catch Error as e {
        return false
    }
}

;-------------------------------------------------------------------
; UIA Tree Inspector Shortcuts
;-------------------------------------------------------------------

; FindFirst/FindElement throw TargetError when no match; FindAll returns []. Use FindAll fallback for reliability.
UIATreeInspector_FindTreeByAutomationId(root, automationId) {
    if (!root)
        return 0
    aid := String(automationId)
    try
        return root.FindFirst({ Type: UIA.Type.Tree, AutomationId: aid })
    catch TargetError {
    }
    try {
        trees := root.FindAll({ Type: UIA.Type.Tree })
        for t in trees {
            if !t
                continue
            try {
                if (String(t.AutomationId) = aid)
                    return t
            } catch {
            }
        }
    } catch {
    }
    return 0
}

UIATreeInspector_FindRightDumpTree(root) {
    t := UIATreeInspector_FindTreeByAutomationId(root, "17")
    if (t)
        return t
    try {
        trees := root.FindAll({ Type: UIA.Type.Tree })
        for x in trees {
            if !x
                continue
            try {
                if (x.Name = "UIA Tree")
                    return x
            } catch {
            }
        }
    } catch {
    }
    bestL := -0x7FFFFFFF
    bestT := 0
    try {
        trees := root.FindAll({ Type: UIA.Type.Tree })
        for t in trees {
            if (!t)
                continue
            try {
                if (String(t.AutomationId) = "4")
                    continue
            } catch {
                continue
            }
            try {
                br := t.BoundingRectangle
                if (br.l > bestL) {
                    bestL := br.l
                    bestT := t
                }
            } catch {
                if (!bestT)
                    bestT := t
            }
        }
    } catch {
    }
    return bestT
}

; Win32 focus on left SysTreeView32 (AutomationId 4, TVWins) so UIA selection and arrow keys stay in sync.
UIATreeInspector_FocusLeftWindowsTree(treeContainer, winHwnd) {
    leftHwnd := 0
    if (!treeContainer || !winHwnd)
        return 0
    try
        leftHwnd := treeContainer.NativeWindowHandle
    catch {
        leftHwnd := 0
    }
    if (leftHwnd) {
        try
            ControlFocus "ahk_id " leftHwnd, "ahk_id " winHwnd
        catch {
            ; Invalid HWND pair or control not targetable; SetFocus below may still work.
        }
    }
    try
        treeContainer.SetFocus()
    catch {
        ; UIA SetFocus can surface COM/Win32 errors (e.g. "Target window not found"); non-fatal.
    }
    return leftHwnd
}

; Jiggle selection with keys guaranteed to go to the windows list TreeView (not filter / middle panels).
UIATreeInspector_JiggleLeftTree(leftHwnd, winHwnd, downDelayMs) {
    if (!WinExist("ahk_id " winHwnd))
        return
    if (!WinActive("ahk_id " winHwnd))
        return
    if (leftHwnd) {
        try {
            ControlSend "{Down}", "ahk_id " leftHwnd, "ahk_id " winHwnd
            Sleep downDelayMs
            ControlSend "{Up}", "ahk_id " leftHwnd, "ahk_id " winHwnd
        } catch {
            Send "{Down}"
            Sleep downDelayMs
            Send "{Up}"
        }
    } else {
        Send "{Down}"
        Sleep downDelayMs
        Send "{Up}"
    }
}

; [SK module] UIA Tree Inspector hotkeys -> Shift keys\hotif_uia_tree.ahk
#include %A_ScriptDir%\Shift keys\hotif_uia_tree.ahk

;-------------------------------------------------------------------
; [SK module] Settle Up hotkeys -> Shift keys\hotif_settleup.ahk
#include %A_ScriptDir%\Shift keys\hotif_settleup.ahk

;-------------------------------------------------------------------
; [SK module] Miro Chrome hotkeys -> Shift keys\hotif_miro.ahk
#include %A_ScriptDir%\Shift keys\hotif_miro.ahk

;-------------------------------------------------------------------
; PowerToys Command Palette Shortcuts
;-------------------------------------------------------------------
; [SK module] Command Palette hotkeys -> Shift keys\hotif_command_palette.ahk
#include %A_ScriptDir%\Shift keys\hotif_command_palette.ahk

; [SK module] ChatGPT loading banner and wait helpers -> Shift keys\chatgpt_loading_helpers.ahk
#include %A_ScriptDir%\Shift keys\chatgpt_loading_helpers.ahk

; [SK module] TEMPORARY M365 Copilot auto-continue (^!#n) -> Shift keys\m365_copilot_temp.ahk
#include %A_ScriptDir%\Shift keys\m365_copilot_temp.ahk

; VS Code evidence -> PDF search loop (^!#o) — see VSCodeEvidenceSearch.ahk
global EVIDENCE_SEARCH_FROM_SHIFT_KEYS := true
#include %A_ScriptDir%\VSCodeEvidenceSearch.ahk
#InputLevel 10
EvidenceSearch_BindHotkey()