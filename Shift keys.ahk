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

; Map that stores the pop-up text for each application.  Extend freely.
cheatSheets := Map()

; --- Mercado Livre (Brazil) -----------------------------------------------
cheatSheets["Mercado Livre"] := "
(
    Mercado Livre (Shift)
    [S] Focus search field
    [C] Carrinho de compras (cart)
    [P] Compras feitas (purchases)
    [Y] Filtro Chegará amanhã
    [F] Filtro Full
    [I] Filtro Internacional
    [N] Filtro Envio local / Nacional
    [G] Filtro Frete grátis
    [O] Ordenar por (menu)
    [R] Faixa de preço (Mín/Máx)
    [L] Seguinte (paginação)
    [K] Anterior (paginação)
    [A] Adicionar ao carrinho
    [V] Favoritos (coração)
    [J] Continuar (fluxo compra/endereço)
)"  ; end Mercado Livre

; --- Shopee (Brazil) -------------------------------------------------------
cheatSheets["Shopee"] := "
(
    Shopee (Shift)
    [S] Buscar na Shopee (campo de busca)
    [C] Carrinho de compras
    [P] Minhas compras / Pedidos (especulativo)
    [Y] Filtro Entrega Rápida (analogia Chegará amanhã)
    [F] Filtro Promoções / Full (especulativo)
    [I] Filtro Internacional
    [N] Filtro Envio Nacional
    [G] Filtro Frete grátis (especulativo)
    [O] Ordenar por (menu)
    [R] Faixa de preço (Mín/Máx)
    [L] Seguinte (paginação) – especulativo
    [K] Anterior (paginação) – especulativo
    [A] Adicionar ao carrinho (página do produto)
    [V] Favoritar produto (coração)
    [J] Continuar (carrinho/checkout, incl. \"Fazer pedido\")
)"  ; end Shopee

;---------------------------------------- Shift + keys ----------------------------------------------
; ----- Assignment policy: use Shift + <key> first. When all Shift slots in the sequence are consumed, continue with Ctrl + Alt + <key> in the same order.
; ----- You can have repeated keys, depending on the software.
; ----- Preferred key sequence (most important first): Y U I O P H J K L N M , . W E R T D F G C V B
; ----- Ctrl + Alt sequence (fallback, same order):    Y U I O P H J K L N M , . W E R T D F G C V B
; ----- Shift + D (Teams chat) -> Fold chat sections (🩶 grey)

; --- WhatsApp desktop -------------------------------------------------------
cheatSheets["WhatsApp"] := "
(
    WhatsApp (Shift)
    🎤 [V]Toggle [V]oice message
    🔍 [S][S]earch chats
    ↩️ [R][R]eply
    😀 [E][E]moji panel
    📬 [U]Toggle [U]nread filter
    💬 [F][F]ocus current chat
    ✅ [M][M]ark as read/unread
    📌 [P][P]in chat or unpin
)"  ; end WhatsApp

; --- Outlook main window ----------------------------------------------------
cheatSheets["OUTLOOK.EXE"] := "
(
    Outlook (Shift)
    📧 [G]Send to [G]eneral
    📰 [N]Send to [N]ewsletter
    📥 [I]Go to [I]nbox
    🥇 [J][J]ump to first mail
    ◧ [H]Toggle high [H] navigation pane
    📝 [S][S]ubject / Title
    👥 [T][T]o / Required
    📝 [B][B]ody (Subject → Body)
    🎯 [F][F]ocused / Other
    🔀 [K]Cycle bac[K]ward pane
    🔀 [L]Cyc[L]e forward pane
    📋 [M]Toggle Mail / Calendar
    📅 [W]eek view
    📅 M[o]nth view
    🪟 [P][P]op Out (Type 50000, Name Pop Out, LocalizedType button, Path {T:33,CN:rctrl_renwnd32}, {T:33}, {T:33}, {T:21}, {T:0, i:-1})
    
    📅 Meeting request (reading pane or popped-out invitation - same shortcuts)
    ✅ [A][A]ccept meeting invitation
    ❌ [D][D]ecline meeting (confirmation first)
    📌 [Alt+F] Follow (updates from organizer)
    ❓ [T][T]entative (More options …) - when a meeting request is open, runs before [T]o / Required
    
    📅 Canceled meeting (organizer canceled — Remove event)
    🗑️ [E]Remove event (no confirmation)
    
    Outlook (Ctrl+Alt)
    📌 Ribbon actions select the Home tab first when needed (e.g. View or Help was active).
    🔎 [Ctrl+Alt+F] Search
    📮 [Ctrl+Alt+M] Mail view
    📅 [Ctrl+Alt+G] Calendar view
    📃 [Ctrl+Alt+L] Focus message list
    📖 [Ctrl+Alt+P] Focus reading pane
    
    ↩️ [Ctrl+Alt+R] Reply
    👥 [Ctrl+Alt+A] Reply all
    ➡️ [Ctrl+Alt+W] Forward
    🗑️ [Ctrl+Alt+D] Delete
    🗄️ [Ctrl+Alt+E] Archive
    ✅ [Ctrl+Alt+U] Read/Unread
    🏷️ [Ctrl+Alt+C] Categorize
    📁 [Ctrl+Alt+V] Move
    🧪 [Ctrl+Alt+I] Filter menu
    ↕️ [Ctrl+Alt+S] Sort menu
    
    🆕 [Ctrl+Alt+N] New (Calendar: New event / Mail: new message)
    🧭 [Ctrl+Alt+T] Today (Calendar)
)"  ; end Outlook

; --- Outlook Reminder window -------------------------------------------------
cheatSheets["OutlookReminder"] := "
(
    Outlook - Reminders (Shift)
    ⏰ [H]Snooze 1 [H]our
    ⏰ [F]Snooze [F]our hours
    ⏰ [T]Snooze 10 minu[T]es
    ⏰ [Y]Snooze 1 da[Y]
    ⏰ [W]Snooze 1 [W]eek
    ✅ [D][D]ismiss reminder (selected item)
    ❌ [X]Dismiss all reminders (confirm)
    🌐 [J][J]oin online (selected item)
)"  ; end Outlook Reminder

; --- Outlook Appointment window ---------------------------------------------
cheatSheets["OutlookAppointment"] := "
(
    Outlook - Appointment (Shift)
    📅 [S][S]tart date (popover)
    🕐 [T]Start [T]ime (popover)
    🕐 [E][E]nd time (popover)
    ☑️ [A][A]ll-day toggle (popover)
    🔄 [C]Re[C]urring / Series (popover)
    🕒 [1]Time suggestion 1
    🕒 [2]Time suggestion 2
    
    📝 [I]T[I]tle field
    👥 [R][R]equired attendees
    📍 [O]L[o]cation / Add a room
    📝 [B][B]ody (main details)
    
    🎥 [M]Tea[M]s meeting
    🧬 [U]Series (recurring)
    📶 A[V]ailability (Free/Busy...)
    ⏰ Reminder (fre[Q]uency)
    🏷️ Cate[G]ory
    🔒 [P]rivate / Not private
    
    🗓️ Prev day [K] / Next day [L]
    🧭 [Y]Today
    📆 [D]ate header
    🧑‍🤝‍🧑 Sc[H]eduler / Scheduling assistant
    ⚙️ Optio[N]s (scheduler view)
    🌐 Time [Z]one
    ➕ [J]Add required attendee
    ➕ [Alt+O]Add optional attendee
    
    ↩️ [Backspace] Back (scheduler view)
    🧙 [W][W]izard (configure)
)"  ; end Outlook Appointment

; --- Outlook Message window ---------------------------------------------------
cheatSheets["OutlookMessage"] := "
(
    Outlook - Message (Shift)
    📅 Meeting invitation (same shortcuts as main Mail reading pane)
    ✅ [A][A]ccept meeting invitation
    ❌ [D][D]ecline meeting (confirmation first)
    📌 [Alt+F] Follow (updates from organizer)
    ❓ [T][T]entative (More options …) - when a meeting request is open, runs before [T]o / Required
    📅 Canceled meeting (Remove event)
    🗑️ [E]Remove event (no confirmation)
    📝 [S][S]ubject / Title
    👥 [T][T]o / Required
    📝 [B][B]ody (Location → Body)
)"  ; end Outlook Message

; --- Microsoft Teams â€" meeting window --------------------------------------
cheatSheets["TeamsMeeting"] := "
(
    Teams - Meeting (Shift)
    💬 [C]Open [C]hat pane
    ⛶ [M]aximize [M]eeting window
    👍 [R]eact / [R]eagir
    🎥 [J][J]oin now (camera + mic on)
    🔊 [A][A]udio settings
)"  ; end TeamsMeeting

; --- Microsoft Teams â€" chat window -----------------------------------------
cheatSheets["TeamsChat"] := "
(
    Teams - Chat (Shift)
    🔙 [K]Back (toolbar)
    ⏩ [L]Forward (toolbar)
    ↩️ [R][R]eply
    📬 [U]View all [U]nread items
    📌 [P][P]in chat
    ✏️ [E][E]dit message
    📎 [A][A]ttach file
    📜 [H][H]istory menu
    📬 [M][M]ark unread
    📌 [X]Unpin (e[X]it pin)
    📁 [C][C]ollapse all folders
    ℹ️ [I][I]nfo / Details panel
    🪟 [.]Detach chat (new [.]window)
    👥 [T][T]eam / Add participants
    📞 [V][V]ideo call
    🩶 [F][F]old chat sections
    👍 [Y] Like reaction
    ❤️ [G][G]ive heart reaction
    😂 [J][J]oke reaction (😂)
    
    --- Search Field (Top) ---
    🔍 [Alt+1]Select 1st search result (↓↓ Enter)
    🔍 [Alt+2]Select 2nd search result (↓↓↓ Enter)
    🔍 [Alt+3]Select 3rd search result (↓↓↓↓ Enter)
    🔍 [Alt+4]Select 4th search result (↓↓↓↓↓ Enter)
    🔍 [Alt+5]Select 5th search result (↓↓↓↓↓↓ Enter)
    
    --- Built-in Shortcuts ---
    Geral:
    [Ctrl + .] > Show keyboard shortcuts
    [Ctrl + E] > Open search
    [Ctrl + /] > Show commands
    [Ctrl + G] > Go to a chat or channel
    [Ctrl + N] > Start new chat
    [Ctrl + Shift + N] > Open a new chat
    [Ctrl + Shift + F] > Open filter
    [Ctrl + ,] > Open Settings
    [F1] > Open Help
    [Ctrl + =] > Zoom in
    [Ctrl + -] > Zoom out
    [Ctrl + 0] > Reset zoom level
    [Ctrl + O] > Open existing conversation in new window
    
    Navegação:
    [Ctrl + 1-9] > Open 1st-9th App in App Bar
    [Ctrl + L] > Move focus to left rail item
    [Ctrl + M] > Move focus to messages panel
    [Ctrl + Alt + T] > Move focus to top system notification
    [Alt + Left] > Back
    [Alt + Right] > Forward
    [Ctrl + H] > Open history menu
    [Ctrl + R] > Go to text box
    [Ctrl + Alt + Enter] > Focus on resizable divider
    [Ctrl + Shift + Enter] > Reset slots to default size
    [Win + Shift + Y] > Move focus to notification
    
    Redigir:
    [Ctrl + Shift + X] > Expand text box
    [Ctrl + Enter] > Send (expanded text box)
    [Alt + Shift + O] > Attach file
    [Shift + Enter] > Start new line
    [Ctrl + B] > Apply bold style
    [Ctrl + I] > Apply italic style
    [Ctrl + U] > Apply underline style
    [Alt + A] > Rewrite with Copilot
    [Alt + Shift + E] > Open video recorder
    [Ctrl + Alt + L] > Add a Loop paragraph
    [Ctrl + Shift + I] > Mark message as important
    [Ctrl + K] > Insert link
    [Ctrl + Alt + Shift + C] > Insert embedded code
    [Ctrl + Alt + Shift + B] > Insert code block
    
    Mensagens:
    [Alt + Q] > Collapse all conversation folders
    [Ctrl + J] > Go to last read/new message
    [Ctrl + Alt + R] > React to last message
    [Alt + P] > Activate/deactivate details panel
    [Alt + Shift + R] > Reply to last message
    [Alt + 1-9] > Open 1st-9th Tab in Chat Panel Header
    [Ctrl + Alt + Z] > Clear all filters
    [Ctrl + Alt + U] > View all unread items
    [Ctrl + Alt + B] > View all meeting items
    [Ctrl + Alt + C] > View all chat conversations
    [Ctrl + Alt + A] > View all channel conversations
    [Ctrl + F] > Search current Chat/Channel messages
    [Alt + T] > Open Threads List
)"  ; end TeamsChat

; --- Spotify ---------------------------------------------------------------
cheatSheets["Spotify.exe"] := "
(
    Spotify (Shift)
    🔗 [C][C]onnect to device
    ⛶ [F][F]ullscreen
    🔍 [S][S]earch
    📋 [P][P]laylists
    🎤 [A][A]rtists
    💿 [B]Al[B]ums
    🏠 [H][H]ome
    🎵 [N][N]ow Playing
    🎯 [M][M]ade For You
    🆕 [R]New [R]eleases
    📊 [X]E[X]plore Charts
    🎵 [V][V]iew (Now Playing)
    📚 [L][L]ibrary sidebar
    ⛶ [E][E]xpand Library
    🎤 [Y]L[Y]rics
    ⏯️ [T][T]oggle Play/Pause
)"  ; end Spotify

; --- OneNote ---------------------------------------------------------------
cheatSheets["ONENOTE.EXE"] := "
(
    OneNote (Shift)
    📈 [Y]Expand [Y]section
    📉 [U]Collapse ([U]nfold reverse)
    📉 [I]Collapse All ([I]nward)
    📈 [O][O]pen All (Expand)
    📝 [P]Select [P]aragraph (line + children)
    🗑️ [D][D]elete line and children
    🗑️ [S][S]ingle delete (keep children)
    🔍 [F][F]ind Advanced (with quotes)
)"  ; end OneNote

; --- Chrome general shortcuts ----------------------------------------------
cheatSheets["chrome.exe"] := "
(
    Chrome (Shift)
    🪟 [W]Pop current tab to new [W]indow
    🏷️ [Ctrl+Alt+Y] [N]ame ChatGPT Window as "ChatGPT"
)"  ; end Chrome

; --- Google Maps (Chrome) ---------------------------------------------------
cheatSheets["Google Maps"] := "
(
    Google Maps (Shift)
    🔍 [S][S]earch box (place / query)
    📍 [L][L]at/long (copy coordinates to clipboard)
)"  ; end Google Maps

; --- Chrome PDF Viewer ------------------------------------------------------
cheatSheets["Chrome PDF Viewer"] := "
(
    Chrome PDF Viewer (Shift)
    ⬇️ [D] [D]ownload PDF
    📏 [F] [F]it to page (zoom to fit)
    🔢 [P] [P]age number field (focus)
    🗂️ [T] [T]humbnails sidebar (toggle)
    🔲 [2] Two-page view ([2] pages)
    🎬 [E] Present mode (pr[E]sent)
)"  ; end Chrome PDF Viewer

; --- Cursor ------------------------------------------------------
cheatSheets["Cursor.exe"] := "
(
    Cursor
    
    === Ctrl (no other modifiers) ===
    🎯 [1] Remove clustering and focus on the code (ahk)
    📁 [2] Copy path (cursor)
    📊 [3] CSV: Edit CSV
    💾 [4] CSV: Apply changes to source file and save
    📋 [5] MarkDown Enhanced: Export in PDF format. 
    📽️ [6] Marp export (PDF)
    🔨 [7] Build LaTeX project
    📄 [8] View LaTeX PDF file
    📄 [9] Markdown Preview Enhanced: Insert Page Break
    🤖 [M]Ask [M]essage, wait 6s, then paste (ahk)
    ⚡ [G]Kill terminal ([G]o away)
    📉 [Y]Fold all (tuck awa[Y])
    📈 [U] [U]nfold all
    📋 [O]Open Paste As... ([O]pen)
    📁 [H]Smart nav: Editor→Explorer / Explorer→Reveal (s[H]ow)
    🔲 [J]Select to Bracket (ad[J]acent)
    📉 [,] Fold all directories
    💬 [.] Toggle chat or agent
    🤖 [E] Maximize chat size — native Cursor (`workbench.action.maximizeChatSize`; user keybinding)
    📂 [R]File open [R]ecent
    🔍 [T]Go to [T]ype symbol in workspace
    💬 [N] [N]ew chat tab (replacing current)
    ➕ [Enter] [I]nsert line below
    🔍 [P]Open [P]roject
    💬 [;] Insert comment
    📝 [D]Duplicate selection to next find match
    🔍 [F] [F]ind
    ↩️ [Z]Undo (common [Z])
    📊 [B]Toggle [B]ar (primary sidebar)
    
    === Shift ===
    📉 [F][F]old (ahk)
    📈 [U][U]nfold (ahk)
    📄 [M][M]arkdown preview (cursor)
    🪟 [W][W]indow (move editor) (cursor)
    💻 [T][T]erminal (ahk)
    💻 [N][N]ew Terminal (ahk)
    📁 [E][E]xplorer (ahk)
    📄🪟 [K] Mar[K]down + window (ahk)
    ⌨️ [C][C]ommand palette (ahk)
    📈 [X] E[X]pand selection (ahk)
    ⚡ [S][S]ymbol in access view (cursor)
    💬 [H][H]istory (chat) (ahk)
    🖼️ [I][I]mage (paste) (cursor)
    📁 [G][G]it repos fold (SCM) (ahk)
    🔍 [Q][Q]uery Search (ahk)
    🍞 [R]B[R]eadcrumbs menu (ahk)
    😀 [O]Emoji selector (em[O]ji) (ahk)
    🌿 [D]Git section ([D]iff) (ahk)
    ❌ [Z]Close all editors (end [Z]one) (ahk)
    🤖 [A][A]I models switch (ahk)
    🧘 [Y]Zen mode (tranquilit[Y]) (cursor)
    ⬇️ [P][P]ull (Git) (cursor)
    ✅ [V]Commit (Git sa[V]e) (cursor)
    ⬆️ [B]Push (Git pu[B]lish) (cursor)
    
    === Alt (ahk = AutoHotkey) ===
    📉 [x] Shri[X]nk selection (ahk)
    📉 [,] Classical Markdown Preview
    📉 [Y] Paste image to Markdown
    ⬇️ [U] Scroll AI feed to bottom (ahk-based)
    📋 [M] Quick shortcut menu (ahk)
    🤖 [A] Add file to AI Context (Cursor Chat) (ahk)
    📌 [Q] Unpin current tab
    📌 [P] [P]in current tab
    📋 [I] Reveal in Explorer + copy file (ahk)
    📂 [H] Reveal in Explorer + open file (ahk)
    📄 [R] Refresh preview
    📄 [F] File: New [F]ile
    📂 [O] File: New F[O]lder
    
    === Ctrl+Shift ===
    📝 [Ctrl+Shift+L] Select all identical words ([L]ines)
    🐛 [Ctrl+Shift+D] [D]ebugging
    
    === Ctrl+Alt ===
    📄 [Ctrl+Alt+L] Markdown Preview Enhanced: Toggle Live Update
    📄 [Ctrl+Alt+T] Markdown Preview Enhanced: Toggle Scroll Sync
    ⬆️ [Ctrl+Alt+Up] Go to [P]arent Fold
    ⬅️ [Ctrl+Alt+Left] Go to sibling fold [P]revious
    ➡️ [Ctrl+Alt+Right] Go to sibling fold [N]ext
    ⬆️ [Ctrl+Alt+↑] Add cursor [A]bove
    ⬇️ [Ctrl+Alt+↓] Add cursor [B]elow
    
    === Alt+Shift ===
    ⬆️ [Shift+Alt+↑] [C]opy line Up
    ⬇️ [Shift+Alt+↓] [C]opy line Down
    
    === Alt (other chords) ===
    👁️ [Alt+F12] [P]eek Definition
    ⬆️ [Alt+↑] [M]ove line Up
    ⬇️ [Alt+↓] [M]ove line Down
    👆 [Alt+Click] [M]ulti-cursor by click
    🔄 [Alt+Z] Toggle word [W]rap
    ⬇️ [Alt+J] Jump to [N]ext review
    ⬆️ [Alt+K] [P]revious review (bac[K])
    
    === Function keys & misc ===
    ✏️ [F2] [R]ename symbol
    🔍 [F8] [N]avigate problems
    🗑️ [Shift+Delete] [D]elete line
)"  ; end Cursor

cheatSheets["Code.exe"] := "
(
    VS Code
    
    === Ctrl (no other modifiers) ===
    🎯 [1] Remove clustering and focus on the code (ahk)
    📁 [2] Copy path (VS Code)
    📊 [3] CSV: Edit CSV
    💾 [4] CSV: Apply changes to source file and save
    📋 [5] MarkDown Enhanced: Export in PDF format. 
    📽️ [6] Marp export (PDF)
    🔨 [7] Build LaTeX project
    📄 [8] View LaTeX PDF file
    📄 [9] Markdown Preview Enhanced: Insert Page Break
    🤖 [M]Ask [M]essage, wait 6s, then paste (ahk)
    ⚡ [G]Kill terminal ([G]o away)
    📉 [Y]Fold all (tuck awa[Y])
    📈 [U] [U]nfold all
    📋 [O]Open Paste As... ([O]pen)
    📁 [H]Smart nav: Editor→Explorer / Explorer→Reveal (s[H]ow)
    🔲 [J]Select to Bracket (ad[J]acent)
    📉 [,] Fold all directories
    💬 [.] Copilot Agent Modes
    🤖 [E] VS Code default behavior (Cursor custom maximize removed)
    📂 [R]File open [R]ecent
    🔍 [T]Go to [T]ype symbol in workspace
    💬 [N] Copilot chat session workflow (pending dedicated remap)
    ➕ [Enter] [I]nsert line below
    🔍 [P]VS Code quick open / project search
    💬 [;] Insert comment
    📝 [D]Duplicate selection to next find match
    🔍 [F] [F]ind
    ↩️ [Z]Undo (common [Z])
    📊 [B]Toggle [B]ar (primary sidebar)
    
    === Shift ===
    📉 [F][F]old (ahk)
    📈 [U][U]nfold (ahk)
    📄 [M][M]arkdown preview (VS Code migration pending)
    🪟 [W][W]indow (move editor) (VS Code migration pending)
    💻 [T][T]erminal (ahk)
    💻 [N][N]ew Terminal (ahk)
    📁 [E][E]xplorer (ahk)
    📄🪟 [K] Mar[K]down + window (ahk)
    ⌨️ [C][C]ommand palette (ahk)
    📈 [X] E[X]pand selection (ahk)
    ⚡ [S][S]ymbol in access view (VS Code migration pending)
    💬 [H][H]istory (chat) (ahk)
    🖼️ [I][I]mage (paste) (VS Code migration pending)
    📁 [G][G]it repos fold (SCM) (ahk)
    🔍 [Q][Q]uery Search (ahk)
    🍞 [R]B[R]eadcrumbs menu (ahk)
    😀 [O]Emoji selector (em[O]ji) (ahk)
    🌿 [D]Git section ([D]iff) (ahk)
    ❌ [Z]Close all editors (end [Z]one) (ahk)
    🤖 [A][A]I models switch (ahk)
    🧘 [Y]Zen mode (tranquilit[Y])
    ⬇️ [P][P]ull (Git)
    ✅ [V]Commit (Git)
    ⬆️ [B]Push (Git)
    
    === Alt (ahk = AutoHotkey) ===
    📉 [x] Shri[X]nk selection (ahk)
    📉 [,] Classical Markdown Preview
    📉 [Y] Paste image to Markdown
    ⬇️ [U] Scroll AI feed to bottom (ahk-based)
    📋 [M] Quick shortcut menu (ahk)
    ➕ [C] Add Context picker (VS Code chat) (ahk)
    🤖 [A] Add file to AI Context (VS Code chat) (ahk)
    📌 [Q] Unpin current tab
    📌 [P] [P]in current tab
    📋 [I] Reveal in Explorer + copy file (ahk)
    📂 [H] Reveal in Explorer + open file (ahk)
    📄 [R] Refresh preview
    📄 [F] File: New [F]ile
    📂 [O] File: New F[O]lder
    
    === Ctrl+Shift ===
    📝 [Ctrl+Shift+L] Select all identical words ([L]ines)
    🐛 [Ctrl+Shift+D] [D]ebugging
    
    === Ctrl+Alt ===
    📄 [Ctrl+Alt+L] Markdown Preview Enhanced: Toggle Live Update
    📄 [Ctrl+Alt+T] Markdown Preview Enhanced: Toggle Scroll Sync
    ⬆️ [Ctrl+Alt+Up] Go to [P]arent Fold
    ⬅️ [Ctrl+Alt+Left] Go to sibling fold [P]revious
    ➡️ [Ctrl+Alt+Right] Go to sibling fold [N]ext
    ⬆️ [Ctrl+Alt+↑] Add cursor [A]bove
    ⬇️ [Ctrl+Alt+↓] Add cursor [B]elow
    
    === Alt+Shift ===
    ⬆️ [Shift+Alt+↑] [C]opy line Up
    ⬇️ [Shift+Alt+↓] [C]opy line Down
    
    === Alt (other chords) ===
    👁️ [Alt+F12] [P]eek Definition
    ⬆️ [Alt+↑] [M]ove line Up
    ⬇️ [Alt+↓] [M]ove line Down
    👆 [Alt+Click] [M]ulti-cursor by click
    🔄 [Alt+Z] Toggle word [W]rap
    ⬇️ [Alt+J] Jump to [N]ext review
    ⬆️ [Alt+K] [P]revious review (bac[K])
    
    === Function keys & misc ===
    ✏️ [F2] [R]ename symbol
    🔍 [F8] [N]avigate problems
    🗑️ [Shift+Delete] [D]elete line
)"  ; end VS Code

; --- Windows Explorer ------------------------------------------------------
cheatSheets["explorer.exe"] := "
(
    Explorer (Shift)
    📄 [F]Select first [F]ile
    🔍 [S][S]earch bar
    📍 [A][A]ddress bar
    📁 [N][N]ew Folder²
    🔗 [H]Create s[H]ortcut
    📋 [C][C]opy as path
    📤 [R]Sha[R]e file
    📌 [P][P]inned item (first in sidebar)
    📌 [L][L]ast item (sidebar)
    📦 [X] WinRAR e[X]tract here (personal); work: 7-Zip extract
    📦 [W] WinRAR add to archive / compact (personal); work: 7-Zip add to archive / compress
)"  ; end Explorer

; --- Microsoft Paint ------------------------------------------------------
cheatSheets["mspaint.exe"] := "
(
    MS Paint (Shift)
    📏 [R][R]esize and Skew (Ctrl+W)
    
    --- Common Shortcuts ---
    [Ctrl+N] > 📄 New
    [Ctrl+O] > 📂 Open
    [Ctrl+S] > 💾 Save
    [F12] > 💾 Save As
    [Ctrl+P] > 🖨️ Print
    [Ctrl+Z] > ↩️ Undo
    [Ctrl+Y] > ↪️ Redo
    [Ctrl+A] > 📄 Select all
    [Ctrl+C] > 📋 Copy
    [Ctrl+X] > ✂️ Cut
    [Ctrl+V] > 📋 Paste
    [Ctrl+W] > 📏 Resize and Skew
    [Ctrl+E] > ℹ️ Image properties
    [Ctrl+R] > 📏 Toggle rulers
    [Ctrl+G] > 🔲 Toggle gridlines
    [Ctrl+I] > 🔄 Invert colors
    [F11] > 🖥️ Fullscreen view
    [Ctrl++] > 🔍 Zoom in
    [Ctrl+-] > 🔍 Zoom outd
)"  ; end MS Paint

; --- ClipAngel -------------------------------------------------------------
cheatSheets["ClipAngel.exe"] := "
(
    ClipAngel (Shift)
    📋 [C][C]opy filtered content
    🔄 [T][T]oggle focus list/text
    🗑️ [D][D]elete all non-favorite
    🧹 [X]E[X]it filters (Clear)
    ⭐ [F]Mark as [F]avorite
    ⭐ [U][U]nmark as favorite
    ✏️ [E][E]dit Text (F4)
    💾 [S][S]ave as file
    🔗 [M][M]erge clips
    🔍 [Y]File t[Y]pe filter (Quick Wizard)
    ⌨️ [Alt+1] [S]elect current item
    ⌨️ [Alt+2] [M]ove down once and select
    ⌨️ [Alt+3] [M]ove down twice and select
    ⌨️ [Alt+4] [M]ove down three times and select
    ⌨️ [Alt+5] [M]ove down four times and select
    📋 [Ctrl+1] Tab, Select All, Copy
    📋 [Ctrl+2] Down 1, Tab, Select All, Copy
    📋 [Ctrl+3] Down 2, Tab, Select All, Copy
    📋 [Ctrl+4] Down 3, Tab, Select All, Copy
    📋 [Ctrl+5] Down 4, Tab, Select All, Copy
)"  ; end ClipAngel

; --- Figma -----------------------------------------------------------------
cheatSheets["Figma.exe"] := "
(
    Figma (Shift)
    👁️ [U]Toggle [U]I visibility
    🔍 [S][S]earch component
    ⬆️ [P]Select [P]arent
    🧩 [C]reate [C]omponent
    🔗 [D][D]etach instance
    📐 [A]dd [A]uto layout
    📐 [R][R]emove auto layout
    💡 [S][S]uggest auto layout
    📤 [E][E]xport
    🖼️ [C][C]opy as PNG
    ⚡ [A][A]ctions...
    ⬅️ [L]Align [L]eft
    ➡️ [R]Align [R]ight
    📏 [V]Distribute [V]ertical spacing
    🧹 [T][T]idy up
    ⬆️ [T]Align [T]op
    ⬇️ [B]Align [B]ottom
    ↔️ [H]Align center [H]orizontal
    ↕️ [V]Align center [V]ertical
    📏 [H]Distribute [H]orizontal spacing
)"  ; end Figma

; --- Gmail ---------------------------------------------------------------
cheatSheets["Gmail"] := "
(
    Gmail (Shift)
    📥 [I][I]nbox
    📰 [U][U]pdates
    💬 [F][F]orums
    📬 [R]Toggle [R]ead status
    ⬅️ [P][P]revious conversation
    ➡️ [N][N]ext conversation
    📦 [A][A]rchive conversation
    ✅ [S][S]elect conversation
    ↩️ [Y]Repl[Y]
    ↩️ [G]Reply to [G]roup (all)
    ➡️ [W]For[W]ard
    ⭐ [T]S[T]ar toggle
    🗑️ [D][D]elete
    🚫 [X]Spam (e[X]clude)
    ✍️ [C][C]ompose new email
    🔍 [Q][Q]uery mail (Search)
    📁 [M][M]ove to folder
    ⌨️ [H][H]elp (keyboard shortcuts)
    📬 [B]Inbox [B]utton
    
    --- Built-in Shortcuts (Windows) ---
    
    Compose & chat:
    [p] > Previous message in an open conversation
    [n] > Next message in an open conversation
    [Shift + Esc] > Focus main window
    [Esc] > Focus latest chat or compose
    [Ctrl + .] > Advance to the next chat or compose
    [Ctrl + ,] > Advance to previous chat or compose
    [Ctrl + Enter] > Send
    [Ctrl + Shift + c] > Add cc recipients
    [Ctrl + Shift + b] > Add bcc recipients
    [Ctrl + Shift + f] > Access custom from
    [Ctrl + k] > Insert a link
    [Ctrl + m] > Open spelling suggestions
    
    Formatting text:
    [Ctrl + Shift + 5] > Previous font
    [Ctrl + Shift + 6] > Next font
    [Ctrl + Shift + -] > Decrease text size
    [Ctrl + Shift + +] > Increase text size
    [Ctrl + b] > Bold
    [Ctrl + i] > Italics
    [Ctrl + u] > Underline
    [Ctrl + Shift + 7] > Numbered list
    [Ctrl + Shift + 8] > Bulleted list
    [Ctrl + Shift + 9] > Quote
    [Ctrl + []] > Indent less
    [Ctrl + ]] > Indent more
    [Ctrl + Shift + l] > Align left
    [Ctrl + Shift + e] > Align center
    [Ctrl + Shift + r] > Align right
    [Ctrl + \] > Remove formatting
    
    Actions (shortcuts on):
    [,] > Move focus to toolbar
    [x] > Select conversation
    [s] > Toggle star/rotate among superstars
    [e] > Archive
    [m] > Mute conversation
    [!] > Report as spam
    [#] > Delete
    [r] > Reply
    [Shift + r] > Reply in a new window
    [a] > Reply all
    [Shift + a] > Reply all in a new window
    [f] > Forward
    [Shift + f] > Forward in a new window
    [Shift + n] > Update conversation
    [] or []] > Archive conversation and go previous/next
    [z] > Undo last action
    [Shift + i] > Mark as read
    [Shift + u] > Mark as unread
    [_] > Mark unread from the selected message
    [+ or =] > Mark as important
    [-] > Mark as not important
    [b] > Snooze (not available in classic Gmail)
    [;] > Expand entire conversation
    [:] > Collapse entire conversation
    [Shift + t] > Add conversation to Tasks
    
    Jumping (shortcuts on):
    [g + i] > Go to Inbox
    [g + s] > Go to Starred conversations
    [g + b] > Go to Snoozed conversations
    [g + t] > Go to Sent messages
    [g + d] > Go to Drafts
    [g + a] > Go to All mail
    [Ctrl + Alt + ,] > Switch to left sidebar (Calendar/Keep/Tasks)
    [Ctrl + Alt + .] > Switch to right (back to inbox)
    [g + k] > Go to Tasks
    [g + l] > Go to label
    
    Threadlist selection (shortcuts on):
    [* + a] > Select all conversations
    [* + n] > Deselect all conversations
    [* + r] > Select read conversations
    [* + u] > Select unread conversations
    [* + s] > Select starred conversations
    [* + t] > Select unstarred conversations
    
    Navigation (shortcuts on):
    [g + n] > Go to next page
    [g + p] > Go to previous page
    [u] > Back to threadlist
    [k] > Newer conversation
    [j] > Older conversation
    [o or Enter] > Open conversation
    [`] > Go to next Inbox section
    [~] > Go to previous Inbox section
    
    Application (shortcuts on):
    [c] > Compose
    [d] > Compose in a new tab
    [/] > Search mail
    [q] > Search chat contacts
    [.] > Open ""more actions"" menu
    [v] > Open ""move to"" menu
    [l] > Open ""label as"" menu
    [?] > Open keyboard shortcut help
)"  ; end Gmail

; --- Google Keep ---------------------------------------------------------------
cheatSheets["Google Keep"] := "
(
    Google Keep (Shift)
    🔍 [S][S]earch and select Note
    📋 [M]Toggle [M]ain menu
)"  ; end Google Keep

; --- File Dialog ---------------------------------------------------------------
cheatSheets["FileDialog"] := "
(
    File Dialog (Shift)
    📄 [F]Select first [F]ile
    🔍 [S][S]earch bar
    📍 [A][A]ddress bar
    📁 [N][N]ew Folder
    📌 [P][P]inned item (first in sidebar)
    💻 [T][T]his PC (sidebar)
    📝 [M]File na[M]e field
    ✅ [O][O]pen/Save button
    ❌ [C][C]ancel button
)"

; --- Settings Window -------------------------------------------------
cheatSheets["Settings"] := "(Settings (Shift))`r`n🔊 [V]Set input [V]olume to 100%"

; --- Command Palette -------------------------------------------------
cheatSheets["Command Palette"] := "
(
    Command Palette (Shift)
    ⌨️ [Ctrl+H] Reveal in file explorer
    ⌨️ [C][C]opy file Path
    ⌨️ [B]Go [H]ome
    ⌨️ [S]Precise [S]earch
    ⌨️ [I][I]nsert Favorite (Add)
    ⌨️ [D][E]xclude Favorite
    ⌨️ [Ctrl+1] [S]elect current item
    ⌨️ [Ctrl+2] [M]ove down once and select
    ⌨️ [Ctrl+3] [M]ove down twice and select
    ⌨️ [Ctrl+4] [M]ove down three times and select
    ⌨️ [Ctrl+5] [M]ove down four times and select
    ⌨️ [Ctrl+6] [M]ove down five times and select
    ⌨️ [Alt+1] [S]elect current item
    ⌨️ [Alt+2] [M]ove down once and select
    ⌨️ [Alt+3] [M]ove down twice and select
    ⌨️ [Alt+4] [M]ove down three times and select
    ⌨️ [Alt+5] [M]ove down four times and select
)"

; --- Excel ------------------------------------------------------------
cheatSheets["EXCEL.EXE"] := "
(
    Excel (Shift)
    ⚪ [W]Select [W]hite Color
    ✏️ [E]Enable [E]diting
    📊 [C][C]SV to columns (semicolon delimited)
    📋 [V]Quickly [V]aste and extract CSV
    ➕ [A][A]dd multiple rows (10 rows)
    🗑️ [R][R]ow removal workflow (remove row, down arrow, repeat 5-7 times)
    📅 [P]Type [P]revious day date
)"

; --- Power BI ------------------------------------------------------------
cheatSheets["Power BI"] := "
(
    Power BI (Shift)
    📊 [I]Report v[I]ew
    📊 [O]Table view ([O]verview)
    📋 [Z]Copy cell Val
    📊 [P]Model view ([P]lan)
    📊 [C]Get data ([C]onnect)
    📊 [T][T]ransform Data
    📊 [U][U]pdate (Close and Apply)
    📊 [E]New M[E]asure
    🔄 [Y]Refresh (read[Y])
    📊 [H]Build visual ([H]andle)
    📊 [J]Format visual (ad[J]ust)
    ⬆️ [B][B]ring forward
    ⬇️ [D]Sen[D] backward
    📐 [K]Keep [A]lign straight
    📄 [V]Fit to Page ([V]iew)
    🎨 [M]Format painter ([M]atch format)
    🔗 [N]Group visuals (Groupi[N]g)
    🖱️ [A][A]ll pages button
    ➕ [W]Ne[W] Page
    📕 [F]Close All Drawers ([F]old)
    📖 [G]Open All Drawers (un[G]roup)
    📁 [R]Collapse Fields ([R]educe)
    🔍 [S][S]earch edit field
    ✅ [L]Confirm moda[L] button (OK)
    ❌ [X]Cancel/E[X]it modal button
)"

; --- UIA Tree Inspector -------------------------------------------------
cheatSheets["UIATreeInspector"] :=
"(UIA Tree Inspector (Shift))`r`n🔄 [R][R]efresh List`r`n🔍 [F]ocus [F]ilter field`r`n🔍 [S]elect [S]earch tree item`r`n📋 [C]Copy full UI tree"
; --- SettleUp Shortcuts -----------------------------------------------------
cheatSheets["Settle Up"] := "
(
    Settle Up (Shift)
    ➕ [A][A]dd Transaction
    📝 [N]Focus expense [N]ame field
    💰 [V]Focus expense [V]alue field
)"

; --- Miro Shortcuts -----------------------------------------------------
cheatSheets["Miro"] := "
(
    Miro (Shift)
    📋 [F][F]rame List
    🔗 [G][G]roup
    🔗 [U][U]ngroup
    🔒 [L][L]ock/Unlock
    🔗 [K]Add/Edit Lin[K]
    ❌ [X]Close sidebar (e[X]it)
    --- Built-in Shortcuts (Windows) ---
    Tools:
    [V / H] > Select tool / Hand
    [T] > Text
    [N] > Sticky notes
    [S] > Shapes
    [R] > Rectangle
    [O] > Oval
    [L] > Connection line / Arrow
    [D] > Card
    [P] > Pen
    [E] > Eraser
    [C] > Comment
    [F] > Frame
    [M] > Minimap
    [Ctrl + K] > Command palette
    [Enter (bulk)] > New sticky note
    [Esc (bulk)] > Exit sticky note bulk mode
    [Ctrl + Shift + Enter] > Open card panel
    [Shift + C] > Show/hide comments
    
    General:
    [Ctrl + C / Ctrl + V] > Copy / Paste
    [Ctrl + X] > Cut
    [Ctrl + D] > Duplicate
    [Alt + drag] > Duplicate by drag
    Alt + â†â†’â†‘â†“        â†’  Duplicate horizontally/vertically
    [Ctrl + click] > Select multiple
    [Ctrl + A] > Select all
    [Enter] > Edit selected
    [Esc] > Deselect / quit edit
    [Backspace] > Delete
    [Ctrl + G] > Group
    [Ctrl + Shift + G] > Ungroup
    [Ctrl + Shift + L] > Lock / Unlock
    [Ctrl + Shift + P] > Protected lock / Unprotected lock
    [PgUp] > Bring to front
    [Shift + PgUp] > Bring forward
    [PgDn] > Send to back
    [Shift + PgDn] > Send backward
    [Ctrl + Shift + K] > Create board in new tab
    [Alt + Ctrl + K] > Add/Edit link to object
    [Ctrl + Backspace] > Clear object contents
    
    Navigation:
    â†â†'â†'              â†'  Move items/canvas
    [Ctrl + +] > Zoom in
    [Ctrl + -] > Zoom out
    [Ctrl + 0] > Zoom to 100%
    [Alt + 1] > Zoom to fit
    [Alt + 2] > Zoom to selected item
    [Space + drag] > Move canvas
    [G] > Toggle grid
    [Ctrl + F] > Search
    
    Text: 
    [Ctrl + B] > Bold
    [Ctrl + I] > Italic
    [Ctrl + U] > Underline
    
    Board navigation:
    [Tab] > Move forwards through objects (TL > BR)
    [Shift + Tab] > Move backwards through objects (TL > BR)
    Ctrl + â†'/+â†"/â†/â†'    â†'  Move through board objects
    [Ctrl + Shift + ↓/↑] > Move in/out of container (e.g., frame)
    [Esc] > Back to menu
    [Enter] > Edit an object
    [Esc] > Stop editing an object
    
    Toolbar navigation:
    [Tab / Shift + Tab] > Move between toolbars
    [Arrow keys] > Move between toolbar items
    [Enter / Space] > Activate a menu item
    
    Desktop app:
    [Ctrl + R] > Reload the tab
    [Ctrl + W] > Close the tab
    [Ctrl + Q] > Exit the app
    [Ctrl + Shift + L] > Copy board link
)"

; --- Wikipedia ---------------------------------------------------------------
cheatSheets["Wikipedia"] := "
(
    Wikipedia (Shift)
    🔍 [S][S]earch button click
    💾 [P]Save scroll [P]osition
)"

; --- YouTube ---------------------------------------------------------------
cheatSheets["YouTube"] := "
(
    YouTube (Shift)
    🔍 [S]Focus [S]earch box
    🎬 [U]Focus first video (filter res[U]lts)
    🎬 [I]Focus first v[I]deo via Explore
    🏠 [H]Navigate to [H]ome
    📜 [R]Navigate to histo[R]y
    📋 [P]Navigate to [P]laylists
)"

; --- Google Search ---------------------------------------------------------------
cheatSheets["Google"] := "
(
    Google (Shift)
    🔍 [S][S]earch box focus
    🥇 [U][U]se first result
)" 44

; --- ChatGPT ---------------------------------------------------------------
cheatSheets["ChatGPT"] := "
(
    ChatGPT (Shift)
    📂 [I]Toggle s[I]debar
    🔄 [O]Re-send rules ([O]rder again)
    📋 [C][C]opy code block
    ⬇️ [J]Go down ([J]ump)
    🤖 [L]Send and show AI P[L]anner
)"

; --- Gemini (web, Chrome) -----------------------------------------------
cheatSheets["Gemini"] := "
(
    Gemini (Shift)
    📂 [D]Toggle the[D]rawer
    💬 [N][N]ew chat
    🔍 [S][S]earch
    🔄 [M]Change[M]odel
    🛠️ [T][T]ools
    🖼️ [I]Create [I]mage (Tools menu; opens if needed)
    🔬 [E]Deep r[E]search (Tools menu; opens if needed)
    ⌨️ [P]Focus[P]rompt field
    📋 [C][C]opy last message
    🔊 [R][R]ead aloud last message
    🤖 [G]Send[G]emini prompt text
    ⛶ [F][F]ullscreen input
    🔔 [Enter / Ctrl+Enter]Send and notify on completion
    
    === Alt (ahk) ===
    ⬇️ [U] Scroll AI feed to bottom — same idea as Cursor
)"

; --- M365 Copilot web (Chrome) — same Shift keys as Gemini -----------------
cheatSheets["Copilot Web"] := "
(
    Copilot Web (Shift)
    📂 [D]Toggle nav [D]rawer
    💬 [N][N]ew chat
    🔍 [S][S]earch (nav drawer)
    🔄 [M]Change [M]odel (opens selector)
    🛠️ [T]Add/manage sources (Tools menu)
    🖼️ [I]Designer / create image (Sources menu)
    🔬 [E]Researcher / deep research (Sources menu)
    ⌨️ [P]Focus [P]rompt field
    📋 [C][C]opy last response
    🔊 [R][R]ead aloud last message
    🤖 [G]Send prompt text (Gemini_Prompt.txt)
    ⛶ [F][F]ullscreen input (expand composer)
    🔔 [Enter / Ctrl+Enter]Send and notify on completion
    
    === Alt (ahk) ===
    ⬇️ [U] Scroll AI feed to bottom — same idea as Cursor
)"

; --- Mobills ---------------------------------------------------------------
cheatSheets["Mobills"] := "
(
    Mobills (Shift)
    
    --- Navigation ---
    📊 [D][D]ashboard
    💳 [A][A]ccounts (Contas)
    💰 [T][T]ransactions (Transações)
    💳 [C]redit [C]ards (Cartões)
    📅 [P][P]lanning (Planejamento)
    📈 [R][R]eports (Relatórios)
    ⚙️ [M]ore [M]enu (Mais opções)
    ⬅️ [K]Previous month (bac[K])
    ➡️ [L]Next month (cyc[L]e)
    
    --- Actions ---
    🚫 [I][I]gnore transaction
    ✏️ [N][N]ame Field
    💸 [E]New [E]xpense
    💵 [Y]New Incom[Y]
    💳 [X]Credit card e[X]pense
    🔄 [F]Funds trans[F]er
    🔘 [W][W]indow (Open button + type MAIN)
)"

; Mirrors the former sequential if-chain: later assignments override earlier ones; Shopee/Google only when still unset.
PickChromeAppSheetKey(chromeTitle) {
    key := ""
    if IsChromePdfViewerActive()
        key := "Chrome PDF Viewer"
    if InStr(chromeTitle, "WhatsApp")
        key := "WhatsApp"
    if InStr(chromeTitle, "Gmail")
        key := "Gmail"
    if InStr(chromeTitle, "chatgpt")
        key := "ChatGPT"
    if InStr(chromeTitle, "Mobills")
        key := "Mobills"
    if InStr(chromeTitle, "Google Keep") || InStr(chromeTitle, "keep.google.com")
        key := "Google Keep"
    if InStr(chromeTitle, "YouTube")
        key := "YouTube"
    if InStr(chromeTitle, "UIATreeInspector")
        key := "UIATreeInspector"
    if InStr(chromeTitle, "Settle Up")
        key := "Settle Up"
    if InStr(chromeTitle, "Miro")
        key := "Miro"
    if InStr(chromeTitle, "Wikipedia", false) || InStr(chromeTitle, "wikipedia.org", false)
        key := "Wikipedia"
    if IsMercadoLivreActive()
        key := "Mercado Livre"
    if (key = "" && IsShopeeActive())
        key := "Shopee"
    if InStr(chromeTitle, "gemini", false)
        key := "Gemini"
    if (key = "" && (InStr(chromeTitle, "M365 Copilot", false) || InStr(chromeTitle, "Chat | M365 Copilot", false)))
        key := "Copilot Web"
    if InStr(chromeTitle, "Google Maps")
        key := "Google Maps"
    if (key = "" && (chromeTitle = "Google" || InStr(chromeTitle, " - Google Search")))
        key := "Google"
    return key
}

; ========== Helper to decide which sheet applies ===========================
GetCheatSheetText() {
    global cheatSheets

    exe := WinGetProcessName("A") ; active process name (e.g. chrome.exe)
    title := WinGetTitle("A")
    hwnd := WinExist("A")

    ; Microsoft Store "new Outlook" runs as olk.exe; use same cheat sheets as OUTLOOK.EXE.
    if (StrLower(exe) = "olk.exe")
        exe := "OUTLOOK.EXE"

    ; (removed temporary tooltip debugging)

    ; Prefer Outlook Reminders over generic File Dialog detection
    if (exe = "OUTLOOK.EXE") {
        if RegExMatch(title, "i)Reminder")
            return cheatSheets.Has("OutlookReminder") ? cheatSheets["OutlookReminder"] : ""
    }

    ; Check for file dialog first (works in any app)
    if WinGetClass("ahk_id " hwnd) = "#32770" {
        txt := WinGetText("ahk_id " hwnd)
        if InStr(txt, "Namespace Tree Control") || InStr(txt, "Controle da Ãrvore de Namespace") {
            return cheatSheets["FileDialog"]
        }
    }

    ; Check for Settings window (both English and Portuguese)
    if (title = "Settings" || title = "ConfiguraÃ§Ãµes") {
        return cheatSheets.Has("Settings") ? cheatSheets["Settings"] : ""
    }

    ; Check for Command Palette window
    if InStr(title, "Command Palette", false) {
        return cheatSheets.Has("Command Palette") ? cheatSheets["Command Palette"] : ""
    }

    ; Check for Power BI (by process name or window title)
    if (exe = "PBIDesktop.exe" || InStr(title, "powerbi", false)) {
        return cheatSheets.Has("Power BI") ? cheatSheets["Power BI"] : ""
    }

    ; Special handling for Chrome-based apps that share chrome.exe
    if (exe = "chrome.exe") {
        chromeShortcuts := cheatSheets.Has("chrome.exe") ? cheatSheets["chrome.exe"] : ""
        appShortcuts := ""

        ; Normalize Chrome window title by removing the trailing " - Google Chrome"
        chromeTitle := RegExReplace(title, "i) - Google Chrome$", "")

        siteKey := ""
        if (hwnd) {
            try {
                if (CopilotWeb_IsCopilotHwnd(hwnd, "fast") || CopilotWeb_IsCopilotHwnd(hwnd, "full")
                || CopilotWeb_TryUiaFingerprint(hwnd))
                    siteKey := "Copilot Web"
            } catch {
            }
        }
        if (siteKey = "")
            siteKey := PickChromeAppSheetKey(chromeTitle)
        if (siteKey != "" && cheatSheets.Has(siteKey))
            appShortcuts := cheatSheets[siteKey]

        ; Combine Chrome general + app-specific shortcuts
        if (appShortcuts != "" && chromeShortcuts != "")
            return chromeShortcuts "`r`n`r`n" appShortcuts
        if (appShortcuts != "")
            return appShortcuts
        if (chromeShortcuts != "")
            return chromeShortcuts
        return ""
    }

    ; UIA Tree Inspector - check both process and window title
    if (exe = "AutoHotkey64.exe" && InStr(title, "UIATreeInspector"))
        return cheatSheets["UIATreeInspector"]

    ; Microsoft Teams â€" differentiate meeting vs chat via helper predicates
    if IsTeamsMeetingActive()
        return cheatSheets.Has("TeamsMeeting") ? cheatSheets["TeamsMeeting"] : ""
    if IsTeamsChatActive()
        return cheatSheets.Has("TeamsChat") ? cheatSheets["TeamsChat"] : ""
    if IsFileDialogActive()
        return cheatSheets["FileDialog"]

    ; Special handling for Outlook-based apps
    if (exe = "OUTLOOK.EXE") {
        ; Detect Reminders window â€" e.g. "3 Reminder(s)" or any title containing "Reminder"
        if RegExMatch(title, "i)Reminder") {
            return cheatSheets.Has("OutlookReminder") ? cheatSheets["OutlookReminder"] : cheatSheets["OUTLOOK.EXE"]
        }
        ; Detect Message inspector windows â€" e.g., " - Message (HTML)"
        if RegExMatch(title, "i) - Message \(") {
            return cheatSheets.Has("OutlookMessage") ? cheatSheets["OutlookMessage"] : cheatSheets["OUTLOOK.EXE"]
        }
        ; Detect Appointment, Meeting, or Event inspector windows
        if RegExMatch(title, "i)(Appointment|Meeting|Event)") {
            return cheatSheets.Has("OutlookAppointment") ? cheatSheets["OutlookAppointment"] : cheatSheets[
                "OUTLOOK.EXE"]
        }
        ; Fallback to generic Outlook sheet
        if cheatSheets.Has("OUTLOOK.EXE")
            return cheatSheets["OUTLOOK.EXE"]
    }

    ; Direct match by process name (generic fallback)
    if cheatSheets.Has(exe)
        return cheatSheets[exe]

    ; Try case-insensitive match for process name
    for key, value in cheatSheets {
        if (StrLower(key) = StrLower(exe))
            return value
    }

    ; Nothing found > blank > caller will show fallback message
    return ""
}

; ========== Shared variables for cheat sheet state ========================
global g_helpGui := 0
global g_helpShown := false
global g_globalGui := 0
global g_globalShown := false
global g_helpSearchEdit := 0
global g_helpCheatCtrl := 0
global g_globalSearchEdit := 0
global g_globalCheatCtrl := 0
global g_cheatSheetAppFullProcessed := ""
global g_cheatSheetGlobalFullProcessed := ""
global g_searchAllGui := 0

GetGlobalCheatSheetRawText() {
    global GLOBAL_CHEAT_SHEET_RAW
    provider := GetGlobalAIProviderLabel()
    return StrReplace(GLOBAL_CHEAT_SHEET_RAW, "{AI_PROVIDER}", provider)
}

; Raw text for long-hold global cheat sheet (also used by SearchCheatSheets).
GLOBAL_CHEAT_SHEET_RAW := "
(
    [Win+Alt+Shift] - PRIMARY triple modifier (most used for system-wide shortcuts)
        [Ctrl+Alt+Win] - SECONDARY triple modifier
    
    === AVAILABLE SECONDARY (Ctrl+Alt+Win) SLOTS ===
    [Ctrl+Alt+Win+N] > TEMPORARY — M365 Copilot auto-continue: send "continue", wait for Stop generating, loop (toggle off with same chord)
    [Ctrl+Alt+Win+O] > Evidence search loop — CSV row substring → PDF find; stop saves not-found rows to data/evidence_not_found.csv + 10s report (VSCodeEvidenceSearch.ahk; toggle)
    Letters available: T, U

    Shift+CAW: A/S/D/F/Q/W/E/R (+B debug, +1/Z/G fallbacks) used for window management; other Shift+letters unassigned.
    [Ctrl+Alt+Win+1] > Available
    [Ctrl+Alt+Win+G] > RESERVED — Handy: cancel dictation (define in Handy only; not bound in AHK)
    [Ctrl+Alt+Win+L] > {AI_PROVIDER} D2C direct submit (Utils.ahk; ZMK hold on L key)
    [Ctrl+Alt+Win+V] > Maximize active window (WindowManagement.ahk; ZMK hold on minimize/close key)
    [Ctrl+Alt+Win+X] > Snap 50/50: half-width active window + pair recent window in other half (WindowManagement.ahk)
    [Ctrl+Alt+Win+Z] > Window tools [1]: maximize lone visible window per monitor (WindowManagement.ahk; also Win+Alt+Shift+W → 1)
    [Ctrl+Alt+Win+6] > Window tools [2]: hidden background window list (WindowManagement.ahk; also Win+Alt+Shift+W → 2)
    [Ctrl+Alt+Win+Y] > Window tools [3]: tile background windows (WindowManagement.ahk; also Win+Alt+Shift+W → 3)
    [Ctrl+Alt+Win+P] > Window tools [4]: exit F11 fullscreen (WindowManagement.ahk; also Win+Alt+Shift+W → 4)
    [Ctrl+Alt+Win+0] > Project Quick Selector (opens project folder in Cursor)
    [Ctrl+Alt+Win+1] > Cursor AI quick action (Project Selector + Selection Mode)
    [Ctrl+Alt+Win+2] > Quick Update to Your Scripts (HotStrings macro)
    [Ctrl+Alt+Win+3] > Toggle Outlook and Teams (HotStrings macro)
    [Ctrl+Alt+Win+5] > Clean the Clipboard (HotStrings macro)
    [Ctrl+Alt+Win+7] > Mark Last Clip as Favorite (HotStrings macro; same as Ctrl+Alt+Win+J if 7 chord fails on keyboard)
    [Ctrl+Alt+Win+J] > Mark Last Clip as Favorite (HotStrings macro; alternate for keyboards that ghost Ctrl+Alt+Win+7)
    [Ctrl+Alt+Win+8] > Moves Desktop to Recycle Bin (HotStrings macro)
    [Ctrl+Alt+Win+9] > Handy: Cohere Portuguese (model slot 4)
    [Ctrl+Alt+Win+B] > Handy: Cohere English (model slot 3)
    
    === MAIN KEY COMBINATIONS ===
    [Symbol Layer] Win+Alt+Shift - Primary combination
    [Window Management] Ctrl+Alt+Win - Secondary combination
    
    [Alt+P] Ope clip angel
    
    [Win+Alt+Shift+L] > Outlook Copilot shortcut modal (1–9); global hotkey
    
    === CURSOR ===
    [Win+Alt+Shift+N] > Context file browser (paste path)
    
    [Win+Alt+Shift+J] > Fast Copy: tap on/off (count Ctrl+C / PrtSc / Alt+PrtSc, then paste N); hold 700ms+ repeats last N (Clip Angel)
    
    === SPOTIFY ===
    [Win+Alt+Shift+S] > Opens or activates Spotify
    
    === CLIP ANGEL ===
    [Win+Alt+Shift+1] > Send top list item from Clip Angel
    
    === AI CHAT (Chrome) ===
    [Win+Alt+Shift+I] > Opens {AI_PROVIDER}
    [Win+Alt+Shift+8] > Get word pronunciation, definition, and Portuguese translation ({AI_PROVIDER})
    [Win+Alt+Shift+O] > Read aloud the last message in {AI_PROVIDER}
    [Win+Alt+Shift+P] > Copy the last message in {AI_PROVIDER}
    [Win+Alt+Shift+7] > Copy selected text and read aloud ({AI_PROVIDER})
    
    === HANDY DICTATION ===
    [Win+Alt+Shift+0] > Start/stop dictation (transcription to clipboard)
    [Ctrl+Alt+Win+G] > Cancel dictation (Handy — user-defined; reserved in cheat sheet, not in AHK)
    [Ctrl+Alt+Win+9] > Handy: Cohere Portuguese (picker slot 4; same as Win+Alt+Shift+C then 4)
    [Ctrl+Alt+Win+B] > Handy: Cohere English (picker slot 3; same as Win+Alt+Shift+C then 3)
    [Win+Alt+Shift+C] > AI model picker (Handy): 1 Parakeet V2, 2 Parakeet V3, 3 Cohere English, 4 Cohere Portuguese
    
    === YOUTUBE ===
    [Win+Alt+Shift+H] > Activates Youtube
    
    === GOOGLE ===
    [Win+Alt+Shift+F] > Opens Google
    
    === CURSOR ===
    [Win+Alt+Shift+,] > Opens or activates Cursor
    [Win+Alt+Shift+C] > Handy AI model picker (see HANDY DICTATION)
    
    === OUTLOOK ===
    [Win+Alt+Shift+B] > Open mail
    [Win+Alt+Shift+V] > Open Reminder
    [Win+Alt+Shift+G] > Open calendar
    [Win+Alt+Shift+D] > Voice aloud the email
    
    === MICROSOFT TEAMS ===
    [Win+Alt+Shift+R] > New conversation
    [Win+Alt+Shift+5] > Toggle Mute (meeting)
    [Win+Alt+Shift+4] > Toggle camera (meeting)
    [Win+Alt+Shift+T] > Screen share (meeting)
    [Win+Alt+Shift+2] > Exit meeting
    [Win+Alt+Shift+E] > Select the chats window
    [Win+Alt+Shift+3] > Select the meeting window
    
    === WHATSAPP ===
    [Win+Alt+Shift+Z] > Opens WhatsApp
    
    === WINDOWS ===
    [Win+Alt+Shift+6] > Minimizes windows
    [Win+Alt+Shift+M] > Maximizes the current window
    [Win+Alt+Shift+W] > Window tools menu: [1] maximize lone; [2] hidden background list; [3] tile background (≤12 total, ≤3/monitor); [4] exit F11 fullscreen — direct CAW: Z=[1], 6=[2], Y=[3], P=[4]
    [Win+Alt+Shift+Y] > Focus Mode: Black out all monitors except the one with the active window (toggle)
    
    === WINDOW MANAGEMENT (Ctrl+Alt+Win) ===
    [Ctrl+Alt+Win+X] > Snap 50/50: half-width active window + pair recent window in other half
    [Ctrl+Alt+Win+Z] > Window tools [1]: maximize lone visible window per monitor (also Win+Alt+Shift+W → 1)
    [Ctrl+Alt+Win+6] > Window tools [2]: hidden background window list (also Win+Alt+Shift+W → 2)
    [Ctrl+Alt+Win+Y] > Window tools [3]: tile background windows (also Win+Alt+Shift+W → 3)
    [Ctrl+Alt+Win+P] > Window tools [4]: exit F11 fullscreen (also Win+Alt+Shift+W → 4)
    [Ctrl+Alt+Win+A] > Move window to monitor 1 (left-most)
    [Ctrl+Alt+Win+S] > Move window to monitor 2
    [Ctrl+Alt+Win+D] > Move window to monitor 3
    [Ctrl+Alt+Win+F] > Move window to monitor 4
    [Ctrl+Alt+Win+Shift+A] > Close window on monitor 1
    [Ctrl+Alt+Win+Shift+S] > Close window on monitor 2
    [Ctrl+Alt+Win+Shift+D] > Close window on monitor 3
    [Ctrl+Alt+Win+Shift+F] > Close window on monitor 4
    [Ctrl+Alt+Win+Q] > Cycle windows on monitor 1
    [Ctrl+Alt+Win+W] > Cycle windows on monitor 2
    [Ctrl+Alt+Win+E] > Cycle windows on monitor 3
    [Ctrl+Alt+Win+R] > Cycle windows on monitor 4
    [Ctrl+Alt+Win+Shift+Q] > Minimize window on monitor 1
    [Ctrl+Alt+Win+Shift+W] > Minimize window on monitor 2
    [Ctrl+Alt+Win+Shift+E] > Minimize window on monitor 3
    [Ctrl+Alt+Win+Shift+R] > Minimize window on monitor 4
    
    === COMMAND PALETTE BOOKMARKS ===
    [Ctrl+Alt+Win+M] > Add bookmark (Command Palette Bookmark extension)
    
    === GENERAL ===
    [Win+Alt+Shift+U] > Quick string shortcuts
    [Ctrl+Alt+Win+4] > Send AI Text Optimizer prompt to {AI_PROVIDER} (same as Win+Alt+Shift+U then L, 4)
    [Win+Alt+Shift+Q] > Jump mouse on the middle
    [Win+Alt+Shift+X] > Peek PDF (tap) / Set PDF path (hold 700ms+)
    [Win+Alt+Shift+→] > Show square selector (right direction)
    [Win+Alt+Shift+←] > Show square selector (left direction)
    [Win+Alt+Shift+↓] > Show square selector (down direction)
    [Win+Alt+Shift+↑] > Show square selector (up direction)
    [Win+Alt+Shift+.] > Clip Angel (copy, paste, and quit)
    
    === COMMAND PALETTE ===
    [Win+Ctrl+Alt+Y] > Command Palette - File search
    [Shift+D] > Command Palette (active): exclude current bookmark (confirm)
    
    === SHORTCUTS ===
    [Win+Alt+Shift+A] > Show app-specific shortcuts (quick press)
    [Win+Alt+Shift+A] > Show global shortcuts (hold 700ms+)
    [Win+Alt+Shift+/] > Search all cheat sheets (cross-context)
    
    === WIKIPEDIA ===
    [Win+Alt+Shift+K] > Opens or activates Wikipedia
)"

; Returns Map of context label -> array of matching processed lines. Empty query => empty map.
SearchCheatSheets(query, includeGlobal := true) {
    global cheatSheets
    q := Trim(query)
    results := Map()
    if (q = "")
        return results

    for key, text in cheatSheets {
        if (text = "")
            continue
        proc := ProcessCheatSheetText(NormalizeMojibake(text))
        lines := StrSplit(proc, "`n")
        hits := []
        for line in lines {
            if CheatSheet_LineMatchesQuery(line, q)
                hits.Push(line)
        }
        if (hits.Length)
            results[key] := hits
    }

    if includeGlobal {
        proc := ProcessCheatSheetText(NormalizeMojibake(GetGlobalCheatSheetRawText()))
        lines := StrSplit(proc, "`n")
        hits := []
        for line in lines {
            if CheatSheet_LineMatchesQuery(line, q)
                hits.Push(line)
        }
        if (hits.Length)
            results["(Global shortcuts)"] := hits
    }
    return results
}

CheatSheet_ResizeBody(editCtrl, gui, fontLinePx := 18, minH := 200, lineCountSource := "", chromeAboveBodyPx := 48) {
    text := lineCountSource != "" ? lineCountSource : editCtrl.Value
    lineCnt := StrLen(text) ? StrSplit(text, "`n").Length : 1
    controlHeight := lineCnt * fontLinePx + 10
    if (controlHeight < minH)
        controlHeight := minH
    ; Same work-rect as standard banners (Utils.GetActiveMonitorWorkArea_StandardBar).
    GetActiveMonitorWorkArea_StandardBar(&ml, &mt, &mr, &mb)
    workH := mb - mt
    margin := 6
    maxHeight := workH - chromeAboveBodyPx - margin
    if (maxHeight < 80)
        maxHeight := 80
    if (controlHeight > maxHeight)
        controlHeight := maxHeight
    ; Pixel height for body; Custom RichEdit50W may not use rN row metrics like a built-in Edit.
    editCtrl.Move(, , 1000, controlHeight)
    gui.Show("AutoSize Hide")
    ; If chrome estimate was low, total GUI height can still exceed work area — shrink body once.
    gui.GetPos(, , &gw, &gh)
    maxGuiH := workH - margin
    if (gh > maxGuiH) {
        newH := controlHeight - (gh - maxGuiH)
        if (newH < 80)
            newH := 80
        editCtrl.Move(, , 1000, newH)
        gui.Show("AutoSize Hide")
    }
    CenterGuiOnActiveMonitor(gui)
    gui.Show()
}

CheatSheet_OnAppFilterChanged(*) {
    global g_helpSearchEdit, g_helpCheatCtrl, g_cheatSheetAppFullProcessed, g_helpGui, g_helpShown
    if (!IsObject(g_helpCheatCtrl) || !IsObject(g_helpSearchEdit))
        return
    q := Trim(g_helpSearchEdit.Value)
    body := g_cheatSheetAppFullProcessed
    displayBody := ""
    if (q = "") {
        displayBody := body
    } else {
        displayBody := CheatSheet_BuildFilteredBodyWithSections(body, q)
    }
    CheatSheet_RichSetProcessedBody(g_helpCheatCtrl, displayBody)
    CheatSheet_ResizeBody(g_helpCheatCtrl, g_helpGui, 18, 200, displayBody, 48)
    CheatSheet_DeferFocusSearch(g_helpSearchEdit)
}

CheatSheet_OnGlobalFilterChanged(*) {
    global g_globalSearchEdit, g_globalCheatCtrl, g_cheatSheetGlobalFullProcessed, g_globalGui, g_globalShown
    if (!IsObject(g_globalCheatCtrl) || !IsObject(g_globalSearchEdit))
        return
    q := Trim(g_globalSearchEdit.Value)
    body := g_cheatSheetGlobalFullProcessed
    displayBody := ""
    if (q = "") {
        displayBody := body
    } else {
        displayBody := CheatSheet_BuildFilteredBodyWithSections(body, q)
    }
    CheatSheet_RichSetProcessedBody(g_globalCheatCtrl, displayBody)
    CheatSheet_ResizeBody(g_globalCheatCtrl, g_globalGui, 16, 180, displayBody, 46)
    CheatSheet_DeferFocusSearch(g_globalSearchEdit)
}

ShowSearchAllCheatSheetsGui() {
    global g_searchAllGui
    static filterEdit := 0, lv := 0

    if !IsObject(g_searchAllGui) {
        g_searchAllGui := Gui("+AlwaysOnTop +Resize +MinSize800x400", "Search cheat sheets")
        g_searchAllGui.SetFont("s10", "Consolas")
        g_searchAllGui.Add("Text", "xm Section",
            "Filter (description text only; space = AND; max 20 chars):")
        filterEdit := g_searchAllGui.Add("Edit", "xs w780 Limit20")
        lv := g_searchAllGui.Add("ListView", "xm w780 h520 Grid", ["Context", "Line"])
        filterEdit.OnEvent("Change", (*) => CheatSheet_RefreshSearchAllList(filterEdit, lv))
        lv.OnEvent("DoubleClick", (ctrl, guiEvent) => CheatSheet_OnSearchAllCopy(ctrl, guiEvent))
        g_searchAllGui.Add("Text", "xm", "Double-click a row to copy the line.")
        g_searchAllGui.OnEvent("Escape", CheatSheet_OnEscapeSearchAll)
    }
    filterEdit.Value := ""
    CheatSheet_RefreshSearchAllList(filterEdit, lv)
    g_searchAllGui.Show("w800 h620")
    CheatSheet_DeferFocusSearch(filterEdit)
}

CheatSheet_RefreshSearchAllList(filterEdit, lv) {
    lv.Delete()
    q := Trim(filterEdit.Value)
    if (q = "") {
        return
    }
    m := SearchCheatSheets(q, true)
    for ctx, lines in m {
        for line in lines {
            lv.Add("", ctx, line)
        }
    }
}

CheatSheet_OnSearchAllCopy(lv, guiEvent) {
    row := 0
    try
        row := guiEvent.EventInfo
    if !row
        row := lv.GetNext(0, "F")
    if !row
        return
    line := lv.GetText(row, 2)
    A_Clipboard := line
}

; ========== GUI creation & showing ========================================
ToggleShortcutHelp() {
    global g_helpGui, g_helpShown, g_helpSearchEdit, g_helpCheatCtrl, g_cheatSheetAppFullProcessed

    ; Toggle off if currently shown
    if (IsObject(g_helpGui) && g_helpShown) {
        g_helpGui.Hide()
        g_helpShown := false
        return
    }

    text := NormalizeMojibake(GetCheatSheetText())
    usedFallback := false
    if (text = "") {
        exe := WinGetProcessName("A")
        text := "No cheat-sheet registered for:`n" exe
        usedFallback := true
    }

    if !IsObject(g_helpGui) {
        CheatSheet_EnsureRichDll()
        g_helpGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Border +Owner +LastFound")
        g_helpGui.BackColor := "000000"
        g_helpGui.SetFont("s11 cFFFF00", "Consolas")
        g_helpSearchEdit := g_helpGui.Add("Edit", "xm w1000 Section Limit20", "")
        g_helpCheatCtrl := g_helpGui.Add("Custom",
            "ClassRichEdit50W xs+0 y+4 +0x44 +Multi -E0x200 +VScroll -HScroll -Border Background000000 w1000 r12")
        g_helpSearchEdit.OnEvent("Change", CheatSheet_OnAppFilterChanged)
        g_helpGui.OnEvent("Escape", CheatSheet_OnEscapeApp)
    }

    processedText := ProcessCheatSheetText(text)
    g_cheatSheetAppFullProcessed := processedText
    g_helpSearchEdit.Value := ""
    CheatSheet_RichSetProcessedBody(g_helpCheatCtrl, processedText)
    CheatSheet_ResizeBody(g_helpCheatCtrl, g_helpGui, 18, 200, processedText, 48)
    g_helpShown := true
    CheatSheet_DeferFocusSearch(g_helpSearchEdit)
}

; ========== Global shortcuts cheat sheet (Win+Alt+Shift+key) ===============
ShowGlobalShortcutsHelp() {
    global g_globalGui, g_globalShown, g_globalSearchEdit, g_globalCheatCtrl, g_cheatSheetGlobalFullProcessed

    if (IsObject(g_globalGui) && g_globalShown) {
        g_globalGui.Hide()
        g_globalShown := false
        return
    }

    if !IsObject(g_globalGui) {
        CheatSheet_EnsureRichDll()
        g_globalGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Border +Owner +LastFound")
        g_globalGui.BackColor := "000000"
        g_globalGui.SetFont("s9 c00BFFF", "Consolas")
        g_globalSearchEdit := g_globalGui.Add("Edit", "xm w1000 Section Limit20", "")
        g_globalCheatCtrl := g_globalGui.Add("Custom",
            "ClassRichEdit50W xs+0 y+4 +0x44 +Multi +VScroll -HScroll -Border Background000000 w1000 r12")
        g_globalSearchEdit.OnEvent("Change", CheatSheet_OnGlobalFilterChanged)
        g_globalGui.OnEvent("Escape", CheatSheet_OnEscapeGlobal)
    }

    normalizedText := NormalizeMojibake(GetGlobalCheatSheetRawText())
    processedText := ProcessCheatSheetText(normalizedText)
    g_cheatSheetGlobalFullProcessed := processedText
    g_globalSearchEdit.Value := ""
    CheatSheet_RichSetProcessedBody(g_globalCheatCtrl, processedText)
    CheatSheet_ResizeBody(g_globalCheatCtrl, g_globalGui, 16, 180, processedText, 46)
    g_globalShown := true
    CheatSheet_DeferFocusSearch(g_globalSearchEdit)
}

; ========== Hotkey with hold detection ====================================
; Win + Alt + Shift + A with hold detection
#!+a::
{
    global g_helpGui, g_helpShown, g_globalGui, g_globalShown, g_searchAllGui

    ; First check if any cheat sheet is currently open - if so, close it
    if (IsObject(g_helpGui) && g_helpShown) {
        g_helpGui.Hide()
        g_helpShown := false
        return
    }

    if (IsObject(g_globalGui) && g_globalShown) {
        g_globalGui.Hide()
        g_globalShown := false
        return
    }

    if (IsObject(g_searchAllGui)) {
        g_searchAllGui.Hide()
        return
    }

    ; No cheat sheet is open, determine which one to show based on hold time
    static pressTime := 0
    pressTime := A_TickCount

    ; Wait for key release or timeout (increased to accommodate 1s+ holds)
    KeyWait "a", "T1"  ; Wait max 1.5s for key release

    holdTime := A_TickCount - pressTime

    if (holdTime >= 700) {
        ; Long hold (1s+) - show global shortcuts
        ShowGlobalShortcutsHelp()
    } else {
        ; Quick press - show app-specific shortcuts
        ToggleShortcutHelp()
    }
}

; Win+Alt+Shift+/ — search all registered cheat sheets (ListView; double-click row copies line)
#!+/::
{
    ShowSearchAllCheatSheetsGui()
}

; =============================================================================
; Send Top List Item from Clip Angel
; Hotkey: Win+Alt+Shift+1
; =============================================================================
#!+1::
{
    ClipAngel_SendTopListItem()
}

; =============================================================================
; Clip Angel: Fast Copy Mode + sequential paste (multiple clips in order)
; Hotkey: Win+Alt+Shift+J — while mode off: tap starts mode; hold 700ms+ repeats last paste count.
;         While mode on: press finishes and pastes N clips (Ctrl+C / PrtSc / Alt+PrtSc counted).
; =============================================================================
global FAST_COPY_HOLD_REPEAT_MS := 700
; Set true to write FastCopyMode_DebugLog NDJSON (development only).
global FAST_COPY_DEBUG := false
global FAST_COPY_CLIPBOARD_READ_CYCLE_MS := 1200
global FAST_COPY_GEMINI_UPLOAD_IDLE_MS := 5000
global gFastCopyModeActive := false
global gFastCopyCount := 0
global gFastCopyPasteTargetHwnd := 0
global gFastCopyLastSuccessfulCount := 0
global gFastCopyScreenshotQueue := []
global gFastCopyLastScreenshotQueue := []

FastCopyMode_DebugLog(hypothesisId, location, message, data := "") {
    ; #region agent log
    global FAST_COPY_DEBUG
    if (!FAST_COPY_DEBUG)
        return
    ; Writes NDJSON to debug-1bed80.log (Debug session: 1bed80)
    try {
        runId := "pre-fix"
        logPath := "C:\Users\fie7ca\Documents\scripts\debug-1bed80.log"
        ; Keep data small and non-sensitive; accept either a string or a Map-like object.
        dataJson := "{}"
        if (IsObject(data)) {
            parts := []
            for k, v in data {
                try parts.Push('"' FastCopyMode_JsonEscape(k) '":"' FastCopyMode_JsonEscape(v) '"')
            }
            joined := ""
            if (parts.Length) {
                for i, p in parts {
                    joined .= (i = 1 ? "" : ",") p
                }
            }
            dataJson := "{" joined "}"
        } else if (data != "") {
            dataJson := '{"value":"' FastCopyMode_JsonEscape(data) '"}'
        }
        line := '{'
            . '"sessionId":"1bed80",'
            . '"timestamp":' A_TickCount + 0 ','
            . '"runId":"' runId '",'
            . '"hypothesisId":"' FastCopyMode_JsonEscape(hypothesisId) '",'
            . '"location":"' FastCopyMode_JsonEscape(location) '",'
            . '"message":"' FastCopyMode_JsonEscape(message) '",'
            . '"data":' dataJson
            . '}'
        FileAppend(line "`n", logPath, "UTF-8")
    } catch {
        ; never break user flow
    }
    ; #endregion agent log
}

FastCopyMode_JsonEscape(s) {
    ; #region agent log
    try {
        if (s = "")
            return ""
    } catch {
        return ""
    }
    s := "" s
    s := StrReplace(s, "\", "\\")
    s := StrReplace(s, '"', '\"')
    s := StrReplace(s, "`r", "\r")
    s := StrReplace(s, "`n", "\n")
    return s
    ; #endregion agent log
}

FastCopyMode_ClipboardHasImage() {
    ; #region agent log
    ; CF_DIB=8, CF_DIBV5=17, CF_BITMAP=2
    try {
        return !!(DllCall("IsClipboardFormatAvailable", "UInt", 8, "Int")
        || DllCall("IsClipboardFormatAvailable", "UInt", 17, "Int")
        || DllCall("IsClipboardFormatAvailable", "UInt", 2, "Int"))
    } catch {
        return false
    }
    ; #endregion agent log
}

FastCopyMode_WaitForClipboardImage(timeoutMs := 1200) {
    ; #region agent log
    start := A_TickCount
    while ((A_TickCount - start) < timeoutMs) {
        if (FastCopyMode_ClipboardHasImage())
            return true
        Sleep 15
    }
    return false
    ; #endregion agent log
}

FastCopyMode_CanOpenClipboardNow() {
    ; #region agent log
    ; Returns true if clipboard can be opened immediately (no long lock).
    ok := false
    try {
        if (DllCall("OpenClipboard", "ptr", 0, "int")) {
            ok := true
            DllCall("CloseClipboard")
        }
    } catch {
        ok := false
    }
    return ok
    ; #endregion agent log
}

FastCopyMode_WaitForClipboardUnlocked(timeoutMs := 2000) {
    ; #region agent log
    start := A_TickCount
    while ((A_TickCount - start) < timeoutMs) {
        if (FastCopyMode_CanOpenClipboardNow())
            return true
        Sleep 15
    }
    return false
    ; #endregion agent log
}

FastCopyMode_WaitForClipboardReadCycle(timeoutMs := 1200) {
    ; #region agent log
    ; Wait until we observe some other process reading/locking the clipboard (OpenClipboard fails)
    ; and then wait until it becomes unlocked again. This reduces the risk of overwriting the
    ; clipboard before the target app has actually consumed the bitmap.
    start := A_TickCount
    sawLock := false
    failCount := 0
    okCount := 0

    ; Phase 1: short window with Sleep 1 (avoid busy-spinning the hotkey thread).
    sampleUntil := start + 120
    while (A_TickCount < sampleUntil) {
        if (!FastCopyMode_CanOpenClipboardNow()) {
            failCount += 1
            sawLock := true
            break
        }
        okCount += 1
        Sleep 1
    }

    ; Phase 2: regular polling until timeout budget (handles longer locks).
    if (!sawLock) {
        while ((A_TickCount - start) < timeoutMs) {
            if (!FastCopyMode_CanOpenClipboardNow()) {
                failCount += 1
                sawLock := true
                break
            }
            okCount += 1
            Sleep 10
        }
    }
    if (sawLock) {
        unlockMs := Min(1500, Max(100, timeoutMs - (A_TickCount - start)))
        if (FastCopyMode_WaitForClipboardUnlocked(unlockMs))
            return "lock_then_unlock(fails=" failCount ",oks=" okCount ")"
        return "lock_timeout(fails=" failCount ",oks=" okCount ")"
    }
    ; Never saw the clipboard get locked (target might not read immediately or at all).
    return "no_lock_seen(fails=" failCount ",oks=" okCount ")"
    ; #endregion agent log
}

FastCopyMode_SendPasteAndWaitForReadCycle(isGeminiSession := false) {
    ; #region agent log
    Send "^v"
    if (isGeminiSession) {
        Sleep 200
        return "gemini_delay"
    } else {
        Sleep 400
        return "baseline_delay"
    }
    ; #endregion agent log
}

FastCopyMode_IsGeminiActiveInChrome() {
    ; #region agent log
    try {
        return (WinActive("ahk_exe chrome.exe") && InStr(WinGetTitle("A"), "gemini", false))
    } catch {
        return false
    }
    ; #endregion agent log
}

FastCopyMode_GetActiveChromeHwnd() {
    ; #region agent log
    try {
        hwnd := WinGetID("A")
        if (hwnd && WinExist("ahk_id " hwnd) && WinActive("ahk_id " hwnd))
            return hwnd
    } catch {
    }
    return 0
    ; #endregion agent log
}

FastCopyMode_UiaForActiveChrome() {
    ; #region agent log
    hwnd := FastCopyMode_GetActiveChromeHwnd()
    if (!hwnd)
        throw Error("no_active_hwnd")
    return UIA_Browser("ahk_id " hwnd)
    ; #endregion agent log
}

FastCopyMode_GetGeminiSearchRoot(uia) {
    ; #region agent log
    ; Same pattern as GetGeminiSearchRoot in Gemini.ahk — smaller UIA subtree than full document.
    try {
        root := uia.GetCurrentMainPaneElement()
        if (root)
            return root
    } catch {
    }
    return uia
    ; #endregion agent log
}

FastCopyMode_FocusGeminiPromptField(uia := "") {
    ; #region agent log
    ; Optional uia: reuse cached UIA_Browser for the whole paste loop (efficiency).
    if (uia = "") {
        try {
            uia := FastCopyMode_UiaForActiveChrome()
        } catch Error as e {
            FastCopyMode_DebugLog("H6", "Shift keys.ahk:FastCopyMode_FocusGeminiPromptField", "uia_attach_error", Map(
                "msg", SubStr(e.Message, 1, 120)
            ))
            return false
        }
    }

    try {
        ; Use shared helper from Utils.ahk (Gemini.ahk uses the same).
        promptField := FindGeminiPromptField(uia)

        if (promptField && IsObject(promptField)) {
            try {
                promptField.SetFocus()
            } catch {
                try promptField.Click()
                catch {
                }
            }
            Sleep 80
            try {
                if (promptField.HasKeyboardFocus)
                    return true
            } catch {
                return true
            }
        }
    } catch Error as e {
        FastCopyMode_DebugLog("H6", "Shift keys.ahk:FastCopyMode_FocusGeminiPromptField", "focus_error", Map(
            "msg", SubStr(e.Message, 1, 120)
        ))
    }
    return false
    ; #endregion agent log
}

FastCopyMode_GeminiIsUploadingImage(uia) {
    ; #region agent log
    ; Heuristic: look for common uploading labels in Gemini UI (English + PT-BR).
    ; Scoped to main pane only; do not treat every ProgressBar as upload (false positives).
    try {
        root := FastCopyMode_GetGeminiSearchRoot(uia)
        texts := root.FindAll({ Type: 50020 }) ; Text
        for t in texts {
            name := t.Name
            if (!name)
                continue
            low := StrLower(name)
            if (InStr(low, "open upload file menu"))
                continue
            if (InStr(low, "upload") || InStr(low, "sending") || InStr(low, "carreg") || InStr(low, "enviando")) {
                return true
            }
        }
    } catch {
    }
    return false
    ; #endregion agent log
}

FastCopyMode_WaitForGeminiUploadIdle(uia := "", timeoutMs := 0) {
    ; #region agent log
    global FAST_COPY_GEMINI_UPLOAD_IDLE_MS
    if (timeoutMs <= 0)
        timeoutMs := FAST_COPY_GEMINI_UPLOAD_IDLE_MS
    ; Wait until Gemini is no longer showing an upload-in-progress indicator.
    start := A_TickCount
    if (uia = "") {
        try {
            uia := FastCopyMode_UiaForActiveChrome()
        } catch Error as e {
            FastCopyMode_DebugLog("H6", "Shift keys.ahk:FastCopyMode_WaitForGeminiUploadIdle", "uia_browser_error", Map(
                "msg", SubStr(e.Message, 1, 120)
            ))
            return "uia_fail"
        }
    }

    while ((A_TickCount - start) < timeoutMs) {
        if (!FastCopyMode_GeminiIsUploadingImage(uia))
            return "idle"
        Sleep 150
    }
    return "timeout"
    ; #endregion agent log
}

; --- Gemini + Clip Angel: sequential paste with upload-aware pacing (same prompt focus as Gemini.ahk #!+i) ---

FastCopyMode_IsGeminiHwnd(hwnd) {
    try {
        if (!hwnd)
            return false
        proc := WinGetProcessName("ahk_id " hwnd)
        if (proc != "chrome.exe")
            return false
        title := ""
        try title := WinGetTitle("ahk_id " hwnd)
        if (InStr(title, "gemini", false))
            return true
        ; Fallback: title can be generic; verify by finding the Gemini prompt field via UIA.
        try {
            uia := UIA_Browser("ahk_id " hwnd)
            pf := FindGeminiPromptField(uia)
            if (pf)
                return true
        } catch {
        }
    } catch {
    }
    return false
}

FastCopyMode_IsGeminiForeground() {
    try {
        return FastCopyMode_IsGeminiHwnd(WinGetID("A"))
    } catch {
        return false
    }
}

Gemini_GetUiaForActiveGeminiChrome() {
    try {
        hwnd := WinGetID("A")
        if (!FastCopyMode_IsGeminiHwnd(hwnd))
            return ""
        if (!hwnd)
            return ""
        return UIA_Browser("ahk_id " hwnd)
    } catch {
        return ""
    }
}

; Gemini_FocusPromptSameAsOpenHotkey — see Utils.ahk (shared with Gemini.ahk #!+i).

; After each Clip Angel / screenshot paste: bounded wait until upload UI clears; refocus prompt while uploading.
; timeoutMs: max wait (default FAST_COPY_GEMINI_UPLOAD_IDLE_MS). minNoIndicatorMs: if we never see "uploading",
; wait at least this long before returning idle (approximates legacy fixed tail when no indicator appears).
Gemini_WaitForUploadIdleWithRefocus(uia, timeoutMs := 0, minNoIndicatorMs := 500) {
    global FAST_COPY_GEMINI_UPLOAD_IDLE_MS
    if (!IsObject(uia))
        return "uia_fail"
    if (timeoutMs <= 0)
        timeoutMs := FAST_COPY_GEMINI_UPLOAD_IDLE_MS
    tStart := A_TickCount
    sawUploading := false
    while ((A_TickCount - tStart) < timeoutMs) {
        up := FastCopyMode_GeminiIsUploadingImage(uia)
        if (up)
            sawUploading := true
        if (!up) {
            if (sawUploading || (A_TickCount - tStart) >= minNoIndicatorMs)
                return "idle"
        } else
            FastCopyMode_FocusGeminiPromptField(uia)
        Sleep 150
    }
    return "timeout"
}

; One Clip Angel item per iteration: open once, ^!b with gaps; upload wait between images.
Gemini_PasteFromClipAngelSequential(count, uia := "") {
    if (!IsInteger(count))
        return
    n := Integer(count)
    if (n < 1)
        return
    if (uia = "") {
        uia := Gemini_GetUiaForActiveGeminiChrome()
        if (!IsObject(uia))
            return
    }
    if !ClipAngel_TryAcquireAutomationLock()
        return
    priorHwnd := 0
    try priorHwnd := WinGetID("A")
    catch {
    }
    try {
        StandardLoadingBar_Show("⏳ Pasting clips in Gemini…", BANNER_ACCENT_INTERMEDIATE, { passive: false,
            centerOnHwnd: 0,
            fontSize: 17 })
    } catch {
    }
    try {
        ClipAngel_WaitChordModifiersReleased()
        ClipAngel_ReleaseChordModifiersForSend()
        loop n {
            try StandardLoadingBar_Update("⏳ Pasting clip " A_Index " / " n " …")
            catch {
            }
            Gemini_FocusPromptSameAsOpenHotkey(uia, false)
            if (A_Index = 1) {
                ClipAngel_SendNativeTopItemKeys(priorHwnd)
            } else {
                Sleep(CLIPANGEL_SEQUENTIAL_PASTE_GAP_MS)
                ClipAngel_ReleaseChordModifiersForSend()
                SendInput "^!b"
            }
            ; Brief settle after paste, then condition-based wait for upload UI (efficiency-canon: bounded
            ; wait vs fixed 2.6s). minNoIndicatorMs 2600 preserves ~legacy tail when no upload indicator.
            Sleep 400
            try FastCopyMode_FocusGeminiPromptField(uia)
            try Gemini_WaitForUploadIdleWithRefocus(uia, 4000, 2600)
            try FastCopyMode_FocusGeminiPromptField(uia)
        }
    } finally {
        EnsureClipAngelClosed()
        ClipAngel_RestorePriorFocus(priorHwnd)
        ClipAngel_ReleaseAutomationLock()
        try StandardLoadingBar_Hide(0)
        try {
            if (FastCopyMode_IsGeminiForeground()) {
                if (IsObject(uia)) {
                    try FastCopyMode_FocusGeminiPromptField(uia)
                } else {
                    aw := WinGetID("A")
                    if (aw)
                        FocusGeminiAskFieldForHwnd(aw, false)
                }
            }
        } catch {
        }
    }
}

FastCopyMode_ReleaseHotkeyModifiers() {
    ; Ensure the Win+Alt+Shift hotkey modifiers can't leak into paste keys.
    ; Releasing modifiers does not activate or focus any other window.
    Send "{LWin up}{RWin up}{Alt up}{Shift up}{Ctrl up}"
    Sleep 30
}

FastCopyMode_CaptureScreenshotToQueue(clipboardAlreadyHasImage := false) {
    global gFastCopyScreenshotQueue

    FastCopyMode_DebugLog("H1", "Shift keys.ahk:FastCopyMode_CaptureScreenshotToQueue", "capture_enter", Map(
        "queueLenBefore", gFastCopyScreenshotQueue.Length,
        "hasImageBefore", FastCopyMode_ClipboardHasImage() ? "1" : "0"
    ))

    ; Give the OS a moment to push the screenshot into the clipboard.
    ; Alt+PrintScreen updates the clipboard with an image; rapid captures can overwrite each other
    ; unless we snapshot the clipboard right away.
    ; Wait for a real *image* to appear (not just "clipboard has something").
    ; Some callers (e.g. Win+Shift+S) already waited for the image and just want to snapshot.
    if (!clipboardAlreadyHasImage) {
        try A_Clipboard := ""
        ok := FastCopyMode_WaitForClipboardImage(1500)
    } else {
        ok := FastCopyMode_ClipboardHasImage()
    }
    FastCopyMode_DebugLog("H2", "Shift keys.ahk:FastCopyMode_CaptureScreenshotToQueue", "wait_image_done", Map(
        "ok", ok ? "1" : "0",
        "hasImageAfter", FastCopyMode_ClipboardHasImage() ? "1" : "0"
    ))
    if (!ok)
        return

    try {
        snap := ClipboardAll()
        gFastCopyScreenshotQueue.Push(snap)
        snapSize := ""
        try snapSize := snap.Size
        FastCopyMode_DebugLog("H1", "Shift keys.ahk:FastCopyMode_CaptureScreenshotToQueue", "capture_pushed", Map(
            "queueLenAfter", gFastCopyScreenshotQueue.Length,
            "hasImageAfter", FastCopyMode_ClipboardHasImage() ? "1" : "0",
            "snapType", Type(snap),
            "snapSize", snapSize
        ))
    } catch {
        ; If clipboard snapshot fails, just skip (count still increments).
        FastCopyMode_DebugLog("H2", "Shift keys.ahk:FastCopyMode_CaptureScreenshotToQueue", "clipboardall_failed", Map(
            "queueLenAfter", gFastCopyScreenshotQueue.Length,
            "hasImageAfter", FastCopyMode_ClipboardHasImage() ? "1" : "0"
        ))
    }
}

FastCopyMode_OnWinShiftS() {
    ; Win+Shift+S region snip: count only on successful clipboard image.
    try A_Clipboard := ""
    Send "#+s"
    ok := FastCopyMode_WaitForClipboardImage(30000)
    if (!ok)
        return
    FastCopyMode_OnCopy()
    FastCopyMode_CaptureScreenshotToQueue(true)
}

FastCopyMode_PasteScreenshotQueue(queue) {
    if (!IsObject(queue) || queue.Length < 1)
        return

    total := queue.Length
    try {
        StandardLoadingBar_Show("⏳ Pasting screenshots…", BANNER_ACCENT_INTERMEDIATE, { passive: false, centerOnHwnd: 0,
            fontSize: 17 })
    } catch {
    }

    clipSave := ""
    try clipSave := ClipboardAll()

    try {
        try {
            proc := WinGetProcessName("A")
            cls := WinGetClass("A")
            title := WinGetTitle("A")
        } catch {
            proc := ""
            cls := ""
            title := ""
        }
        isGeminiSession := (proc = "chrome.exe" && InStr(title, "gemini", false))
        cachedGeminiUia := ""
        if (isGeminiSession) {
            try {
                hwndGem := FastCopyMode_GetActiveChromeHwnd()
                if (hwndGem)
                    cachedGeminiUia := UIA_Browser("ahk_id " hwndGem)
            } catch {
                cachedGeminiUia := ""
            }
        }
        FastCopyMode_DebugLog("H3", "Shift keys.ahk:FastCopyMode_PasteScreenshotQueue", "paste_enter", Map(
            "queueLen", queue.Length,
            "hasImageAtEnter", FastCopyMode_ClipboardHasImage() ? "1" : "0",
            "proc", proc,
            "class", cls,
            "title", SubStr(title, 1, 120),
            "isGeminiSession", isGeminiSession ? "1" : "0"
        ))
        for idx, snap in queue {
            try {
                try StandardLoadingBar_Update("⏳ Pasting image " idx " / " total " …")
                catch {
                }
                ; Gemini: ensure prompt is focused BEFORE each paste.
                if (isGeminiSession) {
                    try {
                        focusedPre := FastCopyMode_FocusGeminiPromptField(cachedGeminiUia) ? "1" : "0"
                    } catch {
                        focusedPre := "0"
                    }
                    FastCopyMode_DebugLog("H6", "Shift keys.ahk:FastCopyMode_PasteScreenshotQueue",
                        "gemini_before_paste", Map(
                            "idx", idx,
                            "focusedPrompt", focusedPre
                        ))
                }

                A_Clipboard := snap
                ClipWait 0.6, 1
                snapSize := ""
                try snapSize := snap.Size
                FastCopyMode_DebugLog("H3", "Shift keys.ahk:FastCopyMode_PasteScreenshotQueue", "paste_iter_ready", Map(
                    "idx", idx,
                    "snapType", Type(snap),
                    "snapSize", snapSize,
                    "hasImageNow", FastCopyMode_ClipboardHasImage() ? "1" : "0"
                ))
                cycle := FastCopyMode_SendPasteAndWaitForReadCycle(isGeminiSession)
                FastCopyMode_DebugLog("H4", "Shift keys.ahk:FastCopyMode_PasteScreenshotQueue",
                    "paste_iter_clipboard_cycle", Map(
                        "idx", idx,
                        "cycle", cycle
                    ))

                ; Gemini: condition-based upload wait (early exit when UI idle) vs fixed 2.6s sleep.
                if (isGeminiSession) {
                    Sleep 400
                    try FastCopyMode_FocusGeminiPromptField(cachedGeminiUia)
                    idleStatus := IsObject(cachedGeminiUia) ? Gemini_WaitForUploadIdleWithRefocus(cachedGeminiUia, 5000,
                        800) : "no_uia"
                    if (idleStatus = "no_uia") {
                        Sleep 2600
                        idleStatus := "fallback_fixed_delay"
                    }
                    try FastCopyMode_FocusGeminiPromptField(cachedGeminiUia)
                    FastCopyMode_DebugLog("H6", "Shift keys.ahk:FastCopyMode_PasteScreenshotQueue",
                        "gemini_after_paste", Map(
                            "idx", idx,
                            "uploadIdle", idleStatus
                        ))
                }
            } catch {
                ; continue to next screenshot
                FastCopyMode_DebugLog("H4", "Shift keys.ahk:FastCopyMode_PasteScreenshotQueue", "paste_iter_failed",
                    Map(
                        "idx", idx
                    ))
            }
        }
    } finally {
        if (clipSave != "")
            try A_Clipboard := clipSave
        try StandardLoadingBar_Hide(0)
        FastCopyMode_DebugLog("H3", "Shift keys.ahk:FastCopyMode_PasteScreenshotQueue", "paste_exit", Map(
            "restoredClipboard", clipSave != "" ? "1" : "0"
        ))
    }
}

ExecuteSequentialPaste(actionCount) {
    if (!IsInteger(actionCount))
        return
    n := Integer(actionCount)
    if (n < 1)
        return
    global gFastCopyPasteTargetHwnd
    if !ClipAngel_TryAcquireAutomationLock()
        return
    priorHwnd := gFastCopyPasteTargetHwnd
    if (!priorHwnd || !WinExist("ahk_id " priorHwnd))
        priorHwnd := ClipAngel_ResolvePriorHwnd(0)
    try {
        StandardLoadingBar_Show("⏳ Pasting from clipboard…", BANNER_ACCENT_INTERMEDIATE, { passive: false, centerOnHwnd: 0,
            fontSize: 17 })
    } catch {
    }
    try {
        ClipAngel_WaitChordModifiersReleased()
        ClipAngel_ReleaseChordModifiersForSend()
        loop n {
            try StandardLoadingBar_Update("⏳ Pasting clip " A_Index " / " n " …")
            catch {
            }
            if (A_Index = 1) {
                ClipAngel_SendNativeTopItemKeys(priorHwnd)
            } else {
                Sleep(CLIPANGEL_SEQUENTIAL_PASTE_GAP_MS)
                ClipAngel_ReleaseChordModifiersForSend()
                SendInput "^!b"
            }
            Sleep(CLIPANGEL_INCREMENTAL_PASTE_SETTLE_MS)
        }
    } finally {
        EnsureClipAngelClosed()
        ClipAngel_RestorePriorFocus(priorHwnd)
        ClipAngel_ReleaseAutomationLock()
        try StandardLoadingBar_Hide(0)
    }
}

FastCopyMode_IsActive() {
    global gFastCopyModeActive
    return gFastCopyModeActive
}

FastCopyMode_OnCopy() {
    global gFastCopyCount
    gFastCopyCount += 1
    FastCopyModeBanner_Update(gFastCopyCount)
}

FastCopyMode_PlayCueSound(fileName) {
    if (!IsSoundEnabled())
        return
    path := A_ScriptDir "\sounds\" fileName
    if (!FileExist(path))
        return
    try {
        ScriptSoundPlay(path)
    } catch {
    }
}

FastCopyMode_Start() {
    global gFastCopyModeActive, gFastCopyCount, gFastCopyPasteTargetHwnd, gFastCopyScreenshotQueue
    try {
        gFastCopyPasteTargetHwnd := WinGetID("A")
    } catch {
        gFastCopyPasteTargetHwnd := 0
    }
    gFastCopyCount := 0
    gFastCopyScreenshotQueue := []
    gFastCopyModeActive := true
    try {
        FastCopyModeBanner_Show()
        FastCopyMode_PlayCueSound("fastcopy-start.mp3")
    } catch Error {
        gFastCopyModeActive := false
        ShowCenteredOverlay_Utils("❌ Could not start Fast Copy Mode", 2000, BANNER_ACCENT_ERROR)
    }
}

FastCopyMode_Finish() {
    global gFastCopyModeActive, gFastCopyCount, gFastCopyPasteTargetHwnd
    global gFastCopyScreenshotQueue, gFastCopyLastScreenshotQueue
    count := gFastCopyCount
    shotCount := IsObject(gFastCopyScreenshotQueue) ? gFastCopyScreenshotQueue.Length : 0
    try {
        FastCopyModeBanner_Hide()
    } finally {
        gFastCopyModeActive := false
        gFastCopyCount := 0
    }
    try {
        ; Paste exclusively into the *currently active* window without activating anything else.
        FastCopyMode_ReleaseHotkeyModifiers()
        if (count > 0) {
            FastCopyMode_PlayCueSound("fastcopy-finish.mp3")
            FastCopyMode_DebugLog("H5", "Shift keys.ahk:FastCopyMode_Finish", "finish_paste_start", Map(
                "count", count,
                "shotCount", shotCount
            ))
            if (shotCount > 0) {
                FastCopyMode_PasteScreenshotQueue(gFastCopyScreenshotQueue)
                ; Save for hold-to-repeat behavior.
                gFastCopyLastScreenshotQueue := gFastCopyScreenshotQueue.Clone()
            } else {
                gFastCopyLastScreenshotQueue := []
            }

            remaining := count - shotCount
            FastCopyMode_DebugLog("H5", "Shift keys.ahk:FastCopyMode_Finish", "finish_remaining", Map(
                "remaining", remaining
            ))
            if (remaining > 0) {
                ; Non-screenshot copies: Clip Angel sequential paste (Gemini uses upload-aware loop).
                if (FastCopyMode_IsGeminiForeground())
                    Gemini_PasteFromClipAngelSequential(remaining)
                else
                    ExecuteSequentialPaste(remaining)
            }

            global gFastCopyLastSuccessfulCount
            gFastCopyLastSuccessfulCount := count
        } else
            ShowCenteredOverlay_Utils("⚠ No copies recorded — nothing to paste", 2500, BANNER_ACCENT_INTERMEDIATE)
    } catch Error as e {
        ShowCenteredOverlay_Utils("❌ Fast Copy Mode: " SubStr(e.Message, 1, 80), 2500, BANNER_ACCENT_ERROR)
    }
}

FastCopyMode_RepeatLastPaste() {
    global gFastCopyLastSuccessfulCount
    global gFastCopyLastScreenshotQueue
    if (gFastCopyLastSuccessfulCount < 1) {
        ShowCenteredOverlay_Utils("⚠ No previous Fast Copy paste to repeat", 2500, BANNER_ACCENT_INTERMEDIATE)
        return
    }
    try {
        ; Repeat paste into the *currently active* window without activating anything else.
        FastCopyMode_ReleaseHotkeyModifiers()
        FastCopyMode_PlayCueSound("fastcopy-finish.mp3")
        if (IsObject(gFastCopyLastScreenshotQueue) && gFastCopyLastScreenshotQueue.Length > 0) {
            FastCopyMode_PasteScreenshotQueue(gFastCopyLastScreenshotQueue)
            remaining := gFastCopyLastSuccessfulCount - gFastCopyLastScreenshotQueue.Length
            if (remaining > 0) {
                if (FastCopyMode_IsGeminiForeground())
                    Gemini_PasteFromClipAngelSequential(remaining)
                else
                    ExecuteSequentialPaste(remaining)
            }
        } else {
            if (FastCopyMode_IsGeminiForeground())
                Gemini_PasteFromClipAngelSequential(gFastCopyLastSuccessfulCount)
            else
                ExecuteSequentialPaste(gFastCopyLastSuccessfulCount)
        }
    } catch Error as e {
        ShowCenteredOverlay_Utils("❌ Repeat paste: " SubStr(e.Message, 1, 80), 2500, BANNER_ACCENT_ERROR)
    }
}

#!+j:: {
    global gFastCopyModeActive, FAST_COPY_HOLD_REPEAT_MS
    if (gFastCopyModeActive) {
        FastCopyMode_Finish()
        return
    }
    pressTime := A_TickCount
    KeyWait "j", "T1"
    holdTime := A_TickCount - pressTime
    if (holdTime >= FAST_COPY_HOLD_REPEAT_MS)
        FastCopyMode_RepeatLastPaste()
    else
        FastCopyMode_Start()
}

; Win+Alt+Shift+L — Outlook Copilot shortcut modal (1–9). Global: works from any app; actions activate Outlook.
#!+l:: {
    SelectOutlookCopilotShortcut()
}

#HotIf FastCopyMode_IsActive()
~^c:: FastCopyMode_OnCopy()
~PrintScreen:: FastCopyMode_OnCopy()
~!PrintScreen:: {
    FastCopyMode_OnCopy()
    FastCopyMode_CaptureScreenshotToQueue()
}
$#+s:: FastCopyMode_OnWinShiftS()
#HotIf

;-------------------------------------------------------------------
; Environment paths (unchanged)
;-------------------------------------------------------------------
global WORK_SCRIPTS_PATH := "C:\Users\fie7ca\Documents\01 - Scripts"
global PERSONAL_SCRIPTS_PATH := "C:\Users\eduev\Meu Drive\17 - Projects\scripts"
; global IS_WORK_ENVIRONMENT   := true    ; set to false on personal rig // This will now be loaded from env.ahk

; ---------------------------------------------------------------------------
; ShowErr(msgOrErr)  â€" uniform MsgBox for any thrown value
; ---------------------------------------------------------------------------
ShowErr(err) {
    text := ""
    try {
        if (err is Error) {
            text := err.Message
            if (HasProp(err, "Extra") && err.Extra != "")
                text .= "`n`nExtra:`n" err.Extra
        } else if (err is String) {
            text := err
        } else {
            ; Covers TargetError and other thrown objects/values.
            text := Format("{}", err)
        }
    } catch {
        text := "<unprintable error value>"
    }
    MsgBox("Error:`n" text, "Error", "IconX")
}

; ---------------------------------------------------------------------------
; Centre cheat sheet on the same monitor as StandardLoadingBar / banners
; (GetActiveMonitorWorkArea_StandardBar in Utils.ahk). Do not clamp to (0,0):
; that pulls the window onto the primary display when wx/wy are negative and
; causes multi-monitor spanning / wrong placement.
; ---------------------------------------------------------------------------
CenterGuiOnActiveMonitor(guiObj) {
    guiObj.GetPos(, , &guiW, &guiH)
    GetActiveMonitorWorkArea_StandardBar(&wx, &wy, &wr, &wb)
    ww := wr - wx
    wh := wb - wy
    guiX := wx + (ww - guiW) / 2
    guiY := wy + (wh - guiH) / 2
    guiX := Max(wx, Min(guiX, wx + ww - guiW))
    guiY := Max(wy, Min(guiY, wy + wh - guiH))
    guiObj.Show("NoActivate x" Round(guiX) " y" Round(guiY))
}

;-------------------------------------------------------------------
; OneNote Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe onenote.exe") && !IsFileDialogActive()

; Shift + P : Onenote: select line and children
+p:: Send("^+-") ; Remaps to Ctrl + Shift + -

; Shift + F : Advanced Searching with double quotes
+f:: {
    Send "^f"
    Sleep 50
    Send "^a"
    Sleep 20
    Send "{Del}"
    Sleep 20
    Send '""'
    Sleep 20
    Send "{Left}"
}

; Shift + D : Onenote: delete line and children
+d:: {
    Send("^+-") ; Remaps to Ctrl + Shift + -
    Send "{Del}"
}

; Shift + S : Onenote: delete only current line (keep children)
+s:: {
    Send("+{Right}")
    Send "{Del}"
}

; Shift + U : Onenote: collapse
+u:: Send("!+{+}")     ; Remaps to Alt + Shift + +

; Shift + Y : Onenote: expand
+y:: Send("!+{-}")     ; Remaps to Alt + Shift + -

; Shift + I : Onenote: collapse all
+i:: Send("!+1")     ; Remaps to Alt + Shift + 1

; Shift + O : Onenote: expand all
+o:: Send("!+0")     ; Remaps to Alt + Shift + 0

#HotIf

;-------------------------------------------------------------------
; ClipAngel Shortcuts
;-------------------------------------------------------------------

; Global variables for ClipAngel filter selector
global g_ClipAngelFilterSelectorGui := false
global g_ClipAngelFilterSelectorActive := false
global g_ClipAngelFilterCharMap := Map()  ; Maps character to file type info
global g_ClipAngelFilterHotkeyHandlers := []  ; Store hotkey handlers for cleanup

; File type list (filtered to essential types only)
global g_ClipAngelFileTypes := [{ name: "img", index: 1, navKey: "I", navCount: 1 }, { name: "file", index: 2, navKey: "F",
    navCount: 1 }, { name: "text", index: 3, navKey: "T", navCount: 1 }, { name: "*url", index: 10, navKey: "*",
        navCount: 5 }, { name: "*filename", index: 13, navKey: "*", navCount: 8 }
]

; Character sequence for file type assignment (5 types)
global g_ClipAngelFilterCharSequence := ["1", "2", "3", "4", "5"]

#HotIf WinActive("ClipAngel")

; Shift + C : Select filtered content and copy
+c:: {
    Send "{Tab}"
    Sleep 100
    Send "{Tab}"
    Sleep 300
    Send "^a"  ; Select all
    Sleep 100
    Send "^c"  ; Copy
    Sleep 100
    Send "{F10}"
}

; Shift + T : Switch focus between list and text (Toggle) (F10)
+t:: Send "{F10}"

; Shift + D : Delete all non-favorite (Ctrl+Alt+K)
+d:: Send "^!k"

; Shift + X : Clear filters (F7)
+x:: Send "{F7}"

; Shift + F : Mark as favorite (Alt+Q)
+f:: Send "!q"

; Shift + U : Unmark as favorite (Unmark) (Alt+W)
+u:: Send "!w"

; Shift + E : Edit text (F4)
+e:: Send "{F4}"

; Shift + S : Save as file (Ctrl+S)
+s:: Send "^s"

; Shift + M : Merge clips
+m:: Send "^!j"

; Alt + 1 : Select current item
!1:: Send "{Enter}"

; Alt + 2 : Move down once and select
!2:: {
    Send "{Down}"
    Send "{Enter}"
}

; Alt + 3 : Move down twice and select
!3:: {
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Alt + 4 : Move down three times and select
!4:: {
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Alt + 5 : Move down four times and select
!5:: {
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Ctrl + 1–5 : Down N, F10, Select All, Copy (150ms between each step)
^1:: {
    Send "{F10}"
    Sleep 150
    Send "^a"
    Sleep 150
    Send "^c"
}
^2:: {
    Send "{Down 1}"
    Sleep 150
    Send "{F10}"
    Sleep 150
    Send "^a"
    Sleep 150
    Send "^c"
    Sleep 150
    Send "{F10}"
}
^3:: {
    Send "{Down 2}"
    Sleep 150
    Send "{F10}"
    Sleep 150
    Send "^a"
    Sleep 150
    Send "^c"
    Sleep 150
    Send "{F10}"
}
^4:: {
    Send "{Down 3}"
    Sleep 150
    Send "{F10}"
    Sleep 150
    Send "^a"
    Sleep 150
    Send "^c"
    Sleep 150
    Send "{F10}"
}
^5:: {
    Send "{Down 4}"
    Sleep 150
    Send "{F10}"
    Sleep 150
    Send "^a"
    Sleep 150
    Send "^c"
    Sleep 150
    Send "{F10}"
}

; =============================================================================
; ClipAngel Filter Selector Functions
; =============================================================================

; Navigate ClipAngel ComboBox to select a specific file type
NavigateClipAngelComboBox(typeIndex) {
    try {
        ; Ensure ClipAngel window is active
        win := WinExist("A")
        WinActivate("ahk_id " win)
        WinWaitActive("ahk_id " win, , 1)
        Sleep 50 ; Brief settle after activation

        root := UIA.ElementFromHandle(win)
        Sleep 50 ; Brief settle for UIA

        ; Find the file type filter ComboBox
        typeFilterCombo := root.FindFirst({ AutomationId: "TypeFilter", Type: 50003 })

        ; Fallback strategies
        if !typeFilterCombo {
            typeFilterCombo := root.FindFirst({ AutomationId: "TypeFilter", Type: "ComboBox" })
        }
        if !typeFilterCombo {
            typeFilterCombo := root.FindFirst({ Name: "all types", Type: 50003 })
        }
        if !typeFilterCombo {
            typeFilterCombo := root.FindFirst({ AutomationId: "TypeFilter" })
        }

        if !typeFilterCombo {
            return ; Silently fail if element not found
        }

        ; Get file type info first to show banner
        if (typeIndex < 0 || typeIndex >= g_ClipAngelFileTypes.Length) {
            return ; Invalid index
        }

        fileType := g_ClipAngelFileTypes[typeIndex + 1] ; +1 because arrays are 1-indexed

        ; Format display name for banner
        displayName := fileType.name
        if (displayName = "img") {
            displayName := "Image"
        } else if (displayName = "file") {
            displayName := "File"
        } else if (displayName = "text") {
            displayName := "Text"
        } else if (displayName = "*url") {
            displayName := "URL"
        } else if (displayName = "*filename") {
            displayName := "Filename"
        }

        ; Show banner notification
        ShowCenteredOverlay_Utils("📌 Selecting: " . displayName, 800, BANNER_ACCENT_INTERMEDIATE)

        ; Set focus and click to open dropdown
        try {
            typeFilterCombo.SetFocus()
            Sleep 30
        } catch {
            ; Continue if SetFocus fails
        }

        typeFilterCombo.Click()
        Sleep 150 ; Reduced delay after click

        ; Ensure focus is maintained
        try {
            typeFilterCombo.SetFocus()
            Sleep 100 ; Reduced delay after SetFocus
        } catch {
            typeFilterCombo.Click()
            Sleep 100 ; Reduced delay for retry
        }

        ; Reset to "all types" first - use Home key to physically anchor at top
        Send "{Home}"
        Sleep 200 ; Reduced delay after Home

        ; Navigate to target type
        if (fileType.navKey = "*") {
            ; For asterisk items, send asterisk multiple times
            loop fileType.navCount {
                Send "*"
                Sleep 100 ; Reduced delay between keystrokes
            }
        } else {
            ; For letter-based items, send the letter
            Send fileType.navKey
            Sleep 100 ; Reduced delay after navigation character
        }

        Sleep 100 ; Reduced pause after navigation

        ; Send Tab to confirm selection (NOT Enter)
        Send "{Tab}"
        Sleep 100 ; Reduced delay after TAB

        ; Return focus to clipboard list using Shift+T logic
        Send "{F10}"
        Sleep 100 ; Reduced delay after F10

        ; Send CTRL+HOME to force focus to very first item in sidebar list
        Send "^{Home}"
        Sleep 50 ; Reduced delay after CTRL+HOME
    } catch Error as e {
        ; Silently fail on error
    }
}

; Handler for character key press in filter selector
HandleClipAngelFilterChar(char) {
    global g_ClipAngelFilterSelectorActive, g_ClipAngelFilterCharMap

    ; Only process if selector is active
    if (!g_ClipAngelFilterSelectorActive) {
        return
    }

    ; Get file type info for this character
    fileTypeInfo := g_ClipAngelFilterCharMap.Get(char, "")
    if (fileTypeInfo = "") {
        ; Try lowercase if uppercase
        fileTypeInfo := g_ClipAngelFilterCharMap.Get(StrLower(char), "")
    }

    if (fileTypeInfo != "") {
        ; Cleanup selector first (closes GUI, disables hotkeys)
        CleanupClipAngelFilterSelector()

        ; Small delay to ensure ClipAngel window has focus
        Sleep 150

        ; Navigate to selected file type - use array index, not ComboBox index
        ; Find the array index by searching for matching fileType
        arrayIndex := -1
        for idx, ft in g_ClipAngelFileTypes {
            if (ft.name = fileTypeInfo.name) {
                arrayIndex := idx - 1 ; Convert to 0-based index
                break
            }
        }

        if (arrayIndex >= 0) {
            NavigateClipAngelComboBox(arrayIndex)
        }
    }
}

; Factory function to create a handler that properly captures the character
CreateClipAngelFilterCharHandler(char) {
    return (*) => HandleClipAngelFilterChar(char)
}

; Handler for Escape key
HandleClipAngelFilterEscape(*) {
    global g_ClipAngelFilterSelectorActive
    if (g_ClipAngelFilterSelectorActive) {
        CleanupClipAngelFilterSelector()
    }
}

; Cleanup function for ClipAngel filter selector
CleanupClipAngelFilterSelector() {
    global g_ClipAngelFilterSelectorActive, g_ClipAngelFilterSelectorGui, g_ClipAngelFilterHotkeyHandlers
    global g_ClipAngelFilterCharMap

    ; Disable active flag
    g_ClipAngelFilterSelectorActive := false

    ; Disable all character hotkeys
    for handler in g_ClipAngelFilterHotkeyHandlers {
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
    g_ClipAngelFilterHotkeyHandlers := []

    ; Close and destroy GUI
    if (IsObject(g_ClipAngelFilterSelectorGui)) {
        try {
            g_ClipAngelFilterSelectorGui.Destroy()
        } catch {
            ; Ignore
        }
        g_ClipAngelFilterSelectorGui := false
    }

    ; Clear char map
    g_ClipAngelFilterCharMap := Map()
}

; Show ClipAngel filter selector GUI
ShowClipAngelFilterSelector() {
    global g_ClipAngelFilterSelectorGui, g_ClipAngelFilterSelectorActive, g_ClipAngelFilterCharMap
    global g_ClipAngelFilterHotkeyHandlers, g_ClipAngelFileTypes, g_ClipAngelFilterCharSequence

    ; Close existing GUI if open
    if (g_ClipAngelFilterSelectorActive && IsObject(g_ClipAngelFilterSelectorGui)) {
        CleanupClipAngelFilterSelector()
        Sleep 50
    }

    ; Build character mapping
    g_ClipAngelFilterCharMap := Map()
    charIndex := 0
    for fileType in g_ClipAngelFileTypes {
        if (charIndex < g_ClipAngelFilterCharSequence.Length) {
            char := g_ClipAngelFilterCharSequence[charIndex + 1] ; +1 for 1-indexed array
            g_ClipAngelFilterCharMap[char] := fileType
            charIndex++
        }
    }

    if (g_ClipAngelFilterCharMap.Count = 0) {
        TrayTip("ClipAngel Filter Selector", "No file types found.", "IconX")
        SetTimer(() => TrayTip(), -5000)
        return
    }

    ; Get monitor dimensions
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
    g_ClipAngelFilterSelectorGui := Gui("+AlwaysOnTop +ToolWindow +E0x08000000", "ClipAngel File Type Filter")
    fontSize := (monitorHeight < 800) ? 9 : 10
    g_ClipAngelFilterSelectorGui.SetFont("s" . fontSize, "Segoe UI")
    g_ClipAngelFilterSelectorGui.MarginX := 10
    g_ClipAngelFilterSelectorGui.MarginY := 5

    ; Build display text
    displayText := "ClipAngel File Type Filter`n`n"
    charIndex := 0
    for fileType in g_ClipAngelFileTypes {
        if (charIndex < g_ClipAngelFilterCharSequence.Length) {
            char := g_ClipAngelFilterCharSequence[charIndex + 1]
            ; Format display name for better readability
            displayName := fileType.name
            if (displayName = "img") {
                displayName := "Image"
            } else if (displayName = "file") {
                displayName := "File"
            } else if (displayName = "text") {
                displayName := "Text"
            } else if (displayName = "*url") {
                displayName := "URL"
            } else if (displayName = "*filename") {
                displayName := "Filename"
            }
            displayText .= "[" . char . "] " . displayName . "`n"
            charIndex++
        }
    }

    ; Calculate GUI size
    baseWidth := 300
    textControlHeight := Min(400, (g_ClipAngelFileTypes.Length * 20) + 60)
    textControlWidth := baseWidth - 20

    ; Add text control with display
    g_ClipAngelFilterSelectorGui.AddEdit("w" . textControlWidth . " h" . textControlHeight . " ReadOnly VScroll",
        displayText)

    ; Add Close button
    closeBtn := g_ClipAngelFilterSelectorGui.AddButton("w100 Default Center", "Close")
    closeBtn.OnEvent("Click", (*) => CleanupClipAngelFilterSelector())

    ; Calculate total height
    totalHeight := 10 + textControlHeight + 40 + 10
    guiWidth := baseWidth

    ; Calculate center position
    marginX := 20
    marginY := 20
    guiX := monitorLeft + (monitorWidth - guiWidth) // 2
    guiY := monitorTop + (monitorHeight - totalHeight) // 2

    ; Ensure GUI stays within monitor bounds
    if (guiX < monitorLeft + marginX)
        guiX := monitorLeft + marginX
    if (guiY < monitorTop + marginY)
        guiY := monitorTop + marginY
    if (guiX + guiWidth > monitorLeft + monitorWidth - marginX)
        guiX := monitorLeft + monitorWidth - guiWidth - marginX
    if (guiY + totalHeight > monitorTop + monitorHeight - marginY)
        guiY := monitorTop + monitorHeight - totalHeight - marginY

    ; Show GUI centered on the active window's monitor
    g_ClipAngelFilterSelectorGui.Show("NA w" . guiWidth . " h" . totalHeight . " x" . guiX . " y" . guiY)

    ; Set active flag
    g_ClipAngelFilterSelectorActive := true

    ; Clear handlers array
    g_ClipAngelFilterHotkeyHandlers := []

    ; Enable hotkeys for all assigned characters
    charIndex := 0
    for fileType in g_ClipAngelFileTypes {
        if (charIndex < g_ClipAngelFilterCharSequence.Length) {
            char := g_ClipAngelFilterCharSequence[charIndex + 1]

            ; Create handler
            handler := CreateClipAngelFilterCharHandler(char)

            ; Store handler for cleanup
            g_ClipAngelFilterHotkeyHandlers.Push({ char: char, handler: handler })

            ; Enable hotkey (both uppercase and lowercase)
            try {
                if (char = ",") {
                    Hotkey("vkBC", handler, "On")
                } else if (char = ".") {
                    Hotkey("vkBE", handler, "On")
                } else {
                    Hotkey(char, handler, "On")
                    if (RegExMatch(char, "^[a-z]$")) {
                        Hotkey(StrUpper(char), handler, "On")
                    }
                }
            } catch {
                ; Silently ignore if we can't create hotkey
            }

            charIndex++
        }
    }

    ; Enable Escape hotkey
    Hotkey("Escape", HandleClipAngelFilterEscape, "On")
}

; Shift + Y : Open file type filter selector (Quick Wizard)
+y:: {
    ; Only show selector if ClipAngel is active
    if WinActive("ClipAngel") {
        ShowClipAngelFilterSelector()
    }
}

#HotIf

;-------------------------------------------------------------------
; WhatsApp Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("WhatsApp")

global isRecording := false          ; persists between hotkey presses

; Shift + V : Toggle voice message - Voice
+v:: ToggleVoiceMessage()

; Shift + S : Search chats - Search
+s:: Send("!k")

; Shift + R : Reply - Reply
+r:: Send("!r")

; Shift + E : Emoji panel - Emoji
+e:: Send("^!s")

; Shift + U : Toggle Unread filter - Unread
+u::
{
    try
    {
        uia := UIA_Browser()
        Sleep 300 ; Give UIA time to attach

        ; Find the "Unread" and "All" filter buttons
        unreadButton := uia.FindElement({ Name: "Unread", AutomationId: "unread-filter", Type: "TabItem" })
        allButton := uia.FindElement({ Name: "All", AutomationId: "all-filter", Type: "TabItem" })

        if (unreadButton && allButton) {
            ; Check if the "Unread" button is currently selected.
            ; The .IsSelected property is part of the SelectionItemPattern.
            if (unreadButton.IsSelected) {
                allButton.Click() ; If Unread is selected, click All
            }
            else {
                unreadButton.Click() ; Otherwise, click Unread
            }
        }
        else if (unreadButton) {
            ; Fallback if only the Unread button is found
            unreadButton.Click()
        }
        else {
            MsgBox "Could not find the 'Unread' filter button."
        }
    }
    catch Error as e {
        MsgBox "An error occurred: " e.Message
    }
}

; Shift + F : Focus current chat - Focus
+f::
{
    try
    {
        ; WhatsApp desktop is Chromium-based, so we can use UIA_Browser.
        ; It should attach to the active window, which is WhatsApp thanks to #HotIf.
        uia := UIA_Browser()
        Sleep 300 ; Give UIA time to attach to the browser. A similar delay is in the reference script.

        ; Find the "Archived" button to use as an anchor.
        ; The user provided: Name:"Archived "
        archivedButton := uia.FindElement({ Name: "Archived ", Type: "Button" })

        if (archivedButton) {
            ; Focus the button without clicking it.
            archivedButton.SetFocus()
            ; Send Tab to move to the main conversation list.
            ; From there, the focus should be on the selected chat.
            SendInput "{Tab}"
        }
        else {
            MsgBox "Could not find the 'Archived' button."
        }
    }
    catch Error as e {
        MsgBox "An error occurred while trying to focus WhatsApp conversation: " e.Message
    }
}

; Shift + M : Mark as read or unread - Mark
+m:: Send "^!+u"

; Shift + P : Pin chat or unpin chat - Pin
+p:: Send "^!+p"

; ---------------------------------------------------------------------------
ToggleVoiceMessage() {
    global isRecording

    try {
        chrome := UIA_Browser()      ; top-level Chrome UIA element
        if !IsObject(chrome) {
            MsgBox "Can't attach to Chrome."
            return
        }

        Sleep 100                    ; reduced from 400ms - let Chrome finish drawing

        ; Exact-name regexes (case-insensitive, anchored ^ $)
        voicePattern := "i)^(Voice message|Record voice message)$"
        sendPattern := "i)^(Send|Stop recording)$"

        ; Helper to grab a button by pattern
        ; Use longer timeout (3000ms) for voice message button to allow WhatsApp UI to restore
        FindBtn(p) => WaitForButton(chrome, p, 3000)

        if (isRecording) {           ; â–º we're supposed to stop & send
            if (btn := FindBtn(sendPattern)) {
                ; Determine if this button supports Invoke
                supportsInvoke := false
                try {
                    supportsInvoke := btn.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)
                } catch {
                    supportsInvoke := false
                }

                ; Try multi-strategy activation: prefer Invoke when available, fallback to Click
                clicked := false
                if (supportsInvoke) {
                    try {
                        btn.Invoke()
                        clicked := true
                    } catch {
                    }
                }
                if (!clicked) {
                    try {
                        btn.Click()
                        clicked := true
                    } catch {
                    }
                }

                isRecording := false
                ; Give WhatsApp time to restore the UI after sending
                Sleep 300
            } else {
                ; Assume you clicked Send manually > reset & start new rec
                isRecording := false
                if (btn := FindBtn(voicePattern)) {
                    btn.Click()
                    isRecording := true
                } else
                    MsgBox "Couldn't restart recording (Voice-message button missing)."
            }
        } else {                     ; â–º start recording
            if (btn := FindBtn(voicePattern)) {
                btn.Click()
                isRecording := true
            } else
                MsgBox "Couldn't find the Voice-message button."
        }
    } catch Error as err {
        MsgBox "Error:`n" err.Message
    }
}

; ---------------------------------------------------------------------------
FocusSourceControlViewForCommitGeneration(delayMs := 450) {
    ; Ensure Source Control has time to render commit actions before UIA lookup.
    Send "+d"
    Sleep delayMs
}

; ---------------------------------------------------------------------------
ClickGenerateCommitMessageButton() {
    try {
        ; Ensure Source Control view is focused so the Generate button is visible.
        FocusSourceControlViewForCommitGeneration()

        ; Use UIA_Browser to get the root element (similar to other functions in the script)
        uia := UIA_Browser()
        if !IsObject(uia) {
            ; Fallback: try Ctrl+M shortcut if UIA fails
            Send "^m"
            return true
        }

        ; Find the "Generate Commit Message (Ctrl+M)" button
        ; Try multiple search strategies
        btn := uia.FindFirst({ Name: "Generate Commit Message (Ctrl+M)", ControlType: "Button" })

        ; If not found by exact name, try partial match
        if !btn {
            btn := uia.FindFirst({ Name: "Generate Commit Message", ControlType: "Button" })
        }

        ; If still not found, try by ControlType only (Type: 50000 = Button)
        if !btn {
            ; Get all buttons and find the one with the right name
            buttons := uia.FindAll({ ControlType: "Button" })
            for button in buttons {
                if InStr(button.Name, "Generate Commit Message") {
                    btn := button
                    break
                }
            }
        }

        if btn {
            btn.Click()
            return true
        } else {
            ; Fallback: try Ctrl+M shortcut
            Send "^m"
            return true
        }
    }
    catch Error as e {
        ; Fallback: try Ctrl+M shortcut if UIA fails
        Send "^m"
        return true
    }
}

; ---------------------------------------------------------------------------
; WaitForButton(root, pattern, timeout := 5000)
;   â€¢ Searches all descendant buttons of `root` until Name matches `pattern`
;   â€¢ Returns the UIA element or 0 if none matched within `timeout` ms
; ---------------------------------------------------------------------------
WaitForButton(root, pattern, timeout := 5000) {
    ; #region agent log
    SafeDebugLog Format(
        "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:1727`",`"message`":`"WaitForButton entry`",`"data`":{`"pattern`":`"{4}`",`"timeout`":{5}},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"B`"}`n",
        A_TickCount, Random(1000, 9999), A_TickCount, pattern, timeout)
    ; #endregion
    if !IsObject(root)
        return 0

    deadline := A_TickCount + timeout
    while (A_TickCount < deadline) {
        buttons := root.FindAll({ Type: "Button" })
        ; #region agent log
        SafeDebugLog Format(
            "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:1733`",`"message`":`"Found buttons count`",`"data`":{`"count`":{4}},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"B`"}`n",
            A_TickCount, Random(1000, 9999), A_TickCount, buttons.Length)
        ; #endregion

        ; Collect all matching buttons and their properties
        matchingButtons := []
        for btn in buttons {
            btnName := ""
            try btnName := btn.Name
            ; #region agent log
            if InStr(pattern, "Connect") {
                SafeDebugLog Format(
                    "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:1738`",`"message`":`"Checking button name`",`"data`":{`"name`":`"{4}`",`"pattern`":`"{5}`"},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"B`"}`n",
                    A_TickCount, Random(1000, 9999), A_TickCount, btnName, pattern)
            }
            ; #endregion
            if RegExMatch(btn.Name, pattern) {
                className := ""
                hasClassName := false
                supportsInvoke := false
                parentName := ""
                parentClass := ""

                try {
                    className := btn.ClassName
                    hasClassName := (className != "")
                } catch {
                    ; ClassName property not available or error reading
                    hasClassName := false
                }

                ; Try to capture parent info for better disambiguation (esp. duplicated "Send" buttons)
                try {
                    parent := btn.Parent
                    parentName := parent.Name
                    parentClass := parent.ClassName
                } catch {
                    parentName := ""
                    parentClass := ""
                }

                ; Check if button supports Invoke pattern
                try {
                    supportsInvoke := btn.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)
                } catch {
                    supportsInvoke := false
                }

                ; #region agent log
                SafeDebugLog Format(
                    "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:1770`",`"message`":`"Matching button found`",`"data`":{`"name`":`"{4}`",`"className`":`"{5}`",`"hasClassName`":{6},`"supportsInvoke`":{7}},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"B,C`"}`n",
                    A_TickCount, Random(1000, 9999), A_TickCount, btnName, className, hasClassName, supportsInvoke)
                ; #endregion
                matchingButtons.Push({ btn: btn, hasClassName: hasClassName, className: className, supportsInvoke: supportsInvoke,
                    parentName: parentName, parentClass: parentClass })
            }
        }

        ; If we found matching buttons, prioritize: 1) hasClassName (the actual clickable button), 2) supportsInvoke, 3) first
        ; Note: In WhatsApp, the button WITH ClassName is the actual clickable one, even if it doesn't support Invoke pattern
        if matchingButtons.Length > 0 {
            bestBtn := ""
            bestClassName := ""
            bestScore := 0

            for match in matchingButtons {
                score := 0
                if match.hasClassName && match.className != "" {
                    score += 10  ; Highest priority: has ClassName (the actual clickable button in WhatsApp)
                }
                if match.supportsInvoke {
                    score += 5   ; Second priority: supports Invoke pattern
                }

                ; Additional heuristic for WhatsApp voice "Send" vs text "Send" (H8)
                ; When using the sendPattern, prefer the inner child button whose parent is also "Send"
                if InStr(pattern, "Send|Stop recording") {
                    try {
                        if (match.parentName = "Send") {
                            score += 3
                        }
                    }
                }

                if (score > bestScore) {
                    bestBtn := match.btn
                    bestClassName := match.className
                    bestScore := score
                }
            }

            ; If no button scored (shouldn't happen), use the first one
            if !bestBtn {
                bestBtn := matchingButtons[1].btn
            }

            ; #region agent log
            finalBtnName := "", finalBtnClassName := "", finalBtnType := ""
            try finalBtnName := bestBtn.Name
            try finalBtnClassName := bestBtn.ClassName
            try finalBtnType := bestBtn.ControlType
            SafeDebugLog Format(
                "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:1813`",`"message`":`"WaitForButton returning button`",`"data`":{`"name`":`"{4}`",`"className`":`"{5}`",`"type`":`"{6}`",`"score`":{7}},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"C`"}`n",
                A_TickCount, Random(1000, 9999), A_TickCount, finalBtnName, finalBtnClassName, finalBtnType, bestScore)
            ; #endregion
            return bestBtn
        }
        Sleep 50  ; reduced from 150ms to 50ms for faster retries
    }

    ; #region agent log
    SafeDebugLog Format(
        "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:1818`",`"message`":`"WaitForButton timeout - no button found`",`"data`":{`"pattern`":`"{4}`"},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"B`"}`n",
        A_TickCount, Random(1000, 9999), A_TickCount, pattern)
    ; #endregion
    return 0
}

#HotIf

;-------------------------------------------------------------------
; Outlook Reminder Window Shortcuts
;-------------------------------------------------------------------
#HotIf (WinActive("ahk_exe OUTLOOK.EXE") || WinActive("ahk_exe olk.exe")) && RegExMatch(WinGetTitle("A"),
"i)Reminders?") && !IsFileDialogActive()

; ativa a janela de lembretes do Outlook
ActivateReminder() {
    if (!WinExist("ahk_exe OUTLOOK.EXE")) {
        ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
        return
    }
    WinActivate("ahk_exe OUTLOOK.EXE")
    WinWaitActive("ahk_exe OUTLOOK.EXE", , 1)
}

; digita o tempo e aperta Alt+S
QuickSnooze(t) {
    ActivateReminder()
    Send("{Tab}")              ; chega ao combo
    Send("{Tab}")              ; chega ao combo
    Sleep 100
    Send("^a{Delete}" . t)     ; substitui o texto
    Sleep 120
    Send("!s")                 ; Alt+S = Snooze
    Sleep 200
    Send("{Tab}")
    Send("{Tab}")
    Send("{Tab}")
}

; caixa de confirmaÃ§Ã£o antes de executar
Confirm(t) {
    if MsgBox("Snooze for " t "?", "Confirm Snooze", "YesNo Icon?") = "Yes"
        QuickSnooze(t)
}

; ---------------------------------------------------------------------------
; New Outlook Reminders (keyboard-only)
; - Uses UIA list extraction + Standard Information Display selection banner
; - Executes item actions via Apps/Menu key + arrow navigation (per screenshots)
; ---------------------------------------------------------------------------
global g_RemindersPickKey := ""
global g_DebugBe11ecLogPath := "C:\Users\fie7ca\Documents\scripts\debug-be11ec.log"
global g_RemindersDebugEnabled := false
global g_RemindersTimingProfile := "very_slow"
global g_RemindersTimingProfiles := Map(
    "normal", Map(
        "context_focus_ms", 80,
        "focus_to_menu_ms", 25,
        "menu_render_ms", 45,
        "menu_home_ms", 60,
        "menu_scan_step_ms", 50,
        "submenu_open_ms", 80,
        "post_select_ms", 60,
        "modal_to_row_ms", 35
    ),
    "slow", Map(
        "context_focus_ms", 120,
        "focus_to_menu_ms", 70,
        "menu_render_ms", 170,
        "menu_home_ms", 100,
        "menu_scan_step_ms", 140,
        "submenu_open_ms", 200,
        "post_select_ms", 110,
        "modal_to_row_ms", 80
    ),
    "very_slow", Map(
        "context_focus_ms", 260,
        "focus_to_menu_ms", 180,
        "menu_render_ms", 420,
        "menu_home_ms", 220,
        "menu_scan_step_ms", 360,
        "submenu_open_ms", 520,
        "post_select_ms", 260,
        "modal_to_row_ms", 220
    )
)

Reminders_GetTimingProfile() {
    global g_RemindersTimingProfile, g_RemindersTimingProfiles
    p := "very_slow"
    try p := StrLower(Trim(g_RemindersTimingProfile))
    if !g_RemindersTimingProfiles.Has(p)
        p := "very_slow"
    return g_RemindersTimingProfiles[p]
}

Reminders_DelayValue(key, fallbackMs := 80) {
    try {
        profile := Reminders_GetTimingProfile()
        if profile.Has(key)
            return profile[key]
    } catch {
    }
    return fallbackMs
}

Reminders_LoadingShow(text) {
    try {
        StandardLoadingBar_Show(text, BANNER_ACCENT_INTERMEDIATE, { passive: false, centerOnHwnd: 0, textWidth: 560,
            fontSize: 17 })
    } catch {
    }
}

Reminders_LoadingHide(delayMs := 0) {
    try StandardLoadingBar_Hide(delayMs)
    catch {
    }
}

Reminders_DebugLog(location, message, data := "", hypothesisId := "A", runId := "pre-fix") {
    if !g_RemindersDebugEnabled
        return
    try {
        Debug_Escape(s) => StrReplace(StrReplace(StrReplace(String(s), "\", "\\"), "`"", "\`""), "`n", "\n")
        Debug_MapToJson(m) {
            out := ""
            for k, v in m {
                if (out != "")
                    out .= ","
                out .= "`"" Debug_Escape(k) "`":"
                if (v is Integer || v is Float)
                    out .= String(v)
                else
                    out .= "`"" Debug_Escape(v) "`""
            }
            return "{" out "}"
        }

        payload := Map()
        payload["sessionId"] := "be11ec"
        payload["id"] := "log_" A_TickCount "_" Random(1000, 9999)
        payload["timestamp"] := A_TickCount
        payload["location"] := location
        payload["message"] := message
        payload["runId"] := runId
        payload["hypothesisId"] := hypothesisId

        d := IsObject(data) ? data : Map("value", data)
        ; Encode only scalar data safely (stringify non-numeric as strings)
        payloadJson := Debug_MapToJson(payload)
        dataJson := Debug_MapToJson(d)
        line := SubStr(payloadJson, 1, StrLen(payloadJson) - 1) . ",`"data`":" . dataJson . "}"
        global g_DebugBe11ecLogPath
        FileAppend(line "`n", g_DebugBe11ecLogPath, "UTF-8")
    } catch {
    }
}

; #region agent log (session 6dacac — NDJSON to debug-6dacac.log; remove after verification)
Debug6dacac_Log(location, message, data := "", hypothesisId := "H1", runId := "post-fix") {
    if !g_RemindersDebugEnabled
        return
    try {
        Esc(s) => StrReplace(StrReplace(StrReplace(String(s), "\", "\\"), "`"", "\`""), "`n", "\n")
        MapToJson(m) {
            out := ""
            for k, v in m {
                if (out != "")
                    out .= ","
                out .= "`"" Esc(k) "`":"
                if (v is Integer || v is Float)
                    out .= String(v)
                else
                    out .= "`"" Esc(v) "`""
            }
            return "{" out "}"
        }
        d := IsObject(data) ? data : Map("value", data)
        id := "log_" A_TickCount "_" Random(1000, 9999)
        ts := A_TickCount
        payload := Map()
        payload["sessionId"] := "6dacac"
        payload["id"] := id
        payload["timestamp"] := ts
        payload["location"] := location
        payload["message"] := message
        payload["runId"] := runId
        payload["hypothesisId"] := hypothesisId
        payloadJson := MapToJson(payload)
        dataJson := MapToJson(d)
        line := SubStr(payloadJson, 1, StrLen(payloadJson) - 1) . ",`"data`":" . dataJson . "}`n"
        FileAppend(line, A_ScriptDir "\debug-6dacac.log", "UTF-8")
    } catch {
    }
}

; #endregion

; Diagnostic-only raw UIA dump for reminder window.
; Captures candidate counts and samples by control type so we can patch selectors based on facts.
Reminders_DebugDumpWindowTree(hwnd, tag := "") {
    if !g_RemindersDebugEnabled
        return
    if !hwnd
        return
    try {
        rootSource := ""
        root := Reminders_RootElementForHwnd(hwnd, &rootSource)
        if !root {
            try Reminders_DebugLog("Shift keys.ahk:Reminders_DebugDumpWindowTree", "Root not found", Map(
                "hwnd", hwnd,
                "tag", tag,
                "rootSource", rootSource
            ), "TREE0", "pre-fix")
            return
        }

        DumpByType(typeName, hyp) {
            arr := 0
            try arr := root.FindAll({ ControlType: typeName })
            cnt := arr ? arr.Length : 0
            sample := ""
            if arr {
                maxSample := Min(25, arr.Length)
                loop maxSample {
                    i := A_Index
                    n := ""
                    try n := arr[i].Name
                    if (n = "")
                        continue
                    if (sample != "")
                        sample .= " | "
                    sample .= n
                }
            }
            try Reminders_DebugLog("Shift keys.ahk:Reminders_DebugDumpWindowTree", "Tree slice", Map(
                "hwnd", hwnd,
                "tag", tag,
                "rootSource", rootSource,
                "type", typeName,
                "count", cnt,
                "sample", sample
            ), hyp, "pre-fix")
        }

        ; Types most likely to contain New Outlook reminder rows.
        DumpByType("Group", "TREE-G")
        DumpByType("Button", "TREE-B")
        DumpByType("List", "TREE-L")
        DumpByType("ListItem", "TREE-LI")
        DumpByType("DataItem", "TREE-DI")
        DumpByType("Text", "TREE-T")
        DumpByType("Hyperlink", "TREE-H")
    } catch {
    }
}

Reminders_IsNewOutlookWindow() {
    try {
        ; Reminders window can run under classic OUTLOOK.EXE or Store olk.exe.
        if !(WinActive("ahk_exe OUTLOOK.EXE") || WinActive("ahk_exe olk.exe"))
            return false
        t := WinGetTitle("A")
        ; NOTE: single backslash in regex. Using \\b would match literal "\b".
        ok := RegExMatch(t, "i)\bReminders\b")
        ; #region agent log
        try Reminders_DebugLog("Shift keys.ahk:Reminders_IsNewOutlookWindow", "Computed isNewReminders", Map(
            "ok", ok,
            "title", t
        ), "H1", "pre-fix")
        ; #endregion
        return ok
    } catch {
        return false
    }
}

Reminders_ItemsListSignature(items, maxLabels := 0) {
    sig := String(items.Length)
    n := (maxLabels > 0) ? Min(items.Length, maxLabels) : items.Length
    loop n {
        sig .= "`n" items[A_Index].label
    }
    return sig
}

; Tiny move + restore so Outlook refreshes the accessible tree (same effect as manually moving the window).
; Skip when maximized — WinMove is unreliable; user can restore the window first if needed.
Reminders_NudgeWindowForUiRefresh(hwnd) {
    if !hwnd || !WinExist("ahk_id " hwnd)
        return
    try {
        if (WinGetMinMax("ahk_id " hwnd) = 1)
            return
        WinGetPos(&x, &y, &w, &h, "ahk_id " hwnd)
        WinMove(x + 1, y, w, h, "ahk_id " hwnd)
        Sleep 20  ; brief beat so the shell/UIA sees the move before restore
        WinMove(x, y, w, h, "ahk_id " hwnd)
    } catch {
    }
}

; Two consecutive identical UIA snapshots (or maxPasses) — reduces races while Outlook refreshes the list.
; NOTE: If we get two consecutive EMPTY matches on early passes, keep trying instead of giving up,
; since a newly-visible window might just have UIA that hasn't rendered reminders yet.
Reminders_GetItemsStable(targetHwnd, delayMs := 60, maxPasses := 3) {
    lastSig := ""
    lastItems := []
    loop maxPasses {
        cur := Reminders_GetItems(targetHwnd)
        sig := Reminders_ItemsListSignature(cur, 0)
        if (lastSig != "" && sig = lastSig) {
            ; Don't accept empty-empty match as stable on early passes (pass 1-2); window might still be rendering.
            ; Only consider pass 3+ empty match as "stable" (window legitimately empty).
            emptyMatch := (cur.Length = 0 && lastItems.Length = 0)
            if (!emptyMatch || A_Index >= maxPasses) {
                try Reminders_DebugLog("Shift keys.ahk:Reminders_GetItemsStable", "Stable snapshot matched", Map(
                    "pass", A_Index,
                    "count", cur.Length,
                    "maxPasses", maxPasses,
                    "emptyMatch", emptyMatch ? 1 : 0,
                    "remHwnd", targetHwnd
                ), "ST1", "pre-fix")
                return cur
            } else {
                try Reminders_DebugLog("Shift keys.ahk:Reminders_GetItemsStable",
                    "Empty-empty match ignored, continuing retry", Map(
                        "pass", A_Index,
                        "maxPasses", maxPasses,
                        "remHwnd", targetHwnd
                    ), "ST-SKIP", "pre-fix")
            }
        }
        lastSig := sig
        lastItems := cur
        Sleep delayMs
    }
    try Reminders_DebugLog("Shift keys.ahk:Reminders_GetItemsStable", "Stable snapshot max passes; using last read",
        Map(
            "count", lastItems.Length,
            "remHwnd", targetHwnd
        ), "ST2", "pre-fix")
    return lastItems
}

Reminders_RootElementForHwnd(hwnd, &source := "") {
    source := ""
    if !hwnd
        return ""
    ; New Outlook surfaces often live under WebView2/Chromium subtree.
    try {
        el := UIA.ElementFromChromium("ahk_id " hwnd, 500)
        if el {
            source := "chromium"
            return el
        }
    } catch {
    }
    try {
        el := UIA.ElementFromHandle(hwnd)
        if el {
            source := "handle"
            return el
        }
    } catch {
    }
    return ""
}

Reminders_GetItems(targetHwnd := 0) {
    items := []
    hwnd := targetHwnd ? targetHwnd : WinExist("A")
    if !hwnd
        return items
    try {
        rootSource := ""
        root := Reminders_RootElementForHwnd(hwnd, &rootSource)
        if !root {
            try Reminders_DebugLog("Shift keys.ahk:Reminders_GetItems", "ERROR: Root element null", Map(
                "hwnd", hwnd,
                "rootSource", rootSource
            ), "X0-ERR", "pre-fix")
            return items
        }

        CollectFromCandidates(cands) {
            out := []
            dropped := 0
            dropSample := ""
            for b in cands {
                n := ""
                try n := b.Name
                if (n = "")
                    continue
                ; Exclude global/window controls
                if (n = "Settings" || n = "Dismiss all" || n = "Dismiss All" || n = "Minimize" || n = "Maximize" || n =
                    "Close")
                    continue
                ; Exclude action/menu-like items that can appear while context UI is open
                if RegExMatch(n, "i)^(Snooze reminder|Dismiss reminder|Join Teams meeting|Chat with participants)$")
                    continue

                ; Keep only actual reminder rows. In our UIA tree, these names include time/all-day + a relative marker.
                ; Examples: "Stretch All day Today", "CIM Journey 3:00 PM Microsoft Teams Meeting 18 hrs ago"
                ; Relative-age tokens vary (e.g. "1 hour ago", "7 days", "4 wks ago", "Today").
                ; NOTE: single backslash in regex. Using \\b would match literal "\b".
                hasTime := RegExMatch(n, "i)(\bAll day\b|\bAM\b|\bPM\b)")
                isAllDay := RegExMatch(n, "i)\bAll day\b")
                hasRel := RegExMatch(n,
                    "i)\b(Today|\d+\s*(min|mins|minute|minutes|hr|hrs|hour|hours|day|days|wk|wks|week|weeks)\b(\s+ago)?)\b"
                )
                hasClockTime := RegExMatch(n, "i)\b\d{1,2}:\d{2}\b")
                hasEventCue := RegExMatch(n,
                    "i)\b(Meeting|Teams|Appointment|Event|Call|Invite|Reminder)\b")
                ; New reminders can be "in 15 minutes" without AM/PM; keep broader candidates
                ; while still excluding known command/window controls above.
                isLikelyRow := ((hasTime || hasClockTime || hasRel || hasEventCue) && !RegExMatch(n,
                    "i)^(Snooze reminder|Dismiss reminder|Join Teams meeting|Chat with participants)$"))
                if !isLikelyRow {
                    dropped++
                    if (dropSample != "" && dropped <= 8)
                        dropSample .= " | "
                    if (dropped <= 8)
                        dropSample .= n
                    continue
                }
                out.Push({ el: b, label: n })
            }
            return Map("items", out, "dropped", dropped, "dropSample", dropSample, "totalCandidates", cands ? cands.Length :
                0)
        }

        ; Primary anchor: reminder summary group text varies ("There are ..." / "There is only ...").
        listGroup := root.FindFirst({ ControlType: "Group", Name: "reminder", matchmode: "Substring" })
        if !listGroup
            listGroup := root.FindFirst({ ControlType: "Group", Name: "There ", matchmode: "Substring" })
        if !listGroup
            listGroup := root.FindFirst({ Name: "reminder", matchmode: "Substring" })

        if !listGroup {
            try Reminders_DebugLog("Shift keys.ahk:Reminders_GetItems",
                "WARNING: 'There are X reminders' group not found in UIA", Map(
                    "hwnd", hwnd,
                    "rootSource", rootSource
                ), "X0-WRN", "pre-fix")
        }

        ; Always re-evaluate source: a short-lived cache reused "listGroup" without re-checking
        ; listSmallerThanRoot and could omit reminders (same class of bug as scanning only the group).
        listBtns := 0
        listBtnLen := 0
        if listGroup {
            try listBtns := listGroup.FindAll({ ControlType: "Button" })
            listBtnLen := listBtns ? listBtns.Length : 0
        }
        rootBtns := root.FindAll({ ControlType: "Button" })
        rootBtnLen := rootBtns ? rootBtns.Length : 0
        listSmallerThanRoot := (listGroup && listBtnLen > 0 && rootBtnLen > listBtnLen)

        ; Diagnostic logging for button search
        try Reminders_DebugLog("Shift keys.ahk:Reminders_GetItems", "Button search results", Map(
            "rootBtnLen", rootBtnLen,
            "listBtnLen", listBtnLen,
            "listGroupFound", listGroup ? 1 : 0,
            "hwnd", hwnd,
            "rootSource", rootSource
        ), "GETITEMS-BTN", "pre-fix")

        btns := 0
        source := ""
        if (listGroup && listBtnLen > 0 && !listSmallerThanRoot) {
            btns := listBtns
            source := "listGroup"
        } else {
            btns := rootBtns
            source := "root"
        }

        ; Debug: log if btns is empty before attempting to collect
        if (!btns || btns.Length = 0) {
            try Reminders_DebugLog("Shift keys.ahk:Reminders_GetItems", "No buttons found to process", Map(
                "source", source,
                "rootBtnLen", rootBtnLen,
                "listBtnLen", listBtnLen,
                "hwnd", hwnd,
                "rootSource", rootSource
            ), "GETITEMS-NOBTNS", "pre-fix")
        }

        r := CollectFromCandidates(btns)
        items := r["items"]
        dropped := r["dropped"]
        dropSample := r["dropSample"]

        ; Fallback for New Outlook builds where reminder rows are exposed as ListItem/DataItem, not Button.
        if (items.Length = 0) {
            listItems := 0
            dataItems := 0
            try listItems := root.FindAll({ ControlType: "ListItem" })
            try dataItems := root.FindAll({ ControlType: "DataItem" })
            liLen := listItems ? listItems.Length : 0
            diLen := dataItems ? dataItems.Length : 0
            try Reminders_DebugLog("Shift keys.ahk:Reminders_GetItems", "Fallback candidate search", Map(
                "listItemLen", liLen,
                "dataItemLen", diLen,
                "hwnd", hwnd,
                "rootSource", rootSource
            ), "GETITEMS-FB", "pre-fix")

            if (liLen > 0) {
                r := CollectFromCandidates(listItems)
                if (r["items"].Length > 0) {
                    items := r["items"]
                    dropped := r["dropped"]
                    dropSample := r["dropSample"]
                    source := "fallback-listitem"
                }
            }
            if (items.Length = 0 && diLen > 0) {
                r := CollectFromCandidates(dataItems)
                if (r["items"].Length > 0) {
                    items := r["items"]
                    dropped := r["dropped"]
                    dropSample := r["dropSample"]
                    source := "fallback-dataitem"
                }
            }
        }
        ; #region agent log (session 6dacac — H1 source + row count)
        try {
            Debug6dacac_Log("Shift keys.ahk:Reminders_GetItems", "source chosen", Map(
                "source", source,
                "keptCount", items.Length,
                "rootBtnLen", rootBtnLen,
                "listBtnLen", listBtnLen,
                "listSmallerThanRoot", listSmallerThanRoot ? 1 : 0,
                "remHwnd", hwnd,
                "rootSource", rootSource
            ), "H1", "post-fix")
        } catch {
        }
        ; #endregion
        ; #region agent log
        try {
            sample := ""
            maxSample := Min(12, items.Length)
            loop maxSample {
                i := A_Index
                if (sample != "")
                    sample .= " | "
                sample .= items[i].label
            }
            Reminders_DebugLog("Shift keys.ahk:Reminders_GetItems", "Extracted reminders sample", Map(
                "count", items.Length,
                "sample", sample,
                "source", source,
                "totalCandidates", r["totalCandidates"],
                "dropped", dropped,
                "dropSample", dropSample,
                "hwnd", hwnd,
                "rootSource", rootSource
            ), "X1", "pre-fix")
        } catch {
        }
        ; #endregion
    } catch {
        return items
    }
    return items
}

Reminders_PickKey(key) {
    global g_RemindersPickKey
    ; #region agent log
    try Reminders_DebugLog("Shift keys.ahk:Reminders_PickKey", "PickKey called", Map(
        "key", key,
        "priorKey", A_PriorKey,
        "thisHotkey", A_ThisHotkey
    ), "P1", "pre-fix")
    ; #endregion

    ; Guard: ignore accidental selection when Windows key (or other modifiers) is involved.
    ; This prevents the selection modal from disappearing due to unrelated Win-key chords.
    try {
        if (GetKeyState("LWin", "P") || GetKeyState("RWin", "P")
        || GetKeyState("Ctrl", "P") || GetKeyState("Alt", "P")) {
            ; #region agent log
            try Reminders_DebugLog("Shift keys.ahk:Reminders_PickKey", "Ignored pick due to modifier down", Map(
                "key", key,
                "lwin", GetKeyState("LWin", "P"),
                "rwin", GetKeyState("RWin", "P"),
                "ctrl", GetKeyState("Ctrl", "P"),
                "alt", GetKeyState("Alt", "P")
            ), "P3", "pre-fix")
            ; #endregion
            return
        }
        if (A_PriorKey = "LWin" || A_PriorKey = "RWin") {
            ; #region agent log
            try Reminders_DebugLog("Shift keys.ahk:Reminders_PickKey", "Ignored pick due to priorKey Win", Map(
                "key", key,
                "priorKey", A_PriorKey
            ), "P4", "pre-fix")
            ; #endregion
            return
        }
    } catch {
    }

    g_RemindersPickKey := key
    try StandardLoadingBar_CloseKeysOverlay()
    try StandardLoadingBar_Hide(0)
}

Reminders_PickTimeout() {
    ; #region agent log
    try Reminders_DebugLog("Shift keys.ahk:Reminders_PickTimeout", "Timeout fired", Map(), "P2", "pre-fix")
    ; #endregion
    Reminders_PickKey("TIMEOUT")
}

Reminders_SelectItem(actionLabel, &items, remHwnd, maxItems := 35) {
    global g_RemindersPickKey
    g_RemindersPickKey := ""

    ; Workaround: one-pixel nudge refreshes UIA before enumeration (not repeated in the modal refresh loop).
    Reminders_NudgeWindowForUiRefresh(remHwnd)
    Sleep 100  ; longer wait for reminder window to stabilize UIA after nudge (highly volatile window)

    ; Volatile window: refresh once right before showing the modal so we don't start with a stale/partial snapshot.
    ; Use longer delays (100ms between passes) to wait for Outlook to stabilize the reminder list in UIA.
    try items := Reminders_GetItemsStable(remHwnd, 100, 5)

    if (items.Length = 0) {
        ; One more try with even longer wait (UIA can lag when reminders are being added rapidly).
        try items := Reminders_GetItemsStable(remHwnd, 150, 6)
        if (items.Length = 0) {
            ShowCenteredOverlay_Utils("❌ No reminders found", 1600, BANNER_ACCENT_ERROR)
            return 0
        }
    }

    ; Stable key set (we'll keep callbacks stable and just refresh the displayed list).
    keys := []
    loop 9
        keys.Push(A_Index)
    for c in StrSplit("ABCDEFGHIJKLMNOPQRSTUVWXYZ")
        keys.Push(c)

    ClampCount(itemsLen) {
        c := itemsLen
        if (c > maxItems)
            c := maxItems
        if (c > keys.Length)
            c := keys.Length
        return c
    }

    BuildMsg(currentItems, currentCount) {
        ; Information-only copy inside the interactive banner (see docs/standard_information_display.md).
        m :=
            "ℹ️ The script nudges this window by one pixel before listing (Outlook UIA quirk). If a row is still missing, focus or move the window, then use the shortcut again.`n`n"
        m .= "❓ Select reminder to " actionLabel ":`n`n"
        loop currentCount {
            i := A_Index
            k := keys[i]
            label := currentItems[i].label
            m .= k ")`n    " label "`n`n"
        }
        if (currentItems.Length > currentCount)
            m .= "`n⚠ Showing first " currentCount " of " currentItems.Length " reminders"
        return m
    }

    ; Derive count using integer-only operations (avoid Float issues in ranges/loops).
    count := ClampCount(items.Length)
    msg := BuildMsg(items, count)

    keyCallbacks := Map()
    ; Register all potential selection keys once (prevents needing to rebuild callbacks on refresh).
    maxKeyCount := maxItems
    if (maxKeyCount > keys.Length)
        maxKeyCount := keys.Length
    loop maxKeyCount {
        i := A_Index
        k := keys[i]
        keyCallbacks.Set(k, Reminders_PickKey.Bind(k))
    }
    keyCallbacks.Set("Escape", Reminders_PickKey.Bind("ESC"))

    ; #region agent log
    try Reminders_DebugLog("Shift keys.ahk:Reminders_SelectItem", "About to show selection modal", Map(
        "actionLabel", actionLabel,
        "count", count,
        "priorKey", A_PriorKey,
        "thisHotkey", A_ThisHotkey,
        "remHwnd", remHwnd
    ), "S1", "pre-fix")
    ; #endregion

    ; Prevent immediate auto-selection / ignored picks caused by modifiers still being held from the trigger hotkey.
    ; If the trigger uses Win/Ctrl/Alt/Shift, wait for release before listening for selection keys.
    try {
        if (A_ThisHotkey != "") {
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
        }
    } catch {
    }

    StandardLoadingBar_CloseKeysOverlay()
    ShowModal() {
        if (!IsBlackoutSuppressed()) {
            if (IsObject(keyCallbacks))
                keyCallbacks["D"] := DisableBlackout7Min
            StandardLoadingBar_ShowWithKeys(
                msg,
                keyCallbacks,
                45000,
                0,
                Reminders_PickTimeout,
                "1E1E2E",
                760,
                14,
                BANNER_ACCENT_INTERMEDIATE,
                false,
                "[1-9/A-Z] Select  [Esc] Cancel",
                true
            )
        }
    }
    lastSig := Reminders_ItemsListSignature(items, maxItems)
    ShowModal()
    try {
        latest := Reminders_GetItems(remHwnd)
        if (latest.Length = 0) {
            ; If empty immediately after show, use more aggressive retry
            latest := Reminders_GetItemsStable(remHwnd, 100, 4)
        }
        latestCount := ClampCount(latest.Length)
        sig := Reminders_ItemsListSignature(latest, maxItems)
        if (sig != lastSig) {
            lastSig := sig
            items := latest
            count := latestCount
            msg := BuildMsg(items, count)
            try StandardLoadingBar_Update(msg)
            ; #region agent log
            try Reminders_DebugLog("Shift keys.ahk:Reminders_SelectItem", "Post-show reminder sync", Map(
                "itemsCount", items.Length,
                "showing", count,
                "remHwnd", remHwnd
            ), "SR0", "pre-fix")
            ; #endregion
        }
    } catch {
    }
    ; #region agent log
    try Reminders_DebugLog("Shift keys.ahk:Reminders_SelectItem", "Selection modal shown", Map(
        "count", count,
        "remHwnd", remHwnd
    ), "S3", "pre-fix")
    ; #endregion

    deadline := A_TickCount + 45000
    reopens := 0
    lastRefreshTick := A_TickCount
    pollMs := 300
    while (A_TickCount < deadline) {
        if (g_RemindersPickKey != "")
            break

        ; Live refresh: if reminders change while the modal is open, update the displayed list.
        if (A_TickCount - lastRefreshTick >= pollMs) {
            lastRefreshTick := A_TickCount
            try {
                latest := Reminders_GetItems(remHwnd)
                if (latest.Length = 0) {
                    ; If empty, try with stability check and longer waits (volatile window)
                    latest := Reminders_GetItemsStable(remHwnd, 100, 4)
                }
                latestCount := ClampCount(latest.Length)
                sig := Reminders_ItemsListSignature(latest, maxItems)
                if (sig != lastSig) {
                    lastSig := sig
                    items := latest
                    count := latestCount
                    msg := BuildMsg(items, count)
                    try StandardLoadingBar_Update(msg)
                    ; #region agent log
                    try Reminders_DebugLog("Shift keys.ahk:Reminders_SelectItem", "Live-refreshed reminder list", Map(
                        "itemsCount", items.Length,
                        "showing", count,
                        "remHwnd", remHwnd
                    ), "SR1", "pre-fix")
                    ; #endregion
                }
            } catch {
            }
        }

        ; Detect unexpected overlay dismissal (another script/banner replaced it).
        try {
            global g_StandardLoadingBarIsKeysOverlay, g_StandardLoadingBarGui
            if (!g_StandardLoadingBarIsKeysOverlay || !IsObject(g_StandardLoadingBarGui)) {
                ; #region agent log
                try Reminders_DebugLog("Shift keys.ahk:Reminders_SelectItem",
                    "Selection modal disappeared unexpectedly", Map(
                        "isKeys", g_StandardLoadingBarIsKeysOverlay ? 1 : 0,
                        "hasGui", IsObject(g_StandardLoadingBarGui) ? 1 : 0,
                        "priorKey", A_PriorKey
                    ), "S4", "pre-fix")
                ; #endregion
                reopens++
                if (reopens >= 3) {
                    ; #region agent log
                    try Reminders_DebugLog("Shift keys.ahk:Reminders_SelectItem",
                        "Too many unexpected closes; giving up",
                        Map("reopens", reopens), "S6", "pre-fix")
                    ; #endregion
                    break
                }
                ; Re-show the modal; another overlay likely replaced it.
                ; #region agent log
                try Reminders_DebugLog("Shift keys.ahk:Reminders_SelectItem", "Reopening selection modal after close",
                    Map("reopens", reopens), "S5", "pre-fix")
                ; #endregion
                try StandardLoadingBar_CloseKeysOverlay()
                try StandardLoadingBar_Hide(0)
                Sleep 50
                ShowModal()
                try {
                    latest := Reminders_GetItems(remHwnd)
                    if (latest.Length = 0) {
                        ; If empty after reopen, use more aggressive retry
                        latest := Reminders_GetItemsStable(remHwnd, 100, 4)
                    }
                    latestCount := ClampCount(latest.Length)
                    sig := Reminders_ItemsListSignature(latest, maxItems)
                    if (sig != lastSig) {
                        lastSig := sig
                        items := latest
                        count := latestCount
                        msg := BuildMsg(items, count)
                        try StandardLoadingBar_Update(msg)
                        try Reminders_DebugLog("Shift keys.ahk:Reminders_SelectItem", "Post-reopen reminder sync", Map(
                            "itemsCount", items.Length,
                            "showing", count,
                            "remHwnd", remHwnd
                        ), "SR0b", "pre-fix")
                    }
                } catch {
                }
            }
        } catch {
        }
        Sleep 50
    }

    picked := g_RemindersPickKey
    ; #region agent log
    try Reminders_DebugLog("Shift keys.ahk:Reminders_SelectItem", "Selection modal resolved", Map(
        "picked", picked
    ), "S2", "pre-fix")
    ; #endregion
    if (picked = "" || picked = "ESC" || picked = "TIMEOUT")
        return 0

    ; Resolve key to index
    loop count {
        i := A_Index
        if (keys[i] = picked)
            return i
    }
    return 0
}

Reminders_OpenContextMenuForItem(itemEl) {
    remHwnd := WinExist("A")
    if remHwnd {
        WinActivate("ahk_id " remHwnd)
        WinWaitActive("ahk_id " remHwnd, , 1)
        Sleep 80  ; Give time for activation to settle
    }
    try itemEl.SetFocus()
    Sleep Reminders_DelayValue("context_focus_ms", 80)
    EnsureFocus()
    Sleep Reminders_DelayValue("focus_to_menu_ms", 25)  ; focus → menu open
    ; Apps/Menu key (keyboard-only)
    Send "{AppsKey}"
    Sleep Reminders_DelayValue("menu_render_ms", 70)  ; context menu rendered before Home/scan (dismiss / snooze / join online)
}

Reminders_MenuGetFocusedName() {
    try {
        fe := UIA.GetFocusedElement()
        if !fe
            return ""
        return fe.Name
    } catch {
        return ""
    }
}

Reminders_MenuFindItemContains(needle, maxSteps := 20, logId := "SN") {
    ; Sequentially navigate the context menu looking for an item whose focused text contains needle.
    ; Returns true when found (focus rests on that item).
    needle := StrLower(needle)
    Send "{Home}"
    Sleep Reminders_DelayValue("menu_home_ms", 60)
    loop maxSteps {
        name := Reminders_MenuGetFocusedName()
        shownName := (name != "") ? name : "(no focused text)"
        try StandardLoadingBar_Update("👁️ Menu scan " A_Index "/" maxSteps ": " shownName,
            BANNER_ACCENT_INTERMEDIATE)
        if (Mod(A_Index, 5) = 0) {
            ; #region agent log
            try Reminders_DebugLog("Shift keys.ahk:Reminders_MenuFindItemContains", "Menu scan step", Map(
                "logId", logId,
                "step", A_Index,
                "name", name
            ), "SN1", "pre-fix")
            ; #endregion
        }
        if (name != "" && InStr(StrLower(name), needle)) {
            try StandardLoadingBar_Update("✅ Menu selected: " name, BANNER_ACCENT_SUCCESS)
            ; #region agent log
            try Reminders_DebugLog("Shift keys.ahk:Reminders_MenuFindItemContains", "Menu item matched", Map(
                "logId", logId,
                "step", A_Index,
                "name", name,
                "needle", needle
            ), "SN2", "pre-fix")
            ; #endregion
            return true
        }
        Send "{Down}"
        Sleep Reminders_DelayValue("menu_scan_step_ms", 50)
    }
    ; #region agent log
    try Reminders_DebugLog("Shift keys.ahk:Reminders_MenuFindItemContains", "Menu item not found", Map(
        "logId", logId,
        "needle", needle,
        "maxSteps", maxSteps
    ), "SN3", "pre-fix")
    ; #endregion
    return false
}

Reminders_MenuOpenSubmenuRight() {
    ; Expand focused menu item.
    Send "{Right}"
    Sleep Reminders_DelayValue("submenu_open_ms", 80)
}

Reminders_MenuFindAndOpenSnooze(maxSteps := 20) {
    ; Find "Snooze reminder" regardless of position, then open its submenu.
    if !Reminders_MenuFindItemContains("snooze", maxSteps, "snooze")
        return false
    Reminders_MenuOpenSubmenuRight()
    return true
}

Reminders_NormalizeDurationNeedle(d) {
    d := StrLower(Trim(d))
    if (d = "30m" || d = "30 min" || d = "30 mins" || d = "30 minutes")
        return "30 minutes"
    if (d = "1h" || d = "1 hr" || d = "1 hour" || d = "one hour")
        return "1 hour"
    if (d = "4h" || d = "4 hrs" || d = "4 hours")
        return "4 hours"
    if (d = "1d" || d = "1 day")
        return "1 day"
    if (d = "1w" || d = "1 week")
        return "1 week"
    return d
}

Reminders_MenuFindAndSelectDuration(durationNeedle, maxSteps := 50) {
    durationNeedle := Reminders_NormalizeDurationNeedle(durationNeedle)
    ; After snooze submenu is open, scan items by focused text and press Enter on match.
    if !Reminders_MenuFindItemContains(durationNeedle, maxSteps, "dur:" durationNeedle)
        return false
    Send "{Enter}"
    Sleep Reminders_DelayValue("post_select_ms", 60)
    ; #region agent log
    try Reminders_DebugLog("Shift keys.ahk:Reminders_MenuFindAndSelectDuration", "Duration selected", Map(
        "duration", durationNeedle
    ), "SD1", "pre-fix")
    ; #endregion
    return true
}

Reminders_TryInvokeJoinOnlineMenuItem() {
    ; Context menus are often hosted outside the window subtree,
    ; so search from the UIA root element (desktop).
    try {
        rootDesktop := UIA.GetRootElement()
        rootWin := UIA.ElementFromHandle(WinExist("A"))

        roots := []
        if rootDesktop
            roots.Push({ root: rootDesktop, label: "desktop" })
        if rootWin
            roots.Push({ root: rootWin, label: "window" })

        for r in roots {
            root := r.root

            ; #region agent log
            try {
                menuItems := root.FindAll({ ControlType: "MenuItem" })
                cnt := menuItems ? menuItems.Length : 0
                sample := ""
                if menuItems {
                    maxSample := Min(20, menuItems.Length)
                    loop maxSample {
                        i := A_Index
                        n := ""
                        try n := menuItems[i].Name
                        if (n = "")
                            continue
                        if (sample != "")
                            sample .= " | "
                        sample .= n
                    }
                }
                Reminders_DebugLog("Shift keys.ahk:Reminders_TryInvokeJoinOnlineMenuItem", "MenuItems snapshot", Map(
                    "root", r.label,
                    "count", cnt,
                    "sample", sample
                ), "C1", "pre-fix")
            } catch {
            }
            ; #endregion

            candidates := [{ Name: "Join online", ControlType: "MenuItem" }, { Name: "Join Online", ControlType: "MenuItem" }, { Name: "Join online",
                ControlType: "Button" }, { Name: "Join Online", ControlType: "Button" }, { Name: "Join", matchmode: "Substring",
                    ControlType: "MenuItem", cs: false }, { Name: "Join", matchmode: "Substring", ControlType: "Button",
                        cs: false }
            ]

            for crit in candidates {
                mi := root.FindFirst(crit)
                if mi {
                    ; #region agent log
                    try Reminders_DebugLog("Shift keys.ahk:Reminders_TryInvokeJoinOnlineMenuItem",
                        "Found Join candidate", Map(
                            "root", r.label,
                            "name", mi.Name,
                            "type", mi.ControlType
                        ), "C2", "pre-fix")
                    ; #endregion
                    try {
                        if mi.GetPropertyValue(UIA.Property.IsOffscreen)
                            continue
                    } catch {
                    }
                    try {
                        if !mi.GetPropertyValue(UIA.Property.IsEnabled)
                            continue
                    } catch {
                    }
                    try mi.Click()
                    catch Error {
                        try mi.Invoke()
                    }
                    return true
                }
            }
        }
    } catch {
    }
    return false
}

Reminders_ExecuteItemAction(action) {
    ; Always hide the loading indicator, even on early returns or exceptions.
    ; (Prevents a stuck indicator if UIA/menu calls throw.)
    try {
        ; #region agent log
        try Reminders_DebugLog("Shift keys.ahk:Reminders_ExecuteItemAction", "Enter ExecuteItemAction", Map(
            "action", action,
            "title", WinGetTitle("A"),
            "class", WinGetClass("A"),
            "remHwnd", WinExist("A")
        ), "J1", "pre-fix")
        ; #endregion
        remHwnd := WinExist("A")

        ; Nudge window to refresh UIA before getting items (same as in SelectItem)
        Reminders_NudgeWindowForUiRefresh(remHwnd)
        Sleep 100

        ; Try 3 times with longer waits before accepting empty result
        items := []
        loop 3 {
            attempt := A_Index
            triedItems := Reminders_GetItemsStable(remHwnd, 100, 5)
            try Reminders_DebugLog("Shift keys.ahk:Reminders_ExecuteItemAction",
                "Attempt " attempt " retrieval", Map(
                    "attempt", attempt,
                    "itemsCount", triedItems.Length
                ), "J2-ATT", "pre-fix")
            if (triedItems.Length > 0) {
                items := triedItems
                break
            }
            if (attempt < 3) {
                Reminders_NudgeWindowForUiRefresh(remHwnd)
                Sleep 150  ; longer wait before retry
            }
        }

        ; Final fallback if still empty
        if (items.Length = 0) {
            try items := Reminders_GetItemsStable(remHwnd, 150, 6)
        }

        ; #region agent log
        remTitle := ""
        remClass := ""
        try remTitle := WinGetTitle("ahk_id " remHwnd)
        try remClass := WinGetClass("ahk_id " remHwnd)
        try Reminders_DebugLog("Shift keys.ahk:Reminders_ExecuteItemAction", "Items extracted", Map(
            "action", action,
            "itemsCount", items.Length,
            "remHwnd", remHwnd,
            "remTitle", remTitle,
            "remClass", remClass,
            "activeTitle", WinGetTitle("A"),
            "activeClass", WinGetClass("A")
        ), "J2", "pre-fix")
        ; #endregion
        actionLabel := ""
        if (action = "snooze_1h")
            actionLabel := "Snooze 1 hour"
        else if (action = "snooze_4h")
            actionLabel := "Snooze 4 hours"
        else if (action = "snooze_10m")
            actionLabel := "Snooze 10 minutes"
        else if (action = "snooze_1d")
            actionLabel := "Snooze 1 day"
        else if (action = "snooze_1w")
            actionLabel := "Snooze 1 week"
        else if (action = "dismiss_item")
            actionLabel := "Dismiss reminder"
        else if (action = "join_online")
            actionLabel := "Join online"
        else
            actionLabel := action

        ; Diagnostic: if items empty, show message and log more info
        if (items.Length = 0) {
            ; Capture raw UIA structure from the reminder window before returning.
            try Reminders_DebugDumpWindowTree(remHwnd, "empty-extraction")
            try {
                global g_DebugBe11ecLogPath
                FileAppend(
                    "{`"sessionId`":`"be11ec`",`"id`":`"diagnostic_" A_TickCount "_" Random(1000, 9999)
                    "`",`"timestamp`":" A_TickCount ",`"location`":`"ExecuteItemAction:empty-items`",`"message`":`"No items found after stable retrieval`",`"data`":{"
                    "`"action`":`"" action "`",`"remHwnd`":" remHwnd ",`"title`":`"" remTitle
                    "`",`"class`":`"" remClass "`",`"activeTitle`":`"" WinGetTitle("A") "`",`"activeClass`":`"" WinGetClass(
                        "A") "`"},`"runId`":`"pre-fix`",`"hypothesisId`":`"Z`"}`n",
                    g_DebugBe11ecLogPath,
                    "UTF-8"
                )
            } catch {
            }
            ShowCenteredOverlay_Utils(
                "❌ No reminders found in window. Check debug log at C:\Users\fie7ca\Documents\scripts\debug-be11ec.log",
                2500, BANNER_ACCENT_ERROR)
            return false
        }

        Reminders_LoadingShow("⏳ Reminders: " actionLabel "…")

        ; The selection modal is interactive; hide loading before it shows.
        Reminders_LoadingHide(0)
        idx := Reminders_SelectItem(actionLabel, &items, remHwnd)
        ; #region agent log
        try Reminders_DebugLog("Shift keys.ahk:Reminders_ExecuteItemAction", "Selection result", Map(
            "action", action,
            "selectedIndex", idx
        ), "J3", "pre-fix")
        ; #endregion
        if (!idx)
            return false

        ; Resume loading while executing the chosen action.
        Reminders_LoadingShow("⏳ Reminders: " actionLabel "…")

        el := items[idx].el
        Sleep Reminders_DelayValue("modal_to_row_ms", 35)  ; modal closed → row target ready before context menu (join / dismiss / snooze)
        ; #region agent log
        try Reminders_DebugLog("Shift keys.ahk:Reminders_ExecuteItemAction", "Selected reminder item", Map(
            "action", action,
            "selectedIndex", idx,
            "itemsCount", items.Length,
            "title", WinGetTitle("A")
        ), "A", "pre-fix")
        ; #endregion
        Reminders_OpenContextMenuForItem(el)
        ; #region agent log
        try Reminders_DebugLog("Shift keys.ahk:Reminders_ExecuteItemAction", "Context menu open attempt sent AppsKey",
            Map(
                "action", action
            ), "B", "pre-fix")
        ; #endregion

        ; Assume first menu item is highlighted (Snooze reminder) as per screenshots.
        if (action = "dismiss_item") {
            ; Menu order varies by reminder item (e.g. meeting reminders show Join/Chat first),
            ; so locate "Dismiss reminder" by focused UIA name instead of fixed offsets.
            ; #region agent log
            try Reminders_DebugLog("Shift keys.ahk:Reminders_ExecuteItemAction", "Dynamic dismiss requested", Map(),
            "DZ0",
            "pre-fix")
            ; #endregion
            ok := Reminders_MenuFindItemContains("dismiss", 20, "dismiss")
            if ok {
                Send "{Enter}"
                Sleep Reminders_DelayValue("post_select_ms", 60)
                ; #region agent log
                try Reminders_DebugLog("Shift keys.ahk:Reminders_ExecuteItemAction", "Dismiss invoked", Map(), "DZ1",
                "pre-fix")
                ; #endregion
                return true
            }
            ; #region agent log
            try Reminders_DebugLog("Shift keys.ahk:Reminders_ExecuteItemAction", "Dismiss not found in menu scan", Map(),
            "DZ2",
            "pre-fix")
            ; #endregion
            return false
        }

        if (action = "snooze_1h" || action = "snooze_4h" || action = "snooze_10m" || action = "snooze_1d" || action =
            "snooze_1w") {
            desired := ""
            if (action = "snooze_1h")
                desired := "1 hour"
            else if (action = "snooze_4h")
                desired := "4 hours"
            else if (action = "snooze_10m")
                desired := "10 minutes"
            else if (action = "snooze_1d")
                desired := "1 day"
            else if (action = "snooze_1w")
                desired := "1 week"
            ; #region agent log
            try Reminders_DebugLog("Shift keys.ahk:Reminders_ExecuteItemAction", "Dynamic snooze requested", Map(
                "desired", desired
            ), "SZ0", "pre-fix")
            ; #endregion
            if !Reminders_MenuFindAndOpenSnooze(20)
                return false
            return Reminders_MenuFindAndSelectDuration(desired, 60)
        }

        if (action = "join_online") {
            ; Preferred: direct UIA invoke (menu items are usually under UIA root).
            ok := false
            ; #region agent log
            try Reminders_DebugLog("Shift keys.ahk:Reminders_ExecuteItemAction", "Attempting UIA root Join invoke", Map(),
            "C",
            "pre-fix")
            ; #endregion
            ok := Reminders_TryInvokeJoinOnlineMenuItem()
            ; #region agent log
            try Reminders_DebugLog("Shift keys.ahk:Reminders_ExecuteItemAction", "UIA root Join invoke result", Map(
                "ok", ok), "C",
            "pre-fix")
            ; #endregion
            if ok
                return true

            ; Fallback 1: first-letter navigation (if supported)
            ; #region agent log
            try Reminders_DebugLog("Shift keys.ahk:Reminders_ExecuteItemAction", "Fallback: type 'j' then Enter", Map(),
            "D",
            "pre-fix")
            ; #endregion
            Send "j"
            Sleep 60
            Send "{Enter}"
            Sleep 80

            ; Fallback 2: bounded arrow scan (best-effort, no UIA reads)
            ; #region agent log
            try Reminders_DebugLog("Shift keys.ahk:Reminders_ExecuteItemAction",
                "Fallback: bounded arrow scan then Enter", Map(),
                "E", "pre-fix")
            ; #endregion
            Send "{Home}"
            loop 12 {
                Send "{Down}"
                Sleep 40
            }
            Send "{Enter}"

            return false
        }

        return false
    } finally {
        Reminders_LoadingHide(0)
    }
}

; Shift + H : Snooze 1 hour - Hour
+H:: {
    isNewRem := Reminders_IsNewOutlookWindow()
    ; #region agent log
    try Reminders_DebugLog("Shift keys.ahk:+H", "Shift+H pressed", Map(
        "isNewReminders", isNewRem,
        "title", WinGetTitle("A"),
        "class", WinGetClass("A")
    ), "H2", "pre-fix")
    ; #endregion
    if isNewRem {
        Reminders_ExecuteItemAction("snooze_1h")
        return
    }
    ; #region agent log
    try Reminders_DebugLog("Shift keys.ahk:+H", "Falling back to classic Confirm(1 hour)", Map(), "H3", "pre-fix")
    ; #endregion
    Confirm("1 hour")
}

; Shift + F : Snooze 4 hours - Four
+F:: {
    if Reminders_IsNewOutlookWindow() {
        Reminders_ExecuteItemAction("snooze_4h")
        return
    }
    Confirm("4 hours")
}

; Shift + T : Snooze 10 minutes - Ten
+T:: {
    if Reminders_IsNewOutlookWindow() {
        Reminders_ExecuteItemAction("snooze_10m")
        return
    }
    Confirm("10 minutes")
}

; Shift + Y : Snooze 1 day - daY
+Y:: {
    if Reminders_IsNewOutlookWindow() {
        Reminders_ExecuteItemAction("snooze_1d")
        return
    }
    Confirm("1 day")
}

; Shift + W : Snooze 1 week - Week
+W:: {
    if Reminders_IsNewOutlookWindow() {
        Reminders_ExecuteItemAction("snooze_1w")
        return
    }
    Confirm("1 week")
}

; Shift + D : Dismiss reminder (New Outlook) / Snooze 1 day (classic)
+D:: {
    if Reminders_IsNewOutlookWindow() {
        Reminders_ExecuteItemAction("dismiss_item")
        return
    }
    Confirm("1 day")
}

; Shift + X : Dismiss all reminders - Dismiss (confirm first: New Outlook + classic)
+X:: {
    if Reminders_IsNewOutlookWindow() {
        if MsgBox("Dismiss all reminders?", "Confirm Dismiss", "YesNo Icon?") != "Yes"
            return
        ; Ensure loading indicator can't get stuck on exceptions.
        try {
            Reminders_LoadingShow("⏳ Reminders: Dismiss all…")
            ; Global action: click "Dismiss all" button (UIA)
            try {
                root := UIA.ElementFromHandle(WinExist("A"))
                btn := root.FindFirst({ Name: "Dismiss all", ControlType: "Button" })
                if !btn
                    btn := root.FindFirst({ Name: "Dismiss All", ControlType: "Button" })
                if btn {
                    btn.Click()
                    return
                }
            } catch {
            }
            return
        } finally {
            Reminders_LoadingHide(0)
        }
    }
    ConfirmDismissAll()
}

; Shift + J : Join Online - Join
+J:: {
    isNewRem := Reminders_IsNewOutlookWindow()
    ; #region agent log
    try Reminders_DebugLog("Shift keys.ahk:+J", "Shift+J pressed", Map(
        "isNewReminders", isNewRem,
        "title", WinGetTitle("A"),
        "class", WinGetClass("A")
    ), "F", "pre-fix")
    ; #endregion
    if isNewRem {
        Reminders_ExecuteItemAction("join_online")
        return
    }
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Find the "Join Online" button (classic reminder UI)
        joinButton := root.FindFirst({ Name: "Join Online", Type: "50000", AutomationId: "8346" })
        if !joinButton
            joinButton := root.FindFirst({ Name: "Join Online", ControlType: "Button" })
        if !joinButton
            joinButton := root.FindFirst({ AutomationId: "8346" })

        if joinButton
            joinButton.Click()
    } catch {
    }
}

#HotIf

;-------------------------------------------------------------------
; Microsoft Teams Helper functions
;-------------------------------------------------------------------
TEAMS_PROCESSES := ["ms-teams.exe", "Teams.exe", "MSTeams.exe"]

IsTeamsMeetingTitle(title) {
    if InStr(title, "Chat |") || InStr(title, "Sharing control bar |")
        return false
    if InStr(title, "Microsoft Teams meeting")
        return true
    return RegExMatch(title, "i)^.*\| Microsoft Teams.*$")
}

IsTeamsChatTitle(title) {
    if InStr(title, "Sharing control bar |") || InStr(title, "Microsoft Teams meeting")
        return false
    return InStr(title, "Chat |") && RegExMatch(title, "i)\| Microsoft Teams$")
}

; -------------------------------------------------------------------
; Helper predicates to detect which Teams window is active
; -------------------------------------------------------------------
IsTeamsMeetingActive() {
    return IsTeamsMeetingTitle(WinGetTitle("A"))
}
IsTeamsChatActive() {
    return IsTeamsChatTitle(WinGetTitle("A"))
}

; -------------------------------------------------------------------
; Microsoft Teams Shortcuts â€" MEETING WINDOW
; -------------------------------------------------------------------
#HotIf IsTeamsMeetingActive()

; Shift + C : Open Chat pane - Chat
+C:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        btn := root.FindFirst({ AutomationId: "chat-button" })
        if !btn
            btn := root.FindFirst({ Name: "Chat", ControlType: "Button" })
        if !btn
            btn := root.FindFirst({ Name: "Bate-papo", ControlType: "Button" })

        if btn
            btn.Click()
        else
            MsgBox("Couldn't find the Chat button.", "Control not found", "IconX")
    }
    catch as e {
        MsgBox("UIA error:`n" e.Message, "Error", "IconX")
    }
}

; Shift + M : Maximize meeting window - Maximize
+M:: {
    ; Get current active window title
    currentTitle := WinGetTitle("A")

    ; Check if current window is a compacted Teams meeting
    isCompacted := false
    if (currentTitle = "Reunião do Microsoft Teams | Microsoft Teams") {
        isCompacted := true
        baseTitle := "Reunião do Microsoft Teams"
    } else if (InStr(currentTitle,
        "Modo de exibição compacto da reunião | Reunião do Microsoft Teams | Microsoft Teams")) {
        isCompacted := true
        baseTitle := "Reunião do Microsoft Teams"
    }

    if (!isCompacted) {
        ; Not a compacted meeting window, do nothing
        return
    }

    ; Search for the corresponding normal meeting window
    normalMeetingHwnd := 0
    for hwnd in WinGetList("ahk_exe ms-teams.exe") {
        title := WinGetTitle(hwnd)

        ; Skip if it's the same window or another compacted window
        if (hwnd = WinGetID("A") || InStr(title, "Modo de exibição compacto da reunião")) {
            continue
        }

        ; Check if it's a normal meeting window with the same base title
        if (InStr(title, baseTitle) && InStr(title, "| Microsoft Teams") && !InStr(title,
            "Modo de exibição compacto da reunião")) {
            normalMeetingHwnd := hwnd
            break
        }
    }

    ; If found, switch to the normal meeting window
    if (normalMeetingHwnd) {
        if (!WinExist("ahk_id " normalMeetingHwnd)) {
            ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        try {
            WinActivate("ahk_id " normalMeetingHwnd)
            ; Optional: Show a brief tooltip to confirm the switch
            ShowCenteredOverlay_Utils("✅ Switched to normal meeting view", 1000, BANNER_ACCENT_SUCCESS)
        } catch as e {
            ; Fallback: try to bring window to front (only if window still exists)
            if (WinExist("ahk_id " normalMeetingHwnd)) {
                WinShow("ahk_id " normalMeetingHwnd)
                WinActivate("ahk_id " normalMeetingHwnd)
            } else {
                ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
            }
        }
    } else {
        ; No corresponding normal window found - show notification
        ShowCenteredOverlay_Utils("⚠ No normal meeting window found", 1500, BANNER_ACCENT_ERROR)
    }
}

; Shift + R : React / Reagir - React
+R:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        btn := 0
        try {
            btn := root.FindFirst({ AutomationId: "reaction-menu-button" })
        } catch {
        }
        if !btn {
            try {
                btn := root.FindFirst({ Name: "Reagir", ControlType: "Button" })
            } catch {
            }
        }
        if !btn {
            try {
                btn := root.FindFirst({ Name: "React", ControlType: "Button" })
            } catch {
            }
        }

        if btn
            btn.Click()
        else
            MsgBox("Couldn't find the Reagir button.", "Control not found", "IconX")
    }
    catch as e {
        MsgBox("UIA error:`n" e.Message, "Error", "IconX")
    }
}

; Shift + J : Join now with camera and microphone on - Join
+J:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Find the "Join now" button by AutomationId first
        btn := root.FindFirst({ AutomationId: "prejoin-join-button" })

        ; Fallback: try finding by name (Portuguese)
        if !btn {
            btn := root.FindFirst({ Name: "Ingressar agora Com a cÃ¢mera ligada e Microfone ligado", ControlType: "Button" })
        }

        ; Fallback: try finding by name (English)
        if !btn {
            btn := root.FindFirst({ Name: "Join now with camera and microphone on", ControlType: "Button" })
        }

        ; Fallback: try finding by partial name (Portuguese)
        if !btn {
            btn := root.FindFirst({ Name: "Ingressar agora", ControlType: "Button" })
        }

        ; Fallback: try finding by partial name (English)
        if !btn {
            btn := root.FindFirst({ Name: "Join now", ControlType: "Button" })
        }

        if btn {
            btn.Click()
        }
        ; No message box as requested - fail silently
    }
    catch {
        ; No message box as requested - fail silently
    }
}

; Shift + A : Audio settings - Audio
+A:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Find the audio settings button by AutomationId first
        btn := root.FindFirst({ AutomationId: "prejoin-audiosettings-button" })

        ; Fallback: try finding by name (Portuguese)
        if !btn {
            btn := root.FindFirst({ Name: "Microfone do computador e controles do alto-falante ConfiguraÃ§Ãµes de Ã¡udio",
                ControlType: "Button" })
        }

        ; Fallback: try finding by name (English)
        if !btn {
            btn := root.FindFirst({ Name: "Computer microphone and speaker controls Audio settings",
                ControlType: "Button" })
        }

        ; Fallback: try finding by partial name (Portuguese)
        if !btn {
            btn := root.FindFirst({ Name: "ConfiguraÃ§Ãµes de Ã¡udio", ControlType: "Button" })
        }

        ; Fallback: try finding by partial name (English)
        if !btn {
            btn := root.FindFirst({ Name: "Audio settings", ControlType: "Button" })
        }

        if btn {
            btn.Click()
        }
        ; No message box as requested - fail silently
    }
    catch {
        ; No message box as requested - fail silently
    }
}

#HotIf

;-------------------------------------------------------------------
; Wikipedia Shortcuts
;-------------------------------------------------------------------
; Global variable to track scroll position history (stack: most recent last)
global g_WikipediaScrollHistory := []

#HotIf WinActive("ahk_exe chrome.exe") && InStr(WinGetTitle("A"), "Wikipedia", false)

; Shift + S: Focus the Wikipedia search field (prefer the field; if hidden, click the Search toggle first)
+s::
{
    try {
        ; Step 1: Always scroll to the beginning of the page first
        Send "^Home"
        Sleep 300 ; Give page time to scroll

        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 200 ; Give UIA time to attach

        ; Prefer document; fall back to browser root
        try {
            root := uia.GetCurrentDocumentElement()
        } catch {
            root := uia.BrowserElement
        }

        ; Step 2: Try to locate the "Search Wikipedia" field by name (combo box / edit)
        searchBox := 0

        ; First try: ComboBox with the expected name (Type 50003)
        try {
            searchBox := root.FindElement({ Type: 50003, Name: "Search Wikipedia", cs: false })
        } catch {
        }

        ; Second: Edit control with the same name (in case UI changes type)
        if (!searchBox) {
            try {
                searchBox := root.FindElement({ Type: 50004, Name: "Search Wikipedia", cs: false })
            } catch {
            }
        }

        ; Third: any element by name "Search Wikipedia"
        if (!searchBox) {
            try {
                searchBox := root.FindElement({ Name: "Search Wikipedia", cs: false })
            } catch {
            }
        }

        ; If we found the field, focus/click it and we're done.
        if (searchBox) {
            try {
                searchBox.SetFocus()
            } catch {
                searchBox.Click()
            }
            return
        }

        ; Step 3: If the field is not available, try clicking the "Search" toggle button/link first.
        searchToggle := 0

        ; Strategy 1: Try finding the search group first, then the button within it (most reliable)
        try {
            searchGroup := root.FindElement({ AutomationId: "p-search", cs: false })
            if (searchGroup) {
                try {
                    searchToggle := searchGroup.FindElement({ Type: 50005, Name: "Search", cs: false })
                } catch {
                }
                if (!searchToggle) {
                    try {
                        searchToggle := searchGroup.FindElement({ Type: 50005, Value: "https://en.wikipedia.org/wiki/Special:Search",
                            cs: false })
                    } catch {
                    }
                }
            }
        } catch {
        }

        ; Strategy 2: Search by Value (URL) directly from root
        if (!searchToggle) {
            try {
                searchToggle := root.FindElement({ Type: 50005, Value: "https://en.wikipedia.org/wiki/Special:Search",
                    cs: false })
            } catch {
            }
        }

        ; Strategy 3: Search by Type and Name (original method)
        if (!searchToggle) {
            try {
                searchToggle := root.FindElement({ Type: 50005, Name: "Search", cs: false })
            } catch {
            }
        }

        ; Strategy 4: Search by ControlType and Name
        if (!searchToggle) {
            try {
                searchToggle := root.FindElement({ ControlType: "Hyperlink", Name: "Search", cs: false })
            } catch {
            }
        }

        ; Strategy 5: Search for any link with the Special:Search URL
        if (!searchToggle) {
            try {
                searchToggle := root.FindElement({ Type: 50005, Value: "*Special:Search*", cs: false })
            } catch {
            }
        }

        ; Strategy 6: Try finding by LocalizedType "link" and Name
        if (!searchToggle) {
            try {
                searchToggle := root.FindElement({ LocalizedType: "link", Name: "Search", cs: false })
            } catch {
            }
        }

        if (searchToggle) {
            ; Validate element before clicking: check if it's offscreen or disabled
            isValid := true
            try {
                isOffscreen := searchToggle.GetPropertyValue(UIA.Property.IsOffscreen)
                if (isOffscreen) {
                    isValid := false
                }
            } catch {
                ; Continue even if property check fails
            }

            try {
                isEnabled := searchToggle.GetPropertyValue(UIA.Property.IsEnabled)
                if (!isEnabled) {
                    isValid := false
                }
            } catch {
                ; Continue even if property check fails
            }

            ; If element is invalid, try re-finding it
            if (!isValid) {
                Sleep 200
                searchToggle := 0
                try {
                    searchToggle := root.FindElement({ Type: 50005, Name: "Search", cs: false })
                } catch {
                }
                if (!searchToggle) {
                    try {
                        searchToggle := root.FindElement({ ControlType: "Hyperlink", Name: "Search", cs: false })
                    } catch {
                    }
                }
            }

            ; Try multiple click strategies if element is still valid
            if (searchToggle) {
                clicked := false

                ; Strategy 1: Try Invoke pattern (most reliable for links/buttons)
                try {
                    searchToggle.Invoke()
                    clicked := true
                } catch Error as invokeErr {
                }

                ; Strategy 2: Try SetFocus then Click
                if (!clicked) {
                    try {
                        searchToggle.SetFocus()
                        Sleep 50
                        searchToggle.Click()
                        clicked := true
                    } catch Error as focusClickErr {
                    }
                }

                ; Strategy 3: Try direct Click
                if (!clicked) {
                    try {
                        searchToggle.Click()
                        clicked := true
                    } catch Error as clickErr {
                    }
                }

                ; If all click strategies failed, fall back to accelerator
                if (!clicked) {
                    try {
                        uia.ControlSend("!f")
                    } catch {
                    }
                }

                ; Give the UI a moment to reveal the field, then try again to find it.
                Sleep 250

                searchBox := 0
                try {
                    searchBox := root.FindElement({ Type: 50003, Name: "Search Wikipedia", cs: false })
                } catch {
                }
                if (!searchBox) {
                    try {
                        searchBox := root.FindElement({ Type: 50004, Name: "Search Wikipedia", cs: false })
                    } catch {
                    }
                }
                if (!searchBox) {
                    try {
                        searchBox := root.FindElement({ Name: "Search Wikipedia", cs: false })
                    } catch {
                    }
                }

                if (searchBox) {
                    try {
                        searchBox.SetFocus()
                    } catch {
                        searchBox.Click()
                    }
                    return
                }
            }
        }

        ; Final fallback: use the accelerator key if all else fails.
        try {
            uia.ControlSend("!f")
            return
        } catch {
        }

        MsgBox "Could not find the 'Search Wikipedia' field."
    } catch Error as e {
        MsgBox "An error occurred: " e.Message
    }
}

; Shift + P: Save Wikipedia scroll position
+p::
{
    SaveWikipediaScrollPositionManually_ShiftKeys()
}

; Helper function to restore scroll position to a given percentage
; Returns true on success, false on failure
RestoreWikipediaScrollPosition(scrollPercentage, bannerText := "📜 Restoring scroll position... Please wait") {
    if (scrollPercentage <= 0.0 || scrollPercentage > 1.0) {
        return false
    }

    try {
        ; Create UIA_Browser once
        uia := UIA_Browser("ahk_exe chrome.exe")
        if (!uia) {
            return false
        }

        ; Show banner (foreground monitor)
        StandardLoadingBar_Show(bannerText, BANNER_ACCENT_INTERMEDIATE, { passive: true, centerOnHwnd: 0,
            textWidth: 500,
            fontSize: 17, passiveBgColor: BANNER_ACCENT_INTERMEDIATE })

        ; Block input during restoration (Phase 4: guaranteed cleanup in finally)
        BlockInput("On")
        try {
            ; Wait for page to be ready (condition-based, up to 500ms) instead of fixed Sleep(500)
            deadline := A_TickCount + 500
            docHeight := ""
            loop {
                docHeight := uia.JSReturnThroughClipboard("document.documentElement.scrollHeight")
                if (docHeight != "" && docHeight != "undefined" && docHeight != "null") {
                    try h := Float(docHeight)
                    catch {
                        h := 0
                    }
                    if (h > 0)
                        break
                }
                if (A_TickCount >= deadline)
                    break
                Sleep(50)
            }
            if (docHeight = "" || docHeight = "undefined" || docHeight = "null") {
                StandardLoadingBar_Hide(0)
                return false
            }
            docHeightFloat := Float(docHeight)
            if (docHeightFloat <= 0) {
                StandardLoadingBar_Hide(0)
                return false
            }

            ; Calculate and execute scroll
            targetScrollY := scrollPercentage * docHeightFloat
            uia.JSExecute("window.scrollTo(0, " . Round(targetScrollY) . ");")
            deadline2 := A_TickCount + 500
            while (A_TickCount < deadline2)
                Sleep(50)

            ; Update banner to show success
            try {
                StandardLoadingBar_Update("Scroll position restored!")
                Sleep(1000)
            } catch {
            }

            try {
                Sleep(500)
                StandardLoadingBar_Hide(0)
            } catch {
            }

            return true
        } catch Error as err {
            try StandardLoadingBar_Hide(0)
            catch {
            }
            return false
        } finally {
            BlockInput("Off")
        }
    } catch Error as err {
        BlockInput("Off")
        try StandardLoadingBar_Hide(0)
        catch {
        }
        return false
    }
}

; Helper function to normalize Wikipedia URLs
NormalizeWikipediaURL(url) {
    if (url = "" || !InStr(url, "wikipedia.org")) {
        return ""
    }
    ; Remove fragments and trailing slashes
    url := RegExReplace(url, "/#.*$", "")
    url := RegExReplace(url, "/+$", "")
    return url
}

; Helper function to get and normalize Wikipedia URL from active window
GetWikipediaURLNormalized() {
    try {
        if (!WinActive("ahk_exe chrome.exe") || !InStr(WinGetTitle("A"), "Wikipedia")) {
            return ""
        }
        uia := UIA_Browser("ahk_exe chrome.exe")
        url := uia.GetCurrentURL()
        return NormalizeWikipediaURL(url)
    } catch Error as err {
        return ""
    }
}

; Wikipedia scroll position save function (duplicated from AppLaunchers.ahk)
SaveWikipediaScrollPositionManually_ShiftKeys() {
    try {
        ; Check if Wikipedia window is currently active
        activeWindow := WinGetTitle("A")
        isChromeActive := WinActive("ahk_exe chrome.exe")
        hasWikipedia := InStr(activeWindow, "Wikipedia")
        if (!isChromeActive || !hasWikipedia) {
            return false
        }
    } catch Error as err {
        return false
    }

    ; Exit fullscreen before scroll position save (REQUIRED: UIA unreliable in fullscreen)
    Send("{F11}")
    Sleep(300)  ; Allow time for fullscreen exit (increased for reliability)

    ; Show banner to inform user that scroll position is being saved (foreground monitor)
    StandardLoadingBar_Show("💾 Saving scroll position... Please wait", BANNER_ACCENT_INTERMEDIATE, { passive: true,
        centerOnHwnd: 0,
        textWidth: 500, fontSize: 17, passiveBgColor: BANNER_ACCENT_INTERMEDIATE })
    fullscreenRestored := false  ; Track if we've re-entered fullscreen
    try {
        ; Get normalized Wikipedia URL
        url := GetWikipediaURLNormalized()
        if (url = "") {
            ; Re-enter fullscreen before returning
            Send("{F11}")
            Sleep(300)
            fullscreenRestored := true
            return false
        }

        ; Create UIA_Browser for getting scroll position
        uia := false
        try {
            uia := UIA_Browser("ahk_exe chrome.exe")
        } catch Error as uiaErr {
            try StandardLoadingBar_Hide(0)
            catch {
            }
            ; Re-enter fullscreen before returning
            Send("{F11}")
            Sleep(300)
            fullscreenRestored := true
            return false
        }

        if (!uia) {
            try StandardLoadingBar_Hide(0)
            catch {
            }
            ; Re-enter fullscreen before returning
            Send("{F11}")
            Sleep(300)
            fullscreenRestored := true
            return false
        }

        ; Wait for page to stabilize before measuring (critical for portrait orientation)
        ; Monitors 3 and 4 are portrait (1080x1920); layout may shift during measurement
        Sleep(500)  ; Brief stabilization wait

        ; Get scroll position with retry for stability
        scrollY := ""
        docHeight := ""
        scrollYRetries := 3
        lastScrollY := -1
        loop scrollYRetries {
            try {
                scrollY := uia.JSReturnThroughClipboard("window.pageYOffset")
                scrollYFloat := Float(scrollY)
                ; Check if scroll position is stable
                if (scrollYFloat = lastScrollY || lastScrollY = -1) {
                    if (scrollYFloat = lastScrollY) {
                        break  ; Stable, use this value
                    }
                    lastScrollY := scrollYFloat
                }
            } catch Error as scrollErr {
            }
            if (A_Index < scrollYRetries) {
                Sleep(200)  ; Wait between attempts
            }
        }

        ; Get document height to calculate percentage (with stability check)
        docHeightRetries := 3
        lastDocHeight := -1
        loop docHeightRetries {
            try {
                docHeight := uia.JSReturnThroughClipboard("document.documentElement.scrollHeight")
                docHeightFloat := Float(docHeight)
                ; Check if document height is stable
                if (docHeightFloat = lastDocHeight || lastDocHeight = -1) {
                    if (docHeightFloat = lastDocHeight) {
                        break  ; Stable, use this value
                    }
                    lastDocHeight := docHeightFloat
                }
            } catch Error as docErr {
            }
            if (A_Index < docHeightRetries) {
                Sleep(200)  ; Wait between attempts
            }
        }

        ; Convert to numbers and calculate percentage
        if (scrollY != "" && scrollY != "undefined" && scrollY != "null" && docHeight != "" && docHeight !=
            "undefined" && docHeight != "null") {
            scrollYFloat := Float(scrollY)
            docHeightFloat := Float(docHeight)
            if (scrollYFloat >= 0 && docHeightFloat > 0) {
                scrollPercentage := scrollYFloat / docHeightFloat
                ; Clamp to valid range
                if (scrollPercentage > 1.0) {
                    scrollPercentage := 1.0
                }

                ; Save to INI file
                scrollPositionsFile := A_ScriptDir "\data\wikipedia_scroll_positions.ini"
                SplitPath(scrollPositionsFile, , &dir)
                if (dir != "" && !DirExist(dir)) {
                    DirCreate(dir)
                }
                ; Always add to history stack for "go back" functionality (independent of INI file save)
                global g_WikipediaScrollHistory
                ; Add current position to history (url, scrollPercentage)
                g_WikipediaScrollHistory.Push({ url: url, scrollPercentage: scrollPercentage })
                ; Limit history to last 10 positions to prevent memory issues
                if (g_WikipediaScrollHistory.Length > 10) {
                    g_WikipediaScrollHistory.RemoveAt(1)
                }

                ; Try to save to INI file (for persistence across sessions)
                saved := false
                try {
                    ; Read existing entries first (before deleting file) to preserve them
                    existingEntries := Map()
                    if (FileExist(scrollPositionsFile)) {
                        try {
                            ; Read all existing entries from the Positions section
                            fileContent := FileRead(scrollPositionsFile)
                            ; Parse INI format manually
                            inPositionsSection := false
                            loop parse fileContent, "`n", "`r" {
                                line := Trim(A_LoopField)
                                if (line = "[Positions]") {
                                    inPositionsSection := true
                                    continue
                                }
                                if (inPositionsSection && SubStr(line, 1, 1) = "[") {
                                    ; Hit another section, stop reading
                                    break
                                }
                                if (inPositionsSection && InStr(line, "=")) {
                                    pos := InStr(line, "=")
                                    key := Trim(SubStr(line, 1, pos - 1))
                                    value := Trim(SubStr(line, pos + 1))
                                    if (key != "" && value != "") {
                                        existingEntries[key] := value
                                    }
                                }
                            }
                        } catch {
                            ; If read fails, we'll just write the new entry
                        }
                    }

                    ; Update with new entry
                    existingEntries[url] := scrollPercentage

                    ; Delete file to recreate in UTF-8
                    if (FileExist(scrollPositionsFile)) {
                        try {
                            FileDelete(scrollPositionsFile)
                            Sleep(100)  ; Small delay to ensure file system updates
                        } catch {
                        }
                    }

                    ; Write all entries back in UTF-8 encoding
                    try {
                        ; Write UTF-8 BOM and section header
                        FileAppend("[Positions]`r`n", scrollPositionsFile, "UTF-8")
                        ; Write each entry
                        for key, value in existingEntries {
                            ; Escape special INI characters in key and value
                            escapedKey := StrReplace(key, "=", "`=")
                            escapedKey := StrReplace(escapedKey, ";", "`;")
                            escapedValue := StrReplace(value, "`n", "`;")
                            escapedValue := StrReplace(escapedValue, "`r", "")
                            FileAppend(escapedKey . "=" . escapedValue . "`r`n", scrollPositionsFile, "UTF-8")
                        }
                        saved := true
                    } catch {
                        ; Fallback to IniWrite if manual write fails
                        saved := IniWrite(scrollPercentage, scrollPositionsFile, "Positions", url)
                    }
                } catch Error as iniErr {
                    saved := false
                }

                if (saved) {
                    ; Update banner to show success
                    try {
                        StandardLoadingBar_Update("Scroll position saved!")
                        Sleep(1000)  ; Show success message for 1 second
                    } catch {
                    }
                    ; Re-enter fullscreen after successful save
                    Send("{F11}")
                    Sleep(300)
                    fullscreenRestored := true
                    return true
                } else {
                    ; INI save failed - show error message
                    try {
                        StandardLoadingBar_Update("Error: Save failed")
                        Sleep(2000)  ; Show error message
                    } catch {
                    }
                    ; Re-enter fullscreen even on failure
                    Send("{F11}")
                    Sleep(300)
                    fullscreenRestored := true
                    return false
                }
            }
        }
    } catch Error as err {
    } finally {
        ; Always hide the banner after save operation completes
        try {
            Sleep(500)  ; Brief delay before hiding
            StandardLoadingBar_Hide(0)
        } catch {
        }
        ; Re-enter fullscreen if we haven't already (e.g., if exception occurred or validation failed)
        if (!fullscreenRestored) {
            try {
                Send("{F11}")
                Sleep(300)
            } catch {
            }
        }
    }
    return false
}

; Restore previous scroll position from history
RestorePreviousWikipediaScrollPosition() {
    global g_WikipediaScrollHistory

    try {
        ; Check if Wikipedia window is currently active
        activeWindow := WinGetTitle("A")
        isChromeActive := WinActive("ahk_exe chrome.exe")
        hasWikipedia := InStr(activeWindow, "Wikipedia")
        if (!isChromeActive || !hasWikipedia) {
            return false
        }
    } catch Error as err {
        return false
    }

    ; If history is empty, try to fall back to INI file (for positions saved via activation restore or previous sessions)
    if (g_WikipediaScrollHistory.Length = 0) {
        ; Get current URL to load from INI
        try {
            url := GetWikipediaURLNormalized()
            if (url = "") {
                ; Show brief message that no history exists
                ShowCenteredOverlay_Utils("⚠ No previous scroll position found", 1500, BANNER_ACCENT_ERROR)
                return false
            }

            ; Load from INI file
            scrollPositionsFile := A_ScriptDir "\data\wikipedia_scroll_positions.ini"
            savedPercentage := IniRead(scrollPositionsFile, "Positions", url, "0")
            savedPercentageFloat := Float(savedPercentage)

            if (savedPercentageFloat > 0.0) {
                ; Found a saved position in INI, restore it using helper function
                return RestoreWikipediaScrollPosition(savedPercentageFloat,
                    "Restoring previous scroll position... Please wait")
            } else {
                ; No saved position found in INI either
                ShowCenteredOverlay_Utils("⚠ No previous scroll position found", 1500, BANNER_ACCENT_ERROR)
                return false
            }
        } catch Error as err {
            ; Show brief message that no history exists
            ShowCenteredOverlay_Utils("⚠ No previous scroll position found", 1500, BANNER_ACCENT_ERROR)
            return false
        }
    }

    ; Get current URL to match with history
    try {
        url := GetWikipediaURLNormalized()
        ; #region agent log
        try {
            FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' . A_TickCount .
            ',"location":"Shift keys.ahk:3415","message":"Got normalized URL in restore","data":{"url":"' . url .
            '","urlLength":' . StrLen(url) . '},"sessionId":"debug-session","runId":"post-fix","hypothesisId":"F"}`n',
            DEBUG_LOG_PATH
        } catch {
        }
        ; #endregion
        if (url = "") {
            return false
        }

        ; Create UIA_Browser for checking scroll differences
        uia := UIA_Browser("ahk_exe chrome.exe")
    } catch Error as err {
        return false
    }

    ; Find the most recent previous position (not the current one)
    previousPosition := 0
    foundIndex := 0
    ; Search backwards through history to find a different position
    loop g_WikipediaScrollHistory.Length {
        idx := g_WikipediaScrollHistory.Length - A_Index + 1
        historyItem := g_WikipediaScrollHistory[idx]
        ; Check if this is a different position (different URL or different scroll percentage)
        if (historyItem.url != url) {
            ; Different article, use this one
            previousPosition := historyItem
            foundIndex := idx
            break
        } else {
            ; Same article, check if scroll position is different
            currentScroll := uia.JSReturnThroughClipboard("window.pageYOffset")
            docHeight := uia.JSReturnThroughClipboard("document.documentElement.scrollHeight")
            if (currentScroll != "" && currentScroll != "undefined" && currentScroll != "null" &&
                docHeight != "" && docHeight != "undefined" && docHeight != "null") {
                currentScrollFloat := Float(currentScroll)
                docHeightFloat := Float(docHeight)
                if (docHeightFloat > 0) {
                    currentPercentage := currentScrollFloat / docHeightFloat
                    diff := Abs(currentPercentage - historyItem.scrollPercentage)
                    ; If the saved percentage is different from current, use it
                    if (diff > 0.01) {
                        previousPosition := historyItem
                        foundIndex := idx
                        break
                    }
                }
            }
        }
    }

    if (!previousPosition) {
        ; No different position found in history
        ShowCenteredOverlay_Utils("⚠ No previous scroll position found", 1500, BANNER_ACCENT_ERROR)
        return false
    }

    ; Restore scroll position using helper function
    success := RestoreWikipediaScrollPosition(previousPosition.scrollPercentage,
        "Restoring previous scroll position... Please wait")

    ; Remove the restored position from history (since we just used it) and all positions after it
    if (success && foundIndex > 0 && foundIndex <= g_WikipediaScrollHistory.Length) {
        ; Remove from foundIndex to end
        loop (g_WikipediaScrollHistory.Length - foundIndex + 1) {
            g_WikipediaScrollHistory.Pop()
        }
    }
    return true
}

#HotIf

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

#HotIf IsChromePdfViewerActive()

ChromePdf_GetViewerRoot(uia) {
    ; Prefer the extension's RootWebArea (most stable for the PDF viewer UI)
    root := 0
    try root := uia.FindElement({ Type: 50030, Value: "chrome-extension://mhjfbmdgcfjbbpaeojofohoefgiehjai",
        matchmode: "Substring" })
    if (root)
        return root

    ; Fallbacks
    try root := uia.GetCurrentDocumentElement()
    if (root)
        return root
    try root := uia.BrowserElement
    return root
}

ChromePdf_ClickByAutomationId(automationId, fallbackNames := 0) {
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 80

        root := ChromePdf_GetViewerRoot(uia)
        if (!root)
            return false

        btn := 0
        try btn := root.FindFirst({ Type: 50000, AutomationId: automationId })
        if (!btn)
            try btn := root.FindFirst({ AutomationId: automationId })

        if (!btn && IsObject(fallbackNames)) {
            for , name in fallbackNames {
                try btn := root.FindFirst({ Type: 50000, Name: name })
                if (btn)
                    break
            }
        }

        if (btn) {
            try btn.Invoke()
            catch {
                try btn.Click()
            }
            return true
        }
    } catch {
    }
    return false
}

; PDF toolbar: two buttons share AutomationId "save" (Save to Google Drive vs Download). Never use FindFirst(save) alone.
ChromePdf_ClickDownload() {
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 80

        root := ChromePdf_GetViewerRoot(uia)
        if (!root)
            return false

        downloadNames := ["Baixar", "Download"]
        btn := 0
        for , name in downloadNames {
            try btn := root.FindFirst({ Type: 50000, Name: name, cs: false })
            if (btn)
                break
        }

        if (!btn) {
            try saves := root.FindAll({ Type: 50000, AutomationId: "save" })
            if (IsObject(saves)) {
                for , cand in saves {
                    n := ""
                    try n := cand.Name
                    if (n = "")
                        continue
                    if InStr(StrLower(n), "drive")
                        continue
                    btn := cand
                    break
                }
            }
        }

        if (btn) {
            try btn.Invoke()
            catch {
                try btn.Click()
            }
            return true
        }
    } catch {
    }
    return false
}

ChromePdf_FocusByAutomationId(automationId, controlType := 0) {
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 80

        root := ChromePdf_GetViewerRoot(uia)
        if (!root)
            return false

        el := 0
        if (controlType) {
            try el := root.FindFirst({ Type: controlType, AutomationId: automationId })
        }
        if (!el)
            try el := root.FindFirst({ AutomationId: automationId })

        if (el) {
            try el.SetFocus()
            catch {
                try el.Click()
            }
            return true
        }
    } catch {
    }
    return false
}

ChromePdf_TogglePresentMode() {
    ; Deterministic primary path: open More actions and invoke Present item by selector.
    ; Legacy directional-key fallback remains optional behind feature flag.
    global USE_CHROME_PDF_PRESENT_FALLBACK

    if !ChromePdf_ClickByAutomationId("more", ["More actions"])
        return false

    deadline := A_TickCount + 700
    selectorNames := [
        "Present",
        "Presentation mode",
        "Present mode",
        "Apresentar",
        "Modo de apresentação"
    ]

    while (A_TickCount <= deadline) {
        try {
            uia := UIA_Browser("ahk_exe chrome.exe")
            root := ChromePdf_GetViewerRoot(uia)
            if (root) {
                presentItem := 0

                ; Prefer stable attributes first, then localized names.
                try presentItem := root.FindFirst({ Type: 50011, AutomationId: "present" })
                if (!presentItem)
                    try presentItem := root.FindFirst({ AutomationId: "present" })
                if (!presentItem)
                    try presentItem := root.FindFirst({ Type: 50000, AutomationId: "present" })

                if (!presentItem) {
                    for , candidateName in selectorNames {
                        try presentItem := root.FindFirst({ Type: 50011, Name: candidateName })
                        if (presentItem)
                            break
                        try presentItem := root.FindFirst({ Type: 50000, Name: candidateName })
                        if (presentItem)
                            break
                    }
                }

                if (presentItem) {
                    try presentItem.Invoke()
                    catch {
                        try presentItem.Click()
                    }
                    AgentDebugLog("H8", "ShiftE_presentmode_selector_success")
                    return true
                }
            }
        } catch {
        }
        Sleep 40
    }

    AgentDebugLog("H9", "ShiftE_presentmode_selector_timeout")

    if (USE_CHROME_PDF_PRESENT_FALLBACK) {
        Sleep 120
        Send "{Up}"
        Send "{Up}"
        Sleep 40
        Send "{Enter}"
        AgentDebugLog("H10", "ShiftE_presentmode_fallback_used")
        return true
    }

    return false
}

; Shift + F : Fit to page (Zoom to Fit) - Fit
+f::
{
    ; UIA tree: AutomationId "fit"
    ; #region agent log
    AgentDebugLog("H5", "ShiftF_handler_enter")
    ; #endregion
    result := ChromePdf_ClickByAutomationId("fit")
    if (result) {
        ; #region agent log
        AgentDebugLog("H6", "ShiftF_handler_click_success")
        ; #endregion
    } else {
        ; #region agent log
        AgentDebugLog("H7", "ShiftF_handler_click_failed")
        ; #endregion
    }
}

; Shift + P : Focus page number field - Page
+p::
{
    ; UIA tree: Edit AutomationId "pageSelector"
    ; Per requirement: focus only (no select-all)
    ChromePdf_FocusByAutomationId("pageSelector", 50004)
}

; Shift + T : Toggle thumbnails sidebar - Thumbnails
+t::
{
    ; UIA tree: AutomationId "sidenavToggle"
    ChromePdf_ClickByAutomationId("sidenavToggle")
}

; Shift + D : Download PDF - Download
+d::
{
    ; Toolbar duplicates AutomationId "save" (Drive save vs file download); use ChromePdf_ClickDownload.
    ChromePdf_ClickDownload()
}

; Shift + 2 : Two-page view (mnemonic: 2 = two pages)
+2::
{
    ; UIA: Button Type 50000, Name "More actions", AutomationId "more"
    if ChromePdf_ClickByAutomationId("more", ["More actions"]) {
        Sleep 150
        Send "{Down}"
        Sleep 50
        Send "{Enter}"
    }
}

; Shift + E : Present mode (mnemonic: E from prEsent)
+E::
{
    ChromePdf_TogglePresentMode()
}

#HotIf

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

#HotIf IsMercadoLivreActive()

; Shift + S: Focus Mercado Livre search field
+s::
{
    ML_EnsureHotkeyReceptivity()
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 200

        ; Prefer the document root; fall back to browser element only if needed
        isDocRoot := false
        try {
            root := uia.GetCurrentDocumentElement()
            isDocRoot := true
        } catch {
            root := uia.BrowserElement
        }

        field := 0

        ; 1) Try AutomationId from the current document (or fallback root)
        try {
            field := root.FindElement({ AutomationId: "cb1-edit" })
        } catch {
        }

        ; 2) From the document root, try the numeric path 1,1,4,2 if available
        if (!field && isDocRoot) {
            try {
                field := root.ElementFromPath("1,1,4,2")
            } catch {
            }
        }

        ; 3) As a last resort, search from the browser element with Descendants scope
        if (!field) {
            try {
                field := uia.BrowserElement.FindElement({ AutomationId: "cb1-edit" }, UIA.TreeScope.Descendants)
            } catch {
            }
        }

        if (field) {
            focusOk := false
            try {
                field.SetFocus()
                focusOk := true
            } catch {
                try {
                    field.Click()
                    focusOk := true
                } catch {
                }
            }
            if (focusOk)
                return
        }

        ; Error-driven workaround: force right-click + close menu, then retry once
        ML_EnsureHotkeyReceptivity(true)
        Sleep 300

        field := 0
        try {
            try
                field := root.FindElement({ AutomationId: "cb1-edit" })
            catch
                field := 0
            if (!field && isDocRoot) {
                try
                    field := root.ElementFromPath("1,1,4,2")
                catch
                    field := 0
            }
            if (!field) {
                try
                    field := uia.BrowserElement.FindElement({ AutomationId: "cb1-edit" }, UIA.TreeScope.Descendants)
                catch
                    field := 0
            }
            if (field) {
                try
                    field.SetFocus()
                catch {
                    try
                        field.Click()
                    catch {
                    }
                }
                return
            }
        } catch {
        }

        MsgBox "Could not find Mercado Livre search field."
    } catch Error as e {
        MsgBox "An error occurred: " e.Message
    }
}

; Shift + C: Carrinho de compras (Cart)
+c::
{
    ML_EnsureHotkeyReceptivity()
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 200
        try {
            root := uia.GetCurrentDocumentElement()
        } catch {
            root := uia.BrowserElement
        }

        cart := 0
        ; Prefer AutomationId
        try {
            cart := root.FindElement({ AutomationId: "nav-cart" })
        } catch {
        }
        if (!cart) {
            ; Try by class name substring
            try {
                cart := root.FindElement({ ClassName: "nav-cart", matchmode: "Substring" })
            } catch {
            }
        }
        if (!cart) {
            ; Try by link name containing 'carrinho'
            try {
                cart := root.FindElement({ Type: 50005, Name: "carrinho", cs: false, matchmode: "Substring" })
            } catch {
            }
        }

        if (cart) {
            try cart.Invoke()
            catch {
                try cart.Click()
            }
            return
        }
        MsgBox "Could not find Mercado Livre cart link."
    } catch Error as e {
        MsgBox "An error occurred: " e.Message
    }
}

; Shift + P: Compras (Purchases)
+p::
{
    ML_EnsureHotkeyReceptivity()
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 200
        try {
            root := uia.GetCurrentDocumentElement()
        } catch {
            root := uia.BrowserElement
        }

        purchases := 0
        ; Try by class name first
        try {
            purchases := root.FindElement({ ClassName: "option-purchases" })
        } catch {
        }
        if (!purchases) {
            ; Try by link name 'Compras'
            try {
                purchases := root.FindElement({ Type: 50005, Name: "Compras", cs: false, matchmode: "Substring" })
            } catch {
            }
        }

        if (purchases) {
            try purchases.Invoke()
            catch {
                try purchases.Click()
            }
            return
        }
        MsgBox "Could not find Mercado Livre purchases link."
    } catch Error as e {
        MsgBox "An error occurred: " e.Message
    }
}

; Shift + Y: Chegará amanhã (filter toggle)
+y::
{
    ML_EnsureHotkeyReceptivity()
    if ML_FindAndInvoke([{ Type: 50000, AutomationId: "shipping_time_highlighted_nextday" }])
        return
    MsgBox "Filtro 'Chegará amanhã' não encontrado."
}

; Shift + F: Full (frete grátis Full)
+f::
{
    ML_EnsureHotkeyReceptivity()
    if ML_FindAndInvoke([{ Type: 50000, AutomationId: "shipping_highlighted_fulfillment" }])
        return
    MsgBox "Filtro 'Full' não encontrado."
}

; Shift + I: Compra Internacional
+i::
{
    ML_EnsureHotkeyReceptivity()
    if ML_FindAndInvoke([{ Type: 50000, AutomationId: "SHIPPING_ORIGIN_HIGHLIGHTED" }])
        return
    MsgBox "Filtro 'Internacional' não encontrado."
}

; Shift + N: Envio local / Produtos com frete nacional
+n::
{
    ML_EnsureHotkeyReceptivity()
    if ML_FindAndInvoke([{ Type: 50000, AutomationId: "SHIPPING_ORIGIN_LOCAL_HIGHLIGHTED" }, { Type: 50000, Name: "Envio local",
        cs: false, matchmode: "Substring" }])
        return
    MsgBox "Filtro 'Produtos com frete nacional' não encontrado."
}

; Shift + G: Frete grátis
+g::
{
    ML_EnsureHotkeyReceptivity()
    if ML_FindAndInvoke([{ Type: 50000, AutomationId: "shipping_cost_highlighted_free" }])
        return
    MsgBox "Filtro 'Frete grátis' não encontrado."
}

; Shift + O: Ordenar por – Handy-style menu and UIA execution
; Returns current option label from list selected item (Layer 1) or trigger (Layer 2). Single authority for quality-check reads.
ML_SortGetCurrentLabel(sortList, trigger) {
    currentLabel := ""
    if (sortList) {
        try {
            children := sortList.FindAll({ Type: 50007 })
            for ch in children {
                try {
                    if (ch.SelectionItemPattern.IsSelected) {
                        currentLabel := Trim(ch.Name)
                        if (currentLabel = "")
                            currentLabel := Trim(ch.Value)
                        if (currentLabel != "")
                            return currentLabel
                    }
                } catch {
                }
            }
        } catch {
        }
    }
    try {
        currentLabel := Trim(trigger.Name)
        if (currentLabel = "")
            currentLabel := Trim(trigger.Value)
    } catch {
    }
    return currentLabel
}

ML_SortClose() {
    try Hotkey("1", "Off")
    catch {
    }
    try Hotkey("2", "Off")
    catch {
    }
    try Hotkey("3", "Off")
    catch {
    }
    try Hotkey("Escape", ML_SortCancel, "Off")
    catch {
    }
    global g_ML_SortGui
    if (g_ML_SortGui && IsObject(g_ML_SortGui) && g_ML_SortGui.Hwnd)
        try g_ML_SortGui.Destroy()
    g_ML_SortGui := 0
}

ML_SortCancel(*) {
    ML_SortClose()
}

ML_SortSelect(idx) {
    ML_SortClose()
    ML_SortApply(idx)
}

ML_SortApply(idx) {
    StandardLoadingBar_Show("⏳ Ordenando...", BANNER_ACCENT_INTERMEDIATE, { passive: false, centerOnHwnd: 0,
        textWidth: 380, fontSize: 17 })
    try {
        root := ML_GetDocRoot()
        if (!root) {
            StandardLoadingBar_Update("❌ Página do Mercado Livre não disponível.")
            Sleep 800
            return
        }
        trigger := ML_Find(root, { Type: 50003, AutomationId: "5clcjae", matchmode: "Substring" })
        if (!trigger)
            trigger := ML_Find(root, { Type: 50003, AutomationId: "trigger", matchmode: "Substring" })
        if (!trigger) {
            StandardLoadingBar_Update("❌ Botão 'Ordenar por' não encontrado.")
            Sleep 800
            return
        }
        ; Current selection: menu opens with this option highlighted. Order in menu: 1=Mais relevantes, 2=Menor preço, 3=Maior preço
        current := 1
        label := ""
        try {
            label := Trim(trigger.Name)
            if (label = "")
                label := Trim(trigger.Value)
            if (InStr(label, "Maior preço"))
                current := 3
            else if (InStr(label, "Menor preço"))
                current := 2
            else if (InStr(label, "Mais relevantes"))
                current := 1
        } catch {
        }
        StandardLoadingBar_Update("⏳ Abrindo menu...")
        clickOk := false
        try {
            trigger.Click()
            clickOk := true
        } catch {
        }
        if (!clickOk) {
            StandardLoadingBar_Update("❌ Não foi possível abrir o menu.")
            Sleep 800
            return
        }
        ; Bounded wait for menu: poll for sortList up to 1.2s (canon: condition-based waits with timeout).
        searchRoot := root
        sortList := 0
        loop 12 {
            Sleep 100
            try {
                dropdownSibling := UIA.TreeWalkerTrue.TryGetNextSiblingElement(trigger)
                if (dropdownSibling)
                    sortList := ML_Find(dropdownSibling, { Type: 50008, AutomationId: "menu-list", matchmode: "Substring" })
            } catch {
            }
            if (sortList)
                break
            sortList := ML_Find(searchRoot, { Type: 50008, AutomationId: "5clcjae_-menu-list", matchmode: "Substring" })
            if (sortList)
                break
            sortList := ML_Find(searchRoot, { Type: 50008, AutomationId: "menu-list", matchmode: "Substring" })
            if (sortList)
                break
            try {
                triggerParent := trigger.Parent
                if (triggerParent)
                    sortList := ML_Find(triggerParent, { Type: 50008, AutomationId: "menu-list", matchmode: "Substring" })
            } catch {
            }
            if (sortList)
                break
            sortList := ML_Find(searchRoot, { AutomationId: "menu-list", matchmode: "Substring" })
            if (sortList)
                break
        }
        StandardLoadingBar_Update("⏳ Selecionando opção...")
        optionSubstrings := ["menu-list-option-relevance", "menu-list-option-price_asc", "menu-list-option-price_desc"]
        sub := optionSubstrings[idx]
        optionNames := ["Mais relevantes", "Menor preço", "Maior preço"]
        item := 0
        if (sortList) {
            item := ML_Find(sortList, { Type: 50007, AutomationId: sub, matchmode: "Substring" })
            if (!item)
                item := ML_Find(sortList, { Type: 50007, Name: optionNames[idx], cs: false })
        }
        if (!item) {
            item := ML_Find(searchRoot, { Type: 50007, AutomationId: sub, matchmode: "Substring" })
            if (!item)
                item := ML_Find(searchRoot, { Type: 50007, Name: optionNames[idx], cs: false })
        }
        if (item) {
            try item.Click()
            catch
                try item.Invoke()
            Sleep 150
            StandardLoadingBar_Update("✅ Ordenação aplicada")
            Sleep 300
        } else {
            ; Keyboard: navigate from current to target (delta), then confirm with Enter after label check.
            try WinActivate("ahk_exe chrome.exe")
            catch {
            }
            Sleep 400
            delta := idx - current
            if (delta > 0) {
                loop delta {
                    Send "{Down}"
                    Sleep 100
                }
            } else if (delta < 0) {
                loop (-delta) {
                    Send "{Up}"
                    Sleep 100
                }
            }
            Sleep 200
            targetName := optionNames[idx]
            currentLabel := ML_SortGetCurrentLabel(sortList, trigger)
            labelMatches := (currentLabel != "" && (InStr(currentLabel, targetName) || InStr(targetName, currentLabel)))
            if (labelMatches) {
                Send "{Enter}"
                StandardLoadingBar_Update("✅ Ordenação aplicada")
                Sleep 300
            } else {
                currentIdx := 1
                if (InStr(currentLabel, "Maior preço"))
                    currentIdx := 3
                else if (InStr(currentLabel, "Menor preço"))
                    currentIdx := 2
                else if (InStr(currentLabel, "Mais relevantes"))
                    currentIdx := 1
                if (idx > currentIdx) {
                    Send "{Down}"
                    Sleep 100
                } else if (idx < currentIdx) {
                    Send "{Up}"
                    Sleep 100
                }
                currentLabel := ML_SortGetCurrentLabel(sortList, trigger)
                labelMatches := (currentLabel != "" && (InStr(currentLabel, targetName) || InStr(targetName,
                    currentLabel)))
                Send "{Enter}"
                StandardLoadingBar_Update("✅ Ordenação aplicada")
                Sleep 300
            }
        }
    } catch Error as err {
        try StandardLoadingBar_Update("❌ Erro: " SubStr(err.Message, 1, 40))
        catch {
        }
        Sleep 700
    } finally {
        try StandardLoadingBar_Hide(0)
        catch {
        }
    }
}

+o::
{
    ML_EnsureHotkeyReceptivity()
    root := ML_GetDocRoot()
    if (!root) {
        ; Retry once after workaround and delay (document may not have been ready on first load)
        ML_EnsureHotkeyReceptivity(true)
        Sleep 400
        root := ML_GetDocRoot()
    }
    if (!root) {
        MsgBox "Página do Mercado Livre não disponível."
        return
    }
    global g_ML_SortGui
    g_ML_SortGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Owner")
    g_ML_SortGui.BackColor := "1E1E2E"
    g_ML_SortGui.MarginX := 20
    g_ML_SortGui.MarginY := 15
    g_ML_SortGui.SetFont("s14 cCDD6F4 Bold", "Segoe UI")
    g_ML_SortGui.Add("Text", "w280 Center", "Ordenar por")
    g_ML_SortGui.Add("Text", "w280 h1 Background45475A")
    g_ML_SortGui.SetFont("s12 cCDD6F4", "Segoe UI")
    g_ML_SortGui.Add("Text", "w280", "[1] Mais relevantes")
    g_ML_SortGui.SetFont("s9 c6C7086", "Segoe UI")
    g_ML_SortGui.Add("Text", "w280 y+2", "    Relevância da busca")
    g_ML_SortGui.SetFont("s12 cCDD6F4", "Segoe UI")
    g_ML_SortGui.Add("Text", "w280", "[2] Menor preço")
    g_ML_SortGui.SetFont("s9 c6C7086", "Segoe UI")
    g_ML_SortGui.Add("Text", "w280 y+2", "    Preço crescente")
    g_ML_SortGui.SetFont("s12 cCDD6F4", "Segoe UI")
    g_ML_SortGui.Add("Text", "w280", "[3] Maior preço")
    g_ML_SortGui.SetFont("s9 c6C7086", "Segoe UI")
    g_ML_SortGui.Add("Text", "w280 y+2", "    Preço decrescente")
    g_ML_SortGui.SetFont("s12 cCDD6F4", "Segoe UI")
    g_ML_SortGui.Add("Text", "w280 h1 Background45475A y+10")
    g_ML_SortGui.SetFont("s9 c6C7086", "Segoe UI")
    g_ML_SortGui.Add("Text", "w280 Center", "Press 1-3 | Esc to cancel")
    activeWin := 0
    try
        activeWin := WinGetID("A")
    catch
        activeWin := 0
    MonitorGetWorkArea(1, &monitorLeft, &monitorTop, &monitorRight, &monitorBottom)
    monitorWidth := monitorRight - monitorLeft
    monitorHeight := monitorBottom - monitorTop
    if (activeWin && activeWin != 0) {
        rect := Buffer(16, 0)
        if (DllCall("GetWindowRect", "ptr", activeWin, "ptr", rect)) {
            centerX := NumGet(rect, 0, "int") + (NumGet(rect, 8, "int") - NumGet(rect, 0, "int")) // 2
            centerY := NumGet(rect, 4, "int") + (NumGet(rect, 12, "int") - NumGet(rect, 4, "int")) // 2
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
    g_ML_SortGui.Show("AutoSize Hide")
    g_ML_SortGui.GetPos(&gx, &gy, &gw, &gh)
    cx := monitorLeft + (monitorWidth - gw) // 2
    cy := monitorTop + (monitorHeight - gh) // 2
    g_ML_SortGui.Show("x" . cx . " y" . cy . " NA")
    Hotkey("1", (*) => ML_SortSelect(1), "On")
    Hotkey("2", (*) => ML_SortSelect(2), "On")
    Hotkey("3", (*) => ML_SortSelect(3), "On")
    Hotkey("Escape", ML_SortCancel, "On")
}

; Shift + L: Paginação – Seguinte
+l::
{
    ML_EnsureHotkeyReceptivity()
    if ML_FindAndInvoke([{ Type: 50005, Name: "Seguinte", cs: false }, { Type: 50000, Name: "Seguinte", cs: false }])
        return
    MsgBox "Botão 'Seguinte' não encontrado."
}

; Shift + K: Paginação – Anterior
+k::
{
    ML_EnsureHotkeyReceptivity()
    if ML_FindAndInvoke([{ Type: 50005, Name: "Anterior", cs: false }, { Type: 50000, Name: "Anterior", cs: false }])
        return
    MsgBox "Botão 'Anterior' não encontrado."
}

; Shift + A: Adicionar ao carrinho (página do produto)
+a::
{
    ML_EnsureHotkeyReceptivity()
    if ML_FindAndInvoke([{ Type: 50000, Name: "Adicionar ao carrinho", cs: false }])
        return
    MsgBox "Botão 'Adicionar ao carrinho' não encontrado."
}

; Shift + V: Adicionar aos favoritos (coração)
+v::
{
    ML_EnsureHotkeyReceptivity()
    if ML_FindAndInvoke([{ Type: 50000, Name: "Adicionar aos favoritos", cs: false }, { Type: 50000, ClassName: "ui-pdp-bookmark",
        matchmode: "Substring" }])
        return
    MsgBox "Botão 'Adicionar aos favoritos' não encontrado."
}

; Shift + J: Continuar fluxo (Continuar a compra / Continuar / OK)
+j::
{
    ML_EnsureHotkeyReceptivity()
    conditions := [{ Type: 50005, Name: "Continuar a compra", cs: false }, { Type: 50000, Name: "Continuar", cs: false }, { Type: 50000,
        Name: "OK", cs: false }, { Type: 50000, AutomationId: "shipping_footer_confirm_button" }, { Type: 50005, Name: "Continuar",
            cs: false }, { Type: 50000, Name: "Seguinte", cs: false }
    ]
    if ML_FindAndInvoke(conditions)
        return
    MsgBox "Botão de continuar não encontrado."
}

;-------------------------------------------------------------------
; Shopee (Brazil) Shortcuts
;-------------------------------------------------------------------
#HotIf IsShopeeActive()

; Shift + S: Focus Shopee search field
+s::
{
    try {
        root := Shopee_GetDocRoot()
        if (!root) {
            MsgBox "Página da Shopee não disponível."
            return
        }

        field := 0
        ; Prefer the main search combo box
        try {
            field := root.FindElement({ Type: 50003, Name: "Buscar na Shopee" })
        } catch {
        }
        ; Fallback: any control with LocalizedType = "search"
        if (!field) {
            try field := root.FindElement({ LocalizedType: "search" })
        }
        ; Fallback: numeric path from document root (see shopping uia3.md)
        if (!field) {
            try field := root.ElementFromPathExist("1,1,2,3,1")
        }

        if (field) {
            try field.SetFocus()
            catch {
                try field.Click()
            }
            return
        }
        MsgBox "Campo de busca da Shopee não encontrado."
    } catch Error as e {
        MsgBox "Erro ao focar busca da Shopee: " e.Message
    }
}

; Shift + C: Carrinho de compras (cart)
+c::
{
    if Shopee_FindAndInvoke([{ Type: 50005, AutomationId: "cart_drawer_target_id" }, { Type: 50005, Name: "Carrinho",
        cs: false, matchmode: "Substring" }, { Type: 50000, Name: "Carrinho", cs: false, matchmode: "Substring" }
    ])
        return
    MsgBox "Link/botão de carrinho da Shopee não encontrado."
}

; Shift + P: Minhas compras / pedidos (speculative)
+p::
{
    if Shopee_FindAndInvoke([{ Type: 50005, Name: "Minhas compras", cs: false, matchmode: "Substring" }, { Type: 50005,
        Name: "Meus pedidos", cs: false, matchmode: "Substring" }, { Type: 50005, Name: "Pedidos", cs: false, matchmode: "Substring" }
    ])
        return
    MsgBox "Link de compras/pedidos da Shopee não encontrado (atalho especulativo)."
}

; Shift + Y: Entrega Rápida (Chegará amanhã analog)
+y::
{
    if Shopee_FindAndInvoke([{ Type: 50002, Name: "Entrega Rápida", cs: false, matchmode: "Substring" }])
        return
    MsgBox "Filtro 'Entrega Rápida' não encontrado (atalho especulativo)."
}

; Shift + F: Promoções / produtos com desconto (Full analog, speculative)
+f::
{
    if Shopee_FindAndInvoke([{ Type: 50002, Name: "Produtos com Desconto", cs: false, matchmode: "Substring" }])
        return
    MsgBox "Filtro de promoções/produtos com desconto não encontrado (atalho especulativo)."
}

; Shift + I: Compra internacional
+i::
{
    if Shopee_FindAndInvoke([{ Type: 50002, Name: "Internacional", cs: false }])
        return
    MsgBox "Filtro 'Internacional' da Shopee não encontrado."
}

; Shift + N: Envio nacional
+n::
{
    if Shopee_FindAndInvoke([{ Type: 50002, Name: "Nacional", cs: false }])
        return
    MsgBox "Filtro 'Nacional' da Shopee não encontrado."
}

; Shift + G: Frete grátis (speculative)
+g::
{
    if Shopee_FindAndInvoke([{ Type: 50020, Name: "Frete grátis", cs: false, matchmode: "Substring" }])
        return
    MsgBox "Indicador/controle de 'Frete grátis' não encontrado (atalho especulativo)."
}

; Shift + O: Ordenar por (open sort menu)
+o::
{
    if Shopee_FindAndInvoke([{ Type: 50000, Name: "Classificar por relevância", cs: false, matchmode: "Substring" }, { Type: 50000,
        Name: "Classificar por", cs: false, matchmode: "Substring" }])
        return
    MsgBox "Botão 'Classificar por' da Shopee não encontrado."
}

; Shift + R: Faixa de preço (focus min/max price edits)
+r::
{
    try {
        root := Shopee_GetDocRoot()
        if (!root) {
            MsgBox "Página da Shopee não disponível."
            return
        }

        minimo := 0
        try minimo := root.FindElement({ Type: 50004, Name: "Preço mínimo" })
        catch {
        }
        maximo := 0
        try maximo := root.FindElement({ Type: 50004, Name: "Preço máximo" })
        catch {
        }

        target := minimo ? minimo : maximo
        if (target) {
            try target.SetFocus()
            catch {
                try target.Click()
            }
            Sleep 50
            Send "^a"
            return
        }

        MsgBox "Campos de faixa de preço da Shopee não encontrados."
    } catch Error as e {
        MsgBox "Erro ao focar faixa de preço da Shopee: " e.Message
    }
}

; Shift + L: Paginação – Seguinte (results)
+l::
{
    if Shopee_NavMove(1)
        return
    MsgBox "Navegação 'Seguinte' da Shopee não encontrada ou não aplicável."
}

; Shift + K: Paginação – Anterior (results)
+k::
{
    if Shopee_NavMove(-1)
        return
    MsgBox "Navegação 'Anterior' da Shopee não encontrada ou não aplicável."
}

; Shift + A: Adicionar ao carrinho (página do produto)
+a::
{
    if Shopee_FindAndInvoke([{ Type: 50000, Name: "Adicionar Ao Carrinho", cs: false, matchmode: "Substring" }])
        return
    MsgBox "Botão 'Adicionar Ao Carrinho' não encontrado na página da Shopee."
}

; Shift + V: Favoritar (coração)
+v::
{
    if Shopee_FindAndInvoke([{ Type: 50000, Name: "Favoritar", cs: false, matchmode: "Substring" }])
        return
    MsgBox "Botão de favoritos da Shopee não encontrado."
}

; Shift + J: Continuar fluxo (Continuar / Fazer pedido)
+j::
{
    if Shopee_FindAndInvoke([{ Type: 50000, Name: "Continuar", cs: false, matchmode: "Substring" }, { Type: 50000, Name: "Fazer pedido",
        cs: false, matchmode: "Substring" }, { Type: 50000, Name: "OK", cs: false }
    ])
        return
    MsgBox "Botão de continuar/fazer pedido da Shopee não encontrado."
}

#HotIf

;-------------------------------------------------------------------
; Microsoft Teams Shortcuts (chat)
;-------------------------------------------------------------------
#HotIf IsTeamsChatActive()

; Shift + K : Toolbar Back (UIA; fallback Alt+Left — see teams.md menur75)
+k::
{
    try {
        root := UIA.ElementFromHandle(WinExist("A"))
        if !root
            return
        backBtn := root.FindFirst({ Name: "Back", Type: "50000", AutomationId: "menur75" })
        if !backBtn
            backBtn := root.FindFirst({ Name: "Back", Type: "50000" })
        if (backBtn) {
            backBtn.Click()
            return
        }
    } catch {
    }
    Send "!{Left}"
}

; Shift + L : Toolbar Forward (UIA; fallback Alt+Right). Like reaction moved to Shift+Y.
+l::
{
    try {
        root := UIA.ElementFromHandle(WinExist("A"))
        if !root
            return
        fwdBtn := root.FindFirst({ Name: "Forward", Type: "50000" })
        if (fwdBtn) {
            fwdBtn.Click()
            return
        }
    } catch {
    }
    Send "!{Right}"
}

; Shift + R : Reply - Reply
+r::
{
    Send "{Enter}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Shift + U : View all unread items - Unread
+u::
{
    Send "^!u"
}

; Shift + P : Pin chat - Pin
+p::
{
    Sleep "150"
    Send "^1"
    Sleep "100"
    Send("{AppsKey}")
    Sleep "100"
    Send "{Down}"
    Send "{Down}"
    Send "{Right}"
    Send "{Enter}"
    SendEscape()
    Sleep "200"
    Send "^1"
    Sleep "500"          ; 80 ms
    Send "^+{Home}"
}

; Shift + E : Edit message - Edit
+e::
{
    Send "{Enter}"
    Sleep 100
    Send "{Left}"
    Sleep 100
    Send "{Left}"
    Sleep 100
    Send "{Enter}"
}

; Shift + A : Attach file - Attach
+a::
{
    Send "!+o"
}

; =============================================================================
; Chat: Select "Quick Views" tree item in the chat rail
; Hotkey: Shift+O
; UIA path (relative to window): 2,1,2,3,2,1,1,1,1,1,1,1,1,9,1,4,1
; Full tree path:                11,2,1,2,3,2,1,1,1,1,1,1,1,1,9,1,4,1
; Uses SelectionItemPattern.Select() — no mouse click.
; =============================================================================
+o::
{
    StandardLoadingBar_Show("⏳ Teams: Quick views…", BANNER_ACCENT_INTERMEDIATE, { centerOnHwnd: WinExist("A") })
    try {
        ; Create a new chat
        Send "^n"
        Sleep 500

        hwnd := 0
        for proc in TEAMS_PROCESSES {
            for wnd in WinGetList("ahk_exe " proc) {
                if IsTeamsChatTitle(WinGetTitle(wnd)) {
                    hwnd := wnd
                    break 2
                }
            }
        }
        if !hwnd {
            ShowCenteredOverlay_Utils("⚠ NO TEAMS CHAT WINDOW", 2000, BANNER_ACCENT_ERROR)
            return
        }

        root := UIA.ElementFromHandle(hwnd)
        ; Path is relative to the window element (child 11 = Chrome_WidgetWin_1, the Chromium host pane)
        quickViews := root.ElementFromPathExist("11,2,1,2,3,2,1,1,1,1,1,1,1,1,9,1,4,1")
        if !quickViews
            quickViews := root.FindFirst(UIA.CreateCondition({ Name: "Quick views", ControlType: "TreeItem" }))
        if !quickViews {
            ShowCenteredOverlay_Utils("❌ QUICK VIEWS NOT FOUND", 2000, BANNER_ACCENT_ERROR)
            return
        }
        quickViews.SelectionItemPattern.Select()
    } catch as e {
        ShowCenteredOverlay_Utils("❌ QUICK VIEWS: " e.Message, 2000, BANNER_ACCENT_ERROR)
    } finally {
        StandardLoadingBar_Hide(0)
    }
}

; Shift + H : Open history menu - History
+h::
{
    Send "^h"
}

; Shift + M : Mark as unread - Mark
+m::
{
    Send "^1"
    Sleep "220"
    Send("{AppsKey}")
    Sleep "220"
    Send "{Down}"
    Send "{Enter}"
}

; Shift + X : Unpin chat - Unpin
+x::
{
    Sleep "150"
    Send "^1"
    Sleep "100"
    Send("{AppsKey}")
    Sleep "100"
    Send "r"
    Send "{Enter}"
}

; Shift + C : Collapse all conversation folders - Collapse
+c::
{
    Send "!q"
}

; Shift + I : Activate/deactivate details panel - Info
+i::
{
    Send "!p"
}

; Shift + . : Detach current chat - Window
+.::
{
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        moreOptionsButton := root.FindFirst({ Name: "More chat options", Type: "50000" })

        if moreOptionsButton {
            moreOptionsButton.Click()
            Sleep 350

            detachMenuItem := root.FindFirst({ Name: "Open in new window", Type: "50011" })

            if !detachMenuItem {
                detachMenuItem := UIA.GetRootElement().FindFirst({ Name: "Open in new window", Type: "50011" })
            }

            if detachMenuItem {
                detachMenuItem.Click()
            } else {
                ShowSmallLoadingIndicator_ChatGPT("Could not find Open in new window")
                SetTimer(() => HideSmallLoadingIndicator_ChatGPT(), -2000)
            }
        } else {
            ShowSmallLoadingIndicator_ChatGPT("Could not find more chat options")
            SetTimer(() => HideSmallLoadingIndicator_ChatGPT(), -2000)
        }
    }
    catch Error as e {
        ShowSmallLoadingIndicator_ChatGPT("Could not detach chat")
        SetTimer(() => HideSmallLoadingIndicator_ChatGPT(), -2000)
    }
}

; Shift + V : Video call - Video
+v::
{
    ; Show confirmation popup
    if MsgBox("Do you want to call this person?", "Confirm Call", "YesNo Icon?") = "Yes" {
        try {
            win := WinExist("A")
            root := UIA.ElementFromHandle(win)

            callButton := 0
            callButtonNames := ["Audio call", "Video call", "Start audio call", "Start video call"]

            for name in callButtonNames {
                candidates := root.FindAll({ Name: name, Type: "50000", matchmode: "Substring", cs: false })
                if candidates {
                    for candidate in candidates {
                        if !candidate.GetPropertyValue(UIA.Property.IsOffscreen) && candidate.GetPropertyValue(UIA.Property
                            .IsEnabled) {
                            callButton := candidate
                            break
                        }
                    }
                }
                if callButton
                    break
            }

            if !callButton {
                candidates := root.FindAll({ Type: "50000" })
                if candidates {
                    for candidate in candidates {
                        if InStr(StrLower(candidate.Name), "call") && !candidate.GetPropertyValue(UIA.Property.IsOffscreen
                        ) && candidate.GetPropertyValue(UIA.Property.IsEnabled) {
                            callButton := candidate
                            break
                        }
                    }
                }
            }

            if callButton {
                callButton.Click()
            } else {
                ; Show error banner
                ShowSmallLoadingIndicator_ChatGPT("Could not find call button")
                SetTimer(() => HideSmallLoadingIndicator_ChatGPT(), -2000)
            }
        }
        catch Error as e {
            ; Show error banner
            ShowSmallLoadingIndicator_ChatGPT("Could not find call button")
            SetTimer(() => HideSmallLoadingIndicator_ChatGPT(), -2000)
        }
    }
}

; Shift + T : Add participants - Team
+t::
{
    try {
        ; Show progress banner
        ShowSmallLoadingIndicator_ChatGPT("Adding participants...")

        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; First, find and click the "More chat options" button
        moreOptionsButton := root.FindFirst({ Name: "More chat options", Type: "50000", matchmode: "Substring" })

        if !moreOptionsButton {
            ; Show error banner
            ShowSmallLoadingIndicator_ChatGPT("Could not find more options button")
            SetTimer(() => HideSmallLoadingIndicator_ChatGPT(), -2000)
            return
        }

        moreOptionsButton.Click()
        Sleep 500  ; Wait for menu to open

        ; Now find and click the "View and add participants" button
        participantsButton := root.FindFirst({ Name: "View and add participants", Type: "50000", matchmode: "Substring" })

        if participantsButton {
            participantsButton.Click()
            Sleep 300
            Send "{Tab}"
            Sleep 300
            Send "{Enter}"
            ; Hide progress banner on success
            SetTimer(() => HideSmallLoadingIndicator_ChatGPT(), -1000)
        } else {
            ; Show error banner
            ShowSmallLoadingIndicator_ChatGPT("Could not find add participants button")
            SetTimer(() => HideSmallLoadingIndicator_ChatGPT(), -2000)
        }
    }
    catch Error as e {
        ; Show error banner
        ShowSmallLoadingIndicator_ChatGPT("Could not find add participants button")
        SetTimer(() => HideSmallLoadingIndicator_ChatGPT(), -2000)
    }
}

; Shift + F : Fold chat sections - Fold
+f::
{
    try {
        win := WinExist("A")
        if !win
            return

        root := UIA.ElementFromHandle(win)

        ; Narrow search to the chat navigation tree to speed up lookups.
        treeCond := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Tree)
        trees := ""
        try trees := root.FindElements(treeCond, UIA.TreeScope.Descendants)

        targetTree := ""
        targetIds := ["menur6a5", "menur6as", "menur6br", "menur6f6"]

        if trees {
            for candidate in trees {
                if !candidate
                    continue
                hasSection := false
                for id in targetIds {
                    if !id
                        continue
                    sectionEl := ""
                    try sectionEl := candidate.FindFirst({ AutomationId: id, Type: "50024" })
                    if sectionEl {
                        hasSection := true
                        break
                    }
                }
                if hasSection {
                    targetTree := candidate
                    break
                }
            }
        }

        if !targetTree
            targetTree := root

        ; Collect all expandable tree items (categories and chat groups).
        treeItemCond := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.TreeItem)
        canCollapseCond := UIA.CreatePropertyCondition(UIA.Property.IsExpandCollapsePatternAvailable, true)
        collapsibleCond := UIA.CreateAndCondition(treeItemCond, canCollapseCond)
        items := targetTree.FindElements(collapsibleCond, UIA.TreeScope.Descendants)

        if !items {
            ShowSmallLoadingIndicator_ChatGPT("No collapsible chat sections found")
            SetTimer(() => HideSmallLoadingIndicator_ChatGPT(), -2000)
            return
        }

        collapsed := 0
        already := 0
        total := 0

        for item in items {
            if !item
                continue
            total++
            try {
                pat := item.ExpandCollapsePattern
                if pat.ExpandCollapseState != UIA.ExpandCollapseState.Collapsed {
                    pat.Collapse()
                    collapsed++
                    Sleep 25
                } else {
                    already++
                }
            } catch Error {
                ; Best-effort fallback: focus and send Left to collapse.
                try {
                    item.SetFocus()
                    Sleep 40
                    Send "{Left}"
                    collapsed++
                } catch {
                }
            }
        }

        msg := ""
        if collapsed {
            msg := Format("Collapsed {} chat section{}", collapsed, collapsed = 1 ? "" : "s")
        } else if already && !collapsed {
            msg := "Chat sections already collapsed"
        } else {
            msg := "Nothing to collapse"
        }

        ShowSmallLoadingIndicator_ChatGPT(msg)
        SetTimer(() => HideSmallLoadingIndicator_ChatGPT(), -2000)
    }
    catch Error {
        ShowSmallLoadingIndicator_ChatGPT("Could not collapse chat sections")
        SetTimer(() => HideSmallLoadingIndicator_ChatGPT(), -2000)
    }

    Send "^{Home}"
    Sleep "200"
    Send "c"
    Send "{Right}"
    Sleep "100"
    Send "^{Home}"
    Sleep "200"
    Send "g"
    Send "{Right}"
    Send "^{Home}"
    Sleep "200"
    Send "f"
    Send "f"
    Send "{Right}"
}

; Shift + Y : Like reaction (moved from Shift+L — now Shift+L is Forward)
+y::
{
    Send "{Enter}"
    Send "{Enter}"
    SendEscape()
}

; Shift + G : Heart reaction - Heart
+g::
{
    Send "{Enter}"
    Send "{Down}"
    Send "{Enter}"
    SendEscape()
}

; Shift + J : Laugh reaction - Laugh
+j::
{
    Send "{Enter}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
    SendEscape()
}

; Alt + 1 : Select 1st search result - Search
!1::
{
    Send "{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Enter}"
}

; Alt + 2 : Select 2nd search result - Search
!2::
{
    Send "{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Enter}"
}

; Alt + 3 : Select 3rd search result - Search
!3::
{
    Send "{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Enter}"
}

; Alt + 4 : Select 4th search result - Search
!4::
{
    Send "{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Enter}"
}

; Alt + 5 : Select 5th search result - Search
!5::
{
    Send "{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Enter}"
}

#HotIf

;-------------------------------------------------------------------
; Outlook Shortcuts
;-------------------------------------------------------------------

IsOutlookMessageActive() {
    return (WinActive("ahk_exe OUTLOOK.EXE") || WinActive("ahk_exe olk.exe"))
    && RegExMatch(WinGetTitle("A"), "i) - Message \(")
}

IsOutlookAppointmentActive() {
    ; Classic Outlook inspector windows use titles like " - Appointment/Meeting/Event".
    ; New Outlook editors often use titles like "New event - Outlook" and run under OUTLOOK.EXE or olk.exe.
    if !(WinActive("ahk_exe OUTLOOK.EXE") || WinActive("ahk_exe olk.exe"))
        return false

    t := ""
    try t := WinGetTitle("A")
    if RegExMatch(t, "i)(Appointment|Meeting|Event)")
        return true

    ; New Outlook: detect by UIA presence of the title field.
    if IsNewOutlookActive() {
        try {
            root := UIA.ElementFromHandle(WinExist("A"))
            if root.FindFirst({ Name: "Add title", ControlType: "Edit" })
                return true
            if root.FindFirst({ Name: "Add title", Type: 50004 })
                return true
            if root.FindFirst({ AutomationId: "4100" })
                return true
        } catch {
        }
    }

    return false
}

IsOutlookReminderActive() {
    return (WinActive("ahk_exe OUTLOOK.EXE") || WinActive("ahk_exe olk.exe"))
    && RegExMatch(WinGetTitle("A"), "i)Reminders?")
}

IsOutlookComposeActive() {
    ; New Outlook compose runs inside the main window and doesn't match the classic " - Message (" title.
    if !IsNewOutlookActive()
        return false
    try {
        root := UIA.ElementFromHandle(WinExist("A"))
        ; Prefer compose-only anchors seen in outlook-mail.md compose capture.
        if root.FindFirst({ AutomationId: "popoutCompose" })
            return true
        if root.FindFirst({ AutomationId: "discardCompose" })
            return true
        if root.FindFirst({ AutomationId: "splitButton-ram0__primaryActionButton" }) ; Send
            return true
        ; Fallback: presence of the compose Subject edit (MSG_*_SUBJECT) + Message body edit
        if root.FindFirst({ Name: "Subject", ControlType: "Edit" }) && root.FindFirst({ Name: "Message body",
            ControlType: "Edit" })
            return true
    } catch {
    }
    return false
}

IsOutlookMainActive() {
    if !(WinActive("ahk_exe OUTLOOK.EXE") || WinActive("ahk_exe olk.exe"))
        return false
    t := ""
    cls := ""
    try t := WinGetTitle("A")
    try cls := WinGetClass("A")
    ; Exclude inspectors and reminders
    if RegExMatch(t, "i) - Message \(")
        return false
    if RegExMatch(t, "i)(Appointment|Meeting|Event)")
        return false
    if RegExMatch(t, "i)^New event")
        return false
    if RegExMatch(t, "i)Reminder")
        return false
    ; Prefer the New Outlook shell window.
    if (cls != "" && InStr(cls, "Outlook Host"))
        return true
    if (t != "" && InStr(t, " - Outlook"))
        return true
    return true
}

Outlook_ActivateMainWindow() {
    ; Bring the main Outlook shell (Mail/Calendar) to front.
    try {
        wins := WinGetList("ahk_class Outlook Host")
        for hwnd in wins {
            t := ""
            try t := WinGetTitle("ahk_id " hwnd)
            if RegExMatch(t, "i)^(Mail|Calendar) - .* - Outlook") {
                try WinActivate("ahk_id " hwnd)
                try WinWaitActive("ahk_id " hwnd, , 1)
                return hwnd
            }
        }
    } catch {
    }
    ; Fallback: best-effort activate any Outlook process window.
    try {
        if WinExist("ahk_exe olk.exe")
            return WinActivate("ahk_exe olk.exe")
    } catch {
    }
    try {
        if WinExist("ahk_exe OUTLOOK.EXE")
            return WinActivate("ahk_exe OUTLOOK.EXE")
    } catch {
    }
    return 0
}

Outlook_FocusMainSearch() {
    Outlook_ActivateMainWindow()
    return OutlookFocusFirst([{ AutomationId: "topSearchInput", ControlType: "ComboBox" }, { AutomationId: "topSearchInput" }])
}

Outlook_SwitchToMail() {
    Outlook_ActivateMainWindow()
    ; Left-rail toggle (see outlook-mail.md): Chromium tree via OutlookMail_* — not OutlookClickFirst (ElementFromHandle misses WebView2).
    return OutlookMail_ClickFirst([{ AutomationId: "ddea774c-382b-47d7-aab5-adc2139a802b", ControlType: "Button" }, { Name: "Mail",
        ControlType: "Button" }, { Name: "Mail", Type: 50000 }
    ])
}

Outlook_SwitchToCalendar() {
    Outlook_ActivateMainWindow()
    return OutlookMail_ClickFirst([{ AutomationId: "8cbeb86f-83e1-43b5-aaba-cd3514322f0b", ControlType: "Button" }, { Name: "Calendar",
        ControlType: "Button" }, { Name: "Calendar", Type: 50000 }
    ])
}

; Calendar toolbar: "Go to today April 6, 2026" (dynamic date; caledar.md). WebView2 — use Chromium root.
; Loading bar during UIA work: docs/standard_information_display.md (Show → Update → Hide).
OutlookCalendar_ClickGoToToday() {
    try {
        StandardLoadingBar_Show("⏳ Outlook: Go to today…", BANNER_ACCENT_INTERMEDIATE, { passive: false, centerOnHwnd: 0,
            textWidth: 480, fontSize: 17 })
    } catch {
    }
    ok := false
    try {
        Outlook_ActivateMainWindow()
        try StandardLoadingBar_Update("⏳ Outlook: Opening Calendar…")
        if !Outlook_SwitchToCalendar()
            return false
        Sleep 150
        try StandardLoadingBar_Update("⏳ Outlook: Finding Go to today…")
        el := OutlookMail_FindFirst([{ Name: "Go to today", matchmode: "Substring", ControlType: "Button" }, { Name: "Go to today",
            matchmode: "Substring", Type: 50000 }, { Name: "Today", ControlType: "Button" }, { Name: "Today", matchmode: "Substring",
                ControlType: "Button" }
        ])
        if !el
            return false
        try el.ScrollIntoView()
        Sleep 40
        try el.SetFocus()
        Sleep 50
        try el.Click()
        catch Error {
            try el.Invoke()
            catch Error {
                return false
            }
        }
        ok := true
    } finally {
        try StandardLoadingBar_Hide(0)
        catch {
        }
    }
    return ok
}

; Left-rail Copilot app (copilot.md). AutomationId may vary by tenant; name fallback "Copilot".
global OUTLOOK_COPILOT_RAIL_AID := "b5abf2ae-c16b-4310-8f8a-d3bcdb52f162"

; Win+Alt+Shift+L modal: same pattern as Utils ShowAiModelSelector (1–9 + Esc).
global g_OutlookCopilotSelectorGui := false
global g_OutlookCopilotSelectorActive := false
global g_OutlookCopilotEscPollPrev := false
global g_OutlookCopilotShortcuts := [{ name: "Toggle Copilot voice chat", desc: "Voice chat / new voice session" }, { name: "New chat",
    desc: "Start a new Copilot chat" }, { name: "Focus Copilot input", desc: "Message Copilot composer" }, { name: "Open navigation panel",
        desc: "Expand Copilot nav drawer" }, { name: "Work scope", desc: "Toggle Work (grounded) scope" }, { name: "Web scope",
            desc: "Toggle Web scope" }, { name: "Model / mode (Auto…)", desc: "Open mode / model menu" }, { name: "Temporary chat",
                desc: "Temporary chat session" }, { name: "Chats and more", desc: "History and more options" }
]

OutlookCopilot_FindRailButton(root) {
    if !root
        return ""
    searchRoot := root
    try {
        r := root.FindFirst({ Name: "left-rail-appbar", matchmode: "Substring", ControlType: "Group" })
        if !r
            r := root.FindFirst({ Name: "left-rail-appbar", matchmode: "Substring" })
        if r
            searchRoot := r
    } catch {
    }
    el := ""
    try el := searchRoot.FindFirst({ AutomationId: OUTLOOK_COPILOT_RAIL_AID, ControlType: "Button" })
    if !el
        try el := searchRoot.FindFirst({ AutomationId: OUTLOOK_COPILOT_RAIL_AID })
    if !el
        try el := searchRoot.FindFirst({ Name: "Copilot", ControlType: "Button" })
    if !el
        try el := root.FindFirst({ AutomationId: OUTLOOK_COPILOT_RAIL_AID, ControlType: "Button" })
    if !el
        try el := root.FindFirst({ Name: "Copilot", ControlType: "Button" })
    return el
}

; Bounded wait: main Outlook shell is foreground (efficiency-canon; Alt-Tab from other apps).
OutlookCopilot_WaitMainOutlookForeground(timeoutMs := 2200) {
    hwnd := 0
    try {
        for h in WinGetList("ahk_class Outlook Host") {
            t := ""
            try t := WinGetTitle("ahk_id " h)
            if RegExMatch(t, "i)^(Mail|Calendar) - .* - Outlook") {
                hwnd := h
                break
            }
        }
    } catch {
    }
    if !hwnd {
        Outlook_ActivateMainWindow()
        deadline := A_TickCount + timeoutMs
        while (A_TickCount < deadline) {
            try {
                p := WinGetProcessName("A")
                if (p = "OUTLOOK.EXE" || p = "olk.exe")
                    return true
            } catch {
            }
            Outlook_ActivateMainWindow()
            Sleep 55
        }
        return false
    }
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        try {
            if (WinGetID("A") = hwnd)
                return true
        } catch {
        }
        try WinActivate("ahk_id " hwnd)
        try WinWaitActive("ahk_id " hwnd, , 0.35)
        Sleep 45
    }
    return false
}

OutlookCopilot_VoiceChatCriteriaList() {
    return [{ Name: "Start a new voice chat", matchmode: "Substring", ControlType: "Button" }, { Name: "Start a new voice chat",
        matchmode: "Substring", Type: 50000 }, { Name: "voice chat", matchmode: "Substring", ControlType: "Button" }
    ]
}

OutlookCopilot_FindVoiceChatButtonInRoot(root) {
    if !root
        return ""
    for criteria in OutlookCopilot_VoiceChatCriteriaList() {
        try {
            el := root.FindFirst(criteria)
            if el
                return el
        } catch {
        }
    }
    return ""
}

; Quality gate: element exists in UIA tree, is on-screen, enabled, and has non-zero bounds (rendered + clickable).
OutlookCopilot_ElementIsRenderedAndClickable(el) {
    if !IsObject(el)
        return false
    try {
        if el.GetPropertyValue(UIA.Property.IsOffscreen)
            return false
    } catch {
        try {
            if el.IsOffscreen
                return false
        } catch {
        }
    }
    try {
        if !el.GetPropertyValue(UIA.Property.IsEnabled)
            return false
    } catch {
        try {
            if !el.IsEnabled
                return false
        } catch {
            return false
        }
    }
    try {
        rect := el.BoundingRectangle
        l := rect.l, t := rect.t, r := rect.r, b := rect.b
        if (Abs((r - l) * (b - t)) < 4)
            return false
    } catch {
    }
    return true
}

; Wait until voice chat button passes quality gates OR timeout (poll + optional ScrollIntoView when stuck offscreen).
OutlookCopilot_WaitVoiceChatButtonReady(timeoutMs := 12000, pollMs := 120) {
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        root := OutlookMail_RootElement()
        if root {
            el := OutlookCopilot_FindVoiceChatButtonInRoot(root)
            if el {
                if OutlookCopilot_ElementIsRenderedAndClickable(el)
                    return el
                try el.ScrollIntoView()
            }
        }
        Sleep pollMs
    }
    return ""
}

; Ensure left-rail Copilot is selected (WebView2 / Chromium root).
OutlookCopilot_EnsureCopilotRailOn() {
    Outlook_ActivateMainWindow()
    root := OutlookMail_RootElement()
    if !root
        return false
    copBtn := OutlookCopilot_FindRailButton(root)
    if !copBtn
        return false
    if Outlook_RailToggleIsPressed(copBtn)
        return true
    try copBtn.ScrollIntoView()
    Sleep 40
    try copBtn.SetFocus()
    Sleep 50
    try copBtn.Click()
    catch Error {
        try copBtn.Invoke()
        catch Error {
            return false
        }
    }
    return true
}

; Bounded wait until Copilot chat surface or voice control appears (copilot.md mainChat / buttons).
OutlookCopilot_WaitCopilotChatSurface(timeoutMs := 2500) {
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        root := OutlookMail_RootElement()
        if root {
            try {
                if root.FindFirst({ AutomationId: "mainChat" })
                    return true
            } catch {
            }
            try {
                if root.FindFirst({ AutomationId: "iframe:" OUTLOOK_COPILOT_RAIL_AID })
                    return true
            } catch {
            }
            try {
                if root.FindFirst({ Name: "Start a new voice chat", matchmode: "Substring" })
                    return true
            } catch {
            }
            try {
                if root.FindFirst({ Name: "Message Copilot", ControlType: "Edit" })
                    return true
            } catch {
            }
        }
        Sleep 80
    }
    return false
}

OutlookCopilot_ToggleVoiceChat() {
    try {
        StandardLoadingBar_Show("⏳ Outlook Copilot: voice chat…", BANNER_ACCENT_INTERMEDIATE, { passive: false,
            centerOnHwnd: 0, textWidth: 480, fontSize: 17 })
    } catch {
    }
    ok := false
    try {
        try StandardLoadingBar_Update("⏳ Outlook Copilot: activating Outlook…")
        if !OutlookCopilot_WaitMainOutlookForeground(2200) {
            ShowCenteredOverlay_Utils("❌ Outlook Copilot: Outlook window not in foreground", 2000, BANNER_ACCENT_ERROR)
            return false
        }
        try StandardLoadingBar_Update("⏳ Outlook Copilot: opening Copilot…")
        if !OutlookCopilot_EnsureCopilotRailOn() {
            ShowCenteredOverlay_Utils("❌ Outlook Copilot: rail button not found", 1600, BANNER_ACCENT_ERROR)
            return false
        }
        try StandardLoadingBar_Update("⏳ Outlook Copilot: loading surface…")
        OutlookCopilot_WaitCopilotChatSurface(4000)
        try StandardLoadingBar_Update("⏳ Outlook Copilot: waiting for voice button (quality gate)…")
        el := OutlookCopilot_WaitVoiceChatButtonReady(12000, 120)
        if !el {
            ShowCenteredOverlay_Utils("❌ Outlook Copilot: voice chat button not ready (timeout)", 2200,
                BANNER_ACCENT_ERROR)
            return false
        }
        if !OutlookCopilot_ElementIsRenderedAndClickable(el) {
            root := OutlookMail_RootElement()
            el := root ? OutlookCopilot_FindVoiceChatButtonInRoot(root) : ""
            if !el || !OutlookCopilot_ElementIsRenderedAndClickable(el) {
                ShowCenteredOverlay_Utils("❌ Outlook Copilot: voice chat button not actionable", 2000,
                    BANNER_ACCENT_ERROR)
                return false
            }
        }
        try el.ScrollIntoView()
        Sleep 40
        try el.SetFocus()
        Sleep 50
        try el.Click()
        catch Error {
            try el.Invoke()
            catch Error {
                ShowCenteredOverlay_Utils("❌ Outlook Copilot: voice chat click failed", 1600, BANNER_ACCENT_ERROR)
                return false
            }
        }
        ok := true
    } finally {
        try StandardLoadingBar_Hide(0)
        catch {
        }
    }
    return ok
}

OutlookCopilot_FocusComposer() {
    el := OutlookMail_FindFirst([{ AutomationId: "m365-chat-editor-target-element", ControlType: "Edit" }, { AutomationId: "m365-chat-editor-target-element" }, { Name: "Message Copilot",
        ControlType: "Edit" }])
    if !el
        return false
    try el.ScrollIntoView()
    try el.SetFocus()
    Sleep 40
    try el.Click()
    catch {
        return true
    }
    return true
}

OutlookCopilot_RunSlot(slot) {
    Outlook_ActivateMainWindow()
    switch slot {
        case 1:
            return OutlookCopilot_ToggleVoiceChat()
        case 2:
            if OutlookMail_ClickFirst([{ AutomationId: "new-chat-button", ControlType: "Button" }, { Name: "New chat",
                matchmode: "Substring", ControlType: "Button" }])
                return true
            ShowCenteredOverlay_Utils("❌ Outlook Copilot: New chat not found", 1600, BANNER_ACCENT_ERROR)
            return false
        case 3:
            if OutlookCopilot_FocusComposer()
                return true
            ShowCenteredOverlay_Utils("❌ Outlook Copilot: composer not found", 1600, BANNER_ACCENT_ERROR)
            return false
        case 4:
            if OutlookMail_ClickFirst([{ AutomationId: "sidepaneExpandButton", ControlType: "Button" }, { Name: "Open navigation panel",
                matchmode: "Substring", ControlType: "Button" }])
                return true
            ShowCenteredOverlay_Utils("❌ Outlook Copilot: navigation panel control not found", 1600,
                BANNER_ACCENT_ERROR)
            return false
        case 5:
            if OutlookMail_ClickFirst([{ AutomationId: "toggle-work", ControlType: "Button" }, { Name: "Work",
                ControlType: "Button" }])
                return true
            ShowCenteredOverlay_Utils("❌ Outlook Copilot: Work scope not found", 1600, BANNER_ACCENT_ERROR)
            return false
        case 6:
            if OutlookMail_ClickFirst([{ AutomationId: "toggle-web", ControlType: "Button" }, { Name: "Web",
                ControlType: "Button" }])
                return true
            ShowCenteredOverlay_Utils("❌ Outlook Copilot: Web scope not found", 1600, BANNER_ACCENT_ERROR)
            return false
        case 7:
            if OutlookMail_ClickFirst([{ AutomationId: "gptModeSwitcher", ControlType: "Button" }, { Name: "Auto",
                matchmode: "Substring", ControlType: "Button" }])
                return true
            ShowCenteredOverlay_Utils("❌ Outlook Copilot: mode menu not found", 1600, BANNER_ACCENT_ERROR)
            return false
        case 8:
            if OutlookMail_ClickFirst([{ AutomationId: "menura", ControlType: "Button" }, { Name: "Temporary chat",
                matchmode: "Substring", ControlType: "Button" }])
                return true
            ShowCenteredOverlay_Utils("❌ Outlook Copilot: Temporary chat not found", 1600, BANNER_ACCENT_ERROR)
            return false
        case 9:
            if OutlookMail_ClickFirst([{ AutomationId: "moreButton", ControlType: "Button" }, { Name: "OpenCopilot chats and more",
                matchmode: "Substring", ControlType: "Button" }, { Name: "chats and more", matchmode: "Substring",
                    ControlType: "Button" }
            ])
                return true
            ShowCenteredOverlay_Utils("❌ Outlook Copilot: More menu not found", 1600, BANNER_ACCENT_ERROR)
            return false
        default:
            return false
    }
}

ShowOutlookCopilotSelector() {
    global g_OutlookCopilotSelectorGui, g_OutlookCopilotSelectorActive, g_OutlookCopilotShortcuts
    if g_OutlookCopilotSelectorActive
        return
    g_OutlookCopilotSelectorGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Owner")
    g_OutlookCopilotSelectorGui.BackColor := "1E1E2E"
    g_OutlookCopilotSelectorGui.MarginX := 20
    g_OutlookCopilotSelectorGui.MarginY := 15
    g_OutlookCopilotSelectorGui.SetFont("s14 cCDD6F4 Bold", "Segoe UI")
    g_OutlookCopilotSelectorGui.Add("Text", "w320 Center", "🤖 Outlook Copilot")
    g_OutlookCopilotSelectorGui.Add("Text", "w320 h1 Background45475A")
    g_OutlookCopilotSelectorGui.SetFont("s12 cCDD6F4", "Segoe UI")
    for num, item in g_OutlookCopilotShortcuts {
        g_OutlookCopilotSelectorGui.Add("Text", "w320", "[" num "] " item.name)
        g_OutlookCopilotSelectorGui.SetFont("s9 c6C7086", "Segoe UI")
        g_OutlookCopilotSelectorGui.Add("Text", "w320 y+2", "    " item.desc)
        g_OutlookCopilotSelectorGui.SetFont("s12 cCDD6F4", "Segoe UI")
    }
    g_OutlookCopilotSelectorGui.Add("Text", "w320 h1 Background45475A y+10")
    g_OutlookCopilotSelectorGui.SetFont("s9 c6C7086", "Segoe UI")
    g_OutlookCopilotSelectorGui.Add("Text", "w320 Center", "Press 1–9 | Esc to cancel")
    try {
        g_OutlookCopilotSelectorGui.OnEvent("Escape", OutlookCopilotSelector_GuiEscape)
    } catch {
    }
    activeWin := 0
    try activeWin := WinGetID("A")
    catch {
        activeWin := 0
    }
    MonitorGetWorkArea(1, &monitorLeft, &monitorTop, &monitorRight, &monitorBottom)
    monitorWidth := monitorRight - monitorLeft
    monitorHeight := monitorBottom - monitorTop
    if activeWin {
        rect := Buffer(16, 0)
        if DllCall("GetWindowRect", "ptr", activeWin, "ptr", rect) {
            winLeft := NumGet(rect, 0, "int")
            winTop := NumGet(rect, 4, "int")
            winRight := NumGet(rect, 8, "int")
            winBottom := NumGet(rect, 12, "int")
            centerX := winLeft + (winRight - winLeft) // 2
            centerY := winTop + (winBottom - winTop) // 2
            loop MonitorGetCount() {
                MonitorGetWorkArea(A_Index, &l, &t, &r, &b)
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
    g_OutlookCopilotSelectorGui.Show("AutoSize Hide")
    g_OutlookCopilotSelectorGui.GetPos(&gx, &gy, &gw, &gh)
    cx := monitorLeft + (monitorWidth - gw) // 2
    cy := monitorTop + (monitorHeight - gh) // 2
    ; Avoid "NA": if focus stays in Outlook/Chromium, Esc is consumed there first. Activate so Esc reaches this Gui.
    g_OutlookCopilotSelectorGui.Show("x" cx " y" cy)
    try WinActivate(g_OutlookCopilotSelectorGui.Hwnd)
    g_OutlookCopilotSelectorActive := true
    loop 9 {
        Hotkey(String(A_Index), OutlookCopilotSelector_HandleKey, "On")
    }
    ; $*Escape: $ forces keyboard hook (ignores synthetic Esc from SendInput in same script). I10 uses g_OnEscapePressed.
    Hotkey("$*Escape", OutlookCopilotSelector_EscapeFromHotkey, "On")
    global g_OnEscapePressed
    g_OnEscapePressed := OutlookCopilotSelector_GlobalEscapeCallback
    Utils_EnsureGlobalEscapeHotkey()
    ; Fallback when hook hotkeys miss Esc (e.g. another AHK process / hook order).
    g_OutlookCopilotEscPollPrev := false
    SetTimer(OutlookCopilotSelector_EscapePoll, 50)
}

; Poll Esc — fallback when $*Escape / g_OnEscapePressed miss. VK_ESCAPE 0x1B via GetAsyncKeyState.
OutlookCopilotSelector_EscapePoll() {
    global g_OutlookCopilotSelectorActive, g_OutlookCopilotEscPollPrev
    if !g_OutlookCopilotSelectorActive {
        SetTimer(OutlookCopilotSelector_EscapePoll, 0)
        return
    }
    escSync := GetKeyState("Escape", "P")
    escAsync := (DllCall("user32\GetAsyncKeyState", "int", 0x1B) & 0x8000) != 0
    escDown := escSync || escAsync
    if escDown {
        if !g_OutlookCopilotEscPollPrev {
            g_OutlookCopilotEscPollPrev := true
            OutlookCopilotSelector_Cancel()
        }
    } else {
        g_OutlookCopilotEscPollPrev := false
    }
}

OutlookCopilotSelector_EscapeFromHotkey(*) {
    OutlookCopilotSelector_Cancel()
}

; Used with g_OnEscapePressed when I0 *Escape does not fire
OutlookCopilotSelector_GlobalEscapeCallback(*) {
    OutlookCopilotSelector_Cancel()
}

OutlookCopilotSelector_GuiEscape(*) {
    OutlookCopilotSelector_Cancel()
}

OutlookCopilotSelector_HandleKey(ThisHotkey, *) {
    global g_OutlookCopilotSelectorActive
    if !g_OutlookCopilotSelectorActive
        return
    sel := Integer(ThisHotkey)
    if sel < 1 || sel > 9
        return
    OutlookCopilotSelector_Close()
    OutlookCopilot_RunSlot(sel)
}

OutlookCopilotSelector_Cancel(*) {
    OutlookCopilotSelector_Close()
}

OutlookCopilotSelector_Close() {
    global g_OutlookCopilotSelectorGui, g_OutlookCopilotSelectorActive
    if !g_OutlookCopilotSelectorActive
        return
    SetTimer(OutlookCopilotSelector_EscapePoll, 0)
    global g_OutlookCopilotEscPollPrev
    g_OutlookCopilotEscPollPrev := false
    g_OutlookCopilotSelectorActive := false
    global g_OnEscapePressed
    g_OnEscapePressed := ""
    loop 9 {
        try Hotkey(String(A_Index), "Off")
    }
    try Hotkey("Escape", OutlookCopilotSelector_Cancel, "Off")
    try Hotkey("*Escape", OutlookCopilotSelector_Cancel, "Off")
    try Hotkey("$*Escape", OutlookCopilotSelector_EscapeFromHotkey, "Off")
    Utils_EnsureGlobalEscapeHotkey()
    if IsObject(g_OutlookCopilotSelectorGui) && g_OutlookCopilotSelectorGui.Hwnd {
        try g_OutlookCopilotSelectorGui.Destroy()
    }
    g_OutlookCopilotSelectorGui := false
}

SelectOutlookCopilotShortcut() {
    global g_OutlookCopilotSelectorActive
    if g_OutlookCopilotSelectorActive {
        OutlookCopilotSelector_Cancel()
    } else
        ShowOutlookCopilotSelector()
}

; True if toggle appears "on" (left-rail app bar toggle buttons).
Outlook_RailToggleIsPressed(el) {
    if !el
        return false
    try {
        if el.IsTogglePatternAvailable
            return (el.TogglePattern.CurrentToggleState = 1)
    } catch {
    }
    try {
        return (el.GetCurrentPropertyValue(UIA.Property.ToggleToggleState) = 1)
    } catch {
    }
    return false
}

; Shift+M: switch between Mail and Calendar using rail toggle state (not window title). See outlook-mail.md left-rail-appbar.
Outlook_ToggleMailCalendarRail() {
    Outlook_ActivateMainWindow()
    root := OutlookMail_RootElement()
    if !root
        return false
    searchRoot := root
    try {
        r := root.FindFirst({ Name: "left-rail-appbar", matchmode: "Substring", ControlType: "Group" })
        if !r
            r := root.FindFirst({ Name: "left-rail-appbar", matchmode: "Substring" })
        if r
            searchRoot := r
    } catch {
    }
    mailBtn := ""
    calBtn := ""
    try mailBtn := searchRoot.FindFirst({ AutomationId: "ddea774c-382b-47d7-aab5-adc2139a802b", ControlType: "Button" })
    if !mailBtn
        try mailBtn := searchRoot.FindFirst({ Name: "Mail", ControlType: "Button" })
    try calBtn := searchRoot.FindFirst({ AutomationId: "8cbeb86f-83e1-43b5-aaba-cd3514322f0b", ControlType: "Button" })
    if !calBtn
        try calBtn := searchRoot.FindFirst({ Name: "Calendar", ControlType: "Button" })
    ; If rail group scope missed (build differences), search full Chromium root
    if !mailBtn
        try mailBtn := root.FindFirst({ AutomationId: "ddea774c-382b-47d7-aab5-adc2139a802b", ControlType: "Button" })
    if !calBtn
        try calBtn := root.FindFirst({ AutomationId: "8cbeb86f-83e1-43b5-aaba-cd3514322f0b", ControlType: "Button" })
    if !mailBtn || !calBtn
        return false

    mailOn := Outlook_RailToggleIsPressed(mailBtn)
    calOn := Outlook_RailToggleIsPressed(calBtn)
    target := ""

    if (mailOn && !calOn)
        target := calBtn
    else if (calOn && !mailOn)
        target := mailBtn
    else {
        ; Both off, both on, or no toggle pattern: infer from title
        t := WinGetTitle("A")
        if RegExMatch(t, "i)Calendar")
            target := mailBtn
        else
            target := calBtn
    }

    if !target
        return false
    try target.SetFocus()
    Sleep 50
    try target.Click()
    catch {
        try target.Invoke()
    }
    return true
}

Outlook_MailList_GetFirstListItem(root) {
    if !root
        return 0
    listEl := 0
    try {
        g := root.FindFirst({ AutomationId: "MailList" })
        if g
            listEl := g.FindFirst({ ControlType: "List" })
    } catch {
    }
    if !listEl {
        try listEl := root.FindFirst({ Name: "Message list", ControlType: "List" })
    }
    if !listEl {
        try listEl := root.FindFirst({ Name: "Message list", Type: 50008 })
    }
    if !listEl
        return 0

    ; New Outlook may expose date/header bars near the top of the list. Pick the first real
    ; message row instead of UI chrome/header artifacts.
    try {
        items := listEl.FindAll({ ControlType: "ListItem" })
        if IsObject(items) {
            for item in items {
                if !Outlook_MailList_IsNonMessageItem(item)
                    return item
            }
        }
    } catch {
    }

    try {
        items := listEl.FindAll({ Type: 50007 })
        if IsObject(items) {
            for item in items {
                if !Outlook_MailList_IsNonMessageItem(item)
                    return item
            }
        }
    } catch {
    }

    item := 0
    try item := listEl.FindFirst({ ControlType: "ListItem" })
    if !item {
        try item := listEl.FindFirst({ Type: 50007 })
    }
    if item && !Outlook_MailList_IsNonMessageItem(item)
        return item
    return 0
}

Outlook_MailList_IsNonMessageItem(item) {
    if !item
        return true
    aid := ""
    name := ""
    hasSelectionPattern := false
    try aid := item.AutomationId
    try name := Trim(item.Name)
    try hasSelectionPattern := !!item.GetCurrentPropertyValue(UIA.Property.IsSelectionItemPatternAvailable)

    ; Date group headers / list chrome can appear before the first actual mail row.
    if (aid != "" && RegExMatch(aid, "i)^groupHeader"))
        return true
    if (name != "" && RegExMatch(name, "i)^((Today|Yesterday)\b|Header action menu)$"))
        return true

    ; Tiny unnamed bars without selection pattern are not message rows.
    if (!hasSelectionPattern && StrLen(name) < 4)
        return true

    return false
}

Outlook_MailList_FocusedIsDrawerHeader() {
    try {
        fe := UIA.GetFocusedElement()
        if !fe
            return false
        name := ""
        aid := ""
        try name := Trim(fe.Name)
        try aid := fe.AutomationId
        if (aid != "" && RegExMatch(aid, "i)^groupHeader"))
            return true
        if (name != "" && RegExMatch(name, "i)^(Today|Yesterday)\b"))
            return true
    } catch {
    }
    return false
}

Outlook_MailList_FocusedIsLikelyMessageRow() {
    try {
        fe := UIA.GetFocusedElement()
        if !fe
            return false
        if Outlook_MailList_IsNonMessageItem(fe)
            return false
        ctlType := 0
        hasSelectionPattern := false
        try ctlType := fe.ControlType
        try hasSelectionPattern := !!fe.GetCurrentPropertyValue(UIA.Property.IsSelectionItemPatternAvailable)
        if (ctlType = UIA.ControlType.ListItem || ctlType = 50007)
            return true
        return hasSelectionPattern
    } catch {
    }
    return false
}

Outlook_MailList_SkipDrawerHeadersByKeyboard(maxSteps := 5) {
    loop maxSteps {
        if Outlook_MailList_FocusedIsLikelyMessageRow()
            return true
        Send "{Down}"
        Sleep 30
    }
    return Outlook_MailList_FocusedIsLikelyMessageRow()
}

Outlook_MailList_TrySelectFirstItem(root) {
    item := Outlook_MailList_GetFirstListItem(root)
    if !item
        return false
    try item.ScrollIntoView()
    catch {
    }
    Sleep 40
    try item.Select()
    catch {
        try item.Click()
        catch {
            try item.Invoke()
            catch {
                return false
            }
        }
    }
    try item.SetFocus()
    return true
}

; selectFirst: when true, select the first message row in the list (if any) for faster triage after moves.
Outlook_FocusMailMessageList(selectFirst := false) {
    Outlook_ActivateMainWindow()
    hwnd := WinExist("A")
    if !hwnd
        return false
    if selectFirst {
        try {
            root := OutlookMail_RootElementForHwnd(hwnd)
            if !root
                root := UIA.ElementFromHandle(hwnd)
            if Outlook_MailList_TrySelectFirstItem(root)
                return true
        } catch {
        }
    }
    if !OutlookFocusFirst([{ AutomationId: "Skip to message list-region" }, { Name: "Message list", matchmode: "Substring" }])
        return false

    if !selectFirst {
        try EnsureFocus()
        return true
    }

    ; Retry selection after focusing list container.
    try {
        root := OutlookMail_RootElementForHwnd(WinExist("A"))
        if !root
            root := UIA.ElementFromHandle(WinExist("A"))
        if Outlook_MailList_TrySelectFirstItem(root)
            return true
    } catch {
    }

    ; Last fallback: Home may land on Today/Yesterday header; step down until a message row is focused.
    Send "{Home}"
    Sleep 40
    return Outlook_MailList_SkipDrawerHeadersByKeyboard()
}

Outlook_FocusMailReadingPane() {
    Outlook_ActivateMainWindow()
    if OutlookFocusFirst([{ AutomationId: "Skip to message-region" }, { Name: "Reading Pane", matchmode: "Substring" }]) {
        try EnsureFocus()
        return true
    }
    return false
}

OutlookMail_ClickReadingPaneCommand(cmdName) {
    Outlook_ActivateMainWindow()
    try {
        root := OutlookMail_RootElement()
        if !root
            root := UIA.ElementFromHandle(WinExist("A"))
        pane := root.FindFirst({ AutomationId: "Skip to message-region" })
        if !pane
            return false
        el := 0
        try el := pane.FindFirst({ Name: cmdName, ControlType: "MenuItem" })
        if !el
            try el := pane.FindFirst({ Name: cmdName, ControlType: "Button" })
        if el {
            try el.SetFocus()
            Sleep 40
            try el.Click()
            catch {
                try el.Invoke()
            }
            return true
        }
    } catch {
    }
    return false
}

; Reading pane group (New Outlook WebView2 UI). Uses the active window only — no Outlook_ActivateMainWindow (see IsOutlookMeetingRequestReadingPaneActive vs pop-out invitation).
OutlookMail_GetReadingPaneElement() {
    try {
        hwnd := WinExist("A")
        if !hwnd
            return ""
        root := OutlookMail_RootElementForHwnd(hwnd)
        if root {
            pane := root.FindFirst({ AutomationId: "Skip to message-region" })
            if pane
                return pane
        }
    } catch {
    }
    try {
        root := UIA.ElementFromHandle(WinExist("A"))
        if root {
            pane := root.FindFirst({ AutomationId: "Skip to message-region" })
            if pane
                return pane
        }
    } catch {
    }
    return ""
}

; True when the main Mail/Calendar shell shows a meeting request in the reading pane (not a popped-out invitation window).
IsOutlookMeetingRequestReadingPaneActive() {
    if !IsOutlookMainActive()
        return false
    try {
        pane := OutlookMail_GetReadingPaneElement()
        if !pane
            return false
        return pane.FindFirst({ Name: "Accept the meeting", ControlType: "MenuItem" }) ? true : false
    } catch {
    }
    return false
}

; Popped-out meeting invitation: " - Message (" inspector with Accept/Decline row (same UI as reading pane, different host window).
IsOutlookMeetingInvitationPopOutActive() {
    if !(WinActive("ahk_exe OUTLOOK.EXE") || WinActive("ahk_exe olk.exe"))
        return false
    if !RegExMatch(WinGetTitle("A"), "i) - Message \(")
        return false
    try {
        root := OutlookMail_RootElementForHwnd(WinExist("A"))
        if !root
            return false
        return root.FindFirst({ Name: "Accept the meeting", ControlType: "MenuItem" }) ? true : false
    } catch {
    }
    return false
}

; Organizer canceled: primary action is Button "Remove event" (not Accept/Decline).
IsOutlookRemoveEventReadingPaneActive() {
    if !IsOutlookMainActive()
        return false
    try {
        root := OutlookMail_RootElementForHwnd(WinExist("A"))
        if !root
            return false
        return root.FindFirst({ Name: "Remove event", ControlType: "Button" }) ? true : false
    } catch {
    }
    return false
}

IsOutlookRemoveEventPopOutActive() {
    if !(WinActive("ahk_exe OUTLOOK.EXE") || WinActive("ahk_exe olk.exe"))
        return false
    if !RegExMatch(WinGetTitle("A"), "i) - Message \(")
        return false
    try {
        root := OutlookMail_RootElementForHwnd(WinExist("A"))
        if !root
            return false
        return root.FindFirst({ Name: "Remove event", ControlType: "Button" }) ? true : false
    } catch {
    }
    return false
}

OutlookMeeting_ClickMenuItemInActiveWindow(criteriaList) {
    hwnd := WinExist("A")
    root := OutlookMail_RootElementForHwnd(hwnd)
    if !root
        return false
    for criteria in criteriaList {
        try {
            el := root.FindFirst(criteria)
            if el {
                try el.SetFocus()
                Sleep 50
                try el.Click()
                catch {
                    try el.Invoke()
                }
                return true
            }
        } catch {
        }
    }
    return false
}

OutlookMeeting_ClickAccept() {
    return OutlookMeeting_ClickMenuItemInActiveWindow([{ Name: "Accept the meeting", ControlType: "MenuItem" }])
}

OutlookMeeting_ClickFollow() {
    return OutlookMeeting_ClickMenuItemInActiveWindow([{ Name: "Follow;", matchmode: "Substring", ControlType: "MenuItem" }])
}

OutlookMeeting_ClickDecline() {
    return OutlookMeeting_ClickMenuItemInActiveWindow([{ Name: "Decline the meeting", ControlType: "MenuItem" }])
}

; True if user confirms (Yes). Default button is Yes (Enter confirms).
OutlookMeeting_ConfirmDecline() {
    return MsgBox("Decline this meeting invitation?", "Confirm decline", "Icon? YesNo Default1") = "Yes"
}

; Canceled meeting: Button "Remove event" (see outlook-remove-evet.md). No confirmation in script.
OutlookMeeting_ClickRemoveEvent() {
    return OutlookMeeting_ClickMenuItemInActiveWindow([{ Name: "Remove event", ControlType: "Button" }])
}

; Opens "More options" (…) then clicks a MenuItem in the overflow menu (see mark-appointment-request.md).
OutlookMeeting_ClickMoreOptionsThen(menuItemName) {
    try {
        hwnd := WinExist("A")
        root := OutlookMail_RootElementForHwnd(hwnd)
        if !root
            return false
        pane := root.FindFirst({ AutomationId: "Skip to message-region" })
        searchRoot := pane ? pane : root
        moreBtn := ""
        try moreBtn := searchRoot.FindFirst({ AutomationId: "menur7c4", ControlType: "Button" })
        if !moreBtn
            try moreBtn := searchRoot.FindFirst({ Name: "More options", ControlType: "Button", matchmode: "Substring" })
        if !moreBtn
            try moreBtn := root.FindFirst({ AutomationId: "menur7c4", ControlType: "Button" })
        if !moreBtn
            return false
        try moreBtn.SetFocus()
        Sleep 50
        try moreBtn.Click()
        catch {
            try moreBtn.Invoke()
        }
        Sleep 120
        el := ""
        try el := root.FindFirst({ Name: menuItemName, ControlType: "MenuItem" })
        if !el
            try el := UIA.ElementFromHandle(WinExist("A")).FindFirst({ Name: menuItemName, ControlType: "MenuItem" })
        if !el
            try el := root.FindFirst({ Name: menuItemName, matchmode: "Substring", ControlType: "MenuItem" })
        if !el
            return false
        try el.SetFocus()
        Sleep 40
        try el.Click()
        catch {
            try el.Invoke()
        }
        return true
    } catch {
    }
    return false
}

OutlookMail_EnsureNavigationPaneVisible() {
    Outlook_ActivateMainWindow()
    try {
        root := OutlookMail_RootElement()
        if !root
            return false
        ; If the navigation pane exists (or at least the folder tree), we’re good.
        try {
            if root.FindFirst({ Name: "Navigation pane", matchmode: "Substring" })
                return true
            if root.FindFirst({ ControlType: "Tree" })
                return true
        } catch {
        }

        ; Otherwise toggle the nav pane (label may be "Show…" or "Hide…", depending on state).
        navToggleCriteria := [{ Name: "navigation pane", matchmode: "Substring", ControlType: "Button" }, { Name: "Navigation pane",
            matchmode: "Substring", ControlType: "Button" }]
        for x in OutlookMail_CriteriaToggleNavigationPaneRibbon()
            navToggleCriteria.Push(x)
        if OutlookMail_ClickFirst(navToggleCriteria) {
            Sleep 120
            try {
                if root.FindFirst({ ControlType: "Tree" })
                    return true
            } catch {
            }
        }
    } catch {
    }
    return false
}

; Ribbon folder-pane control: Type 50000 (Button), Fluent className prefix fui-Button …
; Name is "Show navigation pane" when the folder pane is collapsed, "Hide navigation pane" when expanded.
OutlookMail_CriteriaShowNavigationPaneRibbon() {
    return [{ Name: "Show navigation pane", Type: 50000, matchmode: "Substring" }, { Name: "Show navigation pane",
        ControlType: "Button", matchmode: "Substring" }, { Name: "Show navigation pane", ClassName: "fui-Button",
            matchmode: "Substring" }
    ]
}

OutlookMail_CriteriaHideNavigationPaneRibbon() {
    return [{ Name: "Hide navigation pane", Type: 50000, matchmode: "Substring" }, { Name: "Hide navigation pane",
        ControlType: "Button", matchmode: "Substring" }, { Name: "Hide navigation pane", ClassName: "fui-Button",
            matchmode: "Substring" }
    ]
}

OutlookMail_CriteriaToggleNavigationPaneRibbon() {
    c := []
    for x in OutlookMail_CriteriaHideNavigationPaneRibbon()
        c.Push(x)
    for x in OutlookMail_CriteriaShowNavigationPaneRibbon()
        c.Push(x)
    return c
}

; New Outlook mail surface is WebView2 (Chromium). Ribbon + lists live under Chrome_RenderWidgetHostHWND1;
; UIA.ElementFromHandle(top-level hwnd) may not include that subtree (see UIA.ElementFromChromium).
OutlookMail_RootElement() {
    Outlook_ActivateMainWindow()
    hwnd := WinExist("A")
    if !hwnd
        return ""
    return OutlookMail_RootElementForHwnd(hwnd)
}

; Chromium root for a specific top-level hwnd — does not activate another window (reading pane vs pop-out invitation).
OutlookMail_RootElementForHwnd(hwnd) {
    if !hwnd
        return ""
    try {
        return UIA.ElementFromChromium("ahk_id " hwnd, 500)
    } catch {
    }
    try {
        return UIA.ElementFromHandle(hwnd)
    } catch {
    }
    return ""
}

OutlookMail_FindFirst(criteriaList) {
    root := OutlookMail_RootElement()
    if !root
        return ""
    for criteria in criteriaList {
        try {
            el := root.FindFirst(criteria)
            if el
                return el
        } catch {
        }
    }
    return ""
}

OutlookMail_ClickFirst(criteriaList) {
    el := OutlookMail_FindFirst(criteriaList)
    if !el
        return false
    try el.SetFocus()
    Sleep 50
    try el.Click()
    catch Error {
        try el.Invoke()
        catch Error {
            return false
        }
    }
    return true
}

; New Outlook: ribbon command buttons (Delete, Move, …) live on the Home tab panel only (outlook.md: TabItem Home, AutomationId "1").
; Call before UIA clicks that target AutomationIds on that panel when View/Help might be active.
OutlookMail_EnsureHomeTab() {
    Outlook_ActivateMainWindow()
    homeTab := OutlookMail_FindFirst([{ Name: "Home", ControlType: "TabItem", AutomationId: "1" }, { AutomationId: "1",
        ControlType: "TabItem" }, { Name: "Home", ControlType: "TabItem" }
    ])
    if !homeTab {
        try {
            Send "!h"
            Sleep 100
        } catch {
        }
        return true
    }
    try {
        if homeTab.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable) && homeTab.SelectionItemPattern.IsSelected {
            Sleep 30
            return true
        }
    } catch {
    }
    try homeTab.SetFocus()
    Sleep 30
    try homeTab.Click()
    catch Error {
        try homeTab.Invoke()
        catch Error {
            try {
                Send "!h"
                Sleep 100
            } catch {
            }
            return true
        }
    }
    Sleep 100
    return true
}

; Left folder list collapsed: ribbon shows "Show navigation pane" (outlook-mail.md: Ribbon … 8,1).
OutlookMail_IsLeftSidePanelHidden() {
    Outlook_ActivateMainWindow()
    return OutlookMail_FindFirst(OutlookMail_CriteriaShowNavigationPaneRibbon()) != ""
}

; Ribbon "high navigation" toggle: show the left folder pane (same control as "Hide navigation pane" when open).
OutlookMail_ClickHighNavigationShowPane() {
    Outlook_ActivateMainWindow()
    if OutlookMail_ClickFirst(OutlookMail_CriteriaShowNavigationPaneRibbon())
        return true
    return OutlookMail_EnsureNavigationPaneVisible()
}

; Nav pane already visible: open Inbox via folder tree only (no ribbon toggle).
OutlookMail_GoToInboxShortcut() {
    Outlook_ActivateMainWindow()
    try {
        root := OutlookMail_RootElement()
        if !root
            return false

        ; Scope search to the Navigation pane subtree to avoid colliding with other “Inbox” elements.
        nav := 0
        try nav := root.FindFirst({ Name: "Navigation pane", matchmode: "Substring" })
        if !nav
            nav := root

        inbox := 0
        ; Prefer the uniquely-named selected variant.
        try inbox := nav.FindFirst({ Name: "Inbox selected", ControlType: "TreeItem" })
        if !inbox
            try inbox := nav.FindFirst({ Name: "Inbox", ControlType: "TreeItem" })
        if !inbox
            try inbox := nav.FindFirst({ Name: "Inbox", matchmode: "Substring", ControlType: "TreeItem" })

        if inbox {
            try inbox.ScrollIntoView()
            try inbox.SetFocus()
            Sleep 40
            try inbox.Click()
            catch {
                try inbox.Invoke()
            }
            return true
        }
    } catch {
    }
    return false
}

OutlookMail_ClickInboxFolder() {
    Outlook_ActivateMainWindow()
    if OutlookMail_IsLeftSidePanelHidden() {
        if !OutlookMail_ClickHighNavigationShowPane()
            return false
        Sleep 120
        return OutlookMail_GoToInboxShortcut()
    }
    return OutlookMail_GoToInboxShortcut()
}

; Ribbon: Hide navigation pane / Show navigation pane (outlook-mail.md: Ribbon … 8,1).
OutlookMail_ToggleHighNavigationPane() {
    Outlook_ActivateMainWindow()
    return OutlookMail_ClickFirst(OutlookMail_CriteriaToggleNavigationPaneRibbon())
}

OutlookCompose_FocusToRecipientsField() {
    Outlook_ActivateMainWindow()
    try {
        root := UIA.ElementFromHandle(WinExist("A"))

        ; #region agent log
        OC_ToLog(msg, data := "{}", hypo := "OC_To") {
            try {
                line := "{"
                    . '"sessionId":"b96502",'
                    . '"runId":"shiftT",'
                    . '"hypothesisId":"' hypo '",'
                    . '"timestamp":' A_TickCount ','
                    . '"location":"Shift keys.ahk:OutlookCompose_FocusToRecipientsField",'
                    . '"message":"' StrReplace(msg, '"', '\"') '",'
                    . '"data":' data
                    . "}"
                FileAppend(line "`n", "debug-b96502.log", "UTF-8")
            } catch {
            }
        }
        ; #endregion

        try {
            hwnd := WinExist("A")
            t := WinGetTitle("A")
            c := WinGetClass("A")
            p := WinGetProcessName("A")
            OC_ToLog("entry", '{"hwnd":' hwnd ',"proc":"' StrReplace(p, '"', '\"') '","class":"' StrReplace(c, '"',
                '\"') '","title":"' StrReplace(SubStr(t, 1, 120), '"', '\"') '"}', "OC_To_A")
        } catch {
        }

        ; Step 1: click the To: row (reactive UI may expand recipients editor)
        okTo := OutlookClickFirst([{ AutomationId: "134", ControlType: "Group" }, { AutomationId: "134" }, { Name: "To:",
            matchmode: "Substring" }])
        try OC_ToLog("after_to_click", '{"ok":' (okTo ? 1 : 0) '}', "OC_To_B")
        catch {
        }
        if !okTo
            return false
        Sleep 80

        ; Step 2: click the recipient entity group (AutomationId looks like REK000070; can change)
        recipGroup := 0
        try recipGroup := root.FindFirst({ AutomationId: "REK", matchmode: "Substring", ControlType: "Group" })
        if !recipGroup
            try recipGroup := root.FindFirst({ AutomationId: "REK", matchmode: "Substring" })
        if !recipGroup {
            ; Broad fallback: find any group that looks like a recipient entity (class contains _EType_RECIPIENT_ENTITY)
            try recipGroup := root.FindFirst({ ClassName: "_EType_RECIPIENT_ENTITY", matchmode: "Substring",
                ControlType: "Group" })
        }
        if recipGroup {
            n := "", aid := "", cn := "", ct := "", off := "", en := ""
            try n := recipGroup.Name
            try aid := recipGroup.AutomationId
            try cn := recipGroup.ClassName
            try ct := recipGroup.ControlType
            try off := recipGroup.IsOffscreen
            try en := recipGroup.IsEnabled
            try OC_ToLog("recip_group_found", '{"name":"' StrReplace(SubStr(n, 1, 80), '"', '\"') '","automationId":"' StrReplace(
                SubStr(aid, 1, 80), '"', '\"') '","className":"' StrReplace(SubStr(cn, 1, 80), '"', '\"') '","controlType":"' StrReplace(
                    ct, '"', '\"') '","isOffscreen":' (off ? 1 : 0) ',"isEnabled":' (en ? 1 : 0) '}', "OC_To_C")
            catch {
            }
        } else {
            try OC_ToLog("recip_group_not_found", "{}", "OC_To_C")
            catch {
            }
        }
        if recipGroup {
            try recipGroup.ScrollIntoView()
            try recipGroup.SetFocus()
            Sleep 30
            try recipGroup.Click()
            catch {
                try recipGroup.Invoke()
            }
            Sleep 60

            ; Step 3: focus the editable field inside the recipient group (if exposed)
            edit := 0
            try edit := recipGroup.FindFirst({ ControlType: "Edit" })
            if !edit
                try edit := recipGroup.FindFirst({ Type: 50004 })
            if edit {
                try OC_ToLog("inner_edit_found", "{}", "OC_To_D")
                catch {
                }
                try edit.SetFocus()
                Sleep 20
                try edit.Click()
                try {
                    fe := UIA.GetFocusedElement()
                    fn := "", fa := "", ft := ""
                    try fn := fe.Name
                    try fa := fe.AutomationId
                    try ft := fe.Type
                    OC_ToLog("focused_after_edit", '{"name":"' StrReplace(SubStr(fn, 1, 80), '"', '\"') '","automationId":"' StrReplace(
                        SubStr(fa, 1, 80), '"', '\"') '","type":' (ft = "" ? -1 : ft) '}', "OC_To_E")
                } catch {
                }
                return true
            }
            ; Fallback: click the hover target wrapper (often the direct text/caret host)
            wrap := 0
            try wrap := recipGroup.FindFirst({ ClassName: "lpcWrapper", matchmode: "Substring" })
            if wrap {
                try OC_ToLog("wrapper_found", "{}", "OC_To_D")
                catch {
                }
                try wrap.SetFocus()
                Sleep 20
                try wrap.Click()
                try {
                    fe := UIA.GetFocusedElement()
                    fn := "", fa := "", ft := ""
                    try fn := fe.Name
                    try fa := fe.AutomationId
                    try ft := fe.Type
                    OC_ToLog("focused_after_wrapper", '{"name":"' StrReplace(SubStr(fn, 1, 80), '"', '\"') '","automationId":"' StrReplace(
                        SubStr(fa, 1, 80), '"', '\"') '","type":' (ft = "" ? -1 : ft) '}', "OC_To_E")
                } catch {
                }
                return true
            }
            try OC_ToLog("no_inner_target", "{}", "OC_To_D")
            catch {
            }
            return true
        }
    } catch {
    }
    return false
}

IsNewOutlookActive() {
    if !(WinActive("ahk_exe OUTLOOK.EXE") || WinActive("ahk_exe olk.exe"))
        return false
    cls := ""
    title := ""
    exe := ""
    try cls := WinGetClass("A")
    try title := WinGetTitle("A")
    try exe := WinGetProcessName("A")
    return InStr(cls, "Outlook Host")
    || InStr(title, " - Outlook")
    || RegExMatch(title, "i)^(New event|Reminders?)")
}

OutlookFindFirst(criteriaList) {
    try {
        root := UIA.ElementFromHandle(WinExist("A"))
        for criteria in criteriaList {
            el := root.FindFirst(criteria)
            if el
                return el
        }
    } catch Error {
    }
    return ""
}

OutlookFocusFirst(criteriaList) {
    el := OutlookFindFirst(criteriaList)
    if !el
        return false
    try el.SetFocus()
    return true
}

OutlookClickFirst(criteriaList) {
    el := OutlookFindFirst(criteriaList)
    if !el
        return false
    try el.SetFocus()
    Sleep 50
    try el.Click()
    catch Error {
        try el.Invoke()
        catch Error {
            return false
        }
    }
    return true
}

#HotIf IsOutlookMainActive()

; -------------------------------------------------------------------
; Outlook main window (New Outlook) overflow layer: Ctrl+Alt+…
; -------------------------------------------------------------------

^!f:: {  ; Focus Search
    if !Outlook_FocusMainSearch()
        ShowCenteredOverlay_Utils("❌ Outlook: Search not found", 1200, BANNER_ACCENT_ERROR)
}

^!m:: {  ; Switch to Mail
    if !Outlook_SwitchToMail()
        ShowCenteredOverlay_Utils("❌ Outlook: Mail not found", 1200, BANNER_ACCENT_ERROR)
}

^!g:: {  ; Switch to Calendar
    if !Outlook_SwitchToCalendar()
        ShowCenteredOverlay_Utils("❌ Outlook: Calendar not found", 1200, BANNER_ACCENT_ERROR)
}

^!l:: {  ; Focus message list
    if !Outlook_FocusMailMessageList()
        ShowCenteredOverlay_Utils("❌ Outlook: Message list not found", 1200, BANNER_ACCENT_ERROR)
}

^!p:: {  ; Focus reading pane
    if !Outlook_FocusMailReadingPane()
        ShowCenteredOverlay_Utils("❌ Outlook: Reading pane not found", 1200, BANNER_ACCENT_ERROR)
}

; Mail triage (Reading Pane / Ribbon)
^!r:: {  ; Reply
    if OutlookMail_ClickReadingPaneCommand("Reply")
        return
    OutlookMail_EnsureHomeTab()
    if !OutlookClickFirst([{ Name: "Reply", ControlType: "Button" }, { Name: "Reply",
        ControlType: "MenuItem" }])
        ShowCenteredOverlay_Utils("❌ Outlook: Reply not found", 1200, BANNER_ACCENT_ERROR)
}

^!a:: {  ; Reply all
    if OutlookMail_ClickReadingPaneCommand("Reply all")
        return
    OutlookMail_EnsureHomeTab()
    if !OutlookClickFirst([{ Name: "Reply all", ControlType: "Button" }, { Name: "Reply all",
        ControlType: "MenuItem" }])
        ShowCenteredOverlay_Utils("❌ Outlook: Reply all not found", 1200, BANNER_ACCENT_ERROR)
}

^!w:: {  ; Forward
    if OutlookMail_ClickReadingPaneCommand("Forward")
        return
    OutlookMail_EnsureHomeTab()
    if !OutlookClickFirst([{ Name: "Forward", ControlType: "Button" }, { Name: "Forward",
        ControlType: "MenuItem" }])
        ShowCenteredOverlay_Utils("❌ Outlook: Forward not found", 1200, BANNER_ACCENT_ERROR)
}

^!d:: {  ; Delete
    OutlookMail_EnsureHomeTab()
    if !OutlookClickFirst([{ AutomationId: "519", ControlType: "Button" }, { Name: "Delete", ControlType: "Button" }])
        ShowCenteredOverlay_Utils("❌ Outlook: Delete not found", 1200, BANNER_ACCENT_ERROR)
}

^!e:: {  ; Archive
    OutlookMail_EnsureHomeTab()
    if !OutlookClickFirst([{ AutomationId: "505", ControlType: "Button" }, { Name: "Archive", ControlType: "Button" }])
        ShowCenteredOverlay_Utils("❌ Outlook: Archive not found", 1200, BANNER_ACCENT_ERROR)
}

^!u:: {  ; Read/Unread toggle
    OutlookMail_EnsureHomeTab()
    if !OutlookClickFirst([{ AutomationId: "552", ControlType: "Button" }, { Name: "Read / Unread", ControlType: "Button" }])
        ShowCenteredOverlay_Utils("❌ Outlook: Read/Unread not found", 1200, BANNER_ACCENT_ERROR)
}

^!c:: {  ; Categorize
    OutlookMail_EnsureHomeTab()
    if !OutlookClickFirst([{ AutomationId: "509", ControlType: "Button" }, { Name: "Categorize", ControlType: "Button" }])
        ShowCenteredOverlay_Utils("❌ Outlook: Categorize not found", 1200, BANNER_ACCENT_ERROR)
}

^!v:: {  ; Move
    OutlookMail_EnsureHomeTab()
    if !OutlookClickFirst([{ AutomationId: "540", ControlType: "Button" }, { Name: "Move", ControlType: "Button" }])
        ShowCenteredOverlay_Utils("❌ Outlook: Move not found", 1200, BANNER_ACCENT_ERROR)
    else {
        Sleep 80
        Send "{Down 2}"
    }
}

^!i:: {  ; Mail Filter menu
    if !OutlookClickFirst([{ AutomationId: "mailListFilterMenu", ControlType: "Button" }, { Name: "Filter", ControlType: "Button" }])
        ShowCenteredOverlay_Utils("❌ Outlook: Filter not found", 1200, BANNER_ACCENT_ERROR)
}

^!s:: {  ; Mail Sort menu
    if !OutlookClickFirst([{ AutomationId: "mailListSortMenu", ControlType: "Button" }, { Name: "Sorted", matchmode: "Substring",
        ControlType: "Button" }])
        ShowCenteredOverlay_Utils("❌ Outlook: Sort not found", 1200, BANNER_ACCENT_ERROR)
}

; Calendar (main view)
^!n:: {  ; New item (Mail: new message, Calendar: new event)
    Outlook_ActivateMainWindow()
    ; Calendar capture exposes "New event".
    if OutlookClickFirst([{ Name: "New event", matchmode: "Substring", ControlType: "Button" }])
        return
    ; Mail: fall back to built-in new message.
    Send "^n"
}

^!t:: {  ; Today (Calendar) — toolbar "Go to today …" (Chromium; not ElementFromHandle-only)
    if !OutlookCalendar_ClickGoToToday()
        ShowCenteredOverlay_Utils("❌ Outlook: Today not found", 1200, BANNER_ACCENT_ERROR)
}

; Shift + G : Send to General - General
+G::
{
    if IsNewOutlookActive() {
        ; New Outlook: prefer the Quick Step buttons (stable IDs from outlook-mail.md).
        Outlook_ActivateMainWindow()
        OutlookMail_EnsureHomeTab()
        if OutlookClickFirst([{ AutomationId: "c46846eb-0853-7b70-b484-4d7f31f5d9db", ControlType: "RadioButton" }, ; Move to General
        { AutomationId: "c46846eb-0853-7b70-b484-4d7f31f5d9db" }, { Name: "Move to General", ControlType: "RadioButton" }, { Name: "Move to General",
            matchmode: "Substring" }, { Name: "Move to general", matchmode: "Substring" }, { Name: "Move to Gerais",
                matchmode: "Substring" }
        ]) {
            Sleep 120
            Outlook_FocusMailMessageList(true)
            return
        }
    }
    Send "!5"
    Send "O"
    Send "{Home}"
    Send "00"
    Send "{Enter}"
}

; Shift + N : Send to Newsletter - Newsletter
+N::
{
    if IsNewOutlookActive() {
        ; New Outlook: prefer the Quick Step buttons (stable IDs from outlook-mail.md).
        Outlook_ActivateMainWindow()
        OutlookMail_EnsureHomeTab()
        if OutlookClickFirst([{ AutomationId: "91476b25-0fb7-4460-f695-8905582291db", ControlType: "RadioButton" }, ; Move to Newsletter
        { AutomationId: "91476b25-0fb7-4460-f695-8905582291db" }, { Name: "Move to Newsletter", ControlType: "RadioButton" }, { Name: "Move to Newsletter",
            matchmode: "Substring" }, { Name: "Move to newsletter", matchmode: "Substring" }, { Name: "newsletter",
                matchmode: "Substring", ControlType: "RadioButton" }
        ]) {
            Sleep 120
            Outlook_FocusMailMessageList(true)
            return
        }
    }
    Send "!5"
    Send "O"
    Send "{Home}"
    Send "01"
    Send "{Enter}"
}

; Shift + I : Go to Inbox - Inbox
+I::
{
    if IsNewOutlookActive() {
        Outlook_ActivateMainWindow()
        ; New Outlook: click Inbox in the Navigation pane (unique scope).
        if OutlookMail_ClickInboxFolder()
            return
    }
    Send "{Alt}"
    Sleep 60
    Send "6"
    Sleep 80
    Send "^{Home}"
    Sleep 100
    Send "i"
    Sleep 50
    Send "n"
    Sleep 50
    Send "{Enter}"
}

; Shift + J : Jump to first mail and select it
+J::
{
    if IsNewOutlookActive() {
        Outlook_ActivateMainWindow()
        ; Keep behavior mail-centric: if user is in Calendar, switch first.
        Outlook_SwitchToMail()
        if Outlook_FocusMailMessageList(true)
            return
    }

    ; Classic / fallback: try UIA first-item selection (with non-message filter), then keyboard skip-header path.
    if Outlook_FocusMailMessageList(true)
        return
    if Outlook_FocusMailMessageList() {
        Send "{Home}"
        Sleep 40
        if !Outlook_MailList_SkipDrawerHeadersByKeyboard()
            ShowCenteredOverlay_Utils("❌ Outlook: Could not focus a message row", 1200, BANNER_ACCENT_ERROR)
    } else
        ShowCenteredOverlay_Utils("❌ Outlook: Message list not found", 1200, BANNER_ACCENT_ERROR)
}

; Shift + H : Toggle high navigation pane — ribbon Hide/Show folder pane
+H:: {
    if !IsNewOutlookActive()
        return
    Outlook_ActivateMainWindow()
    if !OutlookMail_ToggleHighNavigationPane()
        ShowCenteredOverlay_Utils("❌ Outlook: Navigation pane toggle not found", 1200, BANNER_ACCENT_ERROR)
}

; Shift + S : Subject / Title - Subject
+S:: {
    if FocusOutlookField({ AutomationId: "4101" }) ; Subject
        return
    if FocusOutlookField({ Name: "Subject", ControlType: "Edit" })
        return
}

; Shift + T : Required / To - To (meeting request in reading pane: Tentative via More options)
+T:: {
    if IsOutlookMeetingRequestReadingPaneActive() {
        if !OutlookMeeting_ClickMoreOptionsThen("Tentative")
            ShowCenteredOverlay_Utils("❌ Outlook: Tentative not found", 1200, BANNER_ACCENT_ERROR)
        return
    }
    ; New Outlook compose: “To:” row is a Group (AutomationId 134) that may be collapsed/hidden.
    if IsNewOutlookActive() && IsOutlookComposeActive() {
        ; #region agent log
        OC_STLog(msg, data := "{}", hypo := "OC_ST") {
            try {
                line := "{"
                    . '"sessionId":"b96502",'
                    . '"runId":"shiftT",'
                    . '"hypothesisId":"' hypo '",'
                    . '"timestamp":' A_TickCount ','
                    . '"location":"Shift keys.ahk:+T(compose)",'
                    . '"message":"' StrReplace(msg, '"', '\"') '",'
                    . '"data":' data
                    . "}"
                FileAppend(line "`n", "debug-b96502.log", "UTF-8")
            } catch {
            }
        }
        ; #endregion

        try {
            hwnd := WinExist("A")
            t := WinGetTitle("A")
            OC_STLog("compose_gate_passed", '{"hwnd":' hwnd ',"title":"' StrReplace(SubStr(t, 1, 120), '"', '\"') '"}',
            "OC_ST_A")
        } catch {
        }

        ; Prefer recipient focus flow (reactive UI): To row -> recipient entity group -> inner field.
        ok := OutlookCompose_FocusToRecipientsField()
        try OC_STLog("after_focus_flow", '{"ok":' (ok ? 1 : 0) '}', "OC_ST_B")
        catch {
        }
        if ok
            return

        ; Fallback experiment (logged): select Bcc then Shift+Tab once.
        try {
            bccOk := OutlookClickFirst([{ Name: "Bcc", matchmode: "Substring", ControlType: "Button" }, { Name: "Bcc",
                matchmode: "Substring" }, { Name: "Show Bcc", matchmode: "Substring", ControlType: "Button" }
            ])
            OC_STLog("bcc_click", '{"ok":' (bccOk ? 1 : 0) '}', "OC_ST_C")
            if bccOk {
                Send "{Tab}"
                Sleep 40
                fe := UIA.GetFocusedElement()
                fn := "", fa := "", ft := ""
                try fn := fe.Name
                try fa := fe.AutomationId
                try ft := fe.Type
                OC_STLog("focused_after_bcc_shift_tab", '{"name":"' StrReplace(SubStr(fn, 1, 80), '"', '\"') '","automationId":"' StrReplace(
                    SubStr(fa, 1, 80), '"', '\"') '","type":' (ft = "" ? -1 : ft) '}', "OC_ST_D")
                return
            }
        } catch {
        }

        ; Fallback: click the To row only.
        if OutlookClickFirst([{ AutomationId: "134", ControlType: "Group" }, { AutomationId: "134" }])
            return
        ; Fallback: any element whose name begins with “To:”.
        if OutlookClickFirst([{ Name: "To:", matchmode: "Substring" }, { Name: "To", matchmode: "Substring",
            ControlType: "Group" }])
            return
    }
    if FocusOutlookField({ AutomationId: "4109" }) ; Required
        return
    if FocusOutlookField({ Name: "Required", ControlType: "Edit" })
        return
    if FocusOutlookField({ AutomationId: "4117" }) ; To
        return
    if FocusOutlookField({ Name: "To", ControlType: "Edit" })
        return
}

; Shift + B : Subject -> Body - Body
+B:: {
    if FocusOutlookField({ AutomationId: "4101" }) {
        Sleep 50
        Send "{Tab}"
        return
    }
    if FocusOutlookField({ Name: "Subject", ControlType: "Edit" }) {
        Sleep 50
        Send "{Tab}"
        return
    }
}

; Shift + F : Toggle Focused / Other - Focused
+F:: {                                  ; toggle Focused / Other
    static nextOutlookButton := "Other"

    try {
        btn := OutlookFindFirst([{ Name: nextOutlookButton, ControlType: "TabItem" }, { Name: nextOutlookButton,
            ControlType: "Button" }, { Name: nextOutlookButton, Type: "Button" }
        ])

        if btn {
            btn.Click()
            nextOutlookButton := (nextOutlookButton = "Other")
                ? "Focused" : "Other"
        } else {
            MsgBox("Couldn't find '" nextOutlookButton "'.", "Button not found", "IconX")
        }

    } catch Error as err {              ; â† **only this form**
        ShowErr(err)
    }
}

; Shift+W : Calendar [W]eek view
+W:: {
    try {
        if IsNewOutlookActive() {
            if OutlookClickFirst([{ AutomationId: "2519", ControlType: "Button" }, { Name: "Week", ControlType: "Button" }])
                return
        }
        if !ClickOutlookByIdThenNameClass("WeeklyView", "Week", "NetUIRibbonButton", 50000)
            Send "^!3"
    } catch {
        Send "^!3"
    }
}

; Shift+O : Calendar m[O]nth view
+O:: {
    try {
        if IsNewOutlookActive() {
            if OutlookClickFirst([{ AutomationId: "2505", ControlType: "Button" }, { Name: "Month", ControlType: "Button" }])
                return
        }
        if !ClickOutlookByIdThenNameClass("MonthlyView", "Month", "NetUIRibbonButton", 50000)
            Send "^!4"
    } catch {
        Send "^!4"
    }
}

; Shift + P : Pop Out current item - Pop Out
+P:: {
    try {
        if !ClickOutlookByIdThenNameClass("", "Pop Out", "", 50000) {
            MsgBox("Couldn't find 'Pop Out' button.", "Outlook Pop Out", "IconX")
        }
    } catch Error as err {
        ShowErr(err)
    }
}

; -------------------------------------------------------------------
; Focus helpers â€" reuse for any field you need
; -------------------------------------------------------------------
FocusOutlookField(criteria) {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)
        ctrl := root.FindFirst(criteria)
        if ctrl {
            ctrl.SetFocus()
            return true
        }
    } catch Error {
    }
    return false
}

; -------------------------------------------------------------------
; Click helper â€" try AutomationId first, then Name+ClassName
; -------------------------------------------------------------------
ClickOutlookByIdThenNameClass(automationId, name, className, controlType := "") {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        if (automationId) {
            el := root.FindFirst({ AutomationId: automationId })
            if (el) {
                el.SetFocus()
                Sleep 50
                el.Click()
                return true
            }
        }

        crit := { Name: name }
        if (className)
            crit.ClassName := className
        if (controlType)
            crit.ControlType := controlType

        el := root.FindFirst(crit)
        if (el) {
            el.SetFocus()
            Sleep 50
            el.Click()
            return true
        }
    } catch Error as err {
        ShowErr(err)
    }
    return false
}

; -------------------------------------------------------------------
; General helper â€" visually confirm focus on the selected element
; Sends Down then Up to force a visible focus cue
; -------------------------------------------------------------------
EnsureFocus() {
    Send "{Down}"
    Send "{Up}"
}

; Helper: Select the first pinned item in Explorer sidebar (Navigation Pane)
; Global so it can be reused by both Explorer and File Dialog contexts
SelectExplorerSidebarFirstPinned() {
    try {
        explorerEl := UIA.ElementFromHandle(WinExist("A"))

        ; Look for the navigation pane (sidebar) - it's typically a Tree control
        navPane := explorerEl.FindFirst({ Type: "Tree" })

        if (navPane) {
            ; If in work environment, prefer selecting the Home tree item directly
            try {
                global IS_WORK_ENVIRONMENT
                if (IS_WORK_ENVIRONMENT) {
                    homeItem := navPane.FindFirst({ Type: "TreeItem", Name: "Home" })
                    if (homeItem) {
                        homeItem.ScrollIntoView()
                        homeItem.Select()    ; select only, no click
                        homeItem.SetFocus()
                        EnsureFocus()
                        return true
                    }
                }
            } catch Error {
                ; ignore and fallback to previous logic
            }
            ; Define the keywords to search for pinned items
            pinnedKeywords := ["fixo", "pinned", "pin", "fixado", "fixada", "fixar", "preso"]

            ; Search for the first TreeItem that contains any of the pinned keywords
            firstPinnedItem := unset
            for keyword in pinnedKeywords {
                firstPinnedItem := navPane.FindFirst({ Type: "TreeItem", Name: keyword, matchmode: "Substring" })
                if (firstPinnedItem)
                    break
            }

            ; If no pinned item found by keywords, try to find Desktop by first letter
            ; Portuguese: "Área de Trabalho" starts with "Á" or "a"
            ; English: "Desktop" starts with "D" or "d"
            if (!firstPinnedItem) {
                allTreeItems := navPane.FindAll({ Type: "TreeItem" })
                for item in allTreeItems {
                    try {
                        itemName := item.Name
                        ; Check if name starts with "a" or "Á" (Portuguese Desktop) or "d" or "D" (English Desktop)
                        ; Case-insensitive check
                        firstChar := SubStr(itemName, 1, 1)
                        if (firstChar = "a" || firstChar = "A" || firstChar = "Á" || firstChar = "á" ||
                            firstChar = "d" || firstChar = "D") {
                            ; Additional check: must be Desktop-related (not just any item starting with a/d)
                            if (InStr(itemName, "Desktop", false) || InStr(itemName, "Área de Trabalho", false) ||
                            InStr(itemName, "Trabalho", false)) {
                                firstPinnedItem := item
                                break
                            }
                        }
                    } catch {
                        ; Skip items without names
                    }
                }
            }

            if (firstPinnedItem) {
                firstPinnedItem.ScrollIntoView()
                firstPinnedItem.Select()
                firstPinnedItem.SetFocus()
                EnsureFocus()
                return true
            }

            ; If we didn't find a pinned item, at least focus the tree and press Home
            navPane.SetFocus()
            Sleep 100
            Send "{Home}"
            EnsureFocus()
            return false
        }
    } catch Error {
        ; swallow and continue to fallback
    }

    ; Robust fallback â€" cycle through panes up to 6 times to reach navigation, then Home
    loop 6 {
        Send "{F6}"
        Sleep 120
        try {
            explorerEl := UIA.ElementFromHandle(WinExist("A"))
            navPane := explorerEl.FindFirst({ Type: "Tree" })
            if (navPane && navPane.HasKeyboardFocus) {
                Send "{Home}"
                EnsureFocus()
                return false
            }
        } catch Error {
        }
    }
    ; Last resort â€" send Home anyway
    Send "{Home}"
    EnsureFocus()
    return false
}

; Shift + K : Send Shift+F6
+K:: Send "+{F6}"

; Shift + L : Send F6
+L:: Send "{F6}"

; Shift + M : Toggle Mail / Calendar - Mail/Calendar
+M:: {
    try {
        if IsNewOutlookActive() {
            if Outlook_ToggleMailCalendarRail()
                return
            ; Fallback if rail toggles not found / no toggle pattern
            t := WinGetTitle("A")
            if RegExMatch(t, "i)Calendar") {
                if Outlook_SwitchToMail()
                    return
            } else {
                if Outlook_SwitchToCalendar()
                    return
            }
        }
        root := UIA.ElementFromHandle(WinExist("A"))

        ; Find Mail and Calendar list items
        mailItem := root.FindFirst({ Name: "Mail", Type: "50007" })
        if !mailItem {
            mailItem := root.FindFirst({ Name: "Mail", ClassName: "NetUIListViewItem" })
        }

        calendarItem := root.FindFirst({ Name: "Calendar", Type: "50007" })
        if !calendarItem {
            calendarItem := root.FindFirst({ Name: "Calendar", ClassName: "NetUIListViewItem" })
        }

        ; Check which is selected and toggle
        if (mailItem && calendarItem) {
            try {
                isMailSelected := mailItem.IsSelected
                isCalendarSelected := calendarItem.IsSelected

                if (isMailSelected) {
                    calendarItem.SetFocus()
                    Sleep 50
                    calendarItem.Click()
                } else {
                    mailItem.SetFocus()
                    Sleep 50
                    mailItem.Click()
                }
            } catch Error as err {
                ; Fallback: if pattern check fails, try clicking Calendar
                calendarItem.SetFocus()
                Sleep 50
                calendarItem.Click()
            }
        } else {
            MsgBox "Could not find Mail or Calendar items.", "Outlook Toggle", "IconX"
        }
    } catch Error as err {
        MsgBox "Error toggling Mail/Calendar:`n" err.Message, "Outlook Toggle", "IconX"
    }
}

; Meeting request - main Mail window reading pane only (not popped-out invitation window).
#HotIf IsOutlookMainActive() && IsOutlookMeetingRequestReadingPaneActive()

; Accept / Follow (header row). Alt+F = Follow (Shift+F stays Focused / Other in the base hotkey above).
+A:: {
    if !OutlookMeeting_ClickAccept()
        ShowCenteredOverlay_Utils("❌ Outlook: Accept the meeting not found", 1200, BANNER_ACCENT_ERROR)
}

!f:: {
    if !OutlookMeeting_ClickFollow()
        ShowCenteredOverlay_Utils("❌ Outlook: Follow not found", 1200, BANNER_ACCENT_ERROR)
}

+D:: {
    if !OutlookMeeting_ConfirmDecline()
        return
    if !OutlookMeeting_ClickDecline()
        ShowCenteredOverlay_Utils("❌ Outlook: Decline the meeting not found", 1200, BANNER_ACCENT_ERROR)
}

; Message inspector-specific hotkeys (Subject/To/Body)
#HotIf IsOutlookMessageActive()

; Shift + S : Subject / Title - Subject
+S:: {
    if FocusOutlookField({ AutomationId: "4101" }) ; Subject
        return
    if FocusOutlookField({ Name: "Subject", ControlType: "Edit" })
        return
}

; Shift + T : Required / To - To
+T:: {
    if FocusOutlookField({ AutomationId: "4109" }) ; Required
        return
    if FocusOutlookField({ Name: "Required", ControlType: "Edit" })
        return
    if FocusOutlookField({ AutomationId: "4117" }) ; To
        return
    if FocusOutlookField({ Name: "To", ControlType: "Edit" })
        return
}

; Shift + B : Body (Subject -> Body) - Body
+B:: {
    if FocusOutlookField({ AutomationId: "4101" }) {
        Sleep 50
        Send "{Tab}"
        return
    }
    if FocusOutlookField({ Name: "Subject", ControlType: "Edit" }) {
        Sleep 50
        Send "{Tab}"
        return
    }
}

; Popped-out meeting invitation (same Accept/Follow/Tentative as reading pane; overrides generic message Shift+T / Alt+F when both match).
#HotIf IsOutlookMeetingInvitationPopOutActive()

+A:: {
    if !OutlookMeeting_ClickAccept()
        ShowCenteredOverlay_Utils("❌ Outlook: Accept the meeting not found", 1200, BANNER_ACCENT_ERROR)
}

!f:: {
    if !OutlookMeeting_ClickFollow()
        ShowCenteredOverlay_Utils("❌ Outlook: Follow not found", 1200, BANNER_ACCENT_ERROR)
}

+T:: {
    if !OutlookMeeting_ClickMoreOptionsThen("Tentative")
        ShowCenteredOverlay_Utils("❌ Outlook: Tentative not found", 1200, BANNER_ACCENT_ERROR)
}

+D:: {
    if !OutlookMeeting_ConfirmDecline()
        return
    if !OutlookMeeting_ClickDecline()
        ShowCenteredOverlay_Utils("❌ Outlook: Decline the meeting not found", 1200, BANNER_ACCENT_ERROR)
}

; Canceled meeting — Button "Remove event" (Shift+E); no confirmation (see outlook-remove-evet.md).
#HotIf IsOutlookMainActive() && IsOutlookRemoveEventReadingPaneActive()

+E:: {
    if !OutlookMeeting_ClickRemoveEvent()
        ShowCenteredOverlay_Utils("❌ Outlook: Remove event not found", 1200, BANNER_ACCENT_ERROR)
}

#HotIf IsOutlookRemoveEventPopOutActive()

+E:: {
    if !OutlookMeeting_ClickRemoveEvent()
        ShowCenteredOverlay_Utils("❌ Outlook: Remove event not found", 1200, BANNER_ACCENT_ERROR)
}

#HotIf

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