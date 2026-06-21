; =============================================================================
; Shift keys module: config.ahk
; Config and cheat-sheet string utilities
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

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
