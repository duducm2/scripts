; =============================================================================
; Utils module: hotstring_selector_core.ahk
; Hotstring selector system core and BuildHotstringCharMap
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

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
; Modal ListView — first-letter row jump (reusable for ListView modals)
; =============================================================================
ModalList_FindFirstByStartingLetter(entries, letter, getNameFn := "") {
    ch := StrLower(SubStr(letter, 1, 1))
    if !RegExMatch(ch, "^[a-z]$")
        return 0
    for i, entry in entries {
        name := getNameFn ? getNameFn(entry) : entry.name
        if (name = "")
            continue
        if (SubStr(StrLower(name), 1, 1) = ch)
            return i
    }
    return 0
}

ListView_SelectRowFocused(lv, rowNum) {
    if (!IsObject(lv) || rowNum < 1)
        return
    lv.Modify(rowNum, "Select Focus Vis")
    try SendMessage(0x1117, rowNum - 1, 0, lv)
    try lv.Focus()
}

ModalListLetterJump_Stop(&hookRef) {
    if (IsObject(hookRef) && hookRef.HasProp("handlers")) {
        for item in hookRef.handlers {
            try Hotkey(item.char, item.handler, "Off")
            try Hotkey(StrUpper(item.char), item.handler, "Off")
        }
    }
    hookRef := ""
}

ModalListLetterJump_CreateHandler(char, isActiveFn, getEntriesFn, getNameFn, onMatchFn) {
    return (*) => ModalListLetterJump_HandleChar(char, isActiveFn, getEntriesFn, getNameFn, onMatchFn)
}

ModalListLetterJump_HandleChar(char, isActiveFn, getEntriesFn, getNameFn, onMatchFn) {
    if (!isActiveFn())
        return
    rowNum := ModalList_FindFirstByStartingLetter(getEntriesFn(), char, getNameFn)
    if (rowNum > 0)
        onMatchFn(rowNum)
}

ModalListLetterJump_Start(&hookRef, isActiveFn, getEntriesFn, getNameFn, onMatchFn) {
    ModalListLetterJump_Stop(&hookRef)
    state := { handlers: [] }
    loop 26 {
        ch := Chr(96 + A_Index)
        handler := ModalListLetterJump_CreateHandler(ch, isActiveFn, getEntriesFn, getNameFn, onMatchFn)
        state.handlers.Push({ char: ch, handler: handler })
        try Hotkey(ch, handler, "On")
        try Hotkey(StrUpper(ch), handler, "On")
    }
    hookRef := state
}
