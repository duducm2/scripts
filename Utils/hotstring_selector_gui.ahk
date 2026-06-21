; =============================================================================
; Utils module: hotstring_selector_gui.ahk
; ShowHotstringSelector GUI and category view
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

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
