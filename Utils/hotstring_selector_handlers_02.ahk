; =============================================================================
; Utils module: hotstring_selector_handlers_02.ahk
; Hotstring selector utility category handlers
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

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
    SetTimer(HotstringSelector_AutoCloseIfIdle, -g_HotstringSelectorAutoCloseMs)
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
