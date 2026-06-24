; =============================================================================
; Utils module: modal_list_letter_jump.ahk
; Modal ListView first-letter row jump (used by context file browser).
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
