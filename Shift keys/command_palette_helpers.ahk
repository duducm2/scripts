; =============================================================================
; Shift keys module: command_palette_helpers.ahk
; Command Palette web-bookmark detection and new-Chrome-window open strategies.
; Loaded via #include into Shift keys.ahk before hotif_command_palette.ahk.
; =============================================================================

; True when the selected result is a web bookmark (footer shows Copy address).
CommandPalette_IsWebBookmarkSelected(hwnd := 0) {
    try {
        hwnd := hwnd ? hwnd : WinExist("A")
        if !hwnd
            return false
        root := UIA.ElementFromHandle(hwnd)
        if root.FindFirst({ Name: "Copy address", matchmode: "Substring" })
            return true
    } catch {
    }
    return CommandPalette_SelectedResultHasUrl(hwnd)
}

; Fallback: selected list item name or child text contains http(s) URL.
CommandPalette_SelectedResultHasUrl(hwnd := 0) {
    try {
        hwnd := hwnd ? hwnd : WinExist("A")
        if !hwnd
            return false
        root := UIA.ElementFromHandle(hwnd)
        for item in root.FindAll({ ControlType: "ListItem" }) {
            try {
                if !(item.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable)
                && item.SelectionItemPattern.IsSelected)
                    continue
                if RegExMatch(item.Name, "i)https?://")
                    return true
                for child in item.FindAll({ matchmode: "Substring", Name: "http" }) {
                    try {
                        if RegExMatch(child.Name, "i)https?://")
                            return true
                    } catch
                        continue
                }
            } catch
                continue
        }
    } catch {
    }
    return false
}

CommandPalette_OpenWebInNewChrome_Detach() {
    Send "{Enter}"
    if !WinWaitActive("ahk_exe chrome.exe", , 5) {
        ToolTip "Could not activate Chrome after opening bookmark.", 10, 10
        SetTimer(() => ToolTip(), -2800)
        return false
    }
    Sleep 150
    if !Chrome_DetachActiveTabToNewWindow() {
        ToolTip "Could not detach tab to a new Chrome window.", 10, 10
        SetTimer(() => ToolTip(), -2800)
        return false
    }
    return true
}

; Copy address (Ctrl+Enter) then chrome.exe --new-window. Returns false on failure.
CommandPalette_OpenWebInNewChrome_CopyRun() {
    savedClip := A_Clipboard
    A_Clipboard := ""
    try {
        Send "^{Enter}"
        if !ClipWait(2) || !RegExMatch(A_Clipboard, "i)^https?://")
            return false
        StudyLink_OpenUrlInChrome(A_Clipboard, true)
        return true
    } finally {
        A_Clipboard := savedClip
    }
}

CommandPalette_OpenWebInNewChrome() {
    global COMMAND_PALETTE_WEB_OPEN_MODE
    if (COMMAND_PALETTE_WEB_OPEN_MODE = "detach")
        return CommandPalette_OpenWebInNewChrome_Detach()
    if CommandPalette_OpenWebInNewChrome_CopyRun()
        return true
    return CommandPalette_OpenWebInNewChrome_Detach()
}

CommandPalette_ActivateSelectedItem() {
    if CommandPalette_IsWebBookmarkSelected() {
        CommandPalette_OpenWebInNewChrome()
        return
    }
    Send "{Enter}"
}

CommandPalette_SelectNthAndActivate(downCount := 0) {
    loop downCount
        Send "{Down}"
    if downCount > 0
        Sleep 30
    CommandPalette_ActivateSelectedItem()
}

; Open Ctrl+K context menu and activate "Editar favorito" / "Edit favorite".
; Focus stays on ContextFilterBox, so use Name FindFirst (not focus-walk / full FindAll).
CommandPalette_EditFavoriteSelected() {
    Send "^k"

    deadline := A_TickCount + 1500
    downCount := 0
    maxDowns := 12
    while (A_TickCount < deadline) {
        el := CommandPalette_FindEditFavoriteElement()
        if el {
            if CommandPalette_ActivateMenuElement(el)
                return
            Send "{Enter}"
            return
        }
        ; Menu items can appear without keyboard focus; Down only if FindFirst still misses.
        if (downCount < maxDowns) {
            Send "{Down}"
            downCount += 1
        }
        Sleep 40
    }

    Send "{Escape}"
    ToolTip "Edit favorite not found in context menu.", 10, 10
    SetTimer(() => ToolTip(), -2800)
}

; Prefer targeted FindFirst (one COM query) over FindAll of every Button/MenuItem/ListItem.
CommandPalette_FindEditFavoriteElement() {
    hwnd := WinExist("A")
    if !hwnd
        return 0
    try {
        root := UIA.ElementFromHandle(hwnd)
        for name in ["Editar favorito", "Edit favorite"] {
            try {
                el := root.FindFirst({ Name: name, matchmode: "Substring", cs: false })
                if el
                    return el
            } catch {
            }
        }
        try {
            el := root.FindFirst({ Name: "favorit", matchmode: "Substring", cs: false })
            if CommandPalette_ElementLooksLikeEditFavorite(el)
                return el
        } catch {
        }
    } catch {
    }
    return 0
}

CommandPalette_ElementLooksLikeEditFavorite(el) {
    if !el
        return false
    nm := ""
    try nm := el.Name
    if RegExMatch(nm, "i)(edit|editar)") && RegExMatch(nm, "i)favorit")
        return true
    haystack := nm
    try haystack .= " " el.GetPropertyValue(UIA.Property.HelpText)
    try haystack .= " " el.AutomationId
    return RegExMatch(haystack, "i)(edit|editar)") && RegExMatch(haystack, "i)favorit")
}

CommandPalette_ActivateMenuElement(el) {
    if !el
        return false
    try {
        if (el.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)) {
            el.InvokePattern.Invoke()
            return true
        }
    } catch {
    }
    try {
        el.Click()
        return true
    } catch {
    }
    try {
        el.SetFocus()
        Sleep 40
        Send "{Enter}"
        return true
    } catch {
    }
    return false
}
