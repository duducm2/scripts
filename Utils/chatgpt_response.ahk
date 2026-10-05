; ChatGPT stop / composer / copy / read-aloud helpers.
; Included from Utils.ahk so D2C_CompanionHasStop can call them in every host
; (AppLaunchers, Gemini, Shift keys). Do not also define these in Shift keys.

ChatGPT_FindStopButton(uia) {
    if (!IsObject(uia))
        return 0
    try {
        saved := UiElements_TrySaved(uia, "ChatGPT", "Stop")
        if (saved)
            return saved
    } catch {
    }
    for n in ["Stop streaming", "Interromper transmissão", "Stop", "Interromper"] {
        try {
            el := uia.FindFirst({ Name: n, Type: 50000 })
            if (el)
                return el
        } catch {
        }
    }
    return 0
}

ChatGPT_FocusComposer(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return 0
    root := 0
    try root := UIA.ElementFromHandle(hwnd)
    catch
        return 0
    el := AiCompanion_FindComposerElement(root, "chatgpt")
    if (!IsObject(el))
        return 0
    try el.SetFocus()
    catch {
    }
    return el
}

; Select-all copy for composers whose Value stays blank. Restores the previous clipboard.
ChatGPT_ComposerGetTextViaClipboard(hwnd := 0) {
    if (!hwnd)
        hwnd := WinExist("A")
    if (!ChatGPT_FocusComposer(hwnd))
        return ""
    Sleep 60
    saved := ClipboardAll()
    try {
        A_Clipboard := ""
        Send "^a"
        Sleep 40
        Send "^c"
        if !ClipWait(1, 1)
            return ""
        text := A_Clipboard
        if (Type(text) != "String")
            text := ""
        text := Trim(text)
        if (text = "" || AiCompanion_IsComposerPlaceholder(text))
            return ""
        return text
    } finally {
        Sleep 40
        try A_Clipboard := saved
        catch {
        }
    }
}

ChatGPT_IsCopyResponseButton(name) {
    if (!name)
        return false
    if (InStr(name, "prompt", false) || InStr(name, "code", false))
        return false
    return (name = "Copy" || name = "Copiar" || InStr(name, "Copy response", false) = 1)
}

ChatGPT_CopyLastMessageToClipboard(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    try ChromeChat_ScrollFeedToBottomFast(hwnd)
    catch {
    }
    root := 0
    try root := UIA.ElementFromHandle(hwnd)
    catch
        return false
    if (!IsObject(root))
        return false
    lastEl := 0
    lastTop := ""
    try {
        buttons := root.FindAll({ Type: 50000 })
        for button in buttons {
            name := ""
            try name := button.Name
            catch
                continue
            if (!ChatGPT_IsCopyResponseButton(name))
                continue
            try br := button.BoundingRectangle
            catch
                continue
            if (!IsObject(br))
                continue
            if (lastEl = 0 || br.t >= lastTop) {
                lastEl := button
                lastTop := br.t
            }
        }
    } catch {
        return false
    }
    if (!IsObject(lastEl))
        return false
    seqBefore := Clipboard_GetSequenceNumber()
    if (!AiCompanionButtons_Click(lastEl))
        return false
    if (!Clipboard_WaitForSequenceChange(seqBefore, 2000, 850))
        return false
    clip := Trim(A_Clipboard)
    return (clip != "" && StrLen(clip) >= 10)
}

ChatGPT_TriggerReadAloud(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    root := 0
    try root := UIA.ElementFromHandle(hwnd)
    catch
        return false
    if (!IsObject(root))
        return false
    for n in ["Read aloud", "Read Aloud", "Ler em voz alta"] {
        try {
            el := root.FindFirst({ Name: n, Type: 50000 })
            if (IsObject(el) && AiCompanionButtons_Click(el))
                return true
        } catch {
        }
        try {
            el := root.FindFirst({ Name: n, Type: 50011 })
            if (IsObject(el) && AiCompanionButtons_Click(el))
                return true
        } catch {
        }
    }
    return false
}
