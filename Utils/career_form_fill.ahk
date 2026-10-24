; =============================================================================
; Career cockpit form fill (standalone, command-line).
; Called by the career server. Does not register hotkeys.
;
;   career_form_fill.ahk --action capture-window --result <json>
;   career_form_fill.ahk --action open-uia --result <json>
;   career_form_fill.ahk --action read-clipboard --out <file> --result <json>
;   career_form_fill.ahk --action send-prompt --file <prompt> --result <json>
;   career_form_fill.ahk --action fill-form --plan <plan> --result <json>
;
; send-prompt expects the caller to press Win+Alt+Shift+I first, so the
; running scripts repo can focus the resolved AI companion. This script then
; pastes and presses Enter.
; fill-form activates the Chrome window whose title contains the saved tab title.
; =============================================================================

#Requires AutoHotkey v2.0
#SingleInstance Off

Utils_EnsureGlobalEscapeHotkey() {
}

#Include "%A_ScriptDir%\standard_loading_bar.ahk"
#Include "%A_ScriptDir%\..\vendor\UIA-v2\Lib\UIA.ahk"

global g_CareerFormResultPath := ""

CareerForm_ParseArgs() {
    args := Map()
    i := 1
    while (i <= A_Args.Length) {
        key := A_Args[i]
        if (SubStr(key, 1, 2) != "--") {
            i += 1
            continue
        }
        name := SubStr(key, 3)
        if (i < A_Args.Length && SubStr(A_Args[i + 1], 1, 2) != "--") {
            args[name] := A_Args[i + 1]
            i += 2
        } else {
            args[name] := "1"
            i += 1
        }
    }
    return args
}

CareerForm_JsonString(value) {
    text := String(value)
    text := StrReplace(text, "\", "\\")
    text := StrReplace(text, '"', '\"')
    text := StrReplace(text, "`r", "\r")
    text := StrReplace(text, "`n", "\n")
    text := StrReplace(text, "`t", "\t")
    return '"' text '"'
}

CareerForm_WriteResult(obj) {
    global g_CareerFormResultPath
    if (g_CareerFormResultPath = "")
        return
    parts := ""
    for key, value in obj.OwnProps() {
        piece := CareerForm_JsonString(key) ": "
        if (key = "windows") {
            items := ""
            for title in value {
                items .= (items = "" ? "" : ",") CareerForm_JsonString(title)
            }
            piece .= "[" items "]"
        } else if (key = "errors") {
            items := ""
            for message in value {
                items .= (items = "" ? "" : ",") CareerForm_JsonString(message)
            }
            piece .= "[" items "]"
        } else if IsInteger(value)
            piece .= String(value)
        else
            piece .= CareerForm_JsonString(value)
        parts .= (parts = "" ? "" : ",") piece
    }
    try {
        file := FileOpen(g_CareerFormResultPath, "w", "UTF-8")
        if (file) {
            file.Write("{" parts "}")
            file.Close()
        }
    } catch {
    }
}

CareerForm_Fail(message) {
    CareerForm_WriteResult({ ok: false, error: message })
    try StandardLoadingBar_Hide(0)
    catch {
    }
    ExitApp 1
}

CareerForm_IsCareerCockpitTitle(title) {
    text := StrLower(title)
    return InStr(text, "job openings") || InStr(text, ":8770")
}

CareerForm_IsCompanionTitle(title) {
    if (CareerForm_IsCareerCockpitTitle(title))
        return false
    text := StrLower(title)
    return InStr(text, "gemini") || InStr(text, "copilot") || InStr(text, "askbosch")
}

CareerForm_CaptureWindow() {
    titles := []
    suggested := ""
    try {
        for hwnd in WinGetList("ahk_exe chrome.exe") {
            title := ""
            try title := Trim(WinGetTitle("ahk_id " hwnd))
            catch {
                continue
            }
            if (title = "" || CareerForm_IsCareerCockpitTitle(title))
                continue
            titles.Push(title)
            if (suggested = "")
                suggested := title
        }
    } catch as err {
        CareerForm_Fail(err.Message)
    }
    CareerForm_WriteResult({ ok: true, window_name: suggested, windows: titles, error: "" })
}

CareerForm_OpenUia() {
    inspector := A_ScriptDir "\..\vendor\UIA-v2\UIATreeInspector.ahk"
    if !FileExist(inspector) {
        CareerForm_Fail("UIA Tree Inspector was not found.")
    }
    hwnd := 0
    try hwnd := WinExist("UIATreeInspector")
    catch {
        hwnd := 0
    }
    if (hwnd) {
        try WinActivate("ahk_id " hwnd)
        catch {
        }
        CareerForm_WriteResult({ ok: true, error: "" })
        return
    }
    try Run('"' A_AhkPath '" "' inspector '"')
    catch as err {
        CareerForm_Fail(err.Message)
    }
    CareerForm_WriteResult({ ok: true, error: "" })
}

CareerForm_ClipboardUnicode() {
    ; A_Clipboard and ClipboardAll can block when another script watches the clipboard.
    opened := false
    loop 20 {
        if DllCall("OpenClipboard", "Ptr", 0) {
            opened := true
            break
        }
        Sleep 50
    }
    if (!opened)
        return ""
    text := ""
    try {
        handle := DllCall("GetClipboardData", "UInt", 13, "Ptr")
        if (handle) {
            ptr := DllCall("GlobalLock", "Ptr", handle, "Ptr")
            if (ptr) {
                text := StrGet(ptr, "UTF-16")
                DllCall("GlobalUnlock", "Ptr", handle)
            }
        }
    } catch {
        text := ""
    }
    DllCall("CloseClipboard")
    return text
}

CareerForm_ReadClipboard(outPath) {
    if (outPath = "")
        CareerForm_Fail("Missing --out for the clipboard file.")
    text := CareerForm_ClipboardUnicode()
    try {
        file := FileOpen(outPath, "w", "UTF-8")
        if (!file)
            CareerForm_Fail("Could not write the clipboard file.")
        file.Write(text)
        file.Close()
    } catch as err {
        CareerForm_Fail(err.Message)
    }
    CareerForm_WriteResult({ ok: true, chars: StrLen(text), error: "" })
}

CareerForm_WaitForCompanion(timeoutMs := 20000) {
    start := A_TickCount
    loop {
        hwnd := 0
        try hwnd := WinExist("A")
        catch {
            hwnd := 0
        }
        if (hwnd) {
            proc := ""
            title := ""
            try proc := WinGetProcessName("ahk_id " hwnd)
            catch {
                proc := ""
            }
            try title := WinGetTitle("ahk_id " hwnd)
            catch {
                title := ""
            }
            if (proc = "chrome.exe" && CareerForm_IsCompanionTitle(title)) {
                Sleep 600
                return hwnd
            }
        }
        if ((A_TickCount - start) >= timeoutMs)
            return 0
        Sleep 200
    }
}

CareerForm_SendPrompt(promptPath) {
    if (promptPath = "" || !FileExist(promptPath))
        CareerForm_Fail("Prompt file was not found.")
    prompt := ""
    try {
        file := FileOpen(promptPath, "r", "UTF-8")
        if (!file)
            CareerForm_Fail("Could not read the prompt file.")
        prompt := file.Read()
        file.Close()
    } catch as err {
        CareerForm_Fail(err.Message)
    }
    if (Trim(prompt) = "")
        CareerForm_Fail("Prompt file is empty.")

    try StandardLoadingBar_Show("Sending the form prompt to the AI companion…", BANNER_ACCENT_INTERMEDIATE, { passive: true })
    catch {
    }
    hwnd := CareerForm_WaitForCompanion()
    if (!hwnd) {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        CareerForm_Fail(
            "The AI companion did not come forward. Leave Gemini, Gemini Enterprise, or Copilot available, then try again."
        )
    }
    saved := CareerForm_ClipboardUnicode()
    errMsg := ""
    try {
        if !WinActive("ahk_id " hwnd) {
            WinActivate("ahk_id " hwnd)
            WinWaitActive("ahk_id " hwnd, , 2)
        }
        A_Clipboard := ""
        A_Clipboard := prompt
        if !ClipWait(1.5)
            throw Error("The prompt did not reach the clipboard.")
        Sleep 120
        Send "^a"
        Sleep 40
        Send "^v"
        Sleep 250
        Send "{Enter}"
    } catch Error as err {
        errMsg := err.Message
    } finally {
        Sleep 200
        if (saved != "") {
            try A_Clipboard := saved
            catch {
            }
        }
        try StandardLoadingBar_Hide(0)
        catch {
        }
    }
    if (errMsg != "")
        CareerForm_Fail(errMsg)
    CareerForm_WriteResult({ ok: true, chars: StrLen(prompt), error: "" })
}

CareerForm_ReadLine(text, &pos) {
    end := InStr(text, "`n", false, pos)
    if (!end)
        return ""
    line := SubStr(text, pos, end - pos)
    line := RTrim(line, "`r")
    pos := end + 1
    return line
}

CareerForm_ReadField(text, &pos) {
    header := CareerForm_ReadLine(text, &pos)
    if !RegExMatch(header, "^([A-Z]+) (\d+)$", &match)
        throw Error("Bad fill-plan header: " header)
    count := Integer(match[2])
    value := ""
    if (count > 0)
        value := SubStr(text, pos, count)
    pos += count
    if (SubStr(text, pos, 1) = "`r")
        pos += 1
    if (SubStr(text, pos, 1) = "`n")
        pos += 1
    return { tag: match[1], value: value }
}

CareerForm_ReadPlan(planPath) {
    file := FileOpen(planPath, "r", "UTF-8")
    if (!file)
        throw Error("Could not read the fill plan.")
    text := file.Read()
    file.Close()
    if (SubStr(text, 1, 1) = Chr(0xFEFF))
        text := SubStr(text, 2)
    pos := 1
    windowField := CareerForm_ReadField(text, &pos)
    if (windowField.tag != "W")
        throw Error("Fill plan is missing the window name.")
    countLine := CareerForm_ReadLine(text, &pos)
    if !RegExMatch(countLine, "^N (\d+)$", &countMatch)
        throw Error("Fill plan is missing the step count.")
    steps := []
    loop Integer(countMatch[1]) {
        action := CareerForm_ReadField(text, &pos)
        typeName := CareerForm_ReadField(text, &pos)
        name := CareerForm_ReadField(text, &pos)
        automationId := CareerForm_ReadField(text, &pos)
        path := CareerForm_ReadField(text, &pos)
        value := CareerForm_ReadField(text, &pos)
        label := CareerForm_ReadField(text, &pos)
        steps.Push({
            action: action.value,
            type: typeName.value,
            name: name.value,
            automationId: automationId.value,
            path: path.value,
            value: value.value,
            label: label.value
        })
    }
    return { window: windowField.value, steps: steps }
}

CareerForm_FindChromeByTitle(windowName) {
    needle := Trim(windowName)
    if (needle = "")
        return 0
    needleLower := StrLower(needle)
    best := 0
    bestScore := 0
    for hwnd in WinGetList("ahk_exe chrome.exe") {
        title := ""
        try title := WinGetTitle("ahk_id " hwnd)
        catch {
            continue
        }
        if (title = "" || CareerForm_IsCareerCockpitTitle(title))
            continue
        titleLower := StrLower(title)
        score := 0
        if (titleLower = needleLower)
            score := 3
        else if InStr(titleLower, needleLower)
            score := 2
        else if InStr(needleLower, titleLower)
            score := 1
        if (score > bestScore) {
            bestScore := score
            best := hwnd
        }
    }
    return best
}

CareerForm_ElementByPath(root, path) {
    if (!IsObject(root) || path = "")
        return root
    el := root
    for part in StrSplit(path, ",") {
        idx := Integer(part)
        kids := 0
        try kids := el.GetChildren()
        catch {
            return 0
        }
        if (!IsObject(kids) || idx < 1 || idx > kids.Length)
            return 0
        el := kids[idx]
        if (!IsObject(el))
            return 0
    }
    return el
}

CareerForm_FindControl(root, step) {
    found := 0
    if (step.path != "") {
        try found := CareerForm_ElementByPath(root, step.path)
        catch {
            found := 0
        }
        if IsObject(found)
            return found
    }
    aid := Trim(step.automationId)
    if (aid != "" && !RegExMatch(aid, "^\d+$")) {
        try found := root.FindFirst({ AutomationId: aid })
        catch {
            found := 0
        }
        if IsObject(found)
            return found
    }
    if (step.name != "" && step.type != "") {
        try found := root.FindFirst({ Type: step.type, Name: step.name })
        catch {
            found := 0
        }
        if IsObject(found)
            return found
    }
    return 0
}

CareerForm_PasteValue(value) {
    A_Clipboard := ""
    A_Clipboard := value
    if !ClipWait(1)
        return false
    Sleep 40
    Send "^a"
    Sleep 30
    Send "^v"
    Sleep 80
    return true
}

CareerForm_SetText(el, value) {
    try el.SetFocus()
    catch {
    }
    Sleep 60
    wrote := false
    try {
        el.Value := value
        wrote := true
    } catch {
        wrote := false
    }
    current := ""
    try current := el.Value
    catch {
        current := ""
    }
    if (wrote && current = value)
        return true
    return CareerForm_PasteValue(value)
}

CareerForm_Yes(value) {
    text := StrLower(Trim(value))
    return text = "1" || text = "true" || text = "yes" || text = "y" || text = "on" || text = "sim" || text = "checked"
}

CareerForm_Check(el, value) {
    want := CareerForm_Yes(value)
    state := -1
    try state := el.TogglePattern.ToggleState
    catch {
        state := -1
    }
    if (state = 1 && want)
        return true
    if (state = 0 && !want)
        return true
    try {
        el.TogglePattern.Toggle()
        return true
    } catch {
    }
    try {
        el.Click()
        return true
    } catch {
        return false
    }
}

CareerForm_Select(el, value) {
    try el.SetFocus()
    catch {
    }
    try el.ExpandCollapsePattern.Expand()
    catch {
        try el.Click()
        catch {
        }
    }
    Sleep 180
    if !CareerForm_PasteValue(value)
        return false
    Sleep 180
    Send "{Enter}"
    Sleep 80
    return true
}

CareerForm_Click(el) {
    try {
        el.Click()
        return true
    } catch {
    }
    try {
        el.SetFocus()
        Send "{Space}"
        return true
    } catch {
        return false
    }
}

CareerForm_FillForm(planPath) {
    if (planPath = "" || !FileExist(planPath))
        CareerForm_Fail("Fill plan was not found.")
    plan := ""
    try plan := CareerForm_ReadPlan(planPath)
    catch as err {
        CareerForm_Fail(err.Message)
    }
    hwnd := CareerForm_FindChromeByTitle(plan.window)
    if (!hwnd)
        CareerForm_Fail("No Chrome window title contains: " plan.window)
    try {
        WinActivate("ahk_id " hwnd)
        if !WinWaitActive("ahk_id " hwnd, , 2)
            CareerForm_Fail("Could not activate the application window.")
    } catch as err {
        CareerForm_Fail(err.Message)
    }
    root := 0
    try root := UIA.ElementFromHandle(hwnd)
    catch as err {
        CareerForm_Fail("Could not read the window UI: " err.Message)
    }
    if (!IsObject(root))
        CareerForm_Fail("Could not read the window UI.")

    try StandardLoadingBar_Show("Filling the application form…", BANNER_ACCENT_INTERMEDIATE, { passive: false,
        centerOnHwnd: hwnd })
    catch {
    }
    filled := 0
    skipped := 0
    errors := []
    saved := CareerForm_ClipboardUnicode()
    try {
        for step in plan.steps {
            label := step.label != "" ? step.label : step.name
            try StandardLoadingBar_Update("Filling " label)
            catch {
            }
            el := CareerForm_FindControl(root, step)
            if (!IsObject(el)) {
                skipped += 1
                errors.Push("Not found: " label)
                continue
            }
            ok := false
            if (step.action = "check")
                ok := CareerForm_Check(el, step.value)
            else if (step.action = "click")
                ok := CareerForm_Click(el)
            else if (step.action = "select")
                ok := CareerForm_Select(el, step.value)
            else
                ok := CareerForm_SetText(el, step.value)
            if (ok)
                filled += 1
            else {
                skipped += 1
                errors.Push("Could not set: " label)
            }
            Sleep 90
        }
    } finally {
        if (saved != "") {
            try A_Clipboard := saved
            catch {
            }
        }
        try {
            StandardLoadingBar_Update("Filled " filled " field" (filled = 1 ? "" : "s"))
            Sleep 700
            StandardLoadingBar_Hide(0)
        } catch {
        }
    }
    CareerForm_WriteResult({
        ok: true,
        filled: filled,
        skipped: skipped,
        errors: errors,
        error: ""
    })
}

try {
    args := CareerForm_ParseArgs()
    if args.Has("result")
        g_CareerFormResultPath := args["result"]
    action := args.Has("action") ? args["action"] : ""
    if (action = "capture-window")
        CareerForm_CaptureWindow()
    else if (action = "open-uia")
        CareerForm_OpenUia()
    else if (action = "read-clipboard")
        CareerForm_ReadClipboard(args.Has("out") ? args["out"] : "")
    else if (action = "send-prompt")
        CareerForm_SendPrompt(args.Has("file") ? args["file"] : "")
    else if (action = "fill-form")
        CareerForm_FillForm(args.Has("plan") ? args["plan"] : "")
    else
        CareerForm_Fail("Unknown action. Use capture-window, open-uia, read-clipboard, send-prompt, or fill-form.")
    ExitApp 0
} catch Error as err {
    CareerForm_Fail(err.Message)
}
