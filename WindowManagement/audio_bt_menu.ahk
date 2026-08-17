; =============================================================================
; WindowManagement module: audio_bt_menu.ahk
; Win+Alt+Shift+9 quick-access menu for Bluetooth audio and Windows sound devices.
; Root picker -> Bluetooth / Input / Output / Help / Ignored submenus.
; Loaded via #include into the WindowManagement.ahk process.
; Backend: infra\tools\AudioBt.ps1 (Core Audio IPolicyConfig + BluetoothAPIs).
; =============================================================================

global g_AudioBtGui := false
global g_AudioBtLv := false
global g_AudioBtHint := false
global g_AudioBtActive := false
global g_AudioBtHotkeysBound := false
global g_AudioBtHotkeyHandlers := []
global g_AudioBtRows := []
global g_AudioBtAllRows := []
global g_AudioBtBusy := false
global g_AudioBtMode := "root"  ; root | BT | In | Out | help | ignore
global g_AudioBtIgnoreItems := []
global g_AudioBtIgnoreIds := Map()
global g_AudioBtOpenFile := A_ScriptDir "\.cursor\wm_audio_bt_open"
global g_AudioBtCloseRequestFile := A_ScriptDir "\.cursor\wm_audio_bt_close_request"
global g_AudioBtCloseCheckTimer := ""

AudioBt_RootItems() {
    return [{ key: "1", kind: "BT", title: "Bluetooth", detail: "Paired audio devices — connect / disconnect" }, { key: "2",
        kind: "In", title: "Input", detail: "Recording devices — default / enable / isolate" }, { key: "3", kind: "Out",
            title: "Output", detail: "Playback devices — default / enable / isolate" }, { key: "4", kind: "help",
                title: "Help", detail: "Command meanings" }, { key: "5", kind: "ignore", title: "Ignored",
                    detail: "Hidden devices — restore" }
    ]
}

AudioBt_HelpRows() {
    return [{ key: "1 / B", command: "Bluetooth", meaning: "Bluetooth submenu: paired audio devices; connect/disconnect." }, { key: "2 / I",
        command: "Input", meaning: "Input submenu: recording devices; default / enable / isolate." }, { key: "3 / O",
            command: "Output", meaning: "Output submenu: playback devices; default / enable / isolate." }, { key: "4 / H",
                command: "Help", meaning: "This help list." }, { key: "5 / G", command: "Ignored",
                    meaning: "Open hidden devices and restore them to the lists." }, { key: "Enter", command: "Open",
                        meaning: "Open the highlighted root item." }, { key: "Esc",
                            command: "Close", meaning: "Close the menu." }, { key: "1-9 / 0", command: "Select",
                                meaning: "Select that row in a device submenu." }, { key: "Enter",
                                    command: "Default", meaning: "Set the selected device as the Windows default (playback or recording)." }, { key: "D",
                                        command: "Disable", meaning: "Disable/block it (Sound Settings Disable)." }, { key: "E",
                                            command: "Enable", meaning: "Enable/unblock it." }, { key: "C", command: "Connect",
                                                meaning: "Connect a paired Bluetooth audio device (A2DP/HFP)." }, { key: "X",
                                                    command: "Disconnect", meaning: "Disconnect that Bluetooth radio link (does not unpair)." }, { key: "I",
                                                        command: "Isolate", meaning: "Enable this device, set it default, disable other active devices of the same flow (output vs input). On a Bluetooth row, isolate that headset's matching endpoints. If it is disconnected, connect first. Isolated rows show ★ plus 🔊 (output) and/or 🎤 (input)." }, { key: "R",
                                                            command: "Refresh", meaning: "Reload the device list from Windows." }, { key: "N",
                                                                command: "Ignore", meaning: "Hide the selected device from the current list only (Input, Output, or Bluetooth)." }, { key: "Enter",
                                                                    command: "Restore", meaning: "In the Ignored list, put that row back on its original list only." }, { key: "Esc",
                                                                        command: "Back", meaning: "Back to the root menu." }
    ]
}

AudioBt_HintText() {
    global g_AudioBtMode
    if (g_AudioBtMode = "root")
        return "1/B Bluetooth   2/I Input   3/O Output   4/H Help   5/G Ignored   Enter open   Esc close"
    if (g_AudioBtMode = "help")
        return "Esc back"
    if (g_AudioBtMode = "ignore")
        return "Enter restore   Esc back"
    return "1-9/0 select   C connect   X disconnect   Enter default   D disable   E enable   I isolate   N ignore   R refresh   Esc back"
}

AudioBt_ModeTitle() {
    global g_AudioBtMode
    switch g_AudioBtMode {
        case "BT": return "Bluetooth devices"
        case "In": return "Input devices"
        case "Out": return "Output devices"
        case "help": return "Help"
        case "ignore": return "Ignored devices"
        default: return "Audio / Bluetooth"
    }
}

AudioBt_Ps1Path() {
    return A_ScriptDir "\infra\tools\AudioBt.ps1"
}

AudioBt_IgnoreIniPath() {
    return A_ScriptDir "\assets\data\audio_bt_ignore.ini"
}

AudioBt_IgnoreEntryKey(id, kind) {
    return StrLower(Trim(id)) "`t" Trim(kind)
}

AudioBt_IgnoreSanitize(text) {
    s := Trim(text)
    s := StrReplace(s, "`r", " ")
    s := StrReplace(s, "`n", " ")
    return s
}

AudioBt_IgnoreParseIni(raw) {
    fields := Map()
    count := 0
    inSection := false
    for line in StrSplit(raw, "`n", "`r") {
        line := Trim(line)
        if (line = "" || SubStr(line, 1, 1) = ";")
            continue
        if (SubStr(line, 1, 1) = "[") {
            inSection := (StrLower(line) = "[ignore]")
            continue
        }
        if (!inSection)
            continue
        eq := InStr(line, "=")
        if (eq < 2)
            continue
        key := StrLower(Trim(SubStr(line, 1, eq - 1)))
        val := Trim(SubStr(line, eq + 1))
        if (key = "count") {
            try count := Integer(val)
            catch {
                count := 0
            }
            continue
        }
        fields[key] := val
    }
    return { count: count, fields: fields }
}

AudioBt_IgnoreLoad() {
    global g_AudioBtIgnoreItems, g_AudioBtIgnoreIds
    g_AudioBtIgnoreItems := []
    g_AudioBtIgnoreIds := Map()
    path := AudioBt_IgnoreIniPath()
    if !FileExist(path)
        return
    raw := ""
    try raw := FileRead(path, "UTF-8")
    catch {
        return
    }
    parsed := AudioBt_IgnoreParseIni(raw)
    loop parsed.count {
        idx := A_Index
        id := parsed.fields.Has("id" idx) ? Trim(parsed.fields["id" idx]) : ""
        kind := parsed.fields.Has("kind" idx) ? Trim(parsed.fields["kind" idx]) : ""
        name := parsed.fields.Has("name" idx) ? Trim(parsed.fields["name" idx]) : ""
        if (id = "")
            continue
        key := AudioBt_IgnoreEntryKey(id, kind)
        if (key = "`t" || g_AudioBtIgnoreIds.Has(key))
            continue
        g_AudioBtIgnoreIds[key] := true
        g_AudioBtIgnoreItems.Push({ id: id, kind: kind, name: name })
    }
}

AudioBt_IgnoreSave() {
    global g_AudioBtIgnoreItems
    path := AudioBt_IgnoreIniPath()
    try DirCreate(A_ScriptDir "\assets\data")
    catch {
    }
    lines := ["[Ignore]", "Count=" g_AudioBtIgnoreItems.Length]
    for item in g_AudioBtIgnoreItems {
        idx := A_Index
        lines.Push("Id" idx "=" AudioBt_IgnoreSanitize(item.id))
        lines.Push("Kind" idx "=" AudioBt_IgnoreSanitize(item.kind))
        lines.Push("Name" idx "=" AudioBt_IgnoreSanitize(item.name))
    }
    text := ""
    for line in lines
        text .= line "`n"
    try FileDelete(path)
    catch {
    }
    FileAppend(text, path, "UTF-8")
}

AudioBt_IsIgnored(id, kind) {
    global g_AudioBtIgnoreIds
    if !IsObject(g_AudioBtIgnoreIds)
        AudioBt_IgnoreLoad()
    key := AudioBt_IgnoreEntryKey(id, kind)
    return Trim(id) != "" && g_AudioBtIgnoreIds.Has(key)
}

AudioBt_IgnoreAdd(row) {
    global g_AudioBtIgnoreItems, g_AudioBtIgnoreIds
    if !IsObject(row)
        return false
    if (Trim(row.id) = "")
        return false
    key := AudioBt_IgnoreEntryKey(row.id, row.kind)
    if (g_AudioBtIgnoreIds.Has(key))
        return false
    g_AudioBtIgnoreIds[key] := true
    g_AudioBtIgnoreItems.Push({ id: row.id, kind: row.kind, name: row.name })
    AudioBt_IgnoreSave()
    return true
}

AudioBt_IgnoreRemove(id, kind) {
    global g_AudioBtIgnoreItems, g_AudioBtIgnoreIds
    key := AudioBt_IgnoreEntryKey(id, kind)
    if (Trim(id) = "" || !g_AudioBtIgnoreIds.Has(key))
        return false
    g_AudioBtIgnoreIds.Delete(key)
    kept := []
    for item in g_AudioBtIgnoreItems {
        if (AudioBt_IgnoreEntryKey(item.id, item.kind) != key)
            kept.Push(item)
    }
    g_AudioBtIgnoreItems := kept
    AudioBt_IgnoreSave()
    return true
}

AudioBt_OnEscape(*) {
    global g_AudioBtActive, g_AudioBtMode
    if (!g_AudioBtActive)
        return false
    if (g_AudioBtMode != "root") {
        AudioBt_ShowRoot()
        return true
    }
    AudioBt_Cleanup()
    return true
}

AudioBt_CheckCloseRequest() {
    global g_AudioBtActive, g_AudioBtOpenFile, g_AudioBtCloseRequestFile
    if (!g_AudioBtActive) {
        try FileDelete(g_AudioBtOpenFile)
        catch {
        }
        try FileDelete(g_AudioBtCloseRequestFile)
        catch {
        }
        return
    }
    if (FileExist(g_AudioBtCloseRequestFile)) {
        try FileDelete(g_AudioBtCloseRequestFile)
        catch {
        }
        AudioBt_Cleanup()
    }
}

AudioBt_UnbindModalHotkeys() {
    global g_AudioBtGui, g_AudioBtHotkeyHandlers, g_AudioBtHotkeysBound
    hwnd := 0
    try {
        if (IsObject(g_AudioBtGui))
            hwnd := g_AudioBtGui.Hwnd
    } catch {
        hwnd := 0
    }
    if (hwnd) {
        try HotIfWinActive("ahk_id " hwnd)
        catch {
        }
    }
    for handler in g_AudioBtHotkeyHandlers {
        try Hotkey(handler.key, "Off")
        catch {
        }
    }
    if (hwnd) {
        try HotIf()
        catch {
        }
    }
    g_AudioBtHotkeyHandlers := []
    g_AudioBtHotkeysBound := false
}

AudioBt_BindOne(key, handler) {
    global g_AudioBtHotkeyHandlers
    try {
        Hotkey(key, handler, "On")
        g_AudioBtHotkeyHandlers.Push({ key: key, handler: handler })
    } catch {
    }
}

AudioBt_BindModalHotkeys() {
    global g_AudioBtGui, g_AudioBtHotkeysBound, g_AudioBtMode
    AudioBt_UnbindModalHotkeys()
    hwnd := 0
    try {
        if (IsObject(g_AudioBtGui))
            hwnd := g_AudioBtGui.Hwnd
    } catch {
        hwnd := 0
    }
    if (!hwnd)
        return
    try HotIfWinActive("ahk_id " hwnd)
    catch {
        return
    }

    if (g_AudioBtMode = "root") {
        AudioBt_BindOne("1", AudioBt_OpenSubmenu.Bind("BT"))
        AudioBt_BindOne("Numpad1", AudioBt_OpenSubmenu.Bind("BT"))
        AudioBt_BindOne("b", AudioBt_OpenSubmenu.Bind("BT"))
        AudioBt_BindOne("B", AudioBt_OpenSubmenu.Bind("BT"))
        AudioBt_BindOne("2", AudioBt_OpenSubmenu.Bind("In"))
        AudioBt_BindOne("Numpad2", AudioBt_OpenSubmenu.Bind("In"))
        AudioBt_BindOne("i", AudioBt_OpenSubmenu.Bind("In"))
        AudioBt_BindOne("I", AudioBt_OpenSubmenu.Bind("In"))
        AudioBt_BindOne("3", AudioBt_OpenSubmenu.Bind("Out"))
        AudioBt_BindOne("Numpad3", AudioBt_OpenSubmenu.Bind("Out"))
        AudioBt_BindOne("o", AudioBt_OpenSubmenu.Bind("Out"))
        AudioBt_BindOne("O", AudioBt_OpenSubmenu.Bind("Out"))
        AudioBt_BindOne("4", AudioBt_ShowHelp)
        AudioBt_BindOne("Numpad4", AudioBt_ShowHelp)
        AudioBt_BindOne("h", AudioBt_ShowHelp)
        AudioBt_BindOne("H", AudioBt_ShowHelp)
        AudioBt_BindOne("5", AudioBt_ShowIgnored)
        AudioBt_BindOne("Numpad5", AudioBt_ShowIgnored)
        AudioBt_BindOne("g", AudioBt_ShowIgnored)
        AudioBt_BindOne("G", AudioBt_ShowIgnored)
        AudioBt_BindOne("Enter", AudioBt_OnRootActivate)
        AudioBt_BindOne("NumpadEnter", AudioBt_OnRootActivate)
    } else if (g_AudioBtMode = "help") {
        AudioBt_BindOne("Backspace", AudioBt_OnEscape)
    } else if (g_AudioBtMode = "ignore") {
        loop 9 {
            dig := String(A_Index)
            AudioBt_BindOne(dig, AudioBt_SelectDigit.Bind(A_Index))
            AudioBt_BindOne("Numpad" . dig, AudioBt_SelectDigit.Bind(A_Index))
        }
        AudioBt_BindOne("0", AudioBt_SelectDigit.Bind(10))
        AudioBt_BindOne("Numpad0", AudioBt_SelectDigit.Bind(10))
        AudioBt_BindOne("Enter", AudioBt_OnUnignore)
        AudioBt_BindOne("NumpadEnter", AudioBt_OnUnignore)
        AudioBt_BindOne("Delete", AudioBt_OnUnignore)
        AudioBt_BindOne("$*u", AudioBt_OnUnignore)
        AudioBt_BindOne("Backspace", AudioBt_OnEscape)
    } else {
        loop 9 {
            dig := String(A_Index)
            AudioBt_BindOne(dig, AudioBt_SelectDigit.Bind(A_Index))
            AudioBt_BindOne("Numpad" . dig, AudioBt_SelectDigit.Bind(A_Index))
        }
        AudioBt_BindOne("0", AudioBt_SelectDigit.Bind(10))
        AudioBt_BindOne("Numpad0", AudioBt_SelectDigit.Bind(10))
        AudioBt_BindOne("Enter", AudioBt_OnDefault)
        AudioBt_BindOne("NumpadEnter", AudioBt_OnDefault)
        AudioBt_BindOne("$*d", AudioBt_OnDisable)
        AudioBt_BindOne("$*e", AudioBt_OnEnable)
        AudioBt_BindOne("$*i", AudioBt_OnIsolate)
        AudioBt_BindOne("$*r", AudioBt_OnRefresh)
        AudioBt_BindOne("$*c", AudioBt_OnConnect)
        AudioBt_BindOne("$*x", AudioBt_OnDisconnect)
        AudioBt_BindOne("$*n", AudioBt_OnIgnore)
        AudioBt_BindOne("Backspace", AudioBt_OnEscape)
    }
    AudioBt_BindOne("Escape", AudioBt_OnEscape)

    try HotIf()
    catch {
    }
    g_AudioBtHotkeysBound := true
}

AudioBt_DestroyGui() {
    global g_AudioBtGui, g_AudioBtLv, g_AudioBtHint
    AudioBt_UnbindModalHotkeys()
    if (IsObject(g_AudioBtGui)) {
        try g_AudioBtGui.Destroy()
        catch {
        }
        g_AudioBtGui := false
    }
    g_AudioBtLv := false
    g_AudioBtHint := false
}

AudioBt_Cleanup() {
    global g_AudioBtActive, g_AudioBtRows, g_AudioBtAllRows
    global g_AudioBtOpenFile, g_AudioBtCloseRequestFile, g_AudioBtCloseCheckTimer
    global g_OnEscapePressed, g_AudioBtBusy, g_AudioBtMode

    g_AudioBtActive := false
    g_AudioBtBusy := false
    g_AudioBtMode := "root"
    SetTimer(AudioBt_CheckCloseRequest, 0)
    g_AudioBtCloseCheckTimer := ""
    try FileDelete(g_AudioBtOpenFile)
    catch {
    }
    try FileDelete(g_AudioBtCloseRequestFile)
    catch {
    }

    AudioBt_BusyHide()
    AudioBt_DestroyGui()
    if (g_OnEscapePressed = AudioBt_OnEscape)
        g_OnEscapePressed := ""
    try Utils_EnsureGlobalEscapeHotkey()
    catch {
    }
    g_AudioBtRows := []
    g_AudioBtAllRows := []
}

AudioBt_WorkArea(&monitorLeft, &monitorTop, &monitorRight, &monitorBottom) {
    MonitorGetWorkArea(1, &monitorLeft, &monitorTop, &monitorRight, &monitorBottom)
    activeWin := 0
    try activeWin := WinGetID("A")
    catch {
        activeWin := 0
    }
    if (!activeWin)
        return
    rect := Buffer(16, 0)
    if !DllCall("GetWindowRect", "ptr", activeWin, "ptr", rect)
        return
    winLeft := NumGet(rect, 0, "int")
    winTop := NumGet(rect, 4, "int")
    winRight := NumGet(rect, 8, "int")
    winBottom := NumGet(rect, 12, "int")
    centerX := winLeft + (winRight - winLeft) // 2
    centerY := winTop + (winBottom - winTop) // 2
    loop MonitorGetCount() {
        idx := A_Index
        MonitorGetWorkArea(idx, &l, &t, &r, &b)
        if (centerX >= l && centerX <= r && centerY >= t && centerY <= b) {
            monitorLeft := l
            monitorTop := t
            monitorRight := r
            monitorBottom := b
            break
        }
    }
}

AudioBt_CreateGui(lvHeight := 360) {
    global g_AudioBtGui, g_AudioBtLv, g_AudioBtHint

    AudioBt_DestroyGui()
    AudioBt_WorkArea(&monitorLeft, &monitorTop, &monitorRight, &monitorBottom)
    monitorWidth := monitorRight - monitorLeft
    monitorHeight := monitorBottom - monitorTop

    g_AudioBtGui := Gui("+AlwaysOnTop +ToolWindow", AudioBt_ModeTitle())
    g_AudioBtGui.SetFont("s10", "Segoe UI")
    g_AudioBtHint := g_AudioBtGui.Add("Text", "w720", AudioBt_HintText())
    if (g_AudioBtMode = "help")
        lvCols := ["Key", "Command", "Meaning"]
    else if (g_AudioBtMode = "ignore")
        lvCols := ["#", "Name", "Kind"]
    else
        lvCols := ["#", "Name", "State"]
    g_AudioBtLv := g_AudioBtGui.Add("ListView", "w720 h" lvHeight " -Multi", lvCols)
    g_AudioBtLv.SetFont("s14", "Segoe UI")
    g_AudioBtLv.OnEvent("DoubleClick", AudioBt_OnListActivate)
    g_AudioBtGui.Add("Button", "w100 Section", (g_AudioBtMode = "root") ? "Close" : "Back").OnEvent("Click",
        AudioBt_OnEscape)
    g_AudioBtGui.OnEvent("Close", (*) => AudioBt_Cleanup())
    g_AudioBtGui.OnEvent("Escape", AudioBt_OnEscape)

    guiW := 750
    guiH := lvHeight + 100
    guiX := monitorLeft + (monitorWidth - guiW) // 2
    guiY := monitorTop + (monitorHeight - guiH) // 2
    if (guiX < monitorLeft + 20)
        guiX := monitorLeft + 20
    if (guiY < monitorTop + 20)
        guiY := monitorTop + 20

    g_AudioBtGui.Show("x" . guiX . " y" . guiY)
    try g_AudioBtLv.Focus()
    catch {
    }
    AudioBt_BindModalHotkeys()
}

AudioBt_Run(action, id := "", note := "") {
    ps1 := AudioBt_Ps1Path()
    if !FileExist(ps1)
        return { ok: false, text: "Missing AudioBt.ps1" }
    outFile := A_Temp "\audiobt-out-" A_TickCount ".txt"
    cmd := 'powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "' ps1 '" -Action ' action
    if (id != "")
        cmd .= ' -Id "' id '"'
    if (action = "connect" || action = "isolate" || action = "disconnect") {
        logDir := A_ScriptDir "\docs\debug"
        try DirCreate(logDir)
        catch {
        }
        cmd .= ' -LogDir "' logDir '"'
        if (note != "") {
            noteSafe := StrReplace(StrReplace(note, '"', "'"), "`n", " ")
            cmd .= ' -Note "' noteSafe '"'
        }
    }
    cmd .= ' -OutFile "' outFile '"'
    exitCode := 1
    try exitCode := RunWait(cmd, , "Hide")
    catch {
        exitCode := 1
    }
    text := ""
    try {
        if FileExist(outFile)
            text := Trim(FileRead(outFile, "UTF-8"), "`r`n")
    } catch {
    }
    try FileDelete(outFile)
    catch {
    }
    ok := (exitCode = 0 && text != "" && SubStr(text, 1, 4) != "ERR`t")
    return { ok: ok, text: text, exitCode: exitCode }
}

AudioBt_ParseList(text) {
    rows := []
    if (text = "" || SubStr(text, 1, 4) = "ERR`t")
        return rows
    for line in StrSplit(text, "`n", "`r") {
        line := Trim(line)
        if (line = "")
            continue
        parts := StrSplit(line, "`t")
        if (parts.Length < 6)
            continue
        if (A_Index = 1 && StrLower(parts[1]) = "id")
            continue
        rows.Push({
            id: parts[1],
            kind: parts[2],
            name: parts[3],
            state: parts[4],
            isDefault: parts[5] = "1",
            canConnect: parts[6] = "1",
            iso: (parts.Length >= 7) ? parts[7] : ""
        })
    }
    return rows
}

AudioBt_FilterRows(allRows, kind) {
    rows := []
    for row in allRows {
        if (row.kind = kind && !AudioBt_IsIgnored(row.id, row.kind))
            rows.Push(row)
    }
    return rows
}

AudioBt_DigitLabel(index) {
    if (index >= 1 && index <= 9)
        return String(index)
    if (index = 10)
        return "0"
    return ""
}

AudioBt_IsolatePrefix(row) {
    iso := ""
    if (IsObject(row) && row.HasProp("iso"))
        iso := row.iso
    if (iso = "" && IsObject(row) && InStr(row.state, "Isolated")) {
        if (row.kind = "Out")
            iso := "Out"
        else if (row.kind = "In")
            iso := "In"
        else
            iso := "InOut"
    }
    if (iso = "")
        return ""
    icons := "★"
    if (iso = "Out" || iso = "InOut")
        icons .= " 🔊"
    if (iso = "In" || iso = "InOut")
        icons .= " 🎤"
    return icons . " "
}

AudioBt_PopulateRootLv() {
    global g_AudioBtLv
    if (!IsObject(g_AudioBtLv))
        return
    g_AudioBtLv.Delete()
    for item in AudioBt_RootItems() {
        g_AudioBtLv.Add("", item.key, item.title, item.detail)
    }
    try g_AudioBtLv.ModifyCol(1, 40)
    try g_AudioBtLv.ModifyCol(2, 140)
    try g_AudioBtLv.ModifyCol(3, 520)
    ListView_SelectRowFocused(g_AudioBtLv, 1)
}

AudioBt_PopulateHelpLv() {
    global g_AudioBtLv
    if (!IsObject(g_AudioBtLv))
        return
    g_AudioBtLv.Delete()
    for row in AudioBt_HelpRows() {
        g_AudioBtLv.Add("", row.key, row.command, row.meaning)
    }
    try g_AudioBtLv.ModifyCol(1, 90)
    try g_AudioBtLv.ModifyCol(2, 110)
    try g_AudioBtLv.ModifyCol(3, 500)
    ListView_SelectRowFocused(g_AudioBtLv, 1)
}

AudioBt_PopulateDeviceLv(keepName := "") {
    global g_AudioBtLv, g_AudioBtRows
    if (!IsObject(g_AudioBtLv))
        return
    g_AudioBtLv.Delete()
    selectRow := 1
    for row in g_AudioBtRows {
        idx := A_Index
        prefix := AudioBt_IsolatePrefix(row)
        g_AudioBtLv.Add("", AudioBt_DigitLabel(idx), prefix . row.name, row.state)
        if (keepName != "" && row.name = keepName)
            selectRow := idx
    }
    try g_AudioBtLv.ModifyCol(1, 40)
    try g_AudioBtLv.ModifyCol(2, 480)
    try g_AudioBtLv.ModifyCol(3, 180)
    if (g_AudioBtRows.Length > 0)
        ListView_SelectRowFocused(g_AudioBtLv, selectRow)
}

AudioBt_PopulateIgnoredLv() {
    global g_AudioBtLv, g_AudioBtRows, g_AudioBtIgnoreItems
    g_AudioBtRows := g_AudioBtIgnoreItems
    if (!IsObject(g_AudioBtLv))
        return
    g_AudioBtLv.Delete()
    for row in g_AudioBtRows {
        idx := A_Index
        kind := row.kind
        if (kind = "BT")
            kind := "Bluetooth"
        else if (kind = "In")
            kind := "Input"
        else if (kind = "Out")
            kind := "Output"
        g_AudioBtLv.Add("", AudioBt_DigitLabel(idx), row.name, kind)
    }
    try g_AudioBtLv.ModifyCol(1, 40)
    try g_AudioBtLv.ModifyCol(2, 480)
    try g_AudioBtLv.ModifyCol(3, 180)
    if (g_AudioBtRows.Length > 0)
        ListView_SelectRowFocused(g_AudioBtLv, 1)
}

AudioBt_SelectedIndex() {
    global g_AudioBtLv, g_AudioBtRows, g_AudioBtMode
    if (!IsObject(g_AudioBtLv))
        return 0
    rowNum := g_AudioBtLv.GetNext(0, "Focused")
    if (rowNum < 1)
        rowNum := g_AudioBtLv.GetNext(0)
    if (g_AudioBtMode = "root") {
        if (rowNum < 1 || rowNum > AudioBt_RootItems().Length)
            return 0
        return rowNum
    }
    if (rowNum < 1 || rowNum > g_AudioBtRows.Length)
        return 0
    return rowNum
}

AudioBt_SelectedRow() {
    global g_AudioBtRows, g_AudioBtMode
    if (g_AudioBtMode = "root")
        return ""
    idx := AudioBt_SelectedIndex()
    if (idx < 1)
        return ""
    return g_AudioBtRows[idx]
}

AudioBt_SetHint(text) {
    global g_AudioBtHint
    if (IsObject(g_AudioBtHint)) {
        try g_AudioBtHint.Value := text
        catch {
        }
    }
}

AudioBt_BusyShow(state) {
    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Show(state, BANNER_ACCENT_INTERMEDIATE, { passive: false, centerOnHwnd: 0 })
}

AudioBt_BusyHide() {
    try StandardLoadingBar_Hide(0)
    catch {
    }
}

AudioBt_RefocusLv() {
    global g_AudioBtLv
    try g_AudioBtLv.Focus()
    catch {
    }
}

AudioBt_ActionBusyText(action, row) {
    verb := ""
    switch action {
        case "disconnect": verb := "Disconnecting"
        case "connect": verb := "Connecting"
        case "default": verb := "Setting default"
        case "enable": verb := "Enabling"
        case "disable": verb := "Disabling"
        case "isolate":
            verb := InStr(row.state, "Disconnected") ? "Connecting and isolating" : "Isolating"
        default: verb := "Working"
    }
    return "⏳ " . verb . " " . row.name . "…"
}

AudioBt_StripResult(text) {
    msg := text
    if (SubStr(msg, 1, 3) = "OK`t")
        msg := SubStr(msg, 4)
    else if (SubStr(msg, 1, 4) = "ERR`t")
        msg := SubStr(msg, 5)
    return Trim(msg)
}

AudioBt_FetchAll(showError := true) {
    global g_AudioBtAllRows, g_AudioBtBusy
    result := AudioBt_Run("list")
    if (!result.ok) {
        g_AudioBtBusy := false
        AudioBt_SetHint(AudioBt_HintText())
        if (showError) {
            msg := result.text
            if (SubStr(msg, 1, 4) = "ERR`t")
                msg := SubStr(msg, 5)
            if (msg = "")
                msg := "Failed to list audio devices"
            if (SubStr(msg, 1, 1) != "❌")
                msg := "❌ " . msg
            ShowNotification_WM(msg)
        }
        return false
    }
    g_AudioBtAllRows := AudioBt_ParseList(result.text)
    return true
}

AudioBt_Refresh(keepName := "", showError := true) {
    global g_AudioBtRows, g_AudioBtAllRows, g_AudioBtBusy, g_AudioBtMode
    if (g_AudioBtMode = "root" || g_AudioBtMode = "help" || g_AudioBtMode = "ignore") {
        g_AudioBtBusy := false
        return true
    }
    if (!AudioBt_FetchAll(showError))
        return false
    g_AudioBtRows := AudioBt_FilterRows(g_AudioBtAllRows, g_AudioBtMode)
    AudioBt_PopulateDeviceLv(keepName)
    AudioBt_SetHint(AudioBt_HintText())
    g_AudioBtBusy := false
    return true
}

AudioBt_SelectDigit(index, *) {
    global g_AudioBtLv, g_AudioBtRows, g_AudioBtBusy, g_AudioBtMode
    if (g_AudioBtBusy || !IsObject(g_AudioBtLv) || g_AudioBtMode = "root" || g_AudioBtMode = "help")
        return
    if (index < 1 || index > g_AudioBtRows.Length)
        return
    ListView_SelectRowFocused(g_AudioBtLv, index)
}

AudioBt_OpenSubmenu(kind, *) {
    global g_AudioBtMode, g_AudioBtBusy, g_AudioBtActive, g_AudioBtRows, g_AudioBtAllRows
    if (!g_AudioBtActive)
        return
    g_AudioBtMode := kind
    AudioBt_CreateGui(360)
    AudioBt_SetHint(AudioBt_HintText())
    g_AudioBtBusy := true
    needFetch := (g_AudioBtAllRows.Length = 0)
    fetchFailed := false
    if (needFetch)
        AudioBt_BusyShow("⏳ Loading " . StrLower(AudioBt_ModeTitle()) . "…")
    try {
        if (needFetch && !AudioBt_FetchAll(false))
            fetchFailed := true
        else {
            g_AudioBtRows := AudioBt_FilterRows(g_AudioBtAllRows, kind)
            AudioBt_PopulateDeviceLv()
            AudioBt_SetHint(AudioBt_HintText())
        }
    } finally {
        g_AudioBtBusy := false
        if (needFetch)
            AudioBt_BusyHide()
        AudioBt_RefocusLv()
    }
    if (fetchFailed)
        ShowNotification_WM("❌ Failed to list audio devices")
}

AudioBt_OnRootActivate(*) {
    global g_AudioBtMode
    if (g_AudioBtMode != "root")
        return
    idx := AudioBt_SelectedIndex()
    items := AudioBt_RootItems()
    if (idx < 1 || idx > items.Length)
        return
    kind := items[idx].kind
    if (kind = "help")
        AudioBt_ShowHelp()
    else if (kind = "ignore")
        AudioBt_ShowIgnored()
    else
        AudioBt_OpenSubmenu(kind)
}

AudioBt_ShowHelp(*) {
    global g_AudioBtMode, g_AudioBtBusy, g_AudioBtActive, g_AudioBtRows
    if (!g_AudioBtActive)
        return
    g_AudioBtMode := "help"
    g_AudioBtBusy := false
    g_AudioBtRows := []
    AudioBt_CreateGui(360)
    AudioBt_PopulateHelpLv()
    AudioBt_SetHint(AudioBt_HintText())
    AudioBt_RefocusLv()
}

AudioBt_ShowIgnored(*) {
    global g_AudioBtMode, g_AudioBtBusy, g_AudioBtActive, g_AudioBtRows
    if (!g_AudioBtActive)
        return
    AudioBt_IgnoreLoad()
    g_AudioBtMode := "ignore"
    g_AudioBtBusy := false
    AudioBt_CreateGui(360)
    AudioBt_PopulateIgnoredLv()
    AudioBt_SetHint(AudioBt_HintText())
    AudioBt_RefocusLv()
    if (g_AudioBtRows.Length = 0)
        ShowCenteredOverlay_Utils("Ignore list is empty", 1400, BANNER_ACCENT_INTERMEDIATE)
}

AudioBt_OnIgnore(*) {
    global g_AudioBtBusy, g_AudioBtActive, g_AudioBtMode, g_AudioBtAllRows
    if (!g_AudioBtActive || g_AudioBtBusy || g_AudioBtMode = "root" || g_AudioBtMode = "help" || g_AudioBtMode =
        "ignore")
        return
    row := AudioBt_SelectedRow()
    if (!IsObject(row)) {
        ShowNotification_WM("Select a device first")
        return
    }
    if (AudioBt_IsIgnored(row.id, row.kind)) {
        ShowNotification_WM("Already ignored")
        return
    }
    if !AudioBt_IgnoreAdd(row) {
        ShowNotification_WM("Could not ignore that device")
        return
    }
    g_AudioBtRows := AudioBt_FilterRows(g_AudioBtAllRows, g_AudioBtMode)
    AudioBt_PopulateDeviceLv()
    AudioBt_RefocusLv()
    ShowCenteredOverlay_Utils("Ignored: " . row.name, 1400, BANNER_ACCENT_SUCCESS)
}

AudioBt_OnUnignore(*) {
    global g_AudioBtBusy, g_AudioBtActive, g_AudioBtMode
    if (!g_AudioBtActive || g_AudioBtBusy || g_AudioBtMode != "ignore")
        return
    row := AudioBt_SelectedRow()
    if (!IsObject(row)) {
        ShowNotification_WM("Select a device first")
        return
    }
    name := row.name
    if !AudioBt_IgnoreRemove(row.id, row.kind) {
        ShowNotification_WM("Could not restore that device")
        return
    }
    AudioBt_PopulateIgnoredLv()
    AudioBt_RefocusLv()
    ShowCenteredOverlay_Utils("Restored: " . name, 1400, BANNER_ACCENT_SUCCESS)
}

AudioBt_DoAction(action, requireBt := false) {
    global g_AudioBtBusy, g_AudioBtActive, g_AudioBtMode
    if (!g_AudioBtActive || g_AudioBtBusy || g_AudioBtMode = "root" || g_AudioBtMode = "help" || g_AudioBtMode =
        "ignore")
        return
    row := AudioBt_SelectedRow()
    if (!IsObject(row)) {
        ShowNotification_WM("Select a device first")
        return
    }
    if (requireBt && row.canConnect != true && row.kind != "BT") {
        ShowNotification_WM("Not a Bluetooth audio device")
        return
    }
    g_AudioBtBusy := true
    AudioBt_BusyShow(AudioBt_ActionBusyText(action, row))
    result := ""
    try {
        result := AudioBt_Run(action, row.id, g_AudioBtMode . "|" . row.name . "|" . row.state)
        StandardLoadingBar_Update("⏳ Refreshing devices…", BANNER_ACCENT_INTERMEDIATE)
        AudioBt_Refresh(row.name, false)
    } finally {
        g_AudioBtBusy := false
        AudioBt_BusyHide()
        AudioBt_RefocusLv()
    }
    msg := AudioBt_StripResult(IsObject(result) ? result.text : "")
    if (IsObject(result) && result.ok)
        ShowCenteredOverlay_Utils("✅ " . (msg = "" ? "Done" : msg), 1400, BANNER_ACCENT_SUCCESS)
    else
        ShowNotification_WM("❌ " . (msg = "" ? "Action failed" : msg))
}

AudioBt_OnDefault(*) {
    AudioBt_DoAction("default")
}

AudioBt_OnDisable(*) {
    AudioBt_DoAction("disable")
}

AudioBt_OnEnable(*) {
    AudioBt_DoAction("enable")
}

AudioBt_OnConnect(*) {
    AudioBt_DoAction("connect", true)
}

AudioBt_OnDisconnect(*) {
    AudioBt_DoAction("disconnect", true)
}

AudioBt_OnIsolate(*) {
    AudioBt_DoAction("isolate")
}

AudioBt_OnRefresh(*) {
    global g_AudioBtBusy, g_AudioBtActive, g_AudioBtMode, g_AudioBtAllRows
    if (!g_AudioBtActive || g_AudioBtBusy || g_AudioBtMode = "root" || g_AudioBtMode = "help" || g_AudioBtMode =
        "ignore")
        return
    row := AudioBt_SelectedRow()
    keep := IsObject(row) ? row.name : ""
    g_AudioBtBusy := true
    g_AudioBtAllRows := []
    AudioBt_BusyShow("⏳ Refreshing devices…")
    try {
        AudioBt_Refresh(keep)
    } finally {
        g_AudioBtBusy := false
        AudioBt_BusyHide()
        AudioBt_RefocusLv()
    }
}

AudioBt_OnListActivate(*) {
    global g_AudioBtMode
    if (g_AudioBtMode = "root")
        AudioBt_OnRootActivate()
    else if (g_AudioBtMode = "ignore")
        AudioBt_OnUnignore()
    else if (g_AudioBtMode != "help")
        AudioBt_OnDefault()
}

AudioBt_ShowRoot() {
    global g_AudioBtMode, g_AudioBtBusy, g_AudioBtRows
    g_AudioBtMode := "root"
    g_AudioBtBusy := false
    g_AudioBtRows := []
    AudioBt_CreateGui(220)
    AudioBt_PopulateRootLv()
    AudioBt_SetHint(AudioBt_HintText())
    try g_AudioBtLv.Focus()
    catch {
    }
}

AudioBt_Show() {
    global g_AudioBtActive, g_AudioBtGui, g_AudioBtAllRows
    global g_OnEscapePressed, g_AudioBtOpenFile, g_AudioBtCloseCheckTimer, g_AudioBtBusy

    if (g_AudioBtActive && IsObject(g_AudioBtGui)) {
        AudioBt_Cleanup()
        Sleep 50
    }

    g_AudioBtAllRows := []
    g_AudioBtBusy := false
    g_AudioBtActive := true
    g_OnEscapePressed := AudioBt_OnEscape
    try {
        DirCreate(A_ScriptDir "\.cursor")
        FileAppend "", g_AudioBtOpenFile
    } catch {
    }
    g_AudioBtCloseCheckTimer := SetTimer(AudioBt_CheckCloseRequest, 120)
    AudioBt_IgnoreLoad()
    AudioBt_ShowRoot()
}

; Win+Alt+Shift+9 tap-dance (400 ms = AI_QD_DOUBLE_TAP_MS / ZMK tap-dance):
;   1× = AI Companion Quick Download (Utils\ai_quick_download.ahk)
;   2× = Audio / Bluetooth quick selector (toggle)
global g_AudioBt_DoubleTapArmed := false
global g_AudioBt_LastPressTick := 0
global g_AudioBt_DoubleTapTimer := 0

class AudioBt_DoubleTapTimerObj {
    static OnSingleTapTimeout() {
        global g_AudioBt_DoubleTapArmed, g_AudioBt_DoubleTapTimer
        if (!g_AudioBt_DoubleTapArmed)
            return
        g_AudioBt_DoubleTapArmed := false
        g_AudioBt_DoubleTapTimer := 0
        AiQuickDownload_Run()
    }
}

#!+9:: {
    global g_AudioBt_DoubleTapArmed, g_AudioBt_LastPressTick, g_AudioBt_DoubleTapTimer
    global g_AudioBtActive, g_AudioBtGui

    thresholdMs := 400
    try thresholdMs := AI_QD_DOUBLE_TAP_MS
    catch {
        thresholdMs := 400
    }

    now := A_TickCount
    elapsed := (g_AudioBt_LastPressTick > 0) ? (now - g_AudioBt_LastPressTick) : 9999

    if (g_AudioBt_DoubleTapArmed && elapsed >= 0 && elapsed < thresholdMs) {
        g_AudioBt_DoubleTapArmed := false
        g_AudioBt_LastPressTick := 0
        if (g_AudioBt_DoubleTapTimer) {
            SetTimer(g_AudioBt_DoubleTapTimer, 0)
            g_AudioBt_DoubleTapTimer := 0
        }
        if (g_AudioBtActive && IsObject(g_AudioBtGui)) {
            AudioBt_Cleanup()
        } else {
            AudioBt_Show()
        }
        return
    }

    g_AudioBt_LastPressTick := now
    g_AudioBt_DoubleTapArmed := true
    g_AudioBt_DoubleTapTimer := ObjBindMethod(AudioBt_DoubleTapTimerObj, "OnSingleTapTimeout")
    SetTimer(g_AudioBt_DoubleTapTimer, -thresholdMs)
}
