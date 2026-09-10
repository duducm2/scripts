; =============================================================================
; WindowManagement module: audio_bt_settings_ui.ahk
; Last-resort Bluetooth connect via Windows Settings (ms-settings:bluetooth).
; Loaded via #include from audio_bt_menu.ahk.
; =============================================================================

AudioBt_SettingsConnectNames() {
    return ["Conectar", "Connect"]
}

AudioBt_SettingsDisconnectNames() {
    return ["Desconectar", "Disconnect"]
}

AudioBt_SettingsExpandNames() {
    return ["Mais opções", "Mais configurações", "More settings", "More options"]
}

AudioBt_SettingsConnectedTextNames() {
    return ["Conectado", "Connected"]
}

AudioBt_SettingsNavNames() {
    return ["Bluetooth e dispositivos", "Bluetooth & devices"]
}

AudioBt_SettingsShowMoreNames() {
    ; Dump may use NBSP between "mais" and "dispositivos" — match via substring helper.
    return ["Exibir mais", "Show more devices", "Show more"]
}

AudioBt_SettingsWinSpec() {
    if WinExist("Configurações ahk_class ApplicationFrameWindow")
        return "Configurações ahk_class ApplicationFrameWindow"
    if WinExist("Settings ahk_class ApplicationFrameWindow")
        return "Settings ahk_class ApplicationFrameWindow"
    if WinExist("ahk_exe SystemSettings.exe")
        return "ahk_exe SystemSettings.exe"
    return ""
}

AudioBt_SettingsRoot(hwnd) {
    if !hwnd
        return 0
    try return UIA.ElementFromHandle(hwnd)
    catch
        return 0
}

AudioBt_SettingsFindNamed(scope, names, typeName := "Button") {
    if !scope
        return 0
    for n in names {
        el := ClipAngel_UiaFindFirst(scope, { Type: typeName, Name: n })
        if el
            return el
    }
    return 0
}

; Substring Name match (tolerates NBSP / extra wording in Settings labels).
AudioBt_SettingsFindNamedSubstring(scope, needles, typeName := "Button") {
    if !scope
        return 0
    try {
        els := scope.FindAll({ Type: typeName })
    } catch {
        return 0
    }
    loop els.Length {
        el := els[A_Index]
        try nm := el.Name
        catch
            continue
        if (nm = "")
            continue
        nmNorm := StrReplace(nm, Chr(0xA0), " ")
        for needle in needles {
            if InStr(nmNorm, needle) || InStr(nm, needle)
                return el
        }
    }
    return 0
}

AudioBt_SettingsScrollIntoView(el) {
    if !el
        return false
    try {
        if el.GetPropertyValue(UIA.Property.IsScrollItemPatternAvailable) {
            el.ScrollItemPattern.ScrollIntoView()
            return true
        }
    } catch {
    }
    try {
        el.ScrollIntoView()
        return true
    } catch {
    }
    return false
}

AudioBt_SettingsClickEl(el) {
    if !el
        return false
    AudioBt_SettingsScrollIntoView(el)
    if ClickSeq_Invoke(el)
        return true
    try {
        el.Click()
        return true
    } catch {
    }
    return false
}

AudioBt_SettingsClickNamed(scope, names) {
    el := AudioBt_SettingsFindNamed(scope, names)
    if !el
        return false
    return AudioBt_SettingsClickEl(el)
}

AudioBt_SettingsRowMatches(elName, deviceName) {
    n := Trim(deviceName)
    if (n = "" || elName = "")
        return false
    if (Trim(elName) = n)
        return true
    return InStr(elName, n ", ") = 1 || InStr(elName, n ",") = 1
}

AudioBt_SettingsRowConnected(elName) {
    if (elName = "")
        return false
    if InStr(elName, "Estado Conectado") || InStr(elName, "Status Connected")
        return true
    t := Trim(elName)
    return (t = "Conectado" || t = "Connected")
}

AudioBt_SettingsRowIsConnected(row) {
    if !row
        return false
    try nm := row.Name
    catch
        nm := ""
    if AudioBt_SettingsRowConnected(nm)
        return true
    scope := AudioBt_SettingsClickScope(row)
    if !scope
        scope := row
    if AudioBt_SettingsFindNamed(scope, AudioBt_SettingsConnectedTextNames(), "Text")
        return true
    return false
}

; Prefer ListItem / DevicesHeroControlButton so Conectar is in scope (sibling of name Group).
AudioBt_SettingsClickScope(row) {
    if !row
        return 0
    try {
        heroBtn := ClipAngel_UiaFindFirst(row, { AutomationId: "DevicesHeroControlButton" })
        if heroBtn
            return heroBtn
    }
    return row
}

AudioBt_SettingsFindRow(root, deviceName) {
    if !root
        return 0
    list := ClipAngel_UiaFindFirst(root, { AutomationId: "SystemSettings_Devices_HeroControlDeviceList_ListView" })
    if list {
        try {
            items := list.FindAll({ Type: "ListItem" })
            loop items.Length {
                el := items[A_Index]
                try nm := el.Name
                catch
                    continue
                if AudioBt_SettingsRowMatches(nm, deviceName)
                    return el
            }
        }
    }
    try {
        els := root.FindAll({ Type: "Group" })
        loop els.Length {
            el := els[A_Index]
            try nm := el.Name
            catch
                continue
            if AudioBt_SettingsRowMatches(nm, deviceName)
                return el
        }
    }
    return 0
}

AudioBt_SettingsWaitPage(hwnd, timeoutMs := 15000) {
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        root := AudioBt_SettingsRoot(hwnd)
        if root {
            toggle := ClipAngel_UiaFindFirst(root, { AutomationId: "SystemSettings_Device_BluetoothRadioToggle_ToggleSwitch" })
            if !toggle
                toggle := ClipAngel_UiaFindFirst(root, { AutomationId: "SystemSettings_Devices_AudioDeviceList_SettingsListItemsRepeater" })
            if !toggle
                toggle := ClipAngel_UiaFindFirst(root, { AutomationId: "SystemSettings_Devices_HeroControlDeviceList_ListView" })
            if toggle
                return root
        }
        Sleep(200)
    }
    return 0
}

AudioBt_SettingsMaximize(hwnd) {
    if !hwnd
        return false
    try WM_MaximizeHwnd(hwnd)
    catch {
        try WinMaximize("ahk_id " hwnd)
        catch {
            try PostMessage(0x0112, 0xF030, , , "ahk_id " hwnd)
            catch {
            }
        }
    }
    deadline := A_TickCount + 2000
    while (A_TickCount < deadline) {
        try {
            if (WinGetMinMax("ahk_id " hwnd) = 1) {
                Sleep(350)
                return true
            }
        } catch {
        }
        Sleep(100)
    }
    Sleep(350)
    return false
}

AudioBt_SettingsClickNavBt(root) {
    if !root
        return false
    for n in AudioBt_SettingsNavNames() {
        el := ClipAngel_UiaFindFirst(root, { Type: "ListItem", Name: n })
        if !el
            el := ClipAngel_UiaFindFirst(root, { Type: "Button", Name: n })
        if el {
            if AudioBt_SettingsClickEl(el)
                return true
        }
    }
    return false
}

; After maximize: ensure Bluetooth & devices page content is present.
AudioBt_SettingsEnsureBtPage(hwnd, timeoutMs := 15000) {
    root := AudioBt_SettingsWaitPage(hwnd, timeoutMs)
    if root
        return root
    root := AudioBt_SettingsRoot(hwnd)
    if root && AudioBt_SettingsClickNavBt(root) {
        Sleep(500)
        root := AudioBt_SettingsWaitPage(hwnd, timeoutMs)
        if root
            return root
    }
    return 0
}

AudioBt_SettingsClickShowMore(root) {
    if !root
        return false
    btn := AudioBt_SettingsFindNamedSubstring(root, AudioBt_SettingsShowMoreNames(), "Button")
    if !btn
        return false
    if !AudioBt_SettingsClickEl(btn)
        return false
    Sleep(600)
    return true
}

; Find device row; if missing, click Exibir mais / Show more and retry once.
AudioBt_SettingsFindRowOrReveal(hwnd, deviceName) {
    root := AudioBt_SettingsRoot(hwnd)
    row := AudioBt_SettingsFindRow(root, deviceName)
    if row
        return row
    if root && AudioBt_SettingsClickShowMore(root) {
        root := AudioBt_SettingsRoot(hwnd)
        row := AudioBt_SettingsFindRow(root, deviceName)
        if row
            return row
    }
    return 0
}

AudioBt_SettingsEnsureOpen(&opened) {
    opened := false
    alreadySpec := AudioBt_SettingsWinSpec()
    try Run("ms-settings:bluetooth")
    catch {
        return 0
    }
    deadline := A_TickCount + 8000
    hwnd := 0
    while (A_TickCount < deadline) {
        spec := AudioBt_SettingsWinSpec()
        if (spec != "") {
            hwnd := WinExist(spec)
            if hwnd
                break
        }
        Sleep(150)
    }
    if !hwnd
        return 0
    opened := (alreadySpec = "")
    try WinActivate("ahk_id " hwnd)
    catch {
    }
    try WinWaitActive("ahk_id " hwnd, , 2)
    catch {
    }
    AudioBt_SettingsMaximize(hwnd)
    return hwnd
}

AudioBt_SettingsExpand(row) {
    if !row
        return false
    scope := AudioBt_SettingsClickScope(row)
    btn := AudioBt_SettingsFindNamed(scope, AudioBt_SettingsExpandNames())
    if !btn
        return false
    if !ClickSeq_Invoke(btn)
        return false
    Sleep(500)
    return true
}

AudioBt_SettingsDisconnectRow(row) {
    scope := AudioBt_SettingsClickScope(row)
    if AudioBt_SettingsClickNamed(scope, AudioBt_SettingsDisconnectNames())
        return true
    if !AudioBt_SettingsExpand(row)
        return false
    scope := AudioBt_SettingsClickScope(row)
    return AudioBt_SettingsClickNamed(scope, AudioBt_SettingsDisconnectNames())
}

AudioBt_SettingsSkipOther(groupName) {
    u := StrLower(groupName)
    for skip in ["teclado", "keyboard", "mouse", "vídeo", "video", "webcam"] {
        if InStr(u, skip)
            return true
    }
    return false
}

AudioBt_SettingsAudioScope(root) {
    if !root
        return 0
    list := ClipAngel_UiaFindFirst(root, { AutomationId: "SystemSettings_Devices_AudioDeviceList_SettingsListItemsRepeater" })
    if list
        return list
    try {
        els := root.FindAll({ Type: "Group" })
        loop els.Length {
            el := els[A_Index]
            try nm := el.Name
            catch
                continue
            if (nm = "Áudio" || nm = "Audio")
                return el
        }
    }
    return root
}

AudioBt_SettingsDropOtherAudio(root, keepName) {
    scope := AudioBt_SettingsAudioScope(root)
    if !scope
        return
    try {
        els := scope.FindAll({ Type: "Group" })
    } catch {
        return
    }
    loop els.Length {
        el := els[A_Index]
        try nm := el.Name
        catch
            continue
        if AudioBt_SettingsRowMatches(nm, keepName)
            continue
        if AudioBt_SettingsSkipOther(nm)
            continue
        if !AudioBt_SettingsRowIsConnected(el)
            continue
        AudioBt_SettingsDisconnectRow(el)
        Sleep(400)
    }
}

AudioBt_SettingsWaitRowConnected(hwnd, deviceName, timeoutMs := 8000) {
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        root := AudioBt_SettingsRoot(hwnd)
        row := AudioBt_SettingsFindRow(root, deviceName)
        if AudioBt_SettingsRowIsConnected(row)
            return true
        Sleep(250)
    }
    return false
}

; Click Conectar under the device row (scroll + invoke). Returns false if button missing.
AudioBt_SettingsClickConnect(row) {
    if !row
        return false
    scope := AudioBt_SettingsClickScope(row)
    btn := AudioBt_SettingsFindNamed(scope, AudioBt_SettingsConnectNames())
    if !btn
        return false
    return AudioBt_SettingsClickEl(btn)
}

; One connect attempt: find Conectar (or bounce disconnect→connect), click, wait for Conectado.
AudioBt_SettingsTryConnectOnce(hwnd, deviceName) {
    row := AudioBt_SettingsFindRowOrReveal(hwnd, deviceName)
    if !row
        return { ok: false, err: "Device not found in Settings: " deviceName, clicked: false }
    scope := AudioBt_SettingsClickScope(row)
    if AudioBt_SettingsFindNamed(scope, AudioBt_SettingsConnectNames()) {
        if !AudioBt_SettingsClickConnect(row)
            return { ok: false, err: "Could not click Connect for " deviceName, clicked: false }
    } else if AudioBt_SettingsRowIsConnected(row) {
        if !AudioBt_SettingsDisconnectRow(row)
            return { ok: false, err: "Could not disconnect " deviceName " in Settings", clicked: false }
        bounceDeadline := A_TickCount + 6000
        while (A_TickCount < bounceDeadline) {
            Sleep(300)
            row := AudioBt_SettingsFindRowOrReveal(hwnd, deviceName)
            scope := AudioBt_SettingsClickScope(row)
            if AudioBt_SettingsFindNamed(scope, AudioBt_SettingsConnectNames())
                break
        }
        if !AudioBt_SettingsClickConnect(row)
            return { ok: false, err: "Could not click Connect for " deviceName, clicked: false }
    } else {
        return { ok: false, err: "No Connect button for " deviceName, clicked: false }
    }
    if AudioBt_SettingsWaitRowConnected(hwnd, deviceName, 8000)
        return { ok: true, err: "", clicked: true }
    return { ok: false, err: "Connect clicked but " deviceName " stayed Emparelhado", clicked: true }
}

AudioBt_SettingsUiConnect(deviceName) {
    deviceName := Trim(deviceName)
    if (deviceName = "")
        return { ok: false, err: "Missing Bluetooth device name" }
    opened := false
    hwnd := AudioBt_SettingsEnsureOpen(&opened)
    if !hwnd
        return { ok: false, err: "Could not open Bluetooth settings" }
    result := { ok: false, err: "Settings connect failed" }
    try {
        root := AudioBt_SettingsEnsureBtPage(hwnd, 15000)
        if !root {
            result.err := "Bluetooth settings page did not load"
            return result
        }
        AudioBt_SettingsDropOtherAudio(root, deviceName)
        attempt := AudioBt_SettingsTryConnectOnce(hwnd, deviceName)
        if (!attempt.ok) {
            ; One retry: fresh page/row + click (covers Emparelhado stuck and transient UIA misses).
            Sleep(400)
            root := AudioBt_SettingsEnsureBtPage(hwnd, 8000)
            if root
                attempt := AudioBt_SettingsTryConnectOnce(hwnd, deviceName)
        }
        if (attempt.ok) {
            result.ok := true
            result.err := ""
            return result
        }
        result.err := attempt.err != "" ? attempt.err : "Settings connect failed"
        return result
    } finally {
        if opened {
            try WinClose("ahk_id " hwnd)
            catch {
            }
        }
    }
}
