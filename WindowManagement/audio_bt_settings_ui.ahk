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
    return ["Mais configurações", "More settings", "More options"]
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

AudioBt_SettingsClickNamed(scope, names) {
    el := AudioBt_SettingsFindNamed(scope, names)
    if !el
        return false
    return ClickSeq_Invoke(el)
}

AudioBt_SettingsRowMatches(groupName, deviceName) {
    n := Trim(deviceName)
    if (n = "" || groupName = "")
        return false
    return InStr(groupName, n ", ") = 1 || InStr(groupName, n ",") = 1
}

AudioBt_SettingsRowConnected(groupName) {
    if InStr(groupName, "Estado Conectado") || InStr(groupName, "Status Connected")
        return true
    return false
}

AudioBt_SettingsFindRow(root, deviceName) {
    if !root
        return 0
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
            if toggle
                return root
        }
        Sleep(200)
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
    return hwnd
}

AudioBt_SettingsExpand(row) {
    if !row
        return false
    btn := AudioBt_SettingsFindNamed(row, AudioBt_SettingsExpandNames())
    if !btn
        return false
    if !ClickSeq_Invoke(btn)
        return false
    Sleep(500)
    return true
}

AudioBt_SettingsDisconnectRow(row) {
    if AudioBt_SettingsClickNamed(row, AudioBt_SettingsDisconnectNames())
        return true
    if !AudioBt_SettingsExpand(row)
        return false
    return AudioBt_SettingsClickNamed(row, AudioBt_SettingsDisconnectNames())
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
        if !AudioBt_SettingsRowConnected(nm)
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
        try nm := row ? row.Name : ""
        catch
            nm := ""
        if AudioBt_SettingsRowConnected(nm)
            return true
        Sleep(250)
    }
    return false
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
        root := AudioBt_SettingsWaitPage(hwnd, 15000)
        if !root {
            result.err := "Bluetooth settings page did not load"
            return result
        }
        AudioBt_SettingsDropOtherAudio(root, deviceName)
        root := AudioBt_SettingsRoot(hwnd)
        row := AudioBt_SettingsFindRow(root, deviceName)
        if !row {
            result.err := "Device not found in Settings: " deviceName
            return result
        }
        try rowName := row.Name
        catch
            rowName := ""
        clicked := false
        if AudioBt_SettingsFindNamed(row, AudioBt_SettingsConnectNames()) {
            clicked := AudioBt_SettingsClickNamed(row, AudioBt_SettingsConnectNames())
        } else if AudioBt_SettingsRowConnected(rowName) {
            if !AudioBt_SettingsDisconnectRow(row) {
                result.err := "Could not disconnect " deviceName " in Settings"
                return result
            }
            bounceDeadline := A_TickCount + 6000
            while (A_TickCount < bounceDeadline) {
                Sleep(300)
                root := AudioBt_SettingsRoot(hwnd)
                row := AudioBt_SettingsFindRow(root, deviceName)
                if AudioBt_SettingsFindNamed(row, AudioBt_SettingsConnectNames())
                    break
            }
            clicked := AudioBt_SettingsClickNamed(row, AudioBt_SettingsConnectNames())
        } else {
            result.err := "No Connect button for " deviceName
            return result
        }
        if !clicked {
            result.err := "Could not click Connect for " deviceName
            return result
        }
        AudioBt_SettingsWaitRowConnected(hwnd, deviceName, 8000)
        result.ok := true
        result.err := ""
        return result
    } finally {
        if opened {
            try WinClose("ahk_id " hwnd)
            catch {
            }
        }
    }
}
