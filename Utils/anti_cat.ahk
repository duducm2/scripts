; =============================================================================
; Utils module: anti_cat.ahk
; Anti-Cat — disable notebook keyboard + trackpad so a cat on the laptop
; cannot type/click. No Bluetooth keyboards are touched (elevated PnP helper).
; Toggle: Macros (#!+W → b). Unlock: same macro again, or hold Escape 4s.
; =============================================================================

global g_AntiCatOn := false
global g_AntiCatOverlay := 0
global g_AntiCatEscDownSince := 0
global g_AntiCatEscTimer := false
global g_AntiCatHeartbeatTimer := false
global g_AntiCatStateDir := ""
global g_AntiCatHelperPs1 := ""
global g_AntiCatEscHoldMs := 4000

AntiCat_StateDir() {
    global g_AntiCatStateDir
    if (g_AntiCatStateDir = "")
        g_AntiCatStateDir := A_ScriptDir "\.cursor\anti_cat"
    return g_AntiCatStateDir
}

AntiCat_HelperPath() {
    global g_AntiCatHelperPs1
    if (g_AntiCatHelperPs1 = "")
        g_AntiCatHelperPs1 := A_ScriptDir "\Utils\anti_cat_helper.ps1"
    return g_AntiCatHelperPs1
}

AntiCat_EnsureStateDir() {
    dir := AntiCat_StateDir()
    if (!DirExist(dir))
        DirCreate(dir)
    return dir
}

AntiCat_WriteFile(name, text) {
    path := AntiCat_StateDir() "\" name
    try FileDelete(path)
    catch {
    }
    try FileAppend(text, path, "UTF-8")
    catch {
    }
}

AntiCat_ReadFile(name) {
    path := AntiCat_StateDir() "\" name
    if (!FileExist(path))
        return ""
    try {
        return Trim(FileRead(path, "UTF-8"))
    } catch {
        return ""
    }
}

AntiCat_IsActive() {
    global g_AntiCatOn
    return !!g_AntiCatOn
}

AntiCat_ClearOverlay() {
    global g_AntiCatOverlay
    if (IsObject(g_AntiCatOverlay)) {
        try g_AntiCatOverlay.Destroy()
        catch {
        }
    }
    g_AntiCatOverlay := 0
}

AntiCat_ShowOverlay() {
    global g_AntiCatOverlay
    AntiCat_ClearOverlay()
    overlay := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x20")
    overlay.BackColor := "1A1A1A"
    overlay.SetFont("s11 cWhite Bold", "Segoe UI")
    overlay.Add("Text", "x16 y10 w520 h22", "Anti-Cat ON — notebook KB + trackpad locked")
    overlay.SetFont("s9 cAAAAAA Norm", "Segoe UI")
    overlay.Add("Text", "x16 y34 w520 h36",
        "Bluetooth keyboards stay active. Unlock: hold Esc 4s, or Macros → Anti-Cat again.")
    overlay.Show("NoActivate x" . (A_ScreenWidth - 560) . " y20 w552 h78")
    try WinSetTransparent(220, overlay)
    catch {
    }
    g_AntiCatOverlay := overlay
}

AntiCat_StopEscMonitor() {
    global g_AntiCatEscTimer, g_AntiCatEscDownSince
    g_AntiCatEscDownSince := 0
    if (g_AntiCatEscTimer) {
        try SetTimer(g_AntiCatEscTimer, 0)
        catch {
            try SetTimer(AntiCat_EscMonitor, 0)
            catch {
            }
        }
        g_AntiCatEscTimer := false
    }
}

AntiCat_StartEscMonitor() {
    global g_AntiCatEscTimer, g_AntiCatEscDownSince
    AntiCat_StopEscMonitor()
    g_AntiCatEscDownSince := 0
    g_AntiCatEscTimer := SetTimer(AntiCat_EscMonitor, 50)
}

AntiCat_EscMonitor(*) {
    global g_AntiCatOn, g_AntiCatEscDownSince, g_AntiCatEscHoldMs
    if (!g_AntiCatOn) {
        AntiCat_StopEscMonitor()
        return
    }
    if (GetKeyState("Escape", "P")) {
        if (!g_AntiCatEscDownSince)
            g_AntiCatEscDownSince := A_TickCount
        else if ((A_TickCount - g_AntiCatEscDownSince) >= g_AntiCatEscHoldMs) {
            g_AntiCatEscDownSince := 0
            DisableAntiCat("esc")
        }
    } else {
        g_AntiCatEscDownSince := 0
    }
}

AntiCat_StopHeartbeat() {
    global g_AntiCatHeartbeatTimer
    if (g_AntiCatHeartbeatTimer) {
        try SetTimer(g_AntiCatHeartbeatTimer, 0)
        catch {
            try SetTimer(AntiCat_Heartbeat, 0)
            catch {
            }
        }
        g_AntiCatHeartbeatTimer := false
    }
}

AntiCat_StartHeartbeat() {
    global g_AntiCatHeartbeatTimer
    AntiCat_StopHeartbeat()
    AntiCat_Heartbeat()
    g_AntiCatHeartbeatTimer := SetTimer(AntiCat_Heartbeat, 2000)
}

AntiCat_Heartbeat(*) {
    if (!AntiCat_IsActive())
        return
    ; Touch the file so the elevated helper can use LastWriteTime as a liveness signal.
    AntiCat_WriteFile("heartbeat", String(A_TickCount))
}

AntiCat_RequestUnlock() {
    AntiCat_WriteFile("cmd", "unlock")
}

AntiCat_WaitStatus(want, timeoutMs := 45000) {
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        st := AntiCat_ReadFile("status")
        if (st = want)
            return st
        if (SubStr(st, 1, 6) = "error:")
            return st
        Sleep 150
    }
    return "error:timeout waiting for helper (" want ")"
}

EnableAntiCat() {
    global g_AntiCatOn
    if (g_AntiCatOn)
        return

    helper := AntiCat_HelperPath()
    if (!FileExist(helper)) {
        try ShowCenteredOverlay_Utils("❌ Anti-Cat helper missing: Utils\anti_cat_helper.ps1", 3000, BANNER_ACCENT_ERROR
        )
        catch {
        }
        return
    }

    dir := AntiCat_EnsureStateDir()
    AntiCat_WriteFile("status", "starting")
    AntiCat_WriteFile("heartbeat", String(A_TickCount))
    try FileDelete(dir "\cmd")
    catch {
    }

    try ShowCenteredOverlay_Utils("🐱 Anti-Cat: approve UAC to lock notebook input…", 2500, BANNER_ACCENT_INTERMEDIATE)
    catch {
    }

    try {
        Run('powershell.exe -NoProfile -ExecutionPolicy Bypass -File "' helper '" -StateDir "' dir '"', , "Hide")
    } catch as e {
        try ShowCenteredOverlay_Utils("❌ Anti-Cat failed to start helper: " e.Message, 3000, BANNER_ACCENT_ERROR)
        catch {
        }
        return
    }

    st := AntiCat_WaitStatus("on", 60000)
    if (st != "on") {
        msg := (st != "") ? st : "error:unknown"
        try ShowCenteredOverlay_Utils("❌ Anti-Cat not armed — " msg, 3500, BANNER_ACCENT_ERROR)
        catch {
        }
        return
    }

    g_AntiCatOn := true
    AntiCat_ShowOverlay()
    AntiCat_StartEscMonitor()
    AntiCat_StartHeartbeat()
    try ShowCenteredOverlay_Utils("🐱 Anti-Cat ON — Esc 4s or macro again to unlock", 2200, BANNER_ACCENT_INFO)
    catch {
    }
}

DisableAntiCat(reason := "macro") {
    global g_AntiCatOn
    wasOn := g_AntiCatOn
    g_AntiCatOn := false
    AntiCat_StopEscMonitor()
    AntiCat_StopHeartbeat()
    AntiCat_ClearOverlay()

    if (wasOn || AntiCat_ReadFile("status") = "on" || AntiCat_ReadFile("status") = "starting") {
        AntiCat_RequestUnlock()
        st := AntiCat_WaitStatus("off", 20000)
        if (st != "off" && SubStr(st, 1, 6) = "error:") {
            try ShowCenteredOverlay_Utils("⚠ Anti-Cat unlock: " st, 3000, BANNER_ACCENT_ERROR)
            catch {
            }
            return
        }
    }

    if (wasOn) {
        label := (reason = "esc") ? "hold Esc" : "macro"
        try ShowCenteredOverlay_Utils("🐱 Anti-Cat OFF (" label ")", 1600, BANNER_ACCENT_SUCCESS)
        catch {
        }
    }
}

ToggleAntiCat(*) {
    if (AntiCat_IsActive())
        DisableAntiCat("macro")
    else
        EnableAntiCat()
}

AntiCat_OnExit(*) {
    if (AntiCat_IsActive() || AntiCat_ReadFile("status") = "on")
        DisableAntiCat("exit")
}

OnExit(AntiCat_OnExit)

RegisterMacro(ToggleAntiCat, "🐱 Anti-Cat (lock notebook KB/trackpad; BT keyboards OK)", "b")