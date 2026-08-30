; =============================================================================
; Utils module: mnemonic_palace_help.ahk
; Memory Palace glossary / vocabulary help
; =============================================================================

Palace_ShowHelp() {
    global g_PalaceGui
    Palace_CloseGui()
    MonitorGetWorkArea(MonitorGetPrimary(), &L, &T, &R, &B)
    workH := B - T
    workW := R - L
    ; -DPIScale below: sizes are physical pixels matching MonitorGetWorkArea
    winW := Min(680, Max(420, workW - 48))
    winH := Min(480, Max(340, Integer(workH * 0.55)))
    chrome := 132
    bodyH := Max(180, winH - chrome)
    bodyW := winW - 32

    g_PalaceGui := Gui("+AlwaysOnTop +ToolWindow -DPIScale", "Memory Palace — Glossary")
    g_PalaceGui.SetFont("s10", "Segoe UI")
    g_PalaceGui.BackColor := "1E1E1E"
    g_PalaceGui.Add("Text", "x16 y12 w" . bodyW . " cWhite", "Memory Palace — vocabulary for this system")
    g_PalaceGui.SetFont("s9 cA0A0A0", "Segoe UI")
    g_PalaceGui.Add("Text", "x16 y36 w" . bodyW,
        "Software model only. Technique rules: mnemonics/technique/README.md (mirrored).")

    body := ""
    for term in Palace_Terms() {
        if (body != "")
            body .= "`r`n`r`n"
        body .= term[1] . "`r`n" . term[2]
    }
    edit := g_PalaceGui.Add("Edit",
        "x16 y64 w" . bodyW . " h" . bodyH
        . " ReadOnly -WantReturn +VScroll Multi Background252526 cD4D4D4",
        body)
    try edit.SetFont("s10", "Segoe UI")
    catch {
    }
    Palace_StyleDarkEdit(edit)

    g_PalaceGui.SetFont("s9 c808080", "Segoe UI")
    g_PalaceGui.Add("Text", "x16 y" . (64 + bodyH + 10) . " w" . bodyW,
    "Esc / Backspace — main menu. [L] Plans · [J] import plan pack · [I] mnemonic pack (.txt|.csv; auto palace numbers) · [O] plans on GitHub · Markdown sync + push via Utility Shortcuts [G]. Browse [L] plans under a study. Dashboard: P plans · M method · L latest · Save (Plans panel)."
    )
    g_PalaceGui.OnEvent("Close", (*) => Palace_CloseGui())
    g_PalaceGui.OnEvent("Escape", (*) => Palace_ShowMainMenu())
    Palace_BindHotkeys([
        ["Backspace", (*) => Palace_ShowMainMenu()],
        ["Escape", (*) => Palace_ShowMainMenu()]
    ])
    try DllCall("dwmapi\DwmSetWindowAttribute", "ptr", g_PalaceGui.Hwnd, "uint", 20, "int*", 1, "int", 4)
    catch {
    }
    Palace_CenterGui(g_PalaceGui, winW, winH)
}

Palace_StyleDarkEdit(editCtrl) {
    hwnd := 0
    try hwnd := editCtrl.Hwnd
    catch {
        return
    }
    if (!hwnd)
        return
    ; Prefer dark Explorer theme so the vertical scrollbar matches dark mode
    try DllCall("uxtheme\SetWindowTheme", "ptr", hwnd, "wstr", "DarkMode_Explorer", "wstr", "")
    catch {
        try DllCall("uxtheme\SetWindowTheme", "ptr", hwnd, "wstr", "DarkMode_CFD", "wstr", "")
        catch {
        }
    }
    ; EM_SETBKGNDCOLOR — COLORREF BGR for #252526
    try SendMessage(0x443, 0, 0x262525, hwnd)
    catch {
    }
}
