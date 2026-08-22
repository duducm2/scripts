; =============================================================================
; Utils module: mnemonic_palace_help.ahk
; Memory Palace glossary / vocabulary help
; =============================================================================

Palace_ShowHelp() {
    global g_PalaceGui
    Palace_CloseGui()
    MonitorGetWorkArea(MonitorGetPrimary(), &L, &T, &R, &B)
    maxH := Max(320, Integer((B - T) * 0.8))
    winW := 680
    winH := maxH
    bodyH := Max(160, winH - 110)

    g_PalaceGui := Gui("+AlwaysOnTop +ToolWindow", "Memory Palace — Glossary")
    g_PalaceGui.SetFont("s10", "Segoe UI")
    g_PalaceGui.BackColor := "1E1E1E"
    g_PalaceGui.Add("Text", "x16 y12 w640 cWhite", "Memory Palace — vocabulary for this system")
    g_PalaceGui.SetFont("s9 cA0A0A0", "Segoe UI")
    g_PalaceGui.Add("Text", "x16 y36 w640",
        "Software model only. Technique rules stay in notes/studies/technique/README.md.")

    body := ""
    for term in Palace_Terms() {
        if (body != "")
            body .= "`r`n`r`n"
        body .= term[1] . "`r`n" . term[2]
    }
    edit := g_PalaceGui.Add("Edit", "x16 y64 w640 h" . bodyH . " ReadOnly -WantReturn +VScroll Multi", body)
    try edit.SetFont("s10", "Segoe UI")
    catch {
    }

    g_PalaceGui.SetFont("s9 c808080", "Segoe UI")
    g_PalaceGui.Add("Text", "x16 y" . (64 + bodyH + 10) . " w640", "Esc / Backspace — return to main menu")
    g_PalaceGui.OnEvent("Close", (*) => Palace_CloseGui())
    g_PalaceGui.OnEvent("Escape", (*) => Palace_ShowMainMenu())
    Palace_BindHotkeys([
        ["Backspace", (*) => Palace_ShowMainMenu()],
        ["Escape", (*) => Palace_ShowMainMenu()]
    ])
    Palace_CenterGui(g_PalaceGui, winW, winH)
}
