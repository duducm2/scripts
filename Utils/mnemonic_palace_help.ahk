; =============================================================================
; Utils module: mnemonic_palace_help.ahk
; Memory Palace glossary / vocabulary help
; =============================================================================

Palace_ShowHelp() {
    global g_PalaceGui
    Palace_CloseGui()
    g_PalaceGui := Gui("+AlwaysOnTop +ToolWindow", "Memory Palace — Glossary")
    g_PalaceGui.SetFont("s10", "Segoe UI")
    g_PalaceGui.BackColor := "1E1E1E"
    g_PalaceGui.Add("Text", "x16 y12 w640 cWhite", "Memory Palace — vocabulary for this system")
    g_PalaceGui.SetFont("s9 cA0A0A0", "Segoe UI")
    g_PalaceGui.Add("Text", "x16 y36 w640",
        "Software model only. Technique rules stay in notes/studies/technique/README.md.")

    y := 64
    for term in Palace_Terms() {
        g_PalaceGui.SetFont("s11 cF1C40F Bold", "Segoe UI")
        g_PalaceGui.Add("Text", "x16 y" . y . " w640", term[1])
        y += 22
        g_PalaceGui.SetFont("s9 cE0E0E0 Norm", "Segoe UI")
        g_PalaceGui.Add("Text", "x16 y" . y . " w640 Wrap", term[2])
        y += 48
    }

    g_PalaceGui.SetFont("s9 c808080", "Segoe UI")
    g_PalaceGui.Add("Text", "x16 y" . y . " w640", "Esc / Backspace — return to main menu")
    g_PalaceGui.OnEvent("Close", (*) => Palace_CloseGui())
    g_PalaceGui.OnEvent("Escape", (*) => Palace_ShowMainMenu())
    Palace_BindHotkeys([
        ["Backspace", (*) => Palace_ShowMainMenu()],
        ["Escape", (*) => Palace_ShowMainMenu()]
    ])
    Palace_CenterGui(g_PalaceGui, 680, y + 50)
}
