; =============================================================================
; Utils module: import_mgmt_launcher.ahk
; Import Management main menu (Utility Shortcuts [J])
; =============================================================================

ImportMgmt_LaunchApp() {
    ImportMgmt_EnsureData()
    ImportMgmt_ShowMainMenu()
}

ImportMgmt_ShowMainMenu() {
    global g_ImportMgmtGui
    ImportMgmt_CloseGui()
    ImportMgmt_EnsureData()

    g_ImportMgmtGui := Gui("+AlwaysOnTop +ToolWindow", "Import Management")
    g_ImportMgmtGui.SetFont("s10", "Segoe UI")
    g_ImportMgmtGui.BackColor := "1E1E1E"
    g_ImportMgmtGui.OnEvent("Close", (*) => ImportMgmt_CloseGui())
    g_ImportMgmtGui.OnEvent("Escape", (*) => ImportMgmt_CloseGui())

    g_ImportMgmtGui.SetFont("s16 cWhite Bold", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x20 y16 w520", "Import Management")
    g_ImportMgmtGui.SetFont("s10 cC0C0C0 Norm", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x20 y48 w520", "Bridge AI-exported job search data into opportunities.csv")

    g_ImportMgmtGui.SetFont("s12 cWhite Bold", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x20 y96 w520", "[I]  AI import")
    g_ImportMgmtGui.SetFont("s9 cA0A0A0 Norm", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x20 y122 w520",
        "Import newest JOB_SEARCH_UPDATE*.txt from Desktop (also .csv / gemini-code*.txt)")

    g_ImportMgmtGui.SetFont("s9 c808080", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x20 y170 w520",
        "Prompt: Utility Shortcuts → Prompts → [j] Job search status update")
    g_ImportMgmtGui.Add("Text", "x20 y190 w520",
        "Backspace utility shortcuts   Esc close")

    ImportMgmt_BindHotkeys([
        ["i", ImportMgmt_OnImport],
        ["Backspace", (*) => ImportMgmt_ReturnToUtilityShortcuts()],
        ["Escape", (*) => ImportMgmt_CloseGui()]
    ])
    ImportMgmt_CenterGui(g_ImportMgmtGui, 560, 240)
}

ImportMgmt_ReturnToUtilityShortcuts() {
    ImportMgmt_CloseGui()
    try ShowHotstringSelector()
    catch {
    }
}

ImportMgmt_OnImport(*) {
    ImportMgmt_ImportFromDesktop()
}
