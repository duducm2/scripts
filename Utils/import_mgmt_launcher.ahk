; =============================================================================
; Utils module: import_mgmt_launcher.ahk
; Import Management main menu (Utility Shortcuts [J])
; Agent docs: docs/prompt-data-output-and-finance-packs.md
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
    g_ImportMgmtGui.Add("Text", "x20 y48 w520",
        "Import AI-exported packs from Desktop into local CSV data")

    g_ImportMgmtGui.SetFont("s12 cWhite Bold", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x20 y84 w520", "[I]  Job search")
    g_ImportMgmtGui.SetFont("s9 cA0A0A0 Norm", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x20 y110 w520",
        "JOB_SEARCH_UPDATE*.txt → opportunities.csv")

    g_ImportMgmtGui.SetFont("s12 cWhite Bold", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x20 y140 w520", "[D]  Finance daily")
    g_ImportMgmtGui.SetFont("s9 cA0A0A0 Norm", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x20 y166 w520",
        "FINANCE_DAILY*.txt → transactions")

    g_ImportMgmtGui.SetFont("s12 cWhite Bold", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x20 y196 w520", "[M]  Finance monthly")
    g_ImportMgmtGui.SetFont("s9 cA0A0A0 Norm", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x20 y222 w520",
        "FINANCE_MONTHLY*.txt → accounts / goals adjustments")

    g_ImportMgmtGui.SetFont("s12 cWhite Bold", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x20 y252 w520", "[P]  Palace mnemonic pack")
    g_ImportMgmtGui.SetFont("s9 cA0A0A0 Norm", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x20 y278 w520",
        "PALACE_PACK*.txt → palaces / beasts / atoms")

    g_ImportMgmtGui.SetFont("s12 cWhite Bold", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x20 y308 w520", "[L]  Study plan pack")
    g_ImportMgmtGui.SetFont("s9 cA0A0A0 Norm", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x20 y334 w520",
        "PLAN_PACK*.txt → study plans")

    g_ImportMgmtGui.SetFont("s9 c808080", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x20 y378 w520",
        "Also available: Finance [F] and Memory Palace [N] launchers")
    g_ImportMgmtGui.Add("Text", "x20 y398 w520",
        "[H] help   Backspace utility shortcuts   Esc close")

    ImportMgmt_BindHotkeys([
        ["i", ImportMgmt_OnImportJobSearch],
        ["d", ImportMgmt_OnImportFinanceDaily],
        ["m", ImportMgmt_OnImportFinanceMonthly],
        ["p", ImportMgmt_OnImportPalacePack],
        ["l", ImportMgmt_OnImportPlanPack],
        ["h", ImportMgmt_OnHelp],
        ["Backspace", (*) => ImportMgmt_ReturnToUtilityShortcuts()],
        ["Escape", (*) => ImportMgmt_CloseGui()]
    ])
    ImportMgmt_CenterGui(g_ImportMgmtGui, 560, 440)
}

ImportMgmt_OnHelp(*) {
    ImportMgmt_ShowHelp()
}

ImportMgmt_ShowHelp() {
    global g_ImportMgmtGui
    ImportMgmt_CloseGui()

    body := ImportMgmt_HelpText()
    winW := 620
    winH := 520
    bodyW := winW - 32
    bodyH := winH - 100

    g_ImportMgmtGui := Gui("+AlwaysOnTop +ToolWindow", "Import Management — Help")
    g_ImportMgmtGui.SetFont("s10", "Segoe UI")
    g_ImportMgmtGui.BackColor := "1E1E1E"
    g_ImportMgmtGui.SetFont("s14 cWhite Bold", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x16 y12 w" . bodyW, "Import rules")
    g_ImportMgmtGui.SetFont("s9 cA0A0A0 Norm", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x16 y38 w" . bodyW,
        "Canonical pack names, overwrite policy, fix-file recovery")
    edit := g_ImportMgmtGui.Add("Edit",
        "x16 y64 w" . bodyW . " h" . bodyH
        . " ReadOnly -WantReturn +VScroll Multi Background252526 cD4D4D4",
        body)
    try edit.SetFont("s10", "Consolas")
    catch {
        try edit.SetFont("s10", "Segoe UI")
        catch {
        }
    }
    g_ImportMgmtGui.SetFont("s9 c808080", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x16 y" . (64 + bodyH + 10) . " w" . bodyW,
    "Esc / Backspace — main menu")
    g_ImportMgmtGui.OnEvent("Close", (*) => ImportMgmt_ShowMainMenu())
    g_ImportMgmtGui.OnEvent("Escape", (*) => ImportMgmt_ShowMainMenu())
    ImportMgmt_BindHotkeys([
        ["Backspace", (*) => ImportMgmt_ShowMainMenu()],
        ["Escape", (*) => ImportMgmt_ShowMainMenu()]
    ])
    ImportMgmt_CenterGui(g_ImportMgmtGui, winW, winH)
}

ImportMgmt_HelpText() {
    return "IMPORT KEYS (Utility Shortcuts [J])`r`n"
    . "[I] Job search       JOB_SEARCH_UPDATE.txt  →  opportunities.csv`r`n"
    . "[D] Finance daily    FINANCE_DAILY.txt      →  transactions`r`n"
    . "[M] Finance monthly  FINANCE_MONTHLY.txt    →  accounts / goals`r`n"
    . "[P] Palace pack      PALACE_PACK.txt        →  palaces / beasts / atoms`r`n"
    . "[L] Study plan       PLAN_PACK.txt          →  study plans`r`n`r`n"
    . "Same imports also live under Finance [F] and Memory Palace [N].`r`n`r`n"
    . "CANONICAL DESKTOP NAMES (always overwrite)`r`n"
    . "Save AI packs with the exact filename above on Desktop.`r`n"
    . "Never add updated, corrected, v2, or similar suffixes.`r`n"
    . "Importer consolidates variants (*_updated*, gemini-code-….txt) to the canonical name before parsing.`r`n`r`n"
    . "WORKFLOW`r`n"
    . "1. Run the pack prompt (#!+U → Prompts, or dictation flow).`r`n"
    . "2. Save the pack to Desktop (Quick Download or copy fence).`r`n"
    . "3. Press the import key here — newest matching pack is imported.`r`n`r`n"
    . "OUTCOMES`r`n"
    . "• Full success: local CSV saved; Desktop pack archived to */data/imported/`r`n"
    . "• Partial failure (job search): good rows saved; pack stays on Desktop; fix file written`r`n"
    . "• Total failure: no save; Desktop fix file written`r`n`r`n"
    . "AI FIX FILES (written on failure — always overwrite)`r`n"
    . "FINANCE_AI_FIX.txt | PALACE_AI_FIX.txt | JOB_SEARCH_AI_FIX.txt`r`n"
    . "Paste fix file into AI → re-deliver full corrected pack → save canonical name → import again.`r`n"
    . "Job search partial import: re-deliver the FULL pack (include rows that already imported).`r`n`r`n"
    . "PACK FORMAT`r`n"
    . "===PREVIEW=== … ===END_PREVIEW===`r`n"
    . "===FILE: PACK.csv=== … ===END_FILE===`r`n"
    . "The AI never writes to your disk — you save the file yourself."
}

ImportMgmt_ReturnToUtilityShortcuts() {
    ImportMgmt_CloseGui()
    try ShowHotstringSelector()
    catch {
    }
}

ImportMgmt_OnImportJobSearch(*) {
    ImportMgmt_ImportFromDesktop()
}

ImportMgmt_OnImportFinanceDaily(*) {
    Finance_ImportDaily()
}

ImportMgmt_OnImportFinanceMonthly(*) {
    Finance_ImportMonthly()
}

ImportMgmt_OnImportPalacePack(*) {
    Palace_ImportMnemonicsFromDesktop()
}

ImportMgmt_OnImportPlanPack(*) {
    Palace_ImportPlanPackFromDesktop()
}
