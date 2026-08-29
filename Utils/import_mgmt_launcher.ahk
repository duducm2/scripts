; =============================================================================
; Utils module: import_mgmt_launcher.ahk
; Import Management main menu (Utility Shortcuts [J]) — finance & palace imports
; Agent docs: docs/prompt-data-output-and-finance-packs.md
; =============================================================================

global g_ImportMgmtGui := false
global g_ImportMgmtHotkeys := []

ImportMgmt_LaunchApp() {
    ImportMgmt_ShowMainMenu()
}

ImportMgmt_CloseGui() {
    global g_ImportMgmtGui
    ImportMgmt_UnbindHotkeys()
    try {
        if (IsObject(g_ImportMgmtGui))
            g_ImportMgmtGui.Destroy()
    } catch {
    }
    g_ImportMgmtGui := false
}

ImportMgmt_UnbindHotkeys() {
    global g_ImportMgmtHotkeys
    try HotIf(ImportMgmt_HotIfKeys)
    catch {
    }
    for key in g_ImportMgmtHotkeys {
        try Hotkey(key, "Off")
        catch {
        }
    }
    g_ImportMgmtHotkeys := []
    try HotIf()
    catch {
    }
}

ImportMgmt_HotIfKeys(*) {
    global g_ImportMgmtGui
    if (!IsObject(g_ImportMgmtGui))
        return false
    try {
        return WinActive("ahk_id " g_ImportMgmtGui.Hwnd)
    } catch {
        return false
    }
}

ImportMgmt_BindHotkeys(pairs) {
    global g_ImportMgmtGui, g_ImportMgmtHotkeys
    ImportMgmt_UnbindHotkeys()
    if (!IsObject(g_ImportMgmtGui))
        return
    try HotIf(ImportMgmt_HotIfKeys)
    catch {
        return
    }
    for p in pairs {
        try {
            Hotkey(p[1], p[2], "On")
            g_ImportMgmtHotkeys.Push(p[1])
        } catch {
        }
    }
    try HotIf()
    catch {
    }
}

ImportMgmt_CenterGui(guiObj, w := 560, h := 220) {
    MonitorGetWorkArea(MonitorGetPrimary(), &L, &T, &R, &B)
    x := L + ((R - L) - w) // 2
    y := T + ((B - T) - h) // 2
    guiObj.Show("x" . x . " y" . y . " w" . w . " h" . h)
}

ImportMgmt_ShowMainMenu() {
    global g_ImportMgmtGui
    ImportMgmt_CloseGui()

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
    g_ImportMgmtGui.Add("Text", "x20 y84 w520", "[D]  Finance daily")
    g_ImportMgmtGui.SetFont("s9 cA0A0A0 Norm", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x20 y110 w520",
        "FINANCE_DAILY*.txt → transactions")

    g_ImportMgmtGui.SetFont("s12 cWhite Bold", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x20 y140 w520", "[M]  Finance monthly")
    g_ImportMgmtGui.SetFont("s9 cA0A0A0 Norm", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x20 y166 w520",
        "FINANCE_MONTHLY*.txt → accounts / goals adjustments")

    g_ImportMgmtGui.SetFont("s12 cWhite Bold", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x20 y196 w520", "[P]  Palace mnemonic pack")
    g_ImportMgmtGui.SetFont("s9 cA0A0A0 Norm", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x20 y222 w520",
        "PALACE_PACK*.txt → palaces / beasts / atoms")

    g_ImportMgmtGui.SetFont("s12 cWhite Bold", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x20 y252 w520", "[L]  Study plan pack")
    g_ImportMgmtGui.SetFont("s9 cA0A0A0 Norm", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x20 y278 w520",
        "PLAN_PACK*.txt → study plans")

    g_ImportMgmtGui.SetFont("s12 cWhite Bold", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x20 y308 w520", "[T]  Task pack")
    g_ImportMgmtGui.SetFont("s9 cA0A0A0 Norm", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x20 y334 w520",
        "TASK_PACK*.txt → projects / tasks / info")

    g_ImportMgmtGui.SetFont("s12 cWhite Bold", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x20 y364 w520", "[N]  Desktop names")
    g_ImportMgmtGui.SetFont("s9 cA0A0A0 Norm", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x20 y390 w520",
        "Manage pack filenames — copy, add, edit, delete")

    g_ImportMgmtGui.SetFont("s9 c808080", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x20 y430 w520",
        "Also available: Finance [F], Memory Palace, and Tasks launchers")
    g_ImportMgmtGui.Add("Text", "x20 y450 w520",
        "[H] help   Backspace utility shortcuts   Esc close")

    ImportMgmt_BindHotkeys([
        ["d", ImportMgmt_OnImportFinanceDaily],
        ["m", ImportMgmt_OnImportFinanceMonthly],
        ["p", ImportMgmt_OnImportPalacePack],
        ["l", ImportMgmt_OnImportPlanPack],
        ["t", ImportMgmt_OnImportTaskPack],
        ["n", ImportMgmt_OnDesktopNames],
        ["h", ImportMgmt_OnHelp],
        ["Backspace", (*) => ImportMgmt_ReturnToUtilityShortcuts()],
        ["Escape", (*) => ImportMgmt_CloseGui()]
    ])
    ImportMgmt_CenterGui(g_ImportMgmtGui, 560, 490)
}

ImportMgmt_OnHelp(*) {
    ImportMgmt_ShowHelp()
}

ImportMgmt_ShowHelp() {
    global g_ImportMgmtGui
    ImportMgmt_CloseGui()

    body := ImportMgmt_HelpText()
    winW := 620
    winH := 480
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
    . "[D] Finance daily    FINANCE_DAILY.txt      →  transactions`r`n"
    . "[M] Finance monthly  FINANCE_MONTHLY.txt    →  accounts / goals`r`n"
    . "[P] Palace pack      PALACE_PACK.txt        →  palaces / beasts / atoms`r`n"
    . "[L] Study plan       PLAN_PACK.txt          →  study plans`r`n"
    . "[T] Task pack        TASK_PACK.txt          →  projects / tasks / info`r`n"
    . "[N] Desktop names    clipangel_desktop_names.csv  →  CRUD + copy`r`n`r`n"
    . "Same imports also live under Finance [F] and Memory Palace. Tasks opens the web app.`r`n`r`n"
    . "CANONICAL DESKTOP NAMES (always overwrite)`r`n"
    . "Save AI packs with the exact filename above on Desktop.`r`n"
    . "Never add updated, corrected, v2, or similar suffixes.`r`n"
    . "Importer consolidates variants (*_updated*, gemini-code-….txt) to the canonical name before parsing.`r`n`r`n"
    . "WORKFLOW`r`n"
    . "1. Run the pack prompt (#!+U → Prompts, or dictation flow).`r`n"
    . "2. Save the pack to Desktop (Quick Download or copy fence).`r`n"
    . "3. Press the import key here — newest matching pack is imported.`r`n`r`n"
    . "PER-IMPORT RULES`r`n"
    . "[D] Finance daily / [M] Finance monthly`r`n"
    . "  • Confirm dialog before save; appends transactions or monthly adjustments.`r`n`r`n"
    . "[P] Palace mnemonic pack / [L] Study plan pack`r`n"
    . "  • Pack upsert with cross-link validation; confirm before save.`r`n`r`n"
    . "[T] Task pack`r`n"
    . "  • Opens Tasks web app (?import=1); preview + confirm in browser.`r`n`r`n"
    . "DESKTOP PACK NAMES ([N])`r`n"
    . "  • Registry: assets/data/clipangel_desktop_names.csv`r`n"
    . "  • [Enter] or [C] copies the bare name (e.g. FINANCE_DAILY) to clipboard`r`n"
    . "  • [A] add, [E] edit, Delete remove; Esc / Backspace returns to this menu`r`n"
    . "  • ClipAngel export uses the same list when renaming Desktop files`r`n`r`n"
    . "OUTCOMES`r`n"
    . "• Full success: local CSV saved; Desktop pack archived to */data/imported/`r`n"
    . "• Failure: Desktop fix file written (FINANCE_AI_FIX / PALACE_AI_FIX / TASK_AI_FIX)`r`n`r`n"
    . "AI FIX FILES (written on failure — always overwrite)`r`n"
    . "FINANCE_AI_FIX.txt | PALACE_AI_FIX.txt | TASK_AI_FIX.txt`r`n"
    . "Paste fix file into AI → re-deliver corrected pack → save canonical name → import again.`r`n`r`n"
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

ImportMgmt_OnImportTaskPack(*) {
    Task_ImportPackFromDesktop()
}

ImportMgmt_OnDesktopNames(*) {
    ImportMgmt_CloseGui()
    ClipAngelExport_ShowNamesManager(ImportMgmt_ShowMainMenu)
}
