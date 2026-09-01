; =============================================================================
; Utils module: import_mgmt_launcher.ahk
; Import Management hub — sole AHK UI for pack imports
; Entry: #!+X, Utility Shortcuts [J], Win+Alt+Shift+F double-tap
; Agent docs: docs/prompt-data-output-and-finance-packs.md
; =============================================================================

global g_ImportMgmtGui := false
global g_ImportMgmtHotkeys := []
global g_ImportMgmtLv := false
global g_ImportMgmtCatalog := []

ImportMgmt_LaunchApp() {
    ImportMgmt_ShowMainMenu()
}

ImportMgmt_CloseGui() {
    global g_ImportMgmtGui, g_ImportMgmtLv
    ImportMgmt_UnbindHotkeys()
    try {
        if (IsObject(g_ImportMgmtGui))
            g_ImportMgmtGui.Destroy()
    } catch {
    }
    g_ImportMgmtGui := false
    g_ImportMgmtLv := false
}

ImportMgmt_IsOpen() {
    global g_ImportMgmtGui
    try {
        return IsObject(g_ImportMgmtGui)
    } catch {
        return false
    }
}

; Close hub if open. Returns true when it was open (caller should skip domain menus).
ImportMgmt_CloseIfOpen() {
    if (!ImportMgmt_IsOpen())
        return false
    ImportMgmt_CloseGui()
    return true
}

; Copy AI fix file to clipboard, show ≥5s orientation toast, close Import Manager if open.
; Returns true when the hub was closed.
ImportMgmt_OnAiFixReady(path, label) {
    if (path = "" || !FileExist(path))
        return ImportMgmt_CloseIfOpen()
    body := ""
    try {
        f := FileOpen(path, "r", "UTF-8")
        if (f) {
            body := f.Read()
            f.Close()
            if (SubStr(body, 1, 1) = Chr(0xFEFF))
                body := SubStr(body, 2)
        }
    } catch {
        body := ""
    }
    if (body != "") {
        try A_Clipboard := body
        catch {
        }
    }
    msg := "AI fix copied — paste into your AI companion · Desktop " . label
    try ShowCenteredOverlay_Utils(msg, 5000, BANNER_ACCENT_ERROR)
    catch {
        TrayTip("Import", msg)
    }
    return ImportMgmt_CloseIfOpen()
}

; Close Import Manager after a successful import from the hub. Returns true when it closed.
ImportMgmt_OnImportSuccess() {
    return ImportMgmt_CloseIfOpen()
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

; Single source for ListView rows, letter accelerators, help sections, and copy-name targets.
ImportMgmt_Catalog() {
    return [
        Map("char", "D", "name", "Finance daily", "detail", "FINANCE_DAILY.txt → transactions",
            "fileName", "FINANCE_DAILY.txt", "run", ImportMgmt_RunFinanceDaily),
        Map("char", "M", "name", "Finance monthly", "detail", "FINANCE_MONTHLY.txt → accounts / goals",
            "fileName", "FINANCE_MONTHLY.txt", "run", ImportMgmt_RunFinanceMonthly),
        Map("char", "P", "name", "Palace mnemonic pack", "detail", "PALACE_PACK.txt → palaces / beasts / atoms",
            "fileName", "PALACE_PACK.txt", "run", ImportMgmt_RunPalacePack),
        Map("char", "L", "name", "Study plan pack", "detail", "PLAN_PACK.txt → study plans",
            "fileName", "PLAN_PACK.txt", "run", ImportMgmt_RunPlanPack),
        Map("char", "T", "name", "Task pack", "detail", "TASK_PACK.txt → projects / tasks / info",
            "fileName", "TASK_PACK.txt", "run", ImportMgmt_RunTaskPack),
        Map("char", "Q", "name", "Palace quick image", "detail", "Newest Desktop PNG/JPG → palace missing image",
            "fileName", "PALACE_QUICK_IMAGE.png", "run", ImportMgmt_RunQuickImage),
        Map("char", "N", "name", "Quick Download names", "detail",
            "#!+9 rename list — clipangel_desktop_names.csv — CRUD + copy",
            "fileName", "clipangel_desktop_names.csv", "run", ImportMgmt_RunDesktopNames),
        Map("char", "H", "name", "Help", "detail", "Per-workflow rules and outcomes",
            "run", ImportMgmt_OnHelp)
    ]
}

ImportMgmt_ShowMainMenu() {
    global g_ImportMgmtGui, g_ImportMgmtLv, g_ImportMgmtCatalog
    ImportMgmt_CloseGui()

    g_ImportMgmtCatalog := ImportMgmt_Catalog()
    contentW := 700
    lvH := 248
    guiW := 740
    guiH := 348

    g_ImportMgmtGui := Gui("+AlwaysOnTop +ToolWindow", "Import Management")
    g_ImportMgmtGui.SetFont("s10", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "w" . contentW,
        "Char = import   Shift+Enter / double-click = import   Shift+C = copy pack name   Shift+Backspace = utility shortcuts   Esc = close"
    )
    g_ImportMgmtLv := g_ImportMgmtGui.Add("ListView", "w" . contentW . " h" . lvH . " -Multi",
        ["Char", "Workflow", "Pack / detail"])
    g_ImportMgmtLv.OnEvent("DoubleClick", ImportMgmt_OnListActivate)
    g_ImportMgmtGui.OnEvent("Close", (*) => ImportMgmt_CloseGui())
    g_ImportMgmtGui.OnEvent("Escape", (*) => ImportMgmt_CloseGui())

    for item in g_ImportMgmtCatalog
        g_ImportMgmtLv.Add("", item["char"], item["name"], item["detail"])
    try g_ImportMgmtLv.ModifyCol(1, 50)
    try g_ImportMgmtLv.ModifyCol(2, 200)
    try g_ImportMgmtLv.ModifyCol(3, 430)
    if (g_ImportMgmtCatalog.Length) {
        try g_ImportMgmtLv.Modify(1, "Select Focus Vis")
        catch {
        }
    }

    pairs := []
    for item in g_ImportMgmtCatalog {
        ch := StrLower(item["char"])
        runFn := item["run"]
        pairs.Push([ch, runFn])
    }
    pairs.Push(["+Enter", ImportMgmt_OnListActivate])
    pairs.Push(["+c", ImportMgmt_CopySelectedFileName])
    pairs.Push(["+Backspace", (*) => ImportMgmt_ReturnToUtilityShortcuts()])
    pairs.Push(["Escape", (*) => ImportMgmt_CloseGui()])
    ImportMgmt_BindHotkeys(pairs)

    ImportMgmt_CenterGui(g_ImportMgmtGui, guiW, guiH)
    try g_ImportMgmtLv.Focus()
    catch {
    }
}

ImportMgmt_SelectedIndex() {
    global g_ImportMgmtLv
    if (!IsObject(g_ImportMgmtLv))
        return 0
    try {
        return g_ImportMgmtLv.GetNext(0)
    } catch {
        return 0
    }
}

ImportMgmt_OnListActivate(*) {
    global g_ImportMgmtCatalog
    idx := ImportMgmt_SelectedIndex()
    if (idx < 1 || idx > g_ImportMgmtCatalog.Length)
        return
    runFn := g_ImportMgmtCatalog[idx]["run"]
    runFn()
}

ImportMgmt_CopySelectedFileName(*) {
    global g_ImportMgmtCatalog
    idx := ImportMgmt_SelectedIndex()
    if (idx < 1 || idx > g_ImportMgmtCatalog.Length)
        return
    item := g_ImportMgmtCatalog[idx]
    if (!item.Has("fileName") || item["fileName"] = "") {
        try ShowCenteredOverlay_Utils("No pack file for this workflow", 1200, BANNER_ACCENT_ERROR)
        catch {
            TrayTip("Import", "No pack file for this workflow")
        }
        return
    }
    name := item["fileName"]
    try A_Clipboard := name
    catch {
    }
    try ShowCenteredOverlay_Utils("Copied: " . name, 1200, BANNER_ACCENT_SUCCESS)
    catch {
        TrayTip("Import", "Copied: " . name)
    }
}

; --- Thin public runners (hub API; domain parsers stay in *_import modules) ---

ImportMgmt_RunFinanceDaily(*) {
    Finance_ImportDaily()
}

ImportMgmt_RunFinanceMonthly(*) {
    Finance_ImportMonthly()
}

ImportMgmt_RunPalacePack(*) {
    Palace_ImportMnemonicsFromDesktop()
}

ImportMgmt_RunPlanPack(*) {
    Palace_ImportPlanPackFromDesktop()
}

ImportMgmt_RunTaskPack(*) {
    Task_ImportPackFromDesktop()
}

ImportMgmt_RunQuickImage(*) {
    ImportMgmt_CloseGui()
    try Palace_QuickAttachDesktopImage()
    catch as e {
        try ShowCenteredOverlay_Utils("Quick image failed: " . e.Message, 2800, BANNER_ACCENT_ERROR)
        catch {
            TrayTip("Import", "Quick image failed")
        }
    }
}

ImportMgmt_RunDesktopNames(*) {
    ImportMgmt_CloseGui()
    ClipAngelExport_ShowNamesManager(ImportMgmt_ShowMainMenu)
}

ImportMgmt_OnHelp(*) {
    ImportMgmt_ShowHelp()
}

ImportMgmt_ShowHelp() {
    global g_ImportMgmtGui, g_ImportMgmtLv
    ImportMgmt_CloseGui()

    body := ImportMgmt_HelpText()
    winW := 680
    winH := 560
    bodyW := winW - 32
    bodyH := winH - 100

    g_ImportMgmtGui := Gui("+AlwaysOnTop +ToolWindow", "Import Management — Help")
    g_ImportMgmtGui.SetFont("s10", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x16 y12 w" . bodyW, "Import rules")
    g_ImportMgmtGui.SetFont("s9", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x16 y38 w" . bodyW,
        "Canonical pack names, overwrite policy, fix-file recovery — Esc / Shift+Backspace return to list")
    edit := g_ImportMgmtGui.Add("Edit",
        "x16 y64 w" . bodyW . " h" . bodyH
        . " ReadOnly -WantReturn +VScroll Multi",
        body)
    try edit.SetFont("s10", "Consolas")
    catch {
        try edit.SetFont("s10", "Segoe UI")
        catch {
        }
    }
    g_ImportMgmtGui.SetFont("s9", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "x16 y" . (64 + bodyH + 10) . " w" . bodyW,
    "Esc / Shift+Backspace — main menu")
    g_ImportMgmtGui.OnEvent("Close", (*) => ImportMgmt_ShowMainMenu())
    g_ImportMgmtGui.OnEvent("Escape", (*) => ImportMgmt_ShowMainMenu())
    ImportMgmt_BindHotkeys([
        ["+Backspace", (*) => ImportMgmt_ShowMainMenu()],
        ["Escape", (*) => ImportMgmt_ShowMainMenu()]
    ])
    ImportMgmt_CenterGui(g_ImportMgmtGui, winW, winH)
}

ImportMgmt_HelpText() {
    return "ENTRY POINTS`r`n"
    . "#!+X  ·  Utility Shortcuts [J]  ·  Win+Alt+Shift+F double-tap`r`n"
    . "This hub is the only AHK import UI. Domain apps no longer expose import menus.`r`n`r`n"
    . "LIST KEYS`r`n"
    . "Char = run workflow   Shift+Enter / double-click = run selected   Shift+C = copy pack name`r`n"
    . "Shift+Backspace = Utility Shortcuts   Esc = close`r`n`r`n"
    . "CANONICAL DESKTOP NAMES (always overwrite)`r`n"
    . "Save AI packs with the exact filename below on Desktop.`r`n"
    . "Never add updated, corrected, v2, or similar suffixes.`r`n"
    . "Importer consolidates variants (*_updated*, gemini-code-….txt) to the canonical name.`r`n`r`n"
    . "GENERAL WORKFLOW`r`n"
    . "1. Run the pack prompt (#!+U → Prompts, or dictation flow).`r`n"
    . "2. Save the pack to Desktop (Quick Download or copy fence).`r`n"
    . "3. Open Import Management → Char / Shift+Enter on the workflow.`r`n`r`n"
    . "========== [D] FINANCE DAILY ==========`r`n"
    . "Pack: FINANCE_DAILY.txt`r`n"
    . "Writes: finances/data transactions (append)`r`n"
    . "Confirm: editable preview before save`r`n"
    . "Success: archive pack → finances/data/imported/; opens transactions; hub closes`r`n"
    . "AI fix: Desktop FINANCE_AI_FIX.txt (copied to clipboard, ≥5s banner, hub closes)`r`n"
    . "Re-run: paste fix → AI re-delivers pack → save FINANCE_DAILY.txt → #!+X → [D]`r`n`r`n"
    . "========== [M] FINANCE MONTHLY ==========`r`n"
    . "Pack: FINANCE_MONTHLY.txt`r`n"
    . "Writes: accounts / goals adjustments + related transactions`r`n"
    . "Confirm: preview before save`r`n"
    . "Success: archive → finances/data/imported/; hub closes (no Finance menu)`r`n"
    . "AI fix: FINANCE_AI_FIX.txt (same clipboard / banner / close behavior)`r`n"
    . "Re-run: #!+X → [M]`r`n`r`n"
    . "========== [P] PALACE MNEMONIC PACK ==========`r`n"
    . "Pack: PALACE_PACK.txt (or PALACE_*.txt|.csv sections)`r`n"
    . "Writes: mnemonics/data palaces / beasts / atoms (upsert + cross-link validation)`r`n"
    . "Confirm: preview before save`r`n"
    . "Success: archive → mnemonics/data/imported/; optional practice MD sync; hub closes`r`n"
    . "AI fix: Desktop PALACE_AI_FIX.txt (clipboard + ≥5s banner + hub closes)`r`n"
    . "Re-run: #!+X → [P]`r`n`r`n"
    . "========== [L] STUDY PLAN PACK ==========`r`n"
    . "Pack: PLAN_PACK.txt`r`n"
    . "Writes: study plans / plan items / resources; syncs plans Markdown`r`n"
    . "Confirm: preview before save`r`n"
    . "Success: archive → mnemonics/data/imported/; hub closes`r`n"
    . "Failure without fix file: toast; hub stays open if launched from here`r`n"
    . "Re-run: #!+X → [L]`r`n`r`n"
    . "========== [T] TASK PACK ==========`r`n"
    . "Pack: TASK_PACK.txt (prefer named TASK_PACK* over gemini-code dumps)`r`n"
    . "Prompt: Utility Prompts [k] Convert to Task (convert-to-task.txt); dictation menu [T] same pack`r`n"
    . "Filters: work | personal | habits only (not freeform categories)`r`n"
    . "Writes: tasks/data projects / tasks / info via Python task_pack_import.py (tasks append-only)`r`n"
    . "Confirm: Palace-style preview ListView before save`r`n"
    . "Success: archive → tasks/data/imported/; hub closes`r`n"
    . "AI fix: Desktop TASK_AI_FIX.txt (copied to clipboard, ≥5s banner, hub closes)`r`n"
    . "Re-run: #!+X → [T]`r`n`r`n"
    . "========== [Q] PALACE QUICK IMAGE ==========`r`n"
    . "Canonical: PALACE_QUICK_IMAGE.png (Shift+P Maps capture uses this name)`r`n"
    . "Batch captures: PALACE_QUICK_IMAGE_02.png, _03.png, … while earlier files remain on Desktop`r`n"
    . "Source: newest PALACE_QUICK_IMAGE*.png/JPG on Desktop, else any newest image`r`n"
    . "Purpose: attach that image to a Memory Palace that has no image_rel_path`r`n"
    . "Flow: pick palace from missing-image list → copy into practice/images/{study}/{n}.{ext}`r`n"
    . "Writes: mnemonics/data palaces.csv image_rel_path; syncs practice Markdown`r`n"
    . "Success: toast with path; hub already closed when picker opened`r`n"
    . "If all palaces have images, or Desktop has no image: error toast`r`n"
    . "Re-run: save palace PNG to Desktop → #!+X → [Q]`r`n`r`n"
    . "========== [N] QUICK DOWNLOAD NAMES ==========`r`n"
    . "Registry: assets/data/clipangel_desktop_names.csv`r`n"
    . "Same list as the #!+9 Quick Download rename picker (Name Desktop file).`r`n"
    . "Enter/C copies bare name (e.g. FINANCE_DAILY); A add, E edit, Delete remove`r`n"
    . "Esc / Backspace returns to this Import Management list`r`n"
    . "Also used when ClipAngel exports a clip to Desktop`r`n`r`n"
    . "OUTCOMES (Finance / Palace / Tasks)`r`n"
    . "• Full success: local CSV saved; Desktop pack archived; Import Manager closes`r`n"
    . "  (Finance daily still opens the transactions page)`r`n"
    . "• Failure with AI fix: fix text copied to clipboard; ≥5s banner; Import Manager closes`r`n"
    . "  → paste into AI companion → re-deliver pack → reopen Import Manager`r`n`r`n"
    . "AI FIX FILES (always overwrite on Desktop)`r`n"
    . "FINANCE_AI_FIX.txt | PALACE_AI_FIX.txt | TASK_AI_FIX.txt`r`n`r`n"
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

; Backward-compatible aliases (prefer ImportMgmt_Run* )
ImportMgmt_OnImportFinanceDaily(*) {
    ImportMgmt_RunFinanceDaily()
}
ImportMgmt_OnImportFinanceMonthly(*) {
    ImportMgmt_RunFinanceMonthly()
}
ImportMgmt_OnImportPalacePack(*) {
    ImportMgmt_RunPalacePack()
}
ImportMgmt_OnImportPlanPack(*) {
    ImportMgmt_RunPlanPack()
}
ImportMgmt_OnImportTaskPack(*) {
    ImportMgmt_RunTaskPack()
}
ImportMgmt_OnDesktopNames(*) {
    ImportMgmt_RunDesktopNames()
}
