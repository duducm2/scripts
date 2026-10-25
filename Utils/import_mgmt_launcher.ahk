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

; Shared outcome tones for every Import Management workflow.
; assets/sounds/import-success.wav and import-failure.wav exist only for this path.
ImportMgmt_PlayOutcomeChime(ok) {
    name := ok ? "import-success.wav" : "import-failure.wav"
    try ScriptSoundPlay(A_ScriptDir . "\assets\sounds\" . name)
    catch {
    }
}

; Copy AI fix file to clipboard, show orientation toast, close Import Manager if open.
; Sends the fix into the companion once. Automatic reply download/copy stays off,
; so this is the send that follows a manual import attempt.
; Returns true when the hub was closed.
ImportMgmt_OnAiFixReady(path, label) {
    ImportMgmt_PlayOutcomeChime(false)
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
    pipelineOwns := false
    try pipelineOwns := PackPipeline_IsActive()
    catch {
        pipelineOwns := false
    }
    if (pipelineOwns) {
        msg := "AI fix ready — pack pipeline will send to companion · Desktop " . label
        try ShowCenteredOverlay_Utils(msg, 2800, BANNER_ACCENT_ERROR)
        catch {
            TrayTip("Import", msg)
        }
    } else {
        msg := "AI fix copied — sending to AI companion · Desktop " . label
        try ShowCenteredOverlay_Utils(msg, 4000, BANNER_ACCENT_ERROR)
        catch {
            TrayTip("Import", msg)
        }
        if (body != "") {
            try ImportWatcher_CompanionPasteAndSubmitFix(body)
            catch {
            }
        }
    }
    return ImportMgmt_CloseIfOpen()
}

; Close Import Manager after a successful import from the hub. Returns true when it closed.
ImportMgmt_OnImportSuccess() {
    ImportMgmt_PlayOutcomeChime(true)
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
        Map("char", "D", "emoji", "💳", "name", "Finance daily", "detail", "FINANCE_DAILY.txt → transactions",
            "fileName", "FINANCE_DAILY.txt", "run", ImportMgmt_RunFinanceDaily),
        Map("char", "M", "emoji", "📈", "name", "Finance monthly", "detail", "FINANCE_MONTHLY.txt → accounts / goals",
            "fileName", "FINANCE_MONTHLY.txt", "run", ImportMgmt_RunFinanceMonthly),
        Map("char", "P", "emoji", "🏰", "name", "Palace mnemonic pack", "detail",
            "PALACE_PACK.txt → palaces / beasts / atoms",
            "fileName", "PALACE_PACK.txt", "run", ImportMgmt_RunPalacePack),
        Map("char", "C", "emoji", "🎒", "name", "Palace collectible", "detail",
            "COLLECTIBLE_PACK.txt → wardrobe relic",
            "fileName", "COLLECTIBLE_PACK.txt", "run", ImportMgmt_RunCollectiblePack),
        Map("char", "L", "emoji", "📋", "name", "Study plan pack", "detail", "PLAN_PACK.txt → study plans",
            "fileName", "PLAN_PACK.txt", "run", ImportMgmt_RunPlanPack),
        Map("char", "T", "emoji", "✅", "name", "Task pack", "detail", "TASK_PACK.txt → projects / tasks / info",
            "fileName", "TASK_PACK.txt", "run", ImportMgmt_RunTaskPack),
        Map("char", "Q", "emoji", "🖼️", "name", "Palace quick image", "detail",
            "Newest Desktop PNG/JPG → palace missing image",
            "fileName", "PALACE_QUICK_IMAGE.png", "run", ImportMgmt_RunQuickImage),
        Map("char", "N", "emoji", "🏷️", "name", "Quick Download names", "detail",
            "#!+9 name-first list — emoji + name — clipangel_desktop_names.csv",
            "fileName", "clipangel_desktop_names.csv", "run", ImportMgmt_RunDesktopNames),
        Map("char", "H", "emoji", "❓", "name", "Help", "detail", "Per-workflow rules and outcomes",
            "run", ImportMgmt_OnHelp)
    ]
}

ImportMgmt_DisplayName(item) {
    if (!IsObject(item))
        return ""
    name := item.Has("name") ? item["name"] : ""
    emoji := item.Has("emoji") ? Trim(item["emoji"]) : ""
    if (emoji != "")
        return emoji . " " . name
    return name
}

ImportMgmt_ShowMainMenu() {
    global g_ImportMgmtGui, g_ImportMgmtLv, g_ImportMgmtCatalog
    ImportMgmt_CloseGui()

    g_ImportMgmtCatalog := ImportMgmt_Catalog()
    contentW := 700
    lvH := 276
    guiW := 740
    guiH := 376

    g_ImportMgmtGui := Gui("+AlwaysOnTop +ToolWindow", "Import Management")
    g_ImportMgmtGui.SetFont("s10", "Segoe UI")
    g_ImportMgmtGui.Add("Text", "w" . contentW,
        "Char = import clipboard, else Desktop file   Shift+Enter / double-click = same   Shift+C = copy pack name   Shift+Backspace = utility shortcuts   Esc = close"
    )
    g_ImportMgmtLv := g_ImportMgmtGui.Add("ListView", "w" . contentW . " h" . lvH . " -Multi",
        ["Char", "Workflow", "Pack / detail"])
    g_ImportMgmtLv.OnEvent("DoubleClick", ImportMgmt_OnListActivate)
    g_ImportMgmtGui.OnEvent("Close", (*) => ImportMgmt_CloseGui())
    g_ImportMgmtGui.OnEvent("Escape", (*) => ImportMgmt_CloseGui())

    for item in g_ImportMgmtCatalog
        g_ImportMgmtLv.Add("", item["char"], ImportMgmt_DisplayName(item), item["detail"])
    try g_ImportMgmtLv.ModifyCol(1, 50)
    try g_ImportMgmtLv.ModifyCol(2, 220)
    try g_ImportMgmtLv.ModifyCol(3, 410)
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

; Catalog run callback when fileName is a pack this hub imports (exact name, e.g. FINANCE_DAILY.txt).
ImportMgmt_RunnerForFileName(fileName) {
    want := StrLower(Trim(fileName))
    if (want = "")
        return ""
    for item in ImportMgmt_Catalog() {
        if (!item.Has("fileName") || !item.Has("run"))
            continue
        if (StrLower(Trim(item["fileName"])) != want)
            continue
        runFn := item["run"]
        name := ""
        try name := runFn.Name
        catch {
            name := ""
        }
        if (name = "ImportMgmt_RunDesktopNames" || name = "ImportMgmt_OnHelp")
            continue
        return runFn
    }
    return ""
}

; --- Clipboard first, then the existing Desktop file (same idea as Quick Download names Shift+T) ---

; The chosen record is the destination, same as Quick Download names Shift+T.
; A task pack says TASK_PROJECTS, not TASK_PACK; a finance pack says FINANCE_DAILY.csv.
; Any pack body is accepted. Empty or non-pack clipboard leaves the Desktop file alone.
ImportMgmt_ClipboardPackMatches(body, fileName) {
    return fileName != "" && body != "" && ClipAngelExport_LooksLikePack(body)
}

ImportMgmt_PackHit(fileName, label, runFn) {
    return Map("fileName", fileName, "label", label, "run", runFn)
}

; FILE section title → catalog pack. Extension and quotes are ignored.
ImportMgmt_TargetForSection(sectionName) {
    key := StrLower(Trim(sectionName))
    key := RegExReplace(key, '[\s"]')
    key := RegExReplace(key, "i)\.(csv|json|txt|ini)$")
    switch key {
        case "finance_daily":
            return ImportMgmt_PackHit("FINANCE_DAILY.txt", "Finance daily", Finance_ImportDaily)
        case "finance_monthly":
            return ImportMgmt_PackHit("FINANCE_MONTHLY.txt", "Finance monthly", Finance_ImportMonthly)
        case "palace_palaces", "palace_beasts", "palace_atoms", "palace_pack":
            return ImportMgmt_PackHit("PALACE_PACK.txt", "Palace mnemonic pack", Palace_ImportMnemonicsFromDesktop)
        case "collectible", "collectible_pack":
            return ImportMgmt_PackHit("COLLECTIBLE_PACK.txt", "Palace collectible", Collectible_ImportFromDesktop)
        case "plans", "plan_items", "plan_pack":
            return ImportMgmt_PackHit("PLAN_PACK.txt", "Study plan pack", Palace_ImportPlanPackFromDesktop)
        case "task_projects", "task_tasks", "task_info", "task_pack":
            return ImportMgmt_PackHit("TASK_PACK.txt", "Task pack", Task_ImportPackFromDesktop)
    }
    return ""
}

; Names inside ===FILE: …===, ---FILE: …---, and line-start =FILE: …=.
ImportMgmt_PackSectionNames(body) {
    found := []
    seen := Map()
    patterns := [
        "i)(?:={2,}|-{2,})\s*FILE:\s*([^`r`n=]+?)\s*(?:={2,}|-{2,})",
        "im)^\s*=\s*FILE:\s*([^=`r`n]+?)\s*="
    ]
    for pat in patterns {
        pos := 1
        while (foundAt := RegExMatch(body, pat, &m, pos)) {
            name := Trim(m[1])
            key := StrLower(name)
            if (name != "" && !seen.Has(key)) {
                seen[key] := true
                found.Push(name)
            }
            pos := foundAt + m.Len(0)
        }
    }
    return found
}

; Copied companion pack → catalog hit. Map(fileName, label, run), or "" when unknown.
; Structure is the FILE section name (FINANCE_DAILY.csv, TASK_TASKS.csv, …).
; If markers exist but no section parsed, a section filename token is enough.
ImportMgmt_ClassifyPack(body) {
    if (body = "")
        return ""
    names := ImportMgmt_PackSectionNames(body)
    for name in names {
        hit := ImportMgmt_TargetForSection(name)
        if (IsObject(hit))
            return hit
    }
    if (names.Length || !ClipAngelExport_LooksLikePack(body))
        return ""
    lower := StrLower(body)
    tokens := [
        ["collectible.json", "collectible"],
        ["finance_daily.csv", "finance_daily"],
        ["finance_monthly.csv", "finance_monthly"],
        ["task_projects.csv", "task_projects"],
        ["task_tasks.csv", "task_tasks"],
        ["task_info.csv", "task_info"],
        ["plan_items.csv", "plan_items"],
        ["plans.csv", "plans"],
        ["palace_palaces.csv", "palace_palaces"],
        ["palace_beasts.csv", "palace_beasts"],
        ["palace_atoms.csv", "palace_atoms"]
    ]
    for pair in tokens {
        if (InStr(lower, pair[1]))
            return ImportMgmt_TargetForSection(pair[2])
    }
    return ""
}

; #!+P [I]: clipboard was just copied. Identify the pack and import that system only.
; Does not fall back to an older Desktop file when this copy is not a known pack.
ImportMgmt_RunClassifiedClipboard() {
    errMsg := ""
    body := ""
    try body := ClipAngelExport_ClipboardAsPackText(&errMsg)
    catch {
        body := ""
        errMsg := "Clipboard unavailable"
    }
    if (body = "") {
        msg := errMsg != "" ? errMsg : "Clipboard empty"
        ImportMgmt_PlayOutcomeChime(false)
        try ShowCenteredOverlay_Utils("❌ " msg, 2500, BANNER_ACCENT_ERROR)
        catch {
            TrayTip("Import", msg)
        }
        return false
    }
    hit := ImportMgmt_ClassifyPack(body)
    if (!IsObject(hit)) {
        ImportMgmt_PlayOutcomeChime(false)
        try ShowCenteredOverlay_Utils("❌ No import matches this copy", 2800, BANNER_ACCENT_ERROR)
        catch {
            TrayTip("Import", "No import matches this copy")
        }
        return false
    }
    placeReason := ""
    placed := false
    try placed := ImportMgmt_TryPlaceClipboardPack(hit["fileName"], &placeReason)
    catch {
        placed := false
        if (placeReason = "")
            placeReason := "Could not write pack to Desktop"
    }
    if (!placed) {
        msg := placeReason != "" ? placeReason : "Could not write pack to Desktop"
        ImportMgmt_PlayOutcomeChime(false)
        try ShowCenteredOverlay_Utils("❌ " msg, 2800, BANNER_ACCENT_ERROR)
        catch {
            TrayTip("Import", msg)
        }
        return false
    }
    try ShowCenteredOverlay_Utils("📥 " hit["label"], 1600, BANNER_ACCENT_INTERMEDIATE)
    catch {
    }
    result := false
    try result := hit["run"].Call()
    catch as e {
        ImportMgmt_PlayOutcomeChime(false)
        try ShowCenteredOverlay_Utils("❌ Import failed: " (e.Message ? e.Message : "unknown"), 2800,
        BANNER_ACCENT_ERROR)
        catch {
        }
        result := false
    }
    try ImportMgmt_CloseIfOpen()
    catch {
    }
    return result
}

; True when Desktop already has this pack (exact name or baseName*.txt|.csv) on search dirs.
ImportMgmt_DesktopPackAvailable(fileName) {
    if (fileName = "")
        return false
    try {
        if (ClipAngelExport_SameNameDesktopPaths(fileName).Length)
            return true
    } catch {
    }
    baseName := fileName
    if (RegExMatch(baseName, "i)\.(txt|csv)$"))
        baseName := RegExReplace(baseName, "i)\.(txt|csv)$", "")
    if (baseName = "")
        return false
    dirs := []
    try dirs := ClipAngelExport_DesktopSearchDirs()
    catch {
        dirs := []
    }
    if (!dirs.Length && DirExist(A_Desktop))
        dirs := [A_Desktop]
    for dir in dirs {
        for ext in ["csv", "txt"] {
            loop files dir . "\" . baseName . "*." . ext, "F"
                return true
        }
    }
    return false
}

; Write a matching clipboard pack to the canonical Desktop name.
; True = written. On false, reasonOut explains why (empty / not a pack / write failed).
ImportMgmt_TryPlaceClipboardPack(fileName, &reasonOut := "") {
    reasonOut := ""
    if (fileName = "") {
        reasonOut := "No pack file name"
        return false
    }
    errMsg := ""
    body := ""
    try body := ClipAngelExport_ClipboardAsPackText(&errMsg)
    catch {
        reasonOut := "Clipboard unavailable"
        return false
    }
    if (body = "") {
        reasonOut := (errMsg != "") ? errMsg : "Clipboard empty"
        return false
    }
    if (!ImportMgmt_ClipboardPackMatches(body, fileName)) {
        reasonOut := "Clipboard is not a pack"
        return false
    }
    desktopPath := ""
    try desktopPath := AiQuickDownload_ResolveDesktopPath()
    catch {
        desktopPath := ""
    }
    if (desktopPath = "" || !DirExist(desktopPath))
        desktopPath := A_Desktop
    if (desktopPath = "" || !DirExist(desktopPath)) {
        reasonOut := "Desktop folder not found"
        return false
    }
    matches := ClipAngelExport_SameNameDesktopPaths(fileName)
    outPath := (matches.Length = 1) ? matches[1] : (desktopPath . "\" . fileName)
    try ClipAngelExport_OverwriteUtf8(outPath, body)
    catch {
        reasonOut := "Could not write pack to Desktop"
        return false
    }
    if !FileExist(outPath) {
        reasonOut := "Could not write pack to Desktop"
        return false
    }
    try outPath := ClipAngelExport_PlaceCanonicalPack(outPath, fileName)
    catch {
    }
    try ImportWatcher_ClaimPath(outPath)
    catch {
    }
    if (FileExist(outPath))
        return true
    reasonOut := "Could not write pack to Desktop"
    return false
}

; Clipboard image → next PALACE_QUICK_IMAGE*.png. False keeps the current Desktop image.
ImportMgmt_TryPlaceClipboardImage() {
    hasImage := false
    try hasImage := Task_ClipboardHasImage()
    catch {
        hasImage := false
    }
    if (!hasImage)
        return false
    dest := ""
    try dest := Palace_DesktopNextQuickImagePath()
    catch {
        dest := ""
    }
    if (dest = "")
        return false
    ok := false
    try ok := Task_SaveClipboardImage(dest)
    catch {
        ok := false
    }
    return ok && FileExist(dest)
}

; Place clipboard pack when present, else use existing Desktop file; toast if neither is available.
; The hub closes after the attempt, whether the import saved, failed, or the confirm was cancelled.
ImportMgmt_RunPreferClipboard(fileName, runFn) {
    placed := false
    placeReason := ""
    try placed := ImportMgmt_TryPlaceClipboardPack(fileName, &placeReason)
    catch {
        placed := false
        if (placeReason = "")
            placeReason := "Clipboard pack place failed"
    }
    if (!placed && !ImportMgmt_DesktopPackAvailable(fileName)) {
        msg := (placeReason != "") ? placeReason : "Clipboard empty"
        ImportMgmt_PlayOutcomeChime(false)
        try ShowCenteredOverlay_Utils(msg, 2800, BANNER_ACCENT_ERROR)
        catch {
            TrayTip("Import", msg)
        }
        ImportMgmt_CloseIfOpen()
        return false
    }
    result := false
    try result := runFn.Call()
    catch {
        ImportMgmt_PlayOutcomeChime(false)
        result := false
    }
    ImportMgmt_CloseIfOpen()
    return result
}

; --- Thin public runners (hub API; domain parsers stay in *_import modules) ---

ImportMgmt_RunFinanceDaily(*) {
    ImportMgmt_RunPreferClipboard("FINANCE_DAILY.txt", Finance_ImportDaily)
}

ImportMgmt_RunFinanceMonthly(*) {
    ImportMgmt_RunPreferClipboard("FINANCE_MONTHLY.txt", Finance_ImportMonthly)
}

ImportMgmt_RunPalacePack(*) {
    ImportMgmt_RunPreferClipboard("PALACE_PACK.txt", Palace_ImportMnemonicsFromDesktop)
}

ImportMgmt_RunCollectiblePack(*) {
    ImportMgmt_RunPreferClipboard("COLLECTIBLE_PACK.txt", Collectible_ImportFromDesktop)
}

ImportMgmt_RunPlanPack(*) {
    ImportMgmt_RunPreferClipboard("PLAN_PACK.txt", Palace_ImportPlanPackFromDesktop)
}

ImportMgmt_RunTaskPack(*) {
    ImportMgmt_RunPreferClipboard("TASK_PACK.txt", Task_ImportPackFromDesktop)
}

ImportMgmt_RunQuickImage(*) {
    ImportMgmt_CloseGui()
    try ImportMgmt_TryPlaceClipboardImage()
    catch {
    }
    try Palace_QuickAttachDesktopImage()
    catch as e {
        try ImportMgmt_PlayOutcomeChime(false)
        catch {
        }
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
    . "#!+P copy menu [I] copies the companion reply, reads its FILE sections, and runs that import.`r`n"
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
    . "2. Copy the pack, or save it on Desktop under the canonical name.`r`n"
    . "3. Open Import Management → Char / Shift+Enter on the workflow.`r`n"
    .
    "Clipboard is tried first (same unwrap as Quick Download names Shift+T). If it is not this pack, the Desktop file is used.`r`n"
    . "A successful import plays import-success.wav. A rejected import plays import-failure.wav. Cancelling the confirm dialog is silent.`r`n`r`n"
    . "========== [D] FINANCE DAILY ==========`r`n"
    . "Pack: FINANCE_DAILY.txt`r`n"
    . "Writes: finances/data transactions (append)`r`n"
    . "Confirm: editable preview before save`r`n"
    . "Success: archive pack → finances/data/imported/; opens transactions; hub closes`r`n"
    . "AI fix: Desktop FINANCE_AI_FIX.txt — sent to the companion; hub closes`r`n"
    . "Re-run: save the corrected FINANCE_DAILY.txt yourself → #!+X → [D]`r`n`r`n"
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
    . "========== [C] PALACE COLLECTIBLE ==========`r`n"
    . "Pack: COLLECTIBLE_PACK.txt (===FILE: COLLECTIBLE.json===)`r`n"
    . "Picture: the newest PNG, JPG, or WEBP on the Desktop. The JSON does not name it.`r`n"
    . "Not watched. A file on the Desktop imports only when you press [C].`r`n"
    .
    "Writes: holds the relic on the avatar. Save in the wardrobe stores it in mnemonics/data/repl/collectibles.json`r`n"
    . "Success: archive pack and picture → mnemonics/data/imported/; refreshes the Memory Palace tab`r`n"
    . "AI fix: Desktop COLLECTIBLE_AI_FIX.txt (clipboard + ≥5s banner + hub closes)`r`n"
    . "Re-run: #!+X → [C]`r`n`r`n"
    . "========== [L] STUDY PLAN PACK ==========`r`n"
    . "Pack: PLAN_PACK.txt`r`n"
    . "Writes: study plans / plan items / resources; syncs plans Markdown`r`n"
    . "Confirm: preview before save`r`n"
    . "Success: archive → mnemonics/data/imported/; hub closes`r`n"
    . "Failure or cancel: toast; hub closes`r`n"
    . "Re-run: #!+X → [L]`r`n`r`n"
    . "========== [T] TASK PACK ==========`r`n"
    . "Pack: TASK_PACK.txt (prefer named TASK_PACK* over gemini-code dumps)`r`n"
    . "Prompt: Utility Prompts [k] Convert to Task (convert-to-task.txt); dictation menu [T] same pack`r`n"
    . "Filters: work | personal | habits only (not freeform categories)`r`n"
    . "Writes: tasks/data projects / tasks / info via Python task_pack_import.py (tasks append-only)`r`n"
    . "Confirm: preview ListView before save. Ctrl+click a task to exclude it and its info points.`r`n"
    . "Success: archive → tasks/data/imported/; hub closes`r`n"
    . "AI fix: Desktop TASK_AI_FIX.txt (copied to clipboard, ≥5s banner, hub closes)`r`n"
    . "Re-run: #!+X → [T]`r`n`r`n"
    . "========== [Q] PALACE QUICK IMAGE ==========`r`n"
    . "Canonical: PALACE_QUICK_IMAGE.png (Shift+P Maps capture uses this name)`r`n"
    . "Batch captures: PALACE_QUICK_IMAGE_02.png, _03.png, … while earlier files remain on Desktop`r`n"
    .
    "Source: clipboard image if one is copied, else newest PALACE_QUICK_IMAGE*.png/JPG on Desktop, else any newest image`r`n"
    . "Purpose: attach that image to a Memory Palace that has no image_rel_path`r`n"
    . "Flow: pick palace from missing-image list → copy into practice/images/{study}/{n}.{ext}`r`n"
    . "Writes: mnemonics/data palaces.csv image_rel_path; syncs practice Markdown`r`n"
    . "Success: toast with path; hub already closed when picker opened`r`n"
    . "If all palaces have images, or Desktop has no image: error toast`r`n"
    . "Re-run: save palace PNG to Desktop → #!+X → [Q]`r`n`r`n"
    . "========== [N] QUICK DOWNLOAD NAMES ==========`r`n"
    . "Registry: assets/data/clipangel_desktop_names.csv (id, emoji, name)`r`n"
    . "Same list as the #!+9 Quick Download name picker (shown before download).`r`n"
    .
    "Each row shows emoji + name (same pattern as this hub). Pack defaults: 💳 daily · 📈 monthly · 🏰 palace · ✅ task · 🖼️ quick image.`r`n"
    .
    "Rename picker: letter jumps list; Enter = apply with current/auto ext; Shift+T = save as .txt;`r`n"
    .
    "Shift+X = change extension (then Enter); Shift+N = type a one-off name; clipboard-first when image/text on clipboard`r`n"
    .
    "Names manager: Enter copies bare name; Shift+C copies and closes every open window; Shift+T writes the clipboard as Desktop\{name}.txt (overwrite; a markdown code fence around a pack is unwrapped), runs that pack's import when the name matches one (FINANCE_DAILY, PALACE_PACK, …), and closes this list without returning here; Shift+A add, Shift+E edit (emoji + name), Delete remove`r`n"
    . "Esc / Backspace returns to this Import Management list`r`n"
    . "Also used when ClipAngel exports a clip to Desktop`r`n`r`n"
    . "OUTCOMES (Finance / Palace / Tasks)`r`n"
    . "• Full success: local CSV saved; Desktop pack archived; Import Manager closes`r`n"
    . "  (Finance daily still opens the transactions page)`r`n"
    . "• Failure with AI fix: fix is sent to the AI companion; Import Manager closes`r`n"
    . "  → save the corrected pack yourself → reopen Import Manager`r`n`r`n"
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
