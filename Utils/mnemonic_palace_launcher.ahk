; =============================================================================
; Utils module: mnemonic_palace_launcher.ahk
; Memory Palace main menu (Utility Shortcuts [N])
; =============================================================================

global g_PalaceDashboardHwnd := 0

Palace_LaunchApp() {
    try Finance_CloseGui()
    try Task_CloseGui()
    try Task_CloseWebApp()
    Palace_EnsureData()
    Palace_ShowMainMenu()
}

Palace_ShowMainMenu() {
    global g_PalaceGui
    Palace_CloseGui()
    Palace_EnsureData()

    studies := Palace_Load("studies")
    palaces := Palace_Load("palaces")
    beasts := Palace_Load("beasts")
    atoms := Palace_Load("atoms")
    plans := Palace_Load("plans")

    g_PalaceGui := Gui("+AlwaysOnTop +ToolWindow", "Memory Palace")
    g_PalaceGui.SetFont("s10", "Segoe UI")
    g_PalaceGui.BackColor := "1E1E1E"
    g_PalaceGui.OnEvent("Close", (*) => Palace_CloseGui())
    g_PalaceGui.OnEvent("Escape", (*) => Palace_CloseGui())

    g_PalaceGui.SetFont("s16 cWhite Bold", "Segoe UI")
    g_PalaceGui.Add("Text", "x20 y16 w880", "Memory Palace")
    g_PalaceGui.SetFont("s10 cC0C0C0 Norm", "Segoe UI")
    g_PalaceGui.Add("Text", "x20 y48 w880",
        studies.Length . " studies  ·  " . palaces.Length . " palaces  ·  "
        . beasts.Length . " beasts  ·  " . atoms.Length . " atoms  ·  "
        . plans.Length . " plans")
    g_PalaceGui.SetFont("s9 cF1C40F", "Segoe UI")
    g_PalaceGui.Add("Text", "x20 y72 w880", "1-3 quick links · letters open a module.")

    items := [
        ["1", "📹 Study Video", "Open / set video link (Google Docs API)"],
        ["2", "📖 Study Article", "Open / set article link (Google Docs API)"],
        ["3", "❤️ Favorite", "Open / set favorite link (Google Docs API)"],
        ["D", "Dashboard", "Study picker and Memory Palace images"],
        ["B", "Browse", "Studies -> palaces -> beasts -> atoms"],
        ["L", "Plans", "Browse / edit study plan checklists (CSV)"],
        ["I", "AI import", "Desktop PALACE_PACK / PALACE_*.txt|.csv (preview)"],
        ["J", "Import plan", "Desktop PLAN_PACK only → sync output/plans"],
        ["Q", "Quick image", "Pick palace without image → newest Desktop PNG/JPG"],
        ["G", "Practice on GitHub", "Synced palace practice notes for mobile"],
        ["O", "Plans on GitHub", "Synced study plan Markdown for mobile"],
        ["R", "Regen Markdown", "Force-create all practice + plan .md files"],
        ["H", "Help", "Vocabulary and mapping rules"],
        ["P", "Push to cloud", "Commit and push scripts repo"]
    ]

    x0 := 20
    y0 := 108
    colW := 440
    rowH := 72
    idx := 0
    for it in items {
        col := Mod(idx, 2)
        row := idx // 2
        x := x0 + col * colW
        y := y0 + row * rowH
        g_PalaceGui.SetFont("s14 cF1C40F Bold", "Segoe UI")
        g_PalaceGui.Add("Text", "x" . (x + 12) . " y" . (y + 8) . " w40 BackgroundTrans", "[" . it[1] . "]")
        g_PalaceGui.SetFont("s12 cWhite Bold", "Segoe UI")
        g_PalaceGui.Add("Text", "x" . (x + 58) . " y" . (y + 8) . " w340 BackgroundTrans", it[2])
        g_PalaceGui.SetFont("s9 cA0A0A0 Norm", "Segoe UI")
        g_PalaceGui.Add("Text", "x" . (x + 58) . " y" . (y + 32) . " w340 BackgroundTrans", it[3])
        idx += 1
    }

    g_PalaceGui.SetFont("s9 c808080", "Segoe UI")
    g_PalaceGui.Add("Text", "x20 y620 w880",
        "Backspace utility shortcuts   Esc close   L plans   J import plan   I mnemonic pack   O plans GitHub   P push   R regen"
    )

    Palace_BindHotkeys([
        ["1", Palace_OnStudyVideo], ["2", Palace_OnStudyArticle], ["3", Palace_OnStudyFavorite],
        ["d", Palace_OnDash], ["b", Palace_OnBrowse], ["l", Palace_OnPlans],
        ["i", Palace_OnImp], ["j", Palace_OnImportPlan],
        ["q", Palace_OnQuickImage], ["g", Palace_OnPracticeGithub],
        ["o", Palace_OnPlansGithub], ["r", Palace_OnRegenMarkdown],
        ["h", Palace_OnHelp], ["p", Palace_OnGitPush],
        ["Backspace", (*) => Palace_ReturnToUtilityShortcuts()],
        ["Escape", (*) => Palace_CloseGui()]
    ])
    Palace_CenterGui(g_PalaceGui, 900, 660)
}

Palace_ReturnToUtilityShortcuts() {
    Palace_CloseGui()
    try ShowHotstringSelector()
    catch {
    }
}

Palace_OnDash(*) {
    Palace_OpenDashboard()
}
Palace_OnStudyVideo(*) {
    Palace_CloseGui()
    StudyTopicSelector_ManageLinks()
}
Palace_OnStudyArticle(*) {
    Palace_CloseGui()
    StudyTopicSelector_ManageArticleLinks()
}
Palace_OnStudyFavorite(*) {
    Palace_CloseGui()
    StudyTopicSelector_ManageFavoriteLinks()
}
Palace_OnBrowse(*) {
    Palace_ShowBrowse()
}
Palace_OnPlans(*) {
    global g_PalaceFilterStudyId, g_PalaceFilterPlanId, g_PalaceFilterPalaceId, g_PalaceFilterBeastId
    Palace_EnsureData()
    pick := Palace_PickStudy()
    if (pick = "") {
        Palace_ShowMainMenu()
        return
    }
    g_PalaceFilterStudyId := pick
    g_PalaceFilterPalaceId := ""
    g_PalaceFilterBeastId := ""
    g_PalaceFilterPlanId := ""
    Palace_ShowPlans()
}
Palace_OnImp(*) {
    Palace_ImportMnemonicsFromDesktop()
}
Palace_OnImportPlan(*) {
    Palace_ImportPlanPackFromDesktop()
}
Palace_OnQuickImage(*) {
    Palace_QuickAttachDesktopImage()
}
Palace_OnPracticeGithub(*) {
    Palace_CloseGui()
    Palace_OpenPracticeGithub()
}
Palace_OnRegenMarkdown(*) {
    Palace_ForceRegenAllMarkdown()
}
Palace_OnHelp(*) {
    Palace_ShowHelp()
}
Palace_OnGitPush(*) {
    Palace_GitSyncPush()
}

Palace_PracticeGithubUrl() {
    return "https://github.com/duducm2/scripts/tree/main/mnemonics/output/practice"
}

Palace_OpenPracticeGithub() {
    url := Palace_PracticeGithubUrl()
    try Run('chrome.exe --new-window "' . url . '"')
    catch as e {
        try Run('"' . url . '"')
        catch {
            Palace_Notify("Could not open GitHub: " . e.Message, 2500, BANNER_ACCENT_ERROR)
            return
        }
    }
    Palace_Notify("Practice folder on GitHub", 1800, BANNER_ACCENT_SUCCESS)
}

Palace_PlansGithubUrl() {
    return "https://github.com/duducm2/scripts/tree/main/mnemonics/output/plans"
}

Palace_OnPlansGithub(*) {
    Palace_CloseGui()
    Palace_OpenPlansGithub()
}

Palace_OpenPlansGithub() {
    url := Palace_PlansGithubUrl()
    try Run('chrome.exe --new-window "' . url . '"')
    catch as e {
        try Run('"' . url . '"')
        catch {
            Palace_Notify("Could not open GitHub: " . e.Message, 2500, BANNER_ACCENT_ERROR)
            return
        }
    }
    Palace_Notify("Plans folder on GitHub", 1800, BANNER_ACCENT_SUCCESS)
}

Palace_PlanSaveServerPort() {
    return 8765
}

; Kill whatever is listening on the plan-save port so [D] always loads fresh Python.
Palace_StopPlanSaveServer(port := 0) {
    if (port = 0)
        port := Palace_PlanSaveServerPort()
    ; netstat+taskkill — kill every LISTENING PID (stale servers can stack).
    q := Chr(39)
    cmd := "for /f `"tokens=5`" %a in (" . q . "netstat -ano ^| findstr :" . port
        . " ^| findstr LISTENING" . q . ") do taskkill /F /PID %a >nul 2>&1"
    loop 3 {
        try RunWait(A_ComSpec . " /c " . cmd, , "Hide")
        catch {
        }
        Sleep 200
        if (!Palace_IsPlanSaveServerRunning(port))
            break
    }
    loop 15 {
        if (!Palace_IsPlanSaveServerRunning(port))
            return
        Sleep 100
    }
}

Palace_EnsurePlanSaveServer(forceRestart := false) {
    port := Palace_PlanSaveServerPort()
    if (forceRestart)
        Palace_StopPlanSaveServer(port)
    else if (Palace_IsPlanSaveServerRunning(port))
        return true
    py := Palace_PythonDir() . "\plan_save_server.py"
    if (!FileExist(py)) {
        Palace_Notify("plan_save_server.py not found", 2200, BANNER_ACCENT_ERROR)
        return false
    }
    pyCmd := Palace_FindPythonCmd()
    if (pyCmd = "") {
        Palace_Notify("Python not found for plan save server", 2500, BANNER_ACCENT_ERROR)
        return false
    }
    dataDir := Palace_DataDir()
    outDir := Palace_OutputDir()
    notesRoot := Palace_NotesStudiesRoot()
    cmd := pyCmd . ' "' . py . '" --data-dir "' . dataDir . '" --output-dir "' . outDir
        . '" --port ' . port
    if (notesRoot != "")
        cmd .= ' --studies-root "' . notesRoot . '"'
    try Run(A_ComSpec . ' /c ' . cmd, A_ScriptDir, "Hide")
    catch as e {
        Palace_Notify("Plan save server failed: " . e.Message, 2800, BANNER_ACCENT_ERROR)
        return false
    }
    loop 20 {
        if (Palace_IsPlanSaveServerRunning(port))
            return true
        Sleep 150
    }
    Palace_Notify("Plan save server did not start", 2800, BANNER_ACCENT_ERROR)
    return false
}

Palace_IsPlanSaveServerRunning(port := 0) {
    if (port = 0)
        port := Palace_PlanSaveServerPort()
    try {
        whr := ComObject("WinHttp.WinHttpRequest.5.1")
        whr.Open("GET", "http://127.0.0.1:" . port . "/health", false)
        whr.Send()
        return (whr.Status = 200)
    } catch {
        return false
    }
}

; Idempotent: import *-plan.md into CSV for studies that have no plans.csv row yet.
; Does not use --force, so existing CSV edits are never overwritten.
Palace_MigratePlansToCsv(showUi := false) {
    py := Palace_PythonDir() . "\migrate_plans_to_csv.py"
    if (!FileExist(py))
        return false
    pyCmd := Palace_FindPythonCmd()
    if (pyCmd = "") {
        if (showUi)
            Palace_Notify("Python not found for plan migration", 2500, BANNER_ACCENT_ERROR)
        return false
    }
    dataDir := Palace_DataDir()
    notesRoot := Palace_NotesStudiesRoot()
    cmd := pyCmd . ' "' . py . '" --data-dir "' . dataDir . '"'
    if (notesRoot != "")
        cmd .= ' --studies-root "' . notesRoot . '"'
    exitCode := 0
    try {
        exitCode := RunWait(A_ComSpec . ' /c ' . cmd, A_ScriptDir, "Hide")
    } catch as e {
        if (showUi)
            Palace_Notify("Plan migration failed: " . e.Message, 2800, BANNER_ACCENT_ERROR)
        return false
    }
    if (exitCode != 0) {
        if (showUi)
            Palace_Notify("Plan migration failed (exit " . exitCode . ")", 2800, BANNER_ACCENT_ERROR)
        return false
    }
    return true
}

Palace_OpenDashboard() {
    try StandardLoadingBar_Show("⏳ Opening Memory Palace dashboard…", BANNER_ACCENT_INTERMEDIATE, {
        passive: false
    })
    catch {
    }
    Palace_EnsureData()
    try StandardLoadingBar_Update("⏳ Migrating plan Markdown into CSV…", BANNER_ACCENT_INTERMEDIATE)
    catch {
    }
    ; Fill empty plans.csv from *-plan.md so backlog Add/Save can find study plans.
    ; Failure is non-blocking — dashboard still opens from MD fallback if needed.
    Palace_MigratePlansToCsv(true)
    try StandardLoadingBar_Update("⏳ Restarting plan save server…", BANNER_ACCENT_INTERMEDIATE)
    catch {
    }
    ; Restart so save-server code (e.g. add_backlog) is never stale after edits.
    Palace_EnsurePlanSaveServer(true)
    py := Palace_PythonDir() . "\chart_generator.py"
    if (!FileExist(py)) {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        Palace_Notify("chart_generator.py not found", 2000, BANNER_ACCENT_ERROR)
        return
    }
    dataDir := Palace_DataDir()
    outDir := Palace_OutputDir()
    notesRoot := Palace_NotesStudiesRoot()
    pyCmd := Palace_FindPythonCmd()
    if (pyCmd = "") {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        Palace_Notify("Python not found. Install Python or enable the py launcher.", 3500, BANNER_ACCENT_ERROR)
        return
    }
    try StandardLoadingBar_Update("⏳ Syncing technique + plans + building dashboard…", BANNER_ACCENT_INTERMEDIATE)
    catch {
    }
    cmd := pyCmd . ' "' . py . '" --data-dir "' . dataDir . '" --output-dir "' . outDir . '"'
    if (notesRoot != "") {
        cmd .= ' --notes-root "' . notesRoot . '"'
        cmd .= ' --studies-root "' . notesRoot . '"'
    }
    lastStudy := Trim(Palace_Setting("General", "LastStudyId", ""))
    if (lastStudy != "")
        cmd .= ' --study-id "' . lastStudy . '"'
    exitCode := 0
    try {
        exitCode := RunWait(A_ComSpec . ' /c ' . cmd, A_ScriptDir, "Hide")
    } catch as e {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        Palace_Notify("Python failed: " . e.Message, 2500, BANNER_ACCENT_ERROR)
        return
    }
    if (exitCode != 0) {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        Palace_Notify("Dashboard Python failed (exit " . exitCode . ")", 3500, BANNER_ACCENT_ERROR)
        return
    }
    html := outDir . "\dashboard.html"
    if (!FileExist(html)) {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        Palace_Notify("dashboard.html was not generated", 2200, BANNER_ACCENT_ERROR)
        return
    }
    try StandardLoadingBar_Update("⏳ Opening Chrome…", BANNER_ACCENT_INTERMEDIATE)
    catch {
    }
    tmpHtml := A_Temp . "\palace_dashboard.html"
    try FileCopy(html, tmpHtml, 1)
    catch {
        tmpHtml := html
    }
    fileUrl := "file:///" . StrReplace(StrReplace(tmpHtml, "\", "/"), " ", "%20") . "?t=" . A_TickCount
    if (!Palace_OpenDashboardInChrome(fileUrl)) {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        Palace_Notify("Chrome failed to open dashboard", 2500, BANNER_ACCENT_ERROR)
        return
    }
    try StandardLoadingBar_Hide(400)
    catch {
    }
    Palace_CloseGui()
}

; Typed contract: positive HWND if chrome.exe window still exists, else 0.
Palace_DashboardHwndValid(hwnd) {
    try hwnd := Integer(hwnd)
    catch {
        return 0
    }
    if (!(hwnd is Integer) || hwnd <= 0)
        return 0
    if (!WinExist("ahk_id " hwnd))
        return 0
    try {
        if (WinGetProcessName("ahk_id " hwnd) != "chrome.exe")
            return 0
    } catch {
        return 0
    }
    return hwnd
}

Palace_DashboardHwndCacheGet() {
    global g_PalaceDashboardHwnd
    hwnd := Palace_DashboardHwndValid(g_PalaceDashboardHwnd)
    if (hwnd)
        return hwnd
    raw := Trim(Palace_Setting("General", "DashboardChromeHwnd", ""))
    if (raw = "" || raw = "0")
        return 0
    hwnd := Palace_DashboardHwndValid(raw)
    if (hwnd) {
        g_PalaceDashboardHwnd := hwnd
        return hwnd
    }
    return 0
}

Palace_DashboardHwndCacheSet(hwnd) {
    global g_PalaceDashboardHwnd
    hwnd := Palace_DashboardHwndValid(hwnd)
    g_PalaceDashboardHwnd := hwnd
    Palace_SetSetting("General", "DashboardChromeHwnd", hwnd ? String(hwnd) : "")
    return hwnd
}

Palace_DashboardHwndCacheClear() {
    Palace_DashboardHwndCacheSet(0)
}

; Dedicated Memory Palace Chrome window title (not Google Search, etc.).
; Includes SPA atom titles like "Memory Palace 3: AI Pricing & Elo Mechanics".
Palace_IsDashboardChromeWindowTitle(title) {
    t := Trim(String(title))
    if (t = "")
        return false
    if (InStr(t, "Google Search") || InStr(t, " - Search") || InStr(t, "Search - "))
        return false
    ; Starts with "Memory Palace" (home, atom view, or "… - Google Chrome").
    if (RegExMatch(t, "i)^Memory Palace(\b|$)"))
        return true
    return false
}

; True if the active document URL is our stable TEMP dashboard file.
Palace_DashboardWindowShowsDashboardFile(hwnd) {
    try {
        root := UIA.ElementFromHandle(hwnd)
        if (!root)
            return false
        doc := root.FindFirst({ Type: "Document" })
        if (!doc)
            return false
        return InStr(String(doc.Value), "palace_dashboard.html") > 0
    } catch {
        return false
    }
}

; Real Chrome omnibox only — never the first page Edit (Notes textarea).
Palace_DashboardFindOmnibox(hwnd) {
    try {
        root := UIA.ElementFromHandle(hwnd)
        if (!root)
            return 0
        for criteria in [{ Type: "Edit", AcceleratorKey: "Ctrl+L" }, { Type: "Edit", Name: "Address and search bar" }, { Type: "Edit",
            AutomationId: "view_1012" }] {
            try {
                el := root.FindFirst(criteria)
                if (el)
                    return el
            } catch {
            }
        }
    } catch {
    }
    return 0
}

; Set omnibox Value + Enter. Does not use UIA_Browser.SetURL (mistargets Notes).
Palace_DashboardNavigateViaOmnibox(hwnd, fileUrl) {
    omnibox := Palace_DashboardFindOmnibox(hwnd)
    if (!omnibox)
        return false
    try {
        omnibox.SetFocus()
        Sleep(40)
        try omnibox.ValuePattern.SetValue(fileUrl)
        catch {
            try omnibox.Value := fileUrl
            catch {
                return false
            }
        }
        Sleep(40)
        if (!InStr(String(omnibox.Value), "palace_dashboard.html"))
            return false
        ControlSend("{Enter}", , "ahk_id " hwnd)
        Sleep(300)
        return true
    } catch {
        return false
    }
}

; Ctrl+L → paste → Enter (after WinActivate). Avoids typing into Notes.
Palace_DashboardNavigateViaClipboard(hwnd, fileUrl) {
    clipSaved := ClipboardAll()
    try {
        A_Clipboard := fileUrl
        if (!ClipWait(1))
            return false
        Send "{LWin Up}{RWin Up}{LAlt Up}{RAlt Up}{LShift Up}{RShift Up}"
        Sleep(40)
        Send "^l"
        Sleep(100)
        Send "^a^v"
        Sleep(80)
        Send "{Enter}"
        Sleep(300)
        return true
    } catch {
        return false
    } finally {
        try A_Clipboard := clipSaved
        catch {
        }
    }
}

; Activate hwnd and load fileUrl. Returns true on success.
; Never call UIA_Browser.Navigate/SetURL — those can Write into the Notes Edit
; and ControlSend Ctrl+L, which leaves a stray "l" in the textarea.
Palace_DashboardNavigate(hwnd, fileUrl) {
    hwnd := Palace_DashboardHwndValid(hwnd)
    if (!hwnd)
        return false
    try {
        WinActivate("ahk_id " hwnd)
        if (!WinWaitActive("ahk_id " hwnd, , 2))
            return false
        ; Already on dashboard file (file overwritten): F5 refresh, no omnibox.
        if (Palace_DashboardWindowShowsDashboardFile(hwnd)) {
            ControlSend("{F5}", , "ahk_id " hwnd)
            Sleep(400)
            Palace_DashboardHwndCacheSet(hwnd)
            return true
        }
        if (Palace_DashboardNavigateViaOmnibox(hwnd, fileUrl)
        || Palace_DashboardNavigateViaClipboard(hwnd, fileUrl)) {
            Palace_DashboardHwndCacheSet(hwnd)
            return true
        }
        return false
    } catch {
        return false
    }
}

; Cache-first open/refresh. Avoids full UIA tab scans. Returns true/false.
Palace_OpenDashboardInChrome(fileUrl) {
    ; Fast path: cached HWND.
    hwnd := Palace_DashboardHwndCacheGet()
    if (hwnd) {
        if (Palace_DashboardNavigate(hwnd, fileUrl))
            return true
        Palace_DashboardHwndCacheClear()
    }

    ; Miss path: lightweight Win32 title scan (no per-window tab UIA).
    candidates := []
    for h in WinGetList("ahk_exe chrome.exe") {
        try title := WinGetTitle("ahk_id " h)
        catch {
            continue
        }
        if (Palace_IsDashboardChromeWindowTitle(title))
            candidates.Push(h)
    }
    if (candidates.Length) {
        keeper := candidates[1]
        if (Palace_DashboardNavigate(keeper, fileUrl)) {
            ; Close other dedicated dashboard windows only.
            loop candidates.Length {
                if (A_Index = 1)
                    continue
                other := candidates[A_Index]
                try {
                    otherTitle := WinGetTitle("ahk_id " other)
                    if (otherTitle = "Memory Palace" || otherTitle = "Memory Palace - Google Chrome")
                        WinClose("ahk_id " other)
                } catch {
                }
            }
            try WinActivate("ahk_id " keeper)
            catch {
            }
            return true
        }
        Palace_DashboardHwndCacheClear()
    }

    ; Cold path: new Chrome window, then cache HWND.
    baseline := Map()
    for h in WinGetList("ahk_exe chrome.exe")
        baseline[h] := true
    try Run('chrome.exe --new-window "' . fileUrl . '"')
    catch {
        return false
    }

    deadline := A_TickCount + 8000
    newHwnd := 0
    while (A_TickCount < deadline) {
        for h in WinGetList("ahk_exe chrome.exe") {
            if (baseline.Has(h))
                continue
            try title := WinGetTitle("ahk_id " h)
            catch {
                title := ""
            }
            if (Palace_IsDashboardChromeWindowTitle(title) || title = "") {
                newHwnd := h
                if (Palace_IsDashboardChromeWindowTitle(title))
                    break 2
            }
        }
        Sleep(100)
    }
    if (!newHwnd) {
        ; Fallback: any Memory Palace titled chrome window.
        for h in WinGetList("ahk_exe chrome.exe") {
            try title := WinGetTitle("ahk_id " h)
            catch {
                continue
            }
            if (Palace_IsDashboardChromeWindowTitle(title)) {
                newHwnd := h
                break
            }
        }
    }
    if (!newHwnd)
        return false
    ; Wait briefly for title to settle to Memory Palace.
    settleDeadline := A_TickCount + 3000
    while (A_TickCount < settleDeadline) {
        try title := WinGetTitle("ahk_id " newHwnd)
        catch {
            title := ""
        }
        if (Palace_IsDashboardChromeWindowTitle(title))
            break
        Sleep(80)
    }
    Palace_DashboardHwndCacheSet(newHwnd)
    try WinActivate("ahk_id " newHwnd)
    catch {
    }
    return true
}
