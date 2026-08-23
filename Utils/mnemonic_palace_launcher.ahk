; =============================================================================
; Utils module: mnemonic_palace_launcher.ahk
; Memory Palace main menu (Utility Shortcuts [N])
; =============================================================================

Palace_LaunchApp() {
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
        . beasts.Length . " beasts  ·  " . atoms.Length . " atoms")
    g_PalaceGui.SetFont("s9 cF1C40F", "Segoe UI")
    g_PalaceGui.Add("Text", "x20 y72 w880", "1-3 quick links · letters open a module.")

    items := [
        ["1", "📹 Study Video", "Open / set video link (Google Docs API)"],
        ["2", "📖 Study Article", "Open / set article link (Google Docs API)"],
        ["3", "❤️ Favorite", "Open / set favorite link (Google Docs API)"],
        ["D", "Dashboard", "Study picker and Memory Palace images"],
        ["B", "Browse", "Studies -> palaces -> beasts -> atoms"],
        ["I", "AI import", "Desktop PALACE pack (preview)"],
        ["Q", "Quick image", "Newest Desktop PNG/JPG → last palace"],
        ["G", "Practice on GitHub", "Synced palace practice notes for mobile"],
        ["O", "Plans on GitHub", "Synced study plan checklists for mobile"],
        ["R", "Regen Markdown", "Force-create all practice + plan .md files"],
        ["H", "Help", "Vocabulary and mapping rules"],
        ["P", "Push to cloud", "Commit and push scripts repo"]
    ]

    x0 := 20
    y0 := 108
    colW := 440
    rowH := 78
    idx := 0
    for it in items {
        col := Mod(idx, 2)
        row := idx // 2
        x := x0 + col * colW
        y := y0 + row * rowH
        g_PalaceGui.SetFont("s14 cF1C40F Bold", "Segoe UI")
        g_PalaceGui.Add("Text", "x" . (x + 12) . " y" . (y + 10) . " w40 BackgroundTrans", "[" . it[1] . "]")
        g_PalaceGui.SetFont("s12 cWhite Bold", "Segoe UI")
        g_PalaceGui.Add("Text", "x" . (x + 58) . " y" . (y + 10) . " w340 BackgroundTrans", it[2])
        g_PalaceGui.SetFont("s9 cA0A0A0 Norm", "Segoe UI")
        g_PalaceGui.Add("Text", "x" . (x + 58) . " y" . (y + 36) . " w340 BackgroundTrans", it[3])
        idx += 1
    }

    g_PalaceGui.SetFont("s9 c808080", "Segoe UI")
    g_PalaceGui.Add("Text", "x20 y588 w880",
        "Backspace utility shortcuts   Esc close   1-3 quick links   R regen   G practice   O plans   P push")

    Palace_BindHotkeys([
        ["1", Palace_OnStudyVideo], ["2", Palace_OnStudyArticle], ["3", Palace_OnStudyFavorite],
        ["d", Palace_OnDash], ["b", Palace_OnBrowse], ["i", Palace_OnImp],
        ["q", Palace_OnQuickImage], ["g", Palace_OnPracticeGithub],
        ["o", Palace_OnPlansGithub], ["r", Palace_OnRegenMarkdown],
        ["h", Palace_OnHelp], ["p", Palace_OnGitPush],
        ["Backspace", (*) => Palace_ReturnToUtilityShortcuts()],
        ["Escape", (*) => Palace_CloseGui()]
    ])
    Palace_CenterGui(g_PalaceGui, 900, 630)
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
Palace_OnImp(*) {
    Palace_ImportMnemonicsFromDesktop()
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

Palace_OpenDashboard() {
    Palace_EnsureData()
    ; Restart so save-server code (e.g. add_backlog) is never stale after edits.
    Palace_EnsurePlanSaveServer(true)
    py := Palace_PythonDir() . "\chart_generator.py"
    if (!FileExist(py)) {
        Palace_Notify("chart_generator.py not found", 2000, BANNER_ACCENT_ERROR)
        return
    }
    dataDir := Palace_DataDir()
    outDir := Palace_OutputDir()
    notesRoot := Palace_NotesStudiesRoot()
    pyCmd := Palace_FindPythonCmd()
    if (pyCmd = "") {
        Palace_Notify("Python not found. Install Python or enable the py launcher.", 3500, BANNER_ACCENT_ERROR)
        return
    }
    try StandardLoadingBar_Show("Syncing technique + plans + building dashboard…", BANNER_ACCENT_INTERMEDIATE)
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
    try StandardLoadingBar_Hide(400)
    catch {
    }
    if (!FileExist(html)) {
        Palace_Notify("dashboard.html was not generated", 2200, BANNER_ACCENT_ERROR)
        return
    }
    tmpHtml := A_Temp . "\palace_dashboard_" . A_TickCount . ".html"
    try FileCopy(html, tmpHtml, 1)
    catch {
        tmpHtml := html
    }
    fileUrl := "file:///" . StrReplace(StrReplace(tmpHtml, "\", "/"), " ", "%20") . "?t=" . A_TickCount
    try Run('chrome.exe --new-window "' . fileUrl . '"')
    catch as e {
        Palace_Notify("Chrome failed: " . e.Message, 2500, BANNER_ACCENT_ERROR)
        return
    }
    Palace_CloseGui()
}
