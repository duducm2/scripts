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
    g_PalaceGui.Add("Text", "x20 y72 w880", "Letters open a module.")

    items := [
        ["D", "Dashboard", "Study picker and Memory Palace images"],
        ["B", "Browse", "Studies -> palaces -> beasts -> atoms"],
        ["I", "AI import", "Desktop PALACE pack (preview)"],
        ["Q", "Quick image", "Newest Desktop PNG/JPG → last palace"],
        ["G", "Practice on GitHub", "Synced palace practice notes for mobile"],
        ["O", "Plans on GitHub", "Synced study plan checklists for mobile"],
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
    g_PalaceGui.Add("Text", "x20 y438 w880",
        "Backspace utility shortcuts   Esc close   G practice   O plans on GitHub   I import   Q image   P push")

    Palace_BindHotkeys([
        ["d", Palace_OnDash], ["b", Palace_OnBrowse], ["i", Palace_OnImp],
        ["q", Palace_OnQuickImage], ["g", Palace_OnPracticeGithub],
        ["o", Palace_OnPlansGithub],
        ["h", Palace_OnHelp], ["p", Palace_OnGitPush],
        ["Backspace", (*) => Palace_ReturnToUtilityShortcuts()],
        ["Escape", (*) => Palace_CloseGui()]
    ])
    Palace_CenterGui(g_PalaceGui, 900, 500)
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
    Palace_OpenPracticeGithub()
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

Palace_OpenDashboard() {
    Palace_EnsureData()
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
