; =============================================================================
; Utils module: task_launcher.ahk
; Tasks main menu (Utility Shortcuts [T])
; =============================================================================

global g_TaskDashboardHwnd := 0

Task_LaunchApp() {
    Task_EnsureData()
    Task_ShowMainMenu()
}

Task_ShowMainMenu() {
    global g_TaskGui, g_TaskFilter, g_TaskBrowseProjectId, g_TaskBrowseTaskId
    Task_CloseGui()
    Task_EnsureData()
    g_TaskBrowseProjectId := ""
    g_TaskBrowseTaskId := ""

    projects := Task_Load("projects")
    tasks := Task_Load("tasks")
    infos := Task_Load("info_points")
    openN := Task_CountOpenTasks()

    g_TaskGui := Gui("+AlwaysOnTop +ToolWindow", "Tasks")
    g_TaskGui.SetFont("s10", "Segoe UI")
    g_TaskGui.BackColor := "1E1E1E"
    g_TaskGui.OnEvent("Close", (*) => Task_CloseGui())
    g_TaskGui.OnEvent("Escape", (*) => Task_CloseGui())
    try DllCall("dwmapi\DwmSetWindowAttribute", "ptr", g_TaskGui.Hwnd, "uint", 20, "int*", 1, "int", 4)
    catch {
    }

    g_TaskGui.SetFont("s16 cWhite Bold", "Segoe UI")
    g_TaskGui.Add("Text", "x20 y16 w880", "Tasks")
    g_TaskGui.SetFont("s10 cC0C0C0 Norm", "Segoe UI")
    g_TaskGui.Add("Text", "x20 y48 w880",
        projects.Length . " projects  ·  " . tasks.Length . " tasks  ·  "
        . infos.Length . " info  ·  " . openN . " open  ·  filter: " . Task_FilterLabel(g_TaskFilter))
    g_TaskGui.SetFont("s9 cF1C40F", "Segoe UI")
    g_TaskGui.Add("Text", "x20 y72 w880", "Keyboard-first project → task → info. Filters: work / personal / habits.")

    items := [
        ["B", "Browse", "Projects → tasks → info points"],
        ["1", "Inbox", "Open tasks (not ✅), current filter"],
        ["K", "Habits due", "Habitual tasks by next_due"],
        ["F", "Cycle filter", "all → work → personal → habits"],
        ["D", "Dashboard", "Charts and lists in Chrome"],
        ["M", "Migrate MD", "Import work / punctual / habits.md"],
        ["H", "Help", "Vocabulary and keys"],
        ["P", "Push", "Commit and push scripts repo"]
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
        g_TaskGui.SetFont("s14 cF1C40F Bold", "Segoe UI")
        g_TaskGui.Add("Text", "x" . (x + 12) . " y" . (y + 10) . " w40 BackgroundTrans", "[" . it[1] . "]")
        g_TaskGui.SetFont("s12 cWhite Bold", "Segoe UI")
        g_TaskGui.Add("Text", "x" . (x + 58) . " y" . (y + 10) . " w340 BackgroundTrans", it[2])
        g_TaskGui.SetFont("s9 cA0A0A0 Norm", "Segoe UI")
        g_TaskGui.Add("Text", "x" . (x + 58) . " y" . (y + 36) . " w340 BackgroundTrans", it[3])
        idx += 1
    }

    g_TaskGui.SetFont("s9 c808080", "Segoe UI")
    g_TaskGui.Add("Text", "x20 y440 w880",
        "Backspace utility shortcuts   Esc close   F filter   M migrate   D dashboard")

    Task_BindHotkeys([
        ["b", Task_OnBrowse], ["1", Task_OnInbox], ["k", Task_OnHabits],
        ["f", Task_OnFilter], ["d", Task_OnDash], ["m", Task_OnMigrate],
        ["h", Task_OnHelp], ["p", Task_OnGitPush],
        ["Backspace", (*) => Task_ReturnToUtilityShortcuts()],
        ["Escape", (*) => Task_CloseGui()]
    ])
    Task_CenterGui(g_TaskGui, 900, 500)
}

Task_ReturnToUtilityShortcuts() {
    Task_CloseGui()
    try ShowHotstringSelector()
    catch {
    }
}

Task_OnBrowse(*) {
    global g_TaskBrowseProjectId, g_TaskBrowseTaskId
    g_TaskBrowseProjectId := ""
    g_TaskBrowseTaskId := ""
    Task_ShowProjects()
}

Task_OnInbox(*) {
    global g_TaskBrowseProjectId, g_TaskBrowseTaskId
    g_TaskBrowseProjectId := ""
    g_TaskBrowseTaskId := ""
    Task_ShowTasksInbox(false)
}

Task_OnHabits(*) {
    global g_TaskBrowseProjectId, g_TaskBrowseTaskId
    g_TaskBrowseProjectId := ""
    g_TaskBrowseTaskId := ""
    Task_ShowTasksInbox(true)
}

Task_OnFilter(*) {
    next := Task_CycleFilter()
    Task_Notify("Filter: " . Task_FilterLabel(next), 1200, BANNER_ACCENT_INTERMEDIATE)
    Task_ShowMainMenu()
}

Task_OnDash(*) {
    Task_OpenDashboard()
}

Task_OnMigrate(*) {
    Task_RunMigrateFromMd()
}

Task_OnHelp(*) {
    Task_ShowHelp()
}

Task_OnGitPush(*) {
    Task_CloseGui()
    try Utility_GitSyncPush()
    catch as e {
        Task_Notify("Push failed: " . e.Message, 2500, BANNER_ACCENT_ERROR)
    }
}

Task_OpenDashboard() {
    Task_EnsureData()
    py := Task_PythonDir() . "\chart_generator.py"
    if (!FileExist(py)) {
        Task_Notify("chart_generator.py not found", 2000, BANNER_ACCENT_ERROR)
        return
    }
    dataDir := Task_DataDir()
    outDir := Task_OutputDir()
    pyCmd := Task_FindPythonCmd()
    if (pyCmd = "") {
        Task_Notify("Python not found.", 3500, BANNER_ACCENT_ERROR)
        return
    }
    try StandardLoadingBar_Show("Building dashboard…", BANNER_ACCENT_INTERMEDIATE)
    catch {
    }
    cmd := pyCmd . ' "' . py . '" --data-dir "' . dataDir . '" --output-dir "' . outDir . '"'
    exitCode := 0
    try {
        exitCode := RunWait(A_ComSpec . ' /c ' . cmd, A_ScriptDir, "Hide")
    } catch as e {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        Task_Notify("Python failed: " . e.Message, 2500, BANNER_ACCENT_ERROR)
        return
    }
    if (exitCode != 0) {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        Task_Notify("Dashboard failed (exit " . exitCode . ")", 3500, BANNER_ACCENT_ERROR)
        return
    }
    html := outDir . "\dashboard.html"
    try StandardLoadingBar_Hide(400)
    catch {
    }
    if (!FileExist(html)) {
        Task_Notify("dashboard.html was not generated", 2200, BANNER_ACCENT_ERROR)
        return
    }
    tmpHtml := A_Temp . "\tasks_dashboard_" . A_TickCount . ".html"
    try FileCopy(html, tmpHtml, 1)
    catch {
        tmpHtml := html
    }
    try Run('chrome.exe --new-window "' . tmpHtml . '"')
    catch {
        try Run('"' . tmpHtml . '"')
        catch as e {
            Task_Notify("Could not open dashboard: " . e.Message, 2500, BANNER_ACCENT_ERROR)
            return
        }
    }
    Task_Notify("Dashboard opened", 1400, BANNER_ACCENT_SUCCESS)
}

Task_ShowHelp() {
    global g_TaskGui
    Task_CloseGui()
    g_TaskGui := Gui("+AlwaysOnTop +ToolWindow", "Tasks Help")
    g_TaskGui.BackColor := "1E1E1E"
    g_TaskGui.SetFont("s12 cWhite Bold", "Segoe UI")
    g_TaskGui.Add("Text", "x20 y16 w700", "Tasks — help")
    y := 48
    for term in Task_Terms() {
        g_TaskGui.SetFont("s10 cF1C40F Bold", "Segoe UI")
        g_TaskGui.Add("Text", "x20 y" . y . " w700", term[1])
        y += 22
        g_TaskGui.SetFont("s9 cC0C0C0 Norm", "Segoe UI")
        g_TaskGui.Add("Text", "x20 y" . y . " w700", term[2])
        y += 36
    }
    g_TaskGui.SetFont("s9 c808080", "Segoe UI")
    g_TaskGui.Add("Text", "x20 y" . y . " w700", "Esc / Backspace → main menu")
    Task_BindHotkeys([
        ["Escape", (*) => Task_ShowMainMenu()],
        ["Backspace", (*) => Task_ShowMainMenu()]
    ])
    Task_CenterGui(g_TaskGui, 760, y + 60)
}
