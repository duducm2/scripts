; =============================================================================
; Utils module: mnemonic_palace_launcher.ahk
; Memory Palace — thin AHK launcher: start localhost :8767 + open Chrome web app
; (CRUD GUIs remain included for import/helpers; primary UX is mnemonics/web)
; =============================================================================

global g_PalaceDashboardHwnd := 0
global g_PalaceServerPid := 0

Palace_ServerPort() {
    return 8767
}

Palace_LaunchApp() {
    Palace_EnsureData()
    existing := Palace_FindExistingWebHwnd()
    if (existing) {
        if (!Palace_IsServerRunning()) {
            if (!Palace_EnsureServer(false))
                return
        }
        if (Palace_ActivateWeb(existing))
            return
        ; Cached HWND went stale — clear and fall through to open/re-find.
        Palace_WebHwndCacheClear()
        existing := Palace_FindExistingWebHwnd()
        if (existing && Palace_ActivateWeb(existing))
            return
    }
    try StandardLoadingBar_Show("⏳ Opening Memory Palace…", BANNER_ACCENT_INTERMEDIATE, { passive: false })
    catch {
    }
    if (!Palace_EnsureServer(false)) {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        return
    }
    ; Race: window may have appeared while the server was starting.
    existing := Palace_FindExistingWebHwnd()
    if (existing) {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        Palace_ActivateWeb(existing)
        return
    }
    url := "http://127.0.0.1:" . Palace_ServerPort() . "/?t=" . A_TickCount
    try StandardLoadingBar_Update("⏳ Opening Chrome…", BANNER_ACCENT_INTERMEDIATE)
    catch {
    }
    ok := Palace_OpenWebInChrome(url)
    try StandardLoadingBar_Hide(300)
    catch {
    }
    if (!ok)
        Palace_Notify("Chrome failed to open Memory Palace", 2500, BANNER_ACCENT_ERROR)
}

Palace_PidPath() {
    return Palace_DataDir() . "\palace_server.pid"
}

Palace_PidRead() {
    global g_PalaceServerPid
    pid := 0
    try pid := Integer(g_PalaceServerPid)
    catch {
        pid := 0
    }
    if (pid > 0) {
        try {
            if ProcessExist(pid)
                return pid
        } catch {
        }
    }
    path := Palace_PidPath()
    if (!FileExist(path))
        return 0
    raw := ""
    try raw := Trim(FileRead(path))
    catch {
        return 0
    }
    try pid := Integer(raw)
    catch {
        return 0
    }
    if (pid <= 0)
        return 0
    try {
        if ProcessExist(pid)
            return pid
    } catch {
    }
    return 0
}

Palace_PidWrite(pid) {
    global g_PalaceServerPid
    try pid := Integer(pid)
    catch {
        pid := 0
    }
    g_PalaceServerPid := pid
    path := Palace_PidPath()
    if (pid > 0) {
        try FileDelete(path)
        catch {
        }
        try FileAppend(String(pid), path, "UTF-8")
        catch {
        }
        return
    }
    try FileDelete(path)
    catch {
    }
}

Palace_KillPidTree(pid) {
    try pid := Integer(pid)
    catch {
        return
    }
    if (pid <= 0)
        return
    try RunWait(A_ComSpec . " /c taskkill /F /T /PID " . pid . " >nul 2>&1", , "Hide")
    catch {
    }
}

Palace_StopServerNetstat(port) {
    q := Chr(39)
    cmd := "for /f `"tokens=5`" %a in (" . q . "netstat -ano ^| findstr :" . port
        . " ^| findstr LISTENING" . q . ") do taskkill /F /PID %a >nul 2>&1"
    try RunWait(A_ComSpec . " /c " . cmd, , "Hide")
    catch {
    }
}

Palace_StopServer(port := 0) {
    if (port = 0)
        port := Palace_ServerPort()
    pid := Palace_PidRead()
    if (pid > 0)
        Palace_KillPidTree(pid)
    deadline := A_TickCount + 1500
    while (A_TickCount < deadline) {
        if (!Palace_IsServerRunning(port))
            break
        Sleep 50
    }
    if (Palace_IsServerRunning(port))
        Palace_StopServerNetstat(port)
    Palace_PidWrite(0)
}

Palace_IsServerRunning(port := 0) {
    if (port = 0)
        port := Palace_ServerPort()
    try {
        whr := ComObject("WinHttp.WinHttpRequest.5.1")
        whr.Open("GET", "http://127.0.0.1:" . port . "/health", false)
        whr.Send()
        return (whr.Status = 200)
    } catch {
        return false
    }
}

Palace_EnsureServer(forceRestart := false) {
    port := Palace_ServerPort()
    if (forceRestart)
        Palace_StopServer(port)
    else if (Palace_IsServerRunning(port))
        return true
    py := Palace_PythonDir() . "\palace_server.py"
    if (!FileExist(py)) {
        Palace_Notify("palace_server.py not found", 2200, BANNER_ACCENT_ERROR)
        return false
    }
    pyCmd := Palace_FindPythonCmd()
    if (pyCmd = "") {
        Palace_Notify("Python not found for Memory Palace server", 2500, BANNER_ACCENT_ERROR)
        return false
    }
    dataDir := Palace_DataDir()
    outDir := Palace_OutputDir()
    scriptsRoot := A_ScriptDir
    studiesRoot := A_ScriptDir . "\mnemonics\studies"
    cmd := pyCmd . ' "' . py . '" --data-dir "' . dataDir . '" --output-dir "' . outDir
        . '" --studies-root "' . studiesRoot . '" --scripts-root "' . scriptsRoot
        . '" --port ' . port
    pid := 0
    try pid := Run(cmd, A_ScriptDir, "Hide")
    catch as e {
        Palace_Notify("Palace server failed: " . e.Message, 2800, BANNER_ACCENT_ERROR)
        return false
    }
    if (pid > 0)
        Palace_PidWrite(pid)
    loop 30 {
        if (Palace_IsServerRunning(port))
            return true
        Sleep 150
    }
    Palace_Notify("Palace server did not start", 2800, BANNER_ACCENT_ERROR)
    return false
}

Palace_IsChromeWindowTitle(title) {
    t := Trim(title)
    if (t = "")
        return false
    ; Chrome notification badge: "(1) Memory Palace - Google Chrome"
    t := RegExReplace(t, "^\(\d+\)\s+", "")
    if (t = "Memory Palace" || t = "Memory Palace - Google Chrome")
        return true
    if (InStr(t, "Memory Palace") = 1)
        return true
    ; Title still on the localhost URL (tab not fully titled yet, or URL bar mode).
    port := String(Palace_ServerPort())
    if (InStr(t, "127.0.0.1:" . port) || InStr(t, "localhost:" . port))
        return true
    return false
}

Palace_WebHwndValid(hwnd) {
    try hwnd := Integer(hwnd)
    catch {
        return 0
    }
    if (!(hwnd is Integer) || hwnd <= 0)
        return 0
    try {
        if !WinExist("ahk_id " hwnd)
            return 0
    } catch {
        return 0
    }
    return hwnd
}

Palace_WebHwndCacheGet() {
    global g_PalaceDashboardHwnd
    hwnd := Palace_WebHwndValid(g_PalaceDashboardHwnd)
    if (hwnd)
        return hwnd
    raw := Trim(IniRead(Palace_SettingsPath(), "General", "WebChromeHwnd", ""))
    if (raw = "")
        raw := Trim(IniRead(Palace_SettingsPath(), "General", "DashboardChromeHwnd", ""))
    hwnd := Palace_WebHwndValid(raw)
    g_PalaceDashboardHwnd := hwnd
    return hwnd
}

Palace_WebHwndCacheSet(hwnd) {
    global g_PalaceDashboardHwnd
    hwnd := Palace_WebHwndValid(hwnd)
    g_PalaceDashboardHwnd := hwnd
    try IniWrite(hwnd ? String(hwnd) : "", Palace_SettingsPath(), "General", "WebChromeHwnd")
    catch {
    }
}

Palace_WebHwndCacheClear() {
    Palace_WebHwndCacheSet(0)
}

Palace_HwndLooksLikeWeb(hwnd) {
    hwnd := Palace_WebHwndValid(hwnd)
    if (!hwnd)
        return 0
    try title := WinGetTitle("ahk_id " hwnd)
    catch {
        return 0
    }
    if !Palace_IsChromeWindowTitle(title)
        return 0
    return hwnd
}

Palace_FindExistingWebHwnd() {
    hwnd := Palace_HwndLooksLikeWeb(Palace_WebHwndCacheGet())
    if (hwnd)
        return hwnd
    ; Fast path: title contains "Memory Palace" (handles badges / suffixes).
    prev := A_TitleMatchMode
    try {
        SetTitleMatchMode(2)
        try hit := WinExist("Memory Palace ahk_exe chrome.exe")
        catch {
            hit := 0
        }
        hwnd := Palace_HwndLooksLikeWeb(hit)
        if (hwnd) {
            Palace_WebHwndCacheSet(hwnd)
            return hwnd
        }
        port := String(Palace_ServerPort())
        try hit := WinExist("127.0.0.1:" . port . " ahk_exe chrome.exe")
        catch {
            hit := 0
        }
        hwnd := Palace_HwndLooksLikeWeb(hit)
        if (hwnd) {
            Palace_WebHwndCacheSet(hwnd)
            return hwnd
        }
    } finally {
        SetTitleMatchMode(prev)
    }
    for h in WinGetList("ahk_exe chrome.exe") {
        hwnd := Palace_HwndLooksLikeWeb(h)
        if (hwnd) {
            Palace_WebHwndCacheSet(hwnd)
            return hwnd
        }
    }
    Palace_WebHwndCacheClear()
    return 0
}

Palace_ActivateWeb(hwnd) {
    hwnd := Palace_WebHwndValid(hwnd)
    if (!hwnd)
        return false
    try {
        if (WinGetMinMax("ahk_id " hwnd) = -1)
            WinRestore("ahk_id " hwnd)
    } catch {
    }
    try WinActivate("ahk_id " hwnd)
    catch {
        return false
    }
    try WinWaitActive("ahk_id " hwnd, , 1)
    catch {
    }
    Palace_WebHwndCacheSet(hwnd)
    return true
}

Palace_OpenWebInChrome(url) {
    ; Never spawn a duplicate if we can still see the app window.
    existing := Palace_FindExistingWebHwnd()
    if (existing) {
        Palace_ActivateWeb(existing)
        return true
    }
    try Run('chrome.exe --new-window "' . url . '"')
    catch {
        try Run(url)
        catch {
            return false
        }
    }
    newHwnd := 0
    prev := A_TitleMatchMode
    try {
        SetTitleMatchMode(2)
        if WinWait("Memory Palace ahk_exe chrome.exe", , 10) {
            try newHwnd := WinExist("Memory Palace ahk_exe chrome.exe")
            catch {
                newHwnd := 0
            }
        }
        if (!newHwnd) {
            port := String(Palace_ServerPort())
            if WinWait("127.0.0.1:" . port . " ahk_exe chrome.exe", , 3) {
                try newHwnd := WinExist("127.0.0.1:" . port . " ahk_exe chrome.exe")
                catch {
                    newHwnd := 0
                }
            }
        }
    } finally {
        SetTitleMatchMode(prev)
    }
    newHwnd := Palace_HwndLooksLikeWeb(newHwnd)
    if (!newHwnd) {
        for h in WinGetList("ahk_exe chrome.exe") {
            newHwnd := Palace_HwndLooksLikeWeb(h)
            if (newHwnd)
                break
        }
    }
    if (newHwnd) {
        Palace_WebHwndCacheSet(newHwnd)
        try WinActivate("ahk_id " newHwnd)
        catch {
        }
        return true
    }
    return true
}

; Legacy menu kept for reference / emergency; primary entry is Palace_LaunchApp → web.
Palace_ShowMainMenu() {
    Palace_LaunchApp()
}

Palace_ReturnToUtilityShortcuts() {
    Palace_CloseGui()
    try ShowHotstringSelector()
    catch {
    }
}

Palace_OnDash(*) {
    Palace_LaunchApp()
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
    ; Legacy entry — primary UX is the web app on :8767 (same as Palace_LaunchApp).
    Palace_LaunchApp()
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
