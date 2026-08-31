; =============================================================================
; Shift keys module: cursor_predicates.ahk
; Cursor/VS Code editor detection and UIA helpers
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

;-----------------------------------------
;  Detect which editor is active
;-----------------------------------------
IsCursorActive() {
    return WinActive("ahk_exe Cursor.exe")
}

IsCodeActive() {
    return WinActive("ahk_exe Code.exe")
}

IsEditorActive() {
    return IsCursorActive() || IsCodeActive()
}

;-----------------------------------------
;  UIA: detect if focus is in Cursor main editor (Monaco inputarea)
;  Uses conditional path to locate editor element, then compares with focused element
;  Conditional path: RootView -> ... -> workbench.parts.editor -> editor-instance -> Edit
;-----------------------------------------
IsCursorMainEditorFocused() {
    try {
        if !(WinActive("ahk_exe Cursor.exe") || WinActive("ahk_exe Code.exe"))
            return false

        winHandle := WinExist("A")
        if (!winHandle)
            return false

        root := UIA.ElementFromHandle(winHandle)
        if (!root)
            return false

        editorEl := root.ElementFromPath({ T: 33, CN: "RootView" }, { T: 33 }, { T: 33 }, { T: 33, CN: "ClientView" }, { T: 33 }, { T: 33 }, { T: 33 }, { T: 30 }, { T: 26 }, { T: 33 }, { T: 26,
            A: "workbench.parts.editor" }, { T: 26, CN: "editor-instance" }, { T: 20 }, { T: 4 }
        )
        if (!editorEl)
            return false

        fe := UIA.GetFocusedElement()
        if (!fe)
            return false

        return UIA.CompareElementsEx(editorEl, fe)
    } catch {
        return false
    }
}

Editor_EnsureCursorWindowActive(editorHwnd) {
    if !(editorHwnd is Integer) || editorHwnd <= 0
        return false
    try {
        if (WinActive("ahk_exe Cursor.exe") || WinActive("ahk_exe Code.exe"))
            return true
        WinActivate("ahk_id " editorHwnd)
        return WinActive("ahk_id " editorHwnd)
    } catch {
        return false
    }
}

; Focus the Cursor "Files Explorer" tree using UIA
FocusCursorFilesExplorer() {
    try {
        if !(WinActive("ahk_exe Cursor.exe") || WinActive("ahk_exe Code.exe"))
            return false

        hwnd := WinExist("A")
        if (!hwnd)
            return false

        root := UIA.ElementFromHandle(hwnd)
        if (!root)
            return false

        treeType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Tree)
        feEn := UIA.CreatePropertyCondition(UIA.Property.Name, "Files Explorer")
        fePt := UIA.CreatePropertyCondition(UIA.Property.Name, "Explorador de Arquivos")
        feName := UIA.CreateOrCondition(feEn, fePt)
        feCond := UIA.CreateAndCondition(treeType, feName)

        fileTree := ""
        try fileTree := root.FindElement(feCond, UIA.TreeScope.Descendants)

        if !fileTree
            return false

        fileTree.SetFocus()
        return true
    } catch {
        return false
    }
}

Editor_FindCursorFilesExplorerTree(root) {
    if !root
        return 0
    treeType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Tree)
    feEn := UIA.CreatePropertyCondition(UIA.Property.Name, "Files Explorer")
    fePt := UIA.CreatePropertyCondition(UIA.Property.Name, "Explorador de Arquivos")
    feName := UIA.CreateOrCondition(feEn, fePt)
    feCond := UIA.CreateAndCondition(treeType, feName)
    fileTree := 0
    try fileTree := root.FindElement(feCond, UIA.TreeScope.Descendants)
    if fileTree
        return fileTree
    for autoId in ["FileExplorer3", "FileExplorer2", "FileExplorer"] {
        try {
            cond := UIA.CreatePropertyCondition(UIA.Property.AutomationId, autoId)
            fileTree := root.FindElement(cond, UIA.TreeScope.Descendants)
            if fileTree
                return fileTree
        }
    }
    return 0
}

Editor_UiaElementHasAncestor(el, ancestor) {
    if !el || !ancestor
        return false
    current := el
    loop 40 {
        if !current
            break
        try {
            if UIA.CompareElements(ancestor, current)
                return true
        } catch {
        }
        try current := UIA.TreeWalkerTrue.GetParentElement(current)
        catch
            break
    }
    return false
}

; True when keyboard focus is in the sidebar Files Explorer tree (safe for Ctrl+2 Copy Path).
Editor_FocusIsInFilesExplorer(editorHwnd := 0) {
    try {
        if !(WinActive("ahk_exe Cursor.exe") || WinActive("ahk_exe Code.exe"))
            return false
        if !editorHwnd
            editorHwnd := WinExist("A")
        if !editorHwnd
            return false
        root := UIA.ElementFromHandle(editorHwnd)
        if !root
            return false
        fileTree := Editor_FindCursorFilesExplorerTree(root)
        if !fileTree
            return false
        fe := UIA.GetFocusedElement()
        if !fe
            return false
        return Editor_UiaElementHasAncestor(fe, fileTree)
    } catch {
        return false
    }
}

; Smart nav Explorer wait: conditional UIA poll vs legacy fixed sleep (efficiency-canon §2).
global EDITOR_USE_CONDITIONAL_EXPLORER_WAIT := true
global EDITOR_COPY_VERIFY_FILEDROP := true
global EDITOR_COPY_USE_EDITOR_FASTPATH := true
global EDITOR_COPY_PREFER_DIRECT_SET := true
global EDITOR_REVEAL_STABLE_POLLS := 2
global EDITOR_REVEAL_FIND_ITEM_FALLBACK := true
global EDITOR_COPY_CLIP_WAIT_MS := 500
global EDITOR_COPY_DIRECT_CLIP_WAIT_MS := 300
global EDITOR_SMARTNAV_TIMING := false
global EDITOR_SMARTNAV_MIN_INTERVAL_MS := 450
global EDITOR_SMARTNAV_USE_LOADING_BAR := false
global g_EditorSmartNavLastTick := 0

Editor_SmartNav_TimingLog(phase, ms) {
    global EDITOR_SMARTNAV_TIMING
    if (!EDITOR_SMARTNAV_TIMING)
        return
    try FileAppend('{"phase":"' phase '","ms":' ms '}' "`n", A_ScriptDir "\.cursor\smartnav-timing.log", "UTF-8")
    catch {
    }
}

Editor_TryCopyFileFromActiveEditor(editorHwnd, expectedBasename := "") {
    if !(editorHwnd is Integer) || editorHwnd <= 0
        return { ok: false, path: "" }
    savedClip := 0
    copyOk := false
    pathText := ""
    try savedClip := ClipboardAll()
    catch {
    }
    try {
        ; Ctrl+2 = Copy Path only when Files Explorer has focus. In the code editor it opens editor group 2.
        if !Editor_FocusIsInFilesExplorer(editorHwnd)
            return { ok: false, path: "" }
        A_Clipboard := ""
        SendInput "^2"
        clipOk := ClipWait(0.6)
        if !clipOk {
            Sleep 80
            A_Clipboard := ""
            SendInput "^2"
            clipOk := ClipWait(0.6)
        }
        if !clipOk
            return { ok: false, path: "" }
        pathText := Trim(Trim(A_Clipboard), Chr(34))
        if !Editor_PathIsExistingFile(pathText)
            return { ok: false, path: "" }
        verifyBasename := expectedBasename != "" ? expectedBasename : Editor_GetBasenameFromEditorTitle(editorHwnd)
        if !(Editor_SetClipboardFiles([pathText])
        && Editor_WaitForClipboardFileDrop(EDITOR_COPY_DIRECT_CLIP_WAIT_MS)
        && Editor_ClipboardMatchesRevealTarget(pathText, verifyBasename))
            return { ok: false, path: "" }
        copyOk := true
        return { ok: true, path: pathText }
    } finally {
        if (!copyOk && IsObject(savedClip)) {
            try A_Clipboard := savedClip
        }
    }
}

Editor_CopyVerifiedFileToClipboard(fullPath, verifyBasename, clipWaitMs := 300) {
    if !Editor_PathIsExistingFile(fullPath)
        return false
    return Editor_SetClipboardFiles([fullPath])
    && Editor_WaitForClipboardFileDrop(clipWaitMs)
    && Editor_ClipboardMatchesRevealTarget(fullPath, verifyBasename)
}

Editor_SmartNavLoadingUpdate(state, editorHwnd := 0) {
    global EDITOR_SMARTNAV_USE_LOADING_BAR
    if (!EDITOR_SMARTNAV_USE_LOADING_BAR)
        return
    try StandardLoadingBar_Update(state, BANNER_ACCENT_INTERMEDIATE)
    catch {
    }
}

Editor_GetSelectedSidebarItemName(editorHwnd) {
    if !(editorHwnd is Integer) || editorHwnd <= 0
        return ""
    ; Do not use UIA.GetFocusedElement — Agents panel / a11y overlays steal focus and return wrong names.
    try {
        root := UIA.ElementFromHandle(editorHwnd)
        if (!root)
            return ""
        items := root.FindAll({ Type: UIA.Type.TreeItem })
        for item in items {
            try {
                if (!item.GetPropertyValue(UIA.Property.IsSelected))
                    continue
                if (item.GetPropertyValue(UIA.Property.IsOffscreen))
                    continue
                nm := item.Name
                if (nm != "" && Editor_IsPlausibleRevealBasename(nm))
                    return nm
            } catch {
            }
        }
    } catch {
    }
    return ""
}

; Strip Cursor binary-tab placeholder text (e.g. "file.pptx, The file is not displayed..."). Any extension.
Editor_NormalizeRevealBasename(raw) {
    if (raw = "")
        return ""
    s := Trim(raw)
    if InStr(s, ",")
        s := Trim(SubStr(s, 1, InStr(s, ",") - 1))
    if (InStr(s, "\") || InStr(s, "/")) {
        SplitPath s, &name
        if (name != "")
            s := name
    }
    return s
}

Editor_GetExpectedRevealBasename(editorHwnd) {
    raw := Editor_GetBasenameFromEditorTitle(editorHwnd)
    if (raw = "") {
        name := Editor_GetSelectedSidebarItemName(editorHwnd)
        if (name != "")
            raw := name
    }
    if (raw = "") {
        try {
            name := Cursor_GetFocusedExplorerItemName()
            if (name != "" && Editor_IsPlausibleRevealBasename(name))
                raw := name
        } catch {
        }
    }
    return Editor_NormalizeRevealBasename(raw)
}

; Single poll: ItemsView + at least one highlighted item (IDE reveal always pre-selects).
Editor_ExplorerRevealReadyOnce(explorerHwnd) {
    if !(explorerHwnd is Integer) || explorerHwnd <= 0 || !WinExist("ahk_id " explorerHwnd)
        return false
    try {
        root := UIA.ElementFromHandle(explorerHwnd)
        if (!root)
            return false
        itemsView := Explorer_FindItemsView(root)
        if (!itemsView)
            return false
        selected := Explorer_GetItemsViewSelection(itemsView)
        return (selected && selected.Length > 0)
    } catch {
        return false
    }
}

Editor_WaitForExplorerItemsView(explorerHwnd, timeoutMs := 600) {
    if !(explorerHwnd is Integer) || explorerHwnd <= 0
        return false
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        try {
            root := UIA.ElementFromHandle(explorerHwnd)
            if Explorer_FindItemsView(root)
                return true
        } catch {
        }
        Sleep 50
    }
    return false
}

; Bounded poll until reveal selection is stable (consecutive OK polls).
Editor_WaitForExplorerRevealReady(explorerHwnd, timeoutMs := 3500) {
    global EDITOR_REVEAL_STABLE_POLLS
    if !(explorerHwnd is Integer) || explorerHwnd <= 0
        return false
    stableRequired := Max(1, EDITOR_REVEAL_STABLE_POLLS)
    deadline := A_TickCount + timeoutMs
    stableCount := 0
    while (A_TickCount < deadline) {
        if (!WinExist("ahk_id " explorerHwnd))
            return false
        if (Editor_ExplorerRevealReadyOnce(explorerHwnd)) {
            stableCount++
            if (stableCount >= stableRequired)
                return true
            Sleep 75
        } else {
            stableCount := 0
            Sleep 50
        }
    }
    return false
}

Editor_WaitForSidebarExplorerFocus(timeoutMs := 800) {
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        if (FocusCursorFilesExplorer())
            return true
        Sleep 50
    }
    return false
}

; Switch to Files Explorer sidebar when focus is elsewhere (SCM, Search, editor, etc.).
Editor_EnsureFilesExplorerSidebarFocused(editorHwnd := 0) {
    if Editor_FocusIsInFilesExplorer(editorHwnd)
        return true
    if !editorHwnd
        editorHwnd := WinExist("A")
    Editor_EnsureCursorWindowActive(editorHwnd)
    Send "^+e"
    if Editor_WaitForSidebarExplorerFocus(800)
        return true
    Send "^!+e"
    return Editor_WaitForSidebarExplorerFocus(400)
}

Editor_FindWorkbenchToggleButton(root, nameSubstring) {
    if !root || !nameSubstring
        return 0
    toggleBtn := 0
    try toggleBtn := root.FindFirst({ Name: nameSubstring, Type: UIA.Type.Button })
    catch
        toggleBtn := 0
    if toggleBtn
        return toggleBtn
    try {
        allButtons := root.FindAll({ Type: UIA.Type.Button })
        if allButtons {
            for btn in allButtons {
                try {
                    if InStr(btn.Name, nameSubstring) {
                        toggleBtn := btn
                        break
                    }
                } catch {
                }
            }
        }
    } catch {
    }
    return toggleBtn
}

; True when a workbench title-bar toggle (e.g. primary/secondary sidebar) is in the "on" state.
Editor_IsWorkbenchToggleOn(root, nameSubstring) {
    toggleBtn := Editor_FindWorkbenchToggleButton(root, nameSubstring)
    if !toggleBtn
        return false
    try {
        if InStr(toggleBtn.ClassName, "checked")
            return true
    } catch {
    }
    try {
        if toggleBtn.GetPropertyValue(UIA.Property.IsTogglePatternAvailable) {
            toggleState := toggleBtn.TogglePattern.ToggleState
            ; ToggleState: 1 = On, 2 = Off, 3 = Indeterminate
            return (toggleState = 1)
        }
    } catch {
    }
    return false
}

; =============================================================================
; Alt+S — robot types git stash / fetch / pull in a new editor terminal
; =============================================================================
global EDITOR_GIT_TERMINAL_READY_MS := 800
global EDITOR_GIT_ROBOT_TIMEOUT_MS := 600000
global g_EditorGitRobotResultPath := ""
global g_EditorGitRobotDeadline := 0
global g_EditorGitRobotQuickUpdate := false

Editor_NormGitDir(p) {
    if (p = "")
        return ""
    return StrLower(StrReplace(RTrim(p, "\"), "/", "\"))
}

; True when this editor is the project-selector slot [s] Scripts (assets/data/projects.ini).
Editor_IsThisScriptsRepo(hwnd, repoDir) {
    global g_Projects
    try ProjectData_Load()
    catch {
    }
    if (hwnd) {
        title := ""
        try title := WinGetTitle("ahk_id " hwnd)
        catch {
        }
        idx := 0
        try idx := CursorTransfer_GetMatchingProjectIndex(hwnd, title)
        catch {
        }
        if (idx > 0 && IsSet(g_Projects) && IsObject(g_Projects) && idx <= g_Projects.Length) {
            project := g_Projects[idx]
            if (IsObject(project) && project.HasProp("char") && StrLower(project.char) = "s")
                return true
        }
    }
    repoNorm := Editor_NormGitDir(repoDir)
    if (repoNorm = "" || !IsSet(g_Projects) || !IsObject(g_Projects))
        return false
    loop g_Projects.Length {
        project := g_Projects[A_Index]
        if !(IsObject(project) && project.HasProp("char") && StrLower(project.char) = "s")
            continue
        if (project.HasProp("path") && project.path != "" && repoNorm = Editor_NormGitDir(project.path))
            return true
        if (project.HasProp("workPath") && project.workPath != "" && repoNorm = Editor_NormGitDir(project.workPath))
            return true
    }
    return false
}

Editor_GitFlowFail(step, reason := "") {
    msg := "❌ " step " failed"
    if (reason != "")
        msg .= ": " reason
    try StandardLoadingBar_Show(msg, BANNER_ACCENT_ERROR)
    try StandardLoadingBar_Hide(1200)
}

Editor_GetEditorProcessCommandLine(editorHwnd := 0) {
    if !editorHwnd
        editorHwnd := WinExist("A")
    if !editorHwnd
        return ""
    pid := 0
    try pid := WinGetPID("ahk_id " editorHwnd)
    if !pid
        return ""
    cmd := ""
    try {
        locator := ComObject("WbemScripting.SWbemLocator")
        svc := locator.ConnectServer(".", "root\cimv2")
        for proc in svc.ExecQuery("SELECT CommandLine FROM Win32_Process WHERE ProcessId = " pid) {
            try cmd := proc.CommandLine
            break
        }
    } catch {
    }
    return cmd ? cmd : ""
}

Editor_DecodeFolderUri(uri) {
    if (uri = "" || !InStr(uri, "file:"))
        return ""
    path := RegExReplace(uri, "i)^file://+", "")
    path := StrReplace(path, "/", "\")
    if RegExMatch(path, "^[A-Za-z]/")
        path := SubStr(path, 1, 1) ":" SubStr(path, 2)
    loop {
        if !RegExMatch(path, "%([0-9A-Fa-f]{2})", &m)
            break
        path := StrReplace(path, m[0], Chr("0x" m[1]))
    }
    return path
}

Editor_ExtractFolderPathsFromCmdLine(cmdLine) {
    paths := []
    seen := Map()
    if (cmdLine = "")
        return paths
    pos := 1
    while RegExMatch(cmdLine, "i)--folder-uri\s+(\S+)", &m, pos) {
        decoded := Editor_DecodeFolderUri(m[1])
        if decoded && DirExist(decoded) && !seen.Has(decoded) {
            seen[decoded] := true
            paths.Push(decoded)
        }
        pos := m.Pos(0) + m.Len(0)
    }
    pos := 1
    while RegExMatch(cmdLine, '"([A-Za-z]:[^"]+)"', &m, pos) {
        p := m[1]
        if DirExist(p) && !seen.Has(p) {
            seen[p] := true
            paths.Push(p)
        }
        pos := m.Pos(0) + m.Len(0)
    }
    return paths
}

Editor_GetScmRepoBasenameFromStatusBar(editorHwnd := 0) {
    try {
        if !editorHwnd
            editorHwnd := WinExist("A")
        if !editorHwnd
            return ""
        root := UIA.ElementFromHandle(editorHwnd)
        if !root
            return ""
        el := root.FindFirst({ AutomationId: "status.scm.0" })
        if !el
            return ""
        if RegExMatch(el.Name, "^([^(]+)\s*\(Git\)", &m)
            return Trim(m[1])
    } catch {
    }
    return ""
}

Editor_DeduplicatePaths(paths) {
    out := []
    seen := Map()
    for p in paths {
        if !p || seen.Has(p)
            continue
        seen[p] := true
        out.Push(p)
    }
    return out
}

Editor_AppendProjectRegistryPaths(&paths) {
    try ProjectData_Load()
    catch {
    }
    if !IsSet(g_Projects) || !IsObject(g_Projects)
        return
    try {
        loop g_Projects.Length {
            project := g_Projects[A_Index]
            for key in ["path", "workPath"] {
                p := ""
                try {
                    if IsObject(project) && project.HasProp(key)
                        p := project.%key%
                } catch {
                }
                if p && DirExist(p)
                    paths.Push(p)
            }
        }
    } catch {
    }
}

Editor_PickGitRootByBasename(candidates, basename) {
    if !basename
        return ""
    baseLow := StrLower(basename)
    for p in candidates {
        folder := RTrim(p, "\")
        SplitPath folder, &dirName
        if (StrLower(dirName) = baseLow) {
            top := GitCli_RevParseTopLevel(p)
            if top
                return top
        }
    }
    return ""
}

Editor_ResolveGitRepoDir(editorHwnd := 0) {
    if !editorHwnd
        editorHwnd := WinExist("A")
    if !editorHwnd
        return ""
    cmdLine := Editor_GetEditorProcessCommandLine(editorHwnd)
    candidates := Editor_ExtractFolderPathsFromCmdLine(cmdLine)
    Editor_AppendProjectRegistryPaths(&candidates)
    candidates := Editor_DeduplicatePaths(candidates)
    basename := Editor_GetScmRepoBasenameFromStatusBar(editorHwnd)
    if !candidates.Length
        return ""
    if (candidates.Length = 1)
        return GitCli_RevParseTopLevel(candidates[1])
    if basename {
        picked := Editor_PickGitRootByBasename(candidates, basename)
        if picked
            return picked
    }
    for p in candidates {
        top := GitCli_RevParseTopLevel(p)
        if top
            return top
    }
    return ""
}

Editor_OpenNewEditorTerminal(editorHwnd) {
    global EDITOR_GIT_TERMINAL_READY_MS
    Editor_EnsureCursorWindowActive(editorHwnd)
    Send '^+"'
    Sleep EDITOR_GIT_TERMINAL_READY_MS
    return (WinActive("ahk_exe Cursor.exe") || WinActive("ahk_exe Code.exe"))
}

Editor_GitQuotePs(s) {
    return "'" StrReplace(s, "'", "''") "'"
}

; Thin launcher: run the on-disk gated flow (later .ps1 fixes apply after one AHK reload).
Editor_GitWriteRobotScript(repoDir, resultPath) {
    flowPs1 := A_ScriptDir "\infra\tools\Editor-GitStashFetchPull.ps1"
    if !FileExist(flowPs1)
        return ""
    repoArg := (repoDir != "") ? repoDir : "."
    script := "& " Editor_GitQuotePs(flowPs1)
    script .= " -RepoDir " Editor_GitQuotePs(repoArg)
    script .= " -ResultPath " Editor_GitQuotePs(resultPath)
    script .= " -TimeoutSec 300 -Robot`r`n"
    ps1Path := A_Temp "\editor-git-robot-" A_TickCount ".ps1"
    try FileDelete(ps1Path)
    FileAppend(script, ps1Path, "UTF-8-RAW")
    return ps1Path
}

Editor_ParseGitRobotResultText(raw) {
    compact := StrReplace(StrReplace(StrReplace(raw, Chr(0), ""), Chr(65279), ""), " ", "")
    compact := Trim(compact, "`r`n `t")
    if InStr(compact, "fail")
        return "fail"
    if InStr(compact, "ok")
        return "ok"
    return ""
}

Editor_GitPlayPullSuccessSound() {
    try {
        soundPath := A_ScriptDir "\assets\sounds\pull-successful.wav"
        if FileExist(soundPath)
            ScriptSoundPlay(soundPath, true)
    } catch {
    }
}

Editor_GitRobotStopPoll() {
    global g_EditorGitRobotResultPath, g_EditorGitRobotDeadline, g_EditorGitRobotQuickUpdate
    try SetTimer(Editor_GitRobotPollResult, 0)
    catch {
    }
    g_EditorGitRobotResultPath := ""
    g_EditorGitRobotDeadline := 0
    g_EditorGitRobotQuickUpdate := false
}

Editor_GitRobotStartPoll(resultPath) {
    global g_EditorGitRobotResultPath, g_EditorGitRobotDeadline, EDITOR_GIT_ROBOT_TIMEOUT_MS
    Editor_GitRobotStopPoll()
    g_EditorGitRobotResultPath := resultPath
    g_EditorGitRobotDeadline := A_TickCount + EDITOR_GIT_ROBOT_TIMEOUT_MS
    SetTimer(Editor_GitRobotPollResult, 250)
}

Editor_GitRobotPollResult(*) {
    global g_EditorGitRobotResultPath, g_EditorGitRobotDeadline, g_EditorGitRobotQuickUpdate
    path := g_EditorGitRobotResultPath
    if (path = "") {
        Editor_GitRobotStopPoll()
        return
    }
    if (A_TickCount > g_EditorGitRobotDeadline) {
        Editor_GitRobotStopPoll()
        Editor_GitFlowFail("Git", "timed out waiting for stash/pull")
        return
    }
    outcome := ""
    try {
        if FileExist(path) {
            raw := FileRead(path)
            outcome := Editor_ParseGitRobotResultText(raw)
            if (outcome != "")
                try FileDelete(path)
        }
    } catch {
    }
    if (outcome = "")
        return
    doQuickUpdate := false
    try doQuickUpdate := g_EditorGitRobotQuickUpdate
    Editor_GitRobotStopPoll()
    if (outcome = "ok") {
        Editor_GitPlayPullSuccessSound()
        if (doQuickUpdate) {
            try StandardLoadingBar_Show("⏳ Restarting scripts...", BANNER_ACCENT_INTERMEDIATE)
            QuickUpdateScripts()
            return
        }
        StandardLoadingBar_Show("✅ Pull complete", BANNER_ACCENT_SUCCESS)
        StandardLoadingBar_Hide(600)
        return
    }
    Editor_GitFlowFail("Git", "stash/pull failed")
}

Editor_GitStashAndPull() {
    hwnd := WinExist("A")
    if !hwnd
        return
    if !(WinActive("ahk_exe Cursor.exe") || WinActive("ahk_exe Code.exe")) {
        Editor_GitFlowFail("Git", "active window is not Cursor or VS Code")
        return
    }
    repoDir := Editor_ResolveGitRepoDir(hwnd)
    resultPath := A_Temp "\editor-git-robot-" A_TickCount ".txt"
    try FileDelete(resultPath)
    StandardLoadingBar_Show("🤖 Opening terminal…", BANNER_ACCENT_INTERMEDIATE, { passive: false,
        centerOnHwnd: hwnd })
    try StandardLoadingBar_ArmForceHide(2500)
    try {
        Editor_OpenNewEditorTerminal(hwnd)
        Sleep 120
        ps1Path := Editor_GitWriteRobotScript(repoDir, resultPath)
        if (ps1Path = "") {
            Editor_GitFlowFail("Git", "Editor-GitStashFetchPull.ps1 not found")
            return
        }
        SendText("powershell -NoProfile -ExecutionPolicy Bypass -File " . Editor_GitQuotePs(ps1Path))
        Send "{Enter}"
    } finally {
        ; Never leave the opening banner up while git runs.
        try StandardLoadingBar_Hide(0)
    }
    Editor_GitRobotStartPoll(resultPath)
    global g_EditorGitRobotQuickUpdate
    g_EditorGitRobotQuickUpdate := Editor_IsThisScriptsRepo(hwnd, repoDir)
}

; ---------------------------------------------------------------------------
; Shift+P: Pull first, then Sync Changes (may also push). Quality gates.
; ---------------------------------------------------------------------------
Editor_GitInvokeUiaElement(el) {
    if !IsObject(el)
        return false
    try {
        if el.GetPropertyValue(UIA.Property.IsInvokePatternAvailable) {
            try {
                el.InvokePattern.Invoke()
                return true
            } catch {
            }
            try {
                el.Invoke()
                return true
            } catch {
            }
        }
    } catch {
    }
    try {
        el.Click()
        return true
    } catch {
    }
    return false
}

Editor_GitIsPullButtonName(name) {
    if (name = "")
        return false
    ; Avoid "Pull Request" / PR actions.
    if InStr(name, "Pull Request")
        return false
    ; Bidirectional sync button is handled as Sync, not Pull-only.
    if RegExMatch(name, "i)\bPull \d+ and push \d+ commits?\b")
        return false
    ; Exact Pull / Pull from… / Git: Pull.
    if RegExMatch(name, "i)^Pull$")
        return true
    if RegExMatch(name, "i)^Pull from\b")
        return true
    if RegExMatch(name, "i)^Git:\s*Pull$")
        return true
    ; Cursor/VS Code SCM + status bar: "Pull 1 commits from origin/main"
    ; or "scripts (Git) - Pull 1 commits from origin/main".
    if RegExMatch(name, "i)\bPull \d+ commits?\b")
        return true
    return false
}

Editor_GitIsSyncButtonName(name) {
    if (name = "")
        return false
    ; In progress or post-commit — do not click.
    if InStr(name, "Synchronizing")
        return false
    if InStr(name, "Commit & Sync") || InStr(name, "Commit and Sync")
        return false
    if InStr(name, "Sync Changes") || InStr(name, "Synchronize Changes")
        return true
    ; SCM / status bar when both ahead and behind:
    ; "Pull 1 and push 1 commits between origin/main"
    if RegExMatch(name, "i)\bPull \d+ and push \d+ commits?\b")
        return true
    ; Ahead-only primary action on the Sync row.
    if RegExMatch(name, "i)\bPush \d+ commits?\b")
        return true
    return false
}

Editor_GitFindPullControl(root) {
    if !IsObject(root)
        return 0
    try {
        for name in ["Pull", "Pull from...", "Git: Pull"] {
            try {
                btn := root.FindFirst({ Name: name, Type: UIA.Type.Button })
                if btn
                    return btn
            } catch {
            }
        }
        for btn in root.FindAll({ Type: UIA.Type.Button }) {
            n := ""
            try n := btn.Name
            if Editor_GitIsPullButtonName(n)
                return btn
        }
    } catch {
    }
    return 0
}

Editor_GitFindSyncControl(root) {
    if !IsObject(root)
        return 0
    try {
        ; Prefer SCM action button (title includes Sync Changes + optional ahead/behind).
        try {
            btn := root.FindFirst({ Name: "Sync Changes", Type: UIA.Type.Button, matchmode: "Substring" })
            if btn {
                n := ""
                try n := btn.Name
                if Editor_GitIsSyncButtonName(n)
                    return btn
            }
        } catch {
        }
        for btn in root.FindAll({ Type: UIA.Type.Button }) {
            n := ""
            try n := btn.Name
            if Editor_GitIsSyncButtonName(n)
                return btn
        }
        ; SCM row is often a TreeItem named "$(sync) Sync Changes …" whose
        ; actionable child is a monaco Button (Pull/Push/Sync). Click that, not the row.
        try {
            for ti in root.FindAll({ Type: UIA.Type.TreeItem }) {
                n := ""
                try n := ti.Name
                if !Editor_GitIsSyncButtonName(n)
                    continue
                try {
                    for btn in ti.FindAll({ Type: UIA.Type.Button }) {
                        if btn
                            return btn
                    }
                } catch {
                }
                return ti
            }
        } catch {
        }
        ; Status bar often uses "Synchronize Changes".
        for needle in ["Synchronize Changes", "Sync Changes"] {
            try {
                el := root.FindFirst({ Name: needle, matchmode: "Substring" })
                if el {
                    n := ""
                    try n := el.Name
                    if Editor_GitIsSyncButtonName(n)
                        return el
                }
            } catch {
            }
        }
    } catch {
    }
    return 0
}

Editor_GitRunPaletteCommand(cmdTitle) {
    if (cmdTitle = "")
        return false
    try {
        Send "^+p"
        Sleep 220
        SendText cmdTitle
        Sleep 180
        Send "{Enter}"
        Sleep 120
        return true
    } catch {
    }
    return false
}

Editor_GitPullOrSync() {
    hwnd := WinExist("A")
    if !hwnd
        return
    if !(WinActive("ahk_exe Cursor.exe") || WinActive("ahk_exe Code.exe")) {
        Editor_GitFlowFail("Git", "active window is not Cursor or VS Code")
        return
    }

    try FocusSourceControlViewForCommitGeneration()
    catch {
        Send "+d"
        Sleep 450
    }

    root := 0
    try root := UIA.ElementFromHandle(hwnd)

    if IsObject(root) {
        ; Gate 2: UIA Pull button
        pullBtn := Editor_GitFindPullControl(root)
        if pullBtn && Editor_GitInvokeUiaElement(pullBtn) {
            try StandardLoadingBar_Show("⬇️ Pull", BANNER_ACCENT_SUCCESS)
            try StandardLoadingBar_Hide(600)
            return
        }
        ; Gates 3–4: UIA Sync Changes (SCM button or status bar)
        syncEl := Editor_GitFindSyncControl(root)
        if syncEl && Editor_GitInvokeUiaElement(syncEl) {
            try StandardLoadingBar_Show("🔄 Sync Changes (pull + push)", BANNER_ACCENT_SUCCESS)
            try StandardLoadingBar_Hide(800)
            return
        }
    }

    ; Gate 5: Command palette Git: Pull
    if Editor_GitRunPaletteCommand("Git: Pull") {
        try StandardLoadingBar_Show("⬇️ Pull (palette)", BANNER_ACCENT_SUCCESS)
        try StandardLoadingBar_Hide(600)
        return
    }

    ; Gate 6: Command palette Git: Sync (only if opening/typing Pull palette failed)
    if Editor_GitRunPaletteCommand("Git: Sync") {
        try StandardLoadingBar_Show("🔄 Sync (palette; pull + push)", BANNER_ACCENT_SUCCESS)
        try StandardLoadingBar_Hide(800)
        return
    }

    Editor_GitFlowFail("Git", "Pull/Sync not found")
}

; Primary sidebar open: Cursor uses sidebarvisible on monaco-workbench; VS Code uses the title-bar toggle.
Editor_IsPrimarySidebarVisible(editorHwnd := 0) {
    try {
        if !editorHwnd
            editorHwnd := WinExist("A")
        if !editorHwnd
            return false
        root := UIA.ElementFromHandle(editorHwnd)
        if !root
            return false
        for el in root.FindAll({ Type: UIA.Type.Pane }) {
            try {
                cls := el.ClassName
                if InStr(cls, "monaco-workbench") && InStr(cls, "sidebarvisible")
                    return true
            } catch {
            }
        }
        if Editor_IsWorkbenchToggleOn(root, "Toggle Primary Side Bar")
            return true
    } catch {
    }
    return false
}

Editor_HidePrimarySidebar(editorHwnd := 0) {
    if !editorHwnd
        editorHwnd := WinExist("A")
    if !editorHwnd || !Editor_IsPrimarySidebarVisible(editorHwnd)
        return true
    Editor_EnsureCursorWindowActive(editorHwnd)
    Send "^b"
    deadline := A_TickCount + 600
    while (A_TickCount < deadline) {
        Sleep 50
        if !Editor_IsPrimarySidebarVisible(editorHwnd)
            return true
    }
    return !Editor_IsPrimarySidebarVisible(editorHwnd)
}

; Shift+E relay (^+e) twice — leave SCM/sidebar chrome and restore focus to the main editor.
Editor_ReturnFocusToMainEditor(editorHwnd := 0) {
    if !editorHwnd
        editorHwnd := WinExist("A")
    if !editorHwnd
        return false
    Editor_EnsureCursorWindowActive(editorHwnd)
    Send "^+e"
    Sleep 80
    Send "^+e"
    return true
}

; Folder path from Explorer window title (path before " - File Explorer" / localized suffix).
Editor_ParseExplorerFolderFromTitle(title) {
    if (title = "")
        return ""
    for suffix in [" - File Explorer", " - Explorador de arquivos", " – Explorador de Arquivos"] {
        pos := InStr(title, suffix, false)
        if (pos > 1)
            return RTrim(SubStr(title, 1, pos - 1), "\")
    }
    return ""
}

Editor_GetExplorerFolderPathFromShell(explorerHwnd) {
    if !(explorerHwnd is Integer) || explorerHwnd <= 0
        return ""
    try {
        shell := ComObject("Shell.Application")
        for window in shell.Windows {
            try {
                if (window.hwnd = explorerHwnd)
                    return window.Document.Folder.Self.Path
            } catch {
            }
        }
    } catch {
    }
    return ""
}

Editor_ClipboardMatchesRevealTarget(fullPath, basename) {
    paths := Editor_GetClipboardFilePaths()
    if (paths.Length = 0)
        return false
    if (fullPath != "" && FileExist(fullPath))
        return Editor_ClipboardContainsFilePath(fullPath)
    targetBasename := Editor_NormalizeRevealBasename(basename)
    if (targetBasename = "")
        return true
    for path in paths {
        try SplitPath path, &name
        if (Editor_NormalizeRevealBasename(name) = targetBasename)
            return true
    }
    return false
}

Editor_GetRevealSelectedItemName(explorerHwnd, expectedBasename := "") {
    try {
        root := UIA.ElementFromHandle(explorerHwnd)
        itemsView := Explorer_FindItemsView(root)
        if (!itemsView)
            return Editor_NormalizeRevealBasename(expectedBasename)
        selected := Explorer_GetItemsViewSelection(itemsView)
        if (selected.Length > 0) {
            try {
                nm := selected[1].Name
                if (nm != "")
                    return Editor_NormalizeRevealBasename(nm)
            } catch {
            }
        }
    } catch {
    }
    return Editor_NormalizeRevealBasename(expectedBasename)
}

; Focus highlighted item only (no tree scan — reveal already selected it).
Editor_EnsureRevealItemSelected(explorerHwnd, expectedBasename := "") {
    if !(explorerHwnd is Integer) || explorerHwnd <= 0
        return false
    try {
        root := UIA.ElementFromHandle(explorerHwnd)
        itemsView := Explorer_FindItemsView(root)
        if (!itemsView)
            return false
        selected := Explorer_GetItemsViewSelection(itemsView)
        if (!selected || selected.Length = 0)
            return false
        try itemsView.SetFocus()
        target := selected[1]
        try target.SetFocus()
        catch {
            try target.Select()
        }
        return true
    } catch {
        return false
    }
}

Editor_SelectRevealListItem(explorerHwnd, itemsView, item) {
    if (!item || !itemsView)
        return false
    try {
        item.ScrollIntoView()
        try item.Select()
        catch {
        }
        try item.SetFocus()
        catch {
        }
        try itemsView.SetFocus()
        itemsView := Explorer_WaitItemsViewKeyboardFocus(explorerHwnd, itemsView, 400)
        selected := Explorer_GetItemsViewSelection(itemsView)
        return (selected && selected.Length > 0)
    } catch {
        return false
    }
}

Editor_TryFindAndSelectRevealItemInExplorer(explorerHwnd, expectedBasename := "") {
    expectedNorm := Editor_NormalizeRevealBasename(expectedBasename)
    if !Editor_IsPlausibleRevealBasename(expectedNorm)
        return false
    if !(explorerHwnd is Integer) || explorerHwnd <= 0
        return false
    try {
        root := UIA.ElementFromHandle(explorerHwnd)
        itemsView := Explorer_FindItemsView(root)
        if (!itemsView)
            return false
        targetItem := 0
        try targetItem := itemsView.FindFirst({ Type: "ListItem", Name: expectedNorm })
        catch {
        }
        if (!targetItem) {
            for li in itemsView.FindAll({ Type: "ListItem" }) {
                try {
                    if (Editor_NormalizeRevealBasename(li.Name) = expectedNorm) {
                        targetItem := li
                        break
                    }
                } catch {
                }
            }
        }
        if (!targetItem)
            return false
        return Editor_SelectRevealListItem(explorerHwnd, itemsView, targetItem)
    } catch {
        return false
    }
}

Editor_TryExplorerSearchSelectRevealItem(explorerHwnd, expectedBasename := "") {
    expectedNorm := Editor_NormalizeRevealBasename(expectedBasename)
    if !Editor_IsPlausibleRevealBasename(expectedNorm)
        return false
    if !(explorerHwnd is Integer) || explorerHwnd <= 0
        return false
    try {
        WinActivate("ahk_id " explorerHwnd)
        if !WinActive("ahk_id " explorerHwnd)
            return false
        Send "^e"
        Sleep 120
        SendText expectedNorm
        deadline := A_TickCount + 1200
        while (A_TickCount < deadline) {
            root := UIA.ElementFromHandle(explorerHwnd)
            itemsView := Explorer_FindItemsView(root)
            if (itemsView) {
                for li in itemsView.FindAll({ Type: "ListItem" }) {
                    try {
                        if (Editor_NormalizeRevealBasename(li.Name) = expectedNorm) {
                            if Editor_SelectRevealListItem(explorerHwnd, itemsView, li)
                                return true
                        }
                    } catch {
                    }
                }
            }
            Sleep 80
        }
    } catch {
    }
    return false
}

Editor_TryRecoverRevealTarget(explorerHwnd, expectedBasename := "") {
    result := Map("ok", false, "fullPath", "", "selected", false)
    if !(explorerHwnd is Integer) || explorerHwnd <= 0
        return result
    expectedNorm := Editor_NormalizeRevealBasename(expectedBasename)
    if !Editor_IsPlausibleRevealBasename(expectedNorm)
        return result

    fullPath := Editor_ResolveRevealFullPath(explorerHwnd, expectedBasename)
    if !Editor_PathIsExistingFile(fullPath) {
        folder := Editor_GetExplorerFolderPathFromShell(explorerHwnd)
        if (folder = "")
            folder := Editor_ParseExplorerFolderFromTitle(WinGetTitle("ahk_id " explorerHwnd))
        if (folder != "") {
            candidate := RTrim(folder, "\") "\" expectedNorm
            if Editor_PathIsExistingFile(candidate)
                fullPath := candidate
        }
    }
    if Editor_PathIsExistingFile(fullPath) {
        result["fullPath"] := fullPath
        result["ok"] := true
        return result
    }

    if Editor_TryFindAndSelectRevealItemInExplorer(explorerHwnd, expectedBasename) {
        result["selected"] := true
        fullPath := Editor_ResolveRevealFullPath(explorerHwnd, expectedBasename)
        if Editor_PathIsExistingFile(fullPath)
            result["fullPath"] := fullPath
        result["ok"] := true
        return result
    }

    if Editor_TryExplorerSearchSelectRevealItem(explorerHwnd, expectedBasename) {
        result["selected"] := true
        fullPath := Editor_ResolveRevealFullPath(explorerHwnd, expectedBasename)
        if Editor_PathIsExistingFile(fullPath)
            result["fullPath"] := fullPath
        result["ok"] := true
        return result
    }

    return result
}

Editor_BuildRevealedFilePath(explorerHwnd, expectedBasename := "") {
    fileName := Editor_GetRevealSelectedItemName(explorerHwnd, expectedBasename)
    if (fileName = "")
        return ""
    try {
        folder := Editor_GetExplorerFolderPathFromShell(explorerHwnd)
        if (folder = "")
            folder := Editor_ParseExplorerFolderFromTitle(WinGetTitle("ahk_id " explorerHwnd))
        if (folder = "")
            return ""
        return RTrim(folder, "\") "\" fileName
    } catch {
        return ""
    }
}

Editor_ResolveRevealFullPath(explorerHwnd, expectedBasename := "") {
    fullPath := Editor_BuildRevealedFilePath(explorerHwnd, expectedBasename)
    if Editor_PathIsExistingFile(fullPath)
        return fullPath
    folderFromShell := Editor_GetExplorerFolderPathFromShell(explorerHwnd)
    basename := Editor_GetRevealSelectedItemName(explorerHwnd, expectedBasename)
    expectedNorm := Editor_NormalizeRevealBasename(expectedBasename)
    if (expectedNorm != "" && Editor_NormalizeRevealBasename(basename) != expectedNorm)
        basename := expectedNorm
    if (basename = "")
        basename := expectedNorm
    if (folderFromShell != "" && basename != "") {
        candidate := RTrim(folderFromShell, "\") "\" basename
        if Editor_PathIsExistingFile(candidate)
            return candidate
    }
    return Editor_PathIsExistingFile(fullPath) ? fullPath : ""
}

Editor_ParsePreRevealExplorerHwnds(preRevealHwnds := "") {
    preSet := Map()
    if (preRevealHwnds = "")
        return preSet
    for part in StrSplit(preRevealHwnds, ",") {
        if (part = "")
            continue
        try preSet[Integer(part)] := true
    }
    return preSet
}

Editor_ScoreExplorerRevealWindow(hwnd, expectedBasename := "", preSet := unset) {
    score := 0
    if (IsSet(preSet) && IsObject(preSet) && !preSet.Has(hwnd))
        score += 100
    expectedNorm := Editor_NormalizeRevealBasename(expectedBasename)
    if !Editor_IsPlausibleRevealBasename(expectedNorm)
        expectedNorm := ""
    folder := Editor_GetExplorerFolderPathFromShell(hwnd)
    selNorm := Editor_NormalizeRevealBasename(Editor_GetRevealSelectedItemName(hwnd, expectedBasename))
    if (expectedNorm != "" && selNorm = expectedNorm && folder != "") {
        candidate := RTrim(folder, "\") "\" expectedNorm
        if Editor_PathIsExistingFile(candidate)
            score += 250
        else
            score += 150
    } else if (expectedNorm != "" && selNorm = expectedNorm)
        score += 40
    if (expectedNorm != "" && folder != "") {
        candidate := RTrim(folder, "\") "\" expectedNorm
        if Editor_PathIsExistingFile(candidate)
            score += 80
    }
    return score
}

Editor_GatherRevealContext(explorerHwnd, expectedBasename := "") {
    ctx := Map()
    ctx["folderFromShell"] := Editor_GetExplorerFolderPathFromShell(explorerHwnd)
    ctx["selectedName"] := ""
    try {
        root := UIA.ElementFromHandle(explorerHwnd)
        itemsView := Explorer_FindItemsView(root)
        if (itemsView) {
            selected := Explorer_GetItemsViewSelection(itemsView)
            if (selected.Length > 0) {
                try {
                    nm := selected[1].Name
                    if (nm != "")
                        ctx["selectedName"] := Editor_NormalizeRevealBasename(nm)
                } catch {
                }
            }
        }
    } catch {
    }
    expectedNorm := Editor_NormalizeRevealBasename(expectedBasename)
    selectedNorm := ctx["selectedName"]
    if (expectedNorm != "" && Editor_IsPlausibleRevealBasename(expectedNorm)
    && (selectedNorm = "" || selectedNorm != expectedNorm))
        verifyBasename := expectedNorm
    else if (selectedNorm != "")
        verifyBasename := selectedNorm
    else
        verifyBasename := expectedNorm
    ctx["verifyBasename"] := verifyBasename
    folderFromShell := ctx["folderFromShell"]
    fullPath := ""
    if (folderFromShell != "" && verifyBasename != "") {
        candidate := RTrim(folderFromShell, "\") "\" verifyBasename
        if Editor_PathIsExistingFile(candidate)
            fullPath := candidate
    }
    if (fullPath = "")
        fullPath := Editor_ResolveRevealFullPath(explorerHwnd, expectedBasename)
    if !Editor_PathIsExistingFile(fullPath)
        fullPath := ""
    ctx["fullPath"] := fullPath
    return ctx
}

; After Enter/Run: wait until Explorer loses foreground or window closes (bounded poll, not fixed sleep).
Editor_WaitForShellDispatchedAfterOpen(explorerHwnd, timeoutMs := 2000) {
    if !(explorerHwnd is Integer) || explorerHwnd <= 0
        return false
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        if !WinExist("ahk_id " explorerHwnd)
            return true
        if !WinActive("ahk_id " explorerHwnd)
            return true
        Sleep 50
    }
    return false
}

; After Reveal in File Explorer: copy/open selected file in Windows Explorer, close window.
Editor_WaitForActiveExplorerWindow(timeoutSec := 2.5, expectedBasename := "", editorHwnd := 0, preRevealHwnds := "") {
    global EDITOR_USE_CONDITIONAL_EXPLORER_WAIT, EDITOR_REVEAL_FIND_ITEM_FALLBACK
    Editor_SmartNavLoadingUpdate("⏳ Waiting for Explorer window…", editorHwnd)
    preSet := Editor_ParsePreRevealExplorerHwnds(preRevealHwnds)
    deadline := A_TickCount + Round(timeoutSec * 1000)
    explorerHwnd := 0
    bestScore := 0
    while (A_TickCount < deadline) {
        bestHwnd := 0
        bestScore := 0
        for hwnd in WinGetList("ahk_exe explorer.exe") {
            score := Editor_ScoreExplorerRevealWindow(hwnd, expectedBasename, preSet)
            if (score > bestScore) {
                bestScore := score
                bestHwnd := hwnd
            }
        }
        if (bestHwnd && bestScore >= 150) {
            try {
                WinActivate("ahk_id " bestHwnd)
                if WinActive("ahk_id " bestHwnd) {
                    explorerHwnd := bestHwnd
                    break
                }
            } catch {
            }
        } else if (bestHwnd && bestScore >= 100) {
            try {
                WinActivate("ahk_id " bestHwnd)
                if WinActive("ahk_id " bestHwnd) {
                    explorerHwnd := bestHwnd
                    break
                }
            } catch {
            }
        }
        Sleep 50
    }
    if (!explorerHwnd || !WinExist("ahk_id " explorerHwnd))
        return 0
    Editor_SmartNavLoadingUpdate("⏳ Loading file list…", editorHwnd)
    Editor_WaitForExplorerItemsView(explorerHwnd, 600)
    if (EDITOR_USE_CONDITIONAL_EXPLORER_WAIT) {
        Editor_SmartNavLoadingUpdate("⏳ Confirming file selection…", editorHwnd)
        revealMs := Max(3500, Round(timeoutSec * 1000))
        if !Editor_WaitForExplorerRevealReady(explorerHwnd, revealMs) {
            if (EDITOR_REVEAL_FIND_ITEM_FALLBACK) {
                Editor_SmartNavLoadingUpdate("⏳ Finding file in Explorer…", editorHwnd)
                recovered := Editor_TryRecoverRevealTarget(explorerHwnd, expectedBasename)
                if (recovered.ok)
                    return explorerHwnd
            }
            return 0
        }
    } else {
        Sleep 2500
    }
    return explorerHwnd
}

Editor_SmartNavRevealShowExplorerTimeout(actionLabel := "") {
    try StandardLoadingBar_Hide(0)
    catch {
    }
    msg := "Explorer opened but the file was not selected in time."
    if (actionLabel != "")
        msg := actionLabel ": " msg
    try ShowCenteredOverlay_Utils("❌ " msg, 2200, BANNER_ACCENT_ERROR)
    catch {
        try TrayTip("Smart Nav", msg, "Iconx")
        SetTimer(() => TrayTip(), -2200)
    }
}

Editor_SmartNavRevealShowSuccess(explorerAction, expectedBasename := "") {
    msg := "✅ Revealed in Explorer"
    if (explorerAction = "copy")
        msg := "✅ File copied to clipboard"
    else if (explorerAction = "open")
        msg := "✅ File opened"
    else if (explorerAction = "share")
        msg := "✅ OneDrive share link copied"
    name := Editor_NormalizeRevealBasename(expectedBasename)
    if (name != "" && Editor_IsPlausibleRevealBasename(name))
        msg .= ": " name
    try ShowCenteredOverlay_Utils(msg, 1400, BANNER_ACCENT_SUCCESS)
    catch {
    }
}

Editor_CopyFromWindowsExplorerAndReturn(editorHwnd, expectedBasename := "", timeoutSec := 2.5, preRevealHwnds := "") {
    global EDITOR_COPY_VERIFY_FILEDROP, EDITOR_COPY_PREFER_DIRECT_SET
    global EDITOR_COPY_CLIP_WAIT_MS, EDITOR_COPY_DIRECT_CLIP_WAIT_MS
    savedClip := 0
    savedTextClip := ""
    try savedClip := ClipboardAll()
    catch {
    }
    try savedTextClip := A_Clipboard
    catch {
    }
    copyOk := false
    explorerHwnd := 0
    try {
        explorerHwnd := Editor_WaitForActiveExplorerWindow(timeoutSec, expectedBasename, editorHwnd, preRevealHwnds)
        if (!explorerHwnd) {
            Editor_SmartNavRevealShowExplorerTimeout("Copy")
            return false
        }
        Editor_SmartNavLoadingUpdate("⏳ Copying file to clipboard…", editorHwnd)

        if (!EDITOR_COPY_VERIFY_FILEDROP) {
            if (!Editor_EnsureRevealItemSelected(explorerHwnd, expectedBasename)) {
                try Explorer_EnsureItemsViewFocusPreserveSelection()
                catch {
                }
            }
            Send "^c"
            if !ClipWait(0.8, 1) && (A_Clipboard = "" || A_Clipboard = savedTextClip) {
                Editor_SmartNavRevealShowExplorerTimeout("Copy")
                return false
            }
            try WinClose("ahk_id " explorerHwnd)
            explorerHwnd := 0
            copyOk := true
            return true
        }

        ctx := Editor_GatherRevealContext(explorerHwnd, expectedBasename)
        fullPath := ctx["fullPath"]
        verifyBasename := ctx["verifyBasename"]

        if (!Editor_PathIsExistingFile(fullPath)) {
            recovered := Editor_TryRecoverRevealTarget(explorerHwnd, expectedBasename)
            if (recovered.ok) {
                if (recovered.fullPath != "")
                    fullPath := recovered.fullPath
                if (recovered.selected)
                    Editor_EnsureRevealItemSelected(explorerHwnd, expectedBasename)
            }
        }

        Editor_EnsureRevealItemSelected(explorerHwnd, expectedBasename)

        if (EDITOR_COPY_PREFER_DIRECT_SET && Editor_PathIsExistingFile(fullPath)) {
            if Editor_CopyVerifiedFileToClipboard(fullPath, verifyBasename, EDITOR_COPY_DIRECT_CLIP_WAIT_MS) {
                try WinClose("ahk_id " explorerHwnd)
                explorerHwnd := 0
                copyOk := true
                return true
            }
        }

        try Explorer_EnsureItemsViewFocusPreserveSelection()
        catch {
        }
        A_Clipboard := ""
        Send "^c"
        if !(Editor_WaitForClipboardFileDrop(EDITOR_COPY_CLIP_WAIT_MS)
        && Editor_ClipboardMatchesRevealTarget(fullPath, verifyBasename)) {
            if (Editor_PathIsExistingFile(fullPath)
            && Editor_CopyVerifiedFileToClipboard(fullPath, verifyBasename, EDITOR_COPY_DIRECT_CLIP_WAIT_MS)) {
                ; direct set succeeded after keyboard miss
            } else {
                recovered := Editor_TryRecoverRevealTarget(explorerHwnd, expectedBasename)
                if (recovered.ok) {
                    if (recovered.fullPath != "")
                        fullPath := recovered.fullPath
                    if (recovered.selected)
                        Editor_EnsureRevealItemSelected(explorerHwnd, expectedBasename)
                }
                if (Editor_PathIsExistingFile(fullPath)
                && Editor_CopyVerifiedFileToClipboard(fullPath, verifyBasename, EDITOR_COPY_DIRECT_CLIP_WAIT_MS)) {
                    ; recovered via find/select or direct path
                } else {
                    Editor_SmartNavRevealShowExplorerTimeout("Copy")
                    return false
                }
            }
        }

        try WinClose("ahk_id " explorerHwnd)
        explorerHwnd := 0
        copyOk := true
        return true
    } catch {
        Editor_SmartNavRevealShowExplorerTimeout("Copy")
        return false
    } finally {
        if (!copyOk) {
            try {
                if (IsObject(savedClip))
                    A_Clipboard := savedClip
            } catch {
            }
        }
        if (editorHwnd) {
            try WinActivate("ahk_id " editorHwnd)
        }
    }
}

Editor_OpenFromWindowsExplorer(editorHwnd, expectedBasename := "", timeoutSec := 2.5, preRevealHwnds := "") {
    try {
        explorerHwnd := Editor_WaitForActiveExplorerWindow(timeoutSec, expectedBasename, editorHwnd, preRevealHwnds)
        if (!explorerHwnd) {
            Editor_SmartNavRevealShowExplorerTimeout("Open")
            if (editorHwnd) {
                try WinActivate("ahk_id " editorHwnd)
            }
            return false
        }

        Editor_EnsureRevealItemSelected(explorerHwnd, expectedBasename)
        fullPath := Editor_ResolveRevealFullPath(explorerHwnd, expectedBasename)
        if (!Editor_PathIsExistingFile(fullPath)) {
            recovered := Editor_TryRecoverRevealTarget(explorerHwnd, expectedBasename)
            if (recovered.ok) {
                if (recovered.fullPath != "")
                    fullPath := recovered.fullPath
                if (recovered.selected)
                    Editor_EnsureRevealItemSelected(explorerHwnd, expectedBasename)
            }
        }
        opened := false

        if (Editor_PathIsExistingFile(fullPath)) {
            Editor_SmartNavLoadingUpdate("⏳ Opening file…", editorHwnd)
            try {
                Run '"' fullPath '"'
                opened := true
                Editor_WaitForShellDispatchedAfterOpen(explorerHwnd, 2500)
            } catch {
            }
        }

        if (!opened) {
            if (!Editor_EnsureRevealItemSelected(explorerHwnd, expectedBasename)) {
                recovered := Editor_TryRecoverRevealTarget(explorerHwnd, expectedBasename)
                if (recovered.ok && recovered.selected)
                    Editor_EnsureRevealItemSelected(explorerHwnd, expectedBasename)
            }
            Editor_SmartNavLoadingUpdate("⏳ Opening file (Enter)…", editorHwnd)
            try Explorer_EnsureItemsViewFocusPreserveSelection()
            catch {
            }
            Send "{Enter}"
            opened := Editor_WaitForShellDispatchedAfterOpen(explorerHwnd, 2500)
        }

        try WinClose("ahk_id " explorerHwnd)

        if (!opened) {
            Editor_SmartNavRevealShowExplorerTimeout("Open")
            if (editorHwnd) {
                try WinActivate("ahk_id " editorHwnd)
            }
            return false
        }
        return true
    } catch {
        Editor_SmartNavRevealShowExplorerTimeout("Open")
        if (editorHwnd) {
            try WinActivate("ahk_id " editorHwnd)
        }
        return false
    }
}

Editor_ShareFromWindowsExplorerAndReturn(editorHwnd, expectedBasename := "", timeoutSec := 4.0, preRevealHwnds := "") {
    explorerHwnd := 0
    try {
        explorerHwnd := Editor_WaitForActiveExplorerWindow(timeoutSec, expectedBasename, editorHwnd, preRevealHwnds)
        if (!explorerHwnd) {
            Editor_SmartNavRevealShowExplorerTimeout("Share")
            return false
        }

        if (!Editor_EnsureRevealItemSelected(explorerHwnd, expectedBasename)) {
            try Explorer_EnsureItemsViewFocusPreserveSelection()
            catch {
            }
        }

        Editor_SmartNavLoadingUpdate("⏳ Sharing on OneDrive…", editorHwnd)
        if !Explorer_CopyOneDriveShareLink_BoschGroup()
            return false

        try WinClose("ahk_id " explorerHwnd)
        explorerHwnd := 0
        return true
    } catch {
        Editor_SmartNavRevealShowExplorerTimeout("Share")
        return false
    } finally {
        if (explorerHwnd) {
            try WinClose("ahk_id " explorerHwnd)
            catch {
            }
        }
        if (editorHwnd) {
            try WinActivate("ahk_id " editorHwnd)
        }
    }
}

Editor_SmartNavRevealAfterSendH(editorHwnd, explorerAction, expectedBasename := "", preRevealHwnds := "") {
    if (explorerAction = "copy")
        return Editor_CopyFromWindowsExplorerAndReturn(editorHwnd, expectedBasename, 2.5, preRevealHwnds)
    if (explorerAction = "open")
        return Editor_OpenFromWindowsExplorer(editorHwnd, expectedBasename, 2.5, preRevealHwnds)
    if (explorerAction = "share")
        return Editor_ShareFromWindowsExplorerAndReturn(editorHwnd, expectedBasename, 4.0, preRevealHwnds)
    return true
}

; Smart navigation - Editor → Explorer, Explorer → Reveal in Explorer (optional copy/open/share in Windows Explorer).
Editor_SmartNavReveal(explorerAction := "") {
    global g_EditorSmartNavLastTick, EDITOR_SMARTNAV_MIN_INTERVAL_MS, EDITOR_COPY_USE_EDITOR_FASTPATH
    global EDITOR_SMARTNAV_USE_LOADING_BAR
    if (g_EditorSmartNavLastTick && (A_TickCount - g_EditorSmartNavLastTick) < EDITOR_SMARTNAV_MIN_INTERVAL_MS)
        return
    g_EditorSmartNavLastTick := A_TickCount

    editorHwnd := WinExist("A")
    ok := false
    expectedBasename := Editor_GetExpectedRevealBasename(editorHwnd)

    try {
        Editor_EnsureCursorWindowActive(editorHwnd)
        if !Editor_FocusIsInFilesExplorer(editorHwnd)
            Editor_EnsureFilesExplorerSidebarFocused(editorHwnd)

        if (explorerAction = "copy" && EDITOR_COPY_USE_EDITOR_FASTPATH) {
            t0 := A_TickCount
            result := Editor_TryCopyFileFromActiveEditor(editorHwnd, expectedBasename)
            Editor_SmartNav_TimingLog("editor_fastpath", A_TickCount - t0)
            if (result.ok)
                ok := true
        }

        if (!ok) {
            if (!Editor_IsPlausibleRevealBasename(expectedBasename))
                expectedBasename := Editor_GetExpectedRevealBasename(editorHwnd)

            preExplorers := ""
            for hwnd in WinGetList("ahk_exe explorer.exe")
                preExplorers .= hwnd ","

            SendLevel 0
            SendInput "^h"

            if (explorerAction = "copy" || explorerAction = "open" || explorerAction = "share")
                ok := Editor_SmartNavRevealAfterSendH(editorHwnd, explorerAction, expectedBasename, preExplorers)
            else
                ok := true
        }
    } catch {
        ok := false
    }
    if (ok)
        Editor_SmartNavRevealShowSuccess(explorerAction, expectedBasename)
}

; UIA: find Type 50020 text by exact Name under scope. Prefer on-screen; if several, pick bottom-most (largest top Y).
Cursor_FindPermissionText50020(scope, exactName, requireOnScreen := true) {
    try
        all := scope.FindAll({ Type: 50020 })
    catch
        return 0
    best := 0
    bestT := -0x7FFFFFFF
    for t in all {
        try
            nm := t.Name
        catch
            continue
        if (nm != exactName)
            continue
        if (requireOnScreen) {
            try {
                if t.GetPropertyValue(UIA.Property.IsOffscreen)
                    continue
            } catch {
            }
        }
        try {
            br := t.BoundingRectangle
            if (br.t > bestT) {
                bestT := br.t
                best := t
            }
        } catch {
            if (!best)
                best := t
        }
    }
    return best
}

; Same as Cursor_FindPermissionText50020 but Name must contain substring (case-insensitive).
Cursor_FindPermissionText50020Contains(scope, substring, requireOnScreen := true) {
    try
        all := scope.FindAll({ Type: 50020 })
    catch
        return 0
    best := 0
    bestT := -0x7FFFFFFF
    for t in all {
        try
            nm := t.Name
        catch
            continue
        if (!InStr(nm, substring, false))
            continue
        if (requireOnScreen) {
            try {
                if t.GetPropertyValue(UIA.Property.IsOffscreen)
                    continue
            } catch {
            }
        }
        try {
            br := t.BoundingRectangle
            if (br.t > bestT) {
                bestT := br.t
                best := t
            }
        } catch {
            if (!best)
                best := t
        }
    }
    return best
}

; Click permission-style Text (50020): prefer parent Invoke/Click, same pattern as !n "Review next file".
Cursor_ClickUiaTextOrParentInvoke(textEl) {
    try {
        parentBtn := UIA.TreeWalkerTrue.GetParentElement(textEl)
        if (parentBtn) {
            try {
                if parentBtn.GetPropertyValue(UIA.Property.IsInvokePatternAvailable) {
                    parentBtn.InvokePattern.Invoke()
                    return true
                }
            } catch {
            }
            try {
                parentBtn.Click()
                return true
            } catch {
            }
        }
    } catch {
    }
    try {
        if textEl.GetPropertyValue(UIA.Property.IsInvokePatternAvailable) {
            textEl.InvokePattern.Invoke()
            return true
        }
    } catch {
    }
    try {
        textEl.Click()
        return true
    } catch {
    }
    return false
}

; Try each name variant (e.g. straight vs curly quotes). Scope to workbench.parts.panel first, then full window.
Cursor_ClickPermissionLabel(variantNames*) {
    hwnd := WinExist("ahk_exe Cursor.exe")
    if (!hwnd)
        return false
    try
        WinActivate(hwnd)
    catch {
    }
    try
        root := UIA.ElementFromHandle(hwnd)
    catch
        return false
    if (!root)
        return false
    Sleep 100
    scope := root
    try {
        panel := root.FindFirst({ AutomationId: "workbench.parts.panel", Type: 50026 })
        if (panel)
            scope := panel
    } catch {
    }
    el := 0
    for name in variantNames {
        el := Cursor_FindPermissionText50020(scope, name, true)
        if (el)
            break
    }
    if (!el) {
        for name in variantNames {
            el := Cursor_FindPermissionText50020(scope, name, false)
            if (el)
                break
        }
    }
    if (!el) {
        for name in variantNames {
            el := Cursor_FindPermissionText50020(root, name, true)
            if (el)
                break
        }
    }
    if (!el) {
        for name in variantNames {
            el := Cursor_FindPermissionText50020(root, name, false)
            if (el)
                break
        }
    }
    if (!el)
        return false
    return Cursor_ClickUiaTextOrParentInvoke(el)
}

; Like Cursor_ClickPermissionLabel but matches any Type 50020 label containing substring (e.g. "Allowlist").
Cursor_ClickPermissionLabelContains(substring) {
    hwnd := WinExist("ahk_exe Cursor.exe")
    if (!hwnd)
        return false
    try
        WinActivate(hwnd)
    catch {
    }
    try
        root := UIA.ElementFromHandle(hwnd)
    catch
        return false
    if (!root)
        return false
    Sleep 100
    scope := root
    try {
        panel := root.FindFirst({ AutomationId: "workbench.parts.panel", Type: 50026 })
        if (panel)
            scope := panel
    } catch {
    }
    el := Cursor_FindPermissionText50020Contains(scope, substring, true)
    if (!el)
        el := Cursor_FindPermissionText50020Contains(scope, substring, false)
    if (!el)
        el := Cursor_FindPermissionText50020Contains(root, substring, true)
    if (!el)
        el := Cursor_FindPermissionText50020Contains(root, substring, false)
    if (!el)
        return false
    return Cursor_ClickUiaTextOrParentInvoke(el)
}

;-------------------------------------------------------------------
; Cursor IDE — Cursor-only Shortcuts
;-------------------------------------------------------------------
