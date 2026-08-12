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

; Delay after stash QuickInput is ready, before Enter (empty stash message).
global EDITOR_GIT_STASH_STEP_MS := 400
global EDITOR_GIT_FLOW_MAX_MS := 45000
global EDITOR_GIT_STEP_TIMEOUT_MS := 12000
global EDITOR_GIT_PULL_TIMEOUT_MS := 15000
global g_EditorGitFlowDeadline := 0
global g_EditorGitDidStashThisFlow := false
global EDITOR_GIT_CMD_STASH := "Git: Stash"
global EDITOR_GIT_CMD_STASH_UNTRACKED := "Git: Stash (Include Untracked)"
global EDITOR_GIT_CMD_STASH_POP := "Git: Stash Pop"

Editor_GitUiaRoot(editorHwnd := 0) {
    try {
        if !editorHwnd
            editorHwnd := WinExist("A")
        if !editorHwnd
            return 0
        return UIA.ElementFromHandle(editorHwnd)
    } catch {
        return 0
    }
}

Editor_GitPollSleep(startTick) {
    Sleep (A_TickCount - startTick < 2000 ? 50 : 200)
}

Editor_GitPollUntil(editorHwnd, timeoutMs, callback, pollLabel := "") {
    global g_EditorGitFlowDeadline
    startTick := A_TickCount
    deadline := g_EditorGitFlowDeadline ? Min(startTick + timeoutMs, g_EditorGitFlowDeadline) : (startTick + timeoutMs)
    lastBarUpdate := 0
    while (A_TickCount < deadline) {
        if Editor_GitFlowWatchdogExpired()
            return false
        root := Editor_GitUiaRoot(editorHwnd)
        if (root && callback(root))
            return true
        if (pollLabel != "" && (A_TickCount - lastBarUpdate >= 1000)) {
            elapsed := Round((A_TickCount - startTick) / 1000)
            try StandardLoadingBar_Update(pollLabel " (" elapsed "s)", BANNER_ACCENT_INTERMEDIATE)
            lastBarUpdate := A_TickCount
        }
        Editor_GitPollSleep(startTick)
    }
    return false
}

Editor_TextHasWorkingTreeBlocker(text) {
    return RegExMatch(text,
        "i)(clean your repository working tree|commit your changes or stash|would be overwritten by merge)")
}

Editor_HasGitWorkingTreeBlocker(editorHwnd := 0, root := 0) {
    try {
        if !root
            root := Editor_GitUiaRoot(editorHwnd)
        if !root
            return false
        for el in root.FindAll({ Type: UIA.Type.Text }) {
            try {
                name := el.Name
                if (name = "")
                    continue
                if Editor_TextHasWorkingTreeBlocker(name)
                    return true
            } catch {
            }
        }
    } catch {
    }
    return false
}

Editor_GetScmSyncStatusNameFromRoot(root) {
    if !root
        return ""
    try {
        el := root.FindFirst({ AutomationId: "status.scm.1" })
        if !el
            return ""
        return el.Name
    } catch {
        return ""
    }
}

Editor_GetScmSyncStatusName(editorHwnd := 0) {
    return Editor_GetScmSyncStatusNameFromRoot(Editor_GitUiaRoot(editorHwnd))
}

Editor_ParsePullBehindCount(syncName := "") {
    if (syncName = "")
        syncName := Editor_GetScmSyncStatusName()
    if RegExMatch(syncName, "i)Pull\s+(\d+)\s+commit", &m)
        return Integer(m[1])
    return 0
}

Editor_IsGitPullPending(syncName := "") {
    if (syncName = "")
        syncName := Editor_GetScmSyncStatusName()
    return RegExMatch(syncName, "i)Pull\s+\d+\s+commit")
}

Editor_GetScmPendingChangesCountFromRoot(root) {
    if !root
        return -1
    try {
        el := root.FindFirst({ Type: UIA.Type.TabItem, Name: "Source Control", matchmode: "Substring" })
        if el {
            name := el.Name
            if RegExMatch(name, "i)(\d+)\s+pending\s+change", &m)
                return Integer(m[1])
            return 0
        }
    } catch {
    }
    return -1
}

Editor_GetScmPendingChangesCount(editorHwnd := 0) {
    return Editor_GetScmPendingChangesCountFromRoot(Editor_GitUiaRoot(editorHwnd))
}

Editor_FindQuickInputFromRoot(root) {
    if !root
        return 0
    try {
        el := root.FindFirst({ ClassName: "quick-input-widget", matchmode: "Substring" })
        if el
            return el
    } catch {
    }
    return 0
}

Editor_QuickInputHasFocusedEdit(quickInputEl) {
    if !quickInputEl
        return false
    try {
        edit := quickInputEl.FindFirst({ Type: UIA.Type.Edit })
        if edit {
            try {
                if edit.GetPropertyValue(UIA.Property.HasKeyboardFocus)
                    return true
            } catch {
            }
        }
    } catch {
    }
    return false
}

Editor_IsQuickInputOpenFromRoot(root) {
    return !!Editor_FindQuickInputFromRoot(root)
}

Editor_IsQuickInputOpen(editorHwnd := 0) {
    return Editor_IsQuickInputOpenFromRoot(Editor_GitUiaRoot(editorHwnd))
}

Editor_WaitForQuickInputClosed(editorHwnd := 0, timeoutMs := 4000) {
    return Editor_GitPollUntil(editorHwnd, timeoutMs, (root) => !Editor_IsQuickInputOpenFromRoot(root))
}

Editor_WaitForStashQuickInput(editorHwnd := 0, timeoutMs := 4000) {
    return Editor_GitPollUntil(editorHwnd, timeoutMs, (root) => (
        (qi := Editor_FindQuickInputFromRoot(root)) && Editor_QuickInputHasFocusedEdit(qi)))
}

Editor_IsGitSyncInProgressFromRoot(root) {
    if !root
        return false
    try {
        syncName := Editor_GetScmSyncStatusNameFromRoot(root)
        if RegExMatch(syncName, "i)(fetching|pulling|syncing)")
            return true
        el := root.FindFirst({ ClassName: "monaco-status", matchmode: "Substring" })
        if el {
            try {
                if RegExMatch(el.Name, "i)(fetching|pulling|syncing)")
                    return true
            } catch {
            }
        }
        el := root.FindFirst({ ClassName: "statusbar-item", matchmode: "Substring" })
        if el {
            try {
                if RegExMatch(el.Name, "i)(fetching|pulling|syncing)")
                    return true
            } catch {
            }
        }
    } catch {
    }
    return false
}

Editor_IsGitSyncInProgress(editorHwnd := 0) {
    return Editor_IsGitSyncInProgressFromRoot(Editor_GitUiaRoot(editorHwnd))
}

Editor_HasGitErrorAlertFromRoot(root) {
    if !root
        return false
    try {
        for el in root.FindAll({ Type: UIA.Type.Text }) {
            try {
                if !InStr(el.ClassName, "monaco-alert")
                    continue
                name := el.Name
                if Editor_TextHasWorkingTreeBlocker(name)
                    continue
                if RegExMatch(name, "i)(error|fatal|conflict|failed|authentication|permission|denied)")
                    return true
            } catch {
            }
        }
    } catch {
    }
    return false
}

Editor_HasGitErrorAlert(editorHwnd := 0) {
    return Editor_HasGitErrorAlertFromRoot(Editor_GitUiaRoot(editorHwnd))
}

Editor_GitFlowWatchdogExpired() {
    global g_EditorGitFlowDeadline
    return g_EditorGitFlowDeadline && (A_TickCount > g_EditorGitFlowDeadline)
}

Editor_GitFlowSnapshot(editorHwnd := 0) {
    if !editorHwnd
        editorHwnd := WinExist("A")
    root := Editor_GitUiaRoot(editorHwnd)
    syncName := Editor_GetScmSyncStatusNameFromRoot(root)
    return Map(
        "syncName", syncName,
        "pullBehind", Editor_ParsePullBehindCount(syncName),
        "pendingChanges", Editor_GetScmPendingChangesCountFromRoot(root),
        "quickInputOpen", Editor_IsQuickInputOpenFromRoot(root),
        "gitBusy", Editor_IsGitSyncInProgressFromRoot(root)
    )
}

Editor_WaitForGitOperationIdle(editorHwnd := 0, timeoutMs := 12000, &failReason := "", pollLabel := "") {
    ok := Editor_GitPollUntil(editorHwnd, timeoutMs, (root) => (
        !Editor_IsQuickInputOpenFromRoot(root) && !Editor_IsGitSyncInProgressFromRoot(root)
        && !Editor_HasGitErrorAlertFromRoot(root)), pollLabel)
    if ok
        return true
    if Editor_GitFlowWatchdogExpired() {
        failReason := "timed out (overall)"
        return false
    }
    if Editor_HasGitErrorAlert(editorHwnd) {
        failReason := "git error alert"
        return false
    }
    failReason := "operation idle timeout"
    return false
}

Editor_EnsureQuickInputClosed(editorHwnd := 0, timeoutMs := 2000) {
    if !Editor_IsQuickInputOpen(editorHwnd)
        return true
    Send "{Escape}"
    return Editor_WaitForQuickInputClosed(editorHwnd, timeoutMs)
}

Editor_GetQuickInputPickItemsFromRoot(root) {
    items := []
    if !root
        return items
    qi := Editor_FindQuickInputFromRoot(root)
    if !qi
        return items
    try {
        tree := qi.FindFirst({ ClassName: "monaco-list", matchmode: "Substring" })
        if tree {
            for el in tree.FindAll({ Type: UIA.Type.TreeItem }) {
                try {
                    if el.Name
                        items.Push(el)
                } catch {
                }
            }
        }
    } catch {
    }
    if items.Length
        return items
    for type in [UIA.Type.TreeItem, UIA.Type.ListItem] {
        try {
            for el in qi.FindAll({ Type: type }) {
                try {
                    if el.Name
                        items.Push(el)
                } catch {
                }
            }
        } catch {
        }
    }
    return items
}

Editor_GetQuickInputSelectedPickItem(root) {
    for el in Editor_GetQuickInputPickItemsFromRoot(root) {
        try {
            if el.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable)
            && el.SelectionItemPattern.IsSelected
                return el
        } catch {
        }
    }
    return 0
}

Editor_FindQuickInputCommandItem(root, exactName) {
    if !root || exactName = ""
        return 0
    qi := Editor_FindQuickInputFromRoot(root)
    searchRoots := []
    if qi
        searchRoots.Push(qi)
    searchRoots.Push(root)
    for sr in searchRoots {
        for type in [UIA.Type.TreeItem, UIA.Type.ListItem] {
            try {
                el := sr.FindFirst({ Name: exactName, Type: type, matchmode: "Exact", cs: false })
                if el
                    return el
            } catch {
            }
        }
    }
    for el in Editor_GetQuickInputPickItemsFromRoot(root) {
        try {
            if (el.Name = exactName)
                return el
        } catch {
        }
    }
    return 0
}

Editor_CountQuickInputExactMatches(root, exactName) {
    count := 0
    for el in Editor_GetQuickInputPickItemsFromRoot(root) {
        try {
            if (el.Name = exactName)
                count += 1
        } catch {
        }
    }
    return count
}

Editor_CommandPaletteSelectionSucceeded(editorHwnd, commandText) {
    if !Editor_IsQuickInputOpen(editorHwnd)
        return true
    root := Editor_GitUiaRoot(editorHwnd)
    qi := Editor_FindQuickInputFromRoot(root)
    if qi && Editor_QuickInputHasFocusedEdit(qi) {
        filter := ""
        try filter := qi.FindFirst({ Type: UIA.Type.Edit }).Value
        catch {
            try filter := qi.FindFirst({ Type: UIA.Type.Edit }).Name
            catch filter := ""
        }
        if (filter != commandText)
            return true
    }
    return false
}

Editor_GitActivateQuickPickItem(el) {
    if !el
        return false
    try {
        if el.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable) {
            el.SelectionItemPattern.Select()
            Sleep 60
            Send "{Enter}"
            return true
        }
    } catch {
    }
    try {
        el.Click()
        Sleep 60
        Send "{Enter}"
        return true
    } catch {
    }
    return Editor_GitActivateElement(el)
}

Editor_GitActivateElement(el) {
    if !el
        return false
    try {
        if el.GetPropertyValue(UIA.Property.IsInvokePatternAvailable) {
            el.InvokePattern.Invoke()
            return true
        }
    } catch {
    }
    try {
        el.Click()
        return true
    } catch {
    }
    try {
        el.SetFocus()
        Sleep 40
        Send "{Enter}"
        return true
    } catch {
    }
    return false
}

; Exact Name match — required when both "Git: Stash" and "Git: Stash (Include Untracked)" appear.
Editor_FindQuickInputListItemExact(root, exactName) {
    return Editor_FindQuickInputCommandItem(root, exactName)
}

Editor_RunCommandPaletteGitCommand(commandText, editorHwnd := 0, &failReason := "", requireExactPick := false) {
    if !editorHwnd
        editorHwnd := WinExist("A")
    Editor_EnsureQuickInputClosed(editorHwnd)
    Send "^+p"
    if !Editor_WaitForStashQuickInput(editorHwnd, 4000) {
        Sleep 600
        if !Editor_WaitForStashQuickInput(editorHwnd, 2000) {
            failReason := "command palette did not open"
            return false
        }
    }
    Sleep 150
    SendText commandText
    Sleep 120
    startTick := A_TickCount
    deadline := A_TickCount + 5000
    while (A_TickCount < deadline) {
        root := Editor_GitUiaRoot(editorHwnd)
        if root {
            el := Editor_FindQuickInputCommandItem(root, commandText)
            if el && Editor_GitActivateQuickPickItem(el) {
                Sleep 120
                if Editor_CommandPaletteSelectionSucceeded(editorHwnd, commandText)
                    return true
            }
            sel := Editor_GetQuickInputSelectedPickItem(root)
            if sel {
                try selName := sel.Name
                catch selName := ""
                    if (selName = commandText) {
                        Send "{Enter}"
                        Sleep 120
                        if Editor_CommandPaletteSelectionSucceeded(editorHwnd, commandText)
                            return true
                    }
            }
            if !requireExactPick && Editor_CountQuickInputExactMatches(root, commandText) = 1 {
                Send "{Enter}"
                Sleep 120
                if Editor_CommandPaletteSelectionSucceeded(editorHwnd, commandText)
                    return true
            }
        }
        Editor_GitPollSleep(startTick)
    }
    failReason := "command not found in palette"
    try Send "{Escape}"
    return false
}

Editor_CompleteStashMessageDialog(editorHwnd, &failReason := "") {
    global EDITOR_GIT_STASH_STEP_MS, EDITOR_GIT_STEP_TIMEOUT_MS
    if !Editor_WaitForStashQuickInput(editorHwnd, 3000) {
        Sleep 600
        if !Editor_WaitForStashQuickInput(editorHwnd, 2000)
            return true
    }
    Sleep EDITOR_GIT_STASH_STEP_MS
    Send "{Enter}"
    if !Editor_WaitForQuickInputClosed(editorHwnd, EDITOR_GIT_STEP_TIMEOUT_MS) {
        failReason := "stash dialog still open"
        return false
    }
    return Editor_WaitForGitOperationIdle(editorHwnd, EDITOR_GIT_STEP_TIMEOUT_MS, &failReason, "⏳ Stashing changes")
}

Editor_RunGitStash(editorHwnd, commandText, &failReason := "") {
    global g_EditorGitDidStashThisFlow
    if !Editor_RunCommandPaletteGitCommand(commandText, editorHwnd, &failReason, true)
        return false
    g_EditorGitDidStashThisFlow := true
    return Editor_CompleteStashMessageDialog(editorHwnd, &failReason)
}

Editor_RunStashIncludeUntracked(editorHwnd, &failReason := "") {
    global EDITOR_GIT_CMD_STASH_UNTRACKED
    return Editor_RunGitStash(editorHwnd, EDITOR_GIT_CMD_STASH_UNTRACKED, &failReason)
}

Editor_GitFlowFail(step, reason := "") {
    msg := "❌ " step " failed"
    if (reason != "")
        msg .= ": " reason
    try StandardLoadingBar_Update(msg, BANNER_ACCENT_ERROR)
    try StandardLoadingBar_Hide(1200)
}

Editor_GitGateStashVerify(editorHwnd, pendingBefore) {
    if Editor_IsQuickInputOpen(editorHwnd)
        return false
    if Editor_HasGitWorkingTreeBlocker(editorHwnd)
        return false
    if (pendingBefore > 0) {
        pendingNow := Editor_GetScmPendingChangesCount(editorHwnd)
        if (pendingNow >= 0 && pendingNow < pendingBefore)
            return true
        if (pendingNow == 0)
            return true
        return false
    }
    return true
}

Editor_GitGateStash(editorHwnd, beforeSnapshot, &failReason := "") {
    global EDITOR_GIT_STEP_TIMEOUT_MS, EDITOR_GIT_CMD_STASH, EDITOR_GIT_CMD_STASH_UNTRACKED
    if Editor_GitFlowWatchdogExpired() {
        failReason := "timed out (overall)"
        return false
    }
    pendingBefore := beforeSnapshot["pendingChanges"]
    try StandardLoadingBar_Update("⏳ Stashing changes…", BANNER_ACCENT_INTERMEDIATE)
    if !Editor_RunGitStash(editorHwnd, EDITOR_GIT_CMD_STASH, &failReason)
        return false
    verified := Editor_GitPollUntil(editorHwnd, EDITOR_GIT_STEP_TIMEOUT_MS, (root) => (
        !Editor_IsQuickInputOpenFromRoot(root)
        && !Editor_HasGitWorkingTreeBlocker(editorHwnd, root)
        && (pendingBefore <= 0
            || (pc := Editor_GetScmPendingChangesCountFromRoot(root)) >= 0 && (pc < pendingBefore || pc == 0))),
    "⏳ Verifying stash")
    if (!verified || !Editor_GitGateStashVerify(editorHwnd, pendingBefore)) && Editor_HasGitWorkingTreeBlocker(
        editorHwnd) {
        try StandardLoadingBar_Update("⏳ Stashing untracked changes…", BANNER_ACCENT_INTERMEDIATE)
        if !Editor_RunGitStash(editorHwnd, EDITOR_GIT_CMD_STASH_UNTRACKED, &failReason)
            return false
        verified := Editor_GitPollUntil(editorHwnd, EDITOR_GIT_STEP_TIMEOUT_MS, (root) => (
            !Editor_IsQuickInputOpenFromRoot(root)
            && !Editor_HasGitWorkingTreeBlocker(editorHwnd, root)
            && (pendingBefore <= 0
                || (pc := Editor_GetScmPendingChangesCountFromRoot(root)) >= 0 && (pc < pendingBefore || pc == 0))),
        "⏳ Verifying stash")
    }
    if verified && Editor_GitGateStashVerify(editorHwnd, pendingBefore)
        return true
    if Editor_HasGitWorkingTreeBlocker(editorHwnd) {
        failReason := "working tree still dirty"
        return false
    }
    failReason := failReason != "" ? failReason : "stash verification timeout"
    return false
}

Editor_GitGateFetchOnce(editorHwnd, beforeSnapshot, &failReason := "") {
    global EDITOR_GIT_STEP_TIMEOUT_MS
    syncBefore := beforeSnapshot["syncName"]
    try StandardLoadingBar_Update("⏳ Fetching from remote…", BANNER_ACCENT_INTERMEDIATE)
    if !Editor_RunCommandPaletteGitCommand("Git: Fetch", editorHwnd, &failReason)
        return false
    if !Editor_WaitForQuickInputClosed(editorHwnd, 4000) {
        failReason := "command palette did not close"
        return false
    }
    fetchStart := A_TickCount
    if !Editor_WaitForGitOperationIdle(editorHwnd, EDITOR_GIT_STEP_TIMEOUT_MS, &failReason, "⏳ Waiting for fetch") {
        return false
    }
    syncAfter := Editor_GetScmSyncStatusName(editorHwnd)
    if (syncAfter != syncBefore)
        return true
    if ((A_TickCount - fetchStart) >= 500) && !Editor_HasGitErrorAlert(editorHwnd)
        return true
    failReason := "fetch did not complete"
    return false
}

Editor_GitGateFetch(editorHwnd, beforeSnapshot, &failReason := "") {
    if Editor_GitGateFetchOnce(editorHwnd, beforeSnapshot, &failReason)
        return true
    if Editor_GitFlowWatchdogExpired() {
        failReason := "timed out (overall)"
        return false
    }
    try StandardLoadingBar_Update("⏳ Retrying fetch…", BANNER_ACCENT_INTERMEDIATE)
    return Editor_GitGateFetchOnce(editorHwnd, beforeSnapshot, &failReason)
}

Editor_GitRunPullCommand(editorHwnd, &failReason := "") {
    global EDITOR_GIT_PULL_TIMEOUT_MS
    if !Editor_RunCommandPaletteGitCommand("Git: Pull", editorHwnd, &failReason)
        return false
    if !Editor_WaitForQuickInputClosed(editorHwnd, 4000) {
        failReason := "command palette did not close"
        return false
    }
    return Editor_WaitForGitOperationIdle(editorHwnd, EDITOR_GIT_PULL_TIMEOUT_MS, &failReason, "⏳ Waiting for pull")
}

Editor_GitPullOutcomeOk(pullBefore, editorHwnd) {
    if (pullBefore <= 0)
        return true
    pullAfter := Editor_ParsePullBehindCount(Editor_GetScmSyncStatusName(editorHwnd))
    if (pullAfter == 0 || !Editor_IsGitPullPending())
        return true
    if (pullAfter < pullBefore)
        return true
    return false
}

Editor_GitGatePullOnce(editorHwnd, &failReason := "", &didRecovery := false) {
    pullBefore := Editor_ParsePullBehindCount(Editor_GetScmSyncStatusName(editorHwnd))
    if (!pullBefore && Editor_IsGitPullPending())
        pullBefore := 1
    try StandardLoadingBar_Update("⏳ Pulling from remote…", BANNER_ACCENT_INTERMEDIATE)
    if !Editor_GitRunPullCommand(editorHwnd, &failReason)
        return false
    if Editor_HasGitWorkingTreeBlocker(editorHwnd) {
        didRecovery := true
        try StandardLoadingBar_Update("⏳ Stashing untracked changes before pull…", BANNER_ACCENT_INTERMEDIATE)
        if !Editor_RunStashIncludeUntracked(editorHwnd, &failReason)
            return false
        try StandardLoadingBar_Update("⏳ Retrying pull…", BANNER_ACCENT_INTERMEDIATE)
        if !Editor_GitRunPullCommand(editorHwnd, &failReason)
            return false
    }
    if Editor_HasGitErrorAlert(editorHwnd) {
        failReason := "git error alert"
        return false
    }
    if Editor_HasGitWorkingTreeBlocker(editorHwnd) {
        failReason := "working tree still dirty"
        return false
    }
    if Editor_GitPullOutcomeOk(pullBefore, editorHwnd)
        return true
    failReason := "pull count unchanged"
    return false
}

Editor_GitGatePull(editorHwnd, &failReason := "") {
    didRecovery := false
    if Editor_GitGatePullOnce(editorHwnd, &failReason, &didRecovery)
        return true
    if didRecovery || Editor_GitFlowWatchdogExpired() {
        if Editor_GitFlowWatchdogExpired()
            failReason := "timed out (overall)"
        return false
    }
    try StandardLoadingBar_Update("⏳ Retrying pull…", BANNER_ACCENT_INTERMEDIATE)
    return Editor_GitGatePullOnce(editorHwnd, &failReason, &didRecovery)
}

Editor_GitPopStashAfterSuccess(editorHwnd) {
    global g_EditorGitDidStashThisFlow, EDITOR_GIT_CMD_STASH_POP
    if !g_EditorGitDidStashThisFlow
        return
    try StandardLoadingBar_Update("⏳ Restoring stashed changes…", BANNER_ACCENT_INTERMEDIATE)
    if !Editor_RunCommandPaletteGitCommand(EDITOR_GIT_CMD_STASH_POP, editorHwnd)
        return
    Editor_WaitForQuickInputClosed(editorHwnd, 4000)
    popReason := ""
    Editor_WaitForGitOperationIdle(editorHwnd, 8000, &popReason, "⏳ Applying stash pop")
    if Editor_HasGitErrorAlert(editorHwnd) || Editor_HasGitWorkingTreeBlocker(editorHwnd) {
        try ShowCenteredOverlay_Utils("ℹ Stashed changes remain on stack", 2200, BANNER_ACCENT_INFO)
    }
    g_EditorGitDidStashThisFlow := false
}

Editor_GitStashAndPull() {
    global g_EditorGitFlowDeadline, EDITOR_GIT_FLOW_MAX_MS, g_EditorGitDidStashThisFlow
    hwnd := WinExist("A")
    if !hwnd
        return
    g_EditorGitDidStashThisFlow := false
    g_EditorGitFlowDeadline := A_TickCount + EDITOR_GIT_FLOW_MAX_MS
    before := Editor_GitFlowSnapshot(hwnd)
    failReason := ""
    StandardLoadingBar_Show("⏳ Git stash, fetch, and pull…", BANNER_ACCENT_INTERMEDIATE, { passive: false, centerOnHwnd: hwnd
    })
    try {
        if !Editor_GitGateStash(hwnd, before, &failReason) {
            Editor_GitFlowFail("Stash", failReason)
            return
        }
        if !Editor_GitGateFetch(hwnd, before, &failReason) {
            Editor_GitFlowFail("Fetch", failReason)
            return
        }
        if !Editor_GitGatePull(hwnd, &failReason) {
            Editor_GitFlowFail("Pull", failReason)
            return
        }
        Editor_GitPopStashAfterSuccess(hwnd)
        StandardLoadingBar_Update("✅ Pull complete", BANNER_ACCENT_SUCCESS)
        try {
            soundPath := A_ScriptDir . "\assets\sounds\pull-successful.wav"
            if FileExist(soundPath)
                ScriptSoundPlay(soundPath, true)
        } catch {
        }
        StandardLoadingBar_Hide(600)
    } catch {
        try StandardLoadingBar_Hide(0)
        catch {
        }
    } finally {
        g_EditorGitFlowDeadline := 0
        g_EditorGitDidStashThisFlow := false
    }
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
