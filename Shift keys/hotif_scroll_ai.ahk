; =============================================================================
; Shift keys module: hotif_scroll_ai.ahk
; Global Alt+U scroll AI feed and related
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

; Alt + U : Scroll AI feed to bottom — must live outside #HotIf IsEditorActive() so Gemini (Chrome) receives it.
!u::
{
    loadingBarShown := false
    try {
        hwnd := WinExist("A")
        if (!hwnd)
            return
        StandardLoadingBar_Show("⏳ Scrolling AI feed to bottom…", BANNER_ACCENT_INTERMEDIATE, { passive: false,
            centerOnHwnd: hwnd, textWidth: 480, fontSize: 17 })
        loadingBarShown := true

        ; Cheap WinGet first — Gemini (Chrome) skips UIA root/composer scan (efficiency-canon: less COM on hot path).
        proc := ""
        try proc := StrLower(WinGetProcessName("ahk_id " hwnd))
        title := ""
        try title := WinGetTitle("ahk_id " hwnd)
        ; #region agent log
        _dbgL := "C:\Users\eduev\Meu Drive\17 - Projects\scripts\debug-97a80a.log"
        try FileAppend(
            '{"sessionId":"97a80a","hypothesisId":"dispatch","location":"hotif_scroll_ai:!u","message":"dispatch entry","data":{"proc":"' proc '","title":"' StrReplace(
                SubStr(title, 1, 80), '"', "'") '"},"timestamp":' A_TickCount '}' "`n", _dbgL)
        ; #endregion agent log
        if (proc = "chrome.exe" && CopilotWeb_IsCopilotHwnd(hwnd, "fast")) {
            CopilotWeb_ScrollFeedToBottom(hwnd)
            return
        }
        if (proc = "chrome.exe" && GeminiEnterprise_IsEnterpriseHwnd(hwnd, "fast")) {
            GeminiEnterprise_ScrollFeedToBottom(hwnd)
            return
        }
        ; #region agent log
        _isConsGem := (proc = "chrome.exe") ? IsConsumerGeminiChromeTitle(title) : false
        try FileAppend(
            '{"sessionId":"97a80a","hypothesisId":"dispatch","location":"hotif_scroll_ai:gemini-check","message":"consumer gemini check","data":{"isConsumerGemini":' (
                _isConsGem ? "true" : "false") ',"proc":"' proc '"},"timestamp":' A_TickCount '}' "`n", _dbgL)
        ; #endregion agent log
        if (proc = "chrome.exe" && _isConsGem) {
            GeminiScrollFeedToBottom_Chrome(hwnd)
            return
        }
        if (proc = "code.exe") {
            if (VSCodeScrollCopilotFeedToBottom(hwnd))
                return
        }

        root := UIA.ElementFromHandle(hwnd)

        ; FindFirst throws if missing — normal for Chrome (no composer-messages-container).
        chatContainer := 0
        try chatContainer := root.FindFirst({ ClassName: "composer-messages-container" })
        catch
            chatContainer := 0
        if (chatContainer) {
            try {
                if (chatContainer.GetPropertyValue(UIA.Property.IsScrollPatternAvailable)) {
                    chatContainer.ScrollPattern.SetScrollPercent(-1, 100)
                    return
                }
            } catch {
            }

            messages := chatContainer.FindAll({ ClassName: "composer-rendered-message", matchmode: "Substring" })
            if (messages && messages.Length > 0) {
                messages[messages.Length].ScrollIntoView()
            }
            return
        }
    } catch {
    } finally {
        if (loadingBarShown) {
            try StandardLoadingBar_Hide(0)
        }
    }
}

; VS Code (Code.exe): scroll Copilot chat feed to bottom using Chat list container.
VSCodeScrollCopilotFeedToBottom(hwnd) {
    try {
        if (!hwnd)
            return false

        root := UIA.ElementFromHandle(hwnd)
        if (!root)
            return false

        chatList := 0
        try {
            lists := root.FindAll({ Type: 50008 })
            if (lists) {
                for lst in lists {
                    try {
                        nm := ""
                        cls := ""
                        try nm := lst.Name
                        try cls := lst.ClassName
                        if (!InStr(cls, "monaco-list"))
                            continue
                        if (!InStr(nm, "Chat"))
                            continue
                        off := false
                        try off := !!lst.GetPropertyValue(UIA.Property.IsOffscreen)
                        if (off)
                            continue
                        chatList := lst
                        break
                    } catch {
                        continue
                    }
                }
            }
        } catch {
            chatList := 0
        }

        if (!chatList)
            return false

        ; Prefer structural scroll first.
        try {
            if (chatList.GetPropertyValue(UIA.Property.IsScrollPatternAvailable)) {
                try chatList.ScrollPattern.SetScrollPercent(-1, 100)
                try chatList.ScrollPattern.SetScrollPercent(100, -1)
            }
        } catch {
        }

        ; Then anchor on the last chat row and ensure keyboard scroll lands at end.
        try {
            rows := chatList.FindAll({ Type: 50007 })
            if (rows && rows.Length > 0)
                rows[rows.Length].ScrollIntoView()
        } catch {
        }

        try {
            chatList.SetFocus()
            Sleep 40
            Send "{End}"
            Sleep 30
            Send "^{End}"
        } catch {
        }

        return true
    } catch {
        return false
    }
}

VSCode_AddFileToAIContext() {
    StandardLoadingBar_Show("⏳ Add file to VS Code chat...", BANNER_ACCENT_INTERMEDIATE, { passive: false,
        centerOnHwnd: 0 })
    try {
        if (!IsCodeActive()) {
            StandardLoadingBar_Update("❌ Failed: VS Code is not active")
            return
        }

        StandardLoadingBar_Update("⏳ Focusing Explorer...")
        Send "^+e"
        Sleep 350
        okFocus := FocusCursorFilesExplorer()
        if (!okFocus) {
            Sleep 150
            okFocus := FocusCursorFilesExplorer()
        }
        if (!okFocus) {
            Sleep 150
            okFocus := FocusCursorFilesExplorer()
        }
        if (!okFocus) {
            StandardLoadingBar_Update("❌ Failed: Could not focus Explorer sidebar")
            return
        }

        StandardLoadingBar_Update("⏳ Opening context menu...")
        result := VSCode_ContextMenuSelectByDownAndActivateAny(
            ["Add file to chat", "Add File to Chat", "Add file to Chat"],
            "{AppsKey}",
            42
        )

        if (result.ok) {
            StandardLoadingBar_Update("✅ File added to chat")
            return
        }

        StandardLoadingBar_Update("❌ Failed: " . result.reason)
    } finally {
        StandardLoadingBar_Hide(600)
    }
}

VSCode_ContextMenuSelectByDownAndActivateAny(targetTexts, openKey := "{AppsKey}", maxSteps := 36) {
    Send openKey
    Sleep 240

    targetMap := Map()
    for t in targetTexts
        targetMap[StrLower(t)] := true

    step := 0
    while (step <= maxSteps) {
        step += 1
        highlightedEl := VSCode_ContextMenuGetHighlightedElement()
        name := ""
        try name := highlightedEl ? highlightedEl.Name : ""

        if (Mod(step, 3) = 0)
            StandardLoadingBar_Update("⏳ Searching menu item... (" step "/" maxSteps ")")

        if (name != "") {
            nameLower := StrLower(name)
            if (targetMap.Has(nameLower)) {
                StandardLoadingBar_Update("⏳ Activating '" . name . "'...")
                if (VSCode_ContextMenuActivateHighlightedItem(highlightedEl))
                    return { ok: true, reason: "" }
                return { ok: false, reason: "Could not activate menu item" }
            }
        }

        Send "{Down}"
        Sleep 55
    }

    return { ok: false, reason: "Menu item not found" }
}

VSCode_ContextMenuGetHighlightedElement() {
    ; Strategy A: focused element is a MenuItem
    try {
        fe := UIA.GetFocusedElement()
        if (fe) {
            try {
                if (fe.ControlType = UIA.Type.MenuItem)
                    return fe
            } catch {
            }
        }
    } catch {
    }

    ; Strategy B: selected MenuItem in VS Code window
    try {
        hwnd := WinExist("ahk_exe Code.exe")
        if (!hwnd)
            return 0
        root := UIA.ElementFromHandle(hwnd)
        all := root.FindAll({ Type: 50011 })
        for mi in all {
            try {
                if (mi.GetPropertyValue(UIA.Property.IsSelected))
                    return mi
            } catch {
            }
        }
    } catch {
    }

    ; Strategy C: menu item with keyboard focus
    try {
        hwnd := WinExist("ahk_exe Code.exe")
        if (!hwnd)
            return 0
        root := UIA.ElementFromHandle(hwnd)
        all := root.FindAll({ Type: 50011 })
        for mi in all {
            try {
                if (mi.GetPropertyValue(UIA.Property.HasKeyboardFocus))
                    return mi
            } catch {
            }
        }
    } catch {
    }

    return 0
}

VSCode_ContextMenuActivateHighlightedItem(menuItemEl) {
    if (!menuItemEl) {
        Send "{Enter}"
        return true
    }

    try {
        menuItemEl.SetFocus()
        Sleep 40
        Send "{Enter}"
        return true
    } catch {
    }

    try {
        if (menuItemEl.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)) {
            menuItemEl.InvokePattern.Invoke()
            return true
        }
    } catch {
    }

    try {
        menuItemEl.Click()
        return true
    } catch {
    }

    return false
}

Cursor_IsElementVisibleByName(name, hwnd := 0, typeList := "", matchmode := "") {
    return !!Cursor_GetVisibleElementByName(name, hwnd, typeList, matchmode)
}

global g_VSCodeShortcutMenuGui := false
global g_VSCodeShortcutMenuActive := false

ShowVSCodeShortcutMenu() {
    global g_VSCodeShortcutMenuGui, g_VSCodeShortcutMenuActive
    if (g_VSCodeShortcutMenuActive)
        return

    g_VSCodeShortcutMenuGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Owner")
    g_VSCodeShortcutMenuGui.BackColor := "1E1E2E"
    g_VSCodeShortcutMenuGui.MarginX := 20
    g_VSCodeShortcutMenuGui.MarginY := 15

    g_VSCodeShortcutMenuGui.SetFont("s14 cCDD6F4 Bold", "Segoe UI")
    g_VSCodeShortcutMenuGui.Add("Text", "w320 Center", "VS Code quick shortcuts")
    g_VSCodeShortcutMenuGui.Add("Text", "w320 h1 Background45475A")
    g_VSCodeShortcutMenuGui.SetFont("s12 cCDD6F4", "Segoe UI")
    g_VSCodeShortcutMenuGui.Add("Text", "w320 Center", "(empty for future use)")
    g_VSCodeShortcutMenuGui.Add("Text", "w320 h1 Background45475A y+10")
    g_VSCodeShortcutMenuGui.SetFont("s9 c6C7086", "Segoe UI")
    g_VSCodeShortcutMenuGui.Add("Text", "w320 Center", "Press Esc to close")

    activeWin := 0
    try
        activeWin := WinGetID("A")
    catch
        activeWin := 0

    MonitorGetWorkArea(1, &monitorLeft, &monitorTop, &monitorRight, &monitorBottom)
    monitorWidth := monitorRight - monitorLeft
    monitorHeight := monitorBottom - monitorTop
    if (activeWin && activeWin != 0) {
        rect := Buffer(16, 0)
        if (DllCall("GetWindowRect", "ptr", activeWin, "ptr", rect)) {
            centerX := NumGet(rect, 0, "int") + (NumGet(rect, 8, "int") - NumGet(rect, 0, "int")) // 2
            centerY := NumGet(rect, 4, "int") + (NumGet(rect, 12, "int") - NumGet(rect, 4, "int")) // 2
            loop MonitorGetCount() {
                MonitorGetWorkArea(A_Index, &l, &t, &r, &b)
                if (centerX >= l && centerX <= r && centerY >= t && centerY <= b) {
                    monitorLeft := l
                    monitorTop := t
                    monitorWidth := r - l
                    monitorHeight := b - t
                    break
                }
            }
        }
    }

    g_VSCodeShortcutMenuGui.Show("AutoSize Hide")
    g_VSCodeShortcutMenuGui.GetPos(&gx, &gy, &gw, &gh)
    cx := monitorLeft + (monitorWidth - gw) // 2
    cy := monitorTop + (monitorHeight - gh) // 2
    g_VSCodeShortcutMenuGui.Show("x" . cx . " y" . cy . " NA")

    g_VSCodeShortcutMenuActive := true
    Hotkey("Escape", VSCodeShortcutMenu_Cancel, "On")
}

VSCodeShortcutMenu_Cancel(*) {
    VSCodeShortcutMenu_Close()
}

VSCodeShortcutMenu_Close() {
    global g_VSCodeShortcutMenuGui, g_VSCodeShortcutMenuActive
    if (!g_VSCodeShortcutMenuActive)
        return
    g_VSCodeShortcutMenuActive := false
    try Hotkey("Escape", VSCodeShortcutMenu_Cancel, "Off")
    if (IsObject(g_VSCodeShortcutMenuGui) && g_VSCodeShortcutMenuGui.Hwnd) {
        try g_VSCodeShortcutMenuGui.Destroy()
    }
    g_VSCodeShortcutMenuGui := false
}

Cursor_GetVisibleElementByName(name, hwnd := 0, typeList := "", matchmode := "") {
    try {
        element := Cursor_FindElementByName(name, hwnd, typeList, matchmode)
        if !element
            return ""

        isOffscreen := true
        try isOffscreen := element.GetPropertyValue(UIA.Property.IsOffscreen)
        if isOffscreen
            return ""

        return element
    } catch Error {
        return ""
    }
}

Cursor_FindElementByName(name, hwnd := 0, typeList := "", matchmode := "") {
    try {
        if !name
            return ""
        if !hwnd
            hwnd := WinExist("A")
        if !hwnd
            return ""

        root := UIA.ElementFromHandle(hwnd)
        if !root
            return ""

        searchConfigs := []
        types := []
        if (Type(typeList) == "Array") {
            types := typeList
        } else if (typeList) {
            types := [typeList]
        }

        if (Type(types) == "Array" && types.Length) {
            for typeVal in types {
                config := { Name: name }
                if matchmode
                    config.matchmode := matchmode
                config.Type := typeVal
                searchConfigs.Push(config)
            }
        } else {
            config := { Name: name }
            if matchmode
                config.matchmode := matchmode
            searchConfigs.Push(config)
        }

        for config in searchConfigs {
            element := ""
            try element := root.FindElement(config)
            if element
                return element
        }

        return ""
    } catch Error {
        return ""
    }
}

;-------------------------------------------------------------------
; AI Mode and Model Switching Functions for Cursor
;-------------------------------------------------------------------

; Fold all Git directories in the Source Control view by collapsing each Git tree root
FoldAllGitDirectoriesInCursor() {
    try {
        hwnd := WinExist("A")
        if !hwnd
            return
        root := UIA.ElementFromHandle(hwnd)

        Sleep(150)
        Send("^+g")
        Sleep(350)

        ; Narrow to the Source Control (SCM) tree area to avoid unrelated matches
        scmCond := UIA.CreatePropertyConditionEx(UIA.Property.Name, "Source Control", UIA.PropertyConditionFlags
            .IgnoreCaseMatchSubstring
        )
        scmCondPt := UIA.CreatePropertyConditionEx(UIA.Property.Name, "Controle de CÃ³digo", UIA.PropertyConditionFlags
            .IgnoreCaseMatchSubstring
        )
        scmName := UIA.CreateOrCondition(scmCond, scmCondPt)
        scmPaneType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Pane)
        scmGroupType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Group)
        scmTreeType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Tree)
        scmScopeCond := UIA.CreateOrCondition(scmPaneType, UIA.CreateOrCondition(scmGroupType, scmTreeType))
        scmRootCond := UIA.CreateAndCondition(scmName, scmScopeCond)
        scmRoot := root.FindElement(scmRootCond, UIA.TreeScope.Descendants)
        if !scmRoot
            scmRoot := root ; fallback if SCM container not found

        ; Find TreeItem nodes whose Name contains " Git" (case-insensitive)
        nameCond := UIA.CreatePropertyConditionEx(UIA.Property.Name, " Git", UIA.PropertyConditionFlags.IgnoreCaseMatchSubstring
        )
        typeCond := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.TreeItem)
        gitItemCond := UIA.CreateAndCondition(typeCond, nameCond)

        items := scmRoot.FindElements(gitItemCond, UIA.TreeScope.Descendants)
        if !items
            return

        for item in items {
            if !item
                continue
            ; Prefer ExpandCollapsePattern when available
            hasExpand := item.GetPropertyValue(UIA.Property.IsExpandCollapsePatternAvailable)
            if hasExpand {
                try {
                    pat := item.ExpandCollapsePattern
                    state := pat.ExpandCollapseState
                    ; Collapse if not already collapsed
                    if state != UIA.ExpandCollapseState.Collapsed
                        pat.Collapse()
                } catch Error {
                    ; Fallback below if pattern fails
                }
            }
            if !hasExpand {
                ; Fallback: try to find the chevron/button and invoke/click it
                btnType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Button)
                txtType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Text)
                dotName := UIA.CreatePropertyCondition(UIA.Property.Name, ".")
                chevronCond := UIA.CreateOrCondition(btnType, UIA.CreateAndCondition(txtType, dotName))
                chevron := item.FindElement(chevronCond, UIA.TreeScope.Children)
                if !chevron
                    chevron := item.FindElement(chevronCond, UIA.TreeScope.Descendants)
                if chevron {
                    ; If it supports Invoke, prefer it; else click
                    if chevron.GetPropertyValue(UIA.Property.IsInvokePatternAvailable) {
                        try chevron.InvokePattern.Invoke()
                    } else {
                        try chevron.Click()
                    }
                }
            }
            Sleep 50
        }
    } catch Error as e {
        try MsgBox "UIA error folding Git directories: " e.Message, "Cursor Git Fold", "IconX"
    }
}

; Collapse all expandable directories in the Explorer (FileExplorer3) for all workspace roots
FoldAllDirectoriesInExplorer() {
    try {
        ; Show progress overlay immediately (yellow for folding)
        StandardLoadingBar_Show("📁 Folding directories...", BANNER_ACCENT_INTERMEDIATE)

        hwnd := WinExist("A")
        if !hwnd {
            StandardLoadingBar_Hide(0)
            return
        }
        root := UIA.ElementFromHandle(hwnd)

        ; Ensure Explorer is focused if not already
        Send "^+e"
        Sleep 280

        ; Find the Explorer container (EN/PT names) and then the Tree control inside it
        expEn := UIA.CreatePropertyConditionEx(UIA.Property.Name, "Explorer", UIA.PropertyConditionFlags.IgnoreCaseMatchSubstring
        )
        expPt := UIA.CreatePropertyConditionEx(UIA.Property.Name, "Explorador", UIA.PropertyConditionFlags.IgnoreCaseMatchSubstring
        )
        expName := UIA.CreateOrCondition(expEn, expPt)
        paneType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Pane)
        groupType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Group)
        scopeCond := UIA.CreateOrCondition(paneType, groupType)
        expRootCond := UIA.CreateAndCondition(expName, scopeCond)
        expRoot := ""
        try expRoot := root.FindElement(expRootCond, UIA.TreeScope.Descendants)
        if !expRoot
            expRoot := root

        treeType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Tree)
        autoId3 := UIA.CreatePropertyCondition(UIA.Property.AutomationId, "FileExplorer3")
        fileTree := ""
        try fileTree := expRoot.FindElement(autoId3, UIA.TreeScope.Descendants)
        if !fileTree {
            try {
                autoId2 := UIA.CreatePropertyCondition(UIA.Property.AutomationId, "FileExplorer2")
                fileTree := expRoot.FindElement(autoId2, UIA.TreeScope.Descendants)
            }
        }
        if !fileTree {
            try {
                autoId := UIA.CreatePropertyCondition(UIA.Property.AutomationId, "FileExplorer")
                fileTree := expRoot.FindElement(autoId, UIA.TreeScope.Descendants)
            }
        }
        if !fileTree {
            try fileTree := expRoot.FindElement(treeType, UIA.TreeScope.Descendants)
        }

        ; Fallback: Try finding by specific Name "Files Explorer" (common in Cursor/VSCode)
        if !fileTree {
            try {
                feEn := UIA.CreatePropertyCondition(UIA.Property.Name, "Files Explorer")
                fePt := UIA.CreatePropertyCondition(UIA.Property.Name, "Explorador de Arquivos")
                feName := UIA.CreateOrCondition(feEn, fePt)
                feCond := UIA.CreateAndCondition(treeType, feName)
                fileTree := root.FindElement(feCond, UIA.TreeScope.Descendants)
            }
        }

        if !fileTree
            return

        ; Capture currently focused tree item (best effort) to restore selection
        hasFocusProp := UIA.CreatePropertyCondition(UIA.Property.HasKeyboardFocus, true)
        focusedEl := ""
        try focusedEl := fileTree.FindElement(hasFocusProp, UIA.TreeScope.Descendants)
        focusedName := ""
        if focusedEl
            focusedName := focusedEl.GetPropertyValue(UIA.Property.Name)

        ; Preserve scroll position when possible
        hPerc := vPerc := ""
        hasScroll := fileTree.GetPropertyValue(UIA.Property.IsScrollPatternAvailable)
        if hasScroll {
            try {
                sp := fileTree.ScrollPattern
                hPerc := sp.HorizontalScrollPercent
                vPerc := sp.VerticalScrollPercent
            }
        }

        ; Get all TreeItem nodes that support expand/collapse (i.e., directories) AND are currently Expanded
        itemType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.TreeItem)
        canExpand := UIA.CreatePropertyCondition(UIA.Property.IsExpandCollapsePatternAvailable, true)
        isExpanded := UIA.CreatePropertyCondition(UIA.Property.ExpandCollapseExpandCollapseState, UIA.ExpandCollapseState
            .Expanded)

        ; Combine conditions: TreeItem AND CanExpand AND IsExpanded
        dirCond := UIA.CreateAndCondition(itemType, UIA.CreateAndCondition(canExpand, isExpanded))

        ; Loop 3 times to ensure all nested directories are collapsed (slower pacing below)
        loop 3 {
            ; Re-find items each iteration as tree structure may change after collapsing
            items := fileTree.FindElements(dirCond, UIA.TreeScope.Descendants)

            if !items || !items.Length
                break

            ; Collapse each expanded directory.
            for item in items {
                if !item
                    continue
                try {
                    ; Since we filtered by Expanded, we know it's expanded (or was when found)
                    pat := item.ExpandCollapsePattern

                    ; Method 2: UIA Collapse Pattern (Primary method) – slowed down to reduce UI stress
                    try {
                        pat.Collapse()
                        Sleep 90
                    } catch {
                    }

                    ; Check if it worked (only check if we really need to try other methods)
                    if item.ExpandCollapsePattern.ExpandCollapseState != UIA.ExpandCollapseState.Collapsed {
                        ; Method 1: Scroll into view (if needed)
                        try {
                            if item.GetPropertyValue(UIA.Property.IsScrollItemPatternAvailable)
                                item.ScrollItemPattern.ScrollIntoView()
                            pat.Collapse()
                        } catch {
                        }
                    }

                    if item.ExpandCollapsePattern.ExpandCollapseState != UIA.ExpandCollapseState.Collapsed {
                        ; Method 3: Keyboard Navigation – slowed down to reduce UI stress
                        try {
                            if item.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable)
                                item.SelectionItemPattern.Select()
                            else
                                item.SetFocus()
                            Send "{Left}"
                            Sleep 90
                        } catch {
                        }
                    }

                    if item.ExpandCollapsePattern.ExpandCollapseState != UIA.ExpandCollapseState.Collapsed {
                        ; Method 4: Click Chevron
                        try {
                            btnType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Button)
                            txtType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Text)
                            glyphName := UIA.CreatePropertyCondition(UIA.Property.Name, "îª´")
                            dotName := UIA.CreatePropertyCondition(UIA.Property.Name, ".")
                            chevronCond := UIA.CreateOrCondition(btnType, UIA.CreateOrCondition(UIA.CreateAndCondition(
                                txtType,
                                glyphName), UIA.CreateAndCondition(txtType, dotName)))
                            chevron := ""
                            try chevron := item.FindElement(chevronCond, UIA.TreeScope.Children)
                            if !chevron
                                try chevron := item.FindElement(chevronCond, UIA.TreeScope.Descendants)
                            if chevron {
                                if chevron.GetPropertyValue(UIA.Property.IsInvokePatternAvailable) {
                                    chevron.InvokePattern.Invoke()
                                } else {
                                    chevron.Click()
                                }
                            }
                        } catch {
                        }
                    }
                } catch Error as e {
                }
                Sleep 30
            }

            ; Brief pause between iterations to allow UI to update
            Sleep 80
        }

        ; Restore scroll position if it changed
        if hasScroll && (hPerc != "" && vPerc != "") {
            try fileTree.ScrollPattern.SetScrollPercent(hPerc, vPerc)
        }

        ; Restore selection/focus if possible
        if focusedName {
            nameCond := UIA.CreatePropertyConditionEx(UIA.Property.Name, focusedName, UIA.PropertyConditionFlags
                .IgnoreCase
            )
            itemType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.TreeItem)
            focusedLookup := UIA.CreateAndCondition(itemType, nameCond)
            newFocus := ""
            try newFocus := fileTree.FindElement(focusedLookup, UIA.TreeScope.Descendants)
            if newFocus {
                try newFocus.SetFocus()
            } else {
                try fileTree.SetFocus()
            }
        } else {
            try fileTree.SetFocus()
        }

        ; Optional brief toast
        StandardLoadingBar_Update("Directories folded")
    } catch Error as e {
        try MsgBox "UIA error folding Explorer directories: " e.Message, "Cursor Explorer Fold", "IconX"
    } finally {
        StandardLoadingBar_Hide(800)
    }
}

; Expand all expandable directories in the Explorer (FileExplorer3) for all workspace roots.
; No hotkey bound — Ctrl+Q used to call this (Cursor + VS Code); assign a hotkey again if needed.
UnfoldAllDirectoriesInExplorer() {
    try {
        ; Show progress overlay immediately (yellow for unfolding)
        StandardLoadingBar_Show("📁 Unfolding directories...", BANNER_ACCENT_INTERMEDIATE)

        hwnd := WinExist("A")
        if !hwnd {
            StandardLoadingBar_Hide(0)
            return
        }
        root := UIA.ElementFromHandle(hwnd)

        ; Ensure Explorer is focused if not already
        Send "^+e"
        Sleep 150

        ; Find the Explorer container
        expEn := UIA.CreatePropertyConditionEx(UIA.Property.Name, "Explorer", UIA.PropertyConditionFlags.IgnoreCaseMatchSubstring
        )
        expPt := UIA.CreatePropertyConditionEx(UIA.Property.Name, "Explorador", UIA.PropertyConditionFlags.IgnoreCaseMatchSubstring
        )
        expName := UIA.CreateOrCondition(expEn, expPt)
        paneType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Pane)
        groupType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Group)
        scopeCond := UIA.CreateOrCondition(paneType, groupType)
        expRootCond := UIA.CreateAndCondition(expName, scopeCond)
        expRoot := ""
        try expRoot := root.FindElement(expRootCond, UIA.TreeScope.Descendants)
        if !expRoot
            expRoot := root

        treeType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Tree)
        fileTree := ""
        try {
            autoId3 := UIA.CreatePropertyCondition(UIA.Property.AutomationId, "FileExplorer3")
            fileTree := expRoot.FindElement(autoId3, UIA.TreeScope.Descendants)
        }
        if !fileTree {
            try {
                autoId2 := UIA.CreatePropertyCondition(UIA.Property.AutomationId, "FileExplorer2")
                fileTree := expRoot.FindElement(autoId2, UIA.TreeScope.Descendants)
            }
        }
        if !fileTree {
            try {
                autoId := UIA.CreatePropertyCondition(UIA.Property.AutomationId, "FileExplorer")
                fileTree := expRoot.FindElement(autoId, UIA.TreeScope.Descendants)
            }
        }
        if !fileTree {
            try fileTree := expRoot.FindElement(treeType, UIA.TreeScope.Descendants)
        }

        ; Fallback: Try finding by specific Name "Files Explorer"
        if !fileTree {
            try {
                feEn := UIA.CreatePropertyCondition(UIA.Property.Name, "Files Explorer")
                fePt := UIA.CreatePropertyCondition(UIA.Property.Name, "Explorador de Arquivos")
                feName := UIA.CreateOrCondition(feEn, fePt)
                feCond := UIA.CreateAndCondition(treeType, feName)
                fileTree := root.FindElement(feCond, UIA.TreeScope.Descendants)
            }
        }

        if !fileTree {
            StandardLoadingBar_Hide(0)
            return
        }

        ; Preserve scroll position when possible
        hPerc := vPerc := ""
        hasScroll := fileTree.GetPropertyValue(UIA.Property.IsScrollPatternAvailable)
        if hasScroll {
            try {
                sp := fileTree.ScrollPattern
                hPerc := sp.HorizontalScrollPercent
                vPerc := sp.VerticalScrollPercent
            }
        }

        ; Get all TreeItem nodes that support expand/collapse (i.e., directories)
        itemType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.TreeItem)
        canExpand := UIA.CreatePropertyCondition(UIA.Property.IsExpandCollapsePatternAvailable, true)
        dirCond := UIA.CreateAndCondition(itemType, canExpand)

        ; Loop 3 times to ensure nested directories are expanded
        loop 3 {
            items := ""
            try items := fileTree.FindElements(dirCond, UIA.TreeScope.Descendants)
            if !items
                break

            ; Expand each collapsed directory. Do not toggle; skip already expanded.
            for item in items {
                if !item
                    continue
                try {
                    pat := item.ExpandCollapsePattern
                    state := pat.ExpandCollapseState

                    if state == UIA.ExpandCollapseState.Collapsed {
                        ; Method 1: Scroll into view
                        try {
                            if item.GetPropertyValue(UIA.Property.IsScrollItemPatternAvailable)
                                item.ScrollItemPattern.ScrollIntoView()
                        } catch {
                        }

                        ; Method 2: UIA Expand Pattern – slowed down to reduce UI stress
                        try {
                            pat.Expand()
                            Sleep 120
                        } catch {
                        }

                        ; Check if it worked
                        if item.ExpandCollapsePattern.ExpandCollapseState == UIA.ExpandCollapseState.Collapsed {
                            ; Method 3: Keyboard Navigation (Select + Right Arrow) – slowed down to reduce UI stress
                            try {
                                if item.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable)
                                    item.SelectionItemPattern.Select()
                                else
                                    item.SetFocus()
                                Send "{Right}"
                                Sleep 120
                            } catch {
                            }
                        }

                        ; Check if it worked
                        if item.ExpandCollapsePattern.ExpandCollapseState == UIA.ExpandCollapseState.Collapsed {
                            ; Method 4: Click Chevron (Moved from catch block)
                            try {
                                btnType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Button
                                )
                                txtType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Text)
                                glyphName := UIA.CreatePropertyCondition(UIA.Property.Name, "îª´")
                                dotName := UIA.CreatePropertyCondition(UIA.Property.Name, ".")
                                chevronCond := UIA.CreateOrCondition(btnType, UIA.CreateOrCondition(UIA.CreateAndCondition(
                                    txtType,
                                    glyphName), UIA.CreateAndCondition(txtType, dotName)))
                                chevron := ""
                                try chevron := item.FindElement(chevronCond, UIA.TreeScope.Children)
                                if !chevron {
                                    try chevron := item.FindElement(chevronCond, UIA.TreeScope.Descendants)
                                }
                                if chevron {
                                    if chevron.GetPropertyValue(UIA.Property.IsInvokePatternAvailable) {
                                        chevron.InvokePattern.Invoke()
                                    } else {
                                        chevron.Click()
                                    }
                                }
                            } catch {
                            }
                        }
                    }
                } catch Error {
                    ; Fallback in case getting pattern fails completely
                    try {
                        btnType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Button)
                        txtType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Text)
                        glyphName := UIA.CreatePropertyCondition(UIA.Property.Name, "îª´")
                        dotName := UIA.CreatePropertyCondition(UIA.Property.Name, ".")
                        chevronCond := UIA.CreateOrCondition(btnType, UIA.CreateOrCondition(UIA.CreateAndCondition(
                            txtType,
                            glyphName), UIA.CreateAndCondition(txtType, dotName)))
                        chevron := ""
                        try chevron := item.FindElement(chevronCond, UIA.TreeScope.Children)
                        if !chevron {
                            try chevron := item.FindElement(chevronCond, UIA.TreeScope.Descendants)
                        }
                        if chevron {
                            if chevron.GetPropertyValue(UIA.Property.IsInvokePatternAvailable) {
                                chevron.InvokePattern.Invoke()
                            } else {
                                chevron.Click()
                            }
                        }
                    } catch {
                    }
                }
                Sleep 10
            }

            ; Brief pause between iterations to allow UI to update
            Sleep 50
        }

        ; Restore scroll position if it changed
        if hasScroll && (hPerc != "" && vPerc != "") {
            try fileTree.ScrollPattern.SetScrollPercent(hPerc, vPerc)
        }

        ; Optional brief toast
        StandardLoadingBar_Update("Directories unfolded")
    } catch Error as e {
        try MsgBox "UIA error unfolding Explorer directories: " e.Message, "Cursor Explorer Unfold", "IconX"
    } finally {
        StandardLoadingBar_Hide(800)
    }
}

; Helper: detect on-screen Text elements for "Agent"/"Ask" and send Ctrl+I/L
HasTextByRegex(pattern) {
    try {
        hwnd := WinExist("A")
        if !hwnd
            return false
        root := UIA.ElementFromHandle(hwnd)
        if !IsObject(root)
            return false
        for el in root.FindAll({ Type: "Text" }) {
            if RegExMatch(el.Name, pattern)
                return true
        }
    } catch Error as e {
        ; ignore and fall through
    }
    return false
}

SendCtrlKeyBasedOnAgentAsk() {
    ; Returns true if a key was sent, false otherwise
    if HasTextByRegex("i)\\bagent\\b") {
        Send "{Ctrl down}i{Ctrl up}"
        return true
    }
    if HasTextByRegex("i)ask") {
        Send "{Ctrl down}l{Ctrl up}"
        return true
    }
    return false
}

; Function to switch between AI modes (agent/ask)
SwitchAIMode() {
    try {
        ; Get user input directly
        userChoice := InputBox("Choose AI Mode:`n`n1. ask`n2. agent`n`nEnter choice (1 or 2):",
            "AI Mode Selection",
            "w250 h150")
        if userChoice.Result != "OK"
            return

        ; Determine the mode string based on choice
        modeString := ""
        switch userChoice.Value {
            case "1":
                modeString := "ask"
            case "2":
                modeString := "agent"
            default:
                MsgBox "Invalid selection. Please choose 1 or 2.", "AI Mode Selection", "IconX"
                return
        }

        ; Send Escape twice, then select the edit field based on on-screen Agent/Ask
        SendEscape(2)
        Sleep 200
        if !SendCtrlKeyBasedOnAgentAsk() {
            ; Fallback to Ctrl+I if no relevant text is found
            Send "{Ctrl down}i{Ctrl up}"
        }
        Sleep 300

        ; Send Ctrl+. and wait for context menu
        Send "^."
        Sleep 500  ; Wait for context menu to appear

        ; Type the selected mode string
        SendText modeString
        Sleep 100

        ; Press Enter to confirm
        Send "{Enter}"

    } catch Error as e {
        MsgBox "Error switching AI mode: " e.Message, "AI Mode Switch Error", "IconX"
    }
}

; Function to switch between AI models
SwitchAIModel() {
    try {
        ; Get user input directly
        userChoice := InputBox(
            "Choose AI Model:`n`n1. auto`n2. CLAUD`n3. GPT`n4. O`n5. DeepSeek`n6. Cursor`n`nEnter choice (1-6):",
            "AI Model Selection", "w250 h200")
        if userChoice.Result != "OK"
            return

        ; Send Escape twice, then select the edit field based on on-screen Agent/Ask
        SendEscape(2)
        Sleep 200
        if !SendCtrlKeyBasedOnAgentAsk() {
            ; Fallback to Ctrl+I if no relevant text is found
            Send "{Ctrl down}i{Ctrl up}"
        }
        Sleep 300

        ; Handle different behaviors based on choice
        switch userChoice.Value {
            case "1":
            {
                ; For auto option: simulate ;, wait for model context menu, then send â†" , Enter
                Send "^;"
                Sleep 300
                SendText "auto"
                Sleep 500
                Send "{Enter}"
                Sleep 300
                SendEscape()
            }
            case "2":
            {
                ; For other options: simulate Ctrl + ., wait, type model string, no Enter
                Send "^;"
                Sleep 500
                SendText "CLAUD"
            }
            case "3":
            {
                Send "^;"
                Sleep 500
                SendText "GPT"
            }
            case "4":
            {
                Send "^;"
                Sleep 500
                SendText "O"
            }
            case "5":
            {
                Send "^;"
                Sleep 500
                SendText "DeepSeek"
            }
            case "6":
            {
                Send "^;"
                Sleep 500
                SendText "Cursor"
            }
            default:
                MsgBox "Invalid selection. Please choose 1-6.", "AI Model Selection", "IconX"
                return
        }

        Sleep 100
        Send "{Enter}"

    } catch Error as e {
        MsgBox "Error switching AI model: " e.Message, "AI Model Switch Error", "IconX"
    }
}
