; =============================================================================
; Shift keys module: hotif_uia_tree.ahk
; UIA Tree Inspector hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

#HotIf WinActive("UIATreeInspector") || WinActive("ahk_exe UIATreeInspectorAutoHotkey64.exe")

; Shift + R : Refresh list
+r:: {
    try {
        root := UIA.ElementFromHandle(WinExist("A"))
        Sleep 200
        btn := root.FindFirst({ Name: "Refresh list", Type: "Button" })
        if !btn
            btn := root.FindFirst({ AutomationId: "5", Type: "Button" })
        if btn {
            btn.Invoke()
        } else {
            MsgBox "Could not find the Refresh list button.", "UIA Tree Inspector", "IconX"
        }
    } catch Error as e {
        MsgBox "Error refreshing list:`n" e.Message, "UIA Tree Inspector", "IconX"
    }
}

; Shift + F : Focus filter field
+f:: {
    try {
        root := UIA.ElementFromHandle(WinExist("A"))
        Sleep 200
        ; Find the "Filter:" text element
        filterText := root.FindFirst({ Name: "Filter:", Type: "Text", AutomationId: "18" })
        if filterText {
            ; Focus on the text element
            filterText.SetFocus()
            Sleep 100

            ; Hit Tab once
            Send "{Tab}"
            Sleep 50
        } else {
            MsgBox "Could not find the 'Filter:' text element.", "Text Focus", "IconX"
        }
    } catch Error as e {
        MsgBox "Error focusing button and performing Shift+Tab sequence:`n" e.Message, "Button Focus", "IconX"
    }
}

; Shift + S : Select tree item by name prefix
+s:: {
    try {
        ; Global variable to store user input
        global g_TreeItemSearchInput := ""

        ; Create GUI dialog
        searchGui := Gui("+AlwaysOnTop +ToolWindow", "Select Tree Item")
        searchGui.SetFont("s10", "Segoe UI")

        ; Add instruction text
        searchGui.AddText("w350 Center", "Enter text to search for tree item (starts with):")

        ; Add text input field
        searchGui.AddEdit("w300 Center vTreeItemInput")

        ; Submit handler
        SubmitTreeItemSearch(ctrl, *) {
            global g_TreeItemSearchInput
            g_TreeItemSearchInput := ctrl.Gui["TreeItemInput"].Text
            ctrl.Gui.Destroy()
        }

        ; Cancel handler
        CancelTreeItemSearch(ctrl, *) {
            global g_TreeItemSearchInput
            g_TreeItemSearchInput := ""
            ctrl.Gui.Destroy()
        }

        ; Add OK and Cancel buttons (OK is default, triggered by Enter)
        okBtn := searchGui.AddButton("w80 xp-40 y+10 Default", "OK")
        okBtn.OnEvent("Click", SubmitTreeItemSearch)
        cancelBtn := searchGui.AddButton("w80 xp+90", "Cancel")
        cancelBtn.OnEvent("Click", CancelTreeItemSearch)

        ; Show GUI and focus input
        searchGui.Show("w350 h150")
        searchGui["TreeItemInput"].Focus()

        ; Wait for dialog to close
        WinWaitClose("ahk_id " searchGui.Hwnd)

        ; Get the input value
        global g_TreeItemSearchInput
        searchText := g_TreeItemSearchInput
        g_TreeItemSearchInput := ""  ; Clear for next use

        ; If user cancelled, exit
        if (searchText = "")
            return

        ; Get root element and find Tree
        root := UIA.ElementFromHandle(WinExist("A"))
        Sleep 200

        ; Find Tree container with AutomationId="4"
        treeContainer := UIATreeInspector_FindTreeByAutomationId(root, "4")
        if (!treeContainer) {
            MsgBox "Could not find the tree container (AutomationId='4').", "UIA Tree Inspector", "IconX"
            return
        }

        inspectorHwnd := WinExist("UIATreeInspector")
        if (!inspectorHwnd)
            inspectorHwnd := WinExist("ahk_exe UIATreeInspectorAutoHotkey64.exe")
        leftHwnd := UIATreeInspector_FocusLeftWindowsTree(treeContainer, inspectorHwnd)

        ; Get all TreeItem children
        treeItems := treeContainer.FindAll({ Type: "TreeItem" })
        if (!treeItems) {
            MsgBox "No tree items found in the tree container.", "UIA Tree Inspector", "IconX"
            return
        }

        ; Search for TreeItem where Name starts with searchText (case-insensitive)
        matchingItem := ""
        searchTextLower := StrLower(searchText)
        for item in treeItems {
            if (!item)
                continue
            try {
                itemName := item.Name
                if (StrLower(SubStr(itemName, 1, StrLen(searchText))) = searchTextLower) {
                    matchingItem := item
                    break
                }
            } catch {
                ; Skip items without names
                continue
            }
        }

        ; Select the matching item
        if (matchingItem) {
            try {
                matchingItem.Select()
                ; Optional: Scroll into view and set focus
                matchingItem.ScrollIntoView()
                matchingItem.SetFocus()

                ; Workaround: force UIA Tree Inspector to refresh the right-side UIA tree
                ; by "jiggling" selection Down then Up after selection via search.
                if (inspectorHwnd) {
                    if !WinActive("ahk_id " inspectorHwnd) {
                        WinActivate "ahk_id " inspectorHwnd
                        WinWaitActive "ahk_id " inspectorHwnd, , 1
                    }

                    if WinActive("ahk_id " inspectorHwnd) {
                        ; Ensure the windows TreeView has Win32 focus before jiggle (arrow keys)
                        leftHwnd := UIATreeInspector_FocusLeftWindowsTree(treeContainer, inspectorHwnd)
                        try matchingItem.SetFocus()
                        Sleep 500
                        UIATreeInspector_JiggleLeftTree(leftHwnd, inspectorHwnd, 1000)
                    }
                }
            } catch Error as e {
                MsgBox "Error selecting tree item:`n" e.Message, "UIA Tree Inspector", "IconX"
            }
        } else {
            MsgBox Format("No tree item found starting with '{}'.", searchText), "UIA Tree Inspector", "IconX"
        }

    } catch Error as e {
        MsgBox "Error in tree item search:`n" e.Message, "UIA Tree Inspector", "IconX"
    }
}

; Shift + C : Search window/control and copy full UIA tree to clipboard
+c:: {
    barShown := false
    try {
        ; Global variable to store user input
        global g_TreeItemSearchInput := ""

        StandardLoadingBar_Show("⏳ UIA Tree Inspector: preparing copy…", BANNER_ACCENT_INTERMEDIATE, { passive: false })
        barShown := true

        ; Create GUI dialog (same as Shift+S)
        searchGui := Gui("+AlwaysOnTop +ToolWindow", "Select Tree Item")
        searchGui.SetFont("s10", "Segoe UI")
        searchGui.AddText("w350 Center", "Enter text to search for tree item (starts with):")
        searchGui.AddEdit("w300 Center vTreeItemInput")

        SubmitTreeItemSearch(ctrl, *) {
            global g_TreeItemSearchInput
            g_TreeItemSearchInput := ctrl.Gui["TreeItemInput"].Text
            ctrl.Gui.Destroy()
        }
        CancelTreeItemSearch(ctrl, *) {
            global g_TreeItemSearchInput
            g_TreeItemSearchInput := ""
            ctrl.Gui.Destroy()
        }

        okBtn := searchGui.AddButton("w80 xp-40 y+10 Default", "OK")
        okBtn.OnEvent("Click", SubmitTreeItemSearch)
        cancelBtn := searchGui.AddButton("w80 xp+90", "Cancel")
        cancelBtn.OnEvent("Click", CancelTreeItemSearch)

        searchGui.Show("w350 h150")
        searchGui["TreeItemInput"].Focus()
        WinWaitClose("ahk_id " searchGui.Hwnd)

        global g_TreeItemSearchInput
        searchText := g_TreeItemSearchInput
        g_TreeItemSearchInput := ""
        if (searchText = "")
            return

        ; Always-on-top loading banner can keep focus after the modal closes; WinActivate(Inspector) then fails silently.
        try
            StandardLoadingBar_Hide(0)
        catch {
        }
        Sleep 80

        ; Prefer the real foreground window if it is already UIATreeInspector (avoids stale WinExist match).
        inspectorHwnd := 0
        try {
            if (WinActive("ahk_exe AutoHotkey64.exe") && InStr(WinGetTitle("A"), "UIATreeInspector"))
                inspectorHwnd := WinGetID("A")
        } catch {
        }
        if (!inspectorHwnd)
            inspectorHwnd := WinExist("UIATreeInspector")
        if (!inspectorHwnd)
            inspectorHwnd := WinExist("ahk_exe UIATreeInspectorAutoHotkey64.exe")
        if (!inspectorHwnd)
            return

        ; AHK v2 WinActivate throws if the window does not exist; wrap so a bad/stale hwnd does not abort +c.
        loop 6 {
            try {
                if WinActive("ahk_id " inspectorHwnd)
                    break
                if !WinExist("ahk_id " inspectorHwnd)
                    break
                WinActivate "ahk_id " inspectorHwnd
                WinWaitActive "ahk_id " inspectorHwnd, , 0.35
            } catch {
                Sleep 50
            }
        }
        if !WinActive("ahk_id " inspectorHwnd) {
            MsgBox "Could not activate UIATreeInspector after the search dialog. Try again.",
                "UIA Tree Inspector",
                "IconX"
            return
        }

        ; Foreground HWND (avoids stale WinExist match when multiple windows match the title pattern).
        inspectorHwnd := WinGetID("A")

        try {
            StandardLoadingBar_Show("🔎 Selecting window/control…", BANNER_ACCENT_INTERMEDIATE, { passive: false,
                centerOnHwnd: inspectorHwnd })
        } catch {
            try
                StandardLoadingBar_Show("🔎 Selecting window/control…", BANNER_ACCENT_INTERMEDIATE, { passive: false,
                    centerOnHwnd: 0 })
            catch {
            }
        }

        ; Select matching item in left tree (AutomationId="4") (same as Shift+S)
        root := UIA.ElementFromHandle(inspectorHwnd)
        Sleep 500
        treeContainer := UIATreeInspector_FindTreeByAutomationId(root, "4")
        if (!treeContainer) {
            MsgBox "Could not find the tree container (AutomationId='4').", "UIA Tree Inspector", "IconX"
            return
        }

        leftHwnd := UIATreeInspector_FocusLeftWindowsTree(treeContainer, inspectorHwnd)

        treeItems := treeContainer.FindAll({ Type: "TreeItem" })
        if (!treeItems) {
            MsgBox "No tree items found in the tree container.", "UIA Tree Inspector", "IconX"
            return
        }

        matchingItem := ""
        searchTextLower := StrLower(searchText)
        for item in treeItems {
            if (!item)
                continue
            try {
                itemName := item.Name
                if (StrLower(SubStr(itemName, 1, StrLen(searchText))) = searchTextLower) {
                    matchingItem := item
                    break
                }
            } catch {
                continue
            }
        }
        if (!matchingItem) {
            MsgBox Format("No tree item found starting with '{}'.", searchText), "UIA Tree Inspector", "IconX"
            return
        }

        try
            matchingItem.Select()
        catch {
        }
        try
            matchingItem.ScrollIntoView()
        catch {
        }
        try
            matchingItem.SetFocus()
        catch {
        }
        Sleep 1500

        ; Refresh workaround (force correct UIA Tree load)
        try
            StandardLoadingBar_Update("🔄 Refreshing UIA Tree…")
        catch {
        }
        if !WinActive("ahk_id " inspectorHwnd) {
            try {
                WinActivate "ahk_id " inspectorHwnd
                WinWaitActive "ahk_id " inspectorHwnd, , 1
            } catch {
            }
        }
        if WinActive("ahk_id " inspectorHwnd) {
            inspectorHwnd := WinGetID("A")
            leftHwnd := UIATreeInspector_FocusLeftWindowsTree(treeContainer, inspectorHwnd)
            try matchingItem.SetFocus()
            UIATreeInspector_JiggleLeftTree(leftHwnd, inspectorHwnd, 2000)
        }

        ; Focus UIA Tree panel (right-side tree) and select root
        try
            StandardLoadingBar_Update("🌳 Focusing UIA Tree panel…")
        catch {
        }
        Sleep 1500
        rightTree := UIATreeInspector_FindRightDumpTree(root)
        if (!rightTree) {
            MsgBox "Could not find the UIA Tree panel (AutomationId='17').", "UIA Tree Inspector", "IconX"
            return
        }

        rootItem := 0
        try
            rootItem := rightTree.FindFirst({ Type: UIA.Type.TreeItem })
        catch TargetError {
            try {
                items := rightTree.FindAll({ Type: UIA.Type.TreeItem })
                if (items.Length)
                    rootItem := items[1]
            } catch {
            }
        }
        if (rootItem) {
            try rootItem.Select()
            try rootItem.ScrollIntoView()
            try rootItem.SetFocus()
        } else {
            try rightTree.SetFocus()
        }
        Sleep 1500

        ; Copy complete UI tree to clipboard via context menu
        try
            StandardLoadingBar_Update("📋 Copying full tree to clipboard…")
        catch {
        }
        A_Clipboard := ""
        Sleep 400
        try {
            Send "{AppsKey}"
            Sleep 600
            Send "{Up}"
            Sleep 400
            Send "{Enter}"
        } catch {
        }
        Sleep 1200
    } catch Error as e {
        try
            StandardLoadingBar_Update("❌ Copy failed: " . SubStr(e.Message, 1, 60))
        catch {
        }
        try
            StandardLoadingBar_Hide(2000)
        catch {
        }
        MsgBox "Error in Shift+C UIA Tree copy:`n" e.Message, "UIA Tree Inspector", "IconX"
        return
    } finally {
        if (barShown)
            try StandardLoadingBar_Hide(0)
    }
}
#HotIf
