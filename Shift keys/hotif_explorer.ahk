; =============================================================================
; Shift keys module: hotif_explorer.ahk
; Windows Explorer hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

#HotIf WinActive("ahk_exe explorer.exe")

; Explorer-specific helper â€" select first pinned item in the sidebar
SelectExplorerSidebarFirstPinned_EX() {
    try {
        explorerEl := UIA.ElementFromHandle(WinExist("A"))
        navPane := explorerEl.FindFirst({ Type: "Tree" })
        if (navPane) {
            ; If in work environment, prefer selecting the Home tree item directly
            try {
                global IS_WORK_ENVIRONMENT
                if (IS_WORK_ENVIRONMENT) {
                    homeItem := navPane.FindFirst({ Type: "TreeItem", Name: "Home" })
                    if (homeItem) {
                        homeItem.ScrollIntoView()
                        homeItem.Select()    ; select only, no click
                        homeItem.SetFocus()
                        EnsureFocus()
                        return true
                    }
                }
            } catch Error {
                ; ignore and fallback to previous logic
            }
            pinnedKeywords := ["fixo", "pinned", "pin", "fixado", "fixada", "fixar", "preso"]
            firstPinnedItem := unset
            for keyword in pinnedKeywords {
                firstPinnedItem := navPane.FindFirst({ Type: "TreeItem", Name: keyword, matchmode: "Substring" })
                if (firstPinnedItem)
                    break
            }
            if (firstPinnedItem) {
                firstPinnedItem.ScrollIntoView()
                firstPinnedItem.Select()
                firstPinnedItem.SetFocus()
                EnsureFocus()
                return true
            }
        }
    } catch Error {
    }
    Send "{F6}"
    Sleep 100
    Send "{Home}"
    return false
}

; Shift + F : Select first file - File
+f::
{
    ; Send a right-click to shift focus into the main pane
    Click "Right"
    Sleep 100
    ; Clear any in-place edits or text focus first
    Send "{ESC}"

    EnsureItemsViewFocus()

    try {
        explorerEl := UIA.ElementFromHandle(WinExist("A"))

        itemsView := explorerEl.FindFirst({ AutomationId: "ItemsView", Type: "List" })
            ? explorerEl.FindFirst({ AutomationId: "ItemsView", Type: "List" })
            : explorerEl.FindFirst({ ClassName: "UIItemsView", Type: "List" })
                ? explorerEl.FindFirst({ ClassName: "UIItemsView", Type: "List" })
                : explorerEl.FindFirst({ Name: "Items View", Type: "List", matchmode: "Substring" })

        ; Fallback to entire window if we still did not find a dedicated list
        listRoot := itemsView ? itemsView : explorerEl

        ; Pick the very first ListItem inside that list root
        firstItem := listRoot.FindFirst({ Type: "ListItem" })

        if (firstItem) {
            firstItem.ScrollIntoView()
            firstItem.Select()
            firstItem.SetFocus()
            EnsureFocus()
            return
        }
    } catch Error {
        ; swallow and fallback below
    }

    ; Last-chance fallback â€" press Home which works if focus is already inside the list
    Send "{Home}"
    EnsureFocus()
}

; Helper to force focus to the ItemsView pane (file list)
EnsureItemsViewFocus() {
    try {
        explorerHwnd := WinExist("A")
        root := UIA.ElementFromHandle(explorerHwnd)

        ; quick check â€" if ItemsView already has keyboard focus, we're done
        iv := root.FindFirst({ AutomationId: "ItemsView", Type: "List" })
        if iv && iv.HasKeyboardFocus
            return

        ; Send up to 6 F6 cycles to reach the pane
        loop 6 {
            Send "{F6}"
            Sleep 120
            iv := root.FindFirst({ AutomationId: "ItemsView", Type: "List" })
            if iv && iv.HasKeyboardFocus
                break
        }
    } catch Error {
    }
}

Explorer_FindItemsView(root) {
    if !root
        return 0
    try {
        itemsView := root.FindFirst({ AutomationId: "ItemsView", Type: "List" })
        if itemsView
            return itemsView
    }
    try {
        itemsView := root.FindFirst({ ClassName: "UIItemsView", Type: "List" })
        if itemsView
            return itemsView
    }
    try {
        itemsView := root.FindFirst({ Name: "Items View", Type: "List", matchmode: "Substring" })
        if itemsView
            return itemsView
    }
    return 0
}

Explorer_GetItemsViewSelection(itemsView) {
    if !itemsView
        return []
    try {
        if itemsView.GetPropertyValue(UIA.Property.IsSelectionPatternAvailable)
            return itemsView.SelectionPattern.GetSelection()
    } catch {
    }
    selected := []
    try {
        for item in itemsView.FindAll({ Type: "ListItem" }) {
            try {
                if item.GetPropertyValue(UIA.Property.IsSelected)
                    selected.Push(item)
            } catch {
            }
        }
    } catch {
    }
    return selected
}

Editor_ClipboardHasFileDrop() {
    return Clipboard_HasFileDrop()
}

Editor_WaitForClipboardFileDrop(timeoutMs := 800) {
    return Clipboard_WaitForFileDrop(timeoutMs)
}

Editor_NormalizeFileDropPath(path) {
    return Clipboard_NormalizeFilePath(path)
}

Editor_PathIsExistingFile(path) {
    return Clipboard_PathIsExistingFile(path)
}

Editor_IsPlausibleRevealBasename(raw) {
    s := Editor_NormalizeRevealBasename(raw)
    if (s = "" || StrLen(s) > 180)
        return false
    lower := StrLower(s)
    if InStr(lower, "not accessible") || InStr(lower, "screen reader") || InStr(lower, "agentswindow")
        return false
    if InStr(s, "`n") || InStr(s, "`r")
        return false
    return true
}

Editor_GetBasenameFromEditorTitle(editorHwnd) {
    if !(editorHwnd is Integer) || editorHwnd <= 0
        return ""
    try {
        title := WinGetTitle("ahk_id " editorHwnd)
        if (title = "")
            return ""
        parts := StrSplit(title, " - ", , 2)
        if (parts.Length >= 1 && parts[1] != "") {
            candidate := Editor_NormalizeRevealBasename(Trim(parts[1]))
            if Editor_IsPlausibleRevealBasename(candidate)
                return candidate
        }
    } catch {
    }
    return ""
}

Editor_GetClipboardFilePaths() {
    return Clipboard_GetFilePaths()
}

Editor_ClipboardContainsFilePath(expectedPath) {
    return Clipboard_ContainsFilePath(expectedPath)
}

Editor_SetClipboardFiles(paths) {
    return Clipboard_SetFiles(paths)
}

Explorer_RestoreItemsViewSelection(itemsView, selected) {
    if !itemsView || !selected || selected.Length = 0
        return false
    try {
        first := selected[1]
        try first.Select()
        catch {
            try first.SelectionItemPattern.Select()
            catch
                return false
        }
        if (selected.Length > 1) {
            loop selected.Length - 1 {
                item := selected[A_Index + 1]
                try item.AddToSelection()
                catch {
                    try item.SelectionItemPattern.AddToSelection()
                    catch {
                    }
                }
            }
        }
        try first.SetFocus()
        catch {
        }
        return true
    } catch {
    }
    return false
}

Explorer_RefreshItemsViewFromHwnd(explorerHwnd, itemsView := 0) {
    try {
        root := UIA.ElementFromHandle(explorerHwnd)
        refreshed := Explorer_FindItemsView(root)
        if refreshed
            return refreshed
    } catch {
    }
    return itemsView
}

Explorer_WaitItemsViewKeyboardFocus(explorerHwnd, itemsView, timeoutMs := 400) {
    if !itemsView
        return 0
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        itemsView := Explorer_RefreshItemsViewFromHwnd(explorerHwnd, itemsView)
        if (itemsView && itemsView.HasKeyboardFocus)
            return itemsView
        Sleep 40
    }
    return itemsView
}

Explorer_EnsureItemsViewFocusPreserveSelection() {
    explorerHwnd := WinExist("A")
    if !explorerHwnd
        throw Error("No active Explorer window.")

    root := UIA.ElementFromHandle(explorerHwnd)
    if !root
        throw Error("Could not read Explorer UI.")

    itemsView := Explorer_FindItemsView(root)
    if !itemsView
        throw Error("Could not find Items View in Explorer.")

    savedSelection := Explorer_GetItemsViewSelection(itemsView)

    if itemsView.HasKeyboardFocus && savedSelection.Length > 0
        return

    if !itemsView.HasKeyboardFocus {
        try itemsView.SetFocus()
        itemsView := Explorer_WaitItemsViewKeyboardFocus(explorerHwnd, itemsView, 400)
    }

    if (!itemsView || !itemsView.HasKeyboardFocus) {
        loop 6 {
            Send "{F6}"
            itemsView := Explorer_WaitItemsViewKeyboardFocus(explorerHwnd, itemsView, 400)
            if (itemsView && itemsView.HasKeyboardFocus)
                break
        }
    }

    if savedSelection.Length > 0
        Explorer_RestoreItemsViewSelection(itemsView, savedSelection)

    if Explorer_GetItemsViewSelection(itemsView).Length = 0
        throw Error("No file selected in Explorer")
}

; Shift + S : Focus search bar - Search
+s:: Send "^e"

; Shift + A : Focus address bar - Address
+a:: Send "!d"

; Shift + N : New folder - New Folder
+n:: Send("^+n")

; Shift + H : Create a shortcut - sHortcut
+h::
{
    ; Ensure focus is in the file list so the keystrokes hit the right target
    EnsureItemsViewFocus()
    Sleep 100
    Send "{Alt}"
    Sleep 50
    Send "{Enter}"
    Sleep 100
    Send "{Down}"
    Sleep 50
    Send "{Enter}"
}

; Shift + C : Copy as path - Copy
+c:: Send "^+c"

; Shift + R : Share file via context menu workflow - shaRe
+r::
{
    Explorer_CopyOneDriveShareLink_BoschGroup()
}

Explorer_CopyOneDriveShareLink_BoschGroup() {
    ; Reliable flow using classic Explorer context menu + UIA (no fixed Tab counts).
    ; Workflow:
    ;   1) Classic context menu -> S -> Enter (Share)
    ;   2) Wait for Share dialog main view
    ;   3) Full permissions: Link settings -> "People in Bosch Group" -> Apply
    ;      Limited sharing: skip settings (main banner) or Back from single-option Link settings
    ;   4) Copy link and confirm clipboard changed

    try {
        ; Ensure focus and file selection in ItemsView before opening the context menu.
        Explorer_EnsureItemsViewFocusPreserveSelection()
        Sleep 200
        ShowSmallLoadingIndicator_ChatGPT("Sharing file…")

        ; 1) Open classic context menu and trigger Share via accelerator.
        ; Shift+F10 is the canonical "classic menu" key, more reliable than AppsKey on some keyboards.
        Send "+{F10}"
        Sleep 200
        Send "s"
        Sleep 80
        Send "{Enter}"

        ; 2) Wait for OneDrive Share dialog (WebView2 host) to appear and load main controls.
        shareHwnd := OneDriveShare_WaitForShareDialogHwnd(20000)
        if !shareHwnd
            throw Error("Timed out waiting for the OneDrive Share dialog window.")

        shareRoot := UIA.ElementFromHandle(shareHwnd)

        ; Wait until main footer controls exist (indicates main share view is loaded).
        OneDriveShare_WaitForAutomationId(shareRoot, "Footer-button-settings", 20000)
        OneDriveShare_WaitForAutomationId(shareRoot, "copy-button", 20000)

        ; 3) Configure link scope when permitted; limited sharing copies the existing-access link as-is.
        limited := OneDriveShare_IsLimitedSharingMainView(shareRoot)
        if !limited {
            settingsBtn := OneDriveShare_WaitForAutomationId(shareRoot, "Footer-button-settings", 5000)
            OneDriveShare_Click(settingsBtn)
            OneDriveShare_WaitForAutomationId(shareRoot, "od-ModifyPermissions-apply-id", 20000)
            OneDriveShare_WaitForLinkSettingsReady(shareRoot, 10000)

            if OneDriveShare_IsLimitedLinkSettings(shareRoot) {
                OneDriveShare_ClickBack(shareRoot)
            } else {
                if !OneDriveShare_SelectRadioByNameContains(shareRoot, "People in Bosch Group", 5000)
                    throw Error("Could not find 'People in Bosch Group' in Link settings.")
                applyBtn := OneDriveShare_WaitForAutomationId(shareRoot, "od-ModifyPermissions-apply-id", 5000)
                OneDriveShare_Click(applyBtn)
                OneDriveShare_WaitForAutomationId(shareRoot, "copy-button", 20000)
            }
        }

        ; 4) Copy link and verify clipboard changed.
        copyBtn := OneDriveShare_WaitForAutomationId(shareRoot, "copy-button", 5000)
        oldClip := A_Clipboard
        A_Clipboard := ""
        OneDriveShare_Click(copyBtn)
        if !OneDriveShare_WaitForClipboardChange(oldClip, 10000)
            throw Error("Clipboard did not update after 'Copy link'.")

        Sleep 1000
        try WinClose("ahk_id " shareHwnd)
        catch {
        }
    } catch Error as e {
        MsgBox("Share macro failed:`n" e.Message, "Shift+R (Share file)", "IconX")
    } finally {
        HideSmallLoadingIndicator_ChatGPT()
    }
}

OneDriveShare_WaitForShareDialogHwnd(timeout := 20000) {
    deadline := A_TickCount + timeout
    while (A_TickCount < deadline) {
        for hwnd in WinGetList("ahk_class WebView2") {
            try {
                title := WinGetTitle("ahk_id " hwnd)
                if RegExMatch(title, "i)^Share\b") {
                    return hwnd
                }
            } catch {
            }
        }
        Sleep 100
    }
    return 0
}

OneDriveShare_WaitForAutomationId(root, automationId, timeout := 5000) {
    if !IsObject(root)
        return 0

    deadline := A_TickCount + timeout
    while (A_TickCount < deadline) {
        try {
            el := root.FindFirst({ AutomationId: automationId })
            if el
                return el
        } catch {
        }
        Sleep 80
    }
    return 0
}

OneDriveShare_Click(el) {
    if !IsObject(el)
        return false
    try {
        if el.GetPropertyValue(UIA.Property.IsInvokePatternAvailable) {
            el.Invoke()
            return true
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

OneDriveShare_LimitedSharingNeedles() {
    return [
        "Ask owner to share",
        "Sharing is limited",
        "can't invite anyone new",
        "only copy links for people who have existing access",
        "Options are limited",
        "Only people with existing access"
    ]
}

OneDriveShare_IsLimitedSharing(root) {
    return OneDriveShare_TreeContainsText(root, OneDriveShare_LimitedSharingNeedles())
}

OneDriveShare_WaitForLinkSettingsReady(root, timeout := 10000) {
    if !IsObject(root)
        return false
    deadline := A_TickCount + timeout
    while (A_TickCount < deadline) {
        try {
            if (root.FindAll({ Type: "RadioButton" }).Length >= 1)
                return true
        } catch {
        }
        if OneDriveShare_TreeContainsText(root, ["The link works for", "Link settings"])
            return true
        Sleep 80
    }
    return false
}

OneDriveShare_TreeContainsText(root, needles) {
    if !IsObject(root)
        return false
    try {
        for el in root.FindAll({ Type: 50020 }) {
            n := ""
            try n := el.Name
            if (n = "")
                continue
            for needle in needles {
                if InStr(n, needle, false)
                    return true
            }
        }
    } catch {
    }
    return false
}

OneDriveShare_IsLimitedSharingMainView(root) {
    return OneDriveShare_IsLimitedSharing(root)
}

OneDriveShare_HasRadioByNameContains(root, nameNeedle) {
    if !IsObject(root)
        return false
    try {
        for radio in root.FindAll({ Type: "RadioButton" }) {
            n := ""
            try n := radio.Name
            if (n != "" && InStr(n, nameNeedle, false))
                return true
        }
    } catch {
    }
    return false
}

OneDriveShare_IsLimitedLinkSettings(root) {
    if OneDriveShare_IsLimitedSharing(root)
        return true
    try {
        if (root.FindAll({ Type: "RadioButton" }).Length = 1)
            return true
    } catch {
    }
    if !OneDriveShare_HasRadioByNameContains(root, "People in Bosch Group")
    && OneDriveShare_HasRadioByNameContains(root, "existing access")
        return true
    return false
}

OneDriveShare_ClickBack(root) {
    backBtn := 0
    try backBtn := root.FindFirst({ Name: "Back", Type: "50000" })
    catch {
    }
    if !backBtn {
        try backBtn := root.FindFirst({ Name: "Back", Type: "50000", matchmode: "Substring" })
        catch {
        }
    }
    if !backBtn
        throw Error("Could not find Back button in Link settings.")
    OneDriveShare_Click(backBtn)
    if !OneDriveShare_WaitForAutomationId(root, "copy-button", 10000)
        throw Error("Timed out returning to main Share view after Back.")
    return true
}

OneDriveShare_SelectRadioByNameContains(root, nameNeedle, timeout := 5000) {
    deadline := A_TickCount + timeout
    while (A_TickCount < deadline) {
        try {
            radios := root.FindAll({ Type: "RadioButton" })
            for radio in radios {
                n := ""
                try n := radio.Name
                if (n != "" && InStr(n, nameNeedle)) {
                    try {
                        if radio.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable) {
                            if !radio.SelectionItemPattern.IsSelected
                                radio.SelectionItemPattern.Select()
                            return true
                        }
                    } catch {
                    }
                    ; Fallback: invoke/click the radio if SelectionItem isn't available.
                    OneDriveShare_Click(radio)
                    return true
                }
            }
        } catch {
        }
        Sleep 80
    }
    return false
}

OneDriveShare_WaitForClipboardChange(oldClip, timeout := 5000) {
    deadline := A_TickCount + timeout
    while (A_TickCount < deadline) {
        if (A_Clipboard != "" && A_Clipboard != oldClip)
            return true
        Sleep 80
    }
    return false
}

; Shift + P : Select first pinned item in Explorer sidebar - Pinned
+p::
{
    SelectExplorerSidebarFirstPinned_EX()
}

; Shift + L : Select the last item of the Explorer sidebar - Last
+l::
{
    ; First, call the same logic as +P to select the desktop (first pinned item)
    SelectExplorerSidebarFirstPinned_EX()
    Sleep 200

    ; Then press END to go down to the bottom of the tree
    Send "{End}"
    Send "{Up}"
    Send "{Up}"
    Send "{Up}"
    Send "{Up}"
}

; Shift + W : WinRAR add to archive / compact (personal); work PC: 7-Zip add to archive / compress
+w::
{
    global IS_WORK_ENVIRONMENT

    if (IS_WORK_ENVIRONMENT) {
        StandardLoadingBar_Show("⏳ Preparing 7-Zip compress...", BANNER_ACCENT_INTERMEDIATE, { passive: false })

        try {
            StandardLoadingBar_Update("📁 Ensuring focus on items view...", BANNER_ACCENT_INTERMEDIATE)
            EnsureItemsViewFocus()
            Sleep 300

            StandardLoadingBar_Update("📋 Opening context menu...", BANNER_ACCENT_INTERMEDIATE)
            Send "{AppsKey}"
            Sleep 1000  ; Context menu takes time to appear

            StandardLoadingBar_Update("🔍 Locating 7-Zip option...", BANNER_ACCENT_INTERMEDIATE)
            Send "7"
            Sleep 800  ; 7-Zip needs time to respond

            StandardLoadingBar_Update("✍️  Sending first confirmation...", BANNER_ACCENT_INTERMEDIATE)
            Send "{Enter}"
            Sleep 600

            StandardLoadingBar_Update("✍️  Sending second confirmation...", BANNER_ACCENT_INTERMEDIATE)
            Send "{Enter}"
            Sleep 800

            StandardLoadingBar_Hide(500)
        } catch as e {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ 7-Zip compress error: " . e.Message, 2000, BANNER_ACCENT_ERROR)
        }
        return
    }

    StandardLoadingBar_Show("⏳ Preparing WinRAR compress...", BANNER_ACCENT_INTERMEDIATE, { passive: false })

    try {
        StandardLoadingBar_Update("📁 Ensuring focus on items view...", BANNER_ACCENT_INTERMEDIATE)
        EnsureItemsViewFocus()
        Sleep 300

        StandardLoadingBar_Update("📋 Opening context menu...", BANNER_ACCENT_INTERMEDIATE)
        Send "{AppsKey}"
        Sleep 1000  ; Context menu takes time to appear

        StandardLoadingBar_Update("🔍 Locating WinRAR option...", BANNER_ACCENT_INTERMEDIATE)
        Send "w"
        Sleep 800

        StandardLoadingBar_Update("✍️  Sending first confirmation...", BANNER_ACCENT_INTERMEDIATE)
        Send "{Enter}"
        Sleep 600

        StandardLoadingBar_Update("✍️  Sending second confirmation...", BANNER_ACCENT_INTERMEDIATE)
        Send "{Enter}"
        Sleep 800

        StandardLoadingBar_Hide(500)
    } catch as e {
        StandardLoadingBar_Hide(0)
        ShowCenteredOverlay_Utils("❌ WinRAR compress error: " . e.Message, 2000, BANNER_ACCENT_ERROR)
    }
}

; Shift + X : WinRAR extract to current folder (personal); work PC: 7-Zip extract
+x::
{
    global IS_WORK_ENVIRONMENT

    if (IS_WORK_ENVIRONMENT) {
        StandardLoadingBar_Show("⏳ Preparing 7-Zip extraction...", BANNER_ACCENT_INTERMEDIATE, { passive: false })

        try {
            StandardLoadingBar_Update("📁 Ensuring focus on items view...", BANNER_ACCENT_INTERMEDIATE)
            EnsureItemsViewFocus()
            Sleep 300

            StandardLoadingBar_Update("📋 Opening context menu...", BANNER_ACCENT_INTERMEDIATE)
            Send "{AppsKey}"
            Sleep 1000  ; Context menu takes time to appear

            StandardLoadingBar_Update("🔍 Locating 7-Zip option...", BANNER_ACCENT_INTERMEDIATE)
            Send "7"
            Sleep 800  ; 7-Zip menu needs time to respond

            StandardLoadingBar_Update("⬇️  Moving to extract option...", BANNER_ACCENT_INTERMEDIATE)
            Send "{Down}"
            Sleep 400

            StandardLoadingBar_Update("✍️  Extracting to current folder...", BANNER_ACCENT_INTERMEDIATE)
            Send "{Enter}"
            Sleep 800

            StandardLoadingBar_Hide(500)
        } catch as e {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ 7-Zip extract error: " . e.Message, 2000, BANNER_ACCENT_ERROR)
        }
        return
    }

    StandardLoadingBar_Show("⏳ Preparing WinRAR extraction...", BANNER_ACCENT_INTERMEDIATE, { passive: false })

    try {
        StandardLoadingBar_Update("📁 Ensuring focus on items view...", BANNER_ACCENT_INTERMEDIATE)
        EnsureItemsViewFocus()
        Sleep 300

        StandardLoadingBar_Update("📋 Opening context menu...", BANNER_ACCENT_INTERMEDIATE)
        Send "{AppsKey}"
        Sleep 1000  ; Context menu takes time to appear

        StandardLoadingBar_Update("🔍 Locating WinRAR option...", BANNER_ACCENT_INTERMEDIATE)
        ; WinRAR shell menu accelerators (English); adjust if UI language differs
        Send "w"
        Sleep 800

        StandardLoadingBar_Update("✍️  Extracting to current folder...", BANNER_ACCENT_INTERMEDIATE)
        Send "x"
        Sleep 800

        StandardLoadingBar_Hide(500)
    } catch as e {
        StandardLoadingBar_Hide(0)
        ShowCenteredOverlay_Utils("❌ WinRAR extract error: " . e.Message, 2000, BANNER_ACCENT_ERROR)
    }
}

