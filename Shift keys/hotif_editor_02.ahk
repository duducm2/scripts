; =============================================================================
; Shift keys module: hotif_editor_02.ahk
; Cursor/VS Code editor hotkeys (part 2)
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

; ---------------------------------------------------------------------------
; VS Code / Cursor: hide (toggle) bottom panel via UIA (no native shortcut)
; ---------------------------------------------------------------------------
ClickHidePanelButton() {
    ; Debounce: if hotkey repeats while key is held, ignore fast repeats.
    prior := ""
    since := -1
    this := ""
    try prior := A_PriorHotkey
    try since := A_TimeSincePriorHotkey
    try this := A_ThisHotkey
    sinceNum := (since = "" ? -1 : (since + 0))
    if (prior = this && sinceNum >= 0 && sinceNum < 350) {
        return false
    }
    try {
        uia := UIA_Browser()
        if !IsObject(uia)
            return false

        btn := 0
        foundAs := ""

        ; VS Code commonly exposes it as a Button with shortcut text.
        try btn := uia.FindFirst({ Name: "Hide Panel (Ctrl+J)", ControlType: "Button" })
        if btn
            foundAs := "Button:Hide Panel (Ctrl+J)"
        if !btn
            try btn := uia.FindFirst({ Name: "Hide Panel", ControlType: "Button" })
        if (!foundAs && btn)
            foundAs := "Button:Hide Panel"

        ; Cursor sometimes exposes it as a CheckBox-style action (still clickable).
        if !btn
            try btn := uia.FindFirst({ Name: "Hide Panel", ControlType: "CheckBox" })
        if (!foundAs && btn)
            foundAs := "CheckBox:Hide Panel"

        ; Substring fallback across UI variants.
        if !btn
            try btn := uia.FindFirst({ Name: "Hide Panel", matchmode: "Substring" })
        if (!foundAs && btn)
            foundAs := "Substring:Hide Panel"

        if btn {
            clicked := false
            supportsInvoke := false
            supportsToggle := false
            try supportsInvoke := btn.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)
            try supportsToggle := btn.GetPropertyValue(UIA.Property.IsTogglePatternAvailable)
            if (supportsToggle) {
                try {
                    btn.TogglePattern.Toggle()
                    clicked := true
                } catch {
                }
            }
            if (supportsInvoke) {
                try {
                    btn.Invoke()
                    clicked := true
                } catch {
                }
            }
            if (!clicked) {
                try {
                    btn.Click()
                    clicked := true
                } catch {
                }
            }
            return clicked
        }
    } catch {
    }

    ; Conservative fallback: keep the old behavior available if UIA fails.
    try Send "^j"
    return false
}

; Ctrl + 6 : Marp export - trigger export, handle Save As and Replace dialogs
^6::
{
    ; Show persistent banner for the entire export flow
    ShowSmallLoadingIndicator_ChatGPT("Exporting with Marp...")
    slowStepMs := 300  ; used throughout this hotkey flow
    mainHwnd := 0
    try {
        ; 1. Trigger Marp export
        try {
            mainHwnd := WinGetID("A")
        } catch {
            ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        Send "^6"
        Sleep 500
        Sleep slowStepMs

        ; 2. Wait for Save As / Export dialog
        ; Try: native #32770, title match, or active window change (modal steals focus)
        prevMatchMode := A_TitleMatchMode
        SetTitleMatchMode 2
        saveDialogHwnd := 0
        deadline := A_TickCount + 25000
        while (A_TickCount < deadline) {
            ; Native Windows dialog (standard Save As)
            h := WinExist("ahk_class #32770")
            if h {
                saveDialogHwnd := h
                break
            }
            ; Title contains Save/Export (any window)
            for str in ["Save As", "Export", "Salvar como", "Guardar como", "Save File", "Save PDF", "Marp",
                "Export PDF"] {
                h := WinExist(str)
                if h {
                    saveDialogHwnd := h
                    break 2
                }
            }
            ; Fallback: modal dialog stole focus (active window changed from main)
            try {
                curr := WinGetID("A")
            } catch {
                ; Transient: no active window at this tick. Keep waiting for the dialog instead of aborting.
                Sleep 250
                continue
            }
            if curr && curr != mainHwnd {
                currTitle := WinGetTitle("ahk_id " curr)
                currClass := WinGetClass("ahk_id " curr)
                ; Likely a dialog: standard dialog class, or a Chromium modal whose TITLE indicates Save/Export.
                ; IMPORTANT: Do NOT treat any Chrome_WidgetWin as dialog (Cursor itself is Chrome_WidgetWin_1).
                isTitleDialogish := InStr(currTitle, "Save") || InStr(currTitle, "Save As")
                || InStr(currTitle, "Export") || InStr(currTitle, "Marp")
                || InStr(currTitle, "Confirm Save As") || InStr(currTitle, "Confirm")
                || InStr(currTitle, "Salvar") || InStr(currTitle, "Guardar")
                if InStr(currClass, "32770") || (InStr(currClass, "Chrome_WidgetWin") && isTitleDialogish) {
                    saveDialogHwnd := curr
                    break
                }
            }
            Sleep 250
        }
        SetTitleMatchMode prevMatchMode
        if !saveDialogHwnd {
            return
        }
        try {
            WinActivate("ahk_id " saveDialogHwnd)
        } catch {
            ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        Sleep 700
        Sleep slowStepMs

        ; 2b. Extract PDF path and filename from the Save dialog via UIA (before confirming save)
        filePath := ""
        fileNameOnly := ""
        fileNameEditEl := 0
        try {
            root := UIA.ElementFromHandle(saveDialogHwnd)
            fileNameEdit := ""

            ; Scan Edit controls and detect the filename field
            try {
                edits := root.FindElements({ Type: "Edit" })
                for el in edits {
                    val := ""
                    try val := el.Value
                    catch {
                    }
                    if (val = "")
                        continue

                    parts := StrSplit(val, "\")
                    suffix := parts.Length ? parts[parts.Length] : val

                    ; If this Edit looks like the File name field (based on name/id and .pdf suffix), capture it
                    if (suffix != "" && InStr(StrLower(suffix), ".pdf")
                    && (el.AutomationId = "1001" || el.Name = "File name:")) {
                        filePath := val
                        SplitPath filePath, , , &ext, &nameNoExt
                        fileNameOnly := (nameNoExt != "") ? (nameNoExt . (ext != "" ? "." ext : "")) : suffix
                    }
                }
            } catch {
            }

            ; First attempt: Edit with AutomationId 1148 (matches file dialog helper)
            fileNameEdit := root.FindFirst({ Type: "Edit", AutomationId: "1148" })

            ; Second attempt: ComboBox 1148 -> inner Edit
            if !fileNameEdit {
                fileNameCombo := root.FindFirst({ Type: "ComboBox", AutomationId: "1148" })
                if fileNameCombo
                    fileNameEdit := fileNameCombo.FindFirst({ Type: "Edit" })
            }

            ; Third attempt: Edit by common localized names
            if !fileNameEdit {
                for name in ["File name:", "Nome:", "Filename:", "File Name:", "Name:", "Nome do arquivo:"] {
                    fileNameEdit := root.FindFirst({ Type: "Edit", Name: name })
                    if fileNameEdit
                        break
                }
            }
            if fileNameEdit {
                fileNameEditEl := fileNameEdit
                pathOrName := Trim(fileNameEdit.Value)
                if (pathOrName != "") {
                    if (InStr(pathOrName, "\") || InStr(pathOrName, ":"))
                        filePath := pathOrName
                    else {
                        ; Filename only: try to get current folder from another Edit in the dialog
                        for el in root.FindElements({ Type: "Edit" }) {
                            try {
                                val := Trim(el.Value)
                                if (val != "" && InStr(val, "\") && InStr(val, ":")) {
                                    filePath := RTrim(val, "\") "\" pathOrName
                                    break
                                }
                            } catch {
                                continue
                            }
                        }
                        if (filePath = "")
                            filePath := pathOrName
                    }
                    SplitPath filePath, , , &ext, &nameNoExt
                    fileNameOnly := (nameNoExt != "") ? (nameNoExt . (ext != "" ? "." ext : "")) : pathOrName
                }
            }
        } catch {
            ; Fallback: no filename extracted; we will just open a new Chrome window at the end
        }

        ; Prepare Chrome context for this PDF: close old windows and open a new one
        if (fileNameOnly != "")
            EnsureSingleChromePdfInstance(filePath, fileNameOnly)
        Sleep slowStepMs

        Send "{Enter}"  ; Confirm initial save
        Sleep slowStepMs

        stillThere := WinExist("ahk_id " saveDialogHwnd) ? 1 : 0

        ; If Enter didn't confirm, click the dialog's Export button via UIA.
        if (stillThere) {
            try WinActivate("ahk_id " saveDialogHwnd)
            catch {
            }
            Sleep 120

            try {
                dlgRoot := UIA.ElementFromHandle(saveDialogHwnd)
                exportBtn := dlgRoot.FindFirst({ Type: "Button", Name: "Export", AutomationId: "1" })
                if !exportBtn
                    exportBtn := dlgRoot.FindFirst({ Type: "Button", Name: "Export" })
                if exportBtn {
                    exportBtn.Invoke()
                }
            } catch {
            }

            Sleep slowStepMs
        }

        ; 3. Handle Confirm Save As / Replace dialog (ClassName #32770, Name: "Confirm Save As")
        ; WinGetText doesn't capture UIA Text elements; use window title. Yes button has Alt+Y.
        SetTitleMatchMode 2
        loop 10 {
            Sleep 500
            replaceHwnd := WinExist("ahk_class #32770")
            if replaceHwnd {
                title := WinGetTitle("ahk_id " replaceHwnd)
                if InStr(title, "Confirm Save As") || InStr(title, "Confirmar Salvar")
                || InStr(title, "Confirmar Guardar") || InStr(title, "Confirm Replace") {
                    try {
                        WinActivate("ahk_id " replaceHwnd)
                    } catch {
                        ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000,
                            BANNER_ACCENT_ERROR)
                        break
                    }
                    Sleep 900  ; Delay for dialog to stabilize before confirming
                    Send "!y"   ; Alt+Y = Yes (per UIA: AcceleratorKey: "Alt+Y")
                    break
                }
            }
        }

        ; 4. Wait for Marp export and viewer open to complete
        Sleep 800
        Sleep slowStepMs
    } finally {
        ; Always hide the banner when the flow completes or aborts
        HideSmallLoadingIndicator_ChatGPT()
    }
}

; Shift + F : Fold - Fold
+f::
{
    Send "^+8"
}

; Shift + U : Unfold - Unfold
+u::
{
    Send "^+9"
}

; Shift + M : Open markdown preview to the side - Markdown
+m:: Send "+i"

; Shift + W : Move editor into new window - Window
+w:: Send "+o"

; Shift + T : Go to terminal - Terminal
+t:: Send "^'"

; Shift + N : New terminal - New Terminal
+n:: Send '^+"'

; Shift + E : Go to file explorer - Explorer
+e:: Send "^+e"

; Shift + K : Open markdown preview and move editor into new window - Keep
+k::
{
    ; Show banner while algorithm is executing
    ; ShowSmallLoadingIndicator_ChatGPT("Processing...")

    ; Step 1: Trigger markdown preview (Shift+M -> +i in Cursor)
    Send "+i"
    Sleep 1800

    ; Step 2: Center mouse in active window (Win+Alt+Shift+Q)
    Send "#!+q"
    Sleep 150

    ; Step 3: Offset mouse right into Markdown Preview Enhanced pane
    PREVIEW_OFFSET_PX := 60
    MouseGetPos(&x, &y)
    x += PREVIEW_OFFSET_PX
    DllCall("SetCursorPos", "int", x, "int", y)
    Sleep 100

    ; Step 4: Click to focus preview area
    Click

    ; Step 5: Detach tab (Shift+W -> +o in Cursor)
    Sleep 2000
    Send "+o"

    Sleep 300

    WinMaximize "A"

    ; Hide banner after completion
    ; HideSmallLoadingIndicator_ChatGPT()
}

; Shift + C : Command palette - Command
+c:: Send "^+p"

; Shift + X : Expand selection - Expand
+x:: Send "+!{Right}"

; Shift + S : Go to symbol in access view - Symbol
+s:: Send "+m"

; Shift + H : Navigate to GitHub Copilot chat history - History
+h::
{
    OpenCopilotChatHistory()
}

; Helper: Detect if the secondary sidebar (Copilot chat panel) is visible in VS Code
; Checks the "checked" state of the "Toggle Secondary Side Bar (Alt+I)" button
IsSecondarySidebarVisible() {
    try {
        hwnd := WinExist("A")
        if !hwnd
            return false
        root := UIA.ElementFromHandle(hwnd)
        if !root
            return false
        return Editor_IsWorkbenchToggleOn(root, "Toggle Secondary Side Bar")
    } catch {
        return false
    }
}

; Helper: Click the "Go Back" button in the GitHub Copilot chat view
; Specifically targets the button ONLY within the chat-view-title-container to avoid clicking the main toolbar's back button
ClickCopilotGoBackButton() {
    try {
        hwnd := WinExist("A")
        if (!hwnd)
            return false

        root := UIA.ElementFromHandle(hwnd)
        if (!root)
            return false

        ; First and foremost: Find the chat-view-title-container
        ; We MUST scope our search to this container to avoid clicking the main toolbar's Go Back button
        chatViewTitle := 0

        try {
            ; Search for the chat-view-title-container by ClassName
            allGroups := root.FindAll({ Type: 50026 })
            if (allGroups) {
                for grp in allGroups {
                    try {
                        className := grp.ClassName
                        if (InStr(className, "chat-view-title-container")) {
                            chatViewTitle := grp
                            break
                        }
                    } catch {
                        continue
                    }
                }
            }
        } catch {
            chatViewTitle := 0
        }

        ; If we found the chat view title container, ONLY search within it
        if (chatViewTitle) {
            try {
                ; Find ALL buttons within this container
                btns := chatViewTitle.FindAll({ Type: 50000 })
                if (btns) {
                    for btn in btns {
                        try {
                            name := btn.Name
                            ; Look for button with "Go Back" in the name within the chat view
                            if (InStr(name, "Go Back")) {
                                ; Try Invoke pattern first
                                try {
                                    if btn.GetPropertyValue(UIA.Property.IsInvokePatternAvailable) {
                                        btn.InvokePattern.Invoke()
                                        return true
                                    }
                                } catch {
                                }

                                ; Fallback to Click
                                try {
                                    btn.Click()
                                    return true
                                } catch {
                                }
                            }
                        } catch {
                            continue
                        }
                    }
                }
            } catch {
            }
        }

        ; If we reach here, either chat view wasn't found or there's no Go Back button in it
        ; This is fine - we're already on the history/sessions page
        return false
    } catch Error as e {
        return false
    }
}

; Main function: Navigate to GitHub Copilot chat history
OpenCopilotChatHistory() {
    try {
        ; Step 1: Check if secondary sidebar is visible
        isSidebarVisible := IsSecondarySidebarVisible()

        if (!isSidebarVisible) {
            ; Step 2: Open the secondary sidebar using Alt+I
            Send "!i"
            Sleep 300  ; Wait for UI to stabilize
        }

        ; Step 3: Try to click the "Go Back" button to navigate to chat history
        ; If the button doesn't exist, we're already on the sessions page - that's fine, just return
        buttonClicked := ClickCopilotGoBackButton()

        if (buttonClicked) {
            ; Step 4: Verify the view updated
            Sleep 200  ; Brief delay to allow UI update
            ; The view should now display the chat history session
            return
        }

        ; If Go Back button doesn't exist, we're already on the sessions/history page
        ; No need to do anything else or fall back to command palette
        ; Just verify we're done
        Sleep 100
    } catch Error as e {
        ; Silent error handling - avoid disrupting user workflow
    }
}

; Click the VS Code Copilot model picker button in the chat input toolbar.
; Targets names like "Pick Model, Auto" and scopes to chat-input-toolbars.
ClickVSCodeCopilotModelButton() {
    try {
        hwnd := WinExist("A")
        if (!hwnd)
            return false

        root := UIA.ElementFromHandle(hwnd)
        if (!root)
            return false

        chatToolbars := 0
        try {
            groups := root.FindAll({ Type: 50026 })
            if (groups) {
                for grp in groups {
                    try {
                        cls := grp.ClassName
                        if (InStr(cls, "chat-input-toolbars")) {
                            chatToolbars := grp
                            break
                        }
                    } catch {
                        continue
                    }
                }
            }
        } catch {
            chatToolbars := 0
        }

        if (!chatToolbars)
            return false

        try {
            btns := chatToolbars.FindAll({ Type: 50000 })
            if (btns) {
                for btn in btns {
                    try {
                        nm := btn.Name
                        if (InStr(nm, "Pick Model,")) {
                            try {
                                btn.SetFocus()
                                Sleep 40
                                Send "{Enter}"
                                return true
                            } catch {
                            }
                            try {
                                btn.Click()
                                return true
                            } catch {
                            }
                        }
                    } catch {
                        continue
                    }
                }
            }
        } catch {
        }

        return false
    } catch {
        return false
    }
}

; Shift + I : Paste Image - Image
+i:: Send "!y"

; Shift + G : Fold Git repos (SCM) - Git Fold (implementation below)

; Shift + Q : Search - Search (Q for Query)
+q:: Send "^+f"

; Shift + R : Open Bread Crumbs menu - Breadcrumbs (R for Route/breadcrumbs)
+r:: Send "+r"

; Shift + D : Git section - Git
+d:: Send "+d"

Editor_FocusScmCommitInput() {
    try {
        hwnd := WinExist("A")
        if (!hwnd)
            return false
        root := UIA.ElementFromHandle(hwnd)
        if (!root)
            return false
        el := 0
        try el := root.FindFirst({ AutomationId: "scm.input" })
        catch {
        }
        if (!el) {
            try {
                ti := root.FindFirst({ Type: UIA.Type.TreeItem, Name: "Source Control Input" })
                if (ti) {
                    try el := ti.FindFirst({ Type: UIA.Type.Edit, ClassName: "inputarea monaco-mouse-cursor-text" })
                    catch {
                    }
                    if (!el) {
                        try el := ti.FindFirst({ Type: UIA.Type.Edit })
                        catch {
                        }
                    }
                }
            } catch {
            }
        }
        if (!el) {
            try {
                for edit in root.FindAll({ Type: UIA.Type.Edit }) {
                    try {
                        if (InStr(edit.Name, "Message", false) || InStr(edit.AutomationId, "scm", false)) {
                            el := edit
                            break
                        }
                    } catch {
                    }
                }
            } catch {
            }
        }
        if (el) {
            try {
                el.SetFocus()
                return true
            } catch {
            }
        }
    } catch {
    }
    Send "{Tab}"
    return false
}

Editor_QuickCommit() {
    hwnd := WinExist("A")
    if !hwnd
        return
    sidebarWasVisible := Editor_IsPrimarySidebarVisible(hwnd)
    FocusSourceControlViewForCommitGeneration()
    Editor_FocusScmCommitInput()
    Sleep 80
    msg := "generic commit " . FormatTime(, "yyyy-MM-dd HH:mm:ss")
    Send "^a"
    Sleep 50
    SendText msg
    Sleep 100
    Send "+v"
    Sleep 500
    Send "+b"
    Sleep 500
    if (sidebarWasVisible) {
        ; Send "+e" types "E" when SCM commit input still has focus — use ^+e twice (Shift+E relay).
        Editor_ReturnFocusToMainEditor(hwnd)
    } else {
        Editor_HidePrimarySidebar(hwnd)
    }
    try ScriptSoundPlay(A_ScriptDir . "\assets\sounds\commit.mp3")
}

; Shift + J : Quick commit — generic timestamped message, commit, push
+j:: Editor_QuickCommit()

; Shift + Z : Close all editors - Close
+z::
{
    Send "+f"
}

; Shift + Y : Zen mode - Zen
+y:: Send "+z"

; Shift + P : Git Pull — native in Cursor/VS Code (no AHK remap)

; Shift + V : Git Commit - Commit
+v:: Send "+v"

; Shift + B : Git Push - Push
+b:: Send "+b"

; Alt + S : Git Stash and Pull (native Alt+S + Enter, then native Shift+P pull)
$!s:: {
    Send "!s"
    Sleep 150
    Send "{Enter}"
    Sleep 150
    Send "+c"
}

; Global variable for commit push selector target window
global gCommitPushTargetWin := 0
; Global variable to store the user's push decision ("push" | "dont_push" | "")
global gCommitPushDecision := ""
; Global variable for non-blocking commit push banner GUI
global g_CommitPushBannerGui := ""
global g_CommitPushBannerBorderGui := ""

; Non-blocking banner: "Don't push? Press N within 5 seconds" (dark background, yellow accent border)
ShowCommitPushBanner() {
    global g_CommitPushBannerGui, g_CommitPushBannerBorderGui
    try {
        if IsObject(g_CommitPushBannerBorderGui) && g_CommitPushBannerBorderGui.Hwnd
            g_CommitPushBannerBorderGui.Destroy()
    } catch {
    }
    g_CommitPushBannerBorderGui := ""
    try {
        if IsObject(g_CommitPushBannerGui) && g_CommitPushBannerGui.Hwnd
            g_CommitPushBannerGui.Destroy()
    } catch {
    }
    bannerGui := Gui("+AlwaysOnTop -Caption +ToolWindow")
    bannerGui.BackColor := "1E1E2E"
    bannerGui.SetFont("s14 cFFFFFF Bold", "Segoe UI")
    bannerGui.Add("Text", "w400 Center", "Don't push? Press N within 5 seconds")
    activeWin := WinGetID("A")
    if (activeWin)
        WinGetPos(&winX, &winY, &winW, &winH, activeWin)
    else {
        workArea := SysGet.MonitorWorkArea(SysGet.MonitorPrimary)
        winX := workArea.Left
        winY := workArea.Top
        winW := workArea.Right - workArea.Left
        winH := workArea.Bottom - workArea.Top
    }
    bannerGui.Show("AutoSize Hide")
    guiW := 0
    guiH := 0
    bannerGui.GetPos(, , &guiW, &guiH)
    guiX := winX + (winW - guiW) / 2
    guiY := winY + (winH - guiH) / 2
    borderWidth := 6
    borderGui := Gui("+AlwaysOnTop -Caption +ToolWindow")
    borderGui.BackColor := BANNER_ACCENT_INTERMEDIATE
    borderGui.Show("NA x" . Round(guiX - borderWidth) . " y" . Round(guiY - borderWidth) . " w" . (guiW + 2 *
        borderWidth) . " h" . (guiH + 2 * borderWidth))
    g_CommitPushBannerBorderGui := borderGui
    bannerGui.Show("x" . Round(guiX) . " y" . Round(guiY) . " NA")
    WinSetTransparent(220, bannerGui)
    g_CommitPushBannerGui := bannerGui
    Hotkey("n", CommitPushBanner_NHandler, "On")
    Hotkey("N", CommitPushBanner_NHandler, "On")
    SetTimer(CloseCommitPushBanner, -5000)
}

CommitPushBanner_NHandler(*) {
    global gCommitPushDecision
    ; User explicitly opted out of pushing; keep commit-only.
    gCommitPushDecision := "dont_push"
    CloseCommitPushBanner()
}

CloseCommitPushBanner() {
    global g_CommitPushBannerGui, g_CommitPushBannerBorderGui
    try {
        if IsObject(g_CommitPushBannerBorderGui) && g_CommitPushBannerBorderGui.Hwnd {
            g_CommitPushBannerBorderGui.Destroy()
            g_CommitPushBannerBorderGui := ""
        }
    } catch {
    }
    try {
        if IsObject(g_CommitPushBannerGui) && g_CommitPushBannerGui.Hwnd {
            g_CommitPushBannerGui.Destroy()
            g_CommitPushBannerGui := ""
        }
    } catch {
    }
    try Hotkey("n", "Off")
    catch {
    }
    try Hotkey("N", "Off")
    catch {
    }
    SetTimer(CloseCommitPushBanner, 0)
}

; Function to get commit push action by number
GetCommitPushActionByNumber(numberText) {
    try number := Integer(numberText)
    catch {
        return ""
    }
    actionMap := Map()
    actionMap[1] := "push"
    actionMap[2] := "dont_push"
    return (actionMap.Has(number)) ? actionMap[number] : ""
}

; Execute stored decision at the exact current push moment
ExecuteStoredCommitPushDecision() {
    global gCommitPushDecision
    global gCommitPushTargetWin
    if (gCommitPushDecision = "push") {
        ; Wait a moment for Cursor to process the commit
        Sleep 500
        ; Ensure the intended window has focus before sending the push hotkey
        if (gCommitPushTargetWin) {
            if (WinExist("ahk_id " gCommitPushTargetWin)) {
                WinActivate gCommitPushTargetWin
                WinWaitActive("ahk_id " gCommitPushTargetWin, , 2)
                Sleep 200
            } else {
                ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
            }
        }
        Send "+b"
    }
    ; Clear after execution to avoid reusing stale decisions
    gCommitPushDecision := ""
}

; Function to execute commit push action
ExecuteCommitPushAction(action) {
    if (action = "")
        return

    if (action = "push") {
        ; Option 1: Push (send Shift+B)
        Send "+b"
    } else if (action = "dont_push") {
        ; Option 2: Don't push (do nothing)
        ; Just close the popup, no action needed
    }
}

; Auto-submit function for commit push selector
AutoSubmitCommitPush(ctrl, *) {
    currentValue := ctrl.Text
    if (currentValue != "" && IsInteger(currentValue)) {
        action := GetCommitPushActionByNumber(currentValue)
        if (action != "") {
            ctrl.Gui.Destroy()
            ExecuteCommitPushAction(action)
        }
    }
}

; Manual submit function for commit push selector (backup)
SubmitCommitPush(ctrl, *) {
    currentValue := ctrl.Gui["CommitPushInput"].Text
    if (currentValue != "" && IsInteger(currentValue)) {
        action := GetCommitPushActionByNumber(currentValue)
        if (action != "") {
            ctrl.Gui.Destroy()
            ExecuteCommitPushAction(action)
        } else {
            MsgBox "Invalid selection. Please choose 1-2.", "Commit Push Selector", "IconX"
        }
    }
}

; Cancel function for commit push selector
CancelCommitPush(ctrl, *) {
    ctrl.Gui.Destroy()
}

; Function to show commit push selector popup
ShowCommitPushSelector() {
    try {
        ; Remember current target window before showing GUI
        gCommitPushTargetWin := WinExist("A")
        ; Create GUI for commit push selection with auto-submit
        commitPushGui := Gui("+AlwaysOnTop +ToolWindow", "Commit Push Selector")
        commitPushGui.SetFont("s10", "Segoe UI")

        ; Add instruction text
        commitPushGui.AddText("w350 Center",
            "Commit sent! Choose next action:`n`n1. Push (Shift+B)`n2. Don't push`n`nType a number (1-2):")

        ; Add input field with auto-submit
        commitPushGui.AddEdit("w50 Center vCommitPushInput", "")
        commitPushGui["CommitPushInput"].OnEvent("Change", AutoSubmitCommitPush)

        ; Add manual submit button (backup)
        commitPushGui.AddButton("w80", "Submit").OnEvent("Click", SubmitCommitPush)

        ; Add cancel button
        commitPushGui.AddButton("w80", "Cancel").OnEvent("Click", CancelCommitPush)

        ; Show GUI and focus input
        commitPushGui.Show("w350 h150")
        commitPushGui["CommitPushInput"].Focus()

    } catch Error as e {
        MsgBox "Error in commit push selector: " e.Message, "Commit Push Selector Error", "IconX"
    }
}

; Auto-submit function - triggers when text changes
global gEmojiTargetWin := 0

GetEmojiByNumber(numberText) {
    try number := Integer(numberText)
    catch {
        return ""
    }
    emojiMap := Map()
    emojiMap[1] := "🔲"
    emojiMap[2] := "⏳"
    emojiMap[3] := "⚡"
    emojiMap[4] := "✅"
    emojiMap[5] := "❓"
    emojiMap[6] := "ℹ️"
    return (emojiMap.Has(number)) ? emojiMap[number] : ""
}

InsertEmojiToTarget(emoji, targetWin := 0) {
    global gEmojiTargetWin
    if (emoji = "")
        return
    hwnd := targetWin ? targetWin : gEmojiTargetWin
    if (hwnd) {
        if (!WinExist("ahk_id " hwnd)) {
            ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        if (!WM_EnsureForegroundForSend(hwnd, 2000)) {
            ShowCenteredOverlay_Utils("❌ Error: Could not focus target window.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        Sleep 50
    }

    ; Use direct text insertion - no clipboard manipulation
    ; This is more reliable and won't interfere with user's clipboard
    SendText(emoji)
}

AutoSubmitEmoji(ctrl, *) {
    global gEmojiTargetWin
    currentValue := ctrl.Text
    if (currentValue != "" && IsInteger(currentValue)) {
        emoji := GetEmojiByNumber(currentValue)
        if (emoji != "") {
            targetWin := gEmojiTargetWin
            ctrl.Gui.Destroy()
            SetTimer(() => InsertEmojiToTarget(emoji, targetWin), -75)
        }
    }
}

; Manual submit function (backup)
SubmitEmoji(ctrl, *) {
    global gEmojiTargetWin
    currentValue := ctrl.Gui["EmojiInput"].Text
    if (currentValue != "" && IsInteger(currentValue)) {
        emoji := GetEmojiByNumber(currentValue)
        if (emoji != "") {
            targetWin := gEmojiTargetWin
            ctrl.Gui.Destroy()
            SetTimer(() => InsertEmojiToTarget(emoji, targetWin), -75)
        } else {
            MsgBox "Invalid selection. Please choose 1-6.", "Emoji Selector", "IconX"
        }
    }
}

; Cancel function
CancelEmoji(ctrl, *) {
    ctrl.Gui.Destroy()
}

; Shift + O : Emoji selector (Auto-submit version) - Emoji
+o::
{
    try {
        ; Remember current target window before showing GUI
        gEmojiTargetWin := WinGetID("A")
        ; Create GUI for emoji selection with auto-submit
        emojiGui := Gui("+AlwaysOnTop +ToolWindow", "Emoji Selector")
        if (gEmojiTargetWin)
            emojiGui.Opt("+Owner" gEmojiTargetWin)
        emojiGui.SetFont("s10", "Segoe UI")

        ; Add instruction text
        emojiGui.AddText("w350 Center",
            "Select emoji to insert:`n`n1. 🔲 Tasks/Checklist items`n2. ⏳ Time-sensitive tasks`n3. ⚡ First priority`n4. ✅ Check`n5. ❓ Questions/Uncertain items`n6. ℹ️ Info/Notes`n`nType a number (1-6):"
        )

        ; Add input field with auto-submit functionality
        emojiGui.AddEdit("w50 Center vEmojiInput Limit1 Number")

        ; Add OK and Cancel buttons (as backup)
        emojiGui.AddButton("w80 xp-40 y+10", "OK").OnEvent("Click", SubmitEmoji)
        emojiGui.AddButton("w80 xp+90", "Cancel").OnEvent("Click", CancelEmoji)

        ; Set up auto-submit on text change
        emojiGui["EmojiInput"].OnEvent("Change", AutoSubmitEmoji)

        ; Show GUI and focus input
        emojiGui.Show("w350 h200")
        emojiGui["EmojiInput"].Focus()

    } catch Error as e {
        MsgBox "Error in emoji selector: " e.Message, "Emoji Selector Error", "IconX"
    }
}

; Global variables for AI model selector
global gAIModelTargetWin := 0

; AI Model auto-submit function
AutoSubmitAIModel(ctrl, *) {
    currentValue := ctrl.Text
    if (currentValue != "" && IsInteger(currentValue)) {
        choice := Integer(currentValue)
        if (choice >= 1 && choice <= 6) {
            ctrl.Gui.Destroy()
            ExecuteAIModelSelection(choice)
        }
    }
}

; Manual submit function for AI model (backup)
SubmitAIModel(ctrl, *) {
    currentValue := ctrl.Gui["AIModelInput"].Text
    if (currentValue != "" && IsInteger(currentValue)) {
        choice := Integer(currentValue)
        if (choice >= 1 && choice <= 6) {
            ctrl.Gui.Destroy()
            ExecuteAIModelSelection(choice)
        } else {
            MsgBox "Invalid selection. Please choose 1-6.", "AI Model Selection", "IconX"
        }
    }
}

; Cancel function for AI model
CancelAIModel(ctrl, *) {
    ctrl.Gui.Destroy()
}

; Execute the AI model selection logic
ExecuteAIModelSelection(choice) {
    try {
        ; Send Escape twice, then select the edit field based on on-screen Agent/Ask
        SendEscape(2)
        Sleep 200
        if !SendCtrlKeyBasedOnAgentAsk() {
            ; Fallback to Ctrl+I if no relevant text is found
            Send "{Ctrl down}i{Ctrl up}"
        }
        Sleep 300

        ; Handle different behaviors based on choice
        switch choice {
            case 1:
            {
                ; For auto option: simulate ;, wait for model context menu, then send ↓, Enter
                Send "^;"
                Sleep 300
                SendText "auto"
                Sleep 500
                Send "{Enter}"
                Sleep 300
                SendEscape()
            }
            case 2:
            {
                ; For other options: simulate Ctrl + ., wait, type model string, no Enter
                Send "^;"
                Sleep 500
                SendText "CLAUD"
            }
            case 3:
            {
                Send "^;"
                Sleep 500
                SendText "GPT"
            }
            case 4:
            {
                Send "^;"
                Sleep 500
                SendText "O"
            }
            case 5:
            {
                Send "^;"
                Sleep 500
                SendText "DeepSeek"
            }
            case 6:
            {
                Send "^;"
                Sleep 500
                SendText "Cursor"
            }
        }

        Sleep 100

    } catch Error as e {
        MsgBox "Error in AI model selection: " e.Message, "AI Model Selection Error", "IconX"
    }
}

; ; Shift + G : Switch between AI models (Auto-submit version)
; +g::
; {
;     try {
;         ; Remember current target window before showing GUI
;         gAIModelTargetWin := WinExist("A")
;         ; Create GUI for AI model selection with auto-submit
;         aiModelGui := Gui("+AlwaysOnTop +ToolWindow", "AI Model Selection")
;         aiModelGui.SetFont("s10", "Segoe UI")

;         ; Add instruction text
;         aiModelGui.AddText("w350 Center",
;             "Choose AI Model:`n`n1. auto`n2. CLAUD`n3. GPT`n4. O`n5. DeepSeek`n6. Cursor`n`nType a number (1-6):")

;         ; Add input field with auto-submit functionality
;         aiModelGui.AddEdit("w50 Center vAIModelInput Limit1 Number")

;         ; Add OK and Cancel buttons (as backup)
;         aiModelGui.AddButton("w80 xp-40 y+10", "OK").OnEvent("Click", SubmitAIModel)
;         aiModelGui.AddButton("w80 xp+90", "Cancel").OnEvent("Click", CancelAIModel)

;         ; Set up auto-submit on text change
;         aiModelGui["AIModelInput"].OnEvent("Change", AutoSubmitAIModel)

;         ; Show GUI and focus input
;         aiModelGui.Show("w350 h200")
;         aiModelGui["AIModelInput"].Focus()

;     } catch Error as e {
;         MsgBox "Error in AI model selector: " e.Message, "AI Model Selector Error", "IconX"
;     }
; }

; Shift + A : Switch AI models - AI
+a:: {
    if (IsCodeActive()) {
        if (ClickVSCodeCopilotModelButton())
            return
    }
    Send "^;"
}

; Shift + G : Fold all Git directories in Source Control (Cursor) - Git Fold
+g:: FoldAllGitDirectoriesInCursor()

; Global variable for commit selector target window
global gCommitTargetWin := 0

; Function to get commit action by number
GetCommitActionByNumber(numberText) {
    try number := Integer(numberText)
    catch {
        return ""
    }
    actionMap := Map()
    actionMap[1] := "workspace"
    actionMap[2] := "repository"
    return (actionMap.Has(number)) ? actionMap[number] : ""
}

; Function to execute commit action
ExecuteCommitAction(action) {
    if (action = "")
        return

    if (action = "workspace") {
        ; Option 1: Commit and push from workspace (original behavior)
        Send "{Right}"
        Send "{Down}"
        Send "{Tab 2}"
        Send "{Enter}"
        Sleep 1500
        Send "{Tab 2}"
        Send "{Enter}"
        Send "{Up 2}"
    }
    else if (action = "repository") {
        ; Option 2: Commit and push from repository (customize this section)

        ; Then execute the commit commands
        Send "^+g"
        Sleep 150

        ; Click on the "Generate Commit Message (Ctrl+M)" button
        ClickGenerateCommitMessageButton()

        Send "{Tab 3}"
        Send "{Enter}"
        Send "{Tab 2}"
        Send "{Enter}"
        Send "{Up 2}"
    }
}

; Auto-submit function for commit selector
AutoSubmitCommit(ctrl, *) {
    currentValue := ctrl.Text
    if (currentValue != "" && IsInteger(currentValue)) {
        action := GetCommitActionByNumber(currentValue)
        if (action != "") {
            ctrl.Gui.Destroy()
            ExecuteCommitAction(action)
        }
    }
}

; Manual submit function for commit selector (backup)
SubmitCommit(ctrl, *) {
    currentValue := ctrl.Gui["CommitInput"].Text
    if (currentValue != "" && IsInteger(currentValue)) {
        action := GetCommitActionByNumber(currentValue)
        if (action != "") {
            ctrl.Gui.Destroy()
            ExecuteCommitAction(action)
        } else {
            MsgBox "Invalid selection. Please choose 1-2.", "Commit Selector", "IconX"
        }
    }
}

; Cancel function for commit selector
CancelCommit(ctrl, *) {
    ctrl.Gui.Destroy()
}

; Ctrl+, / Ctrl+Q : Fold/Unfold all directories — disabled (UIA breaks VS Code/Cursor).
; Functions kept in hotif_scroll_ai.ahk for possible reuse. Documented on cheat sheet only.

; Alt + N : Review next file - Click the button that contains "Review next file" (Type 50020 Text)
; Path from UIA tree: workbench.parts.editor -> editor-instance -> ... -> Group (anysphere-text-button) -> Text "Review next file"
!n::
{
    try {
        win := WinExist("A")
        if (!win) {
            return
        }
        root := UIA.ElementFromHandle(win)
        Sleep 100  ; Allow UI to update

        ; Strategy 1: Scope to editor part (workbench.parts.editor), find "Review next file" Text, then click its parent Group (the button)
        editorPart := ""
        try editorPart := root.FindFirst({ AutomationId: "workbench.parts.editor", Type: 50026 })
        if (editorPart) {
            reviewText := ""
            try reviewText := editorPart.FindFirst({ Name: "Review next file", Type: 50020 })
            if (reviewText) {
                try {
                    parentBtn := UIA.TreeWalkerTrue.GetParentElement(reviewText)
                    if (parentBtn) {
                        try {
                            if parentBtn.GetPropertyValue(UIA.Property.IsInvokePatternAvailable) {
                                parentBtn.InvokePattern.Invoke()
                            } else {
                                parentBtn.Click()
                            }
                            return
                        } catch {
                            try reviewText.Click()
                            return
                        }
                    }
                } catch {
                    try reviewText.Click()
                    return
                }
            }
        }

        ; Strategy 2: Root-level find by Name "Review next file" or "Review" (Type 50020)
        reviewEl := root.FindFirst({ Name: "Review next file", Type: 50020 })
        if !reviewEl {
            reviewEl := root.FindFirst({ Name: "Review", Type: 50020 })
        }
        if !reviewEl {
            allTexts := root.FindAll({ Type: 50020 })
            for text in allTexts {
                name := ""
                try name := text.Name
                if (name = "Review" || name = "Review next file" || InStr(name, "Review next file")) {
                    reviewEl := text
                    break
                }
            }
        }

        if (reviewEl) {
            ; Prefer clicking parent (the button Group) so the clickable area is used
            try {
                parentBtn := UIA.TreeWalkerTrue.GetParentElement(reviewEl)
                if (parentBtn) {
                    try parentBtn.Click()
                    catch {
                        try reviewEl.Click()
                    }
                    return
                }
            } catch {
            }
            try {
                if reviewEl.GetPropertyValue(UIA.Property.IsInvokePatternAvailable) {
                    reviewEl.InvokePattern.Invoke()
                } else {
                    reviewEl.Click()
                }
            } catch {
                try reviewEl.Click()
            }
        }
    } catch Error as e {
        ; If all else fails, silently fail (no fallback action defined)
    }
}

; Gemini (Chrome): scroll the conversation pane to the bottom (last assistant block or prompt field).
; Uses UIA AutomationId prefix from the page tree (model-response-message-content*); see gemini-tree.md.

; After ScrollIntoView on a message node, still scroll the real viewport (nested scroller / Chrome).
; UIA SetScrollPercent: first arg is vertical %, second horizontal (see UIA.ahk); vertical bottom = (100, -1).
; Keys must go to Chrome_RenderWidgetHostHWND1 so they hit the page, not the top-level frame.
GeminiScroll_ApplyGeminiViewportBottom(uia, scope, hwnd) {
    ; FindFirst throws when missing — use per-attempt try so fallbacks run (efficiency-canon: avoid aborted ladders).
    did := ""
    doc := 0
    try doc := scope.FindFirst({ Type: UIA.Type.Document, Name: "Google Gemini" })
    catch
        doc := 0
    if (!doc) {
        try doc := scope.FindFirst({ Type: "Document" })
        catch
            doc := 0
    }
    if (!doc) {
        try doc := uia.FindFirst({ Type: "Document" })
        catch
            doc := 0
    }
    if (doc) {
        try {
            if (doc.GetPropertyValue(UIA.Property.IsScrollPatternAvailable)) {
                doc.ScrollPattern.SetScrollPercent(100, -1)
                did := "document_SetScrollPercent_v100"
            }
        } catch {
        }
    }
    rw := 0
    try rw := ControlGetHwnd("Chrome_RenderWidgetHostHWND1", "ahk_id " hwnd)
    catch
        rw := 0
    if (did = "" && rw) {
        try {
            ControlSend "{Blind}^{End}", , "ahk_id " rw
            did := "ControlSend_RenderWidget_CtrlEnd"
        } catch {
        }
    }
    if (did = "") {
        try {
            ControlSend "{Blind}^{End}", , "ahk_id " hwnd
            did := "ControlSend_Root_CtrlEnd"
        } catch {
            did := "ControlSend_failed"
        }
    }
    ; Overflow divs often ignore UIA + Ctrl+End — wheel on Chromium surface (single HWND resolve).
    if (rw) {
        try {
            ControlClick "Chrome_RenderWidgetHostHWND1", "ahk_id " hwnd, , , , "NA"
            Sleep 40
            ControlSend "{WheelDown 120}", , "ahk_id " rw
        } catch {
        }
    }
}

GeminiScrollFeedToBottom_Chrome(hwnd) {
    try {
        uia := UIA_Browser("ahk_id " hwnd)
        scope := FastCopyMode_GetGeminiSearchRoot(uia)
        blocks := scope.FindAll({ AutomationId: "model-response-message-content", matchmode: "Substring" })
        if (blocks && blocks.Length > 0) {
            try {
                blocks[blocks.Length].ScrollIntoView()
            } catch {
            }
        }
        GeminiScroll_ApplyGeminiViewportBottom(uia, scope, hwnd)
        if (!blocks || blocks.Length = 0) {
            pf := FindGeminiPromptField(uia)
            if (pf) {
                try {
                    pf.ScrollIntoView()
                } catch {
                }
            }
        }
    } catch {
    }
}

; Easy Selection
; Alt + 1 : Easy Selection - 1st item
!1::
{
    Send "{Enter}"
}

; Alt + 2 : Easy Selection - 2nd item
!2::
{
    Send "{Down}"
    Send "{Enter}"
}

; Alt + 3 : Easy Selection - 3rd item
!3::
{
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Alt + 4 : Easy Selection - 4th item
!4::
{
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Alt + 5 : Easy Selection - 5th item
!5::
{
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}
