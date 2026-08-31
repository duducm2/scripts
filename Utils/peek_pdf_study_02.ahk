; =============================================================================
; Utils module: peek_pdf_study_02.ahk
; Peek PDF / QuickLook study helpers (part 2)
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

StudyTopicSelector_BindRobustEscape() {
    global g_StudyTopicSelectorGui, g_OnEscapePressed, g_StudyTopicEscPollPrev
    SetTimer(StudyTopicSelector_EscapePoll, 0)
    if (!StudyTopicSelector_GuiHasWindow(g_StudyTopicSelectorGui))
        return
    Hotkey("$*Escape", StudyTopicSelector_EscapeFromHotkey, "On")
    global g_OnEscapePressed
    g_OnEscapePressed := StudyTopicSelector_GlobalEscapeCallback
    Utils_EnsureGlobalEscapeHotkey()
    g_StudyTopicEscPollPrev := false
    SetTimer(StudyTopicSelector_EscapePoll, 50)
}

StudyTopicSelector_UnbindRobustEscape() {
    global g_OnEscapePressed, g_StudyTopicEscPollPrev
    SetTimer(StudyTopicSelector_EscapePoll, 0)
    g_StudyTopicEscPollPrev := false
    try Hotkey("Escape", StudyTopicSelector_Cancel, "Off")
    catch {
    }
    try Hotkey("*Escape", StudyTopicSelector_Cancel, "Off")
    catch {
    }
    try Hotkey("$*Escape", StudyTopicSelector_EscapeFromHotkey, "Off")
    catch {
    }
    global g_OnEscapePressed
    g_OnEscapePressed := ""
    Utils_EnsureGlobalEscapeHotkey()
}

; Study material (`#!+x`): Gui.Destroy() throws if the window never existed or was already destroyed — swallow and continue.
StudyTopicSelector_SafeDestroyGui(gui) {
    if (!IsObject(gui))
        return
    try gui.Destroy()
    catch {
    }
}

; AHK v2: after Destroy(), IsObject(gui) is still true but reading .Hwnd throws "Gui has no window".
StudyTopicSelector_GuiHasWindow(gui) {
    if !IsObject(gui)
        return false
    try
        return !!gui.Hwnd
    catch
        return false
}

; Category menu (Technique README / Mnemonics / Plans). Bind Escape + Backspace to cancel; Backspace on topic menu goes back via StudyTopicSelector_BackFromTopic.
StudyTopicSelector_ShowCategoryPhase() {
    global g_StudyTopicSelectorGui, g_StudyTopicSelectorLv, g_StudyTopics, g_StudyTopicSelectorActive,
        g_StudyTopicSelectorLastForegroundMonitorIdx

    StudyTopicSelector_StopActiveMonitorTracking()
    StudyTopicSelector_SafeDestroyGui(g_StudyTopicSelectorGui)
    g_StudyTopicSelectorGui := false
    g_StudyTopicSelectorLv := false

    g_StudyTopicSelectorGui := Gui("+AlwaysOnTop +ToolWindow", "Study material")
    g_StudyTopicSelectorGui.SetFont("s10", "Segoe UI")
    g_StudyTopicSelectorGui.Add("Text", "w520",
        "Char = open   Enter/double-click = open   Backspace/Esc = cancel")
    g_StudyTopicSelectorLv := g_StudyTopicSelectorGui.Add("ListView", "w520 h120 -Multi", ["Char", "Category",
        "Description"])
    g_StudyTopicSelectorLv.OnEvent("DoubleClick", StudyTopicSelector_OnCategoryListActivate)
    g_StudyTopicSelectorGui.Add("Button", "w100", "Close").OnEvent("Click", StudyTopicSelector_Cancel)
    g_StudyTopicSelectorGui.OnEvent("Close", StudyTopicSelector_Cancel)
    g_StudyTopicSelectorGui.OnEvent("Escape", StudyTopicSelector_GuiEscape)

    g_StudyTopicSelectorLv.Add("", "1", "Mnemonics", "Practice notes / GitHub")
    g_StudyTopicSelectorLv.Add("", "2", "Plans", "Study plan checklists")
    g_StudyTopicSelectorLv.Add("", "3", "Technique", "Technique README on GitHub")
    try g_StudyTopicSelectorLv.ModifyCol(1, 50)
    try g_StudyTopicSelectorLv.ModifyCol(2, 120)
    try g_StudyTopicSelectorLv.ModifyCol(3, 320)
    try g_StudyTopicSelectorLv.Modify(1, "Select Focus Vis")
    catch {
    }

    StudyTopicSelector_PositionGuiLikeOutlook(g_StudyTopicSelectorGui)
    try g_StudyTopicSelectorLv.Focus()
    catch {
    }
    g_StudyTopicSelectorActive := true
    g_StudyTopicSelectorLastForegroundMonitorIdx := GetMonitorIndexForForeground_StandardBar()

    Hotkey("1", StudyTopicSelector_SelectMnemonics, "On")
    Hotkey("2", StudyTopicSelector_SelectPlans, "On")
    Hotkey("3", StudyTopicSelector_SelectTechnique, "On")
    Hotkey("Enter", StudyTopicSelector_OnCategoryEnter, "On")
    Hotkey("Backspace", StudyTopicSelector_Cancel, "On")
    StudyTopicSelector_BindRobustEscape()
    SetTimer(StudyTopicSelector_TrackActiveMonitorTick, 115)
}

StudyTopicSelector_SelectedChar() {
    global g_StudyTopicSelectorLv
    if (!IsObject(g_StudyTopicSelectorLv))
        return ""
    row := 0
    try row := g_StudyTopicSelectorLv.GetNext(0, "Focused")
    catch {
        row := 0
    }
    if (row < 1) {
        try row := g_StudyTopicSelectorLv.GetNext(0, "Selected")
        catch {
            row := 0
        }
    }
    if (row < 1)
        return ""
    ch := ""
    try ch := g_StudyTopicSelectorLv.GetText(row, 1)
    catch {
        return ""
    }
    return StrLower(Trim(ch))
}

StudyTopicSelector_OnCategoryEnter(*) {
    global g_StudyTopicSelectorActive, g_StudyTopicSelectorPhase
    if (!g_StudyTopicSelectorActive || g_StudyTopicSelectorPhase != "category")
        return
    ch := StudyTopicSelector_SelectedChar()
    if (ch = "1")
        StudyTopicSelector_SelectMnemonics()
    else if (ch = "2")
        StudyTopicSelector_SelectPlans()
    else if (ch = "3")
        StudyTopicSelector_SelectTechnique()
}

StudyTopicSelector_OnCategoryListActivate(*) {
    StudyTopicSelector_OnCategoryEnter()
}

; After closing Manage Links GUI (Esc) or finishing Open Link from submenu — main category Gui still exists.
StudyTopicSelector_ResumeSelectorEscapeAfterLinks(*) {
    global g_StudyTopicSelectorActive, g_StudyTopicSelectorGui
    if (g_StudyTopicSelectorActive && StudyTopicSelector_GuiHasWindow(g_StudyTopicSelectorGui))
        StudyTopicSelector_BindRobustEscape()
}

StudyTopicSelector_ManageLinksEsc(*) {
    global g_StudyLinksGui
    StudyLink_ClosePalaceManageGui(&g_StudyLinksGui)
}

; Persistent global GUI for link management submenu
global g_StudyLinksGui := false
StudyTopicSelector_ManageLinks(*) {
    global g_StudyLinksGui
    StudyLink_EnsureManageSubtopicSentinel()
    StudyTopicSelector_UnbindRobustEscape()
    StudyTopicSelector_SafeDestroyGui(g_StudyLinksGui)
    ytResult := StudyLink_GetResult(STUDYLINK_KEY_YOUTUBE)
    if (ytResult["ok"])
        StudyLink_PlayApiSuccessSound()
    StudyLink_ShowPalaceManageGui(
        &g_StudyLinksGui,
        "Memory Palace — Study Video",
        "📹 Study Video",
        "Current: " . StudyLink_FormatLinkLabel(ytResult),
        [
            ["1", "Open", "Open the stored video link in Chrome"],
            ["2", "Set from clipboard", "Save the http(s) URL on your clipboard"],
            ["3", "Set manually", "Paste or type a URL"]
        ],
        [
            ["1", StudyTopicSelector_ManageLinks_Open],
            ["2", StudyTopicSelector_ManageLinks_Set],
            ["3", StudyTopicSelector_ManageLinks_SetManual]
        ],
        StudyTopicSelector_ManageLinksEsc
    )
}

; [1] Open the saved YouTube subtopic link in Google Chrome
StudyTopicSelector_ManageLinks_Open(*) {
    StudyTopicSelector_Close()
    linkResult := StudyLink_GetResult(STUDYLINK_KEY_YOUTUBE)
    if (!linkResult["ok"]) {
        ShowCenteredOverlay_Utils("❌ Could not load link from API: " . linkResult["err"], 3500, BANNER_ACCENT_ERROR)
        return
    }
    StudyLink_PlayApiSuccessSound()
    url := linkResult["url"]
    if (url != "") {
        StudyLink_OpenUrlInChrome(url, true)
        ShowCenteredOverlay_Utils("✅ Opening YouTube link in a new Chrome window...", 2000, BANNER_ACCENT_SUCCESS)
    } else {
        ShowCenteredOverlay_Utils("⚠ No YouTube link stored. Use [2] Set YouTube link first.", 2500,
            BANNER_ACCENT_INTERMEDIATE)
    }
}

StudyLink_UiaInvokeOrClick(el, preferClick := false) {
    if !IsObject(el)
        return false
    global STUDYLINK_YT_MS_BEFORE_CLICK, STUDYLINK_YT_MS_AFTER_CLICK
    Sleep STUDYLINK_YT_MS_BEFORE_CLICK
    try el.SetFocus()
    catch {
    }
    if (preferClick) {
        try {
            el.Click()
            Sleep STUDYLINK_YT_MS_AFTER_CLICK
            return true
        } catch {
        }
    }
    try {
        if el.GetPropertyValue(UIA.Property.IsInvokePatternAvailable) {
            el.Invoke()
            Sleep STUDYLINK_YT_MS_AFTER_CLICK
            return true
        }
    } catch {
    }
    try {
        el.Click()
        Sleep STUDYLINK_YT_MS_AFTER_CLICK
        return true
    } catch {
    }
    return false
}

StudyLink_UiaWaitFor(root, conditions, timeoutMs := 4000) {
    if !IsObject(root)
        return 0
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        el := ClipAngel_UiaFindFirst(root, conditions)
        if el
            return el
        Sleep 80
    }
    return 0
}

; EN + PT-BR YouTube action-bar / share-panel button labels (Name property).
global STUDYLINK_YT_BTN_SHARE := ["Share", "Compartilhar"]
global STUDYLINK_YT_BTN_COPY := ["Copy", "Copiar"]
global STUDYLINK_YT_BTN_CANCEL := ["Cancel", "Cancelar", "Close", "Fechar"]

; UI pacing (ms) — modest gaps so focus/invoke/click registers on YouTube Web UI.
global STUDYLINK_YT_MS_BEFORE_CLICK := 120
global STUDYLINK_YT_MS_AFTER_CLICK := 80
global STUDYLINK_YT_MS_AFTER_SHARE_CLICK := 350
global STUDYLINK_YT_MS_AFTER_START_AT := 250
global STUDYLINK_YT_MS_AFTER_COPY := 450
global STUDYLINK_YT_MS_BEFORE_CLOSE_PANEL := 280

StudyLink_UiaFindButtonByNames(root, nameList, timeoutMs := 0) {
    if !IsObject(root)
        return 0
    deadline := timeoutMs ? A_TickCount + timeoutMs : 0
    loop {
        for name in nameList {
            el := ClipAngel_UiaFindFirst(root, { Type: 50000, Name: name })
            if el
                return el
        }
        if (!timeoutMs || A_TickCount >= deadline)
            break
        Sleep 80
    }
    return 0
}

StudyLink_IsYoutubeVideoPageUrl(url) {
    return InStr(url, "youtube.com/watch") || InStr(url, "youtube.com/live") || InStr(url, "youtube.com/shorts")
}

StudyLink_CaptureYoutubeTimestampUrl(uia, &errMsg := "") {
    errMsg := ""
    global STUDYLINK_YT_MS_AFTER_SHARE_CLICK, STUDYLINK_YT_MS_AFTER_START_AT
    if !IsObject(uia) {
        errMsg := "Could not attach to Chrome."
        return ""
    }
    try currentUrl := uia.GetCurrentURL()
    catch {
        errMsg := "Could not read the current page URL."
        return ""
    }
    if !StudyLink_IsYoutubeVideoPageUrl(currentUrl) {
        errMsg := "Open a YouTube video page and try again."
        return ""
    }
    shareBtn := StudyLink_UiaFindButtonByNames(uia, STUDYLINK_YT_BTN_SHARE, 3000)
    if !shareBtn {
        errMsg := "Share button not found."
        return ""
    }
    if !StudyLink_UiaInvokeOrClick(shareBtn) {
        errMsg := "Could not open the Share panel."
        return ""
    }
    Sleep STUDYLINK_YT_MS_AFTER_SHARE_CLICK
    startAt := StudyLink_UiaWaitFor(uia, { AutomationId: "start-at-checkbox", Type: 50002 })
    if !startAt {
        errMsg := "Share panel did not open in time."
        return ""
    }
    startAtOn := ClipAngel_FavoriteCellIsOn(startAt)
    if !startAtOn {
        toggled := false
        try {
            if startAt.GetPropertyValue(UIA.Property.IsTogglePatternAvailable) {
                startAt.TogglePattern.Toggle()
                toggled := true
            }
        } catch {
        }
        if (!toggled)
            toggled := StudyLink_UiaInvokeOrClick(startAt)
        if (!toggled) {
            errMsg := "Could not enable Start at."
            return ""
        }
        Sleep STUDYLINK_YT_MS_AFTER_START_AT
    }
    deadline := A_TickCount + 2000
    while (A_TickCount < deadline) {
        shareUrlEl := ClipAngel_UiaFindFirst(uia, { AutomationId: "share-url", Type: 50004 })
        if shareUrlEl {
            try shareVal := Trim(shareUrlEl.Value)
            catch
                shareVal := ""
            if (shareVal != "" && InStr(shareVal, "t="))
                break
        }
        Sleep 80
    }
    shareUrlEl := StudyLink_UiaWaitFor(uia, { AutomationId: "share-url", Type: 50004 }, 2000)
    if !shareUrlEl {
        errMsg := "Share link field not found."
        return ""
    }
    url := ""
    try url := Trim(shareUrlEl.Value)
    catch
        url := ""
    copyBtn := 0
    copyGroup := ClipAngel_UiaFindFirst(uia, { AutomationId: "copy-button" })
    if copyGroup {
        for name in STUDYLINK_YT_BTN_COPY {
            copyBtn := ClipAngel_UiaFindFirst(copyGroup, { Type: 50000, Name: name })
            if copyBtn
                break
        }
    }
    if !copyBtn
        copyBtn := StudyLink_UiaFindButtonByNames(uia, STUDYLINK_YT_BTN_COPY)
    if !copyBtn {
        errMsg := "Copy button not found."
        return ""
    }
    savedClip := A_Clipboard
    if !StudyLink_UiaInvokeOrClick(copyBtn) {
        errMsg := "Could not click Copy."
        return ""
    }
    clipDeadline := A_TickCount + 2000
    clipUrl := ""
    while (A_TickCount < clipDeadline) {
        clipUrl := Trim(A_Clipboard)
        if (clipUrl != "" && clipUrl != savedClip && InStr(clipUrl, "t="))
            break
        Sleep 80
    }
    if (url = "" || !InStr(url, "t=")) {
        if (clipUrl != "" && InStr(clipUrl, "t="))
            url := clipUrl
    }
    if (url = "" || !InStr(url, "youtu") || !InStr(url, "t=")) {
        errMsg := "Could not read the timestamped share link."
        return ""
    }
    global STUDYLINK_YT_MS_AFTER_COPY
    Sleep STUDYLINK_YT_MS_AFTER_COPY
    return url
}

StudyLink_CleanupYoutubeSharePanel(uia, chromeHwnd := 0) {
    global STUDYLINK_YT_MS_BEFORE_CLOSE_PANEL, STUDYLINK_YT_BTN_CANCEL
    if !IsObject(uia)
        return
    if !ClipAngel_UiaFindFirst(uia, { AutomationId: "start-at-checkbox", Type: 50002 })
    && !ClipAngel_UiaFindFirst(uia, { AutomationId: "share-url", Type: 50004 })
        return
    if chromeHwnd {
        try {
            WinActivate("ahk_id " chromeHwnd)
            WinWaitActive("ahk_id " chromeHwnd, , 1)
        } catch {
        }
    }
    Sleep STUDYLINK_YT_MS_BEFORE_CLOSE_PANEL
    cancelBtn := 0
    closeGroup := ClipAngel_UiaFindFirst(uia, { AutomationId: "close-button" })
    if closeGroup {
        for name in STUDYLINK_YT_BTN_CANCEL {
            cancelBtn := ClipAngel_UiaFindFirst(closeGroup, { Type: 50000, Name: name })
            if cancelBtn
                break
        }
        if !cancelBtn {
            try cancelBtn := closeGroup.FindFirst({ Type: 50000, AutomationId: "button" })
            catch {
                cancelBtn := 0
            }
        }
    }
    if !cancelBtn
        cancelBtn := StudyLink_UiaFindButtonByNames(uia, STUDYLINK_YT_BTN_CANCEL)
    if cancelBtn
        StudyLink_UiaInvokeOrClick(cancelBtn, true)
    Sleep 220
    panelStillOpen := !!ClipAngel_UiaFindFirst(uia, { AutomationId: "share-url", Type: 50004 })
    if (!panelStillOpen)
        return
    if chromeHwnd {
        try {
            WinActivate("ahk_id " chromeHwnd)
            WinWaitActive("ahk_id " chromeHwnd, , 1)
        } catch {
        }
    }
    try Send "{Escape}"
    catch {
    }
    Sleep 250
}

; [2] Set the link from clipboard
StudyTopicSelector_ManageLinks_Set(*) {
    StudyTopicSelector_Close()
    StudyLink_SetFromClipboard(STUDYLINK_KEY_YOUTUBE, "video link")
}

; [3] Set YouTube link manually (InputBox → StudyLink_Set API)
StudyTopicSelector_ManageLinks_SetManual(*) {
    StudyTopicSelector_Close()
    StudyLink_SetFromManualInput(STUDYLINK_KEY_YOUTUBE, "YouTube link")
}

ShowStudyTopicSelector() {
    global g_StudyTopicSelectorGui, g_StudyTopicSelectorActive, g_StudyTopics, g_StudyTopicSelectorPhase,
        g_StudyTopicSelectorCategory, g_StudyTopicSelectorLastForegroundMonitorIdx

    if (g_StudyTopicSelectorActive) {
        if (StudyTopicSelector_GuiHasWindow(g_StudyTopicSelectorGui))
            return
        ; Sticky Active after a Gui destroy race — clear and reopen.
        StudyTopicSelector_ForceReset()
    }

    g_StudyTopicSelectorCategory := ""
    g_StudyTopicSelectorPhase := "category"

    StudyTopicSelector_ShowCategoryPhase()
}

StudyTopicSelector_ShowTopicPhase() {
    global g_StudyTopicSelectorActive, g_StudyTopicSelectorPhase, g_StudyTopicSelectorPendingRemove

    if (!g_StudyTopicSelectorActive || g_StudyTopicSelectorPhase != "category")
        return

    StudyTopicSelector_UnbindCategoryHotkeys()
    try Hotkey("Backspace", StudyTopicSelector_Cancel, "Off")
    g_StudyTopicSelectorPendingRemove := false
    g_StudyTopicSelectorPhase := "topic"
    StudyTopicSelector_RebuildTopicPhase()
}

; Re-render topic list while phase == "topic" (open, after add/remove, clear pending-remove).
StudyTopicSelector_RebuildTopicPhase() {
    global g_StudyTopicSelectorGui, g_StudyTopicSelectorLv, g_StudyTopicSelectorActive, g_StudyTopicSelectorPhase,
        g_StudyTopicSelectorCategory, g_StudyTopicSelectorLastForegroundMonitorIdx,
        g_StudyTopicSelectorPendingRemove

    if (!g_StudyTopicSelectorActive || g_StudyTopicSelectorPhase != "topic")
        return

    ; Pause tracker across destroy/recreate so .Hwnd is never read on a destroyed Gui.
    StudyTopicSelector_StopActiveMonitorTracking()
    StudyTopicSelector_UnbindDigitHotkeys()
    try Hotkey("Backspace", "Off")

    catLabel := (g_StudyTopicSelectorCategory = "plans") ? "Plans" : "Mnemonics"
    entries := StudyLinks_GetEntries(g_StudyTopicSelectorCategory)

    StudyTopicSelector_SafeDestroyGui(g_StudyTopicSelectorGui)
    g_StudyTopicSelectorGui := false
    g_StudyTopicSelectorLv := false

    winTitle := g_StudyTopicSelectorPendingRemove
        ? catLabel . " — pick to REMOVE"
            : catLabel . " — topic"
    hint := g_StudyTopicSelectorPendingRemove
        ? "Char = remove   Enter/double-click = remove   Esc/Backspace = cancel remove"
            : "Char = open   Enter/double-click = open   a add   r remove   Backspace = back   Esc = cancel"

    g_StudyTopicSelectorGui := Gui("+AlwaysOnTop +ToolWindow", winTitle)
    g_StudyTopicSelectorGui.SetFont("s10", "Segoe UI")
    g_StudyTopicSelectorGui.Add("Text", "w620", hint)
    g_StudyTopicSelectorLv := g_StudyTopicSelectorGui.Add("ListView", "w620 h280 -Multi", ["Char", "Name",
        "Detail"])
    g_StudyTopicSelectorLv.OnEvent("DoubleClick", StudyTopicSelector_OnTopicListActivate)
    g_StudyTopicSelectorGui.Add("Button", "w100", "Close").OnEvent("Click", StudyTopicSelector_Cancel)
    g_StudyTopicSelectorGui.OnEvent("Close", StudyTopicSelector_Cancel)
    g_StudyTopicSelectorGui.OnEvent("Escape", StudyTopicSelector_GuiEscape)

    if (!g_StudyTopicSelectorPendingRemove) {
        g_StudyTopicSelectorLv.Add("", "a", "Add entry", "Name + URL")
        g_StudyTopicSelectorLv.Add("", "r", "Arm remove", "Next item key deletes")
    }

    if (entries.Length = 0) {
        g_StudyTopicSelectorLv.Add("", "", "(no entries)", "press a to add")
    } else {
        for i, entry in entries {
            if (i > 33)
                break
            label := StudyLinks_LabelForIndex(i)
            g_StudyTopicSelectorLv.Add("", label, entry.name, "")
        }
    }

    try g_StudyTopicSelectorLv.ModifyCol(1, 50)
    try g_StudyTopicSelectorLv.ModifyCol(2, 280)
    try g_StudyTopicSelectorLv.ModifyCol(3, 260)
    focusRow := g_StudyTopicSelectorPendingRemove ? 1 : 3
    if (focusRow > g_StudyTopicSelectorLv.GetCount())
        focusRow := 1
    if (g_StudyTopicSelectorLv.GetCount() > 0) {
        try g_StudyTopicSelectorLv.Modify(focusRow, "Select Focus Vis")
        catch {
        }
    }

    StudyTopicSelector_PositionGuiLikeOutlook(g_StudyTopicSelectorGui)
    try g_StudyTopicSelectorLv.Focus()
    catch {
    }
    g_StudyTopicSelectorLastForegroundMonitorIdx := GetMonitorIndexForForeground_StandardBar()

    count := Min(entries.Length, 33)
    loop count {
        label := StudyLinks_LabelForIndex(A_Index)
        Hotkey(label, StudyTopicSelector_HandleKey, "On")
    }
    Hotkey("a", StudyTopicSelector_AddEntry, "On")
    Hotkey("r", StudyTopicSelector_ArmRemove, "On")
    Hotkey("Enter", StudyTopicSelector_OnTopicEnter, "On")
    Hotkey("Backspace", StudyTopicSelector_BackFromTopic, "On")
    StudyTopicSelector_BindRobustEscape()
    if (g_StudyTopicSelectorActive)
        SetTimer(StudyTopicSelector_TrackActiveMonitorTick, 115)
}

StudyTopicSelector_OnTopicEnter(*) {
    global g_StudyTopicSelectorActive, g_StudyTopicSelectorPhase
    if (!g_StudyTopicSelectorActive || g_StudyTopicSelectorPhase != "topic")
        return
    ch := StudyTopicSelector_SelectedChar()
    if (ch = "")
        return
    if (ch = "a") {
        StudyTopicSelector_AddEntry()
        return
    }
    if (ch = "r") {
        StudyTopicSelector_ArmRemove()
        return
    }
    StudyTopicSelector_HandleKey(ch)
}

StudyTopicSelector_OnTopicListActivate(*) {
    StudyTopicSelector_OnTopicEnter()
}

StudyTopicSelector_BackFromTopic(*) {
    global g_StudyTopicSelectorGui, g_StudyTopicSelectorLv, g_StudyTopicSelectorActive, g_StudyTopicSelectorPhase,
        g_StudyTopicSelectorCategory,
        g_StudyTopicSelectorLastForegroundMonitorIdx, g_StudyTopicSelectorPendingRemove

    if (!g_StudyTopicSelectorActive || g_StudyTopicSelectorPhase != "topic")
        return
    if (g_StudyTopicSelectorPendingRemove) {
        g_StudyTopicSelectorPendingRemove := false
        StudyTopicSelector_RebuildTopicPhase()
        return
    }
    StudyTopicSelector_UnbindDigitHotkeys()
    try Hotkey("Backspace", StudyTopicSelector_BackFromTopic, "Off")
    g_StudyTopicSelectorCategory := ""
    g_StudyTopicSelectorPhase := "category"
    StudyTopicSelector_SafeDestroyGui(g_StudyTopicSelectorGui)
    g_StudyTopicSelectorGui := false
    g_StudyTopicSelectorLv := false
    StudyTopicSelector_ShowCategoryPhase()
    g_StudyTopicSelectorLastForegroundMonitorIdx := GetMonitorIndexForForeground_StandardBar()
}

; [3] Technique: open technique README on GitHub in a new Chrome window (no scroll-to-end).
StudyTopicSelector_SelectTechnique(*) {
    global g_StudyTopicSelectorActive, g_StudyTopicSelectorPhase, g_StudyTopics
    if (!g_StudyTopicSelectorActive || g_StudyTopicSelectorPhase != "category")
        return
    if (!g_StudyTopics.Has(0))
        return
    StudyTopicSelector_Close()
    topic := g_StudyTopics[0]
    StudyTopic_OpenGithubInChrome(StudyTopic_GetUrl(topic, "mnemonics"), false)
}

StudyTopicSelector_SelectMnemonics(*) {
    global g_StudyTopicSelectorActive, g_StudyTopicSelectorPhase, g_StudyTopicSelectorCategory
    if (!g_StudyTopicSelectorActive || g_StudyTopicSelectorPhase != "category")
        return
    g_StudyTopicSelectorCategory := "mnemonics"
    StudyTopicSelector_ShowTopicPhase()
}

StudyTopicSelector_SelectPlans(*) {
    global g_StudyTopicSelectorActive, g_StudyTopicSelectorPhase, g_StudyTopicSelectorCategory
    if (!g_StudyTopicSelectorActive || g_StudyTopicSelectorPhase != "category")
        return
    g_StudyTopicSelectorCategory := "plans"
    StudyTopicSelector_ShowTopicPhase()
}

StudyTopicSelector_ArmRemove(*) {
    global g_StudyTopicSelectorActive, g_StudyTopicSelectorPhase, g_StudyTopicSelectorPendingRemove
    if (!g_StudyTopicSelectorActive || g_StudyTopicSelectorPhase != "topic")
        return
    g_StudyTopicSelectorPendingRemove := true
    StudyTopicSelector_RebuildTopicPhase()
}

StudyTopicSelector_AddEntry(*) {
    global g_StudyTopicSelectorActive, g_StudyTopicSelectorPhase, g_StudyTopicSelectorCategory,
        g_StudyTopicSelectorPendingRemove

    if (!g_StudyTopicSelectorActive || g_StudyTopicSelectorPhase != "topic")
        return
    g_StudyTopicSelectorPendingRemove := false
    StudyTopicSelector_UnbindDigitHotkeys()
    StudyTopicSelector_UnbindRobustEscape()

    catLabel := (g_StudyTopicSelectorCategory = "plans") ? "Plans" : "Mnemonics"
    nameBox := InputBox("Name for the new " . catLabel . " entry:", "Add study link", "w400 h120")
    if (nameBox.Result != "OK") {
        StudyTopicSelector_RebuildTopicPhase()
        return
    }
    name := Trim(nameBox.Value)
    if (name = "") {
        try ShowCenteredOverlay_Utils("⚠ Name cannot be empty.", 2500, BANNER_ACCENT_INTERMEDIATE)
        StudyTopicSelector_RebuildTopicPhase()
        return
    }

    urlBox := InputBox("GitHub (or http) URL:", "Add study link — URL", "w500 h120")
    if (urlBox.Result != "OK") {
        StudyTopicSelector_RebuildTopicPhase()
        return
    }
    url := Trim(urlBox.Value)
    if (url = "" || !(SubStr(url, 1, 7) = "http://" || SubStr(url, 1, 8) = "https://")) {
        try ShowCenteredOverlay_Utils("❌ URL must start with http:// or https://", 3000, BANNER_ACCENT_ERROR)
        StudyTopicSelector_RebuildTopicPhase()
        return
    }

    if StudyLinks_AddEntry(g_StudyTopicSelectorCategory, name, url)
        try ShowCenteredOverlay_Utils("✅ Added: " . name, 2000, BANNER_ACCENT_SUCCESS)
    StudyTopicSelector_RebuildTopicPhase()
}

StudyTopicSelector_HandleKey(thisHotkey := "") {
    global g_StudyTopicSelectorActive, g_StudyTopicSelectorPhase, g_StudyTopicSelectorCategory,
        g_StudyTopicSelectorPendingRemove

    if (!g_StudyTopicSelectorActive || g_StudyTopicSelectorPhase != "topic")
        return

    key := thisHotkey
    if (key = "")
        key := A_ThisHotkey
    ; Strip optional $* / modifiers if present (defensive).
    key := RegExReplace(key, "^[\$\*]*", "")
    key := StrLower(key)

    idx := StudyLinks_IndexFromKey(key)
    if (idx < 1)
        return

    category := g_StudyTopicSelectorCategory
    entries := StudyLinks_GetEntries(category)
    if (idx > entries.Length)
        return

    if (g_StudyTopicSelectorPendingRemove) {
        removedName := entries[idx].name
        g_StudyTopicSelectorPendingRemove := false
        if StudyLinks_RemoveEntry(category, idx)
            try ShowCenteredOverlay_Utils("✅ Removed: " . removedName, 2000, BANNER_ACCENT_SUCCESS)
        StudyTopicSelector_RebuildTopicPhase()
        return
    }

    StudyTopicSelector_Close()
    url := entries[idx].url
    StudyTopic_OpenGithubInChrome(url, category = "mnemonics")
}

StudyTopicSelector_Cancel(*) {
    global g_StudyTopicSelectorActive, g_StudyTopicSelectorPhase, g_StudyTopicSelectorPendingRemove
    if (g_StudyTopicSelectorActive && g_StudyTopicSelectorPhase = "topic" && g_StudyTopicSelectorPendingRemove) {
        g_StudyTopicSelectorPendingRemove := false
        StudyTopicSelector_RebuildTopicPhase()
        return
    }
    StudyTopicSelector_Close()
}

; Always completes teardown even if Active was already false or Gui already destroyed (retry-safe).
StudyTopicSelector_ForceReset() {
    global g_StudyTopicSelectorGui, g_StudyTopicSelectorLv, g_StudyTopicSelectorActive, g_StudyTopicSelectorPhase,
        g_StudyTopicSelectorCategory,
        g_StudyTopicSelectorLastForegroundMonitorIdx, g_StudyLinksGui, g_StudyArticleLinksGui, g_StudyFavoriteLinksGui,
        g_StudyLinkSubmenuGui, g_StudyTopicSelectorPendingRemove

    g_StudyTopicSelectorActive := false
    g_StudyTopicSelectorPhase := ""
    g_StudyTopicSelectorCategory := ""
    g_StudyTopicSelectorPendingRemove := false
    g_StudyTopicSelectorLastForegroundMonitorIdx := 0

    try StudyTopicSelector_StopActiveMonitorTracking()
    try StudyTopicSelector_UnbindRobustEscape()
    try StudyTopicSelector_UnbindCategoryHotkeys()
    try StudyTopicSelector_UnbindDigitHotkeys()
    try Hotkey("Backspace", "Off")
    catch {
    }
    try Utils_EnsureGlobalEscapeHotkey()
    catch {
    }
    try StudyTopicSelector_SafeDestroyGui(g_StudyTopicSelectorGui)
    g_StudyTopicSelectorGui := false
    g_StudyTopicSelectorLv := false
    try StudyTopicSelector_SafeDestroyGui(g_StudyLinksGui)
    g_StudyLinksGui := false
    try StudyTopicSelector_SafeDestroyGui(g_StudyArticleLinksGui)
    g_StudyArticleLinksGui := false
    try StudyTopicSelector_SafeDestroyGui(g_StudyFavoriteLinksGui)
    g_StudyFavoriteLinksGui := false
    try StudyTopicSelector_SafeDestroyGui(g_StudyLinkSubmenuGui)
    g_StudyLinkSubmenuGui := ""
    loop 9 {
        try Hotkey(String(A_Index), "Off")
    }
    try Hotkey("Escape", "Off")
    catch {
    }
}

; Tear-down order aligned with ShowAiModelSelector close (Utils\handy_selector_hotkey.ahk).
StudyTopicSelector_Close() {
    global g_StudyTopicSelectorActive
    if (!g_StudyTopicSelectorActive)
        return
    StudyTopicSelector_ForceReset()
}

PeekPdf_GetIniPath() {
    return A_ScriptDir "\assets\data\peek_pdf.ini"
}

PeekPdf_NormalizePath(path) {
    path := Trim(path)
    q := Chr(34)
    if (SubStr(path, 1, 1) = q && SubStr(path, -1) = q)
        path := SubStr(path, 2, StrLen(path) - 2)
    return Trim(path)
}

; Resolve the Peek executable path.
; Priority: 1) INI [Peek] ExePath  2) Environment-specific path (GetPeekExePath)  3) "peek.exe" (PATH)
PeekPdf_ResolvePeekExePath() {
    iniPath := PeekPdf_GetIniPath()
    exePath := ""
    try exePath := IniRead(iniPath, "Peek", "ExePath", "")
    exePath := PeekPdf_NormalizePath(exePath)
    if (exePath != "" && FileExist(exePath))
        return exePath
    ; Legacy: this previously used GetPeekExePath() and PowerToys Peek.
    ; QuickLook is now the primary study viewer; for compatibility, fall back to QuickLook.
    envExe := GetQuickLookExePath()
    if (FileExist(envExe))
        return envExe
    return "QuickLook.exe"
}

QuickLook_GetIniPath() {
    return A_ScriptDir "\assets\data\quicklook.ini"
}

QuickLook_NormalizePath(path) {
    path := Trim(path)
    q := Chr(34)
    if (SubStr(path, 1, 1) = q && SubStr(path, -1) = q)
        path := SubStr(path, 2, StrLen(path) - 2)
    return Trim(path)
}

; Note: `GetQuickLookExePath()` is defined in `env.ahk` (environment-specific paths).

QuickLook_ResolveExePath() {
    iniPath := QuickLook_GetIniPath()
    exePath := ""
    try exePath := IniRead(iniPath, "QuickLook", "ExePath", "")
    exePath := QuickLook_NormalizePath(exePath)
    if (exePath != "" && FileExist(exePath))
        return exePath
    envExe := GetQuickLookExePath()
    if (FileExist(envExe))
        return envExe
    return "QuickLook.exe"
}

; Detect whether a given monitor index is currently connected and has a usable work area.
; Uses AHK built-ins so we don't depend on custom geometry assumptions.
IsMonitorConnected(monitorIndex) {
    try monitorCount := MonitorGetCount()
    catch
        return false
    if (monitorIndex < 1 || monitorIndex > monitorCount)
        return false
    try {
        MonitorGetWorkArea(monitorIndex, &l, &t, &r, &b)
        if ((r - l) <= 0 || (b - t) <= 0)
            return false
    } catch {
        return false
    }
    return true
}
