; =============================================================================
; Utils module: peek_pdf_study_02.ahk
; Peek PDF / QuickLook study helpers (part 2)
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

StudyTopicSelector_BindRobustEscape() {
    global g_StudyTopicSelectorGui, g_OnEscapePressed, g_StudyTopicEscPollPrev
    SetTimer(StudyTopicSelector_EscapePoll, 0)
    if (!IsObject(g_StudyTopicSelectorGui) || !g_StudyTopicSelectorGui.Hwnd)
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

; Category menu (Technique README / Mnemonics / Plans). Bind Escape + Backspace to cancel; Backspace on topic menu goes back via StudyTopicSelector_BackFromTopic.
StudyTopicSelector_ShowCategoryPhase() {
    global g_StudyTopicSelectorGui, g_StudyTopics

    StudyTopicSelector_SafeDestroyGui(g_StudyTopicSelectorGui)
    g_StudyTopicSelectorGui := false

    g_StudyTopicSelectorGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Owner -DPIScale")
    g_StudyTopicSelectorGui.BackColor := "1E1E2E"
    g_StudyTopicSelectorGui.MarginX := 20
    g_StudyTopicSelectorGui.MarginY := 15

    g_StudyTopicSelectorGui.SetFont("s14 cCDD6F4 Bold", "Segoe UI")
    g_StudyTopicSelectorGui.Add("Text", "w300 Center", "📚 Study material")
    g_StudyTopicSelectorGui.Add("Text", "w300 h1 Background45475A")

    g_StudyTopicSelectorGui.SetFont("s12 cCDD6F4", "Segoe UI")
    g_StudyTopicSelectorGui.Add("Text", "w300", "[1] Mnemonics")
    g_StudyTopicSelectorGui.Add("Text", "w300", "[2] Plans")
    g_StudyTopicSelectorGui.Add("Text", "w300", "[3] Manage Study Subtopic Link 📽️")
    g_StudyTopicSelectorGui.Add("Text", "w300", "[4] Manage Study Article Link 📖")
    g_StudyTopicSelectorGui.Add("Text", "w300", "[5] Manage Study Favorite Link ❤️?")
    g_StudyTopicSelectorGui.Add("Text", "w300", "[6] Technique")
    g_StudyTopicSelectorGui.Add("Text", "w300 h1 Background45475A y+10")
    g_StudyTopicSelectorGui.SetFont("s9 c6C7086", "Segoe UI")
    g_StudyTopicSelectorGui.Add("Text", "w300 Center", "Press 1-6 | Backspace/Esc to cancel")

    try {
        g_StudyTopicSelectorGui.OnEvent("Escape", StudyTopicSelector_GuiEscape)
    } catch {
    }
    StudyTopicSelector_PositionGuiLikeOutlook(g_StudyTopicSelectorGui)
    global g_StudyTopicSelectorActive
    g_StudyTopicSelectorActive := true

    Hotkey("1", StudyTopicSelector_SelectMnemonics, "On")
    Hotkey("2", StudyTopicSelector_SelectPlans, "On")
    Hotkey("3", StudyTopicSelector_ManageLinks, "On")
    Hotkey("4", StudyTopicSelector_ManageArticleLinks, "On")
    Hotkey("5", StudyTopicSelector_ManageFavoriteLinks, "On")
    Hotkey("6", StudyTopicSelector_SelectTechnique, "On")
    Hotkey("Backspace", StudyTopicSelector_Cancel, "On")
    StudyTopicSelector_BindRobustEscape()

}

; After closing Manage Links GUI (Esc) or finishing Open Link from submenu — main category Gui still exists.
StudyTopicSelector_ResumeSelectorEscapeAfterLinks(*) {
    global g_StudyTopicSelectorActive, g_StudyTopicSelectorGui
    if (g_StudyTopicSelectorActive && IsObject(g_StudyTopicSelectorGui) && g_StudyTopicSelectorGui.Hwnd)
        StudyTopicSelector_BindRobustEscape()
}

StudyTopicSelector_ManageLinksEsc(*) {
    global g_StudyLinksGui
    StudyTopicSelector_SafeDestroyGui(g_StudyLinksGui)
    g_StudyLinksGui := false
    StudyTopicSelector_ResumeSelectorEscapeAfterLinks()
}

; Persistent global GUI for link management submenu
global g_StudyLinksGui := false
StudyTopicSelector_ManageLinks(*) {
    global g_StudyLinksGui
    StudyLink_EnsureManageSubtopicSentinel()
    StudyTopicSelector_UnbindRobustEscape()
    StudyTopicSelector_SafeDestroyGui(g_StudyLinksGui)
    g_StudyLinksGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Owner -DPIScale")
    ; Unbind previous hotkeys to avoid conflicts
    loop 9 {
        try Hotkey(String(A_Index), "Off")
    }
    try Hotkey("Escape", "Off")
    g_StudyLinksGui.BackColor := "1E1E2E"
    g_StudyLinksGui.MarginX := 20
    g_StudyLinksGui.MarginY := 15
    g_StudyLinksGui.SetFont("s14 cCDD6F4 Bold", "Segoe UI")
    g_StudyLinksGui.Add("Text", "w400 Center", "🔗 Manage Study Subtopic Link")
    g_StudyLinksGui.Add("Text", "w400 h1 Background45475A")
    g_StudyLinksGui.SetFont("s11 cCDD6F4", "Segoe UI")
    ytResult := StudyLink_GetResult(STUDYLINK_KEY_YOUTUBE)
    if (ytResult["ok"])
        StudyLink_PlayApiSuccessSound()
    g_StudyLinksGui.Add("Text", "w400", "Current YouTube link: " . StudyLink_FormatLinkLabel(ytResult))
    g_StudyLinksGui.Add("Text", "w400", "[1] Open YouTube link")
    g_StudyLinksGui.Add("Text", "w400", "[2] Set YouTube link")
    g_StudyLinksGui.Add("Text", "w400 h1 Background45475A y+10")
    g_StudyLinksGui.SetFont("s9 c6C7086", "Segoe UI")
    g_StudyLinksGui.Add("Text", "w400 Center", "Press 1-2 | Esc to cancel")
    StudyTopicSelector_PositionGuiLikeOutlook(g_StudyLinksGui)
    Hotkey("1", StudyTopicSelector_ManageLinks_Open, "On")
    Hotkey("2", StudyTopicSelector_ManageLinks_Set, "On")
    Hotkey("Escape", StudyTopicSelector_ManageLinksEsc, "On")
    g_StudyLinksGui.Show()
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

; [2] Set the link: YouTube Share + Start at + Copy, save via API
StudyTopicSelector_ManageLinks_Set(*) {
    StudyTopicSelector_Close()
    try {
        chromeHwnd := WinExist("ahk_class Chrome_WidgetWin_1")
        if chromeHwnd {
            WinActivate("ahk_id " chromeHwnd)
            if !WinWaitActive("ahk_id " chromeHwnd, , 2) {
                ShowCenteredOverlay_Utils("❌ Chrome window did not become active.", 2500, BANNER_ACCENT_ERROR)
                return
            }
        } else {
            ShowCenteredOverlay_Utils("❌ Chrome window not found. Switch to Chrome and try again.", 3000,
                BANNER_ACCENT_ERROR)
            return
        }
        Sleep 280
        try UIA.ActivateChromiumAccessibility("ahk_id " chromeHwnd, 300)
        catch {
        }
        uia := UIA_Browser("ahk_id " chromeHwnd)
        errMsg := ""
        url := StudyLink_CaptureYoutubeTimestampUrl(uia, &errMsg)
        if (url != "") {
            ; Close share panel while Chrome still has focus (before loading overlay steals it).
            StudyLink_CleanupYoutubeSharePanel(uia, chromeHwnd)
            StandardLoadingBar_Show("Saving link…", BANNER_ACCENT_INTERMEDIATE, { passive: false })
            setOk := StudyLink_Set(STUDYLINK_KEY_YOUTUBE, url)
            try StandardLoadingBar_Hide(0)
            if setOk
                ShowCenteredOverlay_Utils("✅ Link saved to study notes.", 3000, BANNER_ACCENT_SUCCESS)
            else
                ShowCenteredOverlay_Utils("❌ Could not save the link (API failed).", 3500, BANNER_ACCENT_ERROR)
        } else {
            StudyLink_CleanupYoutubeSharePanel(uia, chromeHwnd)
            ShowCenteredOverlay_Utils(errMsg != "" ? "❌ " errMsg : "❌ Could not capture the link.", 2500,
                BANNER_ACCENT_ERROR)
        }
    } catch as e {
        ShowCenteredOverlay_Utils("❌ Error: " . e.Message, 3000, BANNER_ACCENT_ERROR)
    }
}

ShowStudyTopicSelector() {
    global g_StudyTopicSelectorGui, g_StudyTopicSelectorActive, g_StudyTopics, g_StudyTopicSelectorPhase,
        g_StudyTopicSelectorCategory, g_StudyTopicSelectorLastForegroundMonitorIdx

    if (g_StudyTopicSelectorActive)
        return

    g_StudyTopicSelectorCategory := ""
    g_StudyTopicSelectorPhase := "category"

    StudyTopicSelector_ShowCategoryPhase()

    StudyTopicSelector_StopActiveMonitorTracking()
    g_StudyTopicSelectorLastForegroundMonitorIdx := GetMonitorIndexForForeground_StandardBar()
    SetTimer(StudyTopicSelector_TrackActiveMonitorTick, 115)
}

StudyTopicSelector_ShowTopicPhase() {
    global g_StudyTopicSelectorGui, g_StudyTopicSelectorActive, g_StudyTopics, g_StudyTopicSelectorPhase,
        g_StudyTopicSelectorCategory, g_StudyTopicSelectorLastForegroundMonitorIdx

    if (!g_StudyTopicSelectorActive || g_StudyTopicSelectorPhase != "category")
        return

    StudyTopicSelector_UnbindCategoryHotkeys()
    try Hotkey("Backspace", StudyTopicSelector_Cancel, "Off")

    catLabel := (g_StudyTopicSelectorCategory = "plans") ? "Plans" : "Mnemonics"
    StudyTopicSelector_SafeDestroyGui(g_StudyTopicSelectorGui)
    g_StudyTopicSelectorGui := false

    g_StudyTopicSelectorGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Owner -DPIScale")
    g_StudyTopicSelectorGui.BackColor := "1E1E2E"
    g_StudyTopicSelectorGui.MarginX := 20
    g_StudyTopicSelectorGui.MarginY := 15

    g_StudyTopicSelectorGui.SetFont("s14 cCDD6F4 Bold", "Segoe UI")
    g_StudyTopicSelectorGui.Add("Text", "w300 Center", "📚 " . catLabel . " — topic")
    g_StudyTopicSelectorGui.Add("Text", "w300 h1 Background45475A")

    g_StudyTopicSelectorGui.SetFont("s12 cCDD6F4", "Segoe UI")
    for num, topic in g_StudyTopics {
        n := Integer(num)
        ; [0] Technique is category [6] only — omit from Mnemonics and Plans topic lists.
        if (n = 0 || (g_StudyTopicSelectorCategory = "mnemonics" && topic.mnemonicsPath = ""))
            continue
        g_StudyTopicSelectorGui.Add("Text", "w300", "[" . num . "] " . topic.name)
    }

    g_StudyTopicSelectorGui.Add("Text", "w300 h1 Background45475A y+10")
    g_StudyTopicSelectorGui.SetFont("s9 c6C7086", "Segoe UI")
    footerHint := (g_StudyTopicSelectorCategory = "mnemonics") ? "Press 1-6 | Backspace back | Esc cancel" :
        "Press 1-7 | Backspace back | Esc cancel"
    g_StudyTopicSelectorGui.Add("Text", "w300 Center", footerHint)

    try {
        g_StudyTopicSelectorGui.OnEvent("Escape", StudyTopicSelector_GuiEscape)
    } catch {
    }
    StudyTopicSelector_PositionGuiLikeOutlook(g_StudyTopicSelectorGui)
    g_StudyTopicSelectorLastForegroundMonitorIdx := GetMonitorIndexForForeground_StandardBar()

    g_StudyTopicSelectorPhase := "topic"
    if (g_StudyTopicSelectorCategory = "plans") {
        Hotkey("1", StudyTopicSelector_HandleKey, "On")
        Hotkey("2", StudyTopicSelector_HandleKey, "On")
        Hotkey("3", StudyTopicSelector_HandleKey, "On")
        Hotkey("4", StudyTopicSelector_HandleKey, "On")
        Hotkey("5", StudyTopicSelector_HandleKey, "On")
        Hotkey("6", StudyTopicSelector_HandleKey, "On")
        Hotkey("7", StudyTopicSelector_HandleKey, "On")
    } else {
        Hotkey("1", StudyTopicSelector_HandleKey, "On")
        Hotkey("2", StudyTopicSelector_HandleKey, "On")
        Hotkey("3", StudyTopicSelector_HandleKey, "On")
        Hotkey("4", StudyTopicSelector_HandleKey, "On")
        Hotkey("5", StudyTopicSelector_HandleKey, "On")
        Hotkey("6", StudyTopicSelector_HandleKey, "On")
    }
    Hotkey("Backspace", StudyTopicSelector_BackFromTopic, "On")
    StudyTopicSelector_BindRobustEscape()
}

StudyTopicSelector_BackFromTopic(*) {
    global g_StudyTopicSelectorGui, g_StudyTopicSelectorActive, g_StudyTopicSelectorPhase, g_StudyTopicSelectorCategory,
        g_StudyTopicSelectorLastForegroundMonitorIdx

    if (!g_StudyTopicSelectorActive || g_StudyTopicSelectorPhase != "topic")
        return
    StudyTopicSelector_UnbindDigitHotkeys()
    try Hotkey("Backspace", StudyTopicSelector_BackFromTopic, "Off")
    g_StudyTopicSelectorCategory := ""
    g_StudyTopicSelectorPhase := "category"
    StudyTopicSelector_SafeDestroyGui(g_StudyTopicSelectorGui)
    g_StudyTopicSelectorGui := false
    StudyTopicSelector_ShowCategoryPhase()
    g_StudyTopicSelectorLastForegroundMonitorIdx := GetMonitorIndexForForeground_StandardBar()
}

; [6] Technique: open technique README on GitHub in a new Chrome window (no scroll-to-end).
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

StudyTopicSelector_HandleKey(key) {
    global g_StudyTopicSelectorActive, g_StudyTopics, g_StudyTopicSelectorPhase, g_StudyTopicSelectorCategory

    if (!g_StudyTopicSelectorActive || g_StudyTopicSelectorPhase != "topic")
        return
    selection := Integer(key)
    category := g_StudyTopicSelectorCategory
    StudyTopicSelector_Close()
    if (!g_StudyTopics.Has(selection))
        return

    topic := g_StudyTopics[selection]
    url := StudyTopic_GetUrl(topic, category)
    StudyTopic_OpenGithubInChrome(url, category = "mnemonics")
}

StudyTopicSelector_Cancel(*) {
    StudyTopicSelector_Close()
}

; Tear-down order aligned with ShowAiModelSelector close (Utils\handy_selector_hotkey.ahk).
StudyTopicSelector_Close() {
    global g_StudyTopicSelectorGui, g_StudyTopicSelectorActive, g_StudyTopicSelectorPhase, g_StudyTopicSelectorCategory,
        g_StudyTopicSelectorLastForegroundMonitorIdx, g_StudyLinksGui, g_StudyArticleLinksGui, g_StudyFavoriteLinksGui,
        g_StudyLinkSubmenuGui

    if (!g_StudyTopicSelectorActive)
        return
    StudyTopicSelector_UnbindRobustEscape()
    g_StudyTopicSelectorActive := false
    g_StudyTopicSelectorPhase := ""
    g_StudyTopicSelectorCategory := ""

    StudyTopicSelector_StopActiveMonitorTracking()
    g_StudyTopicSelectorLastForegroundMonitorIdx := 0

    StudyTopicSelector_UnbindCategoryHotkeys()
    StudyTopicSelector_UnbindDigitHotkeys()
    try Hotkey("Backspace", "Off")
    catch {
    }
    Utils_EnsureGlobalEscapeHotkey()
    StudyTopicSelector_SafeDestroyGui(g_StudyTopicSelectorGui)
    g_StudyTopicSelectorGui := false
    StudyTopicSelector_SafeDestroyGui(g_StudyLinksGui)
    g_StudyLinksGui := false
    StudyTopicSelector_SafeDestroyGui(g_StudyArticleLinksGui)
    g_StudyArticleLinksGui := false
    StudyTopicSelector_SafeDestroyGui(g_StudyFavoriteLinksGui)
    g_StudyFavoriteLinksGui := false
    StudyTopicSelector_SafeDestroyGui(g_StudyLinkSubmenuGui)
    g_StudyLinkSubmenuGui := ""
    loop 9 {
        try Hotkey(String(A_Index), "Off")
    }
    try Hotkey("Escape", "Off")
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
