; Module 4 — Manage Study Article Link (textual / web article URLs)
; Included from Utils.ahk after StudyLinkHelpers.ahk (do not #include Helpers here).

global g_StudyArticleLinksGui := false

StudyArticleLink_IsValidHttpUrl(url) {
    u := Trim(url)
    if (u = "")
        return false
    if (SubStr(u, 1, 11) = "javascript:")
        return false
    return (SubStr(u, 1, 7) = "http://" || SubStr(u, 1, 8) = "https://")
}

; Chrome address bar: F6 focus, copy URL to clipboard (keyboard only).
StudyArticleLink_CaptureChromeUrlFromAddressBar(&errMsg := "") {
    errMsg := ""
    chromeHwnd := WinExist("ahk_class Chrome_WidgetWin_1")
    if !chromeHwnd {
        errMsg := "Chrome window not found. Switch to Chrome and try again."
        return ""
    }
    savedClipAll := ClipboardAll()
    savedClipText := A_Clipboard
    try {
        WinActivate("ahk_id " chromeHwnd)
        if !WinWaitActive("ahk_id " chromeHwnd, , 2) {
            errMsg := "Chrome window did not become active."
            return ""
        }
        Sleep 280
        Send "{F6}"
        Sleep 120
        Send "^c"
        clipDeadline := A_TickCount + 2000
        url := ""
        while (A_TickCount < clipDeadline) {
            url := Trim(A_Clipboard)
            if (url != "" && url != Trim(savedClipText) && StudyArticleLink_IsValidHttpUrl(url))
                break
            Sleep 80
        }
        if !StudyArticleLink_IsValidHttpUrl(url) {
            errMsg := "Could not copy a valid URL from the address bar."
            return ""
        }
        return url
    } finally {
        try A_Clipboard := savedClipAll
        catch {
        }
    }
}

StudyTopicSelector_ManageArticleLinksEsc(*) {
    global g_StudyArticleLinksGui
    StudyTopicSelector_SafeDestroyGui(g_StudyArticleLinksGui)
    g_StudyArticleLinksGui := false
    StudyTopicSelector_ResumeSelectorEscapeAfterLinks()
}

StudyTopicSelector_ManageArticleLinks(*) {
    global g_StudyArticleLinksGui
    StudyTopicSelector_UnbindRobustEscape()
    StudyTopicSelector_SafeDestroyGui(g_StudyArticleLinksGui)
    g_StudyArticleLinksGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Owner -DPIScale")
    loop 9 {
        try Hotkey(String(A_Index), "Off")
    }
    try Hotkey("Escape", "Off")
    g_StudyArticleLinksGui.BackColor := "1E1E2E"
    g_StudyArticleLinksGui.MarginX := 20
    g_StudyArticleLinksGui.MarginY := 15
    g_StudyArticleLinksGui.SetFont("s14 cCDD6F4 Bold", "Segoe UI")
    g_StudyArticleLinksGui.Add("Text", "w400 Center", "📄 Manage Study Article Link")
    g_StudyArticleLinksGui.Add("Text", "w400 h1 Background45475A")
    g_StudyArticleLinksGui.SetFont("s11 cCDD6F4", "Segoe UI")
    artResult := StudyLink_GetResult(STUDYLINK_KEY_ARTICLE)
    g_StudyArticleLinksGui.Add("Text", "w400", "Current article link: " . StudyLink_FormatLinkLabel(artResult))
    g_StudyArticleLinksGui.Add("Text", "w400", "[1] Open article link")
    g_StudyArticleLinksGui.Add("Text", "w400", "[2] Set article link (Chrome address bar)")
    g_StudyArticleLinksGui.Add("Text", "w400 h1 Background45475A y+10")
    g_StudyArticleLinksGui.SetFont("s9 c6C7086", "Segoe UI")
    g_StudyArticleLinksGui.Add("Text", "w400 Center", "Press 1-2 | Esc to cancel")
    StudyTopicSelector_PositionGuiLikeOutlook(g_StudyArticleLinksGui)
    Hotkey("1", StudyTopicSelector_ManageArticleLinks_Open, "On")
    Hotkey("2", StudyTopicSelector_ManageArticleLinks_Set, "On")
    Hotkey("Escape", StudyTopicSelector_ManageArticleLinksEsc, "On")
    g_StudyArticleLinksGui.Show()
}

StudyTopicSelector_ManageArticleLinks_Open(*) {
    StudyTopicSelector_Close()
    linkResult := StudyLink_GetResult(STUDYLINK_KEY_ARTICLE)
    if (!linkResult["ok"]) {
        ShowCenteredOverlay_Utils("❌ Could not load link from API: " . linkResult["err"], 3500, BANNER_ACCENT_ERROR)
        return
    }
    url := linkResult["url"]
    if (url != "") {
        StudyLink_OpenUrlInChrome(url)
        ShowCenteredOverlay_Utils("✅ Opening article link in Chrome...", 2000, BANNER_ACCENT_SUCCESS)
    } else {
        ShowCenteredOverlay_Utils("⚠ No article link stored. Use [2] Set article link first.", 2500,
            BANNER_ACCENT_INTERMEDIATE)
    }
}

StudyTopicSelector_ManageArticleLinks_Set(*) {
    StudyTopicSelector_Close()
    try {
        errMsg := ""
        url := StudyArticleLink_CaptureChromeUrlFromAddressBar(&errMsg)
        if (url != "") {
            StandardLoadingBar_Show("Saving article link…", BANNER_ACCENT_INTERMEDIATE, { passive: false })
            setOk := StudyLink_Set(STUDYLINK_KEY_ARTICLE, url)
            try StandardLoadingBar_Hide(0)
            if setOk
                ShowCenteredOverlay_Utils("✅ Article link saved to study notes.", 3000, BANNER_ACCENT_SUCCESS)
            else
                ShowCenteredOverlay_Utils("❌ Could not save the article link (API failed).", 3500, BANNER_ACCENT_ERROR)
        } else {
            ShowCenteredOverlay_Utils(errMsg != "" ? "❌ " errMsg : "❌ Could not capture the URL.", 2500,
                BANNER_ACCENT_ERROR)
        }
    } catch as e {
        ShowCenteredOverlay_Utils("❌ Error: " . e.Message, 3000, BANNER_ACCENT_ERROR)
    }
}
