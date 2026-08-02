; Module 5 — Manage Study Favorite Link (textual / web URLs, same API slot as MacroDroid Favorite)
; Included from Utils.ahk after StudyArticleLink.ahk (do not #include Helpers here).

global g_StudyFavoriteLinksGui := false

StudyTopicSelector_ManageFavoriteLinksEsc(*) {
    global g_StudyFavoriteLinksGui
    StudyTopicSelector_SafeDestroyGui(g_StudyFavoriteLinksGui)
    g_StudyFavoriteLinksGui := false
    StudyTopicSelector_ResumeSelectorEscapeAfterLinks()
}

StudyTopicSelector_ManageFavoriteLinks(*) {
    global g_StudyFavoriteLinksGui
    StudyTopicSelector_UnbindRobustEscape()
    StudyTopicSelector_SafeDestroyGui(g_StudyFavoriteLinksGui)
    g_StudyFavoriteLinksGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Owner -DPIScale")
    loop 9 {
        try Hotkey(String(A_Index), "Off")
    }
    try Hotkey("Escape", "Off")
    g_StudyFavoriteLinksGui.BackColor := "1E1E2E"
    g_StudyFavoriteLinksGui.MarginX := 20
    g_StudyFavoriteLinksGui.MarginY := 15
    g_StudyFavoriteLinksGui.SetFont("s14 cCDD6F4 Bold", "Segoe UI")
    g_StudyFavoriteLinksGui.Add("Text", "w400 Center", "⭐ Manage Study Favorite Link")
    g_StudyFavoriteLinksGui.Add("Text", "w400 h1 Background45475A")
    g_StudyFavoriteLinksGui.SetFont("s11 cCDD6F4", "Segoe UI")
    favResult := StudyLink_GetResult(STUDYLINK_KEY_FAVORITE)
    if (favResult["ok"])
        StudyLink_PlayApiSuccessSound()
    g_StudyFavoriteLinksGui.Add("Text", "w400", "Current favorite link: " . StudyLink_FormatLinkLabel(favResult))
    g_StudyFavoriteLinksGui.Add("Text", "w400", "[1] Open favorite link")
    g_StudyFavoriteLinksGui.Add("Text", "w400", "[2] Set favorite link (Chrome address bar)")
    g_StudyFavoriteLinksGui.Add("Text", "w400", "[3] Set favorite link manually")
    g_StudyFavoriteLinksGui.Add("Text", "w400 h1 Background45475A y+10")
    g_StudyFavoriteLinksGui.SetFont("s9 c6C7086", "Segoe UI")
    g_StudyFavoriteLinksGui.Add("Text", "w400 Center", "Press 1-3 | Esc to cancel")
    StudyTopicSelector_PositionGuiLikeOutlook(g_StudyFavoriteLinksGui)
    Hotkey("1", StudyTopicSelector_ManageFavoriteLinks_Open, "On")
    Hotkey("2", StudyTopicSelector_ManageFavoriteLinks_Set, "On")
    Hotkey("3", StudyTopicSelector_ManageFavoriteLinks_SetManual, "On")
    Hotkey("Escape", StudyTopicSelector_ManageFavoriteLinksEsc, "On")
    g_StudyFavoriteLinksGui.Show()
}

StudyTopicSelector_ManageFavoriteLinks_Open(*) {
    StudyTopicSelector_Close()
    linkResult := StudyLink_GetResult(STUDYLINK_KEY_FAVORITE)
    if (!linkResult["ok"]) {
        ShowCenteredOverlay_Utils("❌ Could not load link from API: " . linkResult["err"], 3500, BANNER_ACCENT_ERROR)
        return
    }
    StudyLink_PlayApiSuccessSound()
    url := linkResult["url"]
    if (url != "") {
        StudyLink_OpenUrlInChrome(url, true)
        ShowCenteredOverlay_Utils("✅ Opening favorite link in a new Chrome window...", 2000, BANNER_ACCENT_SUCCESS)
    } else {
        ShowCenteredOverlay_Utils("⚠ No favorite link stored. Use [2] Set favorite link first.", 2500,
            BANNER_ACCENT_INTERMEDIATE)
    }
}

StudyTopicSelector_ManageFavoriteLinks_Set(*) {
    StudyTopicSelector_Close()
    try {
        errMsg := ""
        url := StudyArticleLink_CaptureChromeUrlFromAddressBar(&errMsg)
        if (url != "") {
            StandardLoadingBar_Show("Saving favorite link…", BANNER_ACCENT_INTERMEDIATE, { passive: false })
            setOk := StudyLink_Set(STUDYLINK_KEY_FAVORITE, url)
            try StandardLoadingBar_Hide(0)
            if setOk
                ShowCenteredOverlay_Utils("✅ Favorite link saved to study notes.", 3000, BANNER_ACCENT_SUCCESS)
            else
                ShowCenteredOverlay_Utils("❌ Could not save the favorite link (API failed).", 3500, BANNER_ACCENT_ERROR
                )
        } else {
            ShowCenteredOverlay_Utils(errMsg != "" ? "❌ " errMsg : "❌ Could not capture the URL.", 2500,
                BANNER_ACCENT_ERROR)
        }
    } catch as e {
        ShowCenteredOverlay_Utils("❌ Error: " . e.Message, 3000, BANNER_ACCENT_ERROR)
    }
}

; [3] Set favorite link manually (InputBox → StudyLink_Set API)
StudyTopicSelector_ManageFavoriteLinks_SetManual(*) {
    StudyTopicSelector_Close()
    StudyLink_SetFromManualInput(STUDYLINK_KEY_FAVORITE, "favorite link")
}
