; Module 5 — Manage Study Favorite Link (textual / web URLs, same API slot as MacroDroid Favorite)
; Included from Utils.ahk after StudyArticleLink.ahk (do not #include Helpers here).

global g_StudyFavoriteLinksGui := false

StudyTopicSelector_ManageFavoriteLinksEsc(*) {
    global g_StudyFavoriteLinksGui
    StudyLink_ClosePalaceManageGui(&g_StudyFavoriteLinksGui)
}

StudyTopicSelector_ManageFavoriteLinks(*) {
    global g_StudyFavoriteLinksGui
    StudyTopicSelector_UnbindRobustEscape()
    StudyTopicSelector_SafeDestroyGui(g_StudyFavoriteLinksGui)
    favResult := StudyLink_GetResult(STUDYLINK_KEY_FAVORITE)
    if (favResult["ok"])
        StudyLink_PlayApiSuccessSound()
    StudyLink_ShowPalaceManageGui(
        &g_StudyFavoriteLinksGui,
        "Memory Palace — Favorite",
        "❤️ Favorite",
        "Current: " . StudyLink_FormatLinkLabel(favResult),
        [
            ["1", "Open", "Open the stored favorite link in Chrome"],
            ["2", "Set from clipboard", "Save the http(s) URL on your clipboard"],
            ["3", "Set manually", "Paste or type a URL"]
        ],
        [
            ["1", StudyTopicSelector_ManageFavoriteLinks_Open],
            ["2", StudyTopicSelector_ManageFavoriteLinks_Set],
            ["3", StudyTopicSelector_ManageFavoriteLinks_SetManual]
        ],
        StudyTopicSelector_ManageFavoriteLinksEsc
    )
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
    StudyLink_SetFromClipboard(STUDYLINK_KEY_FAVORITE, "favorite link")
}

; [3] Set favorite link manually (InputBox → StudyLink_Set API)
StudyTopicSelector_ManageFavoriteLinks_SetManual(*) {
    StudyTopicSelector_Close()
    StudyLink_SetFromManualInput(STUDYLINK_KEY_FAVORITE, "favorite link")
}
