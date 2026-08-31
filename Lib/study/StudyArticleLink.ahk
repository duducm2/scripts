; Module 4 — Manage Study Article Link (textual / web article URLs)
; Included from Utils.ahk after StudyLinkHelpers.ahk (do not #include Helpers here).

global g_StudyArticleLinksGui := false

StudyTopicSelector_ManageArticleLinksEsc(*) {
    global g_StudyArticleLinksGui
    StudyLink_ClosePalaceManageGui(&g_StudyArticleLinksGui)
}

StudyTopicSelector_ManageArticleLinks(*) {
    global g_StudyArticleLinksGui
    StudyTopicSelector_UnbindRobustEscape()
    StudyTopicSelector_SafeDestroyGui(g_StudyArticleLinksGui)
    artResult := StudyLink_GetResult(STUDYLINK_KEY_ARTICLE)
    if (artResult["ok"])
        StudyLink_PlayApiSuccessSound()
    StudyLink_ShowPalaceManageGui(
        &g_StudyArticleLinksGui,
        "Memory Palace — Study Article",
        "📖 Study Article",
        "Current: " . StudyLink_FormatLinkLabel(artResult),
        [
            ["1", "Open", "Open the stored article link in Chrome"],
            ["2", "Set from clipboard", "Save the http(s) URL on your clipboard"],
            ["3", "Set manually", "Paste or type a URL"]
        ],
        [
            ["1", StudyTopicSelector_ManageArticleLinks_Open],
            ["2", StudyTopicSelector_ManageArticleLinks_Set],
            ["3", StudyTopicSelector_ManageArticleLinks_SetManual]
        ],
        StudyTopicSelector_ManageArticleLinksEsc
    )
}

StudyTopicSelector_ManageArticleLinks_Open(*) {
    StudyTopicSelector_Close()
    linkResult := StudyLink_GetResult(STUDYLINK_KEY_ARTICLE)
    if (!linkResult["ok"]) {
        ShowCenteredOverlay_Utils("❌ Could not load link from API: " . linkResult["err"], 3500, BANNER_ACCENT_ERROR)
        return
    }
    StudyLink_PlayApiSuccessSound()
    url := linkResult["url"]
    if (url != "") {
        StudyLink_OpenUrlInChrome(url, true)
        ShowCenteredOverlay_Utils("✅ Opening article link in a new Chrome window...", 2000, BANNER_ACCENT_SUCCESS)
    } else {
        ShowCenteredOverlay_Utils("⚠ No article link stored. Use [2] Set article link first.", 2500,
            BANNER_ACCENT_INTERMEDIATE)
    }
}

StudyTopicSelector_ManageArticleLinks_Set(*) {
    StudyTopicSelector_Close()
    StudyLink_SetFromClipboard(STUDYLINK_KEY_ARTICLE, "article link")
}

; [3] Set article link manually (InputBox → StudyLink_Set API)
StudyTopicSelector_ManageArticleLinks_SetManual(*) {
    StudyTopicSelector_Close()
    StudyLink_SetFromManualInput(STUDYLINK_KEY_ARTICLE, "article link")
}
