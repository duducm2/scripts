; Module 4 — Manage Study Article Link (textual / web article URLs)
; Included from Utils.ahk after StudyLinkHelpers.ahk (do not #include Helpers here).

global g_StudyArticleLinksGui := false

StudyArticleLink_IsValidHttpUrl(url) {
    return StudyLink_IsValidHttpUrl(url)
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
            ["2", "Set from Chrome", "Copy URL from the address bar (F6)"],
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

; [3] Set article link manually (InputBox → StudyLink_Set API)
StudyTopicSelector_ManageArticleLinks_SetManual(*) {
    StudyTopicSelector_Close()
    StudyLink_SetFromManualInput(STUDYLINK_KEY_ARTICLE, "article link")
}
