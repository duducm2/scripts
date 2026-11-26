; =============================================================================
; Utils module: pinned_quick_link.ahk
; One Macros-scoped pinned target (URL, file path, or other Run-able link).
; Persist: assets/data/pinned_quick_link.ini  [Pinned] Target=
; Macros: [4] open · select that row + Shift+E = set from clipboard
; =============================================================================

PinnedQuickLink_Path() {
    return A_ScriptDir "\assets\data\pinned_quick_link.ini"
}

global g_PinnedQuickLinkCacheReady := false
global g_PinnedQuickLinkCacheVal := ""
global g_PinnedQuickLinkCacheTick := 0

PinnedQuickLink_Get() {
    global g_PinnedQuickLinkCacheReady, g_PinnedQuickLinkCacheVal, g_PinnedQuickLinkCacheTick
    if (g_PinnedQuickLinkCacheReady && (A_TickCount - g_PinnedQuickLinkCacheTick) < 5000)
        return g_PinnedQuickLinkCacheVal
    path := PinnedQuickLink_Path()
    if !FileExist(path) {
        g_PinnedQuickLinkCacheReady := true
        g_PinnedQuickLinkCacheVal := ""
        g_PinnedQuickLinkCacheTick := A_TickCount
        return ""
    }
    try {
        v := IniRead(path, "Pinned", "Target", "")
        val := (v = "ERROR") ? "" : Trim(v)
        g_PinnedQuickLinkCacheReady := true
        g_PinnedQuickLinkCacheVal := val
        g_PinnedQuickLinkCacheTick := A_TickCount
        return val
    } catch {
        return ""
    }
}

PinnedQuickLink_Set(target) {
    global g_PinnedQuickLinkCacheReady, g_PinnedQuickLinkCacheVal, g_PinnedQuickLinkCacheTick
    path := PinnedQuickLink_Path()
    dir := A_ScriptDir "\assets\data"
    if !DirExist(dir) {
        try DirCreate(dir)
        catch {
            return false
        }
    }
    trimmed := Trim(target)
    try {
        IniWrite(trimmed, path, "Pinned", "Target")
        g_PinnedQuickLinkCacheReady := true
        g_PinnedQuickLinkCacheVal := trimmed
        g_PinnedQuickLinkCacheTick := A_TickCount
        return true
    } catch {
        return false
    }
}

PinnedQuickLink_Clear() {
    return PinnedQuickLink_Set("")
}

; Prefer http(s) URL in text; else trimmed text with surrounding quotes stripped.
PinnedQuickLink_NormalizeTarget(raw) {
    text := Trim(raw)
    if (text = "")
        return ""
    url := ""
    try url := StudyLink_ExtractUrlFromClipboardText(text)
    catch {
        url := ""
    }
    if (url != "")
        return url
    if (StrLen(text) >= 2) {
        first := SubStr(text, 1, 1)
        last := SubStr(text, -1)
        if ((first = '"' && last = '"') || (first = "'" && last = "'"))
            text := Trim(SubStr(text, 2, StrLen(text) - 2))
    }
    return text
}

PinnedQuickLink_IsHttpUrl(target) {
    try return StudyLink_IsValidHttpUrl(target)
    catch {
        u := Trim(target)
        return (SubStr(u, 1, 7) = "http://" || SubStr(u, 1, 8) = "https://")
    }
}

PinnedQuickLink_TruncateLabel(s, maxLen := 64) {
    t := Trim(s)
    if (t = "")
        return "(none)"
    if (StrLen(t) <= maxLen)
        return t
    return SubStr(t, 1, maxLen - 1) . "…"
}

PinnedQuickLink_MacroTitle() {
    return "📌 Open pinned link — " . PinnedQuickLink_TruncateLabel(PinnedQuickLink_Get())
}

PinnedQuickLink_Open(*) {
    target := PinnedQuickLink_Get()
    if (target = "") {
        ShowCenteredOverlay_Utils(
            "⚠ No pinned link. In Macros, select [4] and press Shift+E (clipboard).", 3000,
            BANNER_ACCENT_INTERMEDIATE)
        return false
    }
    if PinnedQuickLink_IsHttpUrl(target) {
        if !StudyLink_OpenUrlInChrome(target, true) {
            ShowCenteredOverlay_Utils("❌ Could not open Chrome.", 2000, BANNER_ACCENT_ERROR)
            return false
        }
        ShowCenteredOverlay_Utils("✅ Opening pinned link in Chrome…", 1500, BANNER_ACCENT_SUCCESS)
        return true
    }
    try {
        Run(target)
        ShowCenteredOverlay_Utils("✅ Opening pinned link…", 1500, BANNER_ACCENT_SUCCESS)
        return true
    } catch as err {
        ShowCenteredOverlay_Utils("❌ Could not open: " . err.Message, 3000, BANNER_ACCENT_ERROR)
        return false
    }
}

; Used by Utility Shortcuts Macros Shift+E on the pinned-link row.
PinnedQuickLink_SetFromClipboard(*) {
    clip := ""
    try clip := Trim(A_Clipboard)
    catch {
        clip := ""
    }
    fromClip := PinnedQuickLink_NormalizeTarget(clip)
    if (fromClip = "") {
        ShowCenteredOverlay_Utils("❌ Clipboard has no URL, path, or link to pin.", 3000,
            BANNER_ACCENT_ERROR)
        return false
    }
    if !PinnedQuickLink_Set(fromClip) {
        ShowCenteredOverlay_Utils("❌ Could not save pinned link.", 2500, BANNER_ACCENT_ERROR)
        return false
    }
    label := PinnedQuickLink_IsHttpUrl(fromClip) ? "URL" : "target"
    ShowCenteredOverlay_Utils("✅ Pinned " . label . " saved from clipboard.", 2500,
        BANNER_ACCENT_SUCCESS)
    return true
}
