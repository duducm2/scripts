; =============================================================================
; Utils module: pinned_quick_link.ahk
; One Macros-scoped pinned target (URL, file path, or other Run-able link).
; Persist: assets/data/pinned_quick_link.ini  [Pinned] Target=
; Macros: [o] open  [4] set / clear
; =============================================================================

PinnedQuickLink_Path() {
    return A_ScriptDir "\assets\data\pinned_quick_link.ini"
}

PinnedQuickLink_Get() {
    path := PinnedQuickLink_Path()
    if !FileExist(path)
        return ""
    try {
        v := IniRead(path, "Pinned", "Target", "")
        if (v = "ERROR")
            return ""
        return Trim(v)
    } catch {
        return ""
    }
}

PinnedQuickLink_Set(target) {
    path := PinnedQuickLink_Path()
    dir := A_ScriptDir "\assets\data"
    if !DirExist(dir) {
        try DirCreate(dir)
        catch {
            return false
        }
    }
    try {
        IniWrite(Trim(target), path, "Pinned", "Target")
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

PinnedQuickLink_Open(*) {
    target := PinnedQuickLink_Get()
    if (target = "") {
        ShowCenteredOverlay_Utils("⚠ No pinned link. Set with Macros [4] first.", 2500,
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

PinnedQuickLink_SetFromClipboardOrInput(*) {
    clip := ""
    try clip := Trim(A_Clipboard)
    catch {
        clip := ""
    }
    fromClip := PinnedQuickLink_NormalizeTarget(clip)
    if (fromClip != "") {
        if PinnedQuickLink_Set(fromClip) {
            label := PinnedQuickLink_IsHttpUrl(fromClip) ? "URL" : "target"
            ShowCenteredOverlay_Utils("✅ Pinned " . label . " saved from clipboard.", 2500,
                BANNER_ACCENT_SUCCESS)
            return true
        }
        ShowCenteredOverlay_Utils("❌ Could not save pinned link.", 2500, BANNER_ACCENT_ERROR)
        return false
    }

    current := PinnedQuickLink_Get()
    ib := InputBox(
        "Paste a URL, file path, or other link.`nLeave empty and OK to clear.",
        "Set pinned quick link",
        "w520 h140",
        current
    )
    if (ib.Result != "OK") {
        ShowCenteredOverlay_Utils("⚠ Set pinned link cancelled.", 1800, BANNER_ACCENT_INTERMEDIATE)
        return false
    }
    target := PinnedQuickLink_NormalizeTarget(ib.Value)
    if (target = "") {
        if (current = "") {
            ShowCenteredOverlay_Utils("⚠ Nothing to clear.", 1800, BANNER_ACCENT_INTERMEDIATE)
            return false
        }
        if !PinnedQuickLink_Clear() {
            ShowCenteredOverlay_Utils("❌ Could not clear pinned link.", 2500, BANNER_ACCENT_ERROR)
            return false
        }
        ShowCenteredOverlay_Utils("✅ Pinned link cleared.", 2000, BANNER_ACCENT_SUCCESS)
        return true
    }
    if PinnedQuickLink_Set(target) {
        ShowCenteredOverlay_Utils("✅ Pinned link saved.", 2000, BANNER_ACCENT_SUCCESS)
        return true
    }
    ShowCenteredOverlay_Utils("❌ Could not save pinned link.", 2500, BANNER_ACCENT_ERROR)
    return false
}
