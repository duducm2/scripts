; =============================================================================
; Utils module: language_flag_indicator.ahk
; Language flag indicator and status banners
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Status Banner Functions (non-blocking; use standard loading indicator)
; =============================================================================
AiModelBanner_Show(text, bgColor := BANNER_ACCENT_INTERMEDIATE) {
    StandardLoadingBar_Show(text, bgColor, { passive: true, centerOnHwnd: 0, textWidth: 450, fontSize: 17,
        passiveBgColor: bgColor, alpha: 200 })
}

AiModelBanner_Hide() {
    StandardLoadingBar_Hide(0)
}

; =============================================================================
; Persistent Language Flag Indicator (slot 1 = UK, slot 2 = Brazil, slot 3 = multi)
; =============================================================================
; Opaque, always-on-top flag chips pinned to the bottom-right of every monitor.
; Slot 1 = English (UK), slot 2 = Portuguese (Brazil), slot 3 = multi-language (combined flag).
; Use the visible flag as the spoken-language indicator across all screens.
; =============================================================================

LanguageFlag_GetImagePath(slot) {
    rel := (slot = 1) ? "\assets\images\flags\united-kingdom.png"
        : (slot = 2) ? "\assets\images\flags\brazil.png"
            : (slot = 3) ? "\assets\images\flags\brazil-united-kingdom.png"
                : ""
    if (rel = "")
        return ""
    ; Prefer the running script's folder, then the folder that contains Utils.ahk (covers odd layouts).
    candidates := [A_ScriptDir . rel]
    SplitPath(A_LineFile, , &utilsDir)
    if (utilsDir != "" && utilsDir != A_ScriptDir)
        candidates.Push(utilsDir . rel)
    for p in candidates {
        if FileExist(p)
            return p
    }
    return ""
}

LanguageFlag_CreateGui(slot, imagePath) {
    global LANGUAGE_FLAG_WIDTH

    ; Borderless, zero-margin window so the GUI sizes exactly to the bitmap.
    ; Do NOT use WS_EX_TRANSPARENT (part of +E0x80020): that style suppresses
    ; window painting and the flag can disappear entirely.
    flagGui := Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale")
    flagGui.BackColor := "313244"
    flagGui.MarginX := 0
    flagGui.MarginY := 0

    usedPicture := false
    if (imagePath != "") {
        try {
            flagGui.Add("Picture", "w" . LANGUAGE_FLAG_WIDTH . " h-1", imagePath)
            usedPicture := true
        } catch {
            usedPicture := false
        }
    }
    if !usedPicture {
        flagGui.SetFont("s13 cFFFFFF Bold", "Segoe UI")
        label := (slot = 1) ? "EN" : (slot = 2) ? "PT" : (slot = 3) ? "EN+PT" : "?"
        flagGui.Add("Text", "Center w" . LANGUAGE_FLAG_WIDTH . " h31 Background45475A", label)
    }
    flagGui.Show("AutoSize Hide")
    return flagGui
}

LanguageFlag_Show(slot) {
    global g_LanguageFlagGuis, g_LanguageFlagSlot

    if (slot < 1 || slot > 3) {
        LanguageFlag_Hide()
        return
    }

    LanguageFlag_Hide()
    g_LanguageFlagSlot := slot

    imagePath := LanguageFlag_GetImagePath(slot)
    monitorCount := MonitorGetCount()
    if (monitorCount < 1)
        return

    loop monitorCount {
        idx := A_Index
        flagGui := LanguageFlag_CreateGui(slot, imagePath)
        g_LanguageFlagGuis.Push({ monitor: idx, gui: flagGui })
    }

    LanguageFlag_RepositionAllMonitors()
}

LanguageFlag_Hide() {
    global g_LanguageFlagGuis, g_LanguageFlagSlot
    for item in g_LanguageFlagGuis {
        try {
            if IsObject(item.gui)
                item.gui.Destroy()
        } catch {
        }
    }
    g_LanguageFlagGuis := []
    g_LanguageFlagSlot := 0
}

LanguageFlag_RepositionAllMonitors() {
    global g_LanguageFlagGuis, LANGUAGE_FLAG_MARGIN
    if (!g_LanguageFlagGuis.Length)
        return

    for item in g_LanguageFlagGuis {
        monitorIdx := item.monitor
        flagGui := item.gui
        if !IsObject(flagGui)
            continue

        try {
            MonitorGetWorkArea(monitorIdx, &ml, &mt, &mr, &mb)
            flagGui.GetPos(, , &gw, &gh)
        } catch {
            continue
        }

        guiX := mr - gw - LANGUAGE_FLAG_MARGIN
        guiY := mb - gh - LANGUAGE_FLAG_MARGIN
        if (guiX < ml)
            guiX := ml
        if (guiY < mt)
            guiY := mt

        try {
            flagGui.Move(guiX, guiY)
            hwnd := flagGui.Hwnd
            if (hwnd) {
                ; SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE = 0x0001 | 0x0004 | 0x0010 = 0x0015
                DllCall("SetWindowPos", "Ptr", hwnd, "Ptr", 0, "Int", guiX, "Int", guiY, "Int", 0, "Int", 0,
                    "UInt", 0x0015)
            }
            flagGui.Show("NA")
        } catch {
        }
    }
}

LanguageFlag_InitFromPersistedSlot() {
    slot := Handy_GetPersistedAiModelSlot()
    if (slot >= 1 && slot <= 3)
        LanguageFlag_Show(slot)
}

; Small banner for Clip Angel (uses standard loading indicator).
ClipAngelBanner_Show(text, bgColor := BANNER_ACCENT_INTERMEDIATE) {
    StandardLoadingBar_Show(text, bgColor, { passive: true, centerOnHwnd: 0, textWidth: 200, fontSize: 17,
        passiveBgColor: bgColor, alpha: 220 })
}

ClipAngelBanner_Hide() {
    StandardLoadingBar_Hide(0)
}

; Fast Copy Mode (Shift keys - Clip Angel sequential paste): persistent banner with live copy count.
FastCopyModeBanner_Show() {
    StandardLoadingBar_Show("📋 Fast Copy Mode - copies: 0", BANNER_ACCENT_INFO, { passive: true, centerOnHwnd: 0,
        textWidth: 480, fontSize: 17, passiveBgColor: BANNER_ACCENT_INFO, alpha: 220,
        promptKeys: "[Win+Alt+Shift+J] Finish and paste", trackActiveMonitor: true })
}

FastCopyModeBanner_Update(copyCount) {
    StandardLoadingBar_Update("📋 Fast Copy Mode - copies: " copyCount)
}

FastCopyModeBanner_Hide() {
    StandardLoadingBar_Hide(0)
}

; =============================================================================
; Single-character tab banner (uses standard loading indicator). tabNumber 1 = blue, 2 = yellow. Auto-hides after 700 ms.
; =============================================================================
ShowSingleCharTabBanner_Utils(tabNumber) {
    msg := String(tabNumber)
    bgColor := (tabNumber = 1) ? "0000FF" : "FFFF00"
    StandardLoadingBar_Show(msg, bgColor, { passive: true, centerOnHwnd: 0, textWidth: 120, fontSize: 72,
        passiveBgColor: bgColor, alpha: 178 })
    StandardLoadingBar_Hide(700)
}

; =============================================================================
; ExecuteHandyAiModelSelection() - Main automation logic for Handy
; =============================================================================
ExecuteHandyAiModelSelection(selection) {
    global g_HandyAiModels

    modelInfo := g_HandyAiModels[selection]
    modelDisplayName := modelInfo.name
    modelClickName := modelInfo.HasProp("modelClickName") ? modelInfo.modelClickName : modelInfo.name

    try {
        ; Step 1: Launch/activate Handy
        AiModelBanner_Show("🚀 Launching Handy...")
        handyHwnd := Handy_ActivateOrLaunch()
        if (!handyHwnd) {
            AiModelBanner_Show("❌ Failed to launch Handy", "E74C3C")
            Sleep 2000
            AiModelBanner_Hide()
            return
        }

        ; Step 2: Open AI model menu
        AiModelBanner_Show("📋 Opening AI model menu...")
        if (!Handy_OpenAiModelMenu(handyHwnd)) {
            AiModelBanner_Show("❌ Could not open AI menu", "E74C3C")
            Sleep 2000
            AiModelBanner_Hide()
            return
        }

        ; Step 3: Select the model
        AiModelBanner_Show("🎯 Selecting " . modelDisplayName . "...")
        if (!Handy_ClickAiModel(handyHwnd, modelClickName)) {
            AiModelBanner_Show("❌ Model not found: " . modelClickName, "E74C3C")
            Sleep 2000
            AiModelBanner_Hide()
            return
        }
        Handy_SetPersistedAiModelSlot(selection)

        ; Update persistent language flag indicator (slot 1 = UK, slot 2 = BR, slot 3 = multi).
        if (selection >= 1 && selection <= 3)
            LanguageFlag_Show(selection)
        else
            LanguageFlag_Hide()

        ; Step 4: Wait for model to finish loading (poll button name until "loading" disappears)
        AiModelBanner_Show("⏳ Waiting for model...", BANNER_ACCENT_INTERMEDIATE)
        Handy_WaitForModelReady(handyHwnd, 20000)

        ; Step 4.5: Play confirmation sound when model is ready
        soundPath := A_ScriptDir . "\assets\sounds\handy-model-chosen.mp3"
        if (FileExist(soundPath))
            ScriptSoundPlay(soundPath)

        ; Step 5: Close Handy window
        AiModelBanner_Show("✅ Done! Closing Handy...", BANNER_ACCENT_SUCCESS)
        try WinClose("ahk_id " . handyHwnd)
        Sleep 150

        AiModelBanner_Hide()

    } catch Error as e {
        AiModelBanner_Show("❌ Error: " . e.Message, "E74C3C")
        Sleep 2000
        AiModelBanner_Hide()
    }
}
