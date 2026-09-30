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
; Language flag images (slot 1 = UK, slot 2 = Brazil, slot 3 = multi)
; =============================================================================
; The spoken-language flag is shown only while dictating (DictationFlag_*).
; Do not pin a chip to the bottom-right of every monitor between takes.
; LanguageFlag_GetImagePath is shared with the recording flag.
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
; keepOpen: when true, leave Handy visible after success (for History re-transcribe).
; restoreHwnd: window to re-activate after automation (paste/dictation target).
;   0 = capture foreground at start (before any focus steal from fallback path).
; Returns true on success, false on failure.
; Background path: Handy stays off-screen/transparent (BeginSuppress); no WinClose.
; =============================================================================
ExecuteHandyAiModelSelection(selection, keepOpen := false, restoreHwnd := 0, restartDictationIfStopped := false) {
    global g_HandyAiModels, HANDY_AI_MODEL_MAX_ATTEMPTS, HANDY_AI_MODEL_RETRY_DELAY_MS
    global g_HandyModelSwitchBusy, g_HandySuppressActive

    if !g_HandyAiModels.Has(selection)
        return false

    if (g_HandyModelSwitchBusy)
        return false
    g_HandyModelSwitchBusy := true

    if (!restoreHwnd) {
        try restoreHwnd := WinGetID("A")
        catch {
            restoreHwnd := 0
        }
    }

    modelInfo := g_HandyAiModels[selection]
    modelDisplayName := modelInfo.name
    modelClickName := modelInfo.HasProp("modelClickName") ? modelInfo.modelClickName : modelInfo.name

    ; Accidental English <-> multi-lang while recording: stop before the model UI runs.
    stoppedDictation := Handy_StopDictationIfEnglishMultilangSwitch(selection)

    handyHwnd := 0
    try {
        verified := false

        loop HANDY_AI_MODEL_MAX_ATTEMPTS {
            attempt := A_Index
            attemptLabel := (attempt = 1)
                ? "Attempt " . attempt . "/" . HANDY_AI_MODEL_MAX_ATTEMPTS . ": Preparing Handy..."
                : "Retry " . attempt . "/" . HANDY_AI_MODEL_MAX_ATTEMPTS . ": Preparing Handy..."
            AiModelBanner_Show(attemptLabel)
            handyHwnd := Handy_EnsureForAutomation(attempt = 1 ? 2000 : 3000)
            if (!handyHwnd) {
                if (attempt < HANDY_AI_MODEL_MAX_ATTEMPTS) {
                    Sleep HANDY_AI_MODEL_RETRY_DELAY_MS
                    continue
                }
                AiModelBanner_Show("❌ Failed to launch Handy", "E74C3C")
                Sleep 2000
                AiModelBanner_Hide()
                Handy_RestorePrevWindow(restoreHwnd)
                return false
            }

            ; Fast path: already on the requested model (warm suppressed instance).
            if (Handy_VerifyAiModelActive(handyHwnd, modelClickName)) {
                verified := true
                break
            }

            AiModelBanner_Show((attempt > 1 ? "Retry " . attempt . "/" . HANDY_AI_MODEL_MAX_ATTEMPTS . ": " : "")
            . "Selecting " . modelDisplayName . "...")
            if (!Handy_TrySelectAiModel(handyHwnd, modelClickName)) {
                Handy_DismissOpenUi(handyHwnd)
                if (attempt < HANDY_AI_MODEL_MAX_ATTEMPTS) {
                    Sleep HANDY_AI_MODEL_RETRY_DELAY_MS
                    continue
                }
                break
            }

            AiModelBanner_Show("Verifying " . modelDisplayName . "...", BANNER_ACCENT_INTERMEDIATE)
            if (Handy_VerifyAiModelActive(handyHwnd, modelClickName)) {
                verified := true
                break
            }

            Handy_DismissOpenUi(handyHwnd)
            if (attempt < HANDY_AI_MODEL_MAX_ATTEMPTS)
                Sleep HANDY_AI_MODEL_RETRY_DELAY_MS
        }

        if (!verified) {
            AiModelBanner_Show("❌ Could not switch model after " . HANDY_AI_MODEL_MAX_ATTEMPTS . " attempts", "E74C3C")
            Sleep 2000
            AiModelBanner_Hide()
            Handy_RestorePrevWindow(restoreHwnd, handyHwnd)
            return false
        }

        if (!Handy_SetPersistedAiModelSlot(selection)) {
            AiModelBanner_Show("❌ Could not save model preference", BANNER_ACCENT_ERROR)
            Sleep 2000
            AiModelBanner_Hide()
            Handy_RestorePrevWindow(restoreHwnd, handyHwnd)
            return false
        }

        ; Slot is already in the INI. The language flag appears only while Recording is up.

        soundPath := A_ScriptDir . "\assets\sounds\handy-model-chosen.mp3"
        if (FileExist(soundPath))
            ScriptSoundPlay(soundPath)

        if (keepOpen) {
            AiModelBanner_Show("✅ Model ready", BANNER_ACCENT_SUCCESS)
            Sleep 400
            AiModelBanner_Hide()
            ; Interactive follow-up (re-transcribe): end suppress and show Handy.
            if (handyHwnd)
                Handy_EndSuppress(handyHwnd, true)
            try WinActivate("ahk_id " . handyHwnd)
            catch {
            }
            return true
        }

        ; Keep Handy warm and suppressed off-screen — do not WinClose.
        if (handyHwnd)
            Handy_BeginSuppress(handyHwnd)

        AiModelBanner_Show("✅ " . modelDisplayName, BANNER_ACCENT_SUCCESS)
        Sleep 350
        AiModelBanner_Hide()
        Handy_RestorePrevWindow(restoreHwnd, handyHwnd)
        ; Utility Shortcuts K/L during a take: model is switched; start dictation again.
        if (restartDictationIfStopped && stoppedDictation)
            Handy_RestartDictationAfterLanguageSwitch()
        return true

    } catch Error as e {
        AiModelBanner_Show("❌ Error: " . e.Message, "E74C3C")
        Sleep 2000
        AiModelBanner_Hide()
        if (!keepOpen)
            Handy_RestorePrevWindow(restoreHwnd, handyHwnd)
        return false
    } finally {
        g_HandyModelSwitchBusy := false
    }
}

; Utility Shortcuts (# !+U) is registered by every script that includes Utils.
; After reboot the last script to start owns the menu. Model UIA runs in
; AppLaunchers when that script is up, so other hosts hand the slot across.
HANDY_AI_MODEL_REQUEST_MSG_NAME := "EDU_HandyAi_SelectModel"

Handy_FindAppLaunchersHwnd() {
    prevDetect := A_DetectHiddenWindows
    prevMatch := A_TitleMatchMode
    DetectHiddenWindows true
    SetTitleMatchMode 2
    hwnd := 0
    try hwnd := WinExist("AppLaunchers.ahk ahk_class AutoHotkey")
    catch
        hwnd := 0
    if (!hwnd) {
        for exe in ["AutoHotkey64.exe", "AutoHotkey32.exe", "AutoHotkey.exe"] {
            try list := WinGetList("ahk_exe " exe)
            catch
                continue
            for candidate in list {
                title := ""
                try title := WinGetTitle("ahk_id " candidate)
                catch
                    continue
                if (InStr(title, "AppLaunchers.ahk")) {
                    hwnd := candidate
                    break
                }
            }
            if (hwnd)
                break
        }
    }
    DetectHiddenWindows prevDetect
    SetTitleMatchMode prevMatch
    return hwnd
}

Handy_AiModelRequestMsgId() {
    static msg := 0
    if (!msg)
        msg := DllCall("RegisterWindowMessage", "Str", HANDY_AI_MODEL_REQUEST_MSG_NAME, "UInt")
    return msg
}

Handy_OnAiModelRequest(wParam, lParam, *) {
    raw := Integer(wParam)
    slot := raw & 0xFF
    ; Slot only: another process already switched Handy. Do not run model UIA again.
    if (raw & 0x200) {
        if (slot >= 1 && slot <= 3) {
            global g_HandyAiPersistedSlot
            g_HandyAiPersistedSlot := slot
        }
        return
    }
    restartDictationIfStopped := (raw & 0x100) != 0
    restoreHwnd := Integer(lParam)
    SetTimer((*) => ExecuteHandyAiModelSelection(slot, false, restoreHwnd, restartDictationIfStopped), -1)
}

; Owner runs the switch here. Another host asks AppLaunchers when that script is up.
; After a reboot the menu often belongs to a different script while AppLaunchers
; is still down — switch Handy locally then.
; restartDictationIfStopped: Utility Shortcuts K/L only. If a take was running, start again after the switch.
Handy_RequestAiModelSelection(slot, restoreHwnd := 0, restartDictationIfStopped := false) {
    if (HandyAi_IsOwnerProcess()) {
        ExecuteHandyAiModelSelection(slot, false, restoreHwnd, restartDictationIfStopped)
        return true
    }
    msg := Handy_AiModelRequestMsgId()
    target := Handy_FindAppLaunchersHwnd()
    packed := restartDictationIfStopped ? (slot | 0x100) : slot
    if (msg && target) {
        try {
            PostMessage(msg, packed, restoreHwnd, , "ahk_id " target)
            return true
        } catch {
        }
    }
    ExecuteHandyAiModelSelection(slot, false, restoreHwnd, restartDictationIfStopped)
    return true
}

if (HandyAi_IsOwnerProcess())
    OnMessage(Handy_AiModelRequestMsgId(), Handy_OnAiModelRequest)