; =============================================================================
; Utils module: ai_generation_state.ahk
; Global AI generation state U macro
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Global AI generation state: Cursor + Gemini stop-button detectors (Efficiency Canon)
; =============================================================================
; Cursor: Type 50026 (Group), ClassName contains "stop-button".
; Gemini: Chrome window title contains "gemini"; Type 50000, Name "Stop response", ClassName match.
; =============================================================================
Cursor_HasGeneratingStopButton() {
    global UIA
    try {
        cursorHwnds := WinGetList("ahk_exe Cursor.exe")
        if (!cursorHwnds.Length)
            return false
        cr := UIA.CreateCacheRequest(["Type", "ClassName"], , 5)
        for hwnd in cursorHwnds {
            if (!hwnd || !WinExist("ahk_id " hwnd))
                continue
            try {
                root := UIA.ElementFromHandleBuildCache(cr, hwnd)
            } catch {
                try root := UIA.ElementFromHandle(hwnd)
                catch
                    continue
            }
            if (!root)
                continue
            try {
                el := root.FindFirstBuildCache(cr, { Type: 50026, ClassName: "stop-button", matchmode: "Substring" }, 4
                )
                if (el)
                    return true
            } catch {
            }
        }
    } catch {
    }
    return false
}

Gemini_HasGeneratingStopButton() {
    global UIA
    try {
        for hwnd in WinGetList("ahk_exe chrome.exe") {
            if (!hwnd || !WinExist("ahk_id " hwnd))
                continue
            try {
                if (!IsConsumerGeminiChromeTitle(WinGetTitle("ahk_id " hwnd)))
                    continue
            } catch {
                continue
            }
            try {
                cr := UIA.CreateCacheRequest(["Type", "ClassName", "Name"], , 5)
                root := UIA.ElementFromHandleBuildCache(cr, hwnd)
            } catch {
                try root := UIA.ElementFromHandle(hwnd)
                catch
                    continue
            }
            if (!root)
                continue
            ; Stop response: Type 50000, Name "Stop response", ClassName contains "send-button" and "stop"
            try {
                el := root.FindFirstBuildCache(cr, { Type: 50000, Name: "Stop response", ClassName: "send-button",
                    matchmode: "Substring" }, 4)
                if (el)
                    return true
            } catch {
            }
        }
    } catch {
    }
    return false
}

IsAnyAiGenerating() {
    return Cursor_HasGeneratingStopButton() || Gemini_HasGeneratingStopButton()
}

PlayAiWorkingStateSound(isWorking) {
    try {
        if (isWorking)
            ScriptSoundPlay(A_ScriptDir . "\assets\sounds\robots-are-working.wav")
        else
            ScriptSoundPlay(A_ScriptDir . "\assets\sounds\no-robot-working.wav")
    } catch {
    }
}

; =============================================================================
; U macro: Global AI generation state (Cursor + Gemini) with sound and banner
; =============================================================================
; Runs Cursor + Gemini stop-button checks, plays robots-are-working / no-robot-working,
; shows red banner when any AI is working, green when none.
; =============================================================================
Cursor_FindComposerIconAcrossInstances() {
    global BANNER_ACCENT_SUCCESS, BANNER_ACCENT_ERROR
    try {
        isWorking := IsAnyAiGenerating()
        PlayAiWorkingStateSound(isWorking)
        if (isWorking)
            ShowCenteredOverlay_Utils("AI is working (stop button found)", 2000, BANNER_ACCENT_ERROR)
        else
            ShowCenteredOverlay_Utils("No AI is working", 2000, BANNER_ACCENT_SUCCESS)
    } catch Error as e {
        PlayAiWorkingStateSound(false)
        ShowCenteredOverlay_Utils("No AI is working", 2000, BANNER_ACCENT_SUCCESS)
    }
}

; Initialize macros
InitMacros() {
    ; Quick Update to Your Scripts macro
    RegisterMacro(QuickUpdateScripts, "⚡ Quick Update to Your Scripts")
    ; Add specific word to Handy macroh
    RegisterMacro(AddWordToHandy, "➕ Add specific word to Handy")
    ; Toggle Outlook and Teams macro
    RegisterMacro(ToggleOutlookAndTeams, "🔄 Toggle Outlook & Teams")
    ; Clean the Clipboard macro (assigned to "P")
    RegisterMacro(CleanClipboard, "🧹 Clean the Clipboard", "p")
    RegisterMacro(UnescapeMarkdownClipboard, "📋 Unescape markdown clipboard", "e")
    ; Toggle Sound macro
    RegisterMacro(ToggleSoundState, "🔊 Toggle Sound (Mute/Unmute)")
    ; Global AI generation state: Cursor + Gemini (assigned to "U")
    RegisterMacro(Cursor_FindComposerIconAcrossInstances, "🔍 AI working? (Cursor + Gemini)", "u")
    ; Mark Last Clip as Favorite macro (assigned to "J")
    RegisterMacro(MarkLastClipAsFavorite, "⭐ Mark Last Clip as Favorite", "j")
    ; Move Desktop to Recycle Bin (assigned to "N") - red banner, Y/N confirm
    RegisterMacro(DesktopToRecycle_Trigger, "🗑️ Move Desktop to Recycle Bin", "n")
}

InitMacros()

; ------------
; Optional: scope Explorer-only hotstrings used for renaming
; Uncomment to restrict selected triggers to File Explorer or Save dialogs
;------------
;#HotIf WinActive("ahk_exe explorer.exe") || WinActive("ahk_class #32770")
;:o:gdash::
;    InsertText("GS_E&S_CIP Dashboard research and design")
;return
;#HotIf

; --- Hotkeys & Functions -----------------------------------------------------

; Ensure per-monitor DPI awareness so coordinates are physical pixels across mixed scaling
InitDpiAwareness() {
    static PER_MONITOR_AWARE_V2 := -4 ; DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2
    try DllCall("SetProcessDpiAwarenessContext", "ptr", PER_MONITOR_AWARE_V2, "ptr")
}

InitDpiAwareness()

; Auto-execute: show success after QuickUpdateScripts relaunches AppLaunchers.ahk (includes this file) with "/Updated".
if (A_Args.Length > 0 && A_Args[1] = "/Updated") {
    try {
        ShowCenteredOverlay_Utils("✅ Scripts updated and relaunched", 6500, BANNER_ACCENT_SUCCESS)
        soundPath := A_ScriptDir "\assets\sounds\quick-update-success.wav"
        ; Play success chime to completion before scheduling volume: async SoundPlay can register a new session after
        ; the first Apply pass, leaving that session at a low default (~10% in the mixer) until something re-enumerates.
        try {
            if (FileExist(soundPath))
                ScriptSoundPlay(soundPath, true)
        } catch {
        }
        ; After all scripts have been started (this block runs last in include order for AppLaunchers /Updated), apply AHK volume - not at Quick Update start (old sessions / dead timers).
        ScheduleApplyScriptMasterVolumeTargetAfterQuickUpdate()
        if (HandyAi_IsOwnerProcess())
            LanguageFlag_InitFromPersistedSlot()
    } catch {
    }
}
