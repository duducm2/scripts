#Requires AutoHotkey v2.0+
#SingleInstance Force
#include %A_ScriptDir%\env.ahk
#include %A_ScriptDir%\lib\study\StudyLinkHelpers.ahk
#include %A_ScriptDir%\Utils\git_cli.ahk

global g_StudyLinkSubmenuGui := ""

#include vendor\UIA-v2\Lib\UIA.ahk
#include vendor\UIA-v2\Lib\UIA_Browser.ahk
#include %A_ScriptDir%\lib\Media.ahk
#include %A_ScriptDir%\lib\SpotifyWASAPI.ahk
#include %A_ScriptDir%\infra\ipc\ClipboardFiles.ahk

; UIA ControlType constants (Button=50000). Shared with Gemini.ahk focus helpers.
global UIA_ControlType_Button := 50000

; -----------------------------------------------------------------------------
; MODULE MAP - Utils.ahk stays the runnable entry point / shared library and
; #includes each module below. For a given feature, open just its small module.
; Early BANNER_ACCENT_* and GEMINI_PROMPT_* globals stay here (must load first).
; lib\CopilotWeb.ahk #include stays inline before d2c_flow_manager module.
; See Utils/MODULARIZATION_PROGRESS.md for the full module list.
; Pack imports (Finance, Memory Palace, Tasks): docs/prompt-data-output-and-finance-packs.md
; -----------------------------------------------------------------------------

; -----------------------------------------------------------------------------
; Timing and window utilities for performance measurement and lightweight checks
; -----------------------------------------------------------------------------
; StartTimer() -> returns a timer object { start: A_TickCount }
StartTimer() {
    return { start: A_TickCount }
}

; EndTimer(timer, message := "") -> returns elapsed ms; optional message is appended to chrome_detach.log
EndTimer(timer, message := "") {
    elapsed := A_TickCount - timer.start
    if (message != "") {
        try {
            FileAppend(message . ": " . elapsed . " ms`n", A_ScriptDir . "\chrome_detach.log")
        } catch {
            ; swallow logging errors
        }
    }
    return elapsed
}

; WindowExists(windowTitle) - Checks if window exists without waiting
; Returns: True if window exists, false otherwise
WindowExists(windowTitle) {
    prev := A_TitleMatchMode
    SetTitleMatchMode(2)
    exists := !!WinExist(windowTitle)
    SetTitleMatchMode(prev)
    return exists
}

; =============================================================================
; Semantic banner accents (must be defined early)
; Some startup/update helpers call ShowCenteredOverlay_Utils / StandardLoadingBar_Show
; before the later globals block is reached.
; =============================================================================
global BANNER_ACCENT_SUCCESS := "27AE60"      ; Dark green: positive / success
global BANNER_ACCENT_ERROR := "C0392B"        ; Red: negative / error
global BANNER_ACCENT_INTERMEDIATE := "F1C40F" ; Yellow: loading, actionable, neutral
global BANNER_ACCENT_INFO := "2980B9"         ; Blue: info / alternate mode

; Possible Gemini prompt field names (EN and PT) for work/personal env. Used by FindGeminiPromptField.
global GEMINI_PROMPT_FIELD_NAMES := ["Enter a prompt for Gemini", "Enter a prompt here",
    "Digite um prompt para o Gemini", "Digite um prompt aqui"]
global GEMINI_PROMPT_FOCUS_POLL_MS := 25
global GEMINI_PROMPT_FOCUS_TIMEOUT_MS := 300
global GEMINI_OPEN_FAST_SETTLE_MS := 0

; [Utils module] Timing helpers and FindGeminiPromptField -> Utils\helpers_timing_gemini.ahk
#include %A_ScriptDir%\Utils\helpers_timing_gemini.ahk
; [Utils module] F11 fullscreen detection and toggle -> Utils\f11_fullscreen.ahk
#include %A_ScriptDir%\Utils\f11_fullscreen.ahk
; [Utils module] Chrome detach helpers (part 1) -> Utils\chrome_detach_01.ahk
#include %A_ScriptDir%\Utils\chrome_detach_01.ahk
; [Utils module] Chrome detach context menu phases (part 2) -> Utils\chrome_detach_02.ahk
#include %A_ScriptDir%\Utils\chrome_detach_02.ahk

; [Utils module] Chrome detach tab (part 3) and detach entry -> Utils\chrome_detach_03.ahk
#include %A_ScriptDir%\Utils\chrome_detach_03.ahk

; [Utils module] Gemini mode picker mouse + UIA -> Utils\gemini_mode_picker.ahk
#include %A_ScriptDir%\Utils\gemini_mode_picker.ahk

; -----------------------------------------------------------------------------
; This script consolidates various utility hotkeys.
; -----------------------------------------------------------------------------

; [Utils module] Hotstrings core (InsertText, prompt file helpers) -> Utils\hotstrings_core.ahk
#include %A_ScriptDir%\Utils\hotstrings_core.ahk
; [Utils module] Strip prompt human reminders after --- (Shift+H Copilot/Gemini)
#include %A_ScriptDir%\Utils\prompt_strip_reminders.ahk
; [Utils module] Files and links quick-open (InitQuickOpenFiles) -> Utils\files_links.ahk
#include %A_ScriptDir%\Utils\files_links.ahk
; [Utils module] Macros system RegisterMacro and assignments -> Utils\macros_system.ahk
#include %A_ScriptDir%\Utils\macros_system.ahk
; [Utils module] Clip Angel merge non-favorite clips -> Utils\clip_angel_merge.ahk
#include %A_ScriptDir%\Utils\clip_angel_merge.ahk
; [Utils module] Clip Angel activate with focus correction -> Utils\clip_angel_activate.ahk
#include %A_ScriptDir%\Utils\clip_angel_activate.ahk
; [Utils module] Clip Angel mark favorite and related flows -> Utils\clip_angel_favorite.ahk
#include %A_ScriptDir%\Utils\clip_angel_favorite.ahk

; [Utils module] Handy AI model configuration map and persistence -> Utils\handy_ai_model_config.ahk
#include %A_ScriptDir%\Utils\handy_ai_model_config.ahk
; [Utils module] ShowAiModelSelector GUI -> Utils\handy_ai_model_gui.ahk
#include %A_ScriptDir%\Utils\handy_ai_model_gui.ahk

; [Utils module] Gemini-to-Cursor transfer numeric window selector -> Utils\gemini_cursor_transfer.ahk
#include %A_ScriptDir%\Utils\gemini_cursor_transfer.ahk

; [Utils module] Cursor AI text field focus and VS Code chat helpers -> Utils\cursor_composer_focus.ahk
#include %A_ScriptDir%\Utils\cursor_composer_focus.ahk

; [Utils module] Language flag indicator and status banners -> Utils\language_flag_indicator.ahk
#include %A_ScriptDir%\Utils\language_flag_indicator.ahk

; [Utils module] Handy UIA helper functions and ShowAiModelSelector support -> Utils\handy_uia_helpers.ahk
#include %A_ScriptDir%\Utils\handy_uia_helpers.ahk

; [Utils module] Send dictation? [R] Handy History last recording Play -> Utils\handy_replay_last.ahk
#include %A_ScriptDir%\Utils\handy_replay_last.ahk

; [Utils module] Send dictation? [B] toggle model + History re-transcribe + copy -> Utils\handy_retranscribe_last.ahk
#include %A_ScriptDir%\Utils\handy_retranscribe_last.ahk

; [Utils module] SelectAiModelInHandy entry and pre-movement warning -> Utils\handy_selector_entry.ahk
#include %A_ScriptDir%\Utils\handy_selector_entry.ahk

; [Utils module] Standard loading bar show/update/hide lifecycle -> Utils\standard_loading_bar.ahk
#include %A_ScriptDir%\Utils\standard_loading_bar.ahk

; [Utils module] RichEdit mnemonic emphasis for selector modals -> Utils\mnemonic_rich.ahk
#include %A_ScriptDir%\Utils\mnemonic_rich.ahk

; [Utils module] Hotstring Gemini banner and D2C preset helpers -> Utils\hotstring_gemini_banner.ahk
#include %A_ScriptDir%\Utils\hotstring_gemini_banner.ahk

; [Utils module] Gemini navigate/focus/paste helpers (D2C, legacy dictation) -> Utils\gemini_paste_helpers.ahk
#include %A_ScriptDir%\Utils\gemini_paste_helpers.ahk

; [Lib] Chrome chat feed scroll-to-bottom (JS-first) — before companion libs that call it
#include %A_ScriptDir%\lib\ChromeChatScroll.ahk

#include %A_ScriptDir%\lib\CopilotWeb.ahk
#include %A_ScriptDir%\lib\GeminiEnterprise.ahk

; [Lib] Per-companion Fast/Deep/Models INI + apply helpers
#include %A_ScriptDir%\lib\AiCompanionModels.ahk
; [Utils] Shared Shift+L AI companion model list selector
#include %A_ScriptDir%\Utils\ai_companion_model_selector.ahk

; [Utils module] AutoSlot user ignore list (#!+L R/I) -> Utils\autoslot_user_excludes.ahk
#include %A_ScriptDir%\Utils\autoslot_user_excludes.ahk

; [Utils module] Post-dictation visible-window paste picker -> Utils\dictation_visible_paste.ahk
#include %A_ScriptDir%\Utils\dictation_visible_paste.ahk

; [Utils module] Paste field mapping (dynamic app -> text field focus) -> Utils\paste_field_mapping.ahk
#include %A_ScriptDir%\Utils\paste_field_mapping.ahk

; [Utils module] D2C_FlowManager dictation-Gemini-Cursor state machine -> Utils\d2c_flow_manager.ahk
#include %A_ScriptDir%\Utils\d2c_flow_manager.ahk

; [Utils module] Deprecated dictation Gemini confirm banner -> Utils\dictation_legacy.ahk
#include %A_ScriptDir%\Utils\dictation_legacy.ahk
; [Utils module] Global sound toggle and script audio helpers -> Utils\global_sound_audio.ahk
#include %A_ScriptDir%\Utils\global_sound_audio.ahk

; [Utils module] ToggleOutlookAndTeams macro -> Utils\toggle_outlook_teams.ahk
#include %A_ScriptDir%\Utils\toggle_outlook_teams.ahk

; [Utils module] Email note macro (Outlook work / Gmail personal) -> Utils\email_note_macro.ahk
#include %A_ScriptDir%\Utils\email_note_macro.ahk

; [Utils module] CheckAndOpenOutlookTeams prompt helper -> Utils\outlook_teams_check.ahk
#include %A_ScriptDir%\Utils\outlook_teams_check.ahk

; [Utils module] Shared Jump-to-Chat logic for Teams -> Utils\teams_jump_to_chat.ahk
; After outlook_teams_check (uses CheckAndOpenOutlookTeams). No TeamsContext include.
#include %A_ScriptDir%\Utils\teams_jump_to_chat.ahk

; [Utils module] Shared Jump-to-Chat logic for WhatsApp -> Utils\whatsapp_jump_to_chat.ahk
#include %A_ScriptDir%\Utils\whatsapp_jump_to_chat.ahk

; [Utils module] Spotify Immersion (Shift+I + D2C [P]) -> Utils\spotify_immerse.ahk
#include %A_ScriptDir%\Utils\spotify_immerse.ahk

; [Utils module] Dictation → Spotify search + immerse -> Utils\spotify_dictation_play.ahk
#include %A_ScriptDir%\Utils\spotify_dictation_play.ahk

; [Utils module] Google Search first-result click (Shift+U + D2C [C]) -> Utils\google_search_first_result.ahk
#include %A_ScriptDir%\Utils\google_search_first_result.ahk

; [Utils module] Send dictation? [C] new Chrome window, paste address bar, Enter -> Utils\chrome_dictation_navigate.ahk
#include %A_ScriptDir%\Utils\chrome_dictation_navigate.ahk

; [Utils module] Dictation clipboard cleanup countdown -> Utils\dictation_cleanup.ahk
#include %A_ScriptDir%\Utils\dictation_cleanup.ahk

; [Utils module] Dictation merge non-favorite clips countdown -> Utils\dictation_merge.ahk
#include %A_ScriptDir%\Utils\dictation_merge.ahk

; [Utils module] Clean clipboard countdown macro -> Utils\cleanclipboard.ahk
#include %A_ScriptDir%\Utils\cleanclipboard.ahk

; [Utils module] Project data for Cursor window focus selector -> Utils\project_data_cursor.ahk
#include %A_ScriptDir%\Utils\project_data_cursor.ahk
; [Utils module] Prompt registry for Utility Shortcuts -> Utils\prompt_data.ahk
#include %A_ScriptDir%\Utils\prompt_data.ahk
; [Utils module] Prompt render-at-paste variables and includes -> Utils\prompt_render.ahk
#include %A_ScriptDir%\Utils\prompt_render.ahk
; [Utils module] Prompt structural lint on save -> Utils\prompt_lint.ahk
#include %A_ScriptDir%\Utils\prompt_lint.ahk
; [Utils module] Prompt-aware git history and commit -> Utils\prompt_git.ahk
#include %A_ScriptDir%\Utils\prompt_git.ahk
; [Utils module] Prompt usage log -> Utils\prompt_usage.ahk
#include %A_ScriptDir%\Utils\prompt_usage.ahk
; [Utils module] Context file presets for prompt editor -> Utils\prompt_context_presets.ahk
#include %A_ScriptDir%\Utils\prompt_context_presets.ahk
#include %A_ScriptDir%\Utils\hotstring_data.ahk

; [Utils module] Global AI generation state U macro -> Utils\ai_generation_state.ahk
#include %A_ScriptDir%\Utils\ai_generation_state.ahk

; [Utils module] Jump mouse to window center #!+Q -> Utils\jump_mouse_middle.ahk
#include %A_ScriptDir%\Utils\jump_mouse_middle.ahk

; [Utils module] Handy AI model selector hotkey #!+C -> Utils\handy_selector_hotkey.ahk
#include %A_ScriptDir%\Utils\handy_selector_hotkey.ahk

; [Utils module] Desktop to Recycle Bin macro -> Utils\desktop_recycle.ahk
#include %A_ScriptDir%\Utils\desktop_recycle.ahk

; [Utils module] Cut newest Desktop item (#!+O) -> Utils\desktop_cut_newest.ahk
#include %A_ScriptDir%\Utils\desktop_cut_newest.ahk

; [Utils module] Click sequence data (#!+U Sequences / #!+9 runner) -> Utils\click_sequence_data.ahk
#include %A_ScriptDir%\Utils\click_sequence_data.ahk
; [Utils module] Click sequence UIA runner -> Utils\click_sequence_runner.ahk
#include %A_ScriptDir%\Utils\click_sequence_runner.ahk
; [Utils module] Click sequence CRUD GUI -> Utils\click_sequence_gui.ahk
#include %A_ScriptDir%\Utils\click_sequence_gui.ahk
; [Utils module] AI Companion Quick Download (#!+9 single-tap) -> Utils\ai_quick_download.ahk
#include %A_ScriptDir%\Utils\ai_quick_download.ahk
; [Utils module] Click sequence Hardcoded Script registry -> Utils\click_sequence_scripts.ahk
#include %A_ScriptDir%\Utils\click_sequence_scripts.ahk

; [Utils module] Mouse jump helpers and prediction squares -> Utils\mouse_jump_arrows.ahk
#include %A_ScriptDir%\Utils\mouse_jump_arrows.ahk

; [Utils module] Square selector mouse jump (part 1) -> Utils\square_selector_mouse_jump_01.ahk
#include %A_ScriptDir%\Utils\square_selector_mouse_jump_01.ahk
; [Utils module] Square selector mouse jump (part 2) -> Utils\square_selector_mouse_jump_02.ahk
#include %A_ScriptDir%\Utils\square_selector_mouse_jump_02.ahk

; [Utils module] Win+Alt+Shift+Arrow five-step hotkeys -> Utils\mouse_jump_hotkeys.ahk
#include %A_ScriptDir%\Utils\mouse_jump_hotkeys.ahk

; [Utils module] Cross-process AutoSlot place request (Study Topic QuickLook) -> Utils\autoslot_place_ipc.ahk
#include %A_ScriptDir%\Utils\autoslot_place_ipc.ahk

; [Utils module] Peek PDF / QuickLook study helpers (part 1) -> Utils\peek_pdf_study_01.ahk
#include %A_ScriptDir%\Utils\peek_pdf_study_01.ahk
; [Utils module] Peek PDF / QuickLook study helpers (part 2) -> Utils\peek_pdf_study_02.ahk
#include %A_ScriptDir%\Utils\peek_pdf_study_02.ahk

; [Utils module] Peek PDF / QuickLook study helpers (part 3) -> Utils\peek_pdf_study_03.ahk
#include %A_ScriptDir%\Utils\peek_pdf_study_03.ahk

; [Utils module] Win+Alt+Shift+X → Memory Palace -> Utils\study_hotkey_x.ahk
#include %A_ScriptDir%\Utils\study_hotkey_x.ahk

; [Utils module] Modal ListView letter jump helpers -> Utils\modal_list_letter_jump.ahk
#include %A_ScriptDir%\Utils\modal_list_letter_jump.ahk

; [Utils module] Hotstring selector system core -> Utils\hotstring_selector_core.ahk
#include %A_ScriptDir%\Utils\hotstring_selector_core.ahk

; [Utils module] Context file browser (Win+Alt+Shift+N) -> Utils\context_file_browser.ahk
#include %A_ScriptDir%\Utils\context_file_browser.ahk

; [Utils module] CleanupHotstringSelector -> Utils\hotstring_selector_cleanup.ahk
#include %A_ScriptDir%\Utils\hotstring_selector_cleanup.ahk

; [Utils module] HandleHotstringChar and Gemini paste helpers -> Utils\hotstring_selector_handlers_01.ahk
#include %A_ScriptDir%\Utils\hotstring_selector_handlers_01.ahk

; [Utils module] Hotstring selector utility category handlers -> Utils\hotstring_selector_handlers_02.ahk
#include %A_ScriptDir%\Utils\hotstring_selector_handlers_02.ahk

; [Utils module] ShowHotstringSelector GUI and category view -> Utils\hotstring_selector_gui.ahk
#include %A_ScriptDir%\Utils\hotstring_selector_gui.ahk

; [Utils module] Single-form Add/Edit prompt dialog -> Utils\prompt_editor_gui.ahk
#include %A_ScriptDir%\Utils\prompt_editor_gui.ahk
; [Utils module] Context preset manager GUI -> Utils\prompt_context_presets_gui.ahk
#include %A_ScriptDir%\Utils\prompt_context_presets_gui.ahk

; [Utils module] Local finance app (Utility Shortcuts [F])
#include %A_ScriptDir%\Utils\finance_helpers.ahk
#include %A_ScriptDir%\Utils\finance_launcher.ahk
#include %A_ScriptDir%\Utils\finance_transactions.ahk
#include %A_ScriptDir%\Utils\finance_categories.ahk
#include %A_ScriptDir%\Utils\finance_accounts.ahk
#include %A_ScriptDir%\Utils\finance_credit_card.ahk
#include %A_ScriptDir%\Utils\finance_budgets.ahk
#include %A_ScriptDir%\Utils\finance_goals.ahk
#include %A_ScriptDir%\Utils\finance_recurring.ahk
; [Utils module] Canonical Desktop pack paths for AI imports -> Utils\pack_import_desktop.ahk
#include %A_ScriptDir%\Utils\pack_import_desktop.ahk
#include %A_ScriptDir%\Utils\finance_import.ahk
#include %A_ScriptDir%\Utils\finance_git_sync.ahk
; [Utils module] Win+Alt+Shift+D tap-dance → Tasks / Finance / Memory Palace -> Utils\finance_hotkey_d.ahk
#include %A_ScriptDir%\Utils\finance_hotkey_d.ahk

; [Utils module] Memory Palace (Utility Shortcuts [N])
#include %A_ScriptDir%\Utils\mnemonic_palace_helpers.ahk
#include %A_ScriptDir%\Utils\mnemonic_palace_practice_md.ahk
#include %A_ScriptDir%\Utils\mnemonic_palace_launcher.ahk
#include %A_ScriptDir%\Utils\mnemonic_palace_studies.ahk
#include %A_ScriptDir%\Utils\mnemonic_palace_palaces.ahk
#include %A_ScriptDir%\Utils\mnemonic_palace_beasts.ahk
#include %A_ScriptDir%\Utils\mnemonic_palace_atoms.ahk
#include %A_ScriptDir%\Utils\mnemonic_palace_plans.ahk
#include %A_ScriptDir%\Utils\mnemonic_palace_help.ahk
#include %A_ScriptDir%\Utils\mnemonic_palace_import.ahk
#include %A_ScriptDir%\Utils\mnemonic_palace_git_sync.ahk

; [Utils module] Tasks (Utility Shortcuts [T])
#include %A_ScriptDir%\Utils\task_helpers.ahk
#include %A_ScriptDir%\Utils\task_launcher.ahk
#include %A_ScriptDir%\Utils\task_projects.ahk
#include %A_ScriptDir%\Utils\task_tasks.ahk
#include %A_ScriptDir%\Utils\task_info.ahk
#include %A_ScriptDir%\Utils\task_migrate.ahk

; [Utils module] Dynamic prompt context picker -> Utils\prompt_context_picker.ahk
#include %A_ScriptDir%\Utils\prompt_context_picker.ahk

; [Utils module] ClipAngel last clip → Desktop (.txt) -> Utils\clip_angel_export_desktop.ahk
#include %A_ScriptDir%\Utils\clip_angel_export_desktop.ahk
; [Utils module] Import Management hub (finance & palace imports) -> Utils\import_mgmt_launcher.ahk
; Pack-import agent docs -> docs\prompt-data-output-and-finance-packs.md
#include %A_ScriptDir%\Utils\import_mgmt_launcher.ahk
; [Utils module] Utility Shortcuts [G] commit+push scripts + notes repos -> Utils\utility_git_push.ahk
#include %A_ScriptDir%\Utils\utility_git_push.ahk
; [Utils module] Utility shortcuts #!+U / #!+W Macros and ^!# secondary triggers -> Utils\utility_shortcuts.ahk
#include %A_ScriptDir%\Utils\utility_shortcuts.ahk

; [Utils module] Focus mode multi-monitor blackout (#!+Y) -> Utils\focus_mode.ahk
#include %A_ScriptDir%\Utils\focus_mode.ahk

; [Utils module] Print Screen chime, global Escape hotkey -> Utils\print_screen_escape.ahk
#include %A_ScriptDir%\Utils\print_screen_escape.ahk

; [Utils module] Dictation indicator, ~#!+0 hotkey, ToggleDictationMode -> Utils\dictation_toggle.ahk
#include %A_ScriptDir%\Utils\dictation_toggle.ahk

; Handy model hotkeys and language flag: AppLaunchers.ahk only (after #!+C / ^!#9 / ^!#b are registered).
HandyAi_ConfigureProcessOwnership()