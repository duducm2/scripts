/* ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** **
    * Win + Alt + Shift symbol layer shortcuts (AHK v2)
    * â€¢ Provides system - wide symbol shortcuts
    ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** /
    /********************************************************************
     *   AVAILABLE WIN+ALT+SHIFT COMBINATIONS
     *   The following combinations are not currently in use:
     *
     *   Letters still free: P, U
     *
     *   Ctrl+Alt+Win+V: RESERVED — maximize active window (handled in WindowManagement.ahk;
     *   used by ZMK hold on minimize/close key). Do not bind another global ^!#v action here.
     *   Ctrl+Alt+Win+N: TEMPORARY — M365 Copilot auto-continue loop (toggle; remove block at file end)
     *   Numbers: 9 is free; 0-8 are used
     *   Symbols: ; ' [ ] \ | ` ~ @ # $ % ^ & * ( ) - _ = + { } : " < > ? /
     *
     *   Note: Some combinations use Ctrl+Alt+Shift+Arrow keys for extended mouse movement
********************************************************************/
#Requires AutoHotkey v2.0+

#SingleInstance Force

SetTitleMatchMode 2

; -----------------------------------------------------------------------------
; MODULE MAP - this file stays the runnable entry point / source of truth and
; #includes each module below. For a given feature, open just its small module
; (handy for low-context AI). See Shift keys/MODULARIZATION_PROGRESS.md.
;   Shift keys\helpers.ahk through m365_copilot_temp.ahk — 65 modules total
;   Still inline: preamble side effects, #HotIf resets, g_WikipediaScrollHistory,
;   VSCode evidence bootstrap (VSCodeEvidenceSearch.ahk at file end)
; -----------------------------------------------------------------------------

#include %A_ScriptDir%\env.ahk
#include vendor\UIA-v2\Lib\UIA.ahk
#include vendor\UIA-v2\Lib\UIA_Browser.ahk
#include %A_ScriptDir%\Utils.ahk
; Focus dwell watcher + #!+Y (Utils) must share this process so ToggleFocusMode sees the same globals as EnableFocusMode.
FocusBlackoutWatcher_Start()
; Volume: AppLaunchers also schedules retries; this catches Shift keys process when sessions register slightly later.
SetTimer(() => ApplyScriptMasterVolumeTarget(), -3500)
#include %A_ScriptDir%\infra\ipc\ShiftKeysIPC.ahk

; [SK module] Early globals and SafeDebugLog helpers -> Shift keys\helpers.ahk
#include %A_ScriptDir%\Shift keys\helpers.ahk

; [SK module] Config and cheat-sheet string utilities -> Shift keys\config.ahk
#include %A_ScriptDir%\Shift keys\config.ahk
;-------------------------------------------------------------------
; Cheat-sheet overlay (Win + Alt + Shift + A) â€" shows remapped shortcuts
;-------------------------------------------------------------------

; [SK module] Cheat sheet registry (all sheet strings) -> Shift keys\cheat_sheet_registry.ahk
#include %A_ScriptDir%\Shift keys\cheat_sheet_registry.ahk

; [SK module] Cheat sheet GUI, search, hold detection -> Shift keys\cheat_sheet_gui.ahk
#include %A_ScriptDir%\Shift keys\cheat_sheet_gui.ahk

; [SK module] Clip Angel fast copy mode and #!+1/#!+J -> Shift keys\fast_copy_clipangel.ahk
#include %A_ScriptDir%\Shift keys\fast_copy_clipangel.ahk

; [SK module] Env paths, ShowErr, CenterGuiOnActiveMonitor -> Shift keys\env_paths_centergui.ahk
#include %A_ScriptDir%\Shift keys\env_paths_centergui.ahk

; [SK module] OneNote hotkeys -> Shift keys\hotif_onenote.ahk
#include %A_ScriptDir%\Shift keys\hotif_onenote.ahk

; [SK module] ClipAngel hotkeys and filter selector -> Shift keys\hotif_clipangel.ahk
#include %A_ScriptDir%\Shift keys\hotif_clipangel.ahk

; [SK module] WhatsApp desktop hotkeys -> Shift keys\hotif_whatsapp.ahk
#include %A_ScriptDir%\Shift keys\hotif_whatsapp.ahk

;-------------------------------------------------------------------
; Outlook Reminder Window Shortcuts
;-------------------------------------------------------------------
; [SK module] Outlook reminder/appointment hotkeys -> Shift keys\hotif_outlook_reminder.ahk
#include %A_ScriptDir%\Shift keys\hotif_outlook_reminder.ahk

; [SK module] Teams meeting/chat predicate helpers -> Shift keys\teams_predicates.ahk
#include %A_ScriptDir%\Shift keys\teams_predicates.ahk
; [SK module] Teams meeting window hotkeys -> Shift keys\hotif_teams_meeting.ahk
#include %A_ScriptDir%\Shift keys\hotif_teams_meeting.ahk

;-------------------------------------------------------------------
; Wikipedia Shortcuts
;-------------------------------------------------------------------
; Global variable to track scroll position history (stack: most recent last)
global g_WikipediaScrollHistory := []

; [SK module] Wikipedia Chrome hotkeys -> Shift keys\hotif_wikipedia.ahk
#include %A_ScriptDir%\Shift keys\hotif_wikipedia.ahk

;-------------------------------------------------------------------
; [SK module] IsChromePdfViewerActive predicate -> Shift keys\predicates_chrome_pdf.ahk
#include %A_ScriptDir%\Shift keys\predicates_chrome_pdf.ahk

; [SK module] Chrome PDF viewer hotkeys -> Shift keys\hotif_chrome_pdf.ahk
#include %A_ScriptDir%\Shift keys\hotif_chrome_pdf.ahk

; [SK module] IsMercadoLivreActive and ML UIA helpers -> Shift keys\predicates_mercado_livre.ahk
#include %A_ScriptDir%\Shift keys\predicates_mercado_livre.ahk

; [SK module] IsShopeeActive and Shopee UIA helpers -> Shift keys\predicates_shopee.ahk
#include %A_ScriptDir%\Shift keys\predicates_shopee.ahk

; [SK module] Mercado Livre hotkeys -> Shift keys\hotif_mercado_livre.ahk
#include %A_ScriptDir%\Shift keys\hotif_mercado_livre.ahk
; [SK module] Shopee hotkeys -> Shift keys\hotif_shopee.ahk
#include %A_ScriptDir%\Shift keys\hotif_shopee.ahk

;-------------------------------------------------------------------
; Microsoft Teams Shortcuts (chat)
;-------------------------------------------------------------------
; [SK module] Teams chat window hotkeys and UIA -> Shift keys\hotif_teams_chat.ahk
#include %A_ScriptDir%\Shift keys\hotif_teams_chat.ahk

; [SK module] Outlook helper functions (part 1) -> Shift keys\outlook_helpers_01.ahk
#include %A_ScriptDir%\Shift keys\outlook_helpers_01.ahk
; [SK module] Outlook helper functions (part 2) -> Shift keys\outlook_helpers_02.ahk
#include %A_ScriptDir%\Shift keys\outlook_helpers_02.ahk

; [SK module] Outlook main window hotkeys -> Shift keys\hotif_outlook_main.ahk
#include %A_ScriptDir%\Shift keys\hotif_outlook_main.ahk

; Appointment/Meeting inspector-specific hotkeys
; [SK module] Outlook appointment inspector hotkeys (part 1) -> Shift keys\outlook_appointment_hotif_01.ahk
#include %A_ScriptDir%\Shift keys\outlook_appointment_hotif_01.ahk

; [SK module] Outlook appointment date/time helpers and hotkeys -> Shift keys\outlook_appointment_hotif_02.ahk
#include %A_ScriptDir%\Shift keys\outlook_appointment_hotif_02.ahk

; [SK module] Outlook appointment configuration palette -> Shift keys\outlook_appointment_palette.ahk
#include %A_ScriptDir%\Shift keys\outlook_appointment_palette.ahk

; [SK module] Outlook appointment UIA state checking -> Shift keys\outlook_appointment_uia.ahk
#include %A_ScriptDir%\Shift keys\outlook_appointment_uia.ahk

#HotIf

; [SK module] Google Chrome general hotkeys -> Shift keys\hotif_chrome_general.ahk
#include %A_ScriptDir%\Shift keys\hotif_chrome_general.ahk

; [SK module] ChatGPT hotkeys -> Shift keys\hotif_chatgpt.ahk
#include %A_ScriptDir%\Shift keys\hotif_chatgpt.ahk

;-------------------------------------------------------------------
; Settings Window Shortcuts
;-------------------------------------------------------------------
; [SK module] Windows Settings hotkeys -> Shift keys\hotif_settings.ahk
#include %A_ScriptDir%\Shift keys\hotif_settings.ahk
#HotIf

;-------------------------------------------------------------------
; Windows Explorer Shortcuts
;-------------------------------------------------------------------
; [SK module] Windows Explorer hotkeys -> Shift keys\hotif_explorer.ahk
#include %A_ScriptDir%\Shift keys\hotif_explorer.ahk
#HotIf

;-------------------------------------------------------------------
; Microsoft Paint Shortcuts
;-------------------------------------------------------------------
; [SK module] Excel and Paint hotkeys -> Shift keys\hotif_excel_mspaint.ahk
#include %A_ScriptDir%\Shift keys\hotif_excel_mspaint.ahk
#HotIf

;-------------------------------------------------------------------
; Power BI Shortcuts
;-------------------------------------------------------------------
; [SK module] Power BI hotkeys -> Shift keys\hotif_powerbi.ahk
#include %A_ScriptDir%\Shift keys\hotif_powerbi.ahk
#HotIf

; [SK module] Power BI drawer config helpers -> Shift keys\powerbi_helpers.ahk
#include %A_ScriptDir%\Shift keys\powerbi_helpers.ahk

;-------------------------------------------------------------------
; Gmail Shortcuts
;-------------------------------------------------------------------
; [SK module] Gmail hotkeys -> Shift keys\hotif_gmail.ahk
#include %A_ScriptDir%\Shift keys\hotif_gmail.ahk
#HotIf

; [SK module] Cursor/VS Code editor detection and UIA helpers -> Shift keys\cursor_predicates.ahk
#include %A_ScriptDir%\Shift keys\cursor_predicates.ahk
; [SK module] Cursor IDE hotkeys -> Shift keys\hotif_cursor.ahk
#include %A_ScriptDir%\Shift keys\hotif_cursor.ahk
#HotIf

; Shared Editor Shortcuts (Cursor + VS Code)
;-------------------------------------------------------------------
; [SK module] Cursor/VS Code editor hotkeys (part 1) -> Shift keys\hotif_editor_01.ahk
#include %A_ScriptDir%\Shift keys\hotif_editor_01.ahk
; [SK module] Cursor/VS Code editor hotkeys (part 2) -> Shift keys\hotif_editor_02.ahk
#include %A_ScriptDir%\Shift keys\hotif_editor_02.ahk
#HotIf

; [SK module] VSCode_TriggerGenerateCommitMessage helper -> Shift keys\vscode_commit_message.ahk
#include %A_ScriptDir%\Shift keys\vscode_commit_message.ahk

; VS Code IDE — VS Code-specific Shortcuts
;-------------------------------------------------------------------
; [SK module] VS Code hotkeys -> Shift keys\hotif_code.ahk
#include %A_ScriptDir%\Shift keys\hotif_code.ahk
#HotIf

; [SK module] Global Alt+U scroll AI feed and related -> Shift keys\hotif_scroll_ai.ahk
#include %A_ScriptDir%\Shift keys\hotif_scroll_ai.ahk

;-------------------------------------------------------------------
; Spotify Shortcuts
;-------------------------------------------------------------------
; [SK module] Spotify hotkeys -> Shift keys\hotif_spotify.ahk
#include %A_ScriptDir%\Shift keys\hotif_spotify.ahk
#HotIf

;-------------------------------------------------------------------
; Figma Shortcuts
;-------------------------------------------------------------------
; [SK module] Figma hotkeys -> Shift keys\hotif_figma.ahk
#include %A_ScriptDir%\Shift keys\hotif_figma.ahk
#HotIf

; [SK module] Mobills title WinActive hotkeys -> Shift keys\hotif_mobills.ahk
#include %A_ScriptDir%\Shift keys\hotif_mobills.ahk

; K/L month navigation: see #HotIf Mobills_ShouldHandleMonthNavKeys() below (single definition; skips text fields).

; [SK module] Mobills pagination unified -> Shift keys\mobills_pagination.ahk
#include %A_ScriptDir%\Shift keys\mobills_pagination.ahk

; [SK module] Mobills running overlay banner -> Shift keys\mobills_running_banner.ahk
#include %A_ScriptDir%\Shift keys\mobills_running_banner.ahk

; [SK module] Mobills URL-scoped month nav hotkeys -> Shift keys\mobills_hotkeys_fallback.ahk
#include %A_ScriptDir%\Shift keys\mobills_hotkeys_fallback.ahk

; [SK module] Google Keep hotkeys and reminder dismiss helpers -> Shift keys\hotif_google_keep.ahk
#include %A_ScriptDir%\Shift keys\hotif_google_keep.ahk
; [SK module] YouTube Chrome hotkeys -> Shift keys\hotif_youtube.ahk
#include %A_ScriptDir%\Shift keys\hotif_youtube.ahk

;-------------------------------------------------------------------
; Gemini Website Shortcuts
;-------------------------------------------------------------------
; [SK module] Gemini Chrome hotkeys (part 1) -> Shift keys\gemini_chrome_01.ahk
#include %A_ScriptDir%\Shift keys\gemini_chrome_01.ahk
; [SK module] Gemini Chrome tools drawer and hotkeys (part 2) -> Shift keys\gemini_chrome_02.ahk
#include %A_ScriptDir%\Shift keys\gemini_chrome_02.ahk

;-------------------------------------------------------------------
; M365 Copilot web (Chrome) — same Shift shortcuts as Gemini
;-------------------------------------------------------------------
; [SK module] M365 Copilot web Chrome hotkeys -> Shift keys\hotif_copilot_web.ahk
#include %A_ScriptDir%\Shift keys\hotif_copilot_web.ahk

;-------------------------------------------------------------------
; Google Maps Shortcuts (Chrome)
;-------------------------------------------------------------------
; [SK module] Google Maps Chrome hotkeys -> Shift keys\hotif_google_maps.ahk
#include %A_ScriptDir%\Shift keys\hotif_google_maps.ahk

;-------------------------------------------------------------------
; Google Search Shortcuts
;-------------------------------------------------------------------
; [SK module] Google Search Chrome hotkeys -> Shift keys\hotif_google_search.ahk
#include %A_ScriptDir%\Shift keys\hotif_google_search.ahk

;-------------------------------------------------------------------
; File Dialog (Namespace Tree Control) Shortcuts
;-------------------------------------------------------------------
; [SK module] File dialog hotkeys -> Shift keys\hotif_file_dialog.ahk
#include %A_ScriptDir%\Shift keys\hotif_file_dialog.ahk

; [SK module] IsFileDialogActive predicate -> Shift keys\predicates_file_dialog.ahk
#include %A_ScriptDir%\Shift keys\predicates_file_dialog.ahk

; [SK module] UIATreeInspector UIA focus and jiggle helpers -> Shift keys\uia_tree_inspector_helpers.ahk
#include %A_ScriptDir%\Shift keys\uia_tree_inspector_helpers.ahk

; [SK module] UIA Tree Inspector hotkeys -> Shift keys\hotif_uia_tree.ahk
#include %A_ScriptDir%\Shift keys\hotif_uia_tree.ahk

;-------------------------------------------------------------------
; [SK module] Settle Up hotkeys -> Shift keys\hotif_settleup.ahk
#include %A_ScriptDir%\Shift keys\hotif_settleup.ahk

;-------------------------------------------------------------------
; [SK module] Miro Chrome hotkeys -> Shift keys\hotif_miro.ahk
#include %A_ScriptDir%\Shift keys\hotif_miro.ahk

;-------------------------------------------------------------------
; PowerToys Command Palette Shortcuts
;-------------------------------------------------------------------
; [SK module] Command Palette helpers -> Shift keys\command_palette_helpers.ahk
#include %A_ScriptDir%\Shift keys\command_palette_helpers.ahk
; [SK module] Command Palette hotkeys -> Shift keys\hotif_command_palette.ahk
#include %A_ScriptDir%\Shift keys\hotif_command_palette.ahk

; [SK module] ChatGPT loading banner and wait helpers -> Shift keys\chatgpt_loading_helpers.ahk
#include %A_ScriptDir%\Shift keys\chatgpt_loading_helpers.ahk

; [SK module] TEMPORARY M365 Copilot auto-continue (^!#n) -> Shift keys\m365_copilot_temp.ahk
#include %A_ScriptDir%\Shift keys\m365_copilot_temp.ahk

; VS Code evidence -> PDF search loop (^!#o) — see VSCodeEvidenceSearch.ahk
global EVIDENCE_SEARCH_FROM_SHIFT_KEYS := true
#include %A_ScriptDir%\infra\tools\VSCodeEvidenceSearch.ahk
#InputLevel 10
EvidenceSearch_BindHotkey()