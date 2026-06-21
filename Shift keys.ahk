/* ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** **
    * Win + Alt + Shift symbol layer shortcuts (AHK v2)
    * â€¢ Provides system - wide symbol shortcuts
    ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** /
    /********************************************************************
     *   AVAILABLE WIN+ALT+SHIFT COMBINATIONS
     *   The following combinations are not currently in use:
     *
     *   Letters still free: P, U
     *   Win+Alt+Shift+L: Outlook Copilot shortcut modal (global; actions target New Outlook)
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
; (handy for low-context AI). Anything not listed still lives inline in this file.
;   Shift keys\helpers.ahk through m365_copilot_temp.ahk — see #include list below (59 modules)
;   Remaining inline: ML/Shopee predicates, IsChromePdfViewerActive, IsFileDialogActive,
;   UIATreeInspector helpers, VSCode_TriggerGenerateCommitMessage (orchestrator glue between modules)
; -----------------------------------------------------------------------------

#include %A_ScriptDir%\env.ahk
#include UIA-v2\Lib\UIA.ahk
#include UIA-v2\Lib\UIA_Browser.ahk
#include %A_ScriptDir%\Utils.ahk
; Focus dwell watcher + #!+Y (Utils) must share this process so ToggleFocusMode sees the same globals as EnableFocusMode.
FocusBlackoutWatcher_Start()
; Volume: AppLaunchers also schedules retries; this catches Shift keys process when sessions register slightly later.
SetTimer(() => ApplyScriptMasterVolumeTarget(), -3500)
#include %A_ScriptDir%\aux\ShiftKeysIPC.ahk
#include %A_ScriptDir%\CheatSheetRich.ahk

; [SK module] Early globals and SafeDebugLog helpers -> Shift keys\helpers.ahk
#include %A_ScriptDir%\Shift keys\helpers.ahk

; [SK module] Config and cheat-sheet string utilities -> Shift keys\config.ahk
#include %A_ScriptDir%\Shift keys\config.ahk
;-------------------------------------------------------------------
; Cheat-sheet overlay (Win + Alt + Shift + A) â€" shows remapped shortcuts
;-------------------------------------------------------------------

; [SK module] cheatSheets map population (Mercado Livre, Shopee) -> Shift keys\cheat_sheet_data.ahk
#include %A_ScriptDir%\Shift keys\cheat_sheet_data.ahk

; [SK module] Per-app Win+Alt+Shift hotkey cheat sheet definitions -> Shift keys\app_hotkeys.ahk
#include %A_ScriptDir%\Shift keys\app_hotkeys.ahk

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
; Chrome PDF Viewer Shortcuts
;-------------------------------------------------------------------
IsChromePdfViewerActive() {
    ; Hard gate: avoid conflicts with non-Chrome apps
    if !WinActive("ahk_exe chrome.exe")
        return false

    try {
        ; #region agent log
        AgentDebugLog("H1", "IsChromePdfViewerActive_entry")
        ; #endregion
        uia := UIA_Browser("ahk_exe chrome.exe")

        ; Strong fingerprint: Chrome's built-in PDF viewer extension web area
        ; From UIA tree: chrome-extension://mhjfbmdgcfjbbpaeojofohoefgiehjai/index.html
        if (uia.FindElement({ Type: 50030, Value: "chrome-extension://mhjfbmdgcfjbbpaeojofohoefgiehjai", matchmode: "Substring" })) {
            ; #region agent log
            AgentDebugLog("H2", "IsChromePdfViewerActive_extension_match")
            ; #endregion
            return true
        }

        ; Fallback: stable, non-localized PDF toolbar controls
        if (uia.FindElement({ AutomationId: "pageSelector" }) && uia.FindElement({ AutomationId: "save" })) {
            ; #region agent log
            AgentDebugLog("H3", "IsChromePdfViewerActive_toolbar_match")
            ; #endregion
            return true
        }
    } catch {
    }

    ; #region agent log
    AgentDebugLog("H4", "IsChromePdfViewerActive_return_false")
    ; #endregion
    return false
}

; [SK module] Chrome PDF viewer hotkeys -> Shift keys\hotif_chrome_pdf.ahk
#include %A_ScriptDir%\Shift keys\hotif_chrome_pdf.ahk

;-------------------------------------------------------------------
; Mercado Livre (Brazil) Shortcuts
;-------------------------------------------------------------------
; Cache for IsMercadoLivreActive (per efficiency-canon: cache-first with validation).
; Invalidated when foreground HWND changes so we only run UIA once per window/tab focus.
global g_ML_CacheHwnd := 0
global g_ML_CacheResult := false
; Cache for initial-page-load workaround: right-click + close context menu to make hotkeys work (once per window).
global g_ML_ReceptivityHwnd := 0

; Workaround for ML: hotkeys fail on initial page load until page is "initialized". Right-click on empty area then close context menu.
; force=true: run even if cache says we already did (error-driven retry when hotkey action failed).
ML_EnsureHotkeyReceptivity(force := false) {
    global g_ML_ReceptivityHwnd
    if !WinActive("ahk_exe chrome.exe")
        return
    hwnd := WinExist("A")
    if (!hwnd)
        return
    if (!force && hwnd = g_ML_ReceptivityHwnd && WinExist("ahk_id " g_ML_ReceptivityHwnd))
        return
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        try
            root := uia.GetCurrentDocumentElement()
        catch
            root := uia.BrowserElement
        if (!root)
            return
        br := root.BoundingRectangle
        if (!br || (br.r <= br.l) || (br.b <= br.t))
            return
        x := br.l + (br.r - br.l) * 0.15
        y := br.t + (br.b - br.t) * 0.20
        prevMode := A_CoordModeMouse
        CoordMode("Mouse", "Screen")
        MouseClick("Right", x, y)
        SendEscape()
        CoordMode("Mouse", prevMode)
        g_ML_ReceptivityHwnd := hwnd
    } catch {
    }
}

IsMercadoLivreActive() {
    global g_ML_CacheHwnd, g_ML_CacheResult
    if !WinActive("ahk_exe chrome.exe")
        return false
    hwnd := WinExist("A")
    if (!hwnd)
        return false
    ; Cache hit: same window as last check (avoids UIA on every keystroke / cheat sheet open)
    if (hwnd = g_ML_CacheHwnd && WinExist("ahk_id " g_ML_CacheHwnd))
        return g_ML_CacheResult
    ; Platform identification by URL only (do not use window title; it changes to product name). See shopping uia3.md.
    ; URL check via UIA (address bar: Chrome exposes AcceleratorKey "Ctrl+L", not AccessKey). Bounded to this window only.
    try {
        root := UIA.ElementFromHandle(hwnd)
        addressBar := root.FindFirst({ Type: 50004, AcceleratorKey: "Ctrl+L" })
        if (addressBar) {
            url := addressBar.Value
            if InStr(url, "mercadolivre.com") || InStr(url, "mercadolibre.com") {
                g_ML_CacheHwnd := hwnd
                g_ML_CacheResult := true
                return true
            }
        }
    } catch {
        ; UIA failed; do not cache so next call retries
    }
    g_ML_CacheHwnd := hwnd
    g_ML_CacheResult := false
    return false
}

; Mercado Livre UIA helpers: get document root and find/invoke elements (bounded, no global state).
; When UIA_Browser init fails (document not ready), fallback: get Document from foreground window tree (Type 50030 = Document).
ML_GetDocRoot() {
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        try
            return uia.GetCurrentDocumentElement()
        catch
            return uia.BrowserElement
    } catch {
        ; Fallback: document not ready for UIA_Browser init (e.g. on first load). Get Document from active window.
        try {
            hwnd := WinExist("A")
            if (hwnd && WinActive("ahk_exe chrome.exe")) {
                root := UIA.ElementFromHandle(hwnd)
                doc := root.FindFirst({ Type: 50030 })
                if (doc)
                    return doc
            }
        } catch {
        }
        return 0
    }
}

; #region agent log - debug helper
ML_DebugLog(hypothesisId, message, detail, runId := "pre-fix") {
    ; Intentionally no-op.
    return
}
; #endregion agent log

; Locate the Mercado Livre Preço filter container, scoped to the block that
; contains the predefined price links (Até R$ 40 / R$ 40 a R$ 95 / Mais de R$ 95)
; and the Mínimo/Máximo edits. Returns the container element or 0.
ML_GetPriceFilter(root) {
    if !root
        return 0
    priceLinkNames := ["Até R$ 40", "R$ 40 a R$ 95", "Mais de R$ 95"]
    for _, name in priceLinkNames {
        link := ML_Find(root, { Type: 50005, Name: name, cs: false })
        if !link
            continue
        parent := link
        ; Walk up a few levels to find a container that has the Mínimo edit as a descendant.
        loop 6 {
            try parent := parent.Parent
            catch
                break
            if !parent
                break
            minimo := ML_Find(parent, { Type: 50004, Name: "Mínimo", cs: false })
            if (minimo) {
                return parent
            }
        }
    }
    return 0
}

; Best-effort helper to set text of an edit element. Tries ValuePattern first,
; then falls back to focus + clear + send keys.
ML_SetEditText(el, text) {
    if (!el)
        return false
    if (text = "")
        return false
    ok := false
    try {
        if el.GetPropertyValue(UIA.Property.IsValuePatternAvailable) {
            vp := el.ValuePattern
            vp.SetValue(text)
            ok := true
        }
    } catch {
        ML_DebugLog("C", "ValuePattern attempt failed", "exception in SetValue", "run1")
    }
    if (!ok) {
        try el.ScrollIntoView()
        catch {
        }
        try {
            el.SetFocus()
        } catch {
            try el.Click()
            catch {
            }
        }
        Sleep 50
        ML_DebugLog("C", "Focus acquired for edit", text, "run1")
        Send "^a{Del}"
        Sleep 30
        SendText(text)
        ok := true
    }
    return ok
}

; Try conditions in order; invoke or click first match. Returns true if invoked/clicked, false otherwise.
ML_FindAndInvoke(conditionList) {
    root := ML_GetDocRoot()
    if (!root)
        return false
    for cond in conditionList {
        try {
            el := root.FindElement(cond, UIA.TreeScope.Descendants)
            if (el) {
                try el.Invoke()
                catch {
                    try el.Click()
                    catch
                        return false
                }
                return true
            }
        } catch
            continue
    }
    return false
}

; Find single element by condition (Descendants). Returns element or 0.
ML_Find(root, condition) {
    try
        return root.FindElement(condition, UIA.TreeScope.Descendants)
    catch
        return 0
}

;-------------------------------------------------------------------
; Shopee (Brazil) detection and UIA helpers
;-------------------------------------------------------------------
; Cache for IsShopeeActive (same pattern as IsMercadoLivreActive).
global g_Shopee_CacheHwnd := 0
global g_Shopee_CacheResult := false

IsShopeeActive() {
    global g_Shopee_CacheHwnd, g_Shopee_CacheResult
    if !WinActive("ahk_exe chrome.exe")
        return false
    hwnd := WinExist("A")
    if (!hwnd)
        return false
    ; Cache hit: same window as last check (avoids UIA on every keystroke / cheat sheet open)
    if (hwnd = g_Shopee_CacheHwnd && WinExist("ahk_id " g_Shopee_CacheHwnd))
        return g_Shopee_CacheResult
    ; Platform identification by URL only (do not use window title; it changes to product name). See shopping uia3.md.
    ; URL check via UIA (Chrome address bar: AcceleratorKey "Ctrl+L", not AccessKey)
    try {
        root := UIA.ElementFromHandle(hwnd)
        addressBar := root.FindFirst({ Type: 50004, AcceleratorKey: "Ctrl+L" })
        if (addressBar) {
            url := addressBar.Value
            if InStr(url, "shopee.com", false) {
                g_Shopee_CacheHwnd := hwnd
                g_Shopee_CacheResult := true
                return true
            }
        }
    } catch {
        ; UIA failed; do not cache so next call retries
    }
    g_Shopee_CacheHwnd := hwnd
    g_Shopee_CacheResult := false
    return false
}

Shopee_GetDocRoot() {
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        try
            return uia.GetCurrentDocumentElement()
        catch
            return uia.BrowserElement
    } catch
        return 0
}

Shopee_Find(root, condition) {
    try
        return root.FindElement(condition, UIA.TreeScope.Descendants)
    catch
        return 0
}

Shopee_FindAndInvoke(conditionList) {
    root := Shopee_GetDocRoot()
    if (!root)
        return false
    for cond in conditionList {
        try {
            el := root.FindElement(cond, UIA.TreeScope.Descendants)
            if (el) {
                try el.Invoke()
                catch {
                    try el.Click()
                    catch
                        return false
                }
                return true
            }
        } catch
            continue
    }
    return false
}

; Navigate Shopee search results pagination by relative offset (e.g. +1 next, -1 previous).
Shopee_NavMove(offset) {
    if (offset = 0)
        return false
    hwnd := WinExist("A")
    if (!hwnd)
        return false

    ; Get current page index from Chrome address bar (?page=N), defaulting to 0 when absent.
    currentPage := 0
    try {
        rootWin := UIA.ElementFromHandle(hwnd)
        addressBar := rootWin.FindFirst({ Type: 50004, AcceleratorKey: "Ctrl+L" })
        if (addressBar) {
            url := addressBar.Value
            m := ""
            if RegExMatch(url, "i)[?&]page=(\d+)", &m) {
                try currentPage := Integer(m[1])
            }
        }
    } catch {
    }

    root := Shopee_GetDocRoot()
    if (!root)
        return false

    navGroup := Shopee_Find(root, { Type: 50026, Name: "Navegação entre páginas", cs: false })
    if (!navGroup)
        return false

    links := 0
    try links := navGroup.FindAll({ Type: 50005 })
    catch {
        return false
    }
    if (!links)
        return false

    bestEl := 0
    bestDelta := 0x7FFFFFFF

    for link in links {
        value := ""
        try value := link.Value
        catch {
            continue
        }
        if (value = "")
            continue
        m2 := ""
        if !RegExMatch(value, "i)[?&]page=(\d+)", &m2)
            continue
        targetPage := 0
        try targetPage := Integer(m2[1])
        catch {
            continue
        }

        ; Choose the nearest page ahead (offset>0) or behind (offset<0) relative to currentPage.
        if (offset > 0 && targetPage > currentPage) {
            delta := targetPage - currentPage
            if (delta < bestDelta) {
                bestDelta := delta
                bestEl := link
            }
        } else if (offset < 0 && targetPage < currentPage) {
            delta := currentPage - targetPage
            if (delta < bestDelta) {
                bestDelta := delta
                bestEl := link
            }
        }
    }

    if (!bestEl)
        return false

    try {
        bestEl.Invoke()
        return true
    } catch {
        try {
            bestEl.Click()
            return true
        } catch {
            return false
        }
    }
}

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
; [SK module] Command Palette hotkeys -> Shift keys\hotif_command_palette.ahk
#include %A_ScriptDir%\Shift keys\hotif_command_palette.ahk

; [SK module] ChatGPT loading banner and wait helpers -> Shift keys\chatgpt_loading_helpers.ahk
#include %A_ScriptDir%\Shift keys\chatgpt_loading_helpers.ahk

; [SK module] TEMPORARY M365 Copilot auto-continue (^!#n) -> Shift keys\m365_copilot_temp.ahk
#include %A_ScriptDir%\Shift keys\m365_copilot_temp.ahk

; VS Code evidence -> PDF search loop (^!#o) — see VSCodeEvidenceSearch.ahk
global EVIDENCE_SEARCH_FROM_SHIFT_KEYS := true
#include %A_ScriptDir%\VSCodeEvidenceSearch.ahk
#InputLevel 10
EvidenceSearch_BindHotkey()