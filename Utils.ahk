#Requires AutoHotkey v2.0+
#SingleInstance Force
#include %A_ScriptDir%\env.ahk

; #region agent log
DebugBannerLog(location, message, dataStr := "", hypothesisId := "") {
    ; Intentionally no-op.
    ; These agent debug logs are not required for runtime behavior.
    return
}

; Dictation-Gemini flow debug (session 7e3dd7): NDJSON log (disabled)
DebugFlowLog(location, message, dataStr := "", hypothesisId := "") {
    ; Intentionally no-op.
    ; These agent debug logs are not required for runtime behavior.
    return
}
; #endregion

#include UIA-v2\Lib\UIA.ahk
#include UIA-v2\Lib\UIA_Browser.ahk
#include %A_ScriptDir%\lib\Media.ahk

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

; Find the Gemini prompt field via UIA (returns element or 0). Supports EN and PT labels. Used by Gemini.ahk and Utils.ahk.
FindGeminiPromptField(uia) {
    promptField := 0
    for name in GEMINI_PROMPT_FIELD_NAMES {
        try {
            promptField := uia.FindFirst({ Name: name, Type: 50004 })
            if (promptField) {
                return promptField
            }
        } catch
            continue
    }
    try {
        promptField := uia.FindFirst({ Type: "Edit", Name: GEMINI_PROMPT_FIELD_NAMES[1] })
        if (promptField)
            return promptField
    } catch {
    }
    try {
        allEdits := uia.FindAll({ Type: 50004 })
        for edit in allEdits {
            if (InStr(edit.ClassName, "ql-editor") || InStr(edit.ClassName, "new-input-ui")) {
                for name in GEMINI_PROMPT_FIELD_NAMES {
                    if InStr(edit.Name, name) || InStr(edit.Name, "prompt") {
                        return edit
                    }
                }
            }
        }
    } catch {
    }
    try {
        allEdits := uia.FindAll({ Type: 50004 })
        for edit in allEdits {
            if InStr(edit.ClassName, "ql-editor") {
                return edit
            }
        }
    } catch {
    }
    try {
        allEdits := uia.FindAll({ Type: 50004 })
        for edit in allEdits {
            if InStr(edit.ClassName, "ql-editor") {
                for name in GEMINI_PROMPT_FIELD_NAMES {
                    if InStr(edit.Name, name) || InStr(edit.Name, "prompt")
                        return edit
                }
            }
        }
    } catch {
    }
    return 0
}

; Move keyboard focus to the main Ask Gemini field when that Chrome window is already foreground.
; Does not WinActivate (avoids stealing focus from another app). Optional chime matches Gemini.ahk open-hotkey UX.
Utils_PlayGeminiFocusedChime(minIntervalMs := 400) {
    static lastChimeTick := 0
    if (!IsSoundEnabled())
        return false
    now := A_TickCount
    if (lastChimeTick && (now - lastChimeTick) < minIntervalMs)
        return false
    lastChimeTick := now
    try SoundPlay(A_ScriptDir . "\sounds\gemini-focused.wav")
    return true
}

FocusGeminiAskFieldForHwnd(geminiHwnd, playChime := false) {
    if (!geminiHwnd)
        return false
    try {
        if (!WinActive("ahk_id " geminiHwnd))
            return false
    } catch {
        return false
    }
    try {
        uia := UIA_Browser("ahk_id " geminiHwnd)
        Sleep 120
        promptField := FindGeminiPromptField(uia)
        if (!promptField)
            return false
        try {
            if (promptField.HasKeyboardFocus) {
                if (playChime)
                    Utils_PlayGeminiFocusedChime()
                return true
            }
        } catch {
        }
        try promptField.SetFocus()
        Sleep 100
        if (!promptField.HasKeyboardFocus) {
            try promptField.Click()
            Sleep 80
        }
        if (promptField.HasKeyboardFocus) {
            if (playChime)
                Utils_PlayGeminiFocusedChime()
            return true
        }
    } catch {
    }
    return false
}

; 1-based active Chrome tab index via UIA (tab bar). Returns { index, count } or 0. Shared with Gemini.ahk.
GetChromeActiveTabIndex(uia) {
    try {
        uia.GetCurrentMainPaneElement()
        tabs := uia.GetTabs()
        if (!tabs.Length)
            return 0
        current := uia.GetTab("")
        if (!current)
            return 0
        rid := current.RuntimeId
        for i, tab in tabs {
            try {
                if (tab.RuntimeId = rid)
                    return { index: i, count: tabs.Length }
            } catch {
                continue
            }
        }
    } catch {
    }
    return 0
}

; --- Gemini mode picker (Fast / Thinking / Pro), UIA ---------------------------------

FindGeminiModePickerButton(uia) {
    if !IsObject(uia)
        return 0
    try {
        b := uia.FindFirst({ Name: "Open mode picker", Type: 50000 })
        if (b)
            return b
    } catch {
    }
    try {
        b := uia.FindFirst({ Type: "Button", Name: "Open mode picker" })
        if (b)
            return b
    } catch {
    }
    try {
        all := uia.FindAll({ Type: 50000 })
        for btn in all {
            try {
                if InStr(btn.ClassName, "input-area-switch")
                    return btn
            } catch {
            }
        }
    } catch {
    }
    return 0
}

GeminiNormalizeModelName(name) {
    if (name = "")
        return ""
    nl := StrLower(Trim(name))
    if (nl = "fast")
        return "Fast"
    if (nl = "thinking")
        return "Thinking"
    if (nl = "pro")
        return "Pro"
    return ""
}

; Gemini 3 mode menu uses composite Accessible names, e.g. "Fast Answers quickly" (see gemini-tree-model-menu-open.md).
GeminiNormalizeModelLabel(name) {
    if (name = "")
        return ""
    n := GeminiNormalizeModelName(name)
    if (n != "")
        return n
    if RegExMatch(name, "i)^Fast(\s|$)")
        return "Fast"
    if RegExMatch(name, "i)^Thinking(\s|$)")
        return "Thinking"
    if RegExMatch(name, "i)^Pro(\s|$)")
        return "Pro"
    return ""
}

GetGeminiActiveModelFromPickerOnly(uia) {
    picker := FindGeminiModePickerButton(uia)
    if !picker
        return ""
    texts := []
    try {
        texts := picker.FindAll({ Type: 50020 })
    } catch {
    }
    if (!IsObject(texts) || texts.Length = 0) {
        try {
            texts := picker.FindAll({ Type: "Text" })
        } catch {
            texts := []
        }
    }
    for t in texts {
        try {
            tn := t.Name
        } catch {
            continue
        }
        norm := GeminiNormalizeModelLabel(tn)
        if (norm != "") {
            return norm
        }
    }
    return ""
}

GeminiCollectModelOptionButtons(uia) {
    modelPattern := "i)^(Fast|Thinking|Pro)$"
    modelButtons := []
    ; Gemini 3: Material menu uses MenuItem (50011), Name like "Fast Answers quickly" (gemini-tree-model-menu-open.md).
    try {
        menuItems := uia.FindAll({ Type: 50011 })
    } catch {
        menuItems := []
    }
    if (!IsObject(menuItems) || menuItems.Length = 0) {
        try {
            menuItems := uia.FindAll({ Type: "MenuItem" })
        } catch {
            menuItems := []
        }
    }
    for mi in menuItems {
        try {
            fullName := mi.Name
            shortName := GeminiNormalizeModelLabel(fullName)
            if (shortName = "")
                continue
            className := ""
            try {
                className := mi.ClassName
            } catch {
                className := ""
            }
            if (!InStr(className, "bard-mode-list-button"))
                continue
            isDisabled := false
            try {
                if (InStr(className, "disabled") || InStr(className, "mat-mdc-button-disabled"))
                    isDisabled := true
                try {
                    if (!mi.GetPropertyValue(UIA.Property.IsEnabled))
                        isDisabled := true
                } catch {
                }
            } catch {
            }
            isSelected := false
            try {
                if (mi.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable))
                    isSelected := mi.SelectionItemPattern.IsSelected
            } catch {
            }
            if (!isSelected) {
                try {
                    if (InStr(className, "is-selected") || InStr(className, "selected") || InStr(className, "active") ||
                    InStr(className, "mdc-selected"))
                        isSelected := true
                } catch {
                }
            }
            modelButtons.Push({ btn: mi, name: shortName, isSelected: isSelected, isDisabled: isDisabled,
                className: className })
        } catch {
        }
    }
    if (modelButtons.Length > 0)
        return modelButtons
    try {
        allButtons := uia.FindAll({ Type: "Button" })
    } catch {
        return modelButtons
    }
    for btnCandidate in allButtons {
        try {
            btnName := btnCandidate.Name
            if (!RegExMatch(btnName, modelPattern))
                continue
            className := ""
            try {
                className := btnCandidate.ClassName
            } catch {
                className := ""
            }
            isDisabled := false
            try {
                if (InStr(className, "disabled") || InStr(className, "mat-mdc-button-disabled"))
                    isDisabled := true
                try {
                    if (!btnCandidate.GetPropertyValue(UIA.Property.IsEnabled))
                        isDisabled := true
                } catch {
                }
            } catch {
            }
            isSelected := false
            try {
                if (btnCandidate.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable))
                    isSelected := btnCandidate.SelectionItemPattern.IsSelected
            } catch {
            }
            if (!isSelected) {
                try {
                    if (InStr(className, "selected") || InStr(className, "active") || InStr(className, "mdc-selected"))
                        isSelected := true
                } catch {
                }
            }
            modelButtons.Push({ btn: btnCandidate, name: btnName, isSelected: isSelected, isDisabled: isDisabled,
                className: className })
        } catch {
        }
    }
    return modelButtons
}

GeminiInvokeModelButton(btn) {
    if !IsObject(btn)
        return false
    clicked := false
    try {
        btn.SetFocus()
        Sleep 40
    } catch {
    }
    supportsInvoke := false
    try {
        supportsInvoke := btn.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)
    } catch {
    }
    if (supportsInvoke) {
        try {
            btn.Invoke()
            clicked := true
        } catch {
        }
    }
    if (!clicked) {
        try {
            btn.Click()
            clicked := true
        } catch {
        }
    }
    return clicked
}

EnsureGeminiModelViaMenu(expected) {
    exp := GeminiNormalizeModelLabel(expected)
    if (exp = "")
        return false
    try {
        uia := UIA_Browser()
    } catch {
        return false
    }
    if !IsObject(uia)
        return false
    if (GetGeminiActiveModelFromPickerOnly(uia) = exp)
        return true
    picker := FindGeminiModePickerButton(uia)
    if !picker
        return false
    try {
        picker.Click()
    } catch {
        try {
            if (picker.GetPropertyValue(UIA.Property.IsInvokePatternAvailable))
                picker.Invoke()
        } catch {
        }
    }
    Sleep 250
    try {
        uia := UIA_Browser()
    } catch {
        Send "{Escape}"
        return false
    }
    if !IsObject(uia) {
        Send "{Escape}"
        return false
    }
    modelButtons := GeminiCollectModelOptionButtons(uia)
    targetBtn := 0
    for modelBtn in modelButtons {
        if (GeminiNormalizeModelLabel(modelBtn.name) = exp && !modelBtn.isDisabled) {
            targetBtn := modelBtn.btn
            break
        }
    }
    if !targetBtn {
        Send "{Escape}"
        return false
    }
    if !GeminiInvokeModelButton(targetBtn)
        return false
    Sleep 100
    Send "{Escape}"
    return true
}

; -----------------------------------------------------------------------------
; This script consolidates various utility hotkeys.
; -----------------------------------------------------------------------------

; =============================================================================
; Hotstring Core Functions
; =============================================================================

; ----------------------
; Safer hotstrings core
; ----------------------
global g_hotstrings := []
global g_lastExpansion := 0

; Cross-process IPC for Hotstring Selector (symmetric with WindowManagement project selector IPC)
global g_HS_SelectorOpenFile := A_ScriptDir "\.cursor\hs_selector_open"
global g_HS_SelectorCloseRequestFile := A_ScriptDir "\.cursor\hs_selector_close_request"
global g_HS_SelectorCloseCheckTimer := ""

Utils_CheckHotstringSelectorCloseRequest() {
    global g_HotstringSelectorActive, g_HS_SelectorCloseRequestFile
    if (!g_HotstringSelectorActive)
        return
    if (FileExist(g_HS_SelectorCloseRequestFile)) {
        try FileDelete(g_HS_SelectorCloseRequestFile)
        catch {
        }
        CleanupHotstringSelector()
    }
}

; Trigger only on Space or Tab, not Enter or punctuation
Hotstring("EndChars", " `t")

RegisterHotstring(trigger, expansion, category := "", title := "", char := "") {
    global g_hotstrings
    g_hotstrings.Push({ trigger: trigger, expansion: expansion, category: category, title: title, char: char })
}

GetHotstringsCheatSheetText() {
    global g_hotstrings
    if (!IsSet(g_hotstrings) || g_hotstrings.Length = 0)
        return ""
    txt := ""
    for hs in g_hotstrings {
        line := "[" hs.trigger "] > " hs.expansion
        if (txt = "")
            txt := line
        else
            txt := txt . "`n" . line
    }
    return txt
}

; Safe paste insertion to avoid app shortcuts and re-triggers
InsertText(text) {
    global g_lastExpansion
    ; Debounce to prevent rapid duplicate expansions (e.g., double Space)
    if (A_TickCount - g_lastExpansion) < 250
        return
    g_lastExpansion := A_TickCount

    saved := ClipboardAll()
    try {
        A_Clipboard := text
        ClipWait(0.3)
        Sleep 50  ; Give time for clipboard to fully update
        Send "^v"
    } finally {
        Sleep 150  ; Wait longer for paste to complete before restoring clipboard
        A_Clipboard := saved
    }
}

; ----------------------
; Define hotstrings below
; (same triggers, now using InsertText)
; ----------------------

:o:myl::
{
    InsertText("my links")
}

:o:gintegra::
{
    InsertText("GS_UX core team_UX and CIP Integration")
}

:o:gdash::
{
    InsertText("GS_E&S_CIP Dashboard research and design")
}

:o:opex-cim-journey-mapping::
{
    InsertText("opex-cim-journey-mapping")
}

:o:gb2c::
{
    InsertText("GS_B2C_Credit_Management_Strategy_UI_Mentoring")
}

:o:gug::
{
    InsertText("GS_UX Core Team_Monitoring for B2C in Brazil")
}

:o:gpm::
{
    InsertText("project management LA")
}

:o:guxcip::
{
    InsertText("UX and CIP")
}

:o:gtrain::
{
    InsertText("GS_UX core team_Trainings Management")
}

:o:gbp::
{
    InsertText("GS_B2C_Portals and Key Accounts Process POC")
}

GetPromptText(key) {
    promptFile := A_ScriptDir "\prompt\" key ".txt"
    try {
        return FileRead(promptFile)
    } catch {
        return "[PROMPT FILE MISSING: " key "]"
    }
}

:o:cgrammar::
{
    InsertText(GetPromptText("grammar"))
}

:o:ebosch::
{
    InsertText("eduardo.figueiredo@br.bosch.com")
}

:o:egoogle::
{
    InsertText("edu.evangelista.figueiredo@gmail.com")
}

:o:mtask::
{
    InsertText(GetPromptText("mtask"))
}

:o:flog::
{
    InsertText(GetPromptText("flog"))
}

:o:aiopt::
{
    InsertText(GetPromptText("aiopt"))
}

:o:cplant::
{
    InsertText(GetPromptText("cplant"))
}

:o:aibrapid::
{
    InsertText(GetPromptText("aib-rapid-fire-template"))
}

; ----------------------
; Register hotstrings for cheat sheet display
; ----------------------
InitHotstringsCheatSheet() {
    promptDir := A_ScriptDir "\prompt"

    ; Prompts (4 items) - First category
    try {
        RegisterHotstring(":o:cgrammar", FileRead(promptDir "\grammar.txt"), "Prompts",
        "✏️ Grammar & Spelling Corrector")
    } catch {
        RegisterHotstring(":o:cgrammar",
            "Correct grammar, spelling, punctuation, and casing. Give back only the text.`n", "Prompts",
            "✏️ Grammar & Spelling Corrector")
    }
    try {
        RegisterHotstring(":o:mtask", FileRead(promptDir "\mtask.txt"), "Prompts", "🔲 Convert to Task")
    } catch {
        RegisterHotstring(":o:mtask", "Translate this into a task. Output ONLY the task. Start with 🔲.`n", "Prompts",
            "🔲 Convert to Task")
    }
    try {
        RegisterHotstring(":o:flog", FileRead(promptDir "\flog.txt"), "Prompts", "🍽️ Food Log Dictation")
    } catch {
        RegisterHotstring(":o:flog", "Food_Log dictation → Excel CSV. Output ONLY the final CSV block.`n", "Prompts",
            "🍽️ Food Log Dictation")
    }
    try {
        RegisterHotstring(":o:aiopt", FileRead(promptDir "\aiopt.txt"), "Prompts", "🤖 AI Text Optimizer")
    } catch {
        RegisterHotstring(":o:aiopt",
            "Rewrite the input text so it becomes AI-oriented. Preserve all important information.`n", "Prompts",
            "🤖 AI Text Optimizer")
    }

    ; Placeholder prompts (6 slots reserved for future prompts)
    ; Technical Architect & Code Planner (content from prompt/markdown-plan.txt)
    try {
        planPrompt := FileRead(promptDir "\markdown-plan.txt")
        RegisterHotstring(":o:cplan", planPrompt, "Prompts", "📋 Technical Architect & Code Planner")
    } catch {
        RegisterHotstring(":o:cplan",
            "You are an expert Technical Architect and Code Planner. Generate a .plan.md file. **Input Task:**`n",
            "Prompts", "📋 Technical Architect & Code Planner")
    }
    try {
        RegisterHotstring(":o:cplant", FileRead(promptDir "\cplant.txt"), "Prompts", "📝 Plan File Template")
    } catch {
        RegisterHotstring(":o:cplant",
            "---`nname: [Title]`noverview: [Summary]`ntodos:`n  - id: x`n    content: [Step]`n    status: pending`n---`n",
            "Prompts", "📝 Plan File Template")
    }
    ; Character w: creating mnemonic stories (content from notes/studies/technique/story-prompt.txt)
    ; Supports personal (Google Drive) and work (OneDrive - Bosch Group) environments
    mnemonicPromptPath := ""
    mnemonicWorkPath := "C:\Users\fie7ca\OneDrive - Bosch Group\14-my-notes\studies\technique\prompts\story-prompt.txt"
    mnemonicPersonalPath := "C:\Users\eduev\Meu Drive\17 - Projects\notes\studies\technique\prompts\story-prompt.txt"
    if FileExist(mnemonicWorkPath)
        mnemonicPromptPath := mnemonicWorkPath
    else if FileExist(mnemonicPersonalPath)
        mnemonicPromptPath := mnemonicPersonalPath
    else
        mnemonicPromptPath := mnemonicWorkPath
    try {
        mnemonicPrompt := FileRead(mnemonicPromptPath)
        RegisterHotstring(":o:mnemonic", mnemonicPrompt, "Prompts", "📖 Creating mnemonic stories")
    } catch {
        RegisterHotstring("", "", "Prompts", "Reserved 3")
    }
    ; Character e: transcript YouTube video (content from notes/studies/technique/video-transcription-prompt.txt)
    ; Supports personal (Google Drive) and work (OneDrive - Bosch Group) environments
    transcriptPromptPath := ""
    transcriptWorkPath :=
        "C:\Users\fie7ca\OneDrive - Bosch Group\14-my-notes\studies\technique\prompts\video-transcription-prompt.txt"
    transcriptPersonalPath :=
        "C:\Users\eduev\Meu Drive\17 - Projects\notes\studies\technique\prompts\video-transcription-prompt.txt"
    if FileExist(transcriptWorkPath)
        transcriptPromptPath := transcriptWorkPath
    else if FileExist(transcriptPersonalPath)
        transcriptPromptPath := transcriptPersonalPath
    else
        transcriptPromptPath := transcriptWorkPath
    try {
        transcriptPrompt := FileRead(transcriptPromptPath)
        RegisterHotstring(":o:ytranscript", transcriptPrompt, "Prompts", "🎬 Transcript Youtube Video")
    } catch {
        RegisterHotstring("", "", "Prompts", "Reserved 4")
    }
    ; Character r: read aloud this story (content from notes/studies/technique/prompts/read-aloud-prompt.txt)
    ; Supports personal (Google Drive) and work (OneDrive - Bosch Group) environments
    readAloudPromptPath := ""
    readAloudWorkPath :=
        "C:\Users\fie7ca\OneDrive - Bosch Group\14-my-notes\studies\technique\prompts\read-aloud-prompt.txt"
    readAloudPersonalPath :=
        "C:\Users\eduev\Meu Drive\17 - Projects\notes\studies\technique\prompts\read-aloud-prompt.txt"
    if FileExist(readAloudWorkPath)
        readAloudPromptPath := readAloudWorkPath
    else if FileExist(readAloudPersonalPath)
        readAloudPromptPath := readAloudPersonalPath
    else
        readAloudPromptPath := readAloudWorkPath
    try {
        readAloudPrompt := FileRead(readAloudPromptPath)
        RegisterHotstring(":o:readaloud", readAloudPrompt, "Prompts", "📖 read aloud this story", "r")
    } catch {
        RegisterHotstring("", "", "Prompts", "Reserved 5")
    }
    ; Character t: story revision (content from notes/studies/technique/prompts/revision-prompt.txt)
    ; Supports personal (Google Drive) and work (OneDrive - Bosch Group) environments
    revisionPromptPath := ""
    revisionWorkPath :=
        "C:\Users\fie7ca\OneDrive - Bosch Group\14-my-notes\studies\technique\prompts\revision-prompt.txt"
    revisionPersonalPath :=
        "C:\Users\eduev\Meu Drive\17 - Projects\notes\studies\technique\prompts\revision-prompt.txt"
    if FileExist(revisionWorkPath)
        revisionPromptPath := revisionWorkPath
    else if FileExist(revisionPersonalPath)
        revisionPromptPath := revisionPersonalPath
    else
        revisionPromptPath := revisionWorkPath
    try {
        revisionPrompt := FileRead(revisionPromptPath)
        RegisterHotstring(":o:revision", revisionPrompt, "Prompts", "📝 Story revision", "t")
    } catch {
        RegisterHotstring("", "", "Prompts", "Reserved 6")
    }
    ; Character a: story reduction (content from notes/studies/technique/prompts/story-reduction-prompt.txt)
    ; Supports personal (Google Drive) and work (OneDrive - Bosch Group) environments
    storyReductionPromptPath := ""
    storyReductionWorkPath :=
        "C:\Users\fie7ca\OneDrive - Bosch Group\14-my-notes\studies\technique\prompts\story-reduction-prompt.txt"
    storyReductionPersonalPath :=
        "C:\Users\eduev\Meu Drive\17 - Projects\notes\studies\technique\prompts\story-reduction-prompt.txt"
    if FileExist(storyReductionWorkPath)
        storyReductionPromptPath := storyReductionWorkPath
    else if FileExist(storyReductionPersonalPath)
        storyReductionPromptPath := storyReductionPersonalPath
    else
        storyReductionPromptPath := storyReductionWorkPath
    try {
        storyReductionPrompt := FileRead(storyReductionPromptPath)
        RegisterHotstring(":o:storyreduction", storyReductionPrompt, "Prompts", "📝 Story reduction", "a")
    } catch {
        RegisterHotstring("", "", "Prompts", "Reserved 7")
    }
    try {
        aibRapidFireTpl := FileRead(promptDir "\aib-rapid-fire-template.txt")
        RegisterHotstring(":o:aibrapid", aibRapidFireTpl, "Prompts", "📜 Junior AI: ⚡ rapid-fire template")
    } catch {
        RegisterHotstring(":o:aibrapid",
            "Junior AI (AIB): planning doc with ⚡ — conceptual above, execution steps below.`n", "Prompts",
            "📜 Junior AI: ⚡ rapid-fire template")
    }

    ; Hotstrings: emails
    RegisterHotstring(":o:ebosch", "eduardo.figueiredo@br.bosch.com", "Hotstrings", "💼 Bosch Email")
    RegisterHotstring(":o:egoogle", "edu.evangelista.figueiredo@gmail.com", "Hotstrings", "📧 Gmail")

    ; Projects (Cursor workspaces) — keys align with Project Selector 2
    RegisterHotstring(":o:gintegra", "GS_UX core team_UX and CIP Integration", "Projects", "🔄 UX and CIP Integration",
        "u")
    RegisterHotstring(":o:gdash", "GS_E&S_CIP Dashboard research and design", "Projects", "📊 CIP Dashboard", "d")
    RegisterHotstring(":o:boiler-plate", "boiler-plate", "Projects", "🧱 boiler-plate", "0")
    RegisterHotstring(":o:astra", "astra", "Projects", "⭐ astrA", "a")
    RegisterHotstring(":o:opex-cim-journey-mapping", "opex-cim-journey-mapping", "Projects",
        "E&S Opex CIM Journey Mapping",
        "o")

    ; Hotstrings (non-workspace “project-like” names)
    RegisterHotstring(":o:myl", "my links", "Hotstrings", "🔗 my links", "m")
    RegisterHotstring(":o:gpm", "project management LA", "Hotstrings", "📋 project management LA", "p")
    RegisterHotstring(":o:guxcip", "UX and CIP", "Hotstrings", "🔗 UX and CIP", "x")
    RegisterHotstring(":o:gtrain", "GS_UX core team_Trainings Management", "Hotstrings", "🎓 Trainings Management", "t"
    )
}
InitHotstringsCheatSheet()

; =============================================================================
; Files & Links System
; =============================================================================

; Global variables for quick open files
global g_QuickOpenFiles := []
global g_QuickOpenFileCharMap := Map()  ; Maps character to file path

; Register a file for quick opening
RegisterQuickOpenFile(filePath, title) {
    global g_QuickOpenFiles
    g_QuickOpenFiles.Push({ filePath: filePath, title: title, category: "Files & Links" })
}

; Initialize quick open files
InitQuickOpenFiles() {
    ; Register dissertation Power BI file with character 'y'
    RegisterQuickOpenFile(
        "C:\Users\eduev\Meu Drive\04 - Pós-graduação\01 - Mestrado\26-ai-experiment\infoVis\Dissertation InfoVis  - PowerBI - Charts.pbix",
        "📊 Dissertation InfoVis"
    )

    ; Register radio-tiso exercises YouTube link
    RegisterQuickOpenFile(
        "https://www.youtube.com/watch?v=I6ZRH9Mraqw&t=2s",
        "📻 Radio-Tiso Exercises"
    )

    ; Register GS_UX core team_UX and CIP Integration Miro
    RegisterQuickOpenFile(
        "https://miro.com/app/board/uXjVJdbNFkA=/",
        "🎨 GS_UX core team_UX and CIP Integration Miro"
    )

    ; Register GS_E&S_CIP Dashboard research and design Miro
    RegisterQuickOpenFile(
        "https://miro.com/app/board/uXjVJVZSXvk=/",
        "📊 GS_E&S_CIP Dashboard research and design Miro"
    )
}
InitQuickOpenFiles()

; =============================================================================
; Macros System
; =============================================================================

; Global variables for macros
global g_Macros := []
global g_MacroCharMap := Map()  ; Maps character to macro function
global g_ProgrammaticDictationStop := false  ; Skip ~#!+0 when script sends #!+0 programmatically
global g_GeminiToggleTab := 1  ; Last Gemini tab chosen by ^!#4 (UIA-synced); other code may still assume 1/2 toggle state

; Register a macro
RegisterMacro(func, title, char := "") {
    global g_Macros
    g_Macros.Push({ func: func, title: title, category: "Macros", char: char })
}

; Get scripts directory path based on environment
GetScriptsDirectory() {
    global IS_WORK_ENVIRONMENT
    if (IS_WORK_ENVIRONMENT) {
        return "C:\Users\fie7ca\Documents\scripts"
    } else {
        return "C:\Users\eduev\Meu Drive\17 - Projects\scripts"
    }
}

; Get list of script files to update.
; Utils.ahk must be last so QuickUpdateScripts can reload all other scripts before this instance is replaced.
GetScriptFiles() {
    scriptsDir := GetScriptsDirectory()
    return [
        scriptsDir "\WindowManagement.ahk",
        scriptsDir "\Spotify.ahk",
        scriptsDir "\Shift keys.ahk",
        scriptsDir "\Outlook.ahk",
        scriptsDir "\Microsoft Teams.ahk",
        scriptsDir "\Gemini.ahk",
        scriptsDir "\AppLaunchers.ahk",
        scriptsDir "\Utils.ahk"
    ]
}

; Quick Update Scripts macro: PowerShell handoff restart (local-only, no git).
QuickUpdateScripts() {
    static s_isQuickUpdateRunning := false
    if (s_isQuickUpdateRunning) {
        return
    }
    s_isQuickUpdateRunning := true

    try {
        scripts := GetScriptFiles()
        scriptsNames := ""
        for scriptPath in scripts {
            parts := StrSplit(scriptPath, "\")
            scriptsNames .= parts[parts.Length] . ";"
        }

        StandardLoadingBar_Show("⏳ Restarting scripts...", BANNER_ACCENT_INTERMEDIATE, { passive: false })

        utilsPath := ""
        psPaths := []

        for scriptPath in scripts {
            if (InStr(scriptPath, "\Utils.ahk"))
                utilsPath := scriptPath
            psPaths.Push("'" . StrReplace(scriptPath, "'", "''") . "'")
        }

        if (!utilsPath || utilsPath = "") {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ Utils.ahk not found in script list", 2500, BANNER_ACCENT_ERROR)
            return
        }

        psUtils := StrReplace(utilsPath, "'", "''")

        ; PowerShell contract (deterministic): sleep 500ms -> stop AutoHotkey* -> sleep 500ms -> Start-Process scripts.
        ps := ""
        ps .= "Start-Sleep -Milliseconds 500; "
        ps .= "Stop-Process -Name 'AutoHotkey*' -Force -ErrorAction SilentlyContinue; "
        ps .= "Start-Sleep -Milliseconds 500; "
        ps .= "$utilsPath = '" . psUtils . "'; "
        ; AHK Array.Join() isn't available in this environment; build a CSV manually.
        psPathsJoined := ""
        for i, p in psPaths {
            if (i > 1)
                psPathsJoined .= ","
            psPathsJoined .= p
        }
        ps .= "$scripts = @(" . psPathsJoined . "); "
        ps .= "foreach ($s in $scripts) { "
        ps .= "  if ($s -eq $utilsPath) { Start-Process -FilePath $s -ArgumentList '/Updated' | Out-Null } "
        ps .= "  else { Start-Process -FilePath $s | Out-Null } "
        ps .= "}"

        ; Execute asynchronously, then terminate this AHK instance immediately.
        pid := 0
        try {
            pid := Run('powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "' . ps . '"', , "Hide")
        } catch as e {
            throw
        }
        ExitApp
    } catch as e {
        try StandardLoadingBar_Hide(0)
        ShowCenteredOverlay_Utils("❌ QuickUpdateScripts error: " . e.Message, 3500, BANNER_ACCENT_ERROR)
    } finally {
        ; If we didn't ExitApp (error path), reset the guard.
        s_isQuickUpdateRunning := false
    }
}

; Update Gemini script specifically (local-only): restart Gemini.ahk.
UpdateGeminiScript() {
    scriptsDir := GetScriptsDirectory()
    geminiPath := scriptsDir "\Gemini.ahk"

    if (!FileExist(geminiPath)) {
        ShowCenteredOverlay_Utils("❌ Gemini.ahk not found at: " geminiPath, 3000, BANNER_ACCENT_ERROR)
        return
    }

    StandardLoadingBar_Show("⏳ Reloading Gemini...", BANNER_ACCENT_INTERMEDIATE, { passive: false })
    try {
        Run(geminiPath)
        StandardLoadingBar_Hide(0)
        ShowCenteredOverlay_Utils("✅ Gemini script updated!", 1500, BANNER_ACCENT_SUCCESS)
    } catch {
        StandardLoadingBar_Hide(0)
        ShowCenteredOverlay_Utils("❌ Failed to reload Gemini", 2000, BANNER_ACCENT_ERROR)
    }
}

; Add specific word to Handy macro function
AddWordToHandy() {
    targetPath := GetHandyShortcutPath()

    try {
        if WinExist("Handy ahk_class Tauri Window") {
            WinActivate
        } else {
            if (targetPath = "" || !FileExist(targetPath)) {
                MsgBox "Failed to launch Handy.`n`nShortcut not found.", "Utils.ahk", "IconX"
                return
            }
            Run targetPath
            if !WinWait("Handy ahk_class Tauri Window", , 5) {
                MsgBox "Failed to launch Handy."
                return
            }
        }

        if (!WinWaitActive("Handy ahk_class Tauri Window", , 2)) {
            ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        hwnd := WinExist("Handy ahk_class Tauri Window")

        ; Initialize UIA
        el := UIA.ElementFromHandle(hwnd)
        if !el
            return

        ; Locate and click "Advanced" - find Group element directly by Type and ClassName pattern
        ; The clickable target is the Group (50026) with ClassName containing "cursor-pointer" and "flex gap-2 items-center"
        advancedBtn := ""
        try {
            allGroups := el.FindAll({ Type: 50026 })
            if allGroups {
                for group in allGroups {
                    try {
                        groupClassName := group.ClassName
                        ; Look for Group with cursor-pointer and flex gap-2 items-center (Advanced button pattern)
                        if (InStr(groupClassName, "cursor-pointer") && InStr(groupClassName, "flex gap-2 items-center")) {
                            ; Verify it contains "Advanced" text by checking children
                            try {
                                advancedText := group.FindFirst({ Type: 50020, Name: "Advanced" })
                                if advancedText {
                                    advancedBtn := group
                                    break
                                }
                            } catch {
                            }
                        }
                    } catch {
                    }
                }
            }
        } catch {
        }

        ; Fallback: Find by Name "Advanced" and verify parent has correct ClassName
        if !advancedBtn {
            advancedElement := el.FindFirst({ Name: "Advanced" })
            if advancedElement {
                try {
                    parentGroup := advancedElement.GetParentElement()
                    ; Verify parent is a Group with cursor-pointer class (clickable)
                    if (parentGroup && parentGroup.Type = 50026 && InStr(parentGroup.ClassName, "cursor-pointer")) {
                        advancedBtn := parentGroup
                    }
                } catch {
                }
            }
        }

        if advancedBtn {
            try {
                advancedBtn.Click()
            } catch Error as clickErr {
                try {
                    advancedBtn.Invoke()
                } catch {
                }
            }
        }

        Sleep 200

        ; Locate and focus "Add a word" text field
        addWordEdit := el.FindFirst({ Type: 50004, Name: "Add a word" })
        if addWordEdit {
            addWordEdit.SetFocus()
        }
    } catch Error as e {
        MsgBox "Error in AddWordToHandy macro: " e.Message
    }
}

; Resolve Handy shortcut/executable path (environment-aware: work vs home)
; Work: uses Documents\Handy\handy.exe; Home: uses Start Menu shortcuts.
GetHandyShortcutPath() {
    global IS_WORK_ENVIRONMENT

    if (IS_WORK_ENVIRONMENT) {
        ; Work: direct exe path first, then work shortcut fallback
        workExe := "C:\Users\fie7ca\Documents\Handy\handy.exe"
        if (FileExist(workExe))
            return workExe
        for , p in ["C:\Users\fie7ca\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Handy\Handy.lnk",
            "C:\Users\fie7ca\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Handy.lnk"] {
            if (p != "" && FileExist(p))
                return p
        }
        return ""
    }

    ; Home/personal: Start Menu shortcuts only
    candidates := [
        "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Handy\Handy.lnk",
        "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Handy\Handy.lnk",
        "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Handy.lnk"
    ]
    for , p in candidates {
        try {
            if (p != "" && FileExist(p))
                return p
        } catch {
        }
    }
    return ""
}

; Expected Handy process exe path for current environment (for targeting correct instance).
; Work: work exe path; Home: "" (any Handy window).
GetHandyProcessPath() {
    global IS_WORK_ENVIRONMENT
    if (IS_WORK_ENVIRONMENT) {
        workExe := "C:\Users\fie7ca\Documents\Handy\handy.exe"
        return FileExist(workExe) ? workExe : ""
    }
    return ""
}

; =============================================================================
; Clip Angel: Merge Non-Favorite Clips
; =============================================================================
; Ensure Clip Angel window is closed (Alt+V then WinClose fallback)
EnsureClipAngelClosed() {
    ; #region agent log
    DebugLog_0ec0ba("run-clip-favorite", "H4", "EnsureClipAngelClosed enter", "clipAngelExistsBefore=" . DebugBool(!!
        WinExist("ClipAngel")))
    ; #endregion
    if !WinExist("ClipAngel")
        return
    Send "!v"
    Sleep 400
    if WinExist("ClipAngel") {
        Send "!v"
        Sleep 300
    }
    if WinExist("ClipAngel") {
        try WinClose("ClipAngel")
    }
    ; #region agent log
    DebugLog_0ec0ba("run-clip-favorite", "H4", "EnsureClipAngelClosed exit", "clipAngelExistsAfter=" . DebugBool(!!
        WinExist("ClipAngel")))
    ; #endregion
}

; Extract title from first non-favorite clip in ClipAngel
MergeNonFavoriteClips() {
    try {
        ; Show persistent banner for the duration of the algorithm
        AiModelBanner_Show("📋 Merging non-favorite clips...", "FFCC00")

        ; Step 1: Send Alt+B to activate ClipAngel (this opens the window if not visible)
        Send "!b"
        Sleep 500  ; Wait for ClipAngel window to appear

        ; Step 2: Check if ClipAngel window exists now
        if !WinExist("ClipAngel") {
            AiModelBanner_Hide()
            MsgBox "ClipAngel window did not appear. Make sure ClipAngel is running.", "Merge Clips", "IconX"
            return
        }
        try {
            WinActivate("ClipAngel")
        } catch {
            AiModelBanner_Hide()
            ShowCenteredOverlay_Utils("❌ ClipAngel window not found.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        WinWaitActive("ClipAngel", , 2)

        ; Step 3: Initialize UIA on ClipAngel window
        hwnd := WinExist("A")
        el := UIA.ElementFromHandle(hwnd)
        if !el {
            AiModelBanner_Hide()
            EnsureClipAngelClosed()
            MsgBox "Failed to initialize UIA for ClipAngel.", "Merge Clips", "IconX"
            return
        }

        ; Step 4: Find DataGridView by AutomationId
        dataGrid := 0
        try dataGrid := el.FindFirst({ Type: 50036, AutomationId: "dataGridView" })
        if !dataGrid {
            AiModelBanner_Hide()
            EnsureClipAngelClosed()
            MsgBox "Could not find DataGridView in ClipAngel.", "Merge Clips", "IconX"
            return
        }

        ; Step 5: Find Row 0
        row0 := 0
        try row0 := dataGrid.FindFirst({ Type: 50025, Name: "Row 0" })
        if !row0 {
            AiModelBanner_Hide()
            EnsureClipAngelClosed()
            MsgBox "No clips found in Row 0.", "Merge Clips", "IconX"
            return
        }

        ; Step 6: Find "Title Row 0" element
        titleElement := 0
        try titleElement := row0.FindFirst({ Type: 50006, Name: "Title Row 0" })
        if !titleElement {
            AiModelBanner_Hide()
            EnsureClipAngelClosed()
            MsgBox "Could not find Title element in Row 0.", "Merge Clips", "IconX"
            return
        }

        ; Step 7: Extract RTF value
        rtfValue := ""
        try rtfValue := titleElement.Value
        if (rtfValue = "" || rtfValue = "System.Drawing.Bitmap") {
            AiModelBanner_Hide()
            EnsureClipAngelClosed()
            MsgBox "Title Row 0 contains no text data.", "Merge Clips", "IconX"
            return
        }

        ; Step 8: Parse RTF to extract plain text (this is our target favorite clip title)
        favoriteClipTitle := ParseRTFToPlainText(rtfValue)

        ; Step 9: Switch to "All Clips" view to search for this favorite clip
        Send "!b"  ; Close current view
        Sleep 600
        Send "!v"  ; Open "All Clips" view (non-favorites first, favorites second)
        Sleep 500  ; Wait for view to update

        ; Step 10: Re-initialize UIA for the updated view
        hwnd := WinExist("A")
        el := UIA.ElementFromHandle(hwnd)
        if !el {
            AiModelBanner_Hide()
            EnsureClipAngelClosed()
            MsgBox "Failed to re-initialize UIA after switching views.", "Merge Clips", "IconX"
            return
        }

        ; Step 11: Find DataGridView again
        dataGrid := 0
        try dataGrid := el.FindFirst({ Type: 50036, AutomationId: "dataGridView" })
        if !dataGrid {
            AiModelBanner_Hide()
            EnsureClipAngelClosed()
            MsgBox "Could not find DataGridView in All Clips view.", "Merge Clips", "IconX"
            return
        }

        ; Step 12: Focus on Row 0 to start the search
        try {
            row0 := dataGrid.FindFirst({ Type: 50025, Name: "Row 0" })
            if row0 {
                row0.SetFocus()
                Sleep 100
            }
        }

        ; Step 13: Iterative search through rows (max 40 iterations)
        maxIterations := 40
        foundMatch := false
        currentRow := 0

        loop maxIterations {
            currentRow := A_Index - 1  ; 0-based row index

            ; Find current row
            currentRowElement := 0
            try currentRowElement := dataGrid.FindFirst({ Type: 50025, Name: "Row " . currentRow })

            if !currentRowElement {
                ; No more rows, stop searching
                break
            }

            ; Find title element in current row
            currentTitleElement := 0
            try currentTitleElement := currentRowElement.FindFirst({ Type: 50006, Name: "Title Row " . currentRow })

            if currentTitleElement {
                ; Extract and parse the title
                currentRtfValue := ""
                try currentRtfValue := currentTitleElement.Value

                if (currentRtfValue != "" && currentRtfValue != "System.Drawing.Bitmap") {
                    currentTitle := ParseRTFToPlainText(currentRtfValue)

                    ; Compare with the favorite clip title
                    if (currentTitle = favoriteClipTitle) {
                        foundMatch := true

                        ; Step 14: Select and merge non-favorite clips (cursor is on first favorite)
                        ; Move up once to last non-favorite clip
                        Send "{Up}"
                        Sleep 150
                        ; Select from current position to top of list (all non-favorites)
                        Send "^+{Home}"
                        Sleep 150
                        ; Merge the selected clips
                        Send "^!j"
                        Sleep 300  ; Wait for merge to complete

                        ; Step 15: Copy merged clip to clipboard
                        Send "{Tab}"   ; Focus merged content area
                        Sleep 150
                        Send "^a"     ; Select all
                        Sleep 100
                        Send "^c"     ; Copy

                        AiModelBanner_Hide()
                        ShowCenteredOverlay_Utils("✅ Merged non-favorite clips (copied)", 2000, BANNER_ACCENT_SUCCESS)
                        break
                    }
                }
            }

            ; Move to next row
            Send "{Down}"
            Sleep 100  ; Small delay between iterations
        }

        if !foundMatch {
            AiModelBanner_Hide()
            ShowCenteredOverlay_Utils("⚠ Favorite clip not found in first " . maxIterations . " rows", 2000,
                BANNER_ACCENT_INTERMEDIATE)
        }

        ; Guarantee Clip Angel is closed when macro finishes (success or not found)
        EnsureClipAngelClosed()

    } catch Error as e {
        AiModelBanner_Hide()
        EnsureClipAngelClosed()
        MsgBox "Error in MergeNonFavoriteClips: " . e.Message, "Merge Clips", "IconX"
    }
}

; Helper function to parse RTF and extract plain text
ParseRTFToPlainText(rtf) {
    ; Remove RTF header and formatting
    ; Pattern: extract text between last formatting and \par
    plainText := rtf

    ; Remove RTF control sequences (backslash followed by letters/numbers)
    plainText := RegExReplace(plainText, "\\[a-z]+[0-9]*\s?", "")
    ; Remove braces
    plainText := RegExReplace(plainText, "[{}]", "")
    ; Remove everything after \par
    plainText := RegExReplace(plainText, "\\par.*$", "")

    ; Trim whitespace
    plainText := Trim(plainText)

    return plainText
}

; =============================================================================
; Clip Angel: Open/Activate with focus correction (Row 0)
; =============================================================================
; Alt+V: Activate Clip Angel and ensure focus is on "Row 0" (fixes bug where focus
; defaults to upper tabs). Uses UIA: Type 50025, Name "Row 0" per clipangel-tree.txt.
ActivateClipAngelWithFocusCorrection() {
    needBanner := false
    if WinExist("ClipAngel") {
        try {
            WinActivate("ClipAngel")
        } catch {
            ShowCenteredOverlay_Utils("❌ ClipAngel window not found.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        WinWaitActive("ClipAngel", , 2)
    } else {
        needBanner := true
        ClipAngelBanner_Show("📂 Opening Clip Angel...", BANNER_ACCENT_INTERMEDIATE)
        Send "!v"
        if !WinWait("ClipAngel", , 10) {
            ClipAngelBanner_Hide()
            return
        }
        try {
            WinActivate("ClipAngel")
        } catch {
            ClipAngelBanner_Hide()
            ShowCenteredOverlay_Utils("❌ ClipAngel window not found.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        WinWaitActive("ClipAngel", , 2)
    }
    Sleep 50
    ; Must use Clip Angel's HWND, not WinExist("A") — another app can be foreground and UIA targets the wrong tree.
    hwnd := WinExist("ClipAngel")
    if !hwnd {
        if needBanner
            ClipAngelBanner_Hide()
        return
    }
    el := UIA.ElementFromHandle(hwnd)
    if !el {
        if needBanner
            ClipAngelBanner_Hide()
        return
    }
    try {
        dataGrid := ClipAngel_UiaFindFirst(el, { Type: 50036, AutomationId: "dataGridView" })
        if !dataGrid {
            if needBanner
                ClipAngelBanner_Hide()
            return
        }
        row0 := ClipAngel_UiaFindFirst(dataGrid, { Type: 50025, Name: "Row 0" })
        if !row0 {
            if needBanner
                ClipAngelBanner_Hide()
            return
        }
        hasSel := row0.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable)
        isSelected := hasSel && row0.SelectionItemPattern.IsSelected
        if (!isSelected) {
            if !needBanner
                ClipAngelBanner_Show("🎯 Focusing Row 0...", BANNER_ACCENT_INTERMEDIATE)
            needBanner := true
            try {
                if hasSel
                    row0.SelectionItemPattern.Select()
                else
                    row0.SetFocus()
            } catch {
                try row0.SetFocus()
            }
        }
    } catch {
        if needBanner
            ClipAngelBanner_Hide()
        return
    }
    if needBanner {
        ClipAngelBanner_Show("✅ Done", BANNER_ACCENT_SUCCESS)
        SetTimer(ClipAngelBanner_Hide, -500)
    }
}

; =============================================================================
; Clip Angel: Mark Last Clip as Favorite
; =============================================================================
; Shortcut flow (matches app): open Clip Angel, ensure list focus (not Window tab),
; select first or last grid row, Send Alt+Q. Optional: target "last" for bottom row.
; UIA-v2 FindFirst throws TargetError when nothing matches — never chain with if !c without try.
ClipAngel_UiaFindFirst(root, conditions) {
    if !root
        return 0
    try return root.FindFirst(conditions)
    catch
        return 0
}

ClipAngel_FindFavoriteCell(row) {
    if !row
        return 0
    rn := ""
    try rn := row.Name
    catch {
        rn := ""
    }
    suffix := "0"
    if RegExMatch(rn, "i)(?:Row|Linha)\s*(\d+)", &m)
        suffix := m[1]
    else if RegExMatch(rn, "(\d+)\s*$", &m)
        suffix := m[1]
    ; EN + PT-BR column headers seen in Clip Angel / localized WinForms.
    for cand in [
        "Favorite Row " . suffix, "Favourite Row " . suffix, "Favorito Row " . suffix,
        "Favorite Linha " . suffix, "Favorito Linha " . suffix
    ] {
        c := ClipAngel_UiaFindFirst(row, { Type: UIA.Type.CheckBox, Name: cand })
        if c
            return c
        c := ClipAngel_UiaFindFirst(row, { Type: 50002, Name: cand })
        if c
            return c
    }
    try {
        for c in row.FindAll({ Type: 50002 }) {
            try n := c.Name
            catch
                continue
            if RegExMatch(n, "i)favorite|favourite|favorito")
                return c
        }
        boxes := row.FindAll({ Type: 50002 })
        if boxes.Length >= 2
            return boxes[boxes.Length]
    } catch {
    }
    return 0
}

ClipAngel_FavoriteCellIsOn(cell) {
    if !cell
        return false
    try {
        if cell.GetPropertyValue(UIA.Property.IsTogglePatternAvailable)
            return cell.TogglePattern.ToggleState = UIA.ToggleState.On
        ts := cell.GetPropertyValue(UIA.Property.ToggleToggleState)
        if ts != ""
            return ts = UIA.ToggleState.On
    } catch {
    }
    ; Value only for read-only grid cells — Legacy CHECKED (0x10) often false-positives on DataGrid cells.
    try {
        v := cell.Value
        if (v = "true" || v = "True" || v = "1")
            return true
    } catch {
    }
    return false
}

ClipAngel_MainHwnd() {
    h := WinExist("ClipAngel")
    if h
        return h
    return WinExist("ahk_exe ClipAngel.exe")
}

; Macro hotkeys use Ctrl+Alt+Win — if those keys are still down, Send "!q" is not plain Alt+Q (Win+Alt+… hijacks it).
ClipAngel_ReleaseChordModifiersForSend() {
    SendInput "{LWin up}{RWin up}{LControl up}{RControl up}{LAlt up}{RAlt up}{LShift up}{RShift up}"
}

; Wait for physical release (KeyWait) then synthetic up — chord hotkeys often leave keys logically down.
ClipAngel_WaitChordModifiersReleased() {
    tw := "T0.45"
    KeyWait "Ctrl", tw
    KeyWait "Alt", tw
    KeyWait "Shift", tw
    KeyWait "LWin", tw
    KeyWait "RWin", tw
}

; target: "first" = top grid row (Row 0 / newest), "last" = last row returned by UIA FindAll
; (virtualized lists may only expose visible rows — use "first" for reliable top-clip behavior).
MarkLastClipAsFavorite(target := "first") {
    ActivateClipAngelWithFocusCorrection()
    hwnd := ClipAngel_MainHwnd()
    if !hwnd {
        ShowCenteredOverlay_Utils("❌ Clip Angel did not open.", 2000, BANNER_ACCENT_ERROR)
        return
    }
    try {
        try WinActivate("ahk_id " hwnd)
        catch {
            ShowCenteredOverlay_Utils("❌ Clip Angel window not found.", 2000, BANNER_ACCENT_ERROR)
            EnsureClipAngelClosed()
            return
        }
        if !WinWaitActive("ahk_id " hwnd, , 2) {
            ShowCenteredOverlay_Utils("❌ Clip Angel did not become active.", 2000, BANNER_ACCENT_ERROR)
            EnsureClipAngelClosed()
            return
        }
        el := UIA.ElementFromHandle(hwnd)
        if !el {
            ShowCenteredOverlay_Utils("❌ Clip Angel UI not available.", 2000, BANNER_ACCENT_ERROR)
            EnsureClipAngelClosed()
            return
        }
        dataGrid := ClipAngel_UiaFindFirst(el, { Type: 50036, AutomationId: "dataGridView" })
        if !dataGrid {
            ShowCenteredOverlay_Utils("❌ Clip list not found (Window tab may still have focus).", 2500,
                BANNER_ACCENT_ERROR)
            EnsureClipAngelClosed()
            return
        }
        rows := 0
        try rows := dataGrid.FindAll({ Type: 50025 })
        catch {
            rows := 0
        }
        if !rows || rows.Length < 1 {
            ShowCenteredOverlay_Utils("❌ No clips in list.", 2000, BANNER_ACCENT_ERROR)
            EnsureClipAngelClosed()
            return
        }
        rowTarget := rows[1]
        if (target = "last")
            rowTarget := rows[rows.Length]
        hasSel := rowTarget.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable)
        try {
            if hasSel
                rowTarget.SelectionItemPattern.Select()
            else
                rowTarget.SetFocus()
        } catch {
            try rowTarget.SetFocus()
        }
        try {
            if rowTarget.GetPropertyValue(UIA.Property.IsScrollItemPatternAvailable)
                rowTarget.ScrollItemPattern.ScrollIntoView()
        } catch {
        }
        favCell := ClipAngel_FindFavoriteCell(rowTarget)
        if favCell {
            if ClipAngel_FavoriteCellIsOn(favCell) {
                ShowCenteredOverlay_Utils("✅ Selected clip is already a favorite.", 1500, BANNER_ACCENT_SUCCESS)
                EnsureClipAngelClosed()
                return
            }
        }
        if !WinActive("ahk_id " hwnd) {
            try WinActivate("ahk_id " hwnd)
            if !WinWaitActive("ahk_id " hwnd, , 2) {
                ShowCenteredOverlay_Utils("❌ Clip Angel lost focus before Alt+Q.", 2000, BANNER_ACCENT_ERROR)
                EnsureClipAngelClosed()
                return
            }
        }
        ClipAngel_WaitChordModifiersReleased()
        ClipAngel_ReleaseChordModifiersForSend()
        SendInput "!q"
        ShowCenteredOverlay_Utils("✅ Sent Alt+Q — marked focused clip as favorite.", 1500, BANNER_ACCENT_SUCCESS)
        EnsureClipAngelClosed()
    } catch Error as e {
        ShowCenteredOverlay_Utils("❌ Mark favorite failed: " . e.Message, 2500, BANNER_ACCENT_ERROR)
        EnsureClipAngelClosed()
    }
}

DebugBool(v) {
    return v ? "true" : "false"
}

DebugJsonEscape(s) {
    s := StrReplace(s, "\", "\\")
    s := StrReplace(s, "`"", "\`"")
    s := StrReplace(s, "`r", "\r")
    s := StrReplace(s, "`n", "\n")
    return s
}

DebugLog_0ec0ba(runId, hypothesisId, message, dataJson := "{}") {
    try {
        logPath := A_ScriptDir "\debug-0ec0ba.log"
        payload := "{`"sessionId`":`"0ec0ba`",`"runId`":`"" . DebugJsonEscape(runId)
        . "`",`"hypothesisId`":`"" . DebugJsonEscape(hypothesisId)
        . "`",`"location`":`"Utils.ahk`",`"message`":`"" . DebugJsonEscape(message)
        . "`",`"data`":`"" . DebugJsonEscape(dataJson)
        . "`",`"timestamp`":" . A_TickCount . "}"
        FileAppend(payload . "`n", logPath, "UTF-8")
    } catch {
    }
}

; =============================================================================
; AI Model Selection System for Handy
; =============================================================================
; Configuration: Maps selection numbers (1–4) to AI model names.
; These are partial name prefixes used to find buttons in the UIA tree (Type 50000, botão).
; Descriptions match Handy Transcription Models UI for quick verification.
; Slots 3–4: set Cohere Language on General tab before selecting the Cohere model (modelClickName).
global g_HandyAiModels := Map(
    1, { name: "Parakeet V2", desc: "English only. Best model for English speakers." },
    2, { name: "Parakeet V3", desc: "Fast and accurate. Multi-language." },
    3, { name: "Cohere English", desc: "Sets Cohere language to English (General), then activates Cohere.",
        cohereLanguage: "English", modelClickName: "Cohere" },
    4, { name: "Cohere Portuguese", desc: "Sets Cohere language to Portuguese (General), then activates Cohere.",
        cohereLanguage: "Portuguese", modelClickName: "Cohere" }
)

; Picker indices for ^!#9 / ^!#b; update g_HandyAiModels names if Handy renames models.
global HANDY_AI_SLOT_PARAKEET_V3 := 2
global HANDY_AI_SLOT_PARAKEET_V2 := 1

; GUI state for AI model selector
global g_AiModelSelectorGui := false
global g_AiModelSelectorActive := false
global g_AiModelBannerGui := false

; =============================================================================
; ShowAiModelSelector() - Display selection GUI with immediate key capture
; =============================================================================
ShowAiModelSelector() {
    global g_AiModelSelectorGui, g_AiModelSelectorActive, g_HandyAiModels

    ; Don't show if already active
    if (g_AiModelSelectorActive)
        return

    ; Create selection GUI
    g_AiModelSelectorGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Owner")
    g_AiModelSelectorGui.BackColor := "1E1E2E"
    g_AiModelSelectorGui.MarginX := 20
    g_AiModelSelectorGui.MarginY := 15

    ; Title
    g_AiModelSelectorGui.SetFont("s14 cCDD6F4 Bold", "Segoe UI")
    g_AiModelSelectorGui.Add("Text", "w280 Center", "🎙️ Select AI Model")
    g_AiModelSelectorGui.Add("Text", "w280 h1 Background45475A")  ; separator

    ; Model options
    g_AiModelSelectorGui.SetFont("s12 cCDD6F4", "Segoe UI")
    for num, model in g_HandyAiModels {
        g_AiModelSelectorGui.Add("Text", "w280", "[" . num . "] " . model.name)
        g_AiModelSelectorGui.SetFont("s9 c6C7086", "Segoe UI")
        g_AiModelSelectorGui.Add("Text", "w280 y+2", "    " . model.desc)
        g_AiModelSelectorGui.SetFont("s12 cCDD6F4", "Segoe UI")
    }

    ; Footer
    g_AiModelSelectorGui.Add("Text", "w280 h1 Background45475A y+10")
    g_AiModelSelectorGui.SetFont("s9 c6C7086", "Segoe UI")
    g_AiModelSelectorGui.Add("Text", "w280 Center", "Press 1–4 | Esc to cancel")

    ; Get active window to determine which monitor to center on
    activeWin := 0
    try {
        activeWin := WinGetID("A")
    } catch {
        activeWin := 0
    }

    ; Default to primary monitor work area
    MonitorGetWorkArea(1, &monitorLeft, &monitorTop, &monitorRight, &monitorBottom)
    monitorWidth := monitorRight - monitorLeft
    monitorHeight := monitorBottom - monitorTop

    ; If we have an active window, find which monitor contains its center
    if (activeWin && activeWin != 0) {
        rect := Buffer(16, 0)
        if (DllCall("GetWindowRect", "ptr", activeWin, "ptr", rect)) {
            ; Calculate window center
            winLeft := NumGet(rect, 0, "int")
            winTop := NumGet(rect, 4, "int")
            winRight := NumGet(rect, 8, "int")
            winBottom := NumGet(rect, 12, "int")

            centerX := winLeft + (winRight - winLeft) // 2
            centerY := winTop + (winBottom - winTop) // 2

            ; Find which monitor contains the window center
            monitorCount := MonitorGetCount()
            loop monitorCount {
                idx := A_Index
                MonitorGetWorkArea(idx, &l, &t, &r, &b)
                if (centerX >= l && centerX <= r && centerY >= t && centerY <= b) {
                    monitorLeft := l
                    monitorTop := t
                    monitorRight := r
                    monitorBottom := b
                    monitorWidth := r - l
                    monitorHeight := b - t
                    break
                }
            }
        }
    }

    ; Measure GUI size and center on the active monitor
    g_AiModelSelectorGui.Show("AutoSize Hide")
    g_AiModelSelectorGui.GetPos(&gx, &gy, &gw, &gh)
    cx := monitorLeft + (monitorWidth - gw) // 2
    cy := monitorTop + (monitorHeight - gh) // 2
    g_AiModelSelectorGui.Show("x" . cx . " y" . cy . " NA")

    g_AiModelSelectorActive := true

    ; Enable hotkeys for 1–4 and Escape
    Hotkey("1", AiModelSelector_HandleKey, "On")
    Hotkey("2", AiModelSelector_HandleKey, "On")
    Hotkey("3", AiModelSelector_HandleKey, "On")
    Hotkey("4", AiModelSelector_HandleKey, "On")
    Hotkey("Escape", AiModelSelector_Cancel, "On")
}

; Handle key press in AI model selector
AiModelSelector_HandleKey(key) {
    global g_AiModelSelectorGui, g_AiModelSelectorActive, g_HandyAiModels

    if (!g_AiModelSelectorActive)
        return

    ; Get the selection number
    selection := Integer(key)

    ; Close selector GUI
    AiModelSelector_Close()

    ; Execute the selection
    if (g_HandyAiModels.Has(selection)) {
        ExecuteHandyAiModelSelection(selection)
    }
}

; Cancel AI model selector
AiModelSelector_Cancel(*) {
    AiModelSelector_Close()
}

; Close the selector GUI and disable hotkeys
AiModelSelector_Close() {
    global g_AiModelSelectorGui, g_AiModelSelectorActive

    if (!g_AiModelSelectorActive)
        return

    g_AiModelSelectorActive := false

    ; Disable hotkeys
    try Hotkey("1", "Off")
    try Hotkey("2", "Off")
    try Hotkey("3", "Off")
    try Hotkey("4", "Off")
    try Hotkey("Escape", AiModelSelector_Cancel, "Off")

    ; Destroy GUI
    if (IsObject(g_AiModelSelectorGui) && g_AiModelSelectorGui.Hwnd) {
        try g_AiModelSelectorGui.Destroy()
    }
    g_AiModelSelectorGui := false
}

; =============================================================================
; Gemini-to-Cursor transfer: numeric Cursor window selector (1–9) and activate/focus/paste
; Used when user presses [C] Transfer in Gemini copy-decision banner.
; =============================================================================
global g_CursorTransferSelectorGui := false
global g_CursorTransferSelectorActive := false
global g_CursorTransferSelectorResult := ""   ; "" = waiting, 0 = cancel, integer = selected hwnd
global g_CursorTransferWindowList := []      ; up to 9 { hwnd, title }
global g_CursorTransferHotkeyHandlers := []
global g_CursorTransferPidCmdCache := Map()

CursorTransfer_SelectorClose() {
    global g_CursorTransferSelectorActive, g_CursorTransferSelectorGui, g_CursorTransferHotkeyHandlers
    if (!g_CursorTransferSelectorActive)
        return
    g_CursorTransferSelectorActive := false
    for h in g_CursorTransferHotkeyHandlers {
        if (IsObject(h) && h.HasProp("key") && h.HasProp("callback")) {
            try Hotkey(h.key, h.callback, "Off")
        } else {
            try Hotkey(h, "Off")
        }
    }
    g_CursorTransferHotkeyHandlers := []
    if (IsObject(g_CursorTransferSelectorGui) && g_CursorTransferSelectorGui.Hwnd) {
        try g_CursorTransferSelectorGui.Destroy()
    }
    g_CursorTransferSelectorGui := false
}

CursorTransfer_SelectorHandleKey(index, *) {
    global g_CursorTransferSelectorResult, g_CursorTransferWindowList
    if (index >= 1 && index <= g_CursorTransferWindowList.Length)
        g_CursorTransferSelectorResult := g_CursorTransferWindowList[index].hwnd
    CursorTransfer_SelectorClose()
}

CursorTransfer_SelectorEscape(*) {
    global g_CursorTransferSelectorResult
    g_CursorTransferSelectorResult := 0
    CursorTransfer_SelectorClose()
}

; Return project order index from g_Projects for a window title; 0 = no match.
CursorTransfer_GetProjectOrderForTitle(winTitle) {
    idx := CursorTransfer_GetMatchingProjectIndexForTitle(winTitle)
    return idx
}

; Build canonical project-index -> character mapping (same logic as standard project selector).
CursorTransfer_BuildProjectIndexToChar() {
    global g_Projects, g_ProjectCategories, g_ProjectCharSequence
    projectIndexToChar := Map()
    projectIndexToCategory := Map()
    loop g_Projects.Length {
        projectIndex := A_Index
        project := g_Projects[projectIndex]
        category := project.HasProp("category") ? project.category : "Personal"
        projectIndexToCategory[projectIndex] := category
    }
    charIndex := 1
    for category in g_ProjectCategories {
        categoryProjectIndices := []
        for projectIndex, cat in projectIndexToCategory {
            if (cat = category)
                categoryProjectIndices.Push(projectIndex)
        }
        for projectIndex in categoryProjectIndices {
            project := g_Projects[projectIndex]
            if (project.name = "" && project.path = "" && project.workPath = "") {
                charIndex++
                continue
            }
            if (charIndex > g_ProjectCharSequence.Length)
                break
            char := g_ProjectCharSequence[charIndex]
            ; Keep parity with standard selector where 3 is reserved.
            if (char = "3") {
                charIndex++
                if (charIndex > g_ProjectCharSequence.Length)
                    break
                char := g_ProjectCharSequence[charIndex]
            }
            projectIndexToChar[projectIndex] := char
            charIndex++
        }
    }
    return projectIndexToChar
}

; Return matching project index from g_Projects for a window title; 0 = no match.
; Uses longest matching path segment so "user-scripts" wins over "scripts" when both match.
CursorTransfer_GetMatchingProjectIndexForTitle(winTitle) {
    global g_Projects
    if (!winTitle || !IsObject(g_Projects))
        return 0
    try {
        isWork := (IsSet(IS_WORK_ENVIRONMENT) && IS_WORK_ENVIRONMENT)
        winLow := StrLower(winTitle)
        bestIdx := 0
        bestScore := 0
        loop g_Projects.Length {
            project := g_Projects[A_Index]
            if (project.name = "" && project.path = "" && project.workPath = "")
                continue
            projectPath := isWork ? project.workPath : project.path
            if (isWork && (projectPath = ""))
                projectPath := project.path
            if (projectPath = "")
                continue
            matchSegments := ExtractProjectMatchSegments(projectPath)
            for segment in matchSegments {
                if (segment = "")
                    continue
                if (InStr(winLow, StrLower(segment))) {
                    len := StrLen(segment)
                    if (len > bestScore) {
                        bestScore := len
                        bestIdx := A_Index
                    }
                }
            }
        }
        return bestIdx
    } catch {
    }
    return 0
}

; Return project path according to current environment for a given project object.
CursorTransfer_GetEffectiveProjectPath(project) {
    isWork := (IsSet(IS_WORK_ENVIRONMENT) && IS_WORK_ENVIRONMENT)
    projectPath := isWork ? project.workPath : project.path
    if (isWork && (projectPath = ""))
        projectPath := project.path
    return projectPath
}

; True when the OS title has no workspace/file segment (e.g. bare "Cursor" welcome screen).
CursorTransfer_IsUninformativeCursorTitle(winTitle) {
    t := Trim(winTitle)
    if (t = "")
        return true
    tl := StrLower(t)
    if (tl = "cursor")
        return true
    ; Strip trailing " - Cursor"; if nothing remains, title had no file/workspace name.
    rest := RegExReplace(tl, "\s*[-–—]\s*cursor\s*$", "")
    if (Trim(rest) = "")
        return true
    return false
}

; Longest project path wins when multiple g_Projects paths appear in the same command line.
CursorTransfer_GetMatchingProjectIndexByCmdLine(cmdLine) {
    global g_Projects
    if (!cmdLine || !IsObject(g_Projects))
        return 0
    cmdLow := StrLower(cmdLine)
    bestIdx := 0
    bestLen := 0
    try {
        loop g_Projects.Length {
            project := g_Projects[A_Index]
            if (project.name = "" && project.path = "" && project.workPath = "")
                continue
            projectPath := CursorTransfer_GetEffectiveProjectPath(project)
            if (projectPath = "")
                continue
            projLow := StrLower(RTrim(projectPath, "\"))
            if (projLow != "" && InStr(cmdLow, projLow)) {
                len := StrLen(projLow)
                if (len > bestLen) {
                    bestLen := len
                    bestIdx := A_Index
                }
            }
        }
    } catch {
    }
    return bestIdx
}

; Get process command line by PID (cached). Returns "" on failure.
CursorTransfer_GetProcessCommandLine(pid) {
    global g_CursorTransferPidCmdCache
    if (!pid)
        return ""
    if (g_CursorTransferPidCmdCache.Has(pid))
        return g_CursorTransferPidCmdCache[pid]
    cmd := ""
    try {
        locator := ComObject("WbemScripting.SWbemLocator")
        svc := locator.ConnectServer(".", "root\cimv2")
        for proc in svc.ExecQuery("SELECT CommandLine FROM Win32_Process WHERE ProcessId = " pid) {
            try cmd := proc.CommandLine
            break
        }
    } catch {
        cmd := ""
    }
    g_CursorTransferPidCmdCache[pid] := cmd ? cmd : ""
    return g_CursorTransferPidCmdCache[pid]
}

; Title match when the title lists workspace/file; cmd-line when title is bare "Cursor" etc.
; (Same PID often shares one command line — use longest path in cmd for disambiguation.)
CursorTransfer_GetMatchingProjectIndex(hwnd, winTitle := "") {
    global g_Projects
    if (!IsObject(g_Projects))
        return 0
    pid := 0
    try pid := WinGetPID("ahk_id " hwnd)
    cmdLine := CursorTransfer_GetProcessCommandLine(pid)
    idxCmd := CursorTransfer_GetMatchingProjectIndexByCmdLine(cmdLine)
    idxTitle := (winTitle != "") ? CursorTransfer_GetMatchingProjectIndexForTitle(winTitle) : 0
    if (CursorTransfer_IsUninformativeCursorTitle(winTitle)) {
        if (idxCmd > 0)
            return idxCmd
        return idxTitle
    }
    if (idxTitle > 0)
        return idxTitle
    return idxCmd
}

; Stable insertion sort for small arrays by project order, then by title.
CursorTransfer_SortWindowsByProjectOrder(&arr) {
    n := arr.Length
    if (n <= 1)
        return
    loop n - 1 {
        i := A_Index + 1
        key := arr[i]
        j := i - 1
        while (j >= 1) {
            left := arr[j]
            shouldShift := false
            ; Coerce to integer so comparison never sees a string (projectOrder can be string from g_Projects).
            try
                leftOrd := Integer(left.projectOrder)
            catch
                leftOrd := 0
            try
                keyOrd := Integer(key.projectOrder)
            catch
                keyOrd := 0
            if (leftOrd > keyOrd) {
                shouldShift := true
            } else if (leftOrd = keyOrd) {
                leftName := ""
                keyName := ""
                try
                    leftName := String(left.displayName)
                catch
                    leftName := ""
                try
                    keyName := String(key.displayName)
                catch
                    keyName := ""
                if (StrCompare(leftName, keyName) > 0)
                    shouldShift := true
            }
            if (!shouldShift)
                break
            arr[j + 1] := left
            j--
        }
        arr[j + 1] := key
    }
}

; Return project name from g_Projects if window title matches a project path; otherwise "".
CursorTransfer_GetProjectNameForTitle(winTitle) {
    global g_Projects
    idx := CursorTransfer_GetMatchingProjectIndexForTitle(winTitle)
    if (!idx || !IsObject(g_Projects))
        return ""
    try {
        project := g_Projects[idx]
        if (project.name != "")
            return project.name
    } catch {
    }
    return ""
}

CursorTransfer_StripStaticScriptTokenForDisplay(projectName) {
    ; Removes redundant static token(s) like "Script"/"Scripts" from the project label.
    ; If stripping would erase the whole name (e.g. folder is literally "Scripts"), keep the original.
    if (!projectName)
        return ""
    orig := Trim(projectName)
    cleaned := projectName
    ; Match whole-word "Script" or "Scripts" (case-insensitive).
    cleaned := RegExReplace(cleaned, "(?i)\bscript(s)?\b", "")
    cleaned := RegExReplace(cleaned, "\s{2,}", " ")
    cleaned := Trim(cleaned)
    if (StrLen(cleaned) < 2)
        return orig
    return cleaned
}

; Collapse "file.ext (Label) (file.ext)" in Cursor titles to a single filename + label.
CursorTransfer_StripDuplicateFilenameInParens(title) {
    if (!title)
        return ""
    return RegExReplace(title, "(\S+\.\w+)\s+(\([^)]+\))\s+\(\1\)", "$1 $2")
}

; If the window title starts with the project name (already shown in brackets), drop that prefix.
CursorTransfer_StripLeadingProjectFromTitle(title, cleanProjName, projName) {
    if (!title)
        return ""
    ; Longer label first so a shorter prefix cannot steal a match from a longer project name.
    names := []
    if (StrLen(cleanProjName) > StrLen(projName)) {
        if (cleanProjName != "")
            names.Push(cleanProjName)
        if (projName != "")
            names.Push(projName)
    } else {
        if (projName != "")
            names.Push(projName)
        if (cleanProjName != "" && cleanProjName != projName)
            names.Push(cleanProjName)
    }
    for name in names {
        if (StrLen(name) < 2)
            continue
        tl := StrLower(title)
        nl := StrLower(name)
        if (SubStr(tl, 1, StrLen(nl)) != nl)
            continue
        rest := SubStr(title, StrLen(name) + 1)
        rest := Trim(rest)
        if (rest = "")
            return ""
        if (SubStr(rest, 1, 1) = "-" || SubStr(rest, 1, 1) = "|" || SubStr(rest, 1, 1) = ":" || SubStr(rest, 1, 1) =
        "–")
            rest := Trim(SubStr(rest, 2))
        rest := Trim(LTrim(rest, "- "))
        return (rest != "") ? rest : title
    }
    return title
}

; Remove repetitive " - Cursor" suffix (and bare "Cursor") from list labels; every row is already Cursor.
CursorTransfer_StripTrailingCursorAppSuffix(s) {
    if (!s)
        return ""
    t := Trim(s)
    if (StrLower(t) = "cursor")
        return ""
    return Trim(RegExReplace(t, "i)\s*[-–—]\s*Cursor\s*$", ""))
}

Clipboard_GetSequenceNumber() {
    ; WinAPI: https://learn.microsoft.com/windows/win32/api/winuser/nf-winuser-getclipboardsequencenumber
    try {
        return DllCall("GetClipboardSequenceNumber", "uint")
    } catch {
        return 0
    }
}

Clipboard_WaitForSequenceChange(seqBefore, totalTimeoutMs := 2000, fastPhaseMs := 850) {
    ; Tight, bounded wait: returns true as soon as the clipboard sequence changes.
    ; Uses a fast polling phase first, then a slower phase for the remainder.
    start := A_TickCount
    deadline := start + totalTimeoutMs
    fastDeadline := start + fastPhaseMs
    while (A_TickCount < deadline) {
        seqNow := Clipboard_GetSequenceNumber()
        if (seqNow && seqNow != seqBefore)
            return true
        Sleep((A_TickCount < fastDeadline) ? 20 : 50)
    }
    return false
}

GetGeminiScriptMsgTargetHwnd() {
    ; Cache-first resolver for Gemini.ahk AutoHotkey script window.
    static cached := 0
    if (cached && WinExist("ahk_id " cached)) {
        try {
            if (InStr(WinGetTitle("ahk_id " cached), "Gemini.ahk"))
                return cached
        } catch {
        }
    }

    prevMatch := A_TitleMatchMode
    DetectHiddenWindows(true)
    SetTitleMatchMode(2)
    found := 0
    try {
        for hwnd in WinGetList("ahk_exe AutoHotkey64.exe") {
            try {
                if (InStr(WinGetTitle("ahk_id " hwnd), "Gemini.ahk")) {
                    found := hwnd
                    break
                }
            } catch {
                continue
            }
        }
        if (!found) {
            for hwnd in WinGetList("ahk_exe AutoHotkey32.exe") {
                try {
                    if (InStr(WinGetTitle("ahk_id " hwnd), "Gemini.ahk")) {
                        found := hwnd
                        break
                    }
                } catch {
                    continue
                }
            }
        }
    } finally {
        SetTitleMatchMode(prevMatch)
        DetectHiddenWindows(false)
    }

    if (found)
        cached := found
    return found
}

; Returns selected Cursor window HWND or 0 on cancel/timeout/no windows. Blocking with timeout.
; centerOnHwnd: optional window to center the modal on (uses that window's monitor); 0 = foreground monitor.
CursorTransfer_ShowWindowSelector(centerOnHwnd := 0) {
    global g_CursorTransferSelectorGui, g_CursorTransferSelectorActive, g_CursorTransferSelectorResult
    global g_CursorTransferWindowList, g_CursorTransferHotkeyHandlers
    global g_Projects, g_ProjectCharSequence
    CursorTransfer_SelectorClose()
    list := []
    DetectHiddenWindows true
    try {
        for hwnd in WinGetList("ahk_exe Cursor.exe") {
            try {
                title := WinGetTitle("ahk_id " hwnd)
                if (title = "" || InStr(StrLower(title), "preview"))
                    continue
                if (title = "MSCTFIME UI" || title = "Default IME")
                    continue
                list.Push({ hwnd: hwnd, title: title })
                if (list.Length >= 9)
                    break
            } catch {
                continue
            }
        }
    } catch {
    }
    DetectHiddenWindows false
    if (list.Length = 0) {
        ShowCenteredOverlay_Utils("❌ No Cursor windows found", 2000, BANNER_ACCENT_ERROR)
        return 0
    }
    ; Enrich list using path-first project identification, then sort by canonical project order.
    enriched := []
    for w in list {
        winTitle := w.title ? w.title : ""
        projectIndex := CursorTransfer_GetMatchingProjectIndex(w.hwnd, winTitle)
        projectOrder := Integer(projectIndex > 0 ? projectIndex : 10000 + enriched.Length)
        projName := ""
        if (projectIndex > 0) {
            try {
                project := g_Projects[projectIndex]
                projName := project.name
            }
        }
        displayName := ""
        shortAfterParens := ""
        shortAfterLeadStrip := ""
        cleanProjName := ""
        if (projName != "") {
            shortTitle := winTitle ? winTitle : ""
            cleanProjName := CursorTransfer_StripStaticScriptTokenForDisplay(projName)
            if (shortTitle != "") {
                shortTitle := CursorTransfer_StripDuplicateFilenameInParens(shortTitle)
                shortAfterParens := shortTitle
                shortTitle := CursorTransfer_StripLeadingProjectFromTitle(shortTitle, cleanProjName, projName)
                shortAfterLeadStrip := shortTitle
            }
            ; Label = window title only (no g_Projects name prefix). If stripping the duplicate workspace
            ; segment would leave only "Cursor", keep the fuller line (e.g. "scripts - Cursor").
            if (shortAfterParens != "") {
                cand := shortAfterLeadStrip
                if (cand = "" || CursorTransfer_IsUninformativeCursorTitle(cand))
                    displayName := shortAfterParens
                else
                    displayName := cand
                if (CursorTransfer_IsUninformativeCursorTitle(displayName))
                    displayName := displayName . " · #" . w.hwnd
            } else {
                displayName := winTitle ? winTitle : ("Cursor Window " . w.hwnd)
            }
        } else {
            displayName := winTitle ? winTitle : ("Cursor Window " . w.hwnd)
        }
        displayName := CursorTransfer_StripTrailingCursorAppSuffix(displayName)
        if (displayName = "")
            displayName := "#" . w.hwnd
        enriched.Push({
            hwnd: w.hwnd,
            title: winTitle,
            displayName: displayName,
            projectOrder: projectOrder,
            projectIndex: projectIndex,
            hotkeyChar: ""
        })
    }
    if (enriched.Length = 0) {
        ShowCenteredOverlay_Utils("❌ No mapped Cursor projects found", 2000, BANNER_ACCENT_ERROR)
        return 0
    }
    try {
        CursorTransfer_SortWindowsByProjectOrder(&enriched)
        ; Most important first: put active Cursor window at position 1 if it's in the list.
        try {
            activeHwnd := WinGetID("A")
            if (activeHwnd) {
                loop enriched.Length {
                    if (enriched[A_Index].hwnd = activeHwnd) {
                        if (A_Index > 1) {
                            swap := enriched[1]
                            enriched[1] := enriched[A_Index]
                            enriched[A_Index] := swap
                        }
                        break
                    }
                }
            }
        } catch {
        }
        if (enriched.Length > 9) {
            trimmed := []
            loop 9
                trimmed.Push(enriched[A_Index])
            list := trimmed
        } else {
            list := enriched
        }
        ; Number keys 1-9 (like #!+C / SelectAiModelInHandy modal).
        loop list.Length {
            list[A_Index].hotkeyChar := String(A_Index)
        }
        g_CursorTransferWindowList := list
        g_CursorTransferSelectorResult := ""
        g_CursorTransferSelectorActive := true
        g_CursorTransferSelectorGui := Gui("+AlwaysOnTop -Caption +ToolWindow")
        g_CursorTransferSelectorGui.BackColor := "1E1E2E"
        g_CursorTransferSelectorGui.MarginX := 20
        g_CursorTransferSelectorGui.MarginY := 15
        g_CursorTransferSelectorGui.OnEvent("Escape", CursorTransfer_SelectorEscape)
        g_CursorTransferSelectorGui.SetFont("s14 cCDD6F4 Bold", "Segoe UI")
        transferSelGuiW := 720
        g_CursorTransferSelectorGui.Add("Text", "w" . transferSelGuiW . " Center", "📋 Transfer to Cursor")
        g_CursorTransferSelectorGui.Add("Text", "w" . transferSelGuiW . " h1 Background45475A")
        g_CursorTransferSelectorGui.SetFont("s12 cCDD6F4", "Segoe UI")
        for w in list {
            g_CursorTransferSelectorGui.Add("Text", "w" . transferSelGuiW, "[" . w.hotkeyChar . "] " . w.displayName)
        }
        g_CursorTransferSelectorGui.Add("Text", "w" . transferSelGuiW . " h1 Background45475A y+10")
        g_CursorTransferSelectorGui.SetFont("s9 c6C7086", "Segoe UI")
        g_CursorTransferSelectorGui.Add("Text", "w" . transferSelGuiW . " Center", "Press 1-9 | N or Esc to cancel")
        g_CursorTransferSelectorGui.Show("AutoSize Hide")
        g_CursorTransferSelectorGui.GetPos(&gx, &gy, &gw, &gh)
        ; Center on centerOnHwnd's monitor when provided; otherwise foreground window's monitor (dictation / transfer flow).
        monitorIndex := 1
        if (centerOnHwnd && WinExist("ahk_id " centerOnHwnd)) {
            rect := Buffer(16, 0)
            if (DllCall("GetWindowRect", "ptr", centerOnHwnd, "ptr", rect)) {
                cx := NumGet(rect, 0, "int") + (NumGet(rect, 8, "int") - NumGet(rect, 0, "int")) // 2
                cy := NumGet(rect, 4, "int") + (NumGet(rect, 12, "int") - NumGet(rect, 4, "int")) // 2
                monitorCount := MonitorGetCount()
                loop monitorCount {
                    idx := A_Index
                    MonitorGet(idx, &ml, &mt, &mr, &mb)
                    if (cx >= ml && cx <= mr && cy >= mt && cy <= mb) {
                        monitorIndex := idx
                        break
                    }
                }
            }
        } else {
            monitorIndex := GetMonitorIndexForForeground_StandardBar()
        }
        MonitorGetWorkArea(monitorIndex, &ml, &mt, &mr, &mb)
        mw := mr - ml
        mh := mb - mt
        cx := ml + (mw - gw) // 2
        cy := mt + (mh - gh) // 2
        g_CursorTransferSelectorGui.Show("x" . cx . " y" . cy)
    } catch as err {
        ShowCenteredOverlay_Utils("❌ Selector error", 2000, BANNER_ACCENT_ERROR)
        return 0
    }
    ; Wildcard prefix so hotkeys fire even when modifier held (e.g. C still down when modal opens).
    g_CursorTransferHotkeyHandlers := []
    loop list.Length {
        i := A_Index
        keyChar := list[i].hotkeyChar
        if (keyChar = "")
            continue
        try {
            fn := CursorTransfer_SelectorHandleKey.Bind(i)
            hotkeyKey := "*" . keyChar
            Hotkey(hotkeyKey, fn, "On")
            g_CursorTransferHotkeyHandlers.Push({ key: hotkeyKey, callback: fn })
        } catch {
        }
    }
    try {
        Hotkey("*Escape", CursorTransfer_SelectorEscape, "On")
        g_CursorTransferHotkeyHandlers.Push({ key: "*Escape", callback: CursorTransfer_SelectorEscape })
        Hotkey("*N", CursorTransfer_SelectorEscape, "On")
        g_CursorTransferHotkeyHandlers.Push({ key: "*N", callback: CursorTransfer_SelectorEscape })
        Hotkey("*n", CursorTransfer_SelectorEscape, "On")
        g_CursorTransferHotkeyHandlers.Push({ key: "*n", callback: CursorTransfer_SelectorEscape })
    } catch {
    }
    start := A_TickCount
    timeoutMs := 30000
    lastCursorTransferMonitorIdx := monitorIndex
    try {
        while (g_CursorTransferSelectorResult = "") {
            if ((A_TickCount - start) >= timeoutMs)
                break
            Sleep 50
            if (IsObject(g_CursorTransferSelectorGui)) {
                curIdx := GetMonitorIndexForForeground_StandardBar()
                if (curIdx != lastCursorTransferMonitorIdx) {
                    lastCursorTransferMonitorIdx := curIdx
                    MonitorGetWorkArea(curIdx, &ml, &mt, &mr, &mb)
                    mw := mr - ml
                    mh := mb - mt
                    try {
                        g_CursorTransferSelectorGui.GetPos(&gxOld, &gyOld, &gw, &gh)
                        cx := ml + (mw - gw) // 2
                        cy := mt + (mh - gh) // 2
                        g_CursorTransferSelectorGui.Show("x" . cx . " y" . cy)
                    } catch {
                    }
                }
            }
        }
    } catch as loopErr {
        ; ignore loop exceptions
    }
    durationMs := A_TickCount - start
    result := (g_CursorTransferSelectorResult = "") ? 0 : Integer(g_CursorTransferSelectorResult)
    CursorTransfer_SelectorClose()
    return result
}

; =============================================================================
; Cursor AI text field focus (shared logic for Gemini transfer and WindowManagement)
; =============================================================================
Cursor_EnsureComposerHasFocus(editEl) {
    if (!editEl)
        return false
    try {
        editEl.SetFocus()
    } catch {
    }
    loop 3 {
        try {
            if (editEl.HasKeyboardFocus)
                return true
        } catch {
        }
        Sleep 40
    }
    try {
        editEl.ScrollIntoView()
    } catch {
    }
    try {
        editEl.Click()
    } catch {
    }
    Sleep 60
    try {
        return editEl.HasKeyboardFocus
    } catch {
        return false
    }
}

Cursor_FindComposerInput(root) {
    try {
        allEdits := root.FindAll({ Type: UIA.Type.Edit })
        for editEl in allEdits {
            cn := editEl.ClassName
            if (InStr(cn, "aislash-editor-input") && !InStr(cn, "readonly"))
                return editEl
        }
    } catch {
    }
    return ""
}

; Activate Cursor window and focus AI composer input. Returns true on success.
Cursor_FocusAITextField(targetHwnd := 0) {
    try {
        if (targetHwnd) {
            WinActivate("ahk_id " targetHwnd)
            WinWaitActive("ahk_id " targetHwnd, , 2)
        } else {
            targetHwnd := WinExist("ahk_exe Cursor.exe")
            if (!targetHwnd)
                return false
            WinWaitActive("ahk_id " targetHwnd, , 2)
        }
        Sleep 200
        paneWasOpen := false
        focusDone := false
        if (IsSet(UIA)) {
            try {
                root := UIA.ElementFromHandle(targetHwnd)
                if (root) {
                    toggleEl := root.FindFirst({ Type: UIA.Type.CheckBox, Name: "Toggle AI Pane", matchmode: 2 })
                    paneOpen := toggleEl && InStr(toggleEl.ClassName, "checked")
                    paneWasOpen := paneOpen
                    if (paneOpen) {
                        editEl := Cursor_FindComposerInput(root)
                        if (editEl) {
                            if (Cursor_EnsureComposerHasFocus(editEl))
                                focusDone := true
                        }
                    } else {
                        editEl := Cursor_FindComposerInput(root)
                        if (editEl) {
                            if (Cursor_EnsureComposerHasFocus(editEl))
                                focusDone := true
                        } else {
                            Send "^i"
                            loop 15 {
                                Sleep 200
                                root := UIA.ElementFromHandle(targetHwnd)
                                editEl := Cursor_FindComposerInput(root)
                                if (editEl) {
                                    if (Cursor_EnsureComposerHasFocus(editEl))
                                        focusDone := true
                                    break
                                }
                            }
                        }
                    }
                }
            } catch {
            }
        }
        if (!focusDone) {
            if (IsSet(UIA)) {
                try {
                    root := UIA.ElementFromHandle(targetHwnd)
                    if (root) {
                        editEl := Cursor_FindComposerInput(root)
                        if (editEl) {
                            if (Cursor_EnsureComposerHasFocus(editEl))
                                focusDone := true
                        }
                    }
                } catch {
                }
            }
            if (!focusDone) {
                if (!paneWasOpen) {
                    Send "^i"
                    Sleep 1200
                }
                return false
            }
        }
        return true
    } catch {
        return false
    }
}

; Minimum clipboard length for transfer (match Gemini/bridge validation)
CURSOR_TRANSFER_MIN_CLIPBOARD_LENGTH := 10
; Electron/Cursor needs time after paste before Enter; after Enter before foreground changes or submit can drop.
CURSOR_TRANSFER_POST_PASTE_BEFORE_ENTER_MS := 80
CURSOR_TRANSFER_POST_ENTER_BEFORE_RESTORE_MS := 400

; Activate Cursor window, focus AI field, paste clipboard, send Enter. Non-blocking feedback on failure.
; restoreFocusHwnd: if set, WinActivate this window after Enter is processed (before success overlay) so focus does not stay on the target Cursor window.
CursorTransfer_ActivateFocusPaste(targetHwnd, restoreFocusHwnd := 0) {
    if (!targetHwnd || !WinExist("ahk_id " targetHwnd)) {
        ShowCenteredOverlay_Utils("❌ Cursor window not found", 2000, BANNER_ACCENT_ERROR)
        return
    }
    clip := Trim(A_Clipboard)
    if (clip = "" || StrLen(clip) < CURSOR_TRANSFER_MIN_CLIPBOARD_LENGTH) {
        ShowCenteredOverlay_Utils("❌ Clipboard empty or too short", 2000, BANNER_ACCENT_ERROR)
        return
    }
    try {
        WinActivate("ahk_id " targetHwnd)
        if (!WinWaitActive("ahk_id " targetHwnd, , 2)) {
            ShowCenteredOverlay_Utils("❌ Could not activate Cursor", 2000, BANNER_ACCENT_ERROR)
            return
        }
        Sleep 100
        if (!Cursor_FocusAITextField(targetHwnd)) {
            ShowCenteredOverlay_Utils("❌ Could not focus AI field", 2000, BANNER_ACCENT_ERROR)
            return
        }
        try {
            if (WinGetID("A") != targetHwnd) {
                WinActivate("ahk_id " targetHwnd)
                WinWaitActive("ahk_id " targetHwnd, , 2)
            }
        } catch {
        }
        if (Trim(A_Clipboard) = "" || StrLen(Trim(A_Clipboard)) < CURSOR_TRANSFER_MIN_CLIPBOARD_LENGTH) {
            ShowCenteredOverlay_Utils("❌ Clipboard lost before paste", 2000, BANNER_ACCENT_ERROR)
            return
        }
        SendInput "^v"
        Sleep CURSOR_TRANSFER_POST_PASTE_BEFORE_ENTER_MS
        SendInput "{Enter}"
        if (restoreFocusHwnd && WinExist("ahk_id " restoreFocusHwnd)) {
            ; Wait so Cursor keeps foreground until paste + Enter are processed; restoring sooner drops Enter.
            Sleep CURSOR_TRANSFER_POST_ENTER_BEFORE_RESTORE_MS
            try {
                WinActivate("ahk_id " restoreFocusHwnd)
                if (!WinActive("ahk_id " restoreFocusHwnd))
                    WinWaitActive("ahk_id " restoreFocusHwnd, , 0.5)
            } catch {
            }
        }
        ShowCenteredOverlay_Utils("✅ Sent to Cursor", 1500, BANNER_ACCENT_SUCCESS)
    } catch as err {
        ShowCenteredOverlay_Utils("❌ Transfer failed", 2000, BANNER_ACCENT_ERROR)
    }
}

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

; Small banner for Clip Angel (uses standard loading indicator).
ClipAngelBanner_Show(text, bgColor := BANNER_ACCENT_INTERMEDIATE) {
    StandardLoadingBar_Show(text, bgColor, { passive: true, centerOnHwnd: 0, textWidth: 200, fontSize: 17,
        passiveBgColor: bgColor, alpha: 220 })
}

ClipAngelBanner_Hide() {
    StandardLoadingBar_Hide(0)
}

; Fast Copy Mode (Shift keys — Clip Angel sequential paste): persistent banner with live copy count.
FastCopyModeBanner_Show() {
    StandardLoadingBar_Show("📋 Fast Copy Mode — copies: 0", BANNER_ACCENT_INFO, { passive: true, centerOnHwnd: 0,
        textWidth: 480, fontSize: 17, passiveBgColor: BANNER_ACCENT_INFO, alpha: 220,
        promptKeys: "[Win+Alt+Shift+J] Finish and paste", trackActiveMonitor: true })
}

FastCopyModeBanner_Update(copyCount) {
    StandardLoadingBar_Update("📋 Fast Copy Mode — copies: " copyCount)
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
        Sleep 500

        ; Step 1b: Optional — set Cohere language on General (explicit English / Portuguese)
        if (modelInfo.HasProp("cohereLanguage") && modelInfo.cohereLanguage != "") {
            AiModelBanner_Show("🌐 Setting Cohere language: " . modelInfo.cohereLanguage . "...")
            if (!Handy_SetCohereLanguage(handyHwnd, modelInfo.cohereLanguage)) {
                AiModelBanner_Show("❌ Could not set Cohere language", "E74C3C")
                Sleep 2000
                AiModelBanner_Hide()
                return
            }
            Sleep 400
        }

        ; Step 2: Open AI model menu
        AiModelBanner_Show("📋 Opening AI model menu...")
        if (!Handy_OpenAiModelMenu(handyHwnd)) {
            AiModelBanner_Show("❌ Could not open AI menu", "E74C3C")
            Sleep 2000
            AiModelBanner_Hide()
            return
        }
        Sleep 800

        ; Step 3: Select the model
        AiModelBanner_Show("🎯 Selecting " . modelDisplayName . "...")
        if (!Handy_ClickAiModel(handyHwnd, modelClickName)) {
            AiModelBanner_Show("❌ Model not found: " . modelClickName, "E74C3C")
            Sleep 2000
            AiModelBanner_Hide()
            return
        }

        ; Step 4: Wait for model to finish loading (poll button name until "loading" disappears)
        AiModelBanner_Show("⏳ Waiting for model...", BANNER_ACCENT_INTERMEDIATE)
        Handy_WaitForModelReady(handyHwnd, 20000)

        ; Step 4.5: Play confirmation sound when model is ready
        if (IsSoundEnabled()) {
            soundPath := A_ScriptDir . "\sounds\handy-model-chosen.mp3"
            if (FileExist(soundPath)) {
                try SoundPlay(soundPath)
            }
        }

        ; Step 5: Close Handy window
        AiModelBanner_Show("✅ Done! Closing Handy...", BANNER_ACCENT_SUCCESS)
        try WinClose("ahk_id " . handyHwnd)
        Sleep 500

        AiModelBanner_Hide()

    } catch Error as e {
        AiModelBanner_Show("❌ Error: " . e.Message, "E74C3C")
        Sleep 2000
        AiModelBanner_Hide()
    }
}

; =============================================================================
; Handy UIA Helper Functions
; =============================================================================

; Activate existing Handy window or launch it; returns hwnd or 0
Handy_ActivateOrLaunch() {
    targetPath := GetHandyShortcutPath()
    expectedExePath := GetHandyProcessPath()

    ; Find existing Handy window
    matchingHwnd := 0
    for hwnd in WinGetList("Handy ahk_class Tauri Window") {
        try {
            procPath := WinGetProcessPath(hwnd)
            if (expectedExePath = "" || StrCompare(procPath, expectedExePath, false) = 0) {
                matchingHwnd := hwnd
                break
            }
        } catch {
            if (expectedExePath = "") {
                matchingHwnd := hwnd
                break
            }
        }
    }

    if (matchingHwnd) {
        WinActivate("ahk_id " . matchingHwnd)
        WinWaitActive("ahk_id " . matchingHwnd, , 2)
        return matchingHwnd
    }

    ; Launch Handy
    if (targetPath = "" || !FileExist(targetPath))
        return 0

    Run targetPath
    if !WinWait("Handy ahk_class Tauri Window", , 5)
        return 0

    ; Find the window we just launched
    for h in WinGetList("Handy ahk_class Tauri Window") {
        try {
            procPath := WinGetProcessPath(h)
            if (expectedExePath = "" || StrCompare(procPath, expectedExePath, false) = 0) {
                WinActivate("ahk_id " . h)
                WinWaitActive("ahk_id " . h, , 2)
                return h
            }
        } catch {
            if (expectedExePath = "") {
                WinActivate("ahk_id " . h)
                WinWaitActive("ahk_id " . h, , 2)
                return h
            }
        }
    }
    return 0
}

; True when General tab content (COHERE SETTINGS) is visible.
Handy_GeneralTabVisible(el) {
    if !el
        return false
    try {
        return el.FindFirst({ Type: 50020, Name: "COHERE SETTINGS" }) != 0
    } catch {
        return false
    }
}

; Click sidebar "General" so COHERE SETTINGS is shown (needed from Models/About/etc.).
Handy_EnsureGeneralTab(hwnd) {
    el := UIA.ElementFromHandle(hwnd)
    if !el
        return false
    if (Handy_GeneralTabVisible(el))
        return true
    try {
        gen := el.FindFirst({ Type: 50020, Name: "General" })
        if gen {
            try gen.Click()
            catch {
                try gen.Invoke()
            }
            Sleep 450
            el2 := UIA.ElementFromHandle(hwnd)
            return Handy_GeneralTabVisible(el2)
        }
    } catch {
    }
    return false
}

; Language dropdown under COHERE SETTINGS: class uses "rounded min-w-[200px]" (Microphone uses rounded-md).
Handy_FindHandyLanguageButton(el) {
    if !el
        return 0
    try {
        buttons := el.FindAll({ Type: 50000 })
        for btn in buttons {
            cn := ""
            try cn := btn.ClassName
            if (cn != "" && InStr(cn, "rounded min-w-[200px]"))
                return btn
        }
    } catch {
    }
    return 0
}

; With language dropdown open: focus search, type langName, choose row or Enter.
Handy_SetCohereLanguage_PickFromOpenDropdown(hwnd, langName) {
    el := UIA.ElementFromHandle(hwnd)
    if !el
        return
    searchEl := 0
    try {
        for ed in el.FindAll({ Type: UIA.Type.Edit }) {
            searchEl := ed
            break
        }
    } catch {
    }
    if (searchEl) {
        try {
            searchEl.SetFocus()
        } catch {
            try searchEl.Click()
        }
        Sleep 80
    }
    Send "^a"
    SendText langName
    Sleep 280
    picked := false
    try {
        for btn in el.FindAll({ Type: 50000 }) {
            n := ""
            try n := btn.Name
            if (n != langName)
                continue
            cn := ""
            try cn := btn.ClassName
            if (InStr(cn, "w-full px-3 py-2 text-left") || InStr(cn, "w-full px-3 py-2 text-start")) {
                try btn.Click()
                picked := true
                break
            }
        }
    } catch {
    }
    if !picked
        Send "{Enter}"
    Sleep 300
}

; Set Cohere transcription language on General tab (explicit list pick, not Auto Detect).
Handy_SetCohereLanguage(hwnd, langName) {
    el := UIA.ElementFromHandle(hwnd)
    if !el || langName = ""
        return false
    if !Handy_EnsureGeneralTab(hwnd)
        return false
    el := UIA.ElementFromHandle(hwnd)
    langBtn := Handy_FindHandyLanguageButton(el)
    if !langBtn
        return false
    cur := ""
    try cur := langBtn.Name
    if (cur = langName)
        return true

    try langBtn.Click()
    catch {
        try langBtn.Invoke()
    }
    Sleep 400
    Handy_SetCohereLanguage_PickFromOpenDropdown(hwnd, langName)

    el := UIA.ElementFromHandle(hwnd)
    if !el
        return false
    langBtn2 := Handy_FindHandyLanguageButton(el)
    if !langBtn2
        return false
    n2 := ""
    try n2 := langBtn2.Name
    if (n2 = langName)
        return true

    ; Retry once: close stray popup then reopen
    Send "{Escape}"
    Sleep 200
    el := UIA.ElementFromHandle(hwnd)
    if !el
        return false
    langBtn3 := Handy_FindHandyLanguageButton(el)
    if !langBtn3
        return false
    try langBtn3.Click()
    catch {
        try langBtn3.Invoke()
    }
    Sleep 400
    Handy_SetCohereLanguage_PickFromOpenDropdown(hwnd, langName)

    el := UIA.ElementFromHandle(hwnd)
    if !el
        return false
    langBtn4 := Handy_FindHandyLanguageButton(el)
    if !langBtn4
        return false
    n4 := ""
    try n4 := langBtn4.Name
    return n4 = langName
}

; Open the AI model dropdown menu using keyboard navigation
Handy_OpenAiModelMenu(hwnd) {
    el := UIA.ElementFromHandle(hwnd)
    if !el {
        return false
    }

    ; Find anchor: primary "Check for updates" button
    anchor := 0
    try anchor := el.FindFirst({
        Type: 50000,
        ClassName: "transition-colors disabled:opacity-50 tabular-nums text-text/60 hover:text-text/80"
    })
    if (!anchor) {
        try anchor := el.FindFirst({ Type: 50000, Name: "Check for updates" })
    }
    if (!anchor) {
        try anchor := el.FindFirst({ Type: 50000, Name: "Verificar atualizações" })
    }

    ; Fallback: "Update available" anchor when a system update banner is shown
    if (!anchor) {
        try anchor := el.FindFirst({
            Type: 50000,
            ClassName: "transition-colors disabled:opacity-50 tabular-nums text-logo-primary hover:text-logo-primary/80 font-medium"
        })
    }
    if (!anchor) {
        try anchor := el.FindFirst({ Type: 50000, Name: "Update available" })
    }
    if (!anchor) {
        ; Last-resort: use technical condition path to reach the "Update available" button
        try anchor := el.ElementFromPath({ T: 33 }, { T: 33 }, { T: 33 }, { T: 33, CN: "BrowserRootView" }, { T: 33 }, { T: 33,
            CN: "EmbeddedBrowserFrameView" }, { T: 33, CN: "BrowserView" }, { T: 33, CN: "SidebarContentsSplitView" }, { T: 33 }, { T: 33 }, { T: 33 }, { T: 30 }, { T: 26 }, { T: 0,
                CN: "transition-colors disabled:opacity-50 tabular-nums text-logo-primary hover:text-logo-primary/80 font-medium" }
        )
    }

    if (!anchor) {
        return false
    }

    ; Focus anchor, Shift+Tab to model button, Enter to open menu
    try anchor.SetFocus()
    catch {
        try anchor.Click()
    }
    Sleep 100
    Send "+{Tab}"
    Sleep 100
    Send "{Enter}"
    Sleep 300
    return true
}

; Find and click the AI model button by partial name match
Handy_ClickAiModel(hwnd, modelName) {
    el := UIA.ElementFromHandle(hwnd)
    if !el {
        return false
    }

    ; Model buttons have class containing "w-full px-3 py-2 text-left"
    ; and names starting with the model name (e.g., "Whisper Large Good accuracy...")
    ; Try to find by partial name match
    modelBtn := 0
    buttonCount := 0
    nameMatchNoClass := ""

    ; Strategy 1: Find button whose Name starts with modelName
    try {
        buttons := el.FindAll({ Type: 50000 })
        for btn in buttons {
            buttonCount++
            btnName := ""
            try btnName := btn.Name
            if (btnName != "" && InStr(btnName, modelName) = 1) {
                btnClass := ""
                try btnClass := btn.ClassName
                ; Menu items: w-full px-3 py-2 text-left (legacy) or text-start (new Handy UI); header: flex items-center gap-2
                if (InStr(btnClass, "w-full px-3 py-2 text-left") || InStr(btnClass, "w-full px-3 py-2 text-start") ||
                InStr(btnClass, "flex items-center gap-2")) {
                    modelBtn := btn
                    break
                }
                if (nameMatchNoClass = "")
                    nameMatchNoClass := btnClass
            }
        }
    }

    if (!modelBtn)
        return false

    ; Click the model button
    try {
        modelBtn.Click()
        return true
    } catch as e {
        return false
    }
}

; Poll the AI model selection button until Name no longer contains "loading", or maxWaitMs elapses.
; Button: Type 50000, ClassName "flex items-center gap-2 hover:text-text/80 transition-colors "
; Returns true when loading text disappeared, false on timeout or if button not found.
Handy_WaitForModelReady(hwnd, maxWaitMs) {
    global UIA
    pollInterval := 250
    start := A_TickCount
    firstLog := true
    loop {
        if ((A_TickCount - start) >= maxWaitMs) {
            return false
        }
        el := UIA.ElementFromHandle(hwnd)
        if !el {
            Sleep pollInterval
            continue
        }
        btn := 0
        try btn := el.FindFirst({ Type: 50000, ClassName: "flex items-center gap-2 hover:text-text/80 transition-colors " })
        if (!btn) {
            if (firstLog) {
                firstLog := false
            }
            Sleep pollInterval
            continue
        }
        btnName := ""
        try btnName := btn.Name
        if (InStr(btnName, "loading") = 0) {
            return true
        }
        Sleep pollInterval
    }
}

; =============================================================================
; SelectAiModelInHandy() - Opens or closes the selector GUI (hotkey entry point)
; =============================================================================
; Select AI model in Handy via interactive GUI selector.
; Win+Alt+Shift+C toggles: open when closed, close when open.
; Targets the correct Handy instance by environment: work = Documents\Handy\handy.exe, home = any.
SelectAiModelInHandy() {
    global g_AiModelSelectorActive
    if (g_AiModelSelectorActive)
        AiModelSelector_Close()
    else
        ShowAiModelSelector()
}

; =============================================================================
; Helper: Show centered overlay banner (uses standard loading indicator; non-blocking).
; =============================================================================
ShowCenteredOverlay_Utils(text, duration := 1500, bgColor := BANNER_ACCENT_INTERMEDIATE) {
    ; centerOnHwnd 0 = foreground monitor (GetActiveMonitorWorkArea_StandardBar); same intent as prior WinGetID("A") path.
    StandardLoadingBar_Show(text, bgColor, { passive: true, centerOnHwnd: 0, textWidth: 500, fontSize: 17,
        passiveBgColor: bgColor })
    StandardLoadingBar_Hide(duration)
}

; =============================================================================
; Helper: Pre-movement warning (sound + 2s delay) before automated window changes.
; =============================================================================
PlayPreMovementWarning(targetName) {
    if (IsSoundEnabled()) {
        try SoundPlay(A_ScriptDir . "\sounds\pre-movement.wav")
    }
    ShowCenteredOverlay_Utils("✋ Hands off! Moving to " . targetName . "...", 2000, BANNER_ACCENT_INTERMEDIATE)
    Sleep 2000
}

; =============================================================================
; Standard loading bar (monitor-aware, show/update/hide lifecycle)
; Use for long-running shortcuts; replace ad-hoc banners/overlays with this.
; Supports passive (text-only) mode and ShowWithKeys for letter-keystroke commands.
; Semantic accent colors (colorblind accessibility): border only; background stays dark.
; =============================================================================
global BANNER_ACCENT_SUCCESS := "27AE60"      ; Dark green: positive / success
global BANNER_ACCENT_ERROR := "C0392B"        ; Red: negative / error
global BANNER_ACCENT_INTERMEDIATE := "F1C40F" ; Yellow: loading, actionable, neutral
global BANNER_ACCENT_INFO := "2980B9"         ; Blue: distinct from green/yellow for color vision (info / alternate mode)
global g_StandardLoadingBarGui := 0
global g_StandardLoadingBarValue := 0
global g_StandardLoadingBarIsKeysOverlay := false
global g_StandardLoadingBarKeysHotkeys := []
global g_StandardLoadingBarKeysTimeoutTimer := ""
global g_StandardLoadingBarBorderGui := 0
global g_StandardLoadingBarTrackTimer := ""
global g_StandardLoadingBarLastForegroundMonitorIdx := 0

; Return work area { left, top, right, bottom } for the monitor containing hwnd, or "" on failure.
GetWorkAreaForWindow_StandardBar(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return ""
    try {
        WinGetPos(&winX, &winY, &winW, &winH, "ahk_id " hwnd)
        centerX := winX + winW / 2
        centerY := winY + winH / 2
        n := MonitorGetCount()
        loop n {
            MonitorGet(A_Index, &L, &T, &R, &B)
            if (centerX >= L && centerX < R && centerY >= T && centerY < B) {
                MonitorGetWorkArea(A_Index, &wLeft, &wTop, &wRight, &wBottom)
                return { left: wLeft, top: wTop, right: wRight, bottom: wBottom }
            }
        }
    } catch {
    }
    return ""
}

GetActiveMonitorWorkArea_StandardBar(&left, &top, &right, &bottom) {
    left := top := 0
    right := A_ScreenWidth
    bottom := A_ScreenHeight
    activeWin := 0
    try activeWin := WinGetID("A")
    catch
        activeWin := 0
    MonitorGetWorkArea(1, &monitorLeft, &monitorTop, &monitorRight, &monitorBottom)
    mLeft := monitorLeft
    mTop := monitorTop
    mRight := monitorRight
    mBottom := monitorBottom
    if (activeWin && activeWin != 0) {
        rect := Buffer(16, 0)
        if (DllCall("GetWindowRect", "ptr", activeWin, "ptr", rect)) {
            winLeft := NumGet(rect, 0, "int")
            winTop := NumGet(rect, 4, "int")
            winRight := NumGet(rect, 8, "int")
            winBottom := NumGet(rect, 12, "int")
            centerX := winLeft + (winRight - winLeft) // 2
            centerY := winTop + (winBottom - winTop) // 2
            monitorCount := MonitorGetCount()
            loop monitorCount {
                idx := A_Index
                MonitorGetWorkArea(idx, &l, &t, &r, &b)
                if (centerX >= l && centerX <= r && centerY >= t && centerY <= b) {
                    mLeft := l
                    mTop := t
                    mRight := r
                    mBottom := b
                    break
                }
            }
        }
    }
    left := mLeft
    top := mTop
    right := mRight
    bottom := mBottom
}

; 1-based monitor index for the monitor containing the center of the foreground window; 1 if unknown.
GetMonitorIndexForForeground_StandardBar() {
    activeWin := 0
    try activeWin := WinGetID("A")
    catch
        activeWin := 0
    if (!activeWin)
        return 1
    rect := Buffer(16, 0)
    if (!DllCall("GetWindowRect", "ptr", activeWin, "ptr", rect))
        return 1
    winLeft := NumGet(rect, 0, "int")
    winTop := NumGet(rect, 4, "int")
    winRight := NumGet(rect, 8, "int")
    winBottom := NumGet(rect, 12, "int")
    centerX := winLeft + (winRight - winLeft) // 2
    centerY := winTop + (winBottom - winTop) // 2
    monitorCount := MonitorGetCount()
    loop monitorCount {
        idx := A_Index
        MonitorGetWorkArea(idx, &l, &t, &r, &b)
        if (centerX >= l && centerX <= r && centerY >= t && centerY <= b)
            return idx
    }
    return 1
}

StandardLoadingBar_StopActiveMonitorTracking() {
    global g_StandardLoadingBarTrackTimer, g_StandardLoadingBarLastForegroundMonitorIdx
    try SetTimer(g_StandardLoadingBarTrackTimer, 0)
    catch {
    }
    g_StandardLoadingBarTrackTimer := ""
    g_StandardLoadingBarLastForegroundMonitorIdx := 0
}

StandardLoadingBar_RepositionToActiveMonitor() {
    global g_StandardLoadingBarGui, g_StandardLoadingBarBorderGui
    if !IsObject(g_StandardLoadingBarGui)
        return
    GetActiveMonitorWorkArea_StandardBar(&ml, &mt, &mr, &mb)
    monitorWidth := mr - ml
    try {
        g_StandardLoadingBarGui.GetPos(, , &gw, &gh)
    } catch {
        return
    }
    guiX := Round(ml + (monitorWidth - gw) / 2)
    if (guiX < ml)
        guiX := ml
    if (guiX + gw > mr)
        guiX := mr - gw
    guiY := mt + 40
    if (IsObject(g_StandardLoadingBarBorderGui)) {
        borderWidth := 6
        try {
            g_StandardLoadingBarBorderGui.Move(guiX - borderWidth, guiY - borderWidth, gw + 2 * borderWidth, gh + 2 *
                borderWidth)
        } catch {
        }
    }
    try {
        g_StandardLoadingBarGui.Move(guiX, guiY)
        hwnd := g_StandardLoadingBarGui.Hwnd
        if (hwnd)
            DllCall("SetWindowPos", "Ptr", hwnd, "Ptr", 0, "Int", guiX, "Int", guiY, "Int", 0, "Int", 0, "UInt", 0x0015
            )
    } catch {
    }
}

StandardLoadingBar_TrackActiveMonitorTick() {
    global g_StandardLoadingBarGui, g_StandardLoadingBarLastForegroundMonitorIdx
    if !IsObject(g_StandardLoadingBarGui) {
        StandardLoadingBar_StopActiveMonitorTracking()
        return
    }
    newIdx := GetMonitorIndexForForeground_StandardBar()
    if (newIdx != g_StandardLoadingBarLastForegroundMonitorIdx) {
        g_StandardLoadingBarLastForegroundMonitorIdx := newIdx
        StandardLoadingBar_RepositionToActiveMonitor()
    }
}

StandardLoadingBar_Show(state := "Working...", barColor := BANNER_ACCENT_INTERMEDIATE, options := "") {
    global g_StandardLoadingBarGui, g_StandardLoadingBarValue, g_StandardLoadingBarIsKeysOverlay,
        g_StandardLoadingBarBorderGui
    try StandardLoadingBar_CloseKeysOverlay()
    try StandardLoadingBar_Hide(0)
    passive := options && options.HasProp("passive") && options.passive
    centerOnHwnd := options && options.HasProp("centerOnHwnd") ? options.centerOnHwnd : 0
    textWidth := options && options.HasProp("textWidth") ? options.textWidth : 0
    fontSize := options && options.HasProp("fontSize") ? options.fontSize : 17
    alpha := options && options.HasProp("alpha") ? options.alpha : 235
    passiveBgColor := options && options.HasProp("passiveBgColor") ? options.passiveBgColor : ""
    noBorder := options && options.HasProp("noBorder") ? options.noBorder : false
    promptKeys := options && options.HasProp("promptKeys") ? options.promptKeys : ""
    trackActiveMonitor := options && options.HasProp("trackActiveMonitor") && options.trackActiveMonitor

    if (centerOnHwnd) {
        workArea := GetWorkAreaForWindow_StandardBar(centerOnHwnd)
        if (workArea != "") {
            ml := workArea.left
            mt := workArea.top
            mr := workArea.right
            mb := workArea.bottom
        } else
            GetActiveMonitorWorkArea_StandardBar(&ml, &mt, &mr, &mb)
    } else
        GetActiveMonitorWorkArea_StandardBar(&ml, &mt, &mr, &mb)
    monitorWidth := mr - ml
    monitorHeight := mb - mt
    barWidth := textWidth > 0 ? textWidth : Min(900, Max(360, Floor(monitorWidth * 0.6)))
    overlayGui := Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale")
    overlayGui.BackColor := "1E1E2E"
    overlayGui.MarginX := 16
    overlayGui.MarginY := 10
    overlayGui.SetFont("s" . fontSize . " cFFFFFF", "Segoe UI")
    overlayGui.Add("Text", "w" . barWidth . (passive ? " Wrap Center" : " Center"), state)
    if (promptKeys != "") {
        overlayGui.SetFont("s" . fontSize . " cFFFFFF", "Segoe UI")
        overlayGui.Add("Text", "xm w" . barWidth . " Center", promptKeys)
    }
    if (!passive) {
        progressOpts := "w" . barWidth . " h10 c" . barColor . " Background45475A Smooth vOverlayProg"
        overlayGui.Add("Progress", progressOpts, 0)
    }
    overlayGui.Show("AutoSize Hide")
    overlayGui.GetPos(, , &gw, &gh)
    guiX := Round(ml + (monitorWidth - gw) / 2)
    if (guiX < ml)
        guiX := ml
    if (guiX + gw > mr)
        guiX := mr - gw
    guiY := mt + 40

    ; Create border frame behind the overlay for visibility (optional; skip when noBorder to show a single banner). Accent color when passiveBgColor set, else yellow.
    if (!noBorder) {
        borderWidth := 6
        try {
            if IsObject(g_StandardLoadingBarBorderGui)
                g_StandardLoadingBarBorderGui.Destroy()
        } catch {
        }
        borderGui := Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale")
        borderGui.BackColor := (passiveBgColor != "") ? passiveBgColor : BANNER_ACCENT_INTERMEDIATE
        borderGui.Show("NA x" . (guiX - borderWidth) . " y" . (guiY - borderWidth) . " w" . (gw + 2 * borderWidth) .
        " h" .
        (gh + 2 * borderWidth))
        g_StandardLoadingBarBorderGui := borderGui
    } else {
        try {
            if IsObject(g_StandardLoadingBarBorderGui)
                g_StandardLoadingBarBorderGui.Destroy()
        } catch {
        }
        g_StandardLoadingBarBorderGui := 0
    }
    overlayGui.Show("x" . guiX . " y" . guiY . " NA")
    try {
        hwnd := overlayGui.Hwnd
        if (hwnd)
            DllCall("SetWindowPos", "Ptr", hwnd, "Ptr", 0, "Int", guiX, "Int", guiY, "Int", 0, "Int", 0, "UInt", 0x0015
            )
    }
    WinSetTransparent(alpha, overlayGui)
    g_StandardLoadingBarGui := overlayGui
    g_StandardLoadingBarValue := 0
    if (!passive)
        SetTimer(StandardLoadingBar_Tick, 40)
    if (trackActiveMonitor) {
        StandardLoadingBar_StopActiveMonitorTracking()
        g_StandardLoadingBarLastForegroundMonitorIdx := GetMonitorIndexForForeground_StandardBar()
        g_StandardLoadingBarTrackTimer := SetTimer(StandardLoadingBar_TrackActiveMonitorTick, 115)
    }
}

StandardLoadingBar_Tick() {
    global g_StandardLoadingBarGui, g_StandardLoadingBarValue
    if !IsObject(g_StandardLoadingBarGui) {
        SetTimer(StandardLoadingBar_Tick, 0)
        return
    }
    try {
        g_StandardLoadingBarValue += 4
        if (g_StandardLoadingBarValue > 100)
            g_StandardLoadingBarValue := 0
        g_StandardLoadingBarGui["OverlayProg"].Value := g_StandardLoadingBarValue
    } catch {
        SetTimer(StandardLoadingBar_Tick, 0)
    }
}

StandardLoadingBar_Update(state := "", barColor := "") {
    global g_StandardLoadingBarGui
    if !IsObject(g_StandardLoadingBarGui)
        return
    try {
        if (state != "" && g_StandardLoadingBarGui.Controls.Length > 0)
            g_StandardLoadingBarGui.Controls[1].Text := state
    } catch {
    }
    if (barColor != "") {
        try
            g_StandardLoadingBarGui["OverlayProg"].Opt("c" . barColor)
        catch {
        }
    }
}

StandardLoadingBar_Hide(delayMs := 0) {
    global g_StandardLoadingBarGui, g_StandardLoadingBarValue, g_StandardLoadingBarIsKeysOverlay,
        g_StandardLoadingBarBorderGui
    ; #region agent log
    DebugFlowLog("Utils.ahk:StandardLoadingBar_Hide", "entry", "delay=" . delayMs . " isKeys=" . (
        g_StandardLoadingBarIsKeysOverlay ? 1 : 0), "H2")
    ; #endregion
    if (delayMs > 0) {
        SetTimer(() => StandardLoadingBar_Hide(0), -delayMs)
        return
    }
    StandardLoadingBar_StopActiveMonitorTracking()
    if (g_StandardLoadingBarIsKeysOverlay) {
        StandardLoadingBar_CloseKeysOverlay()
        return
    }
    SetTimer(StandardLoadingBar_Tick, 0)
    try {
        if IsObject(g_StandardLoadingBarGui)
            g_StandardLoadingBarGui.Destroy()
    } catch {
    }
    g_StandardLoadingBarGui := 0
    g_StandardLoadingBarValue := 0
    try {
        if IsObject(g_StandardLoadingBarBorderGui)
            g_StandardLoadingBarBorderGui.Destroy()
    } catch {
    }
    g_StandardLoadingBarBorderGui := 0
}

; Unregister keys and timeout timer for the "ShowWithKeys" overlay, then hide. Idempotent.
StandardLoadingBar_CloseKeysOverlay() {
    global g_StandardLoadingBarKeysHotkeys, g_StandardLoadingBarKeysTimeoutTimer
    global g_StandardLoadingBarGui, g_StandardLoadingBarValue, g_StandardLoadingBarIsKeysOverlay,
        g_StandardLoadingBarBorderGui
    g_StandardLoadingBarIsKeysOverlay := false
    try SetTimer(g_StandardLoadingBarKeysTimeoutTimer, 0)
    catch {
    }
    g_StandardLoadingBarKeysTimeoutTimer := ""
    StandardLoadingBar_StopActiveMonitorTracking()
    for key in g_StandardLoadingBarKeysHotkeys {
        try Hotkey(key, "Off")
        catch {
        }
    }
    g_StandardLoadingBarKeysHotkeys := []
    SetTimer(StandardLoadingBar_Tick, 0)
    try {
        if IsObject(g_StandardLoadingBarGui)
            g_StandardLoadingBarGui.Destroy()
    } catch {
    }
    g_StandardLoadingBarGui := 0
    g_StandardLoadingBarValue := 0
    try {
        if IsObject(g_StandardLoadingBarBorderGui)
            g_StandardLoadingBarBorderGui.Destroy()
    } catch {
    }
    g_StandardLoadingBarBorderGui := 0
}

; Show passive overlay and register hotkeys; optional timeout. keyCallbacks: Map/object key -> callback (e.g. "N" -> fn, "R" -> fn).
; timeoutCallback: called when timeout fires (can be empty). Registers both upper and lower case for letter keys.
; passiveBgColor: optional; when set, used as border color. Prefer BANNER_ACCENT_SUCCESS / BANNER_ACCENT_ERROR / BANNER_ACCENT_INTERMEDIATE. Overlay background stays dark.
; noBorder: when true, do not create the yellow border (single banner only).
; promptKeys: optional; fixed bottom strip text (e.g. "[Y] Confirm  [N] Cancel"). Shown in uniform position below main message.
; trackActiveMonitor: when true, reposition the bar to follow the foreground window's monitor while visible (dictation/Gemini flows).
StandardLoadingBar_ShowWithKeys(state, keyCallbacks, timeoutMs := 0, centerOnHwnd := 0, timeoutCallback := "", barColor :=
    BANNER_ACCENT_INTERMEDIATE, textWidth := 500, fontSize := 17, passiveBgColor := "", noBorder := false, promptKeys :=
    "", trackActiveMonitor := false) {
    global g_StandardLoadingBarIsKeysOverlay, g_StandardLoadingBarKeysHotkeys, g_StandardLoadingBarKeysTimeoutTimer
    opts := { passive: true, centerOnHwnd: centerOnHwnd, textWidth: textWidth, fontSize: fontSize }
    if (passiveBgColor != "")
        opts.passiveBgColor := passiveBgColor
    if (noBorder)
        opts.noBorder := true
    if (promptKeys != "")
        opts.promptKeys := promptKeys
    if (trackActiveMonitor)
        opts.trackActiveMonitor := true
    StandardLoadingBar_Show(state, barColor, opts)
    g_StandardLoadingBarIsKeysOverlay := true
    g_StandardLoadingBarKeysHotkeys := []

    ; Register primary and case-variant keys (single letters get * prefix in RegisterKeyHandler for reliable firing).
    for keyName, cb in keyCallbacks {
        if (!cb)
            continue
        StandardLoadingBar_RegisterKeyHandler(keyName, cb)
        if (StrLen(keyName) = 1) {
            o := Ord(keyName)
            alt := ""
            if (o >= Ord("a") && o <= Ord("z"))
                alt := StrUpper(keyName)
            else if (o >= Ord("A") && o <= Ord("Z"))
                alt := StrLower(keyName)
            if (alt != "" && alt != keyName)
                StandardLoadingBar_RegisterKeyHandler(alt, cb)
        }
    }

    if (timeoutMs > 0) {
        g_StandardLoadingBarKeysTimeoutTimer := SetTimer(StandardLoadingBar_KeysTimeoutFired.Bind(timeoutCallback), -
        timeoutMs)
    }
}

StandardLoadingBar_RegisterKeyHandler(key, cb) {
    global g_StandardLoadingBarKeysHotkeys
    if (!cb)
        return
    ; Use * prefix for single letters so the hotkey fires even when overlay has focus or modifiers are held
    keyToReg := key
    if (StrLen(key) = 1) {
        o := Ord(key)
        if ((o >= Ord("a") && o <= Ord("z")) || (o >= Ord("A") && o <= Ord("Z")))
            keyToReg := "*" . key
    }
    fn := StandardLoadingBar_KeyWrapper.Bind(key, cb)
    try {
        Hotkey(keyToReg, fn, "On")
        g_StandardLoadingBarKeysHotkeys.Push(keyToReg)
    } catch as err {
    }
}

StandardLoadingBar_KeyWrapper(key, cb, *) {
    ; Run callback first so it can close the overlay (avoids destroying GUI from hotkey context before callback runs).
    if (cb) {
        try {
            cb.Call()
        }
        catch {
        }
    }
    StandardLoadingBar_CloseKeysOverlay()
}

StandardLoadingBar_KeysTimeoutFired(timeoutCallback) {
    global g_StandardLoadingBarIsKeysOverlay
    ; Only run timeout callback if overlay was not already dismissed (e.g. user pressed N); avoids copy when timer fires after cancel.
    if (g_StandardLoadingBarIsKeysOverlay && timeoutCallback) {
        DebugFlowLog("Utils.ahk:KeysTimeoutFired", "calling timeout callback", "", "H2")
        try timeoutCallback.Call()
        catch {
        }
    } else
        DebugFlowLog("Utils.ahk:KeysTimeoutFired", "skipped callback", "isKeys=" . (g_StandardLoadingBarIsKeysOverlay ?
            1 : 0), "H2")
    StandardLoadingBar_CloseKeysOverlay()
}

; =============================================================================
; Hotstring Selector: Gemini Redirect Banner (non-blocking; uses standard loading indicator)
; =============================================================================
HotstringGeminiBanner_Show(text := "📤 Gemini: inserting prompt...") {
    ; #region agent log
    DebugFlowLog("Utils.ahk:HotstringGeminiBanner_Show", "entry", "text=" . SubStr(text, 1, 40), "H3")
    ; #endregion
    StandardLoadingBar_CloseKeysOverlay()
    Sleep 50
    StandardLoadingBar_Show(text, BANNER_ACCENT_INTERMEDIATE, { passive: true, centerOnHwnd: 0, textWidth: 280,
        fontSize: 17,
        alpha: 204 })
}

HotstringGeminiBanner_Hide(*) {
    StandardLoadingBar_Hide(0)
}

; Dictation → Gemini: join preset prompt and dictated text for InsertText into Gemini prompt field.
D2C_CombinePresetWithDictation(presetText, dictationText) {
    p := Trim(presetText)
    d := Trim(dictationText)
    if (p = "")
        return d
    if (d = "")
        return p
    return p . "`n`n" . d
}

; =============================================================================
; =============================================================================
; D2C_FlowManager: Unified state machine for Dictation → Gemini → Cursor flow.
; Replaces legacy fragmented functions with a central authority to prevent race conditions.
; =============================================================================
class D2C_FlowManager {
    static _instance := 0

    static GetInstance() {
        if (!D2C_FlowManager._instance)
            D2C_FlowManager._instance := D2C_FlowManager()
        return D2C_FlowManager._instance
    }

    __New() {
        this.Reset()
    }

    Reset() {
        this.CurrentPhase := "Idle"
        this.OriginHwnd := 0
        this.GeminiHwnd := 0
        this.CursorHwnd := 0
        this.MonitorTimer := ""
        this.MonitorRetryCount := 0
        this.MonitorMaxRetries := 300 ; 150s
        this.MonitorButtonEverFound := false
        this.MonitorLastCheckTick := 0
        this.HasCopiedForThisResponse := false
    }

    ; --- Entry Points ---

    StartFromDictation() {
        if (this.CurrentPhase != "Idle") {
            return
        }
        this.Reset()
        this.OriginHwnd := WinActive("A")
        this.PromptForGeminiSubmit()
    }

    StartFromHotstring() {
        if (this.CurrentPhase != "Idle") {
            return
        }
        this.Reset()
        this.OriginHwnd := WinActive("A")
        this.ExecuteGeminiSubmit(true)
    }

    ; --- Phase 1: Submit Prompt ---

    PromptForGeminiSubmit() {
        this.CurrentPhase := "PromptingSubmit"
        keyCallbacks := Map(
            "G", this.OnSubmitG.Bind(this),
            "A", this.OnSubmitA.Bind(this),
            "Y", this.OnSubmitY.Bind(this),
            "S", this.OnSubmitS.Bind(this),
            "V", this.OnSubmitV.Bind(this),
            "E", this.OnSubmitE.Bind(this),
            "N", this.OnSubmitN.Bind(this)
        )
        StandardLoadingBar_ShowWithKeys(
            "❓ Send to Gemini? (6s)",
            keyCallbacks,
            6000,
            0,
            this.OnSubmitTimeout.Bind(this),
            "1E1E2E", 440, 17, "", true,
            "[G] Grammar  [A] AI opt  [Y] Send  [S] Paste only  [V] Paste dictated  [E] Paste & send  [N] Cancel",
            true
        )
    }

    OnSubmitG(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return
        this.ExecuteGeminiSubmit(true, "grammar")
    }

    OnSubmitA(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return
        this.ExecuteGeminiSubmit(true, "aiopt")
    }

    OnSubmitY(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return
        this.ExecuteGeminiSubmit(true)
    }

    OnSubmitS(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return
        this.ExecuteGeminiSubmit(false)
    }

    OnSubmitV(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return

        this.CurrentPhase := "PastingDictation"
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        HideDictationIndicator()

        if (this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd))
            WinActivate("ahk_id " this.OriginHwnd)
        Sleep 60
        Send("^v")

        this.Reset()
    }

    ; Paste clipboard at caret in the original window, then Enter (no Gemini). For chat-style fields confident transcription.
    OnSubmitE(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return

        this.CurrentPhase := "PastingSendHere"
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        HideDictationIndicator()

        if (this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd))
            WinActivate("ahk_id " this.OriginHwnd)
        Sleep 60
        Send("^v")
        Sleep 150
        Send("{Enter}")

        this.Reset()
    }

    OnSubmitN(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return
        this.CancelFlow("Gemini submission cancelled")
    }

    OnSubmitTimeout(*) {
        if (this.CurrentPhase != "PromptingSubmit")
            return
        this.ExecuteGeminiSubmit(true)
    }

    ; --- Phase 2: Submit Execute ---

    ; presetMode: "" = Clip Angel first snippet; "grammar" | "aiopt" = preset from prompt/*.txt + clipboard dictation via InsertText.
    ExecuteGeminiSubmit(autoSubmit := true, presetMode := "") {
        this.CurrentPhase := "Submitting"
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        HideDictationIndicator()

        ; Pre-movement warning before activating Gemini for paste (Original → Gemini).
        PlayPreMovementWarning("Gemini")

        ; Paste to Gemini (launches Chrome if needed); then capture active window as Gemini.
        optionalSnippet := ""
        if (presetMode = "grammar" || presetMode = "aiopt") {
            dictation := ""
            try dictation := A_Clipboard
            preset := presetMode = "grammar" ? GetGrammarPromptText() : GetAioptPromptText()
            optionalSnippet := D2C_CombinePresetWithDictation(preset, dictation)
        }
        if (optionalSnippet != "")
            GeminiNavigateFocusAndPasteFirstSnippet(optionalSnippet, false)
        else
            GeminiNavigateFocusAndPasteFirstSnippet("", false)
        this.GeminiHwnd := WinExist("A")

        if (autoSubmit) {
            Sleep 1000 ; Pre-enter delay
            ; Wait for content (guarantee layer)
            endTick := A_TickCount + 5000
            while (A_TickCount < endTick) {
                if (GeminiPromptFieldGetText() != "")
                    break
                Sleep 200
            }
            Send("{Enter}")
            this.StartGeminiMonitor()
        }

        ; Return focus (Gemini → Original): no pre-movement warning on return.
        if (this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd))
            WinActivate("ahk_id " this.OriginHwnd)

        if (!autoSubmit)
            this.Reset()
    }

    ; --- Phase 3: Monitor ---

    StartGeminiMonitor() {
        this.CurrentPhase := "Monitoring"
        this.MonitorRetryCount := 0
        this.MonitorButtonEverFound := false
        this.MonitorTimer := this.CheckGeminiCompletion.Bind(this)
        SetTimer(this.MonitorTimer, 500)
    }

    CheckGeminiCompletion() {
        delta := this.MonitorLastCheckTick ? (A_TickCount - this.MonitorLastCheckTick) : -1
        this.MonitorLastCheckTick := A_TickCount
        if (this.CurrentPhase != "Monitoring") {
            SetTimer(this.MonitorTimer, 0)
            return
        }

        this.MonitorRetryCount++
        if (this.MonitorRetryCount > this.MonitorMaxRetries) {
            SetTimer(this.MonitorTimer, 0)
            this.Reset()
            return
        }

        btn := ""
        buttonNames := ["Stop streaming", "Interromper transmissão", "Stop response"]
        try {
            root := UIA.ElementFromHandle(this.GeminiHwnd)
            for n in buttonNames {
                try {
                    btn := root.FindElement({ Name: n, Type: "Button" })
                } catch {
                    btn := ""
                }
                if (btn)
                    break
            }
        } catch {
            return
        }

        if (btn) {
            this.MonitorButtonEverFound := true
            return
        }

        if (this.MonitorButtonEverFound) {
            ; Suspend timer to prevent re-entrancy during the 800ms Sleep block
            SetTimer(this.MonitorTimer, 0)

            isTrulyGone := true
            loop 4 {
                Sleep 200
                try {
                    for n in buttonNames {
                        if root.ElementExist({ Name: n, Type: "Button" }) {
                            isTrulyGone := false
                            break
                        }
                    }
                } catch {
                    isTrulyGone := true
                }
                if (!isTrulyGone)
                    break
            }

            if (isTrulyGone) {
                ; Timer is already stopped, proceed to next phase
                try {
                    if (IsSoundEnabled())
                        SoundPlay(A_ScriptDir . "\sounds\gemini-completion.wav")
                } catch {
                    ; Ignore chime failures
                }
                this.PromptForResponseAction()
            } else {
                ; False alarm, the button is still there. Resume polling.
                SetTimer(this.MonitorTimer, 500)
            }
        }
    }

    ; --- Phase 4: Action Prompt ---

    PromptForResponseAction() {
        ; Prevent duplicate banner spawns from timer re-entrancy
        if (this.CurrentPhase = "PromptingAction") {
            return
        }
        this.CurrentPhase := "PromptingAction"
        keyCallbacks := Map(
            "Y", this.OnActionY.Bind(this),
            "C", this.OnActionC.Bind(this),
            "R", this.OnActionR.Bind(this),
            "N", this.OnActionN.Bind(this)
        )
        StandardLoadingBar_ShowWithKeys(
            "❓ Copy response?",
            keyCallbacks,
            5000,
            0,
            this.OnActionTimeout.Bind(this),
            BANNER_ACCENT_INTERMEDIATE, 380, 17, "", false,
            "[Y] Copy  [N] No  [R] Copy+Read  [C] Transfer",
            true
        )
    }

    OnActionY(*) {
        if (this.CurrentPhase != "PromptingAction")
            return
        this.ExecuteAction(false, false)
    }

    OnActionC(*) {
        if (this.CurrentPhase != "PromptingAction")
            return
        this.PromptForCursorTransfer()
    }

    OnActionR(*) {
        if (this.CurrentPhase != "PromptingAction")
            return
        this.ExecuteAction(true, false)
    }

    OnActionN(*) {
        if (this.CurrentPhase != "PromptingAction")
            return
        this.CleanupActionPrompt()
        this.Reset()
    }

    OnActionTimeout(*) {
        if (this.CurrentPhase != "PromptingAction")
            return
        this.ExecuteAction(false, false)
    }

    CleanupActionPrompt() {
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
    }

    ExecuteAction(readAloud := false, skipRestoreFocus := false) {
        this.CleanupActionPrompt()
        try {
            this.DoCopyCore(readAloud, skipRestoreFocus)
        } finally {
            this.Reset()
        }
    }

    DoCopyCore(readAloud := false, skipRestoreFocus := false) {
        if (this.HasCopiedForThisResponse) {
            return
        }
        this.HasCopiedForThisResponse := true

        ; Hands off cue before copying Gemini's last response (applies to both manual Y/R/C and timeout auto-copy).
        PlayPreMovementWarning("Gemini")

        if (!WinExist("ahk_id " this.GeminiHwnd)) {
            if (this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd))
                WinActivate("ahk_id " this.OriginHwnd)
            return
        }

        ; By the time DoCopyCore runs, Gemini should already be active. If it is not,
        ; just activate it without a pre-movement warning (source is no longer Original).
        if (!WinActive("ahk_id " this.GeminiHwnd)) {
            try WinActivate("ahk_id " this.GeminiHwnd)
            catch {
                if (this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd))
                    WinActivate("ahk_id " this.OriginHwnd)
                return
            }
            if (!WinWaitActive("ahk_exe chrome.exe", , 0.5)) {
                if (this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd))
                    WinActivate("ahk_id " this.OriginHwnd)
                return
            }
        }

        ; Y / R / C / timeout: same synchronous copy first. R then blocks on read-aloud IPC (wParam=1 skips duplicate Copy in Gemini).
        clipBefore := A_Clipboard
        seqBefore := Clipboard_GetSequenceNumber()
        WM_COPY_LAST_GEMINI := 0x8001
        WM_TRIGGER_READ_ALOUD := 0x8004
        targetHwnd := GetGeminiScriptMsgTargetHwnd()
        sendOk := false
        clipOk := false

        if (targetHwnd) {
            ; lParam = Chrome Gemini hwnd: skip redundant activation in receiver. Timeout 20s (default 5s can abort mid-copy).
            ; Hidden Gemini.ahk main window: SendMessage needs DetectHiddenWindows true or "Target window not found."
            prevDH := A_DetectHiddenWindows
            DetectHiddenWindows true
            try {
                SendMessage(WM_COPY_LAST_GEMINI, 0, this.GeminiHwnd, , "ahk_id " targetHwnd, , , , 20000)
                sendOk := true
            } catch {
            } finally {
                DetectHiddenWindows prevDH
            }
            changed := Clipboard_WaitForSequenceChange(seqBefore, 2000, 850)
            clipOk := sendOk && changed && A_Clipboard != clipBefore && Trim(A_Clipboard) != ""
            if (!sendOk)
                ShowCenteredOverlay_Utils("❌ Gemini copy timed out or IPC failed", 3500, BANNER_ACCENT_ERROR)
            else if (!clipOk)
                ShowCenteredOverlay_Utils("❌ Copy failed or clipboard empty — try again", 3000, BANNER_ACCENT_ERROR)
            else if (readAloud) {
                DetectHiddenWindows true
                try {
                    ; Blocks until Listen flow finishes; wParam 1 => GeminiTriggerReadAloud(false).
                    SendMessage(WM_TRIGGER_READ_ALOUD, 1, 0, , "ahk_id " targetHwnd, , , , 120000)
                } catch {
                    ShowCenteredOverlay_Utils("❌ Read aloud failed or timed out", 4000, BANNER_ACCENT_ERROR)
                } finally {
                    DetectHiddenWindows prevDH
                }
            } else if (IsSoundEnabled())
                try SoundPlay(A_ScriptDir . "\sounds\copy.wav")
        } else {
            ShowCenteredOverlay_Utils("❌ Gemini.ahk not running", 2000, BANNER_ACCENT_ERROR)
        }

        ; Gemini/Clipboard → Original: return transitions are immediate (no warning).
        if (!skipRestoreFocus && this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd) && !WinActive("ahk_id " this.OriginHwnd
        )) {
            WinActivate("ahk_id " this.OriginHwnd)
            ; Fast-path: avoid WinWaitActive if we are already active.
            if (!WinActive("ahk_id " this.OriginHwnd))
                WinWaitActive("ahk_id " this.OriginHwnd, , 0.5)
        }
    }

    ; --- Phase 5: Cursor Transfer ---

    PromptForCursorTransfer() {
        this.CurrentPhase := "Transferring"
        this.CleanupActionPrompt()
        try {
            ; Skip restoring focus so clipboard is not overwritten
            this.DoCopyCore(false, true)

            clipRaw := A_Clipboard
            clip := Trim(clipRaw)
            if (clip = "" || StrLen(clip) < 10) {
                ShowCenteredOverlay_Utils("❌ Copy failed or empty – try again", 2000, BANNER_ACCENT_ERROR)
                if (this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd))
                    WinActivate("ahk_id " this.OriginHwnd)
                return
            }

            ; Restore the pre-handoff anchored window so the user sees the selector/paste
            ; happening in the exact app they were monitoring before the Gemini handoff.
            if (this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd)) {
                try {
                    WinActivate("ahk_id " this.OriginHwnd)
                    ; Fast-path: avoid WinWaitActive if we are already active.
                    if (!WinActive("ahk_id " this.OriginHwnd))
                        WinWaitActive("ahk_id " this.OriginHwnd, , 0.5)
                } catch {
                }
            }
            try A_Clipboard := clipRaw

            tSelectorStart := A_TickCount
            this.CursorHwnd := CursorTransfer_ShowWindowSelector(0)
            tSelectorMs := A_TickCount - tSelectorStart
            if (!this.CursorHwnd) {
                if (this.OriginHwnd && WinExist("ahk_id " this.OriginHwnd))
                    WinActivate("ahk_id " this.OriginHwnd)
                try A_Clipboard := clipRaw
                return
            }

            ; Gemini → Cursor: no pre-movement warning (source is not Original).
            try A_Clipboard := clipRaw
            CursorTransfer_ActivateFocusPaste(this.CursorHwnd, this.OriginHwnd)
        } finally {
            this.Reset()
        }
    }

    ; --- Helpers ---

    CancelFlow(message) {
        StandardLoadingBar_CloseKeysOverlay()
        StandardLoadingBar_Hide(0)
        ShowCenteredOverlay_Utils("⚠ " . message, 1500, BANNER_ACCENT_INTERMEDIATE)
        this.Reset()
    }
}

; Dictation: "Send to Gemini?" confirmation banner (6s, Y to confirm; uses standard loading indicator)
; =============================================================================
; DEPRECATED: Use D2C_FlowManager
DEPRECATED_DictationGeminiConfirm_Show() {
    ; No-op; use DictationGeminiConfirm_ShowAndWait() which uses StandardLoadingBar_ShowWithKeys.
}

DEPRECATED_DictationGeminiConfirm_Hide(*) {
    StandardLoadingBar_CloseKeysOverlay()
}

; submitToGemini=false (N or timeout): terminal. submitToGemini=true: delayed-submit (paste+Enter). pasteOnly=true: paste to Gemini only, no Enter.
DEPRECATED_DictationGeminiConfirm_CleanupAndMaybeSubmit(submitToGemini, pasteOnly := false) {
    global g_DictationGeminiConfirmBannerVisible
    ; #region agent log
    DebugFlowLog("Utils.ahk:CleanupAndMaybeSubmit", "entry", "submit=" . (submitToGemini ? 1 : 0) . " pasteOnly=" . (
        pasteOnly ? 1 : 0), "H2")
    ; #endregion
    ; Only clear banner-visible when proceeding (Y/S/timeout); leave true on N so a stray ShowAndWait does not re-show and register a second 6s timer (logs showed second timeout firing after N).
    if (submitToGemini || pasteOnly)
        g_DictationGeminiConfirmBannerVisible := false
    ; Unregister 6s banner keys (same * prefix as StandardLoadingBar_RegisterKeyHandler uses)
    try Hotkey("*y", "Off")
    try Hotkey("*Y", "Off")
    try Hotkey("*s", "Off")
    try Hotkey("*S", "Off")
    try Hotkey("*n", "Off")
    try Hotkey("*N", "Off")
    SetTimer(DEPRECATED_DictationGeminiConfirm_OnTimeout, 0)
    DEPRECATED_DictationGeminiConfirm_Hide()
    ; S or N at 6s: no submit flow, so stop any running "Copy response?" monitor so it never shows.
    if (!submitToGemini)
        GeminiDelayedSubmitMonitorStopFromUtils()
    if (pasteOnly) {
        ; #region agent log
        DebugFlowLog("Utils.ahk:CleanupAndMaybeSubmit", "running pasteOnly flow", "", "H2")
        ; #endregion
        Sleep 350
        DEPRECATED_GeminiDictationPasteOnlyFlow()
    } else if (submitToGemini) {
        ; #region agent log
        DebugFlowLog("Utils.ahk:CleanupAndMaybeSubmit", "running delayedSubmit flow", "", "H2")
        ; #endregion
        Sleep 350
        DEPRECATED_GeminiDelayedSubmitFlow()
    }
}

DEPRECATED_DictationGeminiConfirm_OnY(*) {
    ; #region agent log
    DebugFlowLog("Utils.ahk:OnY", "Y pressed", "", "H1")
    ; #endregion
    DEPRECATED_DictationGeminiConfirm_CleanupAndMaybeSubmit(true)
}

; S = paste to Gemini only (no Enter, no 4s banner).
DEPRECATED_DictationGeminiConfirm_OnS(*) {
    ; #region agent log
    DebugFlowLog("Utils.ahk:OnS", "S pressed", "", "H1")
    ; #endregion
    DEPRECATED_DictationGeminiConfirm_CleanupAndMaybeSubmit(false, true)
}

; Default action on 6s timeout: proceed as Yes (DelayedSubmitFlow), same as user pressing Y.
DEPRECATED_DictationGeminiConfirm_OnTimeout(*) {
    ; #region agent log
    DebugFlowLog("Utils.ahk:OnTimeout", "6s timeout fired", "", "H4")
    ; #endregion
    DEPRECATED_DictationGeminiConfirm_CleanupAndMaybeSubmit(true)
}

; N = terminate flow: no paste, no Enter, no 4s, no copy; only cleanup and cancel overlay.
DEPRECATED_DictationGeminiConfirm_OnCancel(*) {
    ; #region agent log
    DebugFlowLog("Utils.ahk:OnCancel", "N pressed", "", "H1")
    ; #endregion
    DEPRECATED_DictationGeminiConfirm_CleanupAndMaybeSubmit(false)  ; submitToGemini=false, pasteOnly=false => no flow runs
    ShowCenteredOverlay_Utils("⚠ Gemini submission cancelled", 1500, BANNER_ACCENT_INTERMEDIATE)
}

DEPRECATED_DictationGeminiConfirm_ShowAndWait() {
    global g_DictationGeminiConfirmBannerVisible
    ; Only one banner: atomic check-and-set so only one invocation can pass (prevents duplicate from multiple PlayDictationCompletionChime runs).
    Critical "On"
    if (g_DictationGeminiConfirmBannerVisible) {
        Critical "Off"
        return
    }
    g_DictationGeminiConfirmBannerVisible := true
    Critical "Off"
    ; #region agent log
    DebugFlowLog("Utils.ahk:ShowAndWait", "6s banner showing", "", "H1")
    ; #endregion
    ; Only the official loading bar (standard loading indicator) may show this content. Hide any other bar/overlay first.
    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Hide(0)
    HideDictationIndicator()
    Sleep 50
    keyCallbacks := Map("Y", DEPRECATED_DictationGeminiConfirm_OnY, "S", DEPRECATED_DictationGeminiConfirm_OnS, "N",
        DEPRECATED_DictationGeminiConfirm_OnCancel)
    ; Official loading bar only; no blue; single banner (no border); fixed bottom strip for input.
    StandardLoadingBar_ShowWithKeys("❓ Send to Gemini? (6s)", keyCallbacks,
        6000,
        0,
        DEPRECATED_DictationGeminiConfirm_OnTimeout, "1E1E2E", 380, 17, "", true,
        "[Y] Send  [S] Paste only  [N] Cancel",
        true)
}

; =============================================================================
; Global Sound Toggle System
; File-backed state management for muting/unmuting sounds across all scripts
; =============================================================================

; Check if sound is enabled (reads from INI file for cross-process persistence)
IsSoundEnabled() {
    settingsFile := A_ScriptDir . "\data\settings.ini"
    ; Default to enabled (1) if file doesn't exist or key is missing
    soundEnabled := IniRead(settingsFile, "Settings", "SoundEnabled", "1")
    return (soundEnabled = "1")
}

; Toggle sound state and show visual feedback
ToggleSoundState() {
    settingsFile := A_ScriptDir . "\data\settings.ini"
    currentState := IsSoundEnabled()
    newState := currentState ? "0" : "1"

    ; Update INI file
    IniWrite(newState, settingsFile, "Settings", "SoundEnabled")

    ; Show visual feedback
    if (newState = "1") {
        ShowCenteredOverlay_Utils("🔊 Sound: ON", 2000, BANNER_ACCENT_INTERMEDIATE)
    } else {
        ShowCenteredOverlay_Utils("🔇 Sound: OFF", 2000, BANNER_ACCENT_INTERMEDIATE)
    }
}

; =============================================================================
; Outlook: classic OUTLOOK.EXE and Microsoft Store "new" Outlook (olk.exe)
; =============================================================================
OutlookGetOlkExePath() {
    candidate :=
        "C:\Program Files\WindowsApps\Microsoft.OutlookForWindows_1.2026.317.100_x64__8wekyb3d8bbwe\olk.exe"
    if FileExist(candidate)
        return candidate
    try {
        loop files "C:\Program Files\WindowsApps\Microsoft.OutlookForWindows_*_x64__8wekyb3d8bbwe\olk.exe", "F" {
            return A_LoopFileFullPath
        }
    } catch {
    }
    return ""
}

OutlookProcessRunning() {
    return ProcessExist("OUTLOOK.EXE") || ProcessExist("olk.exe")
}

; =============================================================================
; Toggle Outlook and Teams
; Toggles Outlook and Teams applications to manage RAM usage.
; If both are open: Closes Outlook and minimizes Teams to system tray.
; If one or both are closed: Launches both applications.
; =============================================================================
ToggleOutlookAndTeams() {
    try {
        ; Check if both applications are running
        outlookRunning := OutlookProcessRunning()
        teamsRunning := ProcessExist("ms-teams.exe")

        ; Show start banner
        if (outlookRunning && teamsRunning) {
            ShowCenteredOverlay_Utils("📤 Closing Outlook and Teams...", 1500, BANNER_ACCENT_INTERMEDIATE)
        } else {
            ShowCenteredOverlay_Utils("📤 Opening Outlook and Teams...", 1500, BANNER_ACCENT_INTERMEDIATE)
        }

        if (outlookRunning && teamsRunning) {
            ; Both are open: Close Outlook and minimize Teams to system tray
            ; Close Outlook process(es) — classic and/or Store (olk.exe)
            try {
                if ProcessExist("OUTLOOK.EXE")
                    ProcessClose("OUTLOOK.EXE")
                if ProcessExist("olk.exe")
                    ProcessClose("olk.exe")
            } catch Error as e {
                MsgBox "Error closing Outlook: " e.Message
            }

            ; Close all Teams windows (this keeps Teams in system tray)
            try {
                ; Teams can have multiple process names, check all
                for hwnd in WinGetList("ahk_exe ms-teams.exe") {
                    WinClose(hwnd)
                }
                ; Also check for Teams.exe and MSTeams.exe variants
                for hwnd in WinGetList("ahk_exe Teams.exe") {
                    WinClose(hwnd)
                }
                for hwnd in WinGetList("ahk_exe MSTeams.exe") {
                    WinClose(hwnd)
                }
            } catch Error as e {
                MsgBox "Error closing Teams windows: " e.Message
            }
        } else {
            ; One or both are closed: Launch both applications
            ; Launch Outlook
            if (!outlookRunning) {
                try {
                    outlookPath := ""
                    if (IS_WORK_ENVIRONMENT) {
                        ; Try work environment shortcut path
                        outlookPath := "C:\Users\fie7ca\Documents\Atalhos\Microsoft Outlook.lnk"
                        if (!FileExist(outlookPath)) {
                            outlookPath := ""
                        }
                    } else {
                        ; Try personal environment shortcut path
                        outlookPath :=
                            "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Microsoft Outlook.lnk"
                        if (!FileExist(outlookPath)) {
                            outlookPath := ""
                        }
                    }

                    ; Launch using shortcut if available, otherwise olk.exe or OUTLOOK.EXE
                    if (outlookPath != "") {
                        Run outlookPath
                    } else {
                        olkPath := OutlookGetOlkExePath()
                        if (olkPath != "")
                            Run olkPath
                        else
                            Run "OUTLOOK.EXE"
                    }
                } catch Error as e {
                    MsgBox "Error launching Outlook: " e.Message
                }
            }

            ; Launch/Activate Teams
            ; Simplified approach: Just run the executable. This handles both launching and bringing to front.
            try {
                if (IS_WORK_ENVIRONMENT) {
                    teamsExePath :=
                        "C:\Program Files\WindowsApps\MSTeams_25332.1210.4188.1171_x64__8wekyb3d8bbwe\ms-teams.exe"
                    if (FileExist(teamsExePath)) {
                        Run teamsExePath
                    } else {
                        Run "ms-teams.exe"
                    }
                } else {
                    ; Personal environment
                    teamsPath :=
                        "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Microsoft Teams.lnk"
                    if (FileExist(teamsPath)) {
                        Run teamsPath
                    } else {
                        Run "ms-teams.exe"
                    }
                }

                ; Wait for window to appear and become active
                if (WinWaitActive("ahk_exe ms-teams.exe", , 10)) {
                    ShowCenteredOverlay_Utils("✅ Teams activated", 1500, BANNER_ACCENT_SUCCESS)
                } else {
                    ShowCenteredOverlay_Utils("❌ Teams: Window not found", 2000, BANNER_ACCENT_ERROR)
                }
            } catch Error as e {
                ShowCenteredOverlay_Utils("❌ Teams: Error - " . e.Message, 2000, BANNER_ACCENT_ERROR)
            }

            ; Second: Activate Outlook last (so it gets final focus)
            try {
                if (OutlookProcessRunning()) {
                    ex := ProcessExist("OUTLOOK.EXE") ? "OUTLOOK.EXE" : "olk.exe"
                    WinWait("ahk_exe " ex, , 5)
                    if (!WinExist("ahk_exe " ex)) {
                        ShowCenteredOverlay_Utils("❌ Outlook not running.", 2000, BANNER_ACCENT_ERROR)
                        return
                    } else {
                        WinActivate("ahk_exe " ex)
                        WinWaitActive("ahk_exe " ex, , 2)
                    }
                }
            } catch Error as e {
                ; Silently fail if activation doesn't work
            }
        }

        ; Show finish banner
        ShowCenteredOverlay_Utils("✅ Done", 1500, BANNER_ACCENT_SUCCESS)
    } catch Error as e {
        MsgBox "Error in ToggleOutlookAndTeams macro: " e.Message
    }
}

; =============================================================================
; Check and Prompt to Open Outlook/Teams
; Checks if Outlook or Teams are closed and prompts user to open them if needed
; Parameters:
;   - checkOutlook: true to check Outlook, false otherwise
;   - checkTeams: true to check Teams, false otherwise
; Returns: true if applications are running (or were opened), false if user cancelled
; =============================================================================
CheckAndOpenOutlookTeams(checkOutlook := false, checkTeams := false) {
    outlookClosed := false
    teamsClosed := false

    ; Check Outlook status (classic OUTLOOK.EXE or Store new Outlook olk.exe)
    if (checkOutlook) {
        outlookRunning := OutlookProcessRunning()
        if (!outlookRunning) {
            outlookClosed := true
        }
    }

    ; Check Teams status
    if (checkTeams) {
        teamsRunning := ProcessExist("ms-teams.exe")
        if (!teamsRunning) {
            teamsClosed := true
        }
    }

    ; If both are open, no action needed
    if (!outlookClosed && !teamsClosed) {
        return true
    }

    ; Build message based on what's closed
    message := ""
    if (outlookClosed && teamsClosed) {
        message := "Outlook and Teams are closed. Do you want to open them?"
    } else if (outlookClosed) {
        message := "Outlook is closed. Do you want to open it?"
    } else if (teamsClosed) {
        message := "Teams is closed. Do you want to open it?"
    }

    ; Show message box
    response := MsgBox(message, "Open Applications?", "YesNo Icon?")

    ; If user confirms, open the applications (only open, don't toggle)
    if (response = "Yes") {
        ; Only open the closed applications, don't toggle
        try {
            ; Launch Outlook if closed
            if (outlookClosed) {
                outlookPath := ""
                if (IS_WORK_ENVIRONMENT) {
                    outlookPath := "C:\Users\fie7ca\Documents\Atalhos\Microsoft Outlook.lnk"
                    if (!FileExist(outlookPath)) {
                        outlookPath := ""
                    }
                } else {
                    outlookPath :=
                        "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Microsoft Outlook.lnk"
                    if (!FileExist(outlookPath)) {
                        outlookPath := ""
                    }
                }

                if (outlookPath != "") {
                    Run outlookPath
                } else {
                    olkPath := OutlookGetOlkExePath()
                    if (olkPath != "")
                        Run olkPath
                    else
                        Run "OUTLOOK.EXE"
                }
            }

            ; Launch Teams if closed
            if (teamsClosed) {
                if (IS_WORK_ENVIRONMENT) {
                    teamsExePath :=
                        "C:\Program Files\WindowsApps\MSTeams_25332.1210.4188.1171_x64__8wekyb3d8bbwe\ms-teams.exe"
                    if (FileExist(teamsExePath)) {
                        Run teamsExePath
                    } else {
                        Run "ms-teams.exe"
                    }
                } else {
                    teamsPath :=
                        "C:\Users\eduev\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Microsoft Teams.lnk"
                    if (FileExist(teamsPath)) {
                        Run teamsPath
                    } else {
                        Run "ms-teams.exe"
                    }
                }
            }

            ; Wait a bit for applications to start
            Sleep 2000
            return true
        } catch Error as e {
            MsgBox "Error opening applications: " e.Message, "Error", "IconX"
            return false
        }
    }

    ; User cancelled
    return false
}

; Internal helper: Performs clipboard cleanup without showing prompt
; Used when user has already confirmed they want to clean clipboard
CleanClipboardInternal() {
    Sleep 200

    ; Send Alt+V
    SendInput "!v"

    ; Wait for UI to respond (menu needs time to appear)
    Sleep 600

    ; Send Ctrl+Alt+K
    SendInput "^!k"

    ; Wait for UI to respond (dialog needs time to open)
    Sleep 600

    SendInput "{Enter}"

    ; Wait for UI to respond (processing needs time to complete)
    Sleep 800

    SendInput "!v"

    ; Brief pause to ensure Escape is processed
    Sleep 400
}

; =============================================================================
; Dictation: Non-modal clipboard cleanup countdown (used by Win+Alt+Shift+7)
; =============================================================================
global g_DictationCleanupGui := 0
global g_DictationCleanupBorderGui := 0
global g_DictationCleanupTextCtrl := 0
global g_DictationCleanupRemaining := 0
global g_DictationCleanupCanceled := false

DictationCleanup_SetCancelHotkeys(enable := true) {
    ; Removed ~ prefix to prevent key leakage into active applications
    if (enable) {
        Hotkey("*n", DictationCleanup_Cancel, "On")
        Hotkey("*y", DictationCleanup_Proceed, "On")
        Hotkey("*End", DictationCleanup_Cancel, "On")
    } else {
        try Hotkey("*n", "Off")
        catch {
        }
        try Hotkey("*y", "Off")
        catch {
        }
        try Hotkey("*End", "Off")
        catch {
        }
    }
}

DictationCleanup_ShowBanner() {
    global g_DictationCleanupGui, g_DictationCleanupTextCtrl, g_DictationCleanupRemaining

    ; Destroy any previous banner instance
    try {
        if IsObject(g_DictationCleanupGui)
            g_DictationCleanupGui.Destroy()
    } catch {
    }
    g_DictationCleanupGui := 0
    g_DictationCleanupTextCtrl := 0

    target := WinGetID("A")
    hasWindow := false
    if target && WinExist("ahk_id " target) {
        try {
            WinGetPos(&wx, &wy, &ww, &wh, target)
            hasWindow := (ww > 0 && wh > 0)
        } catch {
            hasWindow := false
        }
    }

    ov := Gui("+AlwaysOnTop -Caption +ToolWindow")
    ov.BackColor := "1E1E2E"
    ov.SetFont("s24 cFFFFFF Bold", "Segoe UI")
    g_DictationCleanupTextCtrl := ov.Add("Text", "w650 Center", "Clearing clipboard in " g_DictationCleanupRemaining "… (press Y to proceed, N or End to cancel)"
    )
    ov.Show("AutoSize Hide")
    ov.GetPos(&gx, &gy, &gw, &gh)

    if hasWindow {
        cx := wx + (ww - gw) // 2
        cy := wy + (wh - gh) // 2
    } else {
        vx := SysGet(76)  ; SM_XVIRTUALSCREEN
        vy := SysGet(77)  ; SM_YVIRTUALSCREEN
        vw := SysGet(78)  ; SM_CXVIRTUALSCREEN
        vh := SysGet(79)  ; SM_CYVIRTUALSCREEN
        cx := vx + (vw - gw) // 2
        cy := vy + (vh - gh) // 2
    }

    borderWidth := 6
    try {
        if IsObject(g_DictationCleanupBorderGui)
            g_DictationCleanupBorderGui.Destroy()
    } catch {
    }
    borderGui := Gui("+AlwaysOnTop -Caption +ToolWindow")
    borderGui.BackColor := BANNER_ACCENT_INTERMEDIATE
    borderGui.Show("NA x" . (cx - borderWidth) . " y" . (cy - borderWidth) . " w" . (gw + 2 * borderWidth) . " h" . (gh +
        2 * borderWidth))
    g_DictationCleanupBorderGui := borderGui

    ov.Show("x" . cx . " y" . cy . " NA")
    WinSetTransparent(178, ov)
    g_DictationCleanupGui := ov
}

DictationCleanup_HideBanner() {
    global g_DictationCleanupGui, g_DictationCleanupBorderGui, g_DictationCleanupTextCtrl
    try {
        if IsObject(g_DictationCleanupBorderGui)
            g_DictationCleanupBorderGui.Destroy()
    } catch {
    }
    g_DictationCleanupBorderGui := 0
    try {
        if IsObject(g_DictationCleanupGui)
            g_DictationCleanupGui.Destroy()
    } catch {
    }
    g_DictationCleanupGui := 0
    g_DictationCleanupTextCtrl := 0
}

DictationCleanup_UpdateBannerText() {
    global g_DictationCleanupTextCtrl, g_DictationCleanupRemaining
    try {
        if IsObject(g_DictationCleanupTextCtrl) {
            g_DictationCleanupTextCtrl.Text := "Clearing clipboard in " g_DictationCleanupRemaining "… (press Y to proceed, N or End to cancel)"
        }
    } catch {
    }
}

DictationCleanup_StartCountdown(seconds := 5) {
    global g_DictationCleanupRemaining, g_DictationCleanupCanceled

    ; Reset state
    g_DictationCleanupCanceled := false
    g_DictationCleanupRemaining := seconds

    ; Show initial banner + enable cancel keys
    DictationCleanup_ShowBanner()
    DictationCleanup_SetCancelHotkeys(true)

    ; Start ticking immediately (1s cadence)
    SetTimer(DictationCleanup_Tick, 0)
    SetTimer(DictationCleanup_Tick, 1000)
}

DictationCleanup_StopCountdown(showCancelledBanner := false) {
    global g_DictationCleanupCanceled

    ; Stop timer + disable cancel keys
    SetTimer(DictationCleanup_Tick, 0)
    DictationCleanup_SetCancelHotkeys(false)

    if (showCancelledBanner) {
        ; Reuse the same banner GUI for a short "cancelled" message (non-blocking)
        global g_DictationCleanupTextCtrl
        try {
            if IsObject(g_DictationCleanupTextCtrl) {
                g_DictationCleanupTextCtrl.Text := "Clipboard cleanup cancelled"
            }
        } catch {
        }
        SetTimer(DictationCleanup_HideBanner, -900)
    } else {
        DictationCleanup_HideBanner()
    }
}

DictationCleanup_Cancel(*) {
    global g_DictationCleanupCanceled
    g_DictationCleanupCanceled := true
    DictationCleanup_StopCountdown(true)
}

DictationCleanup_Proceed(*) {
    ; Immediately proceed with clipboard cleanup, skipping countdown
    DictationCleanup_StopCountdown(false)
    CleanClipboardInternal()
}

DictationCleanup_Tick() {
    global g_DictationCleanupRemaining, g_DictationCleanupCanceled

    ; If already cancelled, ensure everything is stopped
    if (g_DictationCleanupCanceled) {
        DictationCleanup_StopCountdown(true)
        return
    }

    ; Decrement remaining time
    g_DictationCleanupRemaining -= 1

    if (g_DictationCleanupRemaining <= 0) {
        ; Countdown finished -> hide banner and clear clipboard using the existing workflow (Alt+V, Ctrl+Alt+K, etc.)
        DictationCleanup_StopCountdown(false)
        CleanClipboardInternal()
        return
    }

    DictationCleanup_UpdateBannerText()
}

; =============================================================================
; Dictation: Merge non-favorite clips countdown (at end of loop)
; Same UI pattern as clipboard cleanup: 5s banner, N or End to cancel.
; =============================================================================
global g_DictationMergeGui := 0
global g_DictationMergeBorderGui := 0
global g_DictationMergeTextCtrl := 0
global g_DictationMergeRemaining := 0
global g_DictationMergeCanceled := false

DictationMerge_SetCancelHotkeys(enable := true) {
    ; Removed ~ prefix to prevent key leakage into active applications
    if (enable) {
        Hotkey("*n", DictationMerge_Cancel, "On")
        Hotkey("*y", DictationMerge_Proceed, "On")
        Hotkey("*End", DictationMerge_Cancel, "On")
    } else {
        try Hotkey("*n", "Off")
        catch {
        }
        try Hotkey("*y", "Off")
        catch {
        }
        try Hotkey("*End", "Off")
        catch {
        }
    }
}

DictationMerge_ShowBanner() {
    global g_DictationMergeGui, g_DictationMergeTextCtrl, g_DictationMergeRemaining

    try {
        if IsObject(g_DictationMergeGui)
            g_DictationMergeGui.Destroy()
    } catch {
    }
    g_DictationMergeGui := 0
    g_DictationMergeTextCtrl := 0

    target := WinGetID("A")
    hasWindow := false
    if target && WinExist("ahk_id " target) {
        try {
            WinGetPos(&wx, &wy, &ww, &wh, target)
            hasWindow := (ww > 0 && wh > 0)
        } catch {
            hasWindow := false
        }
    }

    ov := Gui("+AlwaysOnTop -Caption +ToolWindow")
    ov.BackColor := "1E1E2E"
    ov.SetFont("s24 cFFFFFF Bold", "Segoe UI")
    g_DictationMergeTextCtrl := ov.Add("Text", "w650 Center", "Merging non-favorite clips in " g_DictationMergeRemaining "… (press Y to proceed, N or End to cancel)"
    )
    ov.Show("AutoSize Hide")
    ov.GetPos(&gx, &gy, &gw, &gh)
    if (hasWindow) {
        cx := wx + (ww - gw) // 2
        cy := wy + (wh - gh) // 2
    } else {
        vx := SysGet(76)
        vy := SysGet(77)
        vw := SysGet(78)
        vh := SysGet(79)
        cx := vx + (vw - gw) // 2
        cy := vy + (vh - gh) // 2
    }
    borderWidth := 6
    try {
        if IsObject(g_DictationMergeBorderGui)
            g_DictationMergeBorderGui.Destroy()
    } catch {
    }
    borderGui := Gui("+AlwaysOnTop -Caption +ToolWindow")
    borderGui.BackColor := BANNER_ACCENT_INTERMEDIATE
    borderGui.Show("NA x" . (cx - borderWidth) . " y" . (cy - borderWidth) . " w" . (gw + 2 * borderWidth) . " h" . (gh +
        2 * borderWidth))
    g_DictationMergeBorderGui := borderGui
    ov.Show("x" . cx . " y" . cy . " NA")
    g_DictationMergeGui := ov
}

DictationMerge_HideBanner() {
    global g_DictationMergeGui, g_DictationMergeBorderGui, g_DictationMergeTextCtrl
    try {
        if IsObject(g_DictationMergeBorderGui)
            g_DictationMergeBorderGui.Destroy()
    } catch {
    }
    g_DictationMergeBorderGui := 0
    try {
        if IsObject(g_DictationMergeGui)
            g_DictationMergeGui.Destroy()
    } catch {
    }
    g_DictationMergeGui := 0
    g_DictationMergeTextCtrl := 0
}

DictationMerge_UpdateBannerText() {
    global g_DictationMergeTextCtrl, g_DictationMergeRemaining
    try {
        if IsObject(g_DictationMergeTextCtrl) {
            g_DictationMergeTextCtrl.Text := "Merging non-favorite clips in " g_DictationMergeRemaining "… (press Y to proceed, N or End to cancel)"
        }
    } catch {
    }
}

DictationMerge_StartCountdown(seconds := 5) {
    global g_DictationMergeRemaining, g_DictationMergeCanceled

    g_DictationMergeCanceled := false
    g_DictationMergeRemaining := seconds

    DictationMerge_ShowBanner()
    DictationMerge_SetCancelHotkeys(true)

    SetTimer(DictationMerge_Tick, 0)
    SetTimer(DictationMerge_Tick, 1000)
}

DictationMerge_StopCountdown(showCancelledBanner := false) {
    SetTimer(DictationMerge_Tick, 0)
    DictationMerge_SetCancelHotkeys(false)

    if (showCancelledBanner) {
        global g_DictationMergeTextCtrl
        try {
            if IsObject(g_DictationMergeTextCtrl) {
                g_DictationMergeTextCtrl.Text := "Merge cancelled"
            }
        } catch {
        }
        SetTimer(DictationMerge_HideBanner, -900)
    } else {
        DictationMerge_HideBanner()
    }
}

DictationMerge_Cancel(*) {
    global g_DictationMergeCanceled
    g_DictationMergeCanceled := true
    DictationMerge_StopCountdown(true)
}

DictationMerge_Proceed(*) {
    ; Immediately proceed with merge, skipping countdown
    DictationMerge_StopCountdown(false)
    MergeNonFavoriteClips()
}

DictationMerge_Tick() {
    global g_DictationMergeRemaining, g_DictationMergeCanceled

    if (g_DictationMergeCanceled) {
        DictationMerge_StopCountdown(true)
        return
    }

    g_DictationMergeRemaining -= 1

    if (g_DictationMergeRemaining <= 0) {
        DictationMerge_StopCountdown(false)
        MergeNonFavoriteClips()
        return
    }

    DictationMerge_UpdateBannerText()
}

; Clean the Clipboard macro function
; Shows a non-modal 4s countdown; auto-continues unless user cancels with N; Y proceeds immediately.
CleanClipboard() {
    CleanClipboard_ShowCountdown()
}

CleanClipboard_ShowCountdown() {
    ; Ensure any previous keys overlay is closed before showing a new one
    try StandardLoadingBar_CloseKeysOverlay()
    catch {
    }
    StandardLoadingBar_Hide(0)
    Sleep 50

    state := "❓ Clean the clipboard? (removes stored clips, 4s)`nPress [Y] to clean now, or [N] within 4s to cancel."
    keyCallbacks := Map("N", CleanClipboard_OnCancel, "Y", CleanClipboard_OnTimeout)

    ; Center on active monitor (centerOnHwnd := 0), use red accent for destructive action.
    StandardLoadingBar_ShowWithKeys(state, keyCallbacks, 4000, 0, CleanClipboard_OnTimeout,
        BANNER_ACCENT_ERROR, 0, 17, "", false, "[Y] Clean now  [N] Cancel (auto-continue in 4s)")
}

CleanClipboard_OnCancel(*) {
    try StandardLoadingBar_CloseKeysOverlay()
    catch {
    }
    StandardLoadingBar_Hide(0)
}

CleanClipboard_OnTimeout(*) {
    try StandardLoadingBar_CloseKeysOverlay()
    catch {
    }
    StandardLoadingBar_Hide(0)
    CleanClipboardInternal()
}

; =============================================================================
; Project Data (for Cursor Window Focus Selector)
; Ported from WindowManagement.ahk to ensure consistent key mapping
; =============================================================================

; Character sequence for assignment: 1 2 3 4 5 q w e r t a s d f g z x c v b 6 7 8 9 0 y u i o p h j k l n m , .
global g_ProjectCharSequence := ["1", "2", "3", "4", "5", "q", "w", "e", "r", "t", "a", "s", "d", "f", "g", "z", "x",
    "c", "v", "b", "6", "7", "8", "9", "0", "y", "u", "i", "o", "p", "h", "j", "k", "l", "n", "m", ",", "."]

; Category display order (General first, Personal second, Work last)
global g_ProjectCategories := ["General", "Personal", "Work"]

; Global project list - must match WindowManagement.ahk for consistent key mapping
; Each project should have: name, path, workPath, and category ("General", "Personal", or "Work")
global g_Projects := [
    ; General category
    { name: "Scripts", path: "C:\Users\eduev\Meu Drive\17 - Projects\scripts", workPath: "C:\Users\fie7ca\Documents\scripts",
        category: "General" }, { name: "14-my-notes", path: "C:\Users\eduev\Meu Drive\17 - Projects\notes", workPath: "C:\Users\fie7ca\OneDrive - Bosch Group\14-my-notes",
            category: "General" }, { name: "", path: "", workPath: "", category: "General" }, { name: "", path: "",
                workPath: "", category: "General" }, { name: "", path: "", workPath: "", category: "General" },
                ; Personal category
                { name: "ZMK Sofle", path: "C:\Users\eduev\Documents\ZMK\zmk-sofle", workPath: "", category: "Personal" }, { name: "AI Experiment",
                    path: "C:\Users\eduev\Meu Drive\04 - Pós-graduação\01 - Mestrado\26-ai-experiment", workPath: "",
                    category: "Personal" }, { name: "my-personal-repo", path: "C:\Users\eduev\Meu Drive\17 - Projects\my-personal-repo",
                        workPath: "C:\Users\fie7ca\OneDrive - Bosch Group\13 - General workspace\my-personal-repo",
                        category: "Personal" }, { name: "",
                            path: "", workPath: "", category: "Personal" }, { name: "", path: "", workPath: "",
                                category: "Personal" },
                            ; Work category
                            { name: "dashboard-model-research", path: "", workPath: "C:\Users\fie7ca\OneDrive - Bosch Group\13 - General workspace\GS_E&S_CIP Dashboard research and design workspace folder\dashboard-model-research",
                                category: "Work" }, { name: "GS_UX core team_UX and CIP Integration", path: "",
                                    workPath: "C:\Users\fie7ca\OneDrive - Bosch Group\13 - General workspace\GS_UX core team_UX and CIP Integration",
                                    category: "Work" }, { name: "🪂 Avante", path: "", workPath: "C:\Users\fie7ca\OneDrive - Bosch Group\General - GS_BDU_Team\00_UX_GS_Team\AM_Planning\Avante",
                                        category: "Work" }, { name: "", path: "", workPath: "", category: "Work" }, { name: "",
                                            path: "", workPath: "", category: "Work" }
]

; Extract matching segments from project path for window title matching
; Cursor window titles have format: "filename - folder-name - Cursor" or "filename - path-segment - Cursor"
ExtractProjectMatchSegments(projectPath) {
    ; Normalize the project path (remove trailing backslashes)
    normalizedPath := RTrim(projectPath, "\")

    ; Split path into segments
    pathSegments := StrSplit(normalizedPath, "\")

    ; Extract the last folder name (e.g., "zmk-sofle", "26-ai-experiment", "scripts")
    lastSegment := pathSegments[pathSegments.Length]

    ; Build list of potential match strings
    matchSegments := [lastSegment]

    ; If we have at least 2 segments, also try the combination
    if (pathSegments.Length >= 2) {
        ; Try last two segments joined with " - " (for cases like "17 - Projects")
        lastTwoJoined := pathSegments[pathSegments.Length - 1] . " - " . pathSegments[pathSegments.Length]
        if (lastTwoJoined != lastSegment) {  ; Only add if different
            matchSegments.Push(lastTwoJoined)
        }
    }

    return matchSegments
}

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
                if (InStr(WinGetTitle("ahk_id " hwnd), "gemini", false) = 0)
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
        if (!IsSoundEnabled())
            return
        if (isWorking)
            SoundPlay(A_ScriptDir . "\sounds\robots-are-working.wav")
        else
            SoundPlay(A_ScriptDir . "\sounds\no-robot-working.wav")
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
    ; Toggle Sound macro
    RegisterMacro(ToggleSoundState, "🔊 Toggle Sound (Mute/Unmute)")
    ; Global AI generation state: Cursor + Gemini (assigned to "U")
    RegisterMacro(Cursor_FindComposerIconAcrossInstances, "🔍 AI working? (Cursor + Gemini)", "u")
    ; Mark Last Clip as Favorite macro (assigned to "J")
    RegisterMacro(MarkLastClipAsFavorite, "⭐ Mark Last Clip as Favorite", "j")
    ; Move Desktop to Recycle Bin (assigned to "N") — red banner, Y/N confirm
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

; Auto-execute: show success after QuickUpdateScripts relaunches Utils.ahk with "/Updated".
if (A_Args.Length > 0 && A_Args[1] = "/Updated") {
    try {
        ShowCenteredOverlay_Utils("✅ Scripts updated and relaunched", 6500, BANNER_ACCENT_SUCCESS)
        soundPath := A_ScriptDir "\sounds\quick-update-success.wav"
        try {
            if (FileExist(soundPath))
                SoundPlay(soundPath)
        } catch {
        }
    } catch {
    }
}

; =============================================================================
; Jump Mouse to Middle of Active Window
; Hotkey: Win+Alt+Shift+3
; Original File: Jump mouse on the middle.ahk
; =============================================================================
#!+Q::
{
    hwnd := WinExist("A")
    if !hwnd {
        return ; silently abort if no active window
    }
    rect := Buffer(16, 0)
    if !DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect) {
        MsgBox "GetWindowRect failed"
        return
    }
    left := NumGet(rect, 0, "int")
    top := NumGet(rect, 4, "int")
    right := NumGet(rect, 8, "int")
    bottom := NumGet(rect, 12, "int")
    centerX := left + (right - left) // 2
    centerY := top + (bottom - top) // 2
    DllCall("SetCursorPos", "int", centerX, "int", centerY)
}

; =============================================================================
; Select AI Model in Handy
; Hotkey: Win+Alt+Shift+C
; =============================================================================
#!+C::
{
    SelectAiModelInHandy()
}

; =============================================================================
; Move all Desktop items to Recycle Bin (recoverable)
; Trigger: Win+Alt+Shift+U selector → letter N
; Target path: OneDrive Desktop. Standard banner with 4s timeout (N = cancel, Y or timeout = run); then success/error banner.
; =============================================================================
global g_DesktopToRecyclePath := ""  ; Set from GetDesktopToRecyclePath() when macro runs
global g_DesktopToRecycleCloseHwnd := 0

DesktopToRecycle_OnConfirm(*) {
    DesktopToRecycle_Run()
}

DesktopToRecycle_OnCancel(*) {
    ShowCenteredOverlay_Utils("⚠ Desktop cleanup cancelled", 1500, BANNER_ACCENT_INTERMEDIATE)
}

DesktopToRecycle_OnTimeout(*) {
    DesktopToRecycle_Run()
}

; Normalize folder path for comparison (trim trailing backslash, lowercase on Windows)
DesktopToRecycle_NormalizePath(p) {
    p := RTrim(p, "\")
    try return StrLower(p)
    return p
}

; Close any Explorer window(s) showing the given folder path (via Shell.Application)
DesktopToRecycle_CloseDesktopExplorer(targetPath) {
    if (!targetPath || targetPath = "")
        return
    normTarget := DesktopToRecycle_NormalizePath(targetPath)
    try {
        shell := ComObject("Shell.Application")
        for window in shell.Windows {
            try {
                if (!window || !window.hwnd)
                    continue
                path := window.Document.Folder.Self.Path
                if (DesktopToRecycle_NormalizePath(path) = normTarget) {
                    window.Quit()
                    return
                }
            } catch
                continue
        }
    } catch {
    }
    ; Fallback: close by hwnd if we had stored it at trigger time
    global g_DesktopToRecycleCloseHwnd
    if (g_DesktopToRecycleCloseHwnd && WinExist("ahk_id " g_DesktopToRecycleCloseHwnd)) {
        try WinClose("ahk_id " g_DesktopToRecycleCloseHwnd)
    }
    g_DesktopToRecycleCloseHwnd := 0
}

DesktopToRecycle_Run() {
    global g_DesktopToRecyclePath, g_DesktopToRecycleCloseHwnd
    ; Play "cleaning desktop" sound immediately when the cleaning starts.
    try {
        soundPath := A_ScriptDir "\sounds\cleaning-desktop.ogg"
        if (FileExist(soundPath))
            SoundPlay(soundPath)
    } catch {
    }
    ; Resolve path: use configured path; if empty or missing, fall back to A_Desktop (works on any PC)
    path := g_DesktopToRecyclePath
    if (!path || path = "" || !DirExist(path))
        path := A_Desktop
    ; Use .NET FileIO.FileSystem SendToRecycleBin (no Shell verbs); process dirs last so parent exists
    ui := "[Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs"
    rec := "[Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin"
    ps := "Add-Type -AssemblyName Microsoft.VisualBasic;$d='" . path .
        "';if(-not(Test-Path -LiteralPath $d)){exit 1};$files=@(Get-ChildItem -LiteralPath $d -Force|Where-Object{-not $_.PSIsContainer});$dirs=@(Get-ChildItem -LiteralPath $d -Force|Where-Object{$_.PSIsContainer});foreach($f in $files){try{[Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($f.FullName," .
        ui . "," . rec .
        ")}catch{}};foreach($dir in $dirs){try{[Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($dir.FullName," .
        ui . "," . rec . ")}catch{}};exit 0"
    try {
        exitCode := RunWait('powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "' . ps . '"', "", "Hide")
        if (exitCode = 0) {
            ShowCenteredOverlay_Utils("✅ Desktop items moved to Recycle Bin", 2000, BANNER_ACCENT_SUCCESS)
            DesktopToRecycle_CloseDesktopExplorer(path)
        } else {
            ShowCenteredOverlay_Utils("❌ Desktop path not found or error: " path, 3500, BANNER_ACCENT_ERROR)
            DesktopToRecycle_CloseDesktopExplorer(path)
        }
    } catch as err {
        ShowCenteredOverlay_Utils("❌ Error moving to Recycle Bin", 2500, BANNER_ACCENT_ERROR)
    }
    g_DesktopToRecycleCloseHwnd := 0
}

; Entry point when "N" is pressed in Win+Alt+Shift+U selector
DesktopToRecycle_Trigger() {
    global g_DesktopToRecycleCloseHwnd, g_DesktopToRecyclePath
    g_DesktopToRecyclePath := GetDesktopToRecyclePath()
    ; Remember active window if it's Explorer showing Desktop - close it after cleaning
    hwnd := WinExist("A")
    g_DesktopToRecycleCloseHwnd := 0
    if (hwnd && WinGetProcessName("ahk_id " hwnd) = "explorer.exe") {
        try {
            if (InStr(WinGetTitle("ahk_id " hwnd), "Desktop", false))
                g_DesktopToRecycleCloseHwnd := hwnd
        } catch {
        }
    }
    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Hide(0)
    Sleep 50
    state := "🗑️ Move all items from:`n" . g_DesktopToRecyclePath . "`nto Recycle Bin? (4s)"
    keyCallbacks := Map("Y", DesktopToRecycle_OnConfirm, "N", DesktopToRecycle_OnCancel)
    ; Center on active monitor (centerOnHwnd := 0), use standard intermediate accent with border.
    StandardLoadingBar_ShowWithKeys(state, keyCallbacks, 4000, 0, DesktopToRecycle_OnTimeout,
        BANNER_ACCENT_INTERMEDIATE, 0, 17, "", false, "[Y] Yes  [N] Cancel")
}

; =============================================================================
; Clip Angel: Open/Activate with focus correction (Row 0)
; Hotkey: Alt+V — when closed: open + focus Row 0; when open: pass Alt+V to close (toggle).
; =============================================================================
!v::
{
    if WinExist("ClipAngel") {
        Send "!v"   ; Already open: close it (Clip Angel toggle)
        return
    }
    ActivateClipAngelWithFocusCorrection()
}

; =============================================================================
; Mouse Jump Shortcuts
; Hotkeys: Win+Alt+Shift+Arrow Keys
; Jump mouse cursor by fixed pixel distance in each direction with multi-monitor support
; =============================================================================

; Set coordinate mode to screen for proper multi-monitor support
CoordMode "Mouse", "Screen"

; Define the movement distance in pixels (increased from 200 to 300)
global MOUSE_JUMP_DISTANCE := 300

; Helper function to get current mouse position using physical screen coordinates
GetMousePos() {
    pt := Buffer(8, 0)
    DllCall("GetCursorPos", "ptr", pt)
    return { x: NumGet(pt, 0, "int"), y: NumGet(pt, 4, "int") }
}

; Helper function to get all monitor information
GetMonitorInfo() {
    ; Get the number of monitors
    monitorCount := SysGet(80)  ; SM_CMONITORS
    monitors := []

    loop monitorCount {
        ; Get work area for each monitor (excludes taskbar)
        MonitorGetWorkArea(A_Index, &left, &top, &right, &bottom)
        monitors.Push({
            left: left,
            top: top,
            right: right,
            bottom: bottom,
            width: right - left,
            height: bottom - top
        })
    }

    return monitors
}

; Helper function: get virtual desktop bounds (supports negative X/Y)
GetVirtualBounds() {
    left := DllCall("GetSystemMetrics", "int", 76)    ; SM_XVIRTUALSCREEN
    top := DllCall("GetSystemMetrics", "int", 77)     ; SM_YVIRTUALSCREEN
    width := DllCall("GetSystemMetrics", "int", 78)   ; SM_CXVIRTUALSCREEN
    height := DllCall("GetSystemMetrics", "int", 79)  ; SM_CYVIRTUALSCREEN
    return { left: left, top: top, right: left + width - 1, bottom: top + height - 1 }
}

Clamp(n, lo, hi) {
    return n < lo ? lo : n > hi ? hi : n
}

; Helper function to find which monitor contains the given coordinates
FindMonitorForCoords(x, y, monitors) {
    for monitor in monitors {
        if (x >= monitor.left && x <= monitor.right && y >= monitor.top && y <= monitor.bottom) {
            return monitor
        }
    }
    return false  ; Not found in any monitor
}

; Helper function to safely move mouse with proper multi-monitor boundary checking
; Always shows both prediction squares (blue for short, red for long) in the direction of movement
SafeMouseMove(deltaX, deltaY) {
    pos := GetMousePos()
    v := GetVirtualBounds()
    ; Calculate target position where mouse will jump to (current + jump distance)
    targetX := Clamp(pos.x + deltaX, v.left, v.right)
    targetY := Clamp(pos.y + deltaY, v.top, v.bottom)

    ; Move the mouse to the target position first
    DllCall("SetCursorPos", "int", targetX, "int", targetY)

    ; After moving, show both prediction squares in the direction of movement
    ; Blue square: shows where mouse will land with short jump (without Control)
    ; Red square: shows where mouse will land with long jump (with Control)
    ShowBothPredictionSquares(targetX, targetY, deltaX, deltaY)
}

; Global array to track all feedback GUI windows
global g_MouseMoveFeedbackGuis := []

; Helper function to close all feedback GUIs
CloseMouseMoveFeedback() {
    global g_MouseMoveFeedbackGuis
    try {
        for gui in g_MouseMoveFeedbackGuis {
            try {
                if (IsObject(gui) && gui.Hwnd) {
                    gui.Destroy()
                }
            } catch {
                ; Silently ignore errors for individual GUIs
            }
        }
        g_MouseMoveFeedbackGuis := []
    } catch {
        ; Silently ignore errors during cleanup
    }
}

; Helper function to show both prediction squares (blue and red) in the direction of movement
; Shows where the mouse will land if user presses short (blue) or long (red) jump in the same direction
ShowBothPredictionSquares(currentX, currentY, deltaX, deltaY) {
    global g_MouseMoveFeedbackGuis
    global MOUSE_JUMP_DISTANCE
    v := GetVirtualBounds()

    ; Close any existing feedback GUIs first
    CloseMouseMoveFeedback()

    ; Determine the direction of movement from the sign of deltaX/deltaY
    ; The squares always use the base MOUSE_JUMP_DISTANCE, regardless of current jump distance

    if (deltaX != 0) {
        ; Horizontal movement - determine direction from sign of deltaX
        directionX := deltaX > 0 ? 1 : -1  ; 1 for right, -1 for left

        ; Blue square: shows where mouse will land with short jump (MOUSE_JUMP_DISTANCE in this direction)
        shortPredictionX := Clamp(currentX + MOUSE_JUMP_DISTANCE * directionX, v.left, v.right)
        ; Red square: shows where mouse will land with long jump (MOUSE_JUMP_DISTANCE * 2 in this direction)
        longPredictionX := Clamp(currentX + MOUSE_JUMP_DISTANCE * 2 * directionX, v.left, v.right)

        ; Show blue square (short distance) in the direction of movement
        ShowPredictionSquare(shortPredictionX, currentY, "0000FF")
        ; Show red square (long distance) in the direction of movement
        ShowPredictionSquare(longPredictionX, currentY, "FF0000")
    } else if (deltaY != 0) {
        ; Vertical movement - determine direction from sign of deltaY
        directionY := deltaY > 0 ? 1 : -1  ; 1 for down, -1 for up

        ; Blue square: shows where mouse will land with short jump (MOUSE_JUMP_DISTANCE in this direction)
        shortPredictionY := Clamp(currentY + MOUSE_JUMP_DISTANCE * directionY, v.top, v.bottom)
        ; Red square: shows where mouse will land with long jump (MOUSE_JUMP_DISTANCE * 2 in this direction)
        longPredictionY := Clamp(currentY + MOUSE_JUMP_DISTANCE * 2 * directionY, v.top, v.bottom)

        ; Show blue square (short distance) in the direction of movement
        ShowPredictionSquare(currentX, shortPredictionY, "0000FF")
        ; Show red square (long distance) in the direction of movement
        ShowPredictionSquare(currentX, longPredictionY, "FF0000")
    }

    ; Auto-hide after 1300ms (1 second longer than before)
    SetTimer(CloseMouseMoveFeedback, -1300)
}

; Helper function to show a single prediction square
ShowPredictionSquare(x, y, color) {
    global g_MouseMoveFeedbackGuis
    squareSize := 40

    ; Create a simple GUI window with specified color background
    squareGui := Gui("+AlwaysOnTop -Caption +ToolWindow")
    squareGui.BackColor := color

    ; Position the square centered at the target position
    guiX := x - (squareSize // 2)
    guiY := y - (squareSize // 2)

    ; Show the square
    squareGui.Show("x" . guiX . " y" . guiY . " w" . squareSize . " h" . squareSize . " NA")
    WinSetTransparent(100, squareGui)  ; Less opaque for better visibility

    ; Store reference for cleanup
    g_MouseMoveFeedbackGuis.Push(squareGui)
}

; =============================================================================
; Square Selector System for Mouse Jump
; Shows 15 red squares with letters in chosen direction, waits for letter selection
; =============================================================================

; Global variables for square selector system
global g_SquareSelectorActive := false
global g_SquareSelectorGuis := []
global g_SquareSelectorPositions := []  ; Array of {x, y} positions for each square
global g_SquareSelectorLetters := ["1", "2", "3", "4", "5", "Q", "W", "E", "R", "T", "A", "S", "D", "F", "G", "Z", "X",
    "C", "V", "B", "6", "7", "8", "9", "0", "y", "u", "i", "o", "p", "h", "j", "k", "l", "n", "m", ",", "."]
global g_SquareSelectorTimer := false
global g_SquareSelectorLetterMap := Map()  ; Map to store letter to index mapping
global g_SquareSelectorSessionID := 0  ; Unique session ID to prevent timer conflicts

; Global array to store hotkey handlers for cleanup
global g_SquareSelectorHotkeyHandlers := []

; Lock flag to prevent multiple square selectors from running simultaneously
global g_SquareSelectorLock := false

; Active direction flag - prevents old selectors from interfering
global g_ActiveDirection := ""

; Loop mode flag - indicates waiting for Escape or arrow key after selection
global g_SquareSelectorLoopMode := false

; Click mode flag - when true, squares are blue and selection will click and exit
global g_SquareSelectorClickMode := false

; Direction indicator GUIs (4 squares around mouse pointer in loop mode)
global g_DirectionIndicatorGuis := []

; Timestamp when squares were last shown (for guaranteed cleanup)
global g_SquareSelectorStartTime := 0

; Backup cleanup timer (guaranteed to fire after 10 seconds)
global g_SquareSelectorBackupTimer := false

; Timer for cleaning up old squares when showing new ones
global g_OldSquaresCleanupTimer := false

; Timer handler for square selector timeout
SquareSelectorTimerHandler(sessionID) {
    global g_SquareSelectorLock, g_ActiveDirection, g_SquareSelectorTimer
    global g_SquareSelectorActive, g_SquareSelectorSessionID

    ; CRITICAL: Check if this timer is for the current session
    ; If session ID doesn't match, this timer is stale and should be ignored
    if (sessionID != g_SquareSelectorSessionID) {
        ; This timer is for an old session, ignore it
        return
    }

    ; Check if selector is still active (might have been cleaned up by new direction)
    if (!g_SquareSelectorActive) {
        ; Already cleaned up, just clear timer reference
        g_SquareSelectorTimer := false
        return
    }

    ; Only cleanup if selector is still active and session matches
    CleanupSquareSelector()
    g_SquareSelectorLock := false
    g_ActiveDirection := ""  ; Clear active direction on timeout
    g_SquareSelectorTimer := false  ; Clear timer reference
}

; Force cleanup function - aggressively destroys all squares regardless of state
; This is a backup mechanism to ensure squares never persist forever
ForceCleanupAllSquares() {
    global g_SquareSelectorGuis, g_DirectionIndicatorGuis
    global g_SquareSelectorActive, g_SquareSelectorLock, g_ActiveDirection
    global g_SquareSelectorLoopMode, g_SquareSelectorClickMode
    global g_SquareSelectorTimer, g_SquareSelectorBackupTimer
    global g_SquareSelectorStartTime

    ; Force disable active flag
    g_SquareSelectorActive := false

    ; Aggressively destroy all square GUIs
    for gui in g_SquareSelectorGuis {
        try {
            if (IsObject(gui)) {
                try {
                    if (gui.Hwnd) {
                        gui.Hide()
                        gui.Destroy()
                    }
                } catch {
                    ; Ignore hide/destroy errors
                }
            }
        } catch {
            ; Silently ignore all errors
        }
    }
    g_SquareSelectorGuis := []

    ; Aggressively destroy all direction indicator GUIs
    DestroyGuiArray(g_DirectionIndicatorGuis)

    ; Cancel all timers
    if (g_SquareSelectorTimer) {
        try {
            SetTimer(g_SquareSelectorTimer, 0)
        } catch {
            ; Ignore
        }
        g_SquareSelectorTimer := false
    }

    if (g_SquareSelectorBackupTimer) {
        try {
            SetTimer(g_SquareSelectorBackupTimer, 0)
        } catch {
            ; Ignore
        }
        g_SquareSelectorBackupTimer := false
    }

    ; Reset all state
    g_SquareSelectorLock := false
    g_ActiveDirection := ""
    g_SquareSelectorLoopMode := false
    g_SquareSelectorClickMode := false
    g_SquareSelectorStartTime := 0

    ; Disable all hotkeys (best effort) to prevent bugs
    try {
        DisableLetterHotkeys()
    } catch {
        ; Ignore
    }
    try {
        DisableDirectionSwitchHotkeys()
    } catch {
        ; Ignore
    }
    try {
        DisableLoopModeHotkeys()
    } catch {
        ; Ignore
    }
    try {
        Hotkey("Ctrl", "Off")
    } catch {
        ; Ignore
    }
    try {
        Hotkey("Escape", "Off")
    } catch {
        ; Ignore
    }
}

; Backup timer handler - guaranteed to fire after 7 seconds
BackupCleanupTimer() {
    global g_SquareSelectorStartTime, g_SquareSelectorGuis, g_SquareSelectorBackupTimer
    global g_SquareSelectorActive

    ; If start time is 0, squares have been cleaned up, stop the timer
    if (g_SquareSelectorStartTime == 0) {
        if (g_SquareSelectorBackupTimer) {
            SetTimer(g_SquareSelectorBackupTimer, 0)
            g_SquareSelectorBackupTimer := false
        }
        return
    }

    ; Check if squares have been visible for more than 7 seconds
    elapsed := (A_TickCount - g_SquareSelectorStartTime) / 1000  ; Convert to seconds
    if (elapsed >= 7) {
        ; Force cleanup if squares have been visible for 7+ seconds
        ForceCleanupAllSquares()
        return
    }

    ; If there are no GUIs and not active, cleanup is done, stop timer
    if (g_SquareSelectorGuis.Length = 0 && !g_SquareSelectorActive) {
        ; No GUIs and not active - cleanup is done, stop timer
        if (g_SquareSelectorBackupTimer) {
            SetTimer(g_SquareSelectorBackupTimer, 0)
            g_SquareSelectorBackupTimer := false
        }
        g_SquareSelectorStartTime := 0
    }
}

; Helper to create a timer handler bound to a specific session ID
CreateTimerHandler(sessionID) {
    return () => SquareSelectorTimerHandler(sessionID)
}

; Helper function to cleanup old square GUIs (used by ShowSquareSelector)
CleanupOldSquareGuis(oldGuis) {
    for gui in oldGuis {
        try {
            if (IsObject(gui) && gui.Hwnd) {
                gui.Destroy()
            }
        } catch {
            ; Silently ignore errors
        }
    }
}

; Function to cleanup square selector system
CleanupSquareSelector() {
    global g_SquareSelectorActive, g_SquareSelectorGuis, g_SquareSelectorTimer
    global g_SquareSelectorLetters, g_SquareSelectorHotkeyHandlers
    global g_SquareSelectorLock, g_ActiveDirection, g_SquareSelectorLoopMode

    ; Disable active flag immediately
    g_SquareSelectorActive := false

    ; Disable all letter hotkeys immediately using stored handlers
    for letter in g_SquareSelectorLetters {
        try {
            Hotkey(letter, "Off")
            Hotkey(StrLower(letter), "Off")
        } catch {
            ; Silently ignore errors if hotkey doesn't exist
        }
    }

    ; ALWAYS disable loop mode hotkeys (including mouse button hotkeys) to prevent blocking clicks
    ; This ensures mouse clicks work even if hotkeys were enabled through a race condition
    try {
        DisableLoopModeHotkeys()
    } catch {
        ; Silently ignore if there's an error
    }
    g_SquareSelectorLoopMode := false

    ; Disable CTRL hotkey (click mode toggle)
    try {
        Hotkey("Ctrl", "Off")
    } catch {
        ; Silently ignore if hotkey doesn't exist
    }

    ; Disable direction switch hotkeys
    DisableDirectionSwitchHotkeys()

    ; Disable Escape hotkey
    try {
        Hotkey("Escape", "Off")
    } catch {
        ; Silently ignore if hotkey doesn't exist
    }

    ; Reset click mode flag
    global g_SquareSelectorClickMode
    g_SquareSelectorClickMode := false

    ; Clear hotkey handlers array
    g_SquareSelectorHotkeyHandlers := []

    ; Destroy all square GUIs
    DestroyGuiArray(g_SquareSelectorGuis)
    g_SquareSelectorPositions := []

    ; Clean up direction indicator squares
    CleanupDirectionIndicators()

    ; Cancel timer if active
    if (g_SquareSelectorTimer) {
        SetTimer(g_SquareSelectorTimer, 0)
        g_SquareSelectorTimer := false
    }

    ; Cancel backup timer if active
    global g_SquareSelectorBackupTimer
    if (g_SquareSelectorBackupTimer) {
        SetTimer(g_SquareSelectorBackupTimer, 0)
        g_SquareSelectorBackupTimer := false
    }

    ; Cancel old squares cleanup timer if active
    global g_OldSquaresCleanupTimer
    if (g_OldSquaresCleanupTimer) {
        SetTimer(g_OldSquaresCleanupTimer, 0)
        g_OldSquaresCleanupTimer := false
    }

    ; Clear start time
    global g_SquareSelectorStartTime
    g_SquareSelectorStartTime := 0

    ; Release lock and clear active direction to prevent bugs
    ; This ensures the hotkeys can be used again after cleanup
    g_SquareSelectorLock := false
    g_ActiveDirection := ""
}

; Function to show 15 squares with letters in a line in the chosen direction
ShowSquareSelector(direction) {
    global g_SquareSelectorActive, g_SquareSelectorGuis, g_SquareSelectorPositions
    global g_SquareSelectorLetters, g_SquareSelectorLock

    ; Don't clear arrays immediately - preserve old squares
    ; We'll clean them up after showing new ones if needed
    oldGuis := g_SquareSelectorGuis.Clone()
    oldPositions := g_SquareSelectorPositions.Clone()

    ; Clear arrays for new squares
    g_SquareSelectorGuis := []
    g_SquareSelectorPositions := []

    ; Don't call CleanupSquareSelector here - it destroys squares
    ; Instead, just disable hotkeys temporarily
    DisableLetterHotkeys()
    try {
        Hotkey("Ctrl", "Off")
    } catch {
        ; Ignore
    }
    DisableDirectionSwitchHotkeys()
    try {
        Hotkey("Escape", "Off")
    } catch {
        ; Ignore
    }

    ; Clean up old squares after a brief delay (allows new squares to appear first)
    ; Cancel any existing old squares cleanup timer first
    global g_OldSquaresCleanupTimer
    if (g_OldSquaresCleanupTimer) {
        SetTimer(g_OldSquaresCleanupTimer, 0)
        g_OldSquaresCleanupTimer := false
    }

    if (oldGuis.Length > 0) {
        g_OldSquaresCleanupTimer := () => CleanupOldSquareGuis(oldGuis)
        SetTimer(g_OldSquaresCleanupTimer, -50)
    }

    Sleep 10

    ; Get current mouse position
    pos := GetMousePos()
    startX := pos.x
    startY := pos.y

    ; Configuration
    squareSize := 24  ; Reduced for more precision
    spacing := 20  ; Reduced for more precision
    numSquares := 38  ; Updated to match total characters in g_SquareSelectorLetters

    ; Normalize direction
    directionLower := StrLower(direction)

    ; STEP 1: Calculate all center positions first
    ; First square (1) starts AFTER mouse position, not centered on it
    ; Initial offset: half square size (12px) + spacing (20px) = 32px from mouse position
    ; This ensures the first square's left edge starts after the mouse cursor
    initialOffset := (squareSize / 2.0) + spacing  ; 12 + 20 = 32 pixels

    calculatedPositions := []
    if (directionLower = "right" || directionLower = "left") {
        ; Horizontal line
        directionMultiplier := directionLower = "right" ? 1 : -1
        loop numSquares {
            i := A_Index
            ; Calculate offset for square i
            ; First square (i=1): initialOffset (32px) - starts after mouse
            ; Subsequent squares: initialOffset + (i-1) * (squareSize + spacing)
            ; For i=1: 32px, for i=2: 32 + 44 = 76px, for i=3: 32 + 88 = 120px, etc.
            offset := (initialOffset + (i - 1) * (squareSize + spacing)) * directionMultiplier
            squareCenterX := Round(startX + offset)
            squareCenterY := startY
            calculatedPositions.Push({ x: squareCenterX, y: squareCenterY })
        }
    } else {
        ; Vertical line (up or down)
        directionMultiplier := directionLower = "down" ? 1 : -1
        loop numSquares {
            i := A_Index
            ; Same calculation for vertical: first square starts after mouse
            offset := (initialOffset + (i - 1) * (squareSize + spacing)) * directionMultiplier
            squareCenterX := startX
            squareCenterY := Round(startY + offset)
            calculatedPositions.Push({ x: squareCenterX, y: squareCenterY })
        }
    }

    ; STEP 2: Create all GUIs at once (don't show yet)
    guiArray := []
    loop numSquares {
        i := A_Index
        pos := calculatedPositions[i]

        ; Create square GUI with letter
        squareGui := Gui("+AlwaysOnTop -Caption +ToolWindow")
        ; Color depends on click mode: blue if click mode active, red otherwise
        global g_SquareSelectorClickMode
        squareGui.BackColor := g_SquareSelectorClickMode ? "0000FF" : "FF0000"  ; Blue or Red
        squareGui.SetFont("s8 Bold cFFFFFF", "Segoe UI")  ; White text, bold, smaller for precision

        ; Set GUI margins to 0 to eliminate any padding that could affect centering
        squareGui.MarginX := 0
        squareGui.MarginY := 0

        ; Create text control that perfectly centers the letter
        ; Center = 0x1 (SS_CENTER) for horizontal centering
        ; 0x200 = SS_CENTERIMAGE for vertical centering
        ; 0x201 combines both (SS_CENTER | SS_CENTERIMAGE) for perfect centering
        ; Text control fills entire square (40x40) to ensure proper centering
        letterText := squareGui.AddText("w" . squareSize . " h" . squareSize . " Center 0x201",
            g_SquareSelectorLetters[i])

        ; Calculate top-left position for this square
        guiX := Round(pos.x - squareSize / 2.0)
        guiY := Round(pos.y - squareSize / 2.0)

        ; Store GUI and position info (not shown yet)
        guiArray.Push({ gui: squareGui, x: guiX, y: guiY, calculatedCenter: pos })
    }

    ; STEP 3: Prepare all GUIs (position while hidden for instant showing)
    for guiInfo in guiArray {
        ; Position while hidden (no rendering delay)
        guiInfo.gui.Show("x" . guiInfo.x . " y" . guiInfo.y . " w" . squareSize . " h" . squareSize . " NA Hide")
        ; Set 80% opacity (204 = 80% opacity, 255 = fully opaque, 0 = fully transparent)
        WinSetTransparent(204, guiInfo.gui)
    }

    ; STEP 4: Show all GUIs simultaneously (batch show for instant appearance)
    ; Use Show() instead of SetWindowPos to ensure windows actually appear
    ; Show all windows using Show() - this is more reliable than SetWindowPos
    for guiInfo in guiArray {
        try {
            ; Show window using Show() - ensure it actually appears
            guiInfo.gui.Show("NA")  ; Show without activating
        } catch {
            ; If Show() fails, try using the position again
            try {
                guiInfo.gui.Show("x" . guiInfo.x . " y" . guiInfo.y . " w" . squareSize . " h" . squareSize . " NA")
            }
        }
    }

    ; STEP 5: Brief delay to ensure all GUIs are fully rendered
    Sleep 20  ; Increased delay to ensure windows are fully rendered before querying positions

    ; STEP 6: Query actual GUI positions and store actual centers for mouse jump
    ; Query actual window positions using GetWindowRect to get exact centers
    ; This accounts for any window borders, padding, or DPI adjustments
    for i, guiInfo in guiArray {
        squareGuiObj := guiInfo.gui  ; Use different variable name to avoid conflict
        g_SquareSelectorGuis.Push(squareGuiObj)

        ; Query actual window rectangle using GetWindowRect
        ; This gives us the actual physical pixel coordinates after DPI adjustments
        rect := Buffer(16, 0)  ; RECT structure: left, top, right, bottom (4 ints)
        if (DllCall("GetWindowRect", "ptr", squareGuiObj.Hwnd, "ptr", rect)) {
            ; Extract rectangle coordinates (physical pixels with DPI awareness)
            winLeft := NumGet(rect, 0, "int")
            winTop := NumGet(rect, 4, "int")
            winRight := NumGet(rect, 8, "int")
            winBottom := NumGet(rect, 12, "int")

            ; Calculate actual center from window rectangle
            actualCenterX := winLeft + (winRight - winLeft) / 2
            actualCenterY := winTop + (winBottom - winTop) / 2

            ; Store actual center position (rounded to nearest pixel)
            g_SquareSelectorPositions.Push({ x: Round(actualCenterX), y: Round(actualCenterY) })
        } else {
            ; Fallback to calculated position if GetWindowRect fails
            g_SquareSelectorPositions.Push({ x: guiInfo.calculatedCenter.x, y: guiInfo.calculatedCenter.y })
        }
    }

    ; Activate letter selection mode
    g_SquareSelectorActive := true
    g_SquareSelectorClickMode := false  ; Reset click mode when showing new squares
    SetupLetterKeyListener()

    ; Enable CTRL hotkey to toggle click mode
    Hotkey("Ctrl", (*) => HandleCtrlToggle(), "On")

    ; Enable arrow keys for immediate direction switching
    EnableDirectionSwitchHotkeys()

    ; Enable Escape key to cancel squares (works in initial mode)
    Hotkey("Escape", (*) => CancelSquareSelector(), "On")

    ; Record start time for guaranteed cleanup
    global g_SquareSelectorStartTime
    g_SquareSelectorStartTime := A_TickCount

    ; Set timer to cleanup after 7 seconds if nothing is pressed
    ; Create cleanup function bound to this session ID (prevents old timers from cleaning up new squares)
    currentSessionID := g_SquareSelectorSessionID
    g_SquareSelectorTimer := CreateTimerHandler(currentSessionID)
    SetTimer(g_SquareSelectorTimer, -7000)  ; 7 second timeout

    ; Set up backup cleanup timer that checks every 2 seconds (guaranteed cleanup after 7 seconds)
    global g_SquareSelectorBackupTimer
    if (g_SquareSelectorBackupTimer) {
        SetTimer(g_SquareSelectorBackupTimer, 0)  ; Cancel old backup timer
    }
    g_SquareSelectorBackupTimer := () => BackupCleanupTimer()
    SetTimer(g_SquareSelectorBackupTimer, 2000)  ; Check every 2 seconds

    ; Lock will be released when timer fires, user selects a letter (enters loop mode), or presses Escape
}

; Function to show 4 direction indicator squares around the mouse pointer
ShowDirectionIndicators() {
    global g_DirectionIndicatorGuis

    ; Clean up any existing direction indicators
    CleanupDirectionIndicators()

    ; Get current mouse position
    pos := GetMousePos()
    mouseX := pos.x
    mouseY := pos.y

    ; Configuration
    squareSize := 24  ; Reduced for more precision
    offset := 35  ; Reduced for more precision

    ; Arrow symbols for each direction
    arrowUp := "↑"
    arrowRight := "→"
    arrowDown := "↓"
    arrowLeft := "←"
    arrows := [arrowUp, arrowRight, arrowDown, arrowLeft]

    ; Positions relative to mouse: Up, Right, Down, Left
    positions := []
    positions.Push({ x: mouseX, y: mouseY - offset })           ; Up
    positions.Push({ x: mouseX + offset, y: mouseY })           ; Right
    positions.Push({ x: mouseX, y: mouseY + offset })           ; Down
    positions.Push({ x: mouseX - offset, y: mouseY })            ; Left

    ; Create all 4 indicator squares
    guiArray := []
    for i, arrow in arrows {
        pos := positions[i]

        ; Create square GUI with arrow
        indicatorGui := Gui("+AlwaysOnTop -Caption +ToolWindow")
        indicatorGui.BackColor := "FF0000"  ; Red
        indicatorGui.SetFont("s10 Bold cFFFFFF", "Segoe UI")  ; White text, bold, smaller for precision

        ; Set GUI margins to 0
        indicatorGui.MarginX := 0
        indicatorGui.MarginY := 0

        ; Create text control that perfectly centers the arrow
        arrowText := indicatorGui.AddText("w" . squareSize . " h" . squareSize . " Center 0x201", arrow)

        ; Calculate top-left position for this square
        guiX := Round(pos.x - squareSize / 2.0)
        guiY := Round(pos.y - squareSize / 2.0)

        ; Store GUI and position info
        guiArray.Push({ gui: indicatorGui, x: guiX, y: guiY })
    }

    ; Position all GUIs while hidden, then show simultaneously
    for guiInfo in guiArray {
        guiInfo.gui.Show("x" . guiInfo.x . " y" . guiInfo.y . " w" . squareSize . " h" . squareSize . " NA Hide")
        ; Set less opaque (same as letter squares)
        WinSetTransparent(80, guiInfo.gui)
    }

    ; Show all GUIs simultaneously
    for guiInfo in guiArray {
        try {
            guiInfo.gui.Show("NA")
        } catch {
            try {
                guiInfo.gui.Show("x" . guiInfo.x . " y" . guiInfo.y . " w" . squareSize . " h" . squareSize . " NA")
            }
        }
        g_DirectionIndicatorGuis.Push(guiInfo.gui)
    }
}

; Helper function to destroy GUI objects in an array (reusable)
DestroyGuiArray(guis) {
    if (!guis || guis.Length = 0) {
        return
    }
    for gui in guis {
        try {
            if (IsObject(gui) && gui.Hwnd) {
                gui.Destroy()
            }
        } catch {
            ; Silently ignore errors
        }
    }
    guis.Length := 0  ; Clear array efficiently
}

; Helper function to cleanup direction indicator squares
CleanupDirectionIndicators() {
    global g_DirectionIndicatorGuis
    DestroyGuiArray(g_DirectionIndicatorGuis)
}

; Factory function to create a handler that properly captures the index
; This ensures each handler gets its own copy of the index value
CreateSquareSelectorHandler(index) {
    ; Return a function that captures the index value at creation time
    return (*) => SelectSquareByIndex(index)
}

; Function to setup hotkey listeners for letter keys
; Uses individual hotkeys that are only active when square selector is shown
SetupLetterKeyListener() {
    global g_SquareSelectorLetters, g_SquareSelectorHotkeyHandlers

    ; Clear any existing handlers
    g_SquareSelectorHotkeyHandlers := []

    ; Create a handler for each letter using factory function
    ; This ensures proper closure capture - each handler gets its own index value
    for i, letter in g_SquareSelectorLetters {
        ; Use factory function to create handler with properly captured index
        handler := CreateSquareSelectorHandler(i)

        ; Store handler reference for cleanup (optional, but good practice)
        g_SquareSelectorHotkeyHandlers.Push({ letter: letter, handler: handler })

        ; Enable both uppercase and lowercase versions
        Hotkey(letter, handler, "On")
        Hotkey(StrLower(letter), handler, "On")
    }
}

; Helper function to disable letter hotkeys (used when entering loop mode)
DisableLetterHotkeys() {
    global g_SquareSelectorLetters

    ; Disable all letter hotkeys
    for letter in g_SquareSelectorLetters {
        try {
            Hotkey(letter, "Off")
            Hotkey(StrLower(letter), "Off")
        } catch {
            ; Silently ignore errors if hotkey doesn't exist
        }
    }
}

; Function to toggle click mode and update square colors
ToggleClickMode() {
    global g_SquareSelectorClickMode, g_SquareSelectorActive, g_SquareSelectorGuis

    ; Only toggle if squares are visible
    if (!g_SquareSelectorActive) {
        return
    }

    ; Toggle click mode flag
    g_SquareSelectorClickMode := !g_SquareSelectorClickMode

    ; Update all square colors based on click mode
    newColor := g_SquareSelectorClickMode ? "0000FF" : "FF0000"  ; Blue or Red
    for gui in g_SquareSelectorGuis {
        try {
            if (IsObject(gui) && gui.Hwnd) {
                gui.BackColor := newColor
                ; Force redraw by hiding and showing
                gui.Show("Hide")
                gui.Show("NA")
            }
        } catch {
            ; Silently ignore errors
        }
    }
}

; Handler for CTRL key to toggle click mode
HandleCtrlToggle() {
    global g_SquareSelectorActive
    ; Only toggle if squares are active
    if (g_SquareSelectorActive) {
        ToggleClickMode()
    }
}

; Handler for letter key press - uses index directly to avoid matching issues
SelectSquareByIndex(index) {
    global g_SquareSelectorActive, g_SquareSelectorPositions, g_ActiveDirection
    global g_SquareSelectorLoopMode, g_SquareSelectorLock

    ; Double-check that selector is active (safety check)
    if (!g_SquareSelectorActive) {
        return
    }

    ; Verify positions array is valid
    if (!g_SquareSelectorPositions || g_SquareSelectorPositions.Length = 0) {
        ; Positions array is empty, cleanup and abort
        CleanupSquareSelector()
        return
    }

    ; Validate index
    if (index < 1 || index > g_SquareSelectorPositions.Length) {
        CleanupSquareSelector()
        return
    }

    ; Get the position for this square (index is 1-based)
    targetPos := g_SquareSelectorPositions[index]

    ; Move mouse to the center of the selected square
    DllCall("SetCursorPos", "int", targetPos.x, "int", targetPos.y)

    ; Check if click mode is active
    global g_SquareSelectorClickMode
    if (g_SquareSelectorClickMode) {
        ; Click mode: perform a click and exit completely
        ; STEP 1: Store target position before cleanup (targetPos is already stored)

        ; STEP 2: Immediately disable all hotkeys and cancel ALL timers
        DisableLetterHotkeys()
        global g_SquareSelectorTimer, g_SquareSelectorBackupTimer, g_OldSquaresCleanupTimer
        if (g_SquareSelectorTimer) {
            SetTimer(g_SquareSelectorTimer, 0)
            g_SquareSelectorTimer := false
        }
        if (g_SquareSelectorBackupTimer) {
            SetTimer(g_SquareSelectorBackupTimer, 0)
            g_SquareSelectorBackupTimer := false
        }
        ; Cancel the old squares cleanup timer from ShowSquareSelector
        if (g_OldSquaresCleanupTimer) {
            SetTimer(g_OldSquaresCleanupTimer, 0)
            g_OldSquaresCleanupTimer := false
        }
        ; Clear start time
        global g_SquareSelectorStartTime
        g_SquareSelectorStartTime := 0

        ; Disable other hotkeys immediately
        try {
            Hotkey("Ctrl", "Off")
        } catch {
            ; Ignore
        }
        DisableDirectionSwitchHotkeys()
        try {
            Hotkey("Escape", "Off")
        } catch {
            ; Ignore
        }

        ; STEP 3: Destroy all square GUIs immediately and aggressively
        ; This must happen BEFORE the click so squares don't block it
        global g_SquareSelectorGuis
        ; Destroy all squares in the array
        for gui in g_SquareSelectorGuis {
            try {
                if (IsObject(gui) && gui.Hwnd) {
                    ; Force immediate destruction - no hiding, just destroy
                    gui.Destroy()
                }
            } catch {
                ; Silently ignore errors
            }
        }
        ; Clear arrays immediately
        g_SquareSelectorGuis := []
        g_SquareSelectorPositions := []

        ; Also destroy direction indicators immediately
        CleanupDirectionIndicators()

        ; Brief delay to ensure GUI destruction is complete
        Sleep 15

        ; STEP 4: Wait briefly for GUI cleanup to complete
        Sleep 25

        ; STEP 5: Find window at target position (now that squares are gone)
        targetHwnd := DllCall("WindowFromPoint", "Int64", (targetPos.y << 32) | (targetPos.x & 0xFFFFFFFF), "Ptr")
        if (targetHwnd) {
            ; Get the root window (in case we got a child window)
            rootHwnd := DllCall("GetAncestor", "Ptr", targetHwnd, "UInt", 2, "Ptr")  ; GA_ROOT = 2
            if (rootHwnd) {
                targetHwnd := rootHwnd
            }
            ; Activate the window
            try {
                WinActivate("ahk_id " . targetHwnd)
                WinWaitActive("ahk_id " . targetHwnd, , 0.35)
            } catch {
                ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
            }
        }

        ; STEP 6: Move mouse and click
        DllCall("SetCursorPos", "int", targetPos.x, "int", targetPos.y)
        Sleep 40

        ; Update last mouse click tick to prevent MonitorActiveWindow interference
        try {
            g_LastMouseClickTick := A_TickCount
        } catch {
            ; Ignore if variable doesn't exist
        }

        ; Perform the click
        Click

        ; STEP 8: Final cleanup - ensure everything is reset
        ; Double-check that all squares are destroyed (defensive cleanup)
        global g_SquareSelectorGuis
        if (g_SquareSelectorGuis.Length > 0) {
            for gui in g_SquareSelectorGuis {
                try {
                    if (IsObject(gui) && gui.Hwnd) {
                        gui.Destroy()
                    }
                } catch {
                    ; Ignore
                }
            }
            g_SquareSelectorGuis := []
        }

        ; Reset all state flags
        g_SquareSelectorLock := false
        g_ActiveDirection := ""
        g_SquareSelectorClickMode := false
        g_SquareSelectorActive := false
        g_SquareSelectorLoopMode := false

        ; Final cleanup of direction indicators (defensive)
        CleanupDirectionIndicators()

        return
    }

    ; Normal mode: Keep letter/number squares visible - don't destroy them
    ; Store the current direction before cleanup (for predicting next direction)
    currentDirection := g_ActiveDirection

    ; Cancel timeout timer since we're entering loop mode
    global g_SquareSelectorTimer, g_SquareSelectorBackupTimer
    if (g_SquareSelectorTimer) {
        SetTimer(g_SquareSelectorTimer, 0)
        g_SquareSelectorTimer := false
    }
    if (g_SquareSelectorBackupTimer) {
        SetTimer(g_SquareSelectorBackupTimer, 0)
        g_SquareSelectorBackupTimer := false
    }
    ; Clear start time since we're entering loop mode
    global g_SquareSelectorStartTime
    g_SquareSelectorStartTime := 0

    ; Predict user wants to continue in same direction - show new squares immediately
    ; This speeds up the workflow (user doesn't need to press arrow key)
    ; Keep old squares visible - don't destroy them, just show new ones
    if (currentDirection) {
        ; Small delay to ensure mouse position is stable
        Sleep 50

        ; Store old squares temporarily so we can clean them up after showing new ones
        global g_SquareSelectorGuis
        oldSquares := g_SquareSelectorGuis.Clone()

        ; Automatically show new squares in the same direction
        ; The mouse is now at the selected square position, so new squares will continue from there
        ; ShowSquareSelector will try to clean up, but we'll preserve old squares
        ShowSquareSelector(currentDirection)

        ; Clean up old squares after a brief delay to allow new squares to appear
        SetTimer(() => CleanupOldSquareGuis(oldSquares), -100)

        ; Cancel the timeout timer that ShowSquareSelector set up - we're in loop mode, no timeout
        global g_SquareSelectorTimer, g_SquareSelectorBackupTimer
        if (g_SquareSelectorTimer) {
            SetTimer(g_SquareSelectorTimer, 0)
            g_SquareSelectorTimer := false
        }
        if (g_SquareSelectorBackupTimer) {
            SetTimer(g_SquareSelectorBackupTimer, 0)
            g_SquareSelectorBackupTimer := false
        }
        ; Clear start time since we're in loop mode
        global g_SquareSelectorStartTime
        g_SquareSelectorStartTime := 0
    } else {
        ; No direction - keep squares visible (don't destroy them)
        ; The algorithm finishes but letters remain displayed
        ; Just disable hotkeys but keep squares visible
        DisableLetterHotkeys()
        try {
            Hotkey("Ctrl", "Off")
        } catch {
            ; Ignore
        }
        DisableDirectionSwitchHotkeys()
        try {
            Hotkey("Escape", "Off")
        } catch {
            ; Ignore
        }
        ; Don't destroy squares - keep them visible
        g_SquareSelectorActive := false
        g_SquareSelectorLock := false
    }

    ; Show direction indicator squares AFTER new squares are shown
    ; (ShowSquareSelector calls CleanupSquareSelector which removes direction indicators,
    ;  so we need to show them after to prevent blinking/vanishing)
    ShowDirectionIndicators()

    ; Enter loop mode
    ; Letter hotkeys are now re-enabled by ShowSquareSelector
    ; Disable letter/number hotkeys will be handled by loop mode handlers

    ; Disable direction switch hotkeys before enabling loop mode hotkeys
    DisableDirectionSwitchHotkeys()

    ; Set loop mode flag
    g_SquareSelectorLoopMode := true

    ; Enable loop mode hotkeys (Escape and arrow keys)
    EnableLoopModeHotkeys()

    ; DO NOT clear g_ActiveDirection - needed for context
    ; DO NOT release lock - maintained during loop mode
}

; Simplified helper function to handle direction hotkey
HandleDirectionHotkey(direction) {
    ; TEST: Uncomment next line to verify hotkey is firing
    ; MsgBox "Hotkey triggered: " . direction, "Debug"

    global g_SquareSelectorActive, g_ActiveDirection, g_SquareSelectorTimer
    global g_SquareSelectorLock, g_SquareSelectorClickMode

    ; STEP 0: Preserve click mode state BEFORE cleanup (so blue squares stay blue when changing direction)
    preservedClickMode := g_SquareSelectorClickMode

    ; STEP 1: IMMEDIATELY disable active flag and clear positions
    ; This prevents letter hotkeys from using old positions
    g_SquareSelectorActive := false
    global g_SquareSelectorPositions
    g_SquareSelectorPositions := []

    ; STEP 2: Increment session ID to invalidate any old timers
    global g_SquareSelectorSessionID
    g_SquareSelectorSessionID++

    ; STEP 3: Cancel any existing timers FIRST (prevents old timers from cleaning up new squares)
    if (g_SquareSelectorTimer) {
        SetTimer(g_SquareSelectorTimer, 0)
        g_SquareSelectorTimer := false
    }
    global g_SquareSelectorBackupTimer
    if (g_SquareSelectorBackupTimer) {
        SetTimer(g_SquareSelectorBackupTimer, 0)
        g_SquareSelectorBackupTimer := false
    }

    ; STEP 4: Clean up old selector completely (GUIs, hotkeys, etc.)
    CleanupSquareSelector()

    ; STEP 5: Reset lock to ensure clean state
    g_SquareSelectorLock := false

    ; STEP 6: Wait a bit for cleanup to complete and brief delay before showing new squares
    Sleep 80

    ; STEP 7: Set new active direction
    g_ActiveDirection := StrLower(direction)

    ; STEP 8: Show new squares (delay already included above)

    ; STEP 9: Disable loop mode if it was active (transitioning from loop mode)
    global g_SquareSelectorLoopMode
    if (g_SquareSelectorLoopMode) {
        DisableLoopModeHotkeys()
        g_SquareSelectorLoopMode := false
    }

    ; STEP 10: Show the new squares (with session ID)
    ShowSquareSelector(g_ActiveDirection)

    ; STEP 11: Restore click mode state if it was active (so blue squares remain blue)
    if (preservedClickMode) {
        g_SquareSelectorClickMode := true
        ; Update all square colors to blue to reflect click mode
        global g_SquareSelectorGuis
        for gui in g_SquareSelectorGuis {
            try {
                if (IsObject(gui) && gui.Hwnd) {
                    gui.BackColor := "0000FF"  ; Blue
                    ; Force redraw by hiding and showing
                    gui.Show("Hide")
                    gui.Show("NA")
                }
            } catch {
                ; Silently ignore errors
            }
        }
    }
}

; Helper function to cancel squares (works in both initial mode and loop mode)
CancelSquareSelector() {
    global g_SquareSelectorLoopMode, g_SquareSelectorLock, g_ActiveDirection
    global g_SquareSelectorActive

    ; Only handle if squares are active
    if (!g_SquareSelectorActive && !g_SquareSelectorLoopMode) {
        return
    }

    ; ALWAYS disable loop mode hotkeys (including mouse button hotkeys) to prevent blocking clicks
    try {
        DisableLoopModeHotkeys()
    } catch {
        ; Silently ignore if there's an error
    }

    ; Cleanup completely
    CleanupSquareSelector()

    ; Reset all state
    g_SquareSelectorLoopMode := false
    g_SquareSelectorLock := false
    g_ActiveDirection := ""
}

; Helper function to exit loop mode (shared by Escape and mouse handlers)
ExitLoopMode() {
    global g_SquareSelectorLoopMode, g_SquareSelectorLock, g_ActiveDirection

    ; Only handle if in loop mode
    if (!g_SquareSelectorLoopMode) {
        return
    }

    ; Disable loop mode hotkeys
    DisableLoopModeHotkeys()

    ; Cleanup completely
    CleanupSquareSelector()

    ; Reset all state
    g_SquareSelectorLoopMode := false
    g_SquareSelectorLock := false
    g_ActiveDirection := ""
}

; Mouse click handlers for loop mode (exit and forward the click)
HandleLoopModeLButton() {
    global g_SquareSelectorLoopMode
    if (g_SquareSelectorLoopMode) {
        ExitLoopMode()
        ; Send the click after exiting loop mode
        Click("Left")
    }
}

HandleLoopModeRButton() {
    global g_SquareSelectorLoopMode
    if (g_SquareSelectorLoopMode) {
        ExitLoopMode()
        ; Send the click after exiting loop mode
        Click("Right")
    }
}

HandleLoopModeMButton() {
    global g_SquareSelectorLoopMode
    if (g_SquareSelectorLoopMode) {
        ExitLoopMode()
        ; Send the click after exiting loop mode
        Click("Middle")
    }
}

; Helper function to enable direction switch hotkeys (arrow keys for switching directions immediately)
EnableDirectionSwitchHotkeys() {
    ; Enable arrow key hotkeys for immediate direction switching (without modifiers)
    Hotkey("Right", (*) => HandleDirectionHotkey("Right"), "On")
    Hotkey("Left", (*) => HandleDirectionHotkey("Left"), "On")
    Hotkey("Down", (*) => HandleDirectionHotkey("Down"), "On")
    Hotkey("Up", (*) => HandleDirectionHotkey("Up"), "On")
}

; Helper function to disable direction switch hotkeys
DisableDirectionSwitchHotkeys() {
    ; Disable arrow key hotkeys
    try {
        Hotkey("Right", "Off")
        Hotkey("Left", "Off")
        Hotkey("Down", "Off")
        Hotkey("Up", "Off")
    } catch {
        ; Silently ignore if hotkeys don't exist
    }
}

; Helper function to enable loop mode hotkeys (Escape, arrow keys, and mouse clicks)
EnableLoopModeHotkeys() {
    ; Enable Escape hotkey for loop mode (uses CancelSquareSelector which works for both modes)
    Hotkey("Escape", (*) => CancelSquareSelector(), "On")

    ; Enable arrow key hotkeys for loop mode (without modifiers)
    Hotkey("Right", (*) => HandleLoopModeRight(), "On")
    Hotkey("Left", (*) => HandleLoopModeLeft(), "On")
    Hotkey("Down", (*) => HandleLoopModeDown(), "On")
    Hotkey("Up", (*) => HandleLoopModeUp(), "On")

    ; Enable mouse click hotkeys to exit loop mode (forward click after exit)
    Hotkey("LButton", (*) => HandleLoopModeLButton(), "On")
    Hotkey("RButton", (*) => HandleLoopModeRButton(), "On")
    Hotkey("MButton", (*) => HandleLoopModeMButton(), "On")
}

; Helper function to disable loop mode hotkeys
DisableLoopModeHotkeys() {
    ; Disable Escape hotkey
    try {
        Hotkey("Escape", "Off")
    } catch {
        ; Silently ignore if hotkey doesn't exist
    }

    ; Disable arrow key hotkeys
    try {
        Hotkey("Right", "Off")
        Hotkey("Left", "Off")
        Hotkey("Down", "Off")
        Hotkey("Up", "Off")
    } catch {
        ; Silently ignore if hotkeys don't exist
    }

    ; Disable mouse click hotkeys
    try {
        Hotkey("LButton", "Off")
        Hotkey("RButton", "Off")
        Hotkey("MButton", "Off")
    } catch {
        ; Silently ignore if hotkeys don't exist
    }
}

; Escape key handler for loop mode
HandleEscapeKey() {
    ExitLoopMode()
}

; Loop mode arrow key handlers (only active when in loop mode)
HandleLoopModeRight() {
    global g_SquareSelectorLoopMode
    if (g_SquareSelectorLoopMode) {
        HandleDirectionHotkey("Right")
    }
}

HandleLoopModeLeft() {
    global g_SquareSelectorLoopMode
    if (g_SquareSelectorLoopMode) {
        HandleDirectionHotkey("Left")
    }
}

HandleLoopModeDown() {
    global g_SquareSelectorLoopMode
    if (g_SquareSelectorLoopMode) {
        HandleDirectionHotkey("Down")
    }
}

HandleLoopModeUp() {
    global g_SquareSelectorLoopMode
    if (g_SquareSelectorLoopMode) {
        HandleDirectionHotkey("Up")
    }
}

; Win+Alt+Shift+Arrow: send that arrow key five times
#!+Right::
{
    Send("{Right 5}")
    return
}

#!+Left::
{
    Send("{Left 5}")
    return
}

#!+Down::
{
    Send("{Down 5}")
    return
}

#!+Up::
{
    Send("{Up 5}")
    return
}

; =============================================================================
; Peek PDF – Win+Alt+Shift+X
; If Peek is open: activate it. Otherwise: show study-topic selector (same aesthetic as Win+Alt+Shift+C).
; =============================================================================

; Study topics for Win+Alt+Shift+X selector. Paths are relative to notes repo (GetNotesRepoPath()).
global g_StudyTopics := Map(
    0, { name: "Technique (how to create studies)", path: "\studies\technique\README.pdf", skipLastPage: true },
    1, { name: "English", path: "\studies\english\lists\1\1.pdf" },
    2, { name: "Piano", path: "\studies\piano\lists\1\1.pdf" },
    3, { name: "Communication", path: "\studies\communication\lists\1\1.pdf" },
    4, { name: "Statistics", path: "\studies\statistics\lists\1\1.pdf" },
    5, { name: "Mnemonics README", path: "\studies\technique\README.pdf" }
)
global g_StudyTopicSelectorGui := false
global g_StudyTopicSelectorActive := false

; PDF focus monitoring for automatic blackout cancellation (Win+Alt+Shift+X)
global g_PdfFocusMonitorTimer := false
global g_PdfFocusTrackedHwnd := 0
global g_PdfFocusLossMode := "Immediate"      ; "Immediate" or "Debounced"
global g_PdfFocusDebounceMs := 1200            ; Allow transient focus loss without un-blackouting
global g_PdfFocusLostSinceTick := 0

; Monitor PDF (Peek) window focus and automatically disable focus mode when it loses focus
MonitorPdfFocus() {
    global g_PdfFocusTrackedHwnd, g_PdfFocusLossMode, g_PdfFocusDebounceMs, g_PdfFocusLostSinceTick

    ; Check if tracked window still exists
    if (g_PdfFocusTrackedHwnd && !WinExist("ahk_id " . g_PdfFocusTrackedHwnd)) {
        DisableFocusMode()
        StopPdfFocusMonitor()
        return
    }

    ; Check if PDF/QuickLook window is still the active window
    if (!WinActive("ahk_id " . g_PdfFocusTrackedHwnd)) {
        if (g_PdfFocusLossMode = "Debounced") {
            ; Wait for focus loss to persist before canceling blackout.
            if (!g_PdfFocusLostSinceTick)
                g_PdfFocusLostSinceTick := A_TickCount

            if ((A_TickCount - g_PdfFocusLostSinceTick) >= g_PdfFocusDebounceMs) {
                DisableFocusMode()
                StopPdfFocusMonitor()
            }
            return
        }

        ; Default: immediate cancellation (keeps behavior for Peek/PDF workflows)
        DisableFocusMode()
        StopPdfFocusMonitor()
    } else {
        ; Focus is stable again; reset debounce timer.
        g_PdfFocusLostSinceTick := 0
    }
}

; Start monitoring PDF window focus
StartPdfFocusMonitor(hwnd := 0, focusLossMode := "Immediate") {
    global g_PdfFocusMonitorTimer, g_PdfFocusTrackedHwnd, g_PdfFocusLossMode, g_PdfFocusLostSinceTick

    StopPdfFocusMonitor()

    g_PdfFocusTrackedHwnd := hwnd ? hwnd : WinExist("A")
    if (!g_PdfFocusTrackedHwnd)
        return

    g_PdfFocusLossMode := focusLossMode
    g_PdfFocusLostSinceTick := 0

    g_PdfFocusMonitorTimer := MonitorPdfFocus
    SetTimer(g_PdfFocusMonitorTimer, 200)
}

; Stop monitoring PDF window focus
StopPdfFocusMonitor() {
    global g_PdfFocusMonitorTimer, g_PdfFocusTrackedHwnd

    if (g_PdfFocusMonitorTimer) {
        SetTimer(g_PdfFocusMonitorTimer, 0)
        g_PdfFocusMonitorTimer := false
    }
    g_PdfFocusTrackedHwnd := 0
}

; YouTube focus session (Win+Alt+Shift+H): toggle on/off; SMTC for Spotify play/pause (not toggle).
global g_YoutubeFocusMonitorTimer := false
global g_YoutubeFocusTrackedHwnd := 0
global g_YoutubeSpotifyPausePending := false
global g_YoutubeFocusSessionActive := false

; Find Spotify's Windows.Media.Control session (SourceAppUserModelId contains "Spotify").
YouTube_FindSpotifyMediaSession() {
    try {
        for session in Media.GetSessions() {
            try {
                id := session.SourceAppUserModelId
                if InStr(id, "Spotify")
                    return session
            } catch {
                continue
            }
        }
    } catch {
        return 0
    }
    return 0
}

; Pause Spotify via SMTC only when status is Playing (avoids starting playback when already paused).
YouTube_PauseSpotifyBeforeYoutube() {
    global g_YoutubeSpotifyPausePending
    if (g_YoutubeSpotifyPausePending)
        return
    try {
        session := YouTube_FindSpotifyMediaSession()
        if !session
            return
        if (session.PlaybackStatus != Media.PlaybackStatus.Playing)
            return
        session.Pause()
        g_YoutubeSpotifyPausePending := true
    } catch {
        ; WinRT/SMTC unavailable — do not fall back to Media_Play_Pause (toggle bug).
    }
}

; Resume Spotify with SMTC Play() only when we paused it and session reports Paused.
YouTube_ResumeSpotifyAfterYoutubeIfPending(restoreHwnd := 0) {
    global g_YoutubeSpotifyPausePending
    if (!g_YoutubeSpotifyPausePending)
        return
    g_YoutubeSpotifyPausePending := false
    try {
        session := YouTube_FindSpotifyMediaSession()
        if (session && session.PlaybackStatus == Media.PlaybackStatus.Paused)
            session.Play()
    } catch {
        ;
    }
    if (restoreHwnd && WinExist("ahk_id " restoreHwnd)) {
        try
            WinActivate("ahk_id " restoreHwnd)
    }
}

; End YouTube focus session: pause YouTube (k), resume Spotify if pending, remove blackout, stop window monitor.
YouTube_EndFocusSession() {
    global g_YoutubeFocusTrackedHwnd, g_YoutubeFocusSessionActive
    restoreHwnd := WinExist("A")
    if (g_YoutubeFocusTrackedHwnd && WinExist("ahk_id " . g_YoutubeFocusTrackedHwnd)) {
        try {
            WinActivate("ahk_id " . g_YoutubeFocusTrackedHwnd)
            Sleep(50)
            Send("k")
            Sleep(100)
        }
    }
    YouTube_ResumeSpotifyAfterYoutubeIfPending(restoreHwnd)
    DisableFocusMode()
    StopYoutubeFocusMonitor()
    g_YoutubeFocusSessionActive := false
}

; Monitor tracked YouTube window: only when it is destroyed, run full session teardown (same as second hotkey).
MonitorYoutubeFocus() {
    global g_YoutubeFocusTrackedHwnd

    if (g_YoutubeFocusTrackedHwnd && !WinExist("ahk_id " . g_YoutubeFocusTrackedHwnd))
        YouTube_EndFocusSession()
}

; Start monitoring YouTube window focus
StartYoutubeFocusMonitor(hwnd := 0) {
    global g_YoutubeFocusMonitorTimer, g_YoutubeFocusTrackedHwnd

    StopYoutubeFocusMonitor()

    g_YoutubeFocusTrackedHwnd := hwnd ? hwnd : WinExist("A")
    if (!g_YoutubeFocusTrackedHwnd)
        return

    g_YoutubeFocusMonitorTimer := MonitorYoutubeFocus
    SetTimer(g_YoutubeFocusMonitorTimer, 200)
}

; Stop monitoring YouTube window focus
StopYoutubeFocusMonitor() {
    global g_YoutubeFocusMonitorTimer, g_YoutubeFocusTrackedHwnd

    if (g_YoutubeFocusMonitorTimer) {
        SetTimer(g_YoutubeFocusMonitorTimer, 0)
        g_YoutubeFocusMonitorTimer := false
    }
    g_YoutubeFocusTrackedHwnd := 0
}

ShowStudyTopicSelector() {
    global g_StudyTopicSelectorGui, g_StudyTopicSelectorActive, g_StudyTopics

    if (g_StudyTopicSelectorActive)
        return

    g_StudyTopicSelectorGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Owner")
    g_StudyTopicSelectorGui.BackColor := "1E1E2E"
    g_StudyTopicSelectorGui.MarginX := 20
    g_StudyTopicSelectorGui.MarginY := 15

    g_StudyTopicSelectorGui.SetFont("s14 cCDD6F4 Bold", "Segoe UI")
    g_StudyTopicSelectorGui.Add("Text", "w280 Center", "📚 Study topic (QuickLook)")
    g_StudyTopicSelectorGui.Add("Text", "w280 h1 Background45475A")

    g_StudyTopicSelectorGui.SetFont("s12 cCDD6F4", "Segoe UI")
    for num, topic in g_StudyTopics {
        g_StudyTopicSelectorGui.Add("Text", "w280", "[" . num . "] " . topic.name)
    }

    g_StudyTopicSelectorGui.Add("Text", "w280 h1 Background45475A y+10")
    g_StudyTopicSelectorGui.SetFont("s9 c6C7086", "Segoe UI")
    g_StudyTopicSelectorGui.Add("Text", "w280 Center", "Press 0–5 | Esc to cancel")

    activeWin := 0
    try {
        activeWin := WinGetID("A")
    } catch {
        activeWin := 0
    }

    MonitorGetWorkArea(1, &monitorLeft, &monitorTop, &monitorRight, &monitorBottom)
    monitorWidth := monitorRight - monitorLeft
    monitorHeight := monitorBottom - monitorTop

    if (activeWin && activeWin != 0) {
        rect := Buffer(16, 0)
        if (DllCall("GetWindowRect", "ptr", activeWin, "ptr", rect)) {
            winLeft := NumGet(rect, 0, "int")
            winTop := NumGet(rect, 4, "int")
            winRight := NumGet(rect, 8, "int")
            winBottom := NumGet(rect, 12, "int")
            centerX := winLeft + (winRight - winLeft) // 2
            centerY := winTop + (winBottom - winTop) // 2
            monitorCount := MonitorGetCount()
            loop monitorCount {
                idx := A_Index
                MonitorGetWorkArea(idx, &l, &t, &r, &b)
                if (centerX >= l && centerX <= r && centerY >= t && centerY <= b) {
                    monitorLeft := l
                    monitorTop := t
                    monitorRight := r
                    monitorBottom := b
                    monitorWidth := r - l
                    monitorHeight := b - t
                    break
                }
            }
        }
    }

    g_StudyTopicSelectorGui.Show("AutoSize Hide")
    g_StudyTopicSelectorGui.GetPos(&gx, &gy, &gw, &gh)
    cx := monitorLeft + (monitorWidth - gw) // 2
    cy := monitorTop + (monitorHeight - gh) // 2
    g_StudyTopicSelectorGui.Show("x" . cx . " y" . cy . " NA")

    g_StudyTopicSelectorActive := true
    Hotkey("0", StudyTopicSelector_HandleKey, "On")
    Hotkey("1", StudyTopicSelector_HandleKey, "On")
    Hotkey("2", StudyTopicSelector_HandleKey, "On")
    Hotkey("3", StudyTopicSelector_HandleKey, "On")
    Hotkey("4", StudyTopicSelector_HandleKey, "On")
    Hotkey("5", StudyTopicSelector_HandleKey, "On")
    Hotkey("Escape", StudyTopicSelector_Cancel, "On")
}

StudyTopicSelector_HandleKey(key) {
    global g_StudyTopicSelectorActive, g_StudyTopics

    if (!g_StudyTopicSelectorActive)
        return
    selection := Integer(key)
    StudyTopicSelector_Close()
    if (!g_StudyTopics.Has(selection))
        return

    topic := g_StudyTopics[selection]
    basePath := GetNotesRepoPath()
    if (basePath = "") {
        try ShowCenteredOverlay_Utils("⚠ Notes repo path not set (env.ahk).", 3000, BANNER_ACCENT_INTERMEDIATE)
        return
    }
    fullPath := RTrim(basePath, "\") . topic.path
    ; Derive Markdown path from the PDF path by replacing the extension.
    if (StrLower(SubStr(fullPath, -3)) = "pdf") {
        mdPath := SubStr(fullPath, 1, StrLen(fullPath) - 3) . "md"
    } else {
        mdPath := fullPath
    }
    if (!FileExist(mdPath)) {
        try ShowCenteredOverlay_Utils("❌ Markdown not found: " mdPath, 3500, BANNER_ACCENT_ERROR)
        return
    }
    QuickLook_OpenPath(mdPath)
}

StudyTopicSelector_Cancel(*) {
    StudyTopicSelector_Close()
}

StudyTopicSelector_Close() {
    global g_StudyTopicSelectorGui, g_StudyTopicSelectorActive

    if (!g_StudyTopicSelectorActive)
        return
    g_StudyTopicSelectorActive := false
    try Hotkey("0", "Off")
    try Hotkey("1", "Off")
    try Hotkey("2", "Off")
    try Hotkey("3", "Off")
    try Hotkey("4", "Off")
    try Hotkey("5", "Off")
    try Hotkey("Escape", StudyTopicSelector_Cancel, "Off")
    if (IsObject(g_StudyTopicSelectorGui) && g_StudyTopicSelectorGui.Hwnd) {
        try g_StudyTopicSelectorGui.Destroy()
    }
    g_StudyTopicSelectorGui := false
}

PeekPdf_GetIniPath() {
    return A_ScriptDir "\data\peek_pdf.ini"
}

PeekPdf_NormalizePath(path) {
    path := Trim(path)
    q := Chr(34)
    if (SubStr(path, 1, 1) = q && SubStr(path, -1) = q)
        path := SubStr(path, 2, StrLen(path) - 2)
    return Trim(path)
}

; Resolve the Peek executable path.
; Priority: 1) INI [Peek] ExePath  2) Environment-specific path (GetPeekExePath)  3) "peek.exe" (PATH)
PeekPdf_ResolvePeekExePath() {
    iniPath := PeekPdf_GetIniPath()
    exePath := ""
    try exePath := IniRead(iniPath, "Peek", "ExePath", "")
    exePath := PeekPdf_NormalizePath(exePath)
    if (exePath != "" && FileExist(exePath))
        return exePath
    ; Legacy: this previously used GetPeekExePath() and PowerToys Peek.
    ; QuickLook is now the primary study viewer; for compatibility, fall back to QuickLook.
    envExe := GetQuickLookExePath()
    if (FileExist(envExe))
        return envExe
    return "QuickLook.exe"
}

QuickLook_GetIniPath() {
    return A_ScriptDir "\data\quicklook.ini"
}

QuickLook_NormalizePath(path) {
    path := Trim(path)
    q := Chr(34)
    if (SubStr(path, 1, 1) = q && SubStr(path, -1) = q)
        path := SubStr(path, 2, StrLen(path) - 2)
    return Trim(path)
}

; Note: `GetQuickLookExePath()` is defined in `env.ahk` (environment-specific paths).

QuickLook_ResolveExePath() {
    iniPath := QuickLook_GetIniPath()
    exePath := ""
    try exePath := IniRead(iniPath, "QuickLook", "ExePath", "")
    exePath := QuickLook_NormalizePath(exePath)
    if (exePath != "" && FileExist(exePath))
        return exePath
    envExe := GetQuickLookExePath()
    if (FileExist(envExe))
        return envExe
    return "QuickLook.exe"
}

; Detect whether a given monitor index is currently connected and has a usable work area.
; Uses AHK built-ins so we don't depend on custom geometry assumptions.
IsMonitorConnected(monitorIndex) {
    try monitorCount := MonitorGetCount()
    catch
        return false
    if (monitorIndex < 1 || monitorIndex > monitorCount)
        return false
    try {
        MonitorGetWorkArea(monitorIndex, &l, &t, &r, &b)
        if ((r - l) <= 0 || (b - t) <= 0)
            return false
    } catch {
        return false
    }
    return true
}

; Map a Windows display device name (e.g. "\\.\DISPLAY2") to the current AHK monitor index.
; Returns 0 when not found/connected.
GetMonitorIndexByDeviceName(deviceName) {
    if (deviceName = "")
        return 0
    try monitorCount := MonitorGetCount()
    catch
        return 0
    loop monitorCount {
        idx := A_Index
        nm := ""
        try nm := MonitorGetName(idx)
        if (nm = deviceName)
            return idx
    }
    return 0
}

; Preferred target for "main monitor (2)" in this setup: Windows DISPLAY2.
; Falls back to the primary monitor if DISPLAY2 isn't present.
GetQuickLookTargetMonitorIndex() {
    idx := GetMonitorIndexByDeviceName("\\.\DISPLAY2")
    if (idx && IsMonitorConnected(idx))
        return idx
    try return MonitorGetPrimary()
    catch
        return 1
}

QuickLook_OpenPath(path) {
    quickLookExe := QuickLook_ResolveExePath()
    if (!FileExist(quickLookExe)) {
        try ShowCenteredOverlay_Utils("❌ QuickLook executable not found: " quickLookExe, 2500, BANNER_ACCENT_ERROR)
        return
    }
    if (!FileExist(path)) {
        try ShowCenteredOverlay_Utils("❌ Markdown not found: " path, 3500, BANNER_ACCENT_ERROR)
        return
    }
    try {
        Run('"' quickLookExe '" "' path '"')
    } catch as e {
        try ShowCenteredOverlay_Utils("❌ Failed to open QuickLook: " e.Message, 3000, BANNER_ACCENT_ERROR)
        return
    }
    if WinWait("ahk_exe QuickLook.exe", , 2) {
        hwnd := WinExist("ahk_exe QuickLook.exe")
        if (hwnd) {
            ; Always prefer the configured physical display (\\.\DISPLAY2); fallback to primary.
            targetMon := GetQuickLookTargetMonitorIndex()

            MoveWindowToMonitor(hwnd, targetMon)
            WinMaximize("ahk_id " hwnd)

            ; Wait until QuickLook's window rect center is on the target monitor (move/maximize can be async).
            ; (Condition-based wait; avoids relying on blind fixed delays.)
            moveOk := false
            deadline := A_TickCount + 1500
            remapIter := 0
            while (A_TickCount < deadline) {
                try {
                    rect := Buffer(16, 0)
                    if (DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect)) {
                        wl := NumGet(rect, 0, "int"), wt := NumGet(rect, 4, "int")
                        wr := NumGet(rect, 8, "int"), wb := NumGet(rect, 12, "int")
                        cx := wl + (wr - wl) // 2
                        cy := wt + (wb - wt) // 2
                        MonitorGetWorkArea(targetMon, &ml, &mt, &mr, &mb)
                        if (cx >= ml && cx <= mr && cy >= mt && cy <= mb) {
                            moveOk := true
                            break
                        }
                    }
                } catch {
                    ; ignore
                }
                remapIter++
                Sleep 50
            }

            ; Ensure QuickLook is the active window before enabling blackout (blackout uses active window monitor).
            try {
                WinShow("ahk_id " hwnd)
                WinActivate("ahk_id " hwnd)
                WinWaitActive("ahk_id " hwnd, , 1)
            } catch {
                ; ignore
            }

            ; Execute blackout only after move is verified and QuickLook is active.
            EnableFocusMode()
            StartPdfFocusMonitor(hwnd, "Debounced")

            ; Click inside QuickLook to ensure the markdown viewer control has keyboard focus.
            try {
                rect2 := Buffer(16, 0)
                if (DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect2)) {
                    wl2 := NumGet(rect2, 0, "int"), wt2 := NumGet(rect2, 4, "int")
                    wr2 := NumGet(rect2, 8, "int"), wb2 := NumGet(rect2, 12, "int")
                    clickX := wl2 + (wr2 - wl2) // 2
                    clickY := wt2 + (wb2 - wt2) // 2
                    CoordMode("Mouse", "Screen")
                    Click clickX, clickY
                }
            } catch {
                ; ignore
            }

            ; Send Ctrl+End using a foreground SendInput sequence (manual Ctrl+End works, so method matters).
            try {
                ; Ensure QuickLook is still active right before sending.
                WinActivate("ahk_id " hwnd)
                WinWaitActive("ahk_id " hwnd, , 1)

                ; Preferred: SendInput to foreground window.
                SendInput("^{End}")
            } catch {
                ; Fallback: explicit down/up (some apps are picky about chord timing)
                try {
                    WinActivate("ahk_id " hwnd)
                    WinWaitActive("ahk_id " hwnd, , 1)
                    Send("{Ctrl down}{End}{Ctrl up}")
                } catch {
                    ; Last resort: ControlSend (often unreliable for WPF, but keep as backup)
                    try {
                        ControlSend("^End", "ahk_id " hwnd)
                    } catch {
                    }
                }
            }
        }
    }
}

; Open a specific PDF in PowerToys Peek and run WaitAndConfigure. Caller must validate pdfPath and exe exist.
; skipGoToLastPage: if true, do not navigate to the last page (e.g. for short docs like technique README).
PeekPdf_OpenPath(pdfPath, skipGoToLastPage := false) {
    peekExe := PeekPdf_ResolvePeekExePath()
    peekEsc := StrReplace(peekExe, "'", "''")
    pdfEsc := StrReplace(pdfPath, "'", "''")
    psArg := "& " . Chr(39) . peekEsc . Chr(39) . " " . Chr(39) . pdfEsc . Chr(39)
    cmd := "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command " . Chr(34) . psArg . Chr(34)
    try Run cmd, "", "Hide"
    catch as e {
        try ShowCenteredOverlay_Utils("❌ Failed to open Peek: " e.Message, 3000, BANNER_ACCENT_ERROR)
        return
    }
    if WinWait("Peek", "", 5) {
        hwnd := WinExist("ahk_exe PowerToys.Peek.UI.exe")
        if (!hwnd)
            hwnd := WinExist("Peek")
        MoveWindowToMonitor(hwnd, 2)
        WinMaximize("ahk_id " hwnd)
        PeekPdf_WaitAndConfigure(skipGoToLastPage)
        EnableFocusMode()
        StartPdfFocusMonitor()
    }
}

PeekPdf_OpenStored() {
    iniPath := PeekPdf_GetIniPath()

    ; Resolve PDF path per environment, with legacy fallback
    pdfPath := ""
    try {
        if (IS_WORK_ENVIRONMENT) {
            pdfPath := IniRead(iniPath, "Peek", "PdfPathWork", "")
            if (pdfPath = "")
                pdfPath := IniRead(iniPath, "Peek", "PdfPath", "")
        } else {
            pdfPath := IniRead(iniPath, "Peek", "PdfPathPersonal", "")
            if (pdfPath = "")
                pdfPath := IniRead(iniPath, "Peek", "PdfPath", "")
        }
    }

    pdfPath := PeekPdf_NormalizePath(pdfPath)
    if (pdfPath = "") {
        try ShowCenteredOverlay_Utils("⚠ No PDF path set. Hold Win+Alt+Shift+X to set.", 3000,
            BANNER_ACCENT_INTERMEDIATE)
        return
    }
    peekExe := PeekPdf_ResolvePeekExePath()
    if (!FileExist(peekExe)) {
        try ShowCenteredOverlay_Utils("❌ Peek executable not found.", 2500, BANNER_ACCENT_ERROR)
        return
    }
    if (!FileExist(pdfPath)) {
        try ShowCenteredOverlay_Utils("❌ PDF file not found: " pdfPath, 3500, BANNER_ACCENT_ERROR)
        return
    }
    peekEsc := StrReplace(peekExe, "'", "''")
    pdfEsc := StrReplace(pdfPath, "'", "''")
    psArg := "& " . Chr(39) . peekEsc . Chr(39) . " " . Chr(39) . pdfEsc . Chr(39)
    cmd := "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command " . Chr(34) . psArg . Chr(34)
    try Run cmd, "", "Hide"
    catch as e {
        try ShowCenteredOverlay_Utils("❌ Failed to open Peek: " e.Message, 3000, BANNER_ACCENT_ERROR)
        return
    }
    if WinWait("Peek", "", 5) {
        hwnd := WinExist("ahk_exe PowerToys.Peek.UI.exe")
        if (!hwnd)
            hwnd := WinExist("Peek")
        MoveWindowToMonitor(hwnd, 2)
        WinMaximize("ahk_id " hwnd)
        PeekPdf_WaitAndConfigure()
    }
}

MoveWindowToMonitor(hwnd, monitorIndex := 2) {
    if (!hwnd)
        return
    try {
        MonitorGetWorkArea(monitorIndex, &l, &t, &r, &b)
    } catch {
        return
    }
    w := r - l
    h := b - t
    try WinMove(l, t, w, h, "ahk_id " hwnd)
}

; Wait for Peek PDF toolbar to load (Page view button), click it, two-page view, focus; optionally go to last page.
; Current state: PDF opening and window maximization are working correctly.
; Execution order: 1) Get Peek hwnd  2) UIA root from hwnd  3) Poll for "Page view" anchor
;  4) Wait for anchor visible + short delay before click  5) Click Page view  6) Right Arrow
;  7) Click window center  8) If not skipGoToLastPage: go to last page (UIA or Ctrl+End). Fallback: Sleep 400 + Click if UIA or anchor fails.
PeekPdf_WaitAndConfigure(skipGoToLastPage := false) {
    global UIA
    ; Standard loading bar: show for the whole process so user knows when we started and when we finished
    StandardLoadingBar_Show("⏳ Peek PDF: configuring...", BANNER_ACCENT_INTERMEDIATE)
    ; 1) Get Peek window hwnd
    hwnd := WinExist("Peek")
    if (!hwnd)
        hwnd := WinExist("ahk_exe PowerToys.Peek.UI.exe")
    if (!hwnd) {
        StandardLoadingBar_Update("❌ Peek PDF: window not found", BANNER_ACCENT_ERROR)
        StandardLoadingBar_Hide(2000)
        Sleep 300
        Click "Left"
        return
    }
    try {
        ; 2) UIA root
        el := UIA.ElementFromHandle(hwnd)
        ; 3) Poll for Page view (layouts) anchor (up to 20s to accommodate Peek load > 5s)
        ; FindFirst throws when no element matches; catch so we keep polling instead of exiting.
        pageViewBtn := ""
        pollIter := 0
        loop 80 {
            pollIter := A_Index
            try
                pageViewBtn := el.FindFirst({ Type: 50000, Name: "Page view", AutomationId: "layouts" })
            catch
                pageViewBtn := ""
            if (pageViewBtn)
                break
            Sleep 150
        }
        if (pageViewBtn) {
            ; 4) Ensure anchor is visible, then short delay so toolbar is ready before click
            visIter := 0
            loop 20 {
                visIter := A_Index
                try {
                    if (!pageViewBtn.GetPropertyValue(UIA.Property.IsOffscreen)) {
                        br := pageViewBtn.BoundingRectangle
                        if (IsObject(br) && (br.r - br.l) > 0 && (br.b - br.t) > 0)
                            break
                    }
                } catch {
                }
                Sleep 30
            }
            ; Short delay so toolbar is ready before clicking Page view
            Sleep 1000
            ; 5) Click Page view button
            invokeOk := false
            clickOk := false
            try {
                pageViewBtn.Invoke()
                invokeOk := true
            } catch as invErr {
                try {
                    pageViewBtn.Click()
                    clickOk := true
                } catch as clickErr {
                }
            }
            Sleep 600
            ; 6) Select "Two page" from the open Page view menu (main window + foreground popup; else ControlSend Right to Peek)
            fgHwnd := 0
            try fgHwnd := WinGetID("A")

            twoPageEl := ""
            twoPageClicked := false

            twoPageScope := "none"
            menuRect := ""
            try {
                abr := pageViewBtn.BoundingRectangle
                if (IsObject(abr))
                    menuRect := { l: abr.l - 700, t: abr.b, r: abr.r + 700, b: abr.b + 650 }
            } catch {
                menuRect := ""
            }

            ; Search in active window region (menu is visible on screen but may not be exposed as Buttons).
            try {
                elActive := UIA.ElementFromHandle(fgHwnd ? fgHwnd : hwnd)
                if (IsObject(menuRect)) {
                    for cand in elActive.FindAll({ IsOffscreen: 0 }) {
                        try {
                            br := cand.BoundingRectangle
                            if (!IsObject(br))
                                continue
                            inRegion := (br.l < menuRect.r && br.r > menuRect.l && br.t < menuRect.b && br.b > menuRect
                                .t)
                            if (!inRegion)
                                continue
                            nm := ""
                            try nm := cand.Name
                            if (nm != "" && InStr(nm, "Two page")) {
                                twoPageEl := cand
                                twoPageScope := (fgHwnd ? "active_region" : "peek_region")
                                break
                            }
                        } catch {
                        }
                    }
                }
            } catch {
            }

            if (twoPageEl) {
                try {
                    ; Click by coordinates for maximum compatibility (works even if Invoke/Click patterns are absent).
                    br := twoPageEl.BoundingRectangle
                    if (IsObject(br)) {
                        cx := br.l + (br.r - br.l) // 2
                        cy := br.t + (br.b - br.t) // 2
                        Click cx, cy
                        twoPageClicked := true
                    }
                } catch {
                    twoPageClicked := false
                }
            } else {
                ; Last-resort: keystroke fallback (kept for robustness)
                try ControlSend "{Right}", "ahk_id " hwnd
                catch
                    Send "{Right}"
            }

            Sleep 400
            ; 7) Click center of Peek window (focus)
            WinGetPos(&wx, &wy, &ww, &wh, "ahk_id " hwnd)
            if (ww > 0 && wh > 0) {
                cx := wx + ww // 2
                cy := wy + wh // 2
                Click cx, cy
            }
            Sleep 150
            ; 8) Go to final page via UIA (keystrokes do not reach the embedded Edge PDF viewer); skip if skipGoToLastPage
            if (!skipGoToLastPage) {
                lastPageSet := false
                try {
                    for doc in el.FindAll({ Type: 50030 }) {
                        try {
                            nm := doc.Name
                            ; Match "containing N pages" (EN) or "N pages"/"N páginas" (avoid "Page 1")
                            if (RegExMatch(nm, "containing\s+(\d+)\s+pages", &m) || RegExMatch(nm,
                                "document.*?(\d+)\s*(?:pages|páginas)", &m)) {
                                totalPages := Integer(m[1])
                                if (totalPages > 0) {
                                    pageSel := el.FindFirst({ Type: 50004, AutomationId: "pageselector" })
                                    if (pageSel) {
                                        try {
                                            pageSel.SetFocus()
                                            Sleep(200)
                                            WinActivate("ahk_id " hwnd)
                                            Sleep(120)
                                            Send("^a")
                                            Sleep(50)
                                            Send(String(totalPages))
                                            Sleep(50)
                                            Send("{Enter}")
                                            lastPageSet := true
                                        } catch {
                                            ; ignore focus/send errors
                                        }
                                    }
                                    break
                                }
                            }
                        } catch {
                            ; ignore per-doc errors
                        }
                    }
                } catch {
                    ; ignore UIA errors for last-page navigation
                }
                if (!lastPageSet) {
                    try
                        ControlSend("^End", "ahk_id " hwnd)
                    catch
                        Send("^End")
                }
            }
            Sleep 100
            StandardLoadingBar_Update("✅ Peek PDF: done", BANNER_ACCENT_SUCCESS)
            StandardLoadingBar_Hide(2000)
        } else {
            StandardLoadingBar_Update("✅ Peek PDF: finished (fallback)", BANNER_ACCENT_SUCCESS)
            StandardLoadingBar_Hide(2000)
            Sleep 400
            Click "Left"
        }
    } catch {
        StandardLoadingBar_Update("✅ Peek PDF: finished (fallback)", BANNER_ACCENT_SUCCESS)
        StandardLoadingBar_Hide(2000)
        Sleep 400
        Click "Left"
    }
}

#!+x::
{
    hwnd := WinExist("ahk_exe QuickLook.exe")
    if hwnd {
        try {
            WinShow("ahk_id " hwnd)
            WinActivate("ahk_id " hwnd)
            WinWaitActive("ahk_id " hwnd, , 1)
        }
        EnableFocusMode()
        StartPdfFocusMonitor(hwnd, "Debounced")
        ; Click inside QuickLook to ensure content receives keyboard focus
        try {
            rect := Buffer(16, 0)
            if (DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect)) {
                wl := NumGet(rect, 0, "int"), wt := NumGet(rect, 4, "int")
                wr := NumGet(rect, 8, "int"), wb := NumGet(rect, 12, "int")
                clickX := wl + (wr - wl) // 2
                clickY := wt + (wb - wt) // 2
                CoordMode("Mouse", "Screen")
                Click clickX, clickY
            }
        } catch {
            ; ignore
        }
        try
            ControlSend("^End", "ahk_id " hwnd)
        catch {
            ; Fallback to active-window keystroke if ControlSend fails
            try {
                WinActivate("ahk_id " hwnd)
                WinWaitActive("ahk_id " hwnd, , 1)
            } catch {
                ; ignore
            }
            try Send("^End")
        }
        return
    }
    ShowStudyTopicSelector()
}

; =============================================================================
; Hotstring Selector System
; =============================================================================
; PURPOSE: Provides a modal GUI-based interface for accessing hotstrings, quick-open files,
;          and executable macros via single-character keyboard shortcuts.
;
; HOTKEY: Windows + Alt + Shift + U (#!+U)
;
; FUNCTIONALITY:
;   - Displays categorized list of available actions (Prompts, General, Projects, Files & Links, Macros)
;   - Each action is assigned a unique character from g_HotstringCharSequence
;   - User presses assigned character to execute corresponding action
;   - Actions include: text expansion (hotstrings), file opening, macro execution
;
; ARCHITECTURE:
;   - Character-to-action mapping built dynamically via BuildHotstringCharMap()
;   - Character assignments follow sequential order within each category
;   - Explicit character assignments (via RegisterMacro/RegisterHotstring char parameter) take precedence
;   - GUI adapts to monitor configuration (landscape/portrait, resolution, scaling)
;
; =============================================================================

; Global state variables for hotstring selector system
global g_HotstringSelectorGui := false          ; GUI object reference (false when not initialized)
global g_HotstringSelectorActive := false       ; Boolean flag indicating selector is currently displayed
global g_HotstringCharMap := Map()              ; Character-to-text-expansion mapping for hotstrings
global g_UtilityHotstringCharMapByCategory := Map() ; Category -> Map(char -> expansion) used by Utility Shortcuts
global g_HotstringHotkeyHandlers := []          ; Array of hotkey handler objects for cleanup on close
global g_HotstringPromptCharMap := Map()        ; Map of prompt-assigned chars => true (rebuilt on each ShowHotstringSelector)
global g_HotstringGeminiArmed := false          ; When true, next Prompts selection is redirected to Gemini
global g_HotstringGeminiAutoSubmit := true      ; During delayed flow: true = send Enter after paste; false = paste only
global g_HotstringGeminiSubmitTimer := false   ; Timer reference for 4s delayed submit (for cleanup if needed)
global g_HotstringGeminiRestoreHwnd := 0        ; Window to restore focus to after paste (set at start of GeminiDelayedSubmitFlow)

; Utility selector hierarchy state
global g_UtilitySelectorMode := "top"           ; "top" | "category"
global g_UtilitySelectorCategory := ""          ; One of g_UtilityTopCategories

; Top-level categories (numbers 1-6 select these)
global g_UtilityTopCategories := ["Prompts", "Projects", "Macros", "General", "Links", "Hotstrings"]
; Top-level trigger keys (lowercase so UtilitySelector_RebindHotkeys auto-binds uppercase too)
global g_UtilityTopCategoryById := Map("r", "Prompts", "p", "Projects", "m", "Macros", "g", "General", "l", "Links",
    "h", "Hotstrings")

; Utility selector cached UI data (rebuilt each time ShowHotstringSelector() runs)
global g_UtilitySelectorAllItems := []          ; Array of {category, char, text, isEmpty, [explicitIndex]}
global g_UtilitySelectorIsPortrait := false
global g_UtilitySelectorMonitor := Map()        ; {left, top, width, height}
global g_UtilitySelectorTitleCtrl := false
global g_UtilitySelectorEditCtrl := false
global g_UtilitySelectorFontSize := 9           ; Base point size for RichEdit rendering (set on ShowHotstringSelector)

; Character assignment sequence: defines order in which characters are assigned to actions
; Format: ["1", "2", "3", "4", "5", "q", "w", "e", "r", "t", "a", "s", "d", "f", "g", "z", "x",
;          "c", "v", "b", "6", "7", "8", "9", "0", "y", "u", "i", "o", "p", "h", "j", "k", "l", "n", "m", ",", "."]
global g_HotstringCharSequence := ["1", "2", "3", "4", "5", "q", "w", "e", "r", "t", "a", "s", "d", "f", "g", "z", "x",
    "c", "v", "b", "6", "7", "8", "9", "0", "y", "u", "i", "o", "p", "h", "j", "k", "l", "n", "m", ",", "."]

; Category display order: defines the sequence in which action categories appear in the GUI
; Order: Prompts → General → Projects → Links → Macros → Hotstrings
; Note: Utility-only views are not included here.
; "Hotstrings" must be present so BuildHotstringCharMap() populates g_UtilityHotstringCharMapByCategory["Hotstrings"].
global g_HotstringCategories := ["Prompts", "General", "Projects", "Files & Links", "Macros", "Hotstrings"]

; Reserved empty character: never assigned to any action; always shows as (empty) in selector
; Set to "" to disable reservation
global g_ReservedEmptyChar := ""

; =============================================================================
; RichEdit helpers (mnemonic emphasis for selectors)
; =============================================================================
global g_MnemonicRichDll := 0

MnemonicRich_EnsureDll() {
    global g_MnemonicRichDll
    ; msftedit.dll must be loaded before creating RichEdit50W controls (ClassRichEdit50W).
    if (!g_MnemonicRichDll)
        g_MnemonicRichDll := DllCall("LoadLibrary", "str", "msftedit.dll", "ptr")
}

; UTF-16 code unit count for RichEdit character indices (BMP = 1, supplementary = 2).
MnemonicRich_Utf16Units(s) {
    n := 0
    for c in StrSplit(s, "") {
        o := Ord(c)
        n += (o > 0xFFFF) ? 2 : 1
    }
    return n
}

; EM_SETTEXTEX = 0x461, ST_UNICODE = 8 — RichEdit’s native UTF-16 path.
MnemonicRich_SetPlainUtf16(ctrl, plain) {
    hwnd := ctrl.Hwnd
    flags := 8 ; ST_UNICODE
    cp := 1200
    settextex := Buffer(8, 0)
    NumPut("uint", flags, settextex, 0)
    NumPut("uint", cp, settextex, 4)
    if (plain = "") {
        emptyBuf := Buffer(2, 0)
        SendMessage(0x461, settextex.Ptr, emptyBuf.Ptr, hwnd)
        return
    }
    textBuf := Buffer((StrLen(plain) + 1) * 2)
    StrPut(plain, textBuf, "UTF-16")
    SendMessage(0x461, settextex.Ptr, textBuf.Ptr, hwnd)
}

MnemonicRich_ThemingOff(ctrl) {
    hwnd := ctrl.Hwnd
    DllCall("uxtheme\SetWindowTheme", "ptr", hwnd, "wstr", "", "wstr", "")
    parent := DllCall("GetParent", "ptr", hwnd, "ptr")
    if (parent)
        DllCall("uxtheme\SetWindowTheme", "ptr", parent, "wstr", "", "wstr", "")
}

; CHARFORMAT2W buffer (116 bytes). textColor is COLORREF (0x00BBGGRR).
MnemonicRich_CharFormat2(faceName, pt, textColor, bold := false) {
    yh := Round(pt * 20)
    cf := Buffer(116, 0)
    NumPut("uint", 116, cf, 0) ; cbSize
    mask := 0x40000000 | 0x80000000 | 0x20000000 | 0x1 ; CFM_COLOR | CFM_SIZE | CFM_FACE | CFM_BOLD
    NumPut("uint", mask, cf, 4) ; dwMask
    NumPut("uint", bold ? 0x1 : 0, cf, 8) ; dwEffects
    NumPut("int", yh, cf, 12) ; yHeight (twips)
    NumPut("int", 0, cf, 16) ; yOffset
    NumPut("uint", textColor, cf, 20) ; crTextColor
    NumPut("uchar", 1, cf, 24) ; bCharSet DEFAULT_CHARSET
    NumPut("uchar", 0, cf, 25) ; bPitchAndFamily
    StrPut(faceName, cf.Ptr + 26, 64, "UTF-16")
    return cf
}

MnemonicRich_ApplyCharFormat(ctrl, scopeAll, cfBuf) {
    w := scopeAll ? 4 : 1 ; SCF_ALL = 4, SCF_SELECTION = 1
    return SendMessage(0x444, w, cfBuf.Ptr, ctrl.Hwnd) ; EM_SETCHARFORMAT
}

MnemonicRich_SetSel(hwnd, cpMin, cpMax) {
    return SendMessage(0xB1, cpMin, cpMax, hwnd) ; EM_SETSEL
}

; Render lines (joined by CR only) and emphasize mnemonic letter (bumpPx) in [key] and first title occurrence.
MnemonicRich_Render(ctrl, lines, basePt, bumpPx := 6, faceName := "Segoe UI", rgbHex := "CDD6F4", bgHex := "1E1E2E") {
    MnemonicRich_EnsureDll()
    MnemonicRich_ThemingOff(ctrl)

    ; Convert RGB hex (RRGGBB) to COLORREF (0x00BBGGRR).
    rr := Integer("0x" . SubStr(rgbHex, 1, 2))
    gg := Integer("0x" . SubStr(rgbHex, 3, 2))
    bb := Integer("0x" . SubStr(rgbHex, 5, 2))
    textColor := (bb << 16) | (gg << 8) | rr

    br := Integer("0x" . SubStr(bgHex, 1, 2))
    bg := Integer("0x" . SubStr(bgHex, 3, 2))
    bb2 := Integer("0x" . SubStr(bgHex, 5, 2))
    bgColor := (bb2 << 16) | (bg << 8) | br

    bumpPt := bumpPx * 72 / 96
    bigPt := basePt + bumpPt

    plain := ""
    spans := [] ; {start,len} in UTF-16 units
    u16Pos := 0
    first := true

    RenderTitleKey(lineText, key, baseU16) {
        if (key = "")
            return
        rb := InStr(lineText, "]")
        if (!rb)
            return
        after := SubStr(lineText, rb + 1)
        tpos := InStr(after, key, false)
        if (!tpos)
            tpos := InStr(after, StrUpper(key), false)
        if (!tpos)
            return
        preNoLast := SubStr(lineText, 1, rb + tpos - 1)
        spans.Push({ start: baseU16 + MnemonicRich_Utf16Units(preNoLast), len: 1 })
    }

    for ln in lines {
        if (!first) {
            plain .= "`r"
            u16Pos += 1
        }
        first := false

        lineText := ln.text
        key := ln.HasProp("key") ? ln.key : ""
        RenderTitleKey(lineText, key, u16Pos)

        ; Optional right-side key emphasis for two-column layouts.
        if (ln.HasProp("keyRight") && ln.keyRight != "" && ln.HasProp("rightStartCharPos") && ln.rightStartCharPos > 1) {
            rightStart := ln.rightStartCharPos
            prefix := SubStr(lineText, 1, rightStart - 1)
            rightText := SubStr(lineText, rightStart)
            baseRightU16 := u16Pos + MnemonicRich_Utf16Units(prefix)
            RenderTitleKey(rightText, ln.keyRight, baseRightU16)
        }

        plain .= lineText
        u16Pos += MnemonicRich_Utf16Units(lineText)
    }

    MnemonicRich_SetPlainUtf16(ctrl, plain)

    hwnd := ctrl.Hwnd
    SendMessage(0x4CF, 0, 0, hwnd) ; EM_SETREADONLY FALSE while formatting
    SendMessage(0x443, 0, bgColor, hwnd) ; EM_SETBKGNDCOLOR

    baseCf := MnemonicRich_CharFormat2(faceName, basePt, textColor, false)
    MnemonicRich_SetSel(hwnd, 0, -1)
    MnemonicRich_ApplyCharFormat(ctrl, false, baseCf)

    bigCf := MnemonicRich_CharFormat2(faceName, bigPt, textColor, false)
    for sp in spans {
        if (sp.len <= 0)
            continue
        MnemonicRich_SetSel(hwnd, sp.start, sp.start + sp.len)
        MnemonicRich_ApplyCharFormat(ctrl, false, bigCf)
    }
    MnemonicRich_SetSel(hwnd, 0, 0)
    SendMessage(0xB7, 0, 0, hwnd) ; EM_SCROLLCARET
    SendMessage(0x4CF, 1, 0, hwnd) ; EM_SETREADONLY TRUE
    SendMessage(0x443, 0, bgColor, hwnd) ; RichEdit can reset bg after readonly
}

; =============================================================================
; BuildHotstringCharMap()
; =============================================================================
; PURPOSE: Constructs character-to-action mappings for all registered items (hotstrings, files, macros)
;          and assigns characters sequentially within each category according to g_HotstringCharSequence.
;
; PROCESS:
;   1. Groups hotstrings by category (Projects, Prompts, General)
;   2. Processes each category in g_HotstringCategories order:
;      - Files & Links: Maps characters to file paths for quick-open functionality
;      - Macros: Maps characters to executable macro functions (explicit assignments first, then sequential)
;      - Other categories: Maps characters to hotstring expansion text
;   3. Returns Map of character → expansion text for hotstrings
;
; RETURNS: Map object where keys are characters and values are expansion text strings
; SIDE EFFECTS: Populates global maps g_QuickOpenFileCharMap and g_MacroCharMap
; =============================================================================
BuildHotstringCharMap() {
    global g_hotstrings, g_QuickOpenFiles, g_HotstringCategories, g_Macros
    charMap := Map()
    global g_QuickOpenFileCharMap := Map()
    global g_MacroCharMap := Map()
    global g_UtilityHotstringCharMapByCategory

    ; Category-scoped hotstring maps used by Utility Shortcuts selector.
    ; This allows the same char to exist in multiple categories (e.g. Prompts 'a' and Projects 'a').
    g_UtilityHotstringCharMapByCategory := Map()
    g_UtilityHotstringCharMapByCategory["Prompts"] := Map()
    g_UtilityHotstringCharMapByCategory["Projects"] := Map()
    g_UtilityHotstringCharMapByCategory["General"] := Map()
    g_UtilityHotstringCharMapByCategory["Hotstrings"] := Map()

    ; Group hotstrings by category
    categorized := Map()
    categorized["Projects"] := []
    categorized["Prompts"] := []
    categorized["General"] := []
    categorized["Hotstrings"] := []

    if (IsSet(g_hotstrings) && g_hotstrings.Length > 0) {
        for hs in g_hotstrings {
            category := hs.category
            if (category = "Projects" || category = "Prompts" || category = "General" || category = "Hotstrings") {
                categorized[category].Push(hs)
            } else {
                categorized["General"].Push(hs)
            }
        }
    }

    ; Assign characters sequentially within each category
    charIndex := 1
    for category in g_HotstringCategories {
        if (category = "Files & Links") {
            ; Handle quick open files
            if (IsSet(g_QuickOpenFiles) && g_QuickOpenFiles.Length > 0) {
                for fileEntry in g_QuickOpenFiles {
                    while (charIndex <= g_HotstringCharSequence.Length && g_ReservedEmptyChar != "" &&
                        g_HotstringCharSequence[charIndex] = g_ReservedEmptyChar)
                        charIndex++
                    if (charIndex <= g_HotstringCharSequence.Length) {
                        char := g_HotstringCharSequence[charIndex]
                        g_QuickOpenFileCharMap[char] := fileEntry.filePath
                        charIndex++
                    }
                }
            }
        } else if (category = "Macros") {
            ; Handle macros
            if (IsSet(g_Macros) && g_Macros.Length > 0) {
                ; First pass: assign macros with explicit character assignments
                for macroEntry in g_Macros {
                    if (macroEntry.HasProp("char") && macroEntry.char != "" && (g_ReservedEmptyChar = "" || macroEntry.char !=
                        g_ReservedEmptyChar)) {
                        ; Check if character is in the sequence and not already assigned
                        charIndexInSequence := 0
                        for idx, seqChar in g_HotstringCharSequence {
                            if (seqChar = macroEntry.char) {
                                charIndexInSequence := idx
                                break
                            }
                        }
                        if (charIndexInSequence > 0) {
                            ; Check if this character is already assigned
                            if (!g_MacroCharMap.Has(macroEntry.char)) {
                                g_MacroCharMap[macroEntry.char] := macroEntry.func
                            }
                        }
                    }
                }
                ; Second pass: assign remaining macros sequentially, skipping already assigned characters
                for macroEntry in g_Macros {
                    ; Skip if this macro already has a character assigned
                    alreadyAssigned := false
                    for assignedChar, assignedFunc in g_MacroCharMap {
                        if (assignedFunc = macroEntry.func) {
                            alreadyAssigned := true
                            break
                        }
                    }
                    if (alreadyAssigned) {
                        continue
                    }

                    ; Find next available character (skip reserved empty char if set)
                    while (charIndex <= g_HotstringCharSequence.Length) {
                        char := g_HotstringCharSequence[charIndex]
                        if (g_ReservedEmptyChar != "" && char = g_ReservedEmptyChar) {
                            charIndex++
                            continue
                        }
                        ; Check if this character is already assigned to a macro
                        if (!g_MacroCharMap.Has(char)) {
                            g_MacroCharMap[char] := macroEntry.func
                            charIndex++
                            break
                        }
                        charIndex++
                    }
                }
            }
        } else {
            ; Handle hotstring categories
            if (categorized.Has(category)) {
                ; Utility Shortcuts: assign within-category (independent) to avoid cross-category collisions.
                utilCharIndex := 1
                utilTaken := Map()

                ; Explicit assignments first
                for hs in categorized[category] {
                    if (hs.HasProp("char") && hs.char != "" && (g_ReservedEmptyChar = "" || hs.char !=
                        g_ReservedEmptyChar)) {
                        if (hs.expansion != "" && !utilTaken.Has(hs.char)) {
                            g_UtilityHotstringCharMapByCategory[category][hs.char] := hs.expansion
                            utilTaken[hs.char] := true
                        }
                    }
                }

                ; Sequential assignments for remaining hotstrings in this category
                for hs in categorized[category] {
                    alreadyAssigned := false
                    for assignedChar, assignedExpansion in g_UtilityHotstringCharMapByCategory[category] {
                        if (assignedExpansion = hs.expansion) {
                            alreadyAssigned := true
                            break
                        }
                    }
                    if (alreadyAssigned)
                        continue

                    while (utilCharIndex <= g_HotstringCharSequence.Length) {
                        ch := g_HotstringCharSequence[utilCharIndex]
                        utilCharIndex++
                        if (g_ReservedEmptyChar != "" && ch = g_ReservedEmptyChar)
                            continue
                        if (!utilTaken.Has(ch)) {
                            if (hs.expansion != "") {
                                g_UtilityHotstringCharMapByCategory[category][ch] := hs.expansion
                                utilTaken[ch] := true
                            }
                            break
                        }
                    }
                }

                ; First pass: assign hotstrings with explicit character assignments
                for hs in categorized[category] {
                    if (hs.HasProp("char") && hs.char != "" && (g_ReservedEmptyChar = "" || hs.char !=
                        g_ReservedEmptyChar)) {
                        ; Check if character is in the sequence and not already assigned
                        charIndexInSequence := 0
                        for idx, seqChar in g_HotstringCharSequence {
                            if (seqChar = hs.char) {
                                charIndexInSequence := idx
                                break
                            }
                        }
                        if (charIndexInSequence > 0) {
                            ; Check if this character is already assigned
                            if (!charMap.Has(hs.char) && hs.expansion != "") {
                                charMap[hs.char] := hs.expansion
                            }
                        }
                    }
                }
                ; Second pass: assign remaining hotstrings sequentially, skipping already assigned characters
                for hs in categorized[category] {
                    ; Skip if this hotstring already has a character assigned
                    alreadyAssigned := false
                    for assignedChar, assignedExpansion in charMap {
                        if (assignedExpansion = hs.expansion) {
                            alreadyAssigned := true
                            break
                        }
                    }
                    if (alreadyAssigned) {
                        continue
                    }

                    ; Find next available character (skip reserved empty char if set)
                    while (charIndex <= g_HotstringCharSequence.Length) {
                        char := g_HotstringCharSequence[charIndex]
                        if (g_ReservedEmptyChar != "" && char = g_ReservedEmptyChar) {
                            charIndex++
                            continue
                        }
                        ; Check if this character is already assigned
                        if (!charMap.Has(char)) {
                            ; Only assign characters to hotstrings that have an expansion (skip empty placeholders)
                            if (hs.expansion != "") {
                                charMap[char] := hs.expansion
                            }
                            charIndex++
                            break
                        }
                        charIndex++
                    }
                }
            }
        }
    }

    return charMap
}

; =============================================================================
; GetCategorizedHotstrings()
; =============================================================================
; PURPOSE: Organizes all registered items (hotstrings, quick-open files, macros) into category-based
;          data structure for GUI display purposes.
;
; PROCESS:
;   1. Initializes empty arrays for each category: Projects, Prompts, General, Files & Links, Macros
;   2. Groups hotstrings by their category property
;   3. Adds quick-open file entries to "Files & Links" category
;   4. Adds macro entries to "Macros" category
;
; RETURNS: Map object where keys are category names and values are arrays of item objects
;          Each item object contains: trigger, expansion, title, category, and optionally char
; =============================================================================
GetCategorizedHotstrings() {
    global g_hotstrings, g_QuickOpenFiles, g_Macros
    categorized := Map()
    categorized["Projects"] := []
    categorized["Prompts"] := []
    categorized["General"] := []
    categorized["Hotstrings"] := []
    categorized["Links"] := []
    categorized["Files & Links"] := []
    categorized["Macros"] := []

    ; Add hotstrings
    if (IsSet(g_hotstrings) && g_hotstrings.Length > 0) {
        for hs in g_hotstrings {
            category := hs.category
            if (category = "Projects" || category = "Prompts" || category = "General" || category = "Hotstrings") {
                categorized[category].Push(hs)
            } else {
                categorized["General"].Push(hs)
            }
        }
    }

    ; Add quick open files (rendered under Links in Utility selector)
    if (IsSet(g_QuickOpenFiles) && g_QuickOpenFiles.Length > 0) {
        for fileEntry in g_QuickOpenFiles {
            categorized["Files & Links"].Push(fileEntry)
        }
    }

    ; Add macros
    if (IsSet(g_Macros) && g_Macros.Length > 0) {
        for macroEntry in g_Macros {
            categorized["Macros"].Push(macroEntry)
        }
    }

    return categorized
}

; Get preview text (truncate long text for display, replace newlines with spaces)
GetPreviewText(text, maxLength := 60) {
    ; Replace newlines and multiple spaces with single space for cleaner preview
    preview := RegExReplace(text, "`r?`n", " ")
    preview := RegExReplace(preview, "\s+", " ")
    preview := Trim(preview)

    if (StrLen(preview) <= maxLength) {
        return preview
    }
    return SubStr(preview, 1, maxLength) . "..."
}

; Find and activate Power BI file, or open it if not already open
FindAndActivatePowerBIFile(filePath) {
    ; Check if file exists
    if (!FileExist(filePath)) {
        return false
    }

    ; Extract filename from path (Power BI window titles don't include .pbix extension)
    SplitPath(filePath, , , , &fileNameNoExt)

    ; Normalize the filename for comparison (trim whitespace, case-insensitive)
    fileNameNoExt := Trim(fileNameNoExt)
    fileNameLower := StrLower(fileNameNoExt)

    ; Search for Power BI Desktop windows with matching filename
    ; Power BI window titles are like "Dissertation InfoVis  - PowerBI - Charts" (no extension)
    try {
        for hwnd in WinGetList("ahk_exe PBIDesktop.exe") {
            try {
                winTitle := WinGetTitle("ahk_id " hwnd)
                winTitleLower := StrLower(Trim(winTitle))

                ; Check if window title matches the filename (case-insensitive)
                ; Power BI window title should be exactly the filename or start with it
                ; Check for exact match first, then check if title starts with filename
                if (winTitleLower = fileNameLower || InStr(winTitleLower, fileNameLower) = 1) {
                    ; Found matching window, activate it
                    WinActivate("ahk_id " hwnd)
                    WinWaitActive("ahk_id " hwnd, , 2)
                    return true
                }
            } catch {
                ; Skip windows we can't access
                continue
            }
        }
    } catch {
        ; No Power BI windows found or error accessing them
    }

    ; No matching window found, open the file
    try {
        Run(filePath)
        return true
    } catch Error as e {
        ; Failed to open file
        return false
    }
}

; Find and activate Miro window by title keywords and URL
; Returns true if window was found and activated, or if URL was opened successfully
FindAndActivateMiroWindow(url, titleKeywords) {
    ; Normalize title keywords for matching (case-insensitive)
    keywordsLower := StrLower(titleKeywords)

    ; Search for Chrome windows with Miro in title
    try {
        for hwnd in WinGetList("ahk_exe chrome.exe") {
            try {
                winTitle := WinGetTitle("ahk_id " hwnd)
                winTitleLower := StrLower(Trim(winTitle))

                ; Check if window is a Miro window and contains the keywords
                if (InStr(winTitleLower, "miro") && InStr(winTitleLower, keywordsLower)) {
                    ; Found matching window, activate it and bring to front
                    ; Use a separate try-catch for activation to ensure we return even if activation fails
                    try {
                        ; Ensure window is not minimized first
                        if (WinGetMinMax("ahk_id " hwnd) = -1) {
                            WinRestore("ahk_id " hwnd)
                        }

                        ; Activate the window
                        WinActivate("ahk_id " hwnd)
                        WinWaitActive("ahk_id " hwnd, , 2)

                        ; Bring to front using AlwaysOnTop trick to ensure it's not hidden
                        WinSetAlwaysOnTop("On", "ahk_id " hwnd)
                        Sleep 50
                        WinSetAlwaysOnTop("Off", "ahk_id " hwnd)
                    } catch Error as activateErr {
                        ; Even if activation fails, we found the window, so return to prevent opening a new one
                    }

                    ; Return immediately after activating - don't continue searching or open new window
                    return true
                }
            } catch Error as e {
                ; Skip windows we can't access
                continue
            }
        }
    } catch {
        ; No Chrome windows found or error accessing them
    }

    ; No matching window found, open the URL
    try {
        ; Open URL in Chrome
        Run("chrome.exe --new-window " . url)

        ; Wait for window to appear and become active
        ; Wait up to 10 seconds for the window to appear
        WinWait("ahk_exe chrome.exe", , 10)

        ; Find the newly opened window by checking for Miro in title
        ; Give it a moment to load
        Sleep(1000)

        ; Try to find the window with Miro in title
        loop 10 {
            for hwnd in WinGetList("ahk_exe chrome.exe") {
                try {
                    winTitle := WinGetTitle("ahk_id " hwnd)
                    winTitleLower := StrLower(Trim(winTitle))

                    if (InStr(winTitleLower, "miro") && InStr(winTitleLower, keywordsLower)) {
                        ; Found the window, activate it
                        ; Ensure window is not minimized first
                        if (WinGetMinMax("ahk_id " hwnd) = -1) {
                            WinRestore("ahk_id " hwnd)
                        }
                        ; Activate the window
                        WinActivate("ahk_id " hwnd)
                        WinWaitActive("ahk_id " hwnd, , 2)
                        ; Bring to front using AlwaysOnTop trick to ensure it's not hidden
                        WinSetAlwaysOnTop("On", "ahk_id " hwnd)
                        Sleep 50
                        WinSetAlwaysOnTop("Off", "ahk_id " hwnd)
                        ; Additional check: ensure window is actually active
                        if (WinActive("ahk_id " hwnd)) {
                            return true
                        }
                    }
                } catch {
                    continue
                }
            }
            Sleep(500)  ; Wait before next attempt
        }

        ; If we couldn't find by title, just activate the most recent Chrome window
        ; This is a fallback in case the title hasn't updated yet
        try {
            chromeWindows := WinGetList("ahk_exe chrome.exe")
            if (chromeWindows.Length > 0) {
                ; Get the first (most recent) Chrome window
                hwnd := chromeWindows[1]

                ; Ensure window is not minimized
                if (WinGetMinMax("ahk_id " hwnd) = -1) {
                    WinRestore("ahk_id " hwnd)
                }
                ; Activate the window
                WinActivate("ahk_id " hwnd)
                WinWaitActive("ahk_id " hwnd, , 2)
                ; Bring to front
                WinSetAlwaysOnTop("On", "ahk_id " hwnd)
                Sleep 50
                WinSetAlwaysOnTop("Off", "ahk_id " hwnd)
                return true
            }
        } catch {
        }

        return true  ; Assume success if we got this far
    } catch Error as e {
        ; Failed to open URL
        return false
    }
}

; =============================================================================
; CleanupHotstringSelector()
; =============================================================================
; PURPOSE: Closes hotstring selector GUI and disables all associated hotkeys.
;          Called when selector is closed via Escape key, character selection, or toggle.
;
; PROCESS:
;   1. Sets g_HotstringSelectorActive = false to prevent further character processing
;   2. Disables all character hotkeys (including uppercase variants for lowercase letters)
;   3. Handles special VK codes for comma (vkBC) and period (vkBE)
;   4. Disables Escape key handler
;   5. Destroys GUI object if it exists
;   6. Clears hotkey handlers array
;   7. Clears character mapping maps
;
; RETURNS: None (void function)
; SIDE EFFECTS: Resets all global state variables to initial values
; =============================================================================
CleanupHotstringSelector() {
    global g_HotstringSelectorActive, g_HotstringSelectorGui, g_HotstringHotkeyHandlers
    global g_HotstringPromptCharMap, g_HotstringGeminiArmed
    global g_HotstringCharMap
    global g_HS_SelectorOpenFile, g_HS_SelectorCloseRequestFile, g_HS_SelectorCloseCheckTimer
    global g_UtilitySelectorMode, g_UtilitySelectorCategory

    ; Disable active flag
    g_HotstringSelectorActive := false

    ; Disable all character hotkeys
    for handler in g_HotstringHotkeyHandlers {
        try {
            key := handler.HasProp("key") ? handler.key : handler.char
            char := handler.HasProp("char") ? handler.char : key
            ; Handle special VK codes
            if (key = "vkBC" || char = ",") {
                Hotkey("vkBC", "Off")
            } else if (key = "vkBE" || char = ".") {
                Hotkey("vkBE", "Off")
            } else {
                Hotkey(key, "Off")
            }
        } catch {
            ; Silently ignore errors
        }
    }

    ; Disable Escape hotkey
    try {
        Hotkey("Escape", "Off")
    } catch {
        ; Ignore
    }

    ; Disable Backspace hotkey (menu back)
    try {
        Hotkey("Backspace", "Off")
    } catch {
        ; Ignore
    }

    ; Stop IPC timer and clear sentinel files
    try SetTimer(Utils_CheckHotstringSelectorCloseRequest, 0)
    g_HS_SelectorCloseCheckTimer := ""
    try FileDelete(g_HS_SelectorOpenFile)
    catch {
    }
    try FileDelete(g_HS_SelectorCloseRequestFile)
    catch {
    }

    ; Clear handlers array
    g_HotstringHotkeyHandlers := []
    g_HotstringPromptCharMap := Map()
    g_HotstringGeminiArmed := false
    g_UtilitySelectorMode := "top"
    g_UtilitySelectorCategory := ""

    ; Close and destroy GUI
    if (IsObject(g_HotstringSelectorGui)) {
        try {
            g_HotstringSelectorGui.Destroy()
        } catch {
            ; Ignore
        }
        g_HotstringSelectorGui := false
    }

    ; Clear char map
    g_HotstringCharMap := Map()
}

; =============================================================================
; HandleHotstringChar(char)
; =============================================================================
; PURPOSE: Processes character key press events when hotstring selector is active.
;          Executes the action (text expansion, file open, or macro) associated with the character.
;
; PARAMETERS:
;   char: String - Single character that was pressed (e.g., "a", "1", ",")
;
; EXECUTION ORDER:
;   1. Special cases: Characters "9" and "0" trigger Miro board activation (hardcoded)
;   2. Quick-open files: Check g_QuickOpenFileCharMap for file path mappings
;   3. Macros: Check g_MacroCharMap for executable macro functions
;   4. Hotstrings: Check g_HotstringCharMap for text expansion mappings
;
; BEHAVIOR:
;   - Performs case-insensitive lookup (tries both original and lowercase)
;   - Closes selector GUI before executing action
;   - For hotstrings: Inserts text using InsertText() after 150ms delay
;   - For files: Opens file based on extension (Power BI files use special handler)
;   - For macros: Executes macro function directly
;
; RETURNS: None (void function)
; =============================================================================

; Navigate to Gemini, focus the prompt field, then paste first clipboard snippet (same as Win+Alt+Shift+1).
; Reference: "order called snippets" – Clip Angel top item sent via !v then ^!b.
; If optionalPromptText is non-empty, inserts that text into the prompt field instead (same as Win+Alt+Shift+U then L, prompt char).
; switchToFirstTab: when true (default), send Ctrl+1 and show tab-1 banner (AI Text Optimizer / ^!#4). When false, use currently active Gemini tab if any, else first Gemini window, without changing tab (delay-submit flow).
GeminiNavigateFocusAndPasteFirstSnippet(optionalPromptText := "", switchToFirstTab := true) {
    SetTitleMatchMode(2)
    geminiHwnd := 0
    if (!switchToFirstTab) {
        ; Prefer the currently active window if it is already a Gemini tab (do not switch tabs).
        try {
            activeHwnd := WinExist("A")
            if (activeHwnd && WinGetProcessName("ahk_id " activeHwnd) = "chrome.exe" && InStr(WinGetTitle("ahk_id " activeHwnd
            ), "gemini", false))
                geminiHwnd := activeHwnd
        } catch {
        }
    }
    if (!geminiHwnd) {
        try {
            for hwnd in WinGetList("ahk_exe chrome.exe") {
                try {
                    if InStr(WinGetTitle("ahk_id " hwnd), "gemini", false) {
                        geminiHwnd := hwnd
                        break
                    }
                } catch {
                }
            }
        } catch {
        }
    }

    if (!geminiHwnd) {
        ; Navigate to Gemini website
        Run "chrome.exe --new-window https://gemini.google.com/"
        if !WinWaitActive("ahk_exe chrome.exe", , 5)
            return
        Sleep 2500  ; Allow page to load
        geminiHwnd := WinExist("A")
    }

    if (geminiHwnd) {
        WinActivate("ahk_id " geminiHwnd)
        WinWaitActive("ahk_id " geminiHwnd, , 2)
        Sleep 350  ; Let UI finish processing the window transition before focus/paste
    } else {
        WinActivate("ahk_exe chrome.exe")
        WinWaitActive("ahk_exe chrome.exe", , 2)
        Sleep 350
    }

    if (switchToFirstTab) {
        ; Switch to first Gemini tab (Ctrl+1) and show number-one banner (consistent with Gemini.ahk)
        Send("^1")
        Sleep 280
        ShowSingleCharTabBanner_Utils(1)
    }

    ; Focus the Gemini prompt field (Anchor & Backtrack strategy)
    try {
        uia := UIA_Browser()
        Sleep 80
        anchorButton := 0
        try {
            anchorButton := uia.FindFirst({ Type: "50000", Name: "Open upload file menu", ControlType: "Button" })
            if (!anchorButton)
                anchorButton := uia.FindFirst({ Type: "50000", Name: "Open upload file menu", cs: false })
        } catch {
        }
        if (!anchorButton) {
            try {
                allButtons := uia.FindAll({ Type: "50000" })
                for button in allButtons {
                    try {
                        if (InStr(button.Name, "Open upload file menu", false)) {
                            anchorButton := button
                            break
                        }
                    } catch {
                        continue
                    }
                }
            } catch {
            }
        }
        if (anchorButton) {
            try {
                anchorButton.SetFocus()
                Sleep 25
                SendInput "+{Tab}"
                Sleep 15
            } catch {
                try {
                    promptField := FindGeminiPromptField(uia)
                    if (promptField)
                        promptField.SetFocus()
                } catch as e {
                }
            }
        } else {
            try {
                promptField := FindGeminiPromptField(uia)
                if (promptField)
                    promptField.SetFocus()
            } catch as e {
            }
        }
    } catch {
    }

    ; Explicitly target Gemini window again before paste so paste goes to Gemini, not the trigger window
    if (geminiHwnd && WinExist("ahk_id " geminiHwnd)) {
        WinActivate("ahk_id " geminiHwnd)
        WinWaitActive("ahk_id " geminiHwnd, , 2)
        Sleep 150
    }

    if (optionalPromptText != "") {
        InsertText(optionalPromptText)
    } else {
        ; Paste first clipboard snippet (same as Win+Alt+Shift+1: order called snippets)
        Send "!v"
        Sleep 50
        Send "^!b"
    }
    ; Brief delay so paste is received and UI/character limits register before any submit or focus change
    Sleep 250
    ; Same sound as when opening Gemini (focus/paste feedback)
    if (IsSoundEnabled())
        SoundPlay(A_ScriptDir . "\sounds\gemini-focused.wav")
}

; Returns grammar preset (from prompt/grammar.txt or fallback). Matches InitHotstringsCheatSheet catch for :o:cgrammar.
GetGrammarPromptText() {
    promptDir := A_ScriptDir "\prompt"
    try {
        return FileRead(promptDir "\grammar.txt")
    } catch {
        return "Correct grammar, spelling, punctuation, and casing. Give back only the text.`n"
    }
}

; Returns the AI Text Optimizer prompt text (from prompt/aiopt.txt or fallback). Used by Ctrl+Alt+Win+4 and L+4 flow.
GetAioptPromptText() {
    promptDir := A_ScriptDir "\prompt"
    try {
        return FileRead(promptDir "\aiopt.txt")
    } catch {
        return "Rewrite the input text so it becomes AI-oriented. Preserve all important information.`n"
    }
}

; Delayed submit flow: show 4s banner, allow N to cancel auto-submit; then navigate+paste and optionally send Enter.
; Tell Gemini.ahk to start background completion monitor (must match WM_START_DELAYED_SUBMIT_MONITOR in Gemini.ahk).
GeminiDelayedSubmitMonitorStartFromUtils(originalHwnd, geminiChromeHwnd) {
    WM_START_DELAYED_SUBMIT_MONITOR := 0x8002
    targetHwnd := GetGeminiScriptMsgTargetHwnd()
    if (targetHwnd) {
        try SendMessage(WM_START_DELAYED_SUBMIT_MONITOR, originalHwnd, geminiChromeHwnd, , "ahk_id " targetHwnd)
    }
}

; Tell Gemini.ahk to stop the delayed-submit monitor so "Copy response?" is not shown (when user chose S or N at 6s).
GeminiDelayedSubmitMonitorStopFromUtils() {
    WM_STOP_DELAYED_SUBMIT_MONITOR := 0x8003
    targetHwnd := GetGeminiScriptMsgTargetHwnd()
    if (targetHwnd) {
        try SendMessage(WM_STOP_DELAYED_SUBMIT_MONITOR, 0, 0, , "ahk_id " targetHwnd)
    }
}

; Paste transcription to Gemini prompt only (no Enter, no 4s banner). Used when user presses S at 6s dictation confirm.
DEPRECATED_GeminiDictationPasteOnlyFlow() {
    restoreHwnd := WinExist("A")
    GeminiNavigateFocusAndPasteFirstSnippet("", false)
    if (restoreHwnd && WinExist("ahk_id " restoreHwnd))
        WinActivate("ahk_id " restoreHwnd)
}

DEPRECATED_GeminiDelayedSubmitFlow() {
    global g_HotstringGeminiAutoSubmit, g_HotstringGeminiRestoreHwnd
    ; #region agent log
    DebugFlowLog("Utils.ahk:GeminiDelayedSubmitFlow", "entry", "", "H3")
    ; #endregion
    g_HotstringGeminiRestoreHwnd := WinExist("A")
    g_HotstringGeminiAutoSubmit := true
    GeminiFinalizeSubmit()
}

; Delay (ms) after paste and before Send Enter in Gemini delayed-submit flow. Prevents premature send and lets the UI register paste + character limits.
global g_GeminiDelayedSubmit_PreEnterDelayMs := 1000
; Max ms to wait for prompt field to show content before sending Enter (guarantee layer). Poll interval = 200 ms.
global g_GeminiDelayedSubmit_WaitContentMaxMs := 5000

; Returns non-empty trimmed text if Gemini prompt field has content (Value or TextPattern); "" on failure or empty. Used to guarantee message is present before submit.
GeminiPromptFieldGetText() {
    try {
        uia := UIA_Browser()
        pf := FindGeminiPromptField(uia)
        if (!pf)
            return ""
        try {
            text := Trim(pf.Value)
            if (text != "")
                return text
        } catch {
        }
        try {
            text := Trim(pf.TextPattern.DocumentRange.GetText(-1))
            if (text != "")
                return text
        } catch {
        }
    } catch {
    }
    return ""
}

GeminiFinalizeSubmit() {
    global g_HotstringGeminiAutoSubmit, g_HotstringGeminiRestoreHwnd, g_GeminiDelayedSubmit_PreEnterDelayMs,
        g_GeminiDelayedSubmit_WaitContentMaxMs

    ; #region agent log
    DebugFlowLog("Utils.ahk:GeminiFinalizeSubmit", "entry", "autoSubmit=" . (g_HotstringGeminiAutoSubmit ? 1 : 0), "H5"
    )
    ; #endregion
    try Hotkey("n", "Off")
    try Hotkey("N", "Off")
    try Hotkey("y", "Off")
    try Hotkey("Y", "Off")
    HotstringGeminiBanner_Hide()

    ; Delay-submit flow: do not switch tabs; paste to currently active Gemini tab
    GeminiNavigateFocusAndPasteFirstSnippet("", false)

    didAutoSubmit := false
    geminiChromeHwnd := 0
    if (g_HotstringGeminiAutoSubmit) {
        ; Execution delay so paste is fully received and UI/character limits register before submit
        Sleep (g_GeminiDelayedSubmit_PreEnterDelayMs)
        ; Guarantee layer: wait until prompt field has content (or timeout) so we don't send Enter prematurely
        pollIntervalMs := 200
        endTick := A_TickCount + g_GeminiDelayedSubmit_WaitContentMaxMs
        contentFound := false
        while (A_TickCount < endTick) {
            if (GeminiPromptFieldGetText() != "") {
                contentFound := true
                break
            }
            Sleep pollIntervalMs
        }
        Send("{Enter}")
        geminiChromeHwnd := WinExist("A")
        didAutoSubmit := true
    }

    g_HotstringGeminiAutoSubmit := true

    ; Return focus to the window the user had before paste (whether Enter was sent or not)
    if (g_HotstringGeminiRestoreHwnd && WinExist("ahk_id " g_HotstringGeminiRestoreHwnd)) {
        WinActivate("ahk_id " g_HotstringGeminiRestoreHwnd)
    }

    ; If we auto-submitted (user did not cancel), ask Gemini.ahk to monitor for completion and show "Copy? [N] [R]" when done
    if (didAutoSubmit && geminiChromeHwnd)
        GeminiDelayedSubmitMonitorStartFromUtils(g_HotstringGeminiRestoreHwnd, geminiChromeHwnd)
}

HandleHotstringChar(char) {
    global g_HotstringSelectorActive, g_HotstringCharMap, g_QuickOpenFileCharMap, g_MacroCharMap
    global g_HotstringPromptCharMap, g_HotstringGeminiArmed
    global g_UtilitySelectorMode, g_UtilityTopCategoryById

    ; Only process if selector is active
    if (!g_HotstringSelectorActive) {
        return
    }

    ; Top-level category selection (1-6)
    if (g_UtilitySelectorMode = "top") {
        ch := StrLower(char)
        if (g_UtilityTopCategoryById.Has(ch)) {
            UtilitySelector_SwitchToCategory(g_UtilityTopCategoryById[ch])
        }
        return
    }

    ; L key: first press = arm Gemini mode (show banner); second press (double-tap) = navigate to Gemini, focus field, paste first snippet.
    if (char = "l" || char = "L") {
        if (g_HotstringGeminiArmed) {
            ; Double-tap L: delayed submit flow (paste + Enter to Gemini).
            CleanupHotstringSelector()
            D2C_FlowManager.GetInstance().StartFromHotstring()
            g_HotstringGeminiArmed := false
            return
        }
        g_HotstringGeminiArmed := true
        ; Show banner when entering Gemini mode (same pattern as Project Selector "Entering Selection Mode").
        HotstringGeminiBanner_Show("⌨ Entering Gemini Mode - Select prompt")
        SetTimer(HotstringGeminiBanner_Hide, -1500)  ; Hide banner after 1.5 s
        SetTimer(DisarmHotstringGeminiMode, -4000)
        return
    }

    ; Consume the armed state on the next selection (any selection), but only redirect Prompts.
    useGemini := false
    if (g_HotstringGeminiArmed) {
        useGemini := g_HotstringPromptCharMap.Has(StrLower(char)) || g_HotstringPromptCharMap.Has(char)
        g_HotstringGeminiArmed := false
    }

    ; Special handling for Miro boards (characters "9" and "0")
    ; Gate to Links category so other views can't trigger it.
    global g_UtilitySelectorCategory
    if (g_UtilitySelectorCategory = "Links") {
        if (char = "9") {
            CleanupHotstringSelector()
            FindAndActivateMiroWindow("https://miro.com/app/board/uXjVJdbNFkA=/", "CIP & UX Integration")
            return
        } else if (char = "0") {
            CleanupHotstringSelector()
            FindAndActivateMiroWindow("https://miro.com/app/board/uXjVJVZSXvk=/", "CIP Dashboard")
            return
        }
    }

    ; Category-scoped dispatch (prevents cross-menu execution when chars overlap)
    global g_UtilityHotstringCharMapByCategory, g_UtilitySelectorCategory

    ResolveExpansion() {
        exp := ""
        try {
            if (IsObject(g_UtilityHotstringCharMapByCategory) && g_UtilityHotstringCharMapByCategory.Has(
                g_UtilitySelectorCategory)) {
                exp := g_UtilityHotstringCharMapByCategory[g_UtilitySelectorCategory].Get(char, "")
                if (exp = "")
                    exp := g_UtilityHotstringCharMapByCategory[g_UtilitySelectorCategory].Get(StrLower(char), "")
            }
        } catch {
            exp := ""
        }
        if (exp = "") {
            exp := g_HotstringCharMap.Get(char, "")
            if (exp = "")
                exp := g_HotstringCharMap.Get(StrLower(char), "")
        }
        return exp
    }

    ResolveFilePath() {
        fp := g_QuickOpenFileCharMap.Get(char, "")
        if (fp = "")
            fp := g_QuickOpenFileCharMap.Get(StrLower(char), "")
        return fp
    }

    ResolveMacro() {
        fn := g_MacroCharMap.Get(char, "")
        if (fn = "")
            fn := g_MacroCharMap.Get(StrLower(char), "")
        return fn
    }

    TryRunFile(fp) {
        if (fp = "")
            return false
        CleanupHotstringSelector()
        SplitPath(fp, , , &ext)
        ext := StrLower(ext)
        if (ext = "pbix") {
            FindAndActivatePowerBIFile(fp)
        } else {
            try Run(fp)
            catch {
            }
        }
        return true
    }

    TryRunMacro(fn) {
        if (fn = "")
            return false
        CleanupHotstringSelector()
        try fn()
        catch {
        }
        return true
    }

    expansion := ""
    filePath := ""
    macroFunc := ""

    if (g_UtilitySelectorCategory = "Links") {
        filePath := ResolveFilePath()
        if (TryRunFile(filePath))
            return
        ; fallback for unexpected collisions
        expansion := ResolveExpansion()
    } else if (g_UtilitySelectorCategory = "Macros") {
        macroFunc := ResolveMacro()
        if (TryRunMacro(macroFunc))
            return
        expansion := ResolveExpansion()
    } else {
        ; Projects / Prompts / Hotstrings / General: hotstring-first
        expansion := ResolveExpansion()
        if (expansion = "") {
            macroFunc := ResolveMacro()
            if (TryRunMacro(macroFunc))
                return
            filePath := ResolveFilePath()
            if (TryRunFile(filePath))
                return
        }
    }

    if (expansion != "") {
        ; Cleanup first (closes GUI, disables hotkeys)
        CleanupHotstringSelector()

        if (useGemini) {
            ; L+Prompt selection: redirect to Gemini (focus prompt field, paste, do NOT submit).
            HotstringGeminiBanner_Show("📤 Gemini: inserting prompt...")
            try {
                SetTitleMatchMode(2)
                geminiHwnd := 0
                try {
                    for hwnd in WinGetList("ahk_exe chrome.exe") {
                        try {
                            if InStr(WinGetTitle("ahk_id " hwnd), "gemini", false) {
                                geminiHwnd := hwnd
                                break
                            }
                        } catch {
                            ; Skip invalid windows
                        }
                    }
                } catch {
                    ; Ignore WinGetList errors
                }

                if (geminiHwnd) {
                    WinActivate("ahk_id " geminiHwnd)
                    WinWaitActive("ahk_id " geminiHwnd, , 2)
                } else {
                    ; Per your preference: fallback to any Chrome window if Gemini isn't found.
                    WinActivate("ahk_exe chrome.exe")
                    WinWaitActive("ahk_exe chrome.exe", , 2)
                }

                ; Explicitly target Gemini tabs based on character:
                ; - L+1/2/3  -> Tab 2 (right Gemini tab, temporary prompts)
                ; - L+4/5/Q/W/E/R/T/A -> Tab 1 (left Gemini tab, main workflow)
                if (char = "1" || char = "2" || char = "3") {
                    ; Chrome convention: Ctrl+2 selects the second tab in the window.
                    Send("^2")
                    Sleep 120
                    ShowSingleCharTabBanner_Utils(2)
                } else if (char = "4" || char = "5"
                    || char = "q" || char = "Q"
                    || char = "w" || char = "W"
                    || char = "e" || char = "E"
                    || char = "r" || char = "R"
                    || char = "t" || char = "T"
                    || char = "a" || char = "A") {
                    ; Ensure Tab 1 is active before inserting the prompt.
                    Send("^1")
                    Sleep 120
                    ShowSingleCharTabBanner_Utils(1)
                }

                ; Focus the Gemini prompt field using the Anchor & Backtrack strategy (copied from Win+Alt+Shift+I),
                ; with sound/file behaviors removed (no external file refs, no side effects).
                try {
                    uia := UIA_Browser()
                    Sleep 80

                    anchorButton := 0
                    try {
                        anchorButton := uia.FindFirst({ Type: "50000", Name: "Open upload file menu", ControlType: "Button" })
                        if (!anchorButton) {
                            anchorButton := uia.FindFirst({ Type: "50000", Name: "Open upload file menu", cs: false })
                        }
                    } catch {
                    }

                    if (!anchorButton) {
                        try {
                            allButtons := uia.FindAll({ Type: "50000" })
                            for button in allButtons {
                                try {
                                    if (InStr(button.Name, "Open upload file menu", false)) {
                                        anchorButton := button
                                        break
                                    }
                                } catch {
                                    continue
                                }
                            }
                        } catch {
                        }
                    }

                    if (anchorButton) {
                        try {
                            anchorButton.SetFocus()
                            Sleep 25
                            SendInput "+{Tab}"
                            Sleep 15
                        } catch {
                            try {
                                promptField := FindGeminiPromptField(uia)
                                if (promptField) {
                                    promptField.SetFocus()
                                }
                            } catch as e {
                            }
                        }
                    } else {
                        try {
                            promptField := FindGeminiPromptField(uia)
                            if (promptField) {
                                promptField.SetFocus()
                            }
                        } catch as e {
                        }
                    }
                } catch {
                    ; If focus fails, we still attempt to paste (user can click manually).
                }

                ; Paste the text (do NOT send Enter)
                InsertText(expansion)
                ; Same sound as when opening Gemini (focus/paste feedback)
                if (IsSoundEnabled())
                    SoundPlay(A_ScriptDir . "\sounds\gemini-focused.wav")
            } finally {
                HotstringGeminiBanner_Hide()
            }
            return
        }

        ; Standard behavior: paste into current active text field.
        ; Small delay to ensure target window has focus before pasting
        Sleep 150
        InsertText(expansion)
    }
}

; =============================================================================
; CreateHotstringCharHandler(char)
; =============================================================================
; PURPOSE: Factory function that creates a hotkey handler function with proper closure over character value.
;          Required because AutoHotkey hotkey handlers need unique function instances per character.
;
; PARAMETERS:
;   char: String - Character to create handler for
;
; RETURNS: Function object that calls HandleHotstringChar(char) when invoked
; =============================================================================
DisarmHotstringGeminiMode(*) {
    global g_HotstringGeminiArmed
    g_HotstringGeminiArmed := false
}

CreateHotstringCharHandler(char) {
    ; Return a function that captures the char value at creation time via closure
    return (*) => HandleHotstringChar(char)
}

; =============================================================================
; HandleHotstringEscape(*)
; =============================================================================
; PURPOSE: Handles Escape key press to close hotstring selector without executing any action.
;
; PARAMETERS: None (varargs function signature for hotkey compatibility)
; RETURNS: None (void function)
; =============================================================================
HandleHotstringEscape(*) {
    global g_HotstringSelectorActive
    if (g_HotstringSelectorActive) {
        CleanupHotstringSelector()
    }
}

HandleUtilitySelectorBack(*) {
    global g_HotstringSelectorActive, g_UtilitySelectorMode
    if (!g_HotstringSelectorActive)
        return
    if (g_UtilitySelectorMode = "category") {
        UtilitySelector_SwitchToTop()
    }
}

UtilitySelector_SwitchToTop() {
    global g_UtilitySelectorMode, g_UtilitySelectorCategory
    g_UtilitySelectorMode := "top"
    g_UtilitySelectorCategory := ""
    try UtilitySelector_RefreshUiAndHotkeys()
    catch {
    }
}

UtilitySelector_SwitchToCategory(category) {
    global g_UtilitySelectorMode, g_UtilitySelectorCategory
    g_UtilitySelectorMode := "category"
    g_UtilitySelectorCategory := category
    try UtilitySelector_RefreshUiAndHotkeys()
    catch {
    }
}

UtilitySelector_MapInternalCategoryToTop(internalCategory) {
    if (internalCategory = "Files & Links")
        return "Links"
    if (internalCategory = "General")
        return "General"
    if (internalCategory = "Links")
        return "Links"
    if (internalCategory = "Hotstrings")
        return "Hotstrings"
    ; Unknown/legacy categories are folded into General now that top-level "Hot Strings" is removed.
    if (internalCategory != "Prompts" && internalCategory != "Projects" && internalCategory != "Macros")
        return "General"
    return internalCategory
}

UtilitySelector_GetAllowedCharsForCurrentView() {
    global g_UtilitySelectorMode, g_UtilitySelectorCategory
    global g_UtilityTopCategoryById, g_UtilitySelectorAllItems
    allowed := Map()

    if (g_UtilitySelectorMode = "top") {
        for id, category in g_UtilityTopCategoryById {
            allowed[id] := true
        }
        return allowed
    }

    ; Category view: enable only actionable items in the selected category.
    ; (Empty placeholders are displayed but not bound.)
    for item in g_UtilitySelectorAllItems {
        if (item.category = g_UtilitySelectorCategory && !item.isEmpty) {
            allowed[item.char] := true
        }
    }

    ; Prompts view: enable Gemini modifier key 'L' workflow
    if (g_UtilitySelectorCategory = "Prompts") {
        allowed["l"] := true
        allowed["L"] := true
    }

    return allowed
}

UtilitySelector_RebindHotkeys() {
    global g_HotstringHotkeyHandlers, g_UtilitySelectorMode
    allowed := UtilitySelector_GetAllowedCharsForCurrentView()

    ; Disable previously-bound hotkeys
    for handler in g_HotstringHotkeyHandlers {
        try {
            key := handler.key
            if (key = "vkBC") {
                Hotkey("vkBC", "Off")
            } else if (key = "vkBE") {
                Hotkey("vkBE", "Off")
            } else {
                Hotkey(key, "Off")
            }
        } catch {
        }
    }
    g_HotstringHotkeyHandlers := []

    ; Bind allowed character hotkeys
    for char, _ in allowed {
        handler := CreateHotstringCharHandler(char)
        try {
            if (char = ",") {
                Hotkey("vkBC", handler, "On")
                g_HotstringHotkeyHandlers.Push({ key: "vkBC", char: char, handler: handler })
            } else if (char = ".") {
                Hotkey("vkBE", handler, "On")
                g_HotstringHotkeyHandlers.Push({ key: "vkBE", char: char, handler: handler })
            } else {
                Hotkey(char, handler, "On")
                g_HotstringHotkeyHandlers.Push({ key: char, char: char, handler: handler })
                if (RegExMatch(char, "^[a-z]$")) {
                    Hotkey(StrUpper(char), handler, "On")
                    g_HotstringHotkeyHandlers.Push({ key: StrUpper(char), char: char, handler: handler })
                }
            }
        } catch {
        }
    }

    ; Back navigation
    if (g_UtilitySelectorMode = "category") {
        try Hotkey("Backspace", HandleUtilitySelectorBack, "On")
    } else {
        try Hotkey("Backspace", "Off")
    }

    ; Escape always closes
    Hotkey("Escape", HandleHotstringEscape, "On")
}

UtilitySelector_BuildTopLevelText() {
    global g_UtilityTopCategories, g_UtilitySelectorAllItems
    ; Count actionable items per category
    counts := Map()
    for cat in g_UtilityTopCategories
        counts[cat] := 0
    for item in g_UtilitySelectorAllItems {
        if (!item.isEmpty && counts.Has(item.category)) {
            counts[item.category] := counts[item.category] + 1
        }
    }

    text := ""
    text .= "[R] Prompts (" . counts["Prompts"] . ")`n"
    text .= "[P] Projects (" . counts["Projects"] . ")`n"
    text .= "[M] Macros (" . counts["Macros"] . ")`n"
    text .= "[G] General (" . counts["General"] . ")`n"
    text .= "[L] Links (" . counts["Links"] . ")`n"
    text .= "[H] Hotstrings (" . counts["Hotstrings"] . ")`n"
    text .= "`nPress R/P/M/G/L/H to open a category.`n"
    return text
}

UtilitySelector_BuildTopLevelRich() {
    global g_UtilityTopCategories, g_UtilitySelectorAllItems
    counts := Map()
    for cat in g_UtilityTopCategories
        counts[cat] := 0
    for item in g_UtilitySelectorAllItems {
        if (!item.isEmpty && counts.Has(item.category)) {
            counts[item.category] := counts[item.category] + 1
        }
    }

    lines := []
    lines.Push({ text: "[R] Prompts (" . counts["Prompts"] . ")", key: "r" })
    lines.Push({ text: "[P] Projects (" . counts["Projects"] . ")", key: "p" })
    lines.Push({ text: "[M] Macros (" . counts["Macros"] . ")", key: "m" })
    lines.Push({ text: "[G] General (" . counts["General"] . ")", key: "g" })
    lines.Push({ text: "[L] Links (" . counts["Links"] . ")", key: "l" })
    lines.Push({ text: "[H] Hotstrings (" . counts["Hotstrings"] . ")", key: "h" })
    lines.Push({ text: "" })
    lines.Push({ text: "Press R/P/M/G/L/H to open a category." })
    return lines
}

UtilitySelector_BuildCategoryText(isPortrait) {
    global g_UtilitySelectorCategory, g_UtilitySelectorAllItems
    ; Filter items for this category
    items := []
    for item in g_UtilitySelectorAllItems {
        if (item.category = g_UtilitySelectorCategory)
            items.Push(item)
    }

    header := "— " . g_UtilitySelectorCategory . " —`n"
    if (items.Length = 0) {
        return header . "(no items)`n`nBackspace = back | Esc = close"
    }

    if (isPortrait) {
        text := header
        for item in items
            text .= item.text . "`n"
        text .= "`nBackspace = back | Esc = close"
        return text
    }

    ; Landscape: two columns
    maxItemLength := 0
    for item in items {
        if (StrLen(item.text) > maxItemLength)
            maxItemLength := StrLen(item.text)
    }
    columnWidth := maxItemLength + 2
    if (columnWidth < 36)
        columnWidth := 36
    columnSpacing := "  "

    midPoint := Ceil(items.Length / 2)
    maxLines := items.Length - midPoint
    if (midPoint > maxLines)
        maxLines := midPoint

    PadString(str, width) {
        len := StrLen(str)
        if (len >= width)
            return str
        padding := width - len
        spaces := ""
        loop padding
            spaces .= " "
        return str . spaces
    }

    text := header
    loop maxLines {
        leftText := ""
        rightText := ""
        if (A_Index <= midPoint)
            leftText := PadString(items[A_Index].text, columnWidth)
        else
            leftText := PadString("", columnWidth)
        rightIdx := A_Index + midPoint
        if (rightIdx <= items.Length)
            rightText := items[rightIdx].text
        text .= leftText . columnSpacing . rightText . "`n"
    }
    text .= "`nBackspace = back | Esc = close"
    return text
}

UtilitySelector_BuildCategoryRich(isPortrait) {
    global g_UtilitySelectorCategory, g_UtilitySelectorAllItems
    items := []
    for item in g_UtilitySelectorAllItems {
        if (item.category = g_UtilitySelectorCategory)
            items.Push(item)
    }

    lines := []
    lines.Push({ text: "— " . g_UtilitySelectorCategory . " —" })
    if (items.Length = 0) {
        lines.Push({ text: "(no items)" })
        lines.Push({ text: "" })
        lines.Push({ text: "Backspace = back | Esc = close" })
        return lines
    }

    if (isPortrait) {
        for item in items
            lines.Push({ text: item.text, key: item.isEmpty ? "" : item.char })
        lines.Push({ text: "" })
        lines.Push({ text: "Backspace = back | Esc = close" })
        return lines
    }

    ; Landscape: two columns (mirror text layout; emphasize both columns)
    maxItemLength := 0
    for item in items {
        if (StrLen(item.text) > maxItemLength)
            maxItemLength := StrLen(item.text)
    }
    columnWidth := maxItemLength + 2
    if (columnWidth < 36)
        columnWidth := 36
    columnSpacing := "  "

    midPoint := Ceil(items.Length / 2)
    maxLines := items.Length - midPoint
    if (midPoint > maxLines)
        maxLines := midPoint

    PadString(str, width) {
        len := StrLen(str)
        if (len >= width)
            return str
        padding := width - len
        spaces := ""
        loop padding
            spaces .= " "
        return str . spaces
    }

    loop maxLines {
        leftText := ""
        rightText := ""
        leftKey := ""
        rightKey := ""
        if (A_Index <= midPoint) {
            leftItem := items[A_Index]
            leftText := PadString(leftItem.text, columnWidth)
            leftKey := leftItem.isEmpty ? "" : leftItem.char
        } else {
            leftText := PadString("", columnWidth)
        }
        rightIdx := A_Index + midPoint
        if (rightIdx <= items.Length) {
            rightItem := items[rightIdx]
            rightText := rightItem.text
            rightKey := rightItem.isEmpty ? "" : rightItem.char
        }
        lineText := leftText . columnSpacing . rightText
        rightStartCharPos := StrLen(leftText . columnSpacing) + 1
        lines.Push({ text: lineText, key: leftKey, keyRight: rightKey, rightStartCharPos: rightStartCharPos })
    }
    lines.Push({ text: "" })
    lines.Push({ text: "Backspace = back | Esc = close" })
    return lines
}

UtilitySelector_BuildDisplayText(isPortrait) {
    global g_UtilitySelectorMode
    if (g_UtilitySelectorMode = "top")
        return UtilitySelector_BuildTopLevelText()
    return UtilitySelector_BuildCategoryText(isPortrait)
}

UtilitySelector_BuildDisplayRich(isPortrait) {
    global g_UtilitySelectorMode
    if (g_UtilitySelectorMode = "top")
        return UtilitySelector_BuildTopLevelRich()
    return UtilitySelector_BuildCategoryRich(isPortrait)
}

UtilitySelector_RefreshUiAndHotkeys() {
    global g_HotstringSelectorGui, g_UtilitySelectorTitleCtrl, g_UtilitySelectorEditCtrl
    global g_UtilitySelectorIsPortrait, g_UtilitySelectorMonitor
    global g_UtilitySelectorMode, g_UtilitySelectorCategory

    if (!IsObject(g_HotstringSelectorGui) || !IsObject(g_UtilitySelectorEditCtrl))
        return

    title := "Utility Shortcuts"
    if (g_UtilitySelectorMode = "category" && g_UtilitySelectorCategory != "")
        title := title . " — " . g_UtilitySelectorCategory

    if (IsObject(g_UtilitySelectorTitleCtrl))
        try g_UtilitySelectorTitleCtrl.Text := title

    global g_UtilitySelectorFontSize
    displayText := UtilitySelector_BuildDisplayText(g_UtilitySelectorIsPortrait)
    displayLines := UtilitySelector_BuildDisplayRich(g_UtilitySelectorIsPortrait)
    try MnemonicRich_Render(g_UtilitySelectorEditCtrl, displayLines, g_UtilitySelectorFontSize, 6, "Segoe UI", "CDD6F4",
        "1E1E2E")

    ; Resize based on new content (reuse existing sizing rules)
    lineCount := 1
    loop parse, displayText, "`n"
        lineCount++
    lineHeight := 14
    textControlHeight := lineCount * lineHeight
    minHeight := 150

    monitorWidth := g_UtilitySelectorMonitor["width"]
    monitorHeight := g_UtilitySelectorMonitor["height"]
    monitorLeft := g_UtilitySelectorMonitor["left"]
    monitorTop := g_UtilitySelectorMonitor["top"]

    if (g_UtilitySelectorIsPortrait) {
        maxHeightPercent := 0.85
        maxHeight := Floor(monitorHeight * maxHeightPercent)
        if (textControlHeight < minHeight)
            textControlHeight := minHeight
        if (textControlHeight > maxHeight)
            textControlHeight := maxHeight
        baseWidth := (monitorWidth < 800) ? 400 : (monitorWidth < 1200) ? 500 : 500
        if (baseWidth > monitorWidth - 40)
            baseWidth := monitorWidth - 40
    } else {
        maxHeightPercent := (monitorHeight < 800) ? 0.90 : 0.75
        maxHeight := Floor(monitorHeight * maxHeightPercent)
        if (textControlHeight < minHeight)
            textControlHeight := minHeight
        if (textControlHeight > maxHeight)
            textControlHeight := maxHeight
        baseWidth := (monitorWidth < 1200) ? 650 : (monitorWidth < 1920) ? 800 : 1000
        if (baseWidth > monitorWidth - 40)
            baseWidth := monitorWidth - 40
    }

    textControlWidth := baseWidth - 20
    try {
        g_UtilitySelectorTitleCtrl.Move(, , textControlWidth)
        g_UtilitySelectorEditCtrl.Move(, , textControlWidth, textControlHeight)
    } catch {
    }

    totalHeight := 10 + 20 + 1 + 4 + textControlHeight + 6 + 18 + 10
    guiWidth := baseWidth

    marginX := 20
    marginY := 20
    guiX := monitorLeft + (monitorWidth - guiWidth) // 2
    guiY := monitorTop + (monitorHeight - totalHeight) // 2
    if (guiX < monitorLeft + marginX)
        guiX := monitorLeft + marginX
    if (guiY < monitorTop + marginY)
        guiY := monitorTop + marginY
    if (guiX + guiWidth > monitorLeft + monitorWidth - marginX)
        guiX := monitorLeft + monitorWidth - guiWidth - marginX
    if (guiY + totalHeight > monitorTop + monitorHeight - marginY)
        guiY := monitorTop + monitorHeight - totalHeight - marginY

    try g_HotstringSelectorGui.Show("NA w" . guiWidth . " h" . totalHeight . " x" . guiX . " y" . guiY)

    UtilitySelector_RebindHotkeys()
}

; =============================================================================
; ShowHotstringSelector()
; =============================================================================
; PURPOSE: Displays the hotstring selector modal GUI and enables character-based hotkeys.
;          GUI shows categorized list of available actions with their assigned characters.
;
; PROCESS:
;   1. Closes existing selector if already open
;   2. Builds character-to-action mappings via BuildHotstringCharMap()
;   3. Validates that at least one action is available (shows tray tip if none)
;   4. Gets categorized hotstring data via GetCategorizedHotstrings()
;   5. Calculates optimal GUI size based on monitor configuration
;   6. Creates and displays GUI with categorized action list
;   7. Enables hotkeys for all assigned characters
;   8. Enables Escape key handler for cancellation
;
; GUI FEATURES:
;   - Responsive layout: Adapts to monitor orientation (landscape/portrait)
;   - Dual-column layout for landscape monitors
;   - Single-column layout for portrait monitors
;   - Categories displayed in order: Prompts → General → Projects → Files & Links → Macros
;
; RETURNS: None (void function)
; SIDE EFFECTS: Sets g_HotstringSelectorActive = true, creates GUI object, enables hotkeys
; =============================================================================
ShowHotstringSelector() {
    global g_HotstringSelectorGui, g_HotstringSelectorActive, g_HotstringCharMap
    global g_HotstringHotkeyHandlers, g_HotstringCategories
    global g_HS_SelectorOpenFile, g_HS_SelectorCloseRequestFile, g_HS_SelectorCloseCheckTimer
    ; Close existing GUI if open
    if (g_HotstringSelectorActive && IsObject(g_HotstringSelectorGui)) {
        CleanupHotstringSelector()
        Sleep 50
    }

    ; In-process mutual exclusion: if project selector is active in this process, close it first
    try {
        if (IsSet(g_ProjectSelectorActive) && g_ProjectSelectorActive && IsSet(CleanupProjectSelector)) {
            CleanupProjectSelector()
            Sleep 50
        }
    } catch {
        ; Ignore failures – hotstring selector should still open
    }

    ; Cross-process safety: if WindowManagement project selector is open in another process,
    ; request it to close via the existing sentinel file mechanism.
    try {
        sentinel := A_ScriptDir "\.cursor\wm_selector_open"
        if (FileExist(sentinel)) {
            closeReq := A_ScriptDir "\.cursor\wm_selector_close_request"
            try FileAppend "", closeReq
            catch {
            }
            Sleep 50
        }
    } catch {
        ; Ignore IPC failures – hotstring selector should still open
    }

    ; Build character mapping
    g_HotstringCharMap := BuildHotstringCharMap()

    ; Check if we have any items to display (hotstrings, quick open files, or macros)
    global g_QuickOpenFileCharMap, g_MacroCharMap
    hasItems := (g_HotstringCharMap.Count > 0) || (g_QuickOpenFileCharMap.Count > 0) || (g_MacroCharMap.Count > 0)
    if (!hasItems) {
        ; Use tray notification to avoid stealing focus/closing other palettes
        TrayTip("Utility Selector", "No items found.", "IconX")
        SetTimer(() => TrayTip(), -5000)  ; Auto-hide after ~5s
        return
    }

    ; Get categorized hotstrings
    categorized := GetCategorizedHotstrings()

    ; =============================================================================
    ; Dynamic Modal UI Adaptation Based on Monitor Configuration
    ;
    ; Monitor dataset (reference only; UI uses live work area from the active window):
    ; {
    ;   "monitor_dataset": [
    ;     {
    ;       "id": 1,
    ;       "resolution": "1920x1080",
    ;       "orientation": "landscape",
    ;       "scale": "125%",
    ;       "ui_strategy": "dual_column_wide"
    ;     },
    ;     {
    ;       "id": 2,
    ;       "resolution": "3840x2160",
    ;       "orientation": "landscape",
    ;       "scale": "150%",
    ;       "ui_strategy": "dual_column_max_width_constrained"
    ;     },
    ;     {
    ;       "id": 3,
    ;       "resolution": "1080x1920",
    ;       "orientation": "portrait",
    ;       "scale": "100%",
    ;       "ui_strategy": "single_column_vertical_stretch"
    ;     },
    ;     {
    ;       "id": 4,
    ;       "resolution": "1080x1920",
    ;       "orientation": "portrait",
    ;       "scale": "100%",
    ;       "ui_strategy": "single_column_vertical_stretch"
    ;     }
    ;   ],
    ;   "instruction_logic": {
    ;     "landscape_rule": "Apply two-column layout; prioritize width expansion.",
    ;     "portrait_rule": "Apply single-column layout; prioritize height expansion."
    ;   }
    ; }
    ; =============================================================================

    ; Get monitor dimensions early for responsive sizing
    activeWin := 0
    try {
        activeWin := WinGetID("A")
    } catch {
        activeWin := 0
    }

    ; Default to primary monitor work area
    MonitorGetWorkArea(1, &monitorLeft, &monitorTop, &monitorRight, &monitorBottom)
    monitorWidth := monitorRight - monitorLeft
    monitorHeight := monitorBottom - monitorTop

    ; If we have an active window, find which monitor contains its center
    if (activeWin && activeWin != 0) {
        rect := Buffer(16, 0)
        if (DllCall("GetWindowRect", "ptr", activeWin, "ptr", rect)) {
            ; Calculate window center
            winLeft := NumGet(rect, 0, "int")
            winTop := NumGet(rect, 4, "int")
            winRight := NumGet(rect, 8, "int")
            winBottom := NumGet(rect, 12, "int")

            centerX := winLeft + (winRight - winLeft) // 2
            centerY := winTop + (winBottom - winTop) // 2

            ; Find which monitor contains the window center
            monitorCount := MonitorGetCount()
            loop monitorCount {
                idx := A_Index
                MonitorGetWorkArea(idx, &l, &t, &r, &b)
                if (centerX >= l && centerX <= r && centerY >= t && centerY <= b) {
                    monitorLeft := l
                    monitorTop := t
                    monitorRight := r
                    monitorBottom := b
                    monitorWidth := r - l
                    monitorHeight := b - t
                    break
                }
            }
        }
    }

    ; Detect monitor orientation: portrait (height > width) vs landscape (width >= height)
    isPortrait := (monitorHeight > monitorWidth)

    ; Create GUI (match Win+Alt+Shift+C AI Model Selector visual style)
    ; Create non-activating GUI so PowerToys Command Palette stays open
    g_HotstringSelectorGui := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x08000000", "Utility Shortcuts")
    g_HotstringSelectorGui.BackColor := "1E1E2E"
    g_HotstringSelectorGui.MarginX := 14
    g_HotstringSelectorGui.MarginY := 10
    ; Use slightly smaller font for compact display; Segoe UI to match C menu
    fontSize := (monitorHeight < 800) ? 9 : 9
    g_HotstringSelectorGui.SetFont("s" . fontSize . " cCDD6F4", "Segoe UI")

    ; Build reverse map: expansion -> character (legacy global mapping; still used elsewhere)
    expansionToChar := Map()
    for char, expansion in g_HotstringCharMap {
        expansionToChar[expansion] := char
    }

    ; Track which selector characters belong to the Prompts category (non-empty expansions only).
    ; Used to restrict the L-modifier redirect behavior to Prompts only.
    global g_HotstringPromptCharMap
    g_HotstringPromptCharMap := Map()
    try {
        global g_UtilityHotstringCharMapByCategory
        if (IsObject(g_UtilityHotstringCharMapByCategory) && g_UtilityHotstringCharMapByCategory.Has("Prompts")) {
            for ch, exp in g_UtilityHotstringCharMapByCategory["Prompts"] {
                try {
                    if (exp != "")
                        g_HotstringPromptCharMap[ch] := true
                } catch {
                }
            }
        }
    } catch {
        ; Ignore prompt-char tracking failures (selector still works normally)
    }

    ; Build items list grouped by category for two-column layout
    ; Collect all items first, then format in two columns
    hotstringCount := 0
    allItems := []  ; Array of {category, char, text, isEmpty}

    ; Build a map of character index to hotstring/file info
    charIndexToHotstring := Map()
    charIndex := 1
    for category in g_HotstringCategories {
        for item in categorized[category] {
            if (charIndex <= g_HotstringCharSequence.Length) {
                charIndexToHotstring[charIndex] := { hotstring: item, category: category }
            }
            charIndex++
        }
    }

    ; First, add explicitly assigned macros at their character positions
    global g_MacroCharMap
    for char, macroFunc in g_MacroCharMap {
        ; Find the macro entry for this function
        macroEntry := ""
        for macro in g_Macros {
            if (macro.func = macroFunc) {
                macroEntry := macro
                break
            }
        }
        if (macroEntry != "") {
            ; Find the index of this character in the sequence
            charIndexInSeq := 0
            for idx, seqChar in g_HotstringCharSequence {
                if (seqChar = char) {
                    charIndexInSeq := idx
                    break
                }
            }
            if (charIndexInSeq > 0) {
                itemText := "[" . char . "] > " . macroEntry.title
                hotstringCount++
                allItems.Push({ category: "Macros", char: char, text: itemText, isEmpty: false, explicitIndex: charIndexInSeq })
            }
        }
    }

    ; Collect all items with their categories (sequential assignment for non-explicit items)
    currentCharIndex := 1
    for category in g_HotstringCategories {
        ; Calculate how many character slots belong to this category
        categorySlotCount := categorized[category].Length

        if (categorySlotCount > 0 || currentCharIndex <= g_HotstringCharSequence.Length) {
            ; Collect all character slots for this category (including empty ones)
            loop categorySlotCount {
                if (currentCharIndex <= g_HotstringCharSequence.Length) {
                    ; Skip reserved empty char if set so it always shows as (empty)
                    while (currentCharIndex <= g_HotstringCharSequence.Length && g_ReservedEmptyChar != "" &&
                        g_HotstringCharSequence[currentCharIndex] = g_ReservedEmptyChar)
                        currentCharIndex++
                    if (currentCharIndex > g_HotstringCharSequence.Length)
                        break
                    char := g_HotstringCharSequence[currentCharIndex]

                    ; Skip if this character is already explicitly assigned to a macro
                    if (g_MacroCharMap.Has(char)) {
                        currentCharIndex++
                        continue
                    }

                    itemText := ""
                    isEmpty := false

                    ; Check if this character has a hotstring assigned
                    if (charIndexToHotstring.Has(currentCharIndex)) {
                        hsInfo := charIndexToHotstring[currentCharIndex]
                        hs := hsInfo.hotstring

                        ; Skip macros that have explicit char assignments
                        if (hsInfo.category = "Macros" && hs.HasProp("char") && hs.char != "") {
                            ; This macro has an explicit assignment, skip it here
                            currentCharIndex++
                            continue
                        }

                        ; Use title if available (for all categories including quick open files), otherwise use preview text
                        if (hs.HasProp("title") && hs.title != "") {
                            itemText := "[" . char . "] > " . hs.title
                            hotstringCount++
                        } else if (hs.HasProp("expansion") && hs.expansion != "") {
                            preview := GetPreviewText(hs.expansion)
                            itemText := "[" . char . "] > " . preview
                            hotstringCount++
                        } else {
                            ; Empty placeholder slot
                            itemText := "[" . char . "] > (empty)"
                            isEmpty := true
                        }
                    } else {
                        ; Character slot exists but no hotstring assigned
                        if (char = "l") {
                            itemText :=
                                "[L] > Gemini: L = arm; L+L = open Gemini + paste first snippet (or Ctrl+Alt+Win+L)"
                            isEmpty := false
                        } else {
                            itemText := "[" . char . "] > (empty)"
                            isEmpty := true
                        }
                    }

                    topCategory := UtilitySelector_MapInternalCategoryToTop(category)
                    allItems.Push({ category: topCategory, char: char, text: itemText, isEmpty: isEmpty })
                    currentCharIndex++
                }
            }
        }
    }

    ; Sort allItems by explicitIndex (if exists) or sequential position, then by category order
    ; Items with explicitIndex should be at their explicit position
    sortedItems := []
    for idx, seqChar in g_HotstringCharSequence {
        ; First check for explicitly assigned items at this position
        found := false
        for item in allItems {
            if (item.HasProp("explicitIndex") && item.explicitIndex = idx) {
                sortedItems.Push(item)
                found := true
                break
            }
        }
        ; If not found as explicit, check for sequential items
        if (!found) {
            for item in allItems {
                if (!item.HasProp("explicitIndex") && item.char = seqChar) {
                    ; Check if this item was already added
                    alreadyAdded := false
                    for added in sortedItems {
                        if (added.char = item.char && added.text = item.text) {
                            alreadyAdded := true
                            break
                        }
                    }
                    if (!alreadyAdded) {
                        sortedItems.Push(item)
                        break
                    }
                }
            }
        }
    }

    allItems := sortedItems

    ; Remove empty placeholder slots (no Unassigned category in the revised hierarchy)
    filtered := []
    for item in allItems {
        if (!item.isEmpty)
            filtered.Push(item)
    }
    allItems := filtered

    ; -------------------------------------------------------------------------
    ; Utility Shortcuts rendering: build items using explicit/per-category maps
    ; -------------------------------------------------------------------------
    ; The legacy block above assigns display chars by sequential slot, which can
    ; differ from mnemonic explicit chars (e.g., Projects 'a' / '0'). For the
    ; hierarchical selector, rebuild the item list from the category-scoped
    ; mapping so display + hotkeys match the selected category.
    try {
        global g_UtilityHotstringCharMapByCategory, g_QuickOpenFileCharMap, g_MacroCharMap, g_HotstringCharSequence

        charOrder := Map()
        for idx, c in g_HotstringCharSequence
            charOrder[c] := idx

        rebuilt := []
        seen := Map() ; key = category "|" char

        BuildExpansionToChar(catMap) {
            m := Map()
            try {
                for ch, exp in catMap
                    m[exp] := ch
            } catch {
            }
            return m
        }

        AddItem(cat, ch, titleText, seenRef, rebuiltRef) {
            if (ch = "" || titleText = "")
                return
            key := cat . "|" . ch
            if (seenRef.Has(key))
                return
            seenRef[key] := true
            rebuiltRef.Push({ category: cat, char: ch, text: "[" . ch . "] > " . titleText, isEmpty: false })
        }

        ; Hotstrings (text expansions) by category using category-scoped maps
        for cat in ["Prompts", "Projects", "General", "Hotstrings"] {
            if (!categorized.Has(cat))
                continue
            catMap := (IsObject(g_UtilityHotstringCharMapByCategory) && g_UtilityHotstringCharMapByCategory.Has(cat)) ?
                g_UtilityHotstringCharMapByCategory[cat] : Map()
            expToChar := BuildExpansionToChar(catMap)

            for hs in categorized[cat] {
                try {
                    if (!hs.HasProp("expansion") || hs.expansion = "")
                        continue

                    ch := (hs.HasProp("char") && hs.char != "") ? hs.char : expToChar.Get(hs.expansion, "")
                    if (ch = "")
                        continue

                    titleText := ""
                    if (hs.HasProp("title") && hs.title != "")
                        titleText := hs.title
                    else
                        titleText := GetPreviewText(hs.expansion)

                    AddItem(cat, ch, titleText, seen, rebuilt)
                } catch {
                }
            }
        }

        ; Files & Links show under Links in Utility menu
        filePathToChar := Map()
        try {
            for ch, fp in g_QuickOpenFileCharMap
                filePathToChar[fp] := ch
        } catch {
        }
        if (categorized.Has("Files & Links")) {
            for fileEntry in categorized["Files & Links"] {
                try {
                    ch := filePathToChar.Get(fileEntry.filePath, "")
                    if (ch = "")
                        continue
                    titleText := fileEntry.HasProp("title") ? fileEntry.title : ""
                    if (titleText = "")
                        titleText := fileEntry.filePath
                    AddItem("Links", ch, titleText, seen, rebuilt)
                } catch {
                }
            }
        }

        ; Macros
        funcToChar := Map()
        try {
            for ch, fn in g_MacroCharMap
                funcToChar[fn] := ch
        } catch {
        }
        if (categorized.Has("Macros")) {
            for macroEntry in categorized["Macros"] {
                try {
                    ch := (macroEntry.HasProp("char") && macroEntry.char != "") ? macroEntry.char : funcToChar.Get(
                        macroEntry.func, "")
                    if (ch = "")
                        continue
                    titleText := macroEntry.HasProp("title") ? macroEntry.title : ""
                    if (titleText = "")
                        titleText := "(macro)"
                    AddItem("Macros", ch, titleText, seen, rebuilt)
                } catch {
                }
            }
        }

        ; Sort by character order for a consistent layout
        try {
            rebuilt.Sort((a, b) => (charOrder.Get(a.char, 9999) = charOrder.Get(b.char, 9999)) ?
                (a.category < b.category ? -1 : 1) :
                (charOrder.Get(a.char, 9999) < charOrder.Get(b.char, 9999) ? -1 : 1))
        } catch {
        }

        allItems := rebuilt
    } catch {
        ; Fallback to legacy list if rebuild fails
    }

    ; Helper function to pad string to specified width
    PadString(str, width) {
        len := StrLen(str)
        if (len >= width)
            return str
        padding := width - len
        spaces := ""
        loop padding {
            spaces .= " "
        }
        return str . spaces
    }

    ; Helper function to center string in specified width
    CenterString(str, width) {
        len := StrLen(str)
        if (len >= width)
            return str
        padding := (width - len) // 2
        leftSpaces := ""
        rightSpaces := ""
        loop padding {
            leftSpaces .= " "
        }
        loop (width - len - padding) {
            rightSpaces .= " "
        }
        return leftSpaces . str . rightSpaces
    }

    ; Helper function to create separator line
    CreateSeparator(width) {
        separator := ""
        loop width {
            separator .= "─"
        }
        return separator
    }

    ; Cache UI data for hierarchical selector refresh
    global g_UtilitySelectorAllItems, g_UtilitySelectorIsPortrait, g_UtilitySelectorMonitor
    g_UtilitySelectorAllItems := allItems
    g_UtilitySelectorIsPortrait := isPortrait
    g_UtilitySelectorMonitor := Map("left", monitorLeft, "top", monitorTop, "width", monitorWidth, "height",
        monitorHeight)

    ; Always open in top-level category screen
    global g_UtilitySelectorMode, g_UtilitySelectorCategory
    g_UtilitySelectorMode := "top"
    g_UtilitySelectorCategory := ""

    displayText := UtilitySelector_BuildDisplayText(isPortrait)
    ; Calculate text control height based on actual content (number of lines)
    ; Count actual lines in displayText (including empty lines for spacing)
    lineCount := 1  ; Start at 1 (first line doesn't have a newline before it)
    loop parse, displayText, "`n" {
        lineCount++
    }
    ; Calculate height: ~14 pixels per line (compact display)
    lineHeight := 14
    textControlHeight := lineCount * lineHeight
    ; Ensure minimum and maximum bounds
    minHeight := 150

    ; Adjust sizing based on orientation
    if (isPortrait) {
        ; PORTRAIT: Prioritize height expansion, use narrower width
        ; Use more vertical space for portrait monitors (up to 85% of height)
        maxHeightPercent := 0.85
        maxHeight := Floor(monitorHeight * maxHeightPercent)
        if (textControlHeight < minHeight)
            textControlHeight := minHeight
        if (textControlHeight > maxHeight)
            textControlHeight := maxHeight

        ; Narrower width for portrait (optimized for vertical scrolling)
        baseWidth := (monitorWidth < 800) ? 400 : (monitorWidth < 1200) ? 500 : 500
        ; Ensure we don't exceed monitor width with margins
        if (baseWidth > monitorWidth - 40)
            baseWidth := monitorWidth - 40
    } else {
        ; LANDSCAPE: Prioritize width expansion, use two-column layout
        ; Use adaptive max height: 85% for large monitors, 90% for small monitors
        maxHeightPercent := (monitorHeight < 800) ? 0.90 : 0.75
        maxHeight := Floor(monitorHeight * maxHeightPercent)
        if (textControlHeight < minHeight)
            textControlHeight := minHeight
        if (textControlHeight > maxHeight)
            textControlHeight := maxHeight

        ; Wide width for landscape (two-column layout)
        ; Further reduced width to eliminate empty space: 650px minimum, scale up to 1000px based on monitor width
        baseWidth := (monitorWidth < 1200) ? 650 : (monitorWidth < 1920) ? 800 : 1000
        ; Ensure we don't exceed monitor width with margins
        if (baseWidth > monitorWidth - 40)
            baseWidth := monitorWidth - 40
    }
    textControlWidth := baseWidth - 20  ; Account for margins

    ; Title and separator (compact)
    g_HotstringSelectorGui.SetFont("s11 cCDD6F4 Bold", "Segoe UI")
    global g_UtilitySelectorTitleCtrl
    g_UtilitySelectorTitleCtrl := g_HotstringSelectorGui.Add("Text", "w" . textControlWidth . " Center",
        "Utility Shortcuts")
    g_HotstringSelectorGui.Add("Text", "w" . textControlWidth . " h1 Background45475A")
    g_HotstringSelectorGui.SetFont("s" . fontSize . " cCDD6F4", "Segoe UI")

    ; Cache base font size for RichEdit rendering in refresh
    global g_UtilitySelectorFontSize
    g_UtilitySelectorFontSize := fontSize

    ; Enable vertical scrolling for long content (RichEdit so we can style mnemonic letters)
    global g_UtilitySelectorEditCtrl
    MnemonicRich_EnsureDll()
    g_UtilitySelectorEditCtrl := g_HotstringSelectorGui.Add("Custom",
        "ClassRichEdit50W w" . textControlWidth . " h" . textControlHeight
        . " +0x44 -E0x200 +VScroll -HScroll -Border Background1E1E2E")
    try MnemonicRich_Render(g_UtilitySelectorEditCtrl, UtilitySelector_BuildDisplayRich(isPortrait), fontSize, 6,
    "Segoe UI",
    "CDD6F4", "1E1E2E")
    g_HotstringSelectorGui.SetFont("s9 c89B4FA", "Segoe UI")
    g_HotstringSelectorGui.Add("Text", "w" . textControlWidth . " Center", "Press Escape to close.")

    ; Total height: margins + title + separator + gap + content + hint + spacing (no button)
    totalHeight := 10 + 20 + 1 + 4 + textControlHeight + 6 + 18 + 10
    guiWidth := baseWidth

    ; Calculate center position for the GUI with margins
    marginX := 20  ; Horizontal margin from screen edges
    marginY := 20  ; Vertical margin from screen edges
    guiX := monitorLeft + (monitorWidth - guiWidth) // 2
    guiY := monitorTop + (monitorHeight - totalHeight) // 2

    ; Ensure the GUI stays within monitor bounds with margins
    if (guiX < monitorLeft + marginX)
        guiX := monitorLeft + marginX
    if (guiY < monitorTop + marginY)
        guiY := monitorTop + marginY
    if (guiX + guiWidth > monitorLeft + monitorWidth - marginX)
        guiX := monitorLeft + monitorWidth - guiWidth - marginX
    if (guiY + totalHeight > monitorTop + monitorHeight - marginY)
        guiY := monitorTop + monitorHeight - totalHeight - marginY

    ; Show GUI centered on the active window's monitor
    g_HotstringSelectorGui.Show("NA w" . guiWidth . " h" . totalHeight . " x" . guiX . " y" . guiY)

    ; Set active flag
    g_HotstringSelectorActive := true

    ; Cross-process IPC: mark Hotstring Selector as open and start close-request timer
    try {
        DirCreate(A_ScriptDir "\.cursor")
        FileAppend("", g_HS_SelectorOpenFile)
    } catch {
    }
    g_HS_SelectorCloseCheckTimer := SetTimer(Utils_CheckHotstringSelectorCloseRequest, 120)

    ; Bind top-level hotkeys (1-6) + Escape; category view binds are applied when user selects a category.
    UtilitySelector_RebindHotkeys()
}

; =============================================================================
; Hotkey Handler: Windows + Alt + Shift + U (#!+U)
; =============================================================================
; PURPOSE: Toggles the hotstring selector GUI on/off.
;
; BEHAVIOR:
;   - If selector is currently open: Closes selector via CleanupHotstringSelector()
;   - If selector is closed: Opens selector via ShowHotstringSelector()
;
; TECHNICAL NOTE: The character sequence displayed in the GUI must remain consistent
;                  and list every slot in order, even when empty, to ensure downstream
;                  AI systems and debugging tools can reliably parse the full character set.
; =============================================================================
#!+U::
{
    global g_HotstringSelectorActive, g_HotstringSelectorGui

    ; Toggle behavior: Close if open, open if closed
    if (g_HotstringSelectorActive && IsObject(g_HotstringSelectorGui)) {
        CleanupHotstringSelector()
    } else {
        ShowHotstringSelector()
    }
}

; Ctrl+Alt+Win+L - direct D2C submit path (paste + Enter, then monitor)
^!#L:: D2C_FlowManager.GetInstance().StartFromHotstring()

; Ctrl+Alt+Win+4 - Start dictation first (~#!+0), then Gemini tab 1/2 toggle + banner (user can speak while tab switches)
^!#4::
{
    global g_GeminiToggleTab, g_DictationActive

    ; SendLevel 1 so generated #!+0 is processed as a hotkey and reaches ~#!+0.
    if (!g_DictationActive) {
        SendLevel 1
        Send "#!+0"
        SendLevel 0
    }

    geminiHwnd := 0
    try {
        for hwnd in WinGetList("ahk_exe chrome.exe") {
            try {
                if InStr(WinGetTitle("ahk_id " hwnd), "gemini", false) {
                    geminiHwnd := hwnd
                    break
                }
            } catch {
            }
        }
    } catch {
    }

    if (!geminiHwnd)
        return

    WinActivate("ahk_id " geminiHwnd)
    if (!WinWaitActive("ahk_id " geminiHwnd, , 2))
        return

    Sleep(120)
    uia := UIA_Browser("ahk_id " geminiHwnd)
    tabInfo := GetChromeActiveTabIndex(uia)
    if (!tabInfo) {
        Sleep(150)
        tabInfo := GetChromeActiveTabIndex(uia)
    }
    targetTab := (tabInfo && tabInfo.index == 1) ? 2 : 1
    g_GeminiToggleTab := targetTab
    Send("^" . targetTab)
    ShowSingleCharTabBanner_Utils(targetTab)
    Sleep(200)

    dictationOk := g_DictationActive
    tabInfoAfter := GetChromeActiveTabIndex(uia)
    if (!tabInfoAfter) {
        Sleep(100)
        tabInfoAfter := GetChromeActiveTabIndex(uia)
    }
    tabOk := tabInfoAfter && tabInfoAfter.index == targetTab
    if (!dictationOk || !tabOk)
        ShowCenteredOverlay_Utils("❌ Shortcut execution failed", 2000, BANNER_ACCENT_ERROR)
}

; Ctrl+Alt+Win+2..8 - same macros as HotStrings panel (Win+Alt+Shift+U); secondary triggers only
^!#2:: QuickUpdateScripts()
^!#3:: ToggleOutlookAndTeams()
^!#5:: CleanClipboard()
^!#6:: Cursor_FindComposerIconAcrossInstances()
; Same macro; InputLevel 10 + hook so chord wins over other low-level handlers (optional ghosting fallback: ^!#j).
#InputLevel 10
#UseHook
^!#7:: MarkLastClipAsFavorite()
^!#j:: MarkLastClipAsFavorite()
#UseHook False
#InputLevel 0
^!#8:: DesktopToRecycle_Trigger()
; Ctrl+Alt+Win+9 / +B — Handy Parakeet V3 / V2 (g_HandyAiModels slots 2 and 1)
^!#9:: ExecuteHandyAiModelSelection(HANDY_AI_SLOT_PARAKEET_V3)
^!#b:: ExecuteHandyAiModelSelection(HANDY_AI_SLOT_PARAKEET_V2)

; =============================================================================
; Alt+Shift+W Shortcut
; Hotkey: Alt+Shift+W
; Sends Alt+Shift+W again, then shows a message box
; =============================================================================
!+W::
{
    ; Send Alt+Shift+W again
    ; Small delay to ensure previous key release is complete
    Sleep 50

    ; Send Win+Ctrl+Alt+Y using SendInput for better reliability
    ; SendInput is more reliable for complex modifier combinations
    SendInput "#^!y"

    ; Show message box

    Sleep 50

    Send '""'

    Sleep 50

    Send "{Left}"

}

; =============================================================================
; Focus Mode (multi-monitor blackout)
; Hotkey: Win+Alt+Shift+Y
; =============================================================================

global g_FocusModeOn := false
global g_FocusModeActiveMonitor := 0
global g_FocusModeOverlays := []  ; array of GUI overlays (one per covered monitor)
global g_FocusModeTrackedWindow := 0  ; window handle that was active when focus mode was enabled
global g_FocusModeMonitorTimer := false  ; timer for monitoring window focus changes

GetActiveMonitorIndex() {
    hwnd := WinExist("A")
    if (!hwnd) {
        return 0
    }

    rect := Buffer(16, 0)
    if (!DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect)) {
        return 0
    }

    left := NumGet(rect, 0, "int")
    top := NumGet(rect, 4, "int")
    right := NumGet(rect, 8, "int")
    bottom := NumGet(rect, 12, "int")

    centerX := left + (right - left) // 2
    centerY := top + (bottom - top) // 2

    monitorCount := MonitorGetCount()
    loop monitorCount {
        MonitorGet(A_Index, &ml, &mt, &mr, &mb)
        if (centerX >= ml && centerX <= mr && centerY >= mt && centerY <= mb) {
            return A_Index
        }
    }

    return 0
}

EnableFocusMode() {
    global g_FocusModeOn, g_FocusModeActiveMonitor, g_FocusModeOverlays, g_FocusModeTrackedWindow

    ; Check if focus mode is already active by verifying state and overlays
    hasActiveOverlays := false
    if (IsObject(g_FocusModeOverlays) && g_FocusModeOverlays.Length > 0) {
        ; Verify at least one overlay window still exists
        for overlay in g_FocusModeOverlays {
            if (IsObject(overlay) && overlay.Hwnd && WinExist("ahk_id " . overlay.Hwnd)) {
                hasActiveOverlays := true
                break
            }
        }
    }

    ; If already enabled (by flag or by existing overlays), clean up and return
    if (g_FocusModeOn || hasActiveOverlays) {
        ; If overlays exist but flag is wrong, fix the state
        if (hasActiveOverlays && !g_FocusModeOn) {
            ; Clean up the orphaned overlays first
            DisableFocusMode()
        } else {
            ; Already properly enabled, just return
            return
        }
    }

    activeMon := GetActiveMonitorIndex()
    if (!activeMon) {
        return
    }

    g_FocusModeActiveMonitor := activeMon
    g_FocusModeOverlays := []

    ; Store the active window handle when focus mode is enabled
    g_FocusModeTrackedWindow := WinExist("A")

    ; Start monitoring for window focus changes
    StartFocusModeWindowMonitor()

    monitorCount := MonitorGetCount()

    loop monitorCount {
        MonitorGet(A_Index, &ml, &mt, &mr, &mb)
        name := ""
        dx := "", dy := ""
        try name := MonitorGetName(A_Index)
    }

    loop monitorCount {
        i := A_Index
        if (i = activeMon) {
            continue
        }

        MonitorGet(i, &l, &t, &r, &b)
        w := r - l
        h := b - t
        if (w <= 0 || h <= 0) {
            continue
        }

        overlay := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x20") ; WS_EX_TRANSPARENT => click-through
        ; Critical: disable GUI DPI scaling so x/y/w/h are interpreted in raw screen coords
        ; Without this, AHK scales the overlay (e.g. 150%) and it spills into other monitors.
        overlay.Opt("-DPIScale")
        overlay.BackColor := "000000"
        overlay.Show("NA x" l " y" t " w" w " h" h)
        g_FocusModeOverlays.Push(overlay)

        ; Query actual overlay window rect (physical pixels) to detect scaling/border issues
        rect := Buffer(16, 0)
        ok := 0
        try ok := DllCall("GetWindowRect", "ptr", overlay.Hwnd, "ptr", rect)
    }

    g_FocusModeOn := true
}

DisableFocusMode() {
    global g_FocusModeOn, g_FocusModeActiveMonitor, g_FocusModeOverlays, g_FocusModeTrackedWindow,
        g_FocusModeMonitorTimer

    ; Stop monitoring window focus changes
    StopFocusModeWindowMonitor()

    for overlay in g_FocusModeOverlays {
        try {
            if (IsObject(overlay) && overlay.Hwnd) {
                overlay.Destroy()
            }
        } catch {
            ; Ignore
        }
    }

    g_FocusModeOverlays := []
    g_FocusModeActiveMonitor := 0
    g_FocusModeTrackedWindow := 0
    g_FocusModeOn := false
}

ToggleFocusMode() {
    global g_FocusModeOn, g_FocusModeOverlays

    ; Check if focus mode is actually active by verifying both state and overlays
    ; This prevents adding layers when state is out of sync
    hasActiveOverlays := false
    if (IsObject(g_FocusModeOverlays) && g_FocusModeOverlays.Length > 0) {
        ; Verify at least one overlay window still exists
        for overlay in g_FocusModeOverlays {
            if (IsObject(overlay) && overlay.Hwnd && WinExist("ahk_id " . overlay.Hwnd)) {
                hasActiveOverlays := true
                break
            }
        }
    }

    ; Determine actual state: if overlays exist OR flag is set, consider it ON
    actualState := g_FocusModeOn || hasActiveOverlays

    if (actualState) {
        ; Focus mode is on - disable it
        DisableFocusMode()
    } else {
        ; Focus mode is off - enable it
        ; First ensure any stale overlays are cleaned up
        if (hasActiveOverlays && !g_FocusModeOn) {
            ; State mismatch: overlays exist but flag says off - clean up first
            DisableFocusMode()
        }
        EnableFocusMode()
    }
}

; Monitor window focus changes and automatically disable focus mode when active window changes
FocusModeWindowMonitor(*) {
    global g_FocusModeOn, g_FocusModeTrackedWindow

    ; Only monitor if focus mode is active
    if (!g_FocusModeOn) {
        return
    }

    ; Check if tracked window still exists
    if (g_FocusModeTrackedWindow && !WinExist("ahk_id " . g_FocusModeTrackedWindow)) {
        ; Tracked window was closed - automatically disable focus mode
        DisableFocusMode()
        return
    }

    ; Window activation monitoring removed - blackout now only disabled via manual toggle (#!+Y)
}

; Start monitoring window focus changes
StartFocusModeWindowMonitor() {
    global g_FocusModeMonitorTimer

    ; Stop any existing timer first
    StopFocusModeWindowMonitor()

    ; Start timer to check window focus every 200ms
    g_FocusModeMonitorTimer := SetTimer(FocusModeWindowMonitor, 200)
}

; Stop monitoring window focus changes
StopFocusModeWindowMonitor() {
    global g_FocusModeMonitorTimer

    if (g_FocusModeMonitorTimer) {
        try {
            SetTimer(g_FocusModeMonitorTimer, 0)  ; Disable timer
        } catch {
            ; Ignore errors
        }
        g_FocusModeMonitorTimer := false
    }
}

#!+Y::
{
    ToggleFocusMode()
}

; =============================================================================
; Print Screen with Chime
; Hotkey: Alt+PrintScreen
; Intercepts the hotkey to prevent other apps from handling it,
; manually triggers the screenshot, and plays a single chime
; =============================================================================
global g_LastPrintScreenSound := 0  ; Track last sound time for debouncing
global g_PrintScreenInProgress := false  ; Prevents recursion from Send

; Audio firewall for PrintScreen: Throttle sounds to prevent duplicates
; Uses Critical section to ensure atomic check-and-update (like dictation mode)
SafePlayPrintScreenSound() {
    Critical  ; Prevents thread interruption - ensures atomic check-and-update sequence
    global g_LastPrintScreenSound

    ; If less than 1000ms has passed since last sound, ignore this call
    if (A_TickCount - g_LastPrintScreenSound < 1000) {
        return
    }

    ; Update timestamp and play sound (if enabled)
    g_LastPrintScreenSound := A_TickCount
    if (IsSoundEnabled()) {
        SoundPlay(A_ScriptDir . "\sounds\print-screen.wav")
    }
}

; Set higher InputLevel to ensure our handler processes before others
#InputLevel 10
!PrintScreen::  ; Removed ~ prefix to CONSUME the hotkey (prevents other apps from receiving it)
{
    global g_PrintScreenInProgress

    ; Prevent recursion: if we're already processing, skip (this handles Send retriggering)
    if (g_PrintScreenInProgress) {
        return
    }

    ; Set flag to prevent recursion from the Send below
    g_PrintScreenInProgress := true

    ; Manually send Alt+PrintScreen to Windows to trigger the screenshot
    ; SendInput is more reliable and won't retrigger our hotkey due to SendLevel
    SendInput("!{PrintScreen}")

    ; Brief delay to ensure screenshot is captured, then play single chime
    ; Uses Critical section to prevent duplicate sounds from concurrent handlers
    Sleep 10
    SafePlayPrintScreenSound()

    ; Reset flag after a brief delay to allow normal operation
    Sleep 100
    g_PrintScreenInProgress := false
}
#InputLevel 0

; True when physical or synthetic Escape must not be sent (matches Escape:: Handy block).
IsHandyDictationEscapeSuppressed() {
    global g_DictationActive
    return g_DictationActive || WinActive("Recording ahk_exe handy.exe") || WinActive(
        "Recording Overlay ahk_exe handy.exe") || WinExist("Recording ahk_exe handy.exe") || WinExist(
            "Recording Overlay ahk_exe handy.exe")
}

; Helper to send Escape while respecting dictation state.
; No-op when Handy recording is active or its Recording/Overlay windows exist (same as Escape::).
SendEscape(count := 1) {
    if (IsHandyDictationEscapeSuppressed()) {
        return
    }
    if (count > 1)
        Send("{Escape " . count . "}")
    else
        Send("{Escape}")
}

; =============================================================================
; Block Escape Key from handy.exe
; Prevents Escape key from closing handy.exe while still allowing
; Escape to work normally in other applications
; =============================================================================

; Optional escape callback: when set (e.g. by WindowManagement for project selector), Utils runs it and consumes Escape.
global g_OnEscapePressed := ""

; Force hook-based hotkey to intercept Escape at a lower level
; This should catch it before handy.exe's hook processes it
#UseHook
#InputLevel 10
Escape::
{
    ; Use state-based blocking: check g_DictationActive instead of checking window each time
    ; This ensures Esc remains restricted for the entire duration of dictation
    global g_DictationActive, g_OnEscapePressed

    ; If a consumer (e.g. project selector) registered an escape handler, run it and consume the key
    if (g_OnEscapePressed) {
        try {
            g_OnEscapePressed.Call()
        } catch {
        }
        return
    }

    ; Cross-process: if project selector is open (WM process), request close via file so WM's timer will run CleanupProjectSelector
    try {
        sentinel := A_ScriptDir "\.cursor\wm_selector_open"
        if (FileExist(sentinel)) {
            closeReq := A_ScriptDir "\.cursor\wm_selector_close_request"
            try FileAppend "", closeReq
            catch {
            }
            return
        }
    } catch {
    }

    ; Block Escape when dictation is active, or Handy Recording/Overlay exists or is focused.
    ; WinExist covers unfocused tool windows that do not hold foreground focus.
    if (IsHandyDictationEscapeSuppressed()) {
        ; Block Escape from reaching handy.exe - do nothing
        ; This prevents the dictation software from closing
        return
    }

    ; Otherwise, forward Escape to the system
    ; Use SendInput for more reliable key forwarding
    SendInput "{Escape}"
}
#InputLevel 0

; =============================================================================
; Dictation Indicator - Red pulsing inner square with yellow border
; Anchored to the top-center of the active window (clamped inside); falls back to
; active monitor work area. Follows focus/window moves via pulse timer. Toggles with Win+Alt+Shift+0.
; =============================================================================

; Global variables for dictation indicator
global g_DictationActive := false
global g_DictationIndicatorGui := false
global g_DictationIndicatorText := false  ; Text control for status messages
global g_DictationPulseTimer := false
global g_DictationCheckTimer := false  ; Timer to check if Recording window still exists
global g_DictationPulseDirection := 1  ; 1 = fading in, -1 = fading out
global g_DictationPulseOpacity := 128  ; Current opacity (50-255)
global g_DictationFollowCache := ""  ; "x,y,w,h" when unchanged skip redundant Move (follow active window)
global g_DictationCompletionChimeScheduled := false  ; Flag to prevent multiple completion chimes
global g_LastDictationSoundTick := 0  ; Timestamp of last dictation sound to throttle audio output
global g_DictationStartSound := A_ScriptDir . "\sounds\speach-start.wav"
global g_DictationStopSound := A_ScriptDir . "\sounds\speach-finished.wav"
global g_PendingDictationAction := ""  ; Action to execute after transcription: "Paste" (reserved for future)
global g_PendingGeminiPromptAfterDictation := false  ; When set by ~#!+0 stop, show "Send to Gemini? Y (4s)" after completion
global g_DictationGeminiConfirmBannerVisible := false  ; Guard: only one "Send to Gemini?" banner at a time
global g_KeepIndicatorVisible := false  ; Flag to keep indicator visible until paste action completes
global g_LastStateTransitionTick := 0  ; Timestamp of last state transition to prevent rapid re-detection
global g_DictationSoundPlayed := false  ; Atomic test-and-set: one start chime per session
global g_DictationStartClipboardText := "" ; Track clipboard content at start to detect changes
global g_DictationHotkeyOwnerHandle := 0 ; Named mutex handle for cross-process single-owner dictation hotkey
global g_DictationHotkeyIsOwner := false ; True only in the single process that owns dictation hotkey handling

; Ensure only one script process handles the dictation hotkey logic.
InitializeDictationHotkeyOwnership() {
    global g_DictationHotkeyOwnerHandle, g_DictationHotkeyIsOwner
    mutexName := "Local\D2C_Dictation_Hotkey_Owner"
    hMutex := DllCall("CreateMutex", "Ptr", 0, "Int", 0, "Str", mutexName, "Ptr")
    if (!hMutex) {
        g_DictationHotkeyIsOwner := false
        return
    }
    err := DllCall("GetLastError", "UInt")
    if (err = 183) { ; ERROR_ALREADY_EXISTS
        g_DictationHotkeyIsOwner := false
        DllCall("CloseHandle", "Ptr", hMutex)
        return
    }
    g_DictationHotkeyOwnerHandle := hMutex
    g_DictationHotkeyIsOwner := true
}

ReleaseDictationHotkeyOwnership(*) {
    global g_DictationHotkeyOwnerHandle
    if (g_DictationHotkeyOwnerHandle) {
        try DllCall("CloseHandle", "Ptr", g_DictationHotkeyOwnerHandle)
        g_DictationHotkeyOwnerHandle := 0
    }
}

InitializeDictationHotkeyOwnership()
OnExit(ReleaseDictationHotkeyOwnership)

; Debug logging helper for dictation workflow
LogDebug(sessionId, runId, hypothesisId, location, message, data := "") {
    logPath := A_ScriptDir "\.cursor\debug.log"
    timestamp := A_Now "." Format("{:03}", A_MSec)
    logEntry := Format(
        '{{"sessionId":"{}","runId":"{}","hypothesisId":"{}","location":"{}","message":"{}","timestamp":"{}","data":{}}}',
        sessionId, runId, hypothesisId, location, message, timestamp, data ? '"' . data . '"' : '""')
    try {
        FileAppend(logEntry . "`n", logPath)
    } catch {
        ; Silently ignore logging errors
    }
}

; Constants for dictation indicator
global DICTATION_SQUARE_SIZE := 150  ; Inner red area (was 50 base scaled up)
global DICTATION_BORDER_PX := 2      ; Yellow border for colorblind visibility (outside red)
global DICTATION_YELLOW_BORDER := "F1C40F"  ; Same hue family as BANNER_ACCENT_INTERMEDIATE
global DICTATION_PULSE_MIN := 50      ; Minimum opacity (~20%)
global DICTATION_PULSE_MAX := 255     ; Maximum opacity (100%)
global DICTATION_PULSE_STEP := 15     ; Opacity change per tick
global DICTATION_PULSE_INTERVAL := 50 ; Timer interval in ms (smooth animation)

; Get the monitor that contains the active window
; Returns monitor index (1-based) or 0 if not found
GetDictationActiveMonitor() {
    hwnd := WinExist("A")
    if (!hwnd) {
        return 1  ; Default to primary monitor
    }

    rect := Buffer(16, 0)
    if (!DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect)) {
        return 1  ; Default to primary monitor
    }

    left := NumGet(rect, 0, "int")
    top := NumGet(rect, 4, "int")
    right := NumGet(rect, 8, "int")
    bottom := NumGet(rect, 12, "int")

    centerX := left + (right - left) // 2
    centerY := top + (bottom - top) // 2

    monitorCount := MonitorGetCount()
    loop monitorCount {
        MonitorGet(A_Index, &ml, &mt, &mr, &mb)
        if (centerX >= ml && centerX <= mr && centerY >= mt && centerY <= mb) {
            return A_Index
        }
    }

    return 1  ; Default to primary monitor
}

; Outer indicator size: yellow border + inner red square.
GetDictationIndicatorOuterSize() {
    global DICTATION_SQUARE_SIZE, DICTATION_BORDER_PX
    return DICTATION_SQUARE_SIZE + 2 * DICTATION_BORDER_PX
}

; Screen rect for the dictation indicator: prefer top-center inside the active window; else top-center of active monitor work area.
GetDictationIndicatorScreenRect(&outX, &outY, &outW, &outH) {
    global DICTATION_SQUARE_SIZE, DICTATION_BORDER_PX
    outW := GetDictationIndicatorOuterSize()
    outH := outW
    marginTop := 8
    hwnd := WinExist("A")
    if (hwnd) {
        try {
            if (WinGetMinMax("ahk_id " hwnd) = -1)
                hwnd := 0
        } catch {
            hwnd := 0
        }
    }
    if (hwnd) {
        rect := Buffer(16, 0)
        if (DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect)) {
            wl := NumGet(rect, 0, "int")
            wt := NumGet(rect, 4, "int")
            wr := NumGet(rect, 8, "int")
            wb := NumGet(rect, 12, "int")
            winW := wr - wl
            winH := wb - wt
            if (winW >= 8 && winH >= 8) {
                outX := wl + (winW - outW) // 2
                outY := wt + marginTop
                if (outX < wl + 2)
                    outX := wl + 2
                if (outY < wt + 2)
                    outY := wt + 2
                if (outX + outW > wr - 2)
                    outX := wr - outW - 2
                if (outY + outH > wb - 2)
                    outY := wb - outH - 2
                if (outX < wl)
                    outX := wl
                if (outY < wt)
                    outY := wt
                return true
            }
        }
    }
    mon := GetDictationActiveMonitor()
    MonitorGetWorkArea(mon, &ml, &mt, &mr, &mb)
    mw := mr - ml
    outX := ml + (mw - outW) // 2
    outY := mt + 12
    if (outX < ml)
        outX := ml
    if (outY < mt)
        outY := mt
    if (outX + outW > mr)
        outX := mr - outW
    if (outY + outH > mb)
        outY := mb - outH
    return true
}

; Reposition indicator when the active window moves or focus changes (called from pulse timer).
DictationIndicator_SyncPosition() {
    global g_DictationIndicatorGui, g_DictationFollowCache
    if (!IsObject(g_DictationIndicatorGui) || !g_DictationIndicatorGui.Hwnd)
        return
    GetDictationIndicatorScreenRect(&sx, &sy, &sw, &sh)
    key := sx . "," . sy . "," . sw . "," . sh
    if (key = g_DictationFollowCache)
        return
    g_DictationFollowCache := key
    try {
        g_DictationIndicatorGui.Show("NA x" . sx . " y" . sy . " w" . sw . " h" . sh)
    } catch {
    }
}

; Show or update the dictation indicator: yellow border + red inner, anchored to active window (or work area fallback).
ShowDictationIndicator() {
    global g_DictationIndicatorGui, g_DictationPulseOpacity, g_DictationFollowCache
    global DICTATION_SQUARE_SIZE, DICTATION_BORDER_PX, DICTATION_YELLOW_BORDER

    GetDictationIndicatorScreenRect(&squareX, &squareY, &outerW, &outerH)

    ; Check if indicator already exists
    if (IsObject(g_DictationIndicatorGui) && g_DictationIndicatorGui.Hwnd) {
        g_DictationFollowCache := ""
        DictationIndicator_SyncPosition()
        UpdateDictationIndicatorText("")
        return
    }

    ; Create new indicator
    ; +AlwaysOnTop: stays on top of all windows
    ; -Caption: no title bar
    ; +ToolWindow: doesn't appear in taskbar
    ; +E0x20: click-through (WS_EX_TRANSPARENT)
    ; -DPIScale: use raw screen coordinates
    g_DictationIndicatorGui := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x20")
    g_DictationIndicatorGui.Opt("-DPIScale")
    g_DictationIndicatorGui.BackColor := DICTATION_YELLOW_BORDER
    g_DictationIndicatorGui.MarginX := 0
    g_DictationIndicatorGui.MarginY := 0

    ; Inner red fill, then status text on top (transparent over red)
    g_DictationIndicatorGui.Add("Text",
        "x" . DICTATION_BORDER_PX . " y" . DICTATION_BORDER_PX . " w" . DICTATION_SQUARE_SIZE . " h" .
        DICTATION_SQUARE_SIZE .
        " BackgroundFF0000", "")
    g_DictationIndicatorGui.SetFont("s14 cFFFFFF Bold", "Segoe UI")
    g_DictationIndicatorText := g_DictationIndicatorGui.Add("Text",
        "x" . DICTATION_BORDER_PX . " y" . DICTATION_BORDER_PX . " w" . DICTATION_SQUARE_SIZE . " h" .
        DICTATION_SQUARE_SIZE .
        " Center +BackgroundTrans", "")

    ; Reset pulse opacity
    g_DictationPulseOpacity := 128
    g_DictationFollowCache := ""

    ; Show the indicator without activating it
    g_DictationIndicatorGui.Show("NA x" . squareX . " y" . squareY . " w" . outerW . " h" . outerH)
    DictationIndicator_SyncPosition()
    ; Apply initial transparency defensively (GUI may have been destroyed concurrently)
    try {
        if (g_DictationIndicatorGui.Hwnd)
            WinSetTransparent(g_DictationPulseOpacity, g_DictationIndicatorGui)
    } catch {
        ; Ignore "target window not found" or similar errors
    }
}

; Update indicator text (for status messages)
UpdateDictationIndicatorText(message := "") {
    global g_DictationIndicatorGui, g_DictationIndicatorText

    if (IsObject(g_DictationIndicatorGui) && g_DictationIndicatorGui.Hwnd && IsObject(g_DictationIndicatorText)) {
        try {
            g_DictationIndicatorText.Value := message
        } catch {
            ; Ignore errors
        }
    }
}

; Hide and destroy the dictation indicator
HideDictationIndicator() {
    global g_DictationIndicatorGui, g_DictationIndicatorText, g_DictationFollowCache

    g_DictationFollowCache := ""
    if (IsObject(g_DictationIndicatorGui)) {
        try {
            if (g_DictationIndicatorGui.Hwnd) {
                g_DictationIndicatorGui.Destroy()
            }
        } catch {
            ; Ignore errors
        }
        g_DictationIndicatorGui := false
        g_DictationIndicatorText := false
    }
}

; Update the pulse animation (called by timer)
UpdateDictationIndicatorPulse() {
    global g_DictationIndicatorGui, g_DictationPulseOpacity, g_DictationPulseDirection
    global DICTATION_PULSE_MIN, DICTATION_PULSE_MAX, DICTATION_PULSE_STEP

    ; Check if indicator GUI is still valid
    if (!IsObject(g_DictationIndicatorGui) || !g_DictationIndicatorGui.Hwnd) {
        return
    }

    DictationIndicator_SyncPosition()

    ; Update opacity based on direction
    g_DictationPulseOpacity += DICTATION_PULSE_STEP * g_DictationPulseDirection

    ; Reverse direction at bounds
    if (g_DictationPulseOpacity >= DICTATION_PULSE_MAX) {
        g_DictationPulseOpacity := DICTATION_PULSE_MAX
        g_DictationPulseDirection := -1  ; Start fading out
    } else if (g_DictationPulseOpacity <= DICTATION_PULSE_MIN) {
        g_DictationPulseOpacity := DICTATION_PULSE_MIN
        g_DictationPulseDirection := 1   ; Start fading in
    }

    ; Apply new transparency
    try {
        WinSetTransparent(g_DictationPulseOpacity, g_DictationIndicatorGui)
    } catch {
        ; Ignore errors (window might have been destroyed)
    }
}

; Start the pulse animation timer
StartDictationPulseTimer() {
    global g_DictationPulseTimer, DICTATION_PULSE_INTERVAL, g_DictationPulseDirection

    ; Reset pulse direction to fade in
    g_DictationPulseDirection := 1

    ; Stop any existing timer
    StopDictationPulseTimer()

    ; Create and start new timer
    g_DictationPulseTimer := UpdateDictationIndicatorPulse
    SetTimer(g_DictationPulseTimer, DICTATION_PULSE_INTERVAL)
}

; Stop the pulse animation timer
StopDictationPulseTimer() {
    global g_DictationPulseTimer

    if (g_DictationPulseTimer) {
        try {
            SetTimer(g_DictationPulseTimer, 0)
        } catch {
            ; Ignore errors
        }
        g_DictationPulseTimer := false
    }
}

; Audio firewall: Throttle dictation sounds to prevent duplicates
; Enforces a minimum 1000ms gap between sounds regardless of how many times logic fires
SafePlayDictationSound(filePath) {
    Critical  ; Prevents thread interruption - ensures atomic check-and-update sequence
    global g_LastDictationSoundTick, g_DictationStartSound
    static lastStartSoundTick := 0

    ; Special handling for start sound: 7 second cooldown to prevent duplicates
    if (InStr(filePath, "speach-start.wav")) {
        if (A_TickCount - lastStartSoundTick < 7000) {
            return
        }
        lastStartSoundTick := A_TickCount
    } else {
        ; Standard 1 second cooldown for other sounds
        if (A_TickCount - g_LastDictationSoundTick < 1000) {
            return
        }
    }

    ; Update timestamp and play sound (if enabled)
    g_LastDictationSoundTick := A_TickCount
    if (IsSoundEnabled() && FileExist(filePath)) {
        try {
            SoundPlay(filePath)
        } catch {
            ; Silently ignore playback failures (missing file, sync placeholder, format, etc.)
        }
    }
}

; Handler for clipboard changes during dictation completion
DictationClipboardHandler(DataType) {
    ; Remove handler immediately to prevent multiple triggers
    OnClipboardChange(DictationClipboardHandler, 0)

    ; Trigger completion logic immediately
    PlayDictationCompletionChime()
}

; Play completion chime after transcription finishes
PlayDictationCompletionChime(*) {
    global g_DictationCompletionChimeScheduled, g_PendingDictationAction,
        g_KeepIndicatorVisible, g_PendingGeminiPromptAfterDictation

    ; Ensure clipboard handler is removed (safe to call even if already removed)
    try {
        OnClipboardChange(DictationClipboardHandler, 0)
    }

    ; Cancel fallback timer to prevent redundant calls
    SetTimer(PlayDictationCompletionChime, 0)

    ; CRITICAL: Test-and-set pattern - clear flag IMMEDIATELY to prevent duplicates
    ; Use Critical to ensure atomicity
    Critical "On"
    chimeShouldPlay := g_DictationCompletionChimeScheduled
    g_DictationCompletionChimeScheduled := false  ; Clear IMMEDIATELY to prevent other calls
    Critical "Off"

    ; Only play if flag was set (prevent duplicate execution)
    if (chimeShouldPlay) {
        SafePlayDictationSound(g_DictationStopSound)

        ; Execute pending action if one was set (reserved for future use).
        pendingAction := g_PendingDictationAction
        g_PendingDictationAction := ""  ; Clear immediately after reading

        if (pendingAction = "Paste") {
            ; Update indicator text to show status
            UpdateDictationIndicatorText("Pasting...")
            ; Execute paste command
            Send "^v"
            ; Wait for paste to complete before hiding indicator
            Sleep 100  ; Small delay to ensure paste completes
            ; Hide indicator only after paste completes
            HideDictationIndicator()
            g_KeepIndicatorVisible := false
        }

        ; If user stopped dictation with Win+Alt+Shift+0 (no pending action), show Gemini confirm banner (once only).
        Critical "On"
        pendingGemini := g_PendingGeminiPromptAfterDictation
        g_PendingGeminiPromptAfterDictation := false  ; Claim atomically so only one invocation shows the banner
        Critical "Off"
        ; #region agent log
        DebugBannerLog("Utils.ahk:PlayDictationCompletionChime", "Completion chime branch",
            "chimeShouldPlay=1 pendingAction=" . pendingAction . " pendingGemini=" . (pendingGemini ? 1 : 0), "H2")
        ; #endregion
        if (pendingGemini && pendingAction = "") {
            ; #region agent log
            DebugBannerLog("Utils.ahk:PlayDictationCompletionChime", "Calling D2C_FlowManager", "pendingAction empty",
                "H3"
            )
            ; #endregion
            D2C_FlowManager.GetInstance().StartFromDictation()
        }
    }
}

; Called when dictation stop detected: play chime now if clipboard already changed, else wait for change
DictationCompletionChimeOrWaitForClipboard() {
    global g_DictationStartClipboardText
    currentClip := ""
    try {
        currentClip := A_Clipboard
    }
    if (currentClip != g_DictationStartClipboardText) {
        PlayDictationCompletionChime()
    } else {
        OnClipboardChange(DictationClipboardHandler)
        SetTimer(PlayDictationCompletionChime, -1500)
    }
}

CheckDictationRecordingWindow() {
    global g_DictationActive, g_LastStateTransitionTick, g_DictationStartClipboardText
    global g_DictationSoundPlayed, g_DictationCompletionChimeScheduled, g_DictationPulseTimer, g_KeepIndicatorVisible
    ; Check if the "Recording" window exists
    windowExists := false
    try {
        windowExists := WinExist("Recording ahk_exe handy.exe")
    } catch {
        windowExists := false
    }

    ; Handle Start: window exists
    if (windowExists) {
        if (!g_DictationActive) {
            g_DictationActive := true
            g_LastStateTransitionTick := A_TickCount

            ; Capture current clipboard content to detect changes later
            try {
                g_DictationStartClipboardText := A_Clipboard
            } catch {
                g_DictationStartClipboardText := ""
            }

            try {
                micVolumeScript := A_ScriptDir "\scripts\Set-MicVolume.ps1"
                if (FileExist(micVolumeScript)) {
                    micRunStart := A_TickCount
                    Run "powershell.exe -ExecutionPolicy Bypass -File `"" micVolumeScript "`"", , "Hide"
                }
            } catch Error as e {
                ; Silently handle errors - don't interrupt dictation if script fails
            }

            ShowDictationIndicator()
            StartDictationPulseTimer()
        }

        ; Atomic test-and-set: one sound per session when window first detected
        Critical "On"
        if (!g_DictationSoundPlayed) {
            g_DictationSoundPlayed := true
            Critical "Off"
            SafePlayDictationSound(g_DictationStartSound)
        } else {
            Critical "Off"
        }
    }
    ; Handle Stop: window gone and was active
    else if (!windowExists && g_DictationActive) {
        Critical "On"
        if (!g_DictationActive || g_DictationCompletionChimeScheduled) {
            Critical "Off"
            return
        }

        if (g_LastStateTransitionTick && (A_TickCount - g_LastStateTransitionTick < 500)) {
            Critical "Off"
            return
        }

        g_DictationCompletionChimeScheduled := true
        g_LastStateTransitionTick := A_TickCount
        g_DictationActive := false
        Critical "Off"
        g_DictationSoundPlayed := false

        StopDictationPulseTimer()
        HideDictationIndicator()
        DictationCompletionChimeOrWaitForClipboard()
    } else if (g_DictationActive && windowExists) {
        ShowDictationIndicator()
        if (!g_DictationPulseTimer) {
            StartDictationPulseTimer()
        }
    }
}

; Start timer to periodically check Recording window state
StartDictationCheckTimer() {
    global g_DictationCheckTimer

    ; Stop any existing timer
    StopDictationCheckTimer()

    ; Check every 500ms
    g_DictationCheckTimer := CheckDictationRecordingWindow
    SetTimer(g_DictationCheckTimer, 500)
}

; Stop the check timer
StopDictationCheckTimer() {
    global g_DictationCheckTimer

    if (g_DictationCheckTimer) {
        try {
            SetTimer(g_DictationCheckTimer, 0)
        } catch {
            ; Ignore errors
        }
        g_DictationCheckTimer := false
    }
}

; Toggle dictation mode on/off
; The check timer handles everything automatically, this just triggers an immediate check
ToggleDictationMode() {
    ; Trigger immediate check (the timer will handle showing/hiding)
    ; This provides instant detection if window already exists
    CheckDictationRecordingWindow()

    ; OPTIMIZED: Ultra-fast polling for instant window detection and audio feedback
    ; Start with 25ms polling (4x faster than normal) for ultra-responsive detection
    ; This ensures zero-delay audio feedback when handy.exe launches
    SetTimer(CheckDictationRecordingWindow, 25)
    ; Revert to normal 500ms polling after 3 seconds (window should be detected by then)
    SetTimer(RevertDictationPolling, -3000)
}

RevertDictationPolling() {
    SetTimer(CheckDictationRecordingWindow, 500)
}

; Force end dictation immediately (e.g., when Ask action is triggered)
; This immediately removes Esc restriction and hides the indicator
EndDictation() {
    global g_DictationActive, g_DictationSoundPlayed

    g_DictationActive := false
    g_DictationSoundPlayed := false

    StopDictationPulseTimer()
    HideDictationIndicator()
}

; Cleanup dictation indicator resources
CleanupDictationIndicator(*) {
    StopDictationPulseTimer()
    StopDictationCheckTimer()
    HideDictationIndicator()
}

; Register cleanup on script exit
OnExit(CleanupDictationIndicator)

; Toggle dictation mode with Win+Alt+Shift+0
; ~ prefix: key passes through to handy.exe. First press starts dictation, second stops and copies.
; Uses KeyWait + state machine + recursion guard to prevent duplicate triggers (typematic repeats).
~#!+0::
{
    global g_DictationActive, g_LastStateTransitionTick, g_DictationStartSound
    global g_ProgrammaticDictationStop, g_PendingGeminiPromptAfterDictation
    global g_DictationHotkeyIsOwner
    static lastHotkeyTick := 0
    static isProcessing := false

    if (!g_DictationHotkeyIsOwner) {
        return
    }

    ; Skip when script sends #!+0 programmatically
    if (g_ProgrammaticDictationStop) {
        g_ProgrammaticDictationStop := false
        return
    }

    if (isProcessing)
        return

    currentTick := A_TickCount
    if (currentTick - lastHotkeyTick < 200)
        return
    lastHotkeyTick := currentTick
    isProcessing := true
    ; Capture before KeyWait: check timer may clear g_DictationActive when Recording window closes,
    ; so by the time we reach if/else it can be false even when user intended to stop.
    dictationWasActiveOnKeyPress := g_DictationActive

    keyWaitStart := A_TickCount
    KeyWait("0", "L")

    if (!g_DictationActive) {
        g_DictationActive := true
        g_LastStateTransitionTick := A_TickCount
        ShowDictationIndicator()
        StartDictationPulseTimer()
        ; Sound: monitoring loop plays when window detected (zero latency)

        try {
            micVolumeScript := A_ScriptDir "\scripts\Set-MicVolume.ps1"
            if (FileExist(micVolumeScript)) {
                micRunStart := A_TickCount
                Run "powershell.exe -ExecutionPolicy Bypass -File `"" micVolumeScript "`"", , "Hide"
            }
        } catch {
        }
    }

    ; User was stopping dictation (had been active when they pressed key) -> show Gemini confirm after completion
    if (dictationWasActiveOnKeyPress) {
        g_PendingGeminiPromptAfterDictation := true
        g_DictationGeminiConfirmBannerVisible := false  ; Allow 6s banner to show for this cycle (reset from previous N cancel)
        ; #region agent log
        DebugBannerLog("Utils.ahk:~#!+0", "Set pending Gemini flag", "dictationWasActiveOnKeyPress=1", "H1")
        ; #endregion
    } else {
    }

    ToggleDictationMode()
    isProcessing := false
}

; Win+Alt+Shift+7 is defined in Gemini.ahk (TTS from selection: repeat exactly + read aloud).
