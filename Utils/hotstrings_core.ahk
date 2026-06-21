; =============================================================================
; Utils module: hotstrings_core.ahk
; Hotstrings core (InitHotstringsCheatSheet)
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

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

; Paste file(s) via CF_HDROP (file attachment), not path text. Returns true on success.
InsertFiles(paths) {
    global g_lastExpansion
    if (A_TickCount - g_lastExpansion) < 250
        return false
    if (!paths || paths.Length = 0)
        return false
    for path in paths {
        if !Clipboard_PathIsExistingFile(path)
            return false
    }

    g_lastExpansion := A_TickCount
    saved := ClipboardAll()
    ok := false
    try {
        if !Clipboard_SetFiles(paths)
            return false
        if !Clipboard_WaitForFileDrop(800)
            return false
        Sleep 50
        Send "^v"
        ok := true
        ; Brief settle so browser upload handlers receive the paste before clipboard restore.
        Sleep 400
    } finally {
        try A_Clipboard := saved
    }
    return ok
}

InsertFiles_IsAiChatForeground() {
    try {
        title := WinGetTitle("A")
        if (title = "")
            return false
        lower := StrLower(title)
        return InStr(lower, "gemini") || InStr(lower, "copilot")
    } catch {
        return false
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

:o:pptslide::
{
    InsertText(GetPromptText("slide-creation"))
}

:o:pptslideref::
{
    InsertText(GetPromptText("slide-creation-with-ref"))
}

:o:boschimg::
{
    InsertText(GetPromptText("bosch-brand-image"))
}

:o:csvfill::
{
    InsertText(GetPromptText("unstructured-to-csv"))
}

:o:mdunesc::
{
    UnescapeMarkdownClipboard()
}

; MyNotes technique prompts: live repo path first, then mirror under prompt\technique (synced by aux\Sync-MyNotesTechniquePrompts.ps1).
GetTechniquePromptFilePath(fileName) {
    repo := GetNotesRepoPath()
    dir := (repo != "") ? repo "\studies\technique\prompts" : ""
    if (dir != "" && FileExist(dir "\" fileName))
        return dir "\" fileName
    mirror := A_ScriptDir "\prompt\technique\" fileName
    if FileExist(mirror)
        return mirror
    if (dir != "")
        return dir "\" fileName
    return mirror
}

InitTechniquePromptHotstrings() {
    ; Five files live in MyNotes: studies\technique\prompts (resolved via GetNotesRepoPath in env.ahk).
    defs := [
        ["story-prompt.txt", ":o:mnemonic", "📖 Creating mnemonic stories", "", "Reserved 3"],
        ["video-transcription-prompt.txt", ":o:ytranscript", "🎬 Transcript Youtube Video", "", "Reserved 4"],
        ["story-reduction-prompt.txt", ":o:storyreduction", "📝 Story reduction", "a", "Reserved 5"],
        ["punctual-beast-append-prompt.txt", ":o:punctualbeast", "🧩 Punctual beast append", "p", "Reserved 6"],
        ["image-background-preservation-prompt.txt", ":o:imgpreserve", "🛡️ Preserve background for image generation",
            "g", "Reserved 7"],
    ]
    for row in defs {
        fileName := row[1]
        trigger := row[2]
        title := row[3]
        exChar := row[4]
        reserved := row[5]
        try {
            body := FileRead(GetTechniquePromptFilePath(fileName))
            if (exChar != "")
                RegisterHotstring(trigger, body, "Prompts", title, exChar)
            else
                RegisterHotstring(trigger, body, "Prompts", title)
        } catch {
            RegisterHotstring("", "", "Prompts", reserved)
        }
    }
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
    InitTechniquePromptHotstrings()
    try {
        aibRapidFireTpl := FileRead(promptDir "\aib-rapid-fire-template.txt")
        RegisterHotstring(":o:aibrapid", aibRapidFireTpl, "Prompts", "📜 Junior AI: ⚡ rapid-fire template")
    } catch {
        RegisterHotstring(":o:aibrapid",
            "Junior AI (AIB): planning doc with ⚡ - conceptual above, execution steps below.`n", "Prompts",
            "📜 Junior AI: ⚡ rapid-fire template")
    }
    try {
        RegisterHotstring(":o:pptslide", FileRead(promptDir "\slide-creation.txt"), "Prompts",
        "📊 Create PowerPoint slide")
    } catch {
        RegisterHotstring(":o:pptslide", "Create one PowerPoint slide as an image.`n", "Prompts",
            "📊 Create PowerPoint slide")
    }
    try {
        RegisterHotstring(":o:pptslideref", FileRead(promptDir "\slide-creation-with-ref.txt"), "Prompts",
        "📊 Create PowerPoint slide (reference)")
    } catch {
        RegisterHotstring(":o:pptslideref",
            "Create one PowerPoint slide as an image using the attached reference as the main visual guide.`n",
            "Prompts",
            "📊 Create PowerPoint slide (reference)")
    }
    try {
        RegisterHotstring(":o:boschimg", FileRead(promptDir "\bosch-brand-image.txt"), "Prompts",
        "🎨 Bosch brand-compliant image")
    } catch {
        RegisterHotstring(":o:boschimg",
            "Generate one Bosch Brand Guide and BDDS compliant image.`n", "Prompts",
            "🎨 Bosch brand-compliant image")
    }
    try {
        RegisterHotstring(":o:csvfill", FileRead(promptDir "\unstructured-to-csv.txt"), "Prompts",
        "📋 Fill CSV from unstructured text")
    } catch {
        RegisterHotstring(":o:csvfill",
            "Extract information from unstructured text and fill/update CSV rows using the provided column schema.`n",
            "Prompts", "📋 Fill CSV from unstructured text")
    }

    ; Hotstrings: emails
    RegisterHotstring(":o:ebosch", "eduardo.figueiredo@br.bosch.com", "Hotstrings", "💼 Bosch Email")
    RegisterHotstring(":o:egoogle", "edu.evangelista.figueiredo@gmail.com", "Hotstrings", "📧 Gmail")

    ; Projects (Cursor workspaces) - keys align with Project Selector 2
    RegisterHotstring(":o:gintegra", "GS_UX core team_UX and CIP Integration", "Projects", "🔄 UX and CIP Integration",
        "u")
    RegisterHotstring(":o:gdash", "GS_E&S_CIP Dashboard research and design", "Projects", "📊 CIP Dashboard", "d")
    RegisterHotstring(":o:boiler-plate", "boiler-plate", "Projects", "🧱 boiler-plate", "0")
    RegisterHotstring(":o:astra", "astra", "Projects", "⭐ astrA", "a")
    RegisterHotstring(":o:opex-cim-journey-mapping", "opex-cim-journey-mapping", "Projects",
        "E&S Opex CIM Journey Mapping",
        "o")
    RegisterHotstring(":o:gpilotb2b", "Piloto PT B2B", "Projects", "🧪 Piloto PT B2B", "b")
    RegisterHotstring(":o:gpython", "17 - Python Scripts", "Projects", "🐍 Python Scripts", "t")

    ; Hotstrings (non-workspace "project-like" names)
    RegisterHotstring(":o:myl", "my links", "Hotstrings", "🔗 my links", "m")
    RegisterHotstring(":o:gpm", "project management LA", "Hotstrings", "📋 project management LA", "p")
    RegisterHotstring(":o:guxcip", "UX and CIP", "Hotstrings", "🔗 UX and CIP", "x")
    RegisterHotstring(":o:gtrain", "GS_UX core team_Trainings Management", "Hotstrings", "🎓 Trainings Management", "t"
    )
}
InitHotstringsCheatSheet()

