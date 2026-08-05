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

RegisterHotstring(trigger, expansion, category := "", title := "", char := "") {
    global g_hotstrings
    g_hotstrings.Push({ trigger: trigger, expansion: expansion, category: category, title: title, char: char })
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

GetPromptDir() {
    return A_ScriptDir "\assets\prompt"
}

; Prompt .txt files are UTF-8; default FileRead uses ANSI on Windows without BOM.
ReadUtf8File(path) {
    return FileRead(path, "UTF-8")
}

GetPromptText(key) {
    try {
        return ReadUtf8File(GetPromptDir() "\" key ".txt")
    } catch {
        return "[PROMPT FILE MISSING: " key "]"
    }
}

; MyNotes technique prompts: live repo path first, then mirror under prompt\technique (synced by aux\Sync-MyNotesTechniquePrompts.ps1).
GetTechniquePromptFilePath(fileName) {
    repo := GetNotesRepoPath()
    dir := (repo != "") ? repo "\studies\technique\prompts" : ""
    if (dir != "" && FileExist(dir "\" fileName))
        return dir "\" fileName
    mirror := A_ScriptDir "\assets\prompt\technique\" fileName
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
            body := ReadUtf8File(GetTechniquePromptFilePath(fileName))
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
    ; Panel registry for #!+U (Win+Alt+Shift+U). IDs are labels only; no typed :o: expanders.
    RegisterHotstring(":o:cgrammar", GetPromptText("grammar"), "Prompts", "✏️ Grammar & Spelling Corrector")
    RegisterHotstring(":o:mtask", GetPromptText("mtask"), "Prompts", "🔲 Convert to Task")
    RegisterHotstring(":o:aiopt", GetPromptText("aiopt"), "Prompts", "🤖 AI Text Optimizer")
    RegisterHotstring(":o:handoff", GetPromptText("handoff-summary"), "Prompts",
    "📝 Summarize for Handoff", "r")
    InitTechniquePromptHotstrings()
    RegisterHotstring(":o:pptdeckcsv", GetPromptText("ppt-content-to-slides-csv"), "Prompts",
    "📊 PPT stage 1: content to slides CSV", "s")
    RegisterHotstring(":o:pptelemcsv", GetPromptText("ppt-slides-to-elements-csv"), "Prompts",
    "🧩 PPT stage 2: slides to elements CSV", "e")
    RegisterHotstring(":o:protscreencsv", GetPromptText("proto-content-to-screens-csv"), "Prompts",
    "📱 Prototype stage 1: content to screens CSV", "w")
    RegisterHotstring(":o:protoelemcsv", GetPromptText("proto-screens-to-elements-csv"), "Prompts",
    "🧩 Prototype stage 2: screens to elements CSV", "f")
    RegisterHotstring(":o:boschimg", GetPromptText("bosch-brand-image"), "Prompts", "🎨 Bosch brand-compliant image")
    RegisterHotstring(":o:csvfill", GetPromptText("unstructured-to-csv"), "Prompts",
    "📋 Fill CSV from unstructured text")
    RegisterHotstring(":o:clipangel", GetPromptText("clipangel-cac"), "Prompts",
    "📎 ClipAngel .cac export", "c")
    RegisterHotstring(":o:howtocsv", GetPromptText("howto-steps-csv"), "Prompts",
    "📝 How-to steps CSV", "h")

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
    RegisterHotstring(":o:gbpm", "BPM", "Projects", "📌 BPM", "m")
    RegisterHotstring(":o:gsrs", "SRS", "Projects", "📡 SRS", "r")
    RegisterHotstring(":o:grise", "RISE", "Projects", "📈 RISE", "g")

    ; Hotstrings (non-workspace "project-like" names)
    RegisterHotstring(":o:myl", "my links", "Hotstrings", "🔗 my links", "m")
    RegisterHotstring(":o:gpm", "project management LA", "Hotstrings", "📋 project management LA", "p")
    RegisterHotstring(":o:guxcip", "UX and CIP", "Hotstrings", "🔗 UX and CIP", "x")
    RegisterHotstring(":o:gtrain", "GS_UX core team_Trainings Management", "Hotstrings", "🎓 Trainings Management", "t"
    )
}
InitHotstringsCheatSheet()