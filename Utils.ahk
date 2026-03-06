#Requires AutoHotkey v2.0+
#SingleInstance Force
#include %A_ScriptDir%\env.ahk

; #region agent log
DebugBannerLog(location, message, dataStr := "", hypothesisId := "") {
    logPath := A_ScriptDir "\debug-5ecf82.log"
    q := Chr(34)
    line := "{" q "sessionId" q ":" q "5ecf82" q "," q "location" q ":" q location q "," q "message" q ":" q message q "," q "timestamp" q ":" A_TickCount
    if (dataStr != "")
        line .= "," q "data" q ":" q dataStr q ""
    if (hypothesisId != "")
        line .= "," q "hypothesisId" q ":" q hypothesisId q ""
    line .= "}"
    try FileAppend line "`n", logPath
}
; #endregion

#include UIA-v2\Lib\UIA.ahk
#include UIA-v2\Lib\UIA_Browser.ahk

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
    InsertText("My Links")
}

:o:gintegra::
{
    InsertText("GS_UX core team_UX and CIP Integration")
}

:o:gdash::
{
    InsertText("GS_E&S_CIP Dashboard research and design")
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
    InsertText("GS_UX_Project_Management_Activities_LA")
}

:o:guxcip::
{
    InsertText("GS_UX_and_CIP")
}

:o:gtrain::
{
    InsertText("GS_UX core team_Trainings Management")
}

:o:gbp::
{
    InsertText("GS_B2C_Portals and Key Accounts Process POC")
}

:o:cgrammar::
{
    InsertText(
        "Correct grammar and spelling. Remove any dashes from the text. The text should be plain with no styles. Give back only the text. Use lininebreaks and a space betetween paragraphs and look like a human."
    )
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
    InsertText(
        "This is a message, summary, text or any textual information that translates into a task for me to do. Translate this way, into a task, make informative and start with the emoji 🔲. Make it clear and consise."
    )
}

:o:flog::
{
    InsertText(
        "Food_Log dictation → Excel CSV`n`nROLE`nYou transcribe my meal dictation (PT/EN) into rows for my Excel Food_Log.`n`nHOW IT WORKS`n- I will dictate one or more meals in free speech.`n- Process immediately without asking questions.`n`nOUTPUT FORMAT (STRICT)`n- Output ONLY the final CSV block. NO bullet points, NO analysis, NO extra text.`n- Plain text CSV. NO markdown, NO header, NO commentary.`n- Output a single block of text. NO empty lines between rows.`n- Separator: semicolon (;)`n`nEXAMPLE OUTPUT`n2025-10-25;Breakfast;08:30;coffee;caffeine;0;`n2025-10-25;Lunch;12:15;rice, beans;protein;0;`n`nOUTPUT RULES`n- STRICTLY CONTINUOUS LINES. Do not insert empty lines between rows.`n- Do not group by meal. List all sequentially in one block.`n- Use single newlines (\n) only. No paragraph breaks.`n- Trim whitespace from each row.`n- Each row format: Date;Meal;Time;Main_Items;Tags;Satisfaction_with_Speech;Notes`n- Sort by Date then Time.`n`nFIELD RULES`n- Date: YYYY‑MM‑DD. Use " "today" " for current date in America/Sao_Paulo timezone.`n  (IMPORTANT: For the date always consider one day before the current one, unless specified otherwise.)`n- Time: HH:MM in 24h; pad leading zeros (e.g., 08:05).`n- Meal: Breakfast | Lunch | Dinner | Snack.`n  PT mapping: café da manhã→Breakfast; almoço→Lunch; jantar/janta→Dinner; lanche→Snack.`n- Main_Items: comma‑separated simple item names (e.g., coffee, bread, butter).`n- Tags: comma‑separated, from this set when present or inferable:`n  caffeine, sugar, alcohol, dairy, gluten, fried, spicy, high-carb, low-carb, processed, protein, fiber, late-night, home-cooked, fast-food.`n  Add " "late-night" " automatically if Time ≥ 22:00.`n- Satisfaction_with_Speech: integer 0–3 (0 = liked a lot; 3 = disliked a lot). If not stated, leave empty.`n- Notes: short free text when I provide context.`n`nMISSING INFO`n- If Date or Time is missing, use " "today" " and infer time from meal type (Breakfast=08:00, Lunch=12:00, Dinner=19:00, Snack=15:00).`n`nACK`n- Process the dictation immediately and output CSV rows only.)"
    )
}

:o:aiopt::
{
    InsertText(
        "Task: Rewrite the input text so it becomes AI-oriented.`n`nGoal: Produce a version that is concise, unambiguous, free of redundancy, and easy for an AI to parse.`n`nContext: You are Gemini, the senior AI. The input text contains instructions and requests intended for a junior AI running in Cursor (AIB). You must preserve ALL important information, especially any instructions, requests, or requirements meant for this junior AI. Do not remove information in the name of clarity or conciseness.`n`nInstructions:`n`n1. Preserve all essential information, especially instructions and requests for the junior AI in Cursor.`n2. Use positive, direct instructions.`n3. Maintain consistent terminology and simple syntax.`n4. Resolve ambiguity and clarify references.`n5. Output in a clean, structured format with no extra commentary.`n6. Do NOT omit any important information, requests, or instructions that the user provided for the junior AI in Cursor.`n7. Do NOT bias the prompt with your own concerns, interpretations, or modifications. Process the text as-is without adding your own perspective or concerns.`n`nInput: <insert text here>`n`nOutput:`nCRITICAL: The output must contain ONLY the processed result with no additional text, commentary, explanations, or formatting. The output must be ready for direct copy and paste without any modification.`n`nA rewritten version of the input text that is optimized for AI interpretation and contains:`n`n* Clear meaning`n* No repeated ideas`n* No filler wording`n* No contradictions`n* Stable terminology`n* Straightforward sentence structure`n* ALL important information preserved, especially instructions for the junior AI in Cursor)"
    )
}

:o:cplant::
{
    InsertText(
        "---`nname: [Title Case Name of the Plan]`noverview: [A concise, 1-2 sentence summary of the high-level objective.]`ntodos:`n  - id: [unique_string_id]`n    content: [Specific, actionable step]`n    status: pending`n    dependencies: [] # Optional: list IDs of prerequisite steps`n---"
    )
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
        RegisterHotstring(":o:cgrammar", "Correct grammar and spelling. Give back only the text.`n", "Prompts",
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

    ; General Information (2 items) - Second category
    RegisterHotstring(":o:ebosch", "eduardo.figueiredo@br.bosch.com", "General", "💼 Bosch Email")
    RegisterHotstring(":o:egoogle", "edu.evangelista.figueiredo@gmail.com", "General", "📧 Gmail")

    ; Project Names (Projects category)
    ; Order here defines character assignment within the Projects segment of g_HotstringCharSequence.
    ; First three keep their current slots, then we introduce "14-my-notes" at the former X slot,
    ; and drop the previous X/C/7 entries so later projects shift left by one as needed.
    RegisterHotstring(":o:myl", "My Links", "Projects", "🔗 My Links")
    RegisterHotstring(":o:gintegra", "GS_UX core team_UX and CIP Integration", "Projects", "🔄 UX and CIP Integration")
    RegisterHotstring(":o:gdash", "GS_E&S_CIP Dashboard research and design", "Projects", "📊 CIP Dashboard")
    ; New project for the X slot in the selector
    RegisterHotstring(":o:14notes", "14-my-notes", "Projects", "📝 14-my-notes")
    ; Remaining projects (previously after X/C/7) shift left one character each
    RegisterHotstring(":o:gpm", "GS_UX_Project_Management_Activities_LA", "Projects", "📋 Project Management LA", "c")
    RegisterHotstring(":o:guxcip", "GS_UX_and_CIP", "Projects", "🔗 UX and CIP")
    RegisterHotstring(":o:gtrain", "GS_UX core team_Trainings Management", "Projects", "🎓 Trainings Management")
    RegisterHotstring(":o:26ai", "26-ai-experiment", "Projects", "🤖 26-ai-experiment")
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
global g_DictationLoopActive := false
global g_ProgrammaticDictationStop := false  ; Skip ~#!+0 handler when script sends #!+0 (loop cycle)

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

; Check which scripts have uncommitted changes or need updates (returns array of script names)
CheckScriptsNeedingUpdates() {
    scriptsDir := GetScriptsDirectory()
    scriptsNeedingUpdates := []

    try {
        ; Change to scripts directory
        SetWorkingDir(scriptsDir)

        ; Check git status for modified files using temp file
        tempFile := A_Temp "\git_status_" A_TickCount ".txt"
        try {
            RunWait('git status --porcelain > "' tempFile '"', scriptsDir, "Hide")
            if (FileExist(tempFile)) {
                gitOutput := FileRead(tempFile)
                FileDelete(tempFile)

                if (gitOutput && StrLen(gitOutput) > 0) {
                    ; Parse git status output to find modified .ahk files
                    lines := StrSplit(gitOutput, "`n")
                    for line in lines {
                        if (InStr(line, ".ahk")) {
                            ; Extract filename from git status line (format: " M filename" or "M  filename")
                            parts := StrSplit(Trim(line), " ")
                            if (parts.Length > 0) {
                                fileName := parts[parts.Length]
                                ; Extract just the script name without path
                                if (InStr(fileName, "\")) {
                                    parts := StrSplit(fileName, "\")
                                    fileName := parts[parts.Length]
                                }
                                scriptsNeedingUpdates.Push(fileName)
                            }
                        }
                    }
                }
            }
        } catch {
            ; Ignore temp file errors
        }

        ; Also check if there are remote updates available
        RunWait("git fetch", scriptsDir, "Hide")
        tempStatusFile := A_Temp "\git_status_sb_" A_TickCount ".txt"
        try {
            RunWait('git status -sb > "' tempStatusFile '"', scriptsDir, "Hide")
            if (FileExist(tempStatusFile)) {
                gitStatusOutput := FileRead(tempStatusFile)
                FileDelete(tempStatusFile)
                if (gitStatusOutput && InStr(gitStatusOutput, "behind")) {
                    ; There are remote updates, add all scripts to the list
                    allScripts := GetScriptFiles()
                    for scriptPath in allScripts {
                        fileName := ""
                        if (InStr(scriptPath, "\")) {
                            parts := StrSplit(scriptPath, "\")
                            fileName := parts[parts.Length]
                        } else {
                            fileName := scriptPath
                        }
                        ; Only add if not already in list
                        if (!scriptsNeedingUpdates.Has(fileName)) {
                            scriptsNeedingUpdates.Push(fileName)
                        }
                    }
                }
            }
        } catch {
            ; Ignore temp file errors
        }
    } catch Error as e {
        ; If git commands fail, return empty array (scripts may not be in git repo)
    }

    return scriptsNeedingUpdates
}

; Quick Update Scripts macro (four layers: Git, sequential reload with Utils last, delay, PowerShell verification).
; Final notification (banner + sound) runs only after all layers and verification succeed. Act.ahk is NOT updated (entry point).
QuickUpdateScripts() {
    scriptsDir := GetScriptsDirectory()
    files := GetScriptFiles()
    failedScripts := []
    deferredUtils := false
    utilsPath := ""

    ; Layer 1: Git synchronization
    try {
        SetWorkingDir(scriptsDir)
        RunWait("git fetch", scriptsDir, "Hide")
        pullResult := RunWait("git pull", scriptsDir, "Hide")
        if (pullResult != 0) {
            ShowCenteredOverlay_Utils("Git pull failed. Proceeding with local reload...", 2000, "FFFF00")
        }
    } catch Error as e {
        ShowCenteredOverlay_Utils("Git update failed: " e.Message, 2000, "FF0000")
    }

    ; Quality check 1: After Git, verify all script files exist (pre-flight so we know what we can run)
    missingPre := []
    for file in files {
        if (!FileExist(file)) {
            parts := StrSplit(file, "\")
            missingPre.Push(parts[parts.Length])
        }
    }
    if (missingPre.Length > 0) {
        list := ""
        for n in missingPre
            list .= n "`n"
        ShowCenteredOverlay_Utils("QC1: Missing after pull:`n" list, 3000, "FF6600")
    }

    ; Layer 2: Sequential script reload; Utils.ahk is deferred (run after verification + notification).
    for index, file in files {
        isUtils := InStr(file, "Utils.ahk")

        if (isUtils) {
            deferredUtils := true
            utilsPath := file
        }

        ; Quality check 2: Before Run - file must exist and be non-empty
        parts := StrSplit(file, "\")
        fileName := parts[parts.Length]
        if (!FileExist(file)) {
            failedScripts.Push(fileName " (not found)")
            continue
        }
        try {
            if (FileGetSize(file) = 0) {
                failedScripts.Push(fileName " (empty file)")
                continue
            }
        } catch {
            failedScripts.Push(fileName " (unreadable)")
            continue
        }

        ; Skip launching Utils.ahk here; it runs after verification and notification
        if (isUtils)
            continue

        ; Reload the script
        try {
            pid := Run(file)
            ; Quality check 3: Verify process started (Run returns PID or 0 on failure)
            if (pid = 0) {
                failedScripts.Push(fileName " (process did not start)")
            }
            Sleep 300
        } catch Error as e {
            failedScripts.Push(fileName " (" e.Message ")")
        }
    }

    ; Brief delay before verification so all processes and filesystem are settled
    Sleep 1500

    ; Layer 3: PowerShell verification - confirm every target script exists, is non-empty, and readable
    pathsFile := A_Temp "\quick-update-paths_" A_TickCount ".txt"
    reportFile := A_Temp "\quick-update-verify-report_" A_TickCount ".txt"
    try {
        pathList := ""
        for file in files
            pathList .= file "`n"
        if (FileExist(pathsFile))
            FileDelete(pathsFile)
        FileAppend(pathList, pathsFile)
    } catch {
        failedScripts.Push("Verify (could not write paths file)")
    }

    verifyExitCode := 0
    if (FileExist(pathsFile)) {
        ; Resolve verification script: first next to Utils.ahk (A_LineFile), then scriptsDir fallback (covers Act.ahk parent launch and direct run)
        utilsDir := SubStr(A_LineFile, 1, InStr(A_LineFile, "\", false, -1) - 1)
        verifyScript := utilsDir "\Verify-ScriptUpdate.ps1"
        if (!FileExist(verifyScript))
            verifyScript := scriptsDir "\Verify-ScriptUpdate.ps1"
        if (FileExist(verifyScript)) {
            verifyExitCode := RunWait('powershell.exe -NoProfile -ExecutionPolicy Bypass -File "' verifyScript '" -ScriptsDir "' scriptsDir '" -PathsFile "' pathsFile '" -ReportFile "' reportFile '"',
                scriptsDir, "Hide")
            try {
                if (FileExist(reportFile)) {
                    verifyReport := FileRead(reportFile)
                    FileDelete(reportFile)
                    lines := StrSplit(Trim(verifyReport), "`n")
                    for line in lines {
                        if (StrLen(Trim(line)) > 0)
                            failedScripts.Push("Verify: " Trim(line))
                    }
                }
            } catch {
                ; Ignore cleanup errors
            }
        } else {
            failedScripts.Push("Verify: Verify-ScriptUpdate.ps1 not found")
        }
        try {
            FileDelete(pathsFile)
        } catch {
        }
    }

    ; Final notification only after all layers (Git, reload, delay, verification) have run
    if (failedScripts.Length > 0) {
        failedList := ""
        for script in failedScripts {
            failedList .= script "`n"
        }
        ShowCenteredOverlay_Utils("Some scripts failed to update:`n" failedList, 4000, "FF0000")
        try {
            if (FileExist(scriptsDir "\sounds\quick-update-failure.wav"))
                SoundPlay(scriptsDir "\sounds\quick-update-failure.wav")
        } catch {
        }
    } else {
        ShowCenteredOverlay_Utils("All scripts updated successfully!", 3500, "00FF00")
        try {
            if (FileExist(scriptsDir "\sounds\quick-update-success.wav"))
                SoundPlay(scriptsDir "\sounds\quick-update-success.wav")
        } catch {
        }
        ; Wait for overlay to be readable before reloading Utils (otherwise reload can replace process and hide it)
        Sleep 5500
    }

    ; Reload Utils.ahk last so this instance is replaced only after notification
    if (deferredUtils && utilsPath != "") {
        try {
            Run(utilsPath)
        } catch {
        }
    }
}

; Update Gemini script specifically, checking other scripts first
UpdateGeminiScript() {
    scriptsDir := GetScriptsDirectory()
    geminiPath := scriptsDir "\Gemini.ahk"

    ; Check which scripts need updates
    scriptsNeedingUpdates := CheckScriptsNeedingUpdates()

    ; Filter out Gemini.ahk from the list
    otherScriptsNeedingUpdates := []
    for script in scriptsNeedingUpdates {
        if (script != "Gemini.ahk") {
            otherScriptsNeedingUpdates.Push(script)
        }
    }

    ; Notify if other scripts need updates
    if (otherScriptsNeedingUpdates.Length > 0) {
        otherScriptsList := ""
        for script in otherScriptsNeedingUpdates {
            otherScriptsList .= script "`n"
        }
        response := MsgBox("Other scripts need updates before Gemini:`n`n" otherScriptsList "`nUpdate all scripts now, or update only Gemini?",
            "Script Updates Available", "YesNo")
        if (response = "Yes") {
            ; Update all scripts
            QuickUpdateScripts()
            return
        }
    }

    ; Update only Gemini
    try {
        ; Pull latest changes first
        SetWorkingDir(scriptsDir)
        RunWait("git fetch", scriptsDir, "Hide")
        RunWait("git pull", scriptsDir, "Hide")

        ; Run Gemini script
        if (!FileExist(geminiPath)) {
            MsgBox "Gemini.ahk not found at: " geminiPath, "Update Failed", "IconX"
            return
        }
        Run geminiPath
        ShowCenteredOverlay_Utils("Gemini script updated!", 1500, "00FF00")
    } catch Error as e {
        MsgBox "Failed to update Gemini script: " e.Message, "Update Failed", "IconX"
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
            ShowCenteredOverlay_Utils("Error: Target window not found.", 2000)
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
}

; Extract title from first non-favorite clip in ClipAngel
MergeNonFavoriteClips() {
    try {
        ; Show persistent banner for the duration of the algorithm
        AiModelBanner_Show("Merging non-favorite clips...", "FFCC00")

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
            ShowCenteredOverlay_Utils("ClipAngel window not found.", 2000)
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
                        ShowCenteredOverlay_Utils("Merged non-favorite clips (copied)", 2000, "FFCC00")
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
            ShowCenteredOverlay_Utils("Favorite clip not found in first " . maxIterations . " rows", 2000)
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
            ShowCenteredOverlay_Utils("ClipAngel window not found.", 2000)
            return
        }
        WinWaitActive("ClipAngel", , 2)
    } else {
        needBanner := true
        ClipAngelBanner_Show("Opening Clip Angel...", "3772FF")
        Send "!v"
        if !WinWait("ClipAngel", , 10) {
            ClipAngelBanner_Hide()
            return
        }
        try {
            WinActivate("ClipAngel")
        } catch {
            ClipAngelBanner_Hide()
            ShowCenteredOverlay_Utils("ClipAngel window not found.", 2000)
            return
        }
        WinWaitActive("ClipAngel", , 2)
    }
    Sleep 50
    hwnd := WinExist("A")
    el := UIA.ElementFromHandle(hwnd)
    if !el {
        if needBanner
            ClipAngelBanner_Hide()
        return
    }
    try {
        dataGrid := el.FindFirst({ Type: 50036, AutomationId: "dataGridView" })
        if !dataGrid {
            if needBanner
                ClipAngelBanner_Hide()
            return
        }
        row0 := dataGrid.FindFirst({ Type: 50025, Name: "Row 0" })
        if !row0 {
            if needBanner
                ClipAngelBanner_Hide()
            return
        }
        hasSel := row0.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable)
        isSelected := hasSel && row0.SelectionItemPattern.IsSelected
        if (!isSelected) {
            if !needBanner
                ClipAngelBanner_Show("Focusing Row 0...", "3772FF")
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
        ClipAngelBanner_Show("Done", "27AE60")
        SetTimer(ClipAngelBanner_Hide, -500)
    }
}

; =============================================================================
; Clip Angel: Mark Last Clip as Favorite
; =============================================================================
; Open Clip Angel, mark current (last) clip as favorite, then close.
MarkLastClipAsFavorite() {
    ; Step 1: Open Clip Angel
    Send "!v"
    Sleep 600

    ; Step 2: Mark current clip as favorite
    Send "!q"
    Sleep 600

    ; Step 3: Close Clip Angel
    Send "!v"
}

; =============================================================================
; AI Model Selection System for Handy
; =============================================================================
; Configuration: Maps selection numbers (1–7) to AI model names.
; These are partial name prefixes used to find buttons in the UIA tree (Type 50000, botão).
; Descriptions match Handy Transcription Models UI for quick verification.
global g_HandyAiModels := Map(
    1, { name: "Whisper Turbo", desc: "Balanced accuracy and speed. Multi-language." },
    2, { name: "Whisper Small", desc: "Fast and fairly accurate. Multi-language, translate to English." },
    3, { name: "Whisper Medium", desc: "Good accuracy, medium speed. Multi-language, translate to English." },
    4, { name: "Whisper Large", desc: "Good accuracy, but slow. Multi-language, translate to English." },
    5, { name: "Parakeet V3", desc: "Fast and accurate. Multi-language." },
    6, { name: "Parakeet V2", desc: "English only. Best model for English speakers." },
    7, { name: "Moonshine Base", desc: "Very fast, English only. Handles accents well." }
)

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
    g_AiModelSelectorGui.Add("Text", "w280 Center", "Press 1–7 | Esc to cancel")

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

    ; Enable hotkeys for 1–6 and Escape
    Hotkey("1", AiModelSelector_HandleKey, "On")
    Hotkey("2", AiModelSelector_HandleKey, "On")
    Hotkey("3", AiModelSelector_HandleKey, "On")
    Hotkey("4", AiModelSelector_HandleKey, "On")
    Hotkey("5", AiModelSelector_HandleKey, "On")
    Hotkey("6", AiModelSelector_HandleKey, "On")
    Hotkey("7", AiModelSelector_HandleKey, "On")
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
    try Hotkey("5", "Off")
    try Hotkey("6", "Off")
    try Hotkey("7", "Off")
    try Hotkey("Escape", AiModelSelector_Cancel, "Off")

    ; Destroy GUI
    if (IsObject(g_AiModelSelectorGui) && g_AiModelSelectorGui.Hwnd) {
        try g_AiModelSelectorGui.Destroy()
    }
    g_AiModelSelectorGui := false
}

; =============================================================================
; Status Banner Functions (non-blocking; use standard loading indicator)
; =============================================================================
AiModelBanner_Show(text, bgColor := "3772FF") {
    centerOnHwnd := 0
    try centerOnHwnd := WinGetID("A")
    catch {
    }
    if (!centerOnHwnd || !WinExist("ahk_id " centerOnHwnd))
        centerOnHwnd := 0
    StandardLoadingBar_Show(text, bgColor, { passive: true, centerOnHwnd: centerOnHwnd, textWidth: 450, fontSize: 18,
        passiveBgColor: bgColor, alpha: 200 })
}

AiModelBanner_Hide() {
    StandardLoadingBar_Hide(0)
}

; Small banner for Clip Angel (uses standard loading indicator).
ClipAngelBanner_Show(text, bgColor := "3772FF") {
    centerOnHwnd := 0
    try centerOnHwnd := WinGetID("A")
    catch {
    }
    if (!centerOnHwnd || !WinExist("ahk_id " centerOnHwnd))
        centerOnHwnd := 0
    StandardLoadingBar_Show(text, bgColor, { passive: true, centerOnHwnd: centerOnHwnd, textWidth: 200, fontSize: 10,
        passiveBgColor: bgColor, alpha: 220 })
}

ClipAngelBanner_Hide() {
    StandardLoadingBar_Hide(0)
}

; =============================================================================
; Single-character tab banner (uses standard loading indicator). tabNumber 1 = blue, 2 = yellow. Auto-hides after 700 ms.
; =============================================================================
ShowSingleCharTabBanner_Utils(tabNumber) {
    msg := String(tabNumber)
    bgColor := (tabNumber = 1) ? "0000FF" : "FFFF00"
    centerOnHwnd := 0
    try centerOnHwnd := WinGetID("A")
    catch {
    }
    if (!centerOnHwnd || !WinExist("ahk_id " centerOnHwnd))
        centerOnHwnd := 0
    StandardLoadingBar_Show(msg, bgColor, { passive: true, centerOnHwnd: centerOnHwnd, textWidth: 120, fontSize: 72,
        passiveBgColor: bgColor, alpha: 178 })
    StandardLoadingBar_Hide(700)
}

; =============================================================================
; ExecuteHandyAiModelSelection() - Main automation logic for Handy
; =============================================================================
ExecuteHandyAiModelSelection(selection) {
    global g_HandyAiModels

    modelInfo := g_HandyAiModels[selection]
    modelName := modelInfo.name

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
        AiModelBanner_Show("🎯 Selecting " . modelName . "...")
        if (!Handy_ClickAiModel(handyHwnd, modelName)) {
            AiModelBanner_Show("❌ Model not found: " . modelName, "E74C3C")
            Sleep 2000
            AiModelBanner_Hide()
            return
        }

        ; Step 4: Wait for model to finish loading (poll button name until "loading" disappears)
        AiModelBanner_Show("⏳ Waiting for model...", "27AE60")
        Handy_WaitForModelReady(handyHwnd, 20000)

        ; Step 4.5: Play confirmation sound when model is ready
        if (IsSoundEnabled()) {
            soundPath := A_ScriptDir . "\sounds\handy-model-chosen.mp3"
            if (FileExist(soundPath)) {
                try SoundPlay(soundPath)
            }
        }

        ; Step 5: Close Handy window
        AiModelBanner_Show("✅ Done! Closing Handy...", "27AE60")
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
; Helper: EnumWindows callback to find Teams windows
; =============================================================================
; Global variables for EnumWindows callback
global g_EnumTeamsPID := 0
global g_EnumTeamsWindows := []

EnumWindowsCallback(hwnd, lParam) {
    global g_EnumTeamsPID, g_EnumTeamsWindows
    try {
        ; Get window's process ID
        winPID := 0
        DllCall("GetWindowThreadProcessId", "Ptr", hwnd, "UInt*", &winPID)

        ; Check if this window belongs to Teams process
        if (winPID = g_EnumTeamsPID) {
            ; Only add main windows (not child windows)
            ; Check if it's a top-level window
            parent := DllCall("GetParent", "Ptr", hwnd, "Ptr")
            if (parent = 0) {
                ; It's a top-level window, add it
                g_EnumTeamsWindows.Push(hwnd)
            }
        }
    } catch {
        ; Ignore errors for inaccessible windows
    }
    return true  ; Continue enumeration
}

; =============================================================================
; Helper: Show centered overlay banner (uses standard loading indicator; non-blocking).
; =============================================================================
ShowCenteredOverlay_Utils(text, duration := 1500, bgColor := "3772FF") {
    centerOnHwnd := WinGetID("A")
    if (!centerOnHwnd || !WinExist("ahk_id " centerOnHwnd))
        centerOnHwnd := 0
    StandardLoadingBar_Show(text, bgColor, { passive: true, centerOnHwnd: centerOnHwnd, textWidth: 500, fontSize: 24,
        passiveBgColor: bgColor })
    StandardLoadingBar_Hide(duration)
}

; =============================================================================
; Standard loading bar (monitor-aware, show/update/hide lifecycle)
; Use for long-running shortcuts; replace ad-hoc banners/overlays with this.
; Supports passive (text-only) mode and ShowWithKeys for letter-keystroke commands.
; =============================================================================
global g_StandardLoadingBarGui := 0
global g_StandardLoadingBarValue := 0
global g_StandardLoadingBarIsKeysOverlay := false
global g_StandardLoadingBarKeysHotkeys := []
global g_StandardLoadingBarKeysTimeoutTimer := ""
global g_StandardLoadingBarBorderGui := 0

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

StandardLoadingBar_Show(state := "Working...", barColor := "3772FF", options := "") {
    global g_StandardLoadingBarGui, g_StandardLoadingBarValue, g_StandardLoadingBarIsKeysOverlay,
        g_StandardLoadingBarBorderGui
    try StandardLoadingBar_CloseKeysOverlay()
    try StandardLoadingBar_Hide(0)
    passive := options && options.HasProp("passive") && options.passive
    centerOnHwnd := options && options.HasProp("centerOnHwnd") ? options.centerOnHwnd : 0
    textWidth := options && options.HasProp("textWidth") ? options.textWidth : 0
    fontSize := options && options.HasProp("fontSize") ? options.fontSize : 9
    alpha := options && options.HasProp("alpha") ? options.alpha : 235
    passiveBgColor := options && options.HasProp("passiveBgColor") ? options.passiveBgColor : ""
    noBorder := options && options.HasProp("noBorder") ? options.noBorder : false

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
    overlayGui.BackColor := (passive && passiveBgColor != "") ? passiveBgColor : "1E1E2E"
    overlayGui.MarginX := 16
    overlayGui.MarginY := 10
    overlayGui.SetFont("s" . fontSize . " cFFFFFF", "Segoe UI")
    overlayGui.Add("Text", "w" . barWidth . (passive ? " Wrap" : ""), state)
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

    ; Create yellow border frame behind the overlay for visibility (optional; skip when noBorder to show a single banner).
    if (!noBorder) {
        borderWidth := 6
        try {
            if IsObject(g_StandardLoadingBarBorderGui)
                g_StandardLoadingBarBorderGui.Destroy()
        } catch {
        }
        borderGui := Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale")
        borderGui.BackColor := "FFFF00"
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
    if (delayMs > 0) {
        SetTimer(() => StandardLoadingBar_Hide(0), -delayMs)
        return
    }
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
; passiveBgColor: optional; when set, used as overlay background (e.g. "FFFF00" for yellow, no blue).
; noBorder: when true, do not create the yellow border (single banner only).
StandardLoadingBar_ShowWithKeys(state, keyCallbacks, timeoutMs := 0, centerOnHwnd := 0, timeoutCallback := "", barColor :=
    "3772FF", textWidth := 500, fontSize := 9, passiveBgColor := "", noBorder := false) {
    global g_StandardLoadingBarIsKeysOverlay, g_StandardLoadingBarKeysHotkeys, g_StandardLoadingBarKeysTimeoutTimer
    opts := { passive: true, centerOnHwnd: centerOnHwnd, textWidth: textWidth, fontSize: fontSize }
    if (passiveBgColor != "")
        opts.passiveBgColor := passiveBgColor
    if (noBorder)
        opts.noBorder := true
    StandardLoadingBar_Show(state, barColor, opts)
    g_StandardLoadingBarIsKeysOverlay := true
    g_StandardLoadingBarKeysHotkeys := []

    ; Register primary and case-variant keys directly from the keyCallbacks map.
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
    ; Bind key first (string), then cb (BoundFunc), to avoid "Invalid base" when cb is first arg
    fn := StandardLoadingBar_KeyWrapper.Bind(key, cb)
    try {
        Hotkey(key, fn, "On")
        g_StandardLoadingBarKeysHotkeys.Push(key)
    } catch {
    }
}

StandardLoadingBar_KeyWrapper(key, cb, *) {
    StandardLoadingBar_CloseKeysOverlay()
    if (cb) {
        try cb.Call()
        catch {
        }
    }
}

StandardLoadingBar_KeysTimeoutFired(timeoutCallback) {
    if (timeoutCallback)
        try timeoutCallback.Call()
        catch {
        }
    StandardLoadingBar_CloseKeysOverlay()
}

; =============================================================================
; Hotstring Selector: Gemini Redirect Banner (non-blocking; uses standard loading indicator)
; =============================================================================
HotstringGeminiBanner_Show(text := "Gemini: inserting prompt...") {
    DictationGeminiConfirm_Hide()
    Sleep 50
    centerOnHwnd := 0
    try centerOnHwnd := WinGetID("A")
    catch {
    }
    if (!centerOnHwnd || !WinExist("ahk_id " centerOnHwnd))
        centerOnHwnd := 0
    StandardLoadingBar_Show(text, "3772FF", { passive: true, centerOnHwnd: centerOnHwnd, textWidth: 160, fontSize: 8,
        alpha: 204 })
}

HotstringGeminiBanner_Hide(*) {
    StandardLoadingBar_Hide(0)
}

; =============================================================================
; Dictation: "Send to Gemini?" confirmation banner (6s, Y to confirm; uses standard loading indicator)
; =============================================================================
DictationGeminiConfirm_Show() {
    ; No-op; use DictationGeminiConfirm_ShowAndWait() which uses StandardLoadingBar_ShowWithKeys.
}

DictationGeminiConfirm_Hide(*) {
    StandardLoadingBar_CloseKeysOverlay()
}

DictationGeminiConfirm_CleanupAndMaybeSubmit(submitToGemini) {
    global g_DictationGeminiConfirmBannerVisible
    g_DictationGeminiConfirmBannerVisible := false  ; Allow future show
    try Hotkey("y", "Off")
    try Hotkey("Y", "Off")
    SetTimer(DictationGeminiConfirm_OnTimeout, 0)
    DictationGeminiConfirm_Hide()
    if (submitToGemini) {
        Sleep 350
        GeminiDelayedSubmitFlow()
    }
}

DictationGeminiConfirm_OnY(*) {
    DictationGeminiConfirm_CleanupAndMaybeSubmit(true)
}

DictationGeminiConfirm_OnTimeout(*) {
    DictationGeminiConfirm_CleanupAndMaybeSubmit(false)
}

DictationGeminiConfirm_ShowAndWait() {
    global g_DictationGeminiConfirmBannerVisible
    ; Only one banner: skip if already visible (prevents duplicate from multiple PlayDictationCompletionChime runs).
    if (g_DictationGeminiConfirmBannerVisible)
        return
    g_DictationGeminiConfirmBannerVisible := true
    ; Only the official loading bar (standard loading indicator) may show this content. Hide any other bar/overlay first.
    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_Hide(0)
    HideDictationIndicator()
    Sleep 50
    centerOnHwnd := 0
    try centerOnHwnd := WinGetID("A")
    catch {
    }
    if (!centerOnHwnd || !WinExist("ahk_id " centerOnHwnd))
        centerOnHwnd := 0
    yCallbacks := Map("Y", DictationGeminiConfirm_OnY)
    ; Official loading bar only; no blue; single banner (no border) so only one GUI shows.
    StandardLoadingBar_ShowWithKeys("Send transcription to Gemini? Press Y (6s)", yCallbacks, 6000, centerOnHwnd,
        DictationGeminiConfirm_OnTimeout, "1E1E2E", 300, 9, "", true)
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
        ShowCenteredOverlay_Utils("🔊 Sound: ON", 2000)
    } else {
        ShowCenteredOverlay_Utils("🔇 Sound: OFF", 2000)
    }
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
        outlookRunning := ProcessExist("OUTLOOK.EXE")
        teamsRunning := ProcessExist("ms-teams.exe")

        ; Show start banner
        if (outlookRunning && teamsRunning) {
            ShowCenteredOverlay_Utils("Closing Outlook and Teams...", 1500)
        } else {
            ShowCenteredOverlay_Utils("Opening Outlook and Teams...", 1500)
        }

        if (outlookRunning && teamsRunning) {
            ; Both are open: Close Outlook and minimize Teams to system tray
            ; Close Outlook process
            try {
                ProcessClose("OUTLOOK.EXE")
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

                    ; Launch using shortcut if available, otherwise use executable
                    if (outlookPath != "") {
                        Run outlookPath
                    } else {
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
                    ShowCenteredOverlay_Utils("Teams activated", 1500)
                } else {
                    ShowCenteredOverlay_Utils("Teams: Window not found", 2000)
                }
            } catch Error as e {
                ShowCenteredOverlay_Utils("Teams: Error - " . e.Message, 2000)
            }

            ; Second: Activate Outlook last (so it gets final focus)
            try {
                if (ProcessExist("OUTLOOK.EXE")) {
                    ; Wait for Outlook window to appear (up to 5 seconds)
                    WinWait("ahk_exe OUTLOOK.EXE", , 5)

                    if (!WinExist("ahk_exe OUTLOOK.EXE")) {
                        ShowCenteredOverlay_Utils("Outlook not running.", 2000)
                        return
                    } else {
                        ; Activate Outlook (this will bring it to foreground, overriding Teams)
                        WinActivate("ahk_exe OUTLOOK.EXE")
                        WinWaitActive("ahk_exe OUTLOOK.EXE", , 2)
                    }
                }
            } catch Error as e {
                ; Silently fail if activation doesn't work
            }
        }

        ; Show finish banner
        ShowCenteredOverlay_Utils("Done", 1500)
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

    ; Check Outlook status
    if (checkOutlook) {
        outlookRunning := ProcessExist("OUTLOOK.EXE")
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

; Infinite Dictation Macro
; Each loop = one 60s cycle; dictation cycles on/off every 15s within a loop to prevent transcription timeouts
ToggleDictationLoop() {
    global g_DictationLoopActive, g_ProgrammaticDictationStop

    if (g_DictationLoopActive) {
        ; Stop Infinite Dictation
        g_DictationLoopActive := false
        ; Turn off timers
        SetTimer(DictationLoopStop, 0)
        SetTimer(DictationLoopStart, 0)
        ; Send Win+Alt+Shift+0 to finish dictation
        g_ProgrammaticDictationStop := true
        SendInput "#!+0"
        ; Merge non-favorite clips: 5s countdown when user finishes Infinite Dictation (N or End to cancel)
        DictationMerge_StartCountdown(5)
    } else {
        ; Start Infinite Dictation (first loop)
        ; Clear any existing timers first to prevent old timers from firing
        SetTimer(DictationLoopStop, 0)
        SetTimer(DictationLoopStart, 0)
        g_DictationLoopActive := true
        ; Begin the first loop
        DictationLoopStart()
    }
}

DictationLoopStart() {
    ; Delegate to Infinite Dictation module when it owns the loop
    if (InfiniteDictation.IsActive) {
        InfiniteDictation.LoopCycle()
        return
    }
    global g_DictationLoopActive, g_DictationStartRetries, g_ProgrammaticDictationStop

    ; Safety check: Only proceed if Infinite Dictation is still active
    ; This prevents starting if user manually stopped it
    if (!g_DictationLoopActive) {
        return
    }

    ; Ensure Handy is running before attempting to start dictation
    if (!ProcessExist("handy.exe")) {
        Handy_ActivateOrLaunch()
        Sleep 2000 ; Wait for launch
    }

    ; Check if already recording to prevent toggling off
    if (WinExist("Recording ahk_exe handy.exe")) {
        ; Already recording, just ensure timer is running
        SetTimer(DictationLoopStop, 0)
        SetTimer(DictationLoopStop, -15000)
        return
    }

    g_DictationStartRetries := 0
    ; Send Win+Alt+Shift+0 to start dictation
    g_ProgrammaticDictationStop := true
    SendEvent "#!+0"

    ; Double-check Infinite Dictation is still active before scheduling next loop
    ; User may have stopped it during the dictation start delay
    if (!g_DictationLoopActive) {
        return
    }

    ; Clear any existing timer first to prevent accumulation
    SetTimer(DictationLoopStop, 0)

    ; Schedule stop after 15s (one loop segment) - negative period = one-shot timer
    ; Only schedules if Infinite Dictation is still active (checked above)
    SetTimer(DictationLoopStop, -15000)

    ; Verification: Check if window appeared after a delay
    SetTimer(VerifyDictationStart, -1500)
}

global g_DictationStartRetries := 0

VerifyDictationStart() {
    global g_DictationLoopActive, g_DictationStartRetries, g_ProgrammaticDictationStop
    if (!g_DictationLoopActive) {
        return
    }

    if (!WinExist("Recording ahk_exe handy.exe")) {
        g_DictationStartRetries++
        if (g_DictationStartRetries <= 3) {
            ; Retry start if window didn't appear
            g_ProgrammaticDictationStop := true
            SendEvent "#!+0"
            ; Reschedule stop timer just in case
            SetTimer(DictationLoopStop, 0)
            SetTimer(DictationLoopStop, -15000)
            SetTimer(VerifyDictationStart, -1500)
        } else {
            ShowCenteredOverlay_Utils("Failed to start dictation", 2000)
            g_DictationLoopActive := false
        }
    }
}

DictationLoopStop() {
    global g_DictationLoopActive, g_DictationLoopSound, g_ProgrammaticDictationStop

    ; Safety check: Only proceed if Infinite Dictation is still active
    ; This prevents restarting next loop if user manually stopped via ToggleDictationLoop()
    if (!g_DictationLoopActive) {
        return
    }

    ; Only send stop command if actually recording
    if (WinExist("Recording ahk_exe handy.exe")) {
        ; Send Win+Alt+Shift+0 to stop dictation (triggers transcription)
        g_ProgrammaticDictationStop := true
        SendEvent "#!+0"

        ; Play sound to notify that transcription has started (if enabled)
        if (IsSoundEnabled()) {
            SoundPlay(g_DictationLoopSound)
        }
    } else {
        ; If not recording, we might have stopped early or crashed.
        ; Restart loop immediately to recover.
        SetTimer(DictationLoopStart, -1000)
    }

    ; Double-check loop is still active before scheduling restart
    ; User may have stopped it during the sound playback delay
    if (!g_DictationLoopActive) {
        return
    }

    ; Clear any existing timer first to prevent accumulation
    SetTimer(DictationLoopStart, 0)

    ; REMOVED: Schedule next start after 1 second
    ; We now rely on PlayDictationCompletionChime to trigger next loop
    ; after transcription is complete.
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
    ov.BackColor := "3772FF"
    ov.SetFont("s24 cFFFFFF Bold", "Segoe UI")
    g_DictationCleanupTextCtrl := ov.Add("Text", "w650 Center", "Clearing clipboard in " g_DictationCleanupRemaining "… (press Y to proceed, N or End to cancel)"
    )
    ov.Show("AutoSize Hide")
    ov.GetPos(&gx, &gy, &gw, &gh)

    if hasWindow {
        cx := wx + (ww - gw) // 2
        cy := wy + (wh - gh) // 2
        ov.Show("x" . cx . " y" . cy . " NA")
    } else {
        vx := SysGet(76)  ; SM_XVIRTUALSCREEN
        vy := SysGet(77)  ; SM_YVIRTUALSCREEN
        vw := SysGet(78)  ; SM_CXVIRTUALSCREEN
        vh := SysGet(79)  ; SM_CYVIRTUALSCREEN
        cx := vx + (vw - gw) // 2
        cy := vy + (vh - gh) // 2
        ov.Show("x" . cx . " y" . cy . " NA")
    }

    WinSetTransparent(178, ov)
    g_DictationCleanupGui := ov
}

DictationCleanup_HideBanner() {
    global g_DictationCleanupGui, g_DictationCleanupTextCtrl
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
    ov.BackColor := "FFCC00"
    ov.SetFont("s24 c000000 Bold", "Segoe UI")
    g_DictationMergeTextCtrl := ov.Add("Text", "w650 Center", "Merging non-favorite clips in " g_DictationMergeRemaining "… (press Y to proceed, N or End to cancel)"
    )
    ov.Show("AutoSize Hide")
    try {
        if (hasWindow)
            ov.Show("AutoSize x" (wx + (ww - 650) // 2) " y" (wy + (wh - 80) // 2))
        else
            ov.Show("AutoSize")
    } catch {
        ov.Show("AutoSize")
    }
    g_DictationMergeGui := ov
}

DictationMerge_HideBanner() {
    global g_DictationMergeGui, g_DictationMergeTextCtrl
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
; Shows a confirmation prompt before executing cleanup
CleanClipboard() {
    ; Initialize Yes/No modal dialog
    result := MsgBox("Do you want to continue running the algorithm to exclude all clips?", "Clean the Clipboard",
        "YesNo")

    ; If user selects "No", terminate macro execution
    if (result = "No") {
        return
    }

    ; User selected "Yes" - proceed with the workflow
    CleanClipboardInternal()
}

; Dictation Toggle with Clipboard Cleanup Option (on start only)
; Toggles Infinite Dictation on/off. When starting, optionally asks to clean clipboard.
; When stopping, does NOT show clipboard cleanup prompt.
DictationStartWithClipboardOption() {
    global g_DictationLoopActive, g_PendingDictationMerge, g_ProgrammaticDictationStop

    if (g_DictationLoopActive) {
        ; Stop Infinite Dictation - show merge countdown when user finishes
        g_DictationLoopActive := false
        ; Turn off timers
        SetTimer(DictationLoopStop, 0)
        SetTimer(DictationLoopStart, 0)
        ; Send Win+Alt+Shift+0 to finish dictation
        g_ProgrammaticDictationStop := true
        SendInput "#!+0"
        ; Set flag to start merge countdown after transcription completes
        ; This ensures AI transcription and handy.exe finish before Clip Angel merge begins
        g_PendingDictationMerge := true
    } else {
        ; Start Infinite Dictation - show clipboard cleanup prompt ONLY when starting
        ; Show message box asking about clipboard cleanup
        result := MsgBox("Would you like to clean up the clipboard?", "Dictation Start", "YesNo")

        if (result = "Yes") {
            ; Execute clipboard cleanup algorithm without showing second prompt
            ; (User already confirmed they want to clean clipboard)
            CleanClipboardInternal()
        }
        ; If No, continue with Infinite Dictation without cleanup

        ; Clear any existing timers first to prevent old timers from firing
        SetTimer(DictationLoopStop, 0)
        SetTimer(DictationLoopStart, 0)
        g_DictationLoopActive := true
        ; Begin the first loop
        DictationLoopStart()
    }
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

; Check if a window title matches a project path
WindowMatchesProject(winTitle, projectPath) {
    if (projectPath = "") {
        return false
    }

    matchSegments := ExtractProjectMatchSegments(projectPath)

    ; Check if window title contains any of the match segments
    for segment in matchSegments {
        if (InStr(winTitle, segment)) {
            return true
        }
    }

    return false
}

; Get the project index that matches a window title
GetMatchingProjectIndex(winTitle) {
    global g_Projects, IS_WORK_ENVIRONMENT

    ; Check each project
    loop g_Projects.Length {
        projectIndex := A_Index
        project := g_Projects[projectIndex]

        ; Skip empty placeholders
        if (project.name = "" && project.path = "" && project.workPath = "") {
            continue
        }

        ; Select path based on environment
        projectPath := IS_WORK_ENVIRONMENT ? project.workPath : project.path

        ; If work environment but no workPath set, fall back to personal path
        if (IS_WORK_ENVIRONMENT && projectPath = "") {
            projectPath := project.path
        }

        ; Check if window matches this project
        if (WindowMatchesProject(winTitle, projectPath)) {
            return projectIndex
        }
    }

    return 0  ; No match found
}

; Build project index to character mapping (replicating ShowProjectSelector logic)
BuildProjectIndexToCharMap() {
    global g_Projects, g_ProjectCategories, g_ProjectCharSequence

    projectIndexToChar := Map()
    projectIndexToCategory := Map()

    ; Build map of project index to category
    loop g_Projects.Length {
        projectIndex := A_Index
        project := g_Projects[projectIndex]
        category := project.HasProp("category") ? project.category : "Personal"
        projectIndexToCategory[projectIndex] := category
    }

    charIndex := 1

    ; Assign characters sequentially within each category (same logic as ShowProjectSelector)
    for category in g_ProjectCategories {
        ; Find all project indices in this category
        categoryProjectIndices := []
        for projectIndex, cat in projectIndexToCategory {
            if (cat = category) {
                categoryProjectIndices.Push(projectIndex)
            }
        }

        ; Assign characters to projects in this category
        for projectIndex in categoryProjectIndices {
            project := g_Projects[projectIndex]

            ; Skip empty placeholders
            if (project.name = "" && project.path = "" && project.workPath = "") {
                charIndex++
                continue
            }

            ; Check if we have a character available
            if (charIndex > g_ProjectCharSequence.Length) {
                break
            }

            char := g_ProjectCharSequence[charIndex]

            ; Skip character "3" - it's reserved for preview window activation
            if (char = "3") {
                charIndex++
                if (charIndex > g_ProjectCharSequence.Length) {
                    break
                }
                char := g_ProjectCharSequence[charIndex]
            }

            projectIndexToChar[projectIndex] := char
            charIndex++
        }
    }

    return projectIndexToChar
}

; =============================================================================
; Cursor Window Focus Selector
; =============================================================================

; Global variables for Cursor focus selector
global g_CursorFocusSelectorGui := false
global g_CursorFocusSelectorActive := false
global g_CursorFocusWindowMap := Map()  ; Maps character to window HWND
global g_CursorFocusHotkeyHandlers := []  ; Store hotkey handlers for cleanup

; Get Cursor windows with assigned keys (matching Project Selector key assignments)
GetCursorWindowsWithKeys() {
    global g_Projects, g_ProjectCharSequence

    ; Build project index to character mapping
    projectIndexToChar := BuildProjectIndexToCharMap()

    ; Get all Cursor windows
    cursorWindows := WinGetList("ahk_exe Cursor.exe")

    ; Build list of windows with their assigned keys
    windowsWithKeys := []
    usedKeys := Map()
    usedProjectIndices := Map()

    ; First pass: assign keys to windows that match projects
    for hwnd in cursorWindows {
        try {
            winTitle := WinGetTitle("ahk_id " . hwnd)
            if (winTitle = "") {
                winTitle := "Untitled"
            }

            ; Check if this window matches a project
            matchingProjectIndex := GetMatchingProjectIndex(winTitle)

            if (matchingProjectIndex > 0 && projectIndexToChar.Has(matchingProjectIndex)) {
                char := projectIndexToChar[matchingProjectIndex]
                ; Only use this key once
                if (!usedKeys.Has(char) && !usedProjectIndices.Has(matchingProjectIndex)) {
                    windowsWithKeys.Push({ hwnd: hwnd, title: winTitle, char: char, projectIndex: matchingProjectIndex })
                    usedKeys[char] := true
                    usedProjectIndices[matchingProjectIndex] := true
                } else {
                    ; Mark as unassigned for now, will assign in second pass
                    windowsWithKeys.Push({ hwnd: hwnd, title: winTitle, char: "", projectIndex: matchingProjectIndex })
                }
            } else {
                ; No project match, will assign in second pass
                windowsWithKeys.Push({ hwnd: hwnd, title: winTitle, char: "", projectIndex: 0 })
            }
        } catch {
            ; Skip windows we can't access
            continue
        }
    }

    ; Second pass: assign remaining keys to unmatched windows
    charIndex := 1
    for window in windowsWithKeys {
        ; Skip if already assigned
        if (window.char != "") {
            continue
        }

        ; Find next available character
        while (charIndex <= g_ProjectCharSequence.Length) {
            char := g_ProjectCharSequence[charIndex]

            ; Skip character "3" - reserved for preview windows
            if (char = "3") {
                charIndex++
                continue
            }

            ; Check if this character is already used
            if (!usedKeys.Has(char)) {
                window.char := char
                usedKeys[char] := true
                charIndex++
                break
            }

            charIndex++
        }
    }

    ; Remove windows without assigned keys (shouldn't happen, but safety check)
    filtered := []
    for window in windowsWithKeys {
        if (window.char != "") {
            filtered.Push(window)
        }
    }

    return filtered
}

; Focus Cursor window and close all others
FocusCursorWindowAndCloseOthers(targetHwnd) {
    ; Get all Cursor windows
    allCursorWindows := WinGetList("ahk_exe Cursor.exe")

    ; Iterate through all windows and close those that don't match target
    for hwnd in allCursorWindows {
        if (hwnd != targetHwnd) {
            try {
                WinClose("ahk_id " . hwnd)
            } catch {
                ; Silently ignore if window close fails
            }
        }
    }

    ; Activate the target window
    if (!WinExist("ahk_id " . targetHwnd)) {
        ShowCenteredOverlay_Utils("Error: Target window not found.", 2000)
        return
    }
    try {
        WinActivate("ahk_id " . targetHwnd)
        WinWaitActive("ahk_id " . targetHwnd, , 1)
    } catch {
        ; Ignore if activation fails
    }
}

; Cleanup function for Cursor focus selector
CleanupCursorFocusSelector() {
    global g_CursorFocusSelectorActive, g_CursorFocusSelectorGui, g_CursorFocusHotkeyHandlers
    global g_CursorFocusWindowMap

    ; Disable active flag
    g_CursorFocusSelectorActive := false

    ; Disable all character hotkeys
    for handler in g_CursorFocusHotkeyHandlers {
        try {
            char := handler.char
            ; Handle special VK codes
            if (char = ",") {
                Hotkey("vkBC", "Off")
            } else if (char = ".") {
                Hotkey("vkBE", "Off")
            } else {
                Hotkey(char, "Off")
                ; Also disable uppercase for lowercase letters
                if (RegExMatch(char, "^[a-z]$")) {
                    Hotkey(StrUpper(char), "Off")
                }
            }
        } catch {
            ; Silently ignore errors
        }
    }

    ; Disable Escape hotkey
    try {
        Hotkey("Escape", HandleCursorFocusEscape, "Off")
    } catch {
        ; Ignore
    }

    ; Clear handlers array
    g_CursorFocusHotkeyHandlers := []

    ; Close and destroy GUI
    if (IsObject(g_CursorFocusSelectorGui)) {
        try {
            g_CursorFocusSelectorGui.Destroy()
        } catch {
            ; Ignore
        }
        g_CursorFocusSelectorGui := false
    }

    ; Clear window map
    g_CursorFocusWindowMap := Map()
}

; Handler for Escape key in Cursor focus selector
HandleCursorFocusEscape(*) {
    global g_CursorFocusSelectorActive
    if (g_CursorFocusSelectorActive) {
        CleanupCursorFocusSelector()
    }
}

; Handler for character key press in Cursor focus selector
HandleCursorFocusChar(char) {
    global g_CursorFocusSelectorActive, g_CursorFocusWindowMap

    ; Only process if selector is active
    if (!g_CursorFocusSelectorActive) {
        return
    }

    ; Get the HWND for this character
    targetHwnd := g_CursorFocusWindowMap.Get(char, "")
    if (targetHwnd = "") {
        ; Try lowercase if uppercase
        targetHwnd := g_CursorFocusWindowMap.Get(StrLower(char), "")
    }

    if (targetHwnd != "") {
        ; Cleanup first (closes GUI, disables hotkeys)
        CleanupCursorFocusSelector()

        ; Focus the window and close others
        FocusCursorWindowAndCloseOthers(targetHwnd)
    }
}

; Factory function to create a handler for Cursor focus selection
CreateCursorFocusCharHandler(char) {
    ; Return a function that captures the char value at creation time
    return (*) => HandleCursorFocusChar(char)
}

; Show Cursor focus selector GUI
ShowCursorFocusSelector() {
    global g_CursorFocusSelectorGui, g_CursorFocusSelectorActive, g_CursorFocusWindowMap
    global g_CursorFocusHotkeyHandlers

    ; Close existing selector if open
    if (g_CursorFocusSelectorActive && IsObject(g_CursorFocusSelectorGui)) {
        CleanupCursorFocusSelector()
        Sleep 50
    }

    ; Also close hotstring selector if it's open (since we're called from within it)
    global g_HotstringSelectorActive
    if (g_HotstringSelectorActive) {
        CleanupHotstringSelector()
        Sleep 50
    }

    ; Get Cursor windows with assigned keys
    windowsWithKeys := GetCursorWindowsWithKeys()

    if (windowsWithKeys.Length = 0) {
        ; Use tray notification to avoid stealing focus
        TrayTip("Focus Cursor Window", "No Cursor windows found.", "IconX")
        SetTimer(() => TrayTip(), -5000)  ; Auto-hide after ~5s
        return
    }

    ; If only one window, just focus it and return
    if (windowsWithKeys.Length = 1) {
        CleanupHotstringSelector()
        FocusCursorWindowAndCloseOthers(windowsWithKeys[1].hwnd)
        return
    }

    ; Clear window map
    g_CursorFocusWindowMap := Map()

    ; Get active monitor for positioning
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

    ; Create GUI - non-activating so it doesn't steal focus
    g_CursorFocusSelectorGui := Gui("+AlwaysOnTop +ToolWindow +E0x08000000", "Focus Window (Close Others)")
    fontSize := (monitorHeight < 800) ? 9 : 10
    g_CursorFocusSelectorGui.SetFont("s" . fontSize, "Segoe UI")
    g_CursorFocusSelectorGui.MarginX := 15
    g_CursorFocusSelectorGui.MarginY := 10

    ; Build display text
    displayText := "=== FOCUS WINDOW (CLOSE OTHERS) ===`n`n"

    for window in windowsWithKeys {
        ; Map character to HWND
        g_CursorFocusWindowMap[window.char] := window.hwnd

        ; Add to display
        displayText .= "[" . window.char . "] " . window.title . "`n"
    }

    displayText .= "`n[ESC] Cancel"

    ; Calculate text dimensions
    baseWidth := 450
    lineHeight := fontSize + 6
    lineCount := StrSplit(displayText, "`n").Length
    textControlHeight := Min(400, lineCount * lineHeight + 20)

    ; Add text control
    g_CursorFocusSelectorGui.Add("Text", "w" . (baseWidth - 30), displayText)

    ; Add close button
    closeBtn := g_CursorFocusSelectorGui.Add("Button", "w80 Center", "Close")
    closeBtn.OnEvent("Click", (*) => CleanupCursorFocusSelector())

    ; Calculate total height
    totalHeight := 20 + textControlHeight + 40 + 10

    ; Calculate center position
    marginX := 20
    marginY := 20
    guiX := monitorLeft + (monitorWidth - baseWidth) // 2
    guiY := monitorTop + (monitorHeight - totalHeight) // 2

    ; Ensure GUI stays within monitor bounds
    if (guiX < monitorLeft + marginX)
        guiX := monitorLeft + marginX
    if (guiY < monitorTop + marginY)
        guiY := monitorTop + marginY
    if (guiX + baseWidth > monitorLeft + monitorWidth - marginX)
        guiX := monitorLeft + monitorWidth - baseWidth - marginX
    if (guiY + totalHeight > monitorTop + monitorHeight - marginY)
        guiY := monitorTop + monitorHeight - totalHeight - marginY

    ; Show GUI
    g_CursorFocusSelectorGui.Show("NA w" . baseWidth . " h" . totalHeight . " x" . guiX . " y" . guiY)

    ; Set active flag
    g_CursorFocusSelectorActive := true

    ; Clear handlers array
    g_CursorFocusHotkeyHandlers := []

    ; Enable hotkeys for assigned characters
    for window in windowsWithKeys {
        char := window.char

        ; Create handler
        handler := CreateCursorFocusCharHandler(char)

        ; Store handler for cleanup
        g_CursorFocusHotkeyHandlers.Push({ char: char, handler: handler })

        ; Enable hotkey
        try {
            ; Handle special characters that need VK codes
            if (char = ",") {
                Hotkey("vkBC", handler, "On")
            } else if (char = ".") {
                Hotkey("vkBE", handler, "On")
            } else {
                Hotkey(char, handler, "On")
                ; Also enable uppercase for lowercase letters
                if (RegExMatch(char, "^[a-z]$")) {
                    Hotkey(StrUpper(char), handler, "On")
                }
            }
        } catch {
            ; Silently ignore if we can't create hotkey for this character
        }
    }

    ; Enable Escape hotkey
    Hotkey("Escape", HandleCursorFocusEscape, "On")
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
    ; Merge Non-Favorite Clips macro (assigned to "U")
    RegisterMacro(MergeNonFavoriteClips, "📋 Merge Non-Favorite Clips", "u")
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
; Target path: OneDrive Desktop. Red confirmation (with path), then success banner.
; =============================================================================
global g_DesktopToRecycleGui := 0
global g_DesktopToRecycleTextCtrl := 0
global g_DesktopToRecyclePath := ""  ; Set from GetDesktopToRecyclePath() when macro runs
global g_DesktopToRecycleCloseHwnd := 0

DesktopToRecycle_ShowBanner() {
    global g_DesktopToRecycleGui, g_DesktopToRecycleTextCtrl, g_DesktopToRecyclePath
    try {
        if IsObject(g_DesktopToRecycleGui)
            g_DesktopToRecycleGui.Destroy()
    } catch {
    }
    g_DesktopToRecycleGui := 0
    g_DesktopToRecycleTextCtrl := 0

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

    confirmText := "Move all items from:`n" . g_DesktopToRecyclePath . "`nto Recycle Bin? (Y = Yes, N = No)"
    ov := Gui("+AlwaysOnTop -Caption +ToolWindow")
    ov.BackColor := "C0392B"
    ov.SetFont("s22 cFFFFFF Bold", "Segoe UI")
    g_DesktopToRecycleTextCtrl := ov.Add("Text", "w700 Center", confirmText)
    ov.Show("AutoSize Hide")
    ov.GetPos(&gx, &gy, &gw, &gh)

    if hasWindow {
        cx := wx + (ww - gw) // 2
        cy := wy + (wh - gh) // 2
        ov.Show("x" . cx . " y" . cy . " NA")
    } else {
        vx := SysGet(76)
        vy := SysGet(77)
        vw := SysGet(78)
        vh := SysGet(79)
        cx := vx + (vw - gw) // 2
        cy := vy + (vh - gh) // 2
        ov.Show("x" . cx . " y" . cy . " NA")
    }

    WinSetTransparent(178, ov)
    g_DesktopToRecycleGui := ov
}

DesktopToRecycle_HideBanner() {
    global g_DesktopToRecycleGui, g_DesktopToRecycleTextCtrl
    try {
        if IsObject(g_DesktopToRecycleGui)
            g_DesktopToRecycleGui.Destroy()
    } catch {
    }
    g_DesktopToRecycleGui := 0
    g_DesktopToRecycleTextCtrl := 0
}

DesktopToRecycle_SetConfirmHotkeys(enable := true) {
    if (enable) {
        Hotkey("*y", DesktopToRecycle_Confirm, "On")
        Hotkey("*n", DesktopToRecycle_Cancel, "On")
    } else {
        try Hotkey("*y", "Off")
        catch {
        }
        try Hotkey("*n", "Off")
        catch {
        }
    }
}

DesktopToRecycle_Cancel(*) {
    DesktopToRecycle_SetConfirmHotkeys(false)
    DesktopToRecycle_HideBanner()
}

DesktopToRecycle_Confirm(*) {
    DesktopToRecycle_SetConfirmHotkeys(false)
    DesktopToRecycle_HideBanner()
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
            ShowCenteredOverlay_Utils("Desktop items moved to Recycle Bin", 2000, "27AE60")
            DesktopToRecycle_CloseDesktopExplorer(path)
        } else {
            ShowCenteredOverlay_Utils("Desktop path not found or error: " path, 3500, "C0392B")
            DesktopToRecycle_CloseDesktopExplorer(path)
        }
    } catch as err {
        ShowCenteredOverlay_Utils("Error moving to Recycle Bin", 2500, "C0392B")
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
    DesktopToRecycle_ShowBanner()
    DesktopToRecycle_SetConfirmHotkeys(true)
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
                ShowCenteredOverlay_Utils("Error: Target window not found.", 2000)
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

; Jump mouse right (short distance) - now shows square selector
#!+Right::
{
    HandleDirectionHotkey("Right")
    return
}

; Jump mouse left (short distance) - now shows square selector
#!+Left::
{
    HandleDirectionHotkey("Left")
}

; Jump mouse down (short distance) - now shows square selector
#!+Down::
{
    HandleDirectionHotkey("Down")
}

; Jump mouse up (short distance) - now shows square selector
#!+Up::
{
    HandleDirectionHotkey("Up")
}

; =============================================================================
; Peek PDF – Win+Alt+Shift+X
; If Peek is open: activate it. Otherwise: show study-topic selector (same aesthetic as Win+Alt+Shift+C).
; =============================================================================

; Study topics for Win+Alt+Shift+X selector. Paths are relative to notes repo (GetNotesRepoPath()).
global g_StudyTopics := Map(
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

; Monitor PDF (Peek) window focus and automatically disable focus mode when it loses focus
MonitorPdfFocus() {
    global g_PdfFocusTrackedHwnd

    ; Check if tracked window still exists
    if (g_PdfFocusTrackedHwnd && !WinExist("ahk_id " . g_PdfFocusTrackedHwnd)) {
        DisableFocusMode()
        StopPdfFocusMonitor()
        return
    }

    ; Check if PDF window is still the active window
    if (!WinActive("ahk_id " . g_PdfFocusTrackedHwnd)) {
        DisableFocusMode()
        StopPdfFocusMonitor()
    }
}

; Start monitoring PDF window focus
StartPdfFocusMonitor(hwnd := 0) {
    global g_PdfFocusMonitorTimer, g_PdfFocusTrackedHwnd

    StopPdfFocusMonitor()

    g_PdfFocusTrackedHwnd := hwnd ? hwnd : WinExist("A")
    if (!g_PdfFocusTrackedHwnd)
        return

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

ShowStudyTopicSelector() {
    global g_StudyTopicSelectorGui, g_StudyTopicSelectorActive, g_StudyTopics

    if (g_StudyTopicSelectorActive)
        return

    g_StudyTopicSelectorGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Owner")
    g_StudyTopicSelectorGui.BackColor := "1E1E2E"
    g_StudyTopicSelectorGui.MarginX := 20
    g_StudyTopicSelectorGui.MarginY := 15

    g_StudyTopicSelectorGui.SetFont("s14 cCDD6F4 Bold", "Segoe UI")
    g_StudyTopicSelectorGui.Add("Text", "w280 Center", "📚 Study topic (Peek)")
    g_StudyTopicSelectorGui.Add("Text", "w280 h1 Background45475A")

    g_StudyTopicSelectorGui.SetFont("s12 cCDD6F4", "Segoe UI")
    for num, topic in g_StudyTopics {
        g_StudyTopicSelectorGui.Add("Text", "w280", "[" . num . "] " . topic.name)
    }

    g_StudyTopicSelectorGui.Add("Text", "w280 h1 Background45475A y+10")
    g_StudyTopicSelectorGui.SetFont("s9 c6C7086", "Segoe UI")
    g_StudyTopicSelectorGui.Add("Text", "w280 Center", "Press 1–5 | Esc to cancel")

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
        try ShowCenteredOverlay_Utils("Notes repo path not set (env.ahk).", 3000, "FFAA00")
        return
    }
    fullPath := RTrim(basePath, "\") . topic.path
    if (!FileExist(fullPath)) {
        try ShowCenteredOverlay_Utils("PDF not found: " fullPath, 3500, "FF0000")
        return
    }
    peekExe := PeekPdf_ResolvePeekExePath()
    if (!FileExist(peekExe)) {
        try ShowCenteredOverlay_Utils("Peek executable not found.", 2500, "FF0000")
        return
    }
    PeekPdf_OpenPath(fullPath)
}

StudyTopicSelector_Cancel(*) {
    StudyTopicSelector_Close()
}

StudyTopicSelector_Close() {
    global g_StudyTopicSelectorGui, g_StudyTopicSelectorActive

    if (!g_StudyTopicSelectorActive)
        return
    g_StudyTopicSelectorActive := false
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
    envExe := GetPeekExePath()
    if (FileExist(envExe))
        return envExe
    return "peek.exe"
}

; Open a specific PDF in PowerToys Peek and run WaitAndConfigure. Caller must validate pdfPath and exe exist.
PeekPdf_OpenPath(pdfPath) {
    peekExe := PeekPdf_ResolvePeekExePath()
    peekEsc := StrReplace(peekExe, "'", "''")
    pdfEsc := StrReplace(pdfPath, "'", "''")
    psArg := "& " . Chr(39) . peekEsc . Chr(39) . " " . Chr(39) . pdfEsc . Chr(39)
    cmd := "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command " . Chr(34) . psArg . Chr(34)
    try Run cmd, "", "Hide"
    catch as e {
        try ShowCenteredOverlay_Utils("Failed to open Peek: " e.Message, 3000, "FF0000")
        return
    }
    if WinWait("Peek", "", 5) {
        WinMaximize
        PeekPdf_WaitAndConfigure()
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
        try ShowCenteredOverlay_Utils("No PDF path set. Hold Win+Alt+Shift+X to set.", 3000, "FFAA00")
        return
    }
    peekExe := PeekPdf_ResolvePeekExePath()
    if (!FileExist(peekExe)) {
        try ShowCenteredOverlay_Utils("Peek executable not found.", 2500, "FF0000")
        return
    }
    if (!FileExist(pdfPath)) {
        try ShowCenteredOverlay_Utils("PDF file not found: " pdfPath, 3500, "FF0000")
        return
    }
    peekEsc := StrReplace(peekExe, "'", "''")
    pdfEsc := StrReplace(pdfPath, "'", "''")
    psArg := "& " . Chr(39) . peekEsc . Chr(39) . " " . Chr(39) . pdfEsc . Chr(39)
    cmd := "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command " . Chr(34) . psArg . Chr(34)
    try Run cmd, "", "Hide"
    catch as e {
        try ShowCenteredOverlay_Utils("Failed to open Peek: " e.Message, 3000, "FF0000")
        return
    }
    if WinWait("Peek", "", 5) {
        WinMaximize
        PeekPdf_WaitAndConfigure()
    }
}

; Wait for Peek PDF toolbar to load (Page view button), click it, two-page view, focus, go to last page.
; Current state: PDF opening and window maximization are working correctly.
; Execution order: 1) Get Peek hwnd  2) UIA root from hwnd  3) Poll for "Page view" anchor
;  4) Wait for anchor visible + short delay before click  5) Click Page view  6) Right Arrow
;  7) Click window center  8) Ctrl+End. Fallback: Sleep 400 + Click if UIA or anchor fails.
PeekPdf_WaitAndConfigure() {
    global UIA
    ; Standard loading bar: show for the whole process so user knows when we started and when we finished
    StandardLoadingBar_Show("Peek PDF: configuring...", "3772FF")
    ; 1) Get Peek window hwnd
    hwnd := WinExist("Peek")
    if (!hwnd)
        hwnd := WinExist("ahk_exe PowerToys.Peek.UI.exe")
    if (!hwnd) {
        StandardLoadingBar_Update("Peek PDF: window not found", "FFAA00")
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
            ; 8) Go to final page via UIA (keystrokes do not reach the embedded Edge PDF viewer)
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
            Sleep 100
            StandardLoadingBar_Update("Peek PDF: done", "27AE60")
            StandardLoadingBar_Hide(2000)
        } else {
            StandardLoadingBar_Update("Peek PDF: finished (fallback)", "FFAA00")
            StandardLoadingBar_Hide(2000)
            Sleep 400
            Click "Left"
        }
    } catch {
        StandardLoadingBar_Update("Peek PDF: finished (fallback)", "FFAA00")
        StandardLoadingBar_Hide(2000)
        Sleep 400
        Click "Left"
    }
}

#!+x::
{
    hwnd := WinExist("ahk_exe PowerToys.Peek.UI.exe")
    if !hwnd
        hwnd := WinExist("Peek")
    if hwnd {
        try {
            WinShow("ahk_id " hwnd)
            WinActivate("ahk_id " hwnd)
        }
        EnableFocusMode()
        StartPdfFocusMonitor(hwnd)
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
global g_HotstringHotkeyHandlers := []          ; Array of hotkey handler objects for cleanup on close
global g_HotstringPromptCharMap := Map()        ; Map of prompt-assigned chars => true (rebuilt on each ShowHotstringSelector)
global g_HotstringGeminiArmed := false          ; When true, next Prompts selection is redirected to Gemini
global g_HotstringGeminiAutoSubmit := true      ; During delayed flow: true = send Enter after paste; false = paste only
global g_HotstringGeminiSubmitTimer := false   ; Timer reference for 4s delayed submit (for cleanup if needed)
global g_HotstringGeminiRestoreHwnd := 0        ; Window to restore focus to after 4s banner + paste (set at start of GeminiDelayedSubmitFlow)

; Character assignment sequence: defines order in which characters are assigned to actions
; Format: ["1", "2", "3", "4", "5", "q", "w", "e", "r", "t", "a", "s", "d", "f", "g", "z", "x",
;          "c", "v", "b", "6", "7", "8", "9", "0", "y", "u", "i", "o", "p", "h", "j", "k", "l", "n", "m", ",", "."]
global g_HotstringCharSequence := ["1", "2", "3", "4", "5", "q", "w", "e", "r", "t", "a", "s", "d", "f", "g", "z", "x",
    "c", "v", "b", "6", "7", "8", "9", "0", "y", "u", "i", "o", "p", "h", "j", "k", "l", "n", "m", ",", "."]

; Category display order: defines the sequence in which action categories appear in the GUI
; Order: Prompts → General → Projects → Files & Links → Macros
global g_HotstringCategories := ["Prompts", "General", "Projects", "Files & Links", "Macros"]

; Reserved empty character: never assigned to any action; always shows as (empty) in selector
; Set to "" to disable reservation
global g_ReservedEmptyChar := ""

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

    ; Group hotstrings by category
    categorized := Map()
    categorized["Projects"] := []
    categorized["Prompts"] := []
    categorized["General"] := []

    if (IsSet(g_hotstrings) && g_hotstrings.Length > 0) {
        for hs in g_hotstrings {
            category := hs.category
            if (category = "Projects" || category = "Prompts" || category = "General") {
                categorized[category].Push(hs)
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
    categorized["Files & Links"] := []
    categorized["Macros"] := []

    ; Add hotstrings
    if (IsSet(g_hotstrings) && g_hotstrings.Length > 0) {
        for hs in g_hotstrings {
            category := hs.category
            if (category = "Projects" || category = "Prompts" || category = "General") {
                categorized[category].Push(hs)
            }
        }
    }

    ; Add quick open files
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

    ; Disable active flag
    g_HotstringSelectorActive := false

    ; Disable all character hotkeys
    for handler in g_HotstringHotkeyHandlers {
        try {
            char := handler.char
            ; Handle special VK codes
            if (char = ",") {
                Hotkey("vkBC", "Off")
            } else if (char = ".") {
                Hotkey("vkBE", "Off")
            } else {
                Hotkey(char, "Off")
                ; Also disable uppercase for lowercase letters
                if (RegExMatch(char, "^[a-z]$")) {
                    Hotkey(StrUpper(char), "Off")
                }
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

    ; Clear handlers array
    g_HotstringHotkeyHandlers := []
    g_HotstringPromptCharMap := Map()
    g_HotstringGeminiArmed := false

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
    DetectHiddenWindows true
    SetTitleMatchMode 2
    geminiPid := 0
    for hwnd in WinGetList("ahk_exe AutoHotkey64.exe") {
        try {
            if (InStr(WinGetTitle("ahk_id " hwnd), "Gemini.ahk")) {
                geminiPid := WinGetPID("ahk_id " hwnd)
                break
            }
        } catch {
            continue
        }
    }
    if (!geminiPid) {
        for hwnd in WinGetList("ahk_exe AutoHotkey32.exe") {
            try {
                if (InStr(WinGetTitle("ahk_id " hwnd), "Gemini.ahk")) {
                    geminiPid := WinGetPID("ahk_id " hwnd)
                    break
                }
            } catch {
                continue
            }
        }
    }
    sent := false
    if (geminiPid) {
        for hwnd in WinGetList("ahk_pid " geminiPid) {
            try {
                SendMessage(WM_START_DELAYED_SUBMIT_MONITOR, originalHwnd, geminiChromeHwnd, , "ahk_id " hwnd)
                sent := true
                break
            } catch {
                continue
            }
        }
    }
    DetectHiddenWindows false
}

GeminiDelayedSubmitFlow() {
    global g_HotstringGeminiAutoSubmit, g_HotstringGeminiRestoreHwnd
    g_HotstringGeminiRestoreHwnd := WinExist("A")  ; Store window to restore focus to after 4s sequence
    g_HotstringGeminiAutoSubmit := true

    HotstringGeminiBanner_Show("Submitting in 4s... Press N to cancel auto-submit")

    Hotkey("n", GeminiCancelAutoSubmit, "On")
    Hotkey("N", GeminiCancelAutoSubmit, "On")

    SetTimer(GeminiFinalizeSubmit, -4000)
}

GeminiCancelAutoSubmit(*) {
    global g_HotstringGeminiAutoSubmit, g_HotstringGeminiRestoreHwnd
    g_HotstringGeminiAutoSubmit := false
    try Hotkey("n", "Off")
    try Hotkey("N", "Off")
    SetTimer(GeminiFinalizeSubmit, 0)  ; cancel 4s timer so paste runs in deferred callback only
    HotstringGeminiBanner_Hide()
    HotstringGeminiBanner_Show("Auto-submit CANCELLED (Paste only)")
    SetTimer(HotstringGeminiBanner_Hide, -1500)
    ; Defer paste so it runs outside hotkey context; sync works and paste goes to Gemini
    SetTimer(GeminiCancelAutoSubmit_DoPaste, -400)
}

GeminiCancelAutoSubmit_DoPaste(*) {
    global g_HotstringGeminiRestoreHwnd
    ; Delay-submit flow: do not switch tabs; paste to currently active Gemini tab
    GeminiNavigateFocusAndPasteFirstSnippet("", false)
    Sleep 200  ; Let paste be received by Gemini before switching focus back
    if (g_HotstringGeminiRestoreHwnd && WinExist("ahk_id " g_HotstringGeminiRestoreHwnd))
        WinActivate("ahk_id " g_HotstringGeminiRestoreHwnd)
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

    try Hotkey("n", "Off")
    try Hotkey("N", "Off")
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

    ; Return focus to the window the user had before the 4s banner + paste (whether Enter was sent or not)
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

    ; Only process if selector is active
    if (!g_HotstringSelectorActive) {
        return
    }

    ; L key: first press = arm Gemini mode (show banner); second press (double-tap) = navigate to Gemini, focus field, paste first snippet.
    if (char = "l" || char = "L") {
        if (g_HotstringGeminiArmed) {
            ; Double-tap L: delayed submit flow (4s banner, N to cancel auto-submit; then navigate+paste, optionally Enter).
            CleanupHotstringSelector()
            GeminiDelayedSubmitFlow()
            g_HotstringGeminiArmed := false
            return
        }
        g_HotstringGeminiArmed := true
        ; Show banner when entering Gemini mode (same pattern as Project Selector "Entering Selection Mode").
        HotstringGeminiBanner_Show("Entering Gemini Mode - Select prompt")
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
    ; Check these first before checking the char maps
    if (char = "9") {
        ; Cleanup first (closes GUI, disables hotkeys)
        CleanupHotstringSelector()
        ; CIP & UX Integration mini workshop - Miro
        FindAndActivateMiroWindow("https://miro.com/app/board/uXjVJdbNFkA=/", "CIP & UX Integration")
        return
    } else if (char = "0") {
        ; Cleanup first (closes GUI, disables hotkeys)
        CleanupHotstringSelector()
        ; CIP Dashboard - Workspace - Miro
        FindAndActivateMiroWindow("https://miro.com/app/board/uXjVJVZSXvk=/", "CIP Dashboard")
        return
    }

    ; First check if character maps to a file path (quick open files)
    filePath := g_QuickOpenFileCharMap.Get(char, "")
    if (filePath = "") {
        ; Try lowercase if uppercase
        filePath := g_QuickOpenFileCharMap.Get(StrLower(char), "")
    }

    if (filePath != "") {
        ; Cleanup first (closes GUI, disables hotkeys)
        CleanupHotstringSelector()

        ; Determine file type and open accordingly
        SplitPath(filePath, , , &ext)
        ext := StrLower(ext)

        if (ext = "pbix") {
            ; Power BI file
            FindAndActivatePowerBIFile(filePath)
        } else {
            ; Generic file opening (fallback)
            try {
                Run(filePath)
            } catch {
                ; File opening failed
            }
        }
        return
    }

    ; Check if character maps to a macro function
    macroFunc := g_MacroCharMap.Get(char, "")
    if (macroFunc = "") {
        ; Try lowercase if uppercase
        macroFunc := g_MacroCharMap.Get(StrLower(char), "")
    }

    if (macroFunc != "") {
        ; Cleanup first (closes GUI, disables hotkeys)
        CleanupHotstringSelector()

        ; Execute the macro function
        try {
            macroFunc()
        } catch Error as e {
            ; Macro execution failed
        }
        return
    }

    ; Check if character maps to a hotstring expansion
    expansion := g_HotstringCharMap.Get(char, "")
    if (expansion = "") {
        ; Try lowercase if uppercase
        expansion := g_HotstringCharMap.Get(StrLower(char), "")
    }

    if (expansion != "") {
        ; Cleanup first (closes GUI, disables hotkeys)
        CleanupHotstringSelector()

        if (useGemini) {
            ; L+Prompt selection: redirect to Gemini (focus prompt field, paste, do NOT submit).
            HotstringGeminiBanner_Show("Gemini: inserting prompt...")
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

    ; Close existing GUI if open
    if (g_HotstringSelectorActive && IsObject(g_HotstringSelectorGui)) {
        CleanupHotstringSelector()
        Sleep 50
    }

    ; Build character mapping
    g_HotstringCharMap := BuildHotstringCharMap()

    ; Check if we have any items to display (hotstrings, quick open files, or macros)
    global g_QuickOpenFileCharMap, g_MacroCharMap
    hasItems := (g_HotstringCharMap.Count > 0) || (g_QuickOpenFileCharMap.Count > 0) || (g_MacroCharMap.Count > 0)
    if (!hasItems) {
        ; Use tray notification to avoid stealing focus/closing other palettes
        TrayTip("Hotstring Selector", "No hotstrings, files, or macros found.", "IconX")
        SetTimer(() => TrayTip(), -5000)  ; Auto-hide after ~5s
        return
    }

    ; Get categorized hotstrings
    categorized := GetCategorizedHotstrings()

    ; =============================================================================
    ; Dynamic Modal UI Adaptation Based on Monitor Configuration
    ;
    ; Monitor Dataset Structure (for reference):
    ; {
    ;   "monitor_dataset": [
    ;     {
    ;       "id": 1,
    ;       "resolution": "1920x1200",
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
    ;       "resolution": "1920x1080",
    ;       "orientation": "landscape",
    ;       "scale": "100%",
    ;       "ui_strategy": "dual_column_wide"
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
    g_HotstringSelectorGui := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x08000000", "Hotstring Shortcuts")
    g_HotstringSelectorGui.BackColor := "1E1E2E"
    g_HotstringSelectorGui.MarginX := 14
    g_HotstringSelectorGui.MarginY := 10
    ; Use slightly smaller font for compact display; Segoe UI to match C menu
    fontSize := (monitorHeight < 800) ? 9 : 9
    g_HotstringSelectorGui.SetFont("s" . fontSize . " cCDD6F4", "Segoe UI")

    ; Build reverse map: expansion -> character
    expansionToChar := Map()
    for char, expansion in g_HotstringCharMap {
        expansionToChar[expansion] := char
    }

    ; Track which selector characters belong to the Prompts category (non-empty expansions only).
    ; Used to restrict the L-modifier redirect behavior to Prompts only.
    global g_HotstringPromptCharMap
    g_HotstringPromptCharMap := Map()
    try {
        if (categorized.Has("Prompts")) {
            for hs in categorized["Prompts"] {
                try {
                    if (hs.HasProp("expansion") && hs.expansion != "" && expansionToChar.Has(hs.expansion)) {
                        g_HotstringPromptCharMap[expansionToChar[hs.expansion]] := true
                    }
                } catch {
                    ; Ignore malformed entries
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

                    allItems.Push({ category: category, char: char, text: itemText, isEmpty: isEmpty })
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

    ; Ensure reserved empty char always appears as (empty) if set
    if (g_ReservedEmptyChar != "") {
        hasReservedEmpty := false
        for item in allItems {
            if (item.char = g_ReservedEmptyChar) {
                hasReservedEmpty := true
                break
            }
        }
        if (!hasReservedEmpty)
            allItems.Push({ category: "Unassigned", char: g_ReservedEmptyChar, text: "[" . g_ReservedEmptyChar .
                "] > (empty)", isEmpty: true })
    }

    ; Show any remaining unassigned character slots
    if (currentCharIndex <= g_HotstringCharSequence.Length) {
        while (currentCharIndex <= g_HotstringCharSequence.Length) {
            char := g_HotstringCharSequence[currentCharIndex]
            if (char = "l") {
                allItems.Push({ category: "Unassigned", char: char, text: "[L] > Gemini: L = arm; L+L = open Gemini + paste first snippet (or Ctrl+Alt+Win+L)",
                    isEmpty: false })
            } else if (g_ReservedEmptyChar != "" && char = g_ReservedEmptyChar) {
                ; Already added above; skip to avoid duplicate
            } else {
                allItems.Push({ category: "Unassigned", char: char, text: "[" . char . "] > (empty)", isEmpty: true })
            }
            currentCharIndex++
        }
    }

    ; Re-sort so explicit "u" (empty) is in sequence order
    sortedItems := []
    for idx, seqChar in g_HotstringCharSequence {
        found := false
        for item in allItems {
            if (item.HasProp("explicitIndex") && item.explicitIndex = idx) {
                sortedItems.Push(item)
                found := true
                break
            }
        }
        if (!found) {
            for item in allItems {
                if (!item.HasProp("explicitIndex") && item.char = seqChar) {
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

    ; Build display text based on monitor orientation
    displayText := ""

    if (isPortrait) {
        ; PORTRAIT MODE: Single-column layout optimized for vertical space
        currentCategory := ""
        for item in allItems {
            if (item.category != currentCategory) {
                if (currentCategory != "")
                    displayText .= "`n"
                currentCategory := item.category
                displayText .= "— " . currentCategory . " —`n"
            }
            displayText .= item.text . "`n"
        }
        displayText .= "`n"
    } else {
        ; LANDSCAPE MODE: Two-column layout optimized for horizontal space
        ; Calculate maximum item text length to determine column width
        maxItemLength := 0
        for item in allItems {
            if (StrLen(item.text) > maxItemLength)
                maxItemLength := StrLen(item.text)
        }
        columnWidth := maxItemLength + 2
        if (columnWidth < 36)
            columnWidth := 36
        totalWidth := columnWidth * 2 + 6
        columnSpacing := "  "

        currentCategory := ""
        categoryItems := []

        for item in allItems {
            if (item.category != currentCategory) {
                if (currentCategory != "" && categoryItems.Length > 0) {
                    displayText .= "— " . currentCategory . " —`n"
                    midPoint := Ceil(categoryItems.Length / 2)
                    maxLines := categoryItems.Length - midPoint
                    if (midPoint > maxLines)
                        maxLines := midPoint
                    loop maxLines {
                        leftText := ""
                        rightText := ""
                        if (A_Index <= midPoint)
                            leftText := PadString(categoryItems[A_Index].text, columnWidth)
                        else
                            leftText := PadString("", columnWidth)
                        rightIdx := A_Index + midPoint
                        if (rightIdx <= categoryItems.Length)
                            rightText := categoryItems[rightIdx].text
                        else
                            rightText := ""
                        displayText .= leftText . columnSpacing . rightText . "`n"
                    }
                    displayText .= "`n"
                }
                currentCategory := item.category
                categoryItems := []
            }
            categoryItems.Push(item)
        }

        if (currentCategory != "" && categoryItems.Length > 0) {
            displayText .= "— " . currentCategory . " —`n"
            midPoint := Ceil(categoryItems.Length / 2)
            maxLines := categoryItems.Length - midPoint
            if (midPoint > maxLines)
                maxLines := midPoint
            loop maxLines {
                leftText := ""
                rightText := ""
                if (A_Index <= midPoint)
                    leftText := PadString(categoryItems[A_Index].text, columnWidth)
                else
                    leftText := PadString("", columnWidth)
                rightIdx := A_Index + midPoint
                if (rightIdx <= categoryItems.Length)
                    rightText := categoryItems[rightIdx].text
                else
                    rightText := ""
                displayText .= leftText . columnSpacing . rightText . "`n"
            }
            displayText .= "`n"
        }
    }

    displayText .= "Press Escape to close."
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
    g_HotstringSelectorGui.Add("Text", "w" . textControlWidth . " Center", "Hotstring Shortcuts")
    g_HotstringSelectorGui.Add("Text", "w" . textControlWidth . " h1 Background45475A")
    g_HotstringSelectorGui.SetFont("s" . fontSize . " cCDD6F4", "Segoe UI")

    ; Enable vertical scrolling for long content
    g_HotstringSelectorGui.AddEdit("w" . textControlWidth . " h" . textControlHeight .
        " ReadOnly VScroll Background1E1E2E", displayText
    )
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

    ; Clear handlers array
    g_HotstringHotkeyHandlers := []

    ; Enable hotkeys for all assigned characters (hotstrings, quick open files, macros),
    ; plus the reserved modifier key 'L' (Gemini redirect mode).
    for char in g_HotstringCharSequence {
        if (char = "l" || char = "L" || g_HotstringCharMap.Has(char) || g_QuickOpenFileCharMap.Has(char) ||
        g_MacroCharMap.Has(char)) {
            ; Use factory function to create handler with properly captured char value
            handler := CreateHotstringCharHandler(char)

            ; Store handler for cleanup
            g_HotstringHotkeyHandlers.Push({ char: char, handler: handler })

            ; Enable hotkey (both uppercase and lowercase)
            try {
                ; Handle special characters that need VK codes
                if (char = ",") {
                    Hotkey("vkBC", handler, "On")  ; VK code for comma
                } else if (char = ".") {
                    Hotkey("vkBE", handler, "On")  ; VK code for period
                } else {
                    Hotkey(char, handler, "On")
                    ; Also enable uppercase for lowercase letters
                    if (RegExMatch(char, "^[a-z]$")) {
                        Hotkey(StrUpper(char), handler, "On")
                    }
                }
            } catch {
                ; Silently ignore if we can't create hotkey for this character
            }
        }
    }

    ; Enable Escape hotkey
    Hotkey("Escape", HandleHotstringEscape, "On")
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

; Ctrl+Alt+Win+L - Same as Win+Alt+Shift+U then L,L: banner, 4s delay, N to cancel; then open Gemini + paste first snippet (+ Enter unless cancelled)
^!#L:: GeminiDelayedSubmitFlow()

; Ctrl+Alt+Win+4 - Send AI Text Optimizer prompt to Gemini (same as Win+Alt+Shift+U then L, 4)
^!#4:: GeminiNavigateFocusAndPasteFirstSnippet(GetAioptPromptText())

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

; Helper to send Escape while respecting dictation state.
; When dictation is active, this becomes a no-op so Handy is not closed.
SendEscape(count := 1) {
    global g_DictationActive
    if (g_DictationActive) {
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

; Force hook-based hotkey to intercept Escape at a lower level
; This should catch it before handy.exe's hook processes it
#UseHook
#InputLevel 10
Escape::
{
    ; Use state-based blocking: check g_DictationActive instead of checking window each time
    ; This ensures Esc remains restricted for the entire duration of dictation
    global g_DictationActive

    ; Block Escape if dictation is active (state-based, no timeout)
    ; This restriction remains for the entire duration of dictation
    if (g_DictationActive) {
        ; Block Escape from reaching handy.exe - do nothing
        ; This prevents the dictation software from closing
        ; Restriction remains active for entire dictation duration (no timeout)
        return
    }

    ; Otherwise, forward Escape to the system
    ; Use SendInput for more reliable key forwarding
    SendInput "{Escape}"
}
#InputLevel 0

; =============================================================================
; Dictation Indicator - Red Pulsing Square
; Shows a red pulsing square at the top center of the active monitor when
; dictation is active with handy.exe. Toggles with Win+Alt+Shift+0.
; =============================================================================

; Global variables for dictation indicator
global g_DictationActive := false
global g_DictationIndicatorGui := false
global g_DictationIndicatorText := false  ; Text control for status messages
global g_DictationPulseTimer := false
global g_DictationCheckTimer := false  ; Timer to check if Recording window still exists
global g_DictationPulseDirection := 1  ; 1 = fading in, -1 = fading out
global g_DictationPulseOpacity := 128  ; Current opacity (50-255)
global g_LastActiveMonitor := 0  ; Track last monitor to detect changes
global g_DictationCompletionChimeScheduled := false  ; Flag to prevent multiple completion chimes
global g_LastDictationSoundTick := 0  ; Timestamp of last dictation sound to throttle audio output
global g_DictationStartSound := A_ScriptDir . "\sounds\speach-start.wav"
global g_DictationStopSound := A_ScriptDir . "\sounds\speach-finished.wav"
global g_DictationLoopSound := A_ScriptDir . "\sounds\retro1.wav"
global g_PendingDictationAction := ""  ; Action to execute after transcription: "Paste" or "PasteEnter"
global g_PendingDictationMerge := false  ; Flag to trigger merge countdown after transcription completes
global g_PendingGeminiPromptAfterDictation := false  ; When set by ~#!+0 stop, show "Send to Gemini? Y (4s)" after completion
global g_DictationGeminiConfirmBannerVisible := false  ; Guard: only one "Send to Gemini?" banner at a time
global g_KeepIndicatorVisible := false  ; Flag to keep indicator visible until paste action completes
global g_LastStateTransitionTick := 0  ; Timestamp of last state transition to prevent rapid re-detection
global g_DictationSoundPlayed := false  ; Atomic test-and-set: one start chime per session
global g_DictationStartClipboardText := "" ; Track clipboard content at start to detect changes

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
global DICTATION_SQUARE_SIZE := 150  ; 3x bigger (was 50)
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

; Show or update the dictation indicator at the top center of the active monitor
ShowDictationIndicator() {
    global g_DictationIndicatorGui, g_DictationPulseOpacity, g_LastActiveMonitor
    global DICTATION_SQUARE_SIZE

    ; Get active monitor bounds
    monitorIndex := GetDictationActiveMonitor()
    MonitorGet(monitorIndex, &monitorLeft, &monitorTop, &monitorRight, &monitorBottom)
    monitorWidth := monitorRight - monitorLeft

    ; Calculate position: top center of monitor with small offset from top
    squareX := monitorLeft + (monitorWidth - DICTATION_SQUARE_SIZE) // 2
    squareY := monitorTop + 20  ; 20 pixels from top

    ; Check if indicator already exists
    if (IsObject(g_DictationIndicatorGui) && g_DictationIndicatorGui.Hwnd) {
        ; Indicator exists - just reposition it if monitor changed
        if (monitorIndex != g_LastActiveMonitor) {
            g_DictationIndicatorGui.Show("NA x" . squareX . " y" . squareY . " w" . DICTATION_SQUARE_SIZE . " h" .
                DICTATION_SQUARE_SIZE)
            g_LastActiveMonitor := monitorIndex
        }
        ; Clear any status text when starting new dictation
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
    g_DictationIndicatorGui.BackColor := "FF0000"  ; Red

    ; Add text control for status messages (centered, white text, bold)
    g_DictationIndicatorGui.SetFont("s14 cFFFFFF Bold", "Segoe UI")
    g_DictationIndicatorGui.MarginX := 10
    g_DictationIndicatorGui.MarginY := 10
    g_DictationIndicatorText := g_DictationIndicatorGui.Add("Text", "w" . (DICTATION_SQUARE_SIZE - 20) . " Center", "")

    ; Reset pulse opacity
    g_DictationPulseOpacity := 128
    g_LastActiveMonitor := monitorIndex

    ; Show the indicator without activating it
    g_DictationIndicatorGui.Show("NA x" . squareX . " y" . squareY . " w" . DICTATION_SQUARE_SIZE . " h" .
        DICTATION_SQUARE_SIZE)
    WinSetTransparent(g_DictationPulseOpacity, g_DictationIndicatorGui)
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
    global g_DictationIndicatorGui, g_DictationIndicatorText

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
    global g_DictationCompletionChimeScheduled, g_PendingDictationAction, g_PendingDictationMerge,
        g_KeepIndicatorVisible, g_PendingGeminiPromptAfterDictation
    global g_DictationLoopActive

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

        ; Execute pending action if one was set (from Win+Alt+Shift+J or 7)
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
        } else if (pendingAction = "PasteEnter") {
            ; Update indicator text to show status
            UpdateDictationIndicatorText("Pasting & Submitting...")
            ; Execute paste command (same as #!+7)
            Send "^v"
            ; Small delay between paste and enter for reliability
            Sleep 50
            ; Execute enter command to submit
            Send "{Enter}"
            ; Wait for both paste and enter to complete before hiding indicator
            Sleep 100  ; Small delay to ensure paste and enter completes
            ; Hide indicator only after both commands complete
            HideDictationIndicator()
            g_KeepIndicatorVisible := false
        }

        ; Check if merge countdown should start after transcription completes
        ; This ensures AI transcription and handy.exe finish before Clip Angel merge begins
        pendingMerge := g_PendingDictationMerge
        g_PendingDictationMerge := false  ; Clear immediately after reading

        if (pendingMerge) {
            ; Transcription is complete, now safe to start merge countdown
            DictationMerge_StartCountdown(5)
        }

        ; If user stopped dictation with Win+Alt+Shift+0 (no Paste/PasteEnter), show Gemini confirm banner (once only).
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
            DebugBannerLog("Utils.ahk:PlayDictationCompletionChime", "Calling ShowAndWait", "pendingAction empty", "H3"
            )
            ; #endregion
            DictationGeminiConfirm_ShowAndWait()
        }

        ; Trigger next loop iteration if active (module or legacy)
        if (InfiniteDictation.IsActive) {
            InfiniteDictation.OnTranscriptionComplete()
        } else if (g_DictationLoopActive) {
            SetTimer(DictationLoopStart, -2000)
        }
    }
}

; Check Recording window (handy.exe) and update indicator; play start chime when detected.
CheckDictationRecordingWindow() {
    global g_DictationActive, g_DictationCompletionChimeScheduled, g_LastStateTransitionTick, g_DictationStartSound,
        g_DictationSoundPlayed, g_DictationStartClipboardText

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
                    RunWait "powershell.exe -ExecutionPolicy Bypass -File `"" micVolumeScript "`"", , "Hide"
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
        global g_KeepIndicatorVisible
        if (!g_KeepIndicatorVisible) {
            HideDictationIndicator()
        }

        ; Check if clipboard has already changed (Handy might have updated it before window closed)
        currentClip := ""
        try {
            currentClip := A_Clipboard
        }

        if (currentClip != g_DictationStartClipboardText) {
            ; Clipboard already updated, trigger sound immediately
            PlayDictationCompletionChime()
        } else {
            ; Clipboard not yet updated, wait for change
            OnClipboardChange(DictationClipboardHandler)
            ; Set fallback timer (reduced to 1.5s)
            SetTimer(PlayDictationCompletionChime, -1500)
        }
    }
    else if (g_DictationActive && windowExists) {
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
    static lastHotkeyTick := 0
    static isProcessing := false

    ; Skip when script sends #!+0 programmatically (Infinite Dictation stop/start, #!+7 stop, etc.)
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

    ; If Infinite Dictation is active, treat as interrupt (same as Win+Alt+Shift+7)
    ; Logic gate: Only allow termination during Recording state; block during Transcribing state
    if (InfiniteDictation.IsActive) {
        if (!WinExist("Recording ahk_exe handy.exe")) {
            ; Transcribing - block termination to prevent interrupting active transcription
            isProcessing := false
            return
        }
        ; Recording - allow termination
        InfiniteDictation.Stop()
        isProcessing := false
        return
    }

    KeyWait("0", "L")

    if (!g_DictationActive) {
        g_DictationActive := true
        g_LastStateTransitionTick := A_TickCount
        ShowDictationIndicator()
        StartDictationPulseTimer()
        ; Sound: monitoring loop plays when window detected (zero latency)

        try {
            micVolumeScript := A_ScriptDir "\scripts\Set-MicVolume.ps1"
            if (FileExist(micVolumeScript))
                RunWait "powershell.exe -ExecutionPolicy Bypass -File `"" micVolumeScript "`"", , "Hide"
        } catch {
        }
    }

    ; User was stopping dictation (had been active when they pressed key) -> show Gemini confirm after completion
    if (dictationWasActiveOnKeyPress) {
        g_PendingGeminiPromptAfterDictation := true
        ; #region agent log
        DebugBannerLog("Utils.ahk:~#!+0", "Set pending Gemini flag", "dictationWasActiveOnKeyPress=1", "H1")
        ; #endregion
    }

    ToggleDictationMode()
    isProcessing := false
}

; Infinite Dictation module (state and loop logic)
#Include "Lib\InfiniteDictation.ahk"

; Win+Alt+Shift+7 is defined in Gemini.ahk (TTS from selection: repeat exactly + read aloud).

; Dictation with paste and submit action - Win+Alt+Shift+J
; Step 1: Programmatically stop dictation (send Win+Alt+Shift+0)
; Step 2: Wait for transcription to complete
; Step 3: Execute paste and enter action
#!+j::
{
    global g_PendingDictationAction, g_DictationActive, g_KeepIndicatorVisible, g_ProgrammaticDictationStop

    ; Play sound signal
    if (IsSoundEnabled()) {
        SoundPlay(A_ScriptDir . "\sounds\retro4.wav")
    }

    ; Only proceed if dictation is currently active
    if (g_DictationActive) {
        ; Set pending action to execute after transcription completes
        g_PendingDictationAction := "PasteEnter"
        ; Keep indicator visible until paste and submit completes
        g_KeepIndicatorVisible := true
        ; Programmatically send Win+Alt+Shift+0 to stop dictation
        ; Use SendInput for reliable key sending
        g_ProgrammaticDictationStop := true
        SendInput "#!+0"
    }
}

; Start the check timer automatically when script loads
; This continuously monitors for the Recording window and updates indicator position
StartDictationCheckTimer()