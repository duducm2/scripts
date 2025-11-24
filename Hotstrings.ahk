#Requires AutoHotkey v2.0+
#SingleInstance Force

; ----------------------
; Safer hotstrings core
; ----------------------
global g_hotstrings := []
global g_lastExpansion := 0

; Trigger only on Space or Tab, not Enter or punctuation
Hotstring("EndChars", " `t")

RegisterHotstring(trigger, expansion) {
    global g_hotstrings
    g_hotstrings.Push({ trigger: trigger, expansion: expansion })
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

    ; Send arrow right then left after pasting
    Send "{Left}"
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

:o:cagent::
{
    InsertText(
        "Continue your browsing. Check for missing radio buttons. Answer everything till you get to the last phase in the TrustMate website."
    )
}

:o:cagentquest::
{
    InsertText("These questions are not fulfilled in the questionnaire. Go back and answer them.")
}

:o:cagentall::
{
    InsertText("Go over the entire form and answer all the questions that are missing.")
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
        "( LTrim`nFood_Log dictation → Excel CSV`n`nROLE`nYou transcribe my meal dictation (PT/EN) into rows for my Excel Food_Log.`n`nHOW IT WORKS`n- I will dictate one or more meals in free speech.`n- Process immediately without asking questions.`n`nOUTPUT (strict)`n- Return ONLY CSV data rows. NO header row. NO markdown, NO code fences, NO commentary.`n- Each row format: Date;Meal;Time;Main_Items;Tags;Satisfaction_with_Speech;Notes`n- Sort by Date then Time.`n`nFIELD RULES`n- Date: YYYY‑MM‑DD. Use " "today" " for current date in America/Sao_Paulo timezone.`n- Time: HH:MM in 24h; pad leading zeros (e.g., 08:05).`n- Meal: Breakfast | Lunch | Dinner | Snack.`n  PT mapping: café da manhã→Breakfast; almoço→Lunch; jantar/janta→Dinner; lanche→Snack.`n- Main_Items: comma‑separated simple item names (e.g., coffee, bread, butter).`n- Tags: comma‑separated, from this set when present or inferable:`n  caffeine, sugar, alcohol, dairy, gluten, fried, spicy, high-carb, low-carb, processed, protein, fiber, late-night, home-cooked, fast-food.`n  Add " "late-night" " automatically if Time ≥ 22:00.`n- Satisfaction_with_Speech: integer 0–3 (0 = liked a lot; 3 = disliked a lot). If not stated, leave empty.`n- Notes: short free text when I provide context.`n`nMISSING INFO`n- If Date or Time is missing, use " "today" " and infer time from meal type (Breakfast=08:00, Lunch=12:00, Dinner=19:00, Snack=15:00).`n`nACK`n- Process the dictation immediately and output CSV rows only.`n For the date always consider one day before the current one.)"
    )
}

:o:aiopt::
{
    InsertText(
        "( LTrim`nTask: Rewrite the input text so it becomes AI-oriented.`n`nGoal: Produce a version that is concise, unambiguous, free of redundancy, and easy for an AI to parse.`n`nContext: The input text contains instructions and requests intended for a second AI (AIB). You must preserve ALL important information, especially any instructions, requests, or requirements meant for AIB. Do not remove information in the name of clarity or conciseness.`n`nInstructions:`n`n1. Preserve all essential information, especially instructions and requests for the second AI.`n2. Use positive, direct instructions.`n3. Maintain consistent terminology and simple syntax.`n4. Resolve ambiguity and clarify references.`n5. Output in a clean, structured format with no extra commentary.`n6. Do NOT omit any important information, requests, or instructions that the user provided for the second AI.`n7. Do NOT bias the prompt with your own concerns, interpretations, or modifications. Process the text as-is without adding your own perspective or concerns.`n`nInput: <insert text here>`n`nOutput:`nCRITICAL: The output must contain ONLY the processed result with no additional text, commentary, explanations, or formatting. The output must be ready for direct copy and paste without any modification.`n`nA rewritten version of the input text that is optimized for AI interpretation and contains:`n`n* Clear meaning`n* No repeated ideas`n* No filler wording`n* No contradictions`n* Stable terminology`n* Straightforward sentence structure`n* ALL important information preserved, especially instructions for the second AI)"
    )
}

; ----------------------
; Register hotstrings for cheat sheet display
; ----------------------
InitHotstringsCheatSheet() {
    RegisterHotstring(":o:myl", "My Links")
    RegisterHotstring(":o:gintegra", "GS_UX core team_UX and CIP Integration")
    RegisterHotstring(":o:gdash", "GS_E&S_CIP Dashboard research and design")
    RegisterHotstring(":o:gb2c", "GS_B2C_Credit_Management_Strategy_UI_Mentoring")
    RegisterHotstring(":o:gug", "GS_UX Core Team_Monitoring for B2C in Brazil")
    RegisterHotstring(":o:gpm", "GS_UX_Project_Management_Activities_LA")
    RegisterHotstring(":o:guxcip", "GS_UX_and_CIP")
    RegisterHotstring(":o:gtrain", "GS_UX core team_Trainings Management")
    RegisterHotstring(":o:gbp", "GS_B2C_Portals and Key Accounts Process POC")
    RegisterHotstring(":o:cgrammar",
        "Correct grammar and spelling. Remove any dashes from the text. The text should be plain with no styles. Give back only the text. Use linebreaks and a space between paragraphs and look like a human."
    )
    RegisterHotstring(":o:cagent",
        "Continue your browsing. Check for missing radio buttons. Answer everything till you get to the last phase in the TrustMate website."
    )
    RegisterHotstring(":o:cagentquest",
        "These questions are not fulfilled in the questionnaire. Go back and answer them.")
    RegisterHotstring(":o:cagentall", "Go over the entire form and answer all the questions that are missing.")
    RegisterHotstring(":o:ebosch", "eduardo.figueiredo@br.bosch.com")
    RegisterHotstring(":o:egoogle", "edu.evangelista.figueiredo@gmail.com")
    RegisterHotstring(":o:mtask",
        "This is a message, summary, text or any textual information that translates into a task for me to do. Translate this way, into a task, make informative and start with the emoji 🔲. Make it clear and consise."
    )
    RegisterHotstring(":o:flog", "Food_Log dictation → Excel CSV prompt (with Speech rating)")
    RegisterHotstring(":o:aiopt", "AI-Optimized text rewriting prompt")
}
InitHotstringsCheatSheet()

; ------------
; Optional: scope Explorer-only hotstrings used for renaming
; Uncomment to restrict selected triggers to File Explorer or Save dialogs
;------------
;#HotIf WinActive("ahk_exe explorer.exe") || WinActive("ahk_class #32770")
;:o:gdash::
;    InsertText("GS_E&S_CIP Dashboard research and design")
;return
;#HotIf
