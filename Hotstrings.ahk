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
    InsertText("Correct grammar and spelling. Remove any dashes from the text.")
}

:o:cagent::
{
    InsertText("Continue your browsing. Check for missing radio buttons. Answer everything till you get to the last phase in the TrustMate website.")
}

:o:cagentquest::
{
    InsertText("These questions are not fulfilled in the questionnaire. Go back and answer them.")
}

:o:cagentall::
{
    InsertText("Go over the entire form and answer all the questions that are missing.")
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
    RegisterHotstring(":o:cgrammar", "Correct grammar and spelling. Remove any dashes from the text.")
    RegisterHotstring(":o:cagent", "Continue your browsing. Check for missing radio buttons. Answer everything till you get to the last phase in the TrustMate website.")
    RegisterHotstring(":o:cagentquest", "These questions are not fulfilled in the questionnaire. Go back and answer them.")
    RegisterHotstring(":o:cagentall", "Go over the entire form and answer all the questions that are missing.")
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
