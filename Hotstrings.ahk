#Requires AutoHotkey v2.0+

; Central registry of hotstrings for cheat-sheet display
global g_hotstrings := []

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

; ----------------------
; Define hotstrings below
; ----------------------

; --- Project Names ---
::myl::My Links
RegisterHotstring("::myl::", "My Links")

::gintegra::GS_UX core team_UX and CIP Integration
RegisterHotstring("::gintegra::", "GS_UX core team_UX and CIP Integration")

::gdash::GS_E&S_CIP Dashboard research and design
RegisterHotstring("::gdash::", "GS_E&S_CIP Dashboard research and design")

::gb2c::GS_B2C_Credit_Management_Strategy_UI_Mentoring
RegisterHotstring("::gb2c::", "GS_B2C_Credit_Management_Strategy_UI_Mentoring")

::gug::GS_UX Core Team_Monitoring for B2C in Brazil
RegisterHotstring("::gug::", "GS_UX Core Team_Monitoring for B2C in Brazil")

::gpm::GS_UX_Project_Management_Activities_LA
RegisterHotstring("::gpm::", "GS_UX_Project_Management_Activities_LA")

::guxcip::GS_UX_and_CIP
RegisterHotstring("::guxcip::", "GS_UX_and_CIP")

::gtrain::GS_UX core team_Trainings Management
RegisterHotstring("::gtrain::", "GS_UX core team_Trainings Management")

::gbp::GS_B2C_Portals and Key Accounts Process POC
RegisterHotstring("::gbp::", "GS_B2C_Portals and Key Accounts Process POC")

::cgrammar:: Coorect grammar and spelling
RegisterHotstring("::cgrammar::", "Coorect grammar and spelling")