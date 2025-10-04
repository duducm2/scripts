#Requires AutoHotkey v2.0+

; Prevent hotstring duplication by changing ending characters
#Hotstring EndChars `n `t

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
:*o:myl::My Links
RegisterHotstring(":*o:myl::", "My Links")

:*o:gintegra::GS_UX core team_UX and CIP Integration
RegisterHotstring(":*o:gintegra::", "GS_UX core team_UX and CIP Integration")

:*o:gdash::GS_E&S_CIP Dashboard research and design
RegisterHotstring(":*o:gdash::", "GS_E&S_CIP Dashboard research and design")

:*o:gb2c::GS_B2C_Credit_Management_Strategy_UI_Mentoring
RegisterHotstring(":*o:gb2c::", "GS_B2C_Credit_Management_Strategy_UI_Mentoring")

:*o:gug::GS_UX Core Team_Monitoring for B2C in Brazil
RegisterHotstring(":*o:gug::", "GS_UX Core Team_Monitoring for B2C in Brazil")

:*o:gpm::GS_UX_Project_Management_Activities_LA
RegisterHotstring(":*o:gpm::", "GS_UX_Project_Management_Activities_LA")

:*o:guxcip::GS_UX_and_CIP
RegisterHotstring(":*o:guxcip::", "GS_UX_and_CIP")

:*o:gtrain::GS_UX core team_Trainings Management
RegisterHotstring(":*o:gtrain::", "GS_UX core team_Trainings Management")

:*o:gbp::GS_B2C_Portals and Key Accounts Process POC
RegisterHotstring(":*o:gbp::", "GS_B2C_Portals and Key Accounts Process POC")

:*o:cgrammar:: Coorect grammar and spelling
RegisterHotstring(":*o:cgrammar::", "Coorect grammar and spelling")