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

::btw::by the way
RegisterHotstring("::btw::", "by the way")

; --- Project Names ---
::guc::GS_UX core team_UX and CIP Integration
RegisterHotstring("::guc::", "GS_UX core team_UX and CIP Integration")

::ged::GS_E&S_CIP Dashboard research and design
RegisterHotstring("::ged::", "GS_E&S_CIP Dashboard research and design")

::gbc::GS_B2C_Credit_Management_Strategy_UI_Mentoring
RegisterHotstring("::gbc::", "GS_B2C_Credit_Management_Strategy_UI_Mentoring")

::gum::GS_UX Core Team_Monitoring for B2C in Brazil
RegisterHotstring("::gum::", "GS_UX Core Team_Monitoring for B2C in Brazil")

::gup::GS_UX_Project_Management_Activities_LA
RegisterHotstring("::gup::", "GS_UX_Project_Management_Activities_LA")

::guc2::GS_UX_and_CIP
RegisterHotstring("::guc2::", "GS_UX_and_CIP")

::gut::GS_UX core team_Trainings Management
RegisterHotstring("::gut::", "GS_UX core team_Trainings Management")

::gub::GS_UX Bootcamp 2024 in Joinville
RegisterHotstring("::gub::", "GS_UX Bootcamp 2024 in Joinville")

::guc3::GS_UX core team_Project PD Calculator
RegisterHotstring("::guc3::", "GS_UX core team_Project PD Calculator")

::gba::GS_B2R_AA CF Automation
RegisterHotstring("::gba::", "GS_B2R_AA CF Automation")

::gpc::GS_P2P_Checklist for Activity Handover
RegisterHotstring("::gpc::", "GS_P2P_Checklist for Activity Handover")

::epc::EXT_PS_CCC Telemetry for light vehicles
RegisterHotstring("::epc::", "EXT_PS_CCC Telemetry for light vehicles")

::emp::EXT_M-PMQ_PPAP Application
RegisterHotstring("::emp::", "EXT_M-PMQ_PPAP Application")

::emu::EXT_M-PUR36_Tooling Management
RegisterHotstring("::emu::", "EXT_M-PUR36_Tooling Management")

::gbp::GS_B2C_Portals and Key Accounts Process POC
RegisterHotstring("::gbp::", "GS_B2C_Portals and Key Accounts Process POC")