; =============================================================================
; TestStudyLinkApi.ahk — Functional test for StudyLink API (GET + SET)
; Target: AIB — validates Google Apps Script endpoint integration
;
; Output follows standard_information_display.md protocol:
;   ✅ Green (BANNER_ACCENT_SUCCESS) = test passed
;   ❌ Red   (BANNER_ACCENT_ERROR)   = test failed
; =============================================================================
#Requires AutoHotkey v2.0+
#SingleInstance Force
#Include %A_ScriptDir%\Utils.ahk

result := StudyLink_RunFunctionalTest()

ytSet := result["youtubeSetOk"] ? "✅ YouTube SET: PASSED" : "❌ YouTube SET: " . result["youtubeSetMsg"]
ytGet := result["youtubeGetOk"] ? "✅ YouTube GET: PASSED" : "❌ YouTube GET: " . result["youtubeGetMsg"]
artSet := result["articleSetOk"] ? "✅ Article SET: PASSED" : "❌ Article SET: " . result["articleSetMsg"]
artGet := result["articleGetOk"] ? "✅ Article GET: PASSED" : "❌ Article GET: " . result["articleGetMsg"]

allPassed := result["setOk"] && result["getOk"]

if (allPassed) {
    ShowCenteredOverlay_Utils(ytSet "`n" ytGet "`n" artSet "`n" artGet, 3500, BANNER_ACCENT_SUCCESS)
} else {
    ShowCenteredOverlay_Utils(ytSet "`n" ytGet "`n" artSet "`n" artGet, 4500, BANNER_ACCENT_ERROR)
}

FileAppend(
    "StudyLink API Test Results`n"
    "======================`n"
    "YouTube SET: " (result["youtubeSetOk"] ? "✅ PASSED" : "❌ FAILED") " (" result["youtubeSetMsg"] ")`n"
    "YouTube GET: " (result["youtubeGetOk"] ? "✅ PASSED" : "❌ FAILED") " (" result["youtubeGetMsg"] ")`n"
    "Article SET: " (result["articleSetOk"] ? "✅ PASSED" : "❌ FAILED") " (" result["articleSetMsg"] ")`n"
    "Article GET: " (result["articleGetOk"] ? "✅ PASSED" : "❌ FAILED") " (" result["articleGetMsg"] ")`n"
    "Overall: " (allPassed ? "✅ ALL PASSED" : "❌ SOME FAILED") "`n",
    "*"
)

ExitApp