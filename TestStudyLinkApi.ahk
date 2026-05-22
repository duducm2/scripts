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

setStatus := result["setOk"] ? "✅ SET: PASSED" : "❌ SET: FAILED"
getStatus := result["getOk"] ? "✅ GET: PASSED" : "❌ GET: FAILED"

allPassed := result["setOk"] && result["getOk"]

if (allPassed) {
    ShowCenteredOverlay_Utils(setStatus "`n" getStatus, 3000, BANNER_ACCENT_SUCCESS)
} else {
    setDetail := result["setOk"]
        ? "✅ SET: PASSED"
        : "❌ SET: " . (result.HasProp("setDetail") ? result.setDetail : "FAILED")
    getDetail := result["getOk"]
        ? "✅ GET: PASSED"
        : "❌ GET: " . (result.HasProp("getDetail") ? result.getDetail : "FAILED")
    ShowCenteredOverlay_Utils(setDetail "`n" getDetail, 4000, BANNER_ACCENT_ERROR)
}

FileAppend(
    "StudyLink API Test Results`n"
    "======================`n"
    "SET: " (result["setOk"] ? "✅ PASSED" : "❌ FAILED") " (" result["setMsg"] ")`n"
    "GET: " (result["getOk"] ? "✅ PASSED" : "❌ FAILED") " (" result["getMsg"] ")`n"
    "Overall: " (allPassed ? "✅ ALL PASSED" : "❌ SOME FAILED") "`n",
    "*"
)

ExitApp
