; Load-only harness: surfaces #Warn LocalSameAsGlobal issues when Utils.ahk is included.

; Run/reload this script and fix every warning until the dialog is silent.

#Requires AutoHotkey v2.0

#Warn

#Include %A_ScriptDir%\..\..\vendor\UIA-v2\Lib\UIA.ahk

#Include %A_ScriptDir%\..\..\Utils.ahk

; Smoke: missing provider functions or unset Copilot needles surface here with #Warn.
_ := GetGlobalAIProviderLabel()
_ := UseCopilotWebForGlobalAI()
_ := CopilotWeb_UrlIsChat("")

; #region agent log

VerifyUtilsWarn_DebugLog(location, message, hypothesisId := "E") {

    try {

        logPath := A_ScriptDir . "\..\debug-ddf6aa.log"

        ts := A_TickCount

        FileAppend(
            Format(
                '{{"sessionId":"ddf6aa","runId":"load-check","hypothesisId":"{1}","location":"{2}","message":"{3}","timestamp":{4}}}`n',
                hypothesisId, location, message, ts),
            logPath, "UTF-8")

    }

}

; #endregion

VerifyUtilsWarn_DebugLog("Verify-UtilsWarn.ahk:load", "Utils included; auto-execute reached without abort", "E")

ToolTip("Verify-UtilsWarn: no LocalSameAsGlobal warnings on load.", 10, 10)

SetTimer(() => ToolTip(), -2500)