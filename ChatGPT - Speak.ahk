#Requires AutoHotkey v2.0+
#SingleInstance Force ; Good practice to add

#include UIA-v2\Lib\UIA.ahk
#include UIA-v2\Lib\UIA_Browser.ahk
#include %A_ScriptDir%\\env.ahk ; Include environment configuration

; Win+Alt+Shift+7 to toggle dictation to ChatGPT
#!+7::
{
    ToggleDictationSpeak()
}

ToggleDictationSpeak(triedFallback := false, forceAction := "") {
    static isDictating := false ; State variable
    static submitFailCount := 0 ; Counter for consecutive submit failures

    ; Define button names for both languages
    pt_dictateName := "Botão de ditado"
    en_dictateName := "Dictate button"
    pt_submitName := "Enviar ditado"
    en_submitName := "Submit dictation"
    pt_transcribingName := "Interromper ditado"
    en_transcribingName := "Stop dictation"
    pt_sendPromptName := "Enviar prompt"
    en_sendPromptName := "Send prompt"
    pt_stopStreamingName := "Interromper transmissão"
    en_stopStreamingName := "Stop streaming"

    ; Select names based on environment (IS_WORK_ENVIRONMENT is true for Work/Portuguese)
    currentDictateName := IS_WORK_ENVIRONMENT ? pt_dictateName : en_dictateName
    currentSubmitName := IS_WORK_ENVIRONMENT ? pt_submitName : en_submitName
    currentTranscribingName := IS_WORK_ENVIRONMENT ? pt_transcribingName : en_transcribingName
    currentSendPromptName := IS_WORK_ENVIRONMENT ? pt_sendPromptName : en_sendPromptName
    currentStopStreamingName := IS_WORK_ENVIRONMENT ? pt_stopStreamingName : en_stopStreamingName

    ; Prepare arrays for FindButton
    dictateNames_to_find := [currentDictateName]
    submitOrStopNames_to_find := [currentSubmitName, currentTranscribingName]

    ; Activate the window containing "chatgpt"
    SetTitleMatchMode 2
    WinActivate "chatgpt"
    WinWaitActive "ahk_exe chrome.exe"
    cUIA := UIA_Browser()
    Sleep 300

    action := forceAction ? forceAction : (!isDictating ? "start" : "stop")

    if (action = "start") {
        try {
            dictateBtn := cUIA.FindElement({ Name: currentDictateName, Type: "Button" })
            if dictateBtn {
                dictateBtn.Click()
                isDictating := true
                submitFailCount := 0
                return
            } else if !triedFallback {
                ToggleDictationSpeak(true, "stop")
                return
            } else {
                MsgBox currentDictateName . " not found, and could not stop dictation either."
            }
        } catch Error as e {
            if !triedFallback {
                ToggleDictationSpeak(true, "stop")
                return
            } else {
                MsgBox "Error during pre-dictation or starting dictation: " e.Message
                isDictating := false
            }
        }
    } else if (action = "stop") {
        try {
            submitBtn := cUIA.FindElement({ Name: currentSubmitName, Type: "Button" })
            if !submitBtn {
                submitBtn := cUIA.FindElement({ Name: currentTranscribingName, Type: "Button" })
            }
            if submitBtn {
                submitBtn.Click()
                isDictating := false
                submitFailCount := 0
                try {
                    Sleep 200
                    finalSendBtn := cUIA.WaitElement({ Name: currentSendPromptName, AutomationId: "composer-submit-button" },
                    10000)
                    if finalSendBtn {
                        SendInput "{Enter}"
                    } else {
                        MsgBox "Timeout: '" . currentSendPromptName .
                            "' button did not reappear after submitting dictation."
                    }
                } catch Error as e_wait {
                    MsgBox "Error waiting for/clicking final " . currentSendPromptName . " button: " e_wait.Message
                }
                return
            } else if !triedFallback {
                ToggleDictationSpeak(true, "start")
                return
            } else {
                MsgBox currentSubmitName . " or " . currentTranscribingName .
                    " button not found, and could not start dictation either."
                submitFailCount++
            }
        } catch Error as e {
            if !triedFallback {
                ToggleDictationSpeak(true, "start")
                return
            } else {
                MsgBox "Error finding or clicking Submit/Stop dictation button: " e.Message
                submitFailCount++
            }
        }
        if submitFailCount >= 1 {
            MsgBox "Failed to submit dictation 1 time. Assuming dictation stopped. Press hotkey again to start."
            isDictating := false
            submitFailCount := 0
        }
    }
}
