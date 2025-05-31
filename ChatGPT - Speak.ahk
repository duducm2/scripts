#Requires AutoHotkey v2.0+
#SingleInstance Force ; Good practice to add

#include UIA-v2\Lib\UIA.ahk
#include UIA-v2\Lib\UIA_Browser.ahk
#include %A_ScriptDir%\\env.ahk ; Include environment configuration

; Win+Alt+Shift+7 to toggle dictation to ChatGPT
#!+7::
{
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
    en_stopStreamingName := "Stop streaming" ; Corrected from "Stop streaming" to "Stop streaming"

    ; Select names based on environment (IS_WORK_ENVIRONMENT is true for Work/Portuguese)
    currentDictateName := IS_WORK_ENVIRONMENT ? pt_dictateName : en_dictateName
    currentSubmitName := IS_WORK_ENVIRONMENT ? pt_submitName : en_submitName
    currentTranscribingName := IS_WORK_ENVIRONMENT ? pt_transcribingName : en_transcribingName
    currentSendPromptName := IS_WORK_ENVIRONMENT ? pt_sendPromptName : en_sendPromptName
    currentStopStreamingName := IS_WORK_ENVIRONMENT ? pt_stopStreamingName : en_stopStreamingName

    ; Create regex patterns for names to handle both languages if IS_WORK_ENVIRONMENT is not strictly one or the other
    ; Or if we want to be more robust to slight variations.
    ; For now, we'll use the selected name. A more robust solution could be:
    dictateNamePattern := IS_WORK_ENVIRONMENT ? pt_dictateName : en_dictateName
    submitNamePattern := IS_WORK_ENVIRONMENT ? pt_submitName : en_submitName
    sendPromptNamePattern := IS_WORK_ENVIRONMENT ? pt_sendPromptName : en_sendPromptName
    ; For buttons that might show one name OR the other depending on state, we can use regex
    ; Example: dictatingOrSubmitPattern := "(" . pt_submitName . "|" . en_submitName . "|" . pt_transcribingName . "|" . en_transcribingName . ")"

    static isDictating := false ; State variable
    static submitFailCount := 0 ; Counter for consecutive submit failures

    ; Activate the window containing "chatgpt"
    SetTitleMatchMode 2
    WinActivate "chatgpt"
    WinWaitActive "ahk_exe chrome.exe" ; Ensure the correct window is active

    ; Initialize UIA Browser for the active window
    cUIA := UIA_Browser()
    Sleep 300 ; Wait a bit for the UI

    if !isDictating { ; If not currently dictating, prepare field, paste prompt and start it
        try {
            ; Find and click the Dictate button
            dictateBtn := cUIA.FindElement({ Name: currentDictateName, Type: "Button" })
            if dictateBtn {
                dictateBtn.Click()
                isDictating := true ; Update state
                submitFailCount := 0 ; Reset fail counter on successful start
            } else {
                MsgBox currentDictateName . " not found."
            }
        } catch Error as e {
            MsgBox "Error during pre-dictation or starting dictation: " e.Message
            isDictating := false ; Reset state on error
        }
    } else { ; If already dictating, stop it (click Submit dictation)
        try {
            ; Find the Submit dictation button using its Name and Type
            ; The button to stop dictation might be "Interromper ditado" or "Stop dictation" (currentTranscribingName)
            ; or "Enviar ditado" / "Submit dictation" (currentSubmitName) if it changes name.
            ; We will try currentSubmitName first, as per original logic, then currentTranscribingName as an alternative.

            submitBtn := cUIA.FindElement({ Name: currentSubmitName, Type: "Button" })
            if !submitBtn { ; If submitName is not found, try transcribingName
                submitBtn := cUIA.FindElement({ Name: currentTranscribingName, Type: "Button" })
            }

            if submitBtn {
                submitBtn.Click()
                isDictating := false ; Update state
                submitFailCount := 0 ; Reset fail counter on successful submit

                ; >>> NEW: Wait for Send prompt button and press Enter <<<
                try {
                    Sleep 200 ; Small pause after clicking submit
                    ; Wait for the button to reappear/enable (Timeout 10 seconds)
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
                ; >>> End NEW <<<

            } else {
                MsgBox currentSubmitName . " or " . currentTranscribingName . " button not found."
                submitFailCount++ ; Increment fail counter
            }
        } catch Error as e {
            MsgBox "Error finding or clicking Submit/Stop dictation button: " e.Message
            submitFailCount++ ; Increment fail counter
            ; Decide if state should be reset here? Maybe not, user might want to try again.
        }

        ; Check if submit failed too many times
        if submitFailCount >= 1 {
            MsgBox "Failed to submit dictation 1 time. Assuming dictation stopped. Press hotkey again to start."
            isDictating := false ; Reset state
            submitFailCount := 0 ; Reset counter
        }
    }
}
