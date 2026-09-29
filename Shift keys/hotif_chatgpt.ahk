; =============================================================================
; Shift keys module: hotif_chatgpt.ahk
; ChatGPT hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

;-------------------------------------------------------------------
; ChatGPT Shortcuts (Phase 2: O(1) predicate via IsChatGPTActiveForHotkey when USE_DAEMON_CONTEXT_CHATGPT)
;-------------------------------------------------------------------
#HotIf IsChatGPTActiveForHotkey()

; Shift + U : (reserved for later script)

; Shift + I: Toggle sidebar
+i:: Send("^+s")

; Shift + O : Re-send rules & ask ChatGPT to correct mistake
+o::
{
    ; Ensure composer is focused
    SendEscape()
    Sleep 150

    promptText := ""
    try promptText := FileRead(PROMPT_FILE, "UTF-8")
    if (StrLen(promptText) = 0)
        promptText := "[Prompt file missing]"

    msg :=
        "It seems you violated one of the conversation rules (e.g., incorrect name spelling). Read the rules below, identify your mistake, and reply ONLY with the corrected content." .
        "`n`n" . promptText

    oldClip := A_Clipboard
    A_Clipboard := ""
    A_Clipboard := msg
    ClipWait 1
    Send "^v"
    Sleep 100
    hwnd := GetChatGPTWindowHwnd()
    if (!hwnd)
        hwnd := WinExist("A")
    if (!AiCompanion_SendAndConfirm(hwnd, "chatgpt", (*) => Send("{Enter}"))) {
        A_Clipboard := oldClip
        return
    }
    A_Clipboard := oldClip

    Send "!{Tab}"
    WaitForButtonAndShowSmallLoading_ChatGPT(["Stop streaming", "Interromper transmissão"], "Waiting for response...")
}

; Shift + C: Copy last code block
+c:: Send("^+;")

; Shift + J: Go down
+j::
{
    Send "d"
    Sleep 50
    Send "{Backspace}"
    Sleep 50
    Send "+{Tab}"
    Sleep 50
    Send "{Enter}"
}

; Shift + L: Send and show AI banner
+l:: SubmitChatGPTMessage()

; Function to submit ChatGPT message and show AI banner
SubmitChatGPTMessage() {
    hwnd := GetChatGPTWindowHwnd()
    if (!hwnd)
        hwnd := WinExist("A")
    ; Escape focuses the composer before the snapshot, so a blank field is not a false send.
    SendEscape()
    Sleep 100
    if (!AiCompanion_SendAndConfirm(hwnd, "chatgpt", (*) => Send("{Enter}")))
        return
    Send "!{Tab}"
    Sleep 300
    ShowSmallLoadingIndicator_ChatGPT("AI is responding...")
    WaitForButtonAndShowSmallLoading_ChatGPT(["Stop streaming", "Interromper transmissão", "Stop", "Interromper"],
        "AI is responding...", 0)
}

#HotIf
