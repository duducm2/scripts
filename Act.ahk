;============ Github eyelash sofle ============================================================================================================================================================
; https://github.com/duducm2/zmk-sofle/blob/main/keymap-drawer/eyelash_sofle.svg

;============ Keysmap ========================================================================================================================================================================

;----- Win + Shift + keys -----
; Win + Shift + E : Opens Desktop

;----- Alt + Shift + keys -----
;     Alt + Shift + Q : Power toys command palette
;     Alt + Shift + W : Open file explorer
;     Alt + Shift + Z : Bookmarks
;     Alt + Shift + S : Windows
;     Alt + Shift + A : Web

;----- Win + Alt + Shift + keys -----
;     Win + Alt + Shift + N : Opens or activates OneNote
;     Win + Alt + Shift + S : Opens or activates Spotify
;     Win + Alt + Shift + 6 : Get word pronunciation, definition, and Portuguese translation from ChatGPT
;     Win + Alt + Shift + 3 : Toggles the microphone (ChatGPT)
;     Win + Alt + Shift + T : Clicks on the last microphone icon in the ChatGPT
;     Win + Alt + Shift + 4 : Activates Youtube
;     Win + Alt + Shift + 5 : Check grammar and improve text in both English and Portuguese
;     Win + Alt + Shift + W : Clear ChatGPT input field for new conversation
;     Win + Alt + Shift + 7 : Speak with ChatGPT
;     Win + Alt + Shift + C : Opens ChatGPT
;     Win + Alt + Shift + B : Copy last ChatGPT prompt
;     Win + Alt + Shift + G : Opens Google
;     Win + Alt + Shift + Z : Opens WhatsApp
;     Win + Alt + Shift + Up/Down/Left/Right : Fast scrolling
;     Win + Alt + Shift + 0 : Hunt and Peck
;     Win + Alt + Shift + ] : Minimizes windows
;     Win + Alt + Shift + L : Pop up the current tab
;     Win + Alt + Shift + D : Spotify - Go to library

;    ----- Work Environment Shortcuts -----
;     Win + Alt + Shift + 8 : Outlook - Send to general
;     Win + Alt + Shift + 9 : Outlook - Send to newsletter
;     Win + Alt + Shift + A : Outlook - Open mail
;     Win + Alt + Shift + R : Outlook - Open Reminder
;     Win + Alt + Shift + X : Outlook - Open calendar
;     Win + Alt + Shift + Y : Microsoft Teams - Mark as unread
;     Win + Alt + Shift + M : Microsoft Teams - New conversation
;     Win + Alt + Shift + I : Microsoft Teams - Laugh
;     Win + Alt + Shift + O : Microsoft Teams - Like
;     Win + Alt + Shift + P : Microsoft Teams - Heart
;     Win + Alt + Shift + ç : Microsoft Teams - Edit message
;     Win + Alt + Shift + H : Microsoft Teams - Reply Message
;     Win + Alt + Shift + J : Microsoft Teams - Pin
;     Win + Alt + Shift + K : Microsoft Teams - Remove pin

;----- Shift + keys -----
;     Shift + Y : ?
;     Shift + U : Onenote: select line and children
;     Shift + I : Onenote: collapse current
;     Shift + O : Onenote: expand current
;     Shift + K : Onenote: collapse all
;     Shift + L : Onenote: expand all
;     Shift + H : What's app - Go to current chat

;----- Other Shortcuts -----
;     Ctrl + Shift + 8 : Fold/unfold (VS Code)
;     Ctrl + Shift + 9 : Fold/unfold (VS Code)
;     Ctrl + Shift + W : Remove unecessary keys

#Include env.ahk

;============ softwares =======================================================================================================================================================================

if (IS_WORK_ENVIRONMENT) {
    ; Run "C:\Users\fie7ca\Documents\CapsLock Indicator\CLIv3-3.16.1.2.exe"
    Run "C:\Users\fie7ca\Documents\HuntAndPack\hap.exe"
    Run "C:\Users\fie7ca\Documents\Caffeine\caffeine64.exe"
    Run "C:\Users\fie7ca\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\apps do Chrome\Mobills"
    Run "C:\Users\fie7ca\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\apps do Chrome\Gmail"
    Run "C:\Users\fie7ca\Documents\HuntAndPack\hap.exe"
} else {
    Run "C:\Users\eduev\OneDrive\Documentos\HuntAndPeck\hap.exe"
}

; Open ChatGPT
Send "^!c"

; Shift + keys
Run GetScriptPath("Shift keys.ahk")

;============ Ctrl + Shift + keys =====================================================================================================================================================================
; Ctrl + Shift + W : Remove unecessary keys
Run GetScriptPath("Remove unecessary keys.ahk")

;============ Win + keys =====================================================================================================================================================================

; Shift + Win + E : Opens Desktop
Run GetScriptPath("Open Desktop.ahk")

;============ Alt + Shift + keys =============================================================================================================================================================
;Alt + Shift + Q: power toys command palette
;Alt + Shift + W: open file explorer
;Alt + Shift + Z: Bookmarks
; Alt + Shift + S: Windows
; Alt + Shift + A: Web

;============ Win + Alt + Shift + keys =================================================================================================================================================================

; Win + Alt + Shift + N : Opens or activates OneNote
Run GetScriptPath("OneNote - Open.ahk")

; Win + Alt + Shift + S : Opens or activates Spotify
Run GetScriptPath("Spotify - Open.ahk")

; Win + Alt + Shift + D : Spotify - Go to library
Run GetScriptPath("Spotify - Go to library.ahk")

; Win + Alt + Shift + 6 : Get word pronunciation, definition, and Portuguese translation from ChatGPT
Run GetScriptPath("ChatGPT - Pronunciation.ahk")

; Win + Alt + Shift + 3 : Toggles the
Run GetScriptPath("ChatGPT - Toggle microphone.ahk")

; Win + Alt + Shift + T : Click on the last microphone icon in ChatGPT
Run GetScriptPath("ChatGPT - Click last microphone.ahk")

; Win + Alt + Shift + 4 : Activates Youtube
Run GetScriptPath("Youtube - Activate.ahk")

; Win + Alt + Shift + 5 : Check grammar and improve text in both English and Portuguese
Run GetScriptPath("ChatGPT - Check for grammar.ahk")

; Win + Alt + Shift + 7: Speak with ChatGPT
Run GetScriptPath("ChatGPT - Speak.ahk")

; Win + Alt + Shift + C : Opens ChatGPT
Run GetScriptPath("Open ChatGPT.ahk")

; Win + Alt + Shift + B : Copy last ChatGPT prompt
Run GetScriptPath("ChatGPT - Copy last message.ahk")

; Win + Alt + Shift + G : Opens Google
Run GetScriptPath("Open Google.ahk")

; Win + Alt + Shift + Z : Opens WhatsApp
Run GetScriptPath("Open WhatsApp.ahk")

; Win + Alt + Shift + 0 : Hunt and Peck
Run GetScriptPath("Hunt and Peck.ahk")

; Win + Alt + Shift + ] : Minimizes windows
Run GetScriptPath("Minimize.ahk")

; Win + Alt + Shift + L : Pop up the current tab
Run GetScriptPath("Google Chrome - pop out tab.ahk")

if (!IS_WORK_ENVIRONMENT) {
    Run GetScriptPath("Overleaf - Pages interaction.ahk")
}

if (IS_WORK_ENVIRONMENT) {
    ; Win + Alt + Shift + 2
    ; Run "C:\Users\fie7ca\OneDrive - Bosch Group\07 - Scripts\Windows Explorer - Copy path.ahk" ❌

    ; Win + Alt + Shift + 8
    Run GetScriptPath("Outlook - Send to general.ahk")

    ; Win + Alt + Shift + 9
    Run GetScriptPath("Outlook - Send to newsletter.ahk")

    ; Win + Alt + Shift + A
    Run GetScriptPath("Outlook - Open mail.ahk")

    ; Win + Alt + Shift + R
    Run GetScriptPath("Outlook - Open Reminder.ahk")

    ; Win + Alt + Shift + X
    Run GetScriptPath("Outlook - Open calendar.ahk")

    ; Win + Alt + Shift + Y
    Run GetScriptPath("Microsoft Teams - Mark as unread.ahk")

    ; Win + Alt + Shift + M
    Run GetScriptPath("Microsoft Teams - New conversation.ahk")

    ; Win + Alt + Shift + I
    Run GetScriptPath("Microsoft Teams - Laugh.ahk")

    ; Win + Alt + Shift + O
    Run GetScriptPath("Microsoft Teams - Like.ahk")

    ; Win + Alt + Shift + P
    Run GetScriptPath("Microsoft Teams - Heart.ahk")

    ; Win + Alt + Shift + ç
    Run GetScriptPath("Microsoft Teams - Edit message.ahk")

    ; Win + Alt + Shift + H
    Run GetScriptPath("Microsoft Teams - Reply Message.ahk")

    ; Win + Alt + Shift + J
    Run GetScriptPath("Microsoft Teams - Pin.ahk")

    ; Win + Alt + Shift + K
    Run GetScriptPath("Microsoft Teams - Remove pin.ahk")

}
