;---------------------------------------- Github ----------------------------------------------------
; https://github.com/duducm2/zmk-sofle/blob/main/keymap-drawer/eyelash_sofle.svg
;---------------------------------------- Alt + Shift + keys ----------------------------------------
;     Alt + Shift + Q : Power toys command palette
;     Alt + Shift + W : Open file explorer
;     Alt + Shift + Z : Bookmarks
;     Alt + Shift + S : Windows
;     Alt + Shift + A : Web
;---------------------------------------- Layer 2 + Keys ----------------------------------

; Layer 2 + 0: Ctrl + Home
; Layer 2 + ESC: Ctrl + End
; Layer 2 + P: Shift + Home
; Layer 2 + ´: Shift + End
; Layer 2 + 8: Ctrl + Right arrow (5 times)
; Layer 2 + 9: Ctrl + Left arrow (5 times)
; Layer 2 + 6: Ctrl + Shift + Home
; Layer 2 + 7: Ctrl + Shift + End
; Layer 2 + y: Ctrl + Shift + Page Up
; Layer 2 + u: Ctrl + Shift + Page Down
; Layer 2 + i: Ctrl + Page Up
; Layer 2 + o: Ctrl + Page Down

;---------------------------------------- Win + Alt + Shift + keys ----------------------------------
; ----- Available -----------
;     Win + Alt + Shift + 1 : Available
;     Win + Alt + Shift + 2 : Available
;     Win + Alt + Shift + L : Available
;     Win + Alt + Shift + 8 : Available
;     Win + Alt + Shift + 9 : Available
;     Win + Alt + Shift + Y : Available
;     Win + Alt + Shift + I : Available
;     Win + Alt + Shift + O : Available
;     Win + Alt + Shift + P : Available
;     Win + Alt + Shift + ç : Available
;     Win + Alt + Shift + J : Available
;     Win + Alt + Shift + K : Available
; ----- Onenote -------------
;     Win + Alt + Shift + N : Opens or activates OneNote
; ----- Spotify -------------
;     Win + Alt + Shift + S : Opens or activates Spotify
;     Win + Alt + Shift + D : Spotify - Go to library
; ----- ChatGPT -------------
;     Win + Alt + Shift + 6 : Get word pronunciation, definition, and Portuguese translation from ChatGPT
;     Win + Alt + Shift + 3 : Toggles the microphone (ChatGPT)
;     Win + Alt + Shift + T : Clicks on the last microphone icon in the ChatGPT
;     Win + Alt + Shift + 5 : Check grammar and improve text in both English and Portuguese
;     Win + Alt + Shift + W : Clear ChatGPT input field for new conversation
;     Win + Alt + Shift + 7 : Speak with ChatGPT
;     Win + Alt + Shift + C : Opens ChatGPT
;     Win + Alt + Shift + B : Copy last ChatGPT prompt
; ----- Youtube -------------
;     Win + Alt + Shift + 4 : Activates Youtube
; ----- Google --------------
;     Win + Alt + Shift + G : Opens Google
; ----- Outlook --------------
;     Win + Alt + Shift + A : Outlook - Open mail
;     Win + Alt + Shift + R : Outlook - Open Reminder
;     Win + Alt + Shift + X : Outlook - Open calendar
; ----- Microsoft Teams ------
;     Win + Alt + Shift + V : Open chat
;     Win + Alt + Shift + M : Microsoft Teams - New conversation
;     Win + Alt + Shift + Q : Microsoft Teams meeting - Toggle Mute
;     Win + Alt + Shift + F : Microsoft Teams meeting - Toggle camera
;     Win + Alt + Shift + E : Microsoft Teams meeting - Screen share
;     Win + Alt + Shift + U : Microsoft Teams meeting - Exit meeting
; ----- Hunt and Peck --------
;     Win + Alt + Shift + 0 : Hunt and Peck
; ----- WhatsApp ---------
;     Win + Alt + Shift + Z : Opens WhatsApp
; ----- Windows ---------
;     Win + Alt + Shift + ] : Minimizes windows
; ----- General ---------
;     Win + Alt + Shift + H : Jump mouse on the middle
;---------------------------------------- Shift + keys ----------------------------------------------
; ----- You can have repeated keys, depending on the software. -----------
;----- General --------------
;     Shift + Y : ?
; ----- Microsoft Teams -----
;     Shift + J : Mark as unread
;     Shift + U : Like
;     Shift + I : Heart
;     Shift + O : Laugh
;     Shift + E : Edit message
;     Shift + R : Reply Message
;     Shift + L : Remove pin
;     Shift + K : Pin
; ----- Onenote -------------
;     Shift + D : Delete
;     Shift + U : Onenote: select line and children
;     Shift + I : Onenote: collapse current
;     Shift + O : Onenote: expand current
;     Shift + K : Onenote: collapse all
;     Shift + L : Onenote: expand all
; ----- WhatsApp -------------
;     Shift + H : What's app - Go to current chat
;     Shift + J : Pin
;     Shift + K : Make it unread
; ----- Google Chrome --------
;     Shift + G : Pop up the current tab
; ----- Outlook --------------
;     Shift + U: Send to general
;     Shift + I: Send to newsletter

;---------------------------------------- Changed Shortcuts in the software itself ------------------
;     Ctrl + Shift + 8 : Fold/unfold (VS Code)
;     Ctrl + Shift + 9 : Fold/unfold (VS Code)
;     Ctrl + Shift + W : Remove unecessary keys (this closes some softwares and was causing issues)

;---------------------------------------- Scripts ---------------------------------------------------

#Include env.ahk

if (IS_WORK_ENVIRONMENT) {

    Run "C:\Users\fie7ca\Documents\HuntAndPack\hap.exe"
    Run "C:\Users\fie7ca\Documents\Caffeine\caffeine64.exe"
    Run "C:\Users\fie7ca\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\apps do Chrome\Mobills"
    Run "C:\Users\fie7ca\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\apps do Chrome\Gmail"
    Run "c:\Users\fie7ca\Documents\Settle Up.lnk"

    Run GetScriptPath("Microsoft Teams - meeting shortcuts.ahk")
    Run GetScriptPath("Outlook - Open mail.ahk")
    Run GetScriptPath("Outlook - Open Reminder.ahk")
    Run GetScriptPath("Outlook - Open calendar.ahk")
    Run GetScriptPath("Microsoft Teams - New conversation.ahk")
} else {
    Run "C:\Users\eduev\OneDrive\Documentos\HuntAndPeck\hap.exe"
}

; Open ChatGPT
Send "^!c"

Run GetScriptPath("Shift keys.ahk")
Run GetScriptPath("Remove unecessary keys.ahk")
Run GetScriptPath("Open Desktop.ahk")
Run GetScriptPath("OneNote - Open.ahk")
Run GetScriptPath("Spotify - Open.ahk")
Run GetScriptPath("Spotify - Go to library.ahk")
Run GetScriptPath("ChatGPT - Pronunciation.ahk")
Run GetScriptPath("ChatGPT - Toggle microphone.ahk")
Run GetScriptPath("ChatGPT - Click last microphone.ahk")
Run GetScriptPath("Youtube - Activate.ahk")
Run GetScriptPath("ChatGPT - Check for grammar.ahk")
Run GetScriptPath("ChatGPT - Speak.ahk")
Run GetScriptPath("Open ChatGPT.ahk")
Run GetScriptPath("ChatGPT - Copy last message.ahk")
Run GetScriptPath("Open Google.ahk")
Run GetScriptPath("Open WhatsApp.ahk")
Run GetScriptPath("Hunt and Peck.ahk")
Run GetScriptPath("Minimize.ahk")
Run GetScriptPath("Jump mouse on the middle.ahk")