;---------------------------------------- Github ----------------------------------------------------
; https://github.com/duducm2/zmk-sofle/blob/main/keymap-drawer/eyelash_sofle.svg
;---------------------------------------- Layer 0 + Keys ----------------------------------
;    Power toys command palette
;    Open file explorer
;    Windows
;---------------------------------------- Layer 1 + Keys ----------------------------------
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
; Layer 2 + A: Go to desktop
;---------------------------------------- Win + Alt + Shift + keys ----------------------------------
; ----- Available -----------
; A C J K M Q U W X Y 3
; ----- Onenote -------------
;     Win + Alt + Shift + N : Opens or activates OneNote
; ----- Spotify -------------
;     Win + Alt + Shift + S : Opens or activates Spotify
;     Win + Alt + Shift + D : Spotify - Go to library
; ----- ChatGPT -------------
;     Win + Alt + Shift + L : Talk with ChatGPT through voice
;     Win + Alt + Shift + 8 : Get word pronunciation, definition, and Portuguese translation from ChatGPT
;     Win + Alt + Shift + 0 : Toggles the microphone
;     Win + Alt + Shift + 9 : Clicks on the last microphone icon in the ChatGPT
;     Win + Alt + Shift + O : Check grammar and improve text in both English and Portuguese
;     Win + Alt + Shift + 7 : Speak with ChatGPT
;     Win + Alt + Shift + I : Opens ChatGPT
;     Win + Alt + Shift + P : Copy last ChatGPT prompt
; ----- Youtube -------------
;     Win + Alt + Shift + H : Activates Youtube
; ----- Google --------------
;     Win + Alt + Shift + F : Opens Google
; ----- Outlook --------------
;     Win + Alt + Shift + B: Outlook - Open mail
;     Win + Alt + Shift + V : Outlook - Open Reminder
;     Win + Alt + Shift + G : Outlook - Open calendar
; ----- Microsoft Teams ------
;     Win + Alt + Shift + R : Microsoft Teams - New conversation
;     Win + Alt + Shift + 5 : Microsoft Teams meeting - Toggle Mute
;     Win + Alt + Shift + 4 : Microsoft Teams meeting - Toggle camera
;     Win + Alt + Shift + T : Microsoft Teams meeting - Screen share
;     Win + Alt + Shift + 2 : Microsoft Teams meeting - Exit meeting
;     Win + Alt + Shift + E : Select the chats window
;     Win + Alt + Shift + 3 : Select the meeting window
; ----- WhatsApp ---------
;     Win + Alt + Shift + Z : Opens WhatsApp
; ----- Windows ---------
;     Win + Alt + Shift + 6 : Minimizes windows
; ----- General ---------
;     Win + Alt + Shift + 1 : Jump mouse on the middle
; ----- Hunt and Peck ---------
;     Win + Alt + Shift + X : Activate hunt and Peck
;---------------------------------------- Shift + keys ----------------------------------------------
; ----- You can have repeated keys, depending on the software.
; ----- Prefered Keys sequences (most important first): Y U I O P H J K L N M , . 6 7 8 9 0 W E R T D F G C V B
; ----- Microsoft Teams -----
;     Shift + J : Mark as unread
;     Shift + Y : Like
;     Shift + U : Heart
;     Shift + I : Laugh
;     Shift + E : Edit message
;     Shift + R : Reply Message
;     Shift + L : Remove pin
;     Shift + K : Pin
; ----- Onenote -------------
;     Shift + D : Delete
;     Shift + Y : Onenote: select line and children
;     Shift + U : Onenote: collapse current
;     Shift + I : Onenote: expand current
;     Shift + J : Onenote: collapse all
;     Shift + K : Onenote: expand all
; ----- WhatsApp -------------
;     Shift + H : What's app - Go to current chat
;     Shift + J : Pin
;     Shift + K : Make it unread
;     Shift + Y : Send voice message
;     Shift + U: Search
;     Shift + I: Reply
;     Shift + O: Sticker panel
;     Shift + P: Toggle unred
; ----- Gmail --------
;     Shift + Y : Go to main inbox
;     Shift + U : Go to updates
;     Shift + I : Mark as read
;     Shift + O : Mark as unread
; ----- Google Chrome --------
;     Shift + G : Pop up the current tab
; ----- Outlook --------------
;     Shift + U: Send to general
;     Shift + I: Send to newsletter
; ----- ChatGPT --------------
;     Shift + Y: Select all and cut
;     Shift + U: Change model
;     Shift + I: Togle sidebar
;     Shift + O: Write chatgpt
;     Shift + P: New chat
;     Shift + H: Copy last code block
; ----- Cursor --------------
;     Shift + Y : Fold/unfold (Cursor)
;     Shift + U : Fold/unfold (Cursor)
; ----- Windows Explorer --------------
;     Shift + Y : Select first file
;     Shift + U : New folder
;     Shift + I : New Shortcut

;---------------------------------------- Changed Shortcuts in the software itself ------------------
;     Ctrl + Shift + W : Remove unecessary keys (this closes some softwares and was causing issues)

;---------------------------------------- Scripts ---------------------------------------------------

#Include env.ahk

if (IS_WORK_ENVIRONMENT) {

    Run "C:\Users\fie7ca\Documents\HuntAndPack\hap.exe"
    Run "C:\Users\fie7ca\Documents\Caffeine\caffeine64.exe"
    Run "C:\Users\fie7ca\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\apps do Chrome\Mobills"
    Run "C:\Users\fie7ca\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\apps do Chrome\Gmail"
    Run "c:\Users\fie7ca\Documents\Settle Up.lnk"
    Run "c:\ProgramData\Microsoft\Windows\Start Menu\Programs\OneNote.lnk"

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
Run GetScriptPath("ChatGPT - Talk trough voice.ahk")
Run GetScriptPath("ChatGPT - Check for grammar.ahk")
Run GetScriptPath("ChatGPT - Speak.ahk")
Run GetScriptPath("Open ChatGPT.ahk")
Run GetScriptPath("ChatGPT - Copy last message.ahk")
Run GetScriptPath("Open Google.ahk")
Run GetScriptPath("Open WhatsApp.ahk")
Run GetScriptPath("Minimize.ahk")
Run GetScriptPath("Jump mouse on the middle.ahk")
Run GetScriptPath("Hunt and Peck.ahk")