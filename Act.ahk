;---------------------------------------- Github ----------------------------------------------------
; https://github.com/duducm2/zmk-sofle/blob/main/keymap-drawer/eyelash_sofle.svg

;---------------------------------------- Win + Alt + Shift + keys ----------------------------------
; ----- Available -----------
; C K W X Y 1
; ----- Show shortcuts with shift + key -------------
;     Win + Alt + Shift + A : Show shortcuts with shift + key
; ----- Onenote -------------
;     Win + Alt + Shift + N : Opens or activates OneNote
; ----- Spotify -------------
;     Win + Alt + Shift + S : Opens or activates Spotify
;     Win + Alt + Shift + D : Spotify - Go to library
; ----- ChatGPT -------------
;     Win + Alt + Shift + 8 : Get word pronunciation, definition, and Portuguese translation from ChatGPT
;     Win + Alt + Shift + 9 : Clicks on the last microphone icon in the ChatGPT
;     Win + Alt + Shift + 0 : Speak with ChatGPT
;     Win + Alt + Shift + 7 : Speak with ChatGPT (send message automatically)
;     Win + Alt + Shift + P : Copy last ChatGPT prompt
;     Win + Alt + Shift + J : Copy last ChatGPT message (in the editing box)
;     Win + Alt + Shift + U : Activate ChatGPT and copy last code block
;     Win + Alt + Shift + O : Check grammar and improve text in both English and Portuguese
;     Win + Alt + Shift + I : Opens ChatGPT
;     Win + Alt + Shift + L : Talk with ChatGPT through voice
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
;     Win + Alt + Shift + M : Maximizes the current window
; ----- General ---------
;     Win + Alt + Shift + Q : Jump mouse on the middle
; ----- Hunt and Peck ---------
;     Win + Alt + Shift + X : Activate hunt and Peck

;---------------------------------------- Scripts ---------------------------------------------------

#Include env.ahk

if (IS_WORK_ENVIRONMENT) {

    Run "C:\Users\fie7ca\Documents\HuntAndPack\hap.exe"
    Run "C:\Users\fie7ca\Documents\Caffeine\caffeine64.exe"
    Run "C:\Users\fie7ca\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\apps do Chrome\Mobills"
    Run "C:\Users\fie7ca\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\apps do Chrome\Gmail"
    Run "c:\Users\fie7ca\Documents\Settle Up.lnk"
    Run "c:\ProgramData\Microsoft\Windows\Start Menu\Programs\OneNote.lnk"
    Run "cc:\Users\fie7ca\Documents\Atalhos\Microsoft Teams - Shortcut.lnk"
    Run GetScriptPath("Microsoft Teams.ahk")
    Run GetScriptPath("Outlook.ahk")
} else {
    Run "C:\Users\eduev\OneDrive\Documentos\HuntAndPeck\hap.exe"
}

; Open ChatGPT
Send "^!c"

Run GetScriptPath("Shift keys.ahk")
Run GetScriptPath("AppLaunchers.ahk")
Run GetScriptPath("Spotify.ahk")
Run GetScriptPath("ChatGPT.ahk")
Run GetScriptPath("WindowManagement.ahk")
Run GetScriptPath("Utils.ahk")