#Include env.ahk

;============ Software changed shortcuts ============
; VS Code
; Ctrl + Shift + 8: fold/unfold all
; Ctrl + Shift + 9: fold/unfold current

;============ Key remapping ============
; Sofle - Key remapping
Run GetScriptPath("Sofle - Key remapping.ahk")

;============ softwares ============

if (IS_WORK_ENVIRONMENT) {
    Run "C:\Users\fie7ca\Documents\CapsLock Indicator\CLIv3-3.16.1.2.exe"
    Run "C:\Users\fie7ca\Documents\HuntAndPack\hap.exe"
    Run "C:\Users\fie7ca\Documents\Caffeine\caffeine64.exe"
    Run "C:\Users\fie7ca\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\apps do Chrome\Mobills"
    Run "C:\Users\fie7ca\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\apps do Chrome\Gmail"
    Run "C:\Users\fie7ca\Documents\HuntAndPack\hap.exe"
} else {
    Run "C:\Users\eduev\OneDrive\Documentos\HuntAndPeck\hap.exe"
}

; Open ChatGPT
Send "#!+c"

;============ Win + keys ============

; Win + E : Opens Desktop
Run GetScriptPath("Open Desktop.ahk")

;============ Alt + Shift + keys ============
;Alt + Shift + Q: power toys command palette
;Alt + Shift + W: open file explorer
;Alt + Shift + Z: Bookmarks
; Alt + Shift + S: Windows
; Alt + Shift + A: Web

;============ Win + Alt + keys ============

; Win + Alt + Right/Left/Up/Down : Fast tab navigation
Run GetScriptPath("Fast Tab.ahk")

;============ CapsLock + keys ============

; CapsLock + N : Opens or activates OneNote
Run GetScriptPath("OneNote - Open.ahk")

; ; CapsLock + S : Opens or activates Spotify
Run GetScriptPath("Spotify - Open.ahk")

; CapsLock + 6 : Get word pronunciation, definition, and Portuguese translation from ChatGPT
Run GetScriptPath("ChatGPT - Pronunciation.ahk")

; CapsLock + 3 : Toggles the
Run GetScriptPath("ChatGPT - Toggle microphone.ahk")

; CapsLock + 4 : Activates Youtube
Run GetScriptPath("Youtube - Activate.ahk")

; CapsLock + 5 : Check grammar and improve text in both English and Portuguese
Run GetScriptPath("ChatGPT - Check for grammar.ahk")

; CapsLock + W : Clear ChatGPT input field for new conversation
Run GetScriptPath("ChatGPT - Free Speach.ahk")

; CapsLock + 7: Speek with ChatGPT
Run GetScriptPath("ChatGPT - Speek.ahk")

; CapsLock + C : Opens ChatGPT
Run GetScriptPath("Open ChatGPT.ahk")
; CapsLock + G : Opens Google
Run GetScriptPath("Open Google.ahk")

; CapsLock + Z : Opens WhatsApp
Run GetScriptPath("Open WhatsApp.ahk")

; CapsLock + Alt + Up/Down/Left/Right : Fast scrolling
Run GetScriptPath("Fast Up and Down.ahk")

; CapsLock + 0 : Hunt and Peck
Run GetScriptPath("Hunt and Peck.ahk")

; CapsLock + ] : Minimizes windows
Run GetScriptPath("Minimize.ahk")

if (!IS_WORK_ENVIRONMENT) {
    Run GetScriptPath("Overleaf - Pages interaction.ahk")
}

if (IS_WORK_ENVIRONMENT) {
    ; CapsLock + 2
    ; Run "C:\Users\fie7ca\OneDrive - Bosch Group\07 - Scripts\Windows Explorer - Copy path.ahk" ❌

    ; CapsLock + 8
    Run GetScriptPath("Outlook - Send to general.ahk")

    ; CapsLock + 9
    Run GetScriptPath("Outlook - Send to newsletter.ahk")

    ; CapsLock + A
    Run GetScriptPath("Outlook - Open mail.ahk")

    ; CapsLock + R
    Run GetScriptPath("Outlook - Open Reminder.ahk")

    ; CapsLock + X
    Run GetScriptPath("Outlook - Open calendar.ahk")

    ; CapsLock + Y
    Run GetScriptPath("Microsoft Teams - Mark as unread.ahk")

    ; CapsLock + M
    Run GetScriptPath("Microsoft Teams - New conversation.ahk")

    ; CapsLock + I
    Run GetScriptPath("Microsoft Teams - Laugh.ahk")

    ; CapsLock + O
    Run GetScriptPath("Microsoft Teams - Like.ahk")

    ; CapsLock + P
    Run GetScriptPath("Microsoft Teams - Heart.ahk")

    ; CapsLock + H
    Run GetScriptPath("Microsoft Teams - Reply Message.ahk")

    ; CapsLock + J
    Run GetScriptPath("Microsoft Teams - Pin.ahk")

    ; CapsLock + K
    Run GetScriptPath("Microsoft Teams - Remove pin.ahk")
}
