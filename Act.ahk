; Act.ahk - Comprehensive AutoHotkey Script for Keyboard Automation and Application Management
; This script provides multi-layer keyboard shortcuts, application launchers, and automation for various software including Teams, Outlook, ChatGPT, Spotify, and more.

;---------------------------------------- Github ----------------------------------------------------
; https://github.com/duducm2/zmk-sofle/blob/main/keymap-drawer/eyelash_sofle.svg
;---------------------------------------- Layer 0 + Keys ----------------------------------
;    Power toys command palette
;    Open file explorer
;    Windows
;---------------------------------------- Layer 1 + Keys ----------------------------------
; Layer 1 + ESC: F11
; Layer 1 + 1: F1
; Layer 1 + 2: F2
; Layer 1 + 3: F3
; Layer 1 + 4: F4
; Layer 1 + 5: F5
; Layer 1 + Up arrow: Mouse move up
; Layer 1 + 6: F6
; Layer 1 + 7: F7
; Layer 1 + 8: F8
; Layer 1 + 9: F9
; Layer 1 + 0: F10
; Layer 1 + Backspace: F12
; Layer 1 + E: Macro Win+Z, 4, Enter
; Layer 1 + R: Macro Win+Z, 5, Enter
; Layer 1 + T: Macro Win+Z, 6, Enter
; Layer 1 + Down arrow: Mouse move down
; Layer 1 + Y: Page Up
; Layer 1 + U: Home
; Layer 1 + I: ]
; Layer 1 + O: /
; Layer 1 + P: ;
; Layer 1 + F: Macro Win+Z, 7, Enter
; Layer 1 + G: Macro Win+Z, 8, Enter
; Layer 1 + Left arrow: Mouse move left
; Layer 1 + H: Page Down
; Layer 1 + J: End
; Layer 1 + K: \
; Layer 1 + L: -
; Layer 1 + ´: '
; Layer 1 + B: Win+L (Lock screen)
; Layer 1 + Right arrow: Mouse move right
; Layer 1 + N: ~
; Layer 1 + M: `
; Layer 1 + ,: / (numpad)
; Layer 1 + .: +
; Layer 1 + Enter (right): [
; Layer 1 + Play/Pause: Mute
; Layer 1 + Enter (thumb): Left mouse click
; Layer 1 + mo 2: ?
; Layer 1 + Printscreen: {
; Layer 1 + Context Menu: |
; Layer 1 + Delete: Alt+Shift+S

;---------------------------------------- Layer 2 + Keys ----------------------------------
; Layer 2 + ESC: Alt+F4
; Layer 2 + 1: Bluetooth Select 0
; Layer 2 + 2: Bluetooth Select 1
; Layer 2 + 3: Bluetooth Select 2
; Layer 2 + 4: Win+Home
; Layer 2 + 5: Win+Shift+E
; Layer 2 + Up arrow: 5 up arrows (Ctrl+Next Track with Ctrl)
; Layer 2 + 6: Ctrl+Shift+Home
; Layer 2 + 7: Ctrl+Shift+End
; Layer 2 + 8: Macro Select to beginning of word (5 times) and delete
; Layer 2 + 9: Macro Select to end of word (5 times) and delete
; Layer 2 + 0: Ctrl+Home
; Layer 2 + Backspace: Ctrl+End
; Layer 2 + Down arrow: 5 down arrows (Ctrl+Previous Track with Ctrl)
; Layer 2 + Y: Shift+Home
; Layer 2 + U: Shift+End
; Layer 2 + I: Macro Select to beginning of line and delete
; Layer 2 + O: Macro Select to end of line and delete
; Layer 2 + P: Ctrl+Page Up
; Layer 2 + ´: Ctrl+Page Down
; Layer 2 + F: Ctrl+Shift+V
; Layer 2 + Left arrow: 5 left arrows (Ctrl+5 left arrows with Ctrl)
; Layer 2 + H: Ctrl+Shift+Page Up
; Layer 2 + J: Ctrl+Shift+Page Down
; Layer 2 + V: Ctrl+Alt+V
; Layer 2 + B: Ctrl+Alt+B
; Layer 2 + Right arrow: 5 right arrows (Ctrl+5 right arrows with Ctrl)
; Layer 2 + Play/Pause: Mute
; Layer 2 + Enter (thumb): Left mouse click
; Layer 2 + Printscreen: Output BLE
; Layer 2 + Context Menu: Bluetooth Clear
; Layer 2 + Delete: Ctrl+A, Delete

;---------------------------------------- Win + Alt + Shift + keys ----------------------------------
; ----- Available -----------
; A C J K W X Y 1
; ----- Onenote -------------
;     Win + Alt + Shift + N : Opens or activates OneNote
; ----- Spotify -------------
;     Win + Alt + Shift + S : Opens or activates Spotify
;     Win + Alt + Shift + D : Spotify - Go to library
; ----- ChatGPT -------------
;     Win + Alt + Shift + 8 : Get word pronunciation, definition, and Portuguese translation from ChatGPT
;     Win + Alt + Shift + 9 : Clicks on the last microphone icon in the ChatGPT
;     Win + Alt + Shift + 0 : Speak with ChatGPT
;     Win + Alt + Shift + P : Copy last ChatGPT prompt
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
;---------------------------------------- Shift + keys ----------------------------------------------
; ----- You can have repeated keys, depending on the software.
; ----- Prefered Keys sequences (most important first): Y U I O P H J K L N M , . 6 7 8 9 0 W E R T D F G C V B
; ----- Spotify -----
;     Shift + Y : Toggle devices
;     Shift + U : Select devices areas
;     Shift + I : Full screen
; ----- Microsoft Teams (chat) -----
;     Shift + J : Mark as unread
;     Shift + Y : Like
;     Shift + U : Heart
;     Shift + I : Laugh
;     Shift + O : Go to home
;     Shift + E : Edit message
;     Shift + R : Reply Message
;     Shift + L : Remove pin
;     Shift + K : Pin
; ----- Microsoft Teams (meeting) -----
;     Shift + Y : Open chat
;     Shift + U : Maximize
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
;     Shift + Y: Send to general
;     Shift + U: Send to newsletter
;     Shift + I: ?
;     Shift + O: Ful screen
;     Shift + P: Normal mode
;     Shift + H: Go to title
;     Shift + J: Go required people
;     Shift + K: Go date
;     Shift + L: Go to message
;     Shift + N: Toggle between current and other tabs
; ----- ChatGPT --------------
;     Shift + Y: Select all and cut
;     Shift + U: Change model
;     Shift + I: Togle sidebar
;     Shift + O: Write chatgpt
;     Shift + P: New chat
;     Shift + H: Copy last code block
; ----- Cursor --------------
;     Shift + Y : Unfold
;     Shift + U : Fold
;     Shift + I : Unfold all
;     Shift + O : Fold all
;     Shift + P : Go to terminal
;     Shift + H : New terminal
;     Shift + J : Go to file explorer
;     Shift + K : Format code
;     Shift + L : command palette
;     Shift + M : Change project
;     Shift + , : Show chat history
;     Shift + . : Extensions
;     Shift + 6 : Switch the brackets open/close
;     Shift + 7 : Search
;     Shift + 8 : Save all documents
; ----- Windows Explorer --------------
;     Shift + Y : Select first file
;     Shift + U : New folder
;     Shift + I : New Shortcut
; ----- ClipAngel --------------
;     Shift + Y : Select filtered content and copy
; ----- Figma --------------
;     Shift + Y : Show/Hide UI (Ctrl + \)
;     Shift + U : Component search (Shift + I)
;     Shift + I : Select parent (\)
;     Shift + O : Create component (Ctrl + Alt + K)
;     Shift + P : Detach instance (conflicting with other ahk) (Ctrl + Alt + B)
;     Shift + H : Add auto layout (Shift + A)
;     Shift + J : Remove auto layout (not working) (Alt + Shift + A)
;     Shift + K : Suggest auto layout (Ctrl + Alt + Shift + A)
;     Shift + L : Export (Ctrl + Shift + E)
;     Shift + N : Copy as PNG (Ctrl + Shift + C)
;     Shift + M : Actions... (Ctrl + K)
;     Shift + , : Align left (Alt + A)
;     Shift + . : Align right (Alt + D)
;     Shift + 6 : Align top (Alt + W)
;     Shift + 7 : Align bottom (Alt + S)
;     Shift + 8 : Align center horizontal (Alt + H)
;     Shift + 9 : Align center vertical (Alt + V) (conflicting with other ahk)
;     Shift + 0 : Distribute horizontal spacing (Alt + Shift + H)
;     Shift + W : Distribute vertical spacing (Alt + Shift + V)
;     Shift + E : Tidy up (Ctrl + Alt + Shift + T)

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