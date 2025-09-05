/********************************************************************
 *   Win+Alt+Shift symbol layer shortcuts (AHK v2)
 *   â€¢ Provides system-wide symbol shortcuts
 ********************************************************************/

/********************************************************************
 *   AVAILABLE WIN+ALT+SHIFT COMBINATIONS
 *   The following combinations are not currently in use:
 *   â€¢ C - Available
 *   â€¢ 1 - Available
 ********************************************************************/

#Requires AutoHotkey v2.0+

#SingleInstance Force

SetTitleMatchMode 2

#include %A_ScriptDir%\env.ahk
#include UIA-v2\Lib\UIA.ahk
#include UIA-v2\Lib\UIA_Browser.ahk
#include %A_ScriptDir%\ChatGPT_Loading.ahk

; --- Global Variables ---
global smallLoadingGuis_ChatGPT := []

; Helper: find ChatGPT chrome window by case-insensitive contains match
GetChatGPTWindowHwnd() {
    for hwnd in WinGetList("ahk_exe chrome.exe") {
        if InStr(WinGetTitle("ahk_id " hwnd), "chatgpt", false)
            return hwnd
    }
    return 0
}

; --- Config ---------------------------------------------------------------
PROMPT_FILE := A_ScriptDir "\\ChatGPT_Prompt.txt"

; Function to send symbol characters
SendSymbol(sym) {
    SendText(sym)
}

; Helper: normalize common UTF-8→CP1252 mojibake so arrows and punctuation display correctly
NormalizeMojibake(str) {
    if (str = "")
        return str
    reps := Map(
        "â†’", "→",   ; right arrow
        "â†�", "←",   ; left arrow
        "â†‘", "↑",   ; up arrow
        "â†“", "↓",   ; down arrow
        "â€¢", "•",
        "â€“", "–",
        "â€”", "—",
        "â€¦", "…",
        "â€˜", "‘",
        "â€™", "’",
        "â€œ", "“",
        "â€", "”",
        "Ã—", "×"
    )
    for k, v in reps
        str := StrReplace(str, k, v)
    return str
}

;-------------------------------------------------------------------
; Cheat-sheet overlay (Win + Alt + Shift + A) â€“ shows remapped shortcuts
;-------------------------------------------------------------------

; Map that stores the pop-up text for each application.  Extend freely.
cheatSheets := Map()

; --- Mercado Livre (Brazil) -----------------------------------------------
cheatSheets["Mercado Livre"] := "
(
Mercado Livre
Shift+Y  ï¿½ï¿½'  Focus search field
Shift+U  ï¿½ï¿½'  Carrinho de compras
Shift+I  ï¿½ï¿½'  Compras feitas
)"  ; end Mercado Livre

;---------------------------------------- Shift + keys ----------------------------------------------
; ----- Assignment policy: use Shift + <key> first. When all Shift slots in the sequence are consumed, continue with Ctrl + Alt + <key> in the same order.
; ----- You can have repeated keys, depending on the software.
; ----- Preferred key sequence (most important first): Y U I O P H J K L N M , . W E R T D F G C V B
; ----- Ctrl + Alt sequence (fallback, same order):    Y U I O P H J K L N M , . W E R T D F G C V B

; --- WhatsApp desktop -------------------------------------------------------
cheatSheets["WhatsApp"] := "
(
WhatsApp
Shift+Y  â†’  Toggle voice message
Shift+U  â†’  Search chats
Shift+I  â†’  Reply
Shift+O  â†’  Sticker panel
Shift+P  â†’  Toggle Unread filter
Shift+H  â†’  Focus current chat
Shift+J  â†’  Mark as read or unread
Shift+K  â†’  Pin chat or unpin chat
)"  ; end WhatsApp

; --- Outlook main window ----------------------------------------------------
cheatSheets["OUTLOOK.EXE"] := "
(
Outlook
Shift+Y  â†’  Send to General
Shift+U  â†’  Send to Newsletter
Shift+I  â†’  Go to Inbox
Shift+N  â†’  Focused / Other
)"  ; end Outlook

; --- Outlook Reminder window -------------------------------------------------
cheatSheets["OutlookReminder"] := "
(
Outlook â€“ Reminders
Shift+Y  â†’  Select first reminder
Shift+U  â†’  Snooze 1 hour
Shift+I  â†’  Snooze 4 hours
Shift+O  â†’  Snooze 1 day
Shift+P  â†’  Dismiss all reminders
Shift+H  â†’  Join Online
)"  ; end Outlook Reminder

; --- Outlook Appointment window ---------------------------------------------
cheatSheets["OutlookAppointment"] := "
(
Outlook â€“ Appointment
Shift+Y  â†’  Start date (combo)
Shift+U  â†’  Start date â€“ Date Picker
Shift+I  â†’  Start time (combo)
Shift+O  â†’  End date (combo)
Shift+P  â†’  End time (combo)
Shift+H  â†’  All day checkbox
Shift+J  â†’  Title field
Shift+L  â†’  Required / To field
Shift+M  â†’  Location â†’ Body
Shift+,  â†’  Make Recurring
)"  ; end Outlook Appointment

; --- Outlook Message window ---------------------------------------------------
cheatSheets["OutlookMessage"] := "
(
Outlook â€“ Message
Shift+Y  â†’  Subject / Title
Shift+U  â†’  Required / To
Shift+M  â†’  Location â†’ Body
)"  ; end Outlook Message

; --- Microsoft Teams â€“ meeting window --------------------------------------
cheatSheets["TeamsMeeting"] := "
(
Teams
Shift+Y  â†’  Open Chat pane
Shift+U  â†’  Maximize meeting window
Shift+I  â†’  Reagir
Shift+O  â†’  Join now with camera and microphone on
Shift+P  â†’  Audio settings
)"  ; end TeamsMeeting

; --- Microsoft Teams â€“ chat window -----------------------------------------
cheatSheets["TeamsChat"] := "
(
Teams
Shift+Y  â†’  Like
Shift+U  â†’  Heart
Shift+I  â†’  Laugh
Shift+O  â†’  Home panel
Shift+J  â†’  Mark unread
Shift+K  â†’  Pin chat
Shift+L  â†’  Remove pin
Shift+E  â†’  Edit message
Shift+R  â†’  Reply
)"  ; end TeamsChat

; --- Spotify ---------------------------------------------------------------
cheatSheets["Spotify.exe"] := "
(
Spotify
Shift+Y  â†’  Toggle Connect panel
Shift+U  â†’  Toggle Full screen
Shift+I  â†’  Open Search
Shift+O  â†’  Go to Playlists
Shift+P  â†’  Go to Artists
Shift+H  â†’  Go to Albums
Shift+J  â†’  Go to Search
Shift+K  â†’  Go to Home
Shift+L  â†’  Go to Now Playing
Shift+N  â†’  Go to Made For You
Shift+M  â†’  Go to New Releases
Shift+,  â†’  Go to Charts
Shift+.  â†’  Toggle Now Playing View
Shift+W  â†’  Toggle Library Sidebar
Shift+E  â†’  Toggle Fullscreen Library
Shift+R  â†’  Toggle lyrics
Shift+T  â†’  Toggle play/pause
)"  ; end Spotify

; --- OneNote ---------------------------------------------------------------
cheatSheets["ONENOTE.EXE"] := "
(
OneNote
Shift+Y  â†’  Collapse
Shift+U  â†’  Expand
Shift+I  â†’  Collapse all
Shift+O  â†’  Expand all
Shift+P  â†’  Select line and children
Shift+D  â†’  Delete line and children
Shift+S  â†’  Delete line (keep children)
Shift+F  â†’  Advanced Searching with double quotes
)"  ; end OneNote

; --- Chrome general shortcuts ----------------------------------------------
cheatSheets["chrome.exe"] := "
(
Chrome
Shift+G  â†’  Pop current tab to new window
)"  ; end Chrome

; --- Cursor ------------------------------------------------------
cheatSheets["Cursor.exe"] := "
(
Cursor
Shift+Y  â†’  Fold
Shift+U  â†’  Unfold
Shift+I  â†’  Fold all
Shift+O  â†’  Unfold all
Shift+P  â†’  Go to terminal
Shift+H  â†’  New terminal
Shift+J  â†’  Go to file explorer
Shift+K  â†’  Format code
Shift+L  â†’  Command palette
Shift+N  â†’  Expand selection
Shift+M  â†’  Change project
Shift+,  â†’  Show chat history
Shift+.  â†’  Extensions
Shift+W  â†’  Switch brackets
Shift+E  â†’  Search
Shift+R  â†’  Save all documents
Shift+T  â†’  Change ML model
Shift+D  â†’  Git section
Shift+F  â†’  Close all editors
Shift+G  â†’  Switch AI models (auto/CLAUD/GPT/O/DeepSeek/Cursor)
Shift+C  â†’  Switch AI modes (agent/ask)
Shift+V  â†’  Fold Git repos (SCM)
Shift+B  â†’  Create AI commit message, then select Commit or Commit and Push
Ctrl + Alt + Y  â†’  Select to Bracket
Ctrl + Alt + U  â†’  Open in file explorer
Ctrl + Alt + I  â†’  Fold all directories
Ctrl + Alt + O  â†’  Unfold all directories
Ctrl + Alt + H  â†’  Kill terminal  [custom: defined in Cursor settings.json]

--- Additional Shortcuts ---
Ctrl + T  â†’  New chat tab
Ctrl + N  â†’  New chat tab (replacing current)
Alt + F12  â†’  Peek Definition
Ctrl + D  â†’  Select next identical word
Ctrl + Shift + L  â†’  Select all identical words
F2  â†’  Rename symbol
F8  â†’  Navigate problems
Ctrl + Enter  â†’  Insert line below
Ctrl + P  â†’  Quick Open
Shift + Delete  â†’  Delete line
Alt + â†‘  â†’  Move line up
Alt + â†“  â†’  Move line down
Ctrl + 1 / Ctrl + 2 / Ctrl + 3 ...  â†’  Switch tabs
Ctrl + N  â†’  New file
Ctrl + Alt + â†‘  â†’  Add cursor above
Ctrl + Alt + â†“  â†’  Add cursor below
Alt + Click  â†’  Multi-cursor by click
Shift + Alt + â†‘  â†’  Copy line up
Shift + Alt + â†“  â†’  Copy line down
Ctrl + ;  â†’  Insert comment
Ctrl + Home  â†’  Go to top
Ctrl + End  â†’  Go to end
Alt + Z  â†’  Toggle word wrap
Ctrl + Shift + D  â†’  Debugging
Ctrl + R  â†’  Quick project switch
Alt + J  â†’  Next review
Alt + K  â†’  Previous review
)"  ; end Cursor

; --- Windows Explorer ------------------------------------------------------
cheatSheets["explorer.exe"] := "
(
Explorer
Shift+Y  â†’  Select first file
Shift+U  â†’  Focus search bar
Shift+I  â†’  Focus address bar
Shift+O  â†’  New folder
Shift+J  â†’  Create a shortcut
Shift+K  â†’  Copy as path
Shift+P  â†’  Select first pinned item in Explorer sidebar
Shift+H  â†’  Select the last item of the Explorer sidebar
)"  ; end Explorer

; --- Microsoft Paint ------------------------------------------------------
cheatSheets["mspaint.exe"] := "
(
MS Paint
Shift+Y  â†’  Resize and Skew (Ctrl+W)

--- Common Shortcuts ---
Ctrl+N   â†’  New
Ctrl+O   â†’  Open
Ctrl+S   â†’  Save
F12      â†’  Save As
Ctrl+P   â†’  Print
Ctrl+Z   â†’  Undo
Ctrl+Y   â†’  Redo
Ctrl+A   â†’  Select all
Ctrl+C   â†’  Copy
Ctrl+X   â†’  Cut
Ctrl+V   â†’  Paste
Ctrl+W   â†’  Resize and Skew
Ctrl+E   â†’  Image properties
Ctrl+R   â†’  Toggle rulers
Ctrl+G   â†’  Toggle gridlines
Ctrl+I   â†’  Invert colors
F11      â†’  Fullscreen view
Ctrl++   â†’  Zoom in
Ctrl+-   â†’  Zoom out
)"  ; end MS Paint

; --- ClipAngel -------------------------------------------------------------
cheatSheets["ClipAngel.exe"] := "
(
ClipAngel
Shift+Y  â†’  Select filtered content and copy
Shift+U  â†’  Switch focus list/text
Shift+I  â†’  Delete all non-favorite
Shift+O  â†’  Clear filters
Shift+P  â†’  Mark as favorite
Shift+H  â†’  Unmark as favorite
Shift+J  â†’  Edit text
Shift+K  â†’  Save as file
Shift+L  â†’  Merge clips
)"  ; end ClipAngel

; --- Figma -----------------------------------------------------------------
cheatSheets["Figma.exe"] := "
(
Figma
Shift+Y  â†’  Show/Hide UI
Shift+U  â†’  Component search
Shift+I  â†’  Select parent
Shift+O  â†’  Create component
Shift+P  â†’  Detach instance
Shift+H  â†’  Add auto layout
Shift+J  â†’  Remove auto layout
Shift+K  â†’  Suggest auto layout
Shift+L  â†’  Export
Shift+N  â†’  Copy as PNG
Shift+M  â†’  Actions...
Shift+,  â†’  Align left
Shift+.  â†’  Align right
Shift+W  â†’  Distribute vertical spacing
Shift+E  â†’  Tidy up
Shift+R  â†’  Align top
Shift+T  â†’  Align bottom
Shift+D  â†’  Align center horizontal
Shift+F  â†’  Align center vertical
Shift+G  â†’  Distribute horizontal spacing
)"  ; end Figma

; --- Gmail ---------------------------------------------------------------
cheatSheets["Gmail"] := "
(
Gmail
Shift+Y  â†’  Go to main inbox
Shift+U  â†’  Go to updates
Shift+I  â†’  Go to forums
Shift+O  â†’  Toggle read/unread
Shift+P  â†’  Previous conversation
Shift+H  â†’  Next conversation
Shift+J  â†’  Archive conversation
Shift+K  â†’  Select conversation
Shift+L  â†’  Reply
Shift+N  â†’  Reply all
Shift+M  â†’  Forward
Shift+,  â†’  Star/unstar conversation
Shift+.  â†’  Delete
Shift+W  â†’  Report as spam
Shift+E  â†’  Compose new email
Shift+R  â†’  Search mail
Shift+T  â†’  Move to folder
Shift+D  â†’  Show keyboard shortcuts help

--- Built-in Shortcuts (Windows) ---

Compose & chat:
p  â†’  Previous message in an open conversation
n  â†’  Next message in an open conversation
Shift + Esc  â†’  Focus main window
Esc  â†’  Focus latest chat or compose
Ctrl + .  â†’  Advance to the next chat or compose
Ctrl + ,  â†’  Advance to previous chat or compose
Ctrl + Enter  â†’  Send
Ctrl + Shift + c  â†’  Add cc recipients
Ctrl + Shift + b  â†’  Add bcc recipients
Ctrl + Shift + f  â†’  Access custom from
Ctrl + k  â†’  Insert a link
Ctrl + m  â†’  Open spelling suggestions

Formatting text:
Ctrl + Shift + 5  â†’  Previous font
Ctrl + Shift + 6  â†’  Next font
Ctrl + Shift + -  â†’  Decrease text size
Ctrl + Shift + +  â†’  Increase text size
Ctrl + b  â†’  Bold
Ctrl + i  â†’  Italics
Ctrl + u  â†’  Underline
Ctrl + Shift + 7  â†’  Numbered list
Ctrl + Shift + 8  â†’  Bulleted list
Ctrl + Shift + 9  â†’  Quote
Ctrl + [  â†’  Indent less
Ctrl + ]  â†’  Indent more
Ctrl + Shift + l  â†’  Align left
Ctrl + Shift + e  â†’  Align center
Ctrl + Shift + r  â†’  Align right
Ctrl + \  â†’  Remove formatting

Actions (shortcuts on):
,  â†’  Move focus to toolbar
x  â†’  Select conversation
s  â†’  Toggle star/rotate among superstars
e  â†’  Archive
m  â†’  Mute conversation
!  â†’  Report as spam
#  â†’  Delete
r  â†’  Reply
Shift + r  â†’  Reply in a new window
a  â†’  Reply all
Shift + a  â†’  Reply all in a new window
f  â†’  Forward
Shift + f  â†’  Forward in a new window
Shift + n  â†’  Update conversation
] or [  â†’  Archive conversation and go previous/next
z  â†’  Undo last action
Shift + i  â†’  Mark as read
Shift + u  â†’  Mark as unread
_  â†’  Mark unread from the selected message
+ or =  â†’  Mark as important
-  â†’  Mark as not important
b  â†’  Snooze (not available in classic Gmail)
;  â†’  Expand entire conversation
:  â†’  Collapse entire conversation
Shift + t  â†’  Add conversation to Tasks

Jumping (shortcuts on):
g + i  â†’  Go to Inbox
g + s  â†’  Go to Starred conversations
g + b  â†’  Go to Snoozed conversations
g + t  â†’  Go to Sent messages
g + d  â†’  Go to Drafts
g + a  â†’  Go to All mail
Ctrl + Alt + ,  â†’  Switch to left sidebar (Calendar/Keep/Tasks)
Ctrl + Alt + .  â†’  Switch to right (back to inbox)
g + k  â†’  Go to Tasks
g + l  â†’  Go to label

Threadlist selection (shortcuts on):
* + a  â†’  Select all conversations
* + n  â†’  Deselect all conversations
* + r  â†’  Select read conversations
* + u  â†’  Select unread conversations
* + s  â†’  Select starred conversations
* + t  â†’  Select unstarred conversations

Navigation (shortcuts on):
g + n  â†’  Go to next page
g + p  â†’  Go to previous page
u  â†’  Back to threadlist
k  â†’  Newer conversation
j  â†’  Older conversation
o or Enter  â†’  Open conversation
``  â†’  Go to next Inbox section
~  â†’  Go to previous Inbox section

Application (shortcuts on):
c  â†’  Compose
d  â†’  Compose in a new tab
/  â†’  Search mail
q  â†’  Search chat contacts
.  â†’  Open ""more actions"" menu
v  â†’  Open ""move to"" menu
l  â†’  Open ""label as"" menu
?  â†’  Open keyboard shortcut help
)"  ; end Gmail

; --- Google Keep ---------------------------------------------------------------
cheatSheets["Google Keep"] := "
(
Google Keep
Shift+Y  â†’  Search and select note
Shift+U  â†’  Toggle main menu
)"  ; end Google Keep

; --- File Dialog ---------------------------------------------------------------
cheatSheets["FileDialog"] := "
(
File Dialog
Shift+Y  â†’  Select first item
Shift+U  â†’  Focus search bar
Shift+I  â†’  Focus address bar
Shift+O  â†’  New folder
Shift+P  â†’  Select first pinned item in sidebar
Shift+H  â†’  Select 'This PC' / 'Este computador' in sidebar
Shift+J  â†’  Focus file name field
Shift+K  â†’  Click Insert/Open/Save button
Shift+L  â†’  Click Cancel button
)"

; --- Settings Window -------------------------------------------------
cheatSheets["Settings"] := "(Settings)`r`nShift+Y â†’ Set input volume to 100%"

; --- UIA Tree Inspector -------------------------------------------------
cheatSheets["UIATreeInspector"] := "(UIA Tree Inspector)`r`nShift+Y → Refresh list`r`nShift+U → Focus filter field"
; --- SettleUp Shortcuts -----------------------------------------------------
cheatSheets["Settle Up"] := "
(
Settle Up
Shift+Y  â†’  Add transaction
Shift+U  â†’  Focus expense value field
Shift+I  â†’  Focus expense name field
)"

; --- Miro Shortcuts -----------------------------------------------------
cheatSheets["Miro"] := "
(
Miro
Shift+U  â†’  Frame list
Shift+I  â†’  Group
Shift+O  â†’  Ungroup
Shift+P  â†’  Lock/Unlock
Shift+H  â†’  Add/Edit link
Q
--- Built-in Shortcuts (Windows) ---
Tools:
V / H             â†’  Select tool / Hand
T                 â†’  Text
N                 â†’  Sticky notes
S                 â†’  Shapes
R                 â†’  Rectangle
O                 â†’  Oval
L                 â†’  Connection line / Arrow
D                 â†’  Card
P                 â†’  Pen
E                 â†’  Eraser
C                 â†’  Comment
F                 â†’  Frame
M                 â†’  Minimap
Ctrl + K          â†’  Command palette
Enter (bulk)      â†’  New sticky note
Esc (bulk)        â†’  Exit sticky note bulk mode
Ctrl + Shift + Enter  â†’  Open card panel
Shift + C         â†’  Show/hide comments

General:
Ctrl + C / Ctrl + V  â†’  Copy / Paste
Ctrl + X          â†’  Cut
Ctrl + D          â†’  Duplicate
Alt + drag        â†’  Duplicate by drag
Alt + â†â†’â†‘â†“        â†’  Duplicate horizontally/vertically
Ctrl + click      â†’  Select multiple
Ctrl + A          â†’  Select all
Enter             â†’  Edit selected
Esc               â†’  Deselect / quit edit
Backspace         â†’  Delete
Ctrl + G          â†’  Group
Ctrl + Shift + G  â†’  Ungroup
Ctrl + Shift + L  â†’  Lock / Unlock
Ctrl + Shift + P  â†’  Protected lock / Unprotected lock
PgUp              â†’  Bring to front
Shift + PgUp      â†’  Bring forward
PgDn              â†’  Send to back
Shift + PgDn      â†’  Send backward
Ctrl + Shift + K  â†’  Create board in new tab
Alt + Ctrl + K    â†’  Add/Edit link to object
Ctrl + Backspace  â†’  Clear object contents

Navigation:
â†â†’â†‘â†“              â†’  Move items/canvas
Ctrl + +          â†’  Zoom in
Ctrl + -          â†’  Zoom out
Ctrl + 0          â†’  Zoom to 100%
Alt + 1           â†’  Zoom to fit
Alt + 2           â†’  Zoom to selected item
Space + drag      â†’  Move canvas
G                 â†’  Toggle grid
Ctrl + F          â†’  Search

Text:
Ctrl + B          â†’  Bold
Ctrl + I          â†’  Italic
Ctrl + U          â†’  Underline

Board navigation:
Tab               â†’  Move forwards through objects (TL â†’ BR)
Shift + Tab       â†’  Move backwards through objects (TL â†’ BR)
Ctrl + â†‘/â†“/â†/â†’    â†’  Move through board objects
Ctrl + Shift + â†“/â†‘â†’  Move in/out of container (e.g., frame)
Esc               â†’  Back to menu
Enter             â†’  Edit an object
Esc               â†’  Stop editing an object

Toolbar navigation:
Tab / Shift + Tab â†’  Move between toolbars
Arrow keys        â†’  Move between toolbar items
Enter / Space     â†’  Activate a menu item

Desktop app:
Ctrl + R          â†’  Reload the tab
Ctrl + W          â†’  Close the tab
Ctrl + Q          â†’  Exit the app
Ctrl + Shift + L  â†’  Copy board link
)"

; --- Wikipedia ---------------------------------------------------------------
cheatSheets["Wikipedia"] := "
(
Wikipedia
Shift+Y  â†’  Click search button
)"

; ========== Helper to decide which sheet applies ===========================
GetCheatSheetText() {
    global cheatSheets

    exe := WinGetProcessName("A") ; active process name (e.g. chrome.exe)
    title := WinGetTitle("A")
    hwnd := WinExist("A")

    ; (removed temporary tooltip debugging)

    ; Prefer Outlook Reminders over generic File Dialog detection
    if (exe = "OUTLOOK.EXE") {
        if RegExMatch(title, "i)Reminder")
            return cheatSheets.Has("OutlookReminder") ? cheatSheets["OutlookReminder"] : ""
    }

    ; Check for file dialog first (works in any app)
    if WinGetClass("ahk_id " hwnd) = "#32770" {
        txt := WinGetText("ahk_id " hwnd)
        if InStr(txt, "Namespace Tree Control") || InStr(txt, "Controle da Ãrvore de Namespace")
            return cheatSheets["FileDialog"]
    }

    ; Check for Settings window (both English and Portuguese)
    if (title = "Settings" || title = "ConfiguraÃ§Ãµes") {
        return cheatSheets.Has("Settings") ? cheatSheets["Settings"] : ""
    }

    ; Special handling for Chrome-based apps that share chrome.exe
    if (exe = "chrome.exe") {
        chromeShortcuts := cheatSheets.Has("chrome.exe") ? cheatSheets["chrome.exe"] : ""
        appShortcuts := ""

        ; Normalize Chrome window title by removing the trailing " - Google Chrome"
        chromeTitle := RegExReplace(title, "i) - Google Chrome$", "")

        if InStr(chromeTitle, "WhatsApp")
            appShortcuts := cheatSheets.Has("WhatsApp") ? cheatSheets["WhatsApp"] : ""
        if InStr(chromeTitle, "Gmail")
            appShortcuts := cheatSheets.Has("Gmail") ? cheatSheets["Gmail"] : ""
        if InStr(chromeTitle, "chatgpt")
            appShortcuts :=
                "ChatGPT`r`nShift+Y â†’ Cut all`r`nShift+U â†’ Model selector`r`nShift+I â†’ Toggle sidebar`r`nShift+O â†’ Re-send rules`r`nShift+H â†’ Copy code block`r`nShift+L â†’ Send and show AI banner"
        if InStr(chromeTitle, "Mobills")
            appShortcuts :=
                "Mobills`r`nShift+Y â†’ Dashboard`r`nShift+U â†’ Contas`r`nShift+I â†’ TransaÃ§Ãµes`r`nShift+O â†’ CartÃµes de crÃ©dito`r`nShift+P â†’ Planejamento`r`nShift+H â†’ RelatÃ³rios`r`nShift+J â†’ Mais opÃ§Ãµes`r`nShift+K â†’ Previous month`r`nShift+L â†’ Next month`r`nShift+N â†’ Ignore transaction`r`nShift+M â†’ Name field"
        if InStr(chromeTitle, "Google Keep") || InStr(chromeTitle, "keep.google.com")
            appShortcuts := cheatSheets.Has("Google Keep") ? cheatSheets["Google Keep"] : ""
        if InStr(chromeTitle, "YouTube")
            appShortcuts :=
                "YouTube`r`nShift+Y â†’ Focus search box`r`nShift+U â†’ Focus first video via Search filters`r`nShift+I â†’ Focus first video via Explore"
        if InStr(chromeTitle, "UIATreeInspector")
            appShortcuts := cheatSheets["UIATreeInspector"]
        if InStr(chromeTitle, "Settle Up")
            appShortcuts := cheatSheets.Has("Settle Up") ? cheatSheets["Settle Up"] : ""
        if InStr(chromeTitle, "Miro")
            appShortcuts := cheatSheets.Has("Miro") ? cheatSheets["Miro"] : ""
        if InStr(chromeTitle, "Wikipedia", false) || InStr(chromeTitle, "wikipedia.org", false)
            appShortcuts := cheatSheets.Has("Wikipedia") ? cheatSheets["Wikipedia"] : ""
        if InStr(chromeTitle, "Mercado Livre", false)
            appShortcuts := cheatSheets.Has("Mercado Livre") ? cheatSheets["Mercado Livre"] : ""
        ; Only set generic Google sheet if nothing else matched and title clearly indicates Google site
        if (appShortcuts = "") {
            if (chromeTitle = "Google" || InStr(chromeTitle, " - Google Search"))
                appShortcuts := "Google`r`nShift+Y â†’ Focus search box"
        }

        ; Combine Chrome general + app-specific shortcuts
        if (appShortcuts != "" && chromeShortcuts != "")
            return chromeShortcuts "`r`n`r`n" appShortcuts
        else if (appShortcuts != "")
            return appShortcuts
        else if (chromeShortcuts != "")
            return chromeShortcuts
        else
            return ""
    }

    ; UIA Tree Inspector - check both process and window title
    if (exe = "AutoHotkey64.exe" && InStr(title, "UIATreeInspector"))
        return cheatSheets["UIATreeInspector"]

    ; Microsoft Teams â€“ differentiate meeting vs chat via helper predicates
    if IsTeamsMeetingActive()
        return cheatSheets.Has("TeamsMeeting") ? cheatSheets["TeamsMeeting"] : ""
    if IsTeamsChatActive()
        return cheatSheets.Has("TeamsChat") ? cheatSheets["TeamsChat"] : ""
    if IsFileDialogActive()
        return cheatSheets["FileDialog"]

    ; Special handling for Outlook-based apps
    if (exe = "OUTLOOK.EXE") {
        ; Detect Reminders window â€“ e.g. "3 Reminder(s)" or any title containing "Reminder"
        if RegExMatch(title, "i)Reminder") {
            return cheatSheets.Has("OutlookReminder") ? cheatSheets["OutlookReminder"] : cheatSheets["OUTLOOK.EXE"]
        }
        ; Detect Message inspector windows â€“ e.g., " - Message (HTML)"
        if RegExMatch(title, "i) - Message \(") {
            return cheatSheets.Has("OutlookMessage") ? cheatSheets["OutlookMessage"] : cheatSheets["OUTLOOK.EXE"]
        }
        ; Detect Appointment or Meeting inspector windows
        if RegExMatch(title, "i)(Appointment|Meeting)") {
            return cheatSheets.Has("OutlookAppointment") ? cheatSheets["OutlookAppointment"] : cheatSheets[
                "OUTLOOK.EXE"]
        }
        ; Fallback to generic Outlook sheet
        if cheatSheets.Has("OUTLOOK.EXE")
            return cheatSheets["OUTLOOK.EXE"]
    }

    ; Direct match by process name (generic fallback)
    if cheatSheets.Has(exe)
        return cheatSheets[exe]

    ; Try case-insensitive match for process name
    for key, value in cheatSheets {
        if (StrLower(key) = StrLower(exe))
            return value
    }

    ; Nothing found â†’ blank â†’ caller will show fallback message
    return ""
}

; ========== Shared variables for cheat sheet state ========================
global g_helpGui := 0
global g_helpShown := false
global g_globalGui := 0
global g_globalShown := false

; ========== GUI creation & showing ========================================
ToggleShortcutHelp() {
    global g_helpGui, g_helpShown

    ; Toggle off if currently shown
    if (IsObject(g_helpGui) && g_helpShown) {
        g_helpGui.Hide()
        g_helpShown := false
        ; Hotkey "Esc", "Off"  ; (disabled)
        return
    }

    ; Ensure text for current context
    text := NormalizeMojibake(GetCheatSheetText())
    if (text = "") {
        exe := WinGetProcessName("A")
        text := "No cheat-sheet registered for:`n" exe
    }

    static cheatCtrl

    if !IsObject(g_helpGui) {
        g_helpGui := Gui(
            "+AlwaysOnTop -Caption +ToolWindow +Border +Owner +LastFound"
        )
        g_helpGui.BackColor := "000000"
        g_helpGui.SetFont("s12 cFFFF00", "Consolas")
        ; Enable vertical scroll so oversized cheat sheets remain usable
        cheatCtrl := g_helpGui.Add("Edit", "ReadOnly +Multi -E0x200 +VScroll -HScroll -Border Background000000 w600 r1"
        )

        ; Esc also hides  ; (disabled â€“ use Win+Alt+Shift+A to hide)
        ; Hotkey "Esc", (*) => (g_helpGui.Hide(), g_helpShown := false), "Off"
    }

    ; Update cheat-sheet text and resize height to fit
    cheatCtrl.Value := text
    lineCnt := StrLen(text) ? StrSplit(text, "`n").Length : 1

    ; Calculate height based on line count (font size 12 â‰ˆ 20px per line + margins)
    ; Apply min/max so content scrolls instead of being cut off
    controlHeight := lineCnt * 20 + 10
    minHeight := 220  ; ensure a decent minimum height
    MonitorGetWorkArea(1, &ml, &mt, &mr, &mb)
    maxHeight := Floor((mb - mt) * 0.7)  ; cap to 70% of monitor work area
    if (controlHeight < minHeight)
        controlHeight := minHeight
    if (controlHeight > maxHeight)
        controlHeight := maxHeight

    ; Resize the control and GUI explicitly
    cheatCtrl.Move(, , 600, controlHeight)
    ; Show â†’ measure â†’ centre
    g_helpGui.Show("AutoSize Hide")
    CenterGuiOnActiveMonitor(g_helpGui)
    g_helpGui.Show("NoActivate")  ; ensure visible after centring
    g_helpShown := true
    ; Hotkey "Esc", "On"  ; (disabled)
}

; ========== Global shortcuts cheat sheet (Win+Alt+Shift+key) ===============
ShowGlobalShortcutsHelp() {
    global g_globalGui, g_globalShown

    ; Toggle off if currently shown
    if (IsObject(g_globalGui) && g_globalShown) {
        g_globalGui.Hide()
        g_globalShown := false
        ; Hotkey "Esc", "Off"  ; (disabled)
        return
    }

    ; Create the global shortcuts text with categories
    globalText := "
(
=== AVAILABLE (unused) ===
Win+Alt+Shift+C  â†’  Available
Win+Alt+Shift+1  â†’  Available

=== ONENOTE ===
Win+Alt+Shift+N  â†’  Opens or activates OneNote

=== SPOTIFY ===
Win+Alt+Shift+S  â†’  Opens or activates Spotify

=== CHATGPT ===
Win+Alt+Shift+8  â†’  Get word pronunciation, definition, and Portuguese translation
Win+Alt+Shift+0  â†’  Speak with ChatGPT
Win+Alt+Shift+7  â†’  Speak with ChatGPT (send message automatically)
Win+Alt+Shift+P  â†’  Copy last ChatGPT prompt
Win+Alt+Shift+J  â†’  Copy last ChatGPT message (in the editing box)
Win+Alt+Shift+U  â†’  Activate ChatGPT and copy last code block
Win+Alt+Shift+O  â†’  Check grammar and improve text in both English and Portuguese
Win+Alt+Shift+I  â†’  Opens ChatGPT
Win+Alt+Shift+L  â†’  Talk with ChatGPT through voice
Win+Alt+Shift+Y  â†’  Copy last message and read it aloud

=== YOUTUBE ===
Win+Alt+Shift+H  â†’  Activates Youtube

=== GOOGLE ===
Win+Alt+Shift+F  â†’  Opens Google

=== GMAIL ===
Win+Alt+Shift+W  â†’  Opens Gmail

=== CURSOR ===
Win+Alt+Shift+,  â†’  Opens or activates Cursor

=== OUTLOOK ===
Win+Alt+Shift+B  â†’  Open mail
Win+Alt+Shift+V  â†’  Open Reminder
Win+Alt+Shift+G  â†’  Open calendar

=== MICROSOFT TEAMS ===
Win+Alt+Shift+R  â†’  New conversation
Win+Alt+Shift+5  â†’  Toggle Mute (meeting)
Win+Alt+Shift+4  â†’  Toggle camera (meeting)
Win+Alt+Shift+T  â†’  Screen share (meeting)
Win+Alt+Shift+2  â†’  Exit meeting
Win+Alt+Shift+E  â†’  Select the chats window
Win+Alt+Shift+3  â†’  Select the meeting window

=== WHATSAPP ===
Win+Alt+Shift+Z  â†’  Opens WhatsApp

=== WINDOWS ===
Win+Alt+Shift+6  â†’  Minimizes windows
Win+Alt+Shift+M  â†’  Maximizes the current window

=== GENERAL ===
Win+Alt+Shift+Q  â†’  Jump mouse on the middle
Win+Alt+Shift+X  â†’  Activate hunt and Peck
Win+Alt+Shift+.  â†’  Set microphone volume to 100
Win+Alt+Shift+9  â†’  Pomodoro


=== SHORTCUTS ===
Win+Alt+Shift+A  â†’  Show app-specific shortcuts (quick press)
Win+Alt+Shift+A  â†’  Show global shortcuts (hold 700ms+)

=== WIKIPEDIA ===
Win+Alt+Shift+K  â†’  Opens or activates Wikipedia
)"

    static globalCtrl

    if !IsObject(g_globalGui) {
        g_globalGui := Gui(
            "+AlwaysOnTop -Caption +ToolWindow +Border +Owner +LastFound"
        )
        g_globalGui.BackColor := "000000"
        g_globalGui.SetFont("s10 cFFFF00", "Consolas")  ; Smaller font for more content
        globalCtrl := g_globalGui.Add("Edit", "ReadOnly +Multi +VScroll -HScroll -Border Background000000 w760 h540")

        ; Esc also hides  ; (disabled â€“ use Win+Alt+Shift+A to hide)
        ; Hotkey "Esc", (*) => (g_globalGui.Hide(), g_globalShown := false), "Off"
    }

    ; Fix mojibake (arrows, punctuation) then update text and show
    globalCtrl.Value := NormalizeMojibake(globalText)
    g_globalGui.Show("AutoSize Hide")
    CenterGuiOnActiveMonitor(g_globalGui)
    g_globalGui.Show("NoActivate")
    g_globalShown := true
    ; Hotkey "Esc", "On"  ; (disabled)
}

; ========== Hotkey with hold detection ====================================
; Win + Alt + Shift + A with hold detection
#!+a::
{
    global g_helpGui, g_helpShown, g_globalGui, g_globalShown

    ; First check if any cheat sheet is currently open - if so, close it
    if (IsObject(g_helpGui) && g_helpShown) {
        g_helpGui.Hide()
        g_helpShown := false
        ; Hotkey "Esc", "Off"  ; (disabled)
        return
    }

    if (IsObject(g_globalGui) && g_globalShown) {
        g_globalGui.Hide()
        g_globalShown := false
        ; Hotkey "Esc", "Off"  ; (disabled)
        return
    }

    ; No cheat sheet is open, determine which one to show based on hold time
    static pressTime := 0
    pressTime := A_TickCount

    ; Wait for key release or timeout (increased to accommodate 1s+ holds)
    KeyWait "a", "T1"  ; Wait max 1.5s for key release

    holdTime := A_TickCount - pressTime

    if (holdTime >= 700) {
        ; Long hold (1s+) - show global shortcuts
        ShowGlobalShortcutsHelp()
    } else {
        ; Quick press - show app-specific shortcuts
        ToggleShortcutHelp()
    }
}

;-------------------------------------------------------------------
; Environment paths (unchanged)
;-------------------------------------------------------------------
global WORK_SCRIPTS_PATH := "C:\Users\fie7ca\Documents\01 - Scripts"
global PERSONAL_SCRIPTS_PATH := "G:\Meu Drive\12 - Scripts"
; global IS_WORK_ENVIRONMENT   := true    ; set to false on personal rig // This will now be loaded from env.ahk

; ---------------------------------------------------------------------------
; ShowErr(msgOrErr)  â€“ uniform MsgBox for any thrown value
; ---------------------------------------------------------------------------
ShowErr(err) {
    text := (Type(err) = "Error") ? err.Message : err
    MsgBox("Error:`n" text, "Error", "IconX")
}

; ---------------------------------------------------------------------------
; Helper: centre a GUI over the active window (fallback: primary monitor)
; ---------------------------------------------------------------------------
CenterGuiOnActiveMonitor(guiObj) {
    ; Ensure GUI has its final size
    guiObj.GetPos(, , &guiW, &guiH)

    activeWin := WinGetID("A")
    ; Default to primary monitor work area (monitor #1)
    MonitorGetWorkArea(1, &lPrim, &tPrim, &rPrim, &bPrim)
    wx := lPrim, wy := tPrim, ww := rPrim - lPrim, wh := bPrim - tPrim

    ; If we know the active window, find which monitor contains its centre
    if (activeWin) {
        rect := Buffer(16, 0)
        if (DllCall("GetWindowRect", "ptr", activeWin, "ptr", rect)) {
            cx := NumGet(rect, 0, "int") + (NumGet(rect, 8, "int") - NumGet(rect, 0, "int")) // 2
            cy := NumGet(rect, 4, "int") + (NumGet(rect, 12, "int") - NumGet(rect, 4, "int")) // 2

            count := MonitorGetCount()
            loop count {
                idx := A_Index
                MonitorGetWorkArea(idx, &l, &t, &r, &b)
                if (cx >= l && cx <= r && cy >= t && cy <= b) {
                    wx := l, wy := t, ww := r - l, wh := b - t
                    break
                }
            }
        }
    }

    guiX := wx + (ww - guiW) / 2
    guiY := wy + (wh - guiH) / 2
    guiObj.Show("NoActivate x" Round(guiX) " y" Round(guiY))
}

;-------------------------------------------------------------------
; OneNote Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe onenote.exe")

; Shift + P : Onenote: select line and children
+p:: Send("^+-") ; Remaps to Ctrl + Shift + -

; Shift + F : Advanced Searching with double quotes
+f:: {
    Send "^f"
    Sleep 50
    Send "^a"
    Sleep 20
    Send "{Del}"
    Sleep 20
    Send '""'
    Sleep 20
    Send "{Left}"
}

; Shift + D : Onenote: delete line and children
+d:: {
    Send("^+-") ; Remaps to Ctrl + Shift + -
    Send "{Del}"
}

; Shift + S : Onenote: delete only current line (keep children)
+s:: {
    Send("+{Right}")
    Send "{Del}"
}

; Shift + U : Onenote: collapse
+u:: Send("!+{+}")     ; Remaps to Alt + Shift + +

; Shift + Y : Onenote: expand
+y:: Send("!+{-}")     ; Remaps to Alt + Shift + -

; Shift + I : Onenote: collapse all
+i:: Send("!+1")     ; Remaps to Alt + Shift + 1

; Shift + O : Onenote: expand all
+o:: Send("!+0")     ; Remaps to Alt + Shift + 0

#HotIf

;-------------------------------------------------------------------
; ClipAngel Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ClipAngel")

; Shift + Y : Select filtered content and copy
+y:: {
    Send "{Tab}"
    Sleep 100
    Send "{Tab}"
    Sleep 300
    Send "^a"  ; Select all
    Sleep 100
    Send "^c"  ; Copy
    Sleep 100
    Send "{F10}"
}

; Shift + U : Switch focus between list and text (F10)
+u:: Send "{F10}"

; Shift + I : Delete all non-favorite (Ctrl+Alt+K)
+i:: Send "^!k"

; Shift + O : Clear filters (F7)
+o:: Send "{F7}"

; Shift + P : Mark as favorite (Alt+Q)
+p:: Send "!q"

; Shift + H : Unmark as favorite (Alt+W)
+h:: Send "!w"

; Shift + J : Edit text (F4)
+j:: Send "{F4}"

; Shift + K : Save as file (Ctrl+S)
+k:: Send "^s"

; Shift + L : Merge clips
+l:: Send "^!j"

#HotIf

;-------------------------------------------------------------------
; WhatsApp Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("WhatsApp")

; Remap Shift+J to Ctrl+Alt+Shift+U
+j:: Send "^!+u"

; Remap Shift+K to Ctrl+Alt+Shift+P
+k:: Send "^!+p"

; Shift + U: Extended search (Alt+K)
+u:: Send("!k")

; Shift + I: Reply (Alt+R)
+i:: Send("!r")

; Shift + O: Sticker panel (Ctrl+Alt+S)
+o:: Send("^!s")

; Shift + P: Edit last message (Win+Up)
; The user updated this to be Toggle Unread
+p::
{
    try
    {
        uia := UIA_Browser()
        Sleep 300 ; Give UIA time to attach

        ; Find the "Unread" and "All" filter buttons
        unreadButton := uia.FindElement({ Name: "Unread", AutomationId: "unread-filter", Type: "TabItem" })
        allButton := uia.FindElement({ Name: "All", AutomationId: "all-filter", Type: "TabItem" })

        if (unreadButton && allButton) {
            ; Check if the "Unread" button is currently selected.
            ; The .IsSelected property is part of the SelectionItemPattern.
            if (unreadButton.IsSelected) {
                allButton.Click() ; If Unread is selected, click All
            }
            else {
                unreadButton.Click() ; Otherwise, click Unread
            }
        }
        else if (unreadButton) {
            ; Fallback if only the Unread button is found
            unreadButton.Click()
        }
        else {
            MsgBox "Could not find the 'Unread' filter button."
        }
    }
    catch Error as e {
        MsgBox "An error occurred: " e.Message
    }
}

global isRecording := false          ; persists between hotkey presses

; ---------- HOTKEY ----------------------------------------------------------
+y:: ToggleVoiceMessage()             ; Shift + Y

; ---------------------------------------------------------------------------
ToggleVoiceMessage() {
    global isRecording

    try {
        chrome := UIA_Browser()      ; top-level Chrome UIA element
        if !IsObject(chrome) {
            MsgBox "Can't attach to Chrome."
            return
        }

        Sleep 400                    ; let Chrome finish drawing

        ; Exact-name regexes (case-insensitive, anchored ^ $)
        voicePattern := "i)^(Voice message|Record voice message)$"
        sendPattern := "i)^(Send|Stop recording)$"

        ; Helper to grab a button by pattern
        FindBtn(p) => WaitForButton(chrome, p)

        if (isRecording) {           ; â–º we're supposed to stop & send
            if (btn := FindBtn(sendPattern)) {
                btn.Invoke()
                isRecording := false
            } else {
                ; Assume you clicked Send manually â†’ reset & start new rec
                isRecording := false
                if (btn := FindBtn(voicePattern)) {
                    btn.Invoke()
                    isRecording := true
                } else
                    MsgBox "Couldn't restart recording (Voice-message button missing)."
            }
        } else {                     ; â–º start recording
            if (btn := FindBtn(voicePattern)) {
                btn.Invoke()
                isRecording := true
            } else
                MsgBox "Couldn't find the Voice-message button."
        }
    } catch Error as err {
        MsgBox "Error:`n" err.Message
    }
}

; ---------------------------------------------------------------------------
; WaitForButton(root, pattern, timeout := 5000)
;   â€¢ Searches all descendant buttons of `root` until Name matches `pattern`
;   â€¢ Returns the UIA element or 0 if none matched within `timeout` ms
; ---------------------------------------------------------------------------
WaitForButton(root, pattern, timeout := 5000) {
    if !IsObject(root)
        return 0
    deadline := A_TickCount + timeout
    while (A_TickCount < deadline) {
        for btn in root.FindAll({ Type: "Button" }) {
            if RegExMatch(btn.Name, pattern)
                return btn
        }
        Sleep 150
    }
    return 0
}

; ---------------------------------------------------------------------------
; WaitForList(root, pattern := "", timeout := 5000)
;   â€¢ Searches descendant List controls; Name must match `pattern` if provided
;   â€¢ Returns the UIA element or 0 after `timeout` ms
; ---------------------------------------------------------------------------
WaitForList(root, pattern := "", timeout := 5000) {
    if !IsObject(root)
        return 0
    deadline := A_TickCount + timeout
    while (A_TickCount < deadline) {
        for lst in root.FindAll({ Type: "List" }) {
            if (!pattern || RegExMatch(lst.Name, pattern))
                return lst
        }
        Sleep 150
    }
    return 0
}

; Shift + h: Focus the current conversation
+h::
{
    try
    {
        ; WhatsApp desktop is Chromium-based, so we can use UIA_Browser.
        ; It should attach to the active window, which is WhatsApp thanks to #HotIf.
        uia := UIA_Browser()
        Sleep 300 ; Give UIA time to attach to the browser. A similar delay is in the reference script.

        ; Find the "Archived" button to use as an anchor.
        ; The user provided: Name:"Archived "
        archivedButton := uia.FindElement({ Name: "Archived ", Type: "Button" })

        if (archivedButton) {
            ; Focus the button without clicking it.
            archivedButton.SetFocus()
            ; Send Tab to move to the main conversation list.
            ; From there, the focus should be on the selected chat.
            SendInput "{Tab}"
        }
        else {
            MsgBox "Could not find the 'Archived' button."
        }
    }
    catch Error as e {
        MsgBox "An error occurred while trying to focus WhatsApp conversation: " e.Message
    }
}

#HotIf

;-------------------------------------------------------------------
; Outlook Reminder Window Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe OUTLOOK.EXE") && RegExMatch(WinGetTitle("A"), "i)Reminder")

; Shift + Y : Select first reminder list item
+Y::
{
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Find the first list item in the reminder window
        ; Looking for ListItem type (50007) with AutomationId pattern "ListViewItem-0"
        firstItem := root.FindFirst({ AutomationId: "ListViewItem-0", ControlType: "ListItem" })

        ; Fallback: if specific AutomationId doesn't work, find first ListItem
        if !firstItem {
            firstItem := root.FindFirst({ ControlType: "ListItem" })
        }

        ; Fallback: try finding by type number if ControlType doesn't work
        if !firstItem {
            firstItem := root.FindFirst({ Type: "50007" })
        }

        if firstItem {
            firstItem.Select()
            ; Alternative method: click the item to ensure selection
            ; firstItem.Click()
        }
        else {
            MsgBox("Could not find the first reminder item.", "Reminder Selection", "IconX")
        }
    }
    catch Error as e {
        MsgBox("UIA error:`n" e.Message, "Outlook Reminder Error", "IconX")
    }
}

; ativa a janela de lembretes do Outlook
ActivateReminder() {
    WinActivate("ahk_exe OUTLOOK.EXE")
    WinWaitActive("ahk_exe OUTLOOK.EXE", , 1)
}

; digita o tempo e aperta Alt+S
QuickSnooze(t) {
    ActivateReminder()
    Send("{Tab}")              ; chega ao combo
    Send("{Tab}")              ; chega ao combo
    Sleep 100
    Send("^a{Delete}" . t)     ; substitui o texto
    Sleep 120
    Send("!s")                 ; Alt+S = Snooze
    Sleep 200
    Send("{Tab}")
    Send("{Tab}")
    Send("{Tab}")
}

; caixa de confirmaÃ§Ã£o antes de executar
Confirm(t) {
    if MsgBox("Snooze for " t "?", "Confirm Snooze", "YesNo Icon?") = "Yes"
        QuickSnooze(t)
}

; HOTKEYS de teste
; Shift+U  â†’ 1 minuto
; Shift+I  â†’ 2 minutos
; Shift+O  â†’ 3 minutos
+U:: Confirm("1 hour")
+I:: Confirm("4 hours")
+O:: Confirm("1 day")
+P:: ConfirmDismissAll()  ; Shift + P : Dismiss all reminders with confirmation

; Shift + H : Join Online
+H::
{
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Find the "Join Online" button
        ; Type: 50000 (Button), Name: "Join Online", AutomationId: "8346", ClassName: "Button"
        joinButton := root.FindFirst({ Name: "Join Online", Type: "50000", AutomationId: "8346" })

        ; Fallback: try finding by name only
        if !joinButton {
            joinButton := root.FindFirst({ Name: "Join Online", ControlType: "Button" })
        }

        ; Fallback: try finding by AutomationId only
        if !joinButton {
            joinButton := root.FindFirst({ AutomationId: "8346" })
        }

        if joinButton {
            joinButton.Click()
        }
        else {
            ; No message box as requested - fail silently
        }
    }
    catch Error as e {
        ; No message box as requested - fail silently
    }
}

#HotIf

;-------------------------------------------------------------------
; Microsoft Teams Helper functions
;-------------------------------------------------------------------
IsTeamsMeetingTitle(title) {
    if InStr(title, "Chat |") || InStr(title, "Sharing control bar |")
        return false
    if InStr(title, "Microsoft Teams meeting")
        return true
    return RegExMatch(title, "i)^.*\| Microsoft Teams.*$")
}

IsTeamsChatTitle(title) {
    if InStr(title, "Sharing control bar |") || InStr(title, "Microsoft Teams meeting")
        return false
    return InStr(title, "Chat |") && RegExMatch(title, "i)\| Microsoft Teams$")
}

; -------------------------------------------------------------------
; Helper predicates to detect which Teams window is active
; -------------------------------------------------------------------
IsTeamsMeetingActive() {
    return IsTeamsMeetingTitle(WinGetTitle("A"))
}
IsTeamsChatActive() {
    return IsTeamsChatTitle(WinGetTitle("A"))
}

; -------------------------------------------------------------------
; Microsoft Teams Shortcuts â€“ MEETING WINDOW
; -------------------------------------------------------------------
#HotIf IsTeamsMeetingActive()

; Shift + Y : Open Chat pane
+Y:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        btn := root.FindFirst({ AutomationId: "chat-button" })
        if !btn
            btn := root.FindFirst({ Name: "Chat", ControlType: "Button" })

        if btn
            btn.Click()
        else
            MsgBox("Couldn't find the Chat button.", "Control not found", "IconX")
    }
    catch as e {
        MsgBox("UIA error:`n" e.Message, "Error", "IconX")
    }
}

; Shift + U : Switch from compacted to normal Teams meeting window
+U:: {
    ; Get current active window title
    currentTitle := WinGetTitle("A")

    ; Check if current window is a compacted Teams meeting
    isCompacted := false
    if (currentTitle = "Reunião do Microsoft Teams | Microsoft Teams") {
        isCompacted := true
        baseTitle := "Reunião do Microsoft Teams"
    } else if (InStr(currentTitle,
        "Modo de exibição compacto da reunião | Reunião do Microsoft Teams | Microsoft Teams")) {
        isCompacted := true
        baseTitle := "Reunião do Microsoft Teams"
    }

    if (!isCompacted) {
        ; Not a compacted meeting window, do nothing
        return
    }

    ; Search for the corresponding normal meeting window
    normalMeetingHwnd := 0
    for hwnd in WinGetList("ahk_exe ms-teams.exe") {
        title := WinGetTitle(hwnd)

        ; Skip if it's the same window or another compacted window
        if (hwnd = WinGetID("A") || InStr(title, "Modo de exibição compacto da reunião")) {
            continue
        }

        ; Check if it's a normal meeting window with the same base title
        if (InStr(title, baseTitle) && InStr(title, "| Microsoft Teams") && !InStr(title,
            "Modo de exibição compacto da reunião")) {
            normalMeetingHwnd := hwnd
            break
        }
    }

    ; If found, switch to the normal meeting window
    if (normalMeetingHwnd) {
        try {
            WinActivate("ahk_id " normalMeetingHwnd)
            ; Optional: Show a brief tooltip to confirm the switch
            ToolTip("Switched to normal meeting view")
            SetTimer(() => ToolTip(), -1000) ; Hide tooltip after 1 second
        } catch as e {
            ; Fallback: try to bring window to front
            WinShow("ahk_id " normalMeetingHwnd)
            WinActivate("ahk_id " normalMeetingHwnd)
        }
    } else {
        ; No corresponding normal window found - show notification
        ToolTip("No normal meeting window found")
        SetTimer(() => ToolTip(), -1500) ; Hide tooltip after 1.5 seconds
    }
}

; Shift + I : Reagir (open reactions menu)
+I:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        btn := 0
        try {
            btn := root.FindFirst({ AutomationId: "reaction-menu-button" })
        } catch {
        }
        if !btn {
            try {
                btn := root.FindFirst({ Name: "Reagir", ControlType: "Button" })
            } catch {
            }
        }

        if btn
            btn.Click()
        else
            MsgBox("Couldn't find the Reagir button.", "Control not found", "IconX")
    }
    catch as e {
        MsgBox("UIA error:`n" e.Message, "Error", "IconX")
    }
}

; Shift + O : Join now with camera and microphone on
+O:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Find the "Join now" button by AutomationId first
        btn := root.FindFirst({ AutomationId: "prejoin-join-button" })

        ; Fallback: try finding by name
        if !btn {
            btn := root.FindFirst({ Name: "Ingressar agora Com a cÃ¢mera ligada e Microfone ligado", ControlType: "Button" })
        }

        ; Fallback: try finding by partial name (in case the name varies)
        if !btn {
            btn := root.FindFirst({ Name: "Ingressar agora", ControlType: "Button" })
        }

        if btn {
            btn.Click()
        }
        ; No message box as requested - fail silently
    }
    catch {
        ; No message box as requested - fail silently
    }
}

; Shift + P : Audio settings
+P:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Find the audio settings button by AutomationId first
        btn := root.FindFirst({ AutomationId: "prejoin-audiosettings-button" })

        ; Fallback: try finding by name
        if !btn {
            btn := root.FindFirst({ Name: "Microfone do computador e controles do alto-falante ConfiguraÃ§Ãµes de Ã¡udio",
                ControlType: "Button" })
        }

        ; Fallback: try finding by partial name
        if !btn {
            btn := root.FindFirst({ Name: "ConfiguraÃ§Ãµes de Ã¡udio", ControlType: "Button" })
        }

        if btn {
            btn.Click()
        }
        ; No message box as requested - fail silently
    }
    catch {
        ; No message box as requested - fail silently
    }
}

#HotIf

;-------------------------------------------------------------------
; Wikipedia Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe chrome.exe") && InStr(WinGetTitle("A"), "Wikipedia", false)

; Shift + Y: Focus the Wikipedia search field (Type 50003 ComboBox)
+y::
{
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 200 ; Give UIA time to attach

        ; Prefer document; fall back to browser root
        try {
            root := uia.GetCurrentDocumentElement()
        } catch {
            root := uia.BrowserElement
        }

        searchBox := 0

        ; Simple, attribute-driven matching
        try {
            searchBox := root.FindElement({ Type: 50003, Name: "Search Wikipedia" })
        } catch {
        }
        if (!searchBox) {
            try {
                searchBox := root.FindElement({ Type: 50003, AcceleratorKey: "Alt+f" })
            } catch {
            }
        }
        if (!searchBox) {
            try {
                searchBox := root.FindElement({ Type: 50003, ClassName: "cdx-text-input__input" })
            } catch {
            }
        }
        if (!searchBox) {
            try {
                searchBox := root.FindElement({ Type: 50003, Name: "Search", mm: 2, cs: false })
            } catch {
            }
        }

        if (searchBox) {
            try {
                searchBox.SetFocus()
            } catch {
                searchBox.Click()
            }
        } else {
            ; Try focusing the profile link, Shift+Tab to Search, then activate
            profileLink := 0
            try {
                profileLink := root.FindElement({ Type: 50005, Name: "Duducm2", cs: false })
            } catch {
            }
            if (!profileLink) {
                try {
                    profileLink := root.FindElement({ Type: 50005, Name: "doodoocm2", cs: false })
                } catch {
                }
            }
            if (!profileLink) {
                try {
                    profileLink := root.FindElement({ ControlType: "Hyperlink", Name: "Duducm2", cs: false })
                } catch {
                }
            }

            if (profileLink) {
                try {
                    profileLink.SetFocus()
                    Sleep 100
                    uia.ControlSend("+{Tab}")
                    Sleep 120
                    uia.ControlSend("{Enter}")
                    Sleep 150
                } catch {
                }

                ; Retry locating the search field after activating Search
                try {
                    searchBox := root.FindElement({ Type: 50003, Name: "Search Wikipedia" })
                } catch {
                }
                if (!searchBox) {
                    try {
                        searchBox := root.FindElement({ Type: 50003, ClassName: "cdx-text-input__input" })
                    } catch {
                    }
                }
                if (!searchBox) {
                    try {
                        searchBox := root.FindElement({ Type: 50003, Name: "Search", mm: 2, cs: false })
                    } catch {
                    }
                }

                if (searchBox) {
                    try {
                        searchBox.SetFocus()
                    } catch {
                        searchBox.Click()
                    }
                    return
                }
            }
            ; Removed UIA click on Search button per request

            ; Final fallback: use the accelerator key
            try {
                uia.ControlSend("!f")
                return
            } catch {
            }
            MsgBox "Could not find the 'Search Wikipedia' field."
        }
    }
    catch Error as e {
        MsgBox "An error occurred: " e.Message
    }
}

#HotIf

;-------------------------------------------------------------------
; Mercado Livre (Brazil) Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe chrome.exe") && InStr(WinGetTitle("A"), "Mercado Livre", false)

; Shift + Y: Focus Mercado Livre search field
+y::
{
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 200
        ; Prefer the document root; fall back to browser element
        try {
            root := uia.GetCurrentDocumentElement()
        } catch {
            root := uia.BrowserElement
        }

        field := 0
        ; Try AutomationId first
        try {
            field := root.FindElement({ AutomationId: "cb1-edit" })
        } catch {
        }
        if (!field) {
            ; Try by ComboBox name (Type 50003)
            try {
                field := root.FindElement({ Type: 50003, Name: "Digite o que vocÃª quer encontrar", cs: false })
            } catch {
            }
        }
        if (!field) {
            ; Try by Edit control name
            try {
                field := root.FindElement({ Type: "Edit", Name: "Digite o que vocÃª quer encontrar", cs: false })
            } catch {
            }
        }
        if (!field) {
            ; Try by class name
            try {
                field := root.FindElement({ ClassName: "nav-search-input" })
            } catch {
            }
        }

        if (field) {
            try {
                field.SetFocus()
            } catch {
                field.Click()
            }
            return
        } else {
            MsgBox "Could not find Mercado Livre search field."
        }
    } catch Error as e {
        MsgBox "An error occurred: " e.Message
    }
}

; Shift + U: Carrinho de compras (Cart)
+u::
{
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 200
        try {
            root := uia.GetCurrentDocumentElement()
        } catch {
            root := uia.BrowserElement
        }

        cart := 0
        ; Prefer AutomationId
        try {
            cart := root.FindElement({ AutomationId: "nav-cart" })
        } catch {
        }
        if (!cart) {
            ; Try by class name substring
            try {
                cart := root.FindElement({ ClassName: "nav-cart", matchmode: "Substring" })
            } catch {
            }
        }
        if (!cart) {
            ; Try by link name containing 'carrinho'
            try {
                cart := root.FindElement({ Type: 50005, Name: "carrinho", cs: false, matchmode: "Substring" })
            } catch {
            }
        }

        if (cart) {
            try cart.Invoke()
            catch {
                try cart.Click()
            }
            return
        }
        MsgBox "Could not find Mercado Livre cart link."
    } catch Error as e {
        MsgBox "An error occurred: " e.Message
    }
}

; Shift + I: Compras (Purchases)
+i::
{
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 200
        try {
            root := uia.GetCurrentDocumentElement()
        } catch {
            root := uia.BrowserElement
        }

        purchases := 0
        ; Try by class name first
        try {
            purchases := root.FindElement({ ClassName: "option-purchases" })
        } catch {
        }
        if (!purchases) {
            ; Try by link name 'Compras'
            try {
                purchases := root.FindElement({ Type: 50005, Name: "Compras", cs: false, matchmode: "Substring" })
            } catch {
            }
        }

        if (purchases) {
            try purchases.Invoke()
            catch {
                try purchases.Click()
            }
            return
        }
        MsgBox "Could not find Mercado Livre purchases link."
    } catch Error as e {
        MsgBox "An error occurred: " e.Message
    }
}

#HotIf

;-------------------------------------------------------------------
; Microsoft Teams Shortcuts (chat)
;-------------------------------------------------------------------
#HotIf IsTeamsChatActive()

; Shift + K : Pin
+K::
{
    Sleep "150"
    Send "^1"
    Sleep "100"
    Send("{AppsKey}")
    Sleep "100"
    Send "{Down}"
    Send "{Down}"
    Send "{Right}"
    Send "{Enter}"
    Send "{Esc}"
    Sleep "200"
    Send "^1"
    Sleep "500"          ; 80 ms
    Send "^+{Home}"
}

; Shift + Y : Like
+y::
{
    Send "{Enter}"
    Send "{Enter}"
    Send "{Esc}"
}

; Shift + U : Heart
+u::
{
    Send "{Enter}"
    Send "{Down}"
    Send "{Enter}"
    Send "{Esc}"
}

; Shift + I : Laugh
+i::
{
    Send "{Enter}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
    Send "{Esc}"
}

; Shift + Ã‡ : Remove pin
+l::
{
    Sleep "150"
    Send "^1"
    Sleep "100"
    Send("{AppsKey}")
    Sleep "100"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Shift + R : Reply Message
+r::
{
    Send "{Enter}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Shift + E : Edit message
+e::
{
    Send "{Enter}"
    Sleep 100
    Send "{Left}"
    Sleep 100
    Send "{Left}"
    Sleep 100
    Send "{Enter}"
}

; Shift + J : Mark as unread
+j::
{
    Send "^1"
    Sleep "220"
    Send("{AppsKey}")
    Sleep "220"
    Send "{Down}"
    Send "{Enter}"
}

; Shift + O : Ctrl+1 then Ctrl+Shift+Home
+o::
{
    Send "^1"
    Sleep "80"          ; 80 ms
    Send "^+{Home}"
}

#HotIf

;-------------------------------------------------------------------
; Outlook Shortcuts
;-------------------------------------------------------------------

; Helper predicates for Outlook window types
IsOutlookMessageActive() {
    return WinActive("ahk_exe OUTLOOK.EXE")
    && RegExMatch(WinGetTitle("A"), "i) - Message \(")
}

IsOutlookAppointmentActive() {
    return WinActive("ahk_exe OUTLOOK.EXE")
    && RegExMatch(WinGetTitle("A"), "i)(Appointment|Meeting)")
}

IsOutlookReminderActive() {
    return WinActive("ahk_exe OUTLOOK.EXE")
    && RegExMatch(WinGetTitle("A"), "i)Reminder")
}

IsOutlookMainActive() {
    if !WinActive("ahk_exe OUTLOOK.EXE")
        return false
    t := WinGetTitle("A")
    ; Exclude inspectors and reminders
    if RegExMatch(t, "i) - Message \(")
        return false
    if RegExMatch(t, "i)(Appointment|Meeting)")
        return false
    if RegExMatch(t, "i)Reminder")
        return false
    return true
}

#HotIf IsOutlookMainActive()

; Shift + U : Send to general
+Y::
{
    Send "!5"
    Send "O"
    Send "{Home}"
    Send "00"
    Send "{Enter}"
}

; Shift + U : Send to newsletter
+U::
{
    Send "!5"
    Send "O"
    Send "{Home}"
    Send "01"
    Send "{Enter}"
}

; (Main window has no inspector navigation hotkeys)

; Shift + I : Go to Inbox
+I::
{
    Send "{Alt}"
    Sleep 60
    Send "6"
    Sleep 80
    Send "^{Home}"
    Sleep 100
    Send "i"
    Sleep 50
    Send "n"
    Sleep 50
    Send "{Enter}"
}

+N:: {                                  ; toggle Focused / Other
    static nextOutlookButton := "Other"

    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        btn := root.FindFirst({
            Name: nextOutlookButton, Type: "Button"
        })

        if btn {
            btn.Click()
            nextOutlookButton := (nextOutlookButton = "Other")
                ? "Focused" : "Other"
        } else {
            MsgBox("Couldn't find '" nextOutlookButton "'.", "Button not found", "IconX")
        }

    } catch Error as err {              ; â† **only this form**
        ShowErr(err)
    }
}

; -------------------------------------------------------------------
; Focus helpers â€“ reuse for any field you need
; -------------------------------------------------------------------
FocusOutlookField(criteria) {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)
        ctrl := root.FindFirst(criteria)
        if ctrl {
            ctrl.SetFocus()
            return true
        }
    } catch Error {
    }
    return false
}

; -------------------------------------------------------------------
; Click helper â€“ try AutomationId first, then Name+ClassName
; -------------------------------------------------------------------
ClickOutlookByIdThenNameClass(automationId, name, className, controlType := "") {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        if (automationId) {
            el := root.FindFirst({ AutomationId: automationId })
            if (el) {
                el.SetFocus()
                Sleep 50
                el.Click()
                return true
            }
        }

        crit := { Name: name }
        if (className)
            crit.ClassName := className
        if (controlType)
            crit.ControlType := controlType

        el := root.FindFirst(crit)
        if (el) {
            el.SetFocus()
            Sleep 50
            el.Click()
            return true
        }
    } catch Error as err {
        ShowErr(err)
    }
    return false
}

; -------------------------------------------------------------------
; General helper â€“ visually confirm focus on the selected element
; Sends Down then Up to force a visible focus cue
; -------------------------------------------------------------------
EnsureFocus() {
    Send "{Down}"
    Send "{Up}"
}

; Helper: Select the first pinned item in Explorer sidebar (Navigation Pane)
; Global so it can be reused by both Explorer and File Dialog contexts
SelectExplorerSidebarFirstPinned() {
    try {
        explorerEl := UIA.ElementFromHandle(WinExist("A"))

        ; Look for the navigation pane (sidebar) - it's typically a Tree control
        navPane := explorerEl.FindFirst({ Type: "Tree" })

        if (navPane) {
            ; If in work environment, prefer selecting the Home tree item directly
            try {
                global IS_WORK_ENVIRONMENT
                if (IS_WORK_ENVIRONMENT) {
                    homeItem := navPane.FindFirst({ Type: "TreeItem", Name: "Home" })
                    if (homeItem) {
                        homeItem.ScrollIntoView()
                        homeItem.Select()    ; select only, no click
                        homeItem.SetFocus()
                        EnsureFocus()
                        return true
                    }
                }
            } catch Error {
                ; ignore and fallback to previous logic
            }
            ; Define the keywords to search for pinned items
            pinnedKeywords := ["fixo", "pinned", "pin", "fixado", "fixada", "fixar", "preso"]

            ; Search for the first TreeItem that contains any of the pinned keywords
            firstPinnedItem := unset
            for keyword in pinnedKeywords {
                firstPinnedItem := navPane.FindFirst({ Type: "TreeItem", Name: keyword, matchmode: "Substring" })
                if (firstPinnedItem)
                    break
            }

            if (firstPinnedItem) {
                firstPinnedItem.ScrollIntoView()
                firstPinnedItem.Select()
                firstPinnedItem.SetFocus()
                EnsureFocus()
                return true
            }

            ; If we didn't find a pinned item, at least focus the tree and press Home
            navPane.SetFocus()
            Sleep 100
            Send "{Home}"
            EnsureFocus()
            return false
        }
    } catch Error {
        ; swallow and continue to fallback
    }

    ; Robust fallback â€“ cycle through panes up to 6 times to reach navigation, then Home
    loop 6 {
        Send "{F6}"
        Sleep 120
        try {
            explorerEl := UIA.ElementFromHandle(WinExist("A"))
            navPane := explorerEl.FindFirst({ Type: "Tree" })
            if (navPane && navPane.HasKeyboardFocus) {
                Send "{Home}"
                EnsureFocus()
                return false
            }
        } catch Error {
        }
    }
    ; Last resort â€“ send Home anyway
    Send "{Home}"
    EnsureFocus()
    return false
}

; Message inspector-specific hotkeys (Subject/To/DatePicker/Body)
#HotIf IsOutlookMessageActive()

; Shift + Y â†’ Subject
+Y:: {
    if FocusOutlookField({ AutomationId: "4101" }) ; Subject
        return
    if FocusOutlookField({ Name: "Subject", ControlType: "Edit" })
        return
}

; Shift + U â†’ Required / To
+U:: {
    if FocusOutlookField({ AutomationId: "4109" }) ; Required
        return
    if FocusOutlookField({ Name: "Required", ControlType: "Edit" })
        return
    if FocusOutlookField({ AutomationId: "4117" }) ; To
        return
    if FocusOutlookField({ Name: "To", ControlType: "Edit" })
        return
}

; Shift + I â†’ Date Picker (if present in this inspector)
; (No Shift + I in Message inspector)

; Shift + M â†’ Subject â†’ Body
+M:: {
    if FocusOutlookField({ AutomationId: "4101" }) {
        Sleep 50
        Send "{Tab}"
        return
    }
    if FocusOutlookField({ Name: "Subject", ControlType: "Edit" }) {
        Sleep 50
        Send "{Tab}"
        return
    }
}

#HotIf

; Appointment/Meeting inspector-specific hotkeys
#HotIf IsOutlookAppointmentActive()

; ----- Outlook Appointment: Date/Time helpers -----
Outlook_ClickStartDate() {
    ClickOutlookByIdThenNameClass("4098", "Start date, combo", "RichEdit20WPT", "Edit")
}

Outlook_ClickStartDatePicker() {
    ; Robust open: focus the Date Picker and press Enter
    if FocusOutlookField({ AutomationId: "4352" }) {
        Sleep 80
        Send "{Enter}"
        return
    }
    if FocusOutlookField({ Name: "Date Picker", ControlType: "Button" }) {
        Sleep 80
        Send "{Enter}"
        return
    }
}

Outlook_ClickStartTime() {
    ClickOutlookByIdThenNameClass("4096", "Start time, combo", "RichEdit20WPT", "Edit")
}

Outlook_ClickStartTime_1100AM() {
    ; Clicks the button showing 11:00 AM (start)
    ClickOutlookByIdThenNameClass("4354", "11:00 AM", "AfxWndW", "Button")
}

Outlook_ClickEndDate() {
    ClickOutlookByIdThenNameClass("4099", "End date, combo", "RichEdit20WPT", "Edit")
}

Outlook_ClickEndDatePicker() {
    ; Date Picker next to End date
    ClickOutlookByIdThenNameClass("4353", "Date Picker", "AfxWndW", "Button")
}

Outlook_ClickEndTime() {
    ClickOutlookByIdThenNameClass("4097", "End time, combo", "RichEdit20WPT", "Edit")
}

Outlook_ClickEndTime_1200PM() {
    ; Clicks the button showing 12:00 PM (end)
    ClickOutlookByIdThenNameClass("4355", "12:00 PM", "AfxWndW", "Button")
}

; (moved +H/+J below to ensure the block starts with Y)

; New Shift hotkeys for date/time controls (Start â†’ End), ordered by key preference
; Shift + Y â†’ Start date (combo)
+Y:: {
    Outlook_ClickStartDate()
}

; Shift + U â†’ Start date â€“ Date Picker
+U:: {
    Outlook_ClickStartDatePicker()
}

; Shift + I â†’ Start time (combo)
+I:: {
    Outlook_ClickStartTime()
}

; Shift + O â†’ End date (combo)
+O:: {
    Outlook_ClickEndDate()
}

; Shift + P â†’ End time (combo)
+P:: {
    Outlook_ClickEndTime()
}

; Shift + H â†’ All day checkbox
+H:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        checkbox := root.FindFirst({ AutomationId: "4226", ControlType: "CheckBox" })
        if !checkbox
            checkbox := root.FindFirst({ Name: "All day", ControlType: "CheckBox" })

        if checkbox {
            checkbox.Invoke()
        } else {
            MsgBox "Couldn't find the All day checkbox.", "Control not found", "IconX"
        }
    } catch Error as err {
        ShowErr(err)
    }
}

; Shift + J â†’ Title field
+J:: {
    if FocusOutlookField({ AutomationId: "4100" }) ; Title
        return
    if FocusOutlookField({ Name: "Title", ControlType: "Edit" })
        return
}

; Shift + L â†’ Required / To field
+L:: {
    if FocusOutlookField({ AutomationId: "4109" }) ; Required
        return
    if FocusOutlookField({ Name: "Required", ControlType: "Edit" })
        return
}

; Shift + M â†’ Location â†’ Body
+M:: {
    if FocusOutlookField({ AutomationId: "4111" }) { ; Location
        Sleep 100
        Send "{Tab}"
        return
    }
    if FocusOutlookField({ Name: "Location", ControlType: "Edit" }) {
        Sleep 100
        Send "{Tab}"
        return
    }
}

; Shift + ; â†’ Make Recurring (click)
; Semicolon key virtual key code is VK_BA; with Shift it's +vkBA
+,:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        btn := root.FindFirst({ AutomationId: "4364", ControlType: "Button" })
        if !btn
            btn := root.FindFirst({ Name: "Make Recurring", ControlType: "Button" })

        if btn {
            btn.Invoke()
        } else {
            MsgBox "Couldn't find the Make Recurring button.", "Control not found", "IconX"
        }
    } catch Error as err {
        ShowErr(err)
    }
}

#HotIf

;-------------------------------------------------------------------
; Google Chrome Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe chrome.exe")

; Shift + G : Pop up the current tab
+g::
{
    Send "{F6}"                        ; Focus address bar (omnibox)
    Sleep 100
    Send "{F6}"                        ; Focus the tab strip (current tab)
    Sleep 100
    Send "{AppsKey}"                   ; Open the tab's context menu (AppsKey or Shift+F10)
    Sleep 100                          ; Wait a moment for menu to open
    Send "m"                           ; Select "Move tab to new window" (press 'm')
    Sleep 100
    Send "{Enter}"                     ; Confirm the action (detach tab)
    Sleep 100
    Send "{Enter}"                     ; Confirm the action (detach tab)
}

#HotIf

;-------------------------------------------------------------------
; ChatGPT Shortcuts
;-------------------------------------------------------------------
#HotIf (hwnd := GetChatGPTWindowHwnd()) && WinActive("ahk_id " hwnd)

; Shift + Y : Select all and cut
+y::
{
    Send "^a"
    Sleep 50
    Send "^x"
}

; Shift + U â†’ click ChatGPT's model selector (any language)
+u:: {
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 300                 ; brief settle-time

        ; Find the model selector button containing "model selector" text
        modelCtl := uia.FindElement({
            Name: "Model selector",
            Type: "Button",
            matchmode: "Substring"
        })

        ; Click or complain
        if (modelCtl)
            modelCtl.Click()
        else
            MsgBox "Couldn't find the model-selector button."
    }
    catch Error as e {
        MsgBox "UIA error: " e.Message
    }
}

; Shift + I: Toggle sidebar
+i:: Send("^+s")

; Shift + O : Re-send rules & ask ChatGPT to correct mistake
+o::
{
    ; Ensure composer is focused
    Send "{Esc}"
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
    Send "{Enter}"
    Sleep 100
    A_Clipboard := oldClip

    ; Step 3: Alt+Tab to previous window
    Send "!{Tab}"

    ; After sending, show loading for Stop streaming
    buttonNames := ["Stop streaming", "Interromper transmissÃ£o"]
    WaitForButtonAndShowSmallLoading_ChatGPT(buttonNames, "Waiting for response...")
}

; Shift + H: Copy last code block
+h:: Send("^+;")

; Shift + L: Send and show AI banner
+l:: {
    ; --- Button Names (EN/PT) ---
    pt_stopStreamingName := "Interromper transmissÃ£o"
    en_stopStreamingName := "Stop streaming"
    currentStopStreamingName := IS_WORK_ENVIRONMENT ? pt_stopStreamingName : en_stopStreamingName

    ; Step 1: Send Escape to ensure composer is focused
    Send "{Esc}"
    Sleep 100
    ; Step 2: Send Enter to submit the prompt
    Send "{Enter}"
    Sleep 100
    ; Step 3: Alt+Tab to previous window
    Send "!{Tab}"
    Sleep 300
    ; Step 4: Show banner immediately (debounced by helper), then wait for completion to auto-hide and chime
    ShowSmallLoadingIndicator_ChatGPT("AI is respondingâ€¦")
    ; Use infinite timeout so the banner persists for long responses
    WaitForButtonAndShowSmallLoading_ChatGPT([currentStopStreamingName, "Stop", "Interromper"], "AI is respondingâ€¦",
    0)
}

#HotIf

;-------------------------------------------------------------------
; Settings Window Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("Settings") || WinActive("ConfiguraÃ§Ãµes")

; Shift + Y : Set input volume to 100%
+Y::
{
    try {
        ; Get the active Settings window
        settingsHwnd := WinExist("A")
        settingsRoot := UIA.ElementFromHandle(settingsHwnd)

        ; Try to find the input volume slider by AutomationId first (most reliable)
        volumeSlider := ""
        try {
            volumeSlider := settingsRoot.FindFirst({ AutomationId: "SystemSettings_Audio_Input_VolumeValue_Slider",
                ControlType: "Slider" })
        } catch {
            ; Fallback: Try by name (both English and Portuguese)
            sliderNames := ["Input volume", "Ajustar o volume de entrada"]
            for sliderName in sliderNames {
                try {
                    volumeSlider := settingsRoot.FindFirst({ Name: sliderName, ControlType: "Slider" })
                    if volumeSlider
                        break
                } catch {
                    continue
                }
            }
        }

        if volumeSlider {
            ; Set slider value to maximum (100)
            volumeSlider.SetValue(100)
            ; Optional: Brief confirmation
            ToolTip("Input volume set to 100%")
            SetTimer(() => ToolTip(), -1000)

            ; Close the Settings window
            Sleep(300) ; Small delay to ensure setting is applied
            Send("!{F4}") ; Alt+F4 to close Settings
        } else {
            MsgBox("Input volume slider not found. Make sure you're on the microphone settings page.", "Error", "IconX"
            )
        }

    } catch Error as e {
        MsgBox("Error setting input volume: " . e.Message, "Error", "IconX")
    }
}

#HotIf

;-------------------------------------------------------------------
; Windows Explorer Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe explorer.exe")

; Explorer-specific helper â€“ select first pinned item in the sidebar
SelectExplorerSidebarFirstPinned_EX() {
    try {
        explorerEl := UIA.ElementFromHandle(WinExist("A"))
        navPane := explorerEl.FindFirst({ Type: "Tree" })
        if (navPane) {
            ; If in work environment, prefer selecting the Home tree item directly
            try {
                global IS_WORK_ENVIRONMENT
                if (IS_WORK_ENVIRONMENT) {
                    homeItem := navPane.FindFirst({ Type: "TreeItem", Name: "Home" })
                    if (homeItem) {
                        homeItem.ScrollIntoView()
                        homeItem.Select()    ; select only, no click
                        homeItem.SetFocus()
                        EnsureFocus()
                        return true
                    }
                }
            } catch Error {
                ; ignore and fallback to previous logic
            }
            pinnedKeywords := ["fixo", "pinned", "pin", "fixado", "fixada", "fixar", "preso"]
            firstPinnedItem := unset
            for keyword in pinnedKeywords {
                firstPinnedItem := navPane.FindFirst({ Type: "TreeItem", Name: keyword, matchmode: "Substring" })
                if (firstPinnedItem)
                    break
            }
            if (firstPinnedItem) {
                firstPinnedItem.ScrollIntoView()
                firstPinnedItem.Select()
                firstPinnedItem.SetFocus()
                EnsureFocus()
                return true
            }
        }
    } catch Error {
    }
    Send "{F6}"
    Sleep 100
    Send "{Home}"
    return false
}

; Shift + Y : Select first item in list (force-focus ItemsView)
+y::
{
    ; Send a right-click to shift focus into the main pane
    Click "Right"
    Sleep 100
    ; Clear any in-place edits or text focus first
    Send "{ESC}"

    EnsureItemsViewFocus()

    try {
        explorerEl := UIA.ElementFromHandle(WinExist("A"))

        itemsView := explorerEl.FindFirst({ AutomationId: "ItemsView", Type: "List" })
            ? explorerEl.FindFirst({ AutomationId: "ItemsView", Type: "List" })
            : explorerEl.FindFirst({ ClassName: "UIItemsView", Type: "List" })
                ? explorerEl.FindFirst({ ClassName: "UIItemsView", Type: "List" })
                : explorerEl.FindFirst({ Name: "Items View", Type: "List", matchmode: "Substring" })

        ; Fallback to entire window if we still did not find a dedicated list
        listRoot := itemsView ? itemsView : explorerEl

        ; Pick the very first ListItem inside that list root
        firstItem := listRoot.FindFirst({ Type: "ListItem" })

        if (firstItem) {
            firstItem.ScrollIntoView()
            firstItem.Select()
            firstItem.SetFocus()
            EnsureFocus()
            return
        }
    } catch Error {
        ; swallow and fallback below
    }

    ; Last-chance fallback â€“ press Home which works if focus is already inside the list
    Send "{Home}"
    EnsureFocus()
}

; Helper to force focus to the ItemsView pane (file list)
EnsureItemsViewFocus() {
    try {
        explorerHwnd := WinExist("A")
        root := UIA.ElementFromHandle(explorerHwnd)

        ; quick check â€“ if ItemsView already has keyboard focus, we're done
        iv := root.FindFirst({ AutomationId: "ItemsView", Type: "List" })
        if iv && iv.HasKeyboardFocus
            return

        ; Send up to 6 F6 cycles to reach the pane
        loop 6 {
            Send "{F6}"
            Sleep 120
            iv := root.FindFirst({ AutomationId: "ItemsView", Type: "List" })
            if iv && iv.HasKeyboardFocus
                break
        }
    } catch Error {
    }
}

; Shift + U : Focus search bar (Ctrl+E/F)
+u:: Send "^e"

; Shift + I : Focus address bar (Alt+D)
+i:: Send "!d"

; Shift + O : New folder
+o:: Send("^+n")

; Shift + J : Create a shortcut (Alt, Enter, Down, Enter)
; Sequence: open context/properties, navigate to "Create shortcut" option, confirm
; Uses keystrokes to avoid UIA flakiness in different Explorer views
+j::
{
    ; Ensure focus is in the file list so the keystrokes hit the right target
    EnsureItemsViewFocus()
    Sleep 100
    Send "{Alt}"
    Sleep 50
    Send "{Enter}"
    Sleep 100
    Send "{Down}"
    Sleep 50
    Send "{Enter}"
}

; Shift + K : Copy as path (Ctrl+Shift+C)
+k:: Send "^+c"

; Shift + P : Select first pinned item in Explorer sidebar
+p::
{
    SelectExplorerSidebarFirstPinned_EX()
}

; Shift + H : Select "This PC" / "Este computador" in Explorer sidebar
+h::
{
    ; First, call the same logic as +P to select the desktop (first pinned item)
    SelectExplorerSidebarFirstPinned_EX()
    Sleep 200

    ; Then press END to go down to the bottom of the tree
    Send "{End}"
    Send "{Up}"
    Send "{Up}"
    Send "{Up}"
    Send "{Up}"
}

#HotIf

;-------------------------------------------------------------------
; Microsoft Paint Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe mspaint.exe")

; Shift + Y : Resize and Skew (Ctrl+W)
+y:: Send "^w"

#HotIf

;-------------------------------------------------------------------
; Gmail Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("Gmail")

; Shift + Y: Go to main inbox (already implemented)
+y:: Send("gi")

; Shift + U: Go to updates
+u::
{
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 300 ; Give UIA time to attach

        ; Find the "Updates" tab (Name may start with "Updates" or include message counts)
        updatesButton := uia.FindElement({ Name: "Updates", Type: "TabItem", matchmode: "Substring" })

        if (updatesButton) {
            updatesButton.Click()
        }
        else {
            MsgBox "Could not find the 'Updates' button."
        }
    }
    catch Error as e {
        MsgBox "An error occurred: " e.Message
    }
}

; Shift + O: Toggle read / unread on the selected message
+o::
{
    try
    {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 300

        ; Regex patterns for the buttons (English & Portuguese)
        readPattern := "i)^(Mark as read|Marcar como lida|Marcar como lido)$"
        unreadPattern := "i)^(Mark as unread|Marcar como n[oÃ³] lida|Marcar como n[oÃ³] lido)$"

        ; Prefer clicking "Mark as read" if present; otherwise "Mark as unread"
        if (btn := WaitForButton(uia, readPattern, 1000)) {
            btn.Invoke()
        }
        else if (btn := WaitForButton(uia, unreadPattern, 1000)) {
            btn.Invoke()
        }
        else {
            MsgBox "Could not find a 'Mark as read' or 'Mark as unread' button."
        }
    }
    catch Error as e {
        MsgBox "An error occurred: " e.Message
    }
}

; Shift + I: Go to forums
+i::
{
    try
    {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 300

        ; Try English and Portuguese names for the Forums tab
        forumsButton := uia.FindElement({ Name: "Forums", Type: "TabItem", matchmode: "Substring" })
        if (!forumsButton)
            forumsButton := uia.FindElement({ Name: "FÃ³runs", Type: "TabItem", matchmode: "Substring" })

        if (forumsButton) {
            forumsButton.Click()
        }
        else {
            MsgBox "Could not find the 'Forums' button."
        }
    }
    catch Error as e {
        MsgBox "An error occurred: " e.Message
    }
}

; Shift + P: Previous conversation
+p:: Send("p")

; Shift + H: Next conversation
+h:: Send("n")

; Shift + J: Archive conversation
+j:: Send("e")

; Shift + K: Select conversation
+k:: Send("x")

; Shift + L: Reply
+l:: Send("r")

; Shift + N: Reply all
+n:: Send("a")

; Shift + M: Forward
+m:: Send("f")

; Shift + ,: Star/unstar conversation
+,:: Send("s")

; Shift + .: Delete
+.:: Send("#")

; Shift + W: Report as spam
+w:: Send("!")

; Shift + E: Compose new email
+e:: Send("c")

; Shift + R: Search mail
+r:: Send("/")

; Shift + T: Move to folder
+t:: Send("v")

; Shift + D: Show keyboard shortcuts help
+d:: Send("?")

#HotIf

;-----------------------------------------
;  Detect which editor is active
;-----------------------------------------
IsEditorActive() {
    return WinActive("ahk_exe Cursor.exe")
}

;-------------------------------------------------------------------
; Cursor Shortcuts
;-------------------------------------------------------------------
#HotIf IsEditorActive()

; Shift + Y : Fold
+y::
{
    Send "^+8"
}

; Shift + U : Unfold
+u::
{
    Send "^+9"
}

; Shift + I : Fold all
+i::
{
    Send "^+p" ; Open command palette
    Sleep 400
    Send "Fold  All"
    Sleep 200
    Send "{Down}"
    Sleep 400
    Send "{Enter}"
}

; Shift + O : Unfold all
+o::
{
    Send "^+p" ; Open command palette
    Sleep 400
    Send "Unfold All"
    Sleep 400
    Send "{Enter}"
}

; Shift + F : Close all editors
+f::
{
    Send "^+p" ; Open command palette
    Sleep 400
    Send "Close All Editors"
    Sleep 400
    Send "{Enter}"
}

; Shift + P : Go to terminal
+p:: Send "^'"

; Shift + H : New terminal
+h:: Send '^+"'

; Shift + J : Go to file explorer
+j:: Send "^+e"

; Shift + K : Format code
+k:: Send "!+f"

; Shift + L : command palette
+l:: Send "^+p"

; Shift + M : Change project
+m:: Send "^r"

; Shift + , : Show chat history
+,::
{
    Send "^+p" ; Open command palette
    Sleep 200
    Send "show history"
    Sleep 200
    Send "{Enter}"
}

; Shift + . : Extensions
+.:: Send "^+x"

; Shift + W : Switch the brackets open/close
+w:: Send "^+]"

; Shift + E : Search
+e:: Send "^+f"

; Shift + R : Save all documents
+r::
{
    Send "^k"
    Send "s"
}

; Shift + T : Trigger Ctrl+;
+t:: Send "^;"

; Shift + D : Git section
+d:: Send "^+g"

; Shift + G : Switch between AI models
+g:: SwitchAIModel()

; Shift + C : Switch between AI modes (agent/ask)
+c:: SwitchAIMode()

; Shift + V : Fold all Git directories in Source Control (Cursor)
+v:: FoldAllGitDirectoriesInCursor()

; Shift + B : Create AI commit message, then select Commit or Commit and Push
+b::
{
    Send "{Right}"
    Send "{Down}"
    Send "{Tab 2}"
    Send "{Enter}"
    Sleep 1500
    Send "{Tab 2}"
    Send "{Enter}"
    Send "{Up 2}"
}

; Ctrl + Alt + Y : Select to Bracket (via Command Palette)
^!y::
{
    oldClip := A_Clipboard
    try {
        A_Clipboard := ""
        A_Clipboard := "select to bracket"
        ClipWait 0.5
        Send "^+p"
        Sleep 300
        Send "^v"
        Sleep 150
        Send "{Enter}"
    } finally {
        A_Clipboard := oldClip
    }
}

; Ctrl + Alt + U : Remap to Shift+Alt+R
^!u:: Send "+!r"

; Ctrl + Alt + I : Fold all directories in VS Code Explorer
^!i:: FoldAllDirectoriesInExplorer()

; Ctrl + Alt + O : Unfold all directories in VS Code Explorer
^!o:: UnfoldAllDirectoriesInExplorer()

; Shift + N : Expand selection (via Command Palette)
+n:: Send "+!{Right}"

#HotIf

;-------------------------------------------------------------------
; AI Mode and Model Switching Functions for Cursor
;-------------------------------------------------------------------

; Fold all Git directories in the Source Control view by collapsing each Git tree root
FoldAllGitDirectoriesInCursor() {
    try {
        hwnd := WinExist("A")
        if !hwnd
            return
        root := UIA.ElementFromHandle(hwnd)

        ; Narrow to the Source Control (SCM) tree area to avoid unrelated matches
        scmCond := UIA.CreatePropertyConditionEx(UIA.Property.Name, "Source Control", UIA.PropertyConditionFlags.IgnoreCaseMatchSubstring
        )
        scmCondPt := UIA.CreatePropertyConditionEx(UIA.Property.Name, "Controle de CÃ³digo", UIA.PropertyConditionFlags
            .IgnoreCaseMatchSubstring
        )
        scmName := UIA.CreateOrCondition(scmCond, scmCondPt)
        scmPaneType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Pane)
        scmGroupType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Group)
        scmTreeType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Tree)
        scmScopeCond := UIA.CreateOrCondition(scmPaneType, UIA.CreateOrCondition(scmGroupType, scmTreeType))
        scmRootCond := UIA.CreateAndCondition(scmName, scmScopeCond)
        scmRoot := root.FindElement(scmRootCond, UIA.TreeScope.Descendants)
        if !scmRoot
            scmRoot := root ; fallback if SCM container not found

        ; Find TreeItem nodes whose Name contains " Git" (case-insensitive)
        nameCond := UIA.CreatePropertyConditionEx(UIA.Property.Name, " Git", UIA.PropertyConditionFlags.IgnoreCaseMatchSubstring
        )
        typeCond := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.TreeItem)
        gitItemCond := UIA.CreateAndCondition(typeCond, nameCond)

        items := scmRoot.FindElements(gitItemCond, UIA.TreeScope.Descendants)
        if !items
            return

        for item in items {
            if !item
                continue
            ; Prefer ExpandCollapsePattern when available
            hasExpand := item.GetPropertyValue(UIA.Property.IsExpandCollapsePatternAvailable)
            if hasExpand {
                try {
                    pat := item.ExpandCollapsePattern
                    state := pat.ExpandCollapseState
                    ; Collapse if not already collapsed
                    if state != UIA.ExpandCollapseState.Collapsed
                        pat.Collapse()
                } catch Error {
                    ; Fallback below if pattern fails
                }
            }
            if !hasExpand {
                ; Fallback: try to find the chevron/button and invoke/click it
                btnType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Button)
                txtType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Text)
                dotName := UIA.CreatePropertyCondition(UIA.Property.Name, ".")
                chevronCond := UIA.CreateOrCondition(btnType, UIA.CreateAndCondition(txtType, dotName))
                chevron := item.FindElement(chevronCond, UIA.TreeScope.Children)
                if !chevron
                    chevron := item.FindElement(chevronCond, UIA.TreeScope.Descendants)
                if chevron {
                    ; If it supports Invoke, prefer it; else click
                    if chevron.GetPropertyValue(UIA.Property.IsInvokePatternAvailable) {
                        try chevron.InvokePattern.Invoke()
                    } else {
                        try chevron.Click()
                    }
                }
            }
            Sleep 50
        }
    } catch Error as e {
        try MsgBox "UIA error folding Git directories: " e.Message, "Cursor Git Fold", "IconX"
    }
}

; Collapse all expandable directories in the Explorer (FileExplorer3) for all workspace roots
FoldAllDirectoriesInExplorer() {
    try {
        hwnd := WinExist("A")
        if !hwnd
            return
        root := UIA.ElementFromHandle(hwnd)

        ; Ensure Explorer is focused if not already
        Send "^+e"
        Sleep 150

        ; Find the Explorer container (EN/PT names) and then the Tree control inside it
        expEn := UIA.CreatePropertyConditionEx(UIA.Property.Name, "Explorer", UIA.PropertyConditionFlags.IgnoreCaseMatchSubstring
        )
        expPt := UIA.CreatePropertyConditionEx(UIA.Property.Name, "Explorador", UIA.PropertyConditionFlags.IgnoreCaseMatchSubstring
        )
        expName := UIA.CreateOrCondition(expEn, expPt)
        paneType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Pane)
        groupType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Group)
        scopeCond := UIA.CreateOrCondition(paneType, groupType)
        expRootCond := UIA.CreateAndCondition(expName, scopeCond)
        expRoot := ""
        try expRoot := root.FindElement(expRootCond, UIA.TreeScope.Descendants)
        if !expRoot
            expRoot := root

        treeType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Tree)
        autoId3 := UIA.CreatePropertyCondition(UIA.Property.AutomationId, "FileExplorer3")
        fileTree := ""
        try fileTree := expRoot.FindElement(autoId3, UIA.TreeScope.Descendants)
        if !fileTree {
            try {
                autoId2 := UIA.CreatePropertyCondition(UIA.Property.AutomationId, "FileExplorer2")
                fileTree := expRoot.FindElement(autoId2, UIA.TreeScope.Descendants)
            }
        }
        if !fileTree {
            try {
                autoId := UIA.CreatePropertyCondition(UIA.Property.AutomationId, "FileExplorer")
                fileTree := expRoot.FindElement(autoId, UIA.TreeScope.Descendants)
            }
        }
        if !fileTree {
            try fileTree := expRoot.FindElement(treeType, UIA.TreeScope.Descendants)
        }
        if !fileTree
            return

        ; Capture currently focused tree item (best effort) to restore selection
        hasFocusProp := UIA.CreatePropertyCondition(UIA.Property.HasKeyboardFocus, true)
        focusedEl := ""
        try focusedEl := fileTree.FindElement(hasFocusProp, UIA.TreeScope.Descendants)
        focusedName := ""
        if focusedEl
            focusedName := focusedEl.GetPropertyValue(UIA.Property.Name)

        ; Preserve scroll position when possible
        hPerc := vPerc := ""
        hasScroll := fileTree.GetPropertyValue(UIA.Property.IsScrollPatternAvailable)
        if hasScroll {
            try {
                sp := fileTree.ScrollPattern
                hPerc := sp.HorizontalScrollPercent
                vPerc := sp.VerticalScrollPercent
            }
        }

        ; Get all TreeItem nodes that support expand/collapse (i.e., directories)
        itemType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.TreeItem)
        canExpand := UIA.CreatePropertyCondition(UIA.Property.IsExpandCollapsePatternAvailable, true)
        dirCond := UIA.CreateAndCondition(itemType, canExpand)
        items := fileTree.FindElements(dirCond, UIA.TreeScope.Descendants)
        if !items
            return

        ; Collapse each expanded directory. Do not toggle; skip already collapsed.
        for item in items {
            if !item
                continue
            try {
                pat := item.ExpandCollapsePattern
                if pat.ExpandCollapseState != UIA.ExpandCollapseState.Collapsed
                    pat.Collapse()
            } catch Error {
                ; Fallback: try clicking the chevron/glyph if found (e.g., text "îª´" or button)
                btnType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Button)
                txtType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Text)
                glyphName := UIA.CreatePropertyCondition(UIA.Property.Name, "îª´")
                dotName := UIA.CreatePropertyCondition(UIA.Property.Name, ".")
                chevronCond := UIA.CreateOrCondition(btnType, UIA.CreateOrCondition(UIA.CreateAndCondition(txtType,
                    glyphName), UIA.CreateAndCondition(txtType, dotName)))
                chevron := ""
                try chevron := item.FindElement(chevronCond, UIA.TreeScope.Children)
                if !chevron
                    try chevron := item.FindElement(chevronCond, UIA.TreeScope.Descendants)
                if chevron {
                    if chevron.GetPropertyValue(UIA.Property.IsInvokePatternAvailable) {
                        try chevron.InvokePattern.Invoke()
                    } else {
                        try chevron.Click()
                    }
                }
            }
            Sleep 10
        }

        ; Restore scroll position if it changed
        if hasScroll && (hPerc != "" && vPerc != "") {
            try fileTree.ScrollPattern.SetScrollPercent(hPerc, vPerc)
        }

        ; Restore selection/focus if possible
        if focusedName {
            nameCond := UIA.CreatePropertyConditionEx(UIA.Property.Name, focusedName, UIA.PropertyConditionFlags.IgnoreCase
            )
            itemType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.TreeItem)
            focusedLookup := UIA.CreateAndCondition(itemType, nameCond)
            newFocus := ""
            try newFocus := fileTree.FindElement(focusedLookup, UIA.TreeScope.Descendants)
            if newFocus {
                try newFocus.SetFocus()
            } else {
                try fileTree.SetFocus()
            }
        } else {
            try fileTree.SetFocus()
        }

        ; Optional brief toast
        ToolTip "Directories folded"
        SetTimer () => ToolTip(), -800
    } catch Error as e {
        try MsgBox "UIA error folding Explorer directories: " e.Message, "Cursor Explorer Fold", "IconX"
    }
}

; Expand all expandable directories in the Explorer (FileExplorer3) for all workspace roots
UnfoldAllDirectoriesInExplorer() {
    try {
        hwnd := WinExist("A")
        if !hwnd
            return
        root := UIA.ElementFromHandle(hwnd)

        ; Ensure Explorer is focused if not already
        Send "^+e"
        Sleep 150

        ; Find the Explorer container
        expEn := UIA.CreatePropertyConditionEx(UIA.Property.Name, "Explorer", UIA.PropertyConditionFlags.IgnoreCaseMatchSubstring
        )
        expPt := UIA.CreatePropertyConditionEx(UIA.Property.Name, "Explorador", UIA.PropertyConditionFlags.IgnoreCaseMatchSubstring
        )
        expName := UIA.CreateOrCondition(expEn, expPt)
        paneType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Pane)
        groupType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Group)
        scopeCond := UIA.CreateOrCondition(paneType, groupType)
        expRootCond := UIA.CreateAndCondition(expName, scopeCond)
        expRoot := ""
        try expRoot := root.FindElement(expRootCond, UIA.TreeScope.Descendants)
        if !expRoot
            expRoot := root

        treeType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Tree)
        fileTree := ""
        try {
            autoId3 := UIA.CreatePropertyCondition(UIA.Property.AutomationId, "FileExplorer3")
            fileTree := expRoot.FindElement(autoId3, UIA.TreeScope.Descendants)
        }
        if !fileTree {
            try {
                autoId2 := UIA.CreatePropertyCondition(UIA.Property.AutomationId, "FileExplorer2")
                fileTree := expRoot.FindElement(autoId2, UIA.TreeScope.Descendants)
            }
        }
        if !fileTree {
            try {
                autoId := UIA.CreatePropertyCondition(UIA.Property.AutomationId, "FileExplorer")
                fileTree := expRoot.FindElement(autoId, UIA.TreeScope.Descendants)
            }
        }
        if !fileTree {
            try fileTree := expRoot.FindElement(treeType, UIA.TreeScope.Descendants)
        }
        if !fileTree
            return

        ; Preserve scroll position when possible
        hPerc := vPerc := ""
        hasScroll := fileTree.GetPropertyValue(UIA.Property.IsScrollPatternAvailable)
        if hasScroll {
            try {
                sp := fileTree.ScrollPattern
                hPerc := sp.HorizontalScrollPercent
                vPerc := sp.VerticalScrollPercent
            }
        }

        ; Get all TreeItem nodes that support expand/collapse (i.e., directories)
        itemType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.TreeItem)
        canExpand := UIA.CreatePropertyCondition(UIA.Property.IsExpandCollapsePatternAvailable, true)
        dirCond := UIA.CreateAndCondition(itemType, canExpand)
        items := ""
        try items := fileTree.FindElements(dirCond, UIA.TreeScope.Descendants)
        if !items
            return

        ; Expand each collapsed directory. Do not toggle; skip already expanded.
        for item in items {
            if !item
                continue
            try {
                pat := item.ExpandCollapsePattern
                if pat.ExpandCollapseState == UIA.ExpandCollapseState.Collapsed
                    pat.Expand()
            } catch Error {
                ; Fallback: try clicking the chevron/glyph if found (e.g., text "îª´" or button)
                btnType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Button)
                txtType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Text)
                glyphName := UIA.CreatePropertyCondition(UIA.Property.Name, "îª´")
                dotName := UIA.CreatePropertyCondition(UIA.Property.Name, ".")
                chevronCond := UIA.CreateOrCondition(btnType, UIA.CreateOrCondition(UIA.CreateAndCondition(txtType,
                    glyphName), UIA.CreateAndCondition(txtType, dotName)))
                chevron := ""
                try chevron := item.FindElement(chevronCond, UIA.TreeScope.Children)
                if !chevron {
                    try chevron := item.FindElement(chevronCond, UIA.TreeScope.Descendants)
                }
                if chevron {
                    if chevron.GetPropertyValue(UIA.Property.IsInvokePatternAvailable) {
                        try chevron.InvokePattern.Invoke()
                    } else {
                        try chevron.Click()
                    }
                }
            }
            Sleep 10
        }

        ; Restore scroll position if it changed
        if hasScroll && (hPerc != "" && vPerc != "") {
            try fileTree.ScrollPattern.SetScrollPercent(hPerc, vPerc)
        }

        ; Optional brief toast
        ToolTip "Directories unfolded"
        SetTimer () => ToolTip(), -800
    } catch Error as e {
        try MsgBox "UIA error unfolding Explorer directories: " e.Message, "Cursor Explorer Unfold", "IconX"
    }
}

; Helper: detect on-screen Text elements for "Agent"/"Ask" and send Ctrl+I/L
HasTextByRegex(pattern) {
    try {
        hwnd := WinExist("A")
        if !hwnd
            return false
        root := UIA.ElementFromHandle(hwnd)
        if !IsObject(root)
            return false
        for el in root.FindAll({ Type: "Text" }) {
            if RegExMatch(el.Name, pattern)
                return true
        }
    } catch Error as e {
        ; ignore and fall through
    }
    return false
}

SendCtrlKeyBasedOnAgentAsk() {
    ; Returns true if a key was sent, false otherwise
    if HasTextByRegex("i)\\bagent\\b") {
        Send "{Ctrl down}i{Ctrl up}"
        return true
    }
    if HasTextByRegex("i)ask") {
        Send "{Ctrl down}l{Ctrl up}"
        return true
    }
    return false
}

; Function to switch between AI modes (agent/ask)
SwitchAIMode() {
    try {
        ; Get user input directly
        userChoice := InputBox("Choose AI Mode:`n`n1. ask`n2. agent`n`nEnter choice (1 or 2):", "AI Mode Selection",
            "w250 h150")
        if userChoice.Result != "OK"
            return

        ; Determine the mode string based on choice
        modeString := ""
        switch userChoice.Value {
            case "1":
                modeString := "ask"
            case "2":
                modeString := "agent"
            default:
                MsgBox "Invalid selection. Please choose 1 or 2.", "AI Mode Selection", "IconX"
                return
        }

        ; Send Escape twice, then select the edit field based on on-screen Agent/Ask
        Send "{Escape 2}"
        Sleep 200
        if !SendCtrlKeyBasedOnAgentAsk() {
            ; Fallback to Ctrl+I if no relevant text is found
            Send "{Ctrl down}i{Ctrl up}"
        }
        Sleep 300

        ; Send Ctrl+. and wait for context menu
        Send "^."
        Sleep 500  ; Wait for context menu to appear

        ; Type the selected mode string
        SendText modeString
        Sleep 100

        ; Press Enter to confirm
        Send "{Enter}"

    } catch Error as e {
        MsgBox "Error switching AI mode: " e.Message, "AI Mode Switch Error", "IconX"
    }
}

; Function to switch between AI models
SwitchAIModel() {
    try {
        ; Get user input directly
        userChoice := InputBox(
            "Choose AI Model:`n`n1. auto`n2. CLAUD`n3. GPT`n4. O`n5. DeepSeek`n6. Cursor`n`nEnter choice (1-6):",
            "AI Model Selection", "w250 h200")
        if userChoice.Result != "OK"
            return

        ; Send Escape twice, then select the edit field based on on-screen Agent/Ask
        Send "{Escape 2}"
        Sleep 200
        if !SendCtrlKeyBasedOnAgentAsk() {
            ; Fallback to Ctrl+I if no relevant text is found
            Send "{Ctrl down}i{Ctrl up}"
        }
        Sleep 300

        ; Handle different behaviors based on choice
        switch userChoice.Value {
            case "1":
            {
                ; For auto option: simulate ;, wait for model context menu, then send â†“, Enter
                Send "^;"
                Sleep 300
                SendText "auto"
                Sleep 500
                Send "{Enter}"
                Sleep 100
                Send "{Escape}"
            }
            case "2":
            {
                ; For other options: simulate Ctrl + ., wait, type model string, no Enter
                Send "^;"
                Sleep 500
                SendText "CLAUD"
            }
            case "3":
            {
                Send "^;"
                Sleep 500
                SendText "GPT"
            }
            case "4":
            {
                Send "^;"
                Sleep 500
                SendText "O"
            }
            case "5":
            {
                Send "^;"
                Sleep 500
                SendText "DeepSeek"
            }
            case "6":
            {
                Send "^;"
                Sleep 500
                SendText "Cursor"
            }
            default:
                MsgBox "Invalid selection. Please choose 1-6.", "AI Model Selection", "IconX"
                return
        }

        Sleep 100
        Send "{Enter}"

    } catch Error as e {
        MsgBox "Error switching AI model: " e.Message, "AI Model Switch Error", "IconX"
    }
}

;-------------------------------------------------------------------
; Spotify Shortcuts
;-------------------------------------------------------------------ww
#HotIf WinActive("ahk_exe Spotify.exe")

; Shift + Y : Toggle Connect panel
+y::
{
    try {
        spot := UIA_Browser("ahk_exe Spotify.exe")
        Sleep 300

        ; Find and click the Connect button
        connectPattern := "i)^(Connect to a device|Conectar a um dispositivo|Connect)$"
        if (connectBtn := WaitForButton(spot, connectPattern)) {
            connectBtn.Invoke()

            ; After connecting, wait a moment for the device list to load
            Sleep 500

            ; Search for Office button (e.g., "Office Google Cast")
            officePattern := "i)Office"
            if (officeBtn := WaitForButton(spot, officePattern, 3000)) {
                officeBtn.Invoke()
            }
            ; If Office button not found, continue without error (as requested)
        }
        else
            MsgBox "Couldn't find the Connect-to-device button."
    } catch Error as e {
        MsgBox "Error: " e.Message
    }
}

; Shift + U : Toggle full screen
+u::
{
    try {
        spot := UIA_Browser("ahk_exe Spotify.exe")
        Sleep 300

        ; Look for either Enter or Exit full screen button with case-insensitive pattern
        enterFsPattern := "i)^Enter Full[- ]?screen$"
        exitFsPattern := "i)^Exit Full[- ]?screen$"

        ; First attempt - immediate check
        enterFsBtn := WaitForButton(spot, enterFsPattern, 500)
        if (!enterFsBtn) {
            exitFsBtn := WaitForButton(spot, exitFsPattern, 500)
            if (!exitFsBtn) {
                ; Wait 1 second and try again
                Sleep 1000
                exitFsBtn := WaitForButton(spot, exitFsPattern, 500)
                if (exitFsBtn)
                    exitFsBtn.Invoke()
            } else {
                exitFsBtn.Invoke()
            }
        } else {
            enterFsBtn.Invoke()
        }
    }
}

; Shift + I : Open Search (Ctrl+K)
+i:: Send "^k"

; Shift + O : Go to Playlists (Alt+Shift+1)
+o:: Send "!+1"

; Shift + P : Go to Artists (Alt+Shift+3)
+p:: Send "!+3"

; Shift + H : Go to Albums (Alt+Shift+4)
+h:: Send "!+4"

; Shift + J : Go to Search (Ctrl+L)
+j:: Send "^l"

; Shift + K : Go to Home (Alt+Shift+H)
+k:: Send "!+h"

; Shift + L : Go to Now Playing (Alt+Shift+J)
+l:: Send "!+j"

; Shift + N : Go to Made For You (Alt+Shift+M)
+n:: Send "!+m"

; Shift + M : Go to New Releases (Alt+Shift+N)
+m:: Send "!+n"

; Shift + , : Go to Charts (Alt+Shift+C)
+,:: Send "!+c"

; Shift + . : Toggle Now Playing View Sidebar (Alt+Shift+R)
+.:: Send "!+r"

; Shift + W : Toggle Your Library Sidebar (Alt+Shift+L)
+w:: Send "!+l"

; Shift + E : Toggle Fullscreen Library
+e::
{
    try {
        spot := UIA_Browser("ahk_exe Spotify.exe")
        Sleep 300
        fullscreenLibBtn := spot.FindElement({ Name: "fullscreen library", Type: "Button" })
        if (fullscreenLibBtn) {
            fullscreenLibBtn.Click()
        } else {
            MsgBox "Could not find the fullscreen library button.", "Spotify Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error toggling fullscreen library: " e.Message, "Spotify Error", "IconX"
    }
}

; Shift + R : Toggle lyrics
+r::
{
    try {
        spot := UIA_Browser("ahk_exe Spotify.exe")
        Sleep 300

        ; Find and click the Lyrics button
        lyricsBtn := spot.FindElement({ Name: "Lyrics", Type: "Button" })
        if (lyricsBtn) {
            lyricsBtn.Click()
        } else {
            MsgBox "Could not find the Lyrics button.", "Spotify Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error toggling lyrics: " e.Message, "Spotify Error", "IconX"
    }
}

; Shift + T : Toggle play/pause
+t::
{
    ; Simple approach: try to find Play button, if not found just send media key
    try {
        spot := UIA_Browser("ahk_exe Spotify.exe")
        Sleep 300
        playBtn := spot.FindElement({ Name: "Play", Type: "Button" })
        if (playBtn) {
            playBtn.Click()
            return
        }
    }

    ; If we get here, either Play button wasn't found or there was an error
    ; In either case, just send media key to toggle play/pause
    Send "{Media_Play_Pause}"
}

#HotIf

;-------------------------------------------------------------------
; Figma Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe Figma.exe")

; Shift + Y : Show/Hide UI (Ctrl + \)
+y:: Send("^]")

; Shift + U : Component search (Shift + I)
+u:: Send("+i")

; Shift + I : Select parent (\)
+i:: Send("]")

; Shift + O : Create component (Ctrl + Alt + K)
+o:: Send("^!k")

; Shift + P : Detach instance (Ctrl + Alt + B)
+p:: Send("^!b")

; Shift + H : Add auto layout (Shift + A)
+h:: Send("+a")

; Shift + J : Remove auto layout (Alt + Shift + A)
+j:: Send("!+a")

; Shift + K : Suggest auto layout (Ctrl + Alt + Shift + A)
+k:: Send("^!+a")

; Shift + L : Export (Ctrl + Shift + E)
+l:: Send("^+e")

; Shift + N : Copy as PNG (Ctrl + Shift + C)
+n:: Send("^+c")

; Shift + M : Actions... (Ctrl + K)
+m:: Send("^k")

; Shift + , : Align left (Alt + A)
+,:: Send("!a")

; Shift + . : Align right (Alt + D)
+.:: Send("!d")

; Shift + R : Align top (Alt + W)
+r:: Send("!w")

; Shift + T : Align bottom (Alt + S)
+t:: Send("!s")

; Shift + D : Align center horizontal (Alt + H)
+d:: Send("!h")

; Shift + F : Align center vertical (Alt + V)
+f:: Send("!v")

; Shift + G : Distribute horizontal spacing (Alt + Shift + H)
+g:: Send("!+h")

; Shift + W : Distribute vertical spacing (Alt + Shift + V)
+w:: Send("!+v")

; Shift + E : Tidy up (Ctrl + Alt + Shift + T)
+e:: Send("^!+t")

#HotIf

;-------------------------------------------------------------------
; Mobills Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("Mobills")

; Shift + Y : Dashboard
+y:: {
    try {
        btn := GetMobillsButton("menu-dashboard-item", "Dashboard")
        if (btn) {
            btn.Click()
        } else {
            MsgBox "Could not find the Dashboard button.", "Mobills Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error navigating to Dashboard: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + U : Contas
+u:: {
    try {
        btn := GetMobillsButton("menu-accounts-item", "Accounts")
        if (btn) {
            btn.Click()
        } else {
            MsgBox "Could not find the Contas/Accounts button.", "Mobills Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error navigating to Contas/Accounts: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + I : TransaÃ§Ãµes
+i:: {
    try {
        btn := GetMobillsButton("menu-transactions-item", "Transactions")
        if (btn) {
            btn.Click()
        } else {
            MsgBox "Could not find the TransaÃ§Ãµes/Transactions button.", "Mobills Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error navigating to TransaÃ§Ãµes/Transactions: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + O : CartÃµes de crÃ©dito
+o:: {
    try {
        btn := GetMobillsButton("menu-creditCards-item", "Credit cards")
        if (btn) {
            btn.Click()
        } else {
            MsgBox "Could not find the CartÃµes de crÃ©dito/Credit cards button.", "Mobills Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error navigating to CartÃµes de crÃ©dito/Credit cards: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + P : Planejamento
+p:: {
    try {
        btn := GetMobillsButton("menu-budgets-item", "Budgets")
        if (btn) {
            btn.Click()
        } else {
            MsgBox "Could not find the Planejamento/Budgets button.", "Mobills Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error navigating to Planejamento/Budgets: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + H : RelatÃ³rios
+h:: {
    try {
        btn := GetMobillsButton("menu-reports-item", "Reports")
        if (btn) {
            btn.Click()
        } else {
            MsgBox "Could not find the RelatÃ³rios/Reports button.", "Mobills Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error navigating to RelatÃ³rios/Reports: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + J : Mais opÃ§Ãµes
+j:: {
    try {
        btn := GetMobillsButton("menu-moreOptions-item", "More options")
        if (btn) {
            btn.Click()
        } else {
            MsgBox "Could not find the Mais opÃ§Ãµes/More options button.", "Mobills Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error navigating to Mais opÃ§Ãµes/More options: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + K : Previous month
+k:: PrevMobillsMonth()

; Shift + L : Next month
+l:: NextMobillsMonth()

PrevMobillsMonth() {
    try {
        uia := TryAttachBrowser()
        if !uia {
            MsgBox "Could not attach to the browser window.", "Mobills Navigation", "IconX"
            return
        }
        grp := FindMonthGroup(uia)
        if !grp {
            MsgBox "Could not locate the month container.", "Mobills Navigation", "IconX"
            return
        }
        ; First try: immediate previous sibling
        btn := grp.WalkTree("-1", { Type: "Button" })
        if !btn {
            ; Fallback: search all buttons inside parent and pick the one to the LEFT of the group
            parent := UIA.TreeWalkerTrue.GetParentElement(grp)
            if (parent) {
                grpPos := grp.Location
                for , el in parent.FindAll({ Type: "Button" }) {
                    pos := el.Location
                    if (pos.y >= grpPos.y - 10 && pos.y <= grpPos.y + grpPos.h + 10 && pos.x < grpPos.x) {
                        btn := el                          ; closest left candidate
                    }
                }
            }
        }
        if btn {
            btn.Click()
        } else {
            MsgBox "Could not find the previous-month button.", "Mobills Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error navigating to previous month:`n" e.Message, "Mobills Error", "IconX"
    }
}

NextMobillsMonth() {
    try {
        uia := TryAttachBrowser()
        if !uia {
            MsgBox "Could not attach to the browser window.", "Mobills Navigation", "IconX"
            return
        }
        grp := FindMonthGroup(uia)
        if !grp {
            MsgBox "Could not locate the month container.", "Mobills Navigation", "IconX"
            return
        }
        ; First try: immediate next sibling
        btn := grp.WalkTree("+1", { Type: "Button" })
        if !btn {
            ; Fallback: search all buttons inside parent and pick the one to the RIGHT of the group
            parent := UIA.TreeWalkerTrue.GetParentElement(grp)
            if (parent) {
                grpPos := grp.Location
                for , el in parent.FindAll({ Type: "Button" }) {
                    pos := el.Location
                    if (pos.y >= grpPos.y - 10 && pos.y <= grpPos.y + grpPos.h + 10 && pos.x > grpPos.x + grpPos.w) {
                        btn := el                          ; closest right candidate
                    }
                }
            }
        }
        if btn {
            btn.Click()
        } else {
            MsgBox "Could not find the next-month button.", "Mobills Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error navigating to next month:`n" e.Message, "Mobills Error", "IconX"
    }
}

; ---- New helper to jump from "Open" button ----
FocusViaOpenButton(tabs, pressSpace := false) {
    try {
        uia := TryAttachBrowser()
        if !uia
            return false
        ; Anchor = Button named "Open"
        openBtn := uia.FindElement({ Name: "Open", Type: "Button" })
        if !openBtn {
            ; fallback by class substring
            openBtn := uia.FindElement({ ClassName: "MuiAutocomplete-popupIndicator", Type: "Button", matchmode: "Substring" })
        }
        if !openBtn
            return false
        openBtn.SetFocus()
        Sleep 200
        ; Tab forward specified times
        loop tabs {
            Send "+{Tab}"
            Sleep 80
        }
        if pressSpace {
            Sleep 80
            Send "{Space}"
        }
        return true
    } catch Error {
        return false
    }
}

; Shift + N : Toggle "Ignore transaction"
+n:: FocusViaOpenButton(3, true)

; Shift + M : Focus expense name field
+m:: FocusViaOpenButton(2, false)

#HotIf

;-------------------------------------------------------------------
; Google Keep Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe chrome.exe") && (WinActive("Google Keep") || WinActive("keep.google.com") || InStr(
    WinGetTitle("A"), "Google Keep"))

; Shift + Y : Search and select note
+y::
{
    ; Store the current active window handle
    currentWindow := WinExist("A")

    ; Show message box to get search text from user
    searchText := InputBox("Enter text to search for in your notes:", "Google Keep Search", "w300 h100")

    if (searchText.Result = "OK" && searchText.Value != "") {
        ; Explicitly activate the Google Keep window to ensure we're working with the right window
        WinActivate("ahk_id " currentWindow)
        WinWaitActive("ahk_id " currentWindow, , 2)

        ; Store the search text in clipboard
        oldClip := A_Clipboard
        A_Clipboard := searchText.Value

        ; Wait a moment for clipboard to be ready
        Sleep 200

        ; Send Escape to clear any current selection/focus
        Send "{Esc}"
        Sleep 300

        ; Open search with Ctrl+F
        Send "^f"
        Sleep 200

        ; Paste the search text
        Send "^v"
        Sleep 900

        ; Press Escape to close search
        Send "{Esc}"
        Sleep 300

        ; Press Enter to confirm selection
        Send "{Enter}"
        Sleep 300

        ; Restore original clipboard
        A_Clipboard := oldClip
    }
}

; Shift + U : Toggle main menu
+u::
{
    try {
        ; Store the current active window handle
        currentWindow := WinExist("A")

        ; Explicitly activate the Google Keep window to ensure we're working with the right window
        WinActivate("ahk_id " currentWindow)
        WinWaitActive("ahk_id " currentWindow, , 2)

        ; Use UIA to find and click the main menu button
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 300 ; Give UIA time to attach

        ; Find the main menu button by its properties
        mainMenuBtn := uia.FindElement({
            Name: "Main menu",
            Type: "Button",
            ClassName: "gb_Lc"
        })

        if (mainMenuBtn) {
            mainMenuBtn.Click()
        } else {
            ; Fallback: try to find by name only
            mainMenuBtn := uia.FindElement({ Name: "Main menu", Type: "Button" })
            if (mainMenuBtn) {
                mainMenuBtn.Click()
            } else {
                MsgBox "Could not find the Main menu button.", "Google Keep", "IconX"
            }
        }
    } catch Error as e {
        MsgBox "Error toggling main menu: " e.Message, "Google Keep Error", "IconX"
    }
}

#HotIf

ConfirmDismissAll() {
    if MsgBox("Dismiss all reminders?", "Confirm Dismiss", "YesNo Icon?") = "Yes"
        DismissAllReminders()
}

DismissAllReminders() {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)
        ; Try by AutomationId first
        btn := root.FindFirst({ AutomationId: "8345", ControlType: "Button" })
        ; Fallback: search by name
        if !btn
            btn := root.FindFirst({ Name: "Dismiss All", ControlType: "Button" })
        if btn {
            btn.Click()
        } else {
            MsgBox("Could not find the 'Dismiss All' button.", "Dismiss All", "IconX")
        }
    } catch Error as e {
        MsgBox("UIA error:`n" e.Message, "Dismiss All Error", "IconX")
    }
}

; ---------------------------------------------------------------------------
; Helper for Mobills buttons â€“ language-neutral search
; ---------------------------------------------------------------------------
GetMobillsButton(autoId, btnName) {
    try {
        root := UIA.ElementFromHandle(WinExist("A"))
        btn := root.FindFirst({ AutomationId: autoId, ControlType: "Button" })
        if !btn
            btn := root.FindFirst({ Name: btnName, ControlType: "Button" })
        return btn
    } catch Error {
        return ""
    }
}

;-------------------------------------------
; Helper functions
;-------------------------------------------
TryAttachBrowser() {
    ; Try Chrome first, then Edge
    try return UIA_Browser("ahk_exe chrome.exe")
    catch {
        try return UIA_Browser("ahk_exe msedge.exe")
        catch {
            return 0
        }
    }
}

FindMonthGroup(uia) {
    ; Strategy 1 â€“ look for known class name on the container
    try {
        grp := uia.FindElement({ Type: "Group", ClassName: "sc-kAyceB", matchmode: "Substring" })
        if grp
            return grp
    }
    catch {
    }
    ; Strategy 2 â€“ locate by month text (any language)
    months := ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
        "November", "December",
        "Janeiro", "Fevereiro", "MarÃ§o", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro",
        "Novembro",
        "Dezembro"]
    for , m in months {
        try {
            el := uia.FindElement({ Name: m, Type: "Text", mm: 1, cs: false })
            if el {
                grp := el.WalkTree("p", { Type: "Group" })
                if grp
                    return grp
            }
        }
        catch {
        }
    }
    return 0
}

;-------------------------------------------------------------------
; YouTube Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe chrome.exe") && InStr(WinGetTitle("A"), "YouTube")

; Shift + Y : Focus search box
+y:: {
    try {
        uia := UIA_Browser()
        Sleep 300

        ; Try multiple search strategies
        searchBox := uia.FindFirst({ Type: "ComboBox", Name: "Search" })
        if !searchBox
            searchBox := uia.FindFirst({ Type: "Edit", Name: "Search" })
        if !searchBox
            searchBox := uia.FindFirst({ ClassName: "ytSearchboxComponentInput" })
        if !searchBox
            searchBox := uia.FindFirst({ Type: "SearchBox" })
        if !searchBox
            searchBox := uia.FindFirst({ AutomationId: "search" })

        if (searchBox) {
            searchBox.SetFocus()
            ; Additional fallback - if SetFocus doesn't work, try sending keyboard shortcut
            Sleep 100
            if !searchBox.HasKeyboardFocus {
                Send "/"  ; YouTube's built-in shortcut to focus search
            }
        } else {
            ; Last resort - just use YouTube's built-in keyboard shortcut
            Send "/"
        }
    } catch Error as e {
        ; If all else fails, use the keyboard shortcut
        Send "/"
    }
}

; Shift + U : Focus first video via Search filters button
+u:: {
    try {
        uia := UIA_Browser()
        Sleep 300

        ; Find the "Search filters" button as anchor
        searchFiltersButton := uia.FindFirst({ Name: "Search filters" })
        if !searchFiltersButton
            searchFiltersButton := uia.FindFirst({ Type: "Button", Name: "Search filters" })
        if !searchFiltersButton
            searchFiltersButton := uia.FindFirst({ AutomationId: "search-filters" })

        if (searchFiltersButton) {
            ; Focus the Search filters button (do not click)
            searchFiltersButton.SetFocus()
            Sleep 200

            ; Send Tab to move focus to the first video list item
            Send "{Tab}"
            Sleep 100

            ; Press Enter to select/play the first video
            Send "{Enter}"
        } else {
            ; Fallback: try to navigate to first video using keyboard shortcuts
            Send "{Home}"  ; Go to top of page
            Sleep 100
            Send "{Tab}"   ; Tab to first focusable element
            Sleep 100
            Send "{Enter}" ; Press Enter
        }
    } catch Error as e {
        ; If all else fails, use basic keyboard navigation
        Send "{Home}"
        Sleep 100
        Send "{Tab}"
        Sleep 100
        Send "{Enter}"
    }
}

#HotIf

;-------------------------------------------------------------------
; Google Search Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe chrome.exe") && InStr(WinGetTitle("A"), "Google")

; Shift + Y : Focus Google search box
+y:: {
    try {
        uia := UIA_Browser()
        Sleep 300

        ; Find the Google search box by Name
        searchBox := uia.FindFirst({ Name: "Search" })
        if !searchBox
            searchBox := uia.FindFirst({ Type: "Edit", Name: "Search" })
        if !searchBox
            searchBox := uia.FindFirst({ Type: "SearchBox", Name: "Search" })
        if !searchBox
            searchBox := uia.FindFirst({ AutomationId: "search" })

        if (searchBox) {
            searchBox.SetFocus()
            Sleep 100
            if !searchBox.HasKeyboardFocus {
                ; Fallback: use Ctrl+L to focus address bar, then Tab to search
                Send "^l"
                Sleep 100
                Send "{Tab}"
            }
        } else {
            ; Last resort: use Ctrl+L to focus address bar, then Tab to search
            Send "^l"
            Sleep 100
            Send "{Tab}"
        }
    } catch Error as e {
        ; If all else fails, use keyboard navigation
        Send "^l"
        Sleep 100
        Send "{Tab}"
    }
}

#HotIf

;-------------------------------------------------------------------
; File Dialog (Namespace Tree Control) Shortcuts
;-------------------------------------------------------------------
#HotIf IsFileDialogActive()

; Shift + Y : select first file (via header focus + Shift+Tab)
+y:: {

    try {
        root := UIA.ElementFromHandle(WinExist("A"))

        ; First attempt: find and focus Items View list directly
        itemsList := root.FindFirst({ Type: "List", ClassName: "UIItemsView" })
        if !itemsList
            itemsList := root.FindFirst({ Type: "List", Name: "Items View" })
        if !itemsList
            itemsList := root.FindFirst({ Type: "List", AutomationId: "ItemsView" })

        if itemsList {
            itemsList.SetFocus()
            Sleep 120
            Send "{Home}"  ; Go to first item
            EnsureFocus()
            return
        }

        ; Second attempt: find header (Header control)
        hdr := root.FindFirst({ Type: "Header" })
        if !hdr {
            hdr := root.FindFirst({ Name: "Header", Type: "Header" })
            if !hdr
                hdr := root.FindFirst({ Name: "CabeÃ§alho", Type: "Header" })
        }
        if hdr {
            hdr.SetFocus()
            Sleep 120
            Send "+{Tab}"
            Send "{Home}"
            EnsureFocus()
            return
        }

        ; Third attempt: find file name ComboBox by AutomationId and Type
        ; This should work regardless of the name (File name: or Nome:)
        fileNameCombo := root.FindFirst({ Type: "ComboBox", AutomationId: "1148" })

        ; If not found, try by various possible names
        if !fileNameCombo {
            possibleNames := ["File name:", "Nome:", "Filename:", "File Name:"]
            for name in possibleNames {
                fileNameCombo := root.FindFirst({ Type: "ComboBox", Name: name })
                if fileNameCombo
                    break
            }
        }

        ; If ComboBox found, use it
        if fileNameCombo {
            fileNameCombo.SetFocus()

            MsgBox "ComboBox found"
            Sleep 120
            Send "+{Tab}"
            Send "{Home}"
            EnsureFocus()
            return
        }
    } catch Error {
    }
    ; Last resort fallback: simple Shift+Tab then Home
    Send "+{Tab}"
    Sleep 120
    Send "{Home}"
    EnsureFocus()

}

; Shift + U : Focus search bar (Ctrl+E/F)
+u:: Send "^e"

; Shift + I : Create new folder (Ctrl+Shift+N), then Shift+Tab twice and Enter
+i:: {
    Send "^e"
    Sleep 120
    Send "+{Tab}"
    Sleep 120
    Send "+{Tab}"
    Sleep 120
    Send "{Enter}"
}

; Shift + O : New folder (Ctrl+Shift+N)
+o:: Send "^+n"

; Shift + P : Select first pinned item in sidebar (reuse Explorer helper)
+p::
{
    SelectExplorerSidebarFirstPinned()
}

; Shift + H : Select "This PC" / "Este computador" in sidebar
+h::
{
    SelectExplorerSidebarFirstPinned()
    Sleep 200
    Send "{End}"
    Send "{Up}"
    Send "{Up}"
    Send "{Up}"
    Send "{Up}"
}

; Shift + J : Focus file name edit field (moved from +U)
+j:: {
    try {
        root := UIA.ElementFromHandle(WinExist("A"))

        ; First attempt: Find by AutomationId and Type (most reliable)
        fileNameEdit := root.FindFirst({ Type: "Edit", AutomationId: "1148" })

        ; Second attempt: Try various possible names
        if !fileNameEdit {
            possibleNames := [
                "File name:",      ; English standard
                "Nome:",          ; Portuguese standard
                "Filename:",      ; Alternative English
                "File Name:",     ; Alternative capitalization
                "Name:",          ; Generic English
                "Nome do arquivo:", ; Full Portuguese
                "Save As:",       ; Save dialog English
                "Salvar como:"    ; Save dialog Portuguese
            ]
            for name in possibleNames {
                fileNameEdit := root.FindFirst({ Type: "Edit", Name: name })
                if fileNameEdit
                    break
            }
        }

        ; Third attempt: Try to find through parent ComboBox
        if !fileNameEdit {
            fileNameCombo := root.FindFirst({ Type: "ComboBox", AutomationId: "1148" })
            if fileNameCombo {
                fileNameEdit := fileNameCombo.FindFirst({ Type: "Edit" })
            }
        }

        if fileNameEdit {
            fileNameEdit.SetFocus()
            Sleep 50
            Send "^a"  ; Select all existing text
            return
        }
    } catch Error {
    }
    ; Fallback: Try to focus using keyboard navigation
    Send "!n"  ; Alt+N is a common shortcut for file name field
}

; Shift + K : Click Insert/Open/Save button (moved from +I)
+k:: {
    try {
        root := UIA.ElementFromHandle(WinExist("A"))

        ; First attempt: Find by AutomationId and Type (most reliable)
        actionBtn := root.FindFirst({ Type: "Button", AutomationId: "1" })

        ; Second attempt: Try various possible names
        if !actionBtn {
            possibleNames := [
                ; English variations
                "Insert",
                "Open",
                "Save",
                "Save As",
                "OK",
                ; Portuguese variations
                "Abrir",
                "Salvar",
                "Salvar como",
                "Inserir",
                ; Spanish variations (common in some systems)
                "Insertar",
                "Guardar",
                "Guardar como",
                ; French variations (common in some systems)
                "InsÃ©rer",
                "Ouvrir",
                "Enregistrer",
                "Enregistrer sous"
            ]
            for name in possibleNames {
                actionBtn := root.FindFirst({ Type: "Button", Name: name })
                if actionBtn
                    break
            }
        }

        ; Third attempt: Try SplitButton type (some dialogs use this instead)
        if !actionBtn {
            actionBtn := root.FindFirst({ Type: "SplitButton", AutomationId: "1" })
            if !actionBtn {
                for name in possibleNames {
                    actionBtn := root.FindFirst({ Type: "SplitButton", Name: name })
                    if actionBtn
                        break
                }
            }
        }

        if actionBtn {
            actionBtn.Click()
            return
        }
    } catch Error {
    }
    ; Fallback: Try common keyboard shortcuts
    Send "!s"  ; Alt+S (Save)
    Sleep 50
    Send "!o"  ; Alt+O (Open)
}

; Shift + L : Click Cancel button (moved from +O)
+l:: {
    try {
        root := UIA.ElementFromHandle(WinExist("A"))

        ; First attempt: Find by AutomationId and Type (most reliable)
        cancelBtn := root.FindFirst({ Type: "Button", AutomationId: "2" })

        ; Second attempt: Try various possible names
        if !cancelBtn {
            possibleNames := [
                ; English variations
                "Cancel",
                "Close",
                "Exit",
                "Dismiss",
                ; Portuguese variations
                "Cancelar",
                "Fechar",
                "Sair",
                ; Spanish variations
                "Cancelar",
                "Cerrar",
                ; French variations
                "Annuler",
                "Fermer",
                ; German variations
                "Abbrechen",
                "SchlieÃŸen",
                ; Italian variations
                "Annulla",
                "Chiudi",
                ; Generic
                "No",
                "NÃ£o",
                "Ã—",  ; Sometimes used as close symbol
                "âœ•"   ; Alternative close symbol
            ]
            for name in possibleNames {
                cancelBtn := root.FindFirst({ Type: "Button", Name: name })
                if cancelBtn
                    break
            }
        }

        if cancelBtn {
            cancelBtn.Click()
            return
        }
    } catch Error {
    }
    ; Fallback: Try common keyboard shortcuts
    Send "{Esc}"  ; Escape key is universal for cancel
}

#HotIf

IsFileDialogActive() {
    hwnd := WinActive("A")
    if !hwnd {
        return false
    }

    winClass := WinGetClass("ahk_id " hwnd)
    if winClass != "#32770" {
        return false
    }

    ; Exclude Outlook Reminders which can share dialog-like traits
    try {
        if (WinGetProcessName("ahk_id " hwnd) = "OUTLOOK.EXE") {
            t := WinGetTitle("ahk_id " hwnd)
            if RegExMatch(t, "i)Reminder")
                return false
        }
    } catch Error {
    }

    try {
        root := UIA.ElementFromHandle(hwnd)
        ; Try to find ANY useful identifiers
        for type in ["List", "Tree", "Pane", "Window"] {
            elements := root.FindAll({ Type: type })
            if elements.Length {
                return true
            }
        }
        return true
    } catch Error as e {
        return false
    }
}

;-------------------------------------------------------------------
; UIA Tree Inspector Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("UIATreeInspector") || WinActive("ahk_exe UIATreeInspectorAutoHotkey64.exe")

; Shift + Y : Refresh list
+y:: {
    try {
        root := UIA.ElementFromHandle(WinExist("A"))
        Sleep 200
        btn := root.FindFirst({ Name: "Refresh list", Type: "Button" })
        if !btn
            btn := root.FindFirst({ AutomationId: "5", Type: "Button" })
        if btn {
            btn.Invoke()
        } else {
            MsgBox "Could not find the Refresh list button.", "UIA Tree Inspector", "IconX"
        }
    } catch Error as e {
        MsgBox "Error refreshing list:`n" e.Message, "UIA Tree Inspector", "IconX"
    }
}

; Shift + U : Focus macro sidebar button and shift tab 6 times
+u:: {
    try {
        root := UIA.ElementFromHandle(WinExist("A"))
        Sleep 200
        ; Find the "Filter:" text element
        filterText := root.FindFirst({ Name: "Filter:", Type: "Text", AutomationId: "18" })
        if filterText {
            ; Focus on the text element
            filterText.SetFocus()
            Sleep 100

            ; Hit Tab once
            Send "{Tab}"
            Sleep 50
        } else {
            MsgBox "Could not find the 'Filter:' text element.", "Text Focus", "IconX"
        }
    } catch Error as e {
        MsgBox "Error focusing button and performing Shift+Tab sequence:`n" e.Message, "Button Focus", "IconX"
    }
}
#HotIf

;-------------------------------------------------------------------
; SettleUp Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("Settle Up")

; Shift + Y : Click "Adicionar transaÃ§Ã£o" button (UIA by Name substring)
+y:: {
    try {
        uia := UIA_Browser()
        Sleep 200
        ; Keep it simple: search only by Name with substring
        btn := uia.FindElement({
            Name: "Adicionar transa",
            matchmode: "Substring"
        })
        if (btn) {
            btn.Click()
        } else {
            MsgBox "Could not find the 'Adicionar transaÃ§Ã£o' button."
        }
    } catch Error as e {
        MsgBox "Error clicking 'Adicionar transaÃ§Ã£o': " e.Message
    }
}

; Shift + U : Focus expense value field (via name field + tab)
+u:: {
    try {
        uia := UIA_Browser()
        Sleep 300

        ; Find the "who paid" combo box
        paidByCombo := uia.FindFirst({
            Type: "ComboBox",
            Name: "Eduardo Figueiredo pagou"
        })

        ; If not found by exact match, try partial matches
        if !paidByCombo {
            possibleNames := [
                " pagou",           ; Portuguese suffix
                " paid",            ; English suffix
                " pagÃ³",            ; Spanish suffix
                " a payÃ©"           ; French suffix
            ]
            for suffix in possibleNames {
                paidByCombo := uia.FindFirst({
                    Type: "ComboBox",
                    Name: A_UserName . suffix,
                    matchmode: "Substring"
                })
                if paidByCombo
                    break
            }
        }

        ; Try by AutomationId if name matching failed
        if !paidByCombo {
            paidByCombo := uia.FindFirst({
                Type: "ComboBox",
                AutomationId: "mat-select-54"
            })
        }

        if paidByCombo {
            paidByCombo.Click()
            Sleep 100
            Send "{Tab}"  ; Move to expense value field
            return
        }
    } catch Error as e {
        ; Silently handle errors
    }
}

; Shift + I : Focus expense name field (via value field + 6 tabs)
+i:: {
    try {
        uia := UIA_Browser()
        Sleep 300

        ; Find the "who paid" combo box (same logic as +u)
        paidByCombo := uia.FindFirst({
            Type: "ComboBox",
            Name: "Eduardo Figueiredo pagou"
        })

        ; If not found by exact match, try partial matches
        if !paidByCombo {
            possibleNames := [
                " pagou",           ; Portuguese suffix
                " paid",            ; English suffix
                " pagÃ³",            ; Spanish suffix
                " a payÃ©"           ; French suffix
            ]
            for suffix in possibleNames {
                paidByCombo := uia.FindFirst({
                    Type: "ComboBox",
                    Name: A_UserName . suffix,
                    matchmode: "Substring"
                })
                if paidByCombo
                    break
            }
        }

        ; Try by AutomationId if name matching failed
        if !paidByCombo {
            paidByCombo := uia.FindFirst({
                Type: "ComboBox",
                AutomationId: "mat-select-54"
            })
        }

        if paidByCombo {
            paidByCombo.Click()
            Sleep 100
            Send "{Tab}"  ; Move to expense value field
            Sleep 200     ; Slow tab timing

            ; Now tab 6 times slowly to reach expense name field
            loop 6 {
                Send "{Tab}"
                Sleep 20  ; Slow timing between tabs
            }
            return
        }
    } catch Error as e {
        ; Silently handle errors
    }
}

#HotIf

;-------------------------------------------------------------------
; Miro Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe chrome.exe") && InStr(WinGetTitle("A"), "Miro")

; (removed) Shift + Y : Command palette (Ctrl+K)

; Shift + U : Frame list (Ctrl+Shift+F)
+u:: Send "^+f"

; Shift + I : Group (Ctrl+G)
+i:: Send "^g"

; Shift + O : Ungroup (Ctrl+Shift+G)
+o:: Send "^+g"

; Shift + P : Lock/Unlock (Ctrl+Shift+L)
+p:: Send "^+l"

; Shift + H : Add/Edit link (Alt+Ctrl+K)
+h:: Send "!^k"

#HotIf

; --- Unified banner helpers for ChatGPT indicators (match ChatGPT.ahk style) ---
global smallLoadingGuis_ChatGPT := []

CreateCenteredBanner_ChatGPT(message, bgColor := "3772FF", fontColor := "FFFFFF", fontSize := 24, alpha := 178) {
    bGui := Gui()
    bGui.Opt("+AlwaysOnTop -Caption +ToolWindow")
    bGui.BackColor := bgColor
    bGui.SetFont("s" . fontSize . " c" . fontColor . " Bold", "Segoe UI")
    bGui.Add("Text", "w500 Center", message)

    activeWin := WinGetID("A")
    if (activeWin) {
        WinGetPos(&winX, &winY, &winW, &winH, activeWin)
    } else {
        workArea := SysGet.MonitorWorkArea(SysGet.MonitorPrimary)
        winX := workArea.Left, winY := workArea.Top, winW := workArea.Right - workArea.Left, winH := workArea.Bottom -
            workArea.Top
    }

    bGui.Show("AutoSize Hide")
    guiW := 0, guiH := 0
    bGui.GetPos(, , &guiW, &guiH)

    guiX := winX + (winW - guiW) / 2
    guiY := winY + (winH - guiH) / 2
    bGui.Show("x" . Round(guiX) . " y" . Round(guiY) . " NA")
    WinSetTransparent(alpha, bGui)
    return bGui
}

ShowSmallLoadingIndicator_ChatGPT(state := "Loadingâ€¦", bgColor := "3772FF") {
    global smallLoadingGuis_ChatGPT
    if (smallLoadingGuis_ChatGPT.Length > 0) {
        try if (smallLoadingGuis_ChatGPT[1].Controls.Length > 0)
            smallLoadingGuis_ChatGPT[1].Controls[1].Text := state
        catch {
        }
        return
    }
    textGui := CreateCenteredBanner_ChatGPT(state, bgColor, "FFFFFF", 24, 178)
    smallLoadingGuis_ChatGPT.Push(textGui)
}

HideSmallLoadingIndicator_ChatGPT() {
    global smallLoadingGuis_ChatGPT
    if (smallLoadingGuis_ChatGPT.Length > 0) {
        for gui in smallLoadingGuis_ChatGPT {
            try gui.Destroy()
        }
        smallLoadingGuis_ChatGPT := []
    }
}

; Short completion chime for ChatGPT responses (debounced)
PlayCompletionChime_ChatGPT() {
    try {
        static lastTick := 0
        if (A_TickCount - lastTick < 1500)
            return
        lastTick := A_TickCount

        played := false
        ; Prefer Windows MessageBeep (reliable through default output)
        try {
            rc := DllCall("User32\\MessageBeep", "UInt", 0xFFFFFFFF)
            if (rc)
                played := true
        } catch {
        }

        ; Fallback to system asterisk sound
        if !played {
            try {
                played := SoundPlay("*64", false)
            } catch {
            }
        }

        ; Last resort, attempt the classic beep
        if !played {
            try SoundBeep(1100, 130)
            catch {
            }
        }
    } catch {
    }
}

WaitForButtonAndShowSmallLoading_ChatGPT(buttonNames, stateText := "Loadingâ€¦", timeout := 15000) {
    ; Store ChatGPT's window handle before Alt+Tab (robust contains-match)
    chatGPTHwnd := GetChatGPTWindowHwnd()
    if !chatGPTHwnd {
        return ; ChatGPT window not found
    }

    ; Obtain UIA context for ChatGPT window specifically
    try cUIA := UIA_Browser("ahk_id " chatGPTHwnd)
    catch {
        return ; Failed to get UIA context
    }

    start := A_TickCount
    btn := ""

    ; Wait for the target button to appear and monitor it until it disappears
    deadline := (timeout > 0) ? (start + timeout) : 0
    while (timeout <= 0 || (A_TickCount < deadline)) {
        btn := ""
        for n in buttonNames {
            try btn := cUIA.FindElement({ Name: n, Type: "Button" })
            catch {
                btn := ""
            }
            if !btn {
                ; Fallback: substring match without strict type (handles UI variations)
                try btn := cUIA.FindElement({ Name: n, matchmode: "Substring" })
                catch {
                    btn := ""
                }
            }
            if btn
                break
        }
        if btn {
            ShowSmallLoadingIndicator_ChatGPT(stateText)
            while btn && (timeout <= 0 || (A_TickCount < deadline)) {
                Sleep 250
                btn := ""
                for n in buttonNames {
                    try btn := cUIA.FindElement({ Name: n, Type: "Button" })
                    catch {
                        btn := ""
                    }
                    if !btn {
                        ; Fallback: substring match without strict type
                        try btn := cUIA.FindElement({ Name: n, matchmode: "Substring" })
                        catch {
                            btn := ""
                        }
                    }
                    if btn
                        break
                }
            }
            break
        }
        Sleep 250
    }

    ; Chime only for real AI answering events (not transcription)
    try {
        if (InStr(StrLower(stateText), "transcrib") = 0)
            PlayCompletionChime_ChatGPT()
    } catch {
    }
    ; Always hide the indicator at the end (debounced safety)
    try HideSmallLoadingIndicator_ChatGPT()
    catch {
    }
}
