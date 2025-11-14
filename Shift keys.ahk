/********************************************************************
 *   Win+Alt+Shift symbol layer shortcuts (AHK v2)
 *   â€¢ Provides system-wide symbol shortcuts
 ********************************************************************/

/********************************************************************
 *   AVAILABLE WIN+ALT+SHIFT COMBINATIONS
 *   The following combinations are not currently in use:
 *   
 *   Letters: P, U
 *   Numbers: (all numbers 0-9 are used)
 *   Symbols: ; ' [ ] \ | ` ~ @ # $ % ^ & * ( ) - _ = + { } : " < > ? /
 *   
 *   Note: Some combinations use Ctrl+Alt+Shift+Arrow keys for extended mouse movement
 ********************************************************************/

#Requires AutoHotkey v2.0+

#SingleInstance Force

SetTitleMatchMode 2

#include %A_ScriptDir%\env.ahk
#include UIA-v2\Lib\UIA.ahk
#include UIA-v2\Lib\UIA_Browser.ahk
#include %A_ScriptDir%\ChatGPT_Loading.ahk
#include %A_ScriptDir%\Hotstrings.ahk

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

; Function to pad shortcuts to consistent width for alignment
PadShortcut(shortcut, targetWidth := 24) {
    ; Extract the content between brackets
    if RegExMatch(shortcut, "\[(.*?)\]", &match) {
        content := match[1]
        ; Calculate padding needed
        padding := targetWidth - StrLen(content)
        if (padding > 0) {
            ; Calculate left and right padding for centering
            leftPadding := Floor(padding / 2)
            rightPadding := padding - leftPadding

            ; Create left padding string
            leftPaddingStr := ""
            loop leftPadding {
                leftPaddingStr .= " "
            }

            ; Create right padding string
            rightPaddingStr := ""
            loop rightPadding {
                rightPaddingStr .= " "
            }

            ; Center the content within the brackets
            return "[" . leftPaddingStr . content . rightPaddingStr . "]"
        }
    }
    return shortcut
}

; Function to process cheat sheet text and pad all shortcuts
ProcessCheatSheetText(text) {
    ; Split into lines
    lines := StrSplit(text, "`n")
    processedLines := []

    for line in lines {
        ; Check if line contains a shortcut pattern
        if RegExMatch(line, "(\[.*?\])", &match) {
            ; Replace the shortcut with padded version
            paddedShortcut := PadShortcut(match[1])
            processedLine := StrReplace(line, match[1], paddedShortcut)

            ; Check if this is a built-in shortcut (contains common built-in patterns)
            if (IsBuiltInShortcut(match[1])) {
                ; Add visual distinction for built-in shortcuts with dashes and brackets
                processedLine := "--- " . processedLine
            } else {
                ; Add visual distinction for custom shortcuts
                processedLine := ">>> " . processedLine
            }

            processedLines.Push(processedLine)
        } else {
            processedLines.Push(line)
        }
    }

    ; Join lines back together manually
    result := ""
    for i, line in processedLines {
        if (i = 1) {
            result := line
        } else {
            result := result . "`n" . line
        }
    }

    ; Add legend at the top if there are shortcuts
    if (InStr(result, ">>>") || InStr(result, "---")) {
        separator := ""
        loop 50 {
            separator .= "="
        }
        legend := ">>> Custom shortcuts  |  --- Built-in shortcuts`n" . separator . "`n`n"
        result := legend . result
    }

    return result
}

; Function to detect if a shortcut is a built-in shortcut
IsBuiltInShortcut(shortcut) {
    ; Remove brackets for easier matching
    content := RegExReplace(shortcut, "\[|\]", "")

    ; Common built-in shortcut patterns
    builtInPatterns := [
        "Ctrl \+ [A-Z]",           ; Ctrl + Letter
        "Ctrl \+ [0-9]",           ; Ctrl + Number
        "Ctrl \+ [F1-F12]",        ; Ctrl + Function keys
        "Alt \+ [A-Z]",            ; Alt + Letter
        "Alt \+ [0-9]",            ; Alt + Number
        "Alt \+ [F1-F12]",         ; Alt + Function keys
        "Alt \+ [↑↓←→]",           ; Alt + Arrow keys
        "Ctrl \+ Shift \+ [A-Z]",  ; Ctrl + Shift + Letter
        "Ctrl \+ Shift \+ [0-9]",  ; Ctrl + Shift + Number
        "Ctrl \+ Enter",           ; Ctrl + Enter
        "Ctrl \+ Space",           ; Ctrl + Space
        "Ctrl \+ Tab",             ; Ctrl + Tab
        "Ctrl \+ Esc",             ; Ctrl + Esc
        "Ctrl \+ Home",            ; Ctrl + Home
        "Ctrl \+ End",             ; Ctrl + End
        "Ctrl \+ PageUp",          ; Ctrl + PageUp
        "Ctrl \+ PageDown",        ; Ctrl + PageDown
        "Ctrl \+ Insert",          ; Ctrl + Insert
        "Ctrl \+ Delete",          ; Ctrl + Delete
        "Ctrl \+ Backspace",       ; Ctrl + Backspace
        "Shift \+ [A-Z]",          ; Shift + Letter
        "Shift \+ [0-9]",          ; Shift + Number
        "Shift \+ [F1-F12]",       ; Shift + Function keys
        "Shift \+ [↑↓←→]",         ; Shift + Arrow keys
        "Shift \+ Enter",          ; Shift + Enter
        "Shift \+ Delete",         ; Shift + Delete
        "Shift \+ Tab",            ; Shift + Tab
        "Shift \+ Esc",            ; Shift + Esc
        "F[1-9]|F1[0-2]",         ; Function keys F1-F12
        "Esc",                     ; Escape key
        "Enter",                   ; Enter key
        "Space",                   ; Space key
        "Tab",                     ; Tab key
        "Backspace",               ; Backspace key
        "Delete",                  ; Delete key
        "Insert",                  ; Insert key
        "Home",                    ; Home key
        "End",                     ; End key
        "PageUp",                  ; PageUp key
        "PageDown",                ; PageDown key
        "↑|↓|←|→"                  ; Arrow keys
    ]

    ; Check against built-in patterns
    for pattern in builtInPatterns {
        if RegExMatch(content, "i)^" . pattern . "$") {
            return true
        }
    }

    return false
}

; Helper: normalize common UTF-8→CP1252 mojibake so arrows and punctuation display correctly
NormalizeMojibake(str) {
    if (str = "")
        return str
    reps := Map(
        "â†’", "→",   ; right arrow
        "â†", "←",   ; left arrow
        "â†‘", "↑",   ; up arrow
        "â†“", "↓",   ; down arrow
        "â€¢", "•",
        "â€“", "–",
        "â€”", "—",
        "â€¦", "…",
        "â€˜", "'",
        "â€™", "'",
        "â€œ", Chr(34),
        "â€", Chr(34),
        "Ã—", "×"
    )
    for k, v in reps
        str := StrReplace(str, k, v)
    return str
}

;-------------------------------------------------------------------
; Cheat-sheet overlay (Win + Alt + Shift + A) â€" shows remapped shortcuts
;-------------------------------------------------------------------

; Map that stores the pop-up text for each application.  Extend freely.
cheatSheets := Map()

; --- Mercado Livre (Brazil) -----------------------------------------------
cheatSheets["Mercado Livre"] := "
(
Mercado Livre
[Shift+Y] > 🔍 Focus search field
[Shift+U] > 🛒 Carrinho de compras
[Shift+I] > 📦 Compras feitas
)"  ; end Mercado Livre

;---------------------------------------- Shift + keys ----------------------------------------------
; ----- Assignment policy: use Shift + <key> first. When all Shift slots in the sequence are consumed, continue with Ctrl + Alt + <key> in the same order.
; ----- You can have repeated keys, depending on the software.
; ----- Preferred key sequence (most important first): Y U I O P H J K L N M , . W E R T D F G C V B
; ----- Ctrl + Alt sequence (fallback, same order):    Y U I O P H J K L N M , . W E R T D F G C V B
; ----- Shift + D (Teams chat) -> Fold chat sections (🩶 grey)

; --- WhatsApp desktop -------------------------------------------------------
cheatSheets["WhatsApp"] := "
(
WhatsApp
[Shift+Y] > 🎤 Toggle voice message
[Shift+U] > 🔍 Search chats
[Shift+I] > ↩️ Reply
[Shift+O] > 😀 Sticker panel
[Shift+,] > 📬 Toggle Unread filter
[Shift+H] > 💬 Focus current chat
[Shift+J] > ✅ Mark as read or unread
[Shift+K] > 📌 Pin chat or unpin chat
)"  ; end WhatsApp

; --- Outlook main window ----------------------------------------------------
cheatSheets["OUTLOOK.EXE"] := "
(
Outlook
[Shift+Y] > 📧 Send to General
[Shift+U] > 📰 Send to Newsletter
[Shift+I] > 📥 Go to Inbox
[Shift+O] > 📝 Subject / Title
[Shift+P] > 👥 Required / To
[Shift+H] > 🚫 Don't send any response
[Shift+J] > ✅ Send response 
[Shift+K] > Send Shift+F6
[Shift+L] > Send F6
[Shift+M] > 📝 Subject -> Body
[Shift+N] > 🎯 Focused / Other
)"  ; end Outlook

; --- Outlook Reminder window -------------------------------------------------
cheatSheets["OutlookReminder"] := "
(
Outlook â€" Reminders
[Shift+Y] > 🔔 Select first reminder
[Shift+U] > ⏰ Snooze 1 hour
[Shift+I] > ⏰ Snooze 4 hours
[Shift+O] > ⏰ Snooze 1 day
[Shift+P] > ❌ Dismiss all reminders
[Shift+H] > 🌐 Join Online
)"  ; end Outlook Reminder

; --- Outlook Appointment window ---------------------------------------------
cheatSheets["OutlookAppointment"] := "
(
Outlook â€" Appointment
[Shift+Y] > 📅 Start date (combo)
[Shift+U] > 📅 Start date â€" Date Picker
[Shift+I] > 🕐 Start time (combo)
[Shift+O] > 📅 End date (combo)
[Shift+P] > 🕐 End time (combo)
[Shift+H] > ☑️ All day checkbox
[Shift+J] > 📝 Title field
[Shift+L] > 👥 Required / To field
[Shift+M] > 📍 Location > Body
[Shift+,] > 🔄 Make Recurring
)"  ; end Outlook Appointment

; --- Outlook Message window ---------------------------------------------------
cheatSheets["OutlookMessage"] := "
(
Outlook â€" Message
[Shift+Y] > 📝 Subject / Title
[Shift+U] > 👥 Required / To
[Shift+M] > 📍 Location > Body
)"  ; end Outlook Message

; --- Microsoft Teams â€" meeting window --------------------------------------
cheatSheets["TeamsMeeting"] := "
(
Teams
[Shift+Y] > 💬 Open Chat pane
[Shift+U] > 🔍 Maximize meeting window
[Shift+I] > 👍 Reagir
[Shift+O] > 🎥 Join now with camera and microphone on
[Shift+P] > 🔊 Audio settings
)"  ; end TeamsMeeting

; --- Microsoft Teams â€" chat window -----------------------------------------
cheatSheets["TeamsChat"] := "
(
Teams

--- Custom Shortcuts ---
[Shift+Y] > 👍 Like
[Shift+U] > ❤️ Heart
[Shift+I] > 😂 Laugh
[Shift+O] > 🏠 Home panel
[Shift+P] > 📎 Attach file
[Shift+H] > 📜 Open history menu
[Shift+J] > 📬 Mark unread
[Shift+K] > 📌 Pin chat
[Shift+L] > 📌 Remove pin
[Shift+N] > 📁 Collapse all conversation folders
[Shift+M] > ℹ️ Activate/deactivate details panel
[Shift+,] > 📬 View all unread items
[Shift+.] > 🪟 Detach current chat
[Shift+E] > ✏️ Edit message
[Shift+R] > ↩️ Reply
[Shift+T] > 👥 Add participants
[Shift+W] > 📞 Start call (audio/video)
[Shift+D] > 🩶 Fold chat sections

--- Built-in Shortcuts ---
Geral:
[Ctrl + .] > Show keyboard shortcuts
[Ctrl + E] > Open search
[Ctrl + /] > Show commands
[Ctrl + G] > Go to a chat or channel
[Ctrl + N] > Start new chat
[Ctrl + Shift + N] > Open a new chat
[Ctrl + Shift + F] > Open filter
[Ctrl + ,] > Open Settings
[F1] > Open Help
[Ctrl + =] > Zoom in
[Ctrl + -] > Zoom out
[Ctrl + 0] > Reset zoom level
[Ctrl + O] > Open existing conversation in new window

Navegação:
[Ctrl + 1-9] > Open 1st-9th App in App Bar
[Ctrl + L] > Move focus to left rail item
[Ctrl + M] > Move focus to messages panel
[Ctrl + Alt + T] > Move focus to top system notification
[Alt + Left] > Back
[Alt + Right] > Forward
[Ctrl + H] > Open history menu
[Ctrl + R] > Go to text box
[Ctrl + Alt + Enter] > Focus on resizable divider
[Ctrl + Shift + Enter] > Reset slots to default size
[Win + Shift + Y] > Move focus to notification

Redigir:
[Ctrl + Shift + X] > Expand text box
[Ctrl + Enter] > Send (expanded text box)
[Alt + Shift + O] > Attach file
[Shift + Enter] > Start new line
[Ctrl + B] > Apply bold style
[Ctrl + I] > Apply italic style
[Ctrl + U] > Apply underline style
[Alt + A] > Rewrite with Copilot
[Alt + Shift + E] > Open video recorder
[Ctrl + Alt + L] > Add a Loop paragraph
[Ctrl + Shift + I] > Mark message as important
[Ctrl + K] > Insert link
[Ctrl + Alt + Shift + C] > Insert embedded code
[Ctrl + Alt + Shift + B] > Insert code block

Mensagens:
[Alt + Q] > Collapse all conversation folders
[Ctrl + J] > Go to last read/new message
[Ctrl + Alt + R] > React to last message
[Alt + P] > Activate/deactivate details panel
[Alt + Shift + R] > Reply to last message
[Alt + 1-9] > Open 1st-9th Tab in Chat Panel Header
[Ctrl + Alt + Z] > Clear all filters
[Ctrl + Alt + U] > View all unread items
[Ctrl + Alt + B] > View all meeting items
[Ctrl + Alt + C] > View all chat conversations
[Ctrl + Alt + A] > View all channel conversations
[Ctrl + F] > Search current Chat/Channel messages
[Alt + T] > Open Threads List
)"  ; end TeamsChat

; --- Spotify ---------------------------------------------------------------
cheatSheets["Spotify.exe"] := "
(
Spotify
[Shift+Y] > 🔗 Toggle Connect panel
[Shift+U] > 🖥️ Toggle Full screen
[Shift+I] > 🔍 Open Search
[Shift+O] > 📋 Go to Playlists
[Shift+P] > 🎤 Go to Artists
[Shift+H] > 💿 Go to Albums
[Shift+J] > 🔍 Go to Search
[Shift+K] > 🏠 Go to Home
[Shift+L] > 🎵 Go to Now Playing
[Shift+N] > 🎯 Go to Made For You
[Shift+M] > 🆕 Go to New Releases
[Shift+,] > 📊 Go to Charts
[Shift+.] > 🎵 Toggle Now Playing View
[Shift+W] > 📚 Toggle Library Sidebar
[Shift+E] > 🖥️ Toggle Fullscreen Library
[Shift+R] > 🎤 Toggle lyrics
[Shift+T] > ⏯️ Toggle play/pause
)"  ; end Spotify

; --- OneNote ---------------------------------------------------------------
cheatSheets["ONENOTE.EXE"] := "
(
OneNote
[Shift+Y] > 📉 Collapse
[Shift+U] > 📈 Expand
[Shift+I] > 📉 Collapse all
[Shift+O] > 📈 Expand all
[Shift+P] > 📝 Select line and children
[Shift+D] > 🗑️ Delete line and children
[Shift+S] > 🗑️ Delete line (keep children)
[Shift+F] > 🔍 Advanced Searching with double quotes
)"  ; end OneNote

; --- Chrome general shortcuts ----------------------------------------------
cheatSheets["chrome.exe"] := "
(
Chrome
[Shift+G] > 🪟 Pop current tab to new window
)"  ; end Chrome

; --- Cursor ------------------------------------------------------
cheatSheets["Cursor.exe"] := "
(
Cursor

--- CTRL Shortcuts (Cursor-defined) ---
[Ctrl+1] > 🎯 Remove clustering and focus on the code (ahk)
[Ctrl+2] > 📁 Copy path (cursor)
[Ctrl+M] > 🤖 Ask (Ctrl+Alt+A), wait 6s, then paste (Shift+V) (ahk)
[Ctrl+G] > ⚡ Kill terminal [custom in settings.json]
[Ctrl+Y] > 📉 Fold all
[Ctrl+U] > 📈 Unfold all
[Ctrl+O] > 📋 Paste As...
[Ctrl+H] > 📁 Reveal in file explorer
[Ctrl+J] > 🔲 Select to Bracket
[Ctrl+,] > 📉 Fold all directories
[Ctrl+.] > 💬 Toggle chat or agent
[Ctrl+Q] > 📈 Unfold all directories
[Ctrl+E] > 🤖 Open Agent Window
[Ctrl+R] > 📂 File open Recent
[Ctrl+T] > 🔍 Go to symbol in workspace
[Ctrl+D] > 📝 Add selection to next find match
[Ctrl+F] > 🔍 Find
[Ctrl+Z] > ↩️ Undo
[Ctrl+B] > 📊 Toggle primary sidebar visibility

--- SHIFT Shortcuts (ahk = AutoHotkey) ---
[Shift+Y] > 📉 Fold (ahk)
[Shift+U] > 📈 Unfold (ahk)
[Shift+I] > 📄 Open markdown preview to the side (cursor)
[Shift+O] > 🪟 Move editor into new window (cursor)
[Shift+P] > 💻 Go to terminal (ahk)
[Shift+H] > 💻 New terminal (ahk)
[Shift+J] > 📁 Go to file explorer (ahk)
[Shift+K] > 📄🪟 Open markdown preview and move editor into new window (ahk)
[Shift+L] > ⌨️ Command palette (ahk)
[Shift+N] > 📈 Expand selection (ahk)
[Shift+M] > ⚡ Go to symbol in access view (cursor)
[Shift+,] > 💬 Show chat history (ahk)
[Shift+.] > 🖼️ Paste Image (cursor)
[Shift+W] > 📁 Fold Git repos (SCM) (ahk)
[Shift+E] > 🔍 Search (ahk)
[Shift+R] > 🍞 Open Bread Crumbs menu (ahk)
[Shift+T] > 😀 Emoji selector (1:🔲 2:⏳ 3:⚡ 4:2️⃣ 5:❓) (ahk)
[Shift+D] > 🌿 Git section (ahk)
[Shift+F] > ❌ Close all editors (ahk)
[Shift+G] > 🤖 Switch AI models (auto/CLAUD/GPT/O/DeepSeek/Cursor) (ahk)
[Shift+Z] > 🧘 Zen mode (cursor)
[Shift+C] > ⬇️ Git Pull (cursor)
[Shift+V] > ✅ Git Commit (cursor)
[Shift+B] > ⬆️ Git Push (cursor)

--- CTRL+ALT Shortcuts (Cursor-defined) ---
[Ctrl+Alt+Up] > ⬆️ Go to Parent Fold
[Ctrl+Alt+Left] > ⬅️ Go to sibling fold previous
[Ctrl+Alt+Right] > ➡️ Go to sibling fold next

--- Additional Shortcuts ---
[Ctrl + T] > 💬 New chat tab
[Ctrl + N] > 💬 New chat tab (replacing current)
[Alt + F12] > 👁️ Peek Definition
[Ctrl + Shift + L] > 📝 Select all identical words
[F2] > ✏️ Rename symbol
[F8] > 🔍 Navigate problems
[Ctrl + Enter] > ➕ Insert line below
[Ctrl + P] > 🔍 Quick Open
[Shift + Delete] > 🗑️ Delete line
[Alt + ↑] > ⬆️ Move line up
[Alt + ↓] > ⬇️ Move line down
[Ctrl + 1 / Ctrl + 2 / Ctrl + 3 ...] > 🔄 Switch tabs
[Ctrl + Alt + ↑] > ⬆️ Add cursor above
[Ctrl + Alt + ↓] > ⬇️ Add cursor below
[Alt + Click] > 👆 Multi-cursor by click
[Shift + Alt + ↑] > ⬆️ Copy line up
[Shift + Alt + ↓] > ⬇️ Copy line down
[Ctrl + ;] > 💬 Insert comment
[Ctrl + Home] > ⬆️ Go to top
[Ctrl + End] > ⬇️ Go to end
[Alt + Z] > 🔄 Toggle word wrap
[Ctrl + Shift + D] > 🐛 Debugging
[Ctrl + R] > 🔄 Quick project switch
[Alt + J] > ⬇️ Next review
[Alt + K] > ⬆️ Previous review
)"  ; end Cursor

; --- Windows Explorer ------------------------------------------------------
cheatSheets["explorer.exe"] := "
(
Explorer
[Shift+Y] > 📄 Select first file
[Shift+U] > 🔍 Focus search bar
[Shift+I] > 📍 Focus address bar
[Shift+O] > 📁 New folder
[Shift+J] > 🔗 Create a shortcut
[Shift+K] > 📋 Copy as path
[Shift+P] > 📌 Select first pinned item in Explorer sidebar
[Shift+H] > 📌 Select the last item of the Explorer sidebar
)"  ; end Explorer

; --- Microsoft Paint ------------------------------------------------------
cheatSheets["mspaint.exe"] := "
(
MS Paint
[Shift+Y] > 📏 Resize and Skew (Ctrl+W)

--- Common Shortcuts ---
[Ctrl+N] > 📄 New
[Ctrl+O] > 📂 Open
[Ctrl+S] > 💾 Save
[F12] > 💾 Save As
[Ctrl+P] > 🖨️ Print
[Ctrl+Z] > ↩️ Undo
[Ctrl+Y] > ↪️ Redo
[Ctrl+A] > 📄 Select all
[Ctrl+C] > 📋 Copy
[Ctrl+X] > ✂️ Cut
[Ctrl+V] > 📋 Paste
[Ctrl+W] > 📏 Resize and Skew
[Ctrl+E] > ℹ️ Image properties
[Ctrl+R] > 📏 Toggle rulers
[Ctrl+G] > 🔲 Toggle gridlines
[Ctrl+I] > 🔄 Invert colors
[F11] > 🖥️ Fullscreen view
[Ctrl++] > 🔍 Zoom in
[Ctrl+-] > 🔍 Zoom out
)"  ; end MS Paint

; --- ClipAngel -------------------------------------------------------------
cheatSheets["ClipAngel.exe"] := "
(
ClipAngel
[Shift+Y] > 📋 Select filtered content and copy
[Shift+U] > 🔄 Switch focus list/text
[Shift+I] > 🗑️ Delete all non-favorite
[Shift+O] > 🧹 Clear filters
[Shift+P] > ⭐ Mark as favorite
[Shift+H] > ⭐ Unmark as favorite
[Shift+J] > ✏️ Edit text
[Shift+K] > 💾 Save as file
[Shift+L] > 🔗 Merge clips
)"  ; end ClipAngel

; --- Figma -----------------------------------------------------------------
cheatSheets["Figma.exe"] := "
(
Figma
[Shift+Y] > 👁️ Show/Hide UI
[Shift+U] > 🔍 Component search
[Shift+I] > ⬆️ Select parent
[Shift+O] > 🧩 Create component
[Shift+P] > 🔗 Detach instance
[Shift+H] > 📐 Add auto layout
[Shift+J] > 📐 Remove auto layout
[Shift+K] > 💡 Suggest auto layout
[Shift+L] > 📤 Export
[Shift+N] > 🖼️ Copy as PNG
[Shift+M] > ⚡ Actions...
[Shift+,] > ⬅️ Align left
[Shift+.] > ➡️ Align right
[Shift+W] > 📏 Distribute vertical spacing
[Shift+E] > 🧹 Tidy up
[Shift+R] > ⬆️ Align top
[Shift+T] > ⬇️ Align bottom
[Shift+D] > ↔️ Align center horizontal
[Shift+F] > ↕️ Align center vertical
[Shift+G] > 📏 Distribute horizontal spacing
)"  ; end Figma

; --- Gmail ---------------------------------------------------------------
cheatSheets["Gmail"] := "
(
Gmail
[Shift+Y] > 📥 Go to main inbox
[Shift+U] > 📰 Go to updates
[Shift+I] > 💬 Go to forums
[Shift+O] > 📬 Toggle read/unread
[Shift+P] > ⬅️ Previous conversation
[Shift+H] > ➡️ Next conversation
[Shift+J] > 📦 Archive conversation
[Shift+K] > ✅ Select conversation
[Shift+L] > ↩️ Reply
[Shift+N] > ↩️ Reply all
[Shift+M] > ➡️ Forward
[Shift+,] > ⭐ Star/unstar conversation
[Shift+.] > 🗑️ Delete
[Shift+W] > 🚫 Report as spam
[Shift+E] > ✍️ Compose new email
[Shift+R] > 🔍 Search mail
[Shift+T] > 📁 Move to folder
[Shift+D] > ⌨️ Show keyboard shortcuts help
[Shift+F] > 📬 Click inbox button

--- Built-in Shortcuts (Windows) ---

Compose & chat:
[p] > Previous message in an open conversation
[n] > Next message in an open conversation
[Shift + Esc] > Focus main window
[Esc] > Focus latest chat or compose
[Ctrl + .] > Advance to the next chat or compose
[Ctrl + ,] > Advance to previous chat or compose
[Ctrl + Enter] > Send
[Ctrl + Shift + c] > Add cc recipients
[Ctrl + Shift + b] > Add bcc recipients
[Ctrl + Shift + f] > Access custom from
[Ctrl + k] > Insert a link
[Ctrl + m] > Open spelling suggestions

Formatting text:
[Ctrl + Shift + 5] > Previous font
[Ctrl + Shift + 6] > Next font
[Ctrl + Shift + -] > Decrease text size
[Ctrl + Shift + +] > Increase text size
[Ctrl + b] > Bold
[Ctrl + i] > Italics
[Ctrl + u] > Underline
[Ctrl + Shift + 7] > Numbered list
[Ctrl + Shift + 8] > Bulleted list
[Ctrl + Shift + 9] > Quote
[Ctrl + []] > Indent less
[Ctrl + ]] > Indent more
[Ctrl + Shift + l] > Align left
[Ctrl + Shift + e] > Align center
[Ctrl + Shift + r] > Align right
[Ctrl + \] > Remove formatting

Actions (shortcuts on):
[,] > Move focus to toolbar
[x] > Select conversation
[s] > Toggle star/rotate among superstars
[e] > Archive
[m] > Mute conversation
[!] > Report as spam
[#] > Delete
[r] > Reply
[Shift + r] > Reply in a new window
[a] > Reply all
[Shift + a] > Reply all in a new window
[f] > Forward
[Shift + f] > Forward in a new window
[Shift + n] > Update conversation
[] or []] > Archive conversation and go previous/next
[z] > Undo last action
[Shift + i] > Mark as read
[Shift + u] > Mark as unread
[_] > Mark unread from the selected message
[+ or =] > Mark as important
[-] > Mark as not important
[b] > Snooze (not available in classic Gmail)
[;] > Expand entire conversation
[:] > Collapse entire conversation
[Shift + t] > Add conversation to Tasks

Jumping (shortcuts on):
[g + i] > Go to Inbox
[g + s] > Go to Starred conversations
[g + b] > Go to Snoozed conversations
[g + t] > Go to Sent messages
[g + d] > Go to Drafts
[g + a] > Go to All mail
[Ctrl + Alt + ,] > Switch to left sidebar (Calendar/Keep/Tasks)
[Ctrl + Alt + .] > Switch to right (back to inbox)
[g + k] > Go to Tasks
[g + l] > Go to label

Threadlist selection (shortcuts on):
[* + a] > Select all conversations
[* + n] > Deselect all conversations
[* + r] > Select read conversations
[* + u] > Select unread conversations
[* + s] > Select starred conversations
[* + t] > Select unstarred conversations

Navigation (shortcuts on):
[g + n] > Go to next page
[g + p] > Go to previous page
[u] > Back to threadlist
[k] > Newer conversation
[j] > Older conversation
[o or Enter] > Open conversation
[`] > Go to next Inbox section
[~] > Go to previous Inbox section

Application (shortcuts on):
[c] > Compose
[d] > Compose in a new tab
[/] > Search mail
[q] > Search chat contacts
[.] > Open ""more actions"" menu
[v] > Open ""move to"" menu
[l] > Open ""label as"" menu
[?] > Open keyboard shortcut help
)"  ; end Gmail

; --- Google Keep ---------------------------------------------------------------
cheatSheets["Google Keep"] := "
(
Google Keep
[Shift+Y] > 🔍 Search and select note
[Shift+U] > 📋 Toggle main menu
)"  ; end Google Keep

; --- File Dialog ---------------------------------------------------------------
cheatSheets["FileDialog"] := "
(
File Dialog
[Shift+Y] > 📄 Select first item
[Shift+U] > 🔍 Focus search bar
[Shift+I] > 📍 Focus address bar
[Shift+O] > 📁 New folder
[Shift+P] > 📌 Select first pinned item in sidebar
[Shift+H] > 💻 Select 'This PC' / 'Este computador' in sidebar
[Shift+J] > 📝 Focus file name field
[Shift+K] > ✅ Click Insert/Open/Save button
[Shift+L] > ❌ Click Cancel button
)"

; --- Settings Window -------------------------------------------------
cheatSheets["Settings"] := "(Settings)`r`n[Shift+Y] > 🔊 Set input volume to 100%"

; --- Command Palette -------------------------------------------------
cheatSheets["Command Palette"] := "
(
Command Palette
[Ctrl+H] > ⌨️ Open in folder (Ctrl+Shift+E)
[Shift+K] > ⌨️ Copy file path (Ctrl+Shift+C)
[Shift+Y] > ⌨️ Send ten backspaces
[Shift+U] > ⌨️ Precise search
[Shift+I] > ⌨️ Add favorit
[Ctrl+1] > ⌨️ Select current item (Enter)
[Ctrl+2] > ⌨️ Move down once and select
[Ctrl+3] > ⌨️ Move down twice and select
[Ctrl+4] > ⌨️ Move down three times and select
[Ctrl+5] > ⌨️ Move down four times and select
[Ctrl+6] > ⌨️ Move down five times and select
)"

; --- Excel ------------------------------------------------------------
cheatSheets["EXCEL.EXE"] := "
(
Excel
[Shift+Y] > ⚪ Select White Color
[Shift+U] > ✏️ Enable Editing
[Shift+I] > 📊 Turn CSV delimited by semicolon into columns
)"

; --- Power BI ------------------------------------------------------------
cheatSheets["Power BI"] := "
(
Power BI
[Shift+Y] > 📊 Transform data
[Shift+U] > 📊 Close and apply
[Shift+I] > 📊 Report view
[Shift+O] > 📊 Table view
[Shift+P] > 📊 Model view
[Shift+H] > 📊 Analytics
[Shift+J] > 📊 Format visual
[Shift+K] > 📊 Build visual
[Shift+L] > ✅ OK/Confirm modal button
[Shift+N] > ❌ Cancel/Exit modal button
[Shift+M] > 🖱️ Right-click Previous pages button
[Shift+,] > 📋 Filter pane collapse/expand
[Shift+.] > 🎨 Visualizations pane toggle
[Shift+W] > ➕ New page
[Shift+E] > 📊 New measure
[Shift+R] > 📁 Collapse Fields tables
[Shift+Q] > 📊 Data pane toggle
)"

; --- UIA Tree Inspector -------------------------------------------------
cheatSheets["UIATreeInspector"] :=
"(UIA Tree Inspector)`r`n[Shift+Y] > 🔄 Refresh list`r`n[Shift+U] > 🔍 Focus filter field"
; --- SettleUp Shortcuts -----------------------------------------------------
cheatSheets["Settle Up"] := "
(
Settle Up
[Shift+Y] > ➕ Add transaction
[Shift+U] > 💰 Focus expense value field
[Shift+I] > 📝 Focus expense name field
)"

; --- Miro Shortcuts -----------------------------------------------------
cheatSheets["Miro"] := "
(
Miro
[Shift+U] > 📋 Frame list
[Shift+I] > 🔗 Group
[Shift+O] > 🔗 Ungroup
[Shift+P] > 🔒 Lock/Unlock
[Shift+H] > 🔗 Add/Edit link
Q
--- Built-in Shortcuts (Windows) ---
Tools:
[V / H] > Select tool / Hand
[T] > Text
[N] > Sticky notes
[S] > Shapes
[R] > Rectangle
[O] > Oval
[L] > Connection line / Arrow
[D] > Card
[P] > Pen
[E] > Eraser
[C] > Comment
[F] > Frame
[M] > Minimap
[Ctrl + K] > Command palette
[Enter (bulk)] > New sticky note
[Esc (bulk)] > Exit sticky note bulk mode
[Ctrl + Shift + Enter] > Open card panel
[Shift + C] > Show/hide comments

General:
[Ctrl + C / Ctrl + V] > Copy / Paste
[Ctrl + X] > Cut
[Ctrl + D] > Duplicate
[Alt + drag] > Duplicate by drag
Alt + â†â†’â†‘â†“        â†’  Duplicate horizontally/vertically
[Ctrl + click] > Select multiple
[Ctrl + A] > Select all
[Enter] > Edit selected
[Esc] > Deselect / quit edit
[Backspace] > Delete
[Ctrl + G] > Group
[Ctrl + Shift + G] > Ungroup
[Ctrl + Shift + L] > Lock / Unlock
[Ctrl + Shift + P] > Protected lock / Unprotected lock
[PgUp] > Bring to front
[Shift + PgUp] > Bring forward
[PgDn] > Send to back
[Shift + PgDn] > Send backward
[Ctrl + Shift + K] > Create board in new tab
[Alt + Ctrl + K] > Add/Edit link to object
[Ctrl + Backspace] > Clear object contents

Navigation:
â†â†'â†'              â†'  Move items/canvas
[Ctrl + +] > Zoom in
[Ctrl + -] > Zoom out
[Ctrl + 0] > Zoom to 100%
[Alt + 1] > Zoom to fit
[Alt + 2] > Zoom to selected item
[Space + drag] > Move canvas
[G] > Toggle grid
[Ctrl + F] > Search

Text:
[Ctrl + B] > Bold
[Ctrl + I] > Italic
[Ctrl + U] > Underline

Board navigation:
[Tab] > Move forwards through objects (TL > BR)
[Shift + Tab] > Move backwards through objects (TL > BR)
Ctrl + â†'/+â†"/â†/â†'    â†'  Move through board objects
[Ctrl + Shift + ↓/↑] > Move in/out of container (e.g., frame)
[Esc] > Back to menu
[Enter] > Edit an object
[Esc] > Stop editing an object

Toolbar navigation:
[Tab / Shift + Tab] > Move between toolbars
[Arrow keys] > Move between toolbar items
[Enter / Space] > Activate a menu item

Desktop app:
[Ctrl + R] > Reload the tab
[Ctrl + W] > Close the tab
[Ctrl + Q] > Exit the app
[Ctrl + Shift + L] > Copy board link
)"

; --- Wikipedia ---------------------------------------------------------------
cheatSheets["Wikipedia"] := "
(
Wikipedia
[Shift+Y] > 🔍 Click search button
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

    ; Check for Command Palette window
    if InStr(title, "Command Palette", false) {
        return cheatSheets.Has("Command Palette") ? cheatSheets["Command Palette"] : ""
    }

    ; Check for Power BI (by process name or window title)
    if (exe = "PBIDesktop.exe" || InStr(title, "powerbi", false)) {
        return cheatSheets.Has("Power BI") ? cheatSheets["Power BI"] : ""
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
                "ChatGPT`r`n[Shift+Y] > Cut all`r`n[Shift+U] > Model selector`r`n[Shift+I] > Toggle sidebar`r`n[Shift+O] > Re-send rules`r`n[Shift+H] > Copy code block`r`n[Shift+J] > Go down`r`n[Shift+L] > Send and show AI banner"
        if InStr(chromeTitle, "Mobills")
            appShortcuts :=
                "Mobills - Navigation`r`n[Shift+Y] > Dashboard`r`n[Shift+U] > Contas`r`n[Shift+I] > TransaÃ§Ãµes`r`n[Shift+O] > CartÃµes de crÃ©dito`r`n[Shift+P] > Planejamento`r`n[Shift+H] > RelatÃ³rios`r`n[Shift+J] > Mais opÃ§Ãµes`r`n[Shift+K] > Previous month`r`n[Shift+L] > Next month`r`n`r`nMobills - Actions`r`n[Shift+N] > Ignore transaction`r`n[Shift+M] > Name field`r`n[Shift+E] > New Expense`r`n[Shift+R] > New Income`r`n[Shift+T] > New Credit expense`r`n[Shift+D] > New Transfer`r`n[Shift+W] > Open button + type MAIN"
        if InStr(chromeTitle, "Google Keep") || InStr(chromeTitle, "keep.google.com")
            appShortcuts := cheatSheets.Has("Google Keep") ? cheatSheets["Google Keep"] : ""
        if InStr(chromeTitle, "YouTube")
            appShortcuts :=
                "YouTube`r`n[Shift+Y] > Focus search box`r`n[Shift+U] > Focus first video via Search filters`r`n[Shift+I] > Focus first video via Explore"
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
                appShortcuts := "Google`r`n[Shift+Y] > Focus search box"
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

    ; Microsoft Teams â€" differentiate meeting vs chat via helper predicates
    if IsTeamsMeetingActive()
        return cheatSheets.Has("TeamsMeeting") ? cheatSheets["TeamsMeeting"] : ""
    if IsTeamsChatActive()
        return cheatSheets.Has("TeamsChat") ? cheatSheets["TeamsChat"] : ""
    if IsFileDialogActive()
        return cheatSheets["FileDialog"]

    ; Special handling for Outlook-based apps
    if (exe = "OUTLOOK.EXE") {
        ; Detect Reminders window â€" e.g. "3 Reminder(s)" or any title containing "Reminder"
        if RegExMatch(title, "i)Reminder") {
            return cheatSheets.Has("OutlookReminder") ? cheatSheets["OutlookReminder"] : cheatSheets["OUTLOOK.EXE"]
        }
        ; Detect Message inspector windows â€" e.g., " - Message (HTML)"
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

    ; Nothing found > blank > caller will show fallback message
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
        cheatCtrl := g_helpGui.Add("Edit",
            "ReadOnly +Multi -E0x200 +VScroll -HScroll -Border Background000000 w1000 r1"
        )

        ; Esc also hides  ; (disabled â€" use Win+Alt+Shift+A to hide)
        ; Hotkey "Esc", (*) => (g_helpGui.Hide(), g_helpShown := false), "Off"
    }

    ; Update cheat-sheet text and resize height to fit
    ; Process the text to pad shortcuts for alignment
    processedText := ProcessCheatSheetText(text)
    cheatCtrl.Value := processedText
    lineCnt := StrLen(processedText) ? StrSplit(processedText, "`n").Length : 1

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
    cheatCtrl.Move(, , 1000, controlHeight)
    ; Show > measure > centre
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
    ; Get hotstrings section first
    hsText := ""
    try {
        hsText := GetHotstringsCheatSheetText()
    } catch {
    }

    globalText := ""

    ; Add hotstrings at the top if any are defined
    if (StrLen(hsText)) {
        globalText .= "=== HOTSTRINGS ===`n" hsText "`n`n"
    }

    globalText .= "
(
=== AVAILABLE (unused) ===
[Win+Alt+Shift+P] 
[Win+Alt+Shift+U]

=== CURSOR ===
[Win+Alt+Shift+N] > Opens or activates Cursor (habits, home, punctual, or work windows)

=== SPOTIFY ===
[Win+Alt+Shift+S] > Opens or activates Spotify

r=== CLIP ANGEL ===
[Win+Alt+Shift+1] > Send top list item from Clip Angel

=== CHATGPT ===
[Win+Alt+Shift+8] > Get word pronunciation, definition, and Portuguese translation
[Win+Alt+Shift+0] > Speak with ChatGPT
[Win+Alt+Shift+7] > Speak with ChatGPT (send message automatically)
[Win+Alt+Shift+J] > Copy last ChatGPT message (in the editing box)
[Win+Alt+Shift+O] > Check grammar and improve text in both English and Portuguese
[Win+Alt+Shift+I] > Opens ChatGPT
[Win+Alt+Shift+L] > Talk with ChatGPT through voice
[Win+Alt+Shift+Y] > Copy last message and read it aloud

=== YOUTUBE ===
[Win+Alt+Shift+H] > Activates Youtube

=== GOOGLE ===
[Win+Alt+Shift+F] > Opens Google

=== GMAIL ===
[Win+Alt+Shift+W] > Opens Gmail

=== CURSOR ===
[Win+Alt+Shift+,] > Opens or activates Cursor
[Win+Alt+Shift+C] > Activates Cursor with action options: 1) Proceed with terminal 2) Hit Enter 3) Allow

=== OUTLOOK ===
[Win+Alt+Shift+B] > Open mail
[Win+Alt+Shift+V] > Open Reminder
[Win+Alt+Shift+G] > Open calendar
[Win+Alt+Shift+D] > Voice aloud the email

=== MICROSOFT TEAMS ===
[Win+Alt+Shift+R] > New conversation
[Win+Alt+Shift+5] > Toggle Mute (meeting)
[Win+Alt+Shift+4] > Toggle camera (meeting)
[Win+Alt+Shift+T] > Screen share (meeting)
[Win+Alt+Shift+2] > Exit meeting
[Win+Alt+Shift+E] > Select the chats window
[Win+Alt+Shift+3] > Select the meeting window

=== WHATSAPP ===
[Win+Alt+Shift+Z] > Opens WhatsApp

=== WINDOWS ===
[Win+Alt+Shift+6] > Minimizes windows
[Win+Alt+Shift+M] > Maximizes the current window

=== GENERAL ===
[Win+Alt+Shift+Q] > Jump mouse on the middle
[Win+Alt+Shift+X] > Activate hunt and Peck
[Win+Alt+Shift+→] > Jump mouse 100px right
[Win+Alt+Shift+←] > Jump mouse 100px left
[Win+Alt+Shift+↓] > Jump mouse 100px down
[Win+Alt+Shift+↑] > Jump mouse 100px up
[Win+Alt+Shift+9] > Pomodoro
[Win+Alt+Shift+.] > Clip Angel (copy, paste, and quit)

=== SHORTCUTS ===
[Win+Alt+Shift+A] > Show app-specific shortcuts (quick press)
[Win+Alt+Shift+A] > Show global shortcuts (hold 700ms+)

=== WIKIPEDIA ===
[Win+Alt+Shift+K] > Opens or activates Wikipedia
)"

    static globalCtrl

    if !IsObject(g_globalGui) {
        g_globalGui := Gui(
            "+AlwaysOnTop -Caption +ToolWindow +Border +Owner +LastFound"
        )
        g_globalGui.BackColor := "000000"
        g_globalGui.SetFont("s10 c00BFFF", "Consolas")  ; Smaller font for more content, blue color to distinguish from specific shortcuts
        globalCtrl := g_globalGui.Add("Edit", "ReadOnly +Multi +VScroll -HScroll -Border Background000000 w1000 h540")

        ; Esc also hides  ; (disabled â€" use Win+Alt+Shift+A to hide)
        ; Hotkey "Esc", (*) => (g_globalGui.Hide(), g_globalShown := false), "Off"
    }

    ; Fix mojibake (arrows, punctuation), pad shortcuts, then update text and show
    normalizedText := NormalizeMojibake(globalText)
    processedText := ProcessCheatSheetText(normalizedText)
    globalCtrl.Value := processedText
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

; =============================================================================
; Send Top List Item from Clip Angel
; Hotkey: Win+Alt+Shift+1
; =============================================================================
#!+1::
{
    Send "!v"
    Sleep 50
    Send "^!b"
}

;-------------------------------------------------------------------
; Environment paths (unchanged)
;-------------------------------------------------------------------
global WORK_SCRIPTS_PATH := "C:\Users\fie7ca\Documents\01 - Scripts"
global PERSONAL_SCRIPTS_PATH := "G:\Meu Drive\12 - Scripts"
; global IS_WORK_ENVIRONMENT   := true    ; set to false on personal rig // This will now be loaded from env.ahk

; ---------------------------------------------------------------------------
; ShowErr(msgOrErr)  â€" uniform MsgBox for any thrown value
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

    ; Get the active window - try multiple methods for better reliability
    activeWin := 0
    try {
        activeWin := WinGetID("A")
    } catch {
        ; Fallback: try to get the foreground window
        activeWin := DllCall("GetForegroundWindow", "ptr")
    }

    ; Default to primary monitor work area (monitor #1)
    MonitorGetWorkArea(1, &lPrim, &tPrim, &rPrim, &bPrim)
    wx := lPrim, wy := tPrim, ww := rPrim - lPrim, wh := bPrim - tPrim

    ; If we have an active window, find which monitor contains its center
    if (activeWin && activeWin != 0) {
        rect := Buffer(16, 0)
        if (DllCall("GetWindowRect", "ptr", activeWin, "ptr", rect)) {
            ; Calculate window center
            winLeft := NumGet(rect, 0, "int")
            winTop := NumGet(rect, 4, "int")
            winRight := NumGet(rect, 8, "int")
            winBottom := NumGet(rect, 12, "int")

            cx := winLeft + (winRight - winLeft) // 2
            cy := winTop + (winBottom - winTop) // 2

            ; Find which monitor contains the window center
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

    ; Calculate center position with bounds checking
    guiX := wx + (ww - guiW) / 2
    guiY := wy + (wh - guiH) / 2

    ; Ensure the GUI stays within monitor bounds
    guiX := Max(wx, Min(guiX, wx + ww - guiW))
    guiY := Max(wy, Min(guiY, wy + wh - guiH))

    ; Ensure minimum position (avoid negative coordinates)
    guiX := Max(0, guiX)
    guiY := Max(0, guiY)

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
+,::
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
                ; Assume you clicked Send manually > reset & start new rec
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
ClickGenerateCommitMessageButton() {
    try {
        ; Use UIA_Browser to get the root element (similar to other functions in the script)
        uia := UIA_Browser()
        if !IsObject(uia) {
            ; Fallback: try Ctrl+M shortcut if UIA fails
            Send "^m"
            return true
        }

        ; Find the "Generate Commit Message (Ctrl+M)" button
        ; Try multiple search strategies
        btn := uia.FindFirst({ Name: "Generate Commit Message (Ctrl+M)", ControlType: "Button" })

        ; If not found by exact name, try partial match
        if !btn {
            btn := uia.FindFirst({ Name: "Generate Commit Message", ControlType: "Button" })
        }

        ; If still not found, try by ControlType only (Type: 50000 = Button)
        if !btn {
            ; Get all buttons and find the one with the right name
            buttons := uia.FindAll({ ControlType: "Button" })
            for button in buttons {
                if InStr(button.Name, "Generate Commit Message") {
                    btn := button
                    break
                }
            }
        }

        if btn {
            btn.Click()
            return true
        } else {
            ; Fallback: try Ctrl+M shortcut
            Send "^m"
            return true
        }
    }
    catch Error as e {
        ; Fallback: try Ctrl+M shortcut if UIA fails
        Send "^m"
        return true
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
; Shift+U  > 1 minuto
; Shift+I  > 2 minutos
; Shift+O  > 3 minutos
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
; Microsoft Teams Shortcuts â€" MEETING WINDOW
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
        if !btn
            btn := root.FindFirst({ Name: "Bate-papo", ControlType: "Button" })

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
        if !btn {
            try {
                btn := root.FindFirst({ Name: "React", ControlType: "Button" })
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

        ; Fallback: try finding by name (Portuguese)
        if !btn {
            btn := root.FindFirst({ Name: "Ingressar agora Com a cÃ¢mera ligada e Microfone ligado", ControlType: "Button" })
        }

        ; Fallback: try finding by name (English)
        if !btn {
            btn := root.FindFirst({ Name: "Join now with camera and microphone on", ControlType: "Button" })
        }

        ; Fallback: try finding by partial name (Portuguese)
        if !btn {
            btn := root.FindFirst({ Name: "Ingressar agora", ControlType: "Button" })
        }

        ; Fallback: try finding by partial name (English)
        if !btn {
            btn := root.FindFirst({ Name: "Join now", ControlType: "Button" })
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

        ; Fallback: try finding by name (Portuguese)
        if !btn {
            btn := root.FindFirst({ Name: "Microfone do computador e controles do alto-falante ConfiguraÃ§Ãµes de Ã¡udio",
                ControlType: "Button" })
        }

        ; Fallback: try finding by name (English)
        if !btn {
            btn := root.FindFirst({ Name: "Computer microphone and speaker controls Audio settings",
                ControlType: "Button" })
        }

        ; Fallback: try finding by partial name (Portuguese)
        if !btn {
            btn := root.FindFirst({ Name: "ConfiguraÃ§Ãµes de Ã¡udio", ControlType: "Button" })
        }

        ; Fallback: try finding by partial name (English)
        if !btn {
            btn := root.FindFirst({ Name: "Audio settings", ControlType: "Button" })
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

; Shift + Y: Focus the Wikipedia search field (Type 50004 Edit)
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

        ; Try finding by AutomationId first (most reliable)
        try {
            searchBox := root.FindElement({ AutomationId: "searchInput" })
        } catch {
        }

        ; Try finding by Type and Name
        if (!searchBox) {
            try {
                searchBox := root.FindElement({ Type: 50004, Name: "Search Wikipedia" })
            } catch {
            }
        }

        ; Try finding by Type and ClassName
        if (!searchBox) {
            try {
                searchBox := root.FindElement({ Type: 50004, ClassName: "cdx-text-input__input mw-searchInput" })
            } catch {
            }
        }

        ; Try finding by partial ClassName
        if (!searchBox) {
            try {
                searchBox := root.FindElement({ Type: 50004, ClassName: "cdx-text-input__input" })
            } catch {
            }
        }

        ; Try finding by AcceleratorKey
        if (!searchBox) {
            try {
                searchBox := root.FindElement({ Type: 50004, AcceleratorKey: "Alt+f" })
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
                    searchBox := root.FindElement({ AutomationId: "searchInput" })
                } catch {
                }
                if (!searchBox) {
                    try {
                        searchBox := root.FindElement({ Type: 50004, Name: "Search Wikipedia" })
                    } catch {
                    }
                }
                if (!searchBox) {
                    try {
                        searchBox := root.FindElement({ Type: 50004, ClassName: "cdx-text-input__input" })
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
    Send "r"
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

; Shift + H : Open history menu
+h::
{
    Send "^h"
}

; Shift + P : Attach file
+p::
{
    Send "!+o"
}

; Shift + N : Collapse all conversation folders
+n::
{
    Send "!q"
}

; Shift + M : Activate/deactivate details panel
+m::
{
    Send "!p"
}

; Shift + , : View all unread items
+,::
{
    Send "^!u"
}

; Shift + . : Detach current chat
+.::
{
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        moreOptionsButton := root.FindFirst({ Name: "More chat options", Type: "50000" })

        if moreOptionsButton {
            moreOptionsButton.Click()
            Sleep 350

            detachMenuItem := root.FindFirst({ Name: "Open in new window", Type: "50011" })

            if !detachMenuItem {
                detachMenuItem := UIA.GetRootElement().FindFirst({ Name: "Open in new window", Type: "50011" })
            }

            if detachMenuItem {
                detachMenuItem.Click()
            } else {
                ShowSmallLoadingIndicator_ChatGPT("Could not find Open in new window")
                SetTimer(() => HideSmallLoadingIndicator_ChatGPT(), -2000)
            }
        } else {
            ShowSmallLoadingIndicator_ChatGPT("Could not find more chat options")
            SetTimer(() => HideSmallLoadingIndicator_ChatGPT(), -2000)
        }
    }
    catch Error as e {
        ShowSmallLoadingIndicator_ChatGPT("Could not detach chat")
        SetTimer(() => HideSmallLoadingIndicator_ChatGPT(), -2000)
    }
}

; Shift + W : Call current chat
+w::
{
    ; Show confirmation popup
    if MsgBox("Do you want to call this person?", "Confirm Call", "YesNo Icon?") = "Yes" {
        try {
            win := WinExist("A")
            root := UIA.ElementFromHandle(win)

            callButton := 0
            callButtonNames := ["Audio call", "Video call", "Start audio call", "Start video call"]

            for name in callButtonNames {
                candidates := root.FindAll({ Name: name, Type: "50000", matchmode: "Substring", cs: false })
                if candidates {
                    for candidate in candidates {
                        if !candidate.GetPropertyValue(UIA.Property.IsOffscreen) && candidate.GetPropertyValue(UIA.Property
                            .IsEnabled) {
                            callButton := candidate
                            break
                        }
                    }
                }
                if callButton
                    break
            }

            if !callButton {
                candidates := root.FindAll({ Type: "50000" })
                if candidates {
                    for candidate in candidates {
                        if InStr(StrLower(candidate.Name), "call") && !candidate.GetPropertyValue(UIA.Property.IsOffscreen
                        ) && candidate.GetPropertyValue(UIA.Property.IsEnabled) {
                            callButton := candidate
                            break
                        }
                    }
                }
            }

            if callButton {
                callButton.Click()
            } else {
                ; Show error banner
                ShowSmallLoadingIndicator_ChatGPT("Could not find call button")
                SetTimer(() => HideSmallLoadingIndicator_ChatGPT(), -2000)
            }
        }
        catch Error as e {
            ; Show error banner
            ShowSmallLoadingIndicator_ChatGPT("Could not find call button")
            SetTimer(() => HideSmallLoadingIndicator_ChatGPT(), -2000)
        }
    }
}

; Shift + T : Add participants
+t::
{
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Find the "View and add participants" button using substring matching
        participantsButton := root.FindFirst({ Name: "View and add participants", Type: "50000", matchmode: "Substring" })

        if participantsButton {
            participantsButton.Click()
            Sleep 300
            Send "{Tab}"
            Sleep 300
            Send "{Enter}"
        } else {
            ; Show error banner
            ShowSmallLoadingIndicator_ChatGPT("Could not find add participants button")
            SetTimer(() => HideSmallLoadingIndicator_ChatGPT(), -2000)
        }
    }
    catch Error as e {
        ; Show error banner
        ShowSmallLoadingIndicator_ChatGPT("Could not find add participants button")
        SetTimer(() => HideSmallLoadingIndicator_ChatGPT(), -2000)
    }
}

; Shift + D : Collapse chat categories
+d::
{
    try {
        win := WinExist("A")
        if !win
            return

        root := UIA.ElementFromHandle(win)

        ; Narrow search to the chat navigation tree to speed up lookups.
        treeCond := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Tree)
        trees := ""
        try trees := root.FindElements(treeCond, UIA.TreeScope.Descendants)

        targetTree := ""
        targetIds := ["menur6a5", "menur6as", "menur6br", "menur6f6"]

        if trees {
            for candidate in trees {
                if !candidate
                    continue
                hasSection := false
                for id in targetIds {
                    if !id
                        continue
                    sectionEl := ""
                    try sectionEl := candidate.FindFirst({ AutomationId: id, Type: "50024" })
                    if sectionEl {
                        hasSection := true
                        break
                    }
                }
                if hasSection {
                    targetTree := candidate
                    break
                }
            }
        }

        if !targetTree
            targetTree := root

        ; Collect all expandable tree items (categories and chat groups).
        treeItemCond := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.TreeItem)
        canCollapseCond := UIA.CreatePropertyCondition(UIA.Property.IsExpandCollapsePatternAvailable, true)
        collapsibleCond := UIA.CreateAndCondition(treeItemCond, canCollapseCond)
        items := targetTree.FindElements(collapsibleCond, UIA.TreeScope.Descendants)

        if !items {
            ShowSmallLoadingIndicator_ChatGPT("No collapsible chat sections found")
            SetTimer(() => HideSmallLoadingIndicator_ChatGPT(), -2000)
            return
        }

        collapsed := 0
        already := 0
        total := 0

        for item in items {
            if !item
                continue
            total++
            try {
                pat := item.ExpandCollapsePattern
                if pat.ExpandCollapseState != UIA.ExpandCollapseState.Collapsed {
                    pat.Collapse()
                    collapsed++
                    Sleep 25
                } else {
                    already++
                }
            } catch Error {
                ; Best-effort fallback: focus and send Left to collapse.
                try {
                    item.SetFocus()
                    Sleep 40
                    Send "{Left}"
                    collapsed++
                } catch {
                }
            }
        }

        msg := ""
        if collapsed {
            msg := Format("Collapsed {} chat section{}", collapsed, collapsed = 1 ? "" : "s")
        } else if already && !collapsed {
            msg := "Chat sections already collapsed"
        } else {
            msg := "Nothing to collapse"
        }

        ShowSmallLoadingIndicator_ChatGPT(msg)
        SetTimer(() => HideSmallLoadingIndicator_ChatGPT(), -2000)
    }
    catch Error {
        ShowSmallLoadingIndicator_ChatGPT("Could not collapse chat sections")
        SetTimer(() => HideSmallLoadingIndicator_ChatGPT(), -2000)
    }

    Send "^{Home}"
    Sleep "200"
    Send "c"
    Send "{Right}"
    Sleep "100"
    Send "^{Home}"
    Sleep "200"
    Send "g"
    Send "{Right}"
    Send "^{Home}"
    Sleep "200"
    Send "f"
    Send "f"
    Send "{Right}"
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

; Shift + O : Subject / Title (same as Message window +Y)
+O:: {
    if FocusOutlookField({ AutomationId: "4101" }) ; Subject
        return
    if FocusOutlookField({ Name: "Subject", ControlType: "Edit" })
        return
}

; Shift + P : Required / To (same as Message window +U)
+P:: {
    if FocusOutlookField({ AutomationId: "4109" }) ; Required
        return
    if FocusOutlookField({ Name: "Required", ControlType: "Edit" })
        return
    if FocusOutlookField({ AutomationId: "4117" }) ; To
        return
    if FocusOutlookField({ Name: "To", ControlType: "Edit" })
        return
}

; Shift + M : Subject -> Body (same as Message window +M)
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

; Shift + H : Don't send any response
+H::
{
    ShowSmallLoadingIndicator_ChatGPT("Don't send any response…")
    Send "+{Tab}"
    Sleep 1500
    Send "d"
    Sleep 450
    Send "{Enter}"
    Sleep 500
    HideSmallLoadingIndicator_ChatGPT()
}

; Shift + J : Send response
+J::
{
    ShowSmallLoadingIndicator_ChatGPT("Send response…")
    Send "+{Tab}"
    Sleep 1500
    Send "s"
    Sleep 50
    Send "{Enter}"
    Sleep 500
    HideSmallLoadingIndicator_ChatGPT()
}

; -------------------------------------------------------------------
; Focus helpers â€" reuse for any field you need
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
; Click helper â€" try AutomationId first, then Name+ClassName
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
; General helper â€" visually confirm focus on the selected element
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

    ; Robust fallback â€" cycle through panes up to 6 times to reach navigation, then Home
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
    ; Last resort â€" send Home anyway
    Send "{Home}"
    EnsureFocus()
    return false
}

; Shift + K : Send Shift+F6
+K:: Send "+{F6}"

; Shift + L : Send F6
+L:: Send "{F6}"

; Message inspector-specific hotkeys (Subject/To/DatePicker/Body)
#HotIf IsOutlookMessageActive()

; Shift + Y > Subject
+Y:: {
    if FocusOutlookField({ AutomationId: "4101" }) ; Subject
        return
    if FocusOutlookField({ Name: "Subject", ControlType: "Edit" })
        return
}

; Shift + U > Required / To
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

; Shift + I > Date Picker (if present in this inspector)
; (No Shift + I in Message inspector)

; Shift + M > Subject > Body
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

; New Shift hotkeys for date/time controls (Start > End), ordered by key preference
; Shift + Y > Start date (combo)
+Y:: {
    Outlook_ClickStartDate()
}

; Shift + U â†' Start date â€" Date Picker
+U:: {
    Outlook_ClickStartDatePicker()
}

; Shift + I â†' Start time (combo)
+I:: {
    Outlook_ClickStartTime()
}

; Shift + O â†' End date (combo)
+O:: {
    Outlook_ClickEndDate()
}

; Shift + P â†' End time (combo)
+P:: {
    Outlook_ClickEndTime()
}

; Shift + H â†' All day checkbox
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

; Shift + J â†' Title field
+J:: {
    if FocusOutlookField({ AutomationId: "4100" }) ; Title
        return
    if FocusOutlookField({ Name: "Title", ControlType: "Edit" })
        return
}

; Shift + L â†' Required / To field
+L:: {
    if FocusOutlookField({ AutomationId: "4109" }) ; Required
        return
    if FocusOutlookField({ Name: "Required", ControlType: "Edit" })
        return
}

; Shift + M â†' Location â†' Body
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

; Shift + ; â†' Make Recurring (click)
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

; Shift + U â†' click ChatGPT's model selector (any language)
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
    ; --- Button Names (EN/PT) ---
    pt_stopStreamingName := "Interromper transmissão"
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

; Explorer-specific helper â€" select first pinned item in the sidebar
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

    ; Last-chance fallback â€" press Home which works if focus is already inside the list
    Send "{Home}"
    EnsureFocus()
}

; Helper to force focus to the ItemsView pane (file list)
EnsureItemsViewFocus() {
    try {
        explorerHwnd := WinExist("A")
        root := UIA.ElementFromHandle(explorerHwnd)

        ; quick check â€" if ItemsView already has keyboard focus, we're done
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
; Excel Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe EXCEL.EXE")

; Shift + Y : Select White Color (Up-Arrow, Ctrl-Home, Ctrl-Home)
+y:: {
    Send "^{PgUp}"
}

; Shift + U : Click Enable Editing button
+u:: {
    try {
        root := UIA.ElementFromHandle(WinExist("A"))
        if (btn := WaitForButton(root, "Enable Editing", 3000)) {
            btn.Invoke()
        } else {
            MsgBox "Couldn't find the Enable Editing button."
        }
    } catch Error as err {
        MsgBox "Error:`n" err.Message
    }
}

; Shift + I : Turn CSV delimited by semicolon into columns (Alt, 0, 5, D, Enter, M, Enter, Enter)
+i:: {
    Send "{Alt}"
    Sleep 100
    Send "0"
    Sleep 100
    Send "5"
    Sleep 100
    Send "d"
    Sleep 100
    Send "{Enter}"
    Sleep 100
    if MsgBox("If 'semicolon' is not selected, hit yes", "Confirm", "YesNo Icon?") = "Yes" {
        Send "m"
        Sleep 100
    }
    Send "{Enter}"
    Sleep 100
    Send "{Enter}"
}

#HotIf

;-------------------------------------------------------------------
; Power BI Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe PBIDesktop.exe") || InStr(WinGetTitle("A"), "powerbi", false)

; Shift + Y : Transform data (Alt, H, T, then UIA click)
+y:: {
    try {
        Send "{Alt down}"
        Send "{Alt up}"
        Sleep 80
        Send "h"
        Sleep 120
        Send "t"
        Sleep 250

        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        possibleNames := ["Transform data", "Transformar dados"]
        transformBtn := ""

        for , name in possibleNames {
            transformBtn := root.FindFirst({ Name: name, Type: "50011" })
            if transformBtn
                break
        }

        if !transformBtn {
            transformBtn := root.FindFirst({ Name: "Transform", Type: "50011", matchmode: "Substring" })
        }

        if transformBtn {
            transformBtn.Click()
        } else {
            MsgBox "Could not find the 'Transform data' menu item.", "Power BI", "IconX"
        }
    } catch Error as e {
        MsgBox "Error triggering Transform data: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + U : Close and apply (Alt, H, C, C)
+u:: {
    Send "{Esc}"
    Send "{Esc}"
    Send "{Alt down}"
    Send "{Alt down}"
    Sleep 200
    Sleep 200
    Send "{Alt up}"
    Send "h"
    Sleep 100
    Send "c"
    Sleep 100
    Send "c"
}

; Shift + I : Report view
+i:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Find the Report view tab by name only
        reportTab := root.FindFirst({ Name: "Report view" })
        if !reportTab {
            reportTab := root.FindFirst({ Name: "Report view", matchmode: "Substring" })
        }

        if reportTab {
            reportTab.Click()
        } else {
            MsgBox "Could not find the 'Report view' tab.", "Power BI", "IconX"
        }
    } catch Error as e {
        MsgBox "Error switching to Report view: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + O : Table view
+o:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Find the Table view tab by name only
        tableTab := root.FindFirst({ Name: "Table view" })
        if !tableTab {
            tableTab := root.FindFirst({ Name: "Table view", matchmode: "Substring" })
        }

        if tableTab {
            tableTab.Click()
        } else {
            MsgBox "Could not find the 'Table view' tab.", "Power BI", "IconX"
        }
    } catch Error as e {
        MsgBox "Error switching to Table view: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + P : Model view
+p:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Find the Model view tab by name only
        modelTab := root.FindFirst({ Name: "Model view" })
        if !modelTab {
            modelTab := root.FindFirst({ Name: "Model view", matchmode: "Substring" })
        }

        if modelTab {
            modelTab.Click()
        } else {
            MsgBox "Could not find the 'Model view' tab.", "Power BI", "IconX"
        }
    } catch Error as e {
        MsgBox "Error switching to Model view: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + H : Analytics
+h:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Find the Analytics tab by name only
        analyticsTab := root.FindFirst({ Name: "Analytics" })
        if !analyticsTab {
            analyticsTab := root.FindFirst({ Name: "Analytics", matchmode: "Substring" })
        }

        if analyticsTab {
            analyticsTab.Click()
        } else {
            MsgBox "Could not find the 'Analytics' tab.", "Power BI", "IconX"
        }
    } catch Error as e {
        MsgBox "Error switching to Analytics: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + J : Format visual
+j:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Find the Format visual tab by name only
        formatTab := root.FindFirst({ Name: "Format visual" })
        if !formatTab {
            formatTab := root.FindFirst({ Name: "Format visual", matchmode: "Substring" })
        }

        if formatTab {
            formatTab.Click()
        } else {
            MsgBox "Could not find the 'Format visual' tab.", "Power BI", "IconX"
        }
    } catch Error as e {
        MsgBox "Error switching to Format visual: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + K : Build visual
+k:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Find the Build visual tab by name only
        buildTab := root.FindFirst({ Name: "Build visual" })
        if !buildTab {
            buildTab := root.FindFirst({ Name: "Build visual", matchmode: "Substring" })
        }

        if buildTab {
            buildTab.Click()
        } else {
            MsgBox "Could not find the 'Build visual' tab.", "Power BI", "IconX"
        }
    } catch Error as e {
        MsgBox "Error switching to Build visual: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + L : Click OK/Confirm button in Power BI modals
+l:: {
    try {
        root := UIA.ElementFromHandle(WinExist("A"))

        ; Try by name first (since we know it's "OK" with Type 50000)
        possibleNames := [
            ; English variations
            "OK",
            "Confirm",
            "Accept",
            "Apply",
            "Done"
            "Yes",
            "Continue",
            "Proceed",
            "Save",
            "Finish",
            ; Portuguese variations
            "Confirmar",
            "Aceitar",
            "Aplicar",
            "Sim",
            "Continuar",
            "Prosseguir",
            "Salvar",
            "Finalizar",
            ; Spanish variations
            "Aceptar",
            "Continuar",
            "Guardar",
            "Finalizar",
            ; French variations
            "Confirmer",
            "Accepter",
            "Continuer",
            "Enregistrer",
            ; German variations
            "Bestätigen",
            "Akzeptieren",
            "Fortfahren",
            "Speichern"
        ]

        ; First attempt: Try by name with Button type (numeric 50000 or string "Button")
        for name in possibleNames {
            confirmBtn := root.FindFirst({ Type: "Button", Name: name })
            if !confirmBtn {
                ; Try with numeric type code
                confirmBtn := root.FindFirst({ Type: 50000, Name: name })
            }
            if confirmBtn
                break
        }

        ; Second attempt: Find by AutomationId and Type
        if !confirmBtn {
            confirmBtn := root.FindFirst({ Type: "Button", AutomationId: "1" })
            if !confirmBtn {
                confirmBtn := root.FindFirst({ Type: 50000, AutomationId: "1" })
            }
        }

        ; Third attempt: Try SplitButton type (some dialogs use this instead)
        if !confirmBtn {
            for name in possibleNames {
                confirmBtn := root.FindFirst({ Type: "SplitButton", Name: name })
                if confirmBtn
                    break
            }
            if !confirmBtn {
                confirmBtn := root.FindFirst({ Type: "SplitButton", AutomationId: "1" })
            }
        }

        ; Fourth attempt: Search all buttons and find by name (more thorough)
        if !confirmBtn {
            allButtons := root.FindAll({ Type: "Button" })
            for btn in allButtons {
                btnName := btn.Name
                for name in possibleNames {
                    if (btnName = name) {
                        confirmBtn := btn
                        break
                    }
                }
                if confirmBtn
                    break
            }
        }

        if confirmBtn {
            confirmBtn.Click()
            return
        }
    } catch Error {
    }
    ; Fallback: Try common keyboard shortcuts
    Send "{Enter}"  ; Enter key is universal for OK/Confirm
}

; Shift + N : Click Cancel/Exit button in Power BI modals
+n:: {
    try {
        root := UIA.ElementFromHandle(WinExist("A"))

        ; First attempt: Find by AutomationId and Type (most reliable)
        cancelBtn := root.FindFirst({ Type: "Button", AutomationId: "2" })

        ; Second attempt: Try various possible names for Cancel/Exit
        if !cancelBtn {
            possibleNames := [
                ; English variations
                "Cancel",
                "Close",
                "Exit",
                "Dismiss",
                "No",
                "Abort",
                "Back",
                "Close",
                ; Portuguese variations
                "Cancelar",
                "Fechar",
                "Sair",
                "Descartar",
                "Não",
                "Voltar",
                ; Spanish variations
                "Cancelar",
                "Cerrar",
                "Salir",
                "Descartar",
                "No",
                ; French variations
                "Annuler",
                "Fermer",
                "Quitter",
                "Ignorer",
                "Non",
                ; German variations
                "Abbrechen",
                "Schließen",
                "Verlassen",
                "Abweisen",
                "Nein"
            ]
            for name in possibleNames {
                cancelBtn := root.FindFirst({ Type: "Button", Name: name })
                if cancelBtn
                    break
            }
        }

        ; Third attempt: Try SplitButton type (some dialogs use this instead)
        if !cancelBtn {
            cancelBtn := root.FindFirst({ Type: "SplitButton", AutomationId: "2" })
            if !cancelBtn {
                for name in possibleNames {
                    cancelBtn := root.FindFirst({ Type: "SplitButton", Name: name })
                    if cancelBtn
                        break
                }
            }
        }

        if cancelBtn {
            cancelBtn.Click()
            return
        }
    } catch Error {
    }
    ; Fallback: Try common keyboard shortcuts
    Send "{Esc}"  ; Escape key is universal for cancels
}

; Shift + M : Right-click Previous pages button in Power BI
+m:: {
    try {
        root := UIA.ElementFromHandle(WinExist("A"))

        ; First attempt: Find by Name and Type
        prevPageBtn := root.FindFirst({ Type: "Button", Name: "Previous pages" })
        if !prevPageBtn {
            prevPageBtn := root.FindFirst({ Type: 50000, Name: "Previous pages" })
        }

        ; Second attempt: Find by ClassName
        if !prevPageBtn {
            prevPageBtn := root.FindFirst({ Type: "Button", ClassName: "carouselNavButton previousPage" })
            if !prevPageBtn {
                prevPageBtn := root.FindFirst({ Type: 50000, ClassName: "carouselNavButton previousPage" })
            }
        }

        ; Third attempt: Find by partial ClassName match
        if !prevPageBtn {
            allButtons := root.FindAll({ Type: "Button" })
            for btn in allButtons {
                btnClassName := btn.ClassName
                if InStr(btnClassName, "previousPage") {
                    prevPageBtn := btn
                    break
                }
            }
        }

        if prevPageBtn {
            ; Get button location and instantly move cursor to that position
            btnPos := prevPageBtn.Location
            x := btnPos.x + btnPos.w // 2
            y := btnPos.y + btnPos.h // 2

            ; Instantly set cursor position (no visible movement)
            DllCall("SetCursorPos", "Int", x, "Int", y)

            ; Perform right-click immediately
            saveCoordMode := A_CoordModeMouse
            CoordMode("Mouse", "Screen")
            Click(x " " y " Right")
            CoordMode("Mouse", saveCoordMode)
            return
        }
    } catch Error {
    }
}

; Shift + , : Click Filter pane collapse/expand button
+,:: {
    try {
        root := UIA.ElementFromHandle(WinExist("A"))

        ; Find by Name (full name)
        filterBtn := root.FindFirst({ Type: "Button", Name: "Collapse or expand the filter pane while editing. This also determines how report readers see it" })
        if !filterBtn {
            filterBtn := root.FindFirst({ Type: 50000, Name: "Collapse or expand the filter pane while editing. This also determines how report readers see it" })
        }

        ; Find by ClassName
        if !filterBtn {
            filterBtn := root.FindFirst({ Type: "Button", ClassName: "btn collapseIcon pbi-borderless-button glyphicon glyph-mini pbi-glyph-doublechevronleft" })
            if !filterBtn {
                filterBtn := root.FindFirst({ Type: 50000, ClassName: "btn collapseIcon pbi-borderless-button glyphicon glyph-mini pbi-glyph-doublechevronleft" })
            }
        }

        ; Find by partial ClassName match
        if !filterBtn {
            allButtons := root.FindAll({ Type: "Button" })
            for btn in allButtons {
                btnClassName := btn.ClassName
                if InStr(btnClassName, "pbi-glyph-doublechevronleft") {
                    filterBtn := btn
                    break
                }
            }
        }

        if filterBtn {
            filterBtn.Click()
            return
        }
    } catch Error {
    }
}

; Shift + . : Click Visualizations button
+.:: {
    try {
        root := UIA.ElementFromHandle(WinExist("A"))

        ; Find by Name
        vizBtn := root.FindFirst({ Type: "Button", Name: "Visualizations" })
        if !vizBtn {
            vizBtn := root.FindFirst({ Type: 50000, Name: "Visualizations" })
        }

        ; Find by ClassName
        if !vizBtn {
            vizBtn := root.FindFirst({ Type: "Button", ClassName: "toggle-button ng-star-inserted" })
            if !vizBtn {
                allButtons := root.FindAll({ Type: "Button" })
                for btn in allButtons {
                    if (btn.Name = "Visualizations" && InStr(btn.ClassName, "toggle-button")) {
                        vizBtn := btn
                        break
                    }
                }
            }
        }

        if vizBtn {
            vizBtn.Click()
            return
        }
    } catch Error {
    }
}

; Shift + W : Click Data button
+q:: {
    try {
        root := UIA.ElementFromHandle(WinExist("A"))

        ; Find by Name
        dataBtn := root.FindFirst({ Type: "Button", Name: "Data" })
        if !dataBtn {
            dataBtn := root.FindFirst({ Type: 50000, Name: "Data" })
        }

        ; Find by ClassName
        if !dataBtn {
            dataBtn := root.FindFirst({ Type: "Button", ClassName: "toggle-button" })
            if !dataBtn {
                allButtons := root.FindAll({ Type: "Button" })
                for btn in allButtons {
                    if (btn.Name = "Data" && InStr(btn.ClassName, "toggle-button")) {
                        dataBtn := btn
                        break
                    }
                }
            }
        }

        if dataBtn {
            dataBtn.Click()
            return
        }
    } catch Error {
    }
}

; Shift + Q : Click New page button
+w:: {
    try {
        root := UIA.ElementFromHandle(WinExist("A"))

        ; Find by Name
        newPageBtn := root.FindFirst({ Type: "Button", Name: "New page" })
        if !newPageBtn {
            newPageBtn := root.FindFirst({ Type: 50000, Name: "New page" })
        }

        ; Find by ClassName
        if !newPageBtn {
            newPageBtn := root.FindFirst({ Type: "Button", ClassName: "section static create" })
            if !newPageBtn {
                newPageBtn := root.FindFirst({ Type: 50000, ClassName: "section static create" })
            }
        }

        ; Find by partial ClassName match
        if !newPageBtn {
            allButtons := root.FindAll({ Type: "Button" })
            for btn in allButtons {
                if (btn.Name = "New page") {
                    newPageBtn := btn
                    break
                }
            }
        }

        if newPageBtn {
            newPageBtn.Click()
            return
        }
    } catch Error {
    }
}

; Shift + E : New measure (Alt, H, N, M)
+e:: {
    try {
        Send "{Escape}"
        Sleep 100
        Send "{Alt down}"
        Sleep 80
        Send "{Alt up}"
        Sleep 120
        Send "h"
        Sleep 120
        Send "n"
        Sleep 120
        Send "m"
    } catch Error as e {
        MsgBox "Error triggering New measure: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + R : Collapse Power BI table tree items
+r:: {
    try {
        win := WinExist("A")
        if !win
            return

        root := UIA.ElementFromHandle(win)

        treeItemCond := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.TreeItem)
        expandCond := UIA.CreatePropertyCondition(UIA.Property.IsExpandCollapsePatternAvailable, true)
        tableCond := UIA.CreatePropertyConditionEx(UIA.Property.Name, "Table ", UIA.PropertyConditionFlags.IgnoreCaseMatchSubstring
        )
        calcTableCond := UIA.CreatePropertyConditionEx(UIA.Property.Name, "Calculated Table", UIA.PropertyConditionFlags
            .IgnoreCaseMatchSubstring)
        nameCond := UIA.CreateOrCondition(tableCond, calcTableCond)
        targetCond := UIA.CreateAndCondition(treeItemCond, UIA.CreateAndCondition(expandCond, nameCond))

        items := ""
        try items := root.FindElements(targetCond, UIA.TreeScope.Descendants)

        if !items {
            MsgBox "Could not find any Power BI tables to collapse.", "Power BI", "IconX"
            return
        }

        collapsed := 0
        already := 0

        for item in items {
            if !item
                continue
            try {
                pat := item.ExpandCollapsePattern
                if pat.ExpandCollapseState != UIA.ExpandCollapseState.Collapsed {
                    pat.Collapse()
                    collapsed++
                    Sleep 35
                } else {
                    already++
                }
            } catch Error {
                try {
                    item.SetFocus()
                    Sleep 40
                    Send "{Left}"
                    collapsed++
                } catch {
                }
            }
        }

        if collapsed {
            ToolTip Format("Collapsed {} table{}", collapsed, collapsed = 1 ? "" : "s")
        } else if already {
            ToolTip "All tables already collapsed"
        } else {
            ToolTip "No tables collapsed"
        }
        SetTimer(() => ToolTip(), -1200)
    } catch Error as e {
        MsgBox "Error collapsing Power BI tables: " e.Message, "Power BI Error", "IconX"
    }
}

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

; Shift + F: Click inbox button
+f::
{
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 300 ; Give UIA time to attach

        ; Try to find inbox link by name (may include unread count)
        inboxLink := uia.FindElement({ Name: "Inbox", Type: "50005", matchmode: "Substring" })

        ; Fallback: try by ClassName
        if (!inboxLink) {
            inboxLink := uia.FindElement({ ClassName: "J-Ke n0", Type: "50005" })
        }

        ; Fallback: try by Value (URL)
        if (!inboxLink) {
            inboxLink := uia.FindElement({ Value: "#inbox", Type: "50005", matchmode: "Substring" })
        }

        if (inboxLink) {
            inboxLink.Click()
        }
        else {
            MsgBox "Could not find the 'Inbox' button."
        }
    }
    catch Error as e {
        MsgBox "An error occurred: " e.Message
    }
}

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

; Ctrl + 1 : Remove clustering and focus on the code
^1::
{
    ; Send ESC two times
    Send "{Escape}"  ; ESC
    Sleep 50
    Send "{Escape}"  ; ESC again
    Sleep 100
    Send "^!n"
    Sleep 100
    Send "^!,"
    Sleep 100
    Send "#!o"
}

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

; Shift + P : Go to terminal
+p:: Send "^'"

; Shift + H : New terminal
+h:: Send '^+"'

; Shift + J : Go to file explorer
+j:: Send "^+e"

+k::
{
    Send "+i"
    Sleep 700
    Send "+o"
    Sleep 500
    WinMaximize "A"
}

; Shift + L : command palette
+l:: Send "^+p"

; Shift + , : Show chat history
+,::
{
    Send "^+p" ; Open command palette
    Sleep 200
    Send "show history"
    Sleep 200
    Send "{Enter}"
}

; Shift + E : Search
+e:: Send "^+f"

; Ctrl + M : Ask, wait banner 8s, then Shift+V
^M::
{
    ; Remember current target window so later keystrokes go to the right app
    gCommitPushTargetWin := WinExist("A")
    ; Prompt push decision upfront (blocking, topmost). Store for later execution.
    PromptCommitPushDecisionBlocking()

    ; Reactivate the target window after dialog closes (dialog steals focus)
    if (gCommitPushTargetWin) {
        WinActivate gCommitPushTargetWin
        WinWaitActive("ahk_id " gCommitPushTargetWin, , 1)
        Sleep 200
    }

    Send "+d"
    Sleep 200
    Send "^!a"
    Sleep 200
    loop 8 {
        secondsLeft := 9 - A_Index
        ShowSmallLoadingIndicator_ChatGPT("Waiting " . secondsLeft . "s…")

        ; Check if the message text is present - ultra simple approach
        elementFound := false

        try {
            if WinExist("A") {
                ; Just search for the most distinctive part
                windowText := WinGetText("A")

                ; Look for the unique shortcut text
                if InStr(windowText, "Ctrl+⏎") {
                    elementFound := true
                    MsgBox "Found Ctrl+⏎ - element present"
                } else if InStr(windowText, "commit on") {
                    elementFound := true
                    MsgBox "Found 'commit on' - element present"
                } else if InStr(windowText, "Message") {
                    elementFound := true
                    MsgBox "Found 'Message' - element present"
                }
            }
        }

        ; Simple logic: if element found, continue; if not found, exit early
        if (elementFound) {
            ShowSmallLoadingIndicator_ChatGPT("Element present, continuing...")
            Sleep 1500
            continue
        } else {
            ; Element not found, exit early
            ShowSmallLoadingIndicator_ChatGPT("Element not found, stopping...")
            Sleep 500
            Sleep 2500
            ; Ensure target window is active before sending commit command
            if (gCommitPushTargetWin) {
                WinActivate gCommitPushTargetWin
                Sleep 200
            }
            Send "^!,"
            Sleep 100
            Send "+v"
            HideSmallLoadingIndicator_ChatGPT()

            ; Execute stored decision (if any) after commit is sent
            ExecuteStoredCommitPushDecision()
            return
        }

        Sleep 1000
    }

    ; If we reach here, the loop completed normally (element was found)
    ; Send the commit and show push selector popup
    ; Ensure target window is active before sending commit command
    if (gCommitPushTargetWin) {
        WinActivate gCommitPushTargetWin
        Sleep 200
    }
    Send "^!,"
    Sleep 100
    Send "+v"
    HideSmallLoadingIndicator_ChatGPT()
    ExecuteStoredCommitPushDecision()
}

; Global variable for commit push selector target window
global gCommitPushTargetWin := 0
; Global variable to store the user's push decision ("push" | "dont_push" | "")
global gCommitPushDecision := ""

; Function to get commit push action by number
GetCommitPushActionByNumber(numberText) {
    try number := Integer(numberText)
    catch {
        return ""
    }
    actionMap := Map()
    actionMap[1] := "push"
    actionMap[2] := "dont_push"
    return (actionMap.Has(number)) ? actionMap[number] : ""
}

; Record-only auto-submit handler for the upfront decision prompt
CommitPushDecision_AutoSubmit(ctrl, *) {
    global gCommitPushDecision
    currentValue := ctrl.Text
    if (currentValue != "" && IsInteger(currentValue)) {
        action := GetCommitPushActionByNumber(currentValue)
        if (action != "") {
            gCommitPushDecision := action
            ctrl.Gui.Destroy()
        }
    }
}

; Blocking, topmost prompt to capture push decision upfront
PromptCommitPushDecisionBlocking() {
    global gCommitPushDecision
    try {
        gCommitPushDecision := ""
        decisionGui := Gui("+AlwaysOnTop +ToolWindow", "Commit Push Selector")
        decisionGui.SetFont("s10", "Segoe UI")
        decisionGui.AddText("w350 Center"
            , "Push after commit?`n`n1. Push (Shift+B)`n2. Don't push`n`nType a number (1-2):")
        decisionGui.AddEdit("w50 Center vCommitPushInput", "")
        decisionGui["CommitPushInput"].OnEvent("Change", CommitPushDecision_AutoSubmit)
        decisionGui.AddButton("w80", "Cancel").OnEvent("Click", (*) => decisionGui.Destroy())
        decisionGui.Show("w350 h150")
        decisionGui["CommitPushInput"].Focus()
        WinWaitClose("ahk_id " decisionGui.Hwnd)
    } catch Error as e {
        MsgBox "Error in upfront push decision: " e.Message, "Commit Push Selector Error", "IconX"
    }
}

; Execute stored decision at the exact current push moment
ExecuteStoredCommitPushDecision() {
    global gCommitPushDecision
    global gCommitPushTargetWin
    if (gCommitPushDecision = "push") {
        ; Wait a moment for Cursor to process the commit
        Sleep 500
        ; Ensure the intended window has focus before sending the push hotkey
        if (gCommitPushTargetWin) {
            WinActivate gCommitPushTargetWin
            WinWaitActive("ahk_id " gCommitPushTargetWin, , 2)
            Sleep 200
        }
        Send "+b"
    }
    ; Clear after execution to avoid reusing stale decisions
    gCommitPushDecision := ""
}

; Function to execute commit push action
ExecuteCommitPushAction(action) {
    if (action = "")
        return

    if (action = "push") {
        ; Option 1: Push (send Shift+B)
        Send "+b"
    } else if (action = "dont_push") {
        ; Option 2: Don't push (do nothing)
        ; Just close the popup, no action needed
    }
}

; Auto-submit function for commit push selector
AutoSubmitCommitPush(ctrl, *) {
    currentValue := ctrl.Text
    if (currentValue != "" && IsInteger(currentValue)) {
        action := GetCommitPushActionByNumber(currentValue)
        if (action != "") {
            ctrl.Gui.Destroy()
            ExecuteCommitPushAction(action)
        }
    }
}

; Manual submit function for commit push selector (backup)
SubmitCommitPush(ctrl, *) {
    currentValue := ctrl.Gui["CommitPushInput"].Text
    if (currentValue != "" && IsInteger(currentValue)) {
        action := GetCommitPushActionByNumber(currentValue)
        if (action != "") {
            ctrl.Gui.Destroy()
            ExecuteCommitPushAction(action)
        } else {
            MsgBox "Invalid selection. Please choose 1-2.", "Commit Push Selector", "IconX"
        }
    }
}

; Cancel function for commit push selector
CancelCommitPush(ctrl, *) {
    ctrl.Gui.Destroy()
}

; Function to show commit push selector popup
ShowCommitPushSelector() {
    try {
        ; Remember current target window before showing GUI
        gCommitPushTargetWin := WinExist("A")
        ; Create GUI for commit push selection with auto-submit
        commitPushGui := Gui("+AlwaysOnTop +ToolWindow", "Commit Push Selector")
        commitPushGui.SetFont("s10", "Segoe UI")

        ; Add instruction text
        commitPushGui.AddText("w350 Center",
            "Commit sent! Choose next action:`n`n1. Push (Shift+B)`n2. Don't push`n`nType a number (1-2):")

        ; Add input field with auto-submit
        commitPushGui.AddEdit("w50 Center vCommitPushInput", "")
        commitPushGui["CommitPushInput"].OnEvent("Change", AutoSubmitCommitPush)

        ; Add manual submit button (backup)
        commitPushGui.AddButton("w80", "Submit").OnEvent("Click", SubmitCommitPush)

        ; Add cancel button
        commitPushGui.AddButton("w80", "Cancel").OnEvent("Click", CancelCommitPush)

        ; Show GUI and focus input
        commitPushGui.Show("w350 h150")
        commitPushGui["CommitPushInput"].Focus()

    } catch Error as e {
        MsgBox "Error in commit push selector: " e.Message, "Commit Push Selector Error", "IconX"
    }
}

; Auto-submit function - triggers when text changes
global gEmojiTargetWin := 0

GetEmojiByNumber(numberText) {
    try number := Integer(numberText)
    catch {
        return ""
    }
    emojiMap := Map()
    emojiMap[1] := "🔲"
    emojiMap[2] := "⏳"
    emojiMap[3] := "⚡"
    emojiMap[4] := "2️⃣"
    emojiMap[5] := "❓"
    return (emojiMap.Has(number)) ? emojiMap[number] : ""
}

InsertEmojiToTarget(emoji) {
    if (emoji = "")
        return
    ; Activate the target window if we have it stored
    if (gEmojiTargetWin) {
        WinActivate gEmojiTargetWin
        Sleep 150
    }

    ; Use direct text insertion - no clipboard manipulation
    ; This is more reliable and won't interfere with user's clipboard
    SendText(emoji)
}

AutoSubmitEmoji(ctrl, *) {
    currentValue := ctrl.Text
    if (currentValue != "" && IsInteger(currentValue)) {
        emoji := GetEmojiByNumber(currentValue)
        if (emoji != "") {
            ctrl.Gui.Destroy()
            InsertEmojiToTarget(emoji)
        }
    }
}

; Manual submit function (backup)
SubmitEmoji(ctrl, *) {
    currentValue := ctrl.Gui["EmojiInput"].Text
    if (currentValue != "" && IsInteger(currentValue)) {
        emoji := GetEmojiByNumber(currentValue)
        if (emoji != "") {
            ctrl.Gui.Destroy()
            InsertEmojiToTarget(emoji)
        } else {
            MsgBox "Invalid selection. Please choose 1-5.", "Emoji Selector", "IconX"
        }
    }
}

; Cancel function
CancelEmoji(ctrl, *) {
    ctrl.Gui.Destroy()
}

; Shift + T : Emoji selector (Auto-submit version)
+t::
{
    try {
        ; Remember current target window before showing GUI
        gEmojiTargetWin := WinExist("A")
        ; Create GUI for emoji selection with auto-submit
        emojiGui := Gui("+AlwaysOnTop +ToolWindow", "Emoji Selector")
        emojiGui.SetFont("s10", "Segoe UI")

        ; Add instruction text
        emojiGui.AddText("w350 Center",
            "Select emoji to insert:`n`n1. 🔲 Tasks/Checklist items`n2. ⏳ Time-sensitive tasks`n3. ⚡ First priority`n4. 2️⃣ Second priority`n5. ❓ Questions/Uncertain items`n`nType a number (1-5):"
        )

        ; Add input field with auto-submit functionality
        emojiGui.AddEdit("w50 Center vEmojiInput Limit1 Number")

        ; Add OK and Cancel buttons (as backup)
        emojiGui.AddButton("w80 xp-40 y+10", "OK").OnEvent("Click", SubmitEmoji)
        emojiGui.AddButton("w80 xp+90", "Cancel").OnEvent("Click", CancelEmoji)

        ; Set up auto-submit on text change
        emojiGui["EmojiInput"].OnEvent("Change", AutoSubmitEmoji)

        ; Show GUI and focus input
        emojiGui.Show("w350 h200")
        emojiGui["EmojiInput"].Focus()

    } catch Error as e {
        MsgBox "Error in emoji selector: " e.Message, "Emoji Selector Error", "IconX"
    }
}

; Global variables for AI model selector
global gAIModelTargetWin := 0

; AI Model auto-submit function
AutoSubmitAIModel(ctrl, *) {
    currentValue := ctrl.Text
    if (currentValue != "" && IsInteger(currentValue)) {
        choice := Integer(currentValue)
        if (choice >= 1 && choice <= 6) {
            ctrl.Gui.Destroy()
            ExecuteAIModelSelection(choice)
        }
    }
}

; Manual submit function for AI model (backup)
SubmitAIModel(ctrl, *) {
    currentValue := ctrl.Gui["AIModelInput"].Text
    if (currentValue != "" && IsInteger(currentValue)) {
        choice := Integer(currentValue)
        if (choice >= 1 && choice <= 6) {
            ctrl.Gui.Destroy()
            ExecuteAIModelSelection(choice)
        } else {
            MsgBox "Invalid selection. Please choose 1-6.", "AI Model Selection", "IconX"
        }
    }
}

; Cancel function for AI model
CancelAIModel(ctrl, *) {
    ctrl.Gui.Destroy()
}

; Execute the AI model selection logic
ExecuteAIModelSelection(choice) {
    try {
        ; Send Escape twice, then select the edit field based on on-screen Agent/Ask
        Send "{Escape 2}"
        Sleep 200
        if !SendCtrlKeyBasedOnAgentAsk() {
            ; Fallback to Ctrl+I if no relevant text is found
            Send "{Ctrl down}i{Ctrl up}"
        }
        Sleep 300

        ; Handle different behaviors based on choice
        switch choice {
            case 1:
            {
                ; For auto option: simulate ;, wait for model context menu, then send ↓, Enter
                Send "^;"
                Sleep 300
                SendText "auto"
                Sleep 500
                Send "{Enter}"
                Sleep 300
                Send "{Escape}"
            }
            case 2:
            {
                ; For other options: simulate Ctrl + ., wait, type model string, no Enter
                Send "^;"
                Sleep 500
                SendText "CLAUD"
            }
            case 3:
            {
                Send "^;"
                Sleep 500
                SendText "GPT"
            }
            case 4:
            {
                Send "^;"
                Sleep 500
                SendText "O"
            }
            case 5:
            {
                Send "^;"
                Sleep 500
                SendText "DeepSeek"
            }
            case 6:
            {
                Send "^;"
                Sleep 500
                SendText "Cursor"
            }
        }

        Sleep 100

    } catch Error as e {
        MsgBox "Error in AI model selection: " e.Message, "AI Model Selection Error", "IconX"
    }
}

; ; Shift + G : Switch between AI models (Auto-submit version)
; +g::
; {
;     try {
;         ; Remember current target window before showing GUI
;         gAIModelTargetWin := WinExist("A")
;         ; Create GUI for AI model selection with auto-submit
;         aiModelGui := Gui("+AlwaysOnTop +ToolWindow", "AI Model Selection")
;         aiModelGui.SetFont("s10", "Segoe UI")

;         ; Add instruction text
;         aiModelGui.AddText("w350 Center",
;             "Choose AI Model:`n`n1. auto`n2. CLAUD`n3. GPT`n4. O`n5. DeepSeek`n6. Cursor`n`nType a number (1-6):")

;         ; Add input field with auto-submit functionality
;         aiModelGui.AddEdit("w50 Center vAIModelInput Limit1 Number")

;         ; Add OK and Cancel buttons (as backup)
;         aiModelGui.AddButton("w80 xp-40 y+10", "OK").OnEvent("Click", SubmitAIModel)
;         aiModelGui.AddButton("w80 xp+90", "Cancel").OnEvent("Click", CancelAIModel)

;         ; Set up auto-submit on text change
;         aiModelGui["AIModelInput"].OnEvent("Change", AutoSubmitAIModel)

;         ; Show GUI and focus input
;         aiModelGui.Show("w350 h200")
;         aiModelGui["AIModelInput"].Focus()

;     } catch Error as e {
;         MsgBox "Error in AI model selector: " e.Message, "AI Model Selector Error", "IconX"
;     }
; }

+g::^;

; Shift + V : Fold all Git directories in Source Control (Cursor)
+w:: FoldAllGitDirectoriesInCursor()

; Global variable for commit selector target window
global gCommitTargetWin := 0

; Function to get commit action by number
GetCommitActionByNumber(numberText) {
    try number := Integer(numberText)
    catch {
        return ""
    }
    actionMap := Map()
    actionMap[1] := "workspace"
    actionMap[2] := "repository"
    return (actionMap.Has(number)) ? actionMap[number] : ""
}

; Function to execute commit action
ExecuteCommitAction(action) {
    if (action = "")
        return

    if (action = "workspace") {
        ; Option 1: Commit and push from workspace (original behavior)
        Send "{Right}"
        Send "{Down}"
        Send "{Tab 2}"
        Send "{Enter}"
        Sleep 1500
        Send "{Tab 2}"
        Send "{Enter}"
        Send "{Up 2}"
    }
    else if (action = "repository") {
        ; Option 2: Commit and push from repository (customize this section)

        ; Then execute the commit commands
        Send "^+g"
        Sleep 150

        ; Click on the "Generate Commit Message (Ctrl+M)" button
        ClickGenerateCommitMessageButton()

        Send "{Tab 3}"
        Send "{Enter}"
        Send "{Tab 2}"
        Send "{Enter}"
        Send "{Up 2}"
    }
}

; Auto-submit function for commit selector
AutoSubmitCommit(ctrl, *) {
    currentValue := ctrl.Text
    if (currentValue != "" && IsInteger(currentValue)) {
        action := GetCommitActionByNumber(currentValue)
        if (action != "") {
            ctrl.Gui.Destroy()
            ExecuteCommitAction(action)
        }
    }
}

; Manual submit function for commit selector (backup)
SubmitCommit(ctrl, *) {
    currentValue := ctrl.Gui["CommitInput"].Text
    if (currentValue != "" && IsInteger(currentValue)) {
        action := GetCommitActionByNumber(currentValue)
        if (action != "") {
            ctrl.Gui.Destroy()
            ExecuteCommitAction(action)
        } else {
            MsgBox "Invalid selection. Please choose 1-2.", "Commit Selector", "IconX"
        }
    }
}

; Cancel function for commit selector
CancelCommit(ctrl, *) {
    ctrl.Gui.Destroy()
}

; Ctrl + Alt + I : Fold all directories in VS Code Explorer
^,:: FoldAllDirectoriesInExplorer()

; Ctrl + Alt + O : Unfold all directories in VS Code Explorer
^q:: UnfoldAllDirectoriesInExplorer()

; Ctrl + 6 : Context-aware agent panel actions in Cursor
^6::
{
    targetHwnd := WinExist("A")

    toggleVisible := Cursor_IsElementVisibleByName("Toggle Agents Side Bar (Ctrl+Shift+S)", targetHwnd)
    if toggleVisible {
        Send "^e"
        return
    }

    Send "^e"
    Sleep 800

    newAgentVisible := Cursor_IsElementVisibleByName("New Agent", targetHwnd)
    if newAgentVisible {
        Sleep 600
        Send "^+s"
    }

    moreActionsEl := Cursor_GetVisibleElementByName("More Actions...", targetHwnd, ["Link", 50005, "Button", 50000])
    if moreActionsEl {
        try {
            if moreActionsEl.GetPropertyValue(UIA.Property.IsInvokePatternAvailable) {
                moreActionsEl.InvokePattern.Invoke()
            } else {
                moreActionsEl.Click()
            }
            Sleep 250
        } catch Error as e {
        }

        maximizeBtn := Cursor_GetVisibleElementByName("Maximize Chat Size", targetHwnd, ["Button", 50000], "Substring")
        if (maximizeBtn) {

            try {
                if maximizeBtn.GetPropertyValue(UIA.Property.IsInvokePatternAvailable) {
                    maximizeBtn.InvokePattern.Invoke()
                } else {
                    maximizeBtn.Click()
                }
            } catch Error as e {
            }
        }
        else {
        }
    } else {

    }
}

; Shift + N : Expand selection (via Command Palette)
+n:: Send "+!{Right}"

#HotIf

Cursor_IsElementVisibleByName(name, hwnd := 0, typeList := "", matchmode := "") {
    return !!Cursor_GetVisibleElementByName(name, hwnd, typeList, matchmode)
}

Cursor_GetVisibleElementByName(name, hwnd := 0, typeList := "", matchmode := "") {
    try {
        element := Cursor_FindElementByName(name, hwnd, typeList, matchmode)
        if !element
            return ""

        isOffscreen := true
        try isOffscreen := element.GetPropertyValue(UIA.Property.IsOffscreen)
        if isOffscreen
            return ""

        return element
    } catch Error {
        return ""
    }
}

Cursor_FindElementByName(name, hwnd := 0, typeList := "", matchmode := "") {
    try {
        if !name
            return ""
        if !hwnd
            hwnd := WinExist("A")
        if !hwnd
            return ""

        root := UIA.ElementFromHandle(hwnd)
        if !root
            return ""

        searchConfigs := []
        types := []
        if (Type(typeList) == "Array") {
            types := typeList
        } else if (typeList) {
            types := [typeList]
        }

        if (Type(types) == "Array" && types.Length) {
            for typeVal in types {
                config := { Name: name }
                if matchmode
                    config.matchmode := matchmode
                config.Type := typeVal
                searchConfigs.Push(config)
            }
        } else {
            config := { Name: name }
            if matchmode
                config.matchmode := matchmode
            searchConfigs.Push(config)
        }

        for config in searchConfigs {
            element := ""
            try element := root.FindElement(config)
            if element
                return element
        }

        return ""
    } catch Error {
        return ""
    }
}

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

        Sleep(150)
        Send("^+g")
        Sleep(350)

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
                ; For auto option: simulate ;, wait for model context menu, then send â†" , Enter
                Send "^;"
                Sleep 300
                SendText "auto"
                Sleep 500
                Send "{Enter}"
                Sleep 300
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
;-------------------------------------------------------------------
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

        ; First, try to find and click "Open Your Library" button (if available)
        try {
            openLibBtn := spot.FindElement({ Name: "Open Your Library", Type: "Button" })
            if (openLibBtn) {
                openLibBtn.Click()
                Sleep 500  ; Wait for the library to open and UI to adjust
            }
        } catch {
            ; "Open Your Library" button not found - this is normal, continue to next step
        }

        ; Then, try to find and click "Expand Your Library" button
        try {
            expandLibBtn := spot.FindElement({ Name: "Expand Your Library", Type: "Button" })
            if (expandLibBtn) {
                expandLibBtn.Click()
                Sleep 300  ; Wait for the expansion to complete
            } else {
                MsgBox "Could not find the 'Expand Your Library' button.", "Spotify Navigation", "IconX"
            }
        } catch {
            MsgBox "Could not find the 'Expand Your Library' button.", "Spotify Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error toggling fullscreen library: " e.Message, "Spotify Error", "IconX"
    }
}

; Shift + R : Send Ctrl+Shift+
+r:: Send("^+")

; Shift + T : Toggle Play/Pause using the "Download" button as anchor
; Improvements:
; - Robust word-based detection: matches any Button 50000 whose Name contains the word "play"
;   (also supports "reproduzir" / "tocar"), prefers Play over Pause when both are seen.
; - No click on the anchor. Only SetFocus/Select.
+t:: {
    try {
        spot := UIA_Browser("ahk_exe Spotify.exe")
        Sleep(200)
        if (!spot) {
            Send("{Media_Play_Pause}")
            return
        }
        WinActivate("ahk_exe Spotify.exe")
        Sleep(150)

        ; 1) Find the exact anchor: Button 50000 named "Download"
        anchor := FindDownloadAnchor(spot)
        if (!anchor) {
            ; Fallback: best-scored Play/Pause button via scan
            btn := FindBestPlayPauseButton(spot)
            if (btn) {
                ActivateElement(btn)
                return
            }
            Send("{Media_Play_Pause}")
            return
        }

        ; 2) Focus the anchor WITHOUT clicking it
        if (!FocusAnchorWithoutClick(anchor)) {
            btn := FindBestPlayPauseButton(spot)
            if (btn) {
                ActivateElement(btn)
                return
            }
            Send("{Media_Play_Pause}")
            return
        }
        Sleep(160)

        ; 3) From the anchor, back-tab up to 12 steps to locate Play/Pause
        if (HuntBackToPlayPausePreferPlay(12))  ; Shift+Tab only
            return

        ; 4) Direct lookup fallback - scan and pick the best-scoring Play/Pause button
        btn := FindBestPlayPauseButton(spot)
        if (btn) {
            ActivateElement(btn)
            return
        }

        ; 5) Final fallback – OS media key
        Send("{Media_Play_Pause}")

    } catch Error as e {
        Send("{Media_Play_Pause}")
    }
}

; ---------------------------
; Helpers
; ---------------------------

FindDownloadAnchor(root) {
    ; Exact spec: Type 50000 (Button), Name "Download"
    try {
        el := root.FindElement({ Type: 50000, Name: "Download" })
        if (el)
            return el
    } catch {
    }
    try {
        el := root.FindElement({ Type: "Button", Name: "Download" })
        if (el)
            return el
    } catch {
    }
    return ""
}

FocusAnchorWithoutClick(el) {
    ; Do NOT click the anchor
    try {
        el.SetFocus()
        return true
    } catch {
    }
    try {
        el.Select()   ; Safe, non-click focus in many UIA wrappers
        return true
    } catch {
    }
    return false
}

HuntBackToPlayPausePreferPlay(steps) {
    global UIA
    ; Prefer Play (> Pause). Keep the first Pause seen as a fallback.
    pauseCandidate := ""

    ; Check current focus first
    try {
        fe := UIA.GetFocusedElement()
        sc := PlayPauseScore(fe)
        if (sc >= 2)
            return ActivateElement(fe)  ; Found Play (or equivalent)
        else if (sc = 1)
            pauseCandidate := fe
    } catch {
    }

    loop steps {
        Send("+{Tab}")            ; Shift+Tab only
        Sleep(80)
        try {
            fe := UIA.GetFocusedElement()
            sc := PlayPauseScore(fe)
            if (sc >= 2)
                return ActivateElement(fe)  ; Prefer Play immediately
            else if (!pauseCandidate && sc = 1)
                pauseCandidate := fe        ; Remember first Pause
        } catch {
        }
    }
    if (pauseCandidate)
        return ActivateElement(pauseCandidate)
    return false
}

; Score-based detector:
; 2 = Play-like (play/reproduzir/tocar)
; 1 = Pause-like (pause/pausar/pausa)
; 0 = not a target
PlayPauseScore(el) {
    try {
        tp := el.Type
        if !(tp == 50000 || tp == "Button")
            return 0

        nm := el.Name
        if (!nm)
            return 0

        norm := NormalizeName(nm)

        if (ContainsWord(norm, "play") || ContainsWord(norm, "reproduzir") || ContainsWord(norm, "tocar"))
            return 2
        if (ContainsWord(norm, "pause") || ContainsWord(norm, "pausar") || ContainsWord(norm, "pausa"))
            return 1
    } catch {
    }
    return 0
}

ActivateElement(el) {
    try {
        el.Invoke()     ; Preferred - UIA Invoke pattern
        return true
    } catch {
    }
    try {
        el.Click()      ; Acceptable on the target (not the anchor)
        return true
    } catch {
    }
    Send("{Enter}")     ; Last resort
    Sleep(60)
    return true
}

; Scan all buttons and pick the best-scoring Play/Pause control
FindBestPlayPauseButton(root) {
    best := ""
    bestScore := 0

    ; First try numeric ControlType
    try {
        btns := root.FindAll({ Type: 50000 })
        if (btns && btns.Length) {
            for _, b in btns {
                sc := PlayPauseScore(b)
                if (sc > bestScore) {
                    best := b, bestScore := sc
                    if (bestScore >= 2)  ; Play found - early exit
                        return best
                }
            }
        }
    } catch {
    }

    ; Then try textual ControlType
    try {
        btns := root.FindAll({ Type: "Button" })
        if (btns && btns.Length) {
            for _, b in btns {
                sc := PlayPauseScore(b)
                if (sc > bestScore) {
                    best := b, bestScore := sc
                    if (bestScore >= 2)
                        return best
                }
            }
        }
    } catch {
    }

    return best
}

; --- text utils ---

NormalizeName(s) {
    ; Lowercase and collapse non-word chars (punctuation, hyphens) to single spaces
    try s := StrLower(s)
    catch {
    }
    try s := RegExReplace(s, "[^\w]+", " ")
    catch {
    }
    return Trim(s)
}

ContainsWord(norm, word) {
    ; Match a whole word boundary after normalization
    return RegExMatch(norm, "(^|\s)" . word . "(\s|$)")
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
+n:: {
    try {
        uia := TryAttachBrowser()
        if !uia {
            MsgBox "Could not attach to the browser window.", "Mobills Navigation", "IconX"
            return
        }

        ; Find anchor and start tabbing
        anchor := ""
        try {
            anchor := uia.FindElement({ Type: "Button", Name: "More details", matchmode: "Substring" })
        } catch {
            try {
                anchor := uia.FindElement({ Type: "Button" })
            } catch {
                anchor := uia.FindFirst()
            }
        }

        if (!anchor) {
            MsgBox "Could not find anchor element.", "Mobills Navigation", "IconX"
            return
        }

        ; Focus anchor
        try {
            anchor.SetFocus()
        } catch {
            anchor.Click()
        }
        Sleep(200)

        ; Tab through elements
        maxTabs := 30
        found := false
        ignoreToggleCount := 0 ; Counter for ignore toggles

        loop maxTabs {
            try {
                focused := UIA.GetFocusedElement()
                if (focused) {
                    name := focused.Name
                    type := focused.Type
                    className := focused.ClassName

                    ; Check for ignore-related elements
                    if (InStr(StrLower(name), "ignore") || InStr(StrLower(name), "ignorar") ||
                    InStr(StrLower(className), "switch") || InStr(StrLower(className), "toggle")) {

                        ignoreToggleCount++
                        if (ignoreToggleCount == 2) { ; Target the second toggle
                            Send("{Space}")
                            found := true
                            break
                        }
                    }

                    ; Check checkbox with ignore text nearby
                    if (type = 50002) {
                        try {
                            parent := UIA.TreeWalkerTrue.GetParentElement(focused)
                            if (parent) {
                                parentChildren := parent.FindAll({ Type: "Text" })
                                for child in parentChildren {
                                    try {
                                        if (InStr(StrLower(child.Name), "ignore") || InStr(StrLower(child.Name),
                                        "ignorar")) {
                                            ignoreToggleCount++
                                            if (ignoreToggleCount == 2) { ; Target the second toggle
                                                Send("{Space}")
                                                found := true
                                                break 2
                                            }
                                        }
                                    } catch {
                                        ; Skip
                                    }
                                }
                            }
                        } catch {
                            ; Continue
                        }
                    }
                }
            } catch {
                ; Continue
            }

            Send("{Tab}")
            Sleep(80)
        }

        if (!found) {
            MsgBox "Could not find the second Ignore transaction toggle.", "Mobills Navigation", "IconX"
        }

    } catch Error as e {
        MsgBox "Error: " e.Message, "Mobills Error", "IconX"
    }
}

; ---- Helper to focus the Description field directly ----y
FocusDescriptionField() {
    try {
        uia := TryAttachBrowser()
        if !uia
            return false

        ; Try multiple simple approaches
        descriptionElement := ""

        ; Method 1: Just AutomationId
        try {
            descriptionElement := uia.FindElement({ AutomationId: "mui-66475" })
        } catch {
        }

        ; Method 2: Name "Description"
        if !descriptionElement {
            try {
                descriptionElement := uia.FindElement({ Name: "Description" })
            } catch {
            }
        }

        ; Method 3: Partial ClassName match
        if !descriptionElement {
            try {
                descriptionElement := uia.FindElement({ ClassName: "MuiAutocomplete-input", matchmode: "Substring" })
            } catch {
            }
        }

        ; Method 4: Any input with "Description" in name
        if !descriptionElement {
            try {
                allInputs := uia.FindAll({ Type: "Edit" })
                for input in allInputs {
                    if InStr(input.Name, "Description") {
                        descriptionElement := input
                        break
                    }
                }
            } catch {
            }
        }

        if !descriptionElement {
            MsgBox "Could not find the Description field.", "Mobills Navigation", "IconX"
            return false
        }

        ; Try to click first, then set focus
        try {
            descriptionElement.Click()
            Sleep 100
        } catch {
        }

        descriptionElement.SetFocus()
        return true

    } catch Error as e {
        MsgBox "Error focusing Description field: " e.Message, "Mobills Error", "IconX"
        return false
    }
}

; Shift + M : Focus description field
+m:: FocusDescriptionField()

; Shift + E : Click action button then Expense menu item
+e:: {
    try {
        uia := TryAttachBrowser()
        if !uia {
            MsgBox "Could not attach to the browser window.", "Mobills Navigation", "IconX"
            return
        }

        ; First, click the action button
        actionBtn := uia.FindElement({ Type: "Button", AutomationId: "action-button" })
        if !actionBtn {
            MsgBox "Could not find the action button.", "Mobills Navigation", "IconX"
            return
        }
        actionBtn.Click()
        Sleep(300)  ; Wait for menu to appear

        ; Then click on the Expense menu item
        expenseItem := uia.FindElement({ Type: "MenuItem", Name: "Expense" })
        if !expenseItem {
            MsgBox "Could not find the Expense menu item.", "Mobills Navigation", "IconX"
            return
        }
        expenseItem.Click()

    } catch Error as e {
        MsgBox "Error clicking action button and Expense menu: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + R : Click action button then Income menu item
+r:: {
    try {
        uia := TryAttachBrowser()
        if !uia {
            MsgBox "Could not attach to the browser window.", "Mobills Navigation", "IconX"
            return
        }

        ; First, click the action button
        actionBtn := uia.FindElement({ Type: "Button", AutomationId: "action-button" })
        if !actionBtn {
            MsgBox "Could not find the action button.", "Mobills Navigation", "IconX"
            return
        }
        actionBtn.Click()
        Sleep(300)  ; Wait for menu to appear

        ; Then click on the Income menu item
        incomeItem := uia.FindElement({ Type: "MenuItem", Name: "Income" })
        if !incomeItem {
            MsgBox "Could not find the Income menu item.", "Mobills Navigation", "IconX"
            return
        }
        incomeItem.Click()

    } catch Error as e {
        MsgBox "Error clicking action button and Income menu: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + T : Click action button then Credit card expense menu item
+t:: {
    try {
        uia := TryAttachBrowser()
        if !uia {
            MsgBox "Could not attach to the browser window.", "Mobills Navigation", "IconX"
            return
        }

        ; First, click the action button
        actionBtn := uia.FindElement({ Type: "Button", AutomationId: "action-button" })
        if !actionBtn {
            MsgBox "Could not find the action button.", "Mobills Navigation", "IconX"
            return
        }
        actionBtn.Click()
        Sleep(300)  ; Wait for menu to appear

        ; Then click on the Credit card expense menu item
        creditItem := uia.FindElement({ Type: "MenuItem", Name: "Credit card expense" })
        if !creditItem {
            MsgBox "Could not find the Credit card expense menu item.", "Mobills Navigation", "IconX"
            return
        }
        creditItem.Click()

    } catch Error as e {
        MsgBox "Error clicking action button and Credit card expense menu: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + D : Click action button then Transfer menu item
+d:: {
    try {
        uia := TryAttachBrowser()
        if !uia {
            MsgBox "Could not attach to the browser window.", "Mobills Navigation", "IconX"
            return
        }

        ; First, click the action button
        actionBtn := uia.FindElement({ Type: "Button", AutomationId: "action-button" })
        if !actionBtn {
            MsgBox "Could not find the action button.", "Mobills Navigation", "IconX"
            return
        }
        actionBtn.Click()
        Sleep(300)  ; Wait for menu to appear

        ; Then click on the Transfer menu item
        transferItem := uia.FindElement({ Type: "MenuItem", Name: "Transfer" })
        if !transferItem {
            MsgBox "Could not find the Transfer menu item.", "Mobills Navigation", "IconX"
            return
        }
        transferItem.Click()

    } catch Error as e {
        MsgBox "Error clicking action button and Transfer menu: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + W : Click "Open" button and type "MAIN"
+w:: {
    try {
        uia := TryAttachBrowser()
        if !uia {
            MsgBox "Could not attach to the browser window.", "Mobills Navigation", "IconX"
            return
        }

        ; Find "Attach File" button as anchor
        anchor := uia.FindElement({ Name: "Attach File" })
        if !anchor {
            MsgBox "Could not find the Attach File button (anchor).", "Mobills Navigation", "IconX"
            return
        }

        ; Focus the anchor
        anchor.SetFocus()
        Sleep(200)  ; Wait for focus to settle

        ; Tab backwards once to reach the "Open" button
        Send("+{Tab}")  ; Shift+Tab to go backwards once
        Sleep(200)      ; Wait for focus to settle

        ; Click the "Open" button
        Send("{Enter}")
        Sleep(200)  ; Wait for any dropdown/menu to appear

        ; Type "MAIN" letter by letter for better performance
        Send("M")
        Sleep(50)
        Send("A")
        Sleep(50)
        Send("I")
        Sleep(50)
        Send("N")

    } catch Error as e {
        MsgBox "Error finding and clicking Open button: " e.Message, "Mobills Error", "IconX"
    }
}

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
; Helper for Mobills buttons â€" language-neutral search
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
    ; Strategy 1 â€" look for known class name on the container
    try {
        grp := uia.FindElement({ Type: "Group", ClassName: "sc-kAyceB", matchmode: "Substring" })
        if grp
            return grp
    }
    catch {
    }
    ; Strategy 2 â€" locate by month text (any language)
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

;-------------------------------------------------------------------
; PowerToys Command Palette Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("Command Palette")

; Ctrl + H : Trigger Ctrl+Shift+E
^h:: Send "^+e"

; Shift + K : Trigger Ctrl+Shift+C
+k:: Send "^+c"

; Shift + Y : Send ten backspaces
+y:: Send "{Backspace 10}"

; Shift + U : Insert double quotes twice, then hit left arrow
+u:: Send '""{Left}'

; Shift + I : Send "fav" letter by letter and Enter
+i:: {
    Send "{Backspace 6}"
    Send "a"
    Sleep 200
    Send "d"
    Sleep 200
    Send "{Enter}"
}

; Ctrl + 1 : Trigger Enter
^1:: Send "{Enter}"

; Ctrl + 2 : Trigger Down then Enter
^2:: {
    Send "{Down}"
    Send "{Enter}"
}

; Ctrl + 3 : Trigger Down twice then Enter
^3:: {
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Ctrl + 4 : Trigger Down three times then Enter
^4:: {
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Ctrl + 5 : Trigger Down four times then Enter
^5:: {
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Ctrl + 6 : Trigger Down five times then Enter
^6:: {
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

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
