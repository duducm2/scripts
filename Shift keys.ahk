/* ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** **
    * Win + Alt + Shift symbol layer shortcuts (AHK v2)
    * â€¢ Provides system - wide symbol shortcuts
    ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** ** /
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
#include %A_ScriptDir%\Utils.ahk
#include %A_ScriptDir%\aux\ShiftKeysIPC.ahk

; --- Global Variables ---
global DEBUG_LOG_PATH := A_ScriptDir "\.cursor\debug.log"
; Phase 5: Gate debug I/O; set to true only when diagnosing (avoids file I/O in hot paths).
global DEBUG_SHIFTKEYS := false

; Debug mode agent logging (runtime evidence for this session only)
; (disabled) agent log debug-31b036

; Helper function for safe debug logging with retry on file lock
; Handles file locking gracefully by retrying with exponential backoff
; No-op when DEBUG_SHIFTKEYS is false (production).
SafeDebugLog(text) {
    if (!DEBUG_SHIFTKEYS)
        return true
    maxRetries := 3
    retryDelay := 10
    loop maxRetries {
        try {
            FileAppend text, DEBUG_LOG_PATH
            return true
        } catch Error as err {
            ; Check if error has Number property before accessing it
            ; File lock error is typically error code 32
            hasNumber := false
            errNumber := 0
            try {
                errNumber := err.Number
                hasNumber := true
            } catch {
                ; Error doesn't have Number property, treat as non-retryable
                hasNumber := false
            }

            ; If it's a file lock error (32) and we have retries left, wait and retry
            if (hasNumber && errNumber = 32 && A_Index < maxRetries) {
                Sleep retryDelay * A_Index  ; Exponential backoff
            } else {
                ; For other errors or final retry, silently fail to not interrupt script execution
                return false
            }
        }
    }
    return false
}

; Helper: write NDJSON log line for debug agent (no-op on failure)
AgentDebugLog(hypothesisId, message, runId := "initial") {
    ; Intentionally no-op.
    return
}

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
global USE_CHROME_PDF_PRESENT_FALLBACK := true

; ShiftKeys daemon IPC: bootstrap connection on load (non-blocking)
ShiftKeysIPC_Bootstrap()

; Function to pad shortcuts to consistent width for alignment
PadShortcut(shortcut, targetWidth := 24) {
    ; Return shortcut without padding (spaces removed)
    return shortcut
}

; Function to process cheat sheet text and pad all shortcuts
ProcessCheatSheetText(text) {
    ; Split into lines
    lines := StrSplit(text, "`n")
    processedLines := []

    for line in lines {
        ; Check if line contains a shortcut pattern at the start (before any mnemonic brackets in text)
        ; Match emoji (if present) and the first bracket pattern: emoji + bracket
        ; Pattern: optional emoji/characters (non-bracket chars), then bracket, then rest of line
        if RegExMatch(line, "^([^\[\]]*?)(\[.*?\])(.*)$", &match) {
            emoji := match[1]  ; Emoji or empty
            bracket := match[2]  ; First bracket [KEY]
            restOfLine := match[3]  ; Rest of the line after bracket

            ; Pad the bracket content to align all brackets
            paddedShortcut := PadShortcut(bracket)

            ; Reconstruct line with spacing: emoji + space + padded bracket + space + rest
            ; Add space after emoji (if present) and space after bracket
            if (emoji != "") {
                processedLine := emoji . " " . paddedShortcut . " " . restOfLine
            } else {
                processedLine := paddedShortcut . " " . restOfLine
            }

            ; Check if this is a built-in shortcut (contains common built-in patterns)
            if (IsBuiltInShortcut(bracket)) {
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
Mercado Livre (Shift)
[S] Focus search field
[C] Carrinho de compras (cart)
[P] Compras feitas (purchases)
[Y] Filtro Chegará amanhã
[F] Filtro Full
[I] Filtro Internacional
[N] Filtro Envio local / Nacional
[G] Filtro Frete grátis
[O] Ordenar por (menu)
[R] Faixa de preço (Mín/Máx)
[L] Seguinte (paginação)
[K] Anterior (paginação)
[A] Adicionar ao carrinho
[V] Favoritos (coração)
[J] Continuar (fluxo compra/endereço)
)"  ; end Mercado Livre

; --- Shopee (Brazil) -------------------------------------------------------
cheatSheets["Shopee"] := "
(
Shopee (Shift)
[S] Buscar na Shopee (campo de busca)
[C] Carrinho de compras
[P] Minhas compras / Pedidos (especulativo)
[Y] Filtro Entrega Rápida (analogia Chegará amanhã)
[F] Filtro Promoções / Full (especulativo)
[I] Filtro Internacional
[N] Filtro Envio Nacional
[G] Filtro Frete grátis (especulativo)
[O] Ordenar por (menu)
[R] Faixa de preço (Mín/Máx)
[L] Seguinte (paginação) – especulativo
[K] Anterior (paginação) – especulativo
[A] Adicionar ao carrinho (página do produto)
[V] Favoritar produto (coração)
[J] Continuar (carrinho/checkout, incl. \"Fazer pedido\")
)"  ; end Shopee

;---------------------------------------- Shift + keys ----------------------------------------------
; ----- Assignment policy: use Shift + <key> first. When all Shift slots in the sequence are consumed, continue with Ctrl + Alt + <key> in the same order.
; ----- You can have repeated keys, depending on the software.
; ----- Preferred key sequence (most important first): Y U I O P H J K L N M , . W E R T D F G C V B
; ----- Ctrl + Alt sequence (fallback, same order):    Y U I O P H J K L N M , . W E R T D F G C V B
; ----- Shift + D (Teams chat) -> Fold chat sections (🩶 grey)

; --- WhatsApp desktop -------------------------------------------------------
cheatSheets["WhatsApp"] := "
(
WhatsApp (Shift)
🎤 [V]Toggle [V]oice message
🔍 [S][S]earch chats
↩️ [R][R]eply
😀 [E][E]moji panel
📬 [U]Toggle [U]nread filter
💬 [F][F]ocus current chat
✅ [M][M]ark as read/unread
📌 [P][P]in chat or unpin
)"  ; end WhatsApp

; --- Outlook main window ----------------------------------------------------
cheatSheets["OUTLOOK.EXE"] := "
(
Outlook (Shift)
📧 [G]Send to [G]eneral
📰 [N]Send to [N]ewsletter
📥 [I]Go to [I]nbox
📝 [S][S]ubject / Title
👥 [T][T]o / Required
🚫 [D][D]on't send response
✅ [E]Send [E]vent response
📝 [B][B]ody (Subject → Body)
🎯 [F][F]ocused / Other
🔀 [K]Cycle bac[K]ward pane
🔀 [L]Cyc[L]e forward pane
📋 [M]Toggle Mail / Calendar
📅 [W]eek view
📅 M[o]nth view
🪟 [P][P]op Out (Type 50000, Name Pop Out, LocalizedType button, Path {T:33,CN:rctrl_renwnd32}, {T:33}, {T:33}, {T:21}, {T:0, i:-1})

⌨️ [Ctrl+1] Move down once and select
⌨️ [Ctrl+2] Move down twice and select
⌨️ [Ctrl+3] Move down three times and select
⌨️ [Ctrl+4] Move down four times and select
⌨️ [Ctrl+5] Move down five times and select
)"  ; end Outlook

; --- Outlook Reminder window -------------------------------------------------
cheatSheets["OutlookReminder"] := "
(
Outlook - Reminders (Shift)
⏰ [H]Snooze 1 [H]our
⏰ [F]Snooze [F]our hours
⏰ [D]Snooze 1 [D]ay
❌ [X]E[X]it all reminders (Dismiss)
🌐 [J][J]oin Online

⌨️ [Ctrl+1] Move down once and select
⌨️ [Ctrl+2] Move down twice and select
⌨️ [Ctrl+3] Move down three times and select
⌨️ [Ctrl+4] Move down four times and select
⌨️ [Ctrl+5] Move down five times and select
)"  ; end Outlook Reminder

; --- Outlook Appointment window ---------------------------------------------
cheatSheets["OutlookAppointment"] := "
(
Outlook - Appointment (Shift)
📅 [S][S]tart date (combo)
📅 [P]Date [P]icker (start)
🕐 [T]Start [T]ime (combo)
📅 [E][E]nd date (combo)
🕐 [H]End [H]our (time combo)
☑️ [A][A]ll-day toggle
📝 [I]T[I]tle field
👥 [R][R]equired / To
📍 [L][L]ocation
📝 [B][B]ody
🔄 [C]Make Re[C]urring
🧙 [W][W]izard (configure)

⌨️ [Ctrl+1] Move down once and select
⌨️ [Ctrl+2] Move down twice and select
⌨️ [Ctrl+3] Move down three times and select
⌨️ [Ctrl+4] Move down four times and select
⌨️ [Ctrl+5] Move down five times and select
)"  ; end Outlook Appointment

; --- Outlook Message window ---------------------------------------------------
cheatSheets["OutlookMessage"] := "
(
Outlook - Message (Shift)
📝 [S][S]ubject / Title
👥 [T][T]o / Required
📝 [B][B]ody (Location → Body)

⌨️ [Ctrl+1] Move down once and select
⌨️ [Ctrl+2] Move down twice and select
⌨️ [Ctrl+3] Move down three times and select
⌨️ [Ctrl+4] Move down four times and select
⌨️ [Ctrl+5] Move down five times and select
)"  ; end Outlook Message

; --- Microsoft Teams â€" meeting window --------------------------------------
cheatSheets["TeamsMeeting"] := "
(
Teams - Meeting (Shift)
💬 [C]Open [C]hat pane
⛶ [M]aximize [M]eeting window
👍 [R]eact / [R]eagir
🎥 [J][J]oin now (camera + mic on)
🔊 [A][A]udio settings
)"  ; end TeamsMeeting

; --- Microsoft Teams â€" chat window -----------------------------------------
cheatSheets["TeamsChat"] := "
(
Teams - Chat (Shift)
↩️ [R][R]eply
📬 [U]View all [U]nread items
📌 [P][P]in chat
✏️ [E][E]dit message
📎 [A][A]ttach file
📜 [H][H]istory menu
📬 [M][M]ark unread
📌 [X]Unpin (e[X]it pin)
📁 [C][C]ollapse all folders
ℹ️ [I][I]nfo / Details panel
🪟 [.]Detach chat (new [.]window)
👥 [T][T]eam / Add participants
📞 [V][V]ideo call
🩶 [F][F]old chat sections
👍 [L][L]ike reaction
❤️ [G][G]ive heart reaction
😂 [J][J]oke reaction (😂)
🏠 [O][O]pen home panel

--- Search Field (Top) ---
🔍 [Alt+1]Select 1st search result (↓↓ Enter)
🔍 [Alt+2]Select 2nd search result (↓↓↓ Enter)
🔍 [Alt+3]Select 3rd search result (↓↓↓↓ Enter)
🔍 [Alt+4]Select 4th search result (↓↓↓↓↓ Enter)
🔍 [Alt+5]Select 5th search result (↓↓↓↓↓↓ Enter)

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
Spotify (Shift)
🔗 [C][C]onnect to device
⛶ [F][F]ullscreen
🔍 [S][S]earch
📋 [P][P]laylists
🎤 [A][A]rtists
💿 [B]Al[B]ums
🏠 [H][H]ome
🎵 [N][N]ow Playing
🎯 [M][M]ade For You
🆕 [R]New [R]eleases
📊 [X]E[X]plore Charts
🎵 [V][V]iew (Now Playing)
📚 [L][L]ibrary sidebar
⛶ [E][E]xpand Library
🎤 [Y]L[Y]rics
⏯️ [T][T]oggle Play/Pause
)"  ; end Spotify

; --- OneNote ---------------------------------------------------------------
cheatSheets["ONENOTE.EXE"] := "
(
OneNote (Shift)
📈 [Y]Expand [Y]section
📉 [U]Collapse ([U]nfold reverse)
📉 [I]Collapse All ([I]nward)
📈 [O][O]pen All (Expand)
📝 [P]Select [P]aragraph (line + children)
🗑️ [D][D]elete line and children
🗑️ [S][S]ingle delete (keep children)
🔍 [F][F]ind Advanced (with quotes)
)"  ; end OneNote

; --- Chrome general shortcuts ----------------------------------------------
cheatSheets["chrome.exe"] := "
(
Chrome (Shift)
🪟 [W]Pop current tab to new [W]indow
🏷️ [Ctrl+Alt+Y] [N]ame ChatGPT Window as "ChatGPT"
)"  ; end Chrome

; --- Chrome PDF Viewer ------------------------------------------------------
cheatSheets["Chrome PDF Viewer"] := "
(
Chrome PDF Viewer (Shift)
⬇️ [D] [D]ownload PDF
📏 [F] [F]it to page (zoom to fit)
🔢 [P] [P]age number field (focus)
🗂️ [T] [T]humbnails sidebar (toggle)
🔲 [2] Two-page view ([2] pages)
🎬 [E] Present mode (pr[E]sent)
)"  ; end Chrome PDF Viewer

; --- Cursor ------------------------------------------------------
cheatSheets["Cursor.exe"] := "
(
Cursor

--- CTRL Shortcuts (Cursor-defined) ---
🎯 [1] Remove clustering and focus on the code (ahk)
📁 [2] Copy path (cursor)
📊 [3] CSV: Edit CSV
💾 [4] CSV: Apply changes to source file and save
📋 [5] MarkDown Enhanced: Export in PDF format. 
📽️ [6] Marp export (PDF)
🔨 [7] Build LaTeX project
📄 [8] View LaTeX PDF file
📄 [9] Markdown Preview Enhanced: Insert Page Break
🤖 [M]Ask [M]essage, wait 6s, then paste (ahk)
⚡ [G]Kill terminal ([G]o away)
📉 [Y]Fold all (tuck awa[Y])
📈 [U] [U]nfold all
📋 [O]Open Paste As... ([O]pen)
📁 [H]Smart nav: Editor→Explorer / Explorer→Reveal (s[H]ow)
🔲 [J]Select to Bracket (ad[J]acent)
📉 [,] Fold all directories
💬 [.] Toggle chat or agent
📈 [Q]Unfold all directories (e[Q]ual)
🤖 [E] Maximize chat size — native Cursor (`workbench.action.maximizeChatSize`; user keybinding)
📂 [R]File open [R]ecent
🔍 [T]Go to [T]ype symbol in workspace
💬 [N] [N]ew chat tab (replacing current)
➕ [Enter] [I]nsert line below
🔍 [P]Open [P]roject
🔄 [1/2/3...] Switch tabs
💬 [;] Insert comment
📝 [D]Duplicate selection to next find match
🔍 [F] [F]ind
↩️ [Z]Undo (common [Z])
📊 [B]Toggle [B]ar (primary sidebar)

--- SHIFT Shortcuts (Shift) (ahk = AutoHotkey) ---
📉 [F][F]old (ahk)
📈 [U][U]nfold (ahk)
📄 [M][M]arkdown preview (cursor)
🪟 [W][W]indow (move editor) (cursor)
💻 [T][T]erminal (ahk)
💻 [N][N]ew Terminal (ahk)
📁 [E][E]xplorer (ahk)
📄🪟 [K] Mar[K]down + window (ahk)
⌨️ [C][C]ommand palette (ahk)
📈 [X] E[X]pand selection (ahk)
⚡ [S][S]ymbol in access view (cursor)
💬 [H][H]istory (chat) (ahk)
🖼️ [I][I]mage (paste) (cursor)
📁 [G][G]it repos fold (SCM) (ahk)
🔍 [Q][Q]uery Search (ahk)
🍞 [R]B[R]eadcrumbs menu (ahk)
😀 [O]Emoji selector (em[O]ji) (ahk)
🌿 [D]Git section ([D]iff) (ahk)
❌ [Z]Close all editors (end [Z]one) (ahk)
🤖 [A][A]I models switch (ahk)
🧘 [Y]Zen mode (tranquilit[Y]) (cursor)
⬇️ [P][P]ull (Git) (cursor)
✅ [V]Commit (Git sa[V]e) (cursor)
⬆️ [B]Push (Git pu[B]lish) (cursor)

--- CTRL+ALT Shortcuts (Cursor-defined) ---
📄 [Ctrl+Alt+L] Markdown Preview Enhanced: Toggle Live Update
📄 [Ctrl+Alt+T] Markdown Preview Enhanced: Toggle Scroll Sync
⬆️ [Ctrl+Alt+Up] Go to [P]arent Fold
⬅️ [Ctrl+Alt+Left] Go to sibling fold [P]revious
➡️ [Ctrl+Alt+Right] Go to sibling fold [N]ext
⬆️ [Ctrl+Alt+↑] Add cursor [A]bove
⬇️ [Ctrl+Alt+↓] Add cursor [B]elow

--- ALT Shortcuts (ahk = AutoHotkey) ---
📉 [x] Shri[X]nk selection (ahk)
📉 [,] Classical Markdown Preview
📉 [Y] Paste image to Markdown
⬇️ [U] Scroll AI feed to bottom (ahk-based)
📋 [M] Quick shortcut menu (ahk)
🤖 [A] Add file to AI Context (Cursor Chat) (ahk)
📄 [N] Review [N]ext file (ahk)
📄 [R] efresh preview
📄 [F] File: New [F]ile
📂 [O] File: New F[O]lder

--- Additional Shortcuts ---
👁️ [Alt+F12] [P]eek Definition
📝 [Ctrl+Shift+L] Select all identical words ([L]ines)
✏️ [F2] [R]ename symbol
🔍 [F8] [N]avigate problems
🗑️ [Shift+Delete] [D]elete line
⬆️ [Alt+↑] [M]ove line Up
⬇️ [Alt+↓] [M]ove line Down
👆 [Alt+Click] [M]ulti-cursor by click
⬆️ [Shift+Alt+↑] [C]opy line Up
⬇️ [Shift+Alt+↓] [C]opy line Down
🔄 [Alt+Z] Toggle word [W]rap
🐛 [Ctrl+Shift+D] [D]ebugging
⬇️ [Alt+J] Jump to [N]ext review
⬆️ [Alt+K] [P]revious review (bac[K])
)"  ; end Cursor

; --- Windows Explorer ------------------------------------------------------
cheatSheets["explorer.exe"] := "
(
Explorer (Shift)
📄 [F]Select first [F]ile
🔍 [S][S]earch bar
📍 [A][A]ddress bar
📁 [N][N]ew Folder
🔗 [H]Create s[H]ortcut
📋 [C][C]opy as path
📤 [R]Sha[R]e file
📌 [P][P]inned item (first in sidebar)
📌 [L][L]ast item (sidebar)
📦 [X] WinRAR e[X]tract here (personal); work: stub MsgBox
)"  ; end Explorer

; --- Microsoft Paint ------------------------------------------------------
cheatSheets["mspaint.exe"] := "
(
MS Paint (Shift)
📏 [R][R]esize and Skew (Ctrl+W)

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
[Ctrl+-] > 🔍 Zoom outd
)"  ; end MS Paint

; --- ClipAngel -------------------------------------------------------------
cheatSheets["ClipAngel.exe"] := "
(
ClipAngel (Shift)
📋 [C][C]opy filtered content
🔄 [T][T]oggle focus list/text
🗑️ [D][D]elete all non-favorite
🧹 [X]E[X]it filters (Clear)
⭐ [F]Mark as [F]avorite
⭐ [U][U]nmark as favorite
✏️ [E][E]dit Text
💾 [S][S]ave as file
🔗 [M][M]erge clips
🔍 [Y]File t[Y]pe filter (Quick Wizard)
⌨️ [Alt+1] [S]elect current item
⌨️ [Alt+2] [M]ove down once and select
⌨️ [Alt+3] [M]ove down twice and select
⌨️ [Alt+4] [M]ove down three times and select
⌨️ [Alt+5] [M]ove down four times and select
📋 [Ctrl+1] Tab, Select All, Copy
📋 [Ctrl+2] Down 1, Tab, Select All, Copy
📋 [Ctrl+3] Down 2, Tab, Select All, Copy
📋 [Ctrl+4] Down 3, Tab, Select All, Copy
📋 [Ctrl+5] Down 4, Tab, Select All, Copy
)"  ; end ClipAngel

; --- Figma -----------------------------------------------------------------
cheatSheets["Figma.exe"] := "
(
Figma (Shift)
👁️ [U]Toggle [U]I visibility
🔍 [S][S]earch component
⬆️ [P]Select [P]arent
🧩 [C]reate [C]omponent
🔗 [D][D]etach instance
📐 [A]dd [A]uto layout
📐 [R][R]emove auto layout
💡 [S][S]uggest auto layout
📤 [E][E]xport
🖼️ [C][C]opy as PNG
⚡ [A][A]ctions...
⬅️ [L]Align [L]eft
➡️ [R]Align [R]ight
📏 [V]Distribute [V]ertical spacing
🧹 [T][T]idy up
⬆️ [T]Align [T]op
⬇️ [B]Align [B]ottom
↔️ [H]Align center [H]orizontal
↕️ [V]Align center [V]ertical
📏 [H]Distribute [H]orizontal spacing
)"  ; end Figma

; --- Gmail ---------------------------------------------------------------
cheatSheets["Gmail"] := "
(
Gmail (Shift)
📥 [I][I]nbox
📰 [U][U]pdates
💬 [F][F]orums
📬 [R]Toggle [R]ead status
⬅️ [P][P]revious conversation
➡️ [N][N]ext conversation
📦 [A][A]rchive conversation
✅ [S][S]elect conversation
↩️ [Y]Repl[Y]
↩️ [G]Reply to [G]roup (all)
➡️ [W]For[W]ard
⭐ [T]S[T]ar toggle
🗑️ [D][D]elete
🚫 [X]Spam (e[X]clude)
✍️ [C][C]ompose new email
🔍 [Q][Q]uery mail (Search)
📁 [M][M]ove to folder
⌨️ [H][H]elp (keyboard shortcuts)
📬 [B]Inbox [B]utton

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
Google Keep (Shift)
🔍 [S][S]earch and select Note
📋 [M]Toggle [M]ain menu
)"  ; end Google Keep

; --- File Dialog ---------------------------------------------------------------
cheatSheets["FileDialog"] := "
(
File Dialog (Shift)
📄 [F]Select first [F]ile
🔍 [S][S]earch bar
📍 [A][A]ddress bar
📁 [N][N]ew Folder
📌 [P][P]inned item (first in sidebar)
💻 [T][T]his PC (sidebar)
📝 [M]File na[M]e field
✅ [O][O]pen/Save button
❌ [C][C]ancel button
)"

; --- Settings Window -------------------------------------------------
cheatSheets["Settings"] := "(Settings (Shift))`r`n🔊 [V]Set input [V]olume to 100%"

; --- Command Palette -------------------------------------------------
cheatSheets["Command Palette"] := "
(
Command Palette (Shift)
⌨️ [Ctrl+H] Reveal in file explorer
⌨️ [C][C]opy file Path
⌨️ [B]Go [H]ome
⌨️ [S]Precise [S]earch
⌨️ [I][I]nsert Favorite (Add)
⌨️ [Ctrl+1] [S]elect current item
⌨️ [Ctrl+2] [M]ove down once and select
⌨️ [Ctrl+3] [M]ove down twice and select
⌨️ [Ctrl+4] [M]ove down three times and select
⌨️ [Ctrl+5] [M]ove down four times and select
⌨️ [Ctrl+6] [M]ove down five times and select
⌨️ [Alt+1] [S]elect current item
⌨️ [Alt+2] [M]ove down once and select
⌨️ [Alt+3] [M]ove down twice and select
⌨️ [Alt+4] [M]ove down three times and select
⌨️ [Alt+5] [M]ove down four times and select
)"

; --- Excel ------------------------------------------------------------
cheatSheets["EXCEL.EXE"] := "
(
Excel (Shift)
⚪ [W]Select [W]hite Color
✏️ [E]Enable [E]diting
📊 [C][C]SV to columns (semicolon delimited)
📋 [V]Quickly [V]aste and extract CSV
➕ [A][A]dd multiple rows (10 rows)
🗑️ [R][R]ow removal workflow (remove row, down arrow, repeat 5-7 times)
📅 [P]Type [P]revious day date
)"

; --- Power BI ------------------------------------------------------------
cheatSheets["Power BI"] := "
(
Power BI (Shift)
📊 [I]Report v[I]ew
📊 [O]Table view ([O]verview)
📊 [P]Model view ([P]lan)
📊 [T][T]ransform Data
📊 [U][U]pdate (Close and Apply)
📊 [E]New M[E]asure
🔄 [Y]Refresh (read[Y])
📊 [H]Build visual ([H]andle)
📊 [J]Format visual (ad[J]ust)
⬆️ [B][B]ring forward
⬇️ [D]Sen[D] backward
📐 [K]Keep [A]lign straight
📄 [V]Fit to Page ([V]iew)
🎨 [M]Format painter ([M]atch format)
🔗 [N]Group visuals (Groupi[N]g)
🖱️ [A][A]ll pages button
➕ [W]Ne[W] Page
📕 [F]Close All Drawers ([F]old)
📖 [G]Open All Drawers (un[G]roup)
📁 [R]Collapse Fields ([R]educe)
🔍 [S][S]earch edit field
✅ [L]Confirm moda[L] button (OK)
❌ [X]Cancel/E[X]it modal button
)"

; --- UIA Tree Inspector -------------------------------------------------
cheatSheets["UIATreeInspector"] :=
"(UIA Tree Inspector (Shift))`r`n🔄 [R][R]efresh List`r`n🔍 [F]ocus [F]ilter field`r`n🔍 [S]elect [S]earch tree item`r`n📋 [C]Copy full UI tree"
; --- SettleUp Shortcuts -----------------------------------------------------
cheatSheets["Settle Up"] := "
(
Settle Up (Shift)
➕ [A][A]dd Transaction
📝 [N]Focus expense [N]ame field
💰 [V]Focus expense [V]alue field
)"

; --- Miro Shortcuts -----------------------------------------------------
cheatSheets["Miro"] := "
(
Miro (Shift)
📋 [F][F]rame List
🔗 [G][G]roup
🔗 [U][U]ngroup
🔒 [L][L]ock/Unlock
🔗 [K]Add/Edit Lin[K]
❌ [X]Close sidebar (e[X]it)
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
Wikipedia (Shift)
🔍 [S][S]earch button click
💾 [P]Save scroll [P]osition
)"

; --- YouTube ---------------------------------------------------------------
cheatSheets["YouTube"] := "
(
YouTube (Shift)
🔍 [S]Focus [S]earch box
🎬 [U]Focus first video (filter res[U]lts)
🎬 [I]Focus first v[I]deo via Explore
🏠 [H]Navigate to [H]ome
📜 [R]Navigate to histo[R]y
📋 [P]Navigate to [P]laylists
)"

; --- Google Search ---------------------------------------------------------------
cheatSheets["Google"] := "
(
Google (Shift)
🔍 [S][S]earch box focus
🥇 [U][U]se first result
)"

; --- ChatGPT ---------------------------------------------------------------
cheatSheets["ChatGPT"] := "
(
ChatGPT (Shift)
📂 [I]Toggle s[I]debar
🔄 [O]Re-send rules ([O]rder again)
📋 [C][C]opy code block
⬇️ [J]Go down ([J]ump)
🤖 [L]Send and show AI P[L]anner
)"

; --- Mobills ---------------------------------------------------------------
cheatSheets["Mobills"] := "
(
Mobills (Shift)

--- Navigation ---
📊 [D][D]ashboard
💳 [A][A]ccounts (Contas)
💰 [T][T]ransactions (Transações)
💳 [C]redit [C]ards (Cartões)
📅 [P][P]lanning (Planejamento)
📈 [R][R]eports (Relatórios)
⚙️ [M]ore [M]enu (Mais opções)
⬅️ [K]Previous month (bac[K])
➡️ [L]Next month (cyc[L]e)

--- Actions ---
🚫 [I][I]gnore transaction
✏️ [N][N]ame Field
💸 [E]New [E]xpense
💵 [Y]New Incom[Y]
💳 [X]Credit card e[X]pense
🔄 [F]Funds trans[F]er
🔘 [W][W]indow (Open button + type MAIN)
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
        if InStr(txt, "Namespace Tree Control") || InStr(txt, "Controle da Ãrvore de Namespace") {
            return cheatSheets["FileDialog"]
        }
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

        ; Chrome PDF Viewer: detect via UIA (more reliable than title substring matching)
        if IsChromePdfViewerActive()
            appShortcuts := cheatSheets.Has("Chrome PDF Viewer") ? cheatSheets["Chrome PDF Viewer"] : ""

        if InStr(chromeTitle, "WhatsApp")
            appShortcuts := cheatSheets.Has("WhatsApp") ? cheatSheets["WhatsApp"] : ""
        if InStr(chromeTitle, "Gmail")
            appShortcuts := cheatSheets.Has("Gmail") ? cheatSheets["Gmail"] : ""
        if InStr(chromeTitle, "chatgpt")
            appShortcuts := cheatSheets.Has("ChatGPT") ? cheatSheets["ChatGPT"] : ""
        if InStr(chromeTitle, "Mobills")
            appShortcuts := cheatSheets.Has("Mobills") ? cheatSheets["Mobills"] : ""
        if InStr(chromeTitle, "Google Keep") || InStr(chromeTitle, "keep.google.com")
            appShortcuts := cheatSheets.Has("Google Keep") ? cheatSheets["Google Keep"] : ""
        if InStr(chromeTitle, "YouTube")
            appShortcuts := cheatSheets.Has("YouTube") ? cheatSheets["YouTube"] : ""
        if InStr(chromeTitle, "UIATreeInspector")
            appShortcuts := cheatSheets["UIATreeInspector"]
        if InStr(chromeTitle, "Settle Up")
            appShortcuts := cheatSheets.Has("Settle Up") ? cheatSheets["Settle Up"] : ""
        if InStr(chromeTitle, "Miro")
            appShortcuts := cheatSheets.Has("Miro") ? cheatSheets["Miro"] : ""
        if InStr(chromeTitle, "Wikipedia", false) || InStr(chromeTitle, "wikipedia.org", false)
            appShortcuts := cheatSheets.Has("Wikipedia") ? cheatSheets["Wikipedia"] : ""
        ; Mercado Livre & Shopee: platform identification by URL only (no window title; it changes to product name). See shopping uia3.md.
        if IsMercadoLivreActive()
            appShortcuts := cheatSheets.Has("Mercado Livre") ? cheatSheets["Mercado Livre"] : ""
        if (appShortcuts = "" && IsShopeeActive())
            appShortcuts := cheatSheets.Has("Shopee") ? cheatSheets["Shopee"] : ""
        if InStr(chromeTitle, "gemini", false)
            appShortcuts :=
                "Gemini (Shift)`r`n📂 [D]Toggle the[D]rawer`r`n💬 [N][N]ew chat`r`n🔍 [S][S]earch`r`n🔄 [M]Change[M]odel`r`n🛠️ [T][T]ools`r`n⌨️ [P]Focus[P]rompt field`r`n📋 [C][C]opy last message`r`n🔊 [R][R]ead aloud last message`r`n🤖 [G]Send[G]emini prompt text`r`n⛶ [F][F]ullscreen input`r`n🔔 [Ctrl+Enter]Send and notify on completion"
        ; Only set generic Google sheet if nothing else matched and title clearly indicates Google site
        if (appShortcuts = "") {
            if (chromeTitle = "Google" || InStr(chromeTitle, " - Google Search"))
                appShortcuts := cheatSheets.Has("Google") ? cheatSheets["Google"] : ""
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
        ; Detect Appointment, Meeting, or Event inspector windows
        if RegExMatch(title, "i)(Appointment|Meeting|Event)") {
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

    globalText := "
(
[Win+Alt+Shift] - PRIMARY triple modifier (most used for system-wide shortcuts)
    [Ctrl+Alt+Win] - SECONDARY triple modifier

=== AVAILABLE SECONDARY (Ctrl+Alt+Win) SLOTS ===
Letters available: B, C, G, H, I, K, M, N, O, P, T, U, V, X, Y, Z
[Ctrl+Alt+Win+0] > Project Quick Selector (opens project folder in Cursor)
[Ctrl+Alt+Win+1] > Cursor AI quick action (Project Selector + Selection Mode)
[Ctrl+Alt+Win+2] > Quick Update to Your Scripts (HotStrings macro)
[Ctrl+Alt+Win+3] > Toggle Outlook and Teams (HotStrings macro)
[Ctrl+Alt+Win+5] > Clean the Clipboard (HotStrings macro)
[Ctrl+Alt+Win+6] > AI working? Cursor + Gemini (HotStrings macro)
[Ctrl+Alt+Win+7] > Mark Last Clip as Favorite (HotStrings macro; same as Ctrl+Alt+Win+J if 7 chord fails on keyboard)
[Ctrl+Alt+Win+J] > Mark Last Clip as Favorite (HotStrings macro; alternate for keyboards that ghost Ctrl+Alt+Win+7)
[Ctrl+Alt+Win+8] > Move Desktop to Recycle Bin (HotStrings macro)
[Ctrl+Alt+Win+9] > Available for future shortcut

=== MAIN KEY COMBINATIONS ===
[Symbol Layer] Win+Alt+Shift - Primary combination
[Window Management] Ctrl+Alt+Win - Secondary combination


[Win+Alt+Shift+L] > Available for future shortcut

=== CURSOR ===
[Win+Alt+Shift+N] > Opens or activates Cursor (habits, home, punctual, or work windows)

[Win+Alt+Shift+J] > Fast Copy: tap on/off (count Ctrl+C / PrtSc / Alt+PrtSc, then paste N); hold 700ms+ repeats last N (Clip Angel)

=== SPOTIFY ===
[Win+Alt+Shift+S] > Opens or activates Spotify

r=== CLIP ANGEL ===
[Win+Alt+Shift+1] > Send top list item from Clip Angel

=== GEMINI ===
[Win+Alt+Shift+I] > Opens Gemini
[Win+Alt+Shift+8] > Get word pronunciation, definition, and Portuguese translation (Gemini)
[Win+Alt+Shift+O] > Read aloud the last message in Gemini
[Win+Alt+Shift+P] > Copy the last message in Gemini
[Win+Alt+Shift+7] > Copy selected text and read aloud (Gemini)

=== HANDY DICTATION ===
[Win+Alt+Shift+0] > Start/stop dictation (transcription to clipboard)

=== YOUTUBE ===
[Win+Alt+Shift+H] > Activates Youtube

=== GOOGLE ===
[Win+Alt+Shift+F] > Opens Google

=== GMAIL ===
[Win+Alt+Shift+W] > Opens Gmail

=== CURSOR ===
[Win+Alt+Shift+,] > Opens or activates Cursor
[Win+Alt+Shift+C] > Select AI model (Handy)

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
[Win+Alt+Shift+Y] > Focus Mode: Black out all monitors except the one with the active window (toggle)

=== WINDOW MANAGEMENT (Ctrl+Alt+Win) ===
[Ctrl+Alt+Win+A] > Move window to monitor 1 (left-most)
[Ctrl+Alt+Win+S] > Move window to monitor 2
[Ctrl+Alt+Win+D] > Move window to monitor 3
[Ctrl+Alt+Win+F] > Move window to monitor 4
[Ctrl+Alt+Win+Shift+A] > Close window on monitor 1
[Ctrl+Alt+Win+Shift+S] > Close window on monitor 2
[Ctrl+Alt+Win+Shift+D] > Close window on monitor 3
[Ctrl+Alt+Win+Shift+F] > Close window on monitor 4
[Ctrl+Alt+Win+Q] > Cycle windows on monitor 1
[Ctrl+Alt+Win+W] > Cycle windows on monitor 2
[Ctrl+Alt+Win+E] > Cycle windows on monitor 3
[Ctrl+Alt+Win+R] > Cycle windows on monitor 4
[Ctrl+Alt+Win+Shift+Q] > Minimize window on monitor 1
[Ctrl+Alt+Win+Shift+W] > Minimize window on monitor 2
[Ctrl+Alt+Win+Shift+E] > Minimize window on monitor 3
[Ctrl+Alt+Win+Shift+R] > Minimize window on monitor 4

=== COMMAND PALETTE BOOKMARKS ===
[Ctrl+Alt+Win+M] > Add bookmark (Command Palette Bookmark extension)

=== GENERAL ===
[Win+Alt+Shift+U] > Quick string shortcuts
[Ctrl+Alt+Win+4] > Send AI Text Optimizer prompt to Gemini (same as Win+Alt+Shift+U then L, 4)
[Win+Alt+Shift+Q] > Jump mouse on the middle
[Win+Alt+Shift+X] > Peek PDF (tap) / Set PDF path (hold 700ms+)
[Win+Alt+Shift+→] > Show square selector (right direction)
[Win+Alt+Shift+←] > Show square selector (left direction)
[Win+Alt+Shift+↓] > Show square selector (down direction)
[Win+Alt+Shift+↑] > Show square selector (up direction)
[Win+Alt+Shift+9] > Pomodoro
[Win+Alt+Shift+.] > Clip Angel (copy, paste, and quit)

=== COMMAND PALETTE ===
[Win+Ctrl+Alt+Y] > Command Palette - File search

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

; =============================================================================
; Clip Angel: Fast Copy Mode + sequential paste (multiple clips in order)
; Hotkey: Win+Alt+Shift+J — while mode off: tap starts mode; hold 700ms+ repeats last paste count.
;         While mode on: press finishes and pastes N clips (Ctrl+C / PrtSc / Alt+PrtSc counted).
; =============================================================================
global FAST_COPY_HOLD_REPEAT_MS := 700
; Set true to write FastCopyMode_DebugLog NDJSON (development only).
global FAST_COPY_DEBUG := false
global FAST_COPY_CLIPBOARD_READ_CYCLE_MS := 1200
global FAST_COPY_GEMINI_UPLOAD_IDLE_MS := 5000
global gFastCopyModeActive := false
global gFastCopyCount := 0
global gFastCopyPasteTargetHwnd := 0
global gFastCopyLastSuccessfulCount := 0
global gFastCopyScreenshotQueue := []
global gFastCopyLastScreenshotQueue := []

FastCopyMode_DebugLog(hypothesisId, location, message, data := "") {
    ; #region agent log
    global FAST_COPY_DEBUG
    if (!FAST_COPY_DEBUG)
        return
    ; Writes NDJSON to debug-1bed80.log (Debug session: 1bed80)
    try {
        runId := "pre-fix"
        logPath := "C:\Users\fie7ca\Documents\scripts\debug-1bed80.log"
        ; Keep data small and non-sensitive; accept either a string or a Map-like object.
        dataJson := "{}"
        if (IsObject(data)) {
            parts := []
            for k, v in data {
                try parts.Push('"' FastCopyMode_JsonEscape(k) '":"' FastCopyMode_JsonEscape(v) '"')
            }
            joined := ""
            if (parts.Length) {
                for i, p in parts {
                    joined .= (i = 1 ? "" : ",") p
                }
            }
            dataJson := "{" joined "}"
        } else if (data != "") {
            dataJson := '{"value":"' FastCopyMode_JsonEscape(data) '"}'
        }
        line := '{'
            . '"sessionId":"1bed80",'
            . '"timestamp":' A_TickCount + 0 ','
            . '"runId":"' runId '",'
            . '"hypothesisId":"' FastCopyMode_JsonEscape(hypothesisId) '",'
            . '"location":"' FastCopyMode_JsonEscape(location) '",'
            . '"message":"' FastCopyMode_JsonEscape(message) '",'
            . '"data":' dataJson
            . '}'
        FileAppend(line "`n", logPath, "UTF-8")
    } catch {
        ; never break user flow
    }
    ; #endregion agent log
}

FastCopyMode_JsonEscape(s) {
    ; #region agent log
    try {
        if (s = "")
            return ""
    } catch {
        return ""
    }
    s := "" s
    s := StrReplace(s, "\", "\\")
    s := StrReplace(s, '"', '\"')
    s := StrReplace(s, "`r", "\r")
    s := StrReplace(s, "`n", "\n")
    return s
    ; #endregion agent log
}

FastCopyMode_ClipboardHasImage() {
    ; #region agent log
    ; CF_DIB=8, CF_DIBV5=17, CF_BITMAP=2
    try {
        return !!(DllCall("IsClipboardFormatAvailable", "UInt", 8, "Int")
            || DllCall("IsClipboardFormatAvailable", "UInt", 17, "Int")
            || DllCall("IsClipboardFormatAvailable", "UInt", 2, "Int"))
    } catch {
        return false
    }
    ; #endregion agent log
}

FastCopyMode_WaitForClipboardImage(timeoutMs := 1200) {
    ; #region agent log
    start := A_TickCount
    while ((A_TickCount - start) < timeoutMs) {
        if (FastCopyMode_ClipboardHasImage())
            return true
        Sleep 15
    }
    return false
    ; #endregion agent log
}

FastCopyMode_CanOpenClipboardNow() {
    ; #region agent log
    ; Returns true if clipboard can be opened immediately (no long lock).
    ok := false
    try {
        if (DllCall("OpenClipboard", "ptr", 0, "int")) {
            ok := true
            DllCall("CloseClipboard")
        }
    } catch {
        ok := false
    }
    return ok
    ; #endregion agent log
}

FastCopyMode_WaitForClipboardUnlocked(timeoutMs := 2000) {
    ; #region agent log
    start := A_TickCount
    while ((A_TickCount - start) < timeoutMs) {
        if (FastCopyMode_CanOpenClipboardNow())
            return true
        Sleep 15
    }
    return false
    ; #endregion agent log
}

FastCopyMode_WaitForClipboardReadCycle(timeoutMs := 1200) {
    ; #region agent log
    ; Wait until we observe some other process reading/locking the clipboard (OpenClipboard fails)
    ; and then wait until it becomes unlocked again. This reduces the risk of overwriting the
    ; clipboard before the target app has actually consumed the bitmap.
    start := A_TickCount
    sawLock := false
    failCount := 0
    okCount := 0

    ; Phase 1: short window with Sleep 1 (avoid busy-spinning the hotkey thread).
    sampleUntil := start + 120
    while (A_TickCount < sampleUntil) {
        if (!FastCopyMode_CanOpenClipboardNow()) {
            failCount += 1
            sawLock := true
            break
        }
        okCount += 1
        Sleep 1
    }

    ; Phase 2: regular polling until timeout budget (handles longer locks).
    if (!sawLock) {
        while ((A_TickCount - start) < timeoutMs) {
            if (!FastCopyMode_CanOpenClipboardNow()) {
                failCount += 1
                sawLock := true
                break
            }
            okCount += 1
            Sleep 10
        }
    }
    if (sawLock) {
        unlockMs := Min(1500, Max(100, timeoutMs - (A_TickCount - start)))
        if (FastCopyMode_WaitForClipboardUnlocked(unlockMs))
            return "lock_then_unlock(fails=" failCount ",oks=" okCount ")"
        return "lock_timeout(fails=" failCount ",oks=" okCount ")"
    }
    ; Never saw the clipboard get locked (target might not read immediately or at all).
    return "no_lock_seen(fails=" failCount ",oks=" okCount ")"
    ; #endregion agent log
}

FastCopyMode_SendPasteAndWaitForReadCycle(maxAttempts := 2) {
    ; #region agent log
    global FAST_COPY_CLIPBOARD_READ_CYCLE_MS
    attempt := 0
    cycle := ""
    while (attempt < maxAttempts) {
        attempt += 1
        seqBefore := Clipboard_GetSequenceNumber()
        Send "^v"
        ; Fast path (Utils.ahk): sequence changes when another process reads/updates clipboard.
        if (Clipboard_WaitForSequenceChange(seqBefore, 1200, 600)) {
            cycle := "seq_changed"
            FastCopyMode_DebugLog("H4", "Shift keys.ahk:FastCopyMode_SendPasteAndWaitForReadCycle", "paste_attempt", Map(
                "attempt", attempt,
                "cycle", cycle
            ))
            return cycle
        }
        cycle := FastCopyMode_WaitForClipboardReadCycle(FAST_COPY_CLIPBOARD_READ_CYCLE_MS)
        FastCopyMode_DebugLog("H4", "Shift keys.ahk:FastCopyMode_SendPasteAndWaitForReadCycle", "paste_attempt", Map(
            "attempt", attempt,
            "cycle", cycle
        ))
        if (InStr(cycle, "no_lock_seen") != 1)
            return cycle
        ; If we didn't observe any clipboard read/lock, try once more (some targets debounce).
        Sleep 60
    }
    return cycle
    ; #endregion agent log
}

FastCopyMode_IsGeminiActiveInChrome() {
    ; #region agent log
    try {
        return (WinActive("ahk_exe chrome.exe") && InStr(WinGetTitle("A"), "gemini", false))
    } catch {
        return false
    }
    ; #endregion agent log
}

FastCopyMode_GetActiveChromeHwnd() {
    ; #region agent log
    try {
        hwnd := WinGetID("A")
        if (hwnd && WinExist("ahk_id " hwnd) && WinActive("ahk_id " hwnd))
            return hwnd
    } catch {
    }
    return 0
    ; #endregion agent log
}

FastCopyMode_UiaForActiveChrome() {
    ; #region agent log
    hwnd := FastCopyMode_GetActiveChromeHwnd()
    if (!hwnd)
        throw Error("no_active_hwnd")
    return UIA_Browser("ahk_id " hwnd)
    ; #endregion agent log
}

FastCopyMode_GetGeminiSearchRoot(uia) {
    ; #region agent log
    ; Same pattern as GetGeminiSearchRoot in Gemini.ahk — smaller UIA subtree than full document.
    try {
        root := uia.GetCurrentMainPaneElement()
        if (root)
            return root
    } catch {
    }
    return uia
    ; #endregion agent log
}

FastCopyMode_FocusGeminiPromptField(uia := "") {
    ; #region agent log
    ; Optional uia: reuse cached UIA_Browser for the whole paste loop (efficiency).
    if (uia = "") {
        try {
            uia := FastCopyMode_UiaForActiveChrome()
        } catch Error as e {
            FastCopyMode_DebugLog("H6", "Shift keys.ahk:FastCopyMode_FocusGeminiPromptField", "uia_attach_error", Map(
                "msg", SubStr(e.Message, 1, 120)
            ))
            return false
        }
    }

    try {
        ; Use shared helper from Utils.ahk (Gemini.ahk uses the same).
        promptField := FindGeminiPromptField(uia)

        if (promptField && IsObject(promptField)) {
            try {
                promptField.SetFocus()
            } catch {
                try promptField.Click()
                catch {
                }
            }
            Sleep 80
            try {
                if (promptField.HasKeyboardFocus)
                    return true
            } catch {
                return true
            }
        }
    } catch Error as e {
        FastCopyMode_DebugLog("H6", "Shift keys.ahk:FastCopyMode_FocusGeminiPromptField", "focus_error", Map(
            "msg", SubStr(e.Message, 1, 120)
        ))
    }
    return false
    ; #endregion agent log
}

FastCopyMode_GeminiIsUploadingImage(uia) {
    ; #region agent log
    ; Heuristic: look for common uploading labels in Gemini UI (English + PT-BR).
    ; Scoped to main pane only; do not treat every ProgressBar as upload (false positives).
    try {
        root := FastCopyMode_GetGeminiSearchRoot(uia)
        texts := root.FindAll({ Type: 50020 }) ; Text
        for t in texts {
            name := t.Name
            if (!name)
                continue
            low := StrLower(name)
            if (InStr(low, "open upload file menu"))
                continue
            if (InStr(low, "upload") || InStr(low, "sending") || InStr(low, "carreg") || InStr(low, "enviando")) {
                return true
            }
        }
    } catch {
    }
    return false
    ; #endregion agent log
}

FastCopyMode_WaitForGeminiUploadIdle(uia := "", timeoutMs := 0) {
    ; #region agent log
    global FAST_COPY_GEMINI_UPLOAD_IDLE_MS
    if (timeoutMs <= 0)
        timeoutMs := FAST_COPY_GEMINI_UPLOAD_IDLE_MS
    ; Wait until Gemini is no longer showing an upload-in-progress indicator.
    start := A_TickCount
    if (uia = "") {
        try {
            uia := FastCopyMode_UiaForActiveChrome()
        } catch Error as e {
            FastCopyMode_DebugLog("H6", "Shift keys.ahk:FastCopyMode_WaitForGeminiUploadIdle", "uia_browser_error", Map(
                "msg", SubStr(e.Message, 1, 120)
            ))
            return "uia_fail"
        }
    }

    while ((A_TickCount - start) < timeoutMs) {
        if (!FastCopyMode_GeminiIsUploadingImage(uia))
            return "idle"
        Sleep 150
    }
    return "timeout"
    ; #endregion agent log
}

FastCopyMode_ReleaseHotkeyModifiers() {
    ; Ensure the Win+Alt+Shift hotkey modifiers can't leak into paste keys.
    ; Releasing modifiers does not activate or focus any other window.
    Send "{LWin up}{RWin up}{Alt up}{Shift up}{Ctrl up}"
    Sleep 30
}

FastCopyMode_CaptureScreenshotToQueue() {
    global gFastCopyScreenshotQueue

    FastCopyMode_DebugLog("H1", "Shift keys.ahk:FastCopyMode_CaptureScreenshotToQueue", "capture_enter", Map(
        "queueLenBefore", gFastCopyScreenshotQueue.Length,
        "hasImageBefore", FastCopyMode_ClipboardHasImage() ? "1" : "0"
    ))

    ; Give the OS a moment to push the screenshot into the clipboard.
    ; Alt+PrintScreen updates the clipboard with an image; rapid captures can overwrite each other
    ; unless we snapshot the clipboard right away.
    ; Wait for a real *image* to appear (not just "clipboard has something").
    try A_Clipboard := ""
    ok := FastCopyMode_WaitForClipboardImage(1500)
    FastCopyMode_DebugLog("H2", "Shift keys.ahk:FastCopyMode_CaptureScreenshotToQueue", "wait_image_done", Map(
        "ok", ok ? "1" : "0",
        "hasImageAfter", FastCopyMode_ClipboardHasImage() ? "1" : "0"
    ))
    if (!ok)
        return

    try {
        snap := ClipboardAll()
        gFastCopyScreenshotQueue.Push(snap)
        snapSize := ""
        try snapSize := snap.Size
        FastCopyMode_DebugLog("H1", "Shift keys.ahk:FastCopyMode_CaptureScreenshotToQueue", "capture_pushed", Map(
            "queueLenAfter", gFastCopyScreenshotQueue.Length,
            "hasImageAfter", FastCopyMode_ClipboardHasImage() ? "1" : "0",
            "snapType", Type(snap),
            "snapSize", snapSize
        ))
    } catch {
        ; If clipboard snapshot fails, just skip (count still increments).
        FastCopyMode_DebugLog("H2", "Shift keys.ahk:FastCopyMode_CaptureScreenshotToQueue", "clipboardall_failed", Map(
            "queueLenAfter", gFastCopyScreenshotQueue.Length,
            "hasImageAfter", FastCopyMode_ClipboardHasImage() ? "1" : "0"
        ))
    }
}

FastCopyMode_PasteScreenshotQueue(queue) {
    if (!IsObject(queue) || queue.Length < 1)
        return

    total := queue.Length
    try {
        StandardLoadingBar_Show("⏳ Pasting screenshots…", BANNER_ACCENT_INTERMEDIATE, { passive: false, centerOnHwnd: 0,
            fontSize: 17 })
    } catch {
    }

    clipSave := ""
    try clipSave := ClipboardAll()

    try {
        try {
            proc := WinGetProcessName("A")
            cls := WinGetClass("A")
            title := WinGetTitle("A")
        } catch {
            proc := ""
            cls := ""
            title := ""
        }
        isGeminiSession := (proc = "chrome.exe" && InStr(title, "gemini", false))
        cachedGeminiUia := ""
        if (isGeminiSession) {
            try {
                hwndGem := FastCopyMode_GetActiveChromeHwnd()
                if (hwndGem)
                    cachedGeminiUia := UIA_Browser("ahk_id " hwndGem)
            } catch {
                cachedGeminiUia := ""
            }
        }
        FastCopyMode_DebugLog("H3", "Shift keys.ahk:FastCopyMode_PasteScreenshotQueue", "paste_enter", Map(
            "queueLen", queue.Length,
            "hasImageAtEnter", FastCopyMode_ClipboardHasImage() ? "1" : "0",
            "proc", proc,
            "class", cls,
            "title", SubStr(title, 1, 120),
            "isGeminiSession", isGeminiSession ? "1" : "0"
        ))
        for idx, snap in queue {
            try {
                try StandardLoadingBar_Update("⏳ Pasting image " idx " / " total " …")
                catch {
                }
                ; Gemini: ensure prompt is focused BEFORE each paste.
                if (isGeminiSession) {
                    try {
                        focusedPre := FastCopyMode_FocusGeminiPromptField(cachedGeminiUia) ? "1" : "0"
                    } catch {
                        focusedPre := "0"
                    }
                    FastCopyMode_DebugLog("H6", "Shift keys.ahk:FastCopyMode_PasteScreenshotQueue", "gemini_before_paste", Map(
                        "idx", idx,
                        "focusedPrompt", focusedPre
                    ))
                }

                A_Clipboard := snap
                ClipWait 0.6, 1
                snapSize := ""
                try snapSize := snap.Size
                FastCopyMode_DebugLog("H3", "Shift keys.ahk:FastCopyMode_PasteScreenshotQueue", "paste_iter_ready", Map(
                    "idx", idx,
                    "snapType", Type(snap),
                    "snapSize", snapSize,
                    "hasImageNow", FastCopyMode_ClipboardHasImage() ? "1" : "0"
                ))
                cycle := FastCopyMode_SendPasteAndWaitForReadCycle(2)
                FastCopyMode_DebugLog("H4", "Shift keys.ahk:FastCopyMode_PasteScreenshotQueue", "paste_iter_clipboard_cycle", Map(
                    "idx", idx,
                    "cycle", cycle
                ))

                ; Gemini-specific: images require upload + prompt field focus for subsequent pastes.
                if (isGeminiSession) {
                    idleStatus := FastCopyMode_WaitForGeminiUploadIdle(cachedGeminiUia)
                    FastCopyMode_DebugLog("H6", "Shift keys.ahk:FastCopyMode_PasteScreenshotQueue", "gemini_after_paste", Map(
                        "idx", idx,
                        "uploadIdle", idleStatus
                    ))
                }
            } catch {
                ; continue to next screenshot
                FastCopyMode_DebugLog("H4", "Shift keys.ahk:FastCopyMode_PasteScreenshotQueue", "paste_iter_failed", Map(
                    "idx", idx
                ))
            }
        }
    } finally {
        if (clipSave != "")
            try A_Clipboard := clipSave
        try StandardLoadingBar_Hide(0)
        FastCopyMode_DebugLog("H3", "Shift keys.ahk:FastCopyMode_PasteScreenshotQueue", "paste_exit", Map(
            "restoredClipboard", clipSave != "" ? "1" : "0"
        ))
    }
}

ExecuteSequentialPaste(actionCount) {
    if (!IsInteger(actionCount))
        return
    n := Integer(actionCount)
    if (n < 1)
        return
    try {
        StandardLoadingBar_Show("⏳ Pasting from clipboard…", BANNER_ACCENT_INTERMEDIATE, { passive: false, centerOnHwnd: 0,
            fontSize: 17 })
    } catch {
    }
    try {
        try StandardLoadingBar_Update("⏳ Pasting clip 1 / " n " …")
        catch {
        }
        Send "!v"
        Sleep 50
        Send "^!b"
        remaining := n - 1
        loop remaining {
            try StandardLoadingBar_Update("⏳ Pasting clip " (A_Index + 1) " / " n " …")
            catch {
            }
            Sleep 300
            Send "^!b"
        }
    } finally {
        try StandardLoadingBar_Hide(0)
    }
}

FastCopyMode_IsActive() {
    global gFastCopyModeActive
    return gFastCopyModeActive
}

FastCopyMode_OnCopy() {
    global gFastCopyCount
    gFastCopyCount += 1
    FastCopyModeBanner_Update(gFastCopyCount)
}

FastCopyMode_Start() {
    global gFastCopyModeActive, gFastCopyCount, gFastCopyPasteTargetHwnd, gFastCopyScreenshotQueue
    try {
        gFastCopyPasteTargetHwnd := WinGetID("A")
    } catch {
        gFastCopyPasteTargetHwnd := 0
    }
    gFastCopyCount := 0
    gFastCopyScreenshotQueue := []
    gFastCopyModeActive := true
    try {
        FastCopyModeBanner_Show()
    } catch Error {
        gFastCopyModeActive := false
        ShowCenteredOverlay_Utils("❌ Could not start Fast Copy Mode", 2000, BANNER_ACCENT_ERROR)
    }
}

FastCopyMode_Finish() {
    global gFastCopyModeActive, gFastCopyCount, gFastCopyPasteTargetHwnd
    global gFastCopyScreenshotQueue, gFastCopyLastScreenshotQueue
    count := gFastCopyCount
    shotCount := IsObject(gFastCopyScreenshotQueue) ? gFastCopyScreenshotQueue.Length : 0
    try {
        FastCopyModeBanner_Hide()
    } finally {
        gFastCopyModeActive := false
        gFastCopyCount := 0
    }
    try {
        ; Paste exclusively into the *currently active* window without activating anything else.
        FastCopyMode_ReleaseHotkeyModifiers()
        if (count > 0) {
            FastCopyMode_DebugLog("H5", "Shift keys.ahk:FastCopyMode_Finish", "finish_paste_start", Map(
                "count", count,
                "shotCount", shotCount
            ))
            if (shotCount > 0) {
                FastCopyMode_PasteScreenshotQueue(gFastCopyScreenshotQueue)
                ; Save for hold-to-repeat behavior.
                gFastCopyLastScreenshotQueue := gFastCopyScreenshotQueue.Clone()
            } else {
                gFastCopyLastScreenshotQueue := []
            }

            remaining := count - shotCount
            FastCopyMode_DebugLog("H5", "Shift keys.ahk:FastCopyMode_Finish", "finish_remaining", Map(
                "remaining", remaining
            ))
            if (remaining > 0) {
                ; For non-screenshot copies, fall back to Clip Angel sequential paste.
                ExecuteSequentialPaste(remaining)
            }

            global gFastCopyLastSuccessfulCount
            gFastCopyLastSuccessfulCount := count
        } else
            ShowCenteredOverlay_Utils("⚠ No copies recorded — nothing to paste", 2500, BANNER_ACCENT_INTERMEDIATE)
    } catch Error as e {
        ShowCenteredOverlay_Utils("❌ Fast Copy Mode: " SubStr(e.Message, 1, 80), 2500, BANNER_ACCENT_ERROR)
    }
}

FastCopyMode_RepeatLastPaste() {
    global gFastCopyLastSuccessfulCount
    global gFastCopyLastScreenshotQueue
    if (gFastCopyLastSuccessfulCount < 1) {
        ShowCenteredOverlay_Utils("⚠ No previous Fast Copy paste to repeat", 2500, BANNER_ACCENT_INTERMEDIATE)
        return
    }
    try {
        ; Repeat paste into the *currently active* window without activating anything else.
        FastCopyMode_ReleaseHotkeyModifiers()
        if (IsObject(gFastCopyLastScreenshotQueue) && gFastCopyLastScreenshotQueue.Length > 0) {
            FastCopyMode_PasteScreenshotQueue(gFastCopyLastScreenshotQueue)
            remaining := gFastCopyLastSuccessfulCount - gFastCopyLastScreenshotQueue.Length
            if (remaining > 0)
                ExecuteSequentialPaste(remaining)
        } else {
            ExecuteSequentialPaste(gFastCopyLastSuccessfulCount)
        }
    } catch Error as e {
        ShowCenteredOverlay_Utils("❌ Repeat paste: " SubStr(e.Message, 1, 80), 2500, BANNER_ACCENT_ERROR)
    }
}

#!+j:: {
    global gFastCopyModeActive, FAST_COPY_HOLD_REPEAT_MS
    if (gFastCopyModeActive) {
        FastCopyMode_Finish()
        return
    }
    pressTime := A_TickCount
    KeyWait "j", "T1"
    holdTime := A_TickCount - pressTime
    if (holdTime >= FAST_COPY_HOLD_REPEAT_MS)
        FastCopyMode_RepeatLastPaste()
    else
        FastCopyMode_Start()
}

#HotIf FastCopyMode_IsActive()
~^c:: FastCopyMode_OnCopy()
~PrintScreen:: FastCopyMode_OnCopy()
~!PrintScreen:: {
    FastCopyMode_OnCopy()
    FastCopyMode_CaptureScreenshotToQueue()
}
#HotIf

;-------------------------------------------------------------------
; Environment paths (unchanged)
;-------------------------------------------------------------------
global WORK_SCRIPTS_PATH := "C:\Users\fie7ca\Documents\01 - Scripts"
global PERSONAL_SCRIPTS_PATH := "C:\Users\eduev\Meu Drive\17 - Projects\scripts"
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
#HotIf WinActive("ahk_exe onenote.exe") && !IsFileDialogActive()

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

; Global variables for ClipAngel filter selector
global g_ClipAngelFilterSelectorGui := false
global g_ClipAngelFilterSelectorActive := false
global g_ClipAngelFilterCharMap := Map()  ; Maps character to file type info
global g_ClipAngelFilterHotkeyHandlers := []  ; Store hotkey handlers for cleanup

; File type list (filtered to essential types only)
global g_ClipAngelFileTypes := [{ name: "img", index: 1, navKey: "I", navCount: 1 }, { name: "file", index: 2, navKey: "F",
    navCount: 1 }, { name: "text", index: 3, navKey: "T", navCount: 1 }, { name: "*url", index: 10, navKey: "*",
        navCount: 5 }, { name: "*filename", index: 13, navKey: "*", navCount: 8 }
]

; Character sequence for file type assignment (5 types)
global g_ClipAngelFilterCharSequence := ["1", "2", "3", "4", "5"]

#HotIf WinActive("ClipAngel")

; Shift + C : Select filtered content and copy
+c:: {
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

; Shift + T : Switch focus between list and text (Toggle) (F10)
+t:: Send "{F10}"

; Shift + D : Delete all non-favorite (Ctrl+Alt+K)
+d:: Send "^!k"

; Shift + X : Clear filters (F7)
+x:: Send "{F7}"

; Shift + F : Mark as favorite (Alt+Q)
+f:: Send "!q"

; Shift + U : Unmark as favorite (Unmark) (Alt+W)
+u:: Send "!w"

; Shift + E : Edit text (F4)
+e:: Send "{F4}"

; Shift + S : Save as file (Ctrl+S)
+s:: Send "^s"

; Shift + M : Merge clips
+m:: Send "^!j"

; Alt + 1 : Select current item
!1:: Send "{Enter}"

; Alt + 2 : Move down once and select
!2:: {
    Send "{Down}"
    Send "{Enter}"
}

; Alt + 3 : Move down twice and select
!3:: {
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Alt + 4 : Move down three times and select
!4:: {
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Alt + 5 : Move down four times and select
!5:: {
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Ctrl + 1–5 : Down N, F10, Select All, Copy (150ms between each step)
^1:: {
    Send "{F10}"
    Sleep 150
    Send "^a"
    Sleep 150
    Send "^c"
}
^2:: {
    Send "{Down 1}"
    Sleep 150
    Send "{F10}"
    Sleep 150
    Send "^a"
    Sleep 150
    Send "^c"
    Sleep 150
    Send "{F10}"
}
^3:: {
    Send "{Down 2}"
    Sleep 150
    Send "{F10}"
    Sleep 150
    Send "^a"
    Sleep 150
    Send "^c"
    Sleep 150
    Send "{F10}"
}
^4:: {
    Send "{Down 3}"
    Sleep 150
    Send "{F10}"
    Sleep 150
    Send "^a"
    Sleep 150
    Send "^c"
    Sleep 150
    Send "{F10}"
}
^5:: {
    Send "{Down 4}"
    Sleep 150
    Send "{F10}"
    Sleep 150
    Send "^a"
    Sleep 150
    Send "^c"
    Sleep 150
    Send "{F10}"
}

; =============================================================================
; ClipAngel Filter Selector Functions
; =============================================================================

; Navigate ClipAngel ComboBox to select a specific file type
NavigateClipAngelComboBox(typeIndex) {
    try {
        ; Ensure ClipAngel window is active
        win := WinExist("A")
        WinActivate("ahk_id " win)
        WinWaitActive("ahk_id " win, , 1)
        Sleep 50 ; Brief settle after activation

        root := UIA.ElementFromHandle(win)
        Sleep 50 ; Brief settle for UIA

        ; Find the file type filter ComboBox
        typeFilterCombo := root.FindFirst({ AutomationId: "TypeFilter", Type: 50003 })

        ; Fallback strategies
        if !typeFilterCombo {
            typeFilterCombo := root.FindFirst({ AutomationId: "TypeFilter", Type: "ComboBox" })
        }
        if !typeFilterCombo {
            typeFilterCombo := root.FindFirst({ Name: "all types", Type: 50003 })
        }
        if !typeFilterCombo {
            typeFilterCombo := root.FindFirst({ AutomationId: "TypeFilter" })
        }

        if !typeFilterCombo {
            return ; Silently fail if element not found
        }

        ; Get file type info first to show banner
        if (typeIndex < 0 || typeIndex >= g_ClipAngelFileTypes.Length) {
            return ; Invalid index
        }

        fileType := g_ClipAngelFileTypes[typeIndex + 1] ; +1 because arrays are 1-indexed

        ; Format display name for banner
        displayName := fileType.name
        if (displayName = "img") {
            displayName := "Image"
        } else if (displayName = "file") {
            displayName := "File"
        } else if (displayName = "text") {
            displayName := "Text"
        } else if (displayName = "*url") {
            displayName := "URL"
        } else if (displayName = "*filename") {
            displayName := "Filename"
        }

        ; Show banner notification
        ShowCenteredOverlay_Utils("📌 Selecting: " . displayName, 800, BANNER_ACCENT_INTERMEDIATE)

        ; Set focus and click to open dropdown
        try {
            typeFilterCombo.SetFocus()
            Sleep 30
        } catch {
            ; Continue if SetFocus fails
        }

        typeFilterCombo.Click()
        Sleep 150 ; Reduced delay after click

        ; Ensure focus is maintained
        try {
            typeFilterCombo.SetFocus()
            Sleep 100 ; Reduced delay after SetFocus
        } catch {
            typeFilterCombo.Click()
            Sleep 100 ; Reduced delay for retry
        }

        ; Reset to "all types" first - use Home key to physically anchor at top
        Send "{Home}"
        Sleep 200 ; Reduced delay after Home

        ; Navigate to target type
        if (fileType.navKey = "*") {
            ; For asterisk items, send asterisk multiple times
            loop fileType.navCount {
                Send "*"
                Sleep 100 ; Reduced delay between keystrokes
            }
        } else {
            ; For letter-based items, send the letter
            Send fileType.navKey
            Sleep 100 ; Reduced delay after navigation character
        }

        Sleep 100 ; Reduced pause after navigation

        ; Send Tab to confirm selection (NOT Enter)
        Send "{Tab}"
        Sleep 100 ; Reduced delay after TAB

        ; Return focus to clipboard list using Shift+T logic
        Send "{F10}"
        Sleep 100 ; Reduced delay after F10

        ; Send CTRL+HOME to force focus to very first item in sidebar list
        Send "^{Home}"
        Sleep 50 ; Reduced delay after CTRL+HOME
    } catch Error as e {
        ; Silently fail on error
    }
}

; Handler for character key press in filter selector
HandleClipAngelFilterChar(char) {
    global g_ClipAngelFilterSelectorActive, g_ClipAngelFilterCharMap

    ; Only process if selector is active
    if (!g_ClipAngelFilterSelectorActive) {
        return
    }

    ; Get file type info for this character
    fileTypeInfo := g_ClipAngelFilterCharMap.Get(char, "")
    if (fileTypeInfo = "") {
        ; Try lowercase if uppercase
        fileTypeInfo := g_ClipAngelFilterCharMap.Get(StrLower(char), "")
    }

    if (fileTypeInfo != "") {
        ; Cleanup selector first (closes GUI, disables hotkeys)
        CleanupClipAngelFilterSelector()

        ; Small delay to ensure ClipAngel window has focus
        Sleep 150

        ; Navigate to selected file type - use array index, not ComboBox index
        ; Find the array index by searching for matching fileType
        arrayIndex := -1
        for idx, ft in g_ClipAngelFileTypes {
            if (ft.name = fileTypeInfo.name) {
                arrayIndex := idx - 1 ; Convert to 0-based index
                break
            }
        }

        if (arrayIndex >= 0) {
            NavigateClipAngelComboBox(arrayIndex)
        }
    }
}

; Factory function to create a handler that properly captures the character
CreateClipAngelFilterCharHandler(char) {
    return (*) => HandleClipAngelFilterChar(char)
}

; Handler for Escape key
HandleClipAngelFilterEscape(*) {
    global g_ClipAngelFilterSelectorActive
    if (g_ClipAngelFilterSelectorActive) {
        CleanupClipAngelFilterSelector()
    }
}

; Cleanup function for ClipAngel filter selector
CleanupClipAngelFilterSelector() {
    global g_ClipAngelFilterSelectorActive, g_ClipAngelFilterSelectorGui, g_ClipAngelFilterHotkeyHandlers
    global g_ClipAngelFilterCharMap

    ; Disable active flag
    g_ClipAngelFilterSelectorActive := false

    ; Disable all character hotkeys
    for handler in g_ClipAngelFilterHotkeyHandlers {
        try {
            char := handler.char
            ; Handle special VK codes
            if (char = ",") {
                Hotkey("vkBC", "Off")
            } else if (char = ".") {
                Hotkey("vkBE", "Off")
            } else {
                Hotkey(char, "Off")
                ; Also disable uppercase for lowercase letters
                if (RegExMatch(char, "^[a-z]$")) {
                    Hotkey(StrUpper(char), "Off")
                }
            }
        } catch {
            ; Silently ignore errors
        }
    }

    ; Disable Escape hotkey
    try {
        Hotkey("Escape", "Off")
    } catch {
        ; Ignore
    }

    ; Clear handlers array
    g_ClipAngelFilterHotkeyHandlers := []

    ; Close and destroy GUI
    if (IsObject(g_ClipAngelFilterSelectorGui)) {
        try {
            g_ClipAngelFilterSelectorGui.Destroy()
        } catch {
            ; Ignore
        }
        g_ClipAngelFilterSelectorGui := false
    }

    ; Clear char map
    g_ClipAngelFilterCharMap := Map()
}

; Show ClipAngel filter selector GUI
ShowClipAngelFilterSelector() {
    global g_ClipAngelFilterSelectorGui, g_ClipAngelFilterSelectorActive, g_ClipAngelFilterCharMap
    global g_ClipAngelFilterHotkeyHandlers, g_ClipAngelFileTypes, g_ClipAngelFilterCharSequence

    ; Close existing GUI if open
    if (g_ClipAngelFilterSelectorActive && IsObject(g_ClipAngelFilterSelectorGui)) {
        CleanupClipAngelFilterSelector()
        Sleep 50
    }

    ; Build character mapping
    g_ClipAngelFilterCharMap := Map()
    charIndex := 0
    for fileType in g_ClipAngelFileTypes {
        if (charIndex < g_ClipAngelFilterCharSequence.Length) {
            char := g_ClipAngelFilterCharSequence[charIndex + 1] ; +1 for 1-indexed array
            g_ClipAngelFilterCharMap[char] := fileType
            charIndex++
        }
    }

    if (g_ClipAngelFilterCharMap.Count = 0) {
        TrayTip("ClipAngel Filter Selector", "No file types found.", "IconX")
        SetTimer(() => TrayTip(), -5000)
        return
    }

    ; Get monitor dimensions
    activeWin := 0
    try {
        activeWin := WinGetID("A")
    } catch {
        activeWin := 0
    }

    ; Default to primary monitor work area
    MonitorGetWorkArea(1, &monitorLeft, &monitorTop, &monitorRight, &monitorBottom)
    monitorWidth := monitorRight - monitorLeft
    monitorHeight := monitorBottom - monitorTop

    ; If we have an active window, find which monitor contains its center
    if (activeWin && activeWin != 0) {
        rect := Buffer(16, 0)
        if (DllCall("GetWindowRect", "ptr", activeWin, "ptr", rect)) {
            winLeft := NumGet(rect, 0, "int")
            winTop := NumGet(rect, 4, "int")
            winRight := NumGet(rect, 8, "int")
            winBottom := NumGet(rect, 12, "int")

            centerX := winLeft + (winRight - winLeft) // 2
            centerY := winTop + (winBottom - winTop) // 2

            monitorCount := MonitorGetCount()
            loop monitorCount {
                idx := A_Index
                MonitorGetWorkArea(idx, &l, &t, &r, &b)
                if (centerX >= l && centerX <= r && centerY >= t && centerY <= b) {
                    monitorLeft := l
                    monitorTop := t
                    monitorRight := r
                    monitorBottom := b
                    monitorWidth := r - l
                    monitorHeight := b - t
                    break
                }
            }
        }
    }

    ; Create GUI
    g_ClipAngelFilterSelectorGui := Gui("+AlwaysOnTop +ToolWindow +E0x08000000", "ClipAngel File Type Filter")
    fontSize := (monitorHeight < 800) ? 9 : 10
    g_ClipAngelFilterSelectorGui.SetFont("s" . fontSize, "Segoe UI")
    g_ClipAngelFilterSelectorGui.MarginX := 10
    g_ClipAngelFilterSelectorGui.MarginY := 5

    ; Build display text
    displayText := "ClipAngel File Type Filter`n`n"
    charIndex := 0
    for fileType in g_ClipAngelFileTypes {
        if (charIndex < g_ClipAngelFilterCharSequence.Length) {
            char := g_ClipAngelFilterCharSequence[charIndex + 1]
            ; Format display name for better readability
            displayName := fileType.name
            if (displayName = "img") {
                displayName := "Image"
            } else if (displayName = "file") {
                displayName := "File"
            } else if (displayName = "text") {
                displayName := "Text"
            } else if (displayName = "*url") {
                displayName := "URL"
            } else if (displayName = "*filename") {
                displayName := "Filename"
            }
            displayText .= "[" . char . "] " . displayName . "`n"
            charIndex++
        }
    }

    ; Calculate GUI size
    baseWidth := 300
    textControlHeight := Min(400, (g_ClipAngelFileTypes.Length * 20) + 60)
    textControlWidth := baseWidth - 20

    ; Add text control with display
    g_ClipAngelFilterSelectorGui.AddEdit("w" . textControlWidth . " h" . textControlHeight . " ReadOnly VScroll",
        displayText)

    ; Add Close button
    closeBtn := g_ClipAngelFilterSelectorGui.AddButton("w100 Default Center", "Close")
    closeBtn.OnEvent("Click", (*) => CleanupClipAngelFilterSelector())

    ; Calculate total height
    totalHeight := 10 + textControlHeight + 40 + 10
    guiWidth := baseWidth

    ; Calculate center position
    marginX := 20
    marginY := 20
    guiX := monitorLeft + (monitorWidth - guiWidth) // 2
    guiY := monitorTop + (monitorHeight - totalHeight) // 2

    ; Ensure GUI stays within monitor bounds
    if (guiX < monitorLeft + marginX)
        guiX := monitorLeft + marginX
    if (guiY < monitorTop + marginY)
        guiY := monitorTop + marginY
    if (guiX + guiWidth > monitorLeft + monitorWidth - marginX)
        guiX := monitorLeft + monitorWidth - guiWidth - marginX
    if (guiY + totalHeight > monitorTop + monitorHeight - marginY)
        guiY := monitorTop + monitorHeight - totalHeight - marginY

    ; Show GUI centered on the active window's monitor
    g_ClipAngelFilterSelectorGui.Show("NA w" . guiWidth . " h" . totalHeight . " x" . guiX . " y" . guiY)

    ; Set active flag
    g_ClipAngelFilterSelectorActive := true

    ; Clear handlers array
    g_ClipAngelFilterHotkeyHandlers := []

    ; Enable hotkeys for all assigned characters
    charIndex := 0
    for fileType in g_ClipAngelFileTypes {
        if (charIndex < g_ClipAngelFilterCharSequence.Length) {
            char := g_ClipAngelFilterCharSequence[charIndex + 1]

            ; Create handler
            handler := CreateClipAngelFilterCharHandler(char)

            ; Store handler for cleanup
            g_ClipAngelFilterHotkeyHandlers.Push({ char: char, handler: handler })

            ; Enable hotkey (both uppercase and lowercase)
            try {
                if (char = ",") {
                    Hotkey("vkBC", handler, "On")
                } else if (char = ".") {
                    Hotkey("vkBE", handler, "On")
                } else {
                    Hotkey(char, handler, "On")
                    if (RegExMatch(char, "^[a-z]$")) {
                        Hotkey(StrUpper(char), handler, "On")
                    }
                }
            } catch {
                ; Silently ignore if we can't create hotkey
            }

            charIndex++
        }
    }

    ; Enable Escape hotkey
    Hotkey("Escape", HandleClipAngelFilterEscape, "On")
}

; Shift + Y : Open file type filter selector (Quick Wizard)
+y:: {
    ; Only show selector if ClipAngel is active
    if WinActive("ClipAngel") {
        ShowClipAngelFilterSelector()
    }
}

#HotIf

;-------------------------------------------------------------------
; WhatsApp Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("WhatsApp")

global isRecording := false          ; persists between hotkey presses

; Shift + V : Toggle voice message - Voice
+v:: ToggleVoiceMessage()

; Shift + S : Search chats - Search
+s:: Send("!k")

; Shift + R : Reply - Reply
+r:: Send("!r")

; Shift + E : Emoji panel - Emoji
+e:: Send("^!s")

; Shift + U : Toggle Unread filter - Unread
+u::
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

; Shift + F : Focus current chat - Focus
+f::
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

; Shift + M : Mark as read or unread - Mark
+m:: Send "^!+u"

; Shift + P : Pin chat or unpin chat - Pin
+p:: Send "^!+p"

; ---------------------------------------------------------------------------
ToggleVoiceMessage() {
    global isRecording

    try {
        chrome := UIA_Browser()      ; top-level Chrome UIA element
        if !IsObject(chrome) {
            MsgBox "Can't attach to Chrome."
            return
        }

        Sleep 100                    ; reduced from 400ms - let Chrome finish drawing

        ; Exact-name regexes (case-insensitive, anchored ^ $)
        voicePattern := "i)^(Voice message|Record voice message)$"
        sendPattern := "i)^(Send|Stop recording)$"

        ; Helper to grab a button by pattern
        ; Use longer timeout (3000ms) for voice message button to allow WhatsApp UI to restore
        FindBtn(p) => WaitForButton(chrome, p, 3000)

        if (isRecording) {           ; â–º we're supposed to stop & send
            if (btn := FindBtn(sendPattern)) {
                ; Determine if this button supports Invoke
                supportsInvoke := false
                try {
                    supportsInvoke := btn.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)
                } catch {
                    supportsInvoke := false
                }

                ; Try multi-strategy activation: prefer Invoke when available, fallback to Click
                clicked := false
                if (supportsInvoke) {
                    try {
                        btn.Invoke()
                        clicked := true
                    } catch {
                    }
                }
                if (!clicked) {
                    try {
                        btn.Click()
                        clicked := true
                    } catch {
                    }
                }

                isRecording := false
                ; Give WhatsApp time to restore the UI after sending
                Sleep 300
            } else {
                ; Assume you clicked Send manually > reset & start new rec
                isRecording := false
                if (btn := FindBtn(voicePattern)) {
                    btn.Click()
                    isRecording := true
                } else
                    MsgBox "Couldn't restart recording (Voice-message button missing)."
            }
        } else {                     ; â–º start recording
            if (btn := FindBtn(voicePattern)) {
                btn.Click()
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
    ; #region agent log
    SafeDebugLog Format(
        "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:1727`",`"message`":`"WaitForButton entry`",`"data`":{`"pattern`":`"{4}`",`"timeout`":{5}},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"B`"}`n",
        A_TickCount, Random(1000, 9999), A_TickCount, pattern, timeout)
    ; #endregion
    if !IsObject(root)
        return 0

    deadline := A_TickCount + timeout
    while (A_TickCount < deadline) {
        buttons := root.FindAll({ Type: "Button" })
        ; #region agent log
        SafeDebugLog Format(
            "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:1733`",`"message`":`"Found buttons count`",`"data`":{`"count`":{4}},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"B`"}`n",
            A_TickCount, Random(1000, 9999), A_TickCount, buttons.Length)
        ; #endregion

        ; Collect all matching buttons and their properties
        matchingButtons := []
        for btn in buttons {
            btnName := ""
            try btnName := btn.Name
            ; #region agent log
            if InStr(pattern, "Connect") {
                SafeDebugLog Format(
                    "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:1738`",`"message`":`"Checking button name`",`"data`":{`"name`":`"{4}`",`"pattern`":`"{5}`"},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"B`"}`n",
                    A_TickCount, Random(1000, 9999), A_TickCount, btnName, pattern)
            }
            ; #endregion
            if RegExMatch(btn.Name, pattern) {
                className := ""
                hasClassName := false
                supportsInvoke := false
                parentName := ""
                parentClass := ""

                try {
                    className := btn.ClassName
                    hasClassName := (className != "")
                } catch {
                    ; ClassName property not available or error reading
                    hasClassName := false
                }

                ; Try to capture parent info for better disambiguation (esp. duplicated "Send" buttons)
                try {
                    parent := btn.Parent
                    parentName := parent.Name
                    parentClass := parent.ClassName
                } catch {
                    parentName := ""
                    parentClass := ""
                }

                ; Check if button supports Invoke pattern
                try {
                    supportsInvoke := btn.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)
                } catch {
                    supportsInvoke := false
                }

                ; #region agent log
                SafeDebugLog Format(
                    "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:1770`",`"message`":`"Matching button found`",`"data`":{`"name`":`"{4}`",`"className`":`"{5}`",`"hasClassName`":{6},`"supportsInvoke`":{7}},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"B,C`"}`n",
                    A_TickCount, Random(1000, 9999), A_TickCount, btnName, className, hasClassName, supportsInvoke)
                ; #endregion
                matchingButtons.Push({ btn: btn, hasClassName: hasClassName, className: className, supportsInvoke: supportsInvoke,
                    parentName: parentName, parentClass: parentClass })
            }
        }

        ; If we found matching buttons, prioritize: 1) hasClassName (the actual clickable button), 2) supportsInvoke, 3) first
        ; Note: In WhatsApp, the button WITH ClassName is the actual clickable one, even if it doesn't support Invoke pattern
        if matchingButtons.Length > 0 {
            bestBtn := ""
            bestClassName := ""
            bestScore := 0

            for match in matchingButtons {
                score := 0
                if match.hasClassName && match.className != "" {
                    score += 10  ; Highest priority: has ClassName (the actual clickable button in WhatsApp)
                }
                if match.supportsInvoke {
                    score += 5   ; Second priority: supports Invoke pattern
                }

                ; Additional heuristic for WhatsApp voice "Send" vs text "Send" (H8)
                ; When using the sendPattern, prefer the inner child button whose parent is also "Send"
                if InStr(pattern, "Send|Stop recording") {
                    try {
                        if (match.parentName = "Send") {
                            score += 3
                        }
                    }
                }

                if (score > bestScore) {
                    bestBtn := match.btn
                    bestClassName := match.className
                    bestScore := score
                }
            }

            ; If no button scored (shouldn't happen), use the first one
            if !bestBtn {
                bestBtn := matchingButtons[1].btn
            }

            ; #region agent log
            finalBtnName := "", finalBtnClassName := "", finalBtnType := ""
            try finalBtnName := bestBtn.Name
            try finalBtnClassName := bestBtn.ClassName
            try finalBtnType := bestBtn.ControlType
            SafeDebugLog Format(
                "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:1813`",`"message`":`"WaitForButton returning button`",`"data`":{`"name`":`"{4}`",`"className`":`"{5}`",`"type`":`"{6}`",`"score`":{7}},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"C`"}`n",
                A_TickCount, Random(1000, 9999), A_TickCount, finalBtnName, finalBtnClassName, finalBtnType, bestScore)
            ; #endregion
            return bestBtn
        }
        Sleep 50  ; reduced from 150ms to 50ms for faster retries
    }

    ; #region agent log
    SafeDebugLog Format(
        "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:1818`",`"message`":`"WaitForButton timeout - no button found`",`"data`":{`"pattern`":`"{4}`"},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"B`"}`n",
        A_TickCount, Random(1000, 9999), A_TickCount, pattern)
    ; #endregion
    return 0
}

#HotIf

;-------------------------------------------------------------------
; Outlook Reminder Window Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe OUTLOOK.EXE") && RegExMatch(WinGetTitle("A"), "i)Reminder") && !IsFileDialogActive()

; ativa a janela de lembretes do Outlook
ActivateReminder() {
    if (!WinExist("ahk_exe OUTLOOK.EXE")) {
        ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
        return
    }
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

; Shift + H : Snooze 1 hour - Hour
+H:: Confirm("1 hour")

; Shift + F : Snooze 4 hours - Four
+F:: Confirm("4 hours")

; Shift + D : Snooze 1 day - Day
+D:: Confirm("1 day")

; Shift + X : Dismiss all reminders - Dismiss
+X:: ConfirmDismissAll()

; Shift + J : Join Online - Join
+J::
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

; Ctrl + 1 : Move down once and select (Command Palette)
^1:: {
    Send "{Down}"
    Send "{Enter}"
}

; Ctrl + 2 : Move down twice and select (Command Palette)
^2:: {
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Ctrl + 3 : Move down three times and select (Command Palette)
^3:: {
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Ctrl + 4 : Move down four times and select (Command Palette)
^4:: {
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Ctrl + 5 : Move down five times and select (Command Palette)
^5:: {
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
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

; Shift + C : Open Chat pane - Chat
+C:: {
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

; Shift + M : Maximize meeting window - Maximize
+M:: {
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
        if (!WinExist("ahk_id " normalMeetingHwnd)) {
            ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        try {
            WinActivate("ahk_id " normalMeetingHwnd)
            ; Optional: Show a brief tooltip to confirm the switch
            ShowCenteredOverlay_Utils("✅ Switched to normal meeting view", 1000, BANNER_ACCENT_SUCCESS)
        } catch as e {
            ; Fallback: try to bring window to front (only if window still exists)
            if (WinExist("ahk_id " normalMeetingHwnd)) {
                WinShow("ahk_id " normalMeetingHwnd)
                WinActivate("ahk_id " normalMeetingHwnd)
            } else {
                ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
            }
        }
    } else {
        ; No corresponding normal window found - show notification
        ShowCenteredOverlay_Utils("⚠ No normal meeting window found", 1500, BANNER_ACCENT_ERROR)
    }
}

; Shift + R : React / Reagir - React
+R:: {
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

; Shift + J : Join now with camera and microphone on - Join
+J:: {
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

; Shift + A : Audio settings - Audio
+A:: {
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
; Global variable to track scroll position history (stack: most recent last)
global g_WikipediaScrollHistory := []

#HotIf WinActive("ahk_exe chrome.exe") && InStr(WinGetTitle("A"), "Wikipedia", false)

; Shift + S: Focus the Wikipedia search field (prefer the field; if hidden, click the Search toggle first)
+s::
{
    try {
        ; Step 1: Always scroll to the beginning of the page first
        Send "^Home"
        Sleep 300 ; Give page time to scroll

        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 200 ; Give UIA time to attach

        ; Prefer document; fall back to browser root
        try {
            root := uia.GetCurrentDocumentElement()
        } catch {
            root := uia.BrowserElement
        }

        ; Step 2: Try to locate the "Search Wikipedia" field by name (combo box / edit)
        searchBox := 0

        ; First try: ComboBox with the expected name (Type 50003)
        try {
            searchBox := root.FindElement({ Type: 50003, Name: "Search Wikipedia", cs: false })
        } catch {
        }

        ; Second: Edit control with the same name (in case UI changes type)
        if (!searchBox) {
            try {
                searchBox := root.FindElement({ Type: 50004, Name: "Search Wikipedia", cs: false })
            } catch {
            }
        }

        ; Third: any element by name "Search Wikipedia"
        if (!searchBox) {
            try {
                searchBox := root.FindElement({ Name: "Search Wikipedia", cs: false })
            } catch {
            }
        }

        ; If we found the field, focus/click it and we're done.
        if (searchBox) {
            try {
                searchBox.SetFocus()
            } catch {
                searchBox.Click()
            }
            return
        }

        ; Step 3: If the field is not available, try clicking the "Search" toggle button/link first.
        searchToggle := 0

        ; Strategy 1: Try finding the search group first, then the button within it (most reliable)
        try {
            searchGroup := root.FindElement({ AutomationId: "p-search", cs: false })
            if (searchGroup) {
                try {
                    searchToggle := searchGroup.FindElement({ Type: 50005, Name: "Search", cs: false })
                } catch {
                }
                if (!searchToggle) {
                    try {
                        searchToggle := searchGroup.FindElement({ Type: 50005, Value: "https://en.wikipedia.org/wiki/Special:Search",
                            cs: false })
                    } catch {
                    }
                }
            }
        } catch {
        }

        ; Strategy 2: Search by Value (URL) directly from root
        if (!searchToggle) {
            try {
                searchToggle := root.FindElement({ Type: 50005, Value: "https://en.wikipedia.org/wiki/Special:Search",
                    cs: false })
            } catch {
            }
        }

        ; Strategy 3: Search by Type and Name (original method)
        if (!searchToggle) {
            try {
                searchToggle := root.FindElement({ Type: 50005, Name: "Search", cs: false })
            } catch {
            }
        }

        ; Strategy 4: Search by ControlType and Name
        if (!searchToggle) {
            try {
                searchToggle := root.FindElement({ ControlType: "Hyperlink", Name: "Search", cs: false })
            } catch {
            }
        }

        ; Strategy 5: Search for any link with the Special:Search URL
        if (!searchToggle) {
            try {
                searchToggle := root.FindElement({ Type: 50005, Value: "*Special:Search*", cs: false })
            } catch {
            }
        }

        ; Strategy 6: Try finding by LocalizedType "link" and Name
        if (!searchToggle) {
            try {
                searchToggle := root.FindElement({ LocalizedType: "link", Name: "Search", cs: false })
            } catch {
            }
        }

        if (searchToggle) {
            ; Validate element before clicking: check if it's offscreen or disabled
            isValid := true
            try {
                isOffscreen := searchToggle.GetPropertyValue(UIA.Property.IsOffscreen)
                if (isOffscreen) {
                    isValid := false
                }
            } catch {
                ; Continue even if property check fails
            }

            try {
                isEnabled := searchToggle.GetPropertyValue(UIA.Property.IsEnabled)
                if (!isEnabled) {
                    isValid := false
                }
            } catch {
                ; Continue even if property check fails
            }

            ; If element is invalid, try re-finding it
            if (!isValid) {
                Sleep 200
                searchToggle := 0
                try {
                    searchToggle := root.FindElement({ Type: 50005, Name: "Search", cs: false })
                } catch {
                }
                if (!searchToggle) {
                    try {
                        searchToggle := root.FindElement({ ControlType: "Hyperlink", Name: "Search", cs: false })
                    } catch {
                    }
                }
            }

            ; Try multiple click strategies if element is still valid
            if (searchToggle) {
                clicked := false

                ; Strategy 1: Try Invoke pattern (most reliable for links/buttons)
                try {
                    searchToggle.Invoke()
                    clicked := true
                } catch Error as invokeErr {
                }

                ; Strategy 2: Try SetFocus then Click
                if (!clicked) {
                    try {
                        searchToggle.SetFocus()
                        Sleep 50
                        searchToggle.Click()
                        clicked := true
                    } catch Error as focusClickErr {
                    }
                }

                ; Strategy 3: Try direct Click
                if (!clicked) {
                    try {
                        searchToggle.Click()
                        clicked := true
                    } catch Error as clickErr {
                    }
                }

                ; If all click strategies failed, fall back to accelerator
                if (!clicked) {
                    try {
                        uia.ControlSend("!f")
                    } catch {
                    }
                }

                ; Give the UI a moment to reveal the field, then try again to find it.
                Sleep 250

                searchBox := 0
                try {
                    searchBox := root.FindElement({ Type: 50003, Name: "Search Wikipedia", cs: false })
                } catch {
                }
                if (!searchBox) {
                    try {
                        searchBox := root.FindElement({ Type: 50004, Name: "Search Wikipedia", cs: false })
                    } catch {
                    }
                }
                if (!searchBox) {
                    try {
                        searchBox := root.FindElement({ Name: "Search Wikipedia", cs: false })
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
        }

        ; Final fallback: use the accelerator key if all else fails.
        try {
            uia.ControlSend("!f")
            return
        } catch {
        }

        MsgBox "Could not find the 'Search Wikipedia' field."
    } catch Error as e {
        MsgBox "An error occurred: " e.Message
    }
}

; Shift + P: Save Wikipedia scroll position
+p::
{
    SaveWikipediaScrollPositionManually_ShiftKeys()
}

; Helper function to restore scroll position to a given percentage
; Returns true on success, false on failure
RestoreWikipediaScrollPosition(scrollPercentage, bannerText := "📜 Restoring scroll position... Please wait") {
    if (scrollPercentage <= 0.0 || scrollPercentage > 1.0) {
        return false
    }

    try {
        ; Create UIA_Browser once
        uia := UIA_Browser("ahk_exe chrome.exe")
        if (!uia) {
            return false
        }

        ; Show banner (foreground monitor)
        StandardLoadingBar_Show(bannerText, BANNER_ACCENT_INTERMEDIATE, { passive: true, centerOnHwnd: 0,
            textWidth: 500,
            fontSize: 17, passiveBgColor: BANNER_ACCENT_INTERMEDIATE })

        ; Block input during restoration (Phase 4: guaranteed cleanup in finally)
        BlockInput("On")
        try {
            ; Wait for page to be ready (condition-based, up to 500ms) instead of fixed Sleep(500)
            deadline := A_TickCount + 500
            docHeight := ""
            loop {
                docHeight := uia.JSReturnThroughClipboard("document.documentElement.scrollHeight")
                if (docHeight != "" && docHeight != "undefined" && docHeight != "null") {
                    try h := Float(docHeight)
                    catch {
                        h := 0
                    }
                    if (h > 0)
                        break
                }
                if (A_TickCount >= deadline)
                    break
                Sleep(50)
            }
            if (docHeight = "" || docHeight = "undefined" || docHeight = "null") {
                StandardLoadingBar_Hide(0)
                return false
            }
            docHeightFloat := Float(docHeight)
            if (docHeightFloat <= 0) {
                StandardLoadingBar_Hide(0)
                return false
            }

            ; Calculate and execute scroll
            targetScrollY := scrollPercentage * docHeightFloat
            uia.JSExecute("window.scrollTo(0, " . Round(targetScrollY) . ");")
            deadline2 := A_TickCount + 500
            while (A_TickCount < deadline2)
                Sleep(50)

            ; Update banner to show success
            try {
                StandardLoadingBar_Update("Scroll position restored!")
                Sleep(1000)
            } catch {
            }

            try {
                Sleep(500)
                StandardLoadingBar_Hide(0)
            } catch {
            }

            return true
        } catch Error as err {
            try StandardLoadingBar_Hide(0)
            catch {
            }
            return false
        } finally {
            BlockInput("Off")
        }
    } catch Error as err {
        BlockInput("Off")
        try StandardLoadingBar_Hide(0)
        catch {
        }
        return false
    }
}

; Helper function to normalize Wikipedia URLs
NormalizeWikipediaURL(url) {
    if (url = "" || !InStr(url, "wikipedia.org")) {
        return ""
    }
    ; Remove fragments and trailing slashes
    url := RegExReplace(url, "/#.*$", "")
    url := RegExReplace(url, "/+$", "")
    return url
}

; Helper function to get and normalize Wikipedia URL from active window
GetWikipediaURLNormalized() {
    try {
        if (!WinActive("ahk_exe chrome.exe") || !InStr(WinGetTitle("A"), "Wikipedia")) {
            return ""
        }
        uia := UIA_Browser("ahk_exe chrome.exe")
        url := uia.GetCurrentURL()
        return NormalizeWikipediaURL(url)
    } catch Error as err {
        return ""
    }
}

; Wikipedia scroll position save function (duplicated from AppLaunchers.ahk)
SaveWikipediaScrollPositionManually_ShiftKeys() {
    try {
        ; Check if Wikipedia window is currently active
        activeWindow := WinGetTitle("A")
        isChromeActive := WinActive("ahk_exe chrome.exe")
        hasWikipedia := InStr(activeWindow, "Wikipedia")
        if (!isChromeActive || !hasWikipedia) {
            return false
        }
    } catch Error as err {
        return false
    }

    ; Exit fullscreen before scroll position save (REQUIRED: UIA unreliable in fullscreen)
    Send("{F11}")
    Sleep(300)  ; Allow time for fullscreen exit (increased for reliability)

    ; Show banner to inform user that scroll position is being saved (foreground monitor)
    StandardLoadingBar_Show("💾 Saving scroll position... Please wait", BANNER_ACCENT_INTERMEDIATE, { passive: true,
        centerOnHwnd: 0,
        textWidth: 500, fontSize: 17, passiveBgColor: BANNER_ACCENT_INTERMEDIATE })
    fullscreenRestored := false  ; Track if we've re-entered fullscreen
    try {
        ; Get normalized Wikipedia URL
        url := GetWikipediaURLNormalized()
        if (url = "") {
            ; Re-enter fullscreen before returning
            Send("{F11}")
            Sleep(300)
            fullscreenRestored := true
            return false
        }

        ; Create UIA_Browser for getting scroll position
        uia := false
        try {
            uia := UIA_Browser("ahk_exe chrome.exe")
        } catch Error as uiaErr {
            try StandardLoadingBar_Hide(0)
            catch {
            }
            ; Re-enter fullscreen before returning
            Send("{F11}")
            Sleep(300)
            fullscreenRestored := true
            return false
        }

        if (!uia) {
            try StandardLoadingBar_Hide(0)
            catch {
            }
            ; Re-enter fullscreen before returning
            Send("{F11}")
            Sleep(300)
            fullscreenRestored := true
            return false
        }

        ; Wait for page to stabilize before measuring (critical for portrait orientation)
        ; Monitors 3 and 4 are portrait (1080x1920); layout may shift during measurement
        Sleep(500)  ; Brief stabilization wait

        ; Get scroll position with retry for stability
        scrollY := ""
        docHeight := ""
        scrollYRetries := 3
        lastScrollY := -1
        loop scrollYRetries {
            try {
                scrollY := uia.JSReturnThroughClipboard("window.pageYOffset")
                scrollYFloat := Float(scrollY)
                ; Check if scroll position is stable
                if (scrollYFloat = lastScrollY || lastScrollY = -1) {
                    if (scrollYFloat = lastScrollY) {
                        break  ; Stable, use this value
                    }
                    lastScrollY := scrollYFloat
                }
            } catch Error as scrollErr {
            }
            if (A_Index < scrollYRetries) {
                Sleep(200)  ; Wait between attempts
            }
        }

        ; Get document height to calculate percentage (with stability check)
        docHeightRetries := 3
        lastDocHeight := -1
        loop docHeightRetries {
            try {
                docHeight := uia.JSReturnThroughClipboard("document.documentElement.scrollHeight")
                docHeightFloat := Float(docHeight)
                ; Check if document height is stable
                if (docHeightFloat = lastDocHeight || lastDocHeight = -1) {
                    if (docHeightFloat = lastDocHeight) {
                        break  ; Stable, use this value
                    }
                    lastDocHeight := docHeightFloat
                }
            } catch Error as docErr {
            }
            if (A_Index < docHeightRetries) {
                Sleep(200)  ; Wait between attempts
            }
        }

        ; Convert to numbers and calculate percentage
        if (scrollY != "" && scrollY != "undefined" && scrollY != "null" && docHeight != "" && docHeight !=
            "undefined" && docHeight != "null") {
            scrollYFloat := Float(scrollY)
            docHeightFloat := Float(docHeight)
            if (scrollYFloat >= 0 && docHeightFloat > 0) {
                scrollPercentage := scrollYFloat / docHeightFloat
                ; Clamp to valid range
                if (scrollPercentage > 1.0) {
                    scrollPercentage := 1.0
                }

                ; Save to INI file
                scrollPositionsFile := A_ScriptDir "\data\wikipedia_scroll_positions.ini"
                SplitPath(scrollPositionsFile, , &dir)
                if (dir != "" && !DirExist(dir)) {
                    DirCreate(dir)
                }
                ; Always add to history stack for "go back" functionality (independent of INI file save)
                global g_WikipediaScrollHistory
                ; Add current position to history (url, scrollPercentage)
                g_WikipediaScrollHistory.Push({ url: url, scrollPercentage: scrollPercentage })
                ; Limit history to last 10 positions to prevent memory issues
                if (g_WikipediaScrollHistory.Length > 10) {
                    g_WikipediaScrollHistory.RemoveAt(1)
                }

                ; Try to save to INI file (for persistence across sessions)
                saved := false
                try {
                    ; Read existing entries first (before deleting file) to preserve them
                    existingEntries := Map()
                    if (FileExist(scrollPositionsFile)) {
                        try {
                            ; Read all existing entries from the Positions section
                            fileContent := FileRead(scrollPositionsFile)
                            ; Parse INI format manually
                            inPositionsSection := false
                            loop parse fileContent, "`n", "`r" {
                                line := Trim(A_LoopField)
                                if (line = "[Positions]") {
                                    inPositionsSection := true
                                    continue
                                }
                                if (inPositionsSection && SubStr(line, 1, 1) = "[") {
                                    ; Hit another section, stop reading
                                    break
                                }
                                if (inPositionsSection && InStr(line, "=")) {
                                    pos := InStr(line, "=")
                                    key := Trim(SubStr(line, 1, pos - 1))
                                    value := Trim(SubStr(line, pos + 1))
                                    if (key != "" && value != "") {
                                        existingEntries[key] := value
                                    }
                                }
                            }
                        } catch {
                            ; If read fails, we'll just write the new entry
                        }
                    }

                    ; Update with new entry
                    existingEntries[url] := scrollPercentage

                    ; Delete file to recreate in UTF-8
                    if (FileExist(scrollPositionsFile)) {
                        try {
                            FileDelete(scrollPositionsFile)
                            Sleep(100)  ; Small delay to ensure file system updates
                        } catch {
                        }
                    }

                    ; Write all entries back in UTF-8 encoding
                    try {
                        ; Write UTF-8 BOM and section header
                        FileAppend("[Positions]`r`n", scrollPositionsFile, "UTF-8")
                        ; Write each entry
                        for key, value in existingEntries {
                            ; Escape special INI characters in key and value
                            escapedKey := StrReplace(key, "=", "`=")
                            escapedKey := StrReplace(escapedKey, ";", "`;")
                            escapedValue := StrReplace(value, "`n", "`;")
                            escapedValue := StrReplace(escapedValue, "`r", "")
                            FileAppend(escapedKey . "=" . escapedValue . "`r`n", scrollPositionsFile, "UTF-8")
                        }
                        saved := true
                    } catch {
                        ; Fallback to IniWrite if manual write fails
                        saved := IniWrite(scrollPercentage, scrollPositionsFile, "Positions", url)
                    }
                } catch Error as iniErr {
                    saved := false
                }

                if (saved) {
                    ; Update banner to show success
                    try {
                        StandardLoadingBar_Update("Scroll position saved!")
                        Sleep(1000)  ; Show success message for 1 second
                    } catch {
                    }
                    ; Re-enter fullscreen after successful save
                    Send("{F11}")
                    Sleep(300)
                    fullscreenRestored := true
                    return true
                } else {
                    ; INI save failed - show error message
                    try {
                        StandardLoadingBar_Update("Error: Save failed")
                        Sleep(2000)  ; Show error message
                    } catch {
                    }
                    ; Re-enter fullscreen even on failure
                    Send("{F11}")
                    Sleep(300)
                    fullscreenRestored := true
                    return false
                }
            }
        }
    } catch Error as err {
    } finally {
        ; Always hide the banner after save operation completes
        try {
            Sleep(500)  ; Brief delay before hiding
            StandardLoadingBar_Hide(0)
        } catch {
        }
        ; Re-enter fullscreen if we haven't already (e.g., if exception occurred or validation failed)
        if (!fullscreenRestored) {
            try {
                Send("{F11}")
                Sleep(300)
            } catch {
            }
        }
    }
    return false
}

; Restore previous scroll position from history
RestorePreviousWikipediaScrollPosition() {
    global g_WikipediaScrollHistory

    try {
        ; Check if Wikipedia window is currently active
        activeWindow := WinGetTitle("A")
        isChromeActive := WinActive("ahk_exe chrome.exe")
        hasWikipedia := InStr(activeWindow, "Wikipedia")
        if (!isChromeActive || !hasWikipedia) {
            return false
        }
    } catch Error as err {
        return false
    }

    ; If history is empty, try to fall back to INI file (for positions saved via activation restore or previous sessions)
    if (g_WikipediaScrollHistory.Length = 0) {
        ; Get current URL to load from INI
        try {
            url := GetWikipediaURLNormalized()
            if (url = "") {
                ; Show brief message that no history exists
                ShowCenteredOverlay_Utils("⚠ No previous scroll position found", 1500, BANNER_ACCENT_ERROR)
                return false
            }

            ; Load from INI file
            scrollPositionsFile := A_ScriptDir "\data\wikipedia_scroll_positions.ini"
            savedPercentage := IniRead(scrollPositionsFile, "Positions", url, "0")
            savedPercentageFloat := Float(savedPercentage)

            if (savedPercentageFloat > 0.0) {
                ; Found a saved position in INI, restore it using helper function
                return RestoreWikipediaScrollPosition(savedPercentageFloat,
                    "Restoring previous scroll position... Please wait")
            } else {
                ; No saved position found in INI either
                ShowCenteredOverlay_Utils("⚠ No previous scroll position found", 1500, BANNER_ACCENT_ERROR)
                return false
            }
        } catch Error as err {
            ; Show brief message that no history exists
            ShowCenteredOverlay_Utils("⚠ No previous scroll position found", 1500, BANNER_ACCENT_ERROR)
            return false
        }
    }

    ; Get current URL to match with history
    try {
        url := GetWikipediaURLNormalized()
        ; #region agent log
        try {
            FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' . A_TickCount .
            ',"location":"Shift keys.ahk:3415","message":"Got normalized URL in restore","data":{"url":"' . url .
            '","urlLength":' . StrLen(url) . '},"sessionId":"debug-session","runId":"post-fix","hypothesisId":"F"}`n',
            DEBUG_LOG_PATH
        } catch {
        }
        ; #endregion
        if (url = "") {
            return false
        }

        ; Create UIA_Browser for checking scroll differences
        uia := UIA_Browser("ahk_exe chrome.exe")
    } catch Error as err {
        return false
    }

    ; Find the most recent previous position (not the current one)
    previousPosition := 0
    foundIndex := 0
    ; Search backwards through history to find a different position
    loop g_WikipediaScrollHistory.Length {
        idx := g_WikipediaScrollHistory.Length - A_Index + 1
        historyItem := g_WikipediaScrollHistory[idx]
        ; Check if this is a different position (different URL or different scroll percentage)
        if (historyItem.url != url) {
            ; Different article, use this one
            previousPosition := historyItem
            foundIndex := idx
            break
        } else {
            ; Same article, check if scroll position is different
            currentScroll := uia.JSReturnThroughClipboard("window.pageYOffset")
            docHeight := uia.JSReturnThroughClipboard("document.documentElement.scrollHeight")
            if (currentScroll != "" && currentScroll != "undefined" && currentScroll != "null" &&
                docHeight != "" && docHeight != "undefined" && docHeight != "null") {
                currentScrollFloat := Float(currentScroll)
                docHeightFloat := Float(docHeight)
                if (docHeightFloat > 0) {
                    currentPercentage := currentScrollFloat / docHeightFloat
                    diff := Abs(currentPercentage - historyItem.scrollPercentage)
                    ; If the saved percentage is different from current, use it
                    if (diff > 0.01) {
                        previousPosition := historyItem
                        foundIndex := idx
                        break
                    }
                }
            }
        }
    }

    if (!previousPosition) {
        ; No different position found in history
        ShowCenteredOverlay_Utils("⚠ No previous scroll position found", 1500, BANNER_ACCENT_ERROR)
        return false
    }

    ; Restore scroll position using helper function
    success := RestoreWikipediaScrollPosition(previousPosition.scrollPercentage,
        "Restoring previous scroll position... Please wait")

    ; Remove the restored position from history (since we just used it) and all positions after it
    if (success && foundIndex > 0 && foundIndex <= g_WikipediaScrollHistory.Length) {
        ; Remove from foundIndex to end
        loop (g_WikipediaScrollHistory.Length - foundIndex + 1) {
            g_WikipediaScrollHistory.Pop()
        }
    }
    return true
}

#HotIf

;-------------------------------------------------------------------
; Chrome PDF Viewer Shortcuts
;-------------------------------------------------------------------
IsChromePdfViewerActive() {
    ; Hard gate: avoid conflicts with non-Chrome apps
    if !WinActive("ahk_exe chrome.exe")
        return false

    try {
        ; #region agent log
        AgentDebugLog("H1", "IsChromePdfViewerActive_entry")
        ; #endregion
        uia := UIA_Browser("ahk_exe chrome.exe")

        ; Strong fingerprint: Chrome's built-in PDF viewer extension web area
        ; From UIA tree: chrome-extension://mhjfbmdgcfjbbpaeojofohoefgiehjai/index.html
        if (uia.FindElement({ Type: 50030, Value: "chrome-extension://mhjfbmdgcfjbbpaeojofohoefgiehjai", matchmode: "Substring" })) {
            ; #region agent log
            AgentDebugLog("H2", "IsChromePdfViewerActive_extension_match")
            ; #endregion
            return true
        }

        ; Fallback: stable, non-localized PDF toolbar controls
        if (uia.FindElement({ AutomationId: "pageSelector" }) && uia.FindElement({ AutomationId: "save" })) {
            ; #region agent log
            AgentDebugLog("H3", "IsChromePdfViewerActive_toolbar_match")
            ; #endregion
            return true
        }
    } catch {
    }

    ; #region agent log
    AgentDebugLog("H4", "IsChromePdfViewerActive_return_false")
    ; #endregion
    return false
}

#HotIf IsChromePdfViewerActive()

ChromePdf_GetViewerRoot(uia) {
    ; Prefer the extension's RootWebArea (most stable for the PDF viewer UI)
    root := 0
    try root := uia.FindElement({ Type: 50030, Value: "chrome-extension://mhjfbmdgcfjbbpaeojofohoefgiehjai",
        matchmode: "Substring" })
    if (root)
        return root

    ; Fallbacks
    try root := uia.GetCurrentDocumentElement()
    if (root)
        return root
    try root := uia.BrowserElement
    return root
}

ChromePdf_ClickByAutomationId(automationId, fallbackNames := 0) {
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 80

        root := ChromePdf_GetViewerRoot(uia)
        if (!root)
            return false

        btn := 0
        try btn := root.FindFirst({ Type: 50000, AutomationId: automationId })
        if (!btn)
            try btn := root.FindFirst({ AutomationId: automationId })

        if (!btn && IsObject(fallbackNames)) {
            for , name in fallbackNames {
                try btn := root.FindFirst({ Type: 50000, Name: name })
                if (btn)
                    break
            }
        }

        if (btn) {
            try btn.Invoke()
            catch {
                try btn.Click()
            }
            return true
        }
    } catch {
    }
    return false
}

ChromePdf_FocusByAutomationId(automationId, controlType := 0) {
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 80

        root := ChromePdf_GetViewerRoot(uia)
        if (!root)
            return false

        el := 0
        if (controlType) {
            try el := root.FindFirst({ Type: controlType, AutomationId: automationId })
        }
        if (!el)
            try el := root.FindFirst({ AutomationId: automationId })

        if (el) {
            try el.SetFocus()
            catch {
                try el.Click()
            }
            return true
        }
    } catch {
    }
    return false
}

ChromePdf_TogglePresentMode() {
    ; Deterministic primary path: open More actions and invoke Present item by selector.
    ; Legacy directional-key fallback remains optional behind feature flag.
    global USE_CHROME_PDF_PRESENT_FALLBACK

    if !ChromePdf_ClickByAutomationId("more", ["More actions"])
        return false

    deadline := A_TickCount + 700
    selectorNames := [
        "Present",
        "Presentation mode",
        "Present mode",
        "Apresentar",
        "Modo de apresentação"
    ]

    while (A_TickCount <= deadline) {
        try {
            uia := UIA_Browser("ahk_exe chrome.exe")
            root := ChromePdf_GetViewerRoot(uia)
            if (root) {
                presentItem := 0

                ; Prefer stable attributes first, then localized names.
                try presentItem := root.FindFirst({ Type: 50011, AutomationId: "present" })
                if (!presentItem)
                    try presentItem := root.FindFirst({ AutomationId: "present" })
                if (!presentItem)
                    try presentItem := root.FindFirst({ Type: 50000, AutomationId: "present" })

                if (!presentItem) {
                    for , candidateName in selectorNames {
                        try presentItem := root.FindFirst({ Type: 50011, Name: candidateName })
                        if (presentItem)
                            break
                        try presentItem := root.FindFirst({ Type: 50000, Name: candidateName })
                        if (presentItem)
                            break
                    }
                }

                if (presentItem) {
                    try presentItem.Invoke()
                    catch {
                        try presentItem.Click()
                    }
                    AgentDebugLog("H8", "ShiftE_presentmode_selector_success")
                    return true
                }
            }
        } catch {
        }
        Sleep 40
    }

    AgentDebugLog("H9", "ShiftE_presentmode_selector_timeout")

    if (USE_CHROME_PDF_PRESENT_FALLBACK) {
        Sleep 120
        Send "{Up}"
        Send "{Up}"
        Sleep 40
        Send "{Enter}"
        AgentDebugLog("H10", "ShiftE_presentmode_fallback_used")
        return true
    }

    return false
}

; Shift + F : Fit to page (Zoom to Fit) - Fit
+f::
{
    ; UIA tree: AutomationId "fit"
    ; #region agent log
    AgentDebugLog("H5", "ShiftF_handler_enter")
    ; #endregion
    result := ChromePdf_ClickByAutomationId("fit")
    if (result) {
        ; #region agent log
        AgentDebugLog("H6", "ShiftF_handler_click_success")
        ; #endregion
    } else {
        ; #region agent log
        AgentDebugLog("H7", "ShiftF_handler_click_failed")
        ; #endregion
    }
}

; Shift + P : Focus page number field - Page
+p::
{
    ; UIA tree: Edit AutomationId "pageSelector"
    ; Per requirement: focus only (no select-all)
    ChromePdf_FocusByAutomationId("pageSelector", 50004)
}

; Shift + T : Toggle thumbnails sidebar - Thumbnails
+t::
{
    ; UIA tree: AutomationId "sidenavToggle"
    ChromePdf_ClickByAutomationId("sidenavToggle")
}

; Shift + D : Download PDF - Download
+d::
{
    ; UIA tree: AutomationId "save" (button label is localized)
    ChromePdf_ClickByAutomationId("save", ["Baixar", "Download"])
}

; Shift + 2 : Two-page view (mnemonic: 2 = two pages)
+2::
{
    ; UIA: Button Type 50000, Name "More actions", AutomationId "more"
    if ChromePdf_ClickByAutomationId("more", ["More actions"]) {
        Sleep 150
        Send "{Down}"
        Sleep 50
        Send "{Enter}"
    }
}

; Shift + E : Present mode (mnemonic: E from prEsent)
+E::
{
    ChromePdf_TogglePresentMode()
}

#HotIf

;-------------------------------------------------------------------
; Mercado Livre (Brazil) Shortcuts
;-------------------------------------------------------------------
; Cache for IsMercadoLivreActive (per efficiency-canon: cache-first with validation).
; Invalidated when foreground HWND changes so we only run UIA once per window/tab focus.
global g_ML_CacheHwnd := 0
global g_ML_CacheResult := false
; Cache for initial-page-load workaround: right-click + close context menu to make hotkeys work (once per window).
global g_ML_ReceptivityHwnd := 0

; Workaround for ML: hotkeys fail on initial page load until page is "initialized". Right-click on empty area then close context menu.
; force=true: run even if cache says we already did (error-driven retry when hotkey action failed).
ML_EnsureHotkeyReceptivity(force := false) {
    global g_ML_ReceptivityHwnd
    if !WinActive("ahk_exe chrome.exe")
        return
    hwnd := WinExist("A")
    if (!hwnd)
        return
    if (!force && hwnd = g_ML_ReceptivityHwnd && WinExist("ahk_id " g_ML_ReceptivityHwnd))
        return
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        try
            root := uia.GetCurrentDocumentElement()
        catch
            root := uia.BrowserElement
        if (!root)
            return
        br := root.BoundingRectangle
        if (!br || (br.r <= br.l) || (br.b <= br.t))
            return
        x := br.l + (br.r - br.l) * 0.15
        y := br.t + (br.b - br.t) * 0.20
        prevMode := A_CoordModeMouse
        CoordMode("Mouse", "Screen")
        MouseClick("Right", x, y)
        SendEscape()
        CoordMode("Mouse", prevMode)
        g_ML_ReceptivityHwnd := hwnd
    } catch {
    }
}

IsMercadoLivreActive() {
    global g_ML_CacheHwnd, g_ML_CacheResult
    if !WinActive("ahk_exe chrome.exe")
        return false
    hwnd := WinExist("A")
    if (!hwnd)
        return false
    ; Cache hit: same window as last check (avoids UIA on every keystroke / cheat sheet open)
    if (hwnd = g_ML_CacheHwnd && WinExist("ahk_id " g_ML_CacheHwnd))
        return g_ML_CacheResult
    ; Platform identification by URL only (do not use window title; it changes to product name). See shopping uia3.md.
    ; URL check via UIA (address bar: Chrome exposes AcceleratorKey "Ctrl+L", not AccessKey). Bounded to this window only.
    try {
        root := UIA.ElementFromHandle(hwnd)
        addressBar := root.FindFirst({ Type: 50004, AcceleratorKey: "Ctrl+L" })
        if (addressBar) {
            url := addressBar.Value
            if InStr(url, "mercadolivre.com") || InStr(url, "mercadolibre.com") {
                g_ML_CacheHwnd := hwnd
                g_ML_CacheResult := true
                return true
            }
        }
    } catch {
        ; UIA failed; do not cache so next call retries
    }
    g_ML_CacheHwnd := hwnd
    g_ML_CacheResult := false
    return false
}

; Mercado Livre UIA helpers: get document root and find/invoke elements (bounded, no global state).
; When UIA_Browser init fails (document not ready), fallback: get Document from foreground window tree (Type 50030 = Document).
ML_GetDocRoot() {
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        try
            return uia.GetCurrentDocumentElement()
        catch
            return uia.BrowserElement
    } catch {
        ; Fallback: document not ready for UIA_Browser init (e.g. on first load). Get Document from active window.
        try {
            hwnd := WinExist("A")
            if (hwnd && WinActive("ahk_exe chrome.exe")) {
                root := UIA.ElementFromHandle(hwnd)
                doc := root.FindFirst({ Type: 50030 })
                if (doc)
                    return doc
            }
        } catch {
        }
        return 0
    }
}

; #region agent log - debug helper
ML_DebugLog(hypothesisId, message, detail, runId := "pre-fix") {
    ; Intentionally no-op.
    return
}
; #endregion agent log

; Locate the Mercado Livre Preço filter container, scoped to the block that
; contains the predefined price links (Até R$ 40 / R$ 40 a R$ 95 / Mais de R$ 95)
; and the Mínimo/Máximo edits. Returns the container element or 0.
ML_GetPriceFilter(root) {
    if !root
        return 0
    priceLinkNames := ["Até R$ 40", "R$ 40 a R$ 95", "Mais de R$ 95"]
    for _, name in priceLinkNames {
        link := ML_Find(root, { Type: 50005, Name: name, cs: false })
        if !link
            continue
        parent := link
        ; Walk up a few levels to find a container that has the Mínimo edit as a descendant.
        loop 6 {
            try parent := parent.Parent
            catch
                break
            if !parent
                break
            minimo := ML_Find(parent, { Type: 50004, Name: "Mínimo", cs: false })
            if (minimo) {
                return parent
            }
        }
    }
    return 0
}

; Best-effort helper to set text of an edit element. Tries ValuePattern first,
; then falls back to focus + clear + send keys.
ML_SetEditText(el, text) {
    if (!el)
        return false
    if (text = "")
        return false
    ok := false
    try {
        if el.GetPropertyValue(UIA.Property.IsValuePatternAvailable) {
            vp := el.ValuePattern
            vp.SetValue(text)
            ok := true
        }
    } catch {
        ML_DebugLog("C", "ValuePattern attempt failed", "exception in SetValue", "run1")
    }
    if (!ok) {
        try el.ScrollIntoView()
        catch {
        }
        try {
            el.SetFocus()
        } catch {
            try el.Click()
            catch {
            }
        }
        Sleep 50
        ML_DebugLog("C", "Focus acquired for edit", text, "run1")
        Send "^a{Del}"
        Sleep 30
        SendText(text)
        ok := true
    }
    return ok
}

; Try conditions in order; invoke or click first match. Returns true if invoked/clicked, false otherwise.
ML_FindAndInvoke(conditionList) {
    root := ML_GetDocRoot()
    if (!root)
        return false
    for cond in conditionList {
        try {
            el := root.FindElement(cond, UIA.TreeScope.Descendants)
            if (el) {
                try el.Invoke()
                catch {
                    try el.Click()
                    catch
                        return false
                }
                return true
            }
        } catch
            continue
    }
    return false
}

; Find single element by condition (Descendants). Returns element or 0.
ML_Find(root, condition) {
    try
        return root.FindElement(condition, UIA.TreeScope.Descendants)
    catch
        return 0
}

;-------------------------------------------------------------------
; Shopee (Brazil) detection and UIA helpers
;-------------------------------------------------------------------
; Cache for IsShopeeActive (same pattern as IsMercadoLivreActive).
global g_Shopee_CacheHwnd := 0
global g_Shopee_CacheResult := false

IsShopeeActive() {
    global g_Shopee_CacheHwnd, g_Shopee_CacheResult
    if !WinActive("ahk_exe chrome.exe")
        return false
    hwnd := WinExist("A")
    if (!hwnd)
        return false
    ; Cache hit: same window as last check (avoids UIA on every keystroke / cheat sheet open)
    if (hwnd = g_Shopee_CacheHwnd && WinExist("ahk_id " g_Shopee_CacheHwnd))
        return g_Shopee_CacheResult
    ; Platform identification by URL only (do not use window title; it changes to product name). See shopping uia3.md.
    ; URL check via UIA (Chrome address bar: AcceleratorKey "Ctrl+L", not AccessKey)
    try {
        root := UIA.ElementFromHandle(hwnd)
        addressBar := root.FindFirst({ Type: 50004, AcceleratorKey: "Ctrl+L" })
        if (addressBar) {
            url := addressBar.Value
            if InStr(url, "shopee.com", false) {
                g_Shopee_CacheHwnd := hwnd
                g_Shopee_CacheResult := true
                return true
            }
        }
    } catch {
        ; UIA failed; do not cache so next call retries
    }
    g_Shopee_CacheHwnd := hwnd
    g_Shopee_CacheResult := false
    return false
}

Shopee_GetDocRoot() {
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        try
            return uia.GetCurrentDocumentElement()
        catch
            return uia.BrowserElement
    } catch
        return 0
}

Shopee_Find(root, condition) {
    try
        return root.FindElement(condition, UIA.TreeScope.Descendants)
    catch
        return 0
}

Shopee_FindAndInvoke(conditionList) {
    root := Shopee_GetDocRoot()
    if (!root)
        return false
    for cond in conditionList {
        try {
            el := root.FindElement(cond, UIA.TreeScope.Descendants)
            if (el) {
                try el.Invoke()
                catch {
                    try el.Click()
                    catch
                        return false
                }
                return true
            }
        } catch
            continue
    }
    return false
}

; Navigate Shopee search results pagination by relative offset (e.g. +1 next, -1 previous).
Shopee_NavMove(offset) {
    if (offset = 0)
        return false
    hwnd := WinExist("A")
    if (!hwnd)
        return false

    ; Get current page index from Chrome address bar (?page=N), defaulting to 0 when absent.
    currentPage := 0
    try {
        rootWin := UIA.ElementFromHandle(hwnd)
        addressBar := rootWin.FindFirst({ Type: 50004, AcceleratorKey: "Ctrl+L" })
        if (addressBar) {
            url := addressBar.Value
            m := ""
            if RegExMatch(url, "i)[?&]page=(\d+)", &m) {
                try currentPage := Integer(m[1])
            }
        }
    } catch {
    }

    root := Shopee_GetDocRoot()
    if (!root)
        return false

    navGroup := Shopee_Find(root, { Type: 50026, Name: "Navegação entre páginas", cs: false })
    if (!navGroup)
        return false

    links := 0
    try links := navGroup.FindAll({ Type: 50005 })
    catch {
        return false
    }
    if (!links)
        return false

    bestEl := 0
    bestDelta := 0x7FFFFFFF

    for link in links {
        value := ""
        try value := link.Value
        catch {
            continue
        }
        if (value = "")
            continue
        m2 := ""
        if !RegExMatch(value, "i)[?&]page=(\d+)", &m2)
            continue
        targetPage := 0
        try targetPage := Integer(m2[1])
        catch {
            continue
        }

        ; Choose the nearest page ahead (offset>0) or behind (offset<0) relative to currentPage.
        if (offset > 0 && targetPage > currentPage) {
            delta := targetPage - currentPage
            if (delta < bestDelta) {
                bestDelta := delta
                bestEl := link
            }
        } else if (offset < 0 && targetPage < currentPage) {
            delta := currentPage - targetPage
            if (delta < bestDelta) {
                bestDelta := delta
                bestEl := link
            }
        }
    }

    if (!bestEl)
        return false

    try {
        bestEl.Invoke()
        return true
    } catch {
        try {
            bestEl.Click()
            return true
        } catch {
            return false
        }
    }
}

#HotIf IsMercadoLivreActive()

; Shift + S: Focus Mercado Livre search field
+s::
{
    ML_EnsureHotkeyReceptivity()
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 200

        ; Prefer the document root; fall back to browser element only if needed
        isDocRoot := false
        try {
            root := uia.GetCurrentDocumentElement()
            isDocRoot := true
        } catch {
            root := uia.BrowserElement
        }

        field := 0

        ; 1) Try AutomationId from the current document (or fallback root)
        try {
            field := root.FindElement({ AutomationId: "cb1-edit" })
        } catch {
        }

        ; 2) From the document root, try the numeric path 1,1,4,2 if available
        if (!field && isDocRoot) {
            try {
                field := root.ElementFromPath("1,1,4,2")
            } catch {
            }
        }

        ; 3) As a last resort, search from the browser element with Descendants scope
        if (!field) {
            try {
                field := uia.BrowserElement.FindElement({ AutomationId: "cb1-edit" }, UIA.TreeScope.Descendants)
            } catch {
            }
        }

        if (field) {
            focusOk := false
            try {
                field.SetFocus()
                focusOk := true
            } catch {
                try {
                    field.Click()
                    focusOk := true
                } catch {
                }
            }
            if (focusOk)
                return
        }

        ; Error-driven workaround: force right-click + close menu, then retry once
        ML_EnsureHotkeyReceptivity(true)
        Sleep 300

        field := 0
        try {
            try
                field := root.FindElement({ AutomationId: "cb1-edit" })
            catch
                field := 0
            if (!field && isDocRoot) {
                try
                    field := root.ElementFromPath("1,1,4,2")
                catch
                    field := 0
            }
            if (!field) {
                try
                    field := uia.BrowserElement.FindElement({ AutomationId: "cb1-edit" }, UIA.TreeScope.Descendants)
                catch
                    field := 0
            }
            if (field) {
                try
                    field.SetFocus()
                catch {
                    try
                        field.Click()
                    catch {
                    }
                }
                return
            }
        } catch {
        }

        MsgBox "Could not find Mercado Livre search field."
    } catch Error as e {
        MsgBox "An error occurred: " e.Message
    }
}

; Shift + C: Carrinho de compras (Cart)
+c::
{
    ML_EnsureHotkeyReceptivity()
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

; Shift + P: Compras (Purchases)
+p::
{
    ML_EnsureHotkeyReceptivity()
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

; Shift + Y: Chegará amanhã (filter toggle)
+y::
{
    ML_EnsureHotkeyReceptivity()
    if ML_FindAndInvoke([{ Type: 50000, AutomationId: "shipping_time_highlighted_nextday" }])
        return
    MsgBox "Filtro 'Chegará amanhã' não encontrado."
}

; Shift + F: Full (frete grátis Full)
+f::
{
    ML_EnsureHotkeyReceptivity()
    if ML_FindAndInvoke([{ Type: 50000, AutomationId: "shipping_highlighted_fulfillment" }])
        return
    MsgBox "Filtro 'Full' não encontrado."
}

; Shift + I: Compra Internacional
+i::
{
    ML_EnsureHotkeyReceptivity()
    if ML_FindAndInvoke([{ Type: 50000, AutomationId: "SHIPPING_ORIGIN_HIGHLIGHTED" }])
        return
    MsgBox "Filtro 'Internacional' não encontrado."
}

; Shift + N: Envio local / Produtos com frete nacional
+n::
{
    ML_EnsureHotkeyReceptivity()
    if ML_FindAndInvoke([{ Type: 50000, AutomationId: "SHIPPING_ORIGIN_LOCAL_HIGHLIGHTED" }, { Type: 50000, Name: "Envio local",
        cs: false, matchmode: "Substring" }])
        return
    MsgBox "Filtro 'Produtos com frete nacional' não encontrado."
}

; Shift + G: Frete grátis
+g::
{
    ML_EnsureHotkeyReceptivity()
    if ML_FindAndInvoke([{ Type: 50000, AutomationId: "shipping_cost_highlighted_free" }])
        return
    MsgBox "Filtro 'Frete grátis' não encontrado."
}

; Shift + O: Ordenar por – Handy-style menu and UIA execution
; Returns current option label from list selected item (Layer 1) or trigger (Layer 2). Single authority for quality-check reads.
ML_SortGetCurrentLabel(sortList, trigger) {
    currentLabel := ""
    if (sortList) {
        try {
            children := sortList.FindAll({ Type: 50007 })
            for ch in children {
                try {
                    if (ch.SelectionItemPattern.IsSelected) {
                        currentLabel := Trim(ch.Name)
                        if (currentLabel = "")
                            currentLabel := Trim(ch.Value)
                        if (currentLabel != "")
                            return currentLabel
                    }
                } catch {
                }
            }
        } catch {
        }
    }
    try {
        currentLabel := Trim(trigger.Name)
        if (currentLabel = "")
            currentLabel := Trim(trigger.Value)
    } catch {
    }
    return currentLabel
}

ML_SortClose() {
    try Hotkey("1", "Off")
    catch {
    }
    try Hotkey("2", "Off")
    catch {
    }
    try Hotkey("3", "Off")
    catch {
    }
    try Hotkey("Escape", ML_SortCancel, "Off")
    catch {
    }
    global g_ML_SortGui
    if (g_ML_SortGui && IsObject(g_ML_SortGui) && g_ML_SortGui.Hwnd)
        try g_ML_SortGui.Destroy()
    g_ML_SortGui := 0
}

ML_SortCancel(*) {
    ML_SortClose()
}

ML_SortSelect(idx) {
    ML_SortClose()
    ML_SortApply(idx)
}

ML_SortApply(idx) {
    StandardLoadingBar_Show("⏳ Ordenando...", BANNER_ACCENT_INTERMEDIATE, { passive: false, centerOnHwnd: 0,
        textWidth: 380, fontSize: 17 })
    try {
        root := ML_GetDocRoot()
        if (!root) {
            StandardLoadingBar_Update("❌ Página do Mercado Livre não disponível.")
            Sleep 800
            return
        }
        trigger := ML_Find(root, { Type: 50003, AutomationId: "5clcjae", matchmode: "Substring" })
        if (!trigger)
            trigger := ML_Find(root, { Type: 50003, AutomationId: "trigger", matchmode: "Substring" })
        if (!trigger) {
            StandardLoadingBar_Update("❌ Botão 'Ordenar por' não encontrado.")
            Sleep 800
            return
        }
        ; Current selection: menu opens with this option highlighted. Order in menu: 1=Mais relevantes, 2=Menor preço, 3=Maior preço
        current := 1
        label := ""
        try {
            label := Trim(trigger.Name)
            if (label = "")
                label := Trim(trigger.Value)
            if (InStr(label, "Maior preço"))
                current := 3
            else if (InStr(label, "Menor preço"))
                current := 2
            else if (InStr(label, "Mais relevantes"))
                current := 1
        } catch {
        }
        StandardLoadingBar_Update("⏳ Abrindo menu...")
        clickOk := false
        try {
            trigger.Click()
            clickOk := true
        } catch {
        }
        if (!clickOk) {
            StandardLoadingBar_Update("❌ Não foi possível abrir o menu.")
            Sleep 800
            return
        }
        ; Bounded wait for menu: poll for sortList up to 1.2s (canon: condition-based waits with timeout).
        searchRoot := root
        sortList := 0
        loop 12 {
            Sleep 100
            try {
                dropdownSibling := UIA.TreeWalkerTrue.TryGetNextSiblingElement(trigger)
                if (dropdownSibling)
                    sortList := ML_Find(dropdownSibling, { Type: 50008, AutomationId: "menu-list", matchmode: "Substring" })
            } catch {
            }
            if (sortList)
                break
            sortList := ML_Find(searchRoot, { Type: 50008, AutomationId: "5clcjae_-menu-list", matchmode: "Substring" })
            if (sortList)
                break
            sortList := ML_Find(searchRoot, { Type: 50008, AutomationId: "menu-list", matchmode: "Substring" })
            if (sortList)
                break
            try {
                triggerParent := trigger.Parent
                if (triggerParent)
                    sortList := ML_Find(triggerParent, { Type: 50008, AutomationId: "menu-list", matchmode: "Substring" })
            } catch {
            }
            if (sortList)
                break
            sortList := ML_Find(searchRoot, { AutomationId: "menu-list", matchmode: "Substring" })
            if (sortList)
                break
        }
        StandardLoadingBar_Update("⏳ Selecionando opção...")
        optionSubstrings := ["menu-list-option-relevance", "menu-list-option-price_asc", "menu-list-option-price_desc"]
        sub := optionSubstrings[idx]
        optionNames := ["Mais relevantes", "Menor preço", "Maior preço"]
        item := 0
        if (sortList) {
            item := ML_Find(sortList, { Type: 50007, AutomationId: sub, matchmode: "Substring" })
            if (!item)
                item := ML_Find(sortList, { Type: 50007, Name: optionNames[idx], cs: false })
        }
        if (!item) {
            item := ML_Find(searchRoot, { Type: 50007, AutomationId: sub, matchmode: "Substring" })
            if (!item)
                item := ML_Find(searchRoot, { Type: 50007, Name: optionNames[idx], cs: false })
        }
        if (item) {
            try item.Click()
            catch
                try item.Invoke()
            Sleep 150
            StandardLoadingBar_Update("✅ Ordenação aplicada")
            Sleep 300
        } else {
            ; Keyboard: navigate from current to target (delta), then confirm with Enter after label check.
            try WinActivate("ahk_exe chrome.exe")
            catch {
            }
            Sleep 400
            delta := idx - current
            if (delta > 0) {
                loop delta {
                    Send "{Down}"
                    Sleep 100
                }
            } else if (delta < 0) {
                loop (-delta) {
                    Send "{Up}"
                    Sleep 100
                }
            }
            Sleep 200
            targetName := optionNames[idx]
            currentLabel := ML_SortGetCurrentLabel(sortList, trigger)
            labelMatches := (currentLabel != "" && (InStr(currentLabel, targetName) || InStr(targetName, currentLabel)))
            if (labelMatches) {
                Send "{Enter}"
                StandardLoadingBar_Update("✅ Ordenação aplicada")
                Sleep 300
            } else {
                currentIdx := 1
                if (InStr(currentLabel, "Maior preço"))
                    currentIdx := 3
                else if (InStr(currentLabel, "Menor preço"))
                    currentIdx := 2
                else if (InStr(currentLabel, "Mais relevantes"))
                    currentIdx := 1
                if (idx > currentIdx) {
                    Send "{Down}"
                    Sleep 100
                } else if (idx < currentIdx) {
                    Send "{Up}"
                    Sleep 100
                }
                currentLabel := ML_SortGetCurrentLabel(sortList, trigger)
                labelMatches := (currentLabel != "" && (InStr(currentLabel, targetName) || InStr(targetName,
                    currentLabel)))
                Send "{Enter}"
                StandardLoadingBar_Update("✅ Ordenação aplicada")
                Sleep 300
            }
        }
    } catch Error as err {
        try StandardLoadingBar_Update("❌ Erro: " SubStr(err.Message, 1, 40))
        catch {
        }
        Sleep 700
    } finally {
        try StandardLoadingBar_Hide(0)
        catch {
        }
    }
}

+o::
{
    ML_EnsureHotkeyReceptivity()
    root := ML_GetDocRoot()
    if (!root) {
        ; Retry once after workaround and delay (document may not have been ready on first load)
        ML_EnsureHotkeyReceptivity(true)
        Sleep 400
        root := ML_GetDocRoot()
    }
    if (!root) {
        MsgBox "Página do Mercado Livre não disponível."
        return
    }
    global g_ML_SortGui
    g_ML_SortGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Owner")
    g_ML_SortGui.BackColor := "1E1E2E"
    g_ML_SortGui.MarginX := 20
    g_ML_SortGui.MarginY := 15
    g_ML_SortGui.SetFont("s14 cCDD6F4 Bold", "Segoe UI")
    g_ML_SortGui.Add("Text", "w280 Center", "Ordenar por")
    g_ML_SortGui.Add("Text", "w280 h1 Background45475A")
    g_ML_SortGui.SetFont("s12 cCDD6F4", "Segoe UI")
    g_ML_SortGui.Add("Text", "w280", "[1] Mais relevantes")
    g_ML_SortGui.SetFont("s9 c6C7086", "Segoe UI")
    g_ML_SortGui.Add("Text", "w280 y+2", "    Relevância da busca")
    g_ML_SortGui.SetFont("s12 cCDD6F4", "Segoe UI")
    g_ML_SortGui.Add("Text", "w280", "[2] Menor preço")
    g_ML_SortGui.SetFont("s9 c6C7086", "Segoe UI")
    g_ML_SortGui.Add("Text", "w280 y+2", "    Preço crescente")
    g_ML_SortGui.SetFont("s12 cCDD6F4", "Segoe UI")
    g_ML_SortGui.Add("Text", "w280", "[3] Maior preço")
    g_ML_SortGui.SetFont("s9 c6C7086", "Segoe UI")
    g_ML_SortGui.Add("Text", "w280 y+2", "    Preço decrescente")
    g_ML_SortGui.SetFont("s12 cCDD6F4", "Segoe UI")
    g_ML_SortGui.Add("Text", "w280 h1 Background45475A y+10")
    g_ML_SortGui.SetFont("s9 c6C7086", "Segoe UI")
    g_ML_SortGui.Add("Text", "w280 Center", "Press 1-3 | Esc to cancel")
    activeWin := 0
    try
        activeWin := WinGetID("A")
    catch
        activeWin := 0
    MonitorGetWorkArea(1, &monitorLeft, &monitorTop, &monitorRight, &monitorBottom)
    monitorWidth := monitorRight - monitorLeft
    monitorHeight := monitorBottom - monitorTop
    if (activeWin && activeWin != 0) {
        rect := Buffer(16, 0)
        if (DllCall("GetWindowRect", "ptr", activeWin, "ptr", rect)) {
            centerX := NumGet(rect, 0, "int") + (NumGet(rect, 8, "int") - NumGet(rect, 0, "int")) // 2
            centerY := NumGet(rect, 4, "int") + (NumGet(rect, 12, "int") - NumGet(rect, 4, "int")) // 2
            monitorCount := MonitorGetCount()
            loop monitorCount {
                idx := A_Index
                MonitorGetWorkArea(idx, &l, &t, &r, &b)
                if (centerX >= l && centerX <= r && centerY >= t && centerY <= b) {
                    monitorLeft := l
                    monitorTop := t
                    monitorRight := r
                    monitorBottom := b
                    monitorWidth := r - l
                    monitorHeight := b - t
                    break
                }
            }
        }
    }
    g_ML_SortGui.Show("AutoSize Hide")
    g_ML_SortGui.GetPos(&gx, &gy, &gw, &gh)
    cx := monitorLeft + (monitorWidth - gw) // 2
    cy := monitorTop + (monitorHeight - gh) // 2
    g_ML_SortGui.Show("x" . cx . " y" . cy . " NA")
    Hotkey("1", (*) => ML_SortSelect(1), "On")
    Hotkey("2", (*) => ML_SortSelect(2), "On")
    Hotkey("3", (*) => ML_SortSelect(3), "On")
    Hotkey("Escape", ML_SortCancel, "On")
}

; Shift + L: Paginação – Seguinte
+l::
{
    ML_EnsureHotkeyReceptivity()
    if ML_FindAndInvoke([{ Type: 50005, Name: "Seguinte", cs: false }, { Type: 50000, Name: "Seguinte", cs: false }])
        return
    MsgBox "Botão 'Seguinte' não encontrado."
}

; Shift + K: Paginação – Anterior
+k::
{
    ML_EnsureHotkeyReceptivity()
    if ML_FindAndInvoke([{ Type: 50005, Name: "Anterior", cs: false }, { Type: 50000, Name: "Anterior", cs: false }])
        return
    MsgBox "Botão 'Anterior' não encontrado."
}

; Shift + A: Adicionar ao carrinho (página do produto)
+a::
{
    ML_EnsureHotkeyReceptivity()
    if ML_FindAndInvoke([{ Type: 50000, Name: "Adicionar ao carrinho", cs: false }])
        return
    MsgBox "Botão 'Adicionar ao carrinho' não encontrado."
}

; Shift + V: Adicionar aos favoritos (coração)
+v::
{
    ML_EnsureHotkeyReceptivity()
    if ML_FindAndInvoke([{ Type: 50000, Name: "Adicionar aos favoritos", cs: false }, { Type: 50000, ClassName: "ui-pdp-bookmark",
        matchmode: "Substring" }])
        return
    MsgBox "Botão 'Adicionar aos favoritos' não encontrado."
}

; Shift + J: Continuar fluxo (Continuar a compra / Continuar / OK)
+j::
{
    ML_EnsureHotkeyReceptivity()
    conditions := [{ Type: 50005, Name: "Continuar a compra", cs: false }, { Type: 50000, Name: "Continuar", cs: false }, { Type: 50000,
        Name: "OK", cs: false }, { Type: 50000, AutomationId: "shipping_footer_confirm_button" }, { Type: 50005, Name: "Continuar",
            cs: false }, { Type: 50000, Name: "Seguinte", cs: false }
    ]
    if ML_FindAndInvoke(conditions)
        return
    MsgBox "Botão de continuar não encontrado."
}

;-------------------------------------------------------------------
; Shopee (Brazil) Shortcuts
;-------------------------------------------------------------------
#HotIf IsShopeeActive()

; Shift + S: Focus Shopee search field
+s::
{
    try {
        root := Shopee_GetDocRoot()
        if (!root) {
            MsgBox "Página da Shopee não disponível."
            return
        }

        field := 0
        ; Prefer the main search combo box
        try {
            field := root.FindElement({ Type: 50003, Name: "Buscar na Shopee" })
        } catch {
        }
        ; Fallback: any control with LocalizedType = "search"
        if (!field) {
            try field := root.FindElement({ LocalizedType: "search" })
        }
        ; Fallback: numeric path from document root (see shopping uia3.md)
        if (!field) {
            try field := root.ElementFromPathExist("1,1,2,3,1")
        }

        if (field) {
            try field.SetFocus()
            catch {
                try field.Click()
            }
            return
        }
        MsgBox "Campo de busca da Shopee não encontrado."
    } catch Error as e {
        MsgBox "Erro ao focar busca da Shopee: " e.Message
    }
}

; Shift + C: Carrinho de compras (cart)
+c::
{
    if Shopee_FindAndInvoke([{ Type: 50005, AutomationId: "cart_drawer_target_id" }, { Type: 50005, Name: "Carrinho",
        cs: false, matchmode: "Substring" }, { Type: 50000, Name: "Carrinho", cs: false, matchmode: "Substring" }
    ])
        return
    MsgBox "Link/botão de carrinho da Shopee não encontrado."
}

; Shift + P: Minhas compras / pedidos (speculative)
+p::
{
    if Shopee_FindAndInvoke([{ Type: 50005, Name: "Minhas compras", cs: false, matchmode: "Substring" }, { Type: 50005,
        Name: "Meus pedidos", cs: false, matchmode: "Substring" }, { Type: 50005, Name: "Pedidos", cs: false, matchmode: "Substring" }
    ])
        return
    MsgBox "Link de compras/pedidos da Shopee não encontrado (atalho especulativo)."
}

; Shift + Y: Entrega Rápida (Chegará amanhã analog)
+y::
{
    if Shopee_FindAndInvoke([{ Type: 50002, Name: "Entrega Rápida", cs: false, matchmode: "Substring" }])
        return
    MsgBox "Filtro 'Entrega Rápida' não encontrado (atalho especulativo)."
}

; Shift + F: Promoções / produtos com desconto (Full analog, speculative)
+f::
{
    if Shopee_FindAndInvoke([{ Type: 50002, Name: "Produtos com Desconto", cs: false, matchmode: "Substring" }])
        return
    MsgBox "Filtro de promoções/produtos com desconto não encontrado (atalho especulativo)."
}

; Shift + I: Compra internacional
+i::
{
    if Shopee_FindAndInvoke([{ Type: 50002, Name: "Internacional", cs: false }])
        return
    MsgBox "Filtro 'Internacional' da Shopee não encontrado."
}

; Shift + N: Envio nacional
+n::
{
    if Shopee_FindAndInvoke([{ Type: 50002, Name: "Nacional", cs: false }])
        return
    MsgBox "Filtro 'Nacional' da Shopee não encontrado."
}

; Shift + G: Frete grátis (speculative)
+g::
{
    if Shopee_FindAndInvoke([{ Type: 50020, Name: "Frete grátis", cs: false, matchmode: "Substring" }])
        return
    MsgBox "Indicador/controle de 'Frete grátis' não encontrado (atalho especulativo)."
}

; Shift + O: Ordenar por (open sort menu)
+o::
{
    if Shopee_FindAndInvoke([{ Type: 50000, Name: "Classificar por relevância", cs: false, matchmode: "Substring" }, { Type: 50000,
        Name: "Classificar por", cs: false, matchmode: "Substring" }])
        return
    MsgBox "Botão 'Classificar por' da Shopee não encontrado."
}

; Shift + R: Faixa de preço (focus min/max price edits)
+r::
{
    try {
        root := Shopee_GetDocRoot()
        if (!root) {
            MsgBox "Página da Shopee não disponível."
            return
        }

        minimo := 0
        try minimo := root.FindElement({ Type: 50004, Name: "Preço mínimo" })
        catch {
        }
        maximo := 0
        try maximo := root.FindElement({ Type: 50004, Name: "Preço máximo" })
        catch {
        }

        target := minimo ? minimo : maximo
        if (target) {
            try target.SetFocus()
            catch {
                try target.Click()
            }
            Sleep 50
            Send "^a"
            return
        }

        MsgBox "Campos de faixa de preço da Shopee não encontrados."
    } catch Error as e {
        MsgBox "Erro ao focar faixa de preço da Shopee: " e.Message
    }
}

; Shift + L: Paginação – Seguinte (results)
+l::
{
    if Shopee_NavMove(1)
        return
    MsgBox "Navegação 'Seguinte' da Shopee não encontrada ou não aplicável."
}

; Shift + K: Paginação – Anterior (results)
+k::
{
    if Shopee_NavMove(-1)
        return
    MsgBox "Navegação 'Anterior' da Shopee não encontrada ou não aplicável."
}

; Shift + A: Adicionar ao carrinho (página do produto)
+a::
{
    if Shopee_FindAndInvoke([{ Type: 50000, Name: "Adicionar Ao Carrinho", cs: false, matchmode: "Substring" }])
        return
    MsgBox "Botão 'Adicionar Ao Carrinho' não encontrado na página da Shopee."
}

; Shift + V: Favoritar (coração)
+v::
{
    if Shopee_FindAndInvoke([{ Type: 50000, Name: "Favoritar", cs: false, matchmode: "Substring" }])
        return
    MsgBox "Botão de favoritos da Shopee não encontrado."
}

; Shift + J: Continuar fluxo (Continuar / Fazer pedido)
+j::
{
    if Shopee_FindAndInvoke([{ Type: 50000, Name: "Continuar", cs: false, matchmode: "Substring" }, { Type: 50000, Name: "Fazer pedido",
        cs: false, matchmode: "Substring" }, { Type: 50000, Name: "OK", cs: false }
    ])
        return
    MsgBox "Botão de continuar/fazer pedido da Shopee não encontrado."
}

#HotIf

;-------------------------------------------------------------------
; Microsoft Teams Shortcuts (chat)
;-------------------------------------------------------------------
#HotIf IsTeamsChatActive()

; Shift + R : Reply - Reply
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

; Shift + U : View all unread items - Unread
+u::
{
    Send "^!u"
}

; Shift + P : Pin chat - Pin
+p::
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
    SendEscape()
    Sleep "200"
    Send "^1"
    Sleep "500"          ; 80 ms
    Send "^+{Home}"
}

; Shift + E : Edit message - Edit
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

; Shift + A : Attach file - Attach
+a::
{
    Send "!+o"
}

; Shift + H : Open history menu - History
+h::
{
    Send "^h"
}

; Shift + M : Mark as unread - Mark
+m::
{
    Send "^1"
    Sleep "220"
    Send("{AppsKey}")
    Sleep "220"
    Send "{Down}"
    Send "{Enter}"
}

; Shift + X : Unpin chat - Unpin
+x::
{
    Sleep "150"
    Send "^1"
    Sleep "100"
    Send("{AppsKey}")
    Sleep "100"
    Send "r"
    Send "{Enter}"
}

; Shift + C : Collapse all conversation folders - Collapse
+c::
{
    Send "!q"
}

; Shift + I : Activate/deactivate details panel - Info
+i::
{
    Send "!p"
}

; Shift + . : Detach current chat - Window
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

; Shift + V : Video call - Video
+v::
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

; Shift + T : Add participants - Team
+t::
{
    try {
        ; Show progress banner
        ShowSmallLoadingIndicator_ChatGPT("Adding participants...")

        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; First, find and click the "More chat options" button
        moreOptionsButton := root.FindFirst({ Name: "More chat options", Type: "50000", matchmode: "Substring" })

        if !moreOptionsButton {
            ; Show error banner
            ShowSmallLoadingIndicator_ChatGPT("Could not find more options button")
            SetTimer(() => HideSmallLoadingIndicator_ChatGPT(), -2000)
            return
        }

        moreOptionsButton.Click()
        Sleep 500  ; Wait for menu to open

        ; Now find and click the "View and add participants" button
        participantsButton := root.FindFirst({ Name: "View and add participants", Type: "50000", matchmode: "Substring" })

        if participantsButton {
            participantsButton.Click()
            Sleep 300
            Send "{Tab}"
            Sleep 300
            Send "{Enter}"
            ; Hide progress banner on success
            SetTimer(() => HideSmallLoadingIndicator_ChatGPT(), -1000)
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

; Shift + F : Fold chat sections - Fold
+f::
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

; Shift + O : Open home panel - Open
+o::
{
    try {
        root := UIA.ElementFromHandle(WinExist("A"))

        ; Step 1: Click the dynamic Mentions tree item (match by stable UIA fields, not Name)
        mentionsEl := root.FindFirst({
            Type: "50024",
            LocalizedType: "tree item",
            ClassName: "fui-TreeItem r15xhw3a ___1k0lugy f10bgyvd fy0dxe2 f1d2rq10"
        })
        if !mentionsEl {
            ; Fallback: keep the same stable attributes but use ControlType alias
            mentionsEl := root.FindFirst({
                ControlType: "TreeItem",
                LocalizedType: "tree item",
                ClassName: "fui-TreeItem r15xhw3a ___1k0lugy f10bgyvd fy0dxe2 f1d2rq10"
            })
        }
        if (mentionsEl) {
            mentionsEl.SetFocus()
            Sleep 50
            mentionsEl.Click()
        } else {
            MsgBox "Could not find the target Mentions tree item.", "Shift+O Error", "IconX"
            return
        }

        ; Step 2: Wait 200 milliseconds
        Sleep 300

        ; Step 3: Click the "Chat (Ctrl+1)" button
        chatBtn := root.FindFirst({ AutomationId: "3b64df9d-7e97-4d9c-ac5c-2e0a5d8e6f40", Type: "50000" })
        if !chatBtn {
            ; Fallback: try by Name
            chatBtn := root.FindFirst({ Name: "Chat (Ctrl+1)", Type: "50000" })
        }
        if !chatBtn {
            ; Fallback: try with ControlType
            chatBtn := root.FindFirst({ AutomationId: "3b64df9d-7e97-4d9c-ac5c-2e0a5d8e6f40", ControlType: "Button" })
        }
        if (chatBtn) {
            chatBtn.SetFocus()
            Sleep 50
            chatBtn.Click()
        } else {
            MsgBox "Could not find the 'Chat (Ctrl+1)' button.", "Shift+O Error", "IconX"
        }

    } catch Error as err {
        MsgBox "Error in Shift+O:`n" err.Message, "Shift+O Error", "IconX"
    }
}

; Shift + L : Like reaction - Like
+l::
{
    Send "{Enter}"
    Send "{Enter}"
    SendEscape()
}

; Shift + G : Heart reaction - Heart
+g::
{
    Send "{Enter}"
    Send "{Down}"
    Send "{Enter}"
    SendEscape()
}

; Shift + J : Laugh reaction - Laugh
+j::
{
    Send "{Enter}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
    SendEscape()
}

; Alt + 1 : Select 1st search result - Search
!1::
{
    Send "{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Enter}"
}

; Alt + 2 : Select 2nd search result - Search
!2::
{
    Send "{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Enter}"
}

; Alt + 3 : Select 3rd search result - Search
!3::
{
    Send "{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Enter}"
}

; Alt + 4 : Select 4th search result - Search
!4::
{
    Send "{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Enter}"
}

; Alt + 5 : Select 5th search result - Search
!5::
{
    Send "{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Enter}"
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
    && RegExMatch(WinGetTitle("A"), "i)(Appointment|Meeting|Event)")
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
    if RegExMatch(t, "i)(Appointment|Meeting|Event)")
        return false
    if RegExMatch(t, "i)Reminder")
        return false
    return true
}

#HotIf IsOutlookMainActive()

; Shift + G : Send to General - General
+G::
{
    Send "!5"
    Send "O"
    Send "{Home}"
    Send "00"
    Send "{Enter}"
}

; Shift + N : Send to Newsletter - Newsletter
+N::
{
    Send "!5"
    Send "O"
    Send "{Home}"
    Send "01"
    Send "{Enter}"
}

; Shift + I : Go to Inbox - Inbox
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

; Shift + S : Subject / Title - Subject
+S:: {
    if FocusOutlookField({ AutomationId: "4101" }) ; Subject
        return
    if FocusOutlookField({ Name: "Subject", ControlType: "Edit" })
        return
}

; Shift + T : Required / To - To
+T:: {
    if FocusOutlookField({ AutomationId: "4109" }) ; Required
        return
    if FocusOutlookField({ Name: "Required", ControlType: "Edit" })
        return
    if FocusOutlookField({ AutomationId: "4117" }) ; To
        return
    if FocusOutlookField({ Name: "To", ControlType: "Edit" })
        return
}

; Shift + B : Subject -> Body - Body
+B:: {
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

; Shift + F : Toggle Focused / Other - Focused
+F:: {                                  ; toggle Focused / Other
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

; Shift + D : Don't send any response - Don't send
+D::
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

; Shift + E : Send response - Send
+E::
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

; Shift+W : Calendar [W]eek view
+W:: {
    try {
        if !ClickOutlookByIdThenNameClass("WeeklyView", "Week", "NetUIRibbonButton", 50000)
            Send "^!3"
    } catch {
        Send "^!3"
    }
}

; Shift+O : Calendar m[O]nth view
+O:: {
    try {
        if !ClickOutlookByIdThenNameClass("MonthlyView", "Month", "NetUIRibbonButton", 50000)
            Send "^!4"
    } catch {
        Send "^!4"
    }
}

; Shift + P : Pop Out current item - Pop Out
+P:: {
    try {
        if !ClickOutlookByIdThenNameClass("", "Pop Out", "", 50000) {
            MsgBox("Couldn't find 'Pop Out' button.", "Outlook Pop Out", "IconX")
        }
    } catch Error as err {
        ShowErr(err)
    }
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

            ; If no pinned item found by keywords, try to find Desktop by first letter
            ; Portuguese: "Área de Trabalho" starts with "Á" or "a"
            ; English: "Desktop" starts with "D" or "d"
            if (!firstPinnedItem) {
                allTreeItems := navPane.FindAll({ Type: "TreeItem" })
                for item in allTreeItems {
                    try {
                        itemName := item.Name
                        ; Check if name starts with "a" or "Á" (Portuguese Desktop) or "d" or "D" (English Desktop)
                        ; Case-insensitive check
                        firstChar := SubStr(itemName, 1, 1)
                        if (firstChar = "a" || firstChar = "A" || firstChar = "Á" || firstChar = "á" ||
                            firstChar = "d" || firstChar = "D") {
                            ; Additional check: must be Desktop-related (not just any item starting with a/d)
                            if (InStr(itemName, "Desktop", false) || InStr(itemName, "Área de Trabalho", false) ||
                            InStr(itemName, "Trabalho", false)) {
                                firstPinnedItem := item
                                break
                            }
                        }
                    } catch {
                        ; Skip items without names
                    }
                }
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

; Shift + M : Toggle Mail / Calendar - Mail/Calendar
+M:: {
    try {
        root := UIA.ElementFromHandle(WinExist("A"))

        ; Find Mail and Calendar list items
        mailItem := root.FindFirst({ Name: "Mail", Type: "50007" })
        if !mailItem {
            mailItem := root.FindFirst({ Name: "Mail", ClassName: "NetUIListViewItem" })
        }

        calendarItem := root.FindFirst({ Name: "Calendar", Type: "50007" })
        if !calendarItem {
            calendarItem := root.FindFirst({ Name: "Calendar", ClassName: "NetUIListViewItem" })
        }

        ; Check which is selected and toggle
        if (mailItem && calendarItem) {
            try {
                isMailSelected := mailItem.IsSelected
                isCalendarSelected := calendarItem.IsSelected

                if (isMailSelected) {
                    calendarItem.SetFocus()
                    Sleep 50
                    calendarItem.Click()
                } else {
                    mailItem.SetFocus()
                    Sleep 50
                    mailItem.Click()
                }
            } catch Error as err {
                ; Fallback: if pattern check fails, try clicking Calendar
                calendarItem.SetFocus()
                Sleep 50
                calendarItem.Click()
            }
        } else {
            MsgBox "Could not find Mail or Calendar items.", "Outlook Toggle", "IconX"
        }
    } catch Error as err {
        MsgBox "Error toggling Mail/Calendar:`n" err.Message, "Outlook Toggle", "IconX"
    }
}

; Ctrl + 1 : Move down once and select (Command Palette)
^1:: {
    Send "{Down}"
    Send "{Enter}"
}

; Ctrl + 2 : Move down twice and select (Command Palette)
^2:: {
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Ctrl + 3 : Move down three times and select (Command Palette)
^3:: {
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Ctrl + 4 : Move down four times and select (Command Palette)
^4:: {
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Ctrl + 5 : Move down five times and select (Command Palette)
^5:: {
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Message inspector-specific hotkeys (Subject/To/DatePicker/Body)
#HotIf IsOutlookMessageActive()

; Shift + S : Subject / Title - Subject
+S:: {
    if FocusOutlookField({ AutomationId: "4101" }) ; Subject
        return
    if FocusOutlookField({ Name: "Subject", ControlType: "Edit" })
        return
}

; Shift + T : Required / To - To
+T:: {
    if FocusOutlookField({ AutomationId: "4109" }) ; Required
        return
    if FocusOutlookField({ Name: "Required", ControlType: "Edit" })
        return
    if FocusOutlookField({ AutomationId: "4117" }) ; To
        return
    if FocusOutlookField({ Name: "To", ControlType: "Edit" })
        return
}

; Shift + B : Body (Subject -> Body) - Body
+B:: {
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

; Ctrl + 1 : Move down once and select (Command Palette)
^1:: {
    Send "{Down}"
    Send "{Enter}"
}

; Ctrl + 2 : Move down twice and select (Command Palette)
^2:: {
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Ctrl + 3 : Move down three times and select (Command Palette)
^3:: {
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Ctrl + 4 : Move down four times and select (Command Palette)
^4:: {
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Ctrl + 5 : Move down five times and select (Command Palette)
^5:: {
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
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

; Shift + S : Start date (combo) - Start Date
+S:: {
    Outlook_ClickStartDate()
}

; Shift + P : Start date picker - Picker
+P:: {
    Outlook_ClickStartDatePicker()
}

; Shift + T : Start time (combo) - Time
+T:: {
    Outlook_ClickStartTime()
}

; Shift + E : End date (combo) - End Date
+E:: {
    Outlook_ClickEndDate()
}

; Shift + H : End time (combo) - End Hour
+H:: {
    Outlook_ClickEndTime()
}

; Shift + A : All day checkbox - All Day
+A:: {
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

; Shift + I : Title field - Title
+I:: {
    if FocusOutlookField({ AutomationId: "4100" }) ; Title
        return
    if FocusOutlookField({ Name: "Title", ControlType: "Edit" })
        return
}

; Shift + R : Required / To field - Required
+R:: {
    if FocusOutlookField({ AutomationId: "4109" }) ; Required
        return
    if FocusOutlookField({ Name: "Required", ControlType: "Edit" })
        return
}

; Shift + L : Location -> Body - Location
+L:: {
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

; Shift + B : Body (from Location) - Body
+B:: {
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

; Shift + C : Make Recurring - Recurring
+C:: {
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

; =============================================================================
; Outlook Appointment Configuration Palette
; Shows a grid of 24 letter-labeled squares for selecting appointment configurations
; Shift + . → Show palette (display only, no actions triggered yet)
; =============================================================================

; Global variables for Outlook Appointment palette
global g_OutlookPaletteActive := false
global g_OutlookPaletteGuis := []
global g_OutlookPaletteTimer := false
global g_OutlookPaletteSessionID := 0

; Letter mapping for Outlook Appointment palette (24 combinations)
; Format: [Letter, Status, All-day, Private, Reminder]
; Status: 1=Free, 2=Busy, 3=Out of office
; All-day: 1=Yes, 2=No
; Private: 1=Off, 2=On
; Reminder: 1=15min, 2=2days
global g_OutlookPaletteMapping := Map(
    "Q", { Status: 1, AllDay: 1, Private: 1, Reminder: 1 },  ; Free, All-day Yes, Private Off, 15min
    "W", { Status: 1, AllDay: 1, Private: 1, Reminder: 2 },  ; Free, All-day Yes, Private Off, 2days
    "E", { Status: 1, AllDay: 1, Private: 2, Reminder: 1 },  ; Free, All-day Yes, Private On, 15min
    "R", { Status: 1, AllDay: 1, Private: 2, Reminder: 2 },  ; Free, All-day Yes, Private On, 2days
    "A", { Status: 1, AllDay: 2, Private: 1, Reminder: 1 },  ; Free, All-day No, Private Off, 15min
    "S", { Status: 1, AllDay: 2, Private: 1, Reminder: 2 },  ; Free, All-day No, Private Off, 2days
    "D", { Status: 1, AllDay: 2, Private: 2, Reminder: 1 },  ; Free, All-day No, Private On, 15min
    "F", { Status: 1, AllDay: 2, Private: 2, Reminder: 2 },  ; Free, All-day No, Private On, 2days
    "Z", { Status: 2, AllDay: 1, Private: 1, Reminder: 1 },  ; Busy, All-day Yes, Private Off, 15min
    "X", { Status: 2, AllDay: 1, Private: 1, Reminder: 2 },  ; Busy, All-day Yes, Private Off, 2days
    "C", { Status: 2, AllDay: 1, Private: 2, Reminder: 1 },  ; Busy, All-day Yes, Private On, 15min
    "V", { Status: 2, AllDay: 1, Private: 2, Reminder: 2 },  ; Busy, All-day Yes, Private On, 2days
    "B", { Status: 2, AllDay: 2, Private: 1, Reminder: 1 },  ; Busy, All-day No, Private Off, 15min
    "N", { Status: 2, AllDay: 2, Private: 1, Reminder: 2 },  ; Busy, All-day No, Private Off, 2days
    "M", { Status: 2, AllDay: 2, Private: 2, Reminder: 1 },  ; Busy, All-day No, Private On, 15min
    ",", { Status: 2, AllDay: 2, Private: 2, Reminder: 2 },  ; Busy, All-day No, Private On, 2days
    "U", { Status: 3, AllDay: 1, Private: 1, Reminder: 1 },  ; Out of office, All-day Yes, Private Off, 15min
    "I", { Status: 3, AllDay: 1, Private: 1, Reminder: 2 },  ; Out of office, All-day Yes, Private Off, 2days
    "O", { Status: 3, AllDay: 1, Private: 2, Reminder: 1 },  ; Out of office, All-day Yes, Private On, 15min
    "P", { Status: 3, AllDay: 1, Private: 2, Reminder: 2 },  ; Out of office, All-day Yes, Private On, 2days
    "J", { Status: 3, AllDay: 2, Private: 1, Reminder: 1 },  ; Out of office, All-day No, Private Off, 15min
    "K", { Status: 3, AllDay: 2, Private: 1, Reminder: 2 },  ; Out of office, All-day No, Private Off, 2days
    "L", { Status: 3, AllDay: 2, Private: 2, Reminder: 1 },  ; Out of office, All-day No, Private On, 15min
    ";", { Status: 3, AllDay: 2, Private: 2, Reminder: 2 }   ; Out of office, All-day No, Private On, 2days
)

; Letters in display order (3 status groups, each with 2 rows × 4 columns)
global g_OutlookPaletteLetters := [
    ; Free status (row 1: All-day Yes, row 2: All-day No)
    "Q", "W", "E", "R",
    "A", "S", "D", "F",
    ; Busy status (row 1: All-day Yes, row 2: All-day No)
    "Z", "X", "C", "V",
    "B", "N", "M", ",",
    ; Out of office status (row 1: All-day Yes, row 2: All-day No)
    "U", "I", "O", "P",
    "J", "K", "L", ";"
]

; Timer handler for palette timeout
OutlookPaletteTimerHandler(sessionID) {
    global g_OutlookPaletteActive, g_OutlookPaletteSessionID, g_OutlookPaletteTimer
    if (sessionID != g_OutlookPaletteSessionID) {
        return
    }
    if (!g_OutlookPaletteActive) {
        g_OutlookPaletteTimer := false
        return
    }
    CleanupOutlookPalette()
    g_OutlookPaletteTimer := false
}

; Cleanup function for Outlook palette
CleanupOutlookPalette() {
    global g_OutlookPaletteGuis, g_OutlookPaletteActive
    g_OutlookPaletteActive := false
    for gui in g_OutlookPaletteGuis {
        try {
            if (IsObject(gui) && gui.Hwnd) {
                gui.Hide()
                gui.Destroy()
            }
        } catch {
            ; Silently ignore errors
        }
    }
    g_OutlookPaletteGuis := []
}

; Show Outlook Appointment palette
ShowOutlookAppointmentPalette() {
    global g_OutlookPaletteActive, g_OutlookPaletteGuis, g_OutlookPaletteLetters
    global g_OutlookPaletteTimer, g_OutlookPaletteSessionID

    ; Cleanup any existing palette
    if (g_OutlookPaletteActive) {
        CleanupOutlookPalette()
    }

    ; Increment session ID
    g_OutlookPaletteSessionID++
    g_OutlookPaletteActive := true

    ; Configuration
    squareSize := 40
    spacing := 8
    statusGroupSpacing := 20  ; Space between status groups
    rowsPerStatus := 2
    colsPerStatus := 4

    ; Get mouse position for palette placement
    MouseGetPos(&startX, &startY)

    ; Calculate positions for 3 status groups (each 2×4 grid)
    ; Layout: 3 groups side by side, each group is 2 rows × 4 columns
    guiArray := []
    statusGroupWidth := (squareSize * colsPerStatus) + (spacing * (colsPerStatus - 1))

    loop 3 {  ; 3 status groups
        statusIndex := A_Index
        groupOffsetX := (statusIndex - 1) * (statusGroupWidth + statusGroupSpacing)

        loop rowsPerStatus {  ; 2 rows per status
            rowIndex := A_Index
            loop colsPerStatus {  ; 4 columns per row
                colIndex := A_Index

                ; Calculate letter index in the flat array
                letterIndex := ((statusIndex - 1) * rowsPerStatus * colsPerStatus) +
                ((rowIndex - 1) * colsPerStatus) + colIndex

                if (letterIndex > g_OutlookPaletteLetters.Length) {
                    continue
                }

                letter := g_OutlookPaletteLetters[letterIndex]

                ; Calculate position
                squareX := startX + groupOffsetX + ((colIndex - 1) * (squareSize + spacing))
                squareY := startY + ((rowIndex - 1) * (squareSize + spacing))

                ; Create square GUI
                squareGui := Gui("+AlwaysOnTop -Caption +ToolWindow")
                squareGui.BackColor := "333333"  ; Dark gray background
                squareGui.SetFont("s10 Bold cFFFFFF", "Segoe UI")
                squareGui.MarginX := 0
                squareGui.MarginY := 0

                ; Add letter text
                letterText := squareGui.AddText("w" . squareSize . " h" . squareSize . " Center 0x201", letter)

                ; Calculate top-left position
                guiX := Round(squareX - squareSize / 2.0)
                guiY := Round(squareY - squareSize / 2.0)

                guiArray.Push({ gui: squareGui, x: guiX, y: guiY })
            }
        }
    }

    ; Position all GUIs while hidden
    for guiInfo in guiArray {
        guiInfo.gui.Show("x" . guiInfo.x . " y" . guiInfo.y . " w" . squareSize . " h" . squareSize . " NA Hide")
        WinSetTransparent(220, guiInfo.gui)  ; ~86% opacity
    }

    ; Show all GUIs simultaneously
    for guiInfo in guiArray {
        try {
            guiInfo.gui.Show("NA")
        } catch {
            try {
                guiInfo.gui.Show("x" . guiInfo.x . " y" . guiInfo.y . " w" . squareSize . " h" . squareSize . " NA")
            }
        }
        g_OutlookPaletteGuis.Push(guiInfo.gui)
    }

    ; Set timeout timer (5 seconds)
    timerHandler := () => OutlookPaletteTimerHandler(g_OutlookPaletteSessionID)
    g_OutlookPaletteTimer := timerHandler
    SetTimer(timerHandler, -5000)
}

; -----------------------------------------------------------------------------
; Outlook Appointment – Cascaded selection via dialogs (good UX, text-focused)
; Uses a 3-step flow:
;   1) Pick Private × All-day (4 options)
;   2) Pick Status (3 options)
;   3) Pick Reminder (2 options)
; Final result is shown as a clear text summary (no fields changed yet).
; -----------------------------------------------------------------------------

; Global variable to store Outlook Appointment selection choice
global g_OutlookAppointmentChoice := ""

; Cancel handler for Outlook Appointment selection dialogs
CancelOutlookOption(optionGui, *) {
    global g_OutlookAppointmentChoice
    g_OutlookAppointmentChoice := ""
    optionGui.Destroy()
}

; Auto-submit handler for Outlook Appointment selection dialogs
Outlook_OptionAutoSubmit(ctrl, optionsMap) {
    global g_OutlookAppointmentChoice
    currentValue := ctrl.Text
    if (currentValue != "" && IsInteger(currentValue)) {
        choice := Integer(currentValue)
        choiceStr := String(choice)
        if (optionsMap.Has(choiceStr)) {
            ; Store the choice in global variable
            g_OutlookAppointmentChoice := choiceStr
            ctrl.Gui.Destroy()
        }
    }
}

; Factory function to create auto-submit handler with captured optionsMap
CreateOutlookOptionHandler(optionsMap) {
    return (ctrl, *) => Outlook_OptionAutoSubmit(ctrl, optionsMap)
}

; Show selection dialog with immediate auto-submit on number entry (no Enter needed)
Outlook_SelectOptionByInputBox(title, basePrompt, optionsMap) {
    ; Build prompt text with better spacing and grouping
    prompt := basePrompt . "`n`n"
    validList := ""
    lastGroup := 0
    for key, opt in optionsMap {
        ; Add extra spacing between groups if this option has a Group property
        if (opt.HasProp("Group") && opt.Group != lastGroup && lastGroup > 0) {
            prompt .= "`n"  ; Add blank line between groups
        }
        if (opt.HasProp("Group")) {
            lastGroup := opt.Group
        }

        prompt .= key . ") " . opt.Label . "`n"
        if (validList != "")
            validList .= ", "
        validList .= key
    }
    prompt .= "`n`nType a number (" . validList . "):"

    ; Create GUI dialog
    try {
        optionGui := Gui("+AlwaysOnTop +ToolWindow", title)
        optionGui.SetFont("s10", "Segoe UI")
        optionGui.AddText("w480 Center", prompt)
        optionGui.AddEdit("w60 Center vOptionInput", "")

        ; Set up auto-submit handler using factory function to capture optionsMap
        handler := CreateOutlookOptionHandler(optionsMap)
        optionGui["OptionInput"].OnEvent("Change", handler)

        ; Add Cancel button
        cancelBtn := optionGui.AddButton("w80", "Cancel")
        cancelBtn.OnEvent("Click", CancelOutlookOption.Bind(optionGui))

        optionGui.Show("w500 h250")
        optionGui["OptionInput"].Focus()

        ; Wait for dialog to close
        WinWaitClose("ahk_id " optionGui.Hwnd)

        ; Retrieve the selected choice from global variable
        global g_OutlookAppointmentChoice
        choice := g_OutlookAppointmentChoice
        g_OutlookAppointmentChoice := ""  ; Clear for next use
        return choice
    } catch Error as e {
        MsgBox "Error in selection dialog: " . e.Message, title . " Error", "IconX"
        return ""
    }
}

; =============================================================================
; Outlook Appointment Control State Checking Functions (UIA-based)
; =============================================================================

Outlook_CheckPrivacyState() {
    try {
        win := WinExist("A")
        if (!win) {
            return ""
        }
        root := UIA.ElementFromHandle(win)
        Sleep 100  ; Allow UI to update

        ; Look for Private checkbox - typically has AutomationId or specific Name
        checkbox := root.FindFirst({ AutomationId: "4227", ControlType: "CheckBox" })
        if (!checkbox) {
            checkbox := root.FindFirst({ Name: "Private", ControlType: "CheckBox" })
        }

        if (checkbox) {
            ; Check if checkbox is checked
            isChecked := checkbox.GetCurrentPropertyValue(UIA.Property.ToggleToggleState)
            ; ToggleState: 0 = Off, 1 = On
            return (isChecked = 1) ? "On" : "Off"
        }
    } catch Error {
        ; Silently fail - return empty string
    }
    return ""
}

Outlook_CheckAllDayState() {
    try {
        win := WinExist("A")
        if (!win) {
            return ""
        }
        root := UIA.ElementFromHandle(win)
        Sleep 100  ; Allow UI to update

        checkbox := root.FindFirst({ AutomationId: "4226", ControlType: "CheckBox" })
        if (!checkbox) {
            checkbox := root.FindFirst({ Name: "All day", ControlType: "CheckBox" })
        }

        if (checkbox) {
            isChecked := checkbox.GetCurrentPropertyValue(UIA.Property.ToggleToggleState)
            return (isChecked = 1) ? "Yes" : "No"
        }
    } catch Error {
        ; Silently fail
    }
    return ""
}

Outlook_CheckStatusState() {
    try {
        win := WinExist("A")
        if (!win) {
            return ""
        }
        root := UIA.ElementFromHandle(win)
        Sleep 100  ; Allow UI to update

        ; Look for Status dropdown/button - may need to find by AutomationId or Name
        statusControl := root.FindFirst({ AutomationId: "4356", ControlType: "Button" })
        if (!statusControl) {
            statusControl := root.FindFirst({ Name: "Busy", ControlType: "Button" })
        }
        if (!statusControl) {
            ; Try to find any control with Status-related names
            statusControl := root.FindFirst({ Name: "Free", ControlType: "Button" })
        }

        if (statusControl) {
            ; Get the text/value of the status control
            statusText := statusControl.GetCurrentPropertyValue(UIA.Property.Name)
            if (InStr(statusText, "Free", false)) {
                return "Free"
            } else if (InStr(statusText, "Busy", false)) {
                return "Busy"
            } else if (InStr(statusText, "Out of office", false) || InStr(statusText, "Out of Office", false)) {
                return "Out of office"
            }
        }
    } catch Error {
        ; Silently fail
    }
    return ""
}

Outlook_CheckCategoryState() {
    try {
        win := WinExist("A")
        if (!win) {
            return ""
        }
        root := UIA.ElementFromHandle(win)
        Sleep 100  ; Allow UI to update

        ; Look for Category control - may be a button or dropdown
        categoryControl := root.FindFirst({ AutomationId: "4357", ControlType: "Button" })
        if (!categoryControl) {
            categoryControl := root.FindFirst({ Name: "Categorize", ControlType: "Button" })
        }

        if (categoryControl) {
            ; Try to get category text/value
            categoryText := categoryControl.GetCurrentPropertyValue(UIA.Property.Name)
            if (InStr(categoryText, "Important", false)) {
                return "Important"
            } else if (InStr(categoryText, "Personal", false)) {
                return "Personal"
            }
        }
    } catch Error {
        ; Silently fail
    }
    return ""
}

Outlook_CheckReminderState() {
    try {
        win := WinExist("A")
        if (!win) {
            return ""
        }
        root := UIA.ElementFromHandle(win)
        Sleep 100  ; Allow UI to update

        ; Look for Reminder dropdown/field
        reminderControl := root.FindFirst({ AutomationId: "4358", ControlType: "ComboBox" })
        if (!reminderControl) {
            reminderControl := root.FindFirst({ Name: "Reminder", ControlType: "ComboBox" })
        }
        if (!reminderControl) {
            reminderControl := root.FindFirst({ Name: "Reminder", ControlType: "Edit" })
        }

        if (reminderControl) {
            reminderText := reminderControl.GetCurrentPropertyValue(UIA.Property.Value)
            if (reminderText) {
                return reminderText
            }
        }
    } catch Error {
        ; Silently fail
    }
    return ""
}

; =============================================================================
; Outlook Appointment Control Application Functions
; =============================================================================

ApplyPrivacy(desiredState) {
    if (desiredState = "Off") {
        return  ; Do nothing if Private Off is desired
    }

    try {
        ; Check current state
        currentState := Outlook_CheckPrivacyState()
        if (currentState = "On") {
            return  ; Already set to On, skip
        }

        ; Ensure Outlook window is active
        if (!IsOutlookAppointmentActive()) {
            return
        }

        ; Apply Private On: Alt+7
        ; Use ! prefix for Alt key combination
        Send "!7"
        Sleep 200
    } catch Error {
        ; Silently continue
    }
}

ApplyAllDay(desiredState) {
    try {
        ; Check current state
        currentState := Outlook_CheckAllDayState()
        if ((desiredState = "No (timed)" && currentState = "No") ||
        (desiredState = "Yes" && currentState = "Yes")) {
            return  ; Already set correctly, skip
        }

        ; Ensure Outlook window is active
        if (!IsOutlookAppointmentActive()) {
            return
        }

        ; Use existing UIA checkbox click logic
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        checkbox := root.FindFirst({ AutomationId: "4226", ControlType: "CheckBox" })
        if (!checkbox) {
            checkbox := root.FindFirst({ Name: "All day", ControlType: "CheckBox" })
        }

        if (checkbox) {
            checkbox.Invoke()
            Sleep 200
        }
    } catch Error {
        ; Silently continue
    }
}

ApplyStatus(desiredState) {
    try {
        ; Check current state
        currentState := Outlook_CheckStatusState()
        if (currentState = desiredState) {
            return  ; Already set correctly, skip
        }

        ; Ensure Outlook window is active
        if (!IsOutlookAppointmentActive()) {
            return
        }

        ; Map status to first letter
        statusLetter := ""
        if (desiredState = "Free") {
            statusLetter := "F"
        } else if (desiredState = "Tentative") {
            statusLetter := "T"
        } else if (desiredState = "Busy") {
            statusLetter := "B"
        } else if (desiredState = "Out of office") {
            statusLetter := "O"
        } else {
            return  ; Unknown status
        }

        ; Sequence: Alt (down and up), then 5, then first letter, then Enter
        Send "{Alt down}{Alt up}"  ; Press Alt (down and up)
        Sleep 200  ; Wait for menu to open
        Send "5"  ; Press number 5
        Sleep 200  ; Wait for submenu
        Send statusLetter  ; Press first letter of status
        Sleep 200  ; Wait before confirming
        Send "{Enter}"  ; Confirm selection
        Sleep 200
    } catch Error {
        ; Silently continue
    }
}

ApplyCategory(desiredState) {
    try {
        ; Check current state
        currentState := Outlook_CheckCategoryState()
        if (currentState = desiredState) {
            return  ; Already set correctly, skip
        }

        ; Ensure Outlook window is active
        if (!IsOutlookAppointmentActive()) {
            return
        }

        ; Open Category menu: Alt+6
        Send "!6"
        Sleep 300  ; Wait for menu to open

        ; Navigate based on desired state
        if (desiredState = "Important") {
            ; Down 1-4 times, use 2 as default
            Send "{Down}"
            Sleep 200
            Send "{Down}"
            Sleep 200
            Send "{Down}"
            Sleep 200
            Send "{Down}"
        } else if (desiredState = "Personal") {
            ; Down 6 times
            loop 6 {
                Send "{Down}"
                Sleep 200
            }
        }

        Sleep 200  ; Wait before confirming
        Send "{Enter}"
        Sleep 200
    } catch Error {
        ; Silently continue
    }
}

ApplyReminder(desiredState) {
    try {
        ; Check current state
        currentState := Outlook_CheckReminderState()
        ; Normalize current state for comparison
        normalizedCurrent := ""
        if (InStr(currentState, "4 hours", false) || InStr(currentState, "4 hour", false)) {
            normalizedCurrent := "4 hours"
        } else if (InStr(currentState, "1 day", false) || InStr(currentState, "one day", false)) {
            normalizedCurrent := "1 day"
        } else if (InStr(currentState, "4 days", false) || InStr(currentState, "four days", false)) {
            normalizedCurrent := "4 days"
        } else if (InStr(currentState, "1 week", false) || InStr(currentState, "one week", false)) {
            normalizedCurrent := "1 week"
        }

        if (normalizedCurrent = desiredState) {
            return  ; Already set correctly, skip
        }

        ; Ensure Outlook window is active
        if (!IsOutlookAppointmentActive()) {
            return
        }

        ; Open Reminder field: Alt+8
        Send "!8"
        Sleep 200

        ; Clear existing text and type new value
        Send "^a"  ; Select all
        Sleep 100
        SendText desiredState  ; Type the reminder text
        Sleep 200
        Send "{Enter}"
        Sleep 200
    } catch Error {
        ; Silently continue
    }
}

; =============================================================================
; Main Function: Apply All Outlook Appointment Settings
; =============================================================================

ApplyOutlookAppointmentSettings(privacy, allDay, status, category, reminder) {
    ; Ensure Outlook appointment window is active before applying settings
    foundWindow := false
    targetHwnd := 0

    ; First, check if current window is an Outlook appointment/meeting/event
    currentHwnd := WinExist("A")
    if (currentHwnd) {
        currentTitle := WinGetTitle("A")
        if (WinGetProcessName("A") = "OUTLOOK.EXE" && RegExMatch(currentTitle, "i)(Appointment|Meeting|Event)")) {
            targetHwnd := currentHwnd
            foundWindow := true
        }
    }

    ; If not found in current window, search all Outlook windows
    if (!foundWindow) {
        for hwnd in WinGetList("ahk_exe OUTLOOK.EXE") {
            title := WinGetTitle("ahk_id " hwnd)
            ; Match Appointment, Meeting, or Event windows (including "Untitled - Event")
            if RegExMatch(title, "i)(Appointment|Meeting|Event)") {
                targetHwnd := hwnd
                foundWindow := true
                break
            }
        }
    }

    ; If no window found, show error and exit
    if (!foundWindow || !targetHwnd) {
        MsgBox "Outlook appointment/meeting/event window not found. Please open an appointment window first.", "Error",
            "IconX"
        return
    }

    ; Forcefully activate the window
    if (!WinExist("ahk_id " targetHwnd)) {
        ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
        return
    }
    WinActivate("ahk_id " targetHwnd)
    WinShow("ahk_id " targetHwnd)  ; Ensure window is visible
    WinRestore("ahk_id " targetHwnd)  ; Restore if minimized
    ; Wait for window to become active
    WinWaitActive("ahk_id " targetHwnd, , 2)
    Sleep 300  ; Additional wait for stability and focus

    ; Show loading banner
    ShowSmallLoadingIndicator_ChatGPT("Applying settings...")

    try {
        ; Apply each setting in sequence
        ApplyPrivacy(privacy.Private)
        Sleep 100

        ApplyAllDay(allDay.AllDay)
        Sleep 100

        ApplyStatus(status.Status)
        Sleep 100

        ; If user chose to skip category, do nothing
        if (category.Category != "") {
            ApplyCategory(category.Category)
            Sleep 100
        }

        ApplyReminder(reminder.Reminder)

        ; Update loading banner to show completion
        ShowSmallLoadingIndicator_ChatGPT("Settings applied")
        Sleep 1000
    } catch Error as e {
        ShowSmallLoadingIndicator_ChatGPT("Error applying settings")
        Sleep 1000
    } finally {
        ; Always hide loading banner
        HideSmallLoadingIndicator_ChatGPT()
    }
}

RunOutlookAppointmentWizard() {
    ; Always find and activate Outlook appointment window before starting wizard
    foundWindow := false
    targetHwnd := 0

    ; First, check if current window is an Outlook appointment/meeting/event
    currentHwnd := WinExist("A")
    if (currentHwnd) {
        currentTitle := WinGetTitle("A")
        if (WinGetProcessName("A") = "OUTLOOK.EXE" && RegExMatch(currentTitle, "i)(Appointment|Meeting|Event)")) {
            targetHwnd := currentHwnd
            foundWindow := true
        }
    }

    ; If not found in current window, search all Outlook windows
    if (!foundWindow) {
        for hwnd in WinGetList("ahk_exe OUTLOOK.EXE") {
            title := WinGetTitle("ahk_id " hwnd)
            ; Match Appointment, Meeting, or Event windows (including "Untitled - Event")
            if RegExMatch(title, "i)(Appointment|Meeting|Event)") {
                targetHwnd := hwnd
                foundWindow := true
                break
            }
        }
    }

    ; If no window found, show error and exit
    if (!foundWindow || !targetHwnd) {
        MsgBox "Outlook appointment/meeting/event window not found. Please open an appointment window first.",
            "Wizard Error", "IconX"
        return
    }

    ; Forcefully activate the window
    if (!WinExist("ahk_id " targetHwnd)) {
        ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
        return
    }
    WinActivate("ahk_id " targetHwnd)
    WinShow("ahk_id " targetHwnd)  ; Ensure window is visible
    WinRestore("ahk_id " targetHwnd)  ; Restore if minimized
    ; Wait for window to become active
    WinWaitActive("ahk_id " targetHwnd, , 2)
    Sleep 300  ; Additional wait for stability and focus

    ; STEP 1 – Status (4 options)
    step1Options := Map()
    step1Options["1"] := { Label: "🟢 Free", Status: "Free" }
    step1Options["2"] := { Label: "🟡 Tentative", Status: "Tentative" }
    step1Options["3"] := { Label: "🔴 Busy", Status: "Busy" }
    step1Options["4"] := { Label: "🔴 Out of office", Status: "Out of office" }

    choice1 := Outlook_SelectOptionByInputBox(
        "📅 Outlook Appointment – Step 1 of 6",
        "Choose status:",
        step1Options
    )
    if (choice1 = "") {
        return  ; user cancelled
    }
    selStatus := step1Options[choice1]

    ; STEP 2 – Privacy (2 options)
    step2Options := Map()
    step2Options["1"] := { Label: "🔓 Private OFF", Private: "Off" }
    step2Options["2"] := { Label: "🔒 Private ON", Private: "On" }

    choice2 := Outlook_SelectOptionByInputBox(
        "📅 Outlook Appointment – Step 2 of 6",
        "Choose privacy:",
        step2Options
    )
    if (choice2 = "") {
        return
    }
    selPrivacy := step2Options[choice2]

    ; STEP 3 – All-day (2 options)
    step3Options := Map()
    step3Options["1"] := { Label: "⏰ All-day NO (timed)", AllDay: "No (timed)" }
    step3Options["2"] := { Label: "📅 All-day YES", AllDay: "Yes" }

    choice3 := Outlook_SelectOptionByInputBox(
        "📅 Outlook Appointment – Step 3 of 6",
        "Choose duration:",
        step3Options
    )
    if (choice3 = "") {
        return
    }
    selAllDay := step3Options[choice3]

    ; STEP 4 – Category (3 options, including none)
    step4Options := Map()
    step4Options["1"] := { Label: "🚫 No category", Category: "" }
    step4Options["2"] := { Label: "⭐ Important", Category: "Important" }
    step4Options["3"] := { Label: "👤 Personal", Category: "Personal" }

    choice4 := Outlook_SelectOptionByInputBox(
        "📅 Outlook Appointment – Step 4 of 6",
        "Choose category:",
        step4Options
    )
    if (choice4 = "") {
        return
    }
    selCategory := step4Options[choice4]

    ; STEP 5 – Reminder (7 options)
    step5Options := Map()
    step5Options["0"] := { Label: "🔕 None", Reminder: "0 minutes" }
    step5Options["1"] := { Label: "⏰ 15 minutes", Reminder: "15 minutes" }
    step5Options["2"] := { Label: "⏰ 4 hours", Reminder: "4 hours" }
    step5Options["3"] := { Label: "🗓️ 1 day", Reminder: "1 day" }
    step5Options["4"] := { Label: "📆 2 days", Reminder: "2 days" }
    step5Options["5"] := { Label: "📅 1 week", Reminder: "1 week" }
    step5Options["6"] := { Label: "📅 2 weeks", Reminder: "2 weeks" }

    choice5 := Outlook_SelectOptionByInputBox(
        "📅 Outlook Appointment – Step 5 of 6",
        "Choose reminder:",
        step5Options
    )
    if (choice5 = "") {
        return
    }
    selReminder := step5Options[choice5]

    ; STEP 6 – Note marker (boolean)
    step6Options := Map()
    step6Options["1"] := { Label: "📝 Mark as NOTE", IsNote: true }
    step6Options["2"] := { Label: "✖️ No note", IsNote: false }

    choice6 := Outlook_SelectOptionByInputBox(
        "📅 Outlook Appointment – Step 6 of 6",
        "Is this a note?",
        step6Options
    )
    if (choice6 = "") {
        return
    }
    selNote := step6Options[choice6]

    ; Apply all settings at the end of the wizard
    ApplyOutlookAppointmentSettings(selPrivacy, selAllDay, selStatus, selCategory, selReminder)

    ; If flagged as note, append NOTE at cursor (title should already be focused)
    if (selNote.IsNote) {
        try {
            if (IsOutlookAppointmentActive()) {
                SendText " NOTE"
            }
        } catch Error {
            ; Silently ignore failures
        }
    }
}

; Shift + w → Cascaded text wizard for Outlook Appointment
+w:: {
    if (!IsOutlookAppointmentActive()) {
        return
    }
    RunOutlookAppointmentWizard()
}

; Ctrl + 1 : Move down once and select (Command Palette)
^1:: {
    Send "{Down}"
    Send "{Enter}"
}

; Ctrl + 2 : Move down twice and select (Command Palette)
^2:: {
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Ctrl + 3 : Move down three times and select (Command Palette)
^3:: {
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Ctrl + 4 : Move down four times and select (Command Palette)
^4:: {
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Ctrl + 5 : Move down five times and select (Command Palette)
^5:: {
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

#HotIf

;-------------------------------------------------------------------
; Google Chrome Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe chrome.exe")

; Shift + W : Pop current tab to new window - Window
+w::
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

; Function to rename ChatGPT window (can be called directly or via hotkey)
RenameChatGPTWindowToChatGPT() {
    try {
        ; Show banner to inform user
        ShowSmallLoadingIndicator_ChatGPT("Renaming ChatGPT window...")

        ; Send F5 to refresh the page
        Send "{F5}"
        Sleep 5000 ; Wait for page refresh

        ; Get the active Chrome window
        chatGPTHwnd := WinExist("A")
        if !chatGPTHwnd {
            HideSmallLoadingIndicator_ChatGPT()
            return
        }

        ; Get UIA browser context for the active Chrome window
        cUIA := UIA_Browser("ahk_id " chatGPTHwnd)
        if !cUIA {
            HideSmallLoadingIndicator_ChatGPT()
            return
        }

        Sleep 200 ; Give UIA time to attach

        ; Get root element (prefer document, fallback to browser root)
        try {
            root := cUIA.GetCurrentDocumentElement()
        } catch {
            root := cUIA.BrowserElement
        }
        if !root {
            HideSmallLoadingIndicator_ChatGPT()
            MsgBox "Failed to get root element", "ChatGPT", "IconX"
            return
        }

        ; Step 0: Ensure sidebar is open (required for "Seus chats" to be visible)
        ; Check if sidebar is open by looking for close sidebar button (Portuguese or English)
        sidebarCloseButton := 0
        sidebarCloseNames := ["Fechar barra lateral", "Close sidebar"]
        for name in sidebarCloseNames {
            try {
                sidebarCloseButton := root.FindElement({ Type: 50000, Name: name, cs: false })
                if (sidebarCloseButton)
                    break
            } catch {
                try {
                    sidebarCloseButton := root.FindElement({ Type: 50000, Name: name })
                    if (sidebarCloseButton)
                        break
                } catch {
                }
            }
        }

        ; If sidebar is not open (button not found), open it using keyboard shortcut
        if (!sidebarCloseButton) {
            ; Try to open sidebar with Ctrl+Shift+S
            Send "^+s"
            Sleep 500 ; Wait for sidebar to open

            ; Verify sidebar is now open by checking for the close button again
            for name in sidebarCloseNames {
                try {
                    sidebarCloseButton := root.FindElement({ Type: 50000, Name: name, cs: false })
                    if (sidebarCloseButton)
                        break
                } catch {
                    try {
                        sidebarCloseButton := root.FindElement({ Type: 50000, Name: name })
                        if (sidebarCloseButton)
                            break
                    } catch {
                    }
                }
            }

            ; If still not found, wait a bit more and try one more time
            if (!sidebarCloseButton) {
                Sleep 500
                for name in sidebarCloseNames {
                    try {
                        sidebarCloseButton := root.FindElement({ Type: 50000, Name: name, cs: false })
                        if (sidebarCloseButton)
                            break
                    } catch {
                    }
                }
            }
        }

        Sleep 1000 ; Wait for sidebar to open

        ; Step 1: Locate the chat button (Type: 50000, Name: "Seus chats" or "Your chats")
        chatButton := 0
        chatButtonNames := ["Seus chats", "Your chats", "Chats"]
        for name in chatButtonNames {
            try {
                chatButton := root.FindElement({ Type: 50000, Name: name, cs: false })
                if (chatButton)
                    break
            } catch {
                try {
                    chatButton := root.FindElement({ Type: 50000, Name: name })
                    if (chatButton)
                        break
                } catch {
                }
            }
        }

        if !chatButton {
            HideSmallLoadingIndicator_ChatGPT()
            MsgBox "Failed to find chat button (tried: Seus chats, Your chats, Chats)", "ChatGPT", "IconX"
            return
        }

        ; Step 2: Get the sibling element (next sibling of chat button)
        siblingElement := UIA.TreeWalkerTrue.TryGetNextSiblingElement(chatButton)
        if !siblingElement {
            HideSmallLoadingIndicator_ChatGPT()
            MsgBox "Failed to find sibling element of chat button", "ChatGPT", "IconX"
            return
        }

        ; Step 2.5: Check if sibling element supports ExpandCollapse pattern and expand it if collapsed
        try {
            hasExpandPattern := siblingElement.GetPropertyValue(UIA.Property.IsExpandCollapsePatternAvailable)
            if (hasExpandPattern) {
                expandPattern := siblingElement.ExpandCollapsePattern
                expandState := expandPattern.ExpandCollapseState

                ; If collapsed, expand it
                if (expandState == UIA.ExpandCollapseState.Collapsed) {
                    expandPattern.Expand()
                    Sleep 300 ; Wait for expansion to complete
                } else if (expandState == UIA.ExpandCollapseState.PartiallyExpanded) {
                    ; If partially expanded, try to expand it fully
                    expandPattern.Expand()
                    Sleep 300
                }
            }
        } catch {
            ; Continue even if expand fails - element might not need expansion
        }

        ; Step 3: Find the OpenConversationOptions button directly using its known properties
        ; Button: Type 50000, Name "Abrir opções de conversa" (PT) or "Open conversation options" (EN), AutomationId "radix-_r_b6_", ClassName "__menu-item-trailing-btn"
        openConversationButton := 0
        conversationOptionNames := ["Abrir opções de conversa", "Abrir opções da conversa", "Open conversation options",
            "Conversation options", "Open options"]

        ; Try 1: Find by Name and Type (most reliable) - try both Portuguese and English
        for name in conversationOptionNames {
            try {
                openConversationButton := siblingElement.FindElement({ Type: 50000, Name: name, cs: false },
                UIA.TreeScope.Descendants)
                if (openConversationButton)
                    break
            } catch {
                try {
                    openConversationButton := siblingElement.FindElement({ Type: 50000, Name: name },
                    UIA.TreeScope.Descendants)
                    if (openConversationButton)
                        break
                } catch {
                }
            }
        }

        ; Try 2: Find by AutomationId (if Name search fails)
        if (!openConversationButton) {
            try {
                openConversationButton := siblingElement.FindElement({ Type: 50000, AutomationId: "radix-_r_b6_" }, UIA
                .TreeScope.Descendants)
            } catch {
            }
        }

        ; Try 3: Find by ClassName (if both above fail)
        if (!openConversationButton) {
            try {
                openConversationButton := siblingElement.FindElement({ Type: 50000, ClassName: "__menu-item-trailing-btn" },
                UIA.TreeScope.Descendants)
            } catch {
            }
        }

        ; Try 4: Fallback to first child button (if specific search fails)
        if (!openConversationButton) {
            try {
                openConversationButton := UIA.TreeWalkerTrue.TryGetFirstChildElement(siblingElement)
                ; Verify it's actually a button
                if (openConversationButton && openConversationButton.Type != 50000) {
                    openConversationButton := 0
                }
            } catch {
            }
        }

        if !openConversationButton {
            HideSmallLoadingIndicator_ChatGPT()
            MsgBox "Failed to find OpenConversationOptions button (tried: Abrir opções de conversa, Open conversation options, etc.)",
                "ChatGPT", "IconX"
            return
        }

        ; Step 4: Click the OpenConversationOptions button
        ; Check if button is enabled and visible
        try {
            if (openConversationButton.GetPropertyValue(UIA.Property.IsOffscreen)) {
                HideSmallLoadingIndicator_ChatGPT()
                MsgBox "OpenConversationOptions button is offscreen", "ChatGPT", "IconX"
                return
            }
            if (!openConversationButton.GetPropertyValue(UIA.Property.IsEnabled)) {
                HideSmallLoadingIndicator_ChatGPT()
                MsgBox "OpenConversationOptions button is disabled", "ChatGPT", "IconX"
                return
            }
        } catch {
            ; Continue even if property check fails
        }

        ; Try multiple click strategies in order of preference
        clicked := false

        ; Strategy 1: Try Invoke pattern (most reliable for buttons)
        try {
            openConversationButton.Invoke()
            clicked := true
        } catch {
        }

        ; Strategy 2: Try SetFocus then Click
        if (!clicked) {
            try {
                openConversationButton.SetFocus()
                Sleep 50
                openConversationButton.Click()
                clicked := true
            } catch {
            }
        }

        ; Strategy 3: Force coordinate-based click using "left" parameter
        if (!clicked) {
            try {
                openConversationButton.Click("left")
                clicked := true
            } catch {
            }
        }

        ; Strategy 4: Direct coordinate click using element Location
        if (!clicked) {
            try {
                pos := openConversationButton.Location
                if (pos && pos.w > 0 && pos.h > 0) {
                    ; Activate window first
                    if (!WinExist("ahk_id " chatGPTHwnd)) {
                        ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
                        return
                    }
                    WinActivate("ahk_id " chatGPTHwnd)
                    WinWaitActive("ahk_id " chatGPTHwnd, , 1)
                    Sleep 100

                    ; Save current mouse position
                    MouseGetPos(&prevX, &prevY)

                    ; Click at center of element
                    CoordMode("Mouse", "Screen")
                    Click(pos.x + pos.w // 2, pos.y + pos.h // 2)
                    Sleep 50

                    ; Restore mouse position
                    MouseMove(prevX, prevY)
                    clicked := true
                }
            } catch {
            }
        }

        if (!clicked) {
            HideSmallLoadingIndicator_ChatGPT()
            MsgBox "Failed to click OpenConversationOptions button (all methods failed)", "ChatGPT", "IconX"
            return
        }

        ; After clicking the button, send DownArrow three times, type "ChatGPT", and press Enter
        Sleep 200 ; Give UI time to respond to button click
        Send "{Down}"
        Sleep 100
        Send "{Down}"
        Sleep 100
        Send "{Down}"
        Sleep 100
        Send "{Enter}"
        Sleep 400
        Send "ChatGPT"
        Sleep 100
        Send "{Enter}"
        Sleep 500 ; Wait for rename to complete

        ; Send F5 to refresh the page
        Send "{F5}"
        Sleep 2000 ; Wait for page refresh

        ; Collapse the sidebar at the end
        try {
            ; Try to find and click the close sidebar button (Portuguese or English)
            sidebarCloseButton := 0
            for name in sidebarCloseNames {
                try {
                    sidebarCloseButton := root.FindElement({ Type: 50000, Name: name, cs: false })
                    if (sidebarCloseButton)
                        break
                } catch {
                    try {
                        sidebarCloseButton := root.FindElement({ Type: 50000, Name: name })
                        if (sidebarCloseButton)
                            break
                    } catch {
                    }
                }
            }

            if (sidebarCloseButton) {
                try {
                    sidebarCloseButton.Invoke()
                } catch {
                    try {
                        sidebarCloseButton.Click()
                    } catch {
                        ; Fallback to keyboard shortcut if button click fails
                        Send "^+s"
                    }
                }
            } else {
                ; If button not found, use keyboard shortcut to close sidebar
                Send "^+s"
            }
        } catch {
            ; If any error occurs, use keyboard shortcut as fallback
            Send "^+s"
        }
        Sleep 300 ; Wait for sidebar to close

        ; Hide banner on success
        HideSmallLoadingIndicator_ChatGPT()
    } catch Error as err {
        ; Hide banner on error
        HideSmallLoadingIndicator_ChatGPT()
        ShowErr(err)
        return false
    }
    return true
}

; Ctrl + Alt + Y : Name ChatGPT window as "ChatGPT"
^!y::
{
    RenameChatGPTWindowToChatGPT()
}

#HotIf

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
    Send "{Enter}"
    Sleep 100
    A_Clipboard := oldClip

    ; Step 3: Alt+Tab to previous window
    Send "!{Tab}"

    ; After sending, show loading for Stop streaming
    buttonNames := ["Stop streaming", "Interromper transmissÃ£o"]
    WaitForButtonAndShowSmallLoading_ChatGPT(buttonNames, "Waiting for response...")
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
    ; --- Button Names (EN/PT) ---
    pt_stopStreamingName := "Interromper transmissão"
    en_stopStreamingName := "Stop streaming"
    currentStopStreamingName := IS_WORK_ENVIRONMENT ? pt_stopStreamingName : en_stopStreamingName

    ; Step 1: Send Escape to ensure composer is focused
    SendEscape()
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

; Shift + V : Set input volume to 100% - Volume
+V::
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

; Shift + F : Select first file - File
+f::
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

; Shift + S : Focus search bar - Search
+s:: Send "^e"

; Shift + A : Focus address bar - Address
+a:: Send "!d"

; Shift + N : New folder - New Folder
+n:: Send("^+n")

; Shift + H : Create a shortcut - sHortcut
+h::
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

; Shift + C : Copy as path - Copy
+c:: Send "^+c"

; Shift + R : Share file via context menu workflow - shaRe
+r::
{
    Explorer_CopyOneDriveShareLink_BoschGroup()
}

Explorer_CopyOneDriveShareLink_BoschGroup() {
    ; Reliable flow using classic Explorer context menu + UIA (no fixed Tab counts).
    ; Workflow:
    ;   1) Classic context menu -> S -> Enter (Share)
    ;   2) Wait for Share dialog main view
    ;   3) Open Link settings
    ;   4) Select "People in Bosch Group" and Apply
    ;   5) Copy link and confirm clipboard changed

    ; Ensure the focus is in the items view so the context menu targets the selected file.
    EnsureItemsViewFocus()
    Sleep 150
    ShowSmallLoadingIndicator_ChatGPT("Sharing file…")

    try {
        ; 1) Open classic context menu and trigger Share via accelerator.
        ; Shift+F10 is the canonical "classic menu" key, more reliable than AppsKey on some keyboards.
        Send "+{F10}"
        Sleep 200
        Send "s"
        Sleep 80
        Send "{Enter}"

        ; 2) Wait for OneDrive Share dialog (WebView2 host) to appear and load main controls.
        shareHwnd := OneDriveShare_WaitForShareDialogHwnd(20000)
        if !shareHwnd
            throw Error("Timed out waiting for the OneDrive Share dialog window.")

        shareRoot := UIA.ElementFromHandle(shareHwnd)

        ; Wait until main footer controls exist (indicates main share view is loaded).
        OneDriveShare_WaitForAutomationId(shareRoot, "Footer-button-settings", 20000)
        OneDriveShare_WaitForAutomationId(shareRoot, "copy-button", 20000)

        ; 3) Open Link settings (gear).
        settingsBtn := OneDriveShare_WaitForAutomationId(shareRoot, "Footer-button-settings", 5000)
        OneDriveShare_Click(settingsBtn)

        ; 4) In Link settings, select People in Bosch Group and Apply.
        applyBtn := OneDriveShare_WaitForAutomationId(shareRoot, "od-ModifyPermissions-apply-id", 20000)
        OneDriveShare_SelectRadioByNameContains(shareRoot, "People in Bosch Group", 5000)
        OneDriveShare_Click(applyBtn)

        ; After Apply, the dialog navigates back to main view.
        OneDriveShare_WaitForAutomationId(shareRoot, "copy-button", 20000)

        ; 5) Copy link and verify clipboard changed.
        copyBtn := OneDriveShare_WaitForAutomationId(shareRoot, "copy-button", 5000)
        oldClip := A_Clipboard
        A_Clipboard := ""
        OneDriveShare_Click(copyBtn)
        if !OneDriveShare_WaitForClipboardChange(oldClip, 10000)
            throw Error("Clipboard did not update after 'Copy link'.")
    } catch Error as e {
        MsgBox("Share macro failed:`n" e.Message, "Shift+R (Share file)", "IconX")
    } finally {
        HideSmallLoadingIndicator_ChatGPT()
    }
}

OneDriveShare_WaitForShareDialogHwnd(timeout := 20000) {
    deadline := A_TickCount + timeout
    while (A_TickCount < deadline) {
        for hwnd in WinGetList("ahk_class WebView2") {
            try {
                title := WinGetTitle("ahk_id " hwnd)
                if RegExMatch(title, "i)^Share\b") {
                    return hwnd
                }
            } catch {
            }
        }
        Sleep 100
    }
    return 0
}

OneDriveShare_WaitForAutomationId(root, automationId, timeout := 5000) {
    if !IsObject(root)
        return 0

    deadline := A_TickCount + timeout
    while (A_TickCount < deadline) {
        try {
            el := root.FindFirst({ AutomationId: automationId })
            if el
                return el
        } catch {
        }
        Sleep 80
    }
    return 0
}

OneDriveShare_Click(el) {
    if !IsObject(el)
        return false
    try {
        if el.GetPropertyValue(UIA.Property.IsInvokePatternAvailable) {
            el.Invoke()
            return true
        }
    } catch {
    }
    try {
        el.Click()
        return true
    } catch {
    }
    return false
}

OneDriveShare_SelectRadioByNameContains(root, nameNeedle, timeout := 5000) {
    deadline := A_TickCount + timeout
    while (A_TickCount < deadline) {
        try {
            radios := root.FindAll({ Type: "RadioButton" })
            for radio in radios {
                n := ""
                try n := radio.Name
                if (n != "" && InStr(n, nameNeedle)) {
                    try {
                        if radio.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable) {
                            if !radio.SelectionItemPattern.IsSelected
                                radio.SelectionItemPattern.Select()
                            return true
                        }
                    } catch {
                    }
                    ; Fallback: invoke/click the radio if SelectionItem isn't available.
                    OneDriveShare_Click(radio)
                    return true
                }
            }
        } catch {
        }
        Sleep 80
    }
    return false
}

OneDriveShare_WaitForClipboardChange(oldClip, timeout := 5000) {
    deadline := A_TickCount + timeout
    while (A_TickCount < deadline) {
        if (A_Clipboard != "" && A_Clipboard != oldClip)
            return true
        Sleep 80
    }
    return false
}

; Shift + P : Select first pinned item in Explorer sidebar - Pinned
+p::
{
    SelectExplorerSidebarFirstPinned_EX()
}

; Shift + L : Select the last item of the Explorer sidebar - Last
+l::
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

; Shift + X : WinRAR extract to current folder (personal); work PC stub
+x::
{
    global IS_WORK_ENVIRONMENT
    if (IS_WORK_ENVIRONMENT) {
        MsgBox("An extraction macro for the work environment still needs to be created.", "Shift+X", "Icon!")
        return
    }
    EnsureItemsViewFocus()
    Sleep 150
    Send "{AppsKey}"
    Sleep 250
    ; WinRAR shell menu accelerators (English); adjust if UI language differs
    Send "w"
    Sleep 150
    Send "x"
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
#HotIf WinActive("ahk_exe EXCEL.EXE") && WinGetClass("A") != "#32770"

; Helper function: Convert CSV delimited by semicolon into columns
; autoSelectSemicolon: If true, automatically selects semicolon without showing dialog. If false, shows confirmation dialog.
Excel_CSVToColumns(autoSelectSemicolon := false) {
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
    if (autoSelectSemicolon) {
        ; Automatically select semicolon without dialog
        Send "m"
        Sleep 100
    } else {
        ; Show confirmation dialog for user to decide
        if MsgBox("If 'semicolon' is not selected, hit yes", "Confirm", "YesNo Icon?") = "Yes" {
            Send "m"
            Sleep 100
        }
    }
    Send "{Enter}"
    Sleep 100
    Send "{Enter}"
}

; Shift + W : Select White Color (Up-Arrow, Ctrl-Home, Ctrl-Home)
+w:: {
    Send "^{PgUp}"
}

; Shift + E : Click Enable Editing button
+e:: {
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

; Shift + V : Quickly paste and extract CSV (Paste, CSV to columns)
+v:: {
    Excel_AddMultipleRows()    ; Add multiple rows first
    Sleep 200
    Send "^{Up}"
    Sleep 300
    Send "{Down}"
    Sleep 200
    Send "^v"           ; Ctrl+V (Paste action)
    Sleep 200
    Excel_CSVToColumns(true)    ; Auto-select semicolon, bypass dialog
    Sleep 200
    Send "^{Down}"
    Send "{Down}"
    Send "{Shift down}"
    Send "^{Down}"
    Send "{Shift up}"
    Send "{Alt}"
    Send "3"
    Send "r"
    ; Send "{Down}"
    ; Excel_RemoveRows(10)
    ; Sleep 200
    ; Send "^{Up}"
}

; Shift + C : Turn CSV delimited by semicolon into columns (Alt, 0, 5, D, Enter, M, Enter, Enter)
+c:: {
    Excel_CSVToColumns()
}

; Helper function: Add multiple rows (repeat Alt, Alt, 0, 2 with delays)
Excel_AddMultipleRows(count := 15) {
    ShowSmallLoadingIndicator_ChatGPT("Adding " . count . " rows...")
    ; Extra initial delay so the first Alt+0,2 sequence isn't too fast
    Sleep 300
    loop count {
        Send "{Alt down}"
        Send "{Alt up}"
        Sleep 100
        Send "0"
        Sleep 50
        Send "2"
        Sleep 50
    }
    HideSmallLoadingIndicator_ChatGPT()
}

; Shift + A : Add multiple rows (repeat Alt, Alt, 0, 2 with delays)
+a:: {
    Excel_AddMultipleRows()    ; Call function directly
}

; Helper function: Row removal workflow (remove row, down arrow, repeat)
; Pre-condition: Place cursor in starting cell
; Step 1: Execute REMOVE ROW SHORTCUT (Alt, Alt, 3, R)
; Step 2: Press DOWN ARROW
; Step 3: Repeat Step 1 and Step 2 for specified iterations
; Purpose: Remove alternating sequences of empty and populated rows
Excel_RemoveRows(iterations := 8) {
    ShowSmallLoadingIndicator_ChatGPT("Removing rows...")
    loop iterations {
        ; Execute Remove Row shortcut (Alt, Alt, 3, R)
        Send "{Alt down}"
        Send "{Alt up}"
        Sleep 150
        Send "3"
        Sleep 100
        Send "r"
        Sleep 150
        ; Press Down Arrow
        Send "{Down}"
        Sleep 150
    }
    HideSmallLoadingIndicator_ChatGPT()
}

; Shift + R : Row removal workflow (remove row, down arrow, repeat 5-7 times)
+r:: {
    Excel_RemoveRows()    ; Call function directly
}

; Shift + P : Type previous day date
+p:: {
    ; Calculate yesterday's date
    ; Get current date/time and subtract exactly 24 hours (86400 seconds)
    currentTime := A_Now
    yesterdayTime := DateAdd(currentTime, -86400, "Seconds")
    ; Format as MM/dd/yyyy (MM/DD/YYYY format)
    dateStr := FormatTime(yesterdayTime, "MM/dd/yyyy")
    ; Small delay to ensure Excel is ready
    Sleep 50
    ; Type the date as a whole word at once
    SendText dateStr
}

#HotIf

;-------------------------------------------------------------------
; Power BI Shortcuts
;-------------------------------------------------------------------
#HotIf (WinActive("ahk_exe PBIDesktop.exe") || InStr(WinGetTitle("A"), "powerbi", false)) && !IsFileDialogActive()

; Shift + T : Transform data (Click Home tab, then T, then UIA click)
+t:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Click the Home tab
        homeTab := root.FindFirst({ Type: "50019", Name: "Home", AutomationId: "home" })
        if !homeTab {
            homeTab := root.FindFirst({ Type: "50019", Name: "Home" })
        }
        if !homeTab {
            homeTab := root.FindFirst({ Type: "50019", AutomationId: "home" })
        }

        if homeTab {
            homeTab.Click()
            Sleep 120
        } else {
            MsgBox "Could not find the 'Home' tab.", "Power BI", "IconX"
            return
        }

        Send "t"
        Sleep 250

        possibleNames := ["Transform data", "Transformar dados"]
        transformBtn := ""

        ; Try to find by Name and Type 50000 (Button)
        for , name in possibleNames {
            transformBtn := root.FindFirst({ Name: name, Type: "50000" })
            if transformBtn
                break
        }

        ; Fallback: try with ClassName
        if !transformBtn {
            transformBtn := root.FindFirst({ Type: "50000", ClassName: "splitPrimaryButton root-332" })
        }

        ; Fallback: try with partial ClassName match
        if !transformBtn {
            transformBtn := root.FindFirst({ Type: "50000", ClassName: "splitPrimaryButton", matchmode: "Substring" })
        }

        ; Fallback: try with partial Name match
        if !transformBtn {
            transformBtn := root.FindFirst({ Name: "Transform", Type: "50000", matchmode: "Substring" })
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
    SendEscape()
    SendEscape()
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

; Shift + H : Build visual
+h:: {
    try {
        win := WinExist("A")
        if !win
            return
        root := UIA.ElementFromHandle(win)

        possibleNames := [
            "Build visual",
            "Build visuals",
            "Build visualization",
            "Build pane",
            "Visualizar",
            "Criar visual",
            "Criar visualização",
            "Construir visual",
            "Construir visualização"
        ]

        buildTab := ""

        for name in possibleNames {
            buildTab := root.FindFirst({ Type: "50019", Name: name })
            if buildTab
                break
            buildTab := root.FindFirst({ Type: "TabItem", Name: name })
            if buildTab
                break
            buildTab := root.FindFirst({ Type: "50019", Name: name, matchmode: "Substring" })
            if buildTab
                break
            buildTab := root.FindFirst({ Type: "TabItem", Name: name, matchmode: "Substring" })
            if buildTab
                break
        }

        if !buildTab {
            tabCond := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.TabItem)
            tabs := ""
            try tabs := root.FindElements(tabCond, UIA.TreeScope.Descendants)
            if tabs {
                for tab in tabs {
                    if !tab
                        continue
                    tabName := tab.Name
                    for name in possibleNames {
                        if InStr(tabName, name) {
                            buildTab := tab
                            break
                        }
                    }
                    if buildTab
                        break
                }
            }
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

; Shift + J : Format visual
+j:: {
    try {
        win := WinExist("A")
        if !win
            return
        root := UIA.ElementFromHandle(win)

        formatTab := ""

        ; Try "Format page" first (this is the working solution and is fast)
        try {
            formatTab := root.FindFirst({ Type: "50019", Name: "Format page" })
        } catch {
            try {
                formatTab := root.FindFirst({ Type: "TabItem", Name: "Format page" })
            } catch {
                ; Continue to fallback searches
            }
        }

        ; If "Format page" not found, try original names (simplified - only most common)
        if !formatTab {
            possibleNames := ["Format visual", "Format visuals", "Formatting"]
            for name in possibleNames {
                try {
                    formatTab := root.FindFirst({ Type: "50019", Name: name })
                    if formatTab
                        break
                } catch {
                    try {
                        formatTab := root.FindFirst({ Type: "TabItem", Name: name })
                        if formatTab
                            break
                    } catch {
                        ; Continue to next name
                    }
                }
            }
        }

        ; Final fallback: use FindElements to search all tabs
        if !formatTab {
            try {
                tabCond := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.TabItem)
                tabs := root.FindElements(tabCond, UIA.TreeScope.Descendants)
                if tabs {
                    for tab in tabs {
                        if !tab
                            continue
                        tabName := tab.Name
                        if InStr(tabName, "Format page") || InStr(tabName, "Format visual") || InStr(tabName,
                            "Formatting") {
                            formatTab := tab
                            break
                        }
                    }
                }
            } catch {
                ; Ignore errors
            }
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

; Shift + S : Select the Power BI search edit field (Data anchor + Tab)
+s:: {
    try {
        win := WinExist("A")
        if !win
            return

        root := UIA.ElementFromHandle(win)

        dataBtn := ""
        try {
            for cfg in PowerBI_GetDrawerConfigs() {
                if (cfg.HasOwnProp("label") && cfg.label = "Data") {
                    dataBtn := PowerBI_FindDrawerButton(root, cfg)
                    if dataBtn
                        break
                }
            }
        }

        if !dataBtn {
            MsgBox "Could not locate the Data button anchor.", "Power BI", "IconX"
            return
        }

        focused := false
        try {
            dataBtn.SetFocus()
            focused := true
        } catch {
            try {
                dataBtn.Select()
                focused := true
            } catch {
            }
        }

        if !focused {
            MsgBox "Could not focus the Data button anchor.", "Power BI", "IconX"
            return
        }

        Sleep 120
        Send "{Tab}"
        Sleep 120
        Send "^a"
    } catch Error as e {
        MsgBox "Error selecting the Power BI search field: " e.Message, "Power BI Error", "IconX"
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

; Shift + X : Click Cancel/Exit button in Power BI modals
+x:: {
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
    SendEscape()  ; Escape key is universal for cancels
}

; Shift + A : Right-click All pages button in Power BI
+a:: {
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

; Shift + W : Click New page button
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

; Shift + E : New measure (Click Home tab, then New measure button)
+e:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Click the Home tab
        homeTab := root.FindFirst({ Type: "50019", Name: "Home", AutomationId: "home" })
        if !homeTab {
            homeTab := root.FindFirst({ Type: "50019", Name: "Home" })
        }
        if !homeTab {
            homeTab := root.FindFirst({ Type: "50019", AutomationId: "home" })
        }

        if homeTab {
            homeTab.Click()
            Sleep 200
        } else {
            MsgBox "Could not find the 'Home' tab.", "Power BI", "IconX"
            return
        }

        ; Click the New measure button
        newMeasureBtn := root.FindFirst({ Type: "50000", Name: "New measure", AutomationId: "newMeasure" })
        if !newMeasureBtn {
            newMeasureBtn := root.FindFirst({ Type: "50000", Name: "New measure" })
        }
        if !newMeasureBtn {
            newMeasureBtn := root.FindFirst({ Type: "50000", AutomationId: "newMeasure" })
        }
        if !newMeasureBtn {
            newMeasureBtn := root.FindFirst({ Type: "50000", ClassName: "ms-Button root-336" })
        }

        if newMeasureBtn {
            newMeasureBtn.Click()
        } else {
            MsgBox "Could not find the 'New measure' button.", "Power BI", "IconX"
        }
    } catch Error as e {
        MsgBox "Error triggering New measure: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + Y : Refresh (Click Home tab, then Refresh button)
+y:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Click the Home tab
        homeTab := root.FindFirst({ Type: "50019", Name: "Home", AutomationId: "home" })
        if !homeTab {
            homeTab := root.FindFirst({ Type: "50019", Name: "Home" })
        }
        if !homeTab {
            homeTab := root.FindFirst({ Type: "50019", AutomationId: "home" })
        }

        if homeTab {
            homeTab.Click()
            Sleep 200
        } else {
            MsgBox "Could not find the 'Home' tab.", "Power BI", "IconX"
            return
        }

        ; Find the Refresh button
        refreshBtn := root.FindFirst({ Type: "50000", Name: "Refresh" })
        if !refreshBtn {
            refreshBtn := root.FindFirst({ Type: "50000", ClassName: "splitPrimaryButton root-332" })
        }
        if !refreshBtn {
            refreshBtn := root.FindFirst({ Type: "50000", ClassName: "splitPrimaryButton", matchmode: "Substring" })
        }

        if refreshBtn {
            refreshBtn.Click()
        } else {
            MsgBox "Could not find the 'Refresh' button.", "Power BI", "IconX"
        }
    } catch Error as e {
        MsgBox "Error triggering Refresh: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + B : Bring forward (Click Format tab, then click button 10 times)
+b:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Click the Format tab
        formatTab := root.FindFirst({ Type: "50019", Name: "Format", AutomationId: "visualFormatting" })
        if !formatTab {
            formatTab := root.FindFirst({ Type: "50019", Name: "Format" })
        }
        if !formatTab {
            formatTab := root.FindFirst({ Type: "50019", AutomationId: "visualFormatting" })
        }

        if formatTab {
            formatTab.Click()
            Sleep 200
        } else {
            MsgBox "Could not find the 'Format' tab.", "Power BI", "IconX"
            return
        }

        ; Find the Bring forward button
        bringForwardBtn := root.FindFirst({ Type: "50000", Name: "Bring forward" })
        if !bringForwardBtn {
            bringForwardBtn := root.FindFirst({ Type: "50000", ClassName: "splitPrimaryButton root-332", Name: "Bring forward" })
        }
        if !bringForwardBtn {
            ; Search all buttons for one named "Bring forward"
            allButtons := root.FindAll({ Type: "50000" })
            for btn in allButtons {
                if (btn.Name = "Bring forward") {
                    bringForwardBtn := btn
                    break
                }
            }
        }

        if !bringForwardBtn {
            MsgBox "Could not find the 'Bring forward' button.", "Power BI", "IconX"
            return
        }

        ; Show execution banner
        ShowSmallLoadingIndicator_ChatGPT("Bringing forward 10 times...")

        ; Click the button 10 times
        loop 10 {
            bringForwardBtn.Click()
            Sleep 50
        }

        ; Hide banner after completion
        Sleep 300
        HideSmallLoadingIndicator_ChatGPT()
    } catch Error as e {
        HideSmallLoadingIndicator_ChatGPT()
        MsgBox "Error triggering Bring forward: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + D : Send backward (Click Format tab, then click button 10 times)
+d:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Click the Format tab
        formatTab := root.FindFirst({ Type: "50019", Name: "Format", AutomationId: "visualFormatting" })
        if !formatTab {
            formatTab := root.FindFirst({ Type: "50019", Name: "Format" })
        }
        if !formatTab {
            formatTab := root.FindFirst({ Type: "50019", AutomationId: "visualFormatting" })
        }

        if formatTab {
            formatTab.Click()
            Sleep 200
        } else {
            MsgBox "Could not find the 'Format' tab.", "Power BI", "IconX"
            return
        }

        ; Find the Send backward button
        sendBackwardBtn := root.FindFirst({ Type: "50000", Name: "Send backward" })
        if !sendBackwardBtn {
            sendBackwardBtn := root.FindFirst({ Type: "50000", ClassName: "splitPrimaryButton root-332", Name: "Send backward" })
        }
        if !sendBackwardBtn {
            ; Search all buttons for one named "Send backward"
            allButtons := root.FindAll({ Type: "50000" })
            for btn in allButtons {
                if (btn.Name = "Send backward") {
                    sendBackwardBtn := btn
                    break
                }
            }
        }

        if !sendBackwardBtn {
            MsgBox "Could not find the 'Send backward' button.", "Power BI", "IconX"
            return
        }

        ; Show execution banner
        ShowSmallLoadingIndicator_ChatGPT("Sending backward 10 times...")

        ; Click the button 10 times
        loop 10 {
            sendBackwardBtn.Click()
            Sleep 50
        }

        ; Hide banner after completion
        Sleep 300
        HideSmallLoadingIndicator_ChatGPT()
    } catch Error as e {
        HideSmallLoadingIndicator_ChatGPT()
        MsgBox "Error triggering Send backward: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + K : Align (Click Format tab, then click Align button)
+k:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Click the Format tab
        formatTab := root.FindFirst({ Type: "50019", Name: "Format", AutomationId: "visualFormatting" })
        if !formatTab {
            formatTab := root.FindFirst({ Type: "50019", Name: "Format" })
        }
        if !formatTab {
            formatTab := root.FindFirst({ Type: "50019", AutomationId: "visualFormatting" })
        }

        if formatTab {
            formatTab.Click()
            Sleep 200
        } else {
            MsgBox "Could not find the 'Format' tab.", "Power BI", "IconX"
            return
        }

        ; Find the Align button
        alignBtn := root.FindFirst({ Type: "50000", Name: "Align", AutomationId: "alignFlyout" })
        if !alignBtn {
            alignBtn := root.FindFirst({ Type: "50000", Name: "Align" })
        }
        if !alignBtn {
            alignBtn := root.FindFirst({ Type: "50000", AutomationId: "alignFlyout" })
        }
        if !alignBtn {
            alignBtn := root.FindFirst({ Type: "50000", ClassName: "ms-Button ms-Button--hasMenu root-337" })
        }
        if !alignBtn {
            ; Search all buttons for one named "Align"
            allButtons := root.FindAll({ Type: "50000" })
            for btn in allButtons {
                if (btn.Name = "Align") {
                    alignBtn := btn
                    break
                }
            }
        }

        if !alignBtn {
            MsgBox "Could not find the 'Align' button.", "Power BI", "IconX"
            return
        }

        ; Click the Align button
        alignBtn.Click()
    } catch Error as e {
        MsgBox "Error triggering Align: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + V : Fit to page
+v:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Find the Fit to page button
        fitToPageBtn := root.FindFirst({ Type: "50000", Name: "Fit to page", AutomationId: "fitToPageButton" })
        if !fitToPageBtn {
            fitToPageBtn := root.FindFirst({ Type: "50000", Name: "Fit to page" })
        }
        if !fitToPageBtn {
            fitToPageBtn := root.FindFirst({ Type: "50000", AutomationId: "fitToPageButton" })
        }
        if !fitToPageBtn {
            fitToPageBtn := root.FindFirst({ Type: "50000", ClassName: "smallImageButton", AutomationId: "fitToPageButton" })
        }
        if !fitToPageBtn {
            ; Search all buttons for one named "Fit to page"
            allButtons := root.FindAll({ Type: "50000" })
            for btn in allButtons {
                if (btn.Name = "Fit to page") {
                    fitToPageBtn := btn
                    break
                }
            }
        }

        if !fitToPageBtn {
            MsgBox "Could not find the 'Fit to page' button.", "Power BI", "IconX"
            return
        }

        ; Click the Fit to page button
        fitToPageBtn.Click()
    } catch Error as e {
        MsgBox "Error triggering Fit to page: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + M : Format painter (Match format)
+m:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Find the Format painter button by AutomationId (primary method)
        formatPainterBtn := root.FindFirst({ Type: "50000", AutomationId: "formatPainter" })
        if !formatPainterBtn {
            formatPainterBtn := root.FindFirst({ Type: 50000, AutomationId: "formatPainter" })
        }

        ; Fallback: Try by Name and Type
        if !formatPainterBtn {
            formatPainterBtn := root.FindFirst({ Type: "50000", Name: "Format painter" })
        }
        if !formatPainterBtn {
            formatPainterBtn := root.FindFirst({ Type: 50000, Name: "Format painter" })
        }

        ; Fallback: Try by ClassName
        if !formatPainterBtn {
            formatPainterBtn := root.FindFirst({ Type: "50000", ClassName: "ms-Button root-345", AutomationId: "formatPainter" })
        }
        if !formatPainterBtn {
            formatPainterBtn := root.FindFirst({ Type: 50000, ClassName: "ms-Button root-345" })
        }

        ; Fallback: Try partial ClassName match
        if !formatPainterBtn {
            formatPainterBtn := root.FindFirst({ Type: "50000", ClassName: "ms-Button", matchmode: "Substring" })
            ; Verify it's the right button by checking AutomationId or Name
            if formatPainterBtn {
                try {
                    if (formatPainterBtn.AutomationId != "formatPainter" && formatPainterBtn.Name != "Format painter") {
                        formatPainterBtn := ""
                    }
                } catch {
                    formatPainterBtn := ""
                }
            }
        }

        ; Last resort: Search all buttons for Format painter
        if !formatPainterBtn {
            allButtons := root.FindAll({ Type: "50000" })
            for btn in allButtons {
                try {
                    if (btn.AutomationId = "formatPainter" || btn.Name = "Format painter") {
                        formatPainterBtn := btn
                        break
                    }
                } catch {
                }
            }
        }

        if !formatPainterBtn {
            MsgBox "Could not find the 'Format painter' button.", "Power BI", "IconX"
            return
        }

        ; Click the Format painter button
        formatPainterBtn.Click()
    } catch Error as e {
        MsgBox "Error triggering Format painter: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + N : Group visuals (Click Format tab, then click Group button)
+n:: {
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Click the Format tab
        formatTab := root.FindFirst({ Type: "50019", Name: "Format", AutomationId: "visualFormatting" })
        if !formatTab {
            formatTab := root.FindFirst({ Type: "50019", Name: "Format" })
        }
        if !formatTab {
            formatTab := root.FindFirst({ Type: "50019", AutomationId: "visualFormatting" })
        }

        if formatTab {
            formatTab.Click()
            Sleep 200
        } else {
            MsgBox "Could not find the 'Format' tab.", "Power BI", "IconX"
            return
        }

        ; Find the Group button by AutomationId (primary method)
        groupBtn := root.FindFirst({ Type: "50000", AutomationId: "groupVisualsFlyout" })
        if !groupBtn {
            groupBtn := root.FindFirst({ Type: 50000, AutomationId: "groupVisualsFlyout" })
        }

        ; Fallback: Try by Name and Type
        if !groupBtn {
            groupBtn := root.FindFirst({ Type: "50000", Name: "Group" })
        }
        if !groupBtn {
            groupBtn := root.FindFirst({ Type: 50000, Name: "Group" })
        }

        ; Fallback: Try by ClassName
        if !groupBtn {
            groupBtn := root.FindFirst({ Type: "50000", ClassName: "ms-Button ms-Button--hasMenu root-337",
                AutomationId: "groupVisualsFlyout" })
        }
        if !groupBtn {
            groupBtn := root.FindFirst({ Type: 50000, ClassName: "ms-Button ms-Button--hasMenu root-337" })
        }

        ; Fallback: Try partial ClassName match
        if !groupBtn {
            groupBtn := root.FindFirst({ Type: "50000", ClassName: "ms-Button--hasMenu", matchmode: "Substring" })
            ; Verify it's the right button by checking AutomationId or Name
            if groupBtn {
                try {
                    if (groupBtn.AutomationId != "groupVisualsFlyout" && groupBtn.Name != "Group") {
                        groupBtn := ""
                    }
                } catch {
                    groupBtn := ""
                }
            }
        }

        ; Last resort: Search all buttons for Group
        if !groupBtn {
            allButtons := root.FindAll({ Type: "50000" })
            for btn in allButtons {
                try {
                    if (btn.AutomationId = "groupVisualsFlyout" || btn.Name = "Group") {
                        groupBtn := btn
                        break
                    }
                } catch {
                }
            }
        }

        if !groupBtn {
            MsgBox "Could not find the 'Group' button.", "Power BI", "IconX"
            return
        }

        ; Click the Group button
        groupBtn.Click()
    } catch Error as e {
        MsgBox "Error triggering Group: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + F : Close all Power BI drawers (Visualizations/Data/Properties/Filters)
+f:: {
    try {
        win := WinExist("A")
        if !win
            return

        root := UIA.ElementFromHandle(win)

        drawerConfigs := PowerBI_GetDrawerConfigs()

        closed := 0
        already := 0
        skipped := 0

        for , cfg in drawerConfigs {
            btn := PowerBI_FindDrawerButton(root, cfg)
            if !btn {
                skipped++
                continue
            }
            result := PowerBI_CollapseDrawerElement(btn)
            if result = 1 {
                closed++
            } else if result = 0 {
                already++
            } else {
                skipped++
            }
        }

        msg := closed
            ? Format("Closed {} drawer{}", closed, closed = 1 ? "" : "s")
                : "No drawers needed closing"

        if already
            msg .= Format(" | {} already closed", already)
        if skipped
            msg .= Format(" | {} skipped", skipped)

        ToolTip msg
        SetTimer(() => ToolTip(), -1500)
    } catch Error as e {
        MsgBox "Error closing Power BI drawers: " e.Message, "Power BI Error", "IconX"
    }
}

; Shift + G : Open all Power BI drawers (Visualizations/Data/Properties/Filters)
+g:: {
    try {
        win := WinExist("A")
        if !win
            return

        root := UIA.ElementFromHandle(win)
        drawerConfigs := PowerBI_GetDrawerConfigs()

        opened := 0
        already := 0
        skipped := 0

        for , cfg in drawerConfigs {
            btn := PowerBI_FindDrawerButton(root, cfg)
            if !btn {
                skipped++
                continue
            }
            result := PowerBI_ExpandDrawerElement(btn)
            if result = 1 {
                opened++
            } else if result = 0 {
                already++
            } else {
                skipped++
            }
        }

        msg := opened
            ? Format("Opened {} drawer{}", opened, opened = 1 ? "" : "s")
                : "No drawers needed opening"

        if already
            msg .= Format(" | {} already open", already)
        if skipped
            msg .= Format(" | {} skipped", skipped)

        ToolTip msg
        SetTimer(() => ToolTip(), -1500)
    } catch Error as e {
        MsgBox "Error opening Power BI drawers: " e.Message, "Power BI Error", "IconX"
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

PowerBI_GetDrawerConfigs() {
    return [{
        label: "Visualizations",
        names: ["Visualizations", "Visualizações"],
        classContains: ["toggle-button"]
    }, {
        label: "Data",
        names: ["Data", "Dados"],
        classContains: ["toggle-button"]
    }, {
        label: "Properties",
        names: ["Properties", "Propriedades"],
        classContains: ["toggle-button"]
    }, {
        label: "Filter pane",
        names: [
            "Collapse or expand the filter pane while editing. This also determines how report readers see it",
            "Filter pane",
            "Pane de filtros"
        ],
        classContains: ["pbi-glyph-doublechevronleft", "pbi-glyph-doublechevronright"]
    }]
}

PowerBI_FindDrawerButton(root, config) {
    try {
        if config.HasOwnProp("names") {
            for , name in config.names {
                if !name
                    continue
                for typeVariant in ["Button", 50000] {
                    btn := root.FindFirst({ Type: typeVariant, Name: name })
                    if btn
                        return btn
                    btn := root.FindFirst({ Type: typeVariant, Name: name, matchmode: "Substring" })
                    if btn
                        return btn
                }
            }
        }

        if config.HasOwnProp("classNames") {
            for , className in config.classNames {
                if !className
                    continue
                for typeVariant in ["Button", 50000] {
                    btn := root.FindFirst({ Type: typeVariant, ClassName: className })
                    if btn
                        return btn
                }
            }
        }

        if config.HasOwnProp("classContains") {
            classNeedles := config.classContains
            if (Type(classNeedles) != "Array")
                classNeedles := [classNeedles]
            allButtons := ""
            try allButtons := root.FindAll({ Type: "Button" })
            if !allButtons
                try allButtons := root.FindAll({ Type: 50000 })
            if allButtons {
                for btn in allButtons {
                    if !btn
                        continue
                    btnClass := ""
                    try btnClass := btn.ClassName
                    for , needle in classNeedles {
                        if needle && InStr(btnClass, needle)
                            return btn
                    }
                }
            }
        }
    } catch Error {
    }
    return ""
}

PowerBI_CollapseDrawerElement(element) {
    current := element
    loop 4 {
        if !current
            break
        result := PowerBI_AttemptCollapse(current)
        if result != -1
            return result
        try current := UIA.TreeWalkerTrue.GetParentElement(current)
        catch {
            current := ""
        }
    }
    return -1
}

PowerBI_AttemptCollapse(element) {
    try {
        hasPattern := element.GetPropertyValue(UIA.Property.IsExpandCollapsePatternAvailable)
        if hasPattern {
            pat := element.ExpandCollapsePattern
            state := pat.ExpandCollapseState
            if state != UIA.ExpandCollapseState.Collapsed {
                pat.Collapse()
                Sleep 40
                return 1
            }
            return 0
        }
    } catch Error {
    }
    return -1
}

PowerBI_ExpandDrawerElement(element) {
    current := element
    loop 4 {
        if !current
            break
        result := PowerBI_AttemptExpand(current)
        if result != -1
            return result
        try current := UIA.TreeWalkerTrue.GetParentElement(current)
        catch {
            current := ""
        }
    }
    return -1
}

PowerBI_AttemptExpand(element) {
    try {
        hasPattern := element.GetPropertyValue(UIA.Property.IsExpandCollapsePatternAvailable)
        if hasPattern {
            pat := element.ExpandCollapsePattern
            state := pat.ExpandCollapseState
            if state != UIA.ExpandCollapseState.Expanded {
                pat.Expand()
                Sleep 40
                return 1
            }
            return 0
        }
    } catch Error {
    }
    return -1
}

;-------------------------------------------------------------------
; Gmail Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("Gmail")

; Shift + I : Go to main inbox - Inbox
+i:: Send("gi")

; Shift + U : Go to updates - Updates
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

; Shift + F : Go to forums - Forums
+f::
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

; Shift + R : Toggle read / unread - Read
+r::
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

; Shift + P : Previous conversation - Previous
+p:: Send("p")

; Shift + N : Next conversation - Next
+n:: Send("n")

; Shift + A : Archive conversation - Archive
+a:: Send("e")

; Shift + S : Select conversation - Select
+s:: Send("x")

; Shift + R : Reply - Reply (Note: conflicts with Read/unread, but Reply is more common)
; Actually, let me use Y for Reply to avoid conflict
+y:: Send("r")

; Shift + A : Reply all - All (conflicts with Archive!)
; Let me use G for Reply all (G for Group/all)
+g:: Send("a")

; Shift + W : Forward - Forward
+w:: Send("f")

; Shift + S : Star/unstar conversation - Star (conflicts with Select!)
; Let me use T for Star (T for sTar)
+t:: Send("s")

; Shift + D : Delete - Delete
+d:: Send("#")

; Shift + X : Report as spam - Spam
+x:: Send("!")

; Shift + C : Compose new email - Compose
+c:: Send("c")

; Shift + M : Move to folder - Move
+m:: Send("v")

; Shift + H : Show keyboard shortcuts help - Help
+h:: Send("?")

; Shift + B : Click inbox button - Button
+b::
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

;-----------------------------------------
;  UIA: detect if focus is in Cursor main editor (Monaco inputarea)
;  Uses conditional path to locate editor element, then compares with focused element
;  Conditional path: RootView -> ... -> workbench.parts.editor -> editor-instance -> Edit
;-----------------------------------------
IsCursorMainEditorFocused() {
    try {
        winHandle := WinExist("ahk_exe Cursor.exe")
        if (!winHandle)
            return false

        root := UIA.ElementFromHandle(winHandle)
        if (!root)
            return false

        editorEl := root.ElementFromPath({ T: 33, CN: "RootView" }, { T: 33 }, { T: 33 }, { T: 33, CN: "ClientView" }, { T: 33 }, { T: 33 }, { T: 33 }, { T: 30 }, { T: 26 }, { T: 33 }, { T: 26,
            A: "workbench.parts.editor" }, { T: 26, CN: "editor-instance" }, { T: 20 }, { T: 4 }
        )
        if (!editorEl)
            return false

        fe := UIA.GetFocusedElement()
        if (!fe)
            return false

        return UIA.CompareElementsEx(editorEl, fe)
    } catch {
        return false
    }
}

; Focus the Cursor "Files Explorer" tree using UIA
FocusCursorFilesExplorer() {
    try {
        hwnd := WinExist("ahk_exe Cursor.exe")
        if (!hwnd)
            return false

        root := UIA.ElementFromHandle(hwnd)
        if (!root)
            return false

        treeType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Tree)
        feEn := UIA.CreatePropertyCondition(UIA.Property.Name, "Files Explorer")
        fePt := UIA.CreatePropertyCondition(UIA.Property.Name, "Explorador de Arquivos")
        feName := UIA.CreateOrCondition(feEn, fePt)
        feCond := UIA.CreateAndCondition(treeType, feName)

        fileTree := ""
        try fileTree := root.FindElement(feCond, UIA.TreeScope.Descendants)

        if !fileTree
            return false

        fileTree.SetFocus()
        return true
    } catch {
        return false
    }
}

; UIA: find Type 50020 text by exact Name under scope. Prefer on-screen; if several, pick bottom-most (largest top Y).
Cursor_FindPermissionText50020(scope, exactName, requireOnScreen := true) {
    try
        all := scope.FindAll({ Type: 50020 })
    catch
        return 0
    best := 0
    bestT := -0x7FFFFFFF
    for t in all {
        try
            nm := t.Name
        catch
            continue
        if (nm != exactName)
            continue
        if (requireOnScreen) {
            try {
                if t.GetPropertyValue(UIA.Property.IsOffscreen)
                    continue
            } catch {
            }
        }
        try {
            br := t.BoundingRectangle
            if (br.t > bestT) {
                bestT := br.t
                best := t
            }
        } catch {
            if (!best)
                best := t
        }
    }
    return best
}

; Same as Cursor_FindPermissionText50020 but Name must contain substring (case-insensitive).
Cursor_FindPermissionText50020Contains(scope, substring, requireOnScreen := true) {
    try
        all := scope.FindAll({ Type: 50020 })
    catch
        return 0
    best := 0
    bestT := -0x7FFFFFFF
    for t in all {
        try
            nm := t.Name
        catch
            continue
        if (!InStr(nm, substring, false))
            continue
        if (requireOnScreen) {
            try {
                if t.GetPropertyValue(UIA.Property.IsOffscreen)
                    continue
            } catch {
            }
        }
        try {
            br := t.BoundingRectangle
            if (br.t > bestT) {
                bestT := br.t
                best := t
            }
        } catch {
            if (!best)
                best := t
        }
    }
    return best
}

; Click permission-style Text (50020): prefer parent Invoke/Click, same pattern as !n "Review next file".
Cursor_ClickUiaTextOrParentInvoke(textEl) {
    try {
        parentBtn := UIA.TreeWalkerTrue.GetParentElement(textEl)
        if (parentBtn) {
            try {
                if parentBtn.GetPropertyValue(UIA.Property.IsInvokePatternAvailable) {
                    parentBtn.InvokePattern.Invoke()
                    return true
                }
            } catch {
            }
            try {
                parentBtn.Click()
                return true
            } catch {
            }
        }
    } catch {
    }
    try {
        if textEl.GetPropertyValue(UIA.Property.IsInvokePatternAvailable) {
            textEl.InvokePattern.Invoke()
            return true
        }
    } catch {
    }
    try {
        textEl.Click()
        return true
    } catch {
    }
    return false
}

; Try each name variant (e.g. straight vs curly quotes). Scope to workbench.parts.panel first, then full window.
Cursor_ClickPermissionLabel(variantNames*) {
    hwnd := WinExist("ahk_exe Cursor.exe")
    if (!hwnd)
        return false
    try
        WinActivate(hwnd)
    catch {
    }
    try
        root := UIA.ElementFromHandle(hwnd)
    catch
        return false
    if (!root)
        return false
    Sleep 100
    scope := root
    try {
        panel := root.FindFirst({ AutomationId: "workbench.parts.panel", Type: 50026 })
        if (panel)
            scope := panel
    } catch {
    }
    el := 0
    for name in variantNames {
        el := Cursor_FindPermissionText50020(scope, name, true)
        if (el)
            break
    }
    if (!el) {
        for name in variantNames {
            el := Cursor_FindPermissionText50020(scope, name, false)
            if (el)
                break
        }
    }
    if (!el) {
        for name in variantNames {
            el := Cursor_FindPermissionText50020(root, name, true)
            if (el)
                break
        }
    }
    if (!el) {
        for name in variantNames {
            el := Cursor_FindPermissionText50020(root, name, false)
            if (el)
                break
        }
    }
    if (!el)
        return false
    return Cursor_ClickUiaTextOrParentInvoke(el)
}

; Like Cursor_ClickPermissionLabel but matches any Type 50020 label containing substring (e.g. "Allowlist").
Cursor_ClickPermissionLabelContains(substring) {
    hwnd := WinExist("ahk_exe Cursor.exe")
    if (!hwnd)
        return false
    try
        WinActivate(hwnd)
    catch {
    }
    try
        root := UIA.ElementFromHandle(hwnd)
    catch
        return false
    if (!root)
        return false
    Sleep 100
    scope := root
    try {
        panel := root.FindFirst({ AutomationId: "workbench.parts.panel", Type: 50026 })
        if (panel)
            scope := panel
    } catch {
    }
    el := Cursor_FindPermissionText50020Contains(scope, substring, true)
    if (!el)
        el := Cursor_FindPermissionText50020Contains(scope, substring, false)
    if (!el)
        el := Cursor_FindPermissionText50020Contains(root, substring, true)
    if (!el)
        el := Cursor_FindPermissionText50020Contains(root, substring, false)
    if (!el)
        return false
    return Cursor_ClickUiaTextOrParentInvoke(el)
}

;-------------------------------------------------------------------
; Cursor Shortcuts
;-------------------------------------------------------------------
#HotIf IsEditorActive() && WinGetClass("A") != "#32770"

; Alt + M : Quick shortcut menu for Cursor
; GUI styled like Select AI Model (Utils.ahk): dark theme, Press 1–2 | R · A · F · P | Esc to cancel.
global g_CursorShortcutMenuGui := false
global g_CursorShortcutMenuActive := false

!m::
{
    ShowCursorShortcutMenu()
}

; Alt + A : Add File to AI Context (Add File to Cursor Chat)
!a::
{
    ; #region agent log
    global g_AgentDebugRunCounter
    if (!IsSet(g_AgentDebugRunCounter))
        g_AgentDebugRunCounter := 0
    g_AgentDebugRunCounter += 1
    runId := "altA-" g_AgentDebugRunCounter "-" A_TickCount
    AgentDebugLog_c605fb(runId, "H0", "Shift keys.ahk:!a entry", "Alt+A invoked", { win: WinGetTitle("A"), editorFocused: IsCursorMainEditorFocused() })
    ; #endregion agent log

    ; If invoked from the main editor, move focus to Explorer so the active file is highlighted there.
    if (IsCursorMainEditorFocused()) {
        Send "^+e"
        Sleep 350
        okFocus := FocusCursorFilesExplorer()
        if (!okFocus) {
            Sleep 150
            okFocus := FocusCursorFilesExplorer()
        }
        Sleep 120
    } else {
        ; #region agent log
        AgentDebugLog_c605fb(runId, "H4", "Shift keys.ahk:!a skip focus", "Not main editor; assume Explorer/other", {})
        ; #endregion agent log
    }

    ; Open context menu for the selected file in Explorer, then navigate to the target item.
    if (Cursor_ContextMenuSelectByDownAndVerify("Add File to Cursor Chat", runId))
        return

    ; Fallback: Shift+F10 is often more reliable than AppsKey on some keyboards.
    if (Cursor_ContextMenuSelectByDownAndVerify("Add File to Cursor Chat", runId, "+{F10}"))
        return
}

Cursor_ContextMenuSelectByDownAndVerify(targetText, runId := "", openKey := "{AppsKey}", maxSteps := 28) {
    ; Open context menu.
    Send openKey
    Sleep 240

    ; #region agent log
    if (runId != "")
        AgentDebugLog_c605fb(runId, "H1", "Shift keys.ahk:Cursor_ContextMenuSelectByDownAndVerify", "Opened context menu via=" openKey, 0)
    ; #endregion agent log

    ; Try to detect and read currently highlighted item and then walk down.
    step := 0
    while (step <= maxSteps) {
        step += 1
        name := Cursor_ContextMenuGetHighlightedName()

        ; #region agent log
        if (runId != "" && step <= 10)
            AgentDebugLog_c605fb(runId, "H2", "Shift keys.ahk:Cursor_ContextMenuSelectByDownAndVerify", "Step=" step " highlighted=" name, 0)
        ; #endregion agent log

        if (name = targetText) {
            Send "{Enter}"

            ; #region agent log
            if (runId != "")
                AgentDebugLog_c605fb(runId, "H2", "Shift keys.ahk:Cursor_ContextMenuSelectByDownAndVerify", "Matched target; sent Enter", 0)
            ; #endregion agent log

            return true
        }

        Send "{Down}"
        Sleep 55
    }

    ; #region agent log
    if (runId != "")
        AgentDebugLog_c605fb(runId, "H3", "Shift keys.ahk:Cursor_ContextMenuSelectByDownAndVerify", "Failed to find target within steps=" maxSteps, 0)
    ; #endregion agent log

    return false
}

Cursor_ContextMenuGetHighlightedName() {
    ; Strategy A: focused element is a MenuItem
    try {
        fe := UIA.GetFocusedElement()
        if (fe) {
            try {
                if (fe.ControlType = UIA.Type.MenuItem)
                    return fe.Name
            } catch {
            }
        }
    } catch {
    }

    ; Strategy B: find selected MenuItem
    try {
        hwnd := WinExist("ahk_exe Cursor.exe")
        root := UIA.ElementFromHandle(hwnd)
        all := root.FindAll({ Type: 50011 })
        for mi in all {
            try {
                if (mi.GetPropertyValue(UIA.Property.IsSelected)) {
                    return mi.Name
                }
            } catch {
            }
        }
    } catch {
    }

    ; Strategy C: find MenuItem with keyboard focus
    try {
        hwnd := WinExist("ahk_exe Cursor.exe")
        root := UIA.ElementFromHandle(hwnd)
        all := root.FindAll({ Type: 50011 })
        for mi in all {
            try {
                if (mi.GetPropertyValue(UIA.Property.HasKeyboardFocus)) {
                    return mi.Name
                }
            } catch {
            }
        }
    } catch {
    }

    return ""
}

Cursor_ClickContextMenuItemByName(itemName, timeout := 1200, runId := "") {
    deadline := A_TickCount + timeout
    didEntryLog := false
    while (A_TickCount < deadline) {
        hwnd := WinExist("ahk_exe Cursor.exe")
        if (!hwnd)
            return false

        root := 0
        try root := UIA.ElementFromHandle(hwnd)

        ; #region agent log
        if (runId != "" && !didEntryLog) {
            didEntryLog := true
            AgentDebugLog_c605fb(runId, "H2", "Shift keys.ahk:Cursor_ClickContextMenuItemByName entry", "Searching for menu item", { itemName: itemName, hwnd: hwnd, haveRoot: !!root })
        }
        ; #endregion agent log

        menuItem := 0
        try menuItem := root ? root.FindFirst({ Type: 50011, Name: itemName }) : 0
        if (!menuItem) {
            try menuItem := UIA.GetRootElement().FindFirst({ Type: 50011, Name: itemName })
        }
        if (!menuItem) {
            try menuItem := root ? root.FindFirst({ Type: 50011, Name: itemName, matchmode: "Substring" }) : 0
        }
        if (!menuItem) {
            try menuItem := UIA.GetRootElement().FindFirst({ Type: 50011, Name: itemName, matchmode: "Substring" })
        }

        if (menuItem) {
            clickedVia := "click"
            clickOk := false
            wasEnabled := ""
            wasOffscreen := ""
            invAvail := ""
            try wasEnabled := menuItem.GetPropertyValue(UIA.Property.IsEnabled)
            try wasOffscreen := menuItem.GetPropertyValue(UIA.Property.IsOffscreen)
            try invAvail := menuItem.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)

            ; Prefer InvokePattern when available (often more reliable than Click for menu actions).
            if (invAvail = 1) {
                clickedVia := "invokePattern"
                try {
                    menuItem.InvokePattern.Invoke()
                    clickOk := true
                } catch {
                    clickOk := false
                }
            }

            if (!clickOk) {
                clickedVia := "click"
                try {
                    menuItem.Click()
                    clickOk := true
                } catch {
                    clickedVia := "invoke"
                    try {
                        menuItem.Invoke()
                        clickOk := true
                    } catch {
                        clickOk := false
                    }
                }
            }

            ; #region agent log
            if (runId != "") {
                nm := ""
                cn := ""
                try nm := menuItem.Name
                try cn := menuItem.ClassName
                AgentDebugLog_c605fb(
                    runId,
                    "H2",
                    "Shift keys.ahk:Cursor_ClickContextMenuItemByName success",
                    "MenuItem pre enabled=" wasEnabled " offscreen=" wasOffscreen " invAvail=" invAvail " | via=" clickedVia " ok=" clickOk " name=" nm " class=" cn,
                    0
                )
            }
            ; #endregion agent log

            if (!clickOk)
                return false

            ; #region agent log
            if (runId != "") {
                stillThere := ""
                errMsg := ""
                try {
                    ; If the menu closes, this usually fails to find the item again.
                    stillThere := UIA.GetRootElement().FindFirst({ Type: 50011, Name: itemName }) ? 1 : 0
                } catch Error as e {
                    stillThere := "err"
                    try errMsg := e.Message
                }
                AgentDebugLog_c605fb(
                    runId,
                    "H1",
                    "Shift keys.ahk:Cursor_ClickContextMenuItemByName post",
                    "Post-click menuItemStillFindable=" stillThere " err=" errMsg,
                    0
                )

                ; If the menu is still open after invoke/click, try the classic activation path:
                ; focus the menu item and press Enter.
                if (stillThere = 1) {
                    enterOk := 0
                    try {
                        menuItem.SetFocus()
                        Sleep 40
                        Send "{Enter}"
                        enterOk := 1
                    } catch {
                        enterOk := 0
                    }
                    AgentDebugLog_c605fb(
                        runId,
                        "H1",
                        "Shift keys.ahk:Cursor_ClickContextMenuItemByName enterFallback",
                        "Menu stayed open; sent Enter fallback ok=" enterOk,
                        0
                    )
                }
            }
            ; #endregion agent log

            return true
        }
        Sleep 80
    }

    ; #region agent log
    if (runId != "") {
        names := []
        try {
            hwnd2 := WinExist("ahk_exe Cursor.exe")
            root2 := UIA.ElementFromHandle(hwnd2)
            all := 0
            try all := root2.FindAll({ Type: 50011 })
            if (all) {
                i := 0
                for mi in all {
                    i += 1
                    if (i > 12)
                        break
                    try names.Push(mi.Name)
                }
            }
        } catch {
        }
        AgentDebugLog_c605fb(runId, "H3", "Shift keys.ahk:Cursor_ClickContextMenuItemByName timeout", "Menu item not found before timeout", { itemName: itemName, sampleMenuItemNames: names })
    }
    ; #endregion agent log

    return false
}

; #region agent log
AgentDebugLog_c605fb(runId, hypothesisId, location, message, dataObj := 0) {
    ; NDJSON log for debug mode (no secrets).
    try {
        ; Minimal JSON escaper for strings we control.
        q := "`""
        ; Keep `data` empty for robustness (avoid enumeration errors).
        dataJson := "{}"
        r := "" runId
        h := "" hypothesisId
        loc := "" location
        msg := "" message
        r := StrReplace(r, "\", "/")
        r := StrReplace(r, q, "'")
        r := StrReplace(r, "`r`n", "\n")
        r := StrReplace(r, "`n", "\n")
        r := StrReplace(r, "`r", "\n")
        h := StrReplace(h, "\", "/")
        h := StrReplace(h, q, "'")
        h := StrReplace(h, "`r`n", "\n")
        h := StrReplace(h, "`n", "\n")
        h := StrReplace(h, "`r", "\n")
        loc := StrReplace(loc, "\", "/")
        loc := StrReplace(loc, q, "'")
        loc := StrReplace(loc, "`r`n", "\n")
        loc := StrReplace(loc, "`n", "\n")
        loc := StrReplace(loc, "`r", "\n")
        msg := StrReplace(msg, "\", "/")
        msg := StrReplace(msg, q, "'")
        msg := StrReplace(msg, "`r`n", "\n")
        msg := StrReplace(msg, "`n", "\n")
        msg := StrReplace(msg, "`r", "\n")
        line := "{"
            . q "sessionId" q ":" q "c605fb" q
            . "," q "runId" q ":" q r q
            . "," q "hypothesisId" q ":" q h q
            . "," q "location" q ":" q loc q
            . "," q "message" q ":" q msg q
            . "," q "data" q ":" dataJson
            . "," q "timestamp" q ":" A_TickCount
            . "}`n"
        FileAppend line, "C:\Users\fie7ca\Documents\scripts\debug-c605fb.log", "UTF-8"
    } catch Error as e {
        ; Failsafe: surface why logging failed (avoid silent failure).
        try MsgBox("Debug log write failed:`n" e.Message, "Alt+A debug", "IconX")
    }
}
; #endregion agent log

ShowCursorShortcutMenu() {
    global g_CursorShortcutMenuGui, g_CursorShortcutMenuActive
    if (g_CursorShortcutMenuActive)
        return

    g_CursorShortcutMenuGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Owner")
    g_CursorShortcutMenuGui.BackColor := "1E1E2E"
    g_CursorShortcutMenuGui.MarginX := 20
    g_CursorShortcutMenuGui.MarginY := 15

    g_CursorShortcutMenuGui.SetFont("s14 cCDD6F4 Bold", "Segoe UI")
    g_CursorShortcutMenuGui.Add("Text", "w300 Center", "Select shortcut")
    g_CursorShortcutMenuGui.Add("Text", "w300 h1 Background45475A")

    g_CursorShortcutMenuGui.SetFont("s12 cCDD6F4", "Segoe UI")
    g_CursorShortcutMenuGui.Add("Text", "w300", "[1] hello world one")
    g_CursorShortcutMenuGui.Add("Text", "w300", "[2] hello world two")

    g_CursorShortcutMenuGui.Add("Text", "w300 h1 Background45475A y+8")
    g_CursorShortcutMenuGui.SetFont("s10 c6C7086", "Segoe UI")
    g_CursorShortcutMenuGui.Add("Text", "w300", "Terminal permissions")
    g_CursorShortcutMenuGui.SetFont("s12 cCDD6F4", "Segoe UI")
    g_CursorShortcutMenuGui.Add("Text", "w300", "[R] Run (terminal permission)")
    g_CursorShortcutMenuGui.Add("Text", "w300", "[A] Allowlist (any permission button)")
    g_CursorShortcutMenuGui.Add("Text", "w300", "[F] Mark as fixed")
    g_CursorShortcutMenuGui.Add("Text", "w300", "[P] Proceed")

    g_CursorShortcutMenuGui.Add("Text", "w300 h1 Background45475A y+10")
    g_CursorShortcutMenuGui.SetFont("s9 c6C7086", "Segoe UI")
    g_CursorShortcutMenuGui.Add("Text", "w300 Center", "Press 1–2 | R · A · F · P | Esc to cancel")

    ; Center on same monitor as active window (same logic as Utils ShowAiModelSelector)
    activeWin := 0
    try
        activeWin := WinGetID("A")
    catch
        activeWin := 0
    MonitorGetWorkArea(1, &monitorLeft, &monitorTop, &monitorRight, &monitorBottom)
    monitorWidth := monitorRight - monitorLeft
    monitorHeight := monitorBottom - monitorTop
    if (activeWin && activeWin != 0) {
        rect := Buffer(16, 0)
        if (DllCall("GetWindowRect", "ptr", activeWin, "ptr", rect)) {
            centerX := NumGet(rect, 0, "int") + (NumGet(rect, 8, "int") - NumGet(rect, 0, "int")) // 2
            centerY := NumGet(rect, 4, "int") + (NumGet(rect, 12, "int") - NumGet(rect, 4, "int")) // 2
            loop MonitorGetCount() {
                MonitorGetWorkArea(A_Index, &l, &t, &r, &b)
                if (centerX >= l && centerX <= r && centerY >= t && centerY <= b) {
                    monitorLeft := l
                    monitorTop := t
                    monitorWidth := r - l
                    monitorHeight := b - t
                    break
                }
            }
        }
    }
    g_CursorShortcutMenuGui.Show("AutoSize Hide")
    g_CursorShortcutMenuGui.GetPos(&gx, &gy, &gw, &gh)
    cx := monitorLeft + (monitorWidth - gw) // 2
    cy := monitorTop + (monitorHeight - gh) // 2
    g_CursorShortcutMenuGui.Show("x" . cx . " y" . cy . " NA")

    g_CursorShortcutMenuActive := true
    Hotkey("1", (*) => CursorShortcutMenu_HandleKey("1"), "On")
    Hotkey("2", (*) => CursorShortcutMenu_HandleKey("2"), "On")
    Hotkey("r", (*) => CursorShortcutMenu_HandleKey("r"), "On")
    Hotkey("a", (*) => CursorShortcutMenu_HandleKey("a"), "On")
    Hotkey("f", (*) => CursorShortcutMenu_HandleKey("f"), "On")
    Hotkey("p", (*) => CursorShortcutMenu_HandleKey("p"), "On")
    Hotkey("Escape", CursorShortcutMenu_Cancel, "On")
}

CursorShortcutMenu_HandleKey(key) {
    global g_CursorShortcutMenuActive
    if (!g_CursorShortcutMenuActive)
        return
    CursorShortcutMenu_Close()
    if (key = "1")
        CursorShortcutMenu_Action1()
    else if (key = "2")
        CursorShortcutMenu_Action2()
    else if (key = "r")
        CursorShortcutMenu_ActionRun()
    else if (key = "a")
        CursorShortcutMenu_ActionAllowlist()
    else if (key = "f")
        CursorShortcutMenu_ActionMarkAsFixed()
    else if (key = "p")
        CursorShortcutMenu_ActionProceed()
}

CursorShortcutMenu_Cancel(*) {
    CursorShortcutMenu_Close()
}

CursorShortcutMenu_Close() {
    global g_CursorShortcutMenuGui, g_CursorShortcutMenuActive
    if (!g_CursorShortcutMenuActive)
        return
    g_CursorShortcutMenuActive := false
    try Hotkey("1", "Off")
    try Hotkey("2", "Off")
    try Hotkey("r", "Off")
    try Hotkey("a", "Off")
    try Hotkey("f", "Off")
    try Hotkey("p", "Off")
    try Hotkey("Escape", CursorShortcutMenu_Cancel, "Off")
    if (IsObject(g_CursorShortcutMenuGui) && g_CursorShortcutMenuGui.Hwnd) {
        try g_CursorShortcutMenuGui.Destroy()
    }
    g_CursorShortcutMenuGui := false
}

CursorShortcutMenu_Action1(*) {
    ; Replace with your command for "hello world one"
    return
}

CursorShortcutMenu_Action2(*) {
    ; Replace with your command for "hello world two"
    return
}

CursorShortcutMenu_ActionRun(*) {
    Cursor_ClickPermissionLabel("Run")
}

CursorShortcutMenu_ActionAllowlist(*) {
    Cursor_ClickPermissionLabelContains("Allowlist")
}

CursorShortcutMenu_ActionMarkAsFixed(*) {
    Cursor_ClickPermissionLabel("Mark as fixed")
}

CursorShortcutMenu_ActionProceed(*) {
    Cursor_ClickPermissionLabel("Proceed")
}

; Ctrl + H : Smart navigation - Editor → Explorer, Explorer → Reveal in Explorer
; Works from main editor even when the left Explorer sidebar is closed (opens it first).
^h::
{
    if (IsCursorMainEditorFocused()) {
        ; Main editor: ensure Explorer sidebar is available, focus Files Explorer, then trigger reveal
        if (FocusCursorFilesExplorer()) {
            Sleep 150
            Send "^h"
        } else {
            ; Sidebar may be closed: open Explorer (Ctrl+Shift+E), then retry UIA focus
            Send "^+e"
            Sleep 350
            if (FocusCursorFilesExplorer()) {
                Sleep 150
                Send "^h"
            } else {
                ; Fallback: legacy keybinding path if UIA still fails
                Send "^!+e"
                Sleep 200
                Send "^h"
            }
        }
    } else {
        ; User is NOT in main editor (likely in Explorer): trigger Reveal in File Explorer
        Send "^h"
    }
}

; Ctrl + 1 : Remove clustering and focus on the code
^1::
{
    ; Send ESC two times
    SendEscape()  ; ESC
    Sleep 50
    SendEscape()  ; ESC again
    Sleep 100
    Send "^!n"
    Sleep 100
    Send "^!,"
    Sleep 100
    Send "#!o"
}

; Ctrl + 5 : Context menu navigation sequence
^5::
{
    Click "Right"         ; 1. Right mouse click
    Sleep 100
    Send "{Down}"         ; 2. Press Down Arrow twice
    Send "{Down}"
    Sleep 50
    Send "{Right}"        ; 3. Press Right Arrow once
    Sleep 50
    Send "{Down}"         ; 4. Press Down Arrow once
    Sleep 50
    Send "{Right}"        ; 5. Press Right Arrow once
    Sleep 50
    Send "{Enter}"        ; 6. Press Enter
}

; Ensure only one Chrome window shows the given PDF: close any Chrome window whose title matches the filename, then open a fresh Chrome window.
EnsureSingleChromePdfInstance(filePath := "", fileNameOnly := "") {
    if (fileNameOnly = "")
        return

    ; Include all Chrome top-level windows (visible, minimized, or hidden)
    prevDetectHidden := A_DetectHiddenWindows
    DetectHiddenWindows true

    ; Collect Chrome window hwnds that appear to be showing this PDF.
    ; Title matching is unreliable because Chrome PDF viewer titles can reflect document content instead of filename.
    toClose := []
    fileNameLower := StrLower(fileNameOnly)
    filePathLower := ""
    if (filePath != "")
        filePathLower := StrLower(StrReplace(filePath, "\", "/"))
    baseNameLower := ""
    try {
        SplitPath fileNameOnly, , , &ext, &baseName
        if (baseName != "")
            baseNameLower := StrLower(baseName)
    }
    catch {
    }
    for hwnd in WinGetList("ahk_exe chrome.exe") {
        try {
            title := WinGetTitle("ahk_id " hwnd)
            if (title = "")
                continue
            titleLower := StrLower(title)
            matched := (fileNameLower != "" && InStr(titleLower, fileNameLower) != 0)
            if (!matched && baseNameLower != "")
                matched := InStr(titleLower, baseNameLower) != 0
            ; UIA-based match: look for a Document element whose Value contains the file path or filename.
            ; This catches Chrome PDF viewer windows whose title doesn't include the filename.
            matchedUia := false
            docValue := ""
            if (!matched) {
                try {
                    root := UIA.ElementFromHandle(hwnd)
                    ; Find a document-like element that exposes the PDF file URL/path.
                    docEl := root.FindFirst({ Type: "Document" })
                    if docEl {
                        try docValue := StrLower(docEl.Value)
                        catch {
                            docValue := ""
                        }
                        if (docValue != "") {
                            docValue := StrReplace(docValue, "\", "/")
                            if (filePathLower != "" && InStr(docValue, filePathLower))
                                matchedUia := true
                            else if (fileNameLower != "" && InStr(docValue, fileNameLower))
                                matchedUia := true
                        }
                    }
                } catch {
                    matchedUia := false
                }
            }

            if (!matched && matchedUia)
                matched := true
            if matched
                toClose.Push(hwnd)
        } catch {
            continue
        }
    }

    ; Restore previous DetectHiddenWindows setting
    DetectHiddenWindows prevDetectHidden

    for hwnd in toClose {
        try {
            WinActivate("ahk_id " hwnd)
            Sleep 150
            Send "!{F4}"
            Sleep 200
        } catch {
            ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
        }
    }
    Sleep 300
    ; Open a fresh, empty Chrome window. Marp will open the PDF itself
    ; in the last activated Chrome window after export completes.
    try Run "chrome.exe --new-window"
}

; Ctrl + 6 : Marp export - trigger export, handle Save As and Replace dialogs
^6::
{
    ; Show persistent banner for the entire export flow
    ShowSmallLoadingIndicator_ChatGPT("Exporting with Marp...")
    slowStepMs := 300  ; used throughout this hotkey flow
    mainHwnd := 0
    try {
        ; 1. Trigger Marp export
        try {
            mainHwnd := WinGetID("A")
        } catch {
            ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        Send "^6"
        Sleep 500
        Sleep slowStepMs

        ; 2. Wait for Save As / Export dialog
        ; Try: native #32770, title match, or active window change (modal steals focus)
        prevMatchMode := A_TitleMatchMode
        SetTitleMatchMode 2
        saveDialogHwnd := 0
        deadline := A_TickCount + 25000
        while (A_TickCount < deadline) {
            ; Native Windows dialog (standard Save As)
            h := WinExist("ahk_class #32770")
            if h {
                saveDialogHwnd := h
                break
            }
            ; Title contains Save/Export (any window)
            for str in ["Save As", "Export", "Salvar como", "Guardar como", "Save File", "Save PDF", "Marp",
                "Export PDF"] {
                h := WinExist(str)
                if h {
                    saveDialogHwnd := h
                    break 2
                }
            }
            ; Fallback: modal dialog stole focus (active window changed from main)
            try {
                curr := WinGetID("A")
            } catch {
                ; Transient: no active window at this tick. Keep waiting for the dialog instead of aborting.
                Sleep 250
                continue
            }
            if curr && curr != mainHwnd {
                currTitle := WinGetTitle("ahk_id " curr)
                currClass := WinGetClass("ahk_id " curr)
                ; Likely a dialog: standard dialog class, or a Chromium modal whose TITLE indicates Save/Export.
                ; IMPORTANT: Do NOT treat any Chrome_WidgetWin as dialog (Cursor itself is Chrome_WidgetWin_1).
                isTitleDialogish := InStr(currTitle, "Save") || InStr(currTitle, "Save As")
                || InStr(currTitle, "Export") || InStr(currTitle, "Marp")
                || InStr(currTitle, "Confirm Save As") || InStr(currTitle, "Confirm")
                || InStr(currTitle, "Salvar") || InStr(currTitle, "Guardar")
                if InStr(currClass, "32770") || (InStr(currClass, "Chrome_WidgetWin") && isTitleDialogish) {
                    saveDialogHwnd := curr
                    break
                }
            }
            Sleep 250
        }
        SetTitleMatchMode prevMatchMode
        if !saveDialogHwnd {
            return
        }
        try {
            WinActivate("ahk_id " saveDialogHwnd)
        } catch {
            ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        Sleep 700
        Sleep slowStepMs

        ; 2b. Extract PDF path and filename from the Save dialog via UIA (before confirming save)
        filePath := ""
        fileNameOnly := ""
        fileNameEditEl := 0
        try {
            root := UIA.ElementFromHandle(saveDialogHwnd)
            fileNameEdit := ""

            ; Scan Edit controls and detect the filename field
            try {
                edits := root.FindElements({ Type: "Edit" })
                for el in edits {
                    val := ""
                    try val := el.Value
                    catch {
                    }
                    if (val = "")
                        continue

                    parts := StrSplit(val, "\")
                    suffix := parts.Length ? parts[parts.Length] : val

                    ; If this Edit looks like the File name field (based on name/id and .pdf suffix), capture it
                    if (suffix != "" && InStr(StrLower(suffix), ".pdf")
                    && (el.AutomationId = "1001" || el.Name = "File name:")) {
                        filePath := val
                        SplitPath filePath, , , &ext, &nameNoExt
                        fileNameOnly := (nameNoExt != "") ? (nameNoExt . (ext != "" ? "." ext : "")) : suffix
                    }
                }
            } catch {
            }

            ; First attempt: Edit with AutomationId 1148 (matches file dialog helper)
            fileNameEdit := root.FindFirst({ Type: "Edit", AutomationId: "1148" })

            ; Second attempt: ComboBox 1148 -> inner Edit
            if !fileNameEdit {
                fileNameCombo := root.FindFirst({ Type: "ComboBox", AutomationId: "1148" })
                if fileNameCombo
                    fileNameEdit := fileNameCombo.FindFirst({ Type: "Edit" })
            }

            ; Third attempt: Edit by common localized names
            if !fileNameEdit {
                for name in ["File name:", "Nome:", "Filename:", "File Name:", "Name:", "Nome do arquivo:"] {
                    fileNameEdit := root.FindFirst({ Type: "Edit", Name: name })
                    if fileNameEdit
                        break
                }
            }
            if fileNameEdit {
                fileNameEditEl := fileNameEdit
                pathOrName := Trim(fileNameEdit.Value)
                if (pathOrName != "") {
                    if (InStr(pathOrName, "\") || InStr(pathOrName, ":"))
                        filePath := pathOrName
                    else {
                        ; Filename only: try to get current folder from another Edit in the dialog
                        for el in root.FindElements({ Type: "Edit" }) {
                            try {
                                val := Trim(el.Value)
                                if (val != "" && InStr(val, "\") && InStr(val, ":")) {
                                    filePath := RTrim(val, "\") "\" pathOrName
                                    break
                                }
                            } catch {
                                continue
                            }
                        }
                        if (filePath = "")
                            filePath := pathOrName
                    }
                    SplitPath filePath, , , &ext, &nameNoExt
                    fileNameOnly := (nameNoExt != "") ? (nameNoExt . (ext != "" ? "." ext : "")) : pathOrName
                }
            }
        } catch {
            ; Fallback: no filename extracted; we will just open a new Chrome window at the end
        }

        ; Prepare Chrome context for this PDF: close old windows and open a new one
        if (fileNameOnly != "")
            EnsureSingleChromePdfInstance(filePath, fileNameOnly)
        Sleep slowStepMs

        Send "{Enter}"  ; Confirm initial save
        Sleep slowStepMs

        stillThere := WinExist("ahk_id " saveDialogHwnd) ? 1 : 0

        ; If Enter didn't confirm, click the dialog's Export button via UIA.
        if (stillThere) {
            try WinActivate("ahk_id " saveDialogHwnd)
            catch {
            }
            Sleep 120

            try {
                dlgRoot := UIA.ElementFromHandle(saveDialogHwnd)
                exportBtn := dlgRoot.FindFirst({ Type: "Button", Name: "Export", AutomationId: "1" })
                if !exportBtn
                    exportBtn := dlgRoot.FindFirst({ Type: "Button", Name: "Export" })
                if exportBtn {
                    exportBtn.Invoke()
                }
            } catch {
            }

            Sleep slowStepMs
        }

        ; 3. Handle Confirm Save As / Replace dialog (ClassName #32770, Name: "Confirm Save As")
        ; WinGetText doesn't capture UIA Text elements; use window title. Yes button has Alt+Y.
        SetTitleMatchMode 2
        loop 10 {
            Sleep 500
            replaceHwnd := WinExist("ahk_class #32770")
            if replaceHwnd {
                title := WinGetTitle("ahk_id " replaceHwnd)
                if InStr(title, "Confirm Save As") || InStr(title, "Confirmar Salvar")
                || InStr(title, "Confirmar Guardar") || InStr(title, "Confirm Replace") {
                    try {
                        WinActivate("ahk_id " replaceHwnd)
                    } catch {
                        ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
                        break
                    }
                    Sleep 900  ; Delay for dialog to stabilize before confirming
                    Send "!y"   ; Alt+Y = Yes (per UIA: AcceleratorKey: "Alt+Y")
                    break
                }
            }
        }

        ; 4. Wait for Marp export and viewer open to complete
        Sleep 800
        Sleep slowStepMs
    } finally {
        ; Always hide the banner when the flow completes or aborts
        HideSmallLoadingIndicator_ChatGPT()
    }
}

; Shift + F : Fold - Fold
+f::
{
    Send "^+8"
}

; Shift + U : Unfold - Unfold
+u::
{
    Send "^+9"
}

; Shift + M : Open markdown preview to the side - Markdown
+m:: Send "+i"

; Shift + W : Move editor into new window - Window
+w:: Send "+o"

; Shift + T : Go to terminal - Terminal
+t:: Send "^'"

; Shift + N : New terminal - New Terminal
+n:: Send '^+"'

; Shift + E : Go to file explorer - Explorer
+e:: Send "^+e"

; Shift + K : Open markdown preview and move editor into new window - Keep
+k::
{
    ; Show banner while algorithm is executing
    ; ShowSmallLoadingIndicator_ChatGPT("Processing...")

    ; Step 1: Trigger markdown preview (Shift+M -> +i in Cursor)
    Send "+i"
    Sleep 1800

    ; Step 2: Center mouse in active window (Win+Alt+Shift+Q)
    Send "#!+q"
    Sleep 150

    ; Step 3: Offset mouse right into Markdown Preview Enhanced pane
    PREVIEW_OFFSET_PX := 60
    MouseGetPos(&x, &y)
    x += PREVIEW_OFFSET_PX
    DllCall("SetCursorPos", "int", x, "int", y)
    Sleep 100

    ; Step 4: Click to focus preview area
    Click

    ; Step 5: Detach tab (Shift+W -> +o in Cursor)
    Sleep 2000
    Send "+o"

    Sleep 300

    WinMaximize "A"

    ; Hide banner after completion
    ; HideSmallLoadingIndicator_ChatGPT()
}

; Shift + C : Command palette - Command
+c:: Send "^+p"

; Shift + X : Expand selection - Expand
+x:: Send "+!{Right}"

; Shift + S : Go to symbol in access view - Symbol
+s:: Send "+m"

; Shift + H : Show chat history - History
+h::
{
    Send "^+p" ; Open command palette
    Sleep 200
    Send "show history"
    Sleep 200
    Send "{Enter}"
}

; Shift + I : Paste Image - Image
+i:: Send "!y"

; Shift + G : Fold Git repos (SCM) - Git Fold (implementation below)

; Shift + Q : Search - Search (Q for Query)
+q:: Send "^+f"

; Shift + R : Open Bread Crumbs menu - Breadcrumbs (R for Route/breadcrumbs)
+r:: Send "+r"

; Shift + D : Git section - Git
+d:: Send "+d"

; Shift + Z : Close all editors - Close
+z::
{
    Send "+f"
}

; Shift + Y : Zen mode - Zen
+y:: Send "+z"

; Shift + P : Git Pull - Pull
+p:: Send "+c"

; Shift + V : Git Commit - Commit
+v:: Send "+v"

; Shift + B : Git Push - Push
+b:: Send "+b"

; ---------------------------------------------------------------------------
; Ctrl + M : 1) Generate immediately. 2) Yellow banner 5s (N to avoid push). 3) Wait 15s.
; 4) Focus Cursor, commit, push by default (skip push only if N pressed). 5) Open Git panel to verify.
; 6) Return to previous window (skipped if push so you can review). Little test
; ---------------------------------------------------------------------------
^M:: {
    global gCommitPushTargetWin
    global gCommitPushDecision
    hwnd := WinExist("A")
    if !hwnd
        return
    gCommitPushTargetWin := hwnd
    ; Default behavior is now: commit + push, unless user opts out.
    gCommitPushDecision := "push"

    ; 1. Trigger generation immediately (Ctrl+Alt+A)
    SoundPlay A_ScriptDir "\sounds\commit-start.wav"
    Send "^!."
    ShowCommitPushBanner()

    ; 2. Wait 15s; user can interact with any window
    Sleep 14000
    ; Handoff Stop Sign: warn + play pre-movement cue right before we
    ; regain focus on Cursor and finalize the commit submission.
    PlayPreMovementWarning("Cursor")

    ; 3. Focus Cursor IDE (save current foreground to return later)
    prevHwnd := WinExist("A")
    WinActivate("ahk_id " hwnd)
    if !WinWaitActive("ahk_id " hwnd, , 3)
        return

    ; 4. Execute commit and push if necessary (default: push; user can opt out)
    Send "+v"
    didPush := (gCommitPushDecision = "push")
    if (didPush) {
        Sleep 500
        Send "+b"
    }

    ; 5. Wait for git operations to complete, then open Git panel to verify
    if (didPush) {
        Sleep 4000
    } else {
        Sleep 1500
    }
    ; Send "+d"

    ; Decide whether to return to previous window: stay in Cursor if we pushed
    shouldReturn := !didPush
    gCommitPushDecision := ""

    ; 6. Return to previous screen (graceful error if window no longer exists)
    if (shouldReturn && prevHwnd && prevHwnd != hwnd) {
        if (!WinExist("ahk_id " prevHwnd)) {
            TrayTip("Commit Push", "Previous window no longer available; staying in Cursor.", "Iconi")
            SetTimer(() => TrayTip(), -5000)
        } else {
            try {
                WinActivate("ahk_id " prevHwnd)
                if (!WinWaitActive("ahk_id " prevHwnd, , 2)) {
                    TrayTip("Commit Push", "Could not switch back to previous window.", "Iconi")
                    SetTimer(() => TrayTip(), -5000)
                }
            } catch {
                TrayTip("Commit Push", "Could not switch back to previous window.", "Iconi")
                SetTimer(() => TrayTip(), -5000)
            }
        }
    }
}

; Global variable for commit push selector target window
global gCommitPushTargetWin := 0
; Global variable to store the user's push decision ("push" | "dont_push" | "")
global gCommitPushDecision := ""
; Global variable for non-blocking commit push banner GUI
global g_CommitPushBannerGui := ""
global g_CommitPushBannerBorderGui := ""

; Non-blocking banner: "Don't push? Press N within 5 seconds" (dark background, yellow accent border)
ShowCommitPushBanner() {
    global g_CommitPushBannerGui, g_CommitPushBannerBorderGui
    try {
        if IsObject(g_CommitPushBannerBorderGui) && g_CommitPushBannerBorderGui.Hwnd
            g_CommitPushBannerBorderGui.Destroy()
    } catch {
    }
    g_CommitPushBannerBorderGui := ""
    try {
        if IsObject(g_CommitPushBannerGui) && g_CommitPushBannerGui.Hwnd
            g_CommitPushBannerGui.Destroy()
    } catch {
    }
    bannerGui := Gui("+AlwaysOnTop -Caption +ToolWindow")
    bannerGui.BackColor := "1E1E2E"
    bannerGui.SetFont("s14 cFFFFFF Bold", "Segoe UI")
    bannerGui.Add("Text", "w400 Center", "Don't push? Press N within 5 seconds")
    activeWin := WinGetID("A")
    if (activeWin)
        WinGetPos(&winX, &winY, &winW, &winH, activeWin)
    else {
        workArea := SysGet.MonitorWorkArea(SysGet.MonitorPrimary)
        winX := workArea.Left
        winY := workArea.Top
        winW := workArea.Right - workArea.Left
        winH := workArea.Bottom - workArea.Top
    }
    bannerGui.Show("AutoSize Hide")
    guiW := 0
    guiH := 0
    bannerGui.GetPos(, , &guiW, &guiH)
    guiX := winX + (winW - guiW) / 2
    guiY := winY + (winH - guiH) / 2
    borderWidth := 6
    borderGui := Gui("+AlwaysOnTop -Caption +ToolWindow")
    borderGui.BackColor := BANNER_ACCENT_INTERMEDIATE
    borderGui.Show("NA x" . Round(guiX - borderWidth) . " y" . Round(guiY - borderWidth) . " w" . (guiW + 2 *
        borderWidth) . " h" . (guiH + 2 * borderWidth))
    g_CommitPushBannerBorderGui := borderGui
    bannerGui.Show("x" . Round(guiX) . " y" . Round(guiY) . " NA")
    WinSetTransparent(220, bannerGui)
    g_CommitPushBannerGui := bannerGui
    Hotkey("n", CommitPushBanner_NHandler, "On")
    Hotkey("N", CommitPushBanner_NHandler, "On")
    SetTimer(CloseCommitPushBanner, -5000)
}

CommitPushBanner_NHandler(*) {
    global gCommitPushDecision
    ; User explicitly opted out of pushing; keep commit-only.
    gCommitPushDecision := "dont_push"
    CloseCommitPushBanner()
}

CloseCommitPushBanner() {
    global g_CommitPushBannerGui, g_CommitPushBannerBorderGui
    try {
        if IsObject(g_CommitPushBannerBorderGui) && g_CommitPushBannerBorderGui.Hwnd {
            g_CommitPushBannerBorderGui.Destroy()
            g_CommitPushBannerBorderGui := ""
        }
    } catch {
    }
    try {
        if IsObject(g_CommitPushBannerGui) && g_CommitPushBannerGui.Hwnd {
            g_CommitPushBannerGui.Destroy()
            g_CommitPushBannerGui := ""
        }
    } catch {
    }
    try Hotkey("n", "Off")
    catch {
    }
    try Hotkey("N", "Off")
    catch {
    }
    SetTimer(CloseCommitPushBanner, 0)
}

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

; Execute stored decision at the exact current push moment
ExecuteStoredCommitPushDecision() {
    global gCommitPushDecision
    global gCommitPushTargetWin
    if (gCommitPushDecision = "push") {
        ; Wait a moment for Cursor to process the commit
        Sleep 500
        ; Ensure the intended window has focus before sending the push hotkey
        if (gCommitPushTargetWin) {
            if (WinExist("ahk_id " gCommitPushTargetWin)) {
                WinActivate gCommitPushTargetWin
                WinWaitActive("ahk_id " gCommitPushTargetWin, , 2)
                Sleep 200
            } else {
                ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
            }
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
    emojiMap[4] := "✅"
    emojiMap[5] := "❓"
    return (emojiMap.Has(number)) ? emojiMap[number] : ""
}

InsertEmojiToTarget(emoji) {
    if (emoji = "")
        return
    ; Activate the target window if we have it stored
    if (gEmojiTargetWin) {
        if (WinExist("ahk_id " gEmojiTargetWin)) {
            WinActivate gEmojiTargetWin
            Sleep 150
        } else {
            ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
        }
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

; Shift + O : Emoji selector (Auto-submit version) - Emoji
+o::
{
    try {
        ; Remember current target window before showing GUI
        gEmojiTargetWin := WinExist("A")
        ; Create GUI for emoji selection with auto-submit
        emojiGui := Gui("+AlwaysOnTop +ToolWindow", "Emoji Selector")
        emojiGui.SetFont("s10", "Segoe UI")

        ; Add instruction text
        emojiGui.AddText("w350 Center",
            "Select emoji to insert:`n`n1. 🔲 Tasks/Checklist items`n2. ⏳ Time-sensitive tasks`n3. ⚡ First priority`n4. ✅ Check`n5. ❓ Questions/Uncertain items`n`nType a number (1-5):"
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
        SendEscape(2)
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
                SendEscape()
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

; Shift + A : Switch AI models - AI
+a::^;

; Shift + G : Fold all Git directories in Source Control (Cursor) - Git Fold
+g:: FoldAllGitDirectoriesInCursor()

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

; Alt + N : Review next file - Click the button that contains "Review next file" (Type 50020 Text)
; Path from UIA tree: workbench.parts.editor -> editor-instance -> ... -> Group (anysphere-text-button) -> Text "Review next file"
!n::
{
    try {
        win := WinExist("A")
        if (!win) {
            return
        }
        root := UIA.ElementFromHandle(win)
        Sleep 100  ; Allow UI to update

        ; Strategy 1: Scope to editor part (workbench.parts.editor), find "Review next file" Text, then click its parent Group (the button)
        editorPart := ""
        try editorPart := root.FindFirst({ AutomationId: "workbench.parts.editor", Type: 50026 })
        if (editorPart) {
            reviewText := ""
            try reviewText := editorPart.FindFirst({ Name: "Review next file", Type: 50020 })
            if (reviewText) {
                try {
                    parentBtn := UIA.TreeWalkerTrue.GetParentElement(reviewText)
                    if (parentBtn) {
                        try {
                            if parentBtn.GetPropertyValue(UIA.Property.IsInvokePatternAvailable) {
                                parentBtn.InvokePattern.Invoke()
                            } else {
                                parentBtn.Click()
                            }
                            return
                        } catch {
                            try reviewText.Click()
                            return
                        }
                    }
                } catch {
                    try reviewText.Click()
                    return
                }
            }
        }

        ; Strategy 2: Root-level find by Name "Review next file" or "Review" (Type 50020)
        reviewEl := root.FindFirst({ Name: "Review next file", Type: 50020 })
        if !reviewEl {
            reviewEl := root.FindFirst({ Name: "Review", Type: 50020 })
        }
        if !reviewEl {
            allTexts := root.FindAll({ Type: 50020 })
            for text in allTexts {
                name := ""
                try name := text.Name
                if (name = "Review" || name = "Review next file" || InStr(name, "Review next file")) {
                    reviewEl := text
                    break
                }
            }
        }

        if (reviewEl) {
            ; Prefer clicking parent (the button Group) so the clickable area is used
            try {
                parentBtn := UIA.TreeWalkerTrue.GetParentElement(reviewEl)
                if (parentBtn) {
                    try parentBtn.Click()
                    catch {
                        try reviewEl.Click()
                    }
                    return
                }
            } catch {
            }
            try {
                if reviewEl.GetPropertyValue(UIA.Property.IsInvokePatternAvailable) {
                    reviewEl.InvokePattern.Invoke()
                } else {
                    reviewEl.Click()
                }
            } catch {
                try reviewEl.Click()
            }
        }
    } catch Error as e {
        ; If all else fails, silently fail (no fallback action defined)
    }
}

; Alt + U : Scroll AI feed to bottom (AutoHotkey-based shortcut)
!u::
{
    try {
        hwnd := WinExist("A")
        if (!hwnd) {
            return
        }
        root := UIA.ElementFromHandle(hwnd)

        chatContainer := root.FindFirst({ ClassName: "composer-messages-container" })
        if (chatContainer) {
            try {
                if (chatContainer.GetPropertyValue(UIA.Property.IsScrollPatternAvailable)) {
                    chatContainer.ScrollPattern.SetScrollPercent(-1, 100)
                    return
                }
            } catch {
            }

            messages := chatContainer.FindAll({ ClassName: "composer-rendered-message", matchmode: "Substring" })
            if (messages && messages.Length > 0) {
                messages[messages.Length].ScrollIntoView()
            }
        }
    } catch Error as e {
    }
}

; Easy Selection
; Alt + 1 : Easy Selection - 1st item
!1::
{
    Send "{Enter}"
}

; Alt + 2 : Easy Selection - 2nd item
!2::
{
    Send "{Down}"
    Send "{Enter}"
}

; Alt + 3 : Easy Selection - 3rd item
!3::
{
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Alt + 4 : Easy Selection - 4th item
!4::
{
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Alt + 5 : Easy Selection - 5th item
!5::
{
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

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
        ; Show progress overlay immediately (yellow for folding)
        StandardLoadingBar_Show("📁 Folding directories...", BANNER_ACCENT_INTERMEDIATE)

        hwnd := WinExist("A")
        if !hwnd {
            StandardLoadingBar_Hide(0)
            return
        }
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

        ; Fallback: Try finding by specific Name "Files Explorer" (common in Cursor/VSCode)
        if !fileTree {
            try {
                feEn := UIA.CreatePropertyCondition(UIA.Property.Name, "Files Explorer")
                fePt := UIA.CreatePropertyCondition(UIA.Property.Name, "Explorador de Arquivos")
                feName := UIA.CreateOrCondition(feEn, fePt)
                feCond := UIA.CreateAndCondition(treeType, feName)
                fileTree := root.FindElement(feCond, UIA.TreeScope.Descendants)
            }
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

        ; Get all TreeItem nodes that support expand/collapse (i.e., directories) AND are currently Expanded
        itemType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.TreeItem)
        canExpand := UIA.CreatePropertyCondition(UIA.Property.IsExpandCollapsePatternAvailable, true)
        isExpanded := UIA.CreatePropertyCondition(UIA.Property.ExpandCollapseExpandCollapseState, UIA.ExpandCollapseState
            .Expanded)

        ; Combine conditions: TreeItem AND CanExpand AND IsExpanded
        dirCond := UIA.CreateAndCondition(itemType, UIA.CreateAndCondition(canExpand, isExpanded))

        ; Loop 2 times to ensure all nested directories are collapsed (reduced from 3)
        loop 2 {
            ; Re-find items each iteration as tree structure may change after collapsing
            items := fileTree.FindElements(dirCond, UIA.TreeScope.Descendants)

            if !items || !items.Length
                break

            ; Collapse each expanded directory.
            for item in items {
                if !item
                    continue
                try {
                    ; Since we filtered by Expanded, we know it's expanded (or was when found)
                    pat := item.ExpandCollapsePattern

                    ; Method 2: UIA Collapse Pattern (Primary method) – slowed down to reduce UI stress
                    try {
                        pat.Collapse()
                        Sleep 40
                    } catch {
                    }

                    ; Check if it worked (only check if we really need to try other methods)
                    if item.ExpandCollapsePattern.ExpandCollapseState != UIA.ExpandCollapseState.Collapsed {
                        ; Method 1: Scroll into view (if needed)
                        try {
                            if item.GetPropertyValue(UIA.Property.IsScrollItemPatternAvailable)
                                item.ScrollItemPattern.ScrollIntoView()
                            pat.Collapse()
                        } catch {
                        }
                    }

                    if item.ExpandCollapsePattern.ExpandCollapseState != UIA.ExpandCollapseState.Collapsed {
                        ; Method 3: Keyboard Navigation – slowed down to reduce UI stress
                        try {
                            if item.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable)
                                item.SelectionItemPattern.Select()
                            else
                                item.SetFocus()
                            Send "{Left}"
                            Sleep 40
                        } catch {
                        }
                    }

                    if item.ExpandCollapsePattern.ExpandCollapseState != UIA.ExpandCollapseState.Collapsed {
                        ; Method 4: Click Chevron
                        try {
                            btnType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Button)
                            txtType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Text)
                            glyphName := UIA.CreatePropertyCondition(UIA.Property.Name, "îª´")
                            dotName := UIA.CreatePropertyCondition(UIA.Property.Name, ".")
                            chevronCond := UIA.CreateOrCondition(btnType, UIA.CreateOrCondition(UIA.CreateAndCondition(
                                txtType,
                                glyphName), UIA.CreateAndCondition(txtType, dotName)))
                            chevron := ""
                            try chevron := item.FindElement(chevronCond, UIA.TreeScope.Children)
                            if !chevron
                                try chevron := item.FindElement(chevronCond, UIA.TreeScope.Descendants)
                            if chevron {
                                if chevron.GetPropertyValue(UIA.Property.IsInvokePatternAvailable) {
                                    chevron.InvokePattern.Invoke()
                                } else {
                                    chevron.Click()
                                }
                            }
                        } catch {
                        }
                    }
                } catch Error as e {
                }
                Sleep 5
            }

            ; Brief pause between iterations to allow UI to update
            Sleep 20
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
        StandardLoadingBar_Update("Directories folded")
    } catch Error as e {
        try MsgBox "UIA error folding Explorer directories: " e.Message, "Cursor Explorer Fold", "IconX"
    } finally {
        StandardLoadingBar_Hide(800)
    }
}

; Expand all expandable directories in the Explorer (FileExplorer3) for all workspace roots
UnfoldAllDirectoriesInExplorer() {
    try {
        ; Show progress overlay immediately (yellow for unfolding)
        StandardLoadingBar_Show("📁 Unfolding directories...", BANNER_ACCENT_INTERMEDIATE)

        hwnd := WinExist("A")
        if !hwnd {
            StandardLoadingBar_Hide(0)
            return
        }
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

        ; Fallback: Try finding by specific Name "Files Explorer"
        if !fileTree {
            try {
                feEn := UIA.CreatePropertyCondition(UIA.Property.Name, "Files Explorer")
                fePt := UIA.CreatePropertyCondition(UIA.Property.Name, "Explorador de Arquivos")
                feName := UIA.CreateOrCondition(feEn, fePt)
                feCond := UIA.CreateAndCondition(treeType, feName)
                fileTree := root.FindElement(feCond, UIA.TreeScope.Descendants)
            }
        }

        if !fileTree {
            StandardLoadingBar_Hide(0)
            return
        }

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

        ; Loop 3 times to ensure nested directories are expanded
        loop 3 {
            items := ""
            try items := fileTree.FindElements(dirCond, UIA.TreeScope.Descendants)
            if !items
                break

            ; Expand each collapsed directory. Do not toggle; skip already expanded.
            for item in items {
                if !item
                    continue
                try {
                    pat := item.ExpandCollapsePattern
                    state := pat.ExpandCollapseState

                    if state == UIA.ExpandCollapseState.Collapsed {
                        ; Method 1: Scroll into view
                        try {
                            if item.GetPropertyValue(UIA.Property.IsScrollItemPatternAvailable)
                                item.ScrollItemPattern.ScrollIntoView()
                        } catch {
                        }

                        ; Method 2: UIA Expand Pattern – slowed down to reduce UI stress
                        try {
                            pat.Expand()
                            Sleep 120
                        } catch {
                        }

                        ; Check if it worked
                        if item.ExpandCollapsePattern.ExpandCollapseState == UIA.ExpandCollapseState.Collapsed {
                            ; Method 3: Keyboard Navigation (Select + Right Arrow) – slowed down to reduce UI stress
                            try {
                                if item.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable)
                                    item.SelectionItemPattern.Select()
                                else
                                    item.SetFocus()
                                Send "{Right}"
                                Sleep 120
                            } catch {
                            }
                        }

                        ; Check if it worked
                        if item.ExpandCollapsePattern.ExpandCollapseState == UIA.ExpandCollapseState.Collapsed {
                            ; Method 4: Click Chevron (Moved from catch block)
                            try {
                                btnType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Button)
                                txtType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Text)
                                glyphName := UIA.CreatePropertyCondition(UIA.Property.Name, "îª´")
                                dotName := UIA.CreatePropertyCondition(UIA.Property.Name, ".")
                                chevronCond := UIA.CreateOrCondition(btnType, UIA.CreateOrCondition(UIA.CreateAndCondition(
                                    txtType,
                                    glyphName), UIA.CreateAndCondition(txtType, dotName)))
                                chevron := ""
                                try chevron := item.FindElement(chevronCond, UIA.TreeScope.Children)
                                if !chevron {
                                    try chevron := item.FindElement(chevronCond, UIA.TreeScope.Descendants)
                                }
                                if chevron {
                                    if chevron.GetPropertyValue(UIA.Property.IsInvokePatternAvailable) {
                                        chevron.InvokePattern.Invoke()
                                    } else {
                                        chevron.Click()
                                    }
                                }
                            } catch {
                            }
                        }
                    }
                } catch Error {
                    ; Fallback in case getting pattern fails completely
                    try {
                        btnType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Button)
                        txtType := UIA.CreatePropertyCondition(UIA.Property.ControlType, UIA.Type.Text)
                        glyphName := UIA.CreatePropertyCondition(UIA.Property.Name, "îª´")
                        dotName := UIA.CreatePropertyCondition(UIA.Property.Name, ".")
                        chevronCond := UIA.CreateOrCondition(btnType, UIA.CreateOrCondition(UIA.CreateAndCondition(
                            txtType,
                            glyphName), UIA.CreateAndCondition(txtType, dotName)))
                        chevron := ""
                        try chevron := item.FindElement(chevronCond, UIA.TreeScope.Children)
                        if !chevron {
                            try chevron := item.FindElement(chevronCond, UIA.TreeScope.Descendants)
                        }
                        if chevron {
                            if chevron.GetPropertyValue(UIA.Property.IsInvokePatternAvailable) {
                                chevron.InvokePattern.Invoke()
                            } else {
                                chevron.Click()
                            }
                        }
                    } catch {
                    }
                }
                Sleep 10
            }

            ; Brief pause between iterations to allow UI to update
            Sleep 50
        }

        ; Restore scroll position if it changed
        if hasScroll && (hPerc != "" && vPerc != "") {
            try fileTree.ScrollPattern.SetScrollPercent(hPerc, vPerc)
        }

        ; Optional brief toast
        StandardLoadingBar_Update("Directories unfolded")
    } catch Error as e {
        try MsgBox "UIA error unfolding Explorer directories: " e.Message, "Cursor Explorer Unfold", "IconX"
    } finally {
        StandardLoadingBar_Hide(800)
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
        SendEscape(2)
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
        SendEscape(2)
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
                SendEscape()
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
#HotIf WinActive("ahk_exe Spotify.exe") && !IsFileDialogActive()

; Shift + C : Toggle Connect to a device - Connect to a device
+c::
{
    ; #region agent log
    SafeDebugLog Format(
        "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:7574`",`"message`":`"Connect button handler entry`",`"data`":{},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"A`"}`n",
        A_TickCount, Random(1000, 9999), A_TickCount)
    ; #endregion
    try {
        spot := UIA_Browser("ahk_exe Spotify.exe")
        Sleep 300

        ; Find and click the Connect button
        connectPattern := "i)^(Connect to a device|Conectar a um dispositivo|Connect)$"
        ; #region agent log
        SafeDebugLog Format(
            "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:7582`",`"message`":`"Before WaitForButton call`",`"data`":{`"pattern`":`"{4}`"},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"B`"}`n",
            A_TickCount, Random(1000, 9999), A_TickCount, connectPattern)
        ; #endregion
        if (connectBtn := WaitForButton(spot, connectPattern)) {
            ; #region agent log
            btnName := "", btnType := "", btnClassName := "", btnLocalizedType := "", supportsInvoke := false,
                supportsToggle := false
            try btnName := connectBtn.Name
            try btnType := connectBtn.ControlType
            try btnClassName := connectBtn.ClassName
            try btnLocalizedType := connectBtn.LocalizedControlType
            try supportsInvoke := connectBtn.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)
            try supportsToggle := connectBtn.GetPropertyValue(UIA.Property.IsTogglePatternAvailable)
            SafeDebugLog Format(
                "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:7583`",`"message`":`"Button found before Invoke`",`"data`":{`"name`":`"{4}`",`"type`":`"{5}`",`"className`":`"{6}`",`"localizedType`":`"{7}`",`"supportsInvoke`":{8},`"supportsToggle`":{9}},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"A,C`"}`n",
                A_TickCount, Random(1000, 9999), A_TickCount, btnName, btnType, btnClassName, btnLocalizedType,
                supportsInvoke, supportsToggle)
            ; #endregion
            ; Try multi-strategy activation: prefer Invoke when available, fallback to Click
            clicked := false
            if (supportsInvoke) {
                try {
                    connectBtn.Invoke()
                    clicked := true
                    ; #region agent log
                    SafeDebugLog Format(
                        "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:7663`",`"message`":`"Invoke() succeeded`",`"data`":{},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"A`"}`n",
                        A_TickCount, Random(1000, 9999), A_TickCount)
                    ; #endregion
                } catch Error as invokeErr {
                    ; #region agent log
                    SafeDebugLog Format(
                        "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:7663`",`"message`":`"Invoke() failed, trying Click()`",`"data`":{`"error`":`"{4}`"},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"A,E`"}`n",
                        A_TickCount, Random(1000, 9999), A_TickCount, invokeErr.Message)
                    ; #endregion
                }
            }
            if (!clicked) {
                try {
                    connectBtn.Click()
                    clicked := true
                    ; #region agent log
                    SafeDebugLog Format(
                        "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:7675`",`"message`":`"Click() succeeded`",`"data`":{},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"A,E`"}`n",
                        A_TickCount, Random(1000, 9999), A_TickCount)
                    ; #endregion
                } catch Error as clickErr {
                    ; #region agent log
                    SafeDebugLog Format(
                        "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:7675`",`"message`":`"Click() failed`",`"data`":{`"error`":`"{4}`"},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"A,E`"}`n",
                        A_TickCount, Random(1000, 9999), A_TickCount, clickErr.Message)
                    ; #endregion
                    MsgBox "Error: " clickErr.Message
                    return
                }
            }

            ; After connecting, wait a moment for the device list to load
            Sleep 500

            ; Search for Office button (e.g., "Office Google Cast")
            officePattern := "i)Office"
            if (officeBtn := WaitForButton(spot, officePattern, 3000)) {
                officeBtn.Invoke()
            }
            ; If Office button not found, continue without error (as requested)
        }
        else {
            ; #region agent log
            SafeDebugLog Format(
                "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:7595`",`"message`":`"Button not found by WaitForButton`",`"data`":{`"pattern`":`"{4}`"},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"B,D`"}`n",
                A_TickCount, Random(1000, 9999), A_TickCount, connectPattern)
            ; #endregion
            MsgBox "Couldn't find the Connect-to-device button."
        }
    } catch Error as e {
        ; #region agent log
        SafeDebugLog Format(
            "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:7597`",`"message`":`"Exception caught`",`"data`":{`"error`":`"{4}`"},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"A,B,C,D,E`"}`n",
            A_TickCount, Random(1000, 9999), A_TickCount, e.Message)
        ; #endregion
        MsgBox "Error: " e.Message
    }
    ; #region agent log
    SafeDebugLog Format(
        "{`"id`":`"log_{1}_{2}`",`"timestamp`":{3},`"location`":`"Shift keys.ahk:7600`",`"message`":`"Connect button handler exit`",`"data`":{},`"sessionId`":`"debug-session`",`"runId`":`"run1`",`"hypothesisId`":`"A`"}`n",
        A_TickCount, Random(1000, 9999), A_TickCount)
    ; #endregion
}

; Shift + F : Toggle full screen - Fullscreen
+f::
{
    try {
        spot := UIA_Browser("ahk_exe Spotify.exe")
        Sleep 300

        ; Look for either Enter or Exit full screen button with case-insensitive pattern
        enterFsPattern := "i)^Enter Full[- ]?screen$"
        exitFsPattern := "i)^Exit Full[- ]?screen$"

        ; Helper function to click button with multi-strategy (Invoke or Click)
        ClickButton(btn) {
            if (!btn)
                return false
            supportsInvoke := false
            try {
                supportsInvoke := btn.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)
            } catch {
                supportsInvoke := false
            }
            clicked := false
            if (supportsInvoke) {
                try {
                    btn.Invoke()
                    clicked := true
                } catch {
                }
            }
            if (!clicked) {
                try {
                    btn.Click()
                    clicked := true
                } catch {
                }
            }
            return clicked
        }

        ; First attempt - immediate check
        enterFsBtn := WaitForButton(spot, enterFsPattern, 500)
        if (!enterFsBtn) {
            exitFsBtn := WaitForButton(spot, exitFsPattern, 500)
            if (!exitFsBtn) {
                ; Wait 1 second and try again
                Sleep 1000
                exitFsBtn := WaitForButton(spot, exitFsPattern, 500)
                if (exitFsBtn)
                    ClickButton(exitFsBtn)
            } else {
                ClickButton(exitFsBtn)
            }
        } else {
            ClickButton(enterFsBtn)
        }
    }
}

; Shift + S : Open Search - Search
+s:: Send "^k"

; Shift + P : Go to Playlists - Playlists
+p:: Send "!+1"

; Shift + A : Go to Artists - Artists
+a:: Send "!+3"

; Shift + B : Go to Albums - Albums
+b:: Send "!+4"

; Shift + H : Go to Home - Home
+h:: Send "!+h"

; Shift + N : Go to Now Playing - Now Playing
+n:: Send "!+j"

; Shift + M : Go to Made For You - Made For You
+m:: Send "!+m"

; Shift + R : Go to New Releases - Releases
+r:: Send "!+n"

; Shift + X : Go to Charts - Charts
+x:: Send "!+c"

; Shift + V : Toggle Now Playing View Sidebar - View
+v:: Send "!+r"

; Shift + L : Toggle Your Library Sidebar - Library
+l:: Send "!+l"

; Shift + E : Toggle Fullscreen Library - Expand Library
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

; Shift + Y : Toggle lyrics - Lyrics
+y:: Send("^+")

; Shift + T : Toggle Play/Pause - Play/Pause
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
        if (!WinExist("ahk_exe Spotify.exe")) {
            ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
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
#HotIf WinActive("ahk_exe Figma.exe") && !IsFileDialogActive()

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

; Shift + D : Dashboard
+d:: {
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

; Shift + A : Contas
+a:: {
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

; Shift + T : TransaÃ§Ãµes
+t:: {
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

; Shift + C : CartÃµes de crÃ©dito
+c:: {
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

; Shift + R : RelatÃ³rios
+r:: {
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

; Shift + M : Mais opÃ§Ãµes
+m:: {
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
+k:: Mobills_Navigate("Prev")

; Shift + L : Next month
+l:: Mobills_Navigate("Next")

; K : Previous month
k:: Mobills_Navigate("Prev")

; L : Next month
l:: Mobills_Navigate("Next")

; =============================================================================
; Mobills pagination (unified)
; =============================================================================
Mobills_GetContext(uia) {
    try {
        url := uia.GetCurrentURL()
        url := StrLower(url)
    } catch {
        url := ""
    }
    if InStr(url, "/transactions")
        return "transactions"
    if InStr(url, "/accounts")
        return "accounts"
    ; Budgets has its own pager controls (avoid selecting month/year text)
    if InStr(url, "/budgets")
        return "budgets"
    if InStr(url, "/planning")
        return "planning"
    return "unknown"
}

Mobills_IsDisabled(el) {
    try {
        cls := ""
        try cls := el.GetPropertyValue(UIA.Property.ClassName)
        if (cls != "" && InStr(cls, "Mui-disabled"))
            return true
    } catch {
    }
    return false
}

Mobills_IsButton(el) {
    if !el
        return false
    try {
        ct := el.GetPropertyValue(UIA.Property.ControlType)
        return (ct = UIA.Type.Button || ct = 50000)
    } catch {
        ; Some wrappers expose .Type directly
        try {
            return (el.Type = 50000 || el.Type = UIA.Type.Button)
        } catch {
            return false
        }
    }
}

Mobills_FindElementByCandidates(uia, candidates) {
    if !uia
        return ""
    for , c in candidates {
        try {
            el := uia.FindElement(c)
            if el
                return el
        } catch {
        }
    }
    return ""
}

Mobills_FindOpenButton(uia, index := 1) {
    if !uia
        return ""
    try {
        openButtons := uia.FindAll({ Name: "Open", Type: 50000 })
        if (openButtons && openButtons.Length >= index)
            return openButtons[index]
    } catch {
    }
    return ""
}

; Budgets: resolve BOTH arrows by index, then pick by dir.
; Target:
;   Prev => {T:30}, {T:26}, {T:0, i:7}
;   Next => {T:30}, {T:26}, {T:0, i:8}
Mobills_GetBudgetsPrevNext(uia, &prevBtn, &nextBtn) {
    prevBtn := ""
    nextBtn := ""

    prev := ""
    next := ""
    try {
        prev := uia.ElementFromPath({ Type: 30 }, { Type: 26 }, { Type: 0, i: 7 })
    } catch {
        prev := ""
    }
    try {
        next := uia.ElementFromPath({ Type: 30 }, { Type: 26 }, { Type: 0, i: 8 })
    } catch {
        next := ""
    }

    if (prev && Mobills_IsButton(prev) && !Mobills_IsDisabled(prev))
        prevBtn := prev
    if (next && Mobills_IsButton(next) && !Mobills_IsDisabled(next))
        nextBtn := next

    ; Sanity check: ensure prev is physically left of next (swap if needed).
    if (prevBtn && nextBtn) {
        try {
            p := prevBtn.Location
            n := nextBtn.Location
            if (p.x > n.x) {
                tmp := prevBtn
                prevBtn := nextBtn
                nextBtn := tmp
            }
        } catch {
        }
    }

    return (prevBtn || nextBtn)
}

Mobills_FindPagerByName(uia, dir) {
    ; Try common labels (EN/PT). Substring match.
    namesPrev := ["Go to previous page", "previous page", "Previous", "Prev", "Anterior", "Página anterior",
        "Ir para a página anterior"]
    namesNext := ["Go to next page", "next page", "Next", "Próximo", "Proximo", "Página seguinte",
        "Ir para a próxima página",
        "Ir para a proxima página"]
    names := (dir = "Prev") ? namesPrev : namesNext

    for , nm in names {
        try {
            btn := uia.FindElement({ Name: nm, Type: 50000, matchmode: "Substring" })
            if btn && !Mobills_IsDisabled(btn)
                return btn
        } catch {
        }
    }
    return ""
}

Mobills_FindPagerByMonthHeader(uia, dir) {
    grp := ""
    try grp := FindMonthGroup(uia)
    if !grp
        return ""

    step := (dir = "Prev") ? "-1" : "+1"
    try {
        btn := grp.WalkTree(step, { Type: "Button" })
        if btn && !Mobills_IsDisabled(btn)
            return btn
    } catch {
    }

    ; fallback: scan sibling buttons in parent and choose closest left/right
    try parent := UIA.TreeWalkerTrue.GetParentElement(grp)
    if !parent
        return ""

    try grpPos := grp.Location
    best := ""
    bestX := ""
    try {
        for , el in parent.FindAll({ Type: "Button" }) {
            if Mobills_IsDisabled(el)
                continue
            pos := el.Location
            sameRow := (pos.y >= grpPos.y - 10 && pos.y <= grpPos.y + grpPos.h + 10)
            if !sameRow
                continue
            if (dir = "Prev") {
                if (pos.x < grpPos.x) {
                    if (best = "" || pos.x > bestX) {
                        best := el
                        bestX := pos.x
                    }
                }
            } else {
                if (pos.x > grpPos.x + grpPos.w) {
                    if (best = "" || pos.x < bestX) {
                        best := el
                        bestX := pos.x
                    }
                }
            }
        }
    } catch {
    }
    return best
}

; Transactions month navigation: click the arrows adjacent to the month/year display
; (e.g. "<  February 2026  >"), NOT the table pagination ("previous page / next page").
Mobills_FindMonthNavByMonthYear(uia, dir) {
    if !uia
        return ""

    ; Find the visible month label (Text) in any supported language
    months := ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November",
        "December",
        "Janeiro", "Fevereiro", "Março", "Marco", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro",
        "Novembro", "Dezembro"]

    monthEl := ""
    for , m in months {
        try {
            el := uia.FindElement({ Name: m, Type: "Text", mm: 1, cs: false })
            if el {
                monthEl := el
                break
            }
        } catch {
        }
    }
    if !monthEl
        return ""

    ; Header container is usually a Group near the month label.
    header := ""
    try header := monthEl.WalkTree("p", { Type: "Group" })
    if !header
        return ""

    ; Use the month element bounds as the anchor; choose closest button left/right on the same row.
    try mPos := monthEl.Location
    buttons := ""
    try buttons := header.FindAll({ Type: "Button" })
    if !buttons
        return ""

    best := ""
    bestDist := ""
    try {
        for , b in buttons {
            if Mobills_IsDisabled(b)
                continue
            pos := b.Location
            ; same row as the month label
            if !(pos.y <= (mPos.y + mPos.h + 12) && (pos.y + pos.h) >= (mPos.y - 12))
                continue

            if (dir = "Prev") {
                if (pos.x + pos.w) >= mPos.x
                    continue
                dist := mPos.x - (pos.x + pos.w)
            } else {
                if pos.x <= (mPos.x + mPos.w)
                    continue
                dist := pos.x - (mPos.x + mPos.w)
            }

            if (best = "" || dist < bestDist) {
                best := b
                bestDist := dist
            }
        }
    } catch {
    }

    return best
}

Mobills_FindPagerByPath(uia, dir, context) {
    ; Budgets page: force the known arrow BUTTONs and avoid adjacent Text elements.
    ; Target:
    ;   Next  => {T:30}, {T:26}, {T:0, i:8}
    ;   Prev  => {T:30}, {T:26}, {T:0, i:7}
    if (context = "budgets") {
        if Mobills_GetBudgetsPrevNext(uia, &prevBtn, &nextBtn) {
            if (dir = "Prev")
                return prevBtn ? prevBtn : ""
            return nextBtn ? nextBtn : ""
        }
        ; IMPORTANT: Do not fall through on budgets (prevents picking the wrong control)
        return ""
    }

    ; Legacy paths (worked across some pages previously)
    try {
        if (dir = "Prev") {
            btn := uia.ElementFromPath({ Type: 30 }, { Type: 26 }, { Type: 26 }, { Type: 8 }, { Type: 7 }, { Type: 0 })
        } else {
            btn := uia.ElementFromPath({ Type: 30 }, { Type: 26 }, { Type: 26 }, { Type: 8 }, { Type: 7, i: -1 }, { Type: 0 })
        }
        if btn && Mobills_IsButton(btn) && !Mobills_IsDisabled(btn)
            return btn
    } catch {
    }

    ; Accounts: prefer toolbar-ish container if present (avoid hardcoding a single index)
    if (context = "accounts") {
        try {
            toolbar := uia.ElementFromPath({ Type: 30 }, { Type: 26 })
            if toolbar {
                buttons := toolbar.FindAll({ Type: "Button" })
                best := ""
                bestKey := ""
                for , b in buttons {
                    if Mobills_IsDisabled(b)
                        continue
                    pos := b.Location
                    ; prefer top-most row
                    key := pos.y * 100000 + pos.x
                    if (dir = "Prev") {
                        ; left-most among top candidates
                        if (best = "" || key < bestKey) {
                            best := b
                            bestKey := key
                        }
                    } else {
                        ; right-most among top candidates (approx: invert x)
                        key2 := pos.y * 100000 - pos.x
                        if (best = "" || key2 < bestKey) {
                            best := b
                            bestKey := key2
                        }
                    }
                }
                if best
                    return best
            }
        } catch {
        }
    }

    ; Planning: sometimes the "button" is a Text element
    if (context = "planning") {
        try {
            el := uia.ElementFromPath({ Type: 30 }, { Type: 26 }, { Type: 20, i: 2 })
            if el
                return el
        } catch {
        }
    }

    return ""
}

Mobills_ClickPager(el) {
    if !el
        return false
    try {
        ; Prefer Invoke when available, otherwise click.
        try {
            if (el.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)) {
                el.Invoke()
                return true
            }
        } catch {
        }
        el.Click()
        return true
    } catch {
        return false
    }
}

; =============================================================================
; Mobills "running" banner (non-blocking, overlay style; dark background, blue accent border)
; =============================================================================
global g_MobillsRunningBannerGui := 0
global g_MobillsRunningBannerBorderGui := 0

Mobills_ShowRunningBanner(dir) {
    global g_MobillsRunningBannerGui, g_MobillsRunningBannerBorderGui

    ; Close any existing banner first
    try {
        if IsObject(g_MobillsRunningBannerBorderGui)
            g_MobillsRunningBannerBorderGui.Destroy()
    } catch {
    }
    g_MobillsRunningBannerBorderGui := 0
    try {
        if IsObject(g_MobillsRunningBannerGui)
            g_MobillsRunningBannerGui.Destroy()
    } catch {
    }
    g_MobillsRunningBannerGui := 0

    text := "Mobills: " . ((dir = "Prev") ? "Previous" : "Next") . " (running...)"

    target := WinGetID("A")
    hasWindow := false
    if target && WinExist("ahk_id " target) {
        try {
            WinGetPos(&wx, &wy, &ww, &wh, target)
            hasWindow := (ww > 0 && wh > 0)
        } catch {
            hasWindow := false
        }
    }

    ov := Gui("+AlwaysOnTop -Caption +ToolWindow")
    ov.BackColor := "1E1E2E"
    ov.SetFont("s24 cFFFFFF Bold", "Segoe UI")
    ov.Add("Text", "w500 Center", text)
    ov.Show("AutoSize Hide")
    ov.GetPos(&gx, &gy, &gw, &gh)

    if hasWindow {
        cx := wx + (ww - gw) // 2
        cy := wy + (wh - gh) // 2
    } else {
        vx := SysGet(76)  ; SM_XVIRTUALSCREEN
        vy := SysGet(77)  ; SM_YVIRTUALSCREEN
        vw := SysGet(78)  ; SM_CXVIRTUALSCREEN
        vh := SysGet(79)  ; SM_CYVIRTUALSCREEN
        cx := vx + (vw - gw) // 2
        cy := vy + (vh - gh) // 2
    }

    borderWidth := 6
    borderGui := Gui("+AlwaysOnTop -Caption +ToolWindow")
    borderGui.BackColor := BANNER_ACCENT_INTERMEDIATE
    borderGui.Show("NA x" . (cx - borderWidth) . " y" . (cy - borderWidth) . " w" . (gw + 2 * borderWidth) . " h" . (gh +
        2 * borderWidth))
    g_MobillsRunningBannerBorderGui := borderGui

    ov.Show("x" . cx . " y" . cy . " NA")
    WinSetTransparent(178, ov)
    g_MobillsRunningBannerGui := ov
}

Mobills_HideRunningBanner() {
    global g_MobillsRunningBannerGui, g_MobillsRunningBannerBorderGui
    try {
        if IsObject(g_MobillsRunningBannerBorderGui)
            g_MobillsRunningBannerBorderGui.Destroy()
    } catch {
    }
    g_MobillsRunningBannerBorderGui := 0
    try {
        if IsObject(g_MobillsRunningBannerGui)
            g_MobillsRunningBannerGui.Destroy()
    } catch {
    }
    g_MobillsRunningBannerGui := 0
}

; One-shot pager resolution (no waiting/retries). Returns element or "".
Mobills_FindPagerOnce(uia, dir, context) {
    btn := ""
    ; Context-specific ordering (per plan)
    if (context = "transactions") {
        ; Prefer the arrows next to the month/year header (mobile/desktop), avoid table pagination.
        btn := Mobills_FindMonthNavByMonthYear(uia, dir)
        if !btn
            btn := Mobills_FindPagerByMonthHeader(uia, dir)
        if !btn
            btn := Mobills_FindPagerByName(uia, dir)
    } else if (context = "accounts") {
        btn := Mobills_FindPagerByPath(uia, dir, context)
        if !btn
            btn := Mobills_FindPagerByName(uia, dir)
    } else if (context = "budgets") {
        ; Budgets: ONLY use the known arrow buttons by index (avoid misclicking "Next").
        btn := Mobills_FindPagerByPath(uia, dir, context)
    } else if (context = "planning") {
        btn := Mobills_FindPagerByPath(uia, dir, context)
        if !btn
            btn := Mobills_FindPagerByName(uia, dir)
    } else {
        ; Generic fallback: try name -> month header -> legacy path
        btn := Mobills_FindPagerByName(uia, dir)
        if !btn
            btn := Mobills_FindPagerByMonthHeader(uia, dir)
        if !btn
            btn := Mobills_FindPagerByPath(uia, dir, context)
    }
    return btn
}

; Multi-layer verification to confirm absence (prevents false negatives).
; Returns an element if found in any layer; otherwise returns "".
Mobills_VerifyPagerMissing(dir, context, uiaCurrent := 0) {
    ; Layer 1: Re-check using the current attachment (cheap).
    try {
        if uiaCurrent {
            Sleep 120
            btn := Mobills_FindPagerOnce(uiaCurrent, dir, context)
            if btn
                return btn
        }
    } catch {
    }

    ; Layer 2: Fresh re-attach (new UIA tree) + wait.
    try {
        uia2 := TryAttachBrowser()
        if uia2 {
            Sleep 350
            btn := Mobills_FindPagerOnce(uia2, dir, context)
            if btn
                return btn
        }
    } catch {
    }

    ; Layer 3 (budgets only): re-attach + long wait + ONLY the known i:7/i:8 buttons.
    if (context = "budgets") {
        try {
            uia3 := TryAttachBrowser()
            if uia3 {
                Sleep 600
                if Mobills_GetBudgetsPrevNext(uia3, &prevBtn, &nextBtn) {
                    return (dir = "Prev") ? (prevBtn ? prevBtn : "") : (nextBtn ? nextBtn : "")
                }
            }
        } catch {
        }
    }

    return ""
}

Mobills_Navigate(dir) {
    Mobills_ShowRunningBanner(dir)
    try {
        uia := TryAttachBrowser()
        if !uia {
            MsgBox "Could not attach to the browser window.", "Mobills Navigation", "IconX"
            return
        }

        context := Mobills_GetContext(uia)
        maxRetries := 2
        retryDelay := 200

        loop maxRetries {
            if (A_Index > 1) {
                ; Refresh UIA tree on subsequent attempts (more reliable than just waiting).
                try uia := TryAttachBrowser()
            }
            btn := Mobills_FindPagerOnce(uia, dir, context)

            if btn && Mobills_ClickPager(btn)
                return

            Sleep retryDelay
        }

        ; Multi-layer verification before declaring "not present"
        verifiedBtn := Mobills_VerifyPagerMissing(dir, context, uia)
        if verifiedBtn {
            if Mobills_ClickPager(verifiedBtn)
                return
            MsgBox "Pager control was found but could not be clicked.", "Mobills Navigation", "IconX"
            return
        }

        MsgBox "Could not find the " . ((dir = "Prev") ? "previous" : "next") .
        " page/month control (verified missing).", "Mobills Navigation", "IconX"
    } catch Error as e {
        MsgBox "Error navigating Mobills:`n" e.Message, "Mobills Error", "IconX"
    } finally {
        Mobills_HideRunningBanner()
    }
}

; Backwards-compatible wrappers (old names kept, logic refactored)
PrevMobillsMonth() => Mobills_Navigate("Prev")
NextMobillsMonth() => Mobills_Navigate("Next")

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

; Shift + I : Toggle "Ignore transaction"
+i:: {
    try {
        uia := TryAttachBrowser()
        if !uia {
            MsgBox "Could not attach to the browser window.", "Mobills Navigation", "IconX"
            return
        }

        ignoreToggle := ""
        try {
            label := uia.FindElement({ Name: "Ignore transaction", Type: 50020, matchmode: "Substring" })
            if label {
                parent := UIA.TreeWalkerTrue.GetParentElement(label)
                if parent {
                    for , cb in parent.FindAll({ Type: 50002 }) {
                        ignoreToggle := cb
                        break
                    }
                }
            }
        } catch {
        }
        if !ignoreToggle {
            ; Fallback: pick the last checkbox in the dialog, which is the ignore toggle in current forms.
            try {
                checkboxes := uia.FindAll({ Type: 50002 })
                if (checkboxes && checkboxes.Length > 0)
                    ignoreToggle := checkboxes[checkboxes.Length]
            } catch {
            }
        }
        if !ignoreToggle {
            MsgBox "Could not find the Ignore transaction toggle.", "Mobills Navigation", "IconX"
            return
        }

        try ignoreToggle.SetFocus()
        Sleep 80
        Send "{Space}"

    } catch Error as e {
        MsgBox "Error toggling Ignore transaction: " e.Message, "Mobills Error", "IconX"
    }
}

; Click "New" and select the requested creation menu item.
Mobills_SelectNewMenuItem(itemName) {
    try {
        uia := TryAttachBrowser()
        if !uia
            return false

        actionBtn := Mobills_FindElementByCandidates(uia, [{ Type: 50000, AutomationId: "action-button" }, { Type: 50000,
            Name: "New", matchmode: "Substring" }])
        if !actionBtn
            return false
        actionBtn.Click()
        Sleep 250

        menuItem := Mobills_FindElementByCandidates(uia, [{ Type: 50011, Name: itemName, matchmode: "Substring" }, { Type: 50000,
            Name: itemName, matchmode: "Substring" }])
        if !menuItem
            return false
        menuItem.Click()
        return true
    } catch {
        return false
    }
}

; ---- Helper to focus the Description field directly ----
FocusDescriptionField() {
    try {
        uia := TryAttachBrowser()
        if !uia
            return false

        descriptionElement := Mobills_FindElementByCandidates(uia, [{ Name: "Description", Type: 50004, matchmode: "Substring" }, { Name: "Description",
            Type: "Edit", matchmode: "Substring" }, { ClassName: "MuiAutocomplete-input", Type: "Edit", matchmode: "Substring" }
        ])

        if !descriptionElement {
            MsgBox "Could not find the Description field.", "Mobills Navigation", "IconX"
            return false
        }

        try descriptionElement.Click()
        Sleep 100
        descriptionElement.SetFocus()
        return true
    } catch Error as e {
        MsgBox "Error focusing Description field: " e.Message, "Mobills Error", "IconX"
        return false
    }
}

; Focuses the first account/category "Open" picker and types MAIN.
Mobills_TypeMainInOpenPicker() {
    try {
        uia := TryAttachBrowser()
        if !uia
            return false

        openBtn := Mobills_FindOpenButton(uia, 1)
        if !openBtn
            return false

        try openBtn.Click()
        try openBtn.SetFocus()
        Sleep 120
        Send "MAIN"
        return true
    } catch {
        return false
    }
}

; Shift + N : Focus name/description field
+n:: FocusDescriptionField()

; Shift + E : Click action button then Expense menu item
+e:: {
    try {
        if !Mobills_SelectNewMenuItem("Expense")
            MsgBox "Could not open New > Expense.", "Mobills Navigation", "IconX"
    } catch Error as e {
        MsgBox "Error clicking action button and Expense menu: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + Y : Click action button then Income menu item
+y:: {
    try {
        if !Mobills_SelectNewMenuItem("Income")
            MsgBox "Could not open New > Income.", "Mobills Navigation", "IconX"
    } catch Error as e {
        MsgBox "Error clicking action button and Income menu: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + X : Click action button then Credit card expense menu item
+x:: {
    try {
        if !Mobills_SelectNewMenuItem("Credit card expense")
            MsgBox "Could not open New > Credit card expense.", "Mobills Navigation", "IconX"
    } catch Error as e {
        MsgBox "Error clicking action button and Credit card expense menu: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + F : Click action button then Transfer menu item
+f:: {
    try {
        if !Mobills_SelectNewMenuItem("Transfer")
            MsgBox "Could not open New > Transfer.", "Mobills Navigation", "IconX"
    } catch Error as e {
        MsgBox "Error clicking action button and Transfer menu: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + W : Focus "Open" picker and type "MAIN"
+w:: {
    try {
        if !Mobills_TypeMainInOpenPicker()
            MsgBox "Could not find the Open picker.", "Mobills Navigation", "IconX"
    } catch Error as e {
        MsgBox "Error finding Open picker: " e.Message, "Mobills Error", "IconX"
    }
}

; =============================================================================
; Mobills hotkeys fallback for mobile/device mode
; - Some Chrome "mobile device" views do not keep "Mobills" in the window title,
;   which prevents the WinActive("Mobills") context from matching.
; - This fallback scopes K/L (+Shift variants) to the Mobills Transactions URL only.
; =============================================================================
global g_MobillsUrlCacheTick := 0
global g_MobillsUrlCacheUrl := ""

Mobills_IsTransactionsUrlActive(cacheMs := 250) {
    global g_MobillsUrlCacheTick, g_MobillsUrlCacheUrl
    now := A_TickCount

    if (g_MobillsUrlCacheTick && (now - g_MobillsUrlCacheTick) < cacheMs) {
        return InStr(g_MobillsUrlCacheUrl, "/transactions")
    }

    g_MobillsUrlCacheTick := now
    g_MobillsUrlCacheUrl := ""

    try {
        ; Prefer Chrome if active, else try Edge.
        if WinActive("ahk_exe chrome.exe")
            uia := UIA_Browser("ahk_exe chrome.exe")
        else if WinActive("ahk_exe msedge.exe")
            uia := UIA_Browser("ahk_exe msedge.exe")
        else
            uia := ""

        if uia {
            try g_MobillsUrlCacheUrl := StrLower(uia.GetCurrentURL())
        }
    } catch {
        g_MobillsUrlCacheUrl := ""
    }

    return InStr(g_MobillsUrlCacheUrl, "/transactions")
}

#HotIf (WinActive("ahk_exe chrome.exe") || WinActive("ahk_exe msedge.exe")) && Mobills_IsTransactionsUrlActive()

; K/L and Shift+K/L: Previous/Next month (Transactions) even when title doesn't match
k:: Mobills_Navigate("Prev")
l:: Mobills_Navigate("Next")
+k:: Mobills_Navigate("Prev")
+l:: Mobills_Navigate("Next")

#HotIf

#HotIf

;-------------------------------------------------------------------
; Google Keep Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe chrome.exe") && (WinActive("Google Keep") || WinActive("keep.google.com") || InStr(
    WinGetTitle("A"), "Google Keep"))

; Shift + S : Search and select note
+s::
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
        SendEscape()
        Sleep 300

        ; Open search with Ctrl+F
        Send "^f"
        Sleep 200

        ; Paste the search text
        Send "^v"
        Sleep 900

        ; Press Escape to close search
        SendEscape()
        Sleep 300

        ; Press Enter to confirm selection
        Send "{Enter}"
        Sleep 300

        ; Restore original clipboard
        A_Clipboard := oldClip
    }
}

; Shift + M : Toggle main menu
+m::
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
; When Mobills is visible but not focused, UIA_Browser attach can fail.
; Recovery: activate browser window, retry once. (No clicking.)
Mobills_ActivateBrowserForAttach(exe := "ahk_exe chrome.exe", titleNeedle := "Mobills") {
    try {
        bestHwnd := 0
        hwnds := WinGetList(exe)
        for hwnd in hwnds {
            try {
                t := WinGetTitle("ahk_id " hwnd)
                if (t != "" && InStr(t, titleNeedle)) {
                    bestHwnd := hwnd
                    break
                }
            } catch {
            }
        }
        if (!bestHwnd && hwnds.Length)
            bestHwnd := hwnds[1]
        if (!bestHwnd)
            return false

        WinActivate("ahk_id " bestHwnd)
        WinWaitActive("ahk_id " bestHwnd, , 1)
        return true
    } catch {
        return false
    }
}

TryAttachBrowser() {
    ; #region agent log
    try {
        FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' . A_TickCount .
        ',"location":"Shift keys.ahk:9806","message":"TryAttachBrowser entry","data":{"attempt":"chrome"},"sessionId":"debug-session","runId":"run1","hypothesisId":"A"}`n',
        DEBUG_LOG_PATH
    } catch {
    }
    ; #endregion
    ; Try Chrome first, then Edge (with one focus recovery retry)
    try {
        result := ""
        try result := UIA_Browser("ahk_exe chrome.exe")
        if (result) {
            ; #region agent log
            try {
                FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' . A_TickCount .
                ',"location":"Shift keys.ahk:9814","message":"TryAttachBrowser success","data":{"browser":"chrome","result":' .
                (result ? 1 : 0) . '},"sessionId":"debug-session","runId":"run1","hypothesisId":"A"}`n', DEBUG_LOG_PATH
            } catch {
            }
            ; #endregion
            return result
        }

        ; Recovery: activate + retry
        Mobills_ActivateBrowserForAttach("ahk_exe chrome.exe", "Mobills")
        try result := UIA_Browser("ahk_exe chrome.exe")
        ; #region agent log
        try {
            FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' . A_TickCount .
            ',"location":"Shift keys.ahk:9814","message":"TryAttachBrowser success","data":{"browser":"chrome","result":' .
            (result ? 1 : 0) . '},"sessionId":"debug-session","runId":"run1","hypothesisId":"A"}`n', DEBUG_LOG_PATH
        } catch {
        }
        ; #endregion
        return result
    }
    catch {
        ; #region agent log
        try {
            FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' . A_TickCount .
            ',"location":"Shift keys.ahk:9822","message":"TryAttachBrowser chrome failed, trying edge","data":{"error":"' .
            A_LastError . '"},"sessionId":"debug-session","runId":"run1","hypothesisId":"A"}`n', DEBUG_LOG_PATH
        } catch {
        }
        ; #endregion
        try {
            result := ""
            try result := UIA_Browser("ahk_exe msedge.exe")
            if (!result) {
                Mobills_ActivateBrowserForAttach("ahk_exe msedge.exe", "Mobills")
                try result := UIA_Browser("ahk_exe msedge.exe")
            }
            ; #region agent log
            try {
                FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' . A_TickCount .
                ',"location":"Shift keys.ahk:9829","message":"TryAttachBrowser success","data":{"browser":"edge","result":' .
                (result ? 1 : 0) . '},"sessionId":"debug-session","runId":"run1","hypothesisId":"A"}`n', DEBUG_LOG_PATH
            } catch {
            }
            ; #endregion
            return result
        }
        catch {
            ; #region agent log
            try {
                FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' . A_TickCount .
                ',"location":"Shift keys.ahk:9837","message":"TryAttachBrowser failed both browsers","data":{"error":"' .
                A_LastError . '"},"sessionId":"debug-session","runId":"run1","hypothesisId":"A"}`n', DEBUG_LOG_PATH
            } catch {
            }
            ; #endregion
            return 0
        }
    }
}

FindMonthGroup(uia) {
    ; #region agent log
    try {
        FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' . A_TickCount .
        ',"location":"Shift keys.ahk:9848","message":"FindMonthGroup entry","data":{"uia":' . (uia ? 1 : 0) .
        '},"sessionId":"debug-session","runId":"run1","hypothesisId":"B"}`n', DEBUG_LOG_PATH
    } catch {
    }
    ; #endregion
    ; Strategy 1 â€" look for known class name on the container
    try {
        ; #region agent log
        try {
            FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' . A_TickCount .
            ',"location":"Shift keys.ahk:9855","message":"FindMonthGroup Strategy 1 attempt","data":{"className":"sc-kAyceB"},"sessionId":"debug-session","runId":"run1","hypothesisId":"B"}`n',
            DEBUG_LOG_PATH
        } catch {
        }
        ; #endregion
        grp := uia.FindElement({ Type: "Group", ClassName: "sc-kAyceB", matchmode: "Substring" })
        if grp {
            ; #region agent log
            try {
                FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' . A_TickCount .
                ',"location":"Shift keys.ahk:9862","message":"FindMonthGroup Strategy 1 success","data":{"found":true},"sessionId":"debug-session","runId":"run1","hypothesisId":"B"}`n',
                DEBUG_LOG_PATH
            } catch {
            }
            ; #endregion
            return grp
        }
        ; #region agent log
        try {
            FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' . A_TickCount .
            ',"location":"Shift keys.ahk:9862","message":"FindMonthGroup Strategy 1 no result","data":{"found":false},"sessionId":"debug-session","runId":"run1","hypothesisId":"B"}`n',
            DEBUG_LOG_PATH
        } catch {
        }
        ; #endregion
    }
    catch Error as e {
        ; #region agent log
        try {
            FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' . A_TickCount .
            ',"location":"Shift keys.ahk:9876","message":"FindMonthGroup Strategy 1 exception","data":{"error":"' . e.Message .
            '"},"sessionId":"debug-session","runId":"run1","hypothesisId":"B"}`n', DEBUG_LOG_PATH
        } catch {
        }
        ; #endregion
    }
    ; Strategy 2 â€" locate by month text (any language)
    ; #region agent log
    try {
        FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' . A_TickCount .
        ',"location":"Shift keys.ahk:9883","message":"FindMonthGroup Strategy 2 attempt","data":{"monthCount":14},"sessionId":"debug-session","runId":"run1","hypothesisId":"C"}`n',
        DEBUG_LOG_PATH
    } catch {
    }
    ; #endregion
    months := ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
        "November", "December",
        "Janeiro", "Fevereiro", "MarÃ§o", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro",
        "Novembro",
        "Dezembro"]
    foundMonths := []
    for , m in months {
        try {
            el := uia.FindElement({ Name: m, Type: "Text", mm: 1, cs: false })
            if el {
                ; #region agent log
                try {
                    FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' . A_TickCount .
                    ',"location":"Shift keys.ahk:9897","message":"FindMonthGroup found month text","data":{"month":"' .
                    m . '"},"sessionId":"debug-session","runId":"run1","hypothesisId":"C"}`n', DEBUG_LOG_PATH
                } catch {
                }
                ; #endregion
                grp := el.WalkTree("p", { Type: "Group" })
                if grp {
                    ; #region agent log
                    try {
                        FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' .
                        A_TickCount .
                        ',"location":"Shift keys.ahk:9904","message":"FindMonthGroup Strategy 2 success","data":{"month":"' .
                        m . '","found":true},"sessionId":"debug-session","runId":"run1","hypothesisId":"C"}`n',
                        DEBUG_LOG_PATH
                    } catch {
                    }
                    ; #endregion
                    return grp
                }
            }
        }
        catch {
        }
    }
    ; #region agent log
    try {
        FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' . A_TickCount .
        ',"location":"Shift keys.ahk:9912","message":"FindMonthGroup Strategy 2 failed","data":{"foundMonths":' .
        foundMonths.Length . '},"sessionId":"debug-session","runId":"run1","hypothesisId":"C"}`n', DEBUG_LOG_PATH
    } catch {
    }
    ; #endregion
    ; #region agent log
    try {
        ; Try to find what Groups exist
        try {
            allGroups := uia.FindAll({ Type: "Group" })
            groupCount := allGroups.Length
            FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' . A_TickCount .
            ',"location":"Shift keys.ahk:9918","message":"FindMonthGroup diagnostic","data":{"totalGroups":' .
            groupCount . '},"sessionId":"debug-session","runId":"run1","hypothesisId":"D"}`n', DEBUG_LOG_PATH
            ; Inspect first few Groups for className and name
            sampleCount := groupCount < 5 ? groupCount : 5
            loop sampleCount {
                try {
                    grp := allGroups[A_Index]
                    className := ""
                    name := ""
                    try className := grp.GetPropertyValue(UIA.Property.ClassName)
                    try name := grp.GetPropertyValue(UIA.Property.Name)
                    FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' . A_TickCount .
                    ',"location":"Shift keys.ahk:9920","message":"FindMonthGroup Group sample","data":{"index":' .
                    A_Index . ',"className":"' . className . '","name":"' . name .
                    '"},"sessionId":"debug-session","runId":"run1","hypothesisId":"D"}`n', DEBUG_LOG_PATH
                } catch {
                }
            }
            ; Try to find pagination elements
            try {
                paginationBtns := uia.FindAll({ Name: "Go to next page", Type: 50000 })
                if !paginationBtns.Length {
                    paginationBtns := uia.FindAll({ Name: "Go to previous page", Type: 50000 })
                }
                FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' . A_TickCount .
                ',"location":"Shift keys.ahk:9930","message":"FindMonthGroup pagination check","data":{"foundPagination":' .
                (paginationBtns.Length > 0 ? 1 : 0) . ',"count":' . paginationBtns.Length .
                '},"sessionId":"debug-session","runId":"run1","hypothesisId":"E"}`n', DEBUG_LOG_PATH
            } catch {
            }
            ; Try to find any text elements that might contain dates/months
            try {
                allTexts := uia.FindAll({ Type: "Text" })
                textCount := allTexts.Length
                FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' . A_TickCount .
                ',"location":"Shift keys.ahk:9935","message":"FindMonthGroup text elements","data":{"totalTexts":' .
                textCount . '},"sessionId":"debug-session","runId":"run1","hypothesisId":"E"}`n', DEBUG_LOG_PATH
                ; Sample first 10 text elements
                sampleTextCount := textCount < 10 ? textCount : 10
                loop sampleTextCount {
                    try {
                        txt := allTexts[A_Index]
                        txtName := ""
                        try txtName := txt.GetPropertyValue(UIA.Property.Name)
                        if txtName && StrLen(txtName) > 0 {
                            FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' .
                            A_TickCount .
                            ',"location":"Shift keys.ahk:9940","message":"FindMonthGroup text sample","data":{"index":' .
                            A_Index . ',"text":"' . txtName .
                            '"},"sessionId":"debug-session","runId":"run1","hypothesisId":"E"}`n', DEBUG_LOG_PATH
                        }
                    } catch {
                    }
                }
            } catch {
            }
        } catch Error as e2 {
            FileAppend '{"id":"log_' . A_TickCount . '_' . Random(1000, 9999) . '","timestamp":' . A_TickCount .
            ',"location":"Shift keys.ahk:9918","message":"FindMonthGroup diagnostic failed","data":{"error":"' . e2.Message .
            '"},"sessionId":"debug-session","runId":"run1","hypothesisId":"D"}`n', DEBUG_LOG_PATH
        }
    } catch {
    }
    ; #endregion
    return 0
}

;-------------------------------------------------------------------
; YouTube Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe chrome.exe") && InStr(WinGetTitle("A"), "YouTube")

; Shift + S : Focus search box
+s:: {
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

; Shift + I : Focus first video via Explore button
+i:: {
    try {
        uia := UIA_Browser()
        Sleep 300

        ; Find the "Explore" button as anchor
        exploreButton := uia.FindFirst({ Name: "Explore" })
        if !exploreButton
            exploreButton := uia.FindFirst({ Type: "Button", Name: "Explore" })
        if !exploreButton
            exploreButton := uia.FindFirst({ AutomationId: "explore" })

        if (exploreButton) {
            ; Focus the Explore button (do not click)
            exploreButton.SetFocus()
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

; Shift + H : Navigate to YouTube Home
+h:: {
    try {
        uia := UIA_Browser()
        uia.Navigate("https://www.youtube.com/")
    } catch Error as e {
        ; Fallback: use address bar navigation with clipboard paste
        clipSave := ClipboardAll()
        A_Clipboard := "https://www.youtube.com/"
        Send "^l"  ; Focus address bar
        Sleep 50
        Send "^v{Enter}"  ; Paste and navigate
        Sleep 50
        A_Clipboard := clipSave
    }
}

; Shift + R : Navigate to YouTube History
+r:: {
    try {
        uia := UIA_Browser()
        uia.Navigate("https://www.youtube.com/feed/history")
    } catch Error as e {
        ; Fallback: use address bar navigation with clipboard paste
        clipSave := ClipboardAll()
        A_Clipboard := "https://www.youtube.com/feed/history"
        Send "^l"  ; Focus address bar
        Sleep 50
        Send "^v{Enter}"  ; Paste and navigate
        Sleep 50
        A_Clipboard := clipSave
    }
}

; Shift + P : Navigate to YouTube Playlists
+p:: {
    try {
        uia := UIA_Browser()
        uia.Navigate("https://www.youtube.com/feed/playlists")
    } catch Error as e {
        ; Fallback: use address bar navigation with clipboard paste
        clipSave := ClipboardAll()
        A_Clipboard := "https://www.youtube.com/feed/playlists"
        Send "^l"  ; Focus address bar
        Sleep 50
        Send "^v{Enter}"  ; Paste and navigate
        Sleep 50
        A_Clipboard := clipSave
    }
}

#HotIf

;-------------------------------------------------------------------
; Gemini Website Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe chrome.exe") && InStr(WinGetTitle("A"), "gemini", false)

; Global state for Gemini drawer (main menu) – mirrors the state‑based toggle pattern
isGeminiDrawerOpen := false

; Global state for Gemini model toggle (Fast, Thinking, or Pro)
isGeminiFastModel := "Fast"  ; Tracks current model: "Fast", "Thinking", or "Pro"

; Global variables for Gemini model selector wizard menu
global g_GeminiModelSelectorGui := false
global g_GeminiModelSelectorActive := false
global g_GeminiModelHotkeyHandlers := []
global g_GeminiModelCharSequence := ["1", "2", "3"]
global g_GeminiModels := [{ name: "Fast", description: "Answers quickly" }, { name: "Thinking", description: "Solves complex problems" }, { name: "Pro",
    description: "Thinks longer for advanced math & code" }]

; Shift + D : Toggle the Main menu button (drawer) using fast state-based pattern
+d:: {
    ToggleGeminiDrawer()
}

; ---------------------------------------------------------------------------
ToggleGeminiDrawer() {
    global isGeminiDrawerOpen

    try {
        uia := UIA_Browser()
        if !IsObject(uia) {
            ; If we can't attach to Chrome, fall back to Escape like before
            SendEscape()
            return
        }

        ; Small settle time only – keep this snappy
        Sleep 100

        ; Exact-name regex (case-insensitive, anchored) with simple localization variant
        ; Same button toggles both open/close, we just keep state on our side
        mainMenuPattern := "i)^(Main menu|Menu principal)$"

        ; Use WaitForButton for robust, fast matching (similar to ToggleVoiceMessage)
        btn := WaitForButton(uia, mainMenuPattern, 1500)

        if (btn) {
            try {
                btn.Click()
            } catch Error as err {
                ; Try again in case of transient UIA glitch
                try {
                    btn.Click()
                } catch Error as err2 {
                    ; Swallow secondary failure – nothing else to do
                }
            }

            ; Flip our state after successful click
            isGeminiDrawerOpen := !isGeminiDrawerOpen

            ; Give Gemini a brief moment to redraw the drawer
            Sleep 200
        } else {
            ; If we couldn't find the button at all, keep behavior similar to previous version
            SendEscape()
        }
    } catch Error as e {
        ; If anything goes wrong, graceful fallback
        SendEscape()
    }
}

; Shift + N : New chat in Gemini (sends Ctrl-Shift-O)
+n:: {
    Send "^+o"
}

; Shift + S : Click the Search button - Search
+s:: {
    try {
        uia := UIA_Browser()
        Sleep 300

        ; Primary strategy: Find by Name "Search" with Type 50000 (Button)
        searchButton := uia.FindFirst({ Name: "Search", Type: 50000 })

        ; Fallback 1: Try by Type "Button" and Name "Search"
        if !searchButton {
            searchButton := uia.FindFirst({ Type: "Button", Name: "Search" })
        }

        ; Fallback 2: Try by ClassName containing "search-button" (substring match)
        if !searchButton {
            allButtons := uia.FindAll({ Type: 50000 })
            for button in allButtons {
                if InStr(button.ClassName, "search-button") && InStr(button.Name, "Search") {
                    searchButton := button
                    break
                }
            }
        }

        ; Fallback 3: Try finding by Name with substring match (in case of localization variations)
        if !searchButton {
            allButtons := uia.FindAll({ Type: 50000 })
            for button in allButtons {
                if InStr(button.Name, "Search") || InStr(button.Name, "Pesquisar") || InStr(button.Name, "Buscar") {
                    ; Additional check to ensure it's the search button (has search-button in className)
                    if InStr(button.ClassName, "search-button") {
                        searchButton := button
                        break
                    }
                }
            }
        }

        if (searchButton) {
            searchButton.Click()
        } else {
            ; Last resort: Could try keyboard navigation if Gemini has a keyboard shortcut for search
            ; For now, we'll just not do anything if we can't find the button
        }
    } catch Error as e {
        ; If all else fails, silently fail (no fallback action defined)
    }
}

; Shift + M : Show model selector wizard menu (Fast, Thinking, Pro) - Model
+m:: {
    ShowGeminiModelSelector()
}

; ---------------------------------------------------------------------------
ToggleGeminiModel() {
    global isGeminiFastModel

    ; Show a small banner while toggling models
    ShowSmallLoadingIndicator_ChatGPT("Switching model...")

    try {
        uia := UIA_Browser()
        if !IsObject(uia) {
            return
        }

        Sleep 100  ; Small settle time – keep this snappy

        ; Combined pattern to find Fast, Thinking, or Pro button in one search
        modelPattern := "i)^(Fast|Thinking|Pro)$"

        ; Helper to grab a button by pattern with shorter timeout for speed
        FindBtn(p) => WaitForButton(uia, p, 1500)
        ; Find all model buttons to see which ones are available
        allButtons := uia.FindAll({ Type: "Button" })
        modelButtons := []
        for btnCandidate in allButtons {
            try {
                btnName := btnCandidate.Name
                if (RegExMatch(btnName, modelPattern)) {
                    ; Get className first to check for disabled state
                    className := ""
                    try {
                        className := btnCandidate.ClassName
                    } catch {
                    }

                    ; Check if button is disabled
                    isDisabled := false
                    try {
                        ; Check if button has "disabled" in class name
                        if (InStr(className, "disabled") || InStr(className, "mat-mdc-button-disabled")) {
                            isDisabled := true
                        }
                        ; Also check IsEnabled property
                        try {
                            if (!btnCandidate.GetPropertyValue(UIA.Property.IsEnabled)) {
                                isDisabled := true
                            }
                        } catch {
                        }
                    } catch {
                    }

                    ; Check if button is selected/active
                    isSelected := false
                    try {
                        isSelected := btnCandidate.GetPropertyValue(UIA.Property.IsSelectionItemPatternAvailable) &&
                        btnCandidate.SelectionItemPattern.IsSelected
                    } catch {
                        ; Try alternative way to check if selected
                        try {
                            ; Check if button has "selected" in class name or other properties
                            if (InStr(className, "selected") || InStr(className, "active") || InStr(className,
                                "mdc-selected")) {
                                isSelected := true
                            }
                        } catch {
                        }
                    }
                    modelButtons.Push({ btn: btnCandidate, name: btnName, isSelected: isSelected, isDisabled: isDisabled,
                        className: className })
                }
            } catch {
            }
        }

        ; Strategy: Find a button that is NOT the current one (to actually toggle)
        ; If we have multiple buttons, click one that's different from current state
        btn := 0
        currentModelLower := StrLower(isGeminiFastModel)

        if (modelButtons.Length > 1) {
            ; Multiple buttons available - find one that's different from current AND enabled
            for modelBtn in modelButtons {
                btnNameLower := StrLower(modelBtn.name)
                if (btnNameLower != currentModelLower && !modelBtn.isDisabled) {
                    btn := modelBtn.btn
                    break
                }
            }

            ; If no different enabled button found, try to find any enabled button (even if same name)
            ; This handles cases where we want to click Pro but the disabled Pro button was found first
            if (!btn) {
                for modelBtn in modelButtons {
                    if (!modelBtn.isDisabled) {
                        btn := modelBtn.btn
                        break
                    }
                }
            }
        }

        ; If we couldn't find a different enabled button, or only one button available, use the first enabled one
        ; (This handles the case where only one model button is visible at a time)
        if (!btn && modelButtons.Length > 0) {
            ; Try to find first enabled button
            for modelBtn in modelButtons {
                if (!modelBtn.isDisabled) {
                    btn := modelBtn.btn
                    break
                }
            }

            ; If all buttons are disabled, still try the first one
            if (!btn) {
                btn := modelButtons[1].btn
            }
        }

        ; Final fallback: if no buttons found in our scan, use the original method
        if (!btn) {
            btn := FindBtn(modelPattern)
        }

        if (btn) {
            ; Check which button we found
            btnName := ""
            try btnName := btn.Name

            ; Determine if this button supports Invoke
            supportsInvoke := false
            try {
                supportsInvoke := btn.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)
            } catch {
                supportsInvoke := false
            }

            ; Try multi-strategy activation: prefer Invoke when available, fallback to Click, then mouse coordinates
            clicked := false

            ; Strategy 1: Try to focus the button first (may help with COM errors)
            try {
                btn.SetFocus()
                Sleep 50
            } catch {
                ; Ignore focus errors
            }

            if (supportsInvoke) {
                try {
                    btn.Invoke()
                    clicked := true
                } catch {
                }
            }
            if (!clicked) {
                try {
                    btn.Click()
                    clicked := true
                } catch {
                }
            }

            ; Strategy 3: Fallback to mouse coordinates if both Invoke and Click fail
            if (!clicked) {
                try {
                    ; Get browser window handle to ensure we activate the correct window
                    browserHwnd := 0
                    try {
                        browserHwnd := uia.BrowserId
                    } catch {
                        ; Fallback: try to find Chrome window with Gemini in title
                        browserHwnd := WinExist("ahk_exe chrome.exe")
                    }

                    ; Activate the browser window BEFORE clicking to prevent activating wrong window
                    if (browserHwnd) {
                        if (!WinExist("ahk_id " browserHwnd)) {
                            ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
                            return
                        }
                        WinActivate("ahk_id " browserHwnd)
                        WinWaitActive("ahk_id " browserHwnd, , 1)
                        Sleep 50  ; Brief pause after activation
                    }

                    ; Get button location and click using mouse coordinates
                    btnLocation := btn.Location
                    if (btnLocation && btnLocation.x >= 0 && btnLocation.y >= 0) {
                        ; Save current mouse position and coordinate mode
                        MouseGetPos(&prevX, &prevY)
                        prevCoordMode := A_CoordModeMouse

                        ; Set coordinate mode to Screen for absolute coordinates
                        CoordMode("Mouse", "Screen")

                        ; Calculate click position
                        clickX := btnLocation.x + btnLocation.w // 2
                        clickY := btnLocation.y + btnLocation.h // 2

                        ; Click at button center
                        Click(clickX, clickY)
                        clicked := true

                        ; Restore mouse position and coordinate mode after a brief delay
                        Sleep 100
                        CoordMode("Mouse", prevCoordMode)
                        MouseMove(prevX, prevY)
                    }
                } catch {
                }
            }

            ; Update state based on which button we actually clicked
            ; Use case-insensitive comparison since button names may vary in casing
            btnNameLower := StrLower(btnName)
            if (clicked) {
                if (btnNameLower = "fast") {
                    isGeminiFastModel := "Fast"
                    ShowSmallLoadingIndicator_ChatGPT("Fast model active")
                } else if (btnNameLower = "thinking") {
                    isGeminiFastModel := "Thinking"
                    ShowSmallLoadingIndicator_ChatGPT("Thinking model active")
                } else if (btnNameLower = "pro") {
                    isGeminiFastModel := "Pro"
                    ShowSmallLoadingIndicator_ChatGPT("Pro model active")
                }
                Sleep 150  ; Minimal sleep – just enough for UI to register
            }
        }
    } catch Error as err {
        ; Silently fail if anything goes wrong
    } finally {
        ; Hide the banner shortly after finishing
        SetTimer(() => HideSmallLoadingIndicator_ChatGPT(), -900)
    }
}

; ---------------------------------------------------------------------------
; Cleanup function for Gemini model selector
CleanupGeminiModelSelector() {
    global g_GeminiModelSelectorActive, g_GeminiModelSelectorGui, g_GeminiModelHotkeyHandlers

    ; Disable active flag
    g_GeminiModelSelectorActive := false

    ; Disable all character hotkeys
    for handler in g_GeminiModelHotkeyHandlers {
        try {
            char := handler.char
            Hotkey(char, "Off")
            ; Also disable uppercase for numbers (though they're already uppercase)
            if (RegExMatch(char, "^[1-9]$")) {
                Hotkey(char, "Off")
            }
        } catch {
            ; Silently ignore errors
        }
    }

    ; Disable Escape hotkey
    try {
        Hotkey("Escape", "Off")
    } catch {
        ; Ignore
    }

    ; Clear handlers array
    g_GeminiModelHotkeyHandlers := []

    ; Close and destroy GUI
    if (IsObject(g_GeminiModelSelectorGui)) {
        try {
            g_GeminiModelSelectorGui.Destroy()
        } catch {
            ; Ignore
        }
        g_GeminiModelSelectorGui := false
    }
}

; ---------------------------------------------------------------------------
; Handler for Escape key in Gemini model selector
HandleGeminiModelSelectorEscape(*) {
    global g_GeminiModelSelectorActive
    if (g_GeminiModelSelectorActive) {
        CleanupGeminiModelSelector()
    }
}

; ---------------------------------------------------------------------------
; Factory function to create a handler that properly captures the model name
CreateGeminiModelCharHandler(char) {
    return (*) => HandleGeminiModelSelection(char)
}

; ---------------------------------------------------------------------------
; Handler for model selection in Gemini model selector
HandleGeminiModelSelection(char) {
    global g_GeminiModelSelectorActive, g_GeminiModels, g_GeminiModelCharSequence
    global isGeminiFastModel

    ; Only process if selector is active
    if (!g_GeminiModelSelectorActive) {
        return
    }

    ; Get model index from character (1-based array index)
    modelIndex := -1
    for idx, ch in g_GeminiModelCharSequence {
        if (ch = char) {
            modelIndex := idx
            break
        }
    }

    if (modelIndex < 1 || modelIndex > g_GeminiModels.Length) {
        return
    }

    ; Get model info
    modelInfo := g_GeminiModels[modelIndex]
    if (!IsObject(modelInfo)) {
        return
    }
    try {
        modelName := modelInfo.name
    } catch {
        return
    }
    if (modelName = "") {
        return
    }

    ; Cleanup selector first (closes GUI, disables hotkeys)
    try {
        CleanupGeminiModelSelector()
    } catch {
        ; Ignore cleanup errors
    }

    ; Small delay to ensure GUI cleanup is complete
    Sleep 100

    try {
        ; Activate Gemini window
        SetTitleMatchMode(2)
        geminiHwnd := 0
        try {
            ; Find Gemini window by searching Chrome windows with "gemini" in title
            for hwnd in WinGetList("ahk_exe chrome.exe") {
                try {
                    if InStr(WinGetTitle("ahk_id " hwnd), "gemini", false) {
                        geminiHwnd := hwnd
                        break
                    }
                } catch {
                    ; Silently skip invalid windows
                }
            }
        } catch {
            ; Silently handle WinGetList errors
        }

        if (geminiHwnd) {
            if (!WinExist("ahk_id " geminiHwnd)) {
                ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
                return
            }
            WinActivate("ahk_id " geminiHwnd)
            if !WinWaitActive("ahk_id " geminiHwnd, , 2) {
                return
            }
        } else {
            ; Fallback: try to activate any Chrome window
            if (!WinExist("ahk_exe chrome.exe")) {
                ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
                return
            }
            WinActivate("ahk_exe chrome.exe")
            if !WinWaitActive("ahk_exe chrome.exe", , 2) {
                return
            }
        }

        ; Small settle time for window activation
        Sleep 50

        ; Focus the Gemini prompt field so we can edit the current prompt in-place
        if (FocusGeminiPromptField()) {
            ; 1) Move caret to beginning, 2) insert @model name, 3) move back to end of prompt, 4) add a trailing space
            Send "^{Home}"
            Sleep 120
            modelCommand := "@" . modelName . " "
            SendText modelCommand
            Sleep 120
            Send "^{End}"
            Sleep 120
            Send " "
        }

        ; Update global state (tracks which model is conceptually active)
        isGeminiFastModel := modelName
    } catch Error as err {
        ; Silently fail if anything goes wrong
    }
}

; ---------------------------------------------------------------------------
; Show Gemini model selector wizard menu
ShowGeminiModelSelector() {
    global g_GeminiModelSelectorGui, g_GeminiModelSelectorActive, g_GeminiModelHotkeyHandlers
    global g_GeminiModels, g_GeminiModelCharSequence

    ; Verify global variables are initialized
    if (!IsObject(g_GeminiModels) || g_GeminiModels.Length = 0) {
        return
    }

    if (!IsObject(g_GeminiModelCharSequence) || g_GeminiModelCharSequence.Length = 0) {
        return
    }

    ; Close existing GUI if open
    if (g_GeminiModelSelectorActive && IsObject(g_GeminiModelSelectorGui)) {
        CleanupGeminiModelSelector()
        Sleep 50
    }

    ; Get monitor dimensions
    activeWin := 0
    try {
        activeWin := WinGetID("A")
    } catch {
        activeWin := 0
    }

    ; Default to primary monitor work area
    MonitorGetWorkArea(1, &monitorLeft, &monitorTop, &monitorRight, &monitorBottom)
    monitorWidth := monitorRight - monitorLeft
    monitorHeight := monitorBottom - monitorTop

    ; If we have an active window, find which monitor contains its center
    if (activeWin && activeWin != 0) {
        rect := Buffer(16, 0)
        if (DllCall("GetWindowRect", "ptr", activeWin, "ptr", rect)) {
            winLeft := NumGet(rect, 0, "int")
            winTop := NumGet(rect, 4, "int")
            winRight := NumGet(rect, 8, "int")
            winBottom := NumGet(rect, 12, "int")

            centerX := winLeft + (winRight - winLeft) // 2
            centerY := winTop + (winBottom - winTop) // 2

            monitorCount := MonitorGetCount()
            loop monitorCount {
                idx := A_Index
                MonitorGetWorkArea(idx, &l, &t, &r, &b)
                if (centerX >= l && centerX <= r && centerY >= t && centerY <= b) {
                    monitorLeft := l
                    monitorTop := t
                    monitorRight := r
                    monitorBottom := b
                    monitorWidth := r - l
                    monitorHeight := b - t
                    break
                }
            }
        }
    }

    ; Create GUI
    g_GeminiModelSelectorGui := Gui("+AlwaysOnTop +ToolWindow +E0x08000000", "Gemini Model Selector")
    fontSize := (monitorHeight < 800) ? 9 : 10
    g_GeminiModelSelectorGui.SetFont("s" . fontSize, "Segoe UI")
    g_GeminiModelSelectorGui.MarginX := 10
    g_GeminiModelSelectorGui.MarginY := 5

    ; Build display text
    displayText := "Gemini Model Selector`n`n"
    charIndex := 0
    for model in g_GeminiModels {
        if (charIndex < g_GeminiModelCharSequence.Length) {
            char := g_GeminiModelCharSequence[charIndex + 1]
            displayText .= "[" . char . "] " . model.name . " - " . model.description . "`n"
            charIndex++
        }
    }

    ; Calculate GUI size
    baseWidth := 400
    textControlHeight := Min(400, (g_GeminiModels.Length * 25) + 60)
    textControlWidth := baseWidth - 20

    ; Add text control with display
    g_GeminiModelSelectorGui.AddEdit("w" . textControlWidth . " h" . textControlHeight . " ReadOnly VScroll",
        displayText)

    ; Add Close button
    closeBtn := g_GeminiModelSelectorGui.AddButton("w100 Default Center", "Close")
    closeBtn.OnEvent("Click", (*) => CleanupGeminiModelSelector())

    ; Calculate total height
    totalHeight := 10 + textControlHeight + 40 + 10
    guiWidth := baseWidth

    ; Calculate center position
    marginX := 20
    marginY := 20
    guiX := monitorLeft + (monitorWidth - guiWidth) // 2
    guiY := monitorTop + (monitorHeight - totalHeight) // 2

    ; Ensure GUI stays within monitor bounds
    if (guiX < monitorLeft + marginX)
        guiX := monitorLeft + marginX
    if (guiY < monitorTop + marginY)
        guiY := monitorTop + marginY
    if (guiX + guiWidth > monitorLeft + monitorWidth - marginX)
        guiX := monitorLeft + monitorWidth - guiWidth - marginX
    if (guiY + totalHeight > monitorTop + monitorHeight - marginY)
        guiY := monitorTop + monitorHeight - totalHeight - marginY

    ; Show GUI centered on the active window's monitor
    try {
        g_GeminiModelSelectorGui.Show("NA w" . guiWidth . " h" . totalHeight . " x" . guiX . " y" . guiY)

        ; Small delay to ensure GUI is actually visible
        Sleep 50
    } catch Error as e {
        return
    }

    ; Set active flag
    g_GeminiModelSelectorActive := true

    ; Clear handlers array
    g_GeminiModelHotkeyHandlers := []

    ; Enable hotkeys for all assigned characters
    charIndex := 0
    for model in g_GeminiModels {
        if (charIndex < g_GeminiModelCharSequence.Length) {
            char := g_GeminiModelCharSequence[charIndex + 1]

            ; Create handler
            handler := CreateGeminiModelCharHandler(char)

            ; Store handler for cleanup
            g_GeminiModelHotkeyHandlers.Push({ char: char, handler: handler })

            ; Enable hotkey
            try {
                Hotkey(char, handler, "On")
            } catch {
                ; Silently ignore if we can't create hotkey
            }

            charIndex++
        }
    }

    ; Enable Escape hotkey
    Hotkey("Escape", HandleGeminiModelSelectorEscape, "On")
}

; Shift + T : Click the Tools button - Tools
+t:: {
    try {
        uia := UIA_Browser()
        Sleep 300

        ; Primary strategy: Find by Name "Tools" with Type 50000 (Button)
        toolsButton := uia.FindFirst({ Name: "Tools", Type: 50000 })

        ; Fallback 1: Try by Type "Button" and Name "Tools"
        if !toolsButton {
            toolsButton := uia.FindFirst({ Type: "Button", Name: "Tools" })
        }

        ; Fallback 2: Try by ClassName containing "toolbox-drawer-button" (substring match)
        if !toolsButton {
            allButtons := uia.FindAll({ Type: 50000 })
            for button in allButtons {
                if InStr(button.ClassName, "toolbox-drawer-button") && InStr(button.Name, "Tools") {
                    toolsButton := button
                    break
                }
            }
        }

        ; Fallback 3: Try finding by Name with substring match (in case of localization variations)
        if !toolsButton {
            allButtons := uia.FindAll({ Type: 50000 })
            for button in allButtons {
                if InStr(button.Name, "Tools") || InStr(button.Name, "Ferramentas") {
                    ; Additional check to ensure it's the tools button (has toolbox-drawer-button in className)
                    if InStr(button.ClassName, "toolbox-drawer-button") {
                        toolsButton := button
                        break
                    }
                }
            }
        }

        if (toolsButton) {
            toolsButton.Click()

            Sleep 100

            Send "{Tab}"
        } else {
            ; Last resort: Could not find Tools button
        }
    } catch Error as e {
        ; If all else fails, silently fail (no fallback action defined)
    }
}

; Helper: Focus the Gemini prompt text field using UIA
FocusGeminiPromptField() {
    try {
        uia := UIA_Browser()
        Sleep 150  ; small settle per README (keep this snappy)

        ; Primary strategy: Find by Name (Gemini updated placeholder in 2025)
        try
            promptField := uia.FindFirst({ Name: "Enter a prompt for Gemini", Type: 50004 })
        catch
            promptField := ""

        ; Fallback 1: Legacy name "Enter a prompt here"
        if !promptField {
            try
                promptField := uia.FindFirst({ Name: "Enter a prompt here", Type: 50004 })
            catch
                promptField := ""
        }

        ; Shared scan for remaining fallbacks (single FindAll + scoring for efficiency)
        if !promptField {
            best := 0, bestScore := -1
            allEdits := uia.FindAll({ Type: 50004 })
            for edit in allEdits {
                cls := edit.ClassName
                name := edit.Name
                score := 0
                if InStr(cls, "ql-editor")
                    score += 3
                if InStr(cls, "new-input-ui")
                    score += 2
                if InStr(name, "Enter a prompt")
                    score += 3
                else if InStr(name, "prompt")
                    score += 2
                else if InStr(name, "Digite um prompt")
                    score += 2
                ; pick the highest scoring candidate
                if (score > bestScore) {
                    bestScore := score
                    best := edit
                }
            }
            if (bestScore >= 0) {
                promptField := best
            }
        }

        if (promptField) {
            promptField.SetFocus()
            Sleep 100
            ; Ensure focus was successful
            if (!promptField.HasKeyboardFocus) {
                ; Fallback: try clicking if SetFocus didn't work
                promptField.Click()
                Sleep 100
            }
            return true
        }
    } catch Error as e {
        ; If all else fails, silently fail (no fallback action defined)
    }
    return false
}

; Shift + P : Focus the prompt text field - Prompt
+p:: {
    FocusGeminiPromptField()
}

; Shift + C : Click the last Copy button (copies the preceding message) - Copy
+c:: {
    try {
        uia := UIA_Browser()
        Sleep 300

        ; Find all Copy buttons
        allCopyButtons := []

        ; Primary strategy: Find all buttons with Name "Copy"
        allButtons := uia.FindAll({ Type: 50000 })
        for button in allButtons {
            if (button.Name = "Copy" || InStr(button.Name, "Copy", false) = 1) {
                ; Additional check: ensure it has the Copy button className pattern
                if (InStr(button.ClassName, "icon-button") || InStr(button.ClassName, "mdc-button")) {
                    allCopyButtons.Push(button)
                }
            }
        }

        ; Fallback: Try by Type "Button" if the above didn't find enough
        if (allCopyButtons.Length = 0) {
            allButtons := uia.FindAll({ Type: "Button" })
            for button in allButtons {
                if (button.Name = "Copy" || InStr(button.Name, "Copy", false) = 1) {
                    allCopyButtons.Push(button)
                }
            }
        }

        if (allCopyButtons.Length = 0) {
            ; No Copy buttons found
            return
        }

        ; Find the last Copy button (the one with the highest Y position, meaning furthest down the page)
        lastCopyButton := 0
        highestY := -1

        for copyButton in allCopyButtons {
            try {
                btnPos := copyButton.Location
                btnBottomY := btnPos.y + btnPos.h

                ; The last button will be the one with the highest bottom Y coordinate
                if (btnBottomY > highestY) {
                    highestY := btnBottomY
                    lastCopyButton := copyButton
                }
            } catch {
                ; If getting location fails, skip this button
            }
        }

        ; If position-based approach didn't work, just use the last one in the array
        if (!lastCopyButton && allCopyButtons.Length > 0) {
            lastCopyButton := allCopyButtons[allCopyButtons.Length]
        }

        if (lastCopyButton) {
            lastCopyButton.Click()
        } else {
            ; Last resort: Could not find last Copy button
        }
    } catch Error as e {
        ; If all else fails, silently fail (no fallback action defined)
    }
}

; Shift + R : Read aloud the last message (click last "Show more options" then "Text to speech") - Read
+r:: {
    try {
        uia := UIA_Browser()
        Sleep 300

        ; Step 1: Find all "Show more options" buttons
        allMoreOptionsButtons := []

        ; Primary strategy: Find all buttons with Name "Show more options"
        allButtons := uia.FindAll({ Type: 50000 })
        for button in allButtons {
            if (button.Name = "Show more options" || InStr(button.Name, "Show more options", false) = 1) {
                ; Additional check: ensure it has the more-menu-button className pattern
                if (InStr(button.ClassName, "more-menu-button") || InStr(button.ClassName, "mdc-button")) {
                    allMoreOptionsButtons.Push(button)
                }
            }
        }

        ; Fallback: Try by Type "Button" if the above didn't find enough
        if (allMoreOptionsButtons.Length = 0) {
            allButtons := uia.FindAll({ Type: "Button" })
            for button in allButtons {
                if (button.Name = "Show more options" || InStr(button.Name, "Show more options", false) = 1) {
                    if (InStr(button.ClassName, "more-menu-button")) {
                        allMoreOptionsButtons.Push(button)
                    }
                }
            }
        }

        if (allMoreOptionsButtons.Length = 0) {
            ; No "Show more options" buttons found
            return
        }

        ; Find the last "Show more options" button (the one with the highest Y position, meaning furthest down the page)
        lastMoreOptionsButton := 0
        highestY := -1

        for moreOptionsButton in allMoreOptionsButtons {
            try {
                btnPos := moreOptionsButton.Location
                btnBottomY := btnPos.y + btnPos.h

                ; The last button will be the one with the highest bottom Y coordinate
                if (btnBottomY > highestY) {
                    highestY := btnBottomY
                    lastMoreOptionsButton := moreOptionsButton
                }
            } catch {
                ; If getting location fails, skip this button
            }
        }

        ; If position-based approach didn't work, just use the last one in the array
        if (!lastMoreOptionsButton && allMoreOptionsButtons.Length > 0) {
            lastMoreOptionsButton := allMoreOptionsButtons[allMoreOptionsButtons.Length]
        }

        if (!lastMoreOptionsButton) {
            ; Could not find last "Show more options" button
            return
        }

        ; Step 2: Click the last "Show more options" button
        lastMoreOptionsButton.Click()
        Sleep 400 ; Wait for menu to appear

        ; Step 3: Find and click the "Text to speech" menu item
        textToSpeechMenuItem := 0

        ; Primary strategy: Find by Name "Text to speech" with Type 50011 (MenuItem)
        textToSpeechMenuItem := uia.FindFirst({ Name: "Text to speech", Type: 50011 })

        ; Fallback 1: Try by Type "MenuItem" and Name "Text to speech"
        if !textToSpeechMenuItem {
            textToSpeechMenuItem := uia.FindFirst({ Type: "MenuItem", Name: "Text to speech" })
        }

        ; Fallback 2: Try by ClassName containing "mat-mdc-menu-item" (substring match)
        if !textToSpeechMenuItem {
            allMenuItems := uia.FindAll({ Type: 50011 })
            for menuItem in allMenuItems {
                if InStr(menuItem.Name, "Text to speech") || InStr(menuItem.Name, "speech") {
                    if InStr(menuItem.ClassName, "mat-mdc-menu-item") {
                        textToSpeechMenuItem := menuItem
                        break
                    }
                }
            }
        }

        ; Fallback 3: Try finding by Name with substring match (in case of localization variations)
        if !textToSpeechMenuItem {
            allMenuItems := uia.FindAll({ Type: 50011 })
            for menuItem in allMenuItems {
                if InStr(menuItem.Name, "Text to speech") || InStr(menuItem.Name, "Texto para fala") || InStr(
                    menuItem.Name,
                    "Ler em voz alta") {
                    if InStr(menuItem.ClassName, "mat-mdc-menu-item") {
                        textToSpeechMenuItem := menuItem
                        break
                    }
                }
            }
        }

        if (textToSpeechMenuItem) {
            textToSpeechMenuItem.Click()
        } else {
            ; Last resort: Could not find "Text to speech" menu item
        }
    } catch Error as e {
        ; If all else fails, silently fail (no fallback action defined)
    }
}

; Shift + G : Focus the prompt text field and send Gemini prompt text - Gemini
+g:: {
    try {
        uia := UIA_Browser()
        Sleep 300

        ; Primary strategy: Find by Name (Gemini updated placeholder in 2025)
        try
            promptField := uia.FindFirst({ Name: "Enter a prompt for Gemini", Type: 50004 })
        catch
            promptField := ""

        ; Fallback 1: Legacy name "Enter a prompt here"
        if !promptField {
            try
                promptField := uia.FindFirst({ Name: "Enter a prompt here", Type: 50004 })
            catch
                promptField := ""
        }

        ; Fallback 2: Try by ClassName containing "ql-editor" or "new-input-ui" (substring match)
        if !promptField {
            allEdits := uia.FindAll({ Type: 50004 })
            for edit in allEdits {
                if (InStr(edit.ClassName, "ql-editor") || InStr(edit.ClassName, "new-input-ui")) {
                    if InStr(edit.Name, "Enter a prompt") || InStr(edit.Name, "prompt") {
                        promptField := edit
                        break
                    }
                }
            }
        }

        ; Fallback 3: Try finding by ClassName containing "ql-editor" (most specific identifier)
        if !promptField {
            allEdits := uia.FindAll({ Type: 50004 })
            for edit in allEdits {
                if InStr(edit.ClassName, "ql-editor") {
                    promptField := edit
                    break
                }
            }
        }

        ; Fallback 4: Try finding by Name with substring match (in case of localization variations)
        if !promptField {
            allEdits := uia.FindAll({ Type: 50004 })
            for edit in allEdits {
                if InStr(edit.Name, "Enter a prompt") || InStr(edit.Name, "Digite um prompt") || InStr(edit.Name,
                    "prompt") {
                    ; Additional check to ensure it's the prompt field (has ql-editor in className)
                    if InStr(edit.ClassName, "ql-editor") {
                        promptField := edit
                        break
                    }
                }
            }
        }

        if (promptField) {
            promptField.SetFocus()
            Sleep 100
            ; Ensure focus was successful
            if (!promptField.HasKeyboardFocus) {
                ; Fallback: try clicking if SetFocus didn't work
                promptField.Click()
                Sleep 100
            }

            ; Read the Gemini_Prompt.txt file and paste its contents via clipboard
            promptFilePath := A_ScriptDir "\data\Gemini_Prompt.txt"
            if FileExist(promptFilePath) {
                ; Save current clipboard
                oldClipboard := A_Clipboard
                try {
                    ; Read and set clipboard
                    promptText := FileRead(promptFilePath, "UTF-8")
                    if (promptText) {
                        A_Clipboard := promptText
                        ClipWait 1, 1  ; Wait for clipboard to be ready

                        ; Clear any existing text first (select all and delete)
                        Send "^a"
                        Sleep 50

                        ; Paste the text from clipboard
                        Send "^v"
                        Sleep 100

                        ; Restore original clipboard
                        A_Clipboard := oldClipboard

                        Sleep 400
                        Send "{Enter}"
                    }
                } catch Error as e {
                    ; If file reading fails, try to restore clipboard
                    try {
                        A_Clipboard := oldClipboard
                    }
                }
            } else {
                ; File not found - could show a message or just silently fail
            }
        } else {
            ; Last resort: Could not find prompt field
        }
    } catch Error as e {
        ; If all else fails, silently fail (no fallback action defined)
    }
}

; Shift + F : Click the Expand input to Fullscreen button
+f:: {
    try {
        uia := UIA_Browser()
        Sleep 300

        ; Primary strategy: Find by Name "Expand input to Fullscreen" with Type 50000 (Button)
        fullscreenButton := uia.FindFirst({ Name: "Expand input to Fullscreen", Type: 50000 })

        ; Fallback 1: Try by Type "Button" and Name "Expand input to Fullscreen"
        if !fullscreenButton {
            fullscreenButton := uia.FindFirst({ Type: "Button", Name: "Expand input to Fullscreen" })
        }

        ; Fallback 2: Try by ClassName containing "fullscreen-button" (substring match)
        if !fullscreenButton {
            allButtons := uia.FindAll({ Type: 50000 })
            for button in allButtons {
                if InStr(button.ClassName, "fullscreen-button") {
                    fullscreenButton := button
                    break
                }
            }
        }

        ; Fallback 3: Try finding by Name with substring match (in case of localization variations)
        if !fullscreenButton {
            allButtons := uia.FindAll({ Type: 50000 })
            for button in allButtons {
                if InStr(button.Name, "Expand input") || InStr(button.Name, "Fullscreen") || InStr(button.Name,
                    "Expandir") {
                    ; Additional check to ensure it's the fullscreen button (has fullscreen-button in className)
                    if InStr(button.ClassName, "fullscreen-button") {
                        fullscreenButton := button
                        break
                    }
                }
            }
        }

        if (fullscreenButton) {
            fullscreenButton.Click()
        } else {
            ; Last resort: Could not find fullscreen button
        }
    } catch Error as e {
        ; If all else fails, silently fail (no fallback action defined)
    }
}

; Enter : Send Enter and automatically monitor for response completion
; This automates the notification process so Ctrl+Enter is no longer needed
; Shift+Enter continues to work normally (for line breaks) as it's not intercepted
Enter:: {
    ; Only trigger if Shift or Ctrl are NOT pressed
    ; Shift+Enter = line break (should not trigger monitoring)
    ; Ctrl+Enter = handled by the ^Enter hotkey below
    if (GetKeyState("Shift", "P") || GetKeyState("Ctrl", "P")) {
        ; Pass through to let browser handle it normally
        Send "{Enter}"
        return
    }

    ; Send Enter key to submit the prompt
    Send "{Enter}"

    ; Phase 3: non-blocking daemon watch or legacy blocking monitor
    if (USE_DAEMON_MONITOR_GEMINI) {
        ShiftKeysIPC_StartGeminiWatch(300000, PlayCompletionChime_Gemini)
        return
    }
    WaitForStopResponseButton_Gemini()
}

; Control + Enter : Send Enter and monitor for response completion
^Enter:: {
    ; Send Enter key to submit the prompt
    Send "{Enter}"

    ; Phase 3: non-blocking daemon watch or legacy blocking monitor
    if (USE_DAEMON_MONITOR_GEMINI) {
        ShiftKeysIPC_StartGeminiWatch(300000, PlayCompletionChime_Gemini)
        return
    }
    WaitForStopResponseButton_Gemini()
}

; ---------------------------------------------------------------------------
; Monitor "Stop response" button and play chime when it disappears
; ---------------------------------------------------------------------------
WaitForStopResponseButton_Gemini(timeout := 300000) {
    ; Store Gemini window handle
    geminiHwnd := WinExist("A")
    if !geminiHwnd {
        return ; Gemini window not found
    }

    ; Obtain UIA context for Gemini window
    try {
        uia := UIA_Browser("ahk_id " geminiHwnd)
    } catch {
        return ; Failed to get UIA context
    }

    start := A_TickCount
    btn := ""
    buttonFound := false

    ; Wait for the "Stop response" button to appear
    deadline := (timeout > 0) ? (start + timeout) : 0
    while (timeout <= 0 || (A_TickCount < deadline)) {
        btn := ""

        ; Try to find the "Stop response" button
        try {
            btn := uia.FindFirst({ Type: "50000", Name: "Stop response" })
        } catch {
            btn := ""
        }

        if !btn {
            ; Fallback: Try by Type "Button" and Name "Stop response"
            try {
                btn := uia.FindFirst({ Type: "Button", Name: "Stop response" })
            } catch {
                btn := ""
            }
        }

        if !btn {
            ; Fallback: Try substring match for localization variations
            try {
                btn := uia.FindFirst({ Name: "Stop response", matchmode: "Substring" })
            } catch {
                btn := ""
            }
        }

        if btn {
            buttonFound := true
            ; Monitor the button until it disappears (with confirmation layer)
            while (timeout <= 0 || (A_TickCount < deadline)) {
                ; Monitor the button while it exists
                while btn && (timeout <= 0 || (A_TickCount < deadline)) {
                    Sleep 250
                    btn := ""

                    ; Check if button still exists
                    try {
                        btn := uia.FindFirst({ Type: "50000", Name: "Stop response" })
                    } catch {
                        btn := ""
                    }

                    if !btn {
                        try {
                            btn := uia.FindFirst({ Type: "Button", Name: "Stop response" })
                        } catch {
                            btn := ""
                        }
                    }

                    if !btn {
                        try {
                            btn := uia.FindFirst({ Name: "Stop response", matchmode: "Substring" })
                        } catch {
                            btn := ""
                        }
                    }
                }

                ; Button has disappeared - add confirmation layer
                ; Wait 1.5 seconds and check if it reappears (to avoid false positives)
                confirmationStart := A_TickCount
                confirmationPeriod := 1500  ; 1.5 seconds
                buttonReappeared := false

                ; Check multiple times during the confirmation period
                while ((A_TickCount - confirmationStart) < confirmationPeriod) && (timeout <= 0 || (A_TickCount <
                    deadline)) {
                    Sleep 300

                    ; Check if button reappeared
                    try {
                        btn := uia.FindFirst({ Type: "50000", Name: "Stop response" })
                    } catch {
                        btn := ""
                    }

                    if !btn {
                        try {
                            btn := uia.FindFirst({ Type: "Button", Name: "Stop response" })
                        } catch {
                            btn := ""
                        }
                    }

                    if !btn {
                        try {
                            btn := uia.FindFirst({ Name: "Stop response", matchmode: "Substring" })
                        } catch {
                            btn := ""
                        }
                    }

                    if btn {
                        ; Button reappeared - break out of confirmation loop and continue monitoring
                        buttonReappeared := true
                        break  ; Exit confirmation loop, will continue outer monitoring loop
                    }
                }

                ; If button didn't reappear during confirmation period, response is truly complete
                if !buttonReappeared {
                    break  ; Exit the outer monitoring loop
                }
                ; Otherwise, continue the outer loop to monitor the reappeared button
            }
            break
        }
        Sleep 250
    }

    ; Play chime when button disappears (only if we found it initially)
    if buttonFound {
        try {
            PlayCompletionChime_Gemini()
        } catch {
        }
    }
}

; ---------------------------------------------------------------------------
; Play completion chime for Gemini responses (debounced)
; ---------------------------------------------------------------------------
PlayCompletionChime_Gemini() {
    try {
        static lastTick := 0
        if (A_TickCount - lastTick < 1500)
            return
        lastTick := A_TickCount

        if (IsSoundEnabled()) {
            SoundPlay(A_ScriptDir . "\sounds\gemini-completion.wav")
        }
    } catch {
    }
}

#HotIf

;-------------------------------------------------------------------
; Google Search Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe chrome.exe") && InStr(WinGetTitle("A"), "Google")

; Shift + S : Focus Google search box
+s:: {
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

; Shift + U : Select first search result
+u:: {
    try {
        uia := UIA_Browser()
        if !uia
            return

        ; Find the main results container
        centerCol := uia.FindFirst({ AutomationId: "center_col" })

        targetLink := ""

        if (centerCol) {
            ; Find the first result title text inside center_col
            ; ClassName "LC20lb" is standard for Google result titles
            titleText := centerCol.FindFirst({ ClassName: "LC20lb", MatchMode: "Substring" })

            if (titleText) {
                ; The link is the parent of the title text
                targetLink := titleText.WalkTree("p") ; Get Parent
            }
        } else {
            ; Fallback: search from root if center_col not found
            titleText := uia.FindFirst({ ClassName: "LC20lb", MatchMode: "Substring" })
            if (titleText) {
                targetLink := titleText.WalkTree("p")
            }
        }

        if (targetLink) {
            ; Try to invoke (click) the link
            try {
                targetLink.Invoke()
            } catch {
                targetLink.Click()
            }
        } else {
            ToolTip("First result not found")
            SetTimer(() => ToolTip(), -2000)
        }

    } catch Error as e {
        ToolTip("Error: " . e.Message)
        SetTimer(() => ToolTip(), -2000)
    }
}

#HotIf

;-------------------------------------------------------------------
; File Dialog (Namespace Tree Control) Shortcuts
;-------------------------------------------------------------------
#HotIf IsFileDialogActive()

; Shift + F : Select first file - File
+f:: {
    if !IsFileDialogActive()
        return
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

; Shift + S : Focus search bar - Search bar
+s:: {
    if !IsFileDialogActive()
        return
    Send "^e"
}

; Shift + A : Focus address bar - Address bar
+a:: {
    if !IsFileDialogActive()
        return
    try {
        root := UIA.ElementFromHandle(WinExist("A"))
        ; Try to find address bar by common names
        addressBar := root.FindFirst({ Type: "Edit", Name: "Address:" })
        if !addressBar
            addressBar := root.FindFirst({ Type: "Edit", Name: "Endereço:" })
        if !addressBar
            addressBar := root.FindFirst({ Type: "ComboBox", AutomationId: "1001" })
        if !addressBar
            addressBar := root.FindFirst({ Type: "Edit", ClassName: "Edit" })

        if (addressBar) {
            addressBar.SetFocus()
            Sleep 50
            Send "^a"  ; Select all existing text
            return
        }
    } catch Error {
    }
    ; Fallback: Use Alt+D (common shortcut for address bar in file dialogs)
    Send "!d"
}

; Shift + N : New folder - New Folder
+n:: {
    if !IsFileDialogActive()
        return
    Send "^+n"
}

; Shift + P : Select first pinned item in sidebar - Pinned item
+p:: {
    if !IsFileDialogActive()
        return
    SelectExplorerSidebarFirstPinned()
}

; Shift + T : Select "This PC" / "Este computador" in sidebar - This PC
+t:: {
    if !IsFileDialogActive()
        return
    SelectExplorerSidebarFirstPinned()
    Sleep 200
    Send "{End}"
    Send "{Up}"
    Send "{Up}"
    Send "{Up}"
    Send "{Up}"
}

; Shift + M : Focus file name edit field - Name
+m:: {
    if !IsFileDialogActive()
        return
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

; Shift + O : Click Insert/Open/Save button - Open/Save
+o:: {
    if !IsFileDialogActive()
        return
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

; Shift + C : Click Cancel button - Cancel
+c:: {
    if !IsFileDialogActive()
        return
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
    SendEscape()  ; Escape key is universal for cancel
}

#HotIf

IsFileDialogActive() {
    hwnd := WinActive("A")
    if !hwnd {
        return false
    }

    winClass := WinGetClass("ahk_id " hwnd)
    winTitle := WinGetTitle("ahk_id " hwnd)
    winExe := WinGetProcessName("ahk_id " hwnd)
    if winClass != "#32770" {
        return false
    }

    ; Exclude Outlook Reminders which can share dialog-like traits
    try {
        if (winExe = "OUTLOOK.EXE") {
            if RegExMatch(winTitle, "i)Reminder") {
                return false
            }
        }
    } catch Error {
    }

    try {
        root := UIA.ElementFromHandle(hwnd)

        ; Check UIA properties: Type = Window (50032) and LocalizedType = "dialog"
        rootType := ""
        rootLocalizedType := ""
        rootName := ""
        try rootType := root.Type
        try rootLocalizedType := root.LocalizedType
        try rootName := root.Name

        ; Primary check: Type must be Window and LocalizedType must be "dialog"
        if (rootType = UIA.Type.Window && rootLocalizedType = "dialog") {
            return true
        }

        ; Fallback 1: Check if window name matches common file dialog names
        ; This handles cases where LocalizedType might vary by language/application
        if (rootType = UIA.Type.Window) {
            fileDialogNames := ["Save As", "Open", "Browse", "Select File", "Choose File",
                "Salvar como", "Abrir", "Procurar", "Selecionar arquivo",
                "Guardar como", "Guardar", "Explorar"]
            for dialogName in fileDialogNames {
                if (InStr(rootName, dialogName, false)) {
                    return true
                }
            }
        }

        ; Fallback 2: Check for file dialog characteristic elements (Namespace Tree Control, file lists, etc.)
        ; This is a last resort if UIA properties don't match but it's still a file dialog
        try {
            ; Check for Namespace Tree Control in window text (original method as backup)
            txt := WinGetText("ahk_id " hwnd)
            if InStr(txt, "Namespace Tree Control") || InStr(txt, "Controle da Ãrvore de Namespace") {
                return true
            }
        } catch {
            ; Window text check failed, continue
        }

        ; If we get here, it's not a file dialog
        return false
    } catch Error as e {
        return false
    }
}

;-------------------------------------------------------------------
; UIA Tree Inspector Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("UIATreeInspector") || WinActive("ahk_exe UIATreeInspectorAutoHotkey64.exe")

; Shift + R : Refresh list
+r:: {
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

; Shift + F : Focus filter field
+f:: {
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

; Shift + S : Select tree item by name prefix
+s:: {
    try {
        ; Global variable to store user input
        global g_TreeItemSearchInput := ""

        ; Create GUI dialog
        searchGui := Gui("+AlwaysOnTop +ToolWindow", "Select Tree Item")
        searchGui.SetFont("s10", "Segoe UI")

        ; Add instruction text
        searchGui.AddText("w350 Center", "Enter text to search for tree item (starts with):")

        ; Add text input field
        searchGui.AddEdit("w300 Center vTreeItemInput")

        ; Submit handler
        SubmitTreeItemSearch(ctrl, *) {
            global g_TreeItemSearchInput
            g_TreeItemSearchInput := ctrl.Gui["TreeItemInput"].Text
            ctrl.Gui.Destroy()
        }

        ; Cancel handler
        CancelTreeItemSearch(ctrl, *) {
            global g_TreeItemSearchInput
            g_TreeItemSearchInput := ""
            ctrl.Gui.Destroy()
        }

        ; Add OK and Cancel buttons (OK is default, triggered by Enter)
        okBtn := searchGui.AddButton("w80 xp-40 y+10 Default", "OK")
        okBtn.OnEvent("Click", SubmitTreeItemSearch)
        cancelBtn := searchGui.AddButton("w80 xp+90", "Cancel")
        cancelBtn.OnEvent("Click", CancelTreeItemSearch)

        ; Show GUI and focus input
        searchGui.Show("w350 h150")
        searchGui["TreeItemInput"].Focus()

        ; Wait for dialog to close
        WinWaitClose("ahk_id " searchGui.Hwnd)

        ; Get the input value
        global g_TreeItemSearchInput
        searchText := g_TreeItemSearchInput
        g_TreeItemSearchInput := ""  ; Clear for next use

        ; If user cancelled, exit
        if (searchText = "")
            return

        ; Get root element and find Tree
        root := UIA.ElementFromHandle(WinExist("A"))
        Sleep 200

        ; Find Tree container with AutomationId="4"
        treeContainer := root.FindFirst({ Type: "Tree", AutomationId: "4" })
        if (!treeContainer) {
            MsgBox "Could not find the tree container (AutomationId='4').", "UIA Tree Inspector", "IconX"
            return
        }

        ; Get all TreeItem children
        treeItems := treeContainer.FindAll({ Type: "TreeItem" })
        if (!treeItems) {
            MsgBox "No tree items found in the tree container.", "UIA Tree Inspector", "IconX"
            return
        }

        ; Search for TreeItem where Name starts with searchText (case-insensitive)
        matchingItem := ""
        searchTextLower := StrLower(searchText)
        for item in treeItems {
            if (!item)
                continue
            try {
                itemName := item.Name
                if (StrLower(SubStr(itemName, 1, StrLen(searchText))) = searchTextLower) {
                    matchingItem := item
                    break
                }
            } catch {
                ; Skip items without names
                continue
            }
        }

        ; Select the matching item
        if (matchingItem) {
            try {
                matchingItem.Select()
                ; Optional: Scroll into view and set focus
                matchingItem.ScrollIntoView()
                matchingItem.SetFocus()

                ; Workaround: force UIA Tree Inspector to refresh the right-side UIA tree
                ; by "jiggling" selection Down then Up after selection via search.
                inspectorHwnd := WinExist("UIATreeInspector")
                if (!inspectorHwnd)
                    inspectorHwnd := WinExist("ahk_exe UIATreeInspectorAutoHotkey64.exe")

                if (inspectorHwnd) {
                    if !WinActive("ahk_id " inspectorHwnd) {
                        WinActivate "ahk_id " inspectorHwnd
                        WinWaitActive "ahk_id " inspectorHwnd, , 1
                    }

                    if WinActive("ahk_id " inspectorHwnd) {
                        ; Ensure the list/tree has focus before sending arrow keys
                        try matchingItem.SetFocus()

                        Sleep 500
                        Send "{Down}"
                        Sleep 1000
                        Send "{Up}"
                    }
                }
            } catch Error as e {
                MsgBox "Error selecting tree item:`n" e.Message, "UIA Tree Inspector", "IconX"
            }
        } else {
            MsgBox Format("No tree item found starting with '{}'.", searchText), "UIA Tree Inspector", "IconX"
        }

    } catch Error as e {
        MsgBox "Error in tree item search:`n" e.Message, "UIA Tree Inspector", "IconX"
    }
}

; Shift + C : Search window/control and copy full UIA tree to clipboard
+c:: {
    barShown := false
    try {
        ; Global variable to store user input
        global g_TreeItemSearchInput := ""

        StandardLoadingBar_Show("⏳ UIA Tree Inspector: preparing copy…", BANNER_ACCENT_INTERMEDIATE, { passive: false })
        barShown := true

        ; Create GUI dialog (same as Shift+S)
        searchGui := Gui("+AlwaysOnTop +ToolWindow", "Select Tree Item")
        searchGui.SetFont("s10", "Segoe UI")
        searchGui.AddText("w350 Center", "Enter text to search for tree item (starts with):")
        searchGui.AddEdit("w300 Center vTreeItemInput")

        SubmitTreeItemSearch(ctrl, *) {
            global g_TreeItemSearchInput
            g_TreeItemSearchInput := ctrl.Gui["TreeItemInput"].Text
            ctrl.Gui.Destroy()
        }
        CancelTreeItemSearch(ctrl, *) {
            global g_TreeItemSearchInput
            g_TreeItemSearchInput := ""
            ctrl.Gui.Destroy()
        }

        okBtn := searchGui.AddButton("w80 xp-40 y+10 Default", "OK")
        okBtn.OnEvent("Click", SubmitTreeItemSearch)
        cancelBtn := searchGui.AddButton("w80 xp+90", "Cancel")
        cancelBtn.OnEvent("Click", CancelTreeItemSearch)

        searchGui.Show("w350 h150")
        searchGui["TreeItemInput"].Focus()
        WinWaitClose("ahk_id " searchGui.Hwnd)

        global g_TreeItemSearchInput
        searchText := g_TreeItemSearchInput
        g_TreeItemSearchInput := ""
        if (searchText = "")
            return

        StandardLoadingBar_Update("🔎 Selecting window/control…")
        inspectorHwnd := WinExist("UIATreeInspector")
        if (!inspectorHwnd)
            inspectorHwnd := WinExist("ahk_exe UIATreeInspectorAutoHotkey64.exe")
        if (!inspectorHwnd)
            return

        if !WinActive("ahk_id " inspectorHwnd) {
            WinActivate "ahk_id " inspectorHwnd
            WinWaitActive "ahk_id " inspectorHwnd, , 1
        }
        if !WinActive("ahk_id " inspectorHwnd)
            return

        ; Select matching item in left tree (AutomationId="4") (same as Shift+S)
        root := UIA.ElementFromHandle(inspectorHwnd)
        Sleep 500
        treeContainer := root.FindFirst({ Type: "Tree", AutomationId: "4" })
        if (!treeContainer)
            return

        treeItems := treeContainer.FindAll({ Type: "TreeItem" })
        if (!treeItems)
            return

        matchingItem := ""
        searchTextLower := StrLower(searchText)
        for item in treeItems {
            if (!item)
                continue
            try {
                itemName := item.Name
                if (StrLower(SubStr(itemName, 1, StrLen(searchText))) = searchTextLower) {
                    matchingItem := item
                    break
                }
            } catch {
                continue
            }
        }
        if (!matchingItem)
            return

        matchingItem.Select()
        matchingItem.ScrollIntoView()
        matchingItem.SetFocus()
        Sleep 1500

        ; Refresh workaround (force correct UIA Tree load)
        StandardLoadingBar_Update("🔄 Refreshing UIA Tree…")
        if !WinActive("ahk_id " inspectorHwnd) {
            WinActivate "ahk_id " inspectorHwnd
            WinWaitActive "ahk_id " inspectorHwnd, , 1
        }
        if WinActive("ahk_id " inspectorHwnd) {
            try matchingItem.SetFocus()
            Send "{Down}"
            Sleep 2000
            Send "{Up}"
        }

        ; Focus UIA Tree panel (right-side tree) and select root
        StandardLoadingBar_Update("🌳 Focusing UIA Tree panel…")
        Sleep 1500
        rightTree := 0
        bestL := -0x7FFFFFFF
        trees := root.FindAll({ Type: "Tree" })
        for t in trees {
            if (!t)
                continue
            try {
                if (t.AutomationId = "4")
                    continue
            } catch {
            }
            try {
                br := t.BoundingRectangle
                if (br.l > bestL) {
                    bestL := br.l
                    rightTree := t
                }
            } catch {
                if (!rightTree)
                    rightTree := t
            }
        }
        if (!rightTree)
            return

        rootItem := rightTree.FindFirst({ Type: "TreeItem" })
        if (rootItem) {
            try rootItem.Select()
            try rootItem.ScrollIntoView()
            try rootItem.SetFocus()
        } else {
            try rightTree.SetFocus()
        }
        Sleep 1500

        ; Copy complete UI tree to clipboard via context menu
        StandardLoadingBar_Update("📋 Copying full tree to clipboard…")
        A_Clipboard := ""
        Sleep 400
        Send "{AppsKey}"
        Sleep 600
        Send "{Up}"
        Sleep 400
        Send "{Enter}"
        Sleep 1200
    } catch Error as e {
        try StandardLoadingBar_Update("❌ Copy failed: " . SubStr(e.Message, 1, 60))
        try StandardLoadingBar_Hide(2000)
        MsgBox "Error in Shift+C UIA Tree copy:`n" e.Message, "UIA Tree Inspector", "IconX"
        return
    } finally {
        if (barShown)
            try StandardLoadingBar_Hide(0)
    }
}
#HotIf

;-------------------------------------------------------------------
; SettleUp Shortcuts
;-------------------------------------------------------------------
SettleUp_GetNewExpenseDialog() {
    try {
        uia := UIA_Browser()
        return uia.FindElement({ Name: "New expense", Type: "Group" })
    } catch
        return 0
}

#HotIf WinActive("Settle Up")

; Shift + A : Click Add Transaction button (UIA by Name, EN/PT)
+a:: {
    try {
        uia := UIA_Browser()
        Sleep 150
        btn := uia.FindElement({ Type: "Button", Name: "Add transaction", matchmode: "Substring" })
        if (!btn)
            btn := uia.FindElement({ Type: "Button", Name: "Adicionar transa", matchmode: "Substring" })
        if (btn) {
            btn.Click()
        } else {
            MsgBox "Could not find the Add Transaction button."
        }
    } catch Error as e {
        MsgBox "Error clicking Add Transaction: " e.Message
    }
}

; Shift + N : Focus expense name (Purpose) field in New expense dialog
+n:: {
    try {
        dialog := SettleUp_GetNewExpenseDialog()
        if (!dialog)
            return
        nameEdit := dialog.FindElement({ Type: 50004, Name: "e.g.", matchmode: "Substring" })
        if (nameEdit)
            nameEdit.SetFocus()
    } catch
        return
}

; Shift + V : Focus expense value (amount) field in New expense dialog
+v:: {
    try {
        dialog := SettleUp_GetNewExpenseDialog()
        if (!dialog)
            return
        valueEdit := dialog.FindElement({ Type: 50004 }, 4, 1)
        if (valueEdit)
            valueEdit.SetFocus()
    } catch
        return
}

#HotIf

;-------------------------------------------------------------------
; Miro Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe chrome.exe") && InStr(WinGetTitle("A"), "Miro", false)

; (removed) Shift + Y : Command palette (Ctrl+K)

; Shift + F : Frame List (Ctrl+Shift+F)
+f:: Send "^+f"

; Shift + G : Group (Ctrl+G)
+g:: Send "^g"

; Shift + U : Ungroup (Ctrl+Shift+G)
+u:: {
    ; #region agent log (6169a3)
    try {
        title := WinGetTitle("A")
        exe := WinGetProcessName("A")
        cls := WinGetClass("A")
        shiftP := GetKeyState("Shift", "P")
        titleEsc := StrReplace(StrReplace(title, "\", "\\"), '"', '\"')
        exeEsc := StrReplace(StrReplace(exe, "\", "\\"), '"', '\"')
        clsEsc := StrReplace(StrReplace(cls, "\", "\\"), '"', '\"')
        FileAppend(
            '{"sessionId":"6169a3","runId":"run1","hypothesisId":"H1-H4","location":"Shift keys.ahk:miro:+u","message":"Miro +u fired (pre-send)","data":{"title":"' titleEsc '","exe":"' exeEsc '","class":"' clsEsc '","shiftPhysical":' (shiftP ? "true" : "false") '},"timestamp":' A_TickCount '}'
            "`n",
            "debug-6169a3.log",
            "UTF-8"
        )
    } catch {
    }
    ; #endregion agent log (6169a3)

    ; Using {Blind} so the physical Shift from +u contributes to Ctrl+Shift+G.
    Send "{Blind}^g"

    ; #region agent log (6169a3)
    try {
        shiftP2 := GetKeyState("Shift", "P")
        FileAppend(
            '{"sessionId":"6169a3","runId":"run1","hypothesisId":"H2-H3","location":"Shift keys.ahk:miro:+u","message":"Miro +u sent {Blind}^g (post-send)","data":{"shiftPhysicalAfter":' (shiftP2 ? "true" : "false") '},"timestamp":' A_TickCount '}'
            "`n",
            "debug-6169a3.log",
            "UTF-8"
        )
    } catch {
    }
    ; #endregion agent log (6169a3)
}

; Shift + L : Lock/Unlock (Ctrl+Shift+L)
+l:: Send "^+l"

; Shift + K : Add/Edit Link (Alt+Ctrl+K)
+k:: Send "!^k"

; Shift + X : Close sidebar - Close sidebar
+x:: {
    try {
        uia := UIA_Browser()
        if !IsObject(uia) {
            ; Fallback: try keyboard shortcut if UIA fails
            Send "^+s"
            return
        }

        ; Retry logic: Try multiple times with delays to allow UI to load
        maxRetries := 3
        retryDelay := 300  ; milliseconds between retries
        closeButton := ""

        loop maxRetries {
            ; Strategy 1: Find by Name "Close sidebar" with Type Button (50000)
            try {
                closeButton := uia.FindFirst({ Type: "50000", Name: "Close sidebar", cs: false })
                if (closeButton) {
                    ; Verify button is valid
                    try {
                        btnName := closeButton.Name
                        if (btnName) {
                            break  ; Found valid button, exit retry loop
                        }
                    } catch {
                        closeButton := ""
                    }
                }
            } catch {
            }

            ; Strategy 2: Try case-sensitive search
            if (!closeButton) {
                try {
                    closeButton := uia.FindFirst({ Type: "50000", Name: "Close sidebar" })
                    if (closeButton) {
                        try {
                            btnName := closeButton.Name
                            if (btnName) {
                                break
                            }
                        } catch {
                            closeButton := ""
                        }
                    }
                } catch {
                }
            }

            ; Strategy 3: Try by ControlType "Button" and Name substring
            if (!closeButton) {
                try {
                    closeButton := uia.FindFirst({ ControlType: "Button", Name: "Close sidebar", matchmode: "Substring" })
                    if (closeButton) {
                        try {
                            btnName := closeButton.Name
                            if (btnName) {
                                break
                            }
                        } catch {
                            closeButton := ""
                        }
                    }
                } catch {
                }
            }

            ; Strategy 4: Search all buttons and find by name match
            if (!closeButton) {
                try {
                    allButtons := uia.FindAll({ Type: "50000" })
                    for button in allButtons {
                        try {
                            btnName := button.Name
                            if (InStr(btnName, "Close sidebar") || InStr(btnName, "close sidebar")) {
                                closeButton := button
                                break
                            }
                        } catch {
                            continue
                        }
                    }
                    if (closeButton) {
                        break
                    }
                } catch {
                }
            }

            Sleep retryDelay  ; Wait before next attempt
        }

        ; Confirmation layer: Verify button was found before clicking
        if (closeButton) {
            ; Additional verification: ensure button is still valid and clickable
            try {
                ; Check if button is enabled and visible
                isEnabled := closeButton.GetPropertyValue(UIA.Property.IsEnabled)
                isOffscreen := closeButton.GetPropertyValue(UIA.Property.IsOffscreen)

                if (!isEnabled || isOffscreen) {
                    ; Button found but not usable, try keyboard shortcut fallback
                    Send "^+s"
                    return
                }
            } catch {
                ; Property check failed, continue with click attempt
            }

            ; Try multiple click strategies in order of preference
            clicked := false

            ; Strategy 1: Try Invoke pattern (most reliable for buttons)
            try {
                closeButton.Invoke()
                clicked := true
            } catch {
            }

            ; Strategy 2: Try SetFocus then Click
            if (!clicked) {
                try {
                    closeButton.SetFocus()
                    Sleep 50
                    closeButton.Click()
                    clicked := true
                } catch {
                }
            }

            ; Strategy 3: Force coordinate-based click using "left" parameter
            if (!clicked) {
                try {
                    closeButton.Click("left")
                    clicked := true
                } catch {
                }
            }

            ; Strategy 4: Direct coordinate click using element Location
            if (!clicked) {
                try {
                    pos := closeButton.Location
                    if (pos && pos.w > 0 && pos.h > 0) {
                        ; Save current mouse position
                        MouseGetPos(&prevX, &prevY)

                        ; Click at center of element
                        CoordMode("Mouse", "Screen")
                        Click(pos.x + pos.w // 2, pos.y + pos.h // 2)
                        Sleep 50

                        ; Restore mouse position
                        MouseMove(prevX, prevY)
                        clicked := true
                    }
                } catch {
                }
            }

            ; If all click strategies failed, use keyboard shortcut fallback
            if (!clicked) {
                Send "^+s"
            }
        } else {
            ; Button not found after all retries, use keyboard shortcut fallback
            Send "^+s"
        }
    } catch Error as err {
        ; If any error occurs, use keyboard shortcut as fallback
        Send "^+s"
    }
}

#HotIf

;-------------------------------------------------------------------
; PowerToys Command Palette Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("Command Palette")

; Ctrl + H : Trigger Ctrl+Shift+E
^h:: Send "^+e"

; Shift + C : Trigger Ctrl+Shift+C (Copy file path)
+c:: Send "^+c"

; Shift + B : Go Home (select all then delete 3 chars)
+b:: {
    Send "^a"
    Sleep 30
    Send "{Backspace 3}"
}

; Shift + U : Insert double quotes twice, then hit left arrow
+u:: Send '""{Left}'

; Shift + O : Focus on Folders Only
+o:: {
    Send "!+w"
    Sleep 120
    Send "{Tab}"
    Sleep 30
    Send "{Enter}"
    Sleep 30
    Send "{Up}"
    Sleep 30
    Send "{Up}"
    Sleep 30
    Send "{Down}"
    Sleep 30
    Send "{Enter}"
    Sleep 50
    Send "{Tab}"
}

; Shift + P : Focus on Files Only
+p:: {
    Send "!+w"
    Sleep 120
    Send "{Tab}"
    Sleep 30
    Send "{Enter}"
    Sleep 30
    Send "{Up}"
    Sleep 30
    Send "{Up}"
    Sleep 30
    Send "{Down}"
    Sleep 30
    Send "{Down}"
    Sleep 30
    Send "{Enter}"
    Sleep 50
    Send "{Tab}"
}

; Shift + I : Trigger Command Palette Bookmark "add new bookmark" shortcut
+i:: {
    Send "^!#m"
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

; Easy Selection
; Alt + 1 : Easy Selection - 1st item
!1::
{
    Send "{Enter}"
}

; Alt + 2 : Easy Selection - 2nd item
!2::
{
    Send "{Down}"
    Send "{Enter}"
}

; Alt + 3 : Easy Selection - 3rd item
!3::
{
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Alt + 4 : Easy Selection - 4th item
!4::
{
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

; Alt + 5 : Easy Selection - 5th item
!5::
{
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Down}"
    Send "{Enter}"
}

#HotIf

; --- Unified banner helpers for ChatGPT indicators (use Utils standard loading bar) ---
ShowSmallLoadingIndicator_ChatGPT(state := "Loading…", bgColor := BANNER_ACCENT_INTERMEDIATE) {
    StandardLoadingBar_Show(state, bgColor, { passive: true, centerOnHwnd: 0, textWidth: 500, fontSize: 17,
        passiveBgColor: bgColor })  ; callers pass BANNER_ACCENT_INTERMEDIATE or other semantic constant
}

HideSmallLoadingIndicator_ChatGPT() {
    StandardLoadingBar_Hide(0)
}

; Short completion chime for ChatGPT responses (debounced)
PlayCompletionChime_ChatGPT() {
    try {
        static lastTick := 0
        if (A_TickCount - lastTick < 1500)
            return
        lastTick := A_TickCount

        if (!IsSoundEnabled()) {
            return
        }

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
