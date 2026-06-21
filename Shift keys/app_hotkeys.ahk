; =============================================================================
; Shift keys module: app_hotkeys.ahk
; Per-app Win+Alt+Shift hotkey cheat sheet definitions
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

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
    🥇 [J][J]ump to first mail
    ◧ [H]Toggle high [H] navigation pane
    📝 [S][S]ubject / Title
    👥 [T][T]o / Required
    📝 [B][B]ody (Subject → Body)
    🎯 [F][F]ocused / Other
    🔀 [K]Cycle bac[K]ward pane
    🔀 [L]Cyc[L]e forward pane
    📋 [M]Toggle Mail / Calendar
    📅 [W]eek view
    📅 M[o]nth view
    🪟 [P][P]op Out (Type 50000, Name Pop Out, LocalizedType button, Path {T:33,CN:rctrl_renwnd32}, {T:33}, {T:33}, {T:21}, {T:0, i:-1})
    
    📅 Meeting request (reading pane or popped-out invitation - same shortcuts)
    ✅ [A][A]ccept meeting invitation
    ❌ [D][D]ecline meeting (confirmation first)
    📌 [Alt+F] Follow (updates from organizer)
    ❓ [T][T]entative (More options …) - when a meeting request is open, runs before [T]o / Required
    
    📅 Canceled meeting (organizer canceled — Remove event)
    🗑️ [E]Remove event (no confirmation)
    
    Outlook (Ctrl+Alt)
    📌 Ribbon actions select the Home tab first when needed (e.g. View or Help was active).
    🔎 [Ctrl+Alt+F] Search
    📮 [Ctrl+Alt+M] Mail view
    📅 [Ctrl+Alt+G] Calendar view
    📃 [Ctrl+Alt+L] Focus message list
    📖 [Ctrl+Alt+P] Focus reading pane
    
    ↩️ [Ctrl+Alt+R] Reply
    👥 [Ctrl+Alt+A] Reply all
    ➡️ [Ctrl+Alt+W] Forward
    🗑️ [Ctrl+Alt+D] Delete
    🗄️ [Ctrl+Alt+E] Archive
    ✅ [Ctrl+Alt+U] Read/Unread
    🏷️ [Ctrl+Alt+C] Categorize
    📁 [Ctrl+Alt+V] Move
    🧪 [Ctrl+Alt+I] Filter menu
    ↕️ [Ctrl+Alt+S] Sort menu
    
    🆕 [Ctrl+Alt+N] New (Calendar: New event / Mail: new message)
    🧭 [Ctrl+Alt+T] Today (Calendar)
)"  ; end Outlook

; --- Outlook Reminder window -------------------------------------------------
cheatSheets["OutlookReminder"] := "
(
    Outlook - Reminders (Shift)
    ⏰ [H]Snooze 1 [H]our
    ⏰ [F]Snooze [F]our hours
    ⏰ [T]Snooze 10 minu[T]es
    ⏰ [Y]Snooze 1 da[Y]
    ⏰ [W]Snooze 1 [W]eek
    ✅ [D][D]ismiss reminder (selected item)
    ❌ [X]Dismiss all reminders (confirm)
    🌐 [J][J]oin online (selected item)
)"  ; end Outlook Reminder

; --- Outlook Appointment window ---------------------------------------------
cheatSheets["OutlookAppointment"] := "
(
    Outlook - Appointment (Shift)
    📅 [S][S]tart date (popover)
    🕐 [T]Start [T]ime (popover)
    🕐 [E][E]nd time (popover)
    ☑️ [A][A]ll-day toggle (popover)
    🔄 [C]Re[C]urring / Series (popover)
    🕒 [1]Time suggestion 1
    🕒 [2]Time suggestion 2
    
    📝 [I]T[I]tle field
    👥 [R][R]equired attendees
    📍 [O]L[o]cation / Add a room
    📝 [B][B]ody (main details)
    
    🎥 [M]Tea[M]s meeting
    🧬 [U]Series (recurring)
    📶 A[V]ailability (Free/Busy...)
    ⏰ Reminder (fre[Q]uency)
    🏷️ Cate[G]ory
    🔒 [P]rivate / Not private
    
    🗓️ Prev day [K] / Next day [L]
    🧭 [Y]Today
    📆 [D]ate header
    🧑‍🤝‍🧑 Sc[H]eduler / Scheduling assistant
    ⚙️ Optio[N]s (scheduler view)
    🌐 Time [Z]one
    ➕ [J]Add required attendee
    ➕ [Alt+O]Add optional attendee
    
    ↩️ [Backspace] Back (scheduler view)
    🧙 [W][W]izard (configure)
)"  ; end Outlook Appointment

; --- Outlook Message window ---------------------------------------------------
cheatSheets["OutlookMessage"] := "
(
    Outlook - Message (Shift)
    📅 Meeting invitation (same shortcuts as main Mail reading pane)
    ✅ [A][A]ccept meeting invitation
    ❌ [D][D]ecline meeting (confirmation first)
    📌 [Alt+F] Follow (updates from organizer)
    ❓ [T][T]entative (More options …) - when a meeting request is open, runs before [T]o / Required
    📅 Canceled meeting (Remove event)
    🗑️ [E]Remove event (no confirmation)
    📝 [S][S]ubject / Title
    👥 [T][T]o / Required
    📝 [B][B]ody (Location → Body)
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
    🔙 [K]Back (toolbar)
    ⏩ [L]Forward (toolbar)
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
    👍 [Y] Like reaction
    ❤️ [G][G]ive heart reaction
    😂 [J][J]oke reaction (😂)
    
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

; --- Google Maps (Chrome) ---------------------------------------------------
cheatSheets["Google Maps"] := "
(
    Google Maps (Shift)
    🔍 [S][S]earch box (place / query)
    📍 [L][L]at/long (copy coordinates to clipboard)
)"  ; end Google Maps

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
    
    === Ctrl (no other modifiers) ===
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
    🤖 [E] Maximize chat size — native Cursor (`workbench.action.maximizeChatSize`; user keybinding)
    📂 [R]File open [R]ecent
    🔍 [T]Go to [T]ype symbol in workspace
    💬 [N] [N]ew chat tab (replacing current)
    ➕ [Enter] [I]nsert line below
    🔍 [P]Open [P]roject
    💬 [;] Insert comment
    📝 [D]Duplicate selection to next find match
    🔍 [F] [F]ind
    ↩️ [Z]Undo (common [Z])
    📊 [B]Toggle [B]ar (primary sidebar)
    
    === Shift ===
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
    
    === Alt (ahk = AutoHotkey) ===
    📉 [x] Shri[X]nk selection (ahk)
    📉 [,] Classical Markdown Preview
    📉 [Y] Paste image to Markdown
    ⬇️ [U] Scroll AI feed to bottom (ahk-based)
    📋 [M] Quick shortcut menu (ahk)
    🤖 [A] Add file to AI Context (Cursor Chat) (ahk)
    📌 [Q] Unpin current tab
    📌 [P] [P]in current tab
    📋 [I] Reveal in Explorer + copy file (ahk)
    📂 [H] Reveal in Explorer + open file (ahk)
    📄 [R] Refresh preview
    📄 [F] File: New [F]ile
    📂 [O] File: New F[O]lder
    
    === Ctrl+Shift ===
    📝 [Ctrl+Shift+L] Select all identical words ([L]ines)
    🐛 [Ctrl+Shift+D] [D]ebugging
    
    === Ctrl+Alt ===
    📄 [Ctrl+Alt+L] Markdown Preview Enhanced: Toggle Live Update
    📄 [Ctrl+Alt+T] Markdown Preview Enhanced: Toggle Scroll Sync
    ⬆️ [Ctrl+Alt+Up] Go to [P]arent Fold
    ⬅️ [Ctrl+Alt+Left] Go to sibling fold [P]revious
    ➡️ [Ctrl+Alt+Right] Go to sibling fold [N]ext
    ⬆️ [Ctrl+Alt+↑] Add cursor [A]bove
    ⬇️ [Ctrl+Alt+↓] Add cursor [B]elow
    
    === Alt+Shift ===
    ⬆️ [Shift+Alt+↑] [C]opy line Up
    ⬇️ [Shift+Alt+↓] [C]opy line Down
    
    === Alt (other chords) ===
    👁️ [Alt+F12] [P]eek Definition
    ⬆️ [Alt+↑] [M]ove line Up
    ⬇️ [Alt+↓] [M]ove line Down
    👆 [Alt+Click] [M]ulti-cursor by click
    🔄 [Alt+Z] Toggle word [W]rap
    ⬇️ [Alt+J] Jump to [N]ext review
    ⬆️ [Alt+K] [P]revious review (bac[K])
    
    === Function keys & misc ===
    ✏️ [F2] [R]ename symbol
    🔍 [F8] [N]avigate problems
    🗑️ [Shift+Delete] [D]elete line
)"  ; end Cursor

cheatSheets["Code.exe"] := "
(
    VS Code
    
    === Ctrl (no other modifiers) ===
    🎯 [1] Remove clustering and focus on the code (ahk)
    📁 [2] Copy path (VS Code)
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
    💬 [.] Copilot Agent Modes
    🤖 [E] VS Code default behavior (Cursor custom maximize removed)
    📂 [R]File open [R]ecent
    🔍 [T]Go to [T]ype symbol in workspace
    💬 [N] Copilot chat session workflow (pending dedicated remap)
    ➕ [Enter] [I]nsert line below
    🔍 [P]VS Code quick open / project search
    💬 [;] Insert comment
    📝 [D]Duplicate selection to next find match
    🔍 [F] [F]ind
    ↩️ [Z]Undo (common [Z])
    📊 [B]Toggle [B]ar (primary sidebar)
    
    === Shift ===
    📉 [F][F]old (ahk)
    📈 [U][U]nfold (ahk)
    📄 [M][M]arkdown preview (VS Code migration pending)
    🪟 [W][W]indow (move editor) (VS Code migration pending)
    💻 [T][T]erminal (ahk)
    💻 [N][N]ew Terminal (ahk)
    📁 [E][E]xplorer (ahk)
    📄🪟 [K] Mar[K]down + window (ahk)
    ⌨️ [C][C]ommand palette (ahk)
    📈 [X] E[X]pand selection (ahk)
    ⚡ [S][S]ymbol in access view (VS Code migration pending)
    💬 [H][H]istory (chat) (ahk)
    🖼️ [I][I]mage (paste) (VS Code migration pending)
    📁 [G][G]it repos fold (SCM) (ahk)
    🔍 [Q][Q]uery Search (ahk)
    🍞 [R]B[R]eadcrumbs menu (ahk)
    😀 [O]Emoji selector (em[O]ji) (ahk)
    🌿 [D]Git section ([D]iff) (ahk)
    ❌ [Z]Close all editors (end [Z]one) (ahk)
    🤖 [A][A]I models switch (ahk)
    🧘 [Y]Zen mode (tranquilit[Y])
    ⬇️ [P][P]ull (Git)
    ✅ [V]Commit (Git)
    ⬆️ [B]Push (Git)
    
    === Alt (ahk = AutoHotkey) ===
    📉 [x] Shri[X]nk selection (ahk)
    📉 [,] Classical Markdown Preview
    📉 [Y] Paste image to Markdown
    ⬇️ [U] Scroll AI feed to bottom (ahk-based)
    📋 [M] Quick shortcut menu (ahk)
    ➕ [C] Add Context picker (VS Code chat) (ahk)
    🤖 [A] Add file to AI Context (VS Code chat) (ahk)
    📌 [Q] Unpin current tab
    📌 [P] [P]in current tab
    📋 [I] Reveal in Explorer + copy file (ahk)
    📂 [H] Reveal in Explorer + open file (ahk)
    📄 [R] Refresh preview
    📄 [F] File: New [F]ile
    📂 [O] File: New F[O]lder
    
    === Ctrl+Shift ===
    📝 [Ctrl+Shift+L] Select all identical words ([L]ines)
    🐛 [Ctrl+Shift+D] [D]ebugging
    
    === Ctrl+Alt ===
    📄 [Ctrl+Alt+L] Markdown Preview Enhanced: Toggle Live Update
    📄 [Ctrl+Alt+T] Markdown Preview Enhanced: Toggle Scroll Sync
    ⬆️ [Ctrl+Alt+Up] Go to [P]arent Fold
    ⬅️ [Ctrl+Alt+Left] Go to sibling fold [P]revious
    ➡️ [Ctrl+Alt+Right] Go to sibling fold [N]ext
    ⬆️ [Ctrl+Alt+↑] Add cursor [A]bove
    ⬇️ [Ctrl+Alt+↓] Add cursor [B]elow
    
    === Alt+Shift ===
    ⬆️ [Shift+Alt+↑] [C]opy line Up
    ⬇️ [Shift+Alt+↓] [C]opy line Down
    
    === Alt (other chords) ===
    👁️ [Alt+F12] [P]eek Definition
    ⬆️ [Alt+↑] [M]ove line Up
    ⬇️ [Alt+↓] [M]ove line Down
    👆 [Alt+Click] [M]ulti-cursor by click
    🔄 [Alt+Z] Toggle word [W]rap
    ⬇️ [Alt+J] Jump to [N]ext review
    ⬆️ [Alt+K] [P]revious review (bac[K])
    
    === Function keys & misc ===
    ✏️ [F2] [R]ename symbol
    🔍 [F8] [N]avigate problems
    🗑️ [Shift+Delete] [D]elete line
)"  ; end VS Code

; --- Windows Explorer ------------------------------------------------------
cheatSheets["explorer.exe"] := "
(
    Explorer (Shift)
    📄 [F]Select first [F]ile
    🔍 [S][S]earch bar
    📍 [A][A]ddress bar
    📁 [N][N]ew Folder²
    🔗 [H]Create s[H]ortcut
    📋 [C][C]opy as path
    📤 [R]Sha[R]e file
    📌 [P][P]inned item (first in sidebar)
    📌 [L][L]ast item (sidebar)
    📦 [X] WinRAR e[X]tract here (personal); work: 7-Zip extract
    📦 [W] WinRAR add to archive / compact (personal); work: 7-Zip add to archive / compress
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
    ✏️ [E][E]dit Text (F4)
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
    ⌨️ [D][E]xclude Favorite
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
    📋 [Z]Copy cell Val
    📊 [P]Model view ([P]lan)
    📊 [C]Get data ([C]onnect)
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
)" 44

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

; --- Gemini (web, Chrome) -----------------------------------------------
cheatSheets["Gemini"] := "
(
    Gemini (Shift)
    📂 [D]Toggle the[D]rawer
    💬 [N][N]ew chat
    🔍 [S][S]earch
    🔄 [M]Change[M]odel
    🛠️ [T][T]ools
    🖼️ [I]Create [I]mage (Tools menu; opens if needed)
    🔬 [E]Deep r[E]search (Tools menu; opens if needed)
    ⌨️ [P]Focus[P]rompt field
    📋 [C][C]opy last message
    🔊 [R][R]ead aloud last message
    🤖 [G]Send[G]emini prompt text
    ⛶ [F][F]ullscreen input
    🔔 [Enter / Ctrl+Enter]Send and notify on completion
    
    === Alt (ahk) ===
    ⬇️ [U] Scroll AI feed to bottom — same idea as Cursor
)"

; --- M365 Copilot web (Chrome) — same Shift keys as Gemini -----------------
cheatSheets["Copilot Web"] := "
(
    Copilot Web (Shift)
    📂 [D]Toggle nav [D]rawer
    💬 [N][N]ew chat
    🔍 [S][S]earch (nav drawer)
    🔄 [M]Change [M]odel (opens selector)
    🛠️ [T]Add/manage sources (Tools menu)
    🖼️ [I]Designer / create image (Sources menu)
    🔬 [E]Researcher / deep research (Sources menu)
    ⌨️ [P]Focus [P]rompt field
    📋 [C][C]opy last response
    🔊 [R][R]ead aloud last message
    🤖 [G]Send prompt text (Gemini_Prompt.txt)
    ⛶ [F][F]ullscreen input (expand composer)
    🔔 [Enter / Ctrl+Enter]Send and notify on completion
    
    === Alt (ahk) ===
    ⬇️ [U] Scroll AI feed to bottom — same idea as Cursor
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

; Mirrors the former sequential if-chain: later assignments override earlier ones; Shopee/Google only when still unset.
PickChromeAppSheetKey(chromeTitle) {
    key := ""
    if IsChromePdfViewerActive()
        key := "Chrome PDF Viewer"
    if InStr(chromeTitle, "WhatsApp")
        key := "WhatsApp"
    if InStr(chromeTitle, "Gmail")
        key := "Gmail"
    if InStr(chromeTitle, "chatgpt")
        key := "ChatGPT"
    if InStr(chromeTitle, "Mobills")
        key := "Mobills"
    if InStr(chromeTitle, "Google Keep") || InStr(chromeTitle, "keep.google.com")
        key := "Google Keep"
    if InStr(chromeTitle, "YouTube")
        key := "YouTube"
    if InStr(chromeTitle, "UIATreeInspector")
        key := "UIATreeInspector"
    if InStr(chromeTitle, "Settle Up")
        key := "Settle Up"
    if InStr(chromeTitle, "Miro")
        key := "Miro"
    if InStr(chromeTitle, "Wikipedia", false) || InStr(chromeTitle, "wikipedia.org", false)
        key := "Wikipedia"
    if IsMercadoLivreActive()
        key := "Mercado Livre"
    if (key = "" && IsShopeeActive())
        key := "Shopee"
    if InStr(chromeTitle, "gemini", false)
        key := "Gemini"
    if (key = "" && (InStr(chromeTitle, "M365 Copilot", false) || InStr(chromeTitle, "Chat | M365 Copilot", false)))
        key := "Copilot Web"
    if InStr(chromeTitle, "Google Maps")
        key := "Google Maps"
    if (key = "" && (chromeTitle = "Google" || InStr(chromeTitle, " - Google Search")))
        key := "Google"
    return key
}
