; =============================================================================
; Shift keys module: cheat_sheet_registry.ahk
; Canonical registry for all cheat sheet strings (per-app cheatSheets map +
; GLOBAL_CHEAT_SHEET_RAW). Loaded via #include into Shift keys.ahk.
; =============================================================================
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
    ❌ [D][D]ecline meeting silently (no email; confirmation first)
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
    💾 [S][S]ave / Send
    📅 [D]Start [D]ate (popover)
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
    📆 [F]Date header
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
    ❌ [D][D]ecline meeting silently (no email; confirmation first)
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
    💬 [C]Open [C]hat pane1
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
    🌊 [I][I]mmerse (header Play + fullscreen)
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
    📉 [C][C]ollapse side panel
    🖼️ [P][P]NG capture (clean map / Street View)
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
    📁 [2] Copy path — Explorer then Ctrl+2 (ahk)
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
    📈 [Q] Unfold all directories
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
    ⬇️ [P][P]ull or Sync (Git) (ahk)
    ✅ [V]Commit (Git sa[V]e) (cursor)
    ⬆️ [B]Push (Git pu[B]lish) (cursor)
    ✅ [J] Quick commit — generic timestamp, commit + push (ahk)
    
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
    📤 [6] Share in OneDrive — Reveal + share link (ahk)
    📄 [R] Refresh preview
    📄 [E] File: Op[E]n File (ahk)
    📄 [F] File: New [F]ile
    📂 [O] File: New F[O]lder
    💾 [S][S]tash and Pull (Git) (ahk)
    
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
    📁 [2] Copy path — Explorer then Ctrl+2 (ahk)
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
    📈 [Q] Unfold all directories
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
    ⬇️ [P][P]ull or Sync (Git) (ahk)
    ✅ [V]Commit (Git)
    ⬆️ [B]Push (Git)
    ✅ [J] Quick commit — generic timestamp, commit + push (ahk)
    
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
    📤 [6] Share in OneDrive — Reveal + share link (ahk)
    📄 [R] Refresh preview
    📄 [E] File: Op[E]n File (ahk)
    📄 [F] File: New [F]ile
    📂 [O] File: New F[O]lder
    💾 [S][S]tash and Pull (Git) (ahk)
    
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
    🖥️ [D]Go to [D]esktop
    📁 [N][N]ew Folder²
    🔗 [H]Create s[H]ortcut wizard (paste URL/path)
    📋 [C][C]opy as path
    📤 [R]Sha[R]e file
    📌 [P][P]inned item (first in sidebar)
    📌 [L][L]ast item (sidebar)
    📦 [X] WinRAR e[X]tract here (personal); work: 7-Zip extract
    📦 [W] WinRAR add to archive / compact (personal); work: 7-Zip add to archive / compress
    🖼️ [B] Set as desktop [B]ackground (PT/EN locale-flexible)
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
    📥 [I][I]mport clips
    🔍 [Y]File t[Y]pe filter (Quick Wizard)
    ⌨️ [Esc] Minimize window
    ⌨️ [Alt+1] Paste current item, then minimize
    ⌨️ [Alt+2] Down 1, paste, then minimize
    ⌨️ [Alt+3] Down 2, paste, then minimize
    ⌨️ [Alt+4] Down 3, paste, then minimize
    ⌨️ [Alt+5] Down 4, paste, then minimize
    📋 [Alt+Enter] Paste [F]ile (Clip menu)
    📋 [Ctrl+1] F10, Select All, Copy, then minimize
    📋 [Ctrl+2] Down 1, F10, Select All, Copy, then minimize
    📋 [Ctrl+3] Down 2, F10, Select All, Copy, then minimize
    📋 [Ctrl+4] Down 3, F10, Select All, Copy, then minimize
    📋 [Ctrl+5] Down 4, F10, Select All, Copy, then minimize
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
    📬 [E]Select all visible + mark r[E]ad + deselect
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
    🖥️ [D]Go to [D]esktop
    📁 [N][N]ew Folder
    📌 [P][P]inned item (first in sidebar)
    💻 [T][T]his PC (sidebar)
    📝 [M]File na[M]e field
    💾 [U]Save CSV [U]TF-8
    📥 [I][I]mport CSV (Load → promote headers → format table → shade outside)
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
    ✏️ [E][E]dit Favorite
    🌐 [Enter] Open web bookmark in new Chrome window
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
    📥 [I][I]mport CSV (clipboard path → From Text/CSV → format → shade → save UTF-8)
    💾 [U]Save CSV [U]TF-8 (clipboard path → F12 Save As)
    📊 [C][C]SV to columns (semicolon delimited)
    📋 [V]Quickly [V]aste and extract CSV
    ➕ [A][A]dd multiple rows (10 rows)
    🗑️ [R][R]ow removal workflow (remove row, down arrow, repeat 5-7 times)
    📅 [P]Type [P]revious day date
    📏 [N][N]arrow oversized columns (autofit, cap >15→5, zoom row 1)
)"

; --- PowerPoint --------------------------------------------------------
cheatSheets["POWERPNT.EXE"] := "
(
    PowerPoint (Shift)
    📄 [P]Save as [P]DF on Desktop (COM)
    🎯 [C]enter on slide (center + middle)
    ⬅️ [L]Align [L]eft
    ➡️ [R]Align [R]ight
    ⬆️ [T]Align [T]op
    ⬇️ [B]Align [B]ottom
    ↔️ [H]Align center [H]orizontal
    ↕️ [V]Align center [V]ertical
    📏 [D][D]istribute horizontal
    📏 [Y]Distribute verticall[Y]
    🔗 [G][G]roup
    🔓 [U][U]ngroup
    ⬆️ [F]Bring to [F]ront
    ⬇️ [K]Send to bac[K]
    ☝️ [E]Bring forward (on[E] step)
    👇 [W]Send back[W]ard
    📋 [Q]Duplicate (clo[Q]ne)
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
    📤 [Q][Q]uick Publish
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

; --- Tasks (localhost dashboard, Chrome title Tasks) ----------------------
cheatSheets["Tasks"] := "
(
    Tasks
    Overlay Win+Alt+Shift+A.
    
    === Function keys & misc ===
    ⬆️ [Up] Previous item
    ⬇️ [Down] Next item
    ⬅️ [Left] Collapse, or move to parent
    ➡️ [Right] Expand, or move into first child
    🏠 [Home] First visible item
    🔚 [End] Last visible item
    🔤 [a-z 0-9] Jump to next title starting with that character
    ⏎ [Enter] Edit selected project, section, or task (not General)
    🗑️ [Delete] Delete selected project, section, or task (not General)
    🏠 [Esc] Close overlay first (inline field, image, Info, modal) — keep Work/Personal/Habits
    🏠 [Esc] Then Home — clear bento, emoji filter, selection (search stays)
    🏠 [Esc][Esc] Home and clear search (400 ms = AI_QD_DOUBLE_TAP_MS / ZMK)
    🏠 [Esc] Info form — back to list
    🏠 [Esc] Info list — close and keep dashboard selection and category filter
    ⬆️ [Up] Info list — previous info
    ⬇️ [Down] Info list — next info
    ⏎ [Enter] Info list — add info
    🔗 [Shift+Enter] Info list — open highlighted info link
    ✏️ [Shift+E] Info list — edit highlighted info
    🗑️ [Delete] Info list — delete highlighted info
    ⬅️ [Backspace] Step back (inline field, image preview, modal, search, filter, bento, then selection)
    
    === Ctrl ===
    💾 [Enter] Save open form (info, task, project, section) or inline create field
    📋 [C] Copy selected or hovered task
    ✂️ [X] Cut selected or hovered task (paste moves it)
    📋 [V] Paste copied/cut task into selected project / section
    
    === Shift ===
    💼 [1] Work column only (toggle)
    👤 [2] Personal column only (toggle)
    🔁 [3] Habits column only (toggle)
    🎯 [J] First search hit, or first item in the other of Personal/Work
    ➕ [A] Inline new task if a column or project is in context; otherwise inline new project
    📁 [P] Inline new project
    📑 [B] Inline new section (select a project, section, or task first)
    ✅ [C] Mark selected task done (not habits)
    ➡️ [D] Copy selected habit into Personal · General
    📝 [N] Open Info for the selected project or task
    🖼️ [V] Paste clipboard image as an info point (selection or Info window)
    📂 [E] Expand all projects and sections, or collapse all projects
    🧹 [F] Clear emoji filter
    ⚡ [T] Important — tag selected task, or filter the list if none selected
    ⏳ [Q] Waiting — tag selected task, or filter if none selected
    🔲 [G] General — tag selected task, or filter if none selected
    ❓ [U] Doubt — tag selected task, or filter if none selected
    [O] Clear emoji on selected task (empty), or filter tasks with no emoji
    ⚡ [I] Important on the selected task only (does not filter)
    
    === Alt ===
    🔍 [S] Focus search
)"

; --- Memory Palace (localhost :8767, Chrome title Memory Palace) ----------
cheatSheets["Memory Palace"] := "
(
    Memory Palace
    Overlay Win+Alt+Shift+A. Open: Utility [N] or Win+Alt+Shift+D hold. Packs / quick image: Import Management (#!+X / Utility [J] → [P]/[L]/[Q]).
    
    === Navigation ===
    🏠 [Esc] Close study picker / palace overlay / leave Browse / return to Practice
    🔁 [P] Toggle Practice ↔ Plans
    🆕 [L] Open latest palace for the selected study
    📖 [M] Method
    🗂 [B] Browse
    🔗 [1] Links (study video / article / favorite)
    ❓ [H] Help (glossary · Practice / Plans GitHub)
    
    === Study picker (first open) ===
    🔤 [a-z 1-9] Pick study by letter
    ⬆️⬇️⬅️➡️ Arrows move highlight
    ⏎ [Enter] / [Space] Confirm study → Practice + latest palace
    
    === Search ===
    🔍 [Alt+S] Focus Knowledge Atom search
    ⏎ [Enter] (while search results open) Select first match → open palace
    
    === Study links (anywhere in the app) ===
    🎬 [Shift+V] Open stored video in Chrome (new window)
    📄 [Shift+A] Open stored article in Chrome (new window)
    ⭐ [Shift+F] Open stored favorite in Chrome (new window)
    🎬 [Alt+V] Set video from active Chrome tab (address bar)
    📄 [Alt+A] Set article from active Chrome tab (address bar)
    ⭐ [Alt+F] Set favorite from active Chrome tab (address bar)
    
    === Palace overlay ===
    🖼 [F] Toggle snapshot full-screen
    📝 [D] Toggle Quote & Story (default = Concept only)
    📋 [C] Copy image prompt
    ⬅️ [←] Older palace
    ➡️ [→] Newer palace
    🏠 [Esc] Exit full-screen first, then close overlay
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
    🔄 [M]Select [M]Deep model (no-op if already active)
    ⚡ [Q]Select Fast / [Q]uick model (no-op if already active)
    📃 [L]Model [L]ist (1-9/letters; Insert/a add; E edit; Delete remove; f/d Fast/Deep)
    🛠️ [T][T]ools
    🖼️ [I]Create [I]mage (Tools menu; opens if needed)
    🔬 [E]Deep r[E]search (Tools menu; opens if needed)
    ⌨️ [P]Focus[P]rompt field
    📋 [C][C]opy last message
    🔊 [R][R]ead aloud last message
    🤖 [G]Send[G]emini prompt text
    ⛶ [F][F]ullscreen input
    ✂️ [H]Strip [H]uman reminders (keep --- + blank lines)
    🔔 [Enter / Ctrl+Enter]Send and notify on completion
    
    === Alt (ahk) ===
    ⬇️ [U] Scroll AI feed to bottom — same idea as Cursor
)"

; --- Gemini Enterprise (web, Chrome) — same Shift mnemonics where UI maps ---
cheatSheets["Gemini Enterprise"] := "
(
    Gemini Enterprise (Shift)
    📂 [D]Toggle nav [D]rawer (Menu)
    💬 [N][N]ew chat
    🔍 [S][S]earch
    🔄 [M]Select [M]Deep model (no-op if already active)
    ⚡ [Q]Select Fast / [Q]uick model (no-op if already active)
    📃 [L]Model [L]ist (1-9/letters; Insert/a add; E edit; Delete remove; f/d Fast/Deep)
    🎨 [A][A]rt: Deep + Create images + bosch-brand-image (strip reminders)
    🛠️ [T]Select [T]ools
    🖼️ [I]Create images (Tools menu; opens if needed)
    🔬 [E]Deep r[E]search (Tools menu; opens if needed)
    ⌨️ [P]Focus[P]rompt field
    ✂️ [H]Strip [H]uman reminders (keep --- + blank lines)
    🔔 [Enter / Ctrl+Enter]Submit (chime if Stop control appears)
    
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
    🔄 [M]Select [M]Deep model (no-op if already active)
    ⚡ [Q]Select Fast / [Q]uick model (no-op if already active)
    📃 [L]Model [L]ist (1-9/letters; Insert/a add; E edit; Delete remove; f/d Fast/Deep)
    🎨 [A][A]rt: Deep + Generate image + bosch-brand-image (strip reminders)
    🛠️ [T]Add/manage sources (Tools menu)
    🖼️ [I]Add capabilities → Generate an image
    🔬 [E]Add capabilities → Research a topic
    📋 [C][C]opy last response
    🔊 [R][R]ead aloud last message
    🎙️ [V]Toggle [V]oice chat (start / end)
    ⛶ [F][F]ullscreen input (expand composer)
    ✂️ [H]Strip [H]uman reminders (keep --- + blank lines)
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

; Raw text for long-hold global cheat sheet (also used by SearchCheatSheets).
GLOBAL_CHEAT_SHEET_RAW := "
(
    [Win+Alt+Shift] - PRIMARY triple modifier (most used for system-wide shortcuts)
        [Ctrl+Alt+Win] - SECONDARY triple modifier
    
    === AVAILABLE SECONDARY (Ctrl+Alt+Win) SLOTS ===
    [Ctrl+Alt+Win+N] > TEMPORARY — M365 Copilot auto-continue: send "continue", wait for Stop generating, loop (toggle off with same chord)
    [Ctrl+Alt+Win+O] > Evidence search loop — CSV row substring → PDF find; stop saves not-found rows to data/evidence_not_found.csv + 10s report (VSCodeEvidenceSearch.ahk; toggle)
    Letters available: T
    [Ctrl+Alt+Win+Y] > RESERVED — PowerToys Command Palette bookmarks (show); not bound in AHK
    Shift+CAW: A/S/D/F/Q/W/E/R (+B debug, +Z/+G/+Numpad1 fallbacks) — see === WINDOW MANAGEMENT === below
    [Ctrl+Alt+Win+G] > RESERVED — Handy: cancel dictation (define in Handy only; not bound in AHK)
    [Ctrl+Alt+Win+L] > {AI_PROVIDER} D2C direct submit (Utils.ahk; ZMK hold on L key)
    [Ctrl+Alt+Win+0] > Project Quick Selector (opens project folder in Cursor)
    [Ctrl+Alt+Win+2] > Quick Update to Your Scripts
    [Ctrl+Alt+Win+3] > Toggle Outlook and Teams
    [Ctrl+Alt+Win+5] > Clean the Clipboard
    [Ctrl+Alt+Win+4] > Mark Last Clip as Favorite (same as Ctrl+Alt+Win+J if 4 chord fails on keyboard)
    [Ctrl+Alt+Win+J] > Mark Last Clip as Favorite (alternate for keyboards that ghost Ctrl+Alt+Win+4)
    [Ctrl+Alt+Win+8] > Moves Desktop to Recycle Bin
    [Ctrl+Alt+Win+9] > Handy: Nemotron Streaming Portuguese (model slot 2)
    [Ctrl+Alt+Win+B] > Handy: Parakeet Unified English (model slot 1)
    
    === AVAILABLE Alt+Shift SLOTS ===
    Letters available: W
    [Alt+Shift+W] > available (not bound in AHK; freed from former CAW+Y placeholder)
    
    === MAIN KEY COMBINATIONS ===
    [Symbol Layer] Win+Alt+Shift - Primary combination
    [Window Management] Ctrl+Alt+Win - Secondary combination
    
    [Alt+P] Open clip angel
    [Esc] (Clip Angel focused) Minimize
    
    === CURSOR ===
    [Win+Alt+Shift+N] > Context file browser (paste path)
    
    [Win+Alt+Shift+J] > Fast Copy: tap on/off (count Ctrl+C / PrtSc / Alt+PrtSc, then paste N); hold 700ms+ repeats last N (Clip Angel)
    
    === SPOTIFY ===
    [Win+Alt+Shift+S] > Opens or activates Spotify
    
    === CLIP ANGEL ===
    [Esc] (Clip Angel focused) Minimize
    [Win+Alt+Shift+1] > Send top list item from Clip Angel
    [Win+Alt+Shift+7] > Clip Angel: tap = Edit Text (F4); hold 200ms+ = Paste file then hide
    
    === AI CHAT (Chrome) ===
    [Win+Alt+Shift+I] > Opens {AI_PROVIDER}
    [Win+Alt+Shift+8] > Pronunciation lookup ({AI_PROVIDER}): 1× English, 2× German, hold = EN/DE/PT ListView
    [Win+Alt+Shift+P] > {AI_PROVIDER}: 1× copy last message; 2× (within 400ms) copy last code; then 5s [Y] Desktop [F] Favorite [C] Transfer [R] Read (1× only) [W] Paste window [O] Clip Angel [N] No
    
    === DESKTOP ===
    [Win+Alt+Shift+O] > Newest Desktop item: 1× cut (return to previous window); 2× open; 3× paste clipboard file copy to Desktop; hold 700ms+ copy path as text
    
    === HANDY DICTATION ===
    [Win+Alt+Shift+0] > Start/stop dictation (transcription to clipboard)
    [Ctrl+Alt+Win+G] > Cancel dictation (Handy — user-defined; reserved in cheat sheet, not in AHK)
    [Ctrl+Alt+Win+9] > Handy: Nemotron Streaming Portuguese (picker slot 2; same as Win+Alt+Shift+C then 2)
    [Ctrl+Alt+Win+B] > Handy: Parakeet Unified English (picker slot 1; same as Win+Alt+Shift+C then 1)
    [Win+Alt+Shift+C] > AI model picker (Handy): 1 Parakeet Unified EN, 2 Nemotron Streaming, 3 Cohere Transcribe
    [Send dictation? B] > Toggle Parakeet Unified EN ↔ Cohere Transcribe, re-transcribe newest History entry, copy to clipboard, re-open menu
    [Send dictation? T] > Convert to Task pack (prompt char k → TASK_PACK.txt); not legacy mtask
    
    === PROMPTS ===
    [Win+Alt+Shift+H] > Utility Shortcuts → Prompts (prompt manager; same as #!+U then R)
    [k] > Convert to Task pack (TASK_PACK.txt; filters work|personal|habits); dictation Send [T] uses same
    [2] > Quick task lines (legacy mtask emoji lines; not TASK_PACK)
    [t] > Fill CSV (unrelated to Tasks pack)
    
    === GOOGLE ===
    [Win+Alt+Shift+F] > Opens Google
    
    === CURSOR ===
    [Win+Alt+Shift+,] > Opens or activates Cursor
    [Win+Alt+Shift+C] > Handy AI model picker (see HANDY DICTATION)
    
    === OUTLOOK ===
    [Win+Alt+Shift+B] > Open mail
    [Win+Alt+Shift+V] > Open Reminder
    [Win+Alt+Shift+G] > Open calendar
    [Win+Alt+Shift+D] > Voice aloud the email
    
    === MICROSOFT TEAMS ===
    [Win+Alt+Shift+R] > 1× Teams jump / 2× WhatsApp jump
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
    [Win+Alt+Shift+9] > 1× AI Quick Download (configured click sequences → Desktop → cut; manage via Win+Alt+Shift+U → Sequences) · 2× Audio / Bluetooth (1/B BT, 2/I Input, 3/O Output, 4/H Help, 5/G Ignored; Enter default, D/E enable, C/X connect, I isolate, N ignore, Esc back) · hold 700ms+ Push scripts+notes (Utility [G])
    [Win+Alt+Shift+M] > Maximizes the current window
    [Win+Alt+Shift+Y] > Focus Mode: Black out all monitors except the one with the active window (toggle)
    [Ctrl+Alt+Shift+B] > Switch to previous window (Alt+Tab once; MEH+B; WindowManagement.ahk)
    [Ctrl+Alt+Shift+C] > Switch to second previous window (Alt+Tab twice; MEH+C; WindowManagement.ahk)
    
    === WINDOW MANAGEMENT (Ctrl+Alt+Win) ===
    Canonical list for WindowManagement.ahk — see also === WINDOWS === for Win+Alt+Shift equivalents
    [Ctrl+Alt+Win+V] > Maximize active window (also Win+Alt+Shift+M; ZMK hold on minimize/close key)
    [Ctrl+Alt+Win+X] > Snap 50/50: half-width active window + pair recent window in other half
    [Ctrl+Alt+Win+Z] > Window tools [1]: maximize lone visible window per monitor
    [Ctrl+Alt+Win+6] > Window tools [3]: fill free AutoSlot slots from background (AutoSlot ON; slotted windows stay); else tile bg ≤2/mon skipping slotted
    [Ctrl+Alt+Win+U] > Window tools [2]: hidden background list — open into free AutoSlot half/empty (AutoSlot ON) or restore in place (OFF); close mode arms with list UI
    [Ctrl+Alt+Win+P] > Window tools [4]: exit F11 fullscreen
    [Ctrl+Alt+Win+A] > Move window to monitor 1 (left-most)
    [Ctrl+Alt+Win+S] > Move window to monitor 2
    [Ctrl+Alt+Win+D] > Move window to monitor 3
    [Ctrl+Alt+Win+F] > Move window to monitor 4
    [Ctrl+Alt+Win+Shift+A] > Close window on monitor 1
    [Ctrl+Alt+Win+Shift+S] > Close window on monitor 2
    [Ctrl+Alt+Win+Shift+D] > Close window on monitor 3
    [Ctrl+Alt+Win+Shift+F] > Close window on monitor 4
    [Ctrl+Alt+Win+Shift+Z] > Close window on monitor 1 (fallback when digit-1 chord fails)
    [Ctrl+Alt+Win+Shift+G] > Close window on monitor 1 (fallback from center/IDE display)
    [Ctrl+Alt+Win+Shift+Numpad1] > Close window on monitor 1 (numpad fallback)
    [Win+Ctrl+Alt+Shift+1] > Close window on monitor 1 (fallback when Win+ is swallowed)
    [Win+Ctrl+Alt+Shift+G] > Close window on monitor 1 (fallback when Win+ is swallowed)
    [Ctrl+Alt+Win+Q] > Cycle windows on monitor 1
    [Ctrl+Alt+Win+W] > Cycle windows on monitor 2
    [Ctrl+Alt+Win+E] > Cycle windows on monitor 3
    [Ctrl+Alt+Win+R] > Cycle windows on monitor 4
    [Ctrl+Alt+Win+Shift+Q] > Minimize window on monitor 1
    [Ctrl+Alt+Win+Shift+W] > Minimize window on monitor 2
    [Ctrl+Alt+Win+Shift+E] > Minimize window on monitor 3
    [Ctrl+Alt+Win+Shift+R] > Minimize window on monitor 4
    [Ctrl+Alt+Win+Shift+B] > DEV: log taskbar-minimized background window scan (WindowManagement.ahk)
    
    === COMMAND PALETTE BOOKMARKS ===
    [Ctrl+Alt+Win+Y] > Show bookmarks (PowerToys Command Palette; not bound in AHK)
    [Ctrl+Alt+Win+M] > Add bookmark (Command Palette Bookmark extension)
    
    === GENERAL ===
    [Win+Alt+Shift+U] > Utility Shortcuts (Prompts, Projects, Macros, Hotstrings, Sequences, Finance, Memory Palace, Push [G] scripts+notes)
    [Win+Alt+Shift+W] > Utility Shortcuts → Macros (same as #!+U then M)
    [Win+Alt+Shift+L] > Paste OS clipboard (^v) to window (visible picker; after pick: Y=paste+Enter, N=paste only, Esc=abort, timeout=paste; focus learned main field if saved; Y/N to save when unknown; same as D2C [W])
    [Ctrl+Alt+Win+7] > Toggle {AI_PROVIDER} Chrome tab 1 <-> 2
    [Win+Alt+Shift+Q] > Jump mouse on the middle
    [Win+Alt+Shift+X] > Import Management (same as Utility Shortcuts [J]; [Q]=Palace quick image)
    [Win+Alt+Shift+D] > tap=Tasks :8766 · 2×=Finance · hold=Memory Palace :8767 ([T]/[F]/[N])

    === MEMORY PALACE (:8767) ===
    Launch: Utility Shortcuts [N] · Win+Alt+Shift+D hold (≥700 ms) · Chrome titled Memory Palace
    Import Management (#!+X / Utility [J] / #!+F×2): [P] PALACE_PACK · [L] PLAN_PACK · [Q] quick image (Desktop PNG/JPG → palace missing image) · [T] TASK_PACK (filters work|personal|habits)
    Markdown (practice+plans): Utility Shortcuts Push [G] when mnemonics/data is dirty (then commit+push)
    App sheet: Win+Alt+Shift+A while Memory Palace is focused (same keys below).
    [Alt+S] > Focus Knowledge Atom search
    [Shift+V] > Open study video in Chrome (new window)
    [Shift+A] > Open study article in Chrome (new window)
    [Shift+F] > Open favorite link in Chrome (new window)
    [Alt+V] > Set study video from active Chrome tab (address bar)
    [Alt+A] > Set study article from active Chrome tab (address bar)
    [Alt+F] > Set favorite link from active Chrome tab (address bar)
    [P] > Toggle Practice ↔ Plans
    [L] > Latest palace (selected study)
    [M] > Method
    [B] > Browse
    [1] > Links
    [H] > Help (glossary · Practice / Plans GitHub)
    Study picker: [a-z]/[1-9] pick · arrows move · Enter/Space confirm · Esc dismiss
    Overlay: [F] full-screen snapshot · [D] toggle Quote/Story (default Concept-only) · [C] copy prompt · ← older · → newer
    Overlay Esc: exit full-screen first, then close overlay; Esc elsewhere returns toward Practice
    [Win+Alt+Shift+→] > Show square selector (right direction)
    [Win+Alt+Shift+←] > Show square selector (left direction)
    [Win+Alt+Shift+↓] > Show square selector (down direction)
    [Win+Alt+Shift+↑] > Show square selector (up direction)
    [Win+Alt+Shift+.] > Clip Angel (copy, paste, and quit)
    
    === COMMAND PALETTE ===
    [Shift+D] > Command Palette (active): exclude current bookmark (confirm)
    [Shift+E] > Command Palette (active): edit favorite (Ctrl+K → Editar favorito / Edit bookmark)
    
    === SHORTCUTS ===
    [Win+Alt+Shift+A] > Show app-specific shortcuts (quick press)
    [Win+Alt+Shift+A] > Show global shortcuts (hold 700ms+)
    [Win+Alt+Shift+/] > Search all cheat sheets (cross-context)
    
    === ZMK KEYBOARD (eyelash_sofle.keymap) ===
    Source: eyelash_sofle.keymap — Sofle split; layers 0–4
    Legend: L0=base, L1=hold left thumb, L2=hold right thumb, L3=sticky (from L2·W), L4=auto when L2+L3
    Tap-dance: 1× / 2× / 3× = single/double/triple tap within tapping-term
    
    --- ZMK Layer 0 (base) ---
    [ZMK L0 · ↑] > Up Arrow
    [ZMK L0 · ↓] > Down Arrow
    [ZMK L0 · ←] > Left Arrow
    [ZMK L0 · →] > Right Arrow
    [ZMK L0 · P] hold > Alt+Shift+S
    [ZMK L0 · P] tap 1× > Alt+Shift+Q — jump mouse to middle
    [ZMK L0 · P] tap 2× > Ctrl+Alt+Win+Y — PowerToys Command Palette bookmarks (show)
    [ZMK L0 · P] tap 3× > Ctrl+Alt+Win+M — Command Palette bookmark
    [ZMK L0 · L] hold > Ctrl+Alt+Win+L — {AI_PROVIDER} D2C direct submit (Utils.ahk)
    [ZMK L0 · L] tap 1× > Win+Alt+Shift+0 — start/stop dictation
    [ZMK L0 · L] tap 2× > Ctrl+Alt+Win+4 — mark last clip favorite
    [ZMK L0 · L] tap 3× > Ctrl+Alt+Win+7 — toggle {AI_PROVIDER} tab 1 <-> 2
    [ZMK L0 · ;] hold > Ctrl+Shift+V — paste plain text
    [ZMK L0 · ;] tap 1× > Win+Alt+Shift+1 — Clip Angel top item
    [ZMK L0 · ;] tap 2× > Ctrl+Alt+B
    [ZMK L0 · ;] tap 3× > Ctrl+Alt+V
    [ZMK L0 · .] hold > Ctrl+Alt+Win+V — maximize active window (WindowManagement.ahk)
    [ZMK L0 · .] tap 1× > Win+Alt+Shift+6 — minimize windows
    [ZMK L0 · .] tap 2× > Alt+F4 — close window
    [ZMK L0 · Left thumb] hold > Layer 1 (ONE)
    [ZMK L0 · Left thumb] tap 1× > Ctrl+Alt+Win+0 — Project Quick Selector
    [ZMK L0 · Right thumb] hold > Layer 2 (TWO)
    [ZMK L0 · Right thumb] tap 1× > Win+Alt+Shift+U — quick string shortcuts
    [ZMK L0 · Right thumb] tap 2× > Win+Alt+Shift+Y — Focus Mode toggle
    [ZMK L0 · Win+Alt+Shift key] > Win+Alt+Shift (modifier chord)
    [ZMK L0 · X thumb] hold > Win+Alt+Shift+X — Import Management
    [ZMK L0 · X thumb] tap 1× > Win+Alt+Shift+I — open {AI_PROVIDER}
    [ZMK L0 · X thumb] tap 2× > Ctrl+Alt+Win+6 — window tools [3]
    [ZMK L0 · X thumb] tap 3× > Ctrl+Alt+Win+U — window tools [2]
    [ZMK L0 · E thumb] hold > Win+Shift+E
    [ZMK L0 · E thumb] tap 1× > Context menu
    [ZMK L0 · E thumb] tap 2× > Ctrl+Alt+Win+B — Handy Parakeet Unified English
    [ZMK L0 · E thumb] tap 3× > Ctrl+Alt+Win+9 — Handy Nemotron Streaming Portuguese
    
    --- ZMK Layer 1 (ONE) — hold left thumb ---
    [ZMK L1 · Esc] > F11
    [ZMK L1 · 2] > F2
    [ZMK L1 · 5] > F5
    [ZMK L1 · ↑] default > mouse move up (slow)
    [ZMK L1 · ↑] +Ctrl > quick jump up (~32767 px)
    [ZMK L1 · ↑] +Shift > Win+Alt+Shift+↑ — square selector up
    [ZMK L1 · ↓] default > mouse move down (slow)
    [ZMK L1 · ↓] +Ctrl > quick jump down (~32767 px)
    [ZMK L1 · ↓] +Shift > Win+Alt+Shift+↓ — square selector down
    [ZMK L1 · ←] default > mouse move left (slow)
    [ZMK L1 · ←] +Ctrl > quick jump left (~32767 px)
    [ZMK L1 · ←] +Shift > Win+Alt+Shift+← — square selector left
    [ZMK L1 · →] default > mouse move right (slow)
    [ZMK L1 · →] +Ctrl > quick jump right (~32767 px)
    [ZMK L1 · →] +Shift > Win+Alt+Shift+→ — square selector right
    [ZMK L1 · 6] > F6
    [ZMK L1 · 7] > Ctrl+A then Ctrl+C (select all + copy)
    [ZMK L1 · 8] > Ctrl+Alt+Win+Z — window tools [1]
    [ZMK L1 · 9] > Ctrl+Alt+Win+P — exit F11 fullscreen
    [ZMK L1 · 0] > F10
    [ZMK L1 · Bksp] > F12
    [ZMK L1 · Q] > Ctrl+Alt+Win+Q then Ctrl+Alt+Win+X — snap 50/50 (monitor 1 pair)
    [ZMK L1 · W] > Ctrl+Alt+Win+W then Ctrl+Alt+Win+X — snap 50/50 (monitor 2 pair)
    [ZMK L1 · E] > Ctrl+Alt+Win+E then Ctrl+Alt+Win+X — snap 50/50 (monitor 3 pair)
    [ZMK L1 · R] > Ctrl+Alt+Win+R then Ctrl+Alt+Win+X — snap 50/50 (monitor 4 pair)
    [ZMK L1 · A] > Ctrl+Alt+Win+A — move window to monitor 1
    [ZMK L1 · S] > Ctrl+Alt+Win+S — move window to monitor 2
    [ZMK L1 · D] > Ctrl+Alt+Win+D — move window to monitor 3
    [ZMK L1 · F] > Ctrl+Alt+Win+F — move window to monitor 4
    [ZMK L1 · T] > Ctrl+Alt+W — cycle windows monitor 2
    [ZMK L1 · ]] > ]
    [ZMK L1 · \] > \
    [ZMK L1 · /] > /
    [ZMK L1 · ;] > ;
    [ZMK L1 · -] > -
    [ZMK L1 · =] > =
    [ZMK L1 · `] > `
    [ZMK L1 · [] > [
    [ZMK L1 · '] > '
    [ZMK L1 · N] > mouse left click
    [ZMK L1 · Shift+'] > Shift+'
    [ZMK L1 · Shift+[] > Shift+[
    [ZMK L1 · Non-US \] > Non-US backslash
    [ZMK L1 · encoder] > Ctrl+Shift+= / Ctrl+- (zoom-style inc/dec)
    
    --- ZMK Layer 2 (TWO) — hold right thumb ---
    [ZMK L2 · Esc] > Ctrl+Alt+Win+2 — Quick Update to Your Scripts
    [ZMK L2 · 1] > Bluetooth select profile 0
    [ZMK L2 · 2] > Bluetooth select profile 3
    [ZMK L2 · 3] > Bluetooth select profile 2
    [ZMK L2 · 5] > Ctrl+Alt+Win+X — snap 50/50
    [ZMK L2 · ↑] > 5× Up Arrow
    [ZMK L2 · ↑] +Ctrl > Page Down
    [ZMK L2 · ↓] > 5× Down Arrow
    [ZMK L2 · ↓] +Ctrl > Page Up
    [ZMK L2 · ←] > 5× Left Arrow
    [ZMK L2 · →] > 5× Right Arrow
    [ZMK L2 · Home slot] hold > Ctrl+Shift+Home then Delete
    [ZMK L2 · Home slot] tap > Shift+Home then Delete
    [ZMK L2 · End slot] hold > Ctrl+Shift+End then Delete
    [ZMK L2 · End slot] tap > Shift+End then Delete
    [ZMK L2 · Home] hold > Ctrl+Home
    [ZMK L2 · Home] tap > Home
    [ZMK L2 · End] hold > Ctrl+End
    [ZMK L2 · End] tap > End
    [ZMK L2 · Tab] > Ctrl+Alt+Win+8 — Moves Desktop to Recycle Bin
    [ZMK L2 · W] > Sticky Layer 3 (RAPID-AIB)
    [ZMK L2 · T] > Win+Ctrl+.
    [ZMK L2 · Y] > Alt+PrtSc
    [ZMK L2 · U] > Up Arrow
    [ZMK L2 · I] > Page Up
    [ZMK L2 · O] > Page Down
    [ZMK L2 · H] > Left Arrow
    [ZMK L2 · J] > Down Arrow
    [ZMK L2 · K] > Right Arrow
    [ZMK L2 · PgUp slot] hold > Ctrl+Shift+Page Up
    [ZMK L2 · PgUp slot] tap > Ctrl+Page Up
    [ZMK L2 · PgDn slot] hold > Ctrl+Shift+Page Down
    [ZMK L2 · PgDn slot] tap > Ctrl+Page Down
    [ZMK L2 · A] > Ctrl+Alt+Win+5 — Clean the Clipboard
    [ZMK L2 · S] > Ctrl+Alt+Win+3 — Toggle Outlook and Teams
    [ZMK L2 · D] > Ctrl+Shift+V — paste plain text
    [ZMK L2 · F] > Ctrl+Alt+C
    [ZMK L2 · G] > Ctrl+Win+Alt+C
    [ZMK L2 · Z] > Win+Alt+Shift+4 — toggle camera (Teams meeting)
    [ZMK L2 · /] > Win+Alt+Shift+5 — toggle mute (Teams meeting)
    [ZMK L2 · Bottom-left] > Bluetooth clear all
    [ZMK L2 · ,] > Shift+,
    [ZMK L2 · .] > Shift+.
    [ZMK L2 · Bottom-right] > Ctrl+A then Delete (select all + delete)
    [ZMK L2 · encoder] > Page Down / Page Up
    
    --- ZMK Layer 3 (RAPID-AIB) — sticky from L2·W ---
    [ZMK L3 · ↑] > 5× Up Arrow
    [ZMK L3 · ↓] > 5× Down Arrow
    [ZMK L3 · ←] > 5× Left Arrow
    [ZMK L3 · →] > 5× Right Arrow
    
    --- ZMK Layer 4 — auto when L2+L3 both active ---
    [ZMK L4 · U] > 5× Up Arrow (overrides L2 single Up on this key)
    [ZMK L4 · H] > 5× Left Arrow
    [ZMK L4 · J] > 5× Down Arrow
    [ZMK L4 · K] > 5× Right Arrow
    
    === WIKIPEDIA ===
    [Win+Alt+Shift+K] > Opens or activates Wikipedia
)"