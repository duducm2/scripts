# User Input Bottlenecks Report

**Generated:** Scripts Codebase Analysis  
**Purpose:** Identify blocking elements (message boxes, input lists, wait commands) that pause execution until user interaction, for future speed optimization.

---

## 1. MsgBox (Modal — Blocks Until Click)

Message boxes halt script execution until the user clicks a button.

### Utils.ahk
| Line | Context |
|------|---------|
| 615 | `MsgBox` — "Other scripts need updates before Gemini" (YesNo/Update choice) |
| 633 | `MsgBox` — "Gemini.ahk not found" (Update Failed) |
| 639 | `MsgBox` — "Failed to update Gemini script" |
| 652 | `MsgBox` — "Failed to launch Handy" (Shortcut not found) |
| 657 | `MsgBox` — "Failed to launch Handy" |
| 732 | `MsgBox` — "Error in AddWordToHandy macro" |
| 812 | `MsgBox` — "ClipAngel window did not appear" |
| 824 | `MsgBox` — "Failed to initialize UIA for ClipAngel" |
| 834 | `MsgBox` — "Could not find DataGridView in ClipAngel" |
| 844 | `MsgBox` — "No clips found in Row 0" |
| 854 | `MsgBox` — "Could not find Title element in Row 0" |
| 864 | `MsgBox` — "Title Row 0 contains no text data" |
| 883 | `MsgBox` — "Failed to re-initialize UIA after switching views" |
| 893 | `MsgBox` — "Could not find DataGridView in All Clips view" |
| 980 | `MsgBox` — "Error in MergeNonFavoriteClips" |
| 1865 | `MsgBox` — "Error closing Outlook" |
| 1882 | `MsgBox` — "Error closing Teams windows" |
| 1912 | `MsgBox` — "Error launching Outlook" |
| 1966 | `MsgBox` — "Error in ToggleOutlookAndTeams macro" |
| **2014** | **`MsgBox` — "Open Applications?" (YesNo — user choice)** |
| 2068 | `MsgBox` — "Error opening applications" |
| **2593** | **`MsgBox` — "Clean the Clipboard" (YesNo — user confirmation)** |
| **2626** | **`MsgBox` — "Would you like to clean up the clipboard?" (YesNo)** |
| 3262 | `MsgBox` — "GetWindowRect failed" |
| 4544 | `MsgBox` — (commented out debug) |

### Act.ahk
| Line | Context |
|------|---------|
| **8** | **`MsgBox` — "Can we proceed with Act?" (YesNo — user confirmation)** |

### Shift keys.ahk
| Line | Context |
|------|---------|
| 1473 | `MsgBox` — "Error" (ShowErr helper) |
| 2237 | `MsgBox` — "Could not find the 'Unread' filter button" |
| 2241 | `MsgBox` — "An error occurred" (WhatsApp) |
| 2267 | `MsgBox` — "Could not find the 'Archived' button" |
| 2271 | `MsgBox` — "Error focusing WhatsApp conversation" |
| 2288 | `MsgBox` — "Can't attach to Chrome" |
| 2339 | `MsgBox` — "Couldn't restart recording (Voice-message button missing)" |
| 2346 | `MsgBox` — "Couldn't find the Voice-message button" |
| 2349 | `MsgBox` — "Error" (Voice message) |
| **2587** | **`MsgBox` — "Snooze for X?" (YesNo — user confirmation)** |
| 2725 | `MsgBox` — "Couldn't find the Chat button" |
| 2728 | `MsgBox` — "UIA error" |
| 2817 | `MsgBox` — "Couldn't find the Reagir button" |
| 2820 | `MsgBox` — "UIA error" |
| 3145 | `MsgBox` — "Could not find the 'Search Wikipedia' field" |
| 3147 | `MsgBox` — "An error occurred" |
| 3895 | `MsgBox` — "Could not find Mercado Livre search field" |
| 3898 | `MsgBox` — "An error occurred" |
| 3942 | `MsgBox` — "Could not find Mercado Livre cart link" |
| 3944 | `MsgBox` — "An error occurred" |
| 3981 | `MsgBox` — "Could not find Mercado Livre purchases link" |
| 3983 | `MsgBox` — "An error occurred" |
| **4130** | **`MsgBox` — "Do you want to call this person?" (YesNo)** |
| 4363 | `MsgBox` — "Could not find the 'Mentions' element" |
| 4385 | `MsgBox` — "Could not find the 'Chat (Ctrl+1)' button" |
| 4389 | `MsgBox` — "Error in Shift+O" |
| 4615 | `MsgBox` — "Couldn't find [Outlook button]" |
| 4879 | `MsgBox` — "Could not find Mail or Calendar items" |
| 4882 | `MsgBox` — "Error toggling Mail/Calendar" |
| 5092 | `MsgBox` — "Couldn't find the All day checkbox" |
| 5156 | `MsgBox` — "Couldn't find the Make Recurring button" |
| 5435 | `MsgBox` — "Error in selection dialog" (Outlook option) |
| 5812 | `MsgBox` — "Outlook appointment window not found" |
| 5889 | `MsgBox` — "Outlook appointment window not found" |
| 6120 | `MsgBox` — "Failed to get root element" (ChatGPT) |
| 6201 | `MsgBox` — "Failed to find chat button" |
| 6209 | `MsgBox` — "Failed to find sibling element of chat button" |
| 6290 | `MsgBox` — "Failed to find OpenConversationOptions button" |
| 6300 | `MsgBox` — "OpenConversationOptions button is offscreen" |
| 6305 | `MsgBox` — "OpenConversationOptions button is disabled" |
| 6370 | `MsgBox` — "Failed to click OpenConversationOptions button" |
| 6582 | `MsgBox` — "Input volume slider not found" |
| 6587 | `MsgBox` — "Error setting input volume" |
| **6847** | **`MsgBox` — "If 'semicolon' is not selected, hit yes" (YesNo)** |
| 6869 | `MsgBox` — "Couldn't find the Enable Editing button" |
| 6872 | `MsgBox` — "Error" |
| 6998 | `MsgBox` — "Could not find the 'Home' tab" (Power BI) |
| 7033 | `MsgBox` — "Could not find the 'Transform data' menu item" |
| 7036 | `MsgBox` — "Error triggering Transform data" |
| 7071 | `MsgBox` — "Could not find the 'Report view' tab" |
| 7074 | `MsgBox` — "Error switching to Report view" |
| 7093 | `MsgBox` — "Could not find the 'Table view' tab" |
| 7096 | `MsgBox` — "Error switching to Table view" |
| 7115 | `MsgBox` — "Could not find the 'Model view' tab" |
| 7118 | `MsgBox` — "Error switching to Model view" |
| 7183 | `MsgBox` — "Could not find the 'Build visual' tab" |
| 7186 | `MsgBox` — "Error switching to Build visual" |
| 7256 | `MsgBox` — "Could not find the 'Format visual' tab" |
| 7259 | `MsgBox` — "Error switching to Format visual" |
| 7284 | `MsgBox` — "Could not locate the Data button anchor" |
| 7301 | `MsgBox` — "Could not focus the Data button anchor" |
| 7310 | `MsgBox` — "Error selecting the Power BI search field" |
| 7598 | `MsgBox` — "Could not find the 'Home' tab" |
| 7617 | `MsgBox` — "Could not find the 'New measure' button" |
| 7620 | `MsgBox` — "Error triggering New measure" |
| 7643 | `MsgBox` — "Could not find the 'Home' tab" |
| 7659 | `MsgBox` — "Could not find the 'Refresh' button" |
| 7662 | `MsgBox` — "Error triggering Refresh" |
| 7685 | `MsgBox` — "Could not find the 'Format' tab" |
| 7706 | `MsgBox` — "Could not find the 'Bring forward' button" |
| 7724 | `MsgBox` — "Error triggering Bring forward" |
| 7747 | `MsgBox` — "Could not find the 'Format' tab" |
| 7768 | `MsgBox` — "Could not find the 'Send backward' button" |
| 7786 | `MsgBox` — "Error triggering Send backward" |
| 7809 | `MsgBox` — "Could not find the 'Format' tab" |
| 7836 | `MsgBox` — "Could not find the 'Align' button" |
| 7843 | `MsgBox` — "Error triggering Align" |
| 7876 | `MsgBox` — "Could not find the 'Fit to page' button" |
| 7883 | `MsgBox` — "Error triggering Fit to page" |
| 7945 | `MsgBox` — "Could not find the 'Format painter' button" |
| 7952 | `MsgBox` — "Error triggering Format painter" |
| 7975 | `MsgBox` — "Could not find the 'Format' tab" |
| 8032 | `MsgBox` — "Could not find the 'Group' button" |
| 8039 | `MsgBox` — "Error triggering Group" |
| 8086 | `MsgBox` — "Error closing Power BI drawers" |
| 8132 | `MsgBox` — "Error opening Power BI drawers" |
| 8158 | `MsgBox` — "Could not find any Power BI tables to collapse" |
| 8197 | `MsgBox` — "Error collapsing Power BI tables" |
| 8372 | `MsgBox` — "Could not find the 'Updates' button" |
| 8376 | `MsgBox` — "An error occurred" |
| 8397 | `MsgBox` — "Could not find the 'Forums' button" |
| 8401 | `MsgBox` — "An error occurred" |
| 8425 | `MsgBox` — "Could not find Mark as read/unread button" |
| 8429 | `MsgBox` — "An error occurred" |
| 8499 | `MsgBox` — "Could not find the 'Inbox' button" |
| 8503 | `MsgBox` — "An error occurred" |
| 8924 | `MsgBox` — "Invalid selection" (Commit Push Selector) |
| 8962 | `MsgBox` — "Error in commit push selector" |
| 9017 | `MsgBox` — "Invalid selection" (Emoji Selector) |
| 9057 | `MsgBox` — "Error in emoji selector" |
| 9085 | `MsgBox` — "Invalid selection" (AI Model Selection) |
| 9156 | `MsgBox` — "Error in AI model selection" |
| 9269 | `MsgBox` — "Invalid selection" (Commit Selector) |
| 9563 | `MsgBox` — "UIA error folding Git directories" |
| 9766 | `MsgBox` — "UIA error folding Explorer directories" |
| 9969 | `MsgBox` — "UIA error unfolding Explorer directories" |
| 10024 | `MsgBox` — "Invalid selection" (AI Mode Selection) |
| 10049 | `MsgBox` — "Error switching AI mode" |
| 10117 | `MsgBox` — "Invalid selection" (AI Model Selection) |
| 10125 | `MsgBox` — "Error switching AI model" |
| 10222 | `MsgBox` — "Couldn't find the Connect-to-device button" |
| 10230 | `MsgBox` — "Error" |
| 10355 | `MsgBox` — "Could not find 'Expand Your Library' button" |
| 10358 | `MsgBox` — "Could not find 'Expand Your Library' button" |
| 10361 | `MsgBox` — "Error toggling fullscreen library" |
| 10678 | `MsgBox` — "Could not find Dashboard button" (Mobills) |
| 10681 | `MsgBox` — "Error navigating to Dashboard" |
| 10692 | `MsgBox` — "Could not find Contas/Accounts button" |
| 10695 | `MsgBox` — "Error navigating to Contas/Accounts" |
| 10706 | `MsgBox` — "Could not find Transações/Transactions button" |
| 10709 | `MsgBox` — "Error navigating to Transações/Transactions" |
| 10720 | `MsgBox` — "Could not find Cartões de crédito button" |
| 10723 | `MsgBox` — "Error navigating to Cartões de crédito" |
| 10734 | `MsgBox` — "Could not find Planejamento/Budgets button" |
| 10737 | `MsgBox` — "Error navigating to Planejamento/Budgets" |
| 10748 | `MsgBox` — "Could not find Relatórios/Reports button" |
| 10751 | `MsgBox` — "Error navigating to Relatórios/Reports" |
| 10762 | `MsgBox` — "Could not find Mais opções button" |
| 10765 | `MsgBox` — "Error navigating to Mais opções" |
| 11174 | `MsgBox` — "Could not attach to browser window" |
| 11200 | `MsgBox` — "Pager control could not be clicked" |
| 11204 | `MsgBox` — "Could not find prev/next" |
| 11207 | `MsgBox` — "Error navigating Mobills" |
| 11253 | `MsgBox` — "Could not attach to browser window" |
| 11270 | `MsgBox` — "Could not find anchor element" |
| 11343 | `MsgBox` — "Could not find second Ignore transaction toggle" |
| 11347 | `MsgBox` — "Mobills Error" |
| 11398 | `MsgBox` — "Could not find Description field" |
| 11413 | `MsgBox` — "Error focusing Description field" |
| 11426 | `MsgBox` — "Could not attach to browser window" |

---

## 2. InputBox (Modal — Blocks Until Submit/Cancel)

### Shift keys.ahk
| Line | Context |
|------|---------|
| 5387 | `Outlook_SelectOptionByInputBox` — Used for Outlook appointment options (Privacy, All-day, Status, etc.) — **blocks until user types number** |
| 10011 | `InputBox` — "Choose AI Mode: 1. ask / 2. agent" — **blocks until user submits** |
| 10057 | `InputBox` — "Choose AI Model: 1–6" — **blocks until user submits** |
| 11602 | `InputBox` — "Enter text to search for in your notes" (Google Keep Search) — **blocks until user submits** |

### Microsoft Teams.ahk (if in scope)
| Line | Context |
|------|---------|
| 659 | `InputBox` — "Enter a Teams contact name" (Jump to Chat) — **blocks until user submits** |

---

## 3. KeyWait (Blocks Until Key State Change or Timeout)

### Utils.ahk
| Line | Context |
|------|---------|
| **5082** | `KeyWait("x", "T0.4")` — Distinguishes short vs long press for Hunt & Peck; **blocks up to 400 ms** |
| **5090** | `KeyWait("x")` — On long press, waits for key release; **no timeout — can block indefinitely** |
| **7640** | `KeyWait("0", "L")` — Dictation hotkey; **waits for key release, no timeout** |

### Shift keys.ahk
| Line | Context |
|------|---------|
| **1437** | `KeyWait "a", "T1"` — Win+Alt+Shift+A (Send top list item); **blocks up to 1 s** |

### AppLaunchers.ahk
| Line | Context |
|------|---------|
| **1931** | `KeyWait("9", "T1")` — Pomodoro timer; **blocks up to 1 s** to distinguish short vs long press |

---

## 4. WinWait / WinWaitActive / WinWaitClose

### Utils.ahk
| Line | Context |
|------|---------|
| 656 | `WinWait("Handy ahk_class Tauri Window", , 5)` |
| 662 | `WinWaitActive("Handy ahk_class Tauri Window", , 2)` |
| 816 | `WinWaitActive("ClipAngel", , 2)` |
| 1012 | `WinWaitActive("ClipAngel", , 2)` |
| 1017 | `WinWait("ClipAngel", , 10)` |
| 1022 | `WinWaitActive("ClipAngel", , 2)` |
| 1473 | `WinWaitActive("ahk_id " . matchingHwnd, , 2)` |
| 1482 | `WinWait("Handy ahk_class Tauri Window", , 5)` |
| 1491 | `WinWaitActive("ahk_id " . h, , 2)` |
| 1497 | `WinWaitActive("ahk_id " . h, , 2)` |
| 1939 | `WinWaitActive("ahk_exe ms-teams.exe", , 10)` |
| 1952 | `WinWait("ahk_exe OUTLOOK.EXE", , 5)` |
| 1956 | `WinWaitActive("ahk_exe OUTLOOK.EXE", , 2)` |
| 2935 | `WinWaitActive("ahk_id " . targetHwnd, , 1)` |
| 4399 | `WinWaitActive("ahk_id " . targetHwnd, , 0.35)` |
| 4908 | `WinWaitActive("ahk_id " g_HnPTargetWindow, "", 0.2)` |
| 5403 | `WinWaitActive("ahk_id " hwnd, , 2)` |
| 5450 | `WinWaitActive("ahk_id " hwnd, , 2)` |
| 5479 | `WinWait("ahk_exe chrome.exe", , 10)` |
| 5500 | `WinWaitActive("ahk_id " hwnd, , 2)` |
| 5531 | `WinWaitActive("ahk_id " hwnd, , 2)` |
| 5667 | `WinWaitActive("ahk_exe chrome.exe", , 5)` |
| 5675 | `WinWaitActive("ahk_id " geminiHwnd, , 2)` |
| 5679 | `WinWaitActive("ahk_exe chrome.exe", , 2)` |
| 5897 | `WinWaitActive("ahk_id " geminiHwnd, , 2)` |
| 5901 | `WinWaitActive("ahk_exe chrome.exe", , 2)` |

### Shift keys.ahk
| Line | Context |
|------|---------|
| 1794 | `WinWaitActive("ahk_id " win, , 1)` |
| 2567 | `WinWaitActive("ahk_exe OUTLOOK.EXE", , 1)` |
| **5427** | **`WinWaitClose("ahk_id " optionGui.Hwnd)`** — Outlook option dialog; **blocks until user closes GUI** |
| 5822 | `WinWaitActive("ahk_id " targetHwnd, , 2)` |
| 5899 | `WinWaitActive("ahk_id " targetHwnd, , 2)` |
| 6349 | `WinWaitActive("ahk_id " chatGPTHwnd, , 1)` |
| 8775 | `WinWaitActive("ahk_id " hwnd, , 3)` |
| 8880 | `WinWaitActive("ahk_id " gCommitPushTargetWin, , 2)` |
| 11607 | `WinWaitActive("ahk_id " currentWindow, , 2)` |
| 11650 | `WinWaitActive("ahk_id " currentWindow, , 2)` |
| 11745 | `WinWaitActive("ahk_id " bestHwnd, , 1)` |
| 12494 | `WinWaitActive("ahk_id " browserHwnd, , 1)` |
| 12678 | `WinWaitActive("ahk_id " geminiHwnd, , 2)` |
| 12684 | `WinWaitActive("ahk_exe chrome.exe", , 2)` |
| **14072** | **`WinWaitClose("ahk_id " searchGui.Hwnd)`** — Tree item search; **blocks until user closes GUI** |

### AppLaunchers.ahk
| Line | Context |
|------|---------|
| 87 | `WinWaitActive(targetWindow, , 2)` |
| 95 | `WinWaitActive(fallbackWindow, , 2)` |
| 189 | `WinWaitActive("ahk_id " targetHwnd, , 0.2)` |
| 210 | `WinWait("ahk_exe chrome.exe", , 10)` |
| 213 | `WinWaitActive("ahk_exe chrome.exe", , 2)` |
| 253 | `WinWaitActive("WhatsApp")` |
| 271 | `WinWaitActive("YouTube")` |
| 291 | `WinWaitActive("Gmail ahk_exe chrome.exe")` |
| 310 | `WinWaitActive("ahk_exe Cursor.exe")` |
| 749 | `WinWait("ahk_exe chrome.exe", , 5)` |
| 754 | `WinWait("Wikipedia", , 10)` |
| 756 | `WinWaitActive("Wikipedia", , 5)` |
| 1302 | `WinWaitActive("Wikipedia", , 2)` |
| 1304 | `WinWaitActive("ahk_exe chrome.exe", , 2)` |

### Gemini.ahk
| Line | Context |
|------|---------|
| 254 | `WinWaitActive("ahk_exe chrome.exe", , 2)` |
| 553 | `WinWaitActive("ahk_exe chrome.exe", , 2)` |
| 667 | `WinWaitActive("ahk_exe chrome.exe", , 5)` |
| 681 | `WinWaitActive("ahk_id " geminiHwnd, , 2)` |
| 744 | `WinWaitActive("ahk_id " hwnd, , 2)` |
| 851 | `WinWaitActive("ahk_exe chrome.exe", , 2)` |
| 886 | `WinWaitActive("ahk_id " origHwnd, , 1)` |
| 963 | `WinWaitActive("ahk_exe chrome.exe", , 2)` |
| 1062 | `WinWaitActive("ahk_exe chrome.exe", , 2)` |
| 1093 | `WinWaitActive("ahk_id " origHwnd, , 1)` |

### WindowManagement.ahk
| Line | Context |
|------|---------|
| 528 | `WinWaitActive("ahk_id " target.hwnd, , 0.3)` |
| 791 | `WinWaitActive("ahk_id " targetHwnd, , 3)` |
| 1011 | `WinWaitActive("ahk_id " targetWindow.hwnd, , 2)` |
| 1103 | `WinWaitActive("ahk_id " targetWindow.hwnd, , 2)` |
| 1191 | `WinWaitActive("ahk_id " targetHwnd, , 2)` |
| 1196 | `WinWaitActive("ahk_id " targetHwnd, , 2)` |
| 1783 | `WinWaitActive("ahk_id " targetWindow.hwnd, , 2)` |
| 1818 | `WinWaitActive("ahk_id " . targetHwnd, , 1)` |
| 1962 | `WinWaitActive("ahk_id " . cursorWindows[1], , 1)` |

### Spotify.ahk
| Line | Context |
|------|---------|
| 29 | `WinWaitActive("ahk_exe Spotify.exe", , 2)` |
| 41 | `WinWaitActive("ahk_exe Spotify.exe", , 5)` |
| 48 | `WinWaitActive("ahk_exe Spotify.exe", , 5)` |
| 54 | `WinWaitActive("ahk_exe Spotify.exe", , 5)` |
| 63 | `WinWaitActive("ahk_exe Spotify.exe", , 2)` |
| 101 | `WinWaitActive("ahk_exe Spotify.exe", , 2)` |
| 157 | `WinWaitActive(win, , 2)` |

### GeminiToCursorBridge.ahk
| Line | Context |
|------|---------|
| 176 | `WinWaitActive("ahk_id " targetWindow.hwnd, , 2)` |
| 189 | `WinWaitActive("ahk_id " targetHwnd, , 2)` |
| 194 | `WinWaitActive("ahk_id " targetHwnd, , 2)` |
| 265 | `WinWaitActive("ahk_id " geminiBrowserHwnd, , 3)` |
| 464 | `WinWaitActive("ahk_id " targetHwnd, , 3)` |
| 475 | `WinWaitActive("ahk_id " targetHwnd, , 2)` |
| 496 | `WinWaitActive("ahk_id " targetHwnd, , 2)` |
| 538 | `WinWaitActive("ahk_id " targetHwnd, , 2)` |
| 551 | `WinWaitActive("ahk_id " targetHwnd, , 2)` |

### Microsoft Teams.ahk
| Line | Context |
|------|---------|
| 36 | `WinWaitActive("ahk_id " hwnd, , waitMs/1000)` |
| 48 | `WinWaitActive("ahk_id " hwnd, , waitMs/1000)` |
| 58 | `WinWaitActive("ahk_id " hwnd, , waitMs/1000)` |
| 69 | `WinWaitActive("ahk_id " hwnd, , waitMs/1000)` |
| 668 | `WinWait(teamsWindow, , 15)` |
| 671 | `WinWaitActive(teamsWindow, , 5)` |

---

## 5. Blocking GUI Dialogs (Show + WinWaitClose or Modal-Like)

These GUIs block the calling thread until the user interacts (closes or selects).

### Shift keys.ahk
| Line | Function/GUI | Behavior |
|------|--------------|----------|
| 5423 | `Outlook_SelectOptionByInputBox` | `optionGui.Show("w500 h250")` + `WinWaitClose` — **blocks until user selects option or cancels** |
| 8958 | `ShowCommitPushSelector` | `commitPushGui.Show("w350 h150")` — **blocks until user selects 1 or 2 (auto-submit or button)** |
| 9053 | Emoji Selector (+O) | `emojiGui.Show("w350 h200")` — **blocks until user selects 1–5** |
| 14068 | Tree item search | `searchGui.Show("w350 h150")` + `WinWaitClose` — **blocks until user submits search or cancels** |

### Utils.ahk
| Line | GUI | Behavior |
|------|-----|----------|
| 6675 | `g_HotstringSelectorGui.Show` | Non-modal (NA); waits via Hotkey callbacks — **user must press a character** |
| 3168 | `g_CursorFocusSelectorGui.Show` | Non-modal (NA); waits via Close button — **user must click Close** |
| 1189–1193 | `g_AiModelSelectorGui` | Non-modal; user presses 1–7 or Esc — **user must press key** |
| 3168 | Cursor focus selector | Same pattern — **user must interact** |

### AppLaunchers.ahk
| Line | GUI | Behavior |
|------|-----|----------|
| 1262 | `g_WikipediaSelectorGui.Show` | Non-modal; waits via Hotkey — **user must press character** |
| 8798–8832 | `ShowCommitPushBanner` | Banner + Hotkey Y; **5 s timeout** — **user must press Y within 5 s** |

---

## 6. RunWait (Blocks Until External Process Exits)

### Utils.ahk
| Line | Context |
|------|---------|
| 466 | `RunWait` — git status |
| 496 | `RunWait` — git fetch |
| 499 | `RunWait` — git status |
| 540 | `RunWait` — git fetch |
| 541 | `RunWait` — git pull |
| 628 | `RunWait` — git fetch |
| 629 | `RunWait` — git pull |
| 3387 | `RunWait` — PowerShell command |
| 7459 | `RunWait` — Set-MicVolume.ps1 |
| 7652 | `RunWait` — Set-MicVolume.ps1 |

### Act.ahk
| Line | Context |
|------|---------|
| 21 | `RunWait` — git fetch |
| 22 | `RunWait` — git pull |
| 35 | `RunWait` — git fetch |
| 36 | `RunWait` — git pull |

---

## 7. BlockInput (Blocks Physical Input — Not User Choice)

Temporarily blocks keyboard/mouse; can be a bottleneck if held too long.

### Shift keys.ahk
| Line | Context |
|------|---------|
| 3175 | `BlockInput("On")` — Wikipedia paste workflow |
| 3183 | `BlockInput("Off")` |
| 3192 | `BlockInput("Off")` |
| 3214 | `BlockInput("Off")` |
| 3225 | `BlockInput("Off")` |

### AppLaunchers.ahk
| Line | Context |
|------|---------|
| 504 | `BlockInput("On")` — Cursor window selection |
| 512 | `BlockInput("Off")` |
| 521 | `BlockInput("Off")` |
| 534 | `BlockInput("Off")` |
| 546 | `BlockInput("Off")` |
| 833 | `BlockInput("On")` |
| 865 | `BlockInput("Off")` |
| 917 | `BlockInput("Off")` |
| 931 | `BlockInput("Off")` |
| 988 | `BlockInput("Off")` |
| 1007 | `BlockInput("Off")` |
| 1020 | `BlockInput("Off")` |
| 1329 | `BlockInput("On")` |
| 1348 | `BlockInput("Off")` |
| 1431 | `BlockInput("Off")` |

---

## 8. High-Impact User Interaction Bottlenecks (Summary)

| Category | Files | Count | Notes |
|----------|-------|-------|-------|
| **MsgBox (YesNo/Choice)** | Act, Utils, Shift keys | ~10 critical | Blocks until user clicks |
| **InputBox** | Shift keys | 4 | Blocks until user types and submits |
| **KeyWait (no timeout)** | Utils | 2 | Can block indefinitely on key hold |
| **KeyWait (with timeout)** | Utils, Shift keys, AppLaunchers | 3 | Blocks 0.4–1 s |
| **WinWaitClose (GUI)** | Shift keys | 2 | Blocks until user closes dialog |
| **Blocking selector GUIs** | Shift keys, Utils | 4+ | Show + wait for input (commit push, emoji, Outlook options, tree search) |
| **RunWait** | Utils, Act | 12 | Blocks until external process exits (git, PowerShell) |

---

## 9. Optimization Suggestions

1. **MsgBox → non-blocking toasts**: Replace informational MsgBoxes with `SetTimer`-based banners (similar to `CreateCenteredBanner`).
2. **InputBox → auto-submit GUIs**: Use the same pattern as Commit Push and Emoji selectors (auto-submit on number/Enter) instead of InputBox.
3. **KeyWait without timeout**: Add reasonable timeouts where possible (e.g. `KeyWait("x", "T5")` for long-press).
4. **Selector GUIs**: Already use auto-submit in several places; extend this pattern to AI Mode/Model selection and Outlook options.
5. **WinWait/WinWaitActive**: Keep timeouts; consider shorter values where UX allows.
6. **RunWait**: Consider async/background execution where user does not need to wait (e.g. `Run` instead of `RunWait` for git fetch).
