; =============================================================================
; Shift keys module: hotif_chrome_general.ahk
; Google Chrome general hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

;-------------------------------------------------------------------
; Google Chrome Shortcuts
;-------------------------------------------------------------------
#HotIf WinActive("ahk_exe chrome.exe")

; Shift + W : Pop current tab to new window - Window
+w:: {
    if !Chrome_DetachActiveTabToNewWindow() {
        if (CHROME_DETACH_LEGACY_KEYS)
            Chrome_DetachActiveTabToNewWindow_Legacy()
    }
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
