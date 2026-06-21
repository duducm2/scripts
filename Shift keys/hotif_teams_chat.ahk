; =============================================================================
; Shift keys module: hotif_teams_chat.ahk
; Teams chat window hotkeys and UIA
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

#HotIf IsTeamsChatActive()

; Shift + K : Toolbar Back (UIA; fallback Alt+Left — see teams.md menur75)
+k::
{
    try {
        root := UIA.ElementFromHandle(WinExist("A"))
        if !root
            return
        backBtn := root.FindFirst({ Name: "Back", Type: "50000", AutomationId: "menur75" })
        if !backBtn
            backBtn := root.FindFirst({ Name: "Back", Type: "50000" })
        if (backBtn) {
            backBtn.Click()
            return
        }
    } catch {
    }
    Send "!{Left}"
}

; Shift + L : Toolbar Forward (UIA; fallback Alt+Right). Like reaction moved to Shift+Y.
+l::
{
    try {
        root := UIA.ElementFromHandle(WinExist("A"))
        if !root
            return
        fwdBtn := root.FindFirst({ Name: "Forward", Type: "50000" })
        if (fwdBtn) {
            fwdBtn.Click()
            return
        }
    } catch {
    }
    Send "!{Right}"
}

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

; =============================================================================
; Chat: Select "Quick Views" tree item in the chat rail
; Hotkey: Shift+O
; UIA path (relative to window): 2,1,2,3,2,1,1,1,1,1,1,1,1,9,1,4,1
; Full tree path:                11,2,1,2,3,2,1,1,1,1,1,1,1,1,9,1,4,1
; Uses SelectionItemPattern.Select() — no mouse click.
; =============================================================================
+o::
{
    StandardLoadingBar_Show("⏳ Teams: Quick views…", BANNER_ACCENT_INTERMEDIATE, { centerOnHwnd: WinExist("A") })
    try {
        ; Create a new chat
        Send "^n"
        Sleep 500

        hwnd := 0
        for proc in TEAMS_PROCESSES {
            for wnd in WinGetList("ahk_exe " proc) {
                if IsTeamsChatTitle(WinGetTitle(wnd)) {
                    hwnd := wnd
                    break 2
                }
            }
        }
        if !hwnd {
            ShowCenteredOverlay_Utils("⚠ NO TEAMS CHAT WINDOW", 2000, BANNER_ACCENT_ERROR)
            return
        }

        root := UIA.ElementFromHandle(hwnd)
        ; Path is relative to the window element (child 11 = Chrome_WidgetWin_1, the Chromium host pane)
        quickViews := root.ElementFromPathExist("11,2,1,2,3,2,1,1,1,1,1,1,1,1,9,1,4,1")
        if !quickViews
            quickViews := root.FindFirst(UIA.CreateCondition({ Name: "Quick views", ControlType: "TreeItem" }))
        if !quickViews {
            ShowCenteredOverlay_Utils("❌ QUICK VIEWS NOT FOUND", 2000, BANNER_ACCENT_ERROR)
            return
        }
        quickViews.SelectionItemPattern.Select()
    } catch as e {
        ShowCenteredOverlay_Utils("❌ QUICK VIEWS: " e.Message, 2000, BANNER_ACCENT_ERROR)
    } finally {
        StandardLoadingBar_Hide(0)
    }
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

; Shift + Y : Like reaction (moved from Shift+L — now Shift+L is Forward)
+y::
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
