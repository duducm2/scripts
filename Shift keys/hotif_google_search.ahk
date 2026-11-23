; =============================================================================
; Shift keys module: hotif_google_search.ahk
; Google Search Chrome hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

#HotIf WinActive("ahk_exe chrome.exe") && IsGoogleSearchTitle(SafeWinGetTitle())

; Shift + S : Focus Google search box
+s:: {
    try {
        uia := UIA_Browser()
        Sleep 300

        ; Find the Google search box by Name
        searchBox := UiElements_TrySaved(uia, "Google", "Search")
        if !searchBox
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

; Shift + U : Select first search result (shared with D2C [C] Chrome navigate)
+u:: {
    if (GoogleSearch_ClickFirstResult())
        return
    ToolTip("First result not found")
    SetTimer(() => ToolTip(), -2000)
}

; Shift + I : Click Images / Imagens tab (keeps current search query)
+i:: {
    try {
        uia := UIA_Browser()
        Sleep 200

        root := uia
        try {
            if (doc := uia.GetCurrentDocumentElement())
                root := doc
        } catch {
        }

        imagesBtn := UiElements_TrySaved(root, "Google", "Images")
        if (!imagesBtn && root != uia)
            imagesBtn := UiElements_TrySaved(uia, "Google", "Images")

        if !imagesBtn {
            for label in ["Images", "Imagens"] {
                for ctrlType in ["Link", "TabItem", "Button"] {
                    try imagesBtn := root.FindFirst({ Name: label, Type: ctrlType })
                    catch
                        imagesBtn := 0
                    if (imagesBtn)
                        break 2
                }
            }
        }
        if !imagesBtn {
            for label in ["Images", "Imagens"] {
                try imagesBtn := root.FindFirst({ Name: label })
                catch
                    imagesBtn := 0
                if (imagesBtn) {
                    ; If we matched a Text/Group inside the tab link, walk up to the clickable ancestor.
                    cur := imagesBtn
                    Loop 4 {
                        try {
                            t := cur.Type
                            if (t = 50005 || t = 50019 || t = 50000) { ; Link, TabItem, Button
                                imagesBtn := cur
                                break
                            }
                            cur := cur.WalkTree("p")
                            if !cur
                                break
                        } catch {
                            break
                        }
                    }
                    break
                }
            }
        }

        if (imagesBtn) {
            try {
                imagesBtn.Invoke()
            } catch {
                try imagesBtn.Click()
                catch {
                    imagesBtn.Select()
                }
            }
            return
        }
        ToolTip("Images button not found")
        SetTimer(() => ToolTip(), -2000)
    } catch {
        ToolTip("Images button not found")
        SetTimer(() => ToolTip(), -2000)
    }
}

#HotIf