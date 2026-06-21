; =============================================================================
; Shift keys module: hotif_google_search.ahk
; Google Search Chrome hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

#HotIf WinActive("ahk_exe chrome.exe") && InStr(WinGetTitle("A"), "Google") && !InStr(WinGetTitle("A"),
"Google Maps")

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
