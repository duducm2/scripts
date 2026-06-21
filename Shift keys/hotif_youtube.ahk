; =============================================================================
; Shift keys module: hotif_youtube.ahk
; YouTube Chrome hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

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
