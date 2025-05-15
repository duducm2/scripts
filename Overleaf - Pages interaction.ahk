#Requires AutoHotkey v2

#include C:\Users\eduev\Documents\UIA-v2\Lib\UIA.ahk
#include C:\Users\eduev\Documents\UIA-v2\Lib\UIA_Browser.ahk

!+d::
{
    ; Activate Overleaf window
    SetTitleMatchMode 2
    WinActivate "[PDF] Online LaTeX Editor Overleaf"
    WinWaitActive "ahk_exe chrome.exe"

    ; Initialize UIA Browser
    cUIA := UIA_Browser()

    ; Wait a brief moment for the UI to load
    Sleep 300

    ; Check for modifier keys to determine action
    if GetKeyState("Left", "P") {
        ; Navigate backward multiple pages if Left is pressed with Ctrl+Alt+D
        i := 0
        while GetKeyState("Left", "P") && i < 30 {  ; Safety limit of 30 pages
            try {
                ; Based on UIATreeInspector, look for the "Previous page" button
                prevBtn := cUIA.FindElement({ Name: "Previous page", Type: "Button", matchmode: "Substring" })
                if prevBtn
                    prevBtn.Click()
            } catch {
                break  ; Exit loop if error
            }
            Sleep 800  ; Longer delay to allow user to see the page
            i++
        }
    }
    else if GetKeyState("Right", "P") {
        ; Navigate forward multiple pages if Right is pressed with Ctrl+Alt+D
        i := 0
        while GetKeyState("Right", "P") && i < 30 {  ; Safety limit of 30 pages
            try {
                ; Based on UIATreeInspector, look for the "Next page" button
                nextBtn := cUIA.FindElement({ Name: "Next page", Type: "Button", matchmode: "Substring" })
                if nextBtn
                    nextBtn.Click()
            } catch {
                break  ; Exit loop if error
            }
            Sleep 800  ; Longer delay to allow user to see the page
            i++
        }
    }
    else {
        ; If no arrow keys, use the navigation button + Tab approach to select the page number
        try {
            ; First find and click either the Next or Previous page button
            navBtn := cUIA.FindElement({ Name: "Next page", Type: "Button", matchmode: "Substring" })
            if navBtn {
                ; navBtn.Click()
                ; Then press Tab to move to the page number field
                Sleep 100
                Send "{Tab}"
                Sleep 100
                Send "^a"  ; Select all text in the field
            } else {
                ; Try previous page button if next page button not found
                navBtn := cUIA.FindElement({ Name: "Previous page", Type: "Button", matchmode: "Substring" })
                if navBtn {
                    ; navBtn.Click()
                    ; Then press Tab to move to the page number field
                    Sleep 100
                    Send "{Tab}"
                    Sleep 100
                    Send "^a"  ; Select all text in the field
                } else {
                    ; If neither button found, try clicking anywhere in the toolbar area and Tab
                    Send "{Tab}"
                    Sleep 100
                    Send "^a"
                }
            }
        } catch {
            ; If that fails, try a direct keyboard approach
            Send "{Tab}"
            Sleep 100
            Send "^a"
        }
    }
}
