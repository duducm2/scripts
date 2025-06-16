#Requires AutoHotkey v2.0+
#SingleInstance Force

; Win+Alt+Shift+M → jump straight to a chat in Microsoft Teams
#!+m::
{
    ; Ask for the contact’s name -------------------------------------------
    contact := Trim(InputBox("Enter a Teams contact name:", "Jump to Chat").Value)
    if contact = ""
        return

    ; Speed tweaks ----------------------------------------------------------
    SetWinDelay 0
    SetKeyDelay 0, 0
    SetControlDelay 0

    ; Locate or launch Microsoft Teams -------------------------------------
    teamsWindow := "Microsoft Teams"
    if !WinExist("ahk_exe ms-teams.exe") && !WinExist("ahk_exe Teams.exe")
    {
        Run "ms-teams:"
        WinWait(teamsWindow, , 15)
    }
    WinActivate(teamsWindow)
    WinWaitActive(teamsWindow, , 5)

    ; Open the Go‑to box (Ctrl+G)
    Send "^g"
    Sleep 100  ; allow the box to appear & highlight

    ; Type the contact name directly (avoids clipboard issues) -------------
    SendText(contact)
    Sleep 1000
    Send "{Enter}"
    Sleep 300  ; wait for chat to load

    ; Place cursor in the compose box --------------------------------------
    Send "^r"
}
