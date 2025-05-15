#Requires AutoHotkey v2

#include C:\Users\fie7ca\Documents\UIA-v2\Lib\UIA.ahk

CapsLock & m::
{

    MyGui := Gui(, "Enter a name:")
    MyGui.Add("Text",, "Name:")
    MyGui.Add("Edit", "vName ym")  ; The ym option starts a new column of controls.
    MyGui.Add("Button", "default", "OK").OnEvent("Click", ProcessUserInput)
    MyGui.OnEvent("Close", ProcessUserInput)
    MyGui.Show()

    ProcessUserInput(*)
    {
        Saved := MyGui.Submit()  ; Save the contents of named controls into an object.

        If WinExist("Chat |") || WinExist("Activity |") || WinExist("Teams and Channels |") || WinExist("Calls |") || WinExist("Search |"){

            WinActivate

            Sleep 200

            Send "^e"

            Sleep 300

            Send "^a"
            Send "{Backspace}"

            Send Saved.Name

            Sleep 2000

            Send "{Down}"
            Send "{Down}"
            ; Send "{Down}"
            Sleep 200
            Send "{Enter}"

            Send "^r"
            Sleep 200
            Send "{Tab}"
            Send "+{Tab}"

            Send "^+x"
        }

    }
    Send("{CapsLock}")
}