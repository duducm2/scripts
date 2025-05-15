#Requires AutoHotkey v2
#include C:\Users\eduev\Documents\UIA-v2\Lib\UIA.ahk

!#s:: {
    ; Get the UIA element from the active window
    win := UIA.ElementFromHandle(WinActive("A"))

    ; Find the "Save" or "Salvar" button (adjust Name or AutomationId as necessary)
    saveBtn := win.FindElement({ Name: ["Save", "Salvar"], ControlType: "Button" })

    ; Check if the "Save" button was found and click it
    if saveBtn
        saveBtn.Click()
    else
        MsgBox("Save button not found.")
}
