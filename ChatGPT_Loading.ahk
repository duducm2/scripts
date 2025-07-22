#Requires AutoHotkey v2.0+

; -----------------------------------------------------------------------------
; ChatGPT loading indicator utilities (green notification)
;   • ShowChatGPTLoadingIndicator(state)   → shows / updates indicator
;   • HideChatGPTLoadingIndicator()        → hides indicator
;   • WaitForChatGPTButtonAndShowLoading(buttonNames, stateText, timeoutMs)
;       - polls for any UIA button in `buttonNames` (array of Name strings)
;       - shows indicator while button exists, hides when it disappears
; -----------------------------------------------------------------------------

#include UIA-v2\Lib\UIA_Browser.ahk   ; rely on UIA_Browser from main scripts

; persistent GUI handle
global chatgptLoadingGui := ""

ShowChatGPTLoadingIndicator(state := "Loading…") {
    global chatgptLoadingGui

    ; If already visible, just update text
    if (IsObject(chatgptLoadingGui) && chatgptLoadingGui.Hwnd) {
        chatgptLoadingGui.Controls[1].Text := state
        return
    }

    chatgptLoadingGui := Gui()
    chatgptLoadingGui.Opt("+AlwaysOnTop -Caption +ToolWindow")
    chatgptLoadingGui.BackColor := "00FF00"                      ; bright green
    chatgptLoadingGui.SetFont("s28 c000000 Bold", "Segoe UI")
    chatgptLoadingGui.Add("Text", "w600 Center", state)

    ; Centre over active window (fallback: primary monitor)
    activeWin := WinGetID("A")
    if (activeWin) {
        WinGetPos(&wx, &wy, &ww, &wh, activeWin)
    } else {
        work := SysGet.MonitorWorkArea(SysGet.MonitorPrimary)
        wx := work.Left, wy := work.Top, ww := work.Right - work.Left, wh := work.Bottom - work.Top
    }

    chatgptLoadingGui.Show("AutoSize Hide")
    chatgptLoadingGui.GetPos(, , &gw, &gh)
    gx := wx + (ww - gw) / 2
    gy := wy + (wh - gh) / 2
    chatgptLoadingGui.Show("x" Round(gx) " y" Round(gy) " NA")
    WinSetTransparent(220, chatgptLoadingGui)
}

HideChatGPTLoadingIndicator() {
    global chatgptLoadingGui
    if (IsObject(chatgptLoadingGui) && chatgptLoadingGui.Hwnd) {
        chatgptLoadingGui.Destroy()
        chatgptLoadingGui := ""
    }
}

; Polls UIA for any of the given button names and shows indicator while present
WaitForChatGPTButtonAndShowLoading(buttonNames, stateText := "Loading…", timeout := 15000) {
    try cUIA := UIA_Browser()
    catch {
        return
    }
    start := A_TickCount
    btn := ""
    while ((A_TickCount - start) < timeout) {
        ; detect button presence
        btn := ""
        for n in buttonNames {
            try {
                btn := cUIA.FindElement({ Name: n, Type: "Button" })
            } catch {
                btn := ""
            }
            if btn
                break
        }
        if btn {
            ShowChatGPTLoadingIndicator(stateText)
            ; Wait until it disappears or timeout
            while btn && ((A_TickCount - start) < timeout) {
                Sleep 250
                btn := ""
                for n in buttonNames {
                    try {
                        btn := cUIA.FindElement({ Name: n, Type: "Button" })
                    } catch {
                        btn := ""
                    }
                    if btn
                        break
                }
            }
            break
        }
        Sleep 250
    }
    HideChatGPTLoadingIndicator()
}
