; =============================================================================
; Utils module: google_search_first_result.ahk
; Click Google's first organic result (same as Shift+U on Google Search)
; =============================================================================

; hwnd: Chrome window to target; 0 = active browser via UIA_Browser().
; Returns true if a result link was invoked/clicked.
GoogleSearch_ClickFirstResult(hwnd := 0) {
    try {
        uia := hwnd > 0 ? UIA_Browser("ahk_id " hwnd) : UIA_Browser()
        if !uia
            return false

        centerCol := uia.FindFirst({ AutomationId: "center_col" })
        targetLink := ""

        if (centerCol) {
            ; ClassName "LC20lb" is standard for Google result titles
            titleText := centerCol.FindFirst({ ClassName: "LC20lb", MatchMode: "Substring" })
            if (titleText)
                targetLink := titleText.WalkTree("p")
        } else {
            titleText := uia.FindFirst({ ClassName: "LC20lb", MatchMode: "Substring" })
            if (titleText)
                targetLink := titleText.WalkTree("p")
        }

        if (!targetLink)
            return false

        try {
            targetLink.Invoke()
        } catch {
            targetLink.Click()
        }
        return true
    } catch {
        return false
    }
}

; Poll until the first result is clickable, then click it.
GoogleSearch_WaitAndClickFirstResult(hwnd := 0, timeoutMs := 15000) {
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        if (GoogleSearch_ClickFirstResult(hwnd))
            return true
        Sleep 200
    }
    return false
}
