; =============================================================================
; Chrome chat feed → bottom (JS-first). Shared by Gemini Personal / Enterprise
; and Copilot Web !u. Efficiency-canon: one JSExecute over FindAll + wheel bursts.
;
; Trade-off: address-bar javascript: can be blocked or briefly focus the omnibox
; on some Chrome policies; thin Ctrl+End + single WheelDown fallback then runs.
; =============================================================================

ChromeChat_ScrollJsPayload() {
    ; Compact IIFE: tallest overflow among scrollingElement/html/body + chat-ish selectors.
    ; CSS attribute values stay unquoted so the javascript: URL stays easy to escape in AHK.
    return "(()=>{const S=['main','[role=main]','[class*=chat]','[class*=conversation]','[class*=messages]','[class*=Conversation]','[class*=Copilot]'];const c=new Set([document.scrollingElement,document.documentElement,document.body]);for(const s of S){try{document.querySelectorAll(s).forEach(e=>c.add(e))}catch(e){}}let best=document.scrollingElement,max=0;for(const el of c){if(!el)continue;const d=el.scrollHeight-el.clientHeight;if(d<=max)continue;const oy=getComputedStyle(el).overflowY;if(el!==document.scrollingElement&&el!==document.documentElement&&el!==document.body&&oy!=='auto'&&oy!=='scroll'&&oy!=='overlay')continue;max=d;best=el}if(best){best.scrollTop=best.scrollHeight;try{best.scrollTo(0,best.scrollHeight)}catch(e){}}try{window.scrollTo(0,Math.max(document.body.scrollHeight,document.documentElement.scrollHeight))}catch(e){}})();"
}

ChromeChat_ScrollFeedToBottomFallback(hwnd) {
    rw := 0
    try rw := ControlGetHwnd("Chrome_RenderWidgetHostHWND1", "ahk_id " hwnd)
    catch
        rw := 0
    target := rw ? rw : hwnd
    ChromeChat_ScrollViaMouseWheel(target, 60)
}

; Scroll by posting WM_MOUSEWHEEL messages to the render widget.
; This never leaks keystrokes into the focused element (unlike ControlSend Ctrl+End).
ChromeChat_ScrollViaMouseWheel(hwnd, clicks := 40) {
    static WM_MOUSEWHEEL := 0x020A
    ; WHEEL_DELTA = 120 per notch; negative = scroll down
    wParam := (-120 * 1) << 16
    ; lParam = cursor pos relative to window; center of the window works fine
    try {
        WinGetPos(&wx, &wy, &ww, &wh, "ahk_id " hwnd)
        cx := ww // 2
        cy := wh // 2
    } catch {
        cx := 400
        cy := 400
    }
    lParam := (cy << 16) | (cx & 0xFFFF)
    loop clicks {
        try PostMessage(WM_MOUSEWHEEL, wParam, lParam, , "ahk_id " hwnd)
        if (Mod(A_Index, 10) = 0)
            Sleep 10
    }
}

; Returns UIA_Browser for hwnd (reuses uia when provided) so callers can FindComposer without re-attach.
ChromeChat_ScrollFeedToBottomFast(hwnd, uia := 0) {
    if (!hwnd)
        return 0
    if (!IsObject(uia)) {
        try uia := UIA_Browser("ahk_id " hwnd)
        catch
            uia := 0
    }
    ; JSExecute intentionally skipped: it types the JS payload into the focused
    ; composer when the omnibox isn't targeted (confirmed via debug session 97a80a).
    ; #region agent log
    try FileAppend(
        '{"sessionId":"97a80a","hypothesisId":"B-fix","location":"ChromeChatScroll:ScrollFeedToBottomFast","message":"scroll via mousewheel","data":{"uiaOk":' (
            IsObject(uia) ? "true" : "false") '},"timestamp":' A_TickCount '}' "`n",
        "C:\Users\eduev\Meu Drive\17 - Projects\scripts\debug-97a80a.log")
    ; #endregion agent log
    ChromeChat_ScrollFeedToBottomFallback(hwnd)
    return IsObject(uia) ? uia : 0
}

; Guard wrapper: snapshots composer text before scrolling and restores it if
; the scroll operation leaked keystrokes into the prompt field.
; composerEl - UIA element for the composer/prompt field (may be 0/empty)
; Returns the composer text that was present before scrolling.
ChromeChat_ComposerSnapshot(composerEl) {
    if (!IsObject(composerEl))
        return ""
    text := ""
    try text := composerEl.Value
    catch {
        try text := composerEl.Name
        catch {
            text := ""
        }
    }
    return text
}

; Restores composer content if it was corrupted by the scroll operation.
; composerEl - same UIA element used for the snapshot
; snapshot   - the string returned by ChromeChat_ComposerSnapshot
ChromeChat_ComposerRestore(composerEl, snapshot) {
    ; #region agent log
    try FileAppend(
        '{"sessionId":"97a80a","hypothesisId":"A-E","location":"ChromeChatScroll:ComposerRestore:entry","message":"restore entry","data":{"hasEl":' (
            IsObject(composerEl) ? "true" : "false") ',"snapshotEmpty":' (snapshot = "" ? "true" : "false") ',"snapshotLen":' StrLen(
                snapshot) '},"timestamp":' A_TickCount '}' "`n",
        "C:\Users\eduev\Meu Drive\17 - Projects\scripts\debug-97a80a.log")
    ; #endregion agent log
    if (!IsObject(composerEl) || snapshot = "")
        return
    current := ""
    try current := composerEl.Value
    catch {
        try current := composerEl.Name
        catch {
            return
        }
    }
    if (current == snapshot)
        return
    try {
        composerEl.ValuePattern.SetValue(snapshot)
        return
    } catch {
    }
    ; Fallback: select-all + paste via clipboard
    try {
        savedClip := A_Clipboard
        A_Clipboard := snapshot
        if ClipWait(1, 1) {
            composerEl.SetFocus()
            Sleep 30
            Send "^a"
            Sleep 30
            Send "^v"
            Sleep 50
        }
        A_Clipboard := savedClip
    } catch {
    }
}
