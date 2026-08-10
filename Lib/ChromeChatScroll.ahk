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
    if (rw) {
        try ControlSend "{Blind}^{End}", , "ahk_id " rw
        catch {
            try ControlSend "{Blind}^{End}", , "ahk_id " hwnd
            catch {
                Send "^{End}"
            }
        }
        try ControlSend "{WheelDown 40}", , "ahk_id " rw
        catch {
        }
    } else {
        try ControlSend "{Blind}^{End}", , "ahk_id " hwnd
        catch {
            Send "^{End}"
        }
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
    jsOk := false
    if (IsObject(uia)) {
        try {
            uia.JSExecute(ChromeChat_ScrollJsPayload())
            jsOk := true
        } catch {
            jsOk := false
        }
    }
    if (!jsOk)
        ChromeChat_ScrollFeedToBottomFallback(hwnd)
    return IsObject(uia) ? uia : 0
}
