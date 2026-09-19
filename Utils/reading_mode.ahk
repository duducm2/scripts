; =============================================================================
; Utils module: reading_mode.ahk
; Reading Mode — Left/Right → PgUp/PgDn + last fully visible line highlight.
; Toggle via Macros (#!+W → k). Process-local state (same as Focus Mode).
; Locate pipeline: UIA TextPattern → browser JS → PDF best-effort (soft fail).
; =============================================================================

global g_ReadingModeOn := false
global g_ReadingModeOverlay := 0
global g_ReadingModeSettleMs := 80
global g_ReadingModeViewportPad := 3
global g_ReadingModeTrackedHwnd := 0
global g_ReadingModeTrackedTitle := ""
global g_ReadingModeContextTimer := false
global g_ReadingModeContextPollMs := 250
global g_ReadingModeBarWidth := 10
global g_ReadingModeHoldMs := 3000
global g_ReadingModeBlinkStep := 0
global g_ReadingModePageOverlapFrac := 0.12

; #region agent log
ReadingMode_DebugLog(hypothesisId, location, message, dataJson := "{}") {
    try {
        logPath := A_ScriptDir "\debug-289db7.log"
        loc := StrReplace(location, '"', "'")
        msg := StrReplace(message, '"', "'")
        line := '{"sessionId":"289db7","hypothesisId":"' hypothesisId '","location":"' loc '","message":"' msg '","data":' dataJson ',"timestamp":' A_TickCount '}`n'
        FileAppend(line, logPath, "UTF-8")
    } catch {
    }
}
ReadingMode_RectJson(r) {
    if (!IsObject(r))
        return "null"
    return '{"x":' r.x ',"y":' r.y ',"w":' r.w ',"h":' r.h '}'
}
; #endregion

ReadingMode_IsActive() {
    global g_ReadingModeOn
    return !!g_ReadingModeOn
}

ReadingMode_CaptureContext() {
    global g_ReadingModeTrackedHwnd, g_ReadingModeTrackedTitle
    hwnd := WinExist("A")
    g_ReadingModeTrackedHwnd := hwnd ? hwnd : 0
    g_ReadingModeTrackedTitle := ""
    if (hwnd) {
        try g_ReadingModeTrackedTitle := WinGetTitle("ahk_id " hwnd)
        catch {
            g_ReadingModeTrackedTitle := ""
        }
    }
}

ReadingMode_ClearTrackedContext() {
    global g_ReadingModeTrackedHwnd, g_ReadingModeTrackedTitle
    g_ReadingModeTrackedHwnd := 0
    g_ReadingModeTrackedTitle := ""
}

ReadingMode_StartContextMonitor() {
    global g_ReadingModeContextTimer, g_ReadingModeContextPollMs
    ReadingMode_StopContextMonitor()
    g_ReadingModeContextTimer := SetTimer(ReadingMode_ContextMonitor, g_ReadingModeContextPollMs)
}

ReadingMode_StopContextMonitor() {
    global g_ReadingModeContextTimer
    if (!IsSet(g_ReadingModeContextTimer))
        g_ReadingModeContextTimer := false
    if (g_ReadingModeContextTimer) {
        try SetTimer(g_ReadingModeContextTimer, 0)
        catch {
            try SetTimer(ReadingMode_ContextMonitor, 0)
            catch {
            }
        }
        g_ReadingModeContextTimer := false
    }
}

; Auto-off when active window changes, or Chrome/Edge tab changes (title changes).
ReadingMode_ContextMonitor(*) {
    global g_ReadingModeOn, g_ReadingModeTrackedHwnd, g_ReadingModeTrackedTitle, g_ReadingModeOverlay
    if (!g_ReadingModeOn || !g_ReadingModeTrackedHwnd)
        return

    fg := WinExist("A")
    if (!fg)
        return

    ; Ignore this script's own GUIs (banner / line overlay) so they don't trip auto-off
    try {
        if (WinGetPID("ahk_id " fg) = ProcessExist()) {
            if (IsObject(g_ReadingModeOverlay) && g_ReadingModeOverlay.Hwnd && fg = g_ReadingModeOverlay.Hwnd)
                return
            if (fg != g_ReadingModeTrackedHwnd)
                return
        }
    } catch {
    }

    if (fg != g_ReadingModeTrackedHwnd) {
        ; #region agent log
        ReadingMode_DebugLog("D", "reading_mode.ahk:ContextMonitor", "auto-off hwnd change", '{"fg":' fg ',"tracked":' g_ReadingModeTrackedHwnd '}'
        )
        ; #endregion
        DisableReadingMode()
        return
    }

    title := ""
    try title := WinGetTitle("ahk_id " fg)
    catch {
        return
    }
    if (title != g_ReadingModeTrackedTitle) {
        ; #region agent log
        ReadingMode_DebugLog("D", "reading_mode.ahk:ContextMonitor", "auto-off title change", '{}')
        ; #endregion
        DisableReadingMode()
    }
}

ReadingMode_ClearOverlay() {
    global g_ReadingModeOverlay
    try SetTimer(ReadingMode_ClearOverlay, 0)
    catch {
    }
    if (IsObject(g_ReadingModeOverlay)) {
        try g_ReadingModeOverlay.Destroy()
        catch {
        }
    }
    g_ReadingModeOverlay := 0
}

ReadingMode_CancelHoldSequence() {
    global g_ReadingModeBlinkStep
    try SetTimer(ReadingMode_AfterHoldBlink, 0)
    catch {
    }
    try SetTimer(ReadingMode_BlinkTick, 0)
    catch {
    }
    g_ReadingModeBlinkStep := 0
}

; Text-element viewport if available; else window bounds.
ReadingMode_GetViewport(hwnd) {
    try {
        root := UIA.ElementFromHandle(hwnd)
        el := ReadingMode_FindTextElement(root)
        if (IsObject(el)) {
            vbr := el.BoundingRectangle
            return { l: vbr.l, t: vbr.t, r: vbr.r, b: vbr.b, h: vbr.b - vbr.t }
        }
    } catch {
    }
    try {
        WinGetPos(&x, &y, &w, &h, "ahk_id " hwnd)
        return { l: x, t: y, r: x + w, b: y + h, h: h }
    } catch {
    }
    return false
}

; After PgDn/PgUp, place the mark on the same two rows (now shifted up/down).
ReadingMode_ContinuityBar(beforeRect, direction, hwnd) {
    global g_ReadingModePageOverlapFrac
    bar := ReadingMode_ToMarkBar(beforeRect)
    if (!IsObject(bar))
        return false
    vp := ReadingMode_GetViewport(hwnd)
    if (!IsObject(vp) || vp.h < 40)
        return false
    shift := Round(vp.h * (1 - g_ReadingModePageOverlapFrac))
    newY := bar.y + (direction > 0 ? -shift : shift)
    ; Keep the continuity mark on-screen
    newY := Max(vp.t + 2, Min(newY, vp.b - bar.h - 2))
    return { x: bar.x, y: newY, w: bar.w, h: bar.h }
}

; After hold: blink, then snap mark to the current last two rows.
ReadingMode_AfterHoldBlink(*) {
    global g_ReadingModeBlinkStep
    if (!ReadingMode_IsActive())
        return
    g_ReadingModeBlinkStep := 0
    ; #region agent log
    ReadingMode_DebugLog("F", "reading_mode.ahk:AfterHoldBlink", "blink start then snap bottom", '{"runId":"post-fix"}'
    )
    ; #endregion
    SetTimer(ReadingMode_BlinkTick, 100)
}

ReadingMode_BlinkTick(*) {
    global g_ReadingModeOverlay, g_ReadingModeBlinkStep
    if (!ReadingMode_IsActive()) {
        try SetTimer(ReadingMode_BlinkTick, 0)
        catch {
        }
        return
    }
    g_ReadingModeBlinkStep += 1
    if (g_ReadingModeBlinkStep > 6) {
        try SetTimer(ReadingMode_BlinkTick, 0)
        catch {
        }
        g_ReadingModeBlinkStep := 0
        ; #region agent log
        ReadingMode_DebugLog("F", "reading_mode.ahk:BlinkTick", "snap to bottom rows", '{"runId":"post-fix"}')
        ; #endregion
        ReadingMode_RefreshMark()
        return
    }
    if (!(IsObject(g_ReadingModeOverlay) && g_ReadingModeOverlay.Hwnd))
        return
    try {
        if (Mod(g_ReadingModeBlinkStep, 2) = 1)
            g_ReadingModeOverlay.Hide()
        else
            g_ReadingModeOverlay.Show("NA")
    } catch {
    }
}

ReadingMode_UnionRects(a, b) {
    if (!IsObject(a))
        return IsObject(b) ? b : false
    if (!IsObject(b))
        return a
    x1 := Min(a.x, b.x)
    y1 := Min(a.y, b.y)
    x2 := Max(a.x + a.w, b.x + b.w)
    y2 := Max(a.y + a.h, b.y + b.h)
    return { x: x1, y: y1, w: x2 - x1, h: y2 - y1 }
}

; Prefer last + previous line; if only one line known, stretch one row upward.
; Reject giant "prev" rects (UIA ExpandToEnclosingUnit can return paragraphs).
ReadingMode_TwoRowRect(last, prev := false) {
    if (!IsObject(last))
        return false
    if (IsObject(prev) && prev.h > 0 && prev.h <= Max(last.h * 2.5, 80))
        return ReadingMode_UnionRects(last, prev)
    h := Max(last.h, 12)
    return { x: last.x, y: last.y - h, w: last.w, h: last.h + h }
}

ReadingMode_ShowAnchor(rect) {
    global g_ReadingModeOverlay
    ; #region agent log
    ReadingMode_DebugLog("C", "reading_mode.ahk:ShowAnchor", "show overlay", '{"rect":' ReadingMode_RectJson(rect) ',"runId":"post-fix"}'
    )
    ; #endregion
    if (!IsObject(rect) || rect.w < 1 || rect.h < 2)
        return
    ; Reposition existing persistent bar when possible
    if (IsObject(g_ReadingModeOverlay) && g_ReadingModeOverlay.Hwnd) {
        try {
            g_ReadingModeOverlay.Show("NA x" rect.x " y" rect.y " w" rect.w " h" rect.h)
            return
        } catch {
        }
    }
    ReadingMode_ClearOverlay()
    try {
        overlay := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x20") ; WS_EX_TRANSPARENT
        overlay.Opt("-DPIScale")
        overlay.BackColor := "F1C40F"
        overlay.Show("NA x" rect.x " y" rect.y " w" rect.w " h" rect.h)
        try WinSetTransparent(140, overlay)
        catch {
        }
        g_ReadingModeOverlay := overlay
    } catch {
        g_ReadingModeOverlay := 0
        ; #region agent log
        ReadingMode_DebugLog("E", "reading_mode.ahk:ShowAnchor", "overlay create failed", '{}')
        ; #endregion
    }
}

; Left-edge mark bar: fixed width, height = last two rows of the current view.
ReadingMode_ToMarkBar(rect) {
    global g_ReadingModeBarWidth
    if (!IsObject(rect) || rect.h < 2)
        return false
    h := rect.h
    ; Guard against broken UIA unions that span the whole viewport
    if (h > 160)
        h := 64
    y := rect.y + rect.h - h
    return { x: rect.x, y: y, w: g_ReadingModeBarWidth, h: h }
}

ReadingMode_RefreshMark() {
    hwnd := WinExist("A")
    if (!hwnd)
        return
    rect := ReadingMode_GetAnchorRect(hwnd)
    bar := ReadingMode_ToMarkBar(rect)
    ; #region agent log
    ReadingMode_DebugLog("B", "reading_mode.ahk:RefreshMark", "mark bar", '{"ok":' (IsObject(bar) ? "true" : "false") ',"rect":' ReadingMode_RectJson(
        rect) ',"bar":' ReadingMode_RectJson(bar) ',"runId":"post-fix"}')
    ; #endregion
    if (IsObject(bar))
        ReadingMode_ShowAnchor(bar)
}

EnableReadingMode() {
    global g_ReadingModeOn
    if (g_ReadingModeOn)
        return
    g_ReadingModeOn := true
    ReadingMode_CaptureContext()
    ReadingMode_StartContextMonitor()
    ReadingMode_RefreshMark()
    try ShowCenteredOverlay_Utils("📖 Reading Mode ON — ←/→ = Page Up/Down", 1800, BANNER_ACCENT_INFO)
    catch {
    }
}

DisableReadingMode() {
    global g_ReadingModeOn
    ReadingMode_CancelHoldSequence()
    ReadingMode_StopContextMonitor()
    ReadingMode_ClearTrackedContext()
    if (!g_ReadingModeOn) {
        ReadingMode_ClearOverlay()
        return
    }
    g_ReadingModeOn := false
    ReadingMode_ClearOverlay()
    try ShowCenteredOverlay_Utils("📖 Reading Mode OFF", 1500, BANNER_ACCENT_SUCCESS)
    catch {
    }
}

ToggleReadingMode() {
    if (ReadingMode_IsActive())
        DisableReadingMode()
    else
        EnableReadingMode()
}

ReadingMode_Page(direction) {
    global g_ReadingModeSettleMs, g_ReadingModeHoldMs
    ReadingMode_CancelHoldSequence()

    hwnd := WinExist("A")
    ; #region agent log
    ReadingMode_DebugLog("A", "reading_mode.ahk:Page", "page start", '{"dir":' direction ',"hwnd":' (hwnd ? hwnd : 0) ',"runId":"post-fix"}'
    )
    ; #endregion

    ; Capture the last two rows BEFORE scrolling (continuity target).
    beforeRect := ReadingMode_GetAnchorRect(hwnd)

    if (direction > 0)
        Send("{PgDn}")
    else
        Send("{PgUp}")
    Sleep g_ReadingModeSettleMs

    hwnd := WinExist("A")
    if (!hwnd) {
        ; #region agent log
        ReadingMode_DebugLog("D", "reading_mode.ahk:Page", "hwnd lost after scroll", '{}')
        ; #endregion
        return
    }

    cont := ReadingMode_ContinuityBar(beforeRect, direction, hwnd)
    ; #region agent log
    ReadingMode_DebugLog("F", "reading_mode.ahk:Page", "continuity mark", '{"ok":' (IsObject(cont) ? "true" : "false") ',"before":' ReadingMode_RectJson(
        beforeRect) ',"cont":' ReadingMode_RectJson(cont) ',"runId":"post-fix"}')
    ; #endregion
    if (IsObject(cont))
        ReadingMode_ShowAnchor(cont)
    else
        ReadingMode_RefreshMark()

    ; Hold on continuity rows, then blink and snap to new bottom two rows.
    SetTimer(ReadingMode_AfterHoldBlink, -g_ReadingModeHoldMs)
}

; ---------------------------------------------------------------------------
; Locate: last fully visible line rect {x,y,w,h} in screen coords, or false
; ---------------------------------------------------------------------------
ReadingMode_GetAnchorRect(hwnd) {
    if (!hwnd)
        return false

    rect := ReadingMode_TryUiaAnchor(hwnd)
    if (IsObject(rect)) {
        ; #region agent log
        ReadingMode_DebugLog("A", "reading_mode.ahk:GetAnchorRect", "uia hit", '{"rect":' ReadingMode_RectJson(rect) '}'
        )
        ; #endregion
        return rect
    }

    if (ReadingMode_IsBrowserHwnd(hwnd)) {
        rect := ReadingMode_TryBrowserJsAnchor(hwnd)
        if (IsObject(rect)) {
            ; #region agent log
            ReadingMode_DebugLog("A", "reading_mode.ahk:GetAnchorRect", "browser js hit", '{"rect":' ReadingMode_RectJson(
                rect) '}')
            ; #endregion
            return rect
        }
    }

    ; PDF / other: UIA already tried; browser JS only if Chromium PDF tab
    if (ReadingMode_LooksLikePdf(hwnd)) {
        if (ReadingMode_IsBrowserHwnd(hwnd)) {
            rect := ReadingMode_TryBrowserJsAnchor(hwnd)
            if (IsObject(rect))
                return rect
        }
        ; Soft fail — remaps still work
        ; #region agent log
        ReadingMode_DebugLog("A", "reading_mode.ahk:GetAnchorRect", "pdf soft fail", '{}')
        ; #endregion
        return false
    }

    ; #region agent log
    ReadingMode_DebugLog("A", "reading_mode.ahk:GetAnchorRect", "no anchor", '{}')
    ; #endregion
    return false
}

ReadingMode_IsBrowserHwnd(hwnd) {
    try {
        exe := WinGetProcessName("ahk_id " hwnd)
        return (exe = "chrome.exe" || exe = "msedge.exe" || exe = "brave.exe" || exe = "vivaldi.exe")
    } catch {
        return false
    }
}

ReadingMode_LooksLikePdf(hwnd) {
    try {
        exe := WinGetProcessName("ahk_id " hwnd)
        if (RegExMatch(exe, "i)^(SumatraPDF|AcroRd32|Acrobat|PDFXCview)\.exe$"))
            return true
        title := WinGetTitle("ahk_id " hwnd)
        if (InStr(title, ".pdf") || InStr(title, ".PDF"))
            return true
        if (ReadingMode_IsBrowserHwnd(hwnd)) {
            try {
                uia := UIA_Browser("ahk_id " hwnd)
                url := ""
                try url := uia.GetCurrentURL()
                catch {
                    try url := uia.URL
                    catch {
                    }
                }
                if (url != "" && (InStr(url, ".pdf") || (InStr(url, "chrome-extension://") && InStr(url, "pdf"))))
                    return true
            } catch {
            }
        }
    } catch {
    }
    return false
}

ReadingMode_RectFullyVisible(rx, ry, rw, rh, vl, vt, vr, vb, pad := 3) {
    if (rw < 2 || rh < 2)
        return false
    bottom := ry + rh
    right := rx + rw
    return (ry >= vt + pad && bottom <= vb - pad && rx >= vl - 2 && right <= vr + 2)
}

; ---------------------------------------------------------------------------
; Layer 1 — UIA TextPattern
; ---------------------------------------------------------------------------
ReadingMode_FindTextElement(root) {
    if (!IsObject(root))
        return 0

    try {
        fe := UIA.GetFocusedElement()
        if (IsObject(fe)) {
            try {
                if (fe.IsTextPatternAvailable)
                    return fe
            } catch {
            }
            p := fe
            loop 10 {
                try p := p.Parent
                catch {
                    break
                }
                if (!IsObject(p))
                    break
                try {
                    if (p.IsTextPatternAvailable)
                        return p
                } catch {
                }
            }
        }
    } catch {
    }

    for typeId in [50030, 50004, 50033] { ; Document, Edit, Document-ish / Pane sometimes
        try {
            el := root.FindFirst({ Type: typeId })
            if (IsObject(el)) {
                try {
                    if (el.IsTextPatternAvailable)
                        return el
                } catch {
                }
            }
        } catch {
        }
    }

    try {
        for el in root.FindAll({ IsTextPatternAvailable: 1 }) {
            try {
                if (el.IsTextPatternAvailable)
                    return el
            } catch {
            }
        }
    } catch {
    }
    return 0
}

ReadingMode_TryUiaAnchor(hwnd) {
    global g_ReadingModeViewportPad
    try {
        root := UIA.ElementFromHandle(hwnd)
    } catch {
        return false
    }
    if (!IsObject(root))
        return false

    el := ReadingMode_FindTextElement(root)
    if (!IsObject(el))
        return false

    try {
        if (!el.IsTextPatternAvailable)
            return false
        tp := el.TextPattern
    } catch {
        return false
    }

    try {
        vbr := el.BoundingRectangle
        vl := vbr.l, vt := vbr.t, vr := vbr.r, vb := vbr.b
    } catch {
        return false
    }
    if (vr - vl < 8 || vb - vt < 8)
        return false

    best := false
    prev := false
    bestBottom := -1

    ; Prefer RangeFromPoint near viewport bottom, walk upward for a fully visible line
    cx := (vl + vr) // 2
    offsets := [6, 14, 28, 48, 72, 100, 140]
    for offset in offsets {
        py := vb - offset
        if (py <= vt)
            break
        try {
            rng := tp.RangeFromPoint(cx, py)
            rng.ExpandToEnclosingUnit(UIA.TextUnit.Line)
            for r in rng.GetBoundingRectangles() {
                if (ReadingMode_RectFullyVisible(r.x, r.y, r.w, r.h, vl, vt, vr, vb, g_ReadingModeViewportPad)) {
                    bottom := r.y + r.h
                    if (bottom > bestBottom) {
                        bestBottom := bottom
                        best := { x: r.x, y: r.y, w: r.w, h: r.h }
                    }
                }
            }
            if (IsObject(best)) {
                ; Line immediately above the last fully visible one
                try {
                    above := rng.Clone()
                    above.ExpandToEnclosingUnit(UIA.TextUnit.Line)
                    if (above.Move(UIA.TextUnit.Line, -1)) {
                        for r in above.GetBoundingRectangles() {
                            if (r.w >= 2 && r.h >= 2) {
                                prev := { x: r.x, y: r.y, w: r.w, h: r.h }
                                break
                            }
                        }
                    }
                } catch {
                }
                break
            }
        } catch {
        }
    }

    if (IsObject(best))
        return ReadingMode_TwoRowRect(best, prev)

    ; Fallback: walk GetVisibleRanges by Line — keep last two fully visible
    try {
        ranges := tp.GetVisibleRanges()
    } catch {
        return false
    }
    if (!IsObject(ranges) || ranges.Length < 1)
        return false

    best := false
    prev := false
    bestBottom := -1
    secondBottom := -1

    for vrng in ranges {
        try {
            line := vrng.Clone()
            line.ExpandToEnclosingUnit(UIA.TextUnit.Line)
        } catch {
            continue
        }
        loop 250 {
            try {
                rects := line.GetBoundingRectangles()
            } catch {
                break
            }
            for r in rects {
                if (ReadingMode_RectFullyVisible(r.x, r.y, r.w, r.h, vl, vt, vr, vb, g_ReadingModeViewportPad)) {
                    bottom := r.y + r.h
                    cand := { x: r.x, y: r.y, w: r.w, h: r.h }
                    if (bottom >= bestBottom) {
                        prev := best
                        secondBottom := bestBottom
                        best := cand
                        bestBottom := bottom
                    } else if (bottom > secondBottom) {
                        prev := cand
                        secondBottom := bottom
                    }
                }
            }
            try {
                moved := line.Move(UIA.TextUnit.Line, 1)
            } catch {
                break
            }
            if (!moved)
                break
        }
    }
    return IsObject(best) ? ReadingMode_TwoRowRect(best, prev) : false
}

; ---------------------------------------------------------------------------
; Layer 2 — Chrome/Edge (and friends) JS line boxes
; ---------------------------------------------------------------------------
ReadingMode_BrowserJsPayload() {
    ; Returns JSON {x,y,w,h,dpr} covering the last two fully visible line boxes, or "".
    return "(function(){var vh=window.innerHeight,vw=window.innerWidth,eps=3,best=null,prev=null,dpr=window.devicePixelRatio||1;function consider(r){if(!r||r.width<2||r.height<2)return;if(r.top<eps||r.bottom>vh-eps||r.left<-2||r.right>vw+2)return;var c={x:r.left,y:r.top,w:r.width,h:r.height,b:r.bottom};if(!best||c.b>best.b){prev=best;best=c;}else if(!prev||(c.b>prev.b&&c.b<best.b-0.5)){prev=c;}}try{var w=document.createTreeWalker(document.body,NodeFilter.SHOW_TEXT,null),n;while(n=w.nextNode()){if(!n.nodeValue||!/\S/.test(n.nodeValue))continue;var rg=document.createRange();rg.selectNodeContents(n);var rs=rg.getClientRects();for(var i=0;i<rs.length;i++)consider(rs[i]);}}catch(e){}if(!best){try{document.querySelectorAll('p,li,h1,h2,h3,h4,h5,h6,div,span,td,th,pre,blockquote,article').forEach(function(el){var rs=el.getClientRects();for(var i=0;i<rs.length;i++)consider(rs[i]);});}catch(e){}}if(!best)return'';var x1=best.x,y1=best.y,x2=best.x+best.w,y2=best.y+best.h;if(prev){x1=Math.min(x1,prev.x);y1=Math.min(y1,prev.y);x2=Math.max(x2,prev.x+prev.w);y2=Math.max(y2,prev.y+prev.h);}else{y1=best.y-best.h;y2=best.y+best.h;}return JSON.stringify({x:Math.round(x1),y:Math.round(y1),w:Math.round(x2-x1),h:Math.round(y2-y1),dpr:dpr});})()"
}

ReadingMode_TryBrowserJsAnchor(hwnd) {
    try {
        uia := UIA_Browser("ahk_id " hwnd)
    } catch {
        return false
    }

    js := ReadingMode_BrowserJsPayload()
    raw := ""
    try raw := uia.JSReturnThroughClipboard(js)
    catch {
        return false
    }
    raw := Trim(raw)
    if (raw = "" || !InStr(raw, '"x"'))
        return false

    x := 0, y := 0, w := 0, h := 0
    if (!RegExMatch(raw, '"x"\s*:\s*(-?\d+)', &mx)
    || !RegExMatch(raw, '"y"\s*:\s*(-?\d+)', &my)
    || !RegExMatch(raw, '"w"\s*:\s*(-?\d+)', &mw)
    || !RegExMatch(raw, '"h"\s*:\s*(-?\d+)', &mh)) {
        return false
    }
    x := Integer(mx[1])
    y := Integer(my[1])
    w := Integer(mw[1])
    h := Integer(mh[1])
    if (w < 2 || h < 2)
        return false

    dpr := 1.0
    if (RegExMatch(raw, '"dpr"\s*:\s*([0-9.]+)', &md)) {
        try dpr := Float(md[1])
        catch {
            dpr := 1.0
        }
    }
    x := Round(x * dpr)
    y := Round(y * dpr)
    w := Round(w * dpr)
    h := Round(h * dpr)

    ox := 0, oy := 0
    converted := false
    try {
        ControlGetPos(&cx, &cy, &ow, &oh, "Chrome_RenderWidgetHostHWND1", "ahk_id " hwnd)
        pt := Buffer(8, 0)
        NumPut("int", cx, pt, 0)
        NumPut("int", cy, pt, 4)
        if (DllCall("ClientToScreen", "ptr", hwnd, "ptr", pt)) {
            ox := NumGet(pt, 0, "int")
            oy := NumGet(pt, 4, "int")
            converted := true
        }
    } catch {
    }
    if (!converted) {
        try {
            br := uia.GetCurrentDocumentElement().GetPos("screen")
            ox := br.x
            oy := br.y
            converted := true
        } catch {
        }
    }
    if (!converted) {
        try WinGetPos(&ox, &oy, , , "ahk_id " hwnd)
        catch {
            return false
        }
    }
    return { x: ox + x, y: oy + y, w: w, h: h }
}

; ---------------------------------------------------------------------------
; Hotkeys — only while Reading Mode is active
; ---------------------------------------------------------------------------
#HotIf ReadingMode_IsActive()
Left:: ReadingMode_Page(-1)
Right:: ReadingMode_Page(1)
#HotIf

RegisterMacro(ToggleReadingMode, "📖 Reading Mode toggle", "k")