; =============================================================================
; Utils module: reading_mode.ahk
; Reading Mode — Left/Right → PgUp/PgDn + last fully visible line highlight.
; Toggle via Macros (#!+W → k). Process-local state (same as Focus Mode).
; Locate pipeline: UIA TextPattern → browser JS → PDF best-effort (soft fail).
; =============================================================================

global g_ReadingModeOn := false
global g_ReadingModeOverlay := 0
global g_ReadingModeHighlightMs := 2500
global g_ReadingModeSettleMs := 80
global g_ReadingModeViewportPad := 3

ReadingMode_IsActive() {
    global g_ReadingModeOn
    return !!g_ReadingModeOn
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

ReadingMode_ShowAnchor(rect) {
    global g_ReadingModeOverlay, g_ReadingModeHighlightMs
    if (!IsObject(rect) || rect.w < 2 || rect.h < 2)
        return
    ReadingMode_ClearOverlay()
    try {
        overlay := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x20") ; WS_EX_TRANSPARENT
        overlay.Opt("-DPIScale")
        overlay.BackColor := "F1C40F"
        overlay.Show("NA x" rect.x " y" rect.y " w" rect.w " h" rect.h)
        try WinSetTransparent(110, overlay)
        catch {
        }
        g_ReadingModeOverlay := overlay
        SetTimer(ReadingMode_ClearOverlay, -g_ReadingModeHighlightMs)
    } catch {
        g_ReadingModeOverlay := 0
    }
}

EnableReadingMode() {
    global g_ReadingModeOn
    if (g_ReadingModeOn)
        return
    g_ReadingModeOn := true
    try ShowCenteredOverlay_Utils("📖 Reading Mode ON — ←/→ = Page Up/Down", 1800, BANNER_ACCENT_INFO)
    catch {
    }
}

DisableReadingMode() {
    global g_ReadingModeOn
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
    global g_ReadingModeSettleMs
    if (direction > 0)
        Send("{PgDn}")
    else
        Send("{PgUp}")
    Sleep g_ReadingModeSettleMs
    hwnd := WinExist("A")
    if (!hwnd)
        return
    rect := ReadingMode_GetAnchorRect(hwnd)
    if (IsObject(rect))
        ReadingMode_ShowAnchor(rect)
}

; ---------------------------------------------------------------------------
; Locate: last fully visible line rect {x,y,w,h} in screen coords, or false
; ---------------------------------------------------------------------------
ReadingMode_GetAnchorRect(hwnd) {
    if (!hwnd)
        return false

    rect := ReadingMode_TryUiaAnchor(hwnd)
    if (IsObject(rect))
        return rect

    if (ReadingMode_IsBrowserHwnd(hwnd)) {
        rect := ReadingMode_TryBrowserJsAnchor(hwnd)
        if (IsObject(rect))
            return rect
    }

    ; PDF / other: UIA already tried; browser JS only if Chromium PDF tab
    if (ReadingMode_LooksLikePdf(hwnd)) {
        if (ReadingMode_IsBrowserHwnd(hwnd)) {
            rect := ReadingMode_TryBrowserJsAnchor(hwnd)
            if (IsObject(rect))
                return rect
        }
        ; Soft fail — remaps still work
        return false
    }

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
            if (IsObject(best))
                break
        } catch {
        }
    }

    if (IsObject(best))
        return best

    ; Fallback: walk GetVisibleRanges by Line
    try {
        ranges := tp.GetVisibleRanges()
    } catch {
        return false
    }
    if (!IsObject(ranges) || ranges.Length < 1)
        return false

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
                    if (bottom >= bestBottom) {
                        bestBottom := bottom
                        best := { x: r.x, y: r.y, w: r.w, h: r.h }
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
    return IsObject(best) ? best : false
}

; ---------------------------------------------------------------------------
; Layer 2 — Chrome/Edge (and friends) JS line boxes
; ---------------------------------------------------------------------------
ReadingMode_BrowserJsPayload() {
    ; Returns JSON {x,y,w,h,dpr} in CSS client (viewport) coords, or "".
    return "(function(){var vh=window.innerHeight,vw=window.innerWidth,eps=3,best=null,dpr=window.devicePixelRatio||1;function consider(r){if(!r||r.width<2||r.height<2)return;if(r.top<eps||r.bottom>vh-eps||r.left<-2||r.right>vw+2)return;if(!best||r.bottom>best.b)best={x:r.left,y:r.top,w:r.width,h:r.height,b:r.bottom};}try{var w=document.createTreeWalker(document.body,NodeFilter.SHOW_TEXT,null),n;while(n=w.nextNode()){if(!n.nodeValue||!/\S/.test(n.nodeValue))continue;var rg=document.createRange();rg.selectNodeContents(n);var rs=rg.getClientRects();for(var i=0;i<rs.length;i++)consider(rs[i]);}}catch(e){}if(!best){try{document.querySelectorAll('p,li,h1,h2,h3,h4,h5,h6,div,span,td,th,pre,blockquote,article').forEach(function(el){var rs=el.getClientRects();for(var i=0;i<rs.length;i++)consider(rs[i]);});}catch(e){}}return best?JSON.stringify({x:Math.round(best.x),y:Math.round(best.y),w:Math.round(best.w),h:Math.round(best.h),dpr:dpr}):'';})()"
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