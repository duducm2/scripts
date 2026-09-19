; =============================================================================
; Utils module: reading_mode.ahk
; Reading Mode — Left/Right → PgUp/PgDn + last fully visible line highlight.
; Toggle via Macros (#!+W → k). Process-local state (same as Focus Mode).
; Locate: UIA TextPattern only (no address-bar javascript: injection).
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
global g_ReadingModeBrowserSettleMs := 120
global g_ReadingModeSnapLineH := 0
global g_ReadingModeSnapCorrect := false

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
        DisableReadingMode()
        return
    }

    title := ""
    try title := WinGetTitle("ahk_id " fg)
    catch {
        return
    }
    if (title != g_ReadingModeTrackedTitle)
        DisableReadingMode()
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
    global g_ReadingModeBlinkStep, g_ReadingModeSnapCorrect
    try SetTimer(ReadingMode_AfterHoldBlink, 0)
    catch {
    }
    try SetTimer(ReadingMode_BlinkTick, 0)
    catch {
    }
    g_ReadingModeBlinkStep := 0
    g_ReadingModeSnapCorrect := false
}

; Prefer Chrome render widget (true page pixels) over Document BR (often full window).
ReadingMode_GetViewport(hwnd) {
    if (ReadingMode_IsBrowserHwnd(hwnd)) {
        try {
            ControlGetPos(&cx, &cy, &cw, &ch, "Chrome_RenderWidgetHostHWND1", "ahk_id " hwnd)
            if (cw > 40 && ch > 40) {
                pt := Buffer(8, 0)
                NumPut("int", cx, pt, 0)
                NumPut("int", cy, pt, 4)
                if (DllCall("ClientToScreen", "ptr", hwnd, "ptr", pt)) {
                    sx := NumGet(pt, 0, "int")
                    sy := NumGet(pt, 4, "int")
                    return { l: sx, t: sy, r: sx + cw, b: sy + ch, h: ch }
                }
            }
        } catch {
        }
    }
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

; Intersect text-element BR with content viewport so probes stay in the visible page.
ReadingMode_IntersectViewport(elVp, contentVp) {
    if (!IsObject(contentVp))
        return elVp
    if (!IsObject(elVp))
        return contentVp
    l := Max(elVp.l, contentVp.l)
    t := Max(elVp.t, contentVp.t)
    r := Min(elVp.r, contentVp.r)
    b := Min(elVp.b, contentVp.b)
    if (r - l < 40 || b - t < 40)
        return contentVp
    return { l: l, t: t, r: r, b: b, h: b - t }
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
    global g_ReadingModeSnapCorrect
    hwnd := WinExist("A")
    if (!hwnd)
        return
    rect := ReadingMode_GetAnchorRect(hwnd)
    bar := ReadingMode_ToMarkBar(rect)
    if (g_ReadingModeSnapCorrect) {
        bar := ReadingMode_CorrectSnapOvershoot(bar, hwnd)
        g_ReadingModeSnapCorrect := false
    }
    if (IsObject(bar))
        ReadingMode_ShowAnchor(bar)
}

; If the post-blink bottom snap sits too close to the viewport edge, nudge up one row.
ReadingMode_CorrectSnapOvershoot(bar, hwnd) {
    global g_ReadingModeSnapLineH
    if (!IsObject(bar))
        return false
    vp := ReadingMode_GetViewport(hwnd)
    if (!IsObject(vp))
        return bar
    lineH := g_ReadingModeSnapLineH
    if (lineH < 8)
        lineH := Max(Round(bar.h / 2), 12)
    margin := Max(4, Round(lineH * 0.35))
    if (bar.y + bar.h > vp.b - margin)
        bar := { x: bar.x, y: Max(vp.t + 2, bar.y - lineH), w: bar.w, h: bar.h }
    return bar
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
    global g_ReadingModeSettleMs, g_ReadingModeHoldMs, g_ReadingModeBrowserSettleMs,
        g_ReadingModeSnapLineH, g_ReadingModeSnapCorrect
    ReadingMode_CancelHoldSequence()

    hwnd := WinExist("A")

    ; Capture the last two rows BEFORE scrolling (continuity target).
    beforeRect := ReadingMode_GetAnchorRect(hwnd)
    g_ReadingModeSnapLineH := 0
    if (IsObject(beforeRect) && beforeRect.h >= 8)
        g_ReadingModeSnapLineH := Max(Round(beforeRect.h / 2), 12)

    if (direction > 0)
        Send("{PgDn}")
    else
        Send("{PgUp}")

    settleMs := (hwnd && ReadingMode_IsBrowserHwnd(hwnd)) ? g_ReadingModeBrowserSettleMs : g_ReadingModeSettleMs
    Sleep settleMs

    hwnd := WinExist("A")
    if (!hwnd)
        return

    cont := ReadingMode_ContinuityBar(beforeRect, direction, hwnd)
    if (IsObject(cont))
        ReadingMode_ShowAnchor(cont)
    else
        ReadingMode_RefreshMark()

    ; Hold on continuity rows, then blink and snap to new bottom two rows.
    g_ReadingModeSnapCorrect := true
    SetTimer(ReadingMode_AfterHoldBlink, -g_ReadingModeHoldMs)
}

; ---------------------------------------------------------------------------
; Locate: last fully visible line rect {x,y,w,h} in screen coords, or false
; ---------------------------------------------------------------------------
ReadingMode_IsReasonableAnchor(rect) {
    if (!IsObject(rect))
        return false
    ; Allow up to ~two rows; reject giant paragraph/viewport unions
    if (rect.w < 8 || rect.h < 8 || rect.h > 160)
        return false
    return true
}

ReadingMode_GetAnchorRect(hwnd) {
    if (!hwnd)
        return false

    ; UIA TextPattern only. Do not use UIA_Browser JSReturnThroughClipboard /
    ; SetURL("javascript:…") — that dumps script into the address bar.
    rect := ReadingMode_TryUiaAnchor(hwnd)
    if (ReadingMode_IsReasonableAnchor(rect))
        return rect

    ; Soft-fail — remaps still work
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

ReadingMode_RectFullyVisible(rx, ry, rw, rh, vl, vt, vr, vb, padTop := 3, padBottom := 3) {
    if (rw < 2 || rh < 2)
        return false
    bottom := ry + rh
    right := rx + rw
    return (ry >= vt + padTop && bottom <= vb - padBottom && rx >= vl - 2 && right <= vr + 2)
}

; True last fully-visible line: not in the cutoff band near the viewport bottom.
ReadingMode_IsSolidBottomLine(r, vt, vb, vpH) {
    if (!IsObject(r) || r.h < 8)
        return false
    bottom := r.y + r.h
    ; Reject lines sitting in the bottom cutoff band (~55% of line height, min 12px)
    padBottom := Max(12, Round(r.h * 0.55))
    if (bottom > vb - padBottom)
        return false
    ; Reject probes that resolved to the top half (bad RangeFromPoint mapping)
    if (r.y < vt + Round(vpH * 0.45))
        return false
    return true
}

; ---------------------------------------------------------------------------
; UIA TextPattern
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

    for typeId in [50030, 50004, 50033] { ; Document, Edit, Pane
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
        elVp := { l: vbr.l, t: vbr.t, r: vbr.r, b: vbr.b, h: vbr.b - vbr.t }
    } catch {
        return false
    }

    contentVp := ReadingMode_GetViewport(hwnd)
    vp := ReadingMode_IntersectViewport(elVp, contentVp)
    vl := vp.l, vt := vp.t, vr := vp.r, vb := vp.b
    if (vr - vl < 8 || vb - vt < 8)
        return false
    vpH := vb - vt

    best := false
    bestBottom := -1
    bestRng := 0

    ; Probe from bottom; keep lowest solid fully-visible line (not the cutoff/"next" row).
    cx := (vl + vr) // 2
    offsets := [8, 16, 28, 44, 64, 88, 120, 160, 220]
    for offset in offsets {
        py := vb - offset
        if (py <= vt + 8)
            break
        try {
            rng := tp.RangeFromPoint(cx, py)
            rng.ExpandToEnclosingUnit(UIA.TextUnit.Line)
            for r in rng.GetBoundingRectangles() {
                padBottom := Max(12, Round(Max(r.h, 16) * 0.55))
                if (!ReadingMode_RectFullyVisible(r.x, r.y, r.w, r.h, vl, vt, vr, vb, g_ReadingModeViewportPad,
                    padBottom))
                    continue
                if (!ReadingMode_IsSolidBottomLine(r, vt, vb, vpH))
                    continue
                bottom := r.y + r.h
                if (bottom > bestBottom) {
                    bestBottom := bottom
                    best := { x: r.x, y: r.y, w: r.w, h: r.h }
                    bestRng := rng
                }
            }
        } catch {
        }
    }

    prev := false
    if (IsObject(best) && IsObject(bestRng)) {
        try {
            above := bestRng.Clone()
            above.ExpandToEnclosingUnit(UIA.TextUnit.Line)
            if (above.Move(UIA.TextUnit.Line, -1)) {
                for r in above.GetBoundingRectangles() {
                    if (r.w >= 2 && r.h >= 2 && r.h <= Max(best.h * 2.5, 80)) {
                        prev := { x: r.x, y: r.y, w: r.w, h: r.h }
                        break
                    }
                }
            }
        } catch {
        }
    }

    if (IsObject(best))
        return ReadingMode_TwoRowRect(best, prev)

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
                padBottom := Max(12, Round(Max(r.h, 16) * 0.55))
                if (!ReadingMode_RectFullyVisible(r.x, r.y, r.w, r.h, vl, vt, vr, vb, g_ReadingModeViewportPad,
                    padBottom))
                    continue
                if (!ReadingMode_IsSolidBottomLine(r, vt, vb, vpH))
                    continue
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
; Hotkeys — only while Reading Mode is active
; ---------------------------------------------------------------------------
#HotIf ReadingMode_IsActive()
Left:: ReadingMode_Page(-1)
Right:: ReadingMode_Page(1)
#HotIf

RegisterMacro(ToggleReadingMode, "📖 Reading Mode toggle", "k")