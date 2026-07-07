; =============================================================================
; WindowManagement module: tile_snap.ahk
; Tile background, snap half-pair, maximize helpers
; Extracted verbatim from WindowManagement.ahk; loaded via #include into the
; WindowManagement.ahk process, which remains the entry point / source of truth.
; =============================================================================

; Maximize foreground window via Win API (reliable vs simulating Win+Up / system menu).
; If WinMaximize fails for a stubborn window, fall back to WM_SYSCOMMAND SC_MAXIMIZE (see AutoHotkey WinMaximize docs).
WM_MaximizeActiveWindow() {
    try {
        WinMaximize "A"
    } catch {
        try PostMessage 0x0112, 0xF030, , , "A"  ; WM_SYSCOMMAND, SC_MAXIMIZE
    }
}

WM_MaximizeHwnd(hwnd) {
    if !hwnd
        return
    try {
        WinMaximize "ahk_id " hwnd
    } catch {
        try PostMessage 0x0112, 0xF030, , , "ahk_id " hwnd  ; WM_SYSCOMMAND, SC_MAXIMIZE
    }
}

; 50/50 half-pair snap: DWM-measured programmatic placement (gapless panes with margin + gutter).
; Loose pane fraction constants (WM_ClassifySnapPane defaults for target-pane detection).
WM_SNAP_PANE_MIN_FRAC := 0.15
WM_SNAP_PANE_MAX_FRAC := 0.85
WM_SNAP_ORTH_MIN_FRAC := 0.20
; Strict validation for gapless 50/50 snap.
WM_SNAP_STRICT_PANE_MIN_FRAC := 0.42
WM_SNAP_STRICT_PANE_MAX_FRAC := 0.58
WM_SNAP_STRICT_COVERAGE_MIN_FRAC := 0.97
WM_SNAP_STRICT_EDGE_TOL := 4
WM_SNAP_STRICT_VALIDATE_TIMEOUT_MS := 400
WM_SNAP_STRICT_VALIDATE_POLL_MS := 25
WM_DWMWA_EXTENDED_FRAME_BOUNDS := 9
; Uniform inset inside work area + gutter between the two panes (same recipe as working M4 portrait baseline).
WM_SNAP_PAIR_MARGIN := 6
WM_SNAP_PAIR_GAP := 4

; Absolute placement vs monitor work area (no before/after size comparison).
WM_GetWindowRectHwnd(hwnd, &left, &top, &right, &bottom) {
    rect := Buffer(16, 0)
    if !DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect)
        return false
    left := NumGet(rect, 0, "int")
    top := NumGet(rect, 4, "int")
    right := NumGet(rect, 8, "int")
    bottom := NumGet(rect, 12, "int")
    return true
}

; DPI-scaled invisible resize border (~7px at 96 DPI; matches observed 9px at 144 DPI).
WM_GetDpiScaledEdgeInset(hwnd) {
    dpi := 96
    try dpi := DllCall("GetDpiForWindow", "ptr", hwnd, "uint")
    return Max(0, Round(7 * dpi / 96))
}

WM_GetDpiScaledEdgeForMonitor(monIdx) {
    if (monIdx < 1 || monIdx > MonitorGetCount())
        return 7
    MonitorGet monIdx, &ml, &mt, &mr, &mb
    cx := (ml + mr) // 2
    cy := (mt + mb) // 2
    point64 := (cy & 0xFFFFFFFF) << 32 | (cx & 0xFFFFFFFF)
    hMon := DllCall("MonitorFromPoint", "int64", point64, "uint", 2, "ptr")
    dpiX := 96, dpiY := 96
    try DllCall("Shcore\GetDpiForMonitor", "ptr", hMon, "uint", 0, "uint*", &dpiX, "uint*", &dpiY)
    return Max(0, Round(7 * dpiX / 96))
}

; Margin-inset work area split into start/end pane visible-target rects (exclusive right/bottom).
WM_ComputeSnapPairPaneRects(monIdx, axis) {
    MonitorGetWorkArea monIdx, &wl, &wt, &wr, &wb
    m := WM_SNAP_PAIR_MARGIN
    g := WM_SNAP_PAIR_GAP
    innerL := wl + m
    innerT := wt + m
    innerR := wr - m
    innerB := wb - m
    if (innerR - innerL < 120 || innerB - innerT < 80)
        return Map()
    if (axis = "h") {
        halfW := (innerR - innerL - g) // 2
        if (halfW < 60)
            return Map()
        return Map("start", [innerL, innerT, innerL + halfW, innerB],
        "end", [innerL + halfW + g, innerT, innerR, innerB])
    }
    halfH := (innerB - innerT - g) // 2
    if (halfH < 60)
        return Map()
    return Map("start", [innerL, innerT, innerR, innerT + halfH],
    "end", [innerL, innerT + halfH + g, innerR, innerB])
}

WM_GetSnapInnerWorkArea(monIdx, &wl, &wt, &wr, &wb) {
    MonitorGetWorkArea monIdx, &wl, &wt, &wr, &wb
    m := WM_SNAP_PAIR_MARGIN
    wl += m
    wt += m
    wr -= m
    wb -= m
}

; Estimated visible frame; use target monitor DPI (not stale hwnd DPI after cross-monitor moves).
WM_GetEstimatedVisibleFrameForMonitor(hwnd, monIdx, &left, &top, &right, &bottom) {
    if (!WM_GetWindowRectHwnd(hwnd, &wl, &wt, &wr, &wb))
        return false
    edge := (monIdx >= 1) ? WM_GetDpiScaledEdgeForMonitor(monIdx) : WM_GetDpiScaledEdgeInset(hwnd)
    left := wl + edge
    top := wt
    right := wr - edge
    bottom := wb - edge
    return true
}

WM_LogicalPointFromPhysical(x, y, &lx, &ly) {
    pt := Buffer(8, 0)
    NumPut("int", x, pt, 0)
    NumPut("int", y, pt, 4)
    if DllCall("user32\PhysicalToLogicalPointForPerMonitorDPI", "ptr", pt, "int", 1) {
        lx := NumGet(pt, 0, "int")
        ly := NumGet(pt, 4, "int")
        return true
    }
    lx := x
    ly := y
    return false
}

; DWM extended frame bounds in the same logical space as GetWindowRect / MonitorGetWorkArea.
WM_GetDwmExtendedFrameLogical(hwnd, &left, &top, &right, &bottom) {
    rc := Buffer(16, 0)
    if (DllCall("dwmapi\DwmGetWindowAttribute", "ptr", hwnd, "uint", WM_DWMWA_EXTENDED_FRAME_BOUNDS, "ptr", rc, "uint",
        16) != 0)
        return false
    pxL := NumGet(rc, 0, "int"), pxT := NumGet(rc, 4, "int")
    pxR := NumGet(rc, 8, "int"), pxB := NumGet(rc, 12, "int")
    if (pxR <= pxL || pxB <= pxT)
        return false
    WM_LogicalPointFromPhysical(pxL, pxT, &c1x, &c1y)
    WM_LogicalPointFromPhysical(pxR, pxT, &c2x, &c2y)
    WM_LogicalPointFromPhysical(pxL, pxB, &c3x, &c3y)
    WM_LogicalPointFromPhysical(pxR, pxB, &c4x, &c4y)
    left := Min(c1x, c2x, c3x, c4x)
    top := Min(c1y, c2y, c3y, c4y)
    right := Max(c1x, c2x, c3x, c4x)
    bottom := Max(c1y, c2y, c3y, c4y)
    return (right > left && bottom > top)
}

; True visible frame for snap placement/validation — DWM when sane, else DPI-scaled estimate.
WM_GetSnapVisibleFrame(hwnd, monIdx, &left, &top, &right, &bottom, &source := "") {
    source := "estimate"
    if (WM_GetDwmExtendedFrameLogical(hwnd, &dl, &dt, &dr, &db)) {
        if (WM_GetWindowRectHwnd(hwnd, &wl, &wt, &wr, &wb)) {
            il := dl - wl, it := dt - wt, ir := wr - dr, ib := wb - db
            if (il >= -2 && it >= -2 && ir >= -2 && ib >= -2 && il <= 40 && it <= 40 && ir <= 40 && ib <= 40) {
                left := dl, top := dt, right := dr, bottom := db
                source := "dwm"
                return true
            }
        } else {
            left := dl, top := dt, right := dr, bottom := db
            source := "dwm"
            return true
        }
    }
    source := "estimate"
    return WM_GetEstimatedVisibleFrameForMonitor(hwnd, monIdx, &left, &top, &right, &bottom)
}

; Measured outer-window insets (GetWindowRect minus visible frame) for target monitor context.
WM_MeasureOuterInsetsForSnap(hwnd, monIdx, &insetLeft, &insetTop, &insetRight, &insetBottom, &source := "") {
    insetLeft := 0, insetTop := 0, insetRight := 0, insetBottom := 0
    if (!WM_GetWindowRectHwnd(hwnd, &wl, &wt, &wr, &wb))
        return false
    if (!WM_GetSnapVisibleFrame(hwnd, monIdx, &fl, &ft, &fr, &fb, &source))
        return false
    insetLeft := Max(0, fl - wl)
    insetTop := Max(0, ft - wt)
    insetRight := Max(0, wr - fr)
    insetBottom := Max(0, wb - fb)
    return true
}

WM_ComputeOuterRectForVisibleTarget(targetLeft, targetTop, targetRight, targetBottom, il, it, ir, ib, &x, &y, &w, &h) {
    paneW := targetRight - targetLeft
    paneH := targetBottom - targetTop
    x := targetLeft - il
    y := targetTop - it
    w := paneW + il + ir
    h := paneH + it + ib
}

WM_ClampOuterRectToMonitor(monIdx, &x, &y, &w, &h) {
    MonitorGetWorkArea monIdx, &mwl, &mwt, &mwr, &mwb
    if (x < mwl) {
        w -= (mwl - x)
        x := mwl
    }
    if (y < mwt) {
        h -= (mwt - y)
        y := mwt
    }
    if (x + w > mwr)
        w := Max(1, mwr - x)
    if (y + h > mwb)
        h := Max(1, mwb - y)
}

; Move then resize separately — avoids cross-DPI width ballooning from combined SetWindowPos.
WM_ForceMoveHwndToRectSplit(hwnd, x, y, w, h) {
    if (!hwnd || w < 1 || h < 1)
        return false
    try {
        if (!DllCall("SetWindowPos", "ptr", hwnd, "ptr", 0, "int", x, "int", y, "int", 0, "int", 0, "uint", 0x0015))
            throw Error("SetWindowPos move failed")
        if (!DllCall("SetWindowPos", "ptr", hwnd, "ptr", 0, "int", 0, "int", 0, "int", w, "int", h, "uint", 0x0016))
            throw Error("SetWindowPos size failed")
        return true
    } catch {
        return WM_ForceMoveHwndToRect(hwnd, x, y, w, h)
    }
}

; Estimated visible frame in the same logical space as GetWindowRect / MonitorGetWorkArea.
WM_GetEstimatedVisibleFrame(hwnd, &left, &top, &right, &bottom) {
    return WM_GetEstimatedVisibleFrameForMonitor(hwnd, 0, &left, &top, &right, &bottom)
}

; Visible frame for snap validation/placement.
WM_GetDwmVisibleFrameHwnd(hwnd, &left, &top, &right, &bottom) {
    src := ""
    return WM_GetSnapVisibleFrame(hwnd, 0, &left, &top, &right, &bottom, &src)
}

; Snap placement move — issues MoveWindow without the background-tile verify tolerance gate.
WM_ForceMoveHwndToRect(hwnd, left, top, width, height) {
    if (!hwnd || width < 1 || height < 1)
        return false
    try WinMove(hwnd, left, top, width, height)
    catch {
        try DllCall("MoveWindow", "ptr", hwnd, "int", left, "int", top, "int", width, "int", height, "int", true)
        catch
            return false
    }
    return true
}

WM_GaplessFrameAligned(hwnd, monIdx, targetLeft, targetTop, targetRight, targetBottom) {
    fl := 0, ft := 0, fr := 0, fb := 0
    if (!WM_GetSnapVisibleFrame(hwnd, monIdx, &fl, &ft, &fr, &fb))
        return false
    edge := WM_GetDpiScaledEdgeForMonitor(monIdx)
    WM_GetSnapInnerWorkArea(monIdx, &iwl, &iwt, &iwr, &iwb)
    tol := WM_SNAP_STRICT_EDGE_TOL
    boundTol := Max(tol, edge)
    leftTol := (targetLeft <= iwl + 1) ? boundTol : tol
    topTol := (targetTop <= iwt + 1) ? boundTol : tol
    rightTol := (targetRight >= iwr - 1) ? boundTol : tol
    bottomTol := (targetBottom >= iwb - 1) ? boundTol : tol
    return (Abs(fl - targetLeft) <= leftTol && Abs(ft - targetTop) <= topTol
    && Abs(fr - targetRight) <= rightTol && Abs(fb - targetBottom) <= bottomTol)
}

; Places hwnd so its visible frame lands on targetLeft..targetRight (exclusive right/bottom).
WM_MoveHwndToRectGapless(hwnd, monIdx, targetLeft, targetTop, targetRight, targetBottom) {
    paneW := targetRight - targetLeft
    paneH := targetBottom - targetTop
    il := WM_GetDpiScaledEdgeForMonitor(monIdx)
    it := 0, ir := il, ib := il
    src := "estimate"
    WM_ComputeOuterRectForVisibleTarget(targetLeft, targetTop, targetRight, targetBottom, il, it, ir, ib, &x, &y, &w, &
        h)
    WM_ClampOuterRectToMonitor(monIdx, &x, &y, &w, &h)

    if (!WM_ForceMoveHwndToRectSplit(hwnd, x, y, w, h))
        return false

    loop 6 {
        if (WM_GaplessFrameAligned(hwnd, monIdx, targetLeft, targetTop, targetRight, targetBottom))
            return true
        fl := 0, ft := 0, fr := 0, fb := 0
        if (!WM_GetSnapVisibleFrame(hwnd, monIdx, &fl, &ft, &fr, &fb, &src))
            return false
        if (!WM_MeasureOuterInsetsForSnap(hwnd, monIdx, &il, &it, &ir, &ib, &src))
            return false
        WM_ComputeOuterRectForVisibleTarget(targetLeft, targetTop, targetRight, targetBottom, il, it, ir, ib, &x, &y, &
            w,
            &h)
        WM_ClampOuterRectToMonitor(monIdx, &x, &y, &w, &h)
        maxOuterW := paneW + il + ir + 8
        maxOuterH := paneH + it + ib + 8
        if (w > maxOuterW)
            w := maxOuterW
        if (h > maxOuterH)
            h := maxOuterH
        if (!WM_ForceMoveHwndToRectSplit(hwnd, x, y, w, h))
            return false
    }
    return WM_GaplessFrameAligned(hwnd, monIdx, targetLeft, targetTop, targetRight, targetBottom)
}

; "h" = left/right panes (landscape); "v" = top/bottom panes (portrait).
WM_GetSnapSplitAxis(monIdx) {
    MonitorGetWorkArea(monIdx, &wl, &wt, &wr, &wb)
    return (wr - wl >= wb - wt) ? "h" : "v"
}

; Classifies window into start/end pane on split axis; sets paneSize on that axis.
WM_ClassifySnapPane(axis, wl, wt, wr, wb, left, top, right, bottom, &pane, &paneSize, minFrac := WM_SNAP_PANE_MIN_FRAC,
    maxFrac := WM_SNAP_PANE_MAX_FRAC) {
    pane := ""
    paneSize := 0
    workW := wr - wl
    workH := wb - wt
    if (workW < 100 || workH < 100)
        return false

    if (axis = "h") {
        tol := Max(40, Round(workW * 0.08))
        orthTol := Max(40, Round(workH * 0.10))
        minOrth := Max(200, Round(workH * WM_SNAP_ORTH_MIN_FRAC))
        w := right - left
        h := bottom - top
        workDim := workW
        paneSize := w
        if (h < minOrth || top < wt - orthTol || bottom > wb + orthTol)
            return false
    } else {
        tol := Max(40, Round(workH * 0.08))
        orthTol := Max(40, Round(workW * 0.10))
        minOrth := Max(200, Round(workW * WM_SNAP_ORTH_MIN_FRAC))
        w := right - left
        h := bottom - top
        workDim := workH
        paneSize := h
        if (w < minOrth || left < wl - orthTol || right > wr + orthTol)
            return false
    }

    minPane := Round(workDim * minFrac)
    maxPane := Round(workDim * maxFrac)
    if (paneSize < minPane || paneSize > maxPane)
        return false

    if (axis = "h") {
        center := wl + workW // 2
        onStart := (left <= wl + tol)
        onEnd := (right >= wr - tol)
        if (onStart && !onEnd) {
            pane := "start"
            return true
        }
        if (onEnd && !onStart) {
            pane := "end"
            return true
        }
        pane := ((left + right) // 2 < center) ? "start" : "end"
        return true
    }
    center := wt + workH // 2
    onStart := (top <= wt + tol)
    onEnd := (bottom >= wb - tol)
    if (onStart && !onEnd) {
        pane := "start"
        return true
    }
    if (onEnd && !onStart) {
        pane := "end"
        return true
    }
    pane := ((top + bottom) // 2 < center) ? "start" : "end"
    return true
}

; Edge alignment for visible frame vs margin-inset work area and shared gutter (edgeTol allows DPI border at outer edges).
WM_SnapPaneEdgesAligned(axis, wl, wt, wr, wb, left, top, right, bottom, pane, edgeTol := 0) {
    g := WM_SNAP_PAIR_GAP
    tol := WM_SNAP_STRICT_EDGE_TOL
    boundTol := edgeTol > 0 ? Max(tol, edgeTol) : tol
    if (axis = "h") {
        halfW := (wr - wl - g) // 2
        startEdge := wl + halfW
        endEdge := wl + halfW + g
        if (Abs(top - wt) > boundTol || Abs(bottom - wb) > boundTol)
            return false
        if (pane = "start")
            return (Abs(left - wl) <= boundTol && Abs(right - startEdge) <= tol)
        return (Abs(left - endEdge) <= tol && Abs(right - wr) <= boundTol)
    }
    halfH := (wb - wt - g) // 2
    startEdge := wt + halfH
    endEdge := wt + halfH + g
    if (Abs(left - wl) > boundTol || Abs(right - wr) > boundTol)
        return false
    if (pane = "start")
        return (Abs(top - wt) <= boundTol && Abs(bottom - startEdge) <= tol)
    return (Abs(top - endEdge) <= tol && Abs(bottom - wb) <= boundTol)
}

WM_TryClassifySnapPartnerPane(monIdx, axis, wl, wt, wr, wb, hwnd, oppPane, &paneSize, edgeTol := 0) {
    winL := 0, winT := 0, winR := 0, winB := 0
    if (!WM_GetSnapVisibleFrame(hwnd, monIdx, &winL, &winT, &winR, &winB))
        return false
    pane := ""
    paneSize := 0
    if (!WM_ClassifySnapPane(axis, wl, wt, wr, wb, winL, winT, winR, winB, &pane, &paneSize,
        WM_SNAP_STRICT_PANE_MIN_FRAC, WM_SNAP_STRICT_PANE_MAX_FRAC))
        return false
    if (pane != oppPane)
        return false
    if (!WM_SnapPaneEdgesAligned(axis, wl, wt, wr, wb, winL, winT, winR, winB, pane, edgeTol))
        return false
    return true
}

WM_ValidateSnapBipartitionStrict(monIdx, primaryHwnd, &failReason := "", partnerHwnd := 0) {
    failReason := ""
    if (!primaryHwnd || monIdx < 1 || monIdx > MonitorGetCount()) {
        failReason := "invalid_args"
        return false
    }
    try {
        if (WinGetMinMax("ahk_id " primaryHwnd) = 1) {
            failReason := "primary_maximized"
            return false
        }
    } catch {
        failReason := "primary_minmax_error"
        return false
    }
    axis := WM_GetSnapSplitAxis(monIdx)
    WM_GetSnapInnerWorkArea(monIdx, &wl, &wt, &wr, &wb)
    workW := wr - wl
    workH := wb - wt
    workDim := (axis = "h") ? workW : workH
    monEdge := WM_GetDpiScaledEdgeForMonitor(monIdx)
    if (!WM_GetSnapVisibleFrame(primaryHwnd, monIdx, &pl, &pt, &pr, &pb)) {
        failReason := "primary_no_rect"
        return false
    }
    primaryPane := ""
    primaryPaneSize := 0
    if (!WM_ClassifySnapPane(axis, wl, wt, wr, wb, pl, pt, pr, pb, &primaryPane, &primaryPaneSize,
        WM_SNAP_STRICT_PANE_MIN_FRAC, WM_SNAP_STRICT_PANE_MAX_FRAC)) {
        failReason := "primary_not_in_pane"
        return false
    }
    if (!WM_SnapPaneEdgesAligned(axis, wl, wt, wr, wb, pl, pt, pr, pb, primaryPane, monEdge)) {
        failReason := "primary_edge_gap"
        return false
    }

    oppPane := (primaryPane = "start") ? "end" : "start"
    bestOppSize := 0
    foundOpp := false
    partnerValidated := false
    if (partnerHwnd && partnerHwnd != primaryHwnd) {
        oppSize := 0
        if (WM_TryClassifySnapPartnerPane(monIdx, axis, wl, wt, wr, wb, partnerHwnd, oppPane, &oppSize, monEdge)) {
            foundOpp := true
            bestOppSize := oppSize
            partnerValidated := true
        }
    }
    if (!partnerValidated) {
        for win in GetVisibleWindowsOnMonitor(monIdx, true) {
            if (win.hwnd = primaryHwnd || (partnerHwnd && win.hwnd = partnerHwnd))
                continue
            oppSize := 0
            if (!WM_TryClassifySnapPartnerPane(monIdx, axis, wl, wt, wr, wb, win.hwnd, oppPane, &oppSize, monEdge))
                continue
            foundOpp := true
            if (oppSize > bestOppSize)
                bestOppSize := oppSize
        }
    }
    if (!foundOpp) {
        failReason := "no_opposite_pane"
        return false
    }
    if ((primaryPaneSize + bestOppSize) / workDim < WM_SNAP_STRICT_COVERAGE_MIN_FRAC) {
        failReason := "insufficient_coverage"
        return false
    }
    return true
}

WM_WaitValidateSnapBipartitionStrict(monIdx, primaryHwnd, partnerHwnd := 0) {
    deadline := A_TickCount + WM_SNAP_STRICT_VALIDATE_TIMEOUT_MS
    while (A_TickCount < deadline) {
        if (WM_ValidateSnapBipartitionStrict(monIdx, primaryHwnd, , partnerHwnd))
            return true
        Sleep WM_SNAP_STRICT_VALIDATE_POLL_MS
    }
    return false
}

WM_SnapPairGaplessRects(monIdx, axis, targetHwnd, targetPane, partnerHwnd, partnerPane) {
    rects := WM_ComputeSnapPairPaneRects(monIdx, axis)
    if (rects.Count = 0)
        return false
    tRect := rects[targetPane]
    pRect := rects[partnerPane]
    ok1 := WM_MoveHwndToRectGapless(targetHwnd, monIdx, tRect[1], tRect[2], tRect[3], tRect[4])
    ok2 := WM_MoveHwndToRectGapless(partnerHwnd, monIdx, pRect[1], pRect[2], pRect[3], pRect[4])
    return ok1 && ok2
}

WM_SnapHalfPairActiveWindow() {
    targetHwnd := 0
    try targetHwnd := WinExist("A")
    catch
        targetHwnd := 0
    if (!targetHwnd) {
        ShowNotification_WM("No active window to snap.")
        return
    }
    if (WM_IsExcludedIndicatorWindow(targetHwnd)) {
        ShowNotification_WM("Cannot snap this window (indicator / overlay).")
        return
    }

    ClipAngel_WaitChordModifiersReleased()
    ClipAngel_ReleaseChordModifiersForSend()

    snapMon := WM_ResolveSnapTargetMonitor(targetHwnd)
    monIdx := snapMon["monIdx"]
    axis := WM_GetSnapSplitAxis(monIdx)

    partnerHwnd := 0
    for hwnd in WM_EnumerateOpenHwndsGlobal() {
        if (hwnd != targetHwnd) {
            partnerHwnd := hwnd
            break
        }
    }
    if (!partnerHwnd) {
        WM_PrepareHwndForTile(targetHwnd)
        WM_MaximizeHwnd(targetHwnd)
        ShowNotification_WM("No other window open anywhere — maximized instead.")
        return
    }

    WM_PrepareHwndForTile(targetHwnd)
    WM_PrepareHwndForTile(partnerHwnd)

    targetPane := "start"
    if (!snapMon["emptyMonitorSnap"] && WM_GetWindowRectHwnd(targetHwnd, &tl, &tt, &tr, &tb)) {
        WM_GetSnapInnerWorkArea(monIdx, &wl, &wt, &wr, &wb)
        pane := ""
        paneSize := 0
        if (WM_ClassifySnapPane(axis, wl, wt, wr, wb, tl, tt, tr, tb, &pane, &paneSize) && pane != "")
            targetPane := pane
    }
    partnerPane := (targetPane = "start") ? "end" : "start"

    ok := WM_SnapPairGaplessRects(monIdx, axis, targetHwnd, targetPane, partnerHwnd, partnerPane)
    finalOk := WM_WaitValidateSnapBipartitionStrict(monIdx, targetHwnd, partnerHwnd)
    if (!ok || !finalOk)
        ShowNotification_WM("Snap failed — window may be resisting resize.")
}

; Per monitor: if exactly one visible non-minimized window and not maximized, maximize it.
WM_MaximizeLonelyVisibleOnAllMonitors() {
    maximized := 0
    loop MonitorGetCount() {
        windows := GetVisibleWindowsOnMonitor(A_Index, true)
        if (windows.Length != 1)
            continue
        hwnd := windows[1].hwnd
        try {
            if (WinGetMinMax("ahk_id " hwnd) = 1)
                continue
        } catch {
            continue
        }
        try {
            WM_MaximizeHwnd(hwnd)
            maximized++
        } catch {
        }
    }
    if (maximized > 0) {
        msg := (maximized = 1) ? "✅ Maximized 1 window" : "✅ Maximized " maximized " windows"
        ShowCenteredOverlay_Utils(msg, 1200, BANNER_ACCENT_SUCCESS)
    }
}

WM_GetHwndMonitorIndex(hwnd) {
    if (!hwnd)
        return 0
    try {
        hMon := DllCall("MonitorFromWindow", "ptr", hwnd, "uint", 2, "ptr")
        loop MonitorGetCount() {
            MonitorGet A_Index, &ml, &mt, &mr, &mb
            cx := (ml + mr) // 2
            cy := (mt + mb) // 2
            point64 := (cy & 0xFFFFFFFF) << 32 | (cx & 0xFFFFFFFF)
            if (Integer(DllCall("MonitorFromPoint", "int64", point64, "uint", 2, "ptr")) = Integer(hMon))
                return A_Index
        }
    } catch {
    }
    return 0
}

; AHK 1-based monitor index under the cursor (matches CycleWindowsOnMonitor empty-monitor navigation).
WM_GetMonitorIndexAtCursor() {
    MouseGetPos(&cx, &cy)
    point64 := (cy & 0xFFFFFFFF) << 32 | (cx & 0xFFFFFFFF)
    try {
        hMon := DllCall("MonitorFromPoint", "int64", point64, "uint", 2, "ptr")
        loop MonitorGetCount() {
            MonitorGet A_Index, &ml, &mt, &mr, &mb
            mcx := (ml + mr) // 2
            mcy := (mt + mb) // 2
            mpoint := (mcy & 0xFFFFFFFF) << 32 | (mcx & 0xFFFFFFFF)
            if (Integer(DllCall("MonitorFromPoint", "int64", mpoint, "uint", 2, "ptr")) = Integer(hMon))
                return A_Index
        }
    } catch {
    }
    return 0
}

; True when the monitor has no unobstructed visible windows (same rule as CycleWindowsOnMonitor).
WM_MonitorIsVisuallyEmpty(monIdx) {
    if (monIdx < 1 || monIdx > MonitorGetCount())
        return false
    try {
        return GetVisibleWindowsOnMonitor(monIdx).Length = 0
    } catch {
        return WM_EnumerateOpenHwndsOnMonitor(monIdx).Length = 0
    }
}

; When cursor sits on an empty monitor, snap there instead of on the still-focused window's monitor.
WM_ResolveSnapTargetMonitor(targetHwnd) {
    fgMonIdx := GetMonitorIndexForForeground_StandardBar()
    cursorMonIdx := WM_GetMonitorIndexAtCursor()
    targetMon := WM_GetHwndMonitorIndex(targetHwnd)
    if (targetMon < 1)
        targetMon := fgMonIdx
    monIdx := targetMon
    emptyMonitorSnap := false
    if (cursorMonIdx >= 1 && cursorMonIdx != targetMon && WM_MonitorIsVisuallyEmpty(cursorMonIdx)) {
        monIdx := cursorMonIdx
        emptyMonitorSnap := true
    }
    return Map("monIdx", monIdx, "fgMonIdx", fgMonIdx, "cursorMonIdx", cursorMonIdx,
        "targetMon", targetMon, "emptyMonitorSnap", emptyMonitorSnap)
}

WM_PrepareHwndForTile(hwnd) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    if (WM_IsExcludedIndicatorWindow(hwnd))
        return false
    try {
        if (WM_WindowIsTaskbarMinimized(hwnd))
            WinRestore "ahk_id " hwnd
        else if (WinGetMinMax(hwnd) = 1)
            WinRestore "ahk_id " hwnd
        WinShow "ahk_id " hwnd
    } catch {
        return false
    }
    return true
}

WM_MoveHwndToRect(hwnd, left, top, width, height) {
    if (!hwnd || width < 1 || height < 1)
        return false
    ok := 0
    try ok := WinMove(hwnd, left, top, width, height)
    catch
        ok := 0
    if !ok {
        try ok := DllCall("MoveWindow", "ptr", hwnd, "int", left, "int", top, "int", width, "int", height, "int", true)
    }
    if (!ok)
        return false
    Sleep 30
    if (!WM_GetWindowRectHwnd(hwnd, &rl, &rt, &rr, &rb))
        return false
    tol := 80
    return (Abs(rl - left) <= tol && Abs(rt - top) <= tol && Abs((rr - rl) - width) <= tol && Abs((rb - rt) - height) <=
    tol)
}

WM_TILE_BG_MAX_TOTAL := 12
WM_TILE_BG_MAX_PER_MON := 3

WM_MonitorIsPortrait(mon) {
    MonitorGetWorkArea mon, &left, &top, &right, &bottom
    return (bottom - top) > (right - left)
}

WM_ResolveHwndMonitorIndex(hwnd, fallbackMon := 0) {
    mon := WM_GetHwndMonitorIndex(hwnd)
    if (mon >= 1)
        return mon
    if (WM_GetWindowRectHwnd(hwnd, &wl, &wt, &wr, &wb)) {
        cx := (wl + wr) // 2
        cy := (wt + wb) // 2
        point64 := (cy & 0xFFFFFFFF) << 32 | (cx & 0xFFFFFFFF)
        try {
            hMon := DllCall("MonitorFromPoint", "int64", point64, "uint", 2, "ptr")
            loop MonitorGetCount() {
                MonitorGet A_Index, &ml, &mt, &mr, &mb
                mcx := (ml + mr) // 2
                mcy := (mt + mb) // 2
                mpoint := (mcy & 0xFFFFFFFF) << 32 | (mcx & 0xFFFFFFFF)
                if (Integer(DllCall("MonitorFromPoint", "int64", mpoint, "uint", 2, "ptr")) = Integer(hMon))
                    return A_Index
            }
        } catch {
        }
    }
    if (fallbackMon >= 1 && fallbackMon <= MonitorGetCount())
        return fallbackMon
    try {
        return GetMonitorIndexForForeground_StandardBar()
    } catch {
        return 0
    }
}

; Eligible windows across ALL monitors, in real global z-order (topmost first == most-recently-used).
WM_EnumerateOpenHwndsGlobal() {
    out := []
    for hwnd in WinGetList() {
        try {
            if (WinGetMinMax(hwnd) = -1)
                continue
            if !DllCall("IsWindowVisible", "ptr", hwnd)
                continue
            exStyle := DllCall("GetWindowLongPtr", "ptr", hwnd, "int", -20, "ptr")
            if (exStyle & 0x00000080)
                continue
            class := WinGetClass(hwnd)
            if (class = "Progman" || class = "WorkerW")
                continue
            if (WinGetTitle(hwnd) = "")
                continue
            if (WM_IsExcludedIndicatorWindow(hwnd))
                continue
            out.Push(hwnd)
        } catch {
        }
    }
    return out
}

; Non-minimized windows on a monitor (includes z-order covered — not only unobstructed visible).
WM_EnumerateOpenHwndsOnMonitor(mon) {
    if (mon < 1 || mon > MonitorGetCount())
        return []
    MonitorGet mon, &ml, &mt, &mr, &mb
    cx := (ml + mr) // 2
    cy := (mt + mb) // 2
    point64 := (cy & 0xFFFFFFFF) << 32 | (cx & 0xFFFFFFFF)
    hTarget := DllCall("MonitorFromPoint", "int64", point64, "uint", 2, "ptr")
    out := []
    for hwnd in WinGetList() {
        try {
            if (WinGetMinMax(hwnd) = -1)
                continue
            if !DllCall("IsWindowVisible", "ptr", hwnd)
                continue
            exStyle := DllCall("GetWindowLongPtr", "ptr", hwnd, "int", -20, "ptr")
            if (exStyle & 0x00000080)
                continue
            hMon := DllCall("MonitorFromWindow", "ptr", hwnd, "uint", 2, "ptr")
            if (Integer(hMon) != Integer(hTarget))
                continue
            class := WinGetClass(hwnd)
            if (class = "Progman" || class = "WorkerW")
                continue
            if (WinGetTitle(hwnd) = "")
                continue
            if (WM_IsExcludedIndicatorWindow(hwnd))
                continue
            out.Push(hwnd)
        } catch {
        }
    }
    return out
}

; Tile organize: user apps on script monitors; skip noise, indicators, and background exclude list.
; Does not skip the foreground window — organize should include every unobstructed app the user sees.
WM_TilePassesOrganizeGates(hwnd) {
    if (!hwnd)
        return false
    try {
        exStyle := DllCall("GetWindowLongPtr", "ptr", hwnd, "int", -20, "ptr")
        if (exStyle & 0x00000080)
            return false
        class := WinGetClass(hwnd)
        if (WM_IsDesktopOrTaskbarClass(class))
            return false
        title := WinGetTitle(hwnd)
        if (title = "")
            return false
        if (WM_BackgroundIsSystemNoiseTitle(title))
            return false
        exe := ""
        try exe := WinGetProcessName("ahk_id " hwnd)
        WM_BackgroundTitleExcludes_Ensure()
        if (WM_BackgroundTitleIsExcluded(title, exe))
            return false
        if (WM_IsExcludedIndicatorWindow(hwnd))
            return false
        if (!WM_BackgroundHwndOnAnyScriptMonitor(hwnd))
            return false
    } catch {
        return false
    }
    return true
}

WM_TileCandidateRegister(&seen, &candidates, hwnd, priority, homeMon, &counters, counterKey) {
    if (!hwnd || seen.Has(hwnd) || homeMon < 1)
        return false
    seen[hwnd] := true
    candidates.Push({ hwnd: hwnd, priority: priority, homeMon: homeMon, order: candidates.Length })
    if (counters.Has(counterKey))
        counters[counterKey]++
    else
        counters[counterKey] := 1
    return true
}

WM_CollectTileEligibleHwnds(foreHwndOverride := 0) {
    foreHwnd := foreHwndOverride
    try {
        if (!foreHwnd)
            foreHwnd := WinGetID("A")
    } catch {
        foreHwnd := 0
    }
    fallbackMon := WM_ResolveHwndMonitorIndex(foreHwnd, 1)
    seen := Map()
    candidates := []
    counters := Map()
    counters["hidden"] := 0
    counters["visible"] := 0
    counters["openOnMon"] := 0
    counters["skippedReject"] := 0
    counters["skippedNoMon"] := 0
    counters["winList"] := 0
    counters["missedVis"] := 0
    counters["foreground"] := 0
    visibleAll := WM_BackgroundBuildVisibleHwndSet()
    prevDetect := A_DetectHiddenWindows
    DetectHiddenWindows true
    try {
        if (foreHwnd && WM_TilePassesOrganizeGates(foreHwnd)) {
            homeMon := WM_ResolveHwndMonitorIndex(foreHwnd, fallbackMon)
            if (homeMon >= 1 && WM_TileCandidateRegister(&seen, &candidates, foreHwnd, 0, homeMon, &counters,
                "foreground"))
                counters["foreground"] := 1
        }
        loop MonitorGetCount() {
            mon := A_Index
            try {
                for win in GetVisibleWindowsOnMonitor(mon, true) {
                    if (seen.Has(win.hwnd))
                        continue
                    if (!WM_TilePassesOrganizeGates(win.hwnd)) {
                        counters["skippedReject"]++
                        continue
                    }
                    homeMon := WM_ResolveHwndMonitorIndex(win.hwnd, mon)
                    if (homeMon < 1) {
                        counters["skippedNoMon"]++
                        continue
                    }
                    WM_TileCandidateRegister(&seen, &candidates, win.hwnd, 1, homeMon, &counters, "visible")
                }
            } catch {
            }
            for hwnd in WM_EnumerateOpenHwndsOnMonitor(mon) {
                if (seen.Has(hwnd))
                    continue
                if (!WM_TilePassesOrganizeGates(hwnd)) {
                    counters["skippedReject"]++
                    continue
                }
                homeMon := WM_ResolveHwndMonitorIndex(hwnd, mon)
                if (homeMon < 1) {
                    counters["skippedNoMon"]++
                    continue
                }
                WM_TileCandidateRegister(&seen, &candidates, hwnd, 2, homeMon, &counters, "openOnMon")
            }
        }
        for hwnd in visibleAll {
            if (seen.Has(hwnd))
                continue
            if (!WM_TilePassesOrganizeGates(hwnd)) {
                counters["skippedReject"]++
                continue
            }
            homeMon := WM_ResolveHwndMonitorIndex(hwnd, fallbackMon)
            if (homeMon < 1) {
                counters["skippedNoMon"]++
                continue
            }
            WM_TileCandidateRegister(&seen, &candidates, hwnd, 1, homeMon, &counters, "missedVis")
        }
        for hwnd in WinGetList() {
            if (seen.Has(hwnd))
                continue
            if (!WM_TilePassesOrganizeGates(hwnd)) {
                counters["skippedReject"]++
                continue
            }
            isMin := false
            try isMin := (WinGetMinMax(hwnd) = -1)
            catch {
            }
            if (isMin) {
                if (!WM_BackgroundPassesVisibleStackGates(hwnd))
                    continue
                priority := 0
            } else {
                if (!DllCall("IsWindowVisible", "ptr", hwnd))
                    continue
                priority := visibleAll.Has(hwnd) ? 1 : 2
            }
            homeMon := WM_ResolveHwndMonitorIndex(hwnd, fallbackMon)
            if (homeMon < 1) {
                counters["skippedNoMon"]++
                continue
            }
            counterKey := (priority = 0) ? "hidden" : ((priority = 1) ? "visible" : "winList")
            WM_TileCandidateRegister(&seen, &candidates, hwnd, priority, homeMon, &counters, counterKey)
        }
    } finally {
        DetectHiddenWindows prevDetect
    }
    return {
        total: candidates.Length,
        hidden: counters.Get("hidden", 0),
        visible: counters.Get("visible", 0),
        openOnMon: counters.Get("openOnMon", 0),
        winList: counters.Get("winList", 0),
        missedVis: counters.Get("missedVis", 0),
        foreground: counters.Get("foreground", 0),
        skippedReject: counters.Get("skippedReject", 0),
        skippedNoMon: counters.Get("skippedNoMon", 0),
        candidates: candidates
    }
}

WM_SelectTileHwndsByPriority(candidates, limit) {
    if (candidates.Length = 0 || limit <= 0)
        return []
    sorted := []
    for c in candidates
        sorted.Push(c)
    n := sorted.Length
    if (n > 1) {
        loop n - 1 {
            loop n - A_Index {
                j := A_Index
                a := sorted[j]
                b := sorted[j + 1]
                swap := false
                if (a.priority > b.priority)
                    swap := true
                else if (a.priority = b.priority && a.order > b.order)
                    swap := true
                if (swap) {
                    tmp := sorted[j]
                    sorted[j] := sorted[j + 1]
                    sorted[j + 1] := tmp
                }
            }
        }
    }
    selected := []
    for c in sorted {
        if (selected.Length >= limit)
            break
        selected.Push(c)
    }
    return selected
}

WM_TileBackgroundReserveCandidates(eligible, selected, foreHwnd := 0) {
    used := Map()
    for item in selected
        used[item.hwnd] := true
    reserve := []
    pool := []
    for c in eligible.candidates
        pool.Push(c)
    n := pool.Length
    if (n > 1) {
        loop n - 1 {
            loop n - A_Index {
                j := A_Index
                a := pool[j]
                b := pool[j + 1]
                swap := false
                if (a.priority > b.priority)
                    swap := true
                else if (a.priority = b.priority && a.order > b.order)
                    swap := true
                if (swap) {
                    tmp := pool[j]
                    pool[j] := pool[j + 1]
                    pool[j + 1] := tmp
                }
            }
        }
    }
    for c in pool {
        if (used.Has(c.hwnd))
            continue
        if (!WM_TilePassesOrganizeGates(c.hwnd))
            continue
        reserve.Push(c)
    }
    return reserve
}

WM_TileBackgroundReplaceFailedWithReserve(&selected, failedHwnds, reserve) {
    if (failedHwnds.Length = 0 || reserve.Length = 0)
        return 0
    failedSet := Map()
    for hwnd in failedHwnds
        failedSet[hwnd] := true
    replaced := 0
    repIdx := 1
    newSelected := []
    for item in selected {
        if (failedSet.Has(item.hwnd)) {
            if (repIdx <= reserve.Length) {
                newSelected.Push(reserve[repIdx++])
                replaced++
            }
        } else
            newSelected.Push(item)
    }
    selected.Length := 0
    for item in newSelected
        selected.Push(item)
    return replaced
}

WM_AssignTileHwndsToMonitors(selectedItems, maxPerMon := WM_TILE_BG_MAX_PER_MON) {
    monCount := MonitorGetCount()
    assignment := Map()
    loop monCount
        assignment[A_Index] := []
    unassigned := []
    ; Balance across monitors: always place on the monitor with the fewest planned windows (fills sparse monitors first).
    for item in selectedItems {
        bestMon := 0
        bestLen := 9999
        loop monCount {
            mon := A_Index
            n := assignment[mon].Length
            if (n < maxPerMon && n < bestLen) {
                bestLen := n
                bestMon := mon
            }
        }
        if (bestMon >= 1)
            assignment[bestMon].Push(item.hwnd)
        else
            unassigned.Push(item.hwnd)
    }
    assigned := 0
    loop monCount
        assigned += assignment[A_Index].Length
    return { plan: assignment, assigned: assigned, unassigned: unassigned }
}

WM_RepositionHwndToMonitor(mon, hwnd) {
    if (!hwnd || mon < 1 || !WinExist("ahk_id " hwnd))
        return false
    if (!WM_PrepareHwndForTile(hwnd))
        return false
    MonitorGetWorkArea mon, &left, &top, &right, &bottom
    margin := 8
    workLeft := left + margin
    workTop := top + margin
    workW := (right - left) - margin * 2
    workH := (bottom - top) - margin * 2
    if (workW < 120 || workH < 80)
        return false
    w := Max(200, workW // 4)
    h := Max(150, workH // 4)
    x := workLeft + (workW - w) // 2
    y := workTop + (workH - h) // 2
    return WM_MoveHwndToRect(hwnd, x, y, w, h)
}

WM_TileHwndsOnMonitorWorkArea(mon, hwnds, &tiledMap := unset) {
    if (!hwnds.Length || mon < 1)
        return 0
    MonitorGetWorkArea mon, &left, &top, &right, &bottom
    margin := 8
    gap := 4
    workLeft := left + margin
    workTop := top + margin
    workW := (right - left) - margin * 2
    workH := (bottom - top) - margin * 2
    if (workW < 120 || workH < 80)
        return 0
    n := hwnds.Length
    portrait := WM_MonitorIsPortrait(mon)
    tiled := 0
    if (n = 1) {
        if (WM_PrepareHwndForTile(hwnds[1])) {
            WM_MaximizeHwnd(hwnds[1])
            if (IsSet(tiledMap))
                tiledMap[hwnds[1]] := true
            tiled := 1
        }
        return tiled
    }
    if (portrait) {
        ; Portrait: full-width bands stacked (wide, short tiles).
        rowH := (workH - gap * (n - 1)) // n
        if (rowH < 80)
            return 0
        loop Min(n, 3) {
            i := A_Index
            y := workTop + (i - 1) * (rowH + gap)
            if (WM_PrepareHwndForTile(hwnds[i]) && WM_MoveHwndToRect(hwnds[i], workLeft, y, workW, rowH)) {
                if (IsSet(tiledMap))
                    tiledMap[hwnds[i]] := true
                tiled++
            }
        }
        return tiled
    }
    if (n = 2) {
        halfW := (workW - gap) // 2
        loop 2 {
            i := A_Index
            x := workLeft + (i = 2 ? halfW + gap : 0)
            if (WM_PrepareHwndForTile(hwnds[i]) && WM_MoveHwndToRect(hwnds[i], x, workTop, halfW, workH)) {
                if (IsSet(tiledMap))
                    tiledMap[hwnds[i]] := true
                tiled++
            }
        }
        return tiled
    }
    colW := (workW - gap * 2) // 3
    loop Min(n, 3) {
        i := A_Index
        x := workLeft + (i - 1) * (colW + gap)
        if (WM_PrepareHwndForTile(hwnds[i]) && WM_MoveHwndToRect(hwnds[i], x, workTop, colW, workH)) {
            if (IsSet(tiledMap))
                tiledMap[hwnds[i]] := true
            tiled++
        }
    }
    return tiled
}

WM_TileBackgroundExecutePlan(plan, &tiledMap) {
    totalTiled := 0
    monitorsTiled := 0
    loop MonitorGetCount() {
        mon := A_Index
        hwndList := plan.Has(mon) ? plan[mon] : []
        if (hwndList.Length = 0)
            continue
        for hwnd in hwndList {
            if (WM_GetHwndMonitorIndex(hwnd) != mon)
                WM_RepositionHwndToMonitor(mon, hwnd)
        }
        Sleep 40
        count := WM_TileHwndsOnMonitorWorkArea(mon, hwndList, &tiledMap)
        if (count > 0) {
            totalTiled += count
            monitorsTiled++
        }
    }
    return { totalTiled: totalTiled, monitorsTiled: monitorsTiled }
}

WM_TileBackgroundQualityLogPath() {
    return A_ScriptDir "\.cursor\wm_tile_quality.log"
}

WM_TileBackgroundQualityCheck(eligible, selected, assignResult, &tiledMap) {
    planned := assignResult.assigned
    tiled := 0
    for , ok in tiledMap {
        if (ok)
            tiled++
    }
    failed := []
    plannedHwnds := Map()
    for mon, list in assignResult.plan {
        for hwnd in list
            plannedHwnds[hwnd] := mon
    }
    for hwnd in plannedHwnds {
        if (!tiledMap.Has(hwnd) || !tiledMap[hwnd])
            failed.Push(hwnd)
    }
    unassigned := assignResult.unassigned.Length
    issues := []
    if (unassigned > 0)
        issues.Push("unassigned:" unassigned)
    if (planned < selected.Length)
        issues.Push("assign_short:" (selected.Length - planned))
    if (tiled < planned)
        issues.Push("tile_short:" (planned - tiled))
    ok := (unassigned = 0 && planned = selected.Length && tiled = planned)
    stats := Map(
        "eligible", eligible.total,
        "hidden", eligible.hidden,
        "visible", eligible.visible,
        "openOnMon", eligible.openOnMon,
        "skippedReject", eligible.skippedReject,
        "skippedNoMon", eligible.skippedNoMon,
        "selected", selected.Length,
        "planned", planned,
        "tiled", tiled,
        "unassigned", unassigned,
        "failed", failed.Length,
        "issues", issues,
        "ok", ok)
    global g_WM_LastTileQualityStats
    g_WM_LastTileQualityStats := stats
    return { ok: ok, failed: failed, plannedHwnds: plannedHwnds, stats: stats }
}

WM_TileBackgroundWriteQualityLog(eligible, selected, assignResult, qc, passLabel := "") {
    try DirCreate(A_ScriptDir "\.cursor")
    catch {
    }
    lines := ["=== WM tile quality " A_Now (passLabel != "" ? " " passLabel : "") " ==="]
    st := qc.stats
    lines.Push(Format("eligible={} hidden={} visible={} openOnMon={} skipReject={} skipNoMon={}",
        st["eligible"], st["hidden"], st["visible"], st["openOnMon"], st["skippedReject"], st["skippedNoMon"]))
    lines.Push(Format("selected={} planned={} tiled={} unassigned={} failed={}",
        st["selected"], st["planned"], st["tiled"], st["unassigned"], st["failed"]))
    if (st["issues"].Length)
        lines.Push("issues: " WM_ArrJoin(st["issues"], ", "))
    for hwnd in qc.failed {
        title := ""
        try title := WinGetTitle(hwnd)
        catch {
        }
        mon := qc.plannedHwnds.Has(hwnd) ? qc.plannedHwnds[hwnd] : "?"
        lines.Push(Format("  FAIL hwnd={} mon={} title={}", hwnd, mon, WM_TruncateTitleForList(title, 60)))
    }
    for hwnd in assignResult.unassigned {
        title := ""
        try title := WinGetTitle(hwnd)
        catch {
        }
        lines.Push(Format("  UNASSIGNED hwnd={} title={}", hwnd, WM_TruncateTitleForList(title, 60)))
    }
    path := WM_TileBackgroundQualityLogPath()
    try {
        if FileExist(path)
            FileAppend(WM_ArrJoin(lines, "`n") "`n", path, "UTF-8")
        else
            FileAppend(WM_ArrJoin(lines, "`n") "`n", path, "UTF-8")
    } catch {
    }
}

WM_TileBackgroundWindowsPerMonitor(maxPerMon := WM_TILE_BG_MAX_PER_MON, foreHwndOverride := 0) {
    WM_BackgroundTitleExcludes_Init()
    eligible := WM_CollectTileEligibleHwnds(foreHwndOverride)
    if (eligible.total = 0) {
        WM_CollectBackgroundWindows(foreHwndOverride)
        return { ok: false, noBackground: true, message: WM_FormatBackgroundCollectEmptyMessage() }
    }
    maxSlots := MonitorGetCount() * maxPerMon
    limit := Min(eligible.total, WM_TILE_BG_MAX_TOTAL, maxSlots)
    selected := WM_SelectTileHwndsByPriority(eligible.candidates, limit)
    assignResult := WM_AssignTileHwndsToMonitors(selected, maxPerMon)
    plan := assignResult.plan
    WMAutomation_SuppressCursorCentering("tile_background", 5000)
    tiledMap := Map()
    exec := WM_TileBackgroundExecutePlan(plan, &tiledMap)
    totalTiled := exec.totalTiled
    monitorsTiled := exec.monitorsTiled
    qc := WM_TileBackgroundQualityCheck(eligible, selected, assignResult, &tiledMap)
    if (!qc.ok && qc.failed.Length > 0) {
        reserve := WM_TileBackgroundReserveCandidates(eligible, selected, foreHwndOverride)
        replaced := WM_TileBackgroundReplaceFailedWithReserve(&selected, qc.failed, reserve)
        if (replaced > 0) {
            assignResult := WM_AssignTileHwndsToMonitors(selected, maxPerMon)
            plan := assignResult.plan
            tiledMap := Map()
            exec := WM_TileBackgroundExecutePlan(plan, &tiledMap)
            totalTiled := 0
            for , ok in tiledMap {
                if (ok)
                    totalTiled++
            }
            monitorsTiled := 0
            for mon, list in plan {
                for hwnd in list {
                    if (tiledMap.Has(hwnd) && tiledMap[hwnd]) {
                        monitorsTiled++
                        break
                    }
                }
            }
            qc := WM_TileBackgroundQualityCheck(eligible, selected, assignResult, &tiledMap)
            WM_TileBackgroundWriteQualityLog(eligible, selected, assignResult, qc, "backfill")
        }
    }
    WMAutomation_ClearCursorSuppression("tile_background")
    if (WM_DebugBackgroundEnabled() || !qc.ok)
        WM_TileBackgroundWriteQualityLog(eligible, selected, assignResult, qc, "final")
    if (totalTiled = 0)
        return { ok: false, message: "Could not tile background windows (" eligible.total " eligible)." }
    planned := assignResult.assigned
    msg := (monitorsTiled = 1)
        ? ("Tiled " totalTiled "/" planned " on 1 monitor (" eligible.total " eligible)")
        : ("Tiled " totalTiled "/" planned " on " monitorsTiled " monitors (" eligible.total " eligible)")
    if (!qc.ok) {
        msg .= " ⚠️ QC: "
        if (qc.stats["unassigned"] > 0)
            msg .= qc.stats["unassigned"] " unassigned "
        if (qc.stats["failed"] > 0)
            msg .= qc.stats["failed"] " failed"
        msg .= " — see wm_tile_quality.log"
    } else if (eligible.total > totalTiled) {
        msg .= " (" (eligible.total - totalTiled) " not selected, cap " WM_TILE_BG_MAX_TOTAL ")"
    }
    return { ok: true, message: msg, quality: qc.stats }
}
