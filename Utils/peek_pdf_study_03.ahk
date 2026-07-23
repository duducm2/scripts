; =============================================================================
; Utils module: peek_pdf_study_03.ahk
; Peek PDF / QuickLook study helpers (part 3)
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; Map a Windows display device name (e.g. "\\.\DISPLAY2") to the current AHK monitor index.
; Returns 0 when not found/connected.
GetMonitorIndexByDeviceName(deviceName) {
    if (deviceName = "")
        return 0
    try monitorCount := MonitorGetCount()
    catch
        return 0
    loop monitorCount {
        idx := A_Index
        nm := ""
        try nm := MonitorGetName(idx)
        if (nm = deviceName)
            return idx
    }
    return 0
}

; Poll until QuickLook.exe has a window (cold start can exceed 2s). Returns hwnd or 0.
QuickLook_WaitForHwnd(timeoutMs := 10000) {
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        hwnd := WinExist("ahk_exe QuickLook.exe")
        if (hwnd)
            return hwnd
        Sleep 75
    }
    return 0
}

QuickLook_FindDocumentElement(hwnd) {
    if (!hwnd)
        return 0
    try {
        root := UIA.ElementFromHandle(hwnd)
        doc := root.FindFirst({ Type: UIA.ControlType.Document })
        if (doc)
            return doc
    } catch {
    }
    return 0
}

; UIA Document present and enabled on two consecutive polls (WebView settled).
QuickLook_WaitForViewerReady(hwnd, timeoutMs := 3000) {
    deadline := A_TickCount + timeoutMs
    stableCount := 0
    while (A_TickCount < deadline) {
        if (!WinExist("ahk_id " hwnd))
            return false
        doc := QuickLook_FindDocumentElement(hwnd)
        ok := false
        if (doc) {
            try {
                ok := doc.GetPropertyValue(UIA.Property.IsEnabled)
            } catch {
                ok := true
            }
        }
        if (ok) {
            stableCount++
            if (stableCount >= 2)
                return true
            Sleep 100
        } else {
            stableCount := 0
            Sleep 75
        }
    }
    return false
}

QuickLook_ClickWindowCenter(hwnd) {
    if (!hwnd)
        return
    try {
        rect := Buffer(16, 0)
        if (DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect)) {
            wl := NumGet(rect, 0, "int"), wt := NumGet(rect, 4, "int")
            wr := NumGet(rect, 8, "int"), wb := NumGet(rect, 12, "int")
            CoordMode("Mouse", "Screen")
            Click wl + (wr - wl) // 2, wt + (wb - wt) // 2
        }
    } catch {
    }
}

QuickLook_ScrollViaUIA(hwnd) {
    doc := QuickLook_FindDocumentElement(hwnd)
    if (!doc)
        return false
    try {
        if (doc.GetPropertyValue(UIA.Property.IsScrollPatternAvailable)) {
            doc.ScrollPattern.SetScrollPercent(-1, 100)
            return true
        }
    } catch {
    }
    return false
}

QuickLook_GetScrollVerticalPercent(hwnd) {
    doc := QuickLook_FindDocumentElement(hwnd)
    if (!doc)
        return -1.0
    try {
        if (doc.GetPropertyValue(UIA.Property.IsScrollPatternAvailable)) {
            p := doc.GetPropertyValue(UIA.Property.ScrollVerticalScrollPercent)
            if (p >= 0)
                return p
        }
    } catch {
    }
    return -1.0
}

QuickLook_ScrollViaKeystroke(hwnd) {
    try {
        WinActivate("ahk_id " hwnd)
        WinWaitActive("ahk_id " hwnd, , 1)
        SendInput("^{End}")
        return true
    } catch {
        try {
            WinActivate("ahk_id " hwnd)
            WinWaitActive("ahk_id " hwnd, , 1)
            Send("{Ctrl down}{End}{Ctrl up}")
            return true
        } catch {
            try {
                ControlSend("^End", "ahk_id " hwnd)
                return true
            } catch {
                return false
            }
        }
    }
}

; UIA scroll first, then Ctrl+End; verify vertical % when available (extraRetries for slow title gate).
QuickLook_ScrollToEnd(hwnd, extraRetries := 0) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return
    maxAttempts := 3 + extraRetries
    deadline := A_TickCount + 2000
    loop maxAttempts {
        if (A_TickCount >= deadline)
            break
        QuickLook_ScrollViaUIA(hwnd)
        pct := QuickLook_GetScrollVerticalPercent(hwnd)
        if (pct >= 95.0)
            return
        QuickLook_ScrollViaKeystroke(hwnd)
        pct := QuickLook_GetScrollVerticalPercent(hwnd)
        if (pct >= 95.0)
            return
        if (A_Index < maxAttempts)
            Sleep 150
    }
}

; Focus QuickLook, optional scroll-to-end, then schedule deferred AutoSlot place
; (QL often resets size after document paint — immediate place does not stick).
QuickLook_ApplyStudyLayout(hwnd, scrollToEnd := true, extraScrollRetries := 0) {
    if (!hwnd || !WinExist("ahk_id " hwnd))
        return false
    try {
        WinShow("ahk_id " hwnd)
        WinActivate("ahk_id " hwnd)
        WinWaitActive("ahk_id " hwnd, , 1)
    } catch {
    }
    QuickLook_ClickWindowCenter(hwnd)
    if (scrollToEnd)
        QuickLook_ScrollToEnd(hwnd, extraScrollRetries)
    try QuickLook_ScheduleAutoSlotPlace()
    catch {
    }
    return true
}

QuickLook_DeferredLayoutAfterStart(*) {
    global g_QuickLookDeferredLayoutScroll, g_QuickLookDeferredLayoutPath
    scrollToEnd := g_QuickLookDeferredLayoutScroll
    path := g_QuickLookDeferredLayoutPath
    g_QuickLookDeferredLayoutPath := ""
    hwnd := WinExist("ahk_exe QuickLook.exe")
    if (!hwnd)
        return
    extra := 0
    if (path != "") {
        gate := QuickLook_WaitForOpenReady(hwnd, path, 6000)
        if (!gate["ok"])
            return
        if (STUDY_TOPIC_QL_STRICT_LAYOUT)
            QuickLook_WaitForViewerReady(hwnd, 4000)
        extra := gate["fallback"] ? 2 : 0
    }
    QuickLook_ApplyStudyLayout(hwnd, scrollToEnd, extra)
}

; Legacy open path (STUDY_TOPIC_QL_STRICT_LAYOUT := false).
QuickLook_OpenPath_Legacy(path, scrollToEnd := true) {
    if WinWait("ahk_exe QuickLook.exe", , 2) {
        hwnd := WinExist("ahk_exe QuickLook.exe")
        if (hwnd) {
            gate := QuickLook_WaitForOpenReady(hwnd, path)
            if (!gate["ok"]) {
                try ShowCenteredOverlay_Utils("⚠ QuickLook closed before the file finished loading.", 3200,
                    BANNER_ACCENT_INTERMEDIATE)
                return
            }
            if (gate["matched"])
                Sleep 75
            else
                Sleep 50
            gateUsedFallback := gate["fallback"]
            try {
                WinShow("ahk_id " hwnd)
                WinActivate("ahk_id " hwnd)
                WinWaitActive("ahk_id " hwnd, , 1)
            } catch {
            }
            QuickLook_ClickWindowCenter(hwnd)
            if (scrollToEnd)
                QuickLook_ScrollToEnd(hwnd, gateUsedFallback ? 2 : 0)
        }
    }
}

; Wait until QuickLook's title reflects the opened file (QL-Win shows basename) or timeout with graceful fallback.
; Returns Map: ok (hwnd still valid), matched (title contained basename), fallback (timed out; caller may retry scroll).
QuickLook_WaitForOpenReady(hwnd, path, timeoutMs := 8000) {
    SplitPath path, &baseName
    if (baseName = "")
        return Map("ok", false, "matched", false, "fallback", false)
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        if (!WinExist("ahk_id " hwnd))
            return Map("ok", false, "matched", false, "fallback", false)
        try {
            title := WinGetTitle("ahk_id " hwnd)
        } catch {
            Sleep 75
            continue
        }
        tL := StrLower(title)
        if InStr(tL, StrLower(baseName))
            return Map("ok", true, "matched", true, "fallback", false)
        ; Windows sometimes truncates long titles — try first 31 chars of basename
        if (StrLen(baseName) > 31) {
            shortNeedle := StrLower(SubStr(baseName, 1, 31))
            if InStr(tL, shortNeedle)
                return Map("ok", true, "matched", true, "fallback", false)
        }
        Sleep 75
    }
    ; Option A: proceed after extra delay when title never matched (atypical QuickLook build)
    Sleep 300
    if (!WinExist("ahk_id " hwnd))
        return Map("ok", false, "matched", false, "fallback", false)
    return Map("ok", true, "matched", false, "fallback", true)
}

; scrollToEnd: after open, focus viewer and send Ctrl+End (mnemonics); false leaves viewport at top (plans).
QuickLook_OpenPath(path, scrollToEnd := true) {
    quickLookExe := QuickLook_ResolveExePath()
    if (!FileExist(quickLookExe)) {
        try ShowCenteredOverlay_Utils("❌ QuickLook executable not found: " quickLookExe, 2500, BANNER_ACCENT_ERROR)
        return
    }
    if (!FileExist(path)) {
        try ShowCenteredOverlay_Utils("❌ Markdown not found: " path, 3500, BANNER_ACCENT_ERROR)
        return
    }
    try {
        Run('"' quickLookExe '" "' path '"')
    } catch as e {
        try ShowCenteredOverlay_Utils("❌ Failed to open QuickLook: " e.Message, 3000, BANNER_ACCENT_ERROR)
        return
    }
    if (!STUDY_TOPIC_QL_STRICT_LAYOUT) {
        QuickLook_OpenPath_Legacy(path, scrollToEnd)
        return
    }
    hwnd := QuickLook_WaitForHwnd(10000)
    if (!hwnd) {
        try ShowCenteredOverlay_Utils("⚠ QuickLook did not start in time.", 3200, BANNER_ACCENT_INTERMEDIATE)
        global g_QuickLookDeferredLayoutScroll, g_QuickLookDeferredLayoutPath
        g_QuickLookDeferredLayoutScroll := scrollToEnd
        g_QuickLookDeferredLayoutPath := path
        SetTimer(QuickLook_DeferredLayoutAfterStart, -800)
        return
    }
    gate := QuickLook_WaitForOpenReady(hwnd, path)
    if (!gate["ok"]) {
        try ShowCenteredOverlay_Utils("⚠ QuickLook closed before the file finished loading.", 3200,
            BANNER_ACCENT_INTERMEDIATE)
        return
    }
    if (gate["matched"])
        Sleep 75
    else
        Sleep 50
    QuickLook_WaitForViewerReady(hwnd, 3000)
    extraScroll := gate["fallback"] ? 2 : 0
    QuickLook_ApplyStudyLayout(hwnd, scrollToEnd, extraScroll)
}

; Open a specific PDF in PowerToys Peek and run WaitAndConfigure. Caller must validate pdfPath and exe exist.
; skipGoToLastPage: if true, do not navigate to the last page (e.g. for short docs like technique README).
PeekPdf_OpenPath(pdfPath, skipGoToLastPage := false) {
    peekExe := PeekPdf_ResolvePeekExePath()
    peekEsc := StrReplace(peekExe, "'", "''")
    pdfEsc := StrReplace(pdfPath, "'", "''")
    psArg := "& " . Chr(39) . peekEsc . Chr(39) . " " . Chr(39) . pdfEsc . Chr(39)
    cmd := "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command " . Chr(34) . psArg . Chr(34)
    try Run cmd, "", "Hide"
    catch as e {
        try ShowCenteredOverlay_Utils("❌ Failed to open Peek: " e.Message, 3000, BANNER_ACCENT_ERROR)
        return
    }
    if WinWait("Peek", "", 5) {
        hwnd := WinExist("ahk_exe PowerToys.Peek.UI.exe")
        if (!hwnd)
            hwnd := WinExist("Peek")
        MoveWindowToMonitor(hwnd, 2)
        WinMaximize("ahk_id " hwnd)
        PeekPdf_WaitAndConfigure(skipGoToLastPage)
        try {
            WinActivate("ahk_id " hwnd)
            WinWaitActive("ahk_id " hwnd, , 1)
        } catch {
        }
    }
}

PeekPdf_OpenStored() {
    iniPath := PeekPdf_GetIniPath()

    ; Resolve PDF path per environment, with legacy fallback
    pdfPath := ""
    try {
        if (IS_WORK_ENVIRONMENT) {
            pdfPath := IniRead(iniPath, "Peek", "PdfPathWork", "")
            if (pdfPath = "")
                pdfPath := IniRead(iniPath, "Peek", "PdfPath", "")
        } else {
            pdfPath := IniRead(iniPath, "Peek", "PdfPathPersonal", "")
            if (pdfPath = "")
                pdfPath := IniRead(iniPath, "Peek", "PdfPath", "")
        }
    }

    pdfPath := PeekPdf_NormalizePath(pdfPath)
    if (pdfPath = "") {
        try ShowCenteredOverlay_Utils("⚠ No PDF path set. Hold Win+Alt+Shift+X to set.", 3000,
            BANNER_ACCENT_INTERMEDIATE)
        return
    }
    peekExe := PeekPdf_ResolvePeekExePath()
    if (!FileExist(peekExe)) {
        try ShowCenteredOverlay_Utils("❌ Peek executable not found.", 2500, BANNER_ACCENT_ERROR)
        return
    }
    if (!FileExist(pdfPath)) {
        try ShowCenteredOverlay_Utils("❌ PDF file not found: " pdfPath, 3500, BANNER_ACCENT_ERROR)
        return
    }
    peekEsc := StrReplace(peekExe, "'", "''")
    pdfEsc := StrReplace(pdfPath, "'", "''")
    psArg := "& " . Chr(39) . peekEsc . Chr(39) . " " . Chr(39) . pdfEsc . Chr(39)
    cmd := "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command " . Chr(34) . psArg . Chr(34)
    try Run cmd, "", "Hide"
    catch as e {
        try ShowCenteredOverlay_Utils("❌ Failed to open Peek: " e.Message, 3000, BANNER_ACCENT_ERROR)
        return
    }
    if WinWait("Peek", "", 5) {
        hwnd := WinExist("ahk_exe PowerToys.Peek.UI.exe")
        if (!hwnd)
            hwnd := WinExist("Peek")
        MoveWindowToMonitor(hwnd, 2)
        WinMaximize("ahk_id " hwnd)
        PeekPdf_WaitAndConfigure()
    }
}

MoveWindowToMonitor(hwnd, monitorIndex := 2) {
    if (!hwnd)
        return
    try {
        MonitorGetWorkArea(monitorIndex, &l, &t, &r, &b)
    } catch {
        return
    }
    w := r - l
    h := b - t
    ; Restore before moving, otherwise some apps "teleport" as a 1px/tiny bar.
    try {
        mm := WinGetMinMax("ahk_id " hwnd) ; 1=min,2=max,0=normal
        if (mm != 0) {
            WinRestore("ahk_id " hwnd)
            Sleep 80
        }
    } catch {
    }
    try WinMove(l, t, w, h, "ahk_id " hwnd)
}

; Maximize by hwnd with WinAPI fallback (more reliable than keystrokes).
TryMaximizeWindow(hwnd) {
    if (!hwnd)
        return false
    try {
        WinMaximize("ahk_id " hwnd)
        return true
    } catch {
        try {
            PostMessage 0x0112, 0xF030, , , "ahk_id " hwnd  ; WM_SYSCOMMAND, SC_MAXIMIZE
            return true
        } catch {
            return false
        }
    }
}

; Map a window handle to an AutoHotkey monitor index (1..MonitorGetCount()).
; Needed because MonitorFromWindow returns an HMONITOR handle, not an AHK index.
GetAhkMonitorIndexFromHwnd(hwnd) {
    if (!hwnd)
        return 0
    hMon := 0
    try hMon := DllCall("user32\MonitorFromWindow", "ptr", hwnd, "uint", 2, "ptr") ; MONITOR_DEFAULTTONEAREST
    catch
        return 0
    if (!hMon)
        return 0

    count := 0
    try count := MonitorGetCount()
    catch
        return 0
    if (count < 1)
        return 0
    loop count {
        i := A_Index
        try MonitorGet i, &l, &t, &r, &b
        catch
            continue
        cx := (l + r) // 2
        cy := (t + b) // 2
        ; MonitorFromPoint takes a POINT passed by value (two 32-bit signed ints packed into an int64).
        pt64 := ((cy & 0xFFFFFFFF) << 32) | (cx & 0xFFFFFFFF)
        cur := 0
        try cur := DllCall("user32\MonitorFromPoint", "int64", pt64, "uint", 2, "ptr") ; MONITOR_DEFAULTTONEAREST
        catch
            cur := 0
        if (cur = hMon)
            return i
    }
    return 0
}

; Wait for Peek PDF toolbar to load (Page view button), click it, two-page view, focus; optionally go to last page.
; Current state: PDF opening and window maximization are working correctly.
; Execution order: 1) Get Peek hwnd  2) UIA root from hwnd  3) Poll for "Page view" anchor
;  4) Wait for anchor visible + short delay before click  5) Click Page view  6) Right Arrow
;  7) Click window center  8) If not skipGoToLastPage: go to last page (UIA or Ctrl+End). Fallback: Sleep 400 + Click if UIA or anchor fails.
PeekPdf_WaitAndConfigure(skipGoToLastPage := false) {
    global UIA
    ; Standard loading bar: show for the whole process so user knows when we started and when we finished
    StandardLoadingBar_Show("⏳ Peek PDF: configuring...", BANNER_ACCENT_INTERMEDIATE)
    ; 1) Get Peek window hwnd
    hwnd := WinExist("Peek")
    if (!hwnd)
        hwnd := WinExist("ahk_exe PowerToys.Peek.UI.exe")
    if (!hwnd) {
        StandardLoadingBar_Update("❌ Peek PDF: window not found", BANNER_ACCENT_ERROR)
        StandardLoadingBar_Hide(2000)
        Sleep 300
        Click "Left"
        return
    }
    try {
        ; 2) UIA root
        el := UIA.ElementFromHandle(hwnd)
        ; 3) Poll for Page view (layouts) anchor (up to 20s to accommodate Peek load > 5s)
        ; FindFirst throws when no element matches; catch so we keep polling instead of exiting.
        pageViewBtn := ""
        pollIter := 0
        loop 80 {
            pollIter := A_Index
            try
                pageViewBtn := el.FindFirst({ Type: 50000, Name: "Page view", AutomationId: "layouts" })
            catch
                pageViewBtn := ""
            if (pageViewBtn)
                break
            Sleep 150
        }
        if (pageViewBtn) {
            ; 4) Ensure anchor is visible, then short delay so toolbar is ready before click
            visIter := 0
            loop 20 {
                visIter := A_Index
                try {
                    if (!pageViewBtn.GetPropertyValue(UIA.Property.IsOffscreen)) {
                        br := pageViewBtn.BoundingRectangle
                        if (IsObject(br) && (br.r - br.l) > 0 && (br.b - br.t) > 0)
                            break
                    }
                } catch {
                }
                Sleep 30
            }
            ; Short delay so toolbar is ready before clicking Page view
            Sleep 1000
            ; 5) Click Page view button
            invokeOk := false
            clickOk := false
            try {
                pageViewBtn.Invoke()
                invokeOk := true
            } catch as invErr {
                try {
                    pageViewBtn.Click()
                    clickOk := true
                } catch as clickErr {
                }
            }
            Sleep 600
            ; 6) Select "Two page" from the open Page view menu (main window + foreground popup; else ControlSend Right to Peek)
            fgHwnd := 0
            try fgHwnd := WinGetID("A")

            twoPageEl := ""
            twoPageClicked := false

            twoPageScope := "none"
            menuRect := ""
            try {
                abr := pageViewBtn.BoundingRectangle
                if (IsObject(abr))
                    menuRect := { l: abr.l - 700, t: abr.b, r: abr.r + 700, b: abr.b + 650 }
            } catch {
                menuRect := ""
            }

            ; Search in active window region (menu is visible on screen but may not be exposed as Buttons).
            try {
                elActive := UIA.ElementFromHandle(fgHwnd ? fgHwnd : hwnd)
                if (IsObject(menuRect)) {
                    for cand in elActive.FindAll({ IsOffscreen: 0 }) {
                        try {
                            br := cand.BoundingRectangle
                            if (!IsObject(br))
                                continue
                            inRegion := (br.l < menuRect.r && br.r > menuRect.l && br.t < menuRect.b && br.b > menuRect
                                .t)
                            if (!inRegion)
                                continue
                            nm := ""
                            try nm := cand.Name
                            if (nm != "" && InStr(nm, "Two page")) {
                                twoPageEl := cand
                                twoPageScope := (fgHwnd ? "active_region" : "peek_region")
                                break
                            }
                        } catch {
                        }
                    }
                }
            } catch {
            }

            if (twoPageEl) {
                try {
                    ; Click by coordinates for maximum compatibility (works even if Invoke/Click patterns are absent).
                    br := twoPageEl.BoundingRectangle
                    if (IsObject(br)) {
                        cx := br.l + (br.r - br.l) // 2
                        cy := br.t + (br.b - br.t) // 2
                        Click cx, cy
                        twoPageClicked := true
                    }
                } catch {
                    twoPageClicked := false
                }
            } else {
                ; Last-resort: keystroke fallback (kept for robustness)
                try ControlSend "{Right}", "ahk_id " hwnd
                catch
                    Send "{Right}"
            }

            Sleep 400
            ; 7) Click center of Peek window (focus)
            WinGetPos(&wx, &wy, &ww, &wh, "ahk_id " hwnd)
            if (ww > 0 && wh > 0) {
                cx := wx + ww // 2
                cy := wy + wh // 2
                Click cx, cy
            }
            Sleep 150
            ; 8) Go to final page via UIA (keystrokes do not reach the embedded Edge PDF viewer); skip if skipGoToLastPage
            if (!skipGoToLastPage) {
                lastPageSet := false
                try {
                    for doc in el.FindAll({ Type: 50030 }) {
                        try {
                            nm := doc.Name
                            ; Match "containing N pages" (EN) or "N pages"/"N páginas" (avoid "Page 1")
                            if (RegExMatch(nm, "containing\s+(\d+)\s+pages", &m) || RegExMatch(nm,
                                "document.*?(\d+)\s*(?:pages|páginas)", &m)) {
                                totalPages := Integer(m[1])
                                if (totalPages > 0) {
                                    pageSel := el.FindFirst({ Type: 50004, AutomationId: "pageselector" })
                                    if (pageSel) {
                                        try {
                                            pageSel.SetFocus()
                                            Sleep(200)
                                            WinActivate("ahk_id " hwnd)
                                            Sleep(120)
                                            Send("^a")
                                            Sleep(50)
                                            Send(String(totalPages))
                                            Sleep(50)
                                            Send("{Enter}")
                                            lastPageSet := true
                                        } catch {
                                            ; ignore focus/send errors
                                        }
                                    }
                                    break
                                }
                            }
                        } catch {
                            ; ignore per-doc errors
                        }
                    }
                } catch {
                    ; ignore UIA errors for last-page navigation
                }
                if (!lastPageSet) {
                    try
                        ControlSend("^End", "ahk_id " hwnd)
                    catch
                        Send("^End")
                }
            }
            Sleep 100
            StandardLoadingBar_Update("✅ Peek PDF: done", BANNER_ACCENT_SUCCESS)
            StandardLoadingBar_Hide(2000)
        } else {
            StandardLoadingBar_Update("✅ Peek PDF: finished (fallback)", BANNER_ACCENT_SUCCESS)
            StandardLoadingBar_Hide(2000)
            Sleep 400
            Click "Left"
        }
    } catch {
        StandardLoadingBar_Update("✅ Peek PDF: finished (fallback)", BANNER_ACCENT_SUCCESS)
        StandardLoadingBar_Hide(2000)
        Sleep 400
        Click "Left"
    }
}
