; =============================================================================
; Shift keys module: hotif_google_maps.ahk
; Google Maps Chrome hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

; On while diagnosing Maps capture. Each Shift+P replaces the environment log under assets/data/.
global MAPS_CAPTURE_DEBUG := true
global g_MapsDebugSession := ""
global g_MapsDebugSeq := 0
global g_MapsDebugMethod := ""
global g_MapsDebugAttempt := 0
global g_MapsDebugKept := Map()

Maps_DebugEnabled() {
    global MAPS_CAPTURE_DEBUG
    return IsSet(MAPS_CAPTURE_DEBUG) && MAPS_CAPTURE_DEBUG
}

Maps_DebugIsWork() {
    global IS_WORK_ENVIRONMENT
    try {
        return IsSet(IS_WORK_ENVIRONMENT) && IS_WORK_ENVIRONMENT
    } catch {
        return false
    }
}

; Tracked file so a work run can be pushed and pulled. Same split as quick_update_debug_*.txt.
Maps_DebugLogPath() {
    name := Maps_DebugIsWork() ? "maps_capture_debug_work.txt" : "maps_capture_debug_personal.txt"
    return A_ScriptDir "\assets\data\" name
}

Maps_DebugBegin() {
    global g_MapsDebugSession, g_MapsDebugSeq, g_MapsDebugMethod, g_MapsDebugAttempt, g_MapsDebugKept
    g_MapsDebugSession := A_Now
    g_MapsDebugSeq := 0
    g_MapsDebugMethod := ""
    g_MapsDebugAttempt := 0
    g_MapsDebugKept := Map()
    if !Maps_DebugEnabled()
        return
    path := Maps_DebugLogPath()
    try DirCreate(A_ScriptDir "\assets\data")
    catch {
    }
    try FileDelete(path)
    catch {
    }
    Maps_DebugLog("machine", Map(
        "computer", A_ComputerName,
        "user", A_UserName,
        "work", Maps_DebugIsWork() ? "yes" : "no",
        "scriptDir", A_ScriptDir,
        "log", path
    ))
}

Maps_DebugJsonEscape(s) {
    try {
        if (s = "")
            return ""
    } catch {
        return ""
    }
    s := "" s
    s := StrReplace(s, "\", "\\")
    s := StrReplace(s, '"', '\"')
    s := StrReplace(s, "`r", "\r")
    s := StrReplace(s, "`n", "\n")
    return s
}

Maps_DebugLog(step, data := "") {
    global g_MapsDebugSession, g_MapsDebugSeq
    if !Maps_DebugEnabled()
        return
    try {
        g_MapsDebugSeq += 1
        dataJson := "{}"
        if (IsObject(data)) {
            parts := []
            for k, v in data {
                try parts.Push('"' Maps_DebugJsonEscape(k) '":"' Maps_DebugJsonEscape(v) '"')
            }
            joined := ""
            if (parts.Length) {
                for i, p in parts
                    joined .= (i = 1 ? "" : ",") p
            }
            dataJson := "{" joined "}"
        } else if (data != "") {
            dataJson := '{"value":"' Maps_DebugJsonEscape(data) '"}'
        }
        line := '{'
            . '"session":"' Maps_DebugJsonEscape(g_MapsDebugSession) '",'
            . '"seq":' g_MapsDebugSeq ','
            . '"t":' A_TickCount ','
            . '"step":"' Maps_DebugJsonEscape(step) '",'
            . '"data":' dataJson
            . '}'
        FileAppend(line "`n", Maps_DebugLogPath(), "UTF-8")
    } catch {
    }
}

Maps_DebugFileFacts(path) {
    facts := Map("bytes", 0, "w", 0, "h", 0, "headerErr", "")
    imgW := 0
    imgH := 0
    bytes := 0
    err := ""
    try Maps_PngHeaderOk(path, &imgW, &imgH, &bytes, &err, 1, 1)
    catch {
        err := "capture file missing"
    }
    facts["bytes"] := bytes
    facts["w"] := imgW
    facts["h"] := imgH
    facts["headerErr"] := err
    return facts
}

; Copy a rejected PNG into debug-maps-capture/ before it is deleted. Same file is kept once.
Maps_DebugKeepPng(path, reason := "") {
    global g_MapsDebugMethod, g_MapsDebugAttempt, g_MapsDebugKept
    if !Maps_DebugEnabled()
        return ""
    if (path = "" || !FileExist(path))
        return ""
    try {
        sz := FileGetSize(path)
        modified := FileGetTime(path, "M")
        sig := StrLower(path) "|" sz "|" modified
        if g_MapsDebugKept.Has(sig)
            return g_MapsDebugKept[sig]
        dir := A_ScriptDir "\debug-maps-capture"
        DirCreate(dir)
        method := g_MapsDebugMethod != "" ? g_MapsDebugMethod : "capture"
        attempt := g_MapsDebugAttempt > 0 ? g_MapsDebugAttempt : 1
        name := FormatTime(, "yyyyMMdd-HHmmss") "-" method "-a" attempt ".png"
        dest := dir "\" name
        if FileExist(dest)
            dest := dir "\" FormatTime(, "yyyyMMdd-HHmmss") "-" method "-a" attempt "-" A_TickCount ".png"
        FileCopy(path, dest, 1)
        g_MapsDebugKept[sig] := dest
        Maps_DebugLog("keep_png", Map("src", path, "dest", dest, "reason", reason, "bytes", sz))
        return dest
    } catch {
        return ""
    }
}

Maps_DiscardCaptureFile(path, reason := "") {
    Maps_DebugKeepPng(path, reason)
    Maps_DeleteCaptureFile(path)
}

Maps_DebugHideMs() {
    return Maps_DebugEnabled() ? 8000 : 2200
}

Maps_GetDocumentRoot(uia) {
    root := 0
    try root := uia.GetCurrentDocumentElement()
    catch {
        try root := uia.BrowserElement
    }
    return root
}

Maps_CollapseSidePanel(root) {
    if !root
        return false
    btn := 0
    try btn := root.FindFirst({ Type: 50000, Name: "Collapse side panel", cs: false })
    if !btn {
        try btn := root.FindFirst({ Name: "Collapse side panel", cs: false })
    }
    if !btn
        return false
    try {
        if btn.GetPropertyValue(UIA.Property.IsInvokePatternAvailable)
            btn.InvokePattern.Invoke()
        else
            btn.Click()
    } catch {
        try btn.Click()
        catch {
            return false
        }
    }
    return true
}

Maps_HideChrome(uia) {
    ; Hide every sibling along the largest canvas ancestor chain; tag for restore.
    js :=
        "(function(){var best=null,ba=0;document.querySelectorAll('canvas').forEach(function(c){var a=(c.width||0)*(c.height||0);if(a>ba){ba=a;best=c;}});if(!best)return 0;var n=best;while(n&&n!==document.body){var p=n.parentElement;if(!p)break;for(var i=0;i<p.children.length;i++){var s=p.children[i];if(s!==n){s.style.setProperty('visibility','hidden','important');s.setAttribute('data-ahk-maps-hide','1');}}n=p;}return 1;})()"
    try {
        result := uia.JSReturnThroughClipboard(js)
        return (result = "1" || result = 1)
    } catch {
        try uia.JSExecute(js)
        catch {
            return false
        }
        return true
    }
}

Maps_RestoreChrome(uia) {
    js :=
        "(function(){document.querySelectorAll('[data-ahk-maps-hide]').forEach(function(el){el.style.removeProperty('visibility');el.removeAttribute('data-ahk-maps-hide');});return 1;})()"
    try uia.JSReturnThroughClipboard(js)
    catch {
        try uia.JSExecute(js)
        catch {
        }
    }
}

Maps_FindMapPane(root) {
    if !root
        return 0
    pane := 0
    try pane := root.FindFirst({ Type: 50033, Name: "Street View", cs: false })
    if pane
        return pane
    try {
        for el in root.FindAll({ Type: 50033 }) {
            n := el.Name
            if (n = "")
                continue
            if InStr(n, "Street View") || (SubStr(n, 1, 3) = "Map") {
                return el
            }
        }
    } catch {
    }
    return 0
}

Maps_MaximizeBrowserWindow(uia) {
    if !uia
        return false
    hwnd := 0
    try hwnd := uia.BrowserId
    if !hwnd
        return false
    try {
        if (WinGetMinMax("ahk_id " hwnd) = 1)
            return true
    } catch {
    }
    try {
        WinMaximize("ahk_id " hwnd)
    } catch {
        try PostMessage(0x0112, 0xF030, 0, 0, "ahk_id " hwnd)  ; WM_SYSCOMMAND, SC_MAXIMIZE
        catch {
            return false
        }
    }
    Sleep 500  ; maximize animation + map layout settle
    return true
}

Maps_LoadingOptions(hwnd := 0) {
    return { passive: false, centerOnHwnd: hwnd, textWidth: 680 }
}

; Update the current Loading Indication, or show it if a screen grab dismissed it.
Maps_Loading(msg, hwnd := 0) {
    global g_StandardLoadingBarGui
    if IsObject(g_StandardLoadingBarGui) {
        StandardLoadingBar_Update(msg, BANNER_ACCENT_INTERMEDIATE)
        return
    }
    StandardLoadingBar_Show(msg, BANNER_ACCENT_INTERMEDIATE, Maps_LoadingOptions(hwnd))
}

Maps_ShowSaved(name, hwnd := 0) {
    global g_StandardLoadingBarGui
    opts := Maps_LoadingOptions(hwnd)
    opts.passive := true
    opts.passiveBgColor := BANNER_ACCENT_SUCCESS
    StandardLoadingBar_Show("✅ Saved on Desktop: " name, BANNER_ACCENT_SUCCESS, opts)
    try {
        if IsObject(g_StandardLoadingBarGui) {
            barHwnd := g_StandardLoadingBarGui.Hwnd
            ; HWND_TOPMOST so the maximized map does not cover the banner.
            DllCall("SetWindowPos", "ptr", barHwnd, "ptr", -1, "int", 0, "int", 0, "int", 0, "int", 0, "uint", 0x0013)
        }
    } catch {
    }
    StandardLoadingBar_Hide(4500)
}

Maps_Fail(msg, hwnd := 0, hideMs := 2200) {
    global g_StandardLoadingBarGui
    if !IsObject(g_StandardLoadingBarGui)
        StandardLoadingBar_Show(msg, BANNER_ACCENT_ERROR, Maps_LoadingOptions(hwnd))
    StandardLoadingBar_Update(msg, BANNER_ACCENT_ERROR)
    StandardLoadingBar_Hide(hideMs)
}

Maps_DeleteCaptureFile(path) {
    if (path = "" || !FileExist(path))
        return
    try FileDelete(path)
    catch {
    }
}

Maps_CaptureSearchDirs() {
    dirs := []
    downloads := EnvGet("USERPROFILE") "\Downloads"
    if DirExist(downloads)
        dirs.Push(downloads)
    desktop := Palace_ResolveDesktopDir()
    if (desktop != "" && DirExist(desktop))
        dirs.Push(desktop)
    alt := RTrim(A_Desktop, "\")
    if (alt != "" && StrLower(alt) != StrLower(desktop) && DirExist(alt))
        dirs.Push(alt)
    return dirs
}

Maps_PurgeCaptureTemps() {
    for dir in Maps_CaptureSearchDirs() {
        loop files dir "\ahk-maps-cap*.png", "F" {
            Maps_DeleteCaptureFile(A_LoopFileFullPath)
        }
        loop files dir "\ahk-maps-cap*.crdownload", "F" {
            Maps_DeleteCaptureFile(A_LoopFileFullPath)
        }
    }
    Maps_DeleteCaptureFile(A_Temp "\ahk-maps-cap.png")
}

Maps_IsDesktopPath(path) {
    desktop := Palace_ResolveDesktopDir()
    if (desktop = "")
        return false
    SplitPath(path, , &dir)
    return (StrLower(RTrim(dir, "\")) = StrLower(desktop))
}

Maps_GdipStartup() {
    static token := 0
    if (token)
        return true
    si := Buffer(32, 0)
    NumPut("uint", 1, si, 0)
    newToken := 0
    status := 1
    try status := DllCall("gdiplus\GdiplusStartup", "uptr*", &newToken, "ptr", si, "ptr", 0)
    catch {
        return false
    }
    if (status != 0 || !newToken)
        return false
    token := newToken
    return true
}

Maps_PngEncoderClsid() {
    static clsid := 0
    if IsObject(clsid)
        return clsid
    ; image/png {557CF406-1A04-11D3-9A73-0000F81EF32E}
    clsid := Buffer(16, 0)
    NumPut("uint", 0x557CF406, clsid, 0)
    NumPut("ushort", 0x1A04, clsid, 4)
    NumPut("ushort", 0x11D3, clsid, 6)
    NumPut("uchar", 0x9A, clsid, 8)
    NumPut("uchar", 0x73, clsid, 9)
    NumPut("uchar", 0x00, clsid, 10)
    NumPut("uchar", 0x00, clsid, 11)
    NumPut("uchar", 0xF8, clsid, 12)
    NumPut("uchar", 0x1E, clsid, 13)
    NumPut("uchar", 0xF3, clsid, 14)
    NumPut("uchar", 0x2E, clsid, 15)
    return clsid
}

Maps_GdipSaveHBitmap(hbm, path, &err) {
    err := ""
    if !hbm {
        err := "empty bitmap"
        return false
    }
    if !Maps_GdipStartup() {
        err := "gdiplus startup failed"
        return false
    }
    pBitmap := 0
    status := 1
    try status := DllCall("gdiplus\GdipCreateBitmapFromHBITMAP", "ptr", hbm, "ptr", 0, "ptr*", &pBitmap)
    catch as e {
        err := e.Message
        return false
    }
    if (status != 0 || !pBitmap) {
        err := "bitmap convert failed " status
        return false
    }
    try status := DllCall("gdiplus\GdipSaveImageToFile", "ptr", pBitmap, "wstr", path, "ptr", Maps_PngEncoderClsid(),
    "ptr", 0)
    catch as e {
        status := 1
        err := e.Message
    }
    DllCall("gdiplus\GdipDisposeImage", "ptr", pBitmap)
    if (status != 0) {
        if (err = "")
            err := "png save failed " status
        return false
    }
    if !FileExist(path) {
        err := "png save failed"
        return false
    }
    return true
}

; Same thresholds as the old PowerShell gate: std under 4.5, or 90% near-black, is blank.
Maps_GdipSampleCode(pBitmap, w, h, &err) {
    err := ""
    rect := Buffer(16, 0)
    NumPut("int", w, rect, 8)
    NumPut("int", h, rect, 12)
    bd := Buffer(48, 0)
    status := 1
    try status := DllCall("gdiplus\GdipBitmapLockBits", "ptr", pBitmap, "ptr", rect, "uint", 1, "int", 0x26200A, "ptr",
        bd)
    catch as e {
        err := e.Message
        return 4
    }
    if (status != 0) {
        err := "lock bits failed " status
        return 4
    }
    stride := NumGet(bd, 8, "int")
    scan0 := NumGet(bd, 16, "ptr")
    code := 0
    if (!scan0 || stride = 0) {
        err := "lock bits failed"
        code := 4
    } else {
        stepX := Max(1, w // 32)
        stepY := Max(1, h // 32)
        n := 0
        black := 0
        sum := 0.0
        sum2 := 0.0
        px := Buffer(4, 0)
        y := stepY // 2
        while (y < h) {
            x := stepX // 2
            while (x < w) {
                addr := scan0 + (y * stride) + (x * 4)
                DllCall("RtlMoveMemory", "ptr", px, "ptr", addr, "uptr", 4)
                bb := NumGet(px, 0, "UChar")
                gg := NumGet(px, 1, "UChar")
                rr := NumGet(px, 2, "UChar")
                aa := NumGet(px, 3, "UChar")
                n++
                lum := (0.2126 * rr) + (0.7152 * gg) + (0.0722 * bb)
                sum += lum
                sum2 += lum * lum
                if (aa < 12 || (rr < 14 && gg < 14 && bb < 14))
                    black++
                x += stepX
            }
            y += stepY
        }
        if (n < 16) {
            err := "not a valid PNG"
            code := 4
        } else {
            mean := sum / n
            variance := (sum2 / n) - (mean * mean)
            if (variance < 0)
                variance := 0
            std := Sqrt(variance)
            ratio := black / n
            if (std < 4.5 || ratio >= 0.90) {
                err := "blank or black image"
                code := 2
            }
        }
    }
    DllCall("gdiplus\GdipBitmapUnlockBits", "ptr", pBitmap, "ptr", bd)
    return code
}

Maps_InvokePixelGate(path, &errDetail) {
    errDetail := ""
    if (path = "" || !FileExist(path)) {
        errDetail := "capture file missing"
        return 4
    }
    try {
        if (FileGetSize(path) < 8000) {
            errDetail := "blank or black image"
            return 3
        }
    } catch {
        errDetail := "capture file missing"
        return 4
    }
    if !Maps_GdipStartup() {
        errDetail := "gdiplus startup failed"
        return 4
    }
    pBitmap := 0
    status := 1
    try status := DllCall("gdiplus\GdipLoadImageFromFile", "wstr", path, "ptr*", &pBitmap)
    catch as e {
        errDetail := e.Message
        return 4
    }
    if (status != 0 || !pBitmap) {
        errDetail := "png load failed " status
        return 4
    }
    w := 0
    h := 0
    DllCall("gdiplus\GdipGetImageWidth", "ptr", pBitmap, "uint*", &w)
    DllCall("gdiplus\GdipGetImageHeight", "ptr", pBitmap, "uint*", &h)
    if (w < 200 || h < 150 || w > 10000 || h > 10000) {
        errDetail := "image too small"
        DllCall("gdiplus\GdipDisposeImage", "ptr", pBitmap)
        return 3
    }
    code := Maps_GdipSampleCode(pBitmap, w, h, &errDetail)
    DllCall("gdiplus\GdipDisposeImage", "ptr", pBitmap)
    return code
}

Maps_ClientScreenRect(hwnd, &l, &t, &r, &b) {
    l := 0
    t := 0
    r := 0
    b := 0
    if !hwnd
        return false
    cr := Buffer(16, 0)
    if !DllCall("GetClientRect", "ptr", hwnd, "ptr", cr)
        return false
    pt := Buffer(8, 0)
    if !DllCall("ClientToScreen", "ptr", hwnd, "ptr", pt)
        return false
    l := NumGet(pt, 0, "int")
    t := NumGet(pt, 4, "int")
    r := l + NumGet(cr, 8, "int")
    b := t + NumGet(cr, 12, "int")
    return (r > l && b > t)
}

; Prefer the pane when it sits inside the browser client. Otherwise intersect, then the client itself.
Maps_FitCaptureRect(hwnd, paneL, paneT, paneR, paneB, &outL, &outT, &outW, &outH, &rule, &clientL, &clientT, &clientR,
    &clientB) {
    rule := "no_client"
    outL := 0
    outT := 0
    outW := 0
    outH := 0
    clientL := 0
    clientT := 0
    clientR := 0
    clientB := 0
    if !Maps_ClientScreenRect(hwnd, &clientL, &clientT, &clientR, &clientB)
        return false
    paneL := Round(paneL)
    paneT := Round(paneT)
    paneR := Round(paneR)
    paneB := Round(paneB)
    if (paneR > paneL && paneB > paneT && paneL >= clientL && paneT >= clientT && paneR <= clientR && paneB <= clientB) {
        rule := "pane"
        outL := paneL
        outT := paneT
        outW := paneR - paneL
        outH := paneB - paneT
        return true
    }
    il := Max(paneL, clientL)
    it := Max(paneT, clientT)
    ir := Min(paneR, clientR)
    ib := Min(paneB, clientB)
    if ((ir - il) >= 200 && (ib - it) >= 150) {
        rule := "intersect"
        outL := il
        outT := it
        outW := ir - il
        outH := ib - it
        return true
    }
    rule := "client"
    outL := clientL
    outT := clientT
    outW := clientR - clientL
    outH := clientB - clientT
    return (outW >= 200 && outH >= 150)
}

Maps_QualityReason(code) {
    if (code = 2)
        return "blank or black image"
    if (code = 3)
        return "image too small"
    if (code = 4)
        return "not a valid PNG"
    if (code = 5)
        return "window capture failed"
    return "capture failed quality check"
}

Maps_PngHeaderOk(path, &imgW, &imgH, &bytes, &err, minW, minH) {
    imgW := 0
    imgH := 0
    bytes := 0
    err := ""
    try bytes := FileGetSize(path)
    catch {
        err := "capture file missing"
        return false
    }
    buf := Buffer(24, 0)
    f := 0
    try f := FileOpen(path, "r")
    catch {
        err := "not a valid PNG"
        return false
    }
    if !f {
        err := "not a valid PNG"
        return false
    }
    nRead := 0
    try nRead := f.RawRead(buf, 24)
    catch {
        nRead := 0
    }
    f.Close()
    if (nRead < 24) {
        err := "not a valid PNG"
        return false
    }
    sig := [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]
    for i, expected in sig {
        if (NumGet(buf, i - 1, "UChar") != expected) {
            err := "not a valid PNG"
            return false
        }
    }
    imgW := (NumGet(buf, 16, "UChar") << 24) | (NumGet(buf, 17, "UChar") << 16)
    | (NumGet(buf, 18, "UChar") << 8) | NumGet(buf, 19, "UChar")
    imgH := (NumGet(buf, 20, "UChar") << 24) | (NumGet(buf, 21, "UChar") << 16)
    | (NumGet(buf, 22, "UChar") << 8) | NumGet(buf, 23, "UChar")
    if (imgW < 1 || imgH < 1 || imgW > 10000 || imgH > 10000) {
        err := "not a valid PNG"
        return false
    }
    if (imgW < minW || imgH < minH) {
        err := "image too small"
        return false
    }
    ; A real map/Street View frame does not compress to a few kilobytes. Solid black/blank PNGs do.
    if (bytes < 8000) {
        err := "blank or black image"
        return false
    }
    return true
}

Maps_WaitStableFile(path, timeoutMs := 2500) {
    deadline := A_TickCount + timeoutMs
    prev := -1
    hits := 0
    while (A_TickCount < deadline) {
        if !FileExist(path) {
            hits := 0
            prev := -1
            Sleep 80
            continue
        }
        sz := -1
        try sz := FileGetSize(path)
        catch {
            sz := -1
        }
        if (sz > 0 && sz = prev) {
            hits += 1
            if (hits >= 2)
                return true
        } else {
            hits := 0
        }
        prev := sz
        Sleep 120
    }
    return false
}

Maps_MoveOntoDesktop(srcPath, destPath, &err) {
    err := ""
    if (srcPath = "" || !FileExist(srcPath)) {
        err := "capture file missing"
        return false
    }
    if !Maps_IsDesktopPath(destPath) {
        err := "image is not on the Desktop"
        return false
    }
    if (StrLower(srcPath) = StrLower(destPath))
        return true
    try {
        if FileExist(destPath)
            FileDelete(destPath)
    } catch {
    }
    try {
        FileMove(srcPath, destPath, 1)
    } catch {
        try FileCopy(srcPath, destPath, 1)
        catch {
            err := "could not paste image on Desktop"
            return false
        }
        Maps_DeleteCaptureFile(srcPath)
    }
    if !FileExist(destPath) {
        err := "Desktop paste did not land"
        return false
    }
    return true
}

; Pixel-check the temp capture, then paste it onto the Desktop and confirm the landed file.
Maps_CommitDesktopPng(srcPath, destPath, hwnd, minW, minH, &err) {
    err := ""
    Maps_DebugLog("commit_start", Map("src", srcPath, "dest", destPath, "minW", minW, "minH", minH))
    if (srcPath = "" || !FileExist(srcPath)) {
        err := "capture file missing"
        Maps_DebugLog("commit_fail", Map("err", err))
        return false
    }
    Maps_Loading("⏳ Checking capture...", hwnd)
    if !Maps_WaitStableFile(srcPath, 2000) {
        err := "capture file was not ready"
        Maps_DebugLog("commit_fail", Map("err", err))
        Maps_DiscardCaptureFile(srcPath, err)
        return false
    }
    imgW := 0
    imgH := 0
    bytes := 0
    if !Maps_PngHeaderOk(srcPath, &imgW, &imgH, &bytes, &err, minW, minH) {
        Maps_DebugLog("commit_header", Map("ok", 0, "err", err, "w", imgW, "h", imgH, "bytes", bytes))
        Maps_DiscardCaptureFile(srcPath, err)
        return false
    }
    Maps_DebugLog("commit_header", Map("ok", 1, "err", "", "w", imgW, "h", imgH, "bytes", bytes))
    gateErr := ""
    gate := Maps_InvokePixelGate(srcPath, &gateErr)
    Maps_DebugLog("pixel_gate", Map("code", gate, "reason", gate = 0 ? "ok" : Maps_QualityReason(gate), "err", gateErr))
    if (gate != 0) {
        err := Maps_QualityReason(gate)
        Maps_DiscardCaptureFile(srcPath, err)
        return false
    }
    Maps_Loading("⏳ Pasting image on Desktop...", hwnd)
    if !Maps_MoveOntoDesktop(srcPath, destPath, &err) {
        Maps_DebugLog("commit_fail", Map("err", err))
        Maps_DiscardCaptureFile(srcPath, err)
        return false
    }
    if !Maps_WaitStableFile(destPath, 2500) {
        err := "Desktop file was not ready"
        Maps_DebugLog("commit_fail", Map("err", err, "dest", destPath))
        Maps_DiscardCaptureFile(destPath, err)
        return false
    }
    landedW := 0
    landedH := 0
    landedBytes := 0
    if !Maps_PngHeaderOk(destPath, &landedW, &landedH, &landedBytes, &err, minW, minH) {
        Maps_DebugLog("commit_fail", Map("err", err, "w", landedW, "h", landedH, "bytes", landedBytes))
        Maps_DiscardCaptureFile(destPath, err)
        return false
    }
    if (landedBytes != bytes || landedW != imgW || landedH != imgH) {
        err := "Desktop file did not match the capture"
        Maps_DebugLog("commit_fail", Map("err", err, "srcBytes", bytes, "destBytes", landedBytes, "srcW", imgW, "destW",
            landedW, "srcH", imgH, "destH", landedH))
        Maps_DiscardCaptureFile(destPath, err)
        return false
    }
    if !Maps_IsDesktopPath(destPath) || !FileExist(destPath) {
        err := "image is not on the Desktop"
        Maps_DebugLog("commit_fail", Map("err", err, "dest", destPath))
        Maps_DiscardCaptureFile(destPath, err)
        return false
    }
    Maps_DebugLog("commit_ok", Map("dest", destPath, "w", landedW, "h", landedH, "bytes", landedBytes))
    return true
}

Maps_RestoreMapTitle(uia) {
    if !uia
        return
    js :=
        "(function(){if(window.__ahkMapsTitle){document.title=window.__ahkMapsTitle;delete window.__ahkMapsTitle;}void(0);})()"
    try uia.JSExecute(js)
    catch {
    }
}

Maps_FindFreshCaptureFile(startStamp, timeoutMs) {
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        best := ""
        bestTime := ""
        for dir in Maps_CaptureSearchDirs() {
            loop files dir "\ahk-maps-cap*.png", "F" {
                if (A_LoopFileTimeModified < startStamp && A_LoopFileTimeCreated < startStamp)
                    continue
                if FileExist(A_LoopFileFullPath ".crdownload")
                    continue
                if (best = "" || A_LoopFileTimeModified > bestTime) {
                    best := A_LoopFileFullPath
                    bestTime := A_LoopFileTimeModified
                }
            }
        }
        if (best != "")
            return best
        Maps_Loading("⏳ Waiting for the capture file...")
        Sleep 120
    }
    return ""
}

; In-page canvas export. One javascript: navigation (a second one can cancel the download).
; Returns the downloaded temp PNG path, or "" . Does not write the Desktop file.
Maps_CaptureCanvasViaDownload(uia, startStamp, &errMsg) {
    errMsg := ""
    if !uia {
        errMsg := "canvas export failed"
        Maps_DebugLog("canvas_js", Map("err", "no uia"))
        return ""
    }
    Maps_PurgeCaptureTemps()
    js :=
        "(function(){window.__ahkMapsTitle=window.__ahkMapsTitle||document.title;document.title='AHKMAPS:pending';requestAnimationFrame(function(){try{var canvases=Array.from(document.querySelectorAll('canvas')).filter(function(c){return c.offsetWidth>0&&c.offsetHeight>0;});if(!canvases.length){document.title='AHKMAPS:0:nocanvas';return;}var best=canvases.reduce(function(a,b){return (a.width*a.height)>(b.width*b.height)?a:b;});var w=best.width,h=best.height;if(!(w>0&&h>0)){document.title='AHKMAPS:0:nocanvas';return;}var off=document.createElement('canvas');off.width=w;off.height=h;var ctx=off.getContext('2d');canvases.forEach(function(c){try{if(c.width===w&&c.height===h)ctx.drawImage(c,0,0);}catch(e){}});var pts=[[0.15,0.2],[0.5,0.15],[0.85,0.2],[0.2,0.45],[0.5,0.5],[0.8,0.45],[0.2,0.75],[0.5,0.8],[0.8,0.75],[0.35,0.6],[0.65,0.35],[0.5,0.92]];var n=0,minR=255,minG=255,minB=255,maxR=0,maxG=0,maxB=0,black=0,pi,px;for(pi=0;pi<pts.length;pi++){try{px=ctx.getImageData(Math.max(0,Math.min(w-1,Math.floor(pts[pi][0]*w))),Math.max(0,Math.min(h-1,Math.floor(pts[pi][1]*h))),1,1).data;}catch(e){var msg=String((e&&(e.name||e.message))||e||'');document.title=(/Security|taint/i.test(msg))?'AHKMAPS:0:tainted':'AHKMAPS:0:blank';return;}n++;if(px[3]<12||(px[0]<14&&px[1]<14&&px[2]<14))black++;if(px[0]<minR)minR=px[0];if(px[1]<minG)minG=px[1];if(px[2]<minB)minB=px[2];if(px[0]>maxR)maxR=px[0];if(px[1]>maxG)maxG=px[1];if(px[2]>maxB)maxB=px[2];}if(n<8||(maxR-minR<8&&maxG-minG<8&&maxB-minB<8)||(black/n>=0.90)){document.title='AHKMAPS:0:blank';return;}var url=off.toDataURL('image/png');if(!url||url.length<500){document.title='AHKMAPS:0:failed';return;}var a=document.createElement('a');a.href=url;a.download='ahk-maps-cap.png';document.body.appendChild(a);a.click();a.remove();document.title='AHKMAPS:1';}catch(e){var msg=String((e&&(e.name||e.message))||e||'');document.title=(/Security|taint/i.test(msg))?'AHKMAPS:0:tainted':'AHKMAPS:0:failed';}});})();void(0);"
    try uia.JSExecute(js)
    catch Error as jsErr {
        errMsg := "canvas export failed"
        Maps_DebugLog("canvas_js", Map("err", jsErr.Message))
        return ""
    }
    token := ""
    title := ""
    deadline := A_TickCount + 2500
    while (A_TickCount < deadline && token = "") {
        title := ""
        try title := WinGetTitle("ahk_id " uia.BrowserId)
        catch {
            title := ""
        }
        if (RegExMatch(title, "AHKMAPS:(\S+)", &m) && m[1] != "pending")
            token := m[1]
        else
            Sleep 60
    }
    Maps_DebugLog("canvas_token", Map("token", token = "" ? "(none)" : token, "title", title))
    if (token = "0:tainted" || InStr(token, "tainted")) {
        errMsg := "canvas tainted"
        return ""
    }
    if (token = "" || token = "0:failed") {
        errMsg := "canvas export failed"
        return ""
    }
    if (token = "0:nocanvas") {
        errMsg := "canvas not found"
        return ""
    }
    if (token = "0:blank" || InStr(token, "blank")) {
        errMsg := "blank or black image"
        return ""
    }
    if (token != "1") {
        errMsg := "canvas export failed"
        return ""
    }
    found := Maps_FindFreshCaptureFile(startStamp, 3000)
    if (found = "") {
        errMsg := "download timed out"
        Maps_DebugLog("canvas_download", Map("err", errMsg))
        return ""
    }
    Maps_DebugLog("canvas_download", Map("path", found))
    return found
}

Maps_CapturePrintWindow(hwnd, x, y, w, h, outPath, &errMsg) {
    errMsg := ""
    ok := false
    hdcScreen := 0
    hdcMem := 0
    hbm := 0
    hdcCrop := 0
    hbmCrop := 0
    old := 0
    old2 := 0
    x := Integer(Round(x))
    y := Integer(Round(y))
    w := Integer(Round(w))
    h := Integer(Round(h))
    if (!hwnd || w < 40 || h < 40 || outPath = "") {
        errMsg := "window capture failed"
    } else {
        try {
            wr := Buffer(16, 0)
            if !DllCall("GetWindowRect", "ptr", hwnd, "ptr", wr)
                errMsg := "GetWindowRect failed"
            else {
                wl := NumGet(wr, 0, "int")
                wt := NumGet(wr, 4, "int")
                ww := NumGet(wr, 8, "int") - wl
                wh := NumGet(wr, 12, "int") - wt
                if (ww < 50 || wh < 50)
                    errMsg := "window capture failed"
                else {
                    hdcScreen := DllCall("GetDC", "ptr", 0, "ptr")
                    hdcMem := DllCall("CreateCompatibleDC", "ptr", hdcScreen, "ptr")
                    hbm := DllCall("CreateCompatibleBitmap", "ptr", hdcScreen, "int", ww, "int", wh, "ptr")
                    if (!hdcScreen || !hdcMem || !hbm)
                        errMsg := "CreateCompatibleBitmap failed"
                    else {
                        old := DllCall("SelectObject", "ptr", hdcMem, "ptr", hbm, "ptr")
                        printed := DllCall("PrintWindow", "ptr", hwnd, "ptr", hdcMem, "uint", 2)
                        DllCall("SelectObject", "ptr", hdcMem, "ptr", old)
                        old := 0
                        if !printed
                            errMsg := "PrintWindow failed"
                        else {
                            cx := x - wl
                            cy := y - wt
                            cw := w
                            ch := h
                            if (cx < 0) {
                                cw += cx
                                cx := 0
                            }
                            if (cy < 0) {
                                ch += cy
                                cy := 0
                            }
                            if (cx + cw > ww)
                                cw := ww - cx
                            if (cy + ch > wh)
                                ch := wh - cy
                            if (cw < 40 || ch < 40)
                                errMsg := "image too small"
                            else {
                                hdcCrop := DllCall("CreateCompatibleDC", "ptr", hdcScreen, "ptr")
                                hbmCrop := DllCall("CreateCompatibleBitmap", "ptr", hdcScreen, "int", cw, "int", ch,
                                    "ptr")
                                if (!hdcCrop || !hbmCrop)
                                    errMsg := "CreateCompatibleBitmap failed"
                                else {
                                    old2 := DllCall("SelectObject", "ptr", hdcCrop, "ptr", hbmCrop, "ptr")
                                    blt := DllCall("BitBlt", "ptr", hdcCrop, "int", 0, "int", 0, "int", cw, "int", ch,
                                        "ptr", hdcMem, "int", cx, "int", cy, "uint", 0x00CC0020)
                                    DllCall("SelectObject", "ptr", hdcCrop, "ptr", old2)
                                    old2 := 0
                                    if !blt
                                        errMsg := "crop BitBlt failed"
                                    else
                                        ok := Maps_GdipSaveHBitmap(hbmCrop, outPath, &errMsg)
                                }
                            }
                        }
                    }
                }
            }
        } catch Error as e {
            ok := false
            errMsg := e.Message
        }
    }
    if (old && hdcMem)
        DllCall("SelectObject", "ptr", hdcMem, "ptr", old)
    if (old2 && hdcCrop)
        DllCall("SelectObject", "ptr", hdcCrop, "ptr", old2)
    if (hbmCrop)
        DllCall("DeleteObject", "ptr", hbmCrop)
    if (hbm)
        DllCall("DeleteObject", "ptr", hbm)
    if (hdcCrop)
        DllCall("DeleteDC", "ptr", hdcCrop)
    if (hdcMem)
        DllCall("DeleteDC", "ptr", hdcMem)
    if (hdcScreen)
        DllCall("ReleaseDC", "ptr", 0, "ptr", hdcScreen)
    Maps_DebugLog("print_window", Map("ok", ok ? 1 : 0, "err", errMsg, "hwnd", hwnd, "x", x, "y", y, "w", w, "h", h,
        "out", outPath, "exists", (outPath != "" && FileExist(outPath)) ? 1 : 0))
    if (!ok || outPath = "" || !FileExist(outPath)) {
        if (errMsg = "")
            errMsg := "window capture failed"
        Maps_DiscardCaptureFile(outPath, errMsg)
        return false
    }
    return true
}

Maps_CaptureCopyFromScreen(x, y, w, h, outPath, &errMsg) {
    errMsg := ""
    ok := false
    hdcScreen := 0
    hdcMem := 0
    hbm := 0
    old := 0
    x := Integer(Round(x))
    y := Integer(Round(y))
    w := Integer(Round(w))
    h := Integer(Round(h))
    if (w < 40 || h < 40 || outPath = "") {
        errMsg := "capture failed"
    } else {
        try {
            hdcScreen := DllCall("GetDC", "ptr", 0, "ptr")
            hdcMem := DllCall("CreateCompatibleDC", "ptr", hdcScreen, "ptr")
            hbm := DllCall("CreateCompatibleBitmap", "ptr", hdcScreen, "int", w, "int", h, "ptr")
            if (!hdcScreen || !hdcMem || !hbm)
                errMsg := "CreateCompatibleBitmap failed"
            else {
                old := DllCall("SelectObject", "ptr", hdcMem, "ptr", hbm, "ptr")
                ; CAPTUREBLT so the visible Chrome frame is included.
                blt := DllCall("BitBlt", "ptr", hdcMem, "int", 0, "int", 0, "int", w, "int", h, "ptr", hdcScreen, "int",
                    x, "int", y, "uint", 0x40CC0020)
                DllCall("SelectObject", "ptr", hdcMem, "ptr", old)
                old := 0
                if !blt
                    errMsg := "BitBlt failed"
                else
                    ok := Maps_GdipSaveHBitmap(hbm, outPath, &errMsg)
            }
        } catch Error as e {
            ok := false
            errMsg := e.Message
        }
    }
    if (old && hdcMem)
        DllCall("SelectObject", "ptr", hdcMem, "ptr", old)
    if (hbm)
        DllCall("DeleteObject", "ptr", hbm)
    if (hdcMem)
        DllCall("DeleteDC", "ptr", hdcMem)
    if (hdcScreen)
        DllCall("ReleaseDC", "ptr", 0, "ptr", hdcScreen)
    Maps_DebugLog("copy_from_screen", Map("ok", ok ? 1 : 0, "err", errMsg, "x", x, "y", y, "w", w, "h", h, "out",
        outPath, "exists", (outPath != "" && FileExist(outPath)) ? 1 : 0))
    if (!ok || outPath = "" || !FileExist(outPath)) {
        if (errMsg = "")
            errMsg := "capture failed"
        Maps_DiscardCaptureFile(outPath, errMsg)
        return false
    }
    return true
}

#HotIf WinActive("ahk_exe chrome.exe") && InStr(SafeWinGetTitle(), "Google Maps")

; Shift + S : Focus "Search Google Maps" field
+s:: {
    try {
        uia := UIA_Browser()
        if !uia
            return
        Sleep 200
        root := Maps_GetDocumentRoot(uia)
        if !root
            return

        searchBox := 0
        try searchBox := root.FindFirst({ AutomationId: "ucc-1" })
        if !searchBox {
            try searchBox := root.FindFirst({ Type: 50003, Name: "Search Google Maps", cs: false })
        }
        if !searchBox {
            try searchBox := root.FindFirst({ Name: "Search Google Maps", cs: false })
        }

        if (searchBox) {
            try searchBox.SetFocus()
            catch {
                try searchBox.Click()
            }
            Sleep 100
            if searchBox.HasKeyboardFocus
                return
            try searchBox.Click()
        }
    } catch {
    }
}

; Shift + L : Copy latitude, longitude (from place card button or Maps URL)
+l:: {
    coordOut := ""
    try {
        uia := UIA_Browser()
        if !uia {
            ToolTip("Maps: could not attach to browser")
            SetTimer(() => ToolTip(), -2000)
            return
        }
        Sleep 150
        root := Maps_GetDocumentRoot(uia)
        if root {
            try {
                for btn in root.FindAll({ Type: 50000 }) {
                    n := btn.Name
                    if RegExMatch(n, "^-?\d+\.\d+\s*,\s*-?\d+\.\d+$") {
                        if RegExMatch(n, "^(-?\d+\.\d+)\s*,\s*(-?\d+\.\d+)$", &m) {
                            coordOut := m[1] . ", " . m[2]
                        } else {
                            coordOut := Trim(n)
                        }
                        break
                    }
                }
            } catch {
            }
        }
        if (coordOut = "") {
            try {
                url := uia.GetCurrentURL()
                if RegExMatch(url, "i)google\.[^/]+/maps/@(-?\d+\.\d+),(-?\d+\.\d+)", &um) {
                    coordOut := um[1] . ", " . um[2]
                }
            } catch {
            }
        }
        if (coordOut != "") {
            A_Clipboard := coordOut
            ToolTip("Copied: " . coordOut)
            SetTimer(() => ToolTip(), -1500)
        } else {
            ToolTip("Maps: coordinates not found")
            SetTimer(() => ToolTip(), -2000)
        }
    } catch Error as e {
        ToolTip("Maps: " . e.Message)
        SetTimer(() => ToolTip(), -2000)
    }
}

; Shift + C : Collapse side panel
+c:: {
    try {
        uia := UIA_Browser()
        if !uia
            return
        Sleep 150
        root := Maps_GetDocumentRoot(uia)
        if !root
            return
        Maps_CollapseSidePanel(root)
    } catch {
    }
}

; Shift + P : Clean PNG capture (hide chrome, screenshot map / Street View, paste onto Desktop)
+p:: {
    global IS_WORK_ENVIRONMENT, g_MapsDebugMethod, g_MapsDebugAttempt
    uia := 0
    chromeHidden := false
    loadingShown := false
    browserHwnd := 0
    outPath := ""
    saved := false
    lastErr := ""
    lastMethod := ""
    failTrail := ""
    try {
        Maps_DebugBegin()
        pageTitle := ""
        try pageTitle := SafeWinGetTitle()
        catch {
            pageTitle := ""
        }
        Maps_DebugLog("enter", Map("title", pageTitle, "dpi", A_ScreenDPI, "desktop", Palace_ResolveDesktopDir()))
        Maps_Loading("⏳ Preparing Maps capture...")
        loadingShown := true
        sessionStamp := A_Now

        uia := UIA_Browser()
        if !uia {
            Maps_DebugLog("attach", Map("ok", 0))
            Maps_Fail("❌ Maps: could not attach to browser", 0, Maps_DebugHideMs())
            loadingShown := false
            return
        }
        try browserHwnd := uia.BrowserId
        Maps_DebugLog("attach", Map("ok", 1, "hwnd", browserHwnd))

        Maps_Loading("🔄 Maximizing window...", browserHwnd)
        maxOk := Maps_MaximizeBrowserWindow(uia)
        Maps_DebugLog("maximize", Map("ok", maxOk ? 1 : 0, "hwnd", browserHwnd))
        root := Maps_GetDocumentRoot(uia)
        if !root {
            Maps_DebugLog("document", Map("ok", 0))
            Maps_Fail("❌ Maps: document not found", browserHwnd, Maps_DebugHideMs())
            loadingShown := false
            return
        }
        Maps_DebugLog("document", Map("ok", 1))

        Maps_Loading("🔄 Collapsing side panel...", browserHwnd)
        collapsed := Maps_CollapseSidePanel(root)
        Maps_DebugLog("collapse_panel", Map("ok", collapsed ? 1 : 0))
        Sleep 600  ; side-panel collapse animation

        Maps_Loading("🔄 Hiding map chrome...", browserHwnd)
        hideOk := Maps_HideChrome(uia)
        chromeHidden := true
        Maps_DebugLog("hide_chrome", Map("ok", hideOk ? 1 : 0))
        Maps_Loading("⏳ Loading map tiles...", browserHwnd)
        Sleep 1200  ; let canvas resize and tiles finish loading before capture

        outPath := Palace_DesktopNextQuickImagePath()
        isWork := IsSet(IS_WORK_ENVIRONMENT) && IS_WORK_ENVIRONMENT
        methods := isWork ? ["canvas", "print", "screen"] : ["print", "screen", "canvas"]
        methodList := ""
        for methodName in methods
            methodList .= (methodList = "" ? "" : ",") methodName
        Maps_DebugLog("session_start", Map("isWork", isWork ? 1 : 0, "methods", methodList, "desktop",
            Palace_ResolveDesktopDir(), "dpi", A_ScreenDPI, "outPath", outPath))
        loop 2 {
            attempt := A_Index
            if (attempt > 1) {
                Maps_DebugLog("retry", Map("attempt", attempt, "lastErr", lastErr, "lastMethod", lastMethod))
                Maps_Loading("⏳ Capture did not pass quality check — retrying...", browserHwnd)
                Sleep 800
            }
            if (browserHwnd) {
                try WinActivate("ahk_id " browserHwnd)
            }
            root := Maps_GetDocumentRoot(uia)
            pane := Maps_FindMapPane(root)
            if !pane {
                lastErr := "map pane not found"
                Maps_DebugLog("pane", Map("attempt", attempt, "ok", 0, "err", lastErr))
                continue
            }
            br := pane.BoundingRectangle
            paneW := br.r - br.l
            paneH := br.b - br.t
            paneName := ""
            try paneName := pane.Name
            catch {
                paneName := ""
            }
            if (paneW <= 0 || paneH <= 0) {
                lastErr := "invalid capture region"
                Maps_DebugLog("pane", Map("attempt", attempt, "ok", 0, "name", paneName, "l", br.l, "t", br.t, "r", br.r,
                    "b", br.b, "w", paneW, "h", paneH, "err", lastErr))
                continue
            }
            Maps_DebugLog("pane", Map("attempt", attempt, "ok", 1, "name", paneName, "l", br.l, "t", br.t, "r", br.r,
                "b", br.b, "w", paneW, "h", paneH))
            fitL := 0
            fitT := 0
            fitW := 0
            fitH := 0
            fitRule := ""
            clientL := 0
            clientT := 0
            clientR := 0
            clientB := 0
            fitOk := Maps_FitCaptureRect(browserHwnd, br.l, br.t, br.r, br.b, &fitL, &fitT, &fitW, &fitH, &fitRule,
                &clientL, &clientT, &clientR, &clientB)
            Maps_DebugLog("rect_fit", Map("attempt", attempt, "rule", fitRule, "ok", fitOk ? 1 : 0, "paneL", br.l,
                "paneT", br.t, "paneW", paneW, "paneH", paneH, "clientL", clientL, "clientT", clientT, "clientR",
                clientR, "clientB", clientB, "l", fitL, "t", fitT, "w", fitW, "h", fitH))
            if !fitOk {
                lastErr := "invalid capture region"
                continue
            }
            for method in methods {
                g_MapsDebugMethod := method
                g_MapsDebugAttempt := attempt
                lastMethod := method
                Maps_DebugLog("method_start", Map("method", method, "attempt", attempt))
                Maps_Loading("📸 Capturing map...", browserHwnd)
                capErr := ""
                src := ""
                minW := 320
                minH := 200
                if (method = "canvas") {
                    src := Maps_CaptureCanvasViaDownload(uia, sessionStamp, &capErr)
                    Maps_RestoreMapTitle(uia)
                } else if (method = "print") {
                    src := A_Temp "\ahk-maps-cap.png"
                    Maps_DeleteCaptureFile(src)
                    if !Maps_CapturePrintWindow(browserHwnd, fitL, fitT, fitW, fitH, src, &capErr)
                        src := ""
                    minW := Max(200, Floor(fitW * 0.8))
                    minH := Max(150, Floor(fitH * 0.8))
                } else {
                    src := A_Temp "\ahk-maps-cap.png"
                    Maps_DeleteCaptureFile(src)
                    ; CopyFromScreen includes AlwaysOnTop windows, so the bar must not be in the frame.
                    StandardLoadingBar_Hide(0)
                    Sleep 90
                    grabbed := false
                    try grabbed := Maps_CaptureCopyFromScreen(fitL, fitT, fitW, fitH, src, &capErr)
                    catch Error as screenErr {
                        grabbed := false
                        capErr := "capture failed"
                        Maps_DebugLog("copy_from_screen", Map("ok", 0, "err", screenErr.Message, "x", fitL, "y", fitT,
                            "w", fitW, "h", fitH))
                    }
                    Maps_Loading("📸 Capturing map...", browserHwnd)
                    if !grabbed
                        src := ""
                    minW := Max(200, Floor(fitW * 0.8))
                    minH := Max(150, Floor(fitH * 0.8))
                }
                if (src = "" || !FileExist(src)) {
                    lastErr := capErr != "" ? capErr : "capture failed"
                    failTrail .= (failTrail = "" ? "" : " | ") "a" attempt " " method ": " lastErr
                    facts := Maps_DebugFileFacts(src)
                    Maps_DebugLog("method_fail", Map("method", method, "attempt", attempt, "err", lastErr, "bytes",
                        facts["bytes"], "w", facts["w"], "h", facts["h"], "headerErr", facts["headerErr"]))
                    Maps_DiscardCaptureFile(src, lastErr)
                    continue
                }
                facts := Maps_DebugFileFacts(src)
                Maps_DebugLog("method_file", Map("method", method, "attempt", attempt, "src", src, "bytes", facts[
                    "bytes"], "w", facts["w"], "h", facts["h"], "headerErr", facts["headerErr"]))
                if Maps_CommitDesktopPng(src, outPath, browserHwnd, minW, minH, &capErr) {
                    saved := true
                    break
                }
                lastErr := capErr != "" ? capErr : "capture failed quality check"
                failTrail .= (failTrail = "" ? "" : " | ") "a" attempt " " method ": " lastErr
                Maps_DebugLog("method_fail", Map("method", method, "attempt", attempt, "err", lastErr))
                Maps_DiscardCaptureFile(src, lastErr)
                Maps_DiscardCaptureFile(outPath, lastErr)
            }
            if (saved)
                break
        }
        if (!saved || outPath = "" || !FileExist(outPath) || !Maps_IsDesktopPath(outPath)) {
            bannerErr := (Maps_DebugEnabled() && failTrail != "") ? failTrail : (lastErr != "" ? lastErr :
                "capture failed")
            Maps_DebugLog("fail", Map("err", bannerErr, "lastErr", lastErr, "lastMethod", lastMethod, "outPath",
                outPath))
            Maps_Fail("❌ Maps: " bannerErr, browserHwnd, Maps_DebugHideMs())
            loadingShown := false
            return
        }

        if (chromeHidden && uia) {
            try Maps_RestoreChrome(uia)
            catch {
            }
            chromeHidden := false
        }
        Maps_RestoreMapTitle(uia)
        SplitPath(outPath, &savedName)
        Maps_DebugLog("saved", Map("path", outPath, "name", savedName))
        Maps_ShowSaved(savedName, browserHwnd)
        loadingShown := false
    } catch Error as e {
        extraText := ""
        try extraText := "" e.Extra
        catch {
            extraText := ""
        }
        Maps_DebugLog("exception", Map("message", e.Message, "what", e.What, "extra", extraText, "file", e.File, "line",
            e.Line))
        Maps_Fail("❌ Maps: " . e.Message, browserHwnd, Maps_DebugHideMs())
        loadingShown := false
    } finally {
        if (chromeHidden && uia) {
            try Maps_RestoreChrome(uia)
            catch {
            }
        }
        if (uia) {
            try Maps_RestoreMapTitle(uia)
            catch {
            }
        }
        Maps_PurgeCaptureTemps()
        if (loadingShown)
            StandardLoadingBar_Hide(0)
    }
}

#HotIf