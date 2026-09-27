; =============================================================================
; Shift keys module: hotif_google_maps.ahk
; Google Maps Chrome hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

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

Maps_Fail(msg, hwnd := 0) {
    global g_StandardLoadingBarGui
    if !IsObject(g_StandardLoadingBarGui)
        StandardLoadingBar_Show(msg, BANNER_ACCENT_ERROR, Maps_LoadingOptions(hwnd))
    StandardLoadingBar_Update(msg, BANNER_ACCENT_ERROR)
    StandardLoadingBar_Hide(2200)
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

Maps_RunPs(script) {
    ps1 := A_Temp "\ahk_maps_capture.ps1"
    try FileDelete(ps1)
    catch {
    }
    if !FileAppend(script, ps1, "UTF-8")
        return -1
    exitCode := -1
    try exitCode := RunWait('powershell.exe -NoProfile -ExecutionPolicy Bypass -File "' ps1 '"', , "Hide")
    catch {
        exitCode := -1
    }
    try FileDelete(ps1)
    catch {
    }
    if (exitCode = "")
        return -1
    return Integer(exitCode)
}

Maps_PixelGatePs() {
    return "
(
function Test-AhkMapPng([string]$path) {
    if (-not (Test-Path -LiteralPath $path)) { return 4 }
    $len = (Get-Item -LiteralPath $path).Length
    if ($len -lt 8000) { return 3 }
    Add-Type -AssemblyName System.Drawing
    $fs = $null
    $img = $null
    $bmp = $null
    try {
        $fs = [System.IO.File]::Open($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $img = [System.Drawing.Image]::FromStream($fs, $false, $false)
        $bmp = New-Object System.Drawing.Bitmap ([int]$img.Width), ([int]$img.Height), ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.DrawImage($img, 0, 0, $img.Width, $img.Height)
        $g.Dispose()
        $img.Dispose()
        $img = $null
    } catch {
        if ($bmp) { $bmp.Dispose() }
        return 4
    } finally {
        if ($fs) { $fs.Dispose() }
    }
    $w = $bmp.Width
    $h = $bmp.Height
    if ($w -lt 200 -or $h -lt 150 -or $w -gt 10000 -or $h -gt 10000) {
        $bmp.Dispose()
        return 3
    }
    $rect = New-Object System.Drawing.Rectangle 0, 0, $w, $h
    $bits = $bmp.LockBits($rect, [System.Drawing.Imaging.ImageLockMode]::ReadOnly, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $stride = [Math]::Abs($bits.Stride)
    $raw = New-Object byte[] ($stride * $h)
    [System.Runtime.InteropServices.Marshal]::Copy($bits.Scan0, $raw, 0, $raw.Length)
    $bmp.UnlockBits($bits)
    $bmp.Dispose()
    $stepX = [Math]::Max(1, [int][Math]::Floor($w / 32))
    $stepY = [Math]::Max(1, [int][Math]::Floor($h / 32))
    $n = 0
    $black = 0
    $sum = 0.0
    $sum2 = 0.0
    for ($y = [int]($stepY / 2); $y -lt $h; $y += $stepY) {
        for ($x = [int]($stepX / 2); $x -lt $w; $x += $stepX) {
            $i = ($y * $stride) + ($x * 4)
            if (($i + 3) -ge $raw.Length) { continue }
            $bb = $raw[$i]
            $gg = $raw[$i + 1]
            $rr = $raw[$i + 2]
            $aa = $raw[$i + 3]
            $n++
            $lum = (0.2126 * $rr) + (0.7152 * $gg) + (0.0722 * $bb)
            $sum += $lum
            $sum2 += ($lum * $lum)
            if ($aa -lt 12 -or ($rr -lt 14 -and $gg -lt 14 -and $bb -lt 14)) { $black++ }
        }
    }
    if ($n -lt 16) { return 4 }
    $mean = $sum / $n
    $variance = ($sum2 / $n) - ($mean * $mean)
    if ($variance -lt 0) { $variance = 0 }
    $std = [Math]::Sqrt($variance)
    $ratio = $black / $n
    if ($std -lt 4.5) { return 2 }
    if ($ratio -ge 0.90) { return 2 }
    return 0
}
)"
}

Maps_InvokePixelGate(path) {
    if (path = "" || !FileExist(path))
        return 4
    safe := StrReplace(path, "'", "''")
    script := Maps_PixelGatePs() "`r`n`$code = Test-AhkMapPng '" safe "'`r`nexit `$code`r`n"
    return Maps_RunPs(script)
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
    if (srcPath = "" || !FileExist(srcPath)) {
        err := "capture file missing"
        return false
    }
    Maps_Loading("⏳ Checking capture...", hwnd)
    if !Maps_WaitStableFile(srcPath, 2000) {
        err := "capture file was not ready"
        Maps_DeleteCaptureFile(srcPath)
        return false
    }
    imgW := 0
    imgH := 0
    bytes := 0
    if !Maps_PngHeaderOk(srcPath, &imgW, &imgH, &bytes, &err, minW, minH) {
        Maps_DeleteCaptureFile(srcPath)
        return false
    }
    gate := Maps_InvokePixelGate(srcPath)
    if (gate != 0) {
        err := Maps_QualityReason(gate)
        Maps_DeleteCaptureFile(srcPath)
        return false
    }
    Maps_Loading("⏳ Pasting image on Desktop...", hwnd)
    if !Maps_MoveOntoDesktop(srcPath, destPath, &err) {
        Maps_DeleteCaptureFile(srcPath)
        return false
    }
    if !Maps_WaitStableFile(destPath, 2500) {
        err := "Desktop file was not ready"
        Maps_DeleteCaptureFile(destPath)
        return false
    }
    landedW := 0
    landedH := 0
    landedBytes := 0
    if !Maps_PngHeaderOk(destPath, &landedW, &landedH, &landedBytes, &err, minW, minH) {
        Maps_DeleteCaptureFile(destPath)
        return false
    }
    if (landedBytes != bytes || landedW != imgW || landedH != imgH) {
        err := "Desktop file did not match the capture"
        Maps_DeleteCaptureFile(destPath)
        return false
    }
    if !Maps_IsDesktopPath(destPath) || !FileExist(destPath) {
        err := "image is not on the Desktop"
        Maps_DeleteCaptureFile(destPath)
        return false
    }
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
        return ""
    }
    Maps_PurgeCaptureTemps()
    js :=
        "(function(){window.__ahkMapsTitle=window.__ahkMapsTitle||document.title;document.title='AHKMAPS:pending';requestAnimationFrame(function(){try{var canvases=Array.from(document.querySelectorAll('canvas')).filter(function(c){return c.offsetWidth>0&&c.offsetHeight>0;});if(!canvases.length){document.title='AHKMAPS:0:nocanvas';return;}var best=canvases.reduce(function(a,b){return (a.width*a.height)>(b.width*b.height)?a:b;});var w=best.width,h=best.height;if(!(w>0&&h>0)){document.title='AHKMAPS:0:nocanvas';return;}var off=document.createElement('canvas');off.width=w;off.height=h;var ctx=off.getContext('2d');canvases.forEach(function(c){try{if(c.width===w&&c.height===h)ctx.drawImage(c,0,0);}catch(e){}});var pts=[[0.15,0.2],[0.5,0.15],[0.85,0.2],[0.2,0.45],[0.5,0.5],[0.8,0.45],[0.2,0.75],[0.5,0.8],[0.8,0.75],[0.35,0.6],[0.65,0.35],[0.5,0.92]];var n=0,minR=255,minG=255,minB=255,maxR=0,maxG=0,maxB=0,black=0,pi,px;for(pi=0;pi<pts.length;pi++){try{px=ctx.getImageData(Math.max(0,Math.min(w-1,Math.floor(pts[pi][0]*w))),Math.max(0,Math.min(h-1,Math.floor(pts[pi][1]*h))),1,1).data;}catch(e){document.title='AHKMAPS:0:blank';return;}n++;if(px[3]<12||(px[0]<14&&px[1]<14&&px[2]<14))black++;if(px[0]<minR)minR=px[0];if(px[1]<minG)minG=px[1];if(px[2]<minB)minB=px[2];if(px[0]>maxR)maxR=px[0];if(px[1]>maxG)maxG=px[1];if(px[2]>maxB)maxB=px[2];}if(n<8||(maxR-minR<8&&maxG-minG<8&&maxB-minB<8)||(black/n>=0.90)){document.title='AHKMAPS:0:blank';return;}var url=off.toDataURL('image/png');if(!url||url.length<500){document.title='AHKMAPS:0:failed';return;}var a=document.createElement('a');a.href=url;a.download='ahk-maps-cap.png';document.body.appendChild(a);a.click();a.remove();document.title='AHKMAPS:1';}catch(e){document.title='AHKMAPS:0:failed';}});})();void(0);"
    try uia.JSExecute(js)
    catch {
        errMsg := "canvas export failed"
        return ""
    }
    token := ""
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
        return ""
    }
    return found
}

Maps_CapturePrintWindow(hwnd, x, y, w, h, outPath, &errMsg) {
    errMsg := ""
    if (!hwnd || w <= 0 || h <= 0 || outPath = "") {
        errMsg := "window capture failed"
        return false
    }
    safe := StrReplace(outPath, "'", "''")
    script := "Add-Type -AssemblyName System.Drawing`r`n"
    script .= "Add-Type -TypeDefinition @'`r`n"
    script .= "using System;`r`n"
    script .= "using System.Runtime.InteropServices;`r`n"
    script .= "public class AhkMapCap {`r`n"
    script .=
        "  [DllImport(`"user32.dll`")] public static extern bool PrintWindow(IntPtr hwnd, IntPtr hdc, uint flags);`r`n"
    script .= "  [DllImport(`"user32.dll`")] public static extern bool GetWindowRect(IntPtr hwnd, out RECT rect);`r`n"
    script .= "  [StructLayout(LayoutKind.Sequential)]`r`n"
    script .= "  public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }`r`n"
    script .=
        "  public static RECT WindowRect(IntPtr hwnd) { RECT r = new RECT(); GetWindowRect(hwnd, out r); return r; }`r`n"
    script .= "}`r`n"
    script .= "'@`r`n"
    script .= "$hwnd = [IntPtr]" Integer(hwnd) "`r`n"
    script .= "$wantX = " Integer(x) "`r`n"
    script .= "$wantY = " Integer(y) "`r`n"
    script .= "$wantW = " Integer(w) "`r`n"
    script .= "$wantH = " Integer(h) "`r`n"
    script .= "$rect = [AhkMapCap]::WindowRect($hwnd)`r`n"
    script .= "$ww = $rect.Right - $rect.Left`r`n"
    script .= "$wh = $rect.Bottom - $rect.Top`r`n"
    script .= "if ($ww -lt 50 -or $wh -lt 50) { exit 5 }`r`n"
    script .= "$full = New-Object System.Drawing.Bitmap $ww, $wh`r`n"
    script .= "$g = [System.Drawing.Graphics]::FromImage($full)`r`n"
    script .= "$hdc = $g.GetHdc()`r`n"
    script .= "$printed = [AhkMapCap]::PrintWindow($hwnd, $hdc, 2)`r`n"
    script .= "$g.ReleaseHdc($hdc)`r`n"
    script .= "$g.Dispose()`r`n"
    script .= "if (-not $printed) { $full.Dispose(); exit 5 }`r`n"
    script .= "$scaleX = $full.Width / [double]$ww`r`n"
    script .= "$scaleY = $full.Height / [double]$wh`r`n"
    script .= "$cx = [int][Math]::Round(($wantX - $rect.Left) * $scaleX)`r`n"
    script .= "$cy = [int][Math]::Round(($wantY - $rect.Top) * $scaleY)`r`n"
    script .= "$cw = [int][Math]::Round($wantW * $scaleX)`r`n"
    script .= "$ch = [int][Math]::Round($wantH * $scaleY)`r`n"
    script .= "if ($cx -lt 0) { $cw += $cx; $cx = 0 }`r`n"
    script .= "if ($cy -lt 0) { $ch += $cy; $cy = 0 }`r`n"
    script .= "if ($cx + $cw -gt $full.Width) { $cw = $full.Width - $cx }`r`n"
    script .= "if ($cy + $ch -gt $full.Height) { $ch = $full.Height - $cy }`r`n"
    script .= "if ($cw -lt 40 -or $ch -lt 40) { $full.Dispose(); exit 3 }`r`n"
    script .= "$crop = New-Object System.Drawing.Bitmap $cw, $ch`r`n"
    script .= "$cg = [System.Drawing.Graphics]::FromImage($crop)`r`n"
    script .= "$dest = New-Object System.Drawing.Rectangle 0, 0, $cw, $ch`r`n"
    script .= "$src = New-Object System.Drawing.Rectangle $cx, $cy, $cw, $ch`r`n"
    script .= "$cg.DrawImage($full, $dest, $src, [System.Drawing.GraphicsUnit]::Pixel)`r`n"
    script .= "$cg.Dispose()`r`n"
    script .= "$full.Dispose()`r`n"
    script .= "$crop.Save('" safe "', [System.Drawing.Imaging.ImageFormat]::Png)`r`n"
    script .= "$crop.Dispose()`r`n"
    script .= "exit 0`r`n"
    code := Maps_RunPs(script)
    if (code != 0 || !FileExist(outPath)) {
        Maps_DeleteCaptureFile(outPath)
        errMsg := Maps_QualityReason(code = 0 ? 5 : code)
        return false
    }
    return true
}

Maps_CaptureCopyFromScreen(x, y, w, h, outPath, &errMsg) {
    errMsg := ""
    if (w <= 0 || h <= 0 || outPath = "") {
        errMsg := "capture failed"
        return false
    }
    safe := StrReplace(outPath, "'", "''")
    script := "Add-Type -AssemblyName System.Drawing`r`n"
    script .= "$b = New-Object System.Drawing.Bitmap " Integer(w) ", " Integer(h) "`r`n"
    script .= "$g = [System.Drawing.Graphics]::FromImage($b)`r`n"
    script .= "$g.CopyFromScreen(" Integer(x) ", " Integer(y) ", 0, 0, $b.Size)`r`n"
    script .= "$b.Save('" safe "', [System.Drawing.Imaging.ImageFormat]::Png)`r`n"
    script .= "$g.Dispose()`r`n"
    script .= "$b.Dispose()`r`n"
    script .= "exit 0`r`n"
    code := Maps_RunPs(script)
    if (code != 0 || !FileExist(outPath)) {
        Maps_DeleteCaptureFile(outPath)
        errMsg := "capture failed"
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
    global IS_WORK_ENVIRONMENT
    uia := 0
    chromeHidden := false
    loadingShown := false
    browserHwnd := 0
    outPath := ""
    saved := false
    try {
        Maps_Loading("⏳ Preparing Maps capture...")
        loadingShown := true
        sessionStamp := A_Now

        uia := UIA_Browser()
        if !uia {
            Maps_Fail("❌ Maps: could not attach to browser")
            loadingShown := false
            return
        }
        try browserHwnd := uia.BrowserId

        Maps_Loading("🔄 Maximizing window...", browserHwnd)
        Maps_MaximizeBrowserWindow(uia)
        root := Maps_GetDocumentRoot(uia)
        if !root {
            Maps_Fail("❌ Maps: document not found", browserHwnd)
            loadingShown := false
            return
        }

        Maps_Loading("🔄 Collapsing side panel...", browserHwnd)
        Maps_CollapseSidePanel(root)
        Sleep 600  ; side-panel collapse animation

        Maps_Loading("🔄 Hiding map chrome...", browserHwnd)
        Maps_HideChrome(uia)
        chromeHidden := true
        Maps_Loading("⏳ Loading map tiles...", browserHwnd)
        Sleep 1200  ; let canvas resize and tiles finish loading before capture

        outPath := Palace_DesktopNextQuickImagePath()
        isWork := IsSet(IS_WORK_ENVIRONMENT) && IS_WORK_ENVIRONMENT
        methods := isWork ? ["canvas", "print", "screen"] : ["print", "screen", "canvas"]
        lastErr := ""
        loop 2 {
            if (A_Index > 1) {
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
                continue
            }
            br := pane.BoundingRectangle
            paneW := br.r - br.l
            paneH := br.b - br.t
            if (paneW <= 0 || paneH <= 0) {
                lastErr := "invalid capture region"
                continue
            }
            for method in methods {
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
                    if !Maps_CapturePrintWindow(browserHwnd, br.l, br.t, paneW, paneH, src, &capErr)
                        src := ""
                    minW := Max(200, Floor(paneW * 0.8))
                    minH := Max(150, Floor(paneH * 0.8))
                } else {
                    src := A_Temp "\ahk-maps-cap.png"
                    Maps_DeleteCaptureFile(src)
                    ; CopyFromScreen includes AlwaysOnTop windows, so the bar must not be in the frame.
                    StandardLoadingBar_Hide(0)
                    Sleep 90
                    grabbed := false
                    try grabbed := Maps_CaptureCopyFromScreen(br.l, br.t, paneW, paneH, src, &capErr)
                    catch {
                        grabbed := false
                    }
                    Maps_Loading("📸 Capturing map...", browserHwnd)
                    if !grabbed
                        src := ""
                    minW := Max(200, Floor(paneW * 0.8))
                    minH := Max(150, Floor(paneH * 0.8))
                }
                if (src = "" || !FileExist(src)) {
                    lastErr := capErr != "" ? capErr : "capture failed"
                    Maps_DeleteCaptureFile(src)
                    continue
                }
                if Maps_CommitDesktopPng(src, outPath, browserHwnd, minW, minH, &capErr) {
                    saved := true
                    break
                }
                lastErr := capErr != "" ? capErr : "capture failed quality check"
                Maps_DeleteCaptureFile(src)
                Maps_DeleteCaptureFile(outPath)
            }
            if (saved)
                break
        }
        if (!saved || outPath = "" || !FileExist(outPath) || !Maps_IsDesktopPath(outPath)) {
            Maps_Fail("❌ Maps: " . (lastErr != "" ? lastErr : "capture failed"), browserHwnd)
            loadingShown := false
            return
        }

        Maps_Loading("⏳ Restoring map...", browserHwnd)
        if (chromeHidden && uia) {
            try Maps_RestoreChrome(uia)
            catch {
            }
            chromeHidden := false
        }
        Maps_RestoreMapTitle(uia)
        SplitPath(outPath, &savedName)
        StandardLoadingBar_Update("✅ Saved on Desktop: " savedName, BANNER_ACCENT_SUCCESS)
        StandardLoadingBar_Hide(2800)
        loadingShown := false
    } catch Error as e {
        Maps_Fail("❌ Maps: " . e.Message, browserHwnd)
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