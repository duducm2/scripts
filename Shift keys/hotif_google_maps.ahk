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

Maps_CaptureRectToDesktop(x, y, w, h) {
    if (w <= 0 || h <= 0)
        return ""
    outPath := A_Desktop "\maps-" FormatTime(, "yyyyMMdd-HHmmss") ".png"
    ps1 := A_Temp "\ahk_maps_capture.ps1"
    safePath := StrReplace(outPath, "'", "''")
    script := (
        "Add-Type -AssemblyName System.Drawing`r`n"
        "$b = New-Object System.Drawing.Bitmap(" w ", " h ")`r`n"
        "$g = [System.Drawing.Graphics]::FromImage($b)`r`n"
        "$g.CopyFromScreen(" Integer(x) ", " Integer(y) ", 0, 0, $b.Size)`r`n"
        "$b.Save('" safePath "')`r`n"
        "$g.Dispose()`r`n"
        "$b.Dispose()`r`n"
    )
    try FileDelete(ps1)
    catch {
    }
    FileAppend(script, ps1, "UTF-8")
    exitCode := RunWait('powershell.exe -NoProfile -ExecutionPolicy Bypass -File "' ps1 '"', , "Hide")
    try FileDelete(ps1)
    catch {
    }
    if (exitCode != 0 || !FileExist(outPath))
        return ""
    return outPath
}

#HotIf WinActive("ahk_exe chrome.exe") && InStr(WinGetTitle("A"), "Google Maps")

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

; Shift + P : Clean PNG capture (hide chrome, screenshot map / Street View pane)
+p:: {
    uia := 0
    chromeHidden := false
    try {
        uia := UIA_Browser()
        if !uia {
            ToolTip("Maps: could not attach to browser")
            SetTimer(() => ToolTip(), -2000)
            return
        }
        Sleep 150
        root := Maps_GetDocumentRoot(uia)
        if !root {
            ToolTip("Maps: document not found")
            SetTimer(() => ToolTip(), -2000)
            return
        }

        Maps_CollapseSidePanel(root)
        Sleep 600  ; side-panel collapse animation

        if !Maps_HideChrome(uia) {
            ToolTip("Maps: canvas not found")
            SetTimer(() => ToolTip(), -2000)
            return
        }
        chromeHidden := true
        Sleep 1500  ; let canvas resize and tiles finish loading before capture

        root := Maps_GetDocumentRoot(uia)
        pane := Maps_FindMapPane(root)
        if !pane {
            ToolTip("Maps: map pane not found")
            SetTimer(() => ToolTip(), -2000)
            return
        }

        br := pane.BoundingRectangle
        w := br.r - br.l
        h := br.b - br.t
        if (w <= 0 || h <= 0) {
            ToolTip("Maps: invalid capture region")
            SetTimer(() => ToolTip(), -2000)
            return
        }

        outPath := Maps_CaptureRectToDesktop(br.l, br.t, w, h)
        if (outPath = "") {
            ToolTip("Maps: capture failed")
            SetTimer(() => ToolTip(), -2000)
            return
        }

        ToolTip("Saved: " . outPath)
        SetTimer(() => ToolTip(), -2500)
    } catch Error as e {
        ToolTip("Maps: " . e.Message)
        SetTimer(() => ToolTip(), -2000)
    } finally {
        if (chromeHidden && uia) {
            try Maps_RestoreChrome(uia)
            catch {
            }
        }
    }
}

#HotIf