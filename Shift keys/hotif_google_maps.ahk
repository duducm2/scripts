; =============================================================================
; Shift keys module: hotif_google_maps.ahk
; Google Maps Chrome hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

#HotIf WinActive("ahk_exe chrome.exe") && InStr(WinGetTitle("A"), "Google Maps")

; Shift + S : Focus "Search Google Maps" field
+s:: {
    try {
        uia := UIA_Browser()
        if !uia
            return
        Sleep 200
        root := 0
        try root := uia.GetCurrentDocumentElement()
        catch {
            try root := uia.BrowserElement
        }
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
        root := 0
        try root := uia.GetCurrentDocumentElement()
        catch {
            try root := uia.BrowserElement
        }
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

#HotIf
