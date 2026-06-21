; =============================================================================
; AppLaunchers module: wikipedia_scroll.ahk
; Wikipedia scroll position save/load/restore
; Extracted verbatim from AppLaunchers.ahk; loaded via #include into the
; AppLaunchers.ahk process, which remains the entry point / source of truth.
; =============================================================================

; =============================================================================
; Wikipedia Scroll Position Storage Functions
; =============================================================================

; True when the active window is on an AHK monitor where Wikipedia fullscreen scroll restore runs.
; Current layout: monitors 3 and 4 are portrait (1080x1920); both use the same restore path as M3.
IsWindowOnWikipediaScrollRestoreMonitor() {
    hwnd := WinExist("A")

    if (!hwnd) {
        return false
    }

    rect := Buffer(16, 0)
    if (!DllCall("GetWindowRect", "ptr", hwnd, "ptr", rect)) {
        return false
    }

    left := NumGet(rect, 0, "int")
    top := NumGet(rect, 4, "int")
    right := NumGet(rect, 8, "int")
    bottom := NumGet(rect, 12, "int")

    centerX := left + (right - left) // 2
    centerY := top + (bottom - top) // 2

    monitorCount := MonitorGetCount()
    loop monitorCount {
        MonitorGet(A_Index, &ml, &mt, &mr, &mb)
        if (centerX >= ml && centerX <= mr && centerY >= mt && centerY <= mb) {
            idx := A_Index
            return (idx = 3 || idx = 4)
        }
    }

    return false
}

; Helper function to normalize Wikipedia URLs
NormalizeWikipediaURL(url) {
    if (url = "" || !InStr(url, "wikipedia.org")) {
        return ""
    }
    ; Remove fragments and trailing slashes
    url := RegExReplace(url, "/#.*$", "")
    url := RegExReplace(url, "/+$", "")
    return url
}

; Get current Wikipedia article URL from the active Chrome window
GetWikipediaURL() {
    try {
        if (!WinActive("ahk_exe chrome.exe")) {
            return ""
        }

        winTitle := WinGetTitle("A")
        if (!InStr(winTitle, "Wikipedia")) {
            return ""
        }

        uia := UIA_Browser("ahk_exe chrome.exe")
        if (!uia) {
            return ""
        }

        url := uia.GetCurrentURL()
        normalizedUrl := NormalizeWikipediaURL(url)
        return normalizedUrl
    } catch Error as err {
        return ""
    }
}

; Helper function to restore scroll position to a given percentage
; Returns true on success, false on failure
RestoreWikipediaScrollPosition(scrollPercentage, bannerText := "Restoring scroll position... Please wait") {
    if (scrollPercentage <= 0.0 || scrollPercentage > 1.0) {
        return false
    }

    try {
        StandardLoadingBar_Show(bannerText, BANNER_ACCENT_INTERMEDIATE)
        Sleep(10)

        uia := UIA_Browser("ahk_exe chrome.exe")
        if (!uia) {
            StandardLoadingBar_Hide(0)
            return false
        }

        ; Block input during restoration
        AL_InstallInputGuard()

        ; Wait for page to be ready
        Sleep(500)

        ; Get document height
        docHeight := uia.JSReturnThroughClipboard("document.documentElement.scrollHeight")
        if (docHeight = "" || docHeight = "undefined" || docHeight = "null") {
            AL_RemoveInputGuard()
            StandardLoadingBar_Hide(0)
            return false
        }

        docHeightFloat := Float(docHeight)
        if (docHeightFloat <= 0) {
            AL_RemoveInputGuard()
            StandardLoadingBar_Hide(0)
            return false
        }

        ; Calculate and execute scroll
        targetScrollY := scrollPercentage * docHeightFloat
        uia.JSExecute("window.scrollTo(0, " . Round(targetScrollY) . ");")
        Sleep(500)

        AL_RemoveInputGuard()
        StandardLoadingBar_Update("Scroll position restored!")
        StandardLoadingBar_Hide(500)
        return true
    } catch Error as err {
        AL_RemoveInputGuard()
        StandardLoadingBar_Hide(0)
        return false
    }
}

; Save scroll position for a Wikipedia article URL
; Now saves as percentage (0.0 to 1.0) instead of absolute pixels
SaveWikipediaScrollPosition(url, scrollPercentage) {
    global g_WikipediaScrollPositionsFile
    try {
        if (url = "" || scrollPercentage = "" || scrollPercentage < 0 || scrollPercentage > 1) {
            return false
        }
        ; Normalize URL to match load format - remove trailing slashes and fragments
        normalizedUrl := RegExReplace(url, "/#.*$", "")
        normalizedUrl := RegExReplace(normalizedUrl, "/+$", "")
        ; Ensure directory exists
        SplitPath(g_WikipediaScrollPositionsFile, , &dir)
        if (dir != "" && !DirExist(dir)) {
            DirCreate(dir)
        }
        ; Read existing entries first (before deleting file) to preserve them
        existingEntries := Map()
        if (FileExist(g_WikipediaScrollPositionsFile)) {
            try {
                ; Read all existing entries from the Positions section
                ; We'll read the file manually to handle UTF-16 encoding issues
                fileContent := FileRead(g_WikipediaScrollPositionsFile)
                ; Parse INI format manually
                inPositionsSection := false
                loop parse fileContent, "`n", "`r" {
                    line := Trim(A_LoopField)
                    if (line = "[Positions]") {
                        inPositionsSection := true
                        continue
                    }
                    if (inPositionsSection && SubStr(line, 1, 1) = "[") {
                        ; Hit another section, stop reading
                        break
                    }
                    if (inPositionsSection && InStr(line, "=")) {
                        pos := InStr(line, "=")
                        key := Trim(SubStr(line, 1, pos - 1))
                        value := Trim(SubStr(line, pos + 1))
                        if (key != "" && value != "") {
                            existingEntries[key] := value
                        }
                    }
                }
            } catch {
                ; If read fails, try IniRead as fallback
                try {
                    ; Get all keys in Positions section (this is a workaround)
                    ; We'll just update the one we need
                } catch {
                }
            }
        }

        ; Update with new entry
        existingEntries[normalizedUrl] := scrollPercentage

        ; Delete file to recreate in UTF-8
        if (FileExist(g_WikipediaScrollPositionsFile)) {
            try {
                FileDelete(g_WikipediaScrollPositionsFile)
                Sleep(100)  ; Small delay to ensure file system updates
            } catch {
            }
        }

        ; Write all entries back in UTF-8 encoding
        try {
            ; Write UTF-8 BOM and section header
            FileAppend("[Positions]`r`n", g_WikipediaScrollPositionsFile, "UTF-8")
            ; Write each entry
            for key, value in existingEntries {
                ; Escape special INI characters in key and value
                escapedKey := StrReplace(key, "=", "`=")
                escapedKey := StrReplace(escapedKey, ";", "`;")
                escapedValue := StrReplace(value, "`n", "`;")
                escapedValue := StrReplace(escapedValue, "`r", "")
                FileAppend(escapedKey . "=" . escapedValue . "`r`n", g_WikipediaScrollPositionsFile, "UTF-8")
            }
        } catch {
            ; Fallback to IniWrite if manual write fails
            IniWrite(scrollPercentage, g_WikipediaScrollPositionsFile, "Positions", normalizedUrl)
        }
        return true
    } catch Error as err {
        return false
    }
}

; Load saved scroll position for a Wikipedia article URL
; Returns percentage (0.0 to 1.0) instead of absolute pixels
LoadWikipediaScrollPosition(url) {
    global g_WikipediaScrollPositionsFile
    try {
        if (url = "") {
            return 0.0
        }

        ; Verify INI file path is set
        if (!g_WikipediaScrollPositionsFile) {
            return 0.0
        }

        ; Normalize URL to match save format - remove trailing slashes and fragments
        normalizedUrl := RegExReplace(url, "/#.*$", "")
        normalizedUrl := RegExReplace(normalizedUrl, "/+$", "")

        ; Ensure directory exists (in case it was deleted)
        SplitPath(g_WikipediaScrollPositionsFile, , &dir)
        if (dir != "" && !DirExist(dir)) {
            DirCreate(dir)
        }

        ; Read from INI file
        ; Try manual parsing first (handles UTF-8 BOM and encoding issues better)
        scrollPos := "0"
        try {
            if (FileExist(g_WikipediaScrollPositionsFile)) {
                fileContent := FileRead(g_WikipediaScrollPositionsFile)
                ; Parse INI format manually
                inPositionsSection := false
                loop parse fileContent, "`n", "`r" {
                    line := Trim(A_LoopField)
                    ; Skip empty lines and comments
                    if (line = "" || SubStr(line, 1, 1) = ";") {
                        continue
                    }
                    if (line = "[Positions]") {
                        inPositionsSection := true
                        continue
                    }
                    if (inPositionsSection && SubStr(line, 1, 1) = "[") {
                        ; Hit another section, stop reading
                        break
                    }
                    if (inPositionsSection && InStr(line, "=")) {
                        pos := InStr(line, "=")
                        key := Trim(SubStr(line, 1, pos - 1))
                        value := Trim(SubStr(line, pos + 1))
                        ; Unescape special characters that were escaped during save
                        key := StrReplace(key, "`=", "=")
                        key := StrReplace(key, "`;", ";")
                        ; Compare normalized URLs (case-sensitive for Wikipedia URLs)
                        if (key = normalizedUrl && value != "" && value != "0") {
                            scrollPos := value
                            break
                        }
                    }
                }
            }
        } catch {
            ; If manual parsing fails, fall back to IniRead
            try {
                scrollPos := IniRead(g_WikipediaScrollPositionsFile, "Positions", normalizedUrl, "0")
            } catch {
                scrollPos := "0"
            }
        }

        scrollPercentage := Float(scrollPos)
        return scrollPercentage
    } catch Error as err {
        return 0.0
    }
}
