; =============================================================================
; Utils module: clip_angel_merge.ahk
; Clip Angel merge non-favorite clips
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; Clip Angel: Merge Non-Favorite Clips
; =============================================================================

; Prefer DB merge (no window). On helper failure, fall back to UI walk.
MergeNonFavoriteClips() {
    try {
        StandardLoadingBar_Show("⏳ Merging non-favorite clips...", BANNER_ACCENT_INTERMEDIATE, {
            passive: false,
            fontSize: 17
        })
        payload := ClipAngelDb_MergePayload()
        if (payload) {
            MergeNonFavoriteClips_ViaDb(payload)
            return
        }
        StandardLoadingBar_Update("⏳ DB helper unavailable — using UI...", BANNER_ACCENT_INTERMEDIATE)
        MergeNonFavoriteClips_ViaUI()
    } catch Error as e {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        ShowCenteredOverlay_Utils("❌ Merge failed: " . e.Message, 2500, BANNER_ACCENT_ERROR)
    }
}

; DB path: join text/html clips above newest favorite, put on clipboard, wait for ingest.
MergeNonFavoriteClips_ViaDb(payload) {
    total := payload["total"]
    skipped := payload["skipped"]
    text := payload["text"]
    mergedCount := total - skipped
    if (mergedCount < 1 || Trim(text) = "") {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        msg := "⚠ No non-favorite text clips to merge"
        if (skipped > 0)
            msg .= " (" skipped " non-text skipped)"
        ShowCenteredOverlay_Utils(msg, 2000, BANNER_ACCENT_INTERMEDIATE)
        return
    }
    prevId := ClipAngelDb_MaxId()
    if (prevId = "")
        prevId := -1
    else
        prevId := Integer(prevId)
    A_Clipboard := text
    try ClipWait(0.4)
    catch {
    }
    StandardLoadingBar_Update("⏳ Waiting for Clip Angel ingest...", BANNER_ACCENT_INTERMEDIATE)
    ClipAngel_WaitForClipboardIngest(CLIPANGEL_PRE_FAVORITE_INGEST_DELAY_MS, prevId)
    try StandardLoadingBar_Hide(0)
    catch {
    }
    msg := "✅ Merged " mergedCount " clip" (mergedCount = 1 ? "" : "s")
    if (skipped > 0)
        msg .= " (" skipped " non-text skipped)"
    ShowCenteredOverlay_Utils(msg, 2000, BANNER_ACCENT_SUCCESS)
}

; Extract title from first favorite under favorites filter, then merge non-favorites above it in all-marks.
MergeNonFavoriteClips_ViaUI() {
    try {
        hwnd := 0
        root := 0
        ; Start on favorites to capture the first favorite title (Row 0).
        if !ClipAngel_OpenForAutomation("favorites", 0, false, &hwnd, &root) {
            try StandardLoadingBar_Hide(0)
            catch {
            }
            ShowCenteredOverlay_Utils("❌ Clip Angel did not open (favorites).", 2000, BANNER_ACCENT_ERROR)
            return
        }
        dataGrid := ClipAngel_UiaGetDataGrid(hwnd, root)
        row0 := dataGrid ? ClipAngel_UiaResolveRow0(dataGrid) : 0
        if !row0 {
            try StandardLoadingBar_Hide(0)
            catch {
            }
            ShowCenteredOverlay_Utils("❌ No clips found in favorites.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        titleElement := ClipAngel_UiaFindFirst(row0, { Type: 50006, Name: "Title Row 0" })
        rtfValue := ""
        try rtfValue := titleElement ? titleElement.Value : ""
        catch {
            rtfValue := ""
        }
        if (rtfValue = "" || rtfValue = "System.Drawing.Bitmap") {
            try StandardLoadingBar_Hide(0)
            catch {
            }
            ShowCenteredOverlay_Utils("❌ Favorite title empty.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        favoriteClipTitle := ParseRTFToPlainText(rtfValue)

        StandardLoadingBar_Update("⏳ Searching all marks...", BANNER_ACCENT_INTERMEDIATE)
        if !ClipAngel_OpenForAutomation("all", 0, false, &hwnd, &root) {
            try StandardLoadingBar_Hide(0)
            catch {
            }
            ShowCenteredOverlay_Utils("❌ Could not switch to all marks.", 2000, BANNER_ACCENT_ERROR)
            return
        }
        dataGrid := ClipAngel_UiaGetDataGrid(hwnd, root)
        if !dataGrid {
            try StandardLoadingBar_Hide(0)
            catch {
            }
            ShowCenteredOverlay_Utils("❌ Clip list not found.", 2000, BANNER_ACCENT_ERROR)
            return
        }

        maxIterations := 40
        foundMatch := false
        loop maxIterations {
            currentRow := A_Index - 1
            currentRowElement := ClipAngel_UiaFindFirst(dataGrid, { Type: 50025, Name: "Row " . currentRow })
            if !currentRowElement
                break
            currentTitleElement := ClipAngel_UiaFindFirst(currentRowElement, { Type: 50006, Name: "Title Row " .
                currentRow })
            if currentTitleElement {
                currentRtfValue := ""
                try currentRtfValue := currentTitleElement.Value
                catch {
                    currentRtfValue := ""
                }
                if (currentRtfValue != "" && currentRtfValue != "System.Drawing.Bitmap") {
                    currentTitle := ParseRTFToPlainText(currentRtfValue)
                    if (currentTitle = favoriteClipTitle) {
                        foundMatch := true
                        StandardLoadingBar_Update("⏳ Merging...", BANNER_ACCENT_INTERMEDIATE)
                        ClipAngel_ReleaseChordModifiersForSend()
                        priorSendLevel := A_SendLevel
                        SendLevel 0
                        SendInput "{Up}"
                        ; Brief poll for selection change rather than fixed 150 ms.
                        Sleep 40
                        SendInput "^+{Home}"
                        deadline := A_TickCount + 200
                        while (A_TickCount < deadline)
                            Sleep 15
                        SendInput "^!j"
                        deadline := A_TickCount + 400
                        while (A_TickCount < deadline)
                            Sleep 20
                        SendInput "{Tab}"
                        Sleep 40
                        SendInput "^a"
                        try ClipWait(0.3)
                        catch {
                        }
                        SendInput "^c"
                        try ClipWait(0.4)
                        catch {
                        }
                        SendLevel priorSendLevel
                        try StandardLoadingBar_Hide(0)
                        catch {
                        }
                        ShowCenteredOverlay_Utils("✅ Merged non-favorite clips (copied)", 2000, BANNER_ACCENT_SUCCESS)
                        break
                    }
                }
            }
            ClipAngel_ReleaseChordModifiersForSend()
            SendInput "{Down}"
            Sleep 30
            ; Re-resolve grid if elements go stale after Down.
            if (Mod(A_Index, 5) = 0) {
                try root := UIA.ElementFromHandle(hwnd)
                catch
                    root := 0
                dataGrid := ClipAngel_UiaGetDataGrid(hwnd, root)
                if !dataGrid
                    break
            }
        }

        if !foundMatch {
            try StandardLoadingBar_Hide(0)
            catch {
            }
            ShowCenteredOverlay_Utils("⚠ Favorite clip not found in first " . maxIterations . " rows", 2000,
                BANNER_ACCENT_INTERMEDIATE)
        }
    } catch Error as e {
        try StandardLoadingBar_Hide(0)
        catch {
        }
        ShowCenteredOverlay_Utils("❌ Merge failed: " . e.Message, 2500, BANNER_ACCENT_ERROR)
    } finally {
        ClipAngel_CloseAndRestoreFocus(0)
    }
}

; Helper function to parse RTF and extract plain text
ParseRTFToPlainText(rtf) {
    ; Remove RTF header and formatting
    ; Pattern: extract text between last formatting and \par
    plainText := rtf
    ; Remove RTF control words
    plainText := RegExReplace(plainText, "\\[a-z]+\d*\s?", "")
    ; Remove braces
    plainText := RegExReplace(plainText, "[{}]", "")
    ; Clean up whitespace
    plainText := Trim(plainText)
    ; Take first line if multiline
    if InStr(plainText, "`n")
        plainText := StrSplit(plainText, "`n")[1]
    return Trim(plainText)
}
