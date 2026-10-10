; =============================================================================
; Utils module: collectible_import.ahk
; Desktop COLLECTIBLE_PACK.txt → held on the avatar until Save writes mnemonics/data/repl/collectibles.json
; AI fix: Desktop COLLECTIBLE_AI_FIX.txt
; =============================================================================

Collectible_DesktopPackPath() {
    path := Palace_DesktopNewestCsvOrTxt("COLLECTIBLE_PACK")
    if (path != "")
        return PackImport_NormalizeDesktopSource(path, "COLLECTIBLE_PACK.txt")
    newest := ""
    newestTime := 0
    loop files A_Desktop . "\gemini-code*.txt", "F" {
        text := Palace_ReadUtf8(A_LoopFileFullPath)
        if (!InStr(text, "FILE: COLLECTIBLE", false))
            continue
        ts := Number(A_LoopFileTimeModified)
        if (ts > newestTime) {
            newestTime := ts
            newest := A_LoopFileFullPath
        }
    }
    if (newest = "")
        return ""
    return PackImport_NormalizeDesktopSource(newest, "COLLECTIBLE_PACK.txt")
}

Collectible_JsonField(text, key) {
    q := Chr(34)
    pattern := q . key . q . "\s*:\s*" . q . "([^" . q . "]*)" . q
    if (!RegExMatch(text, pattern, &m))
        return ""
    return m[1]
}

Collectible_WriteAiFix(errorMsg, extraNotes := "") {
    errorMsg := Trim(errorMsg)
    if (errorMsg = "")
        return ""
    body := "The Desktop collectible importer rejected my last COLLECTIBLE_PACK. Fix and re-deliver.`r`n`r`n"
        . "IMPORT ERROR`r`n"
        . errorMsg . "`r`n`r`n"
    if (Trim(extraNotes) != "")
        body .= "EXTRA NOTES`r`n" . Trim(extraNotes) . "`r`n`r`n"
    body .= "WHAT YOU MUST DO`r`n"
        . "- Re-emit one COLLECTIBLE_PACK.txt whose only data is ===FILE: COLLECTIBLE.json=== … ===END_FILE===.`r`n"
        . "- The JSON object needs palace_id (unchanged), slot, name, blurb, svg, anim, and anchor.`r`n"
        . "- slot: armor, gloves, pants, shoes, hat, ring, staff, sword, pet, cape, mount, accessory.`r`n"
        . "- anim: bob, sway, flicker, orbit, float.`r`n"
        . "- anchor: head, shoulders, hands, feet, side, back, below.`r`n"
        . "- svg is one small <svg> illustration. No scripts, no event handlers, no javascript: links.`r`n"
        . "- If the error is JSON, remove trailing commas and markdown outside the FILE markers.`r`n`r`n"
        . "DELIVERY RULES (mandatory)`r`n"
        . "- Deliver one complete COLLECTIBLE_PACK.txt (download chip preferred; else one marked code fence).`r`n"
        . "- Never claim you saved to Desktop / disk. I save the file myself.`r`n"
        . "- Re-deliver using the exact canonical filename (COLLECTIBLE_PACK.txt).`r`n"
        . "- Never add updated, corrected, v2, or similar suffixes to the filename.`r`n`r`n"
        . "After you fix it, I will save COLLECTIBLE_PACK.txt to Desktop and run Import Management [C] again.`r`n"
    path := A_Desktop . "\COLLECTIBLE_AI_FIX.txt"
    try {
        Palace_WriteUtf8(path, body)
        return path
    } catch {
        return ""
    }
}

Collectible_Fail(errorMsg, extraNotes := "") {
    path := Collectible_WriteAiFix(errorMsg, extraNotes)
    if (path != "")
        return ImportMgmt_OnAiFixReady(path, "COLLECTIBLE_AI_FIX.txt")
    Palace_Notify(errorMsg, 5000, BANNER_ACCENT_ERROR)
    return false
}

; Import Management [C] and the Desktop watcher.
Collectible_ImportFromDesktop(*) {
    Palace_EnsureData()
    path := Collectible_DesktopPackPath()
    if (path = "") {
        Palace_Notify("No COLLECTIBLE_PACK.txt on Desktop", 2800, BANNER_ACCENT_ERROR)
        Palace_ReturnAfterImport()
        return false
    }
    py := Palace_PythonDir() . "\collectibles.py"
    pyCmd := Palace_FindPythonCmd()
    if (!FileExist(py) || pyCmd = "") {
        Collectible_Fail("Collectible importer is missing (collectibles.py).")
        Palace_ReturnAfterImport()
        return false
    }
    outPath := A_Temp . "\collectible_import_result.json"
    cmd := pyCmd . ' "' . py . '" import-desktop --data-dir "' . Palace_DataDir()
    . '" --pack "' . path . '" > "' . outPath . '" 2>&1'
    try RunWait(A_ComSpec . " /c " . cmd, Palace_PythonDir(), "Hide")
    catch as e {
        Collectible_Fail("Could not run the collectible importer: " . e.Message)
        Palace_ReturnAfterImport()
        return false
    }
    out := Palace_ReadUtf8(outPath)
    if (!InStr(out, '"ok": true') && !InStr(out, '"ok":true')) {
        err := Collectible_JsonField(out, "error")
        if (err = "")
            err := Trim(out) != "" ? Trim(out) : "COLLECTIBLE_PACK could not be imported."
        Collectible_Fail(err)
        Palace_ReturnAfterImport()
        return false
    }
    name := Collectible_JsonField(out, "name")
    slot := Collectible_JsonField(out, "slot")
    stamp := FormatTime(, "yyyyMMdd-HHmmss")
    try FileCopy(path, Palace_DataDir() . "\imported\" . stamp . "_COLLECTIBLE_PACK.txt", 1)
    catch {
    }
    label := name != "" ? name : "Relic"
    if (slot != "")
        label .= " (" . slot . ")"
    Palace_Notify("Holding " . label . " — Save it in the wardrobe to keep it", 3200, BANNER_ACCENT_SUCCESS)
    Palace_AfterDataWriteRefreshUi()
    Palace_ReturnAfterImport()
    return true
}

; The palace page writes this file when Gemini is open. Send it once.
Collectible_SendPendingPrompt() {
    path := Palace_DataDir() . "\collectible_prompt_pending.txt"
    if (!FileExist(path))
        return false
    text := Trim(Palace_ReadUtf8(path))
    try FileDelete(path)
    catch {
    }
    if (text = "")
        return false
    try A_Clipboard := text
    catch {
    }
    hwnd := 0
    try hwnd := FindGeminiChromeHwnd()
    catch {
        hwnd := 0
    }
    if (!hwnd) {
        Palace_Notify("Gemini is not open — relic prompt is on the clipboard", 4000, BANNER_ACCENT_INFO)
        return true
    }
    sent := false
    try {
        hwnd := GeminiNavigateFocusAndPasteFirstSnippet(text, false)
        if (hwnd)
            sent := !!PromptPaste_SubmitWhenReady(hwnd, "gemini", 0)
    } catch {
        sent := false
    }
    if (!sent)
        Palace_Notify("Could not send the relic prompt — it is on the clipboard", 4000, BANNER_ACCENT_ERROR)
    else
        Palace_Notify("Relic prompt sent to Gemini", 2200, BANNER_ACCENT_SUCCESS)
    return true
}
