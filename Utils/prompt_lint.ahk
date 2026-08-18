; =============================================================================
; Utils module: prompt_lint.ahk
; Structural checks before saving a prompt
; =============================================================================

PromptLint_CsvLineCount(path) {
    if (path = "" || !FileExist(path))
        return 0
    try {
        content := FileRead(path, "UTF-8")
    } catch {
        return 0
    }
    content := StrReplace(StrReplace(content, "`r`n", "`n"), "`r", "`n")
    if (content = "")
        return 0
    return StrSplit(content, "`n").Length
}

PromptLint_CheckPrompt(prompt, personalEntries, workEntries) {
    issues := []
    warnings := []
    if (!IsObject(prompt))
        return { ok: false, issues: ["Invalid prompt"], warnings: [] }

    bodyPath := PromptData_ResolvePath(prompt)
    body := ""
    if (bodyPath != "" && FileExist(bodyPath)) {
        try body := ReadUtf8File(bodyPath)
        catch {
            issues.Push("Cannot read prompt file: " bodyPath)
        }
    } else if (bodyPath != "") {
        issues.Push("Prompt file missing: " bodyPath)
    }

    draftPath := PromptData_ResolveDraftPath(prompt)
    if (draftPath != "" && !FileExist(draftPath))
        warnings.Push("Draft file path set but file missing: " draftPath)

    declared := prompt.HasProp("variables") ? Trim(prompt.variables) : ""
    declaredMap := Map()
    if (declared != "") {
        for part in StrSplit(declared, ",") {
            n := Trim(part)
            if (n != "")
                declaredMap[StrLower(n)] := true
        }
    }

    if (body != "") {
        for name, _ in PromptRender_ExtractPlaceholders(body) {
            lower := StrLower(name)
            if (SubStr(lower, 1, 8) = "include:")
                continue
            if (PromptRender_IsBuiltin(lower))
                continue
            if (!declaredMap.Has(lower))
                warnings.Push("Undeclared variable {{" name "}} (add to Variables field)")
        }
        if (StrLen(body) > 800 && !InStr(body, "`n---`n") && !InStr(body, "`n---`r`n"))
            warnings.Push("Long prompt has no --- author-notes block (reminders may be sent)")
    }

    for side in ["personal", "work"] {
        entries := side = "work" ? workEntries : personalEntries
        if (!IsObject(entries))
            continue
        for e in PromptData_ParseContextEntries(entries) {
            p := PromptData_ContextEntryPath(e)
            if (p = "")
                continue
            if !FileExist(p)
                issues.Push("Missing " side " context file: " p)
            if (PromptData_IsCsvPath(p)) {
                from := e.HasProp("csvKeepFrom") ? e.csvKeepFrom : 0
                to := e.HasProp("csvKeepTo") ? e.csvKeepTo : 0
                if (from >= 1 && to >= 1) {
                    lines := PromptLint_CsvLineCount(p)
                    if (lines > 0 && to > lines)
                        warnings.Push("CSV keep to=" to " exceeds " lines " lines in " p)
                }
            }
        }
    }

    pasteMode := prompt.HasProp("pasteMode") ? StrLower(Trim(prompt.pasteMode)) : ""
    validModes := Map("default", 1, "body_only", 1, "body_plus_clipboard", 1, "body_attach_clipboard", 1,
        "attach_only", 1, "auto_send", 1)
    if (pasteMode != "" && !validModes.Has(pasteMode))
        warnings.Push("Unknown PasteMode: " pasteMode)

    return { ok: issues.Length = 0, issues: issues, warnings: warnings }
}

PromptLint_ConfirmSave(prompt, personalEntries, workEntries) {
    r := PromptLint_CheckPrompt(prompt, personalEntries, workEntries)
    if (r.issues.Length = 0 && r.warnings.Length = 0)
        return true

    lines := ""
    for msg in r.issues
        lines .= "ERROR: " msg "`n"
    for msg in r.warnings
        lines .= "WARN: " msg "`n"
    lines := RTrim(lines, "`n")

    if (r.issues.Length > 0) {
        UtilitySelector_Notify("Lint failed:`n" lines)
        return false
    }
    result := MsgBox(lines "`n`nSave anyway?", "Prompt lint warnings", "YesNo Icon!")
    return (result = "Yes")
}
