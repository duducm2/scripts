; =============================================================================
; Utils module: prompt_render.ahk
; {{var}} render-at-paste, includes, optional fill dialog
; =============================================================================

global g_PromptRenderFillGui := false
global g_PromptRenderFillResult := { ok: false, values: Map() }

PromptRender_BuiltinNames() {
    return Map(
        "clipboard", true,
        "dictation", true,
        "date", true,
        "time", true,
        "datetime", true,
        "env", true,
        "selection", true
    )
}

PromptRender_IsBuiltin(name) {
    return PromptRender_BuiltinNames().Has(StrLower(Trim(name)))
}

PromptRender_BuiltinValue(name) {
    n := StrLower(Trim(name))
    if (n = "clipboard" || n = "dictation") {
        try return A_Clipboard
        catch {
            return ""
        }
    }
    if (n = "date")
        return FormatTime(, "yyyy-MM-dd")
    if (n = "time")
        return FormatTime(, "HH:mm:ss")
    if (n = "datetime")
        return FormatTime(, "yyyy-MM-dd HH:mm:ss")
    if (n = "env") {
        global IS_WORK_ENVIRONMENT
        try {
            if (IsSet(IS_WORK_ENVIRONMENT) && IS_WORK_ENVIRONMENT)
                return "work"
        } catch {
        }
        return "personal"
    }
    if (n = "selection")
        return ""
    return ""
}

PromptRender_ExtractPlaceholders(text) {
    found := Map()
    pos := 1
    while RegExMatch(text, "\{\{([^}]+)\}\}", &m, pos) {
        inner := Trim(m[1])
        if (inner != "" && !found.Has(inner))
            found[inner] := true
        pos := m.Pos(0) + m.Len(0)
    }
    return found
}

PromptRender_ResolveIncludePath(spec) {
    spec := Trim(spec)
    if (spec = "")
        return ""
    if (RegExMatch(spec, "^[a-zA-Z]:\\") || SubStr(spec, 1, 2) = "\\")
        return spec
    return A_ScriptDir "\" spec
}

PromptRender_ExpandIncludes(text, depth := 0, seen := "") {
    if (depth > 8)
        return text
    if (!IsObject(seen))
        seen := Map()
    out := text
    loop 32 {
        if !RegExMatch(out, "\{\{include:([^}]+)\}\}", &m)
            break
        spec := Trim(m[1])
        path := PromptRender_ResolveIncludePath(spec)
        key := StrLower(path)
        if (seen.Has(key)) {
            out := StrReplace(out, m[0], "[include cycle: " spec "]", , 1)
            continue
        }
        seen[key] := true
        chunk := ""
        if (path != "" && FileExist(path)) {
            try chunk := ReadUtf8File(path)
            catch {
                chunk := "[include missing: " spec "]"
            }
        } else {
            chunk := "[include missing: " spec "]"
        }
        chunk := PromptRender_ExpandIncludes(chunk, depth + 1, seen)
        out := StrReplace(out, m[0], chunk, , 1)
    }
    return out
}

PromptRender_ApplyBuiltins(text) {
    out := text
    for name, _ in PromptRender_BuiltinNames() {
        needle := "{{" . name . "}}"
        if InStr(out, needle)
            out := StrReplace(out, needle, PromptRender_BuiltinValue(name))
    }
    return out
}

PromptRender_ApplyValues(text, values) {
    out := text
    if (!IsObject(values))
        return out
    for name, val in values
        out := StrReplace(out, "{{" . name . "}}", val)
    return out
}

PromptRender_UnresolvedCustom(text, declared := "") {
    declaredMap := Map()
    if (declared != "") {
        for part in StrSplit(declared, ",") {
            n := Trim(part)
            if (n != "")
                declaredMap[StrLower(n)] := true
        }
    }
    need := []
    for name, _ in PromptRender_ExtractPlaceholders(text) {
        lower := StrLower(name)
        if (SubStr(lower, 1, 8) = "include:")
            continue
        if (PromptRender_IsBuiltin(lower))
            continue
        if (declaredMap.Has(lower))
            need.Push(name)
        else
            need.Push(name)
    }
    return need
}

PromptRender_FillDialog(names) {
    global g_PromptRenderFillGui, g_PromptRenderFillResult
    if (!IsObject(names) || names.Length = 0)
        return Map()
    g_PromptRenderFillResult := { ok: false, values: Map() }
    g_PromptRenderFillGui := Gui("+AlwaysOnTop +ToolWindow", "Prompt variables")
    g_PromptRenderFillGui.SetFont("s10", "Segoe UI")
    g_PromptRenderFillGui.Add("Text", "xm w360", "Fill variables before paste:")
    edits := Map()
    for name in names {
        g_PromptRenderFillGui.Add("Text", "xm w120", name)
        edits[name] := g_PromptRenderFillGui.Add("Edit", "yp w240")
    }
    g_PromptRenderFillGui.Add("Button", "xm+180 w80 Default", "OK").OnEvent("Click", (*) => PromptRender_FillOk(edits))
    g_PromptRenderFillGui.Add("Button", "x+8 yp w80", "Cancel").OnEvent("Click", PromptRender_FillCancel)
    g_PromptRenderFillGui.OnEvent("Close", PromptRender_FillCancel)
    g_PromptRenderFillGui.OnEvent("Escape", PromptRender_FillCancel)
    mon := UtilitySelector_ActiveMonitorWorkArea()
    g_PromptRenderFillGui.Show("Hide")
    g_PromptRenderFillGui.GetPos(, , &gw, &gh)
    g_PromptRenderFillGui.Show("x" (mon.left + (mon.width - gw) // 2) " y" (mon.top + (mon.height - gh) // 2))
    try WinWaitClose("ahk_id " g_PromptRenderFillGui.Hwnd)
    catch {
    }
    g_PromptRenderFillGui := false
    if (!g_PromptRenderFillResult.ok)
        return false
    return g_PromptRenderFillResult.values
}

PromptRender_FillOk(edits) {
    global g_PromptRenderFillResult, g_PromptRenderFillGui
    vals := Map()
    for name, ctrl in edits {
        try vals[name] := ctrl.Value
        catch {
            vals[name] := ""
        }
    }
    g_PromptRenderFillResult := { ok: true, values: vals }
    try g_PromptRenderFillGui.Destroy()
    catch {
    }
}

PromptRender_FillCancel(*) {
    global g_PromptRenderFillResult, g_PromptRenderFillGui
    g_PromptRenderFillResult := { ok: false, values: Map() }
    try g_PromptRenderFillGui.Destroy()
    catch {
    }
}

; Returns rendered body, or "" if user cancelled fill dialog.
PromptRender_PrepareBody(rawBody, prompt := false) {
    if (rawBody = "")
        return ""
    text := PromptRender_ExpandIncludes(rawBody)
    text := PromptRender_ApplyBuiltins(text)
    declared := (IsObject(prompt) && prompt.HasProp("variables")) ? prompt.variables : ""
    unresolved := []
    for name, _ in PromptRender_ExtractPlaceholders(text) {
        lower := StrLower(name)
        if (SubStr(lower, 1, 8) = "include:")
            continue
        if (PromptRender_IsBuiltin(lower))
            continue
        unresolved.Push(name)
    }
    if (unresolved.Length = 0)
        return text
    filled := PromptRender_FillDialog(unresolved)
    if (filled = false)
        return ""
    return PromptRender_ApplyValues(text, filled)
}

PromptRender_Prepare(prompt) {
    if (!IsObject(prompt))
        return ""
    raw := PromptData_ReadBody(prompt)
    return PromptRender_PrepareBody(raw, prompt)
}
