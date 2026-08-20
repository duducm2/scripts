; =============================================================================
; Utils module: prompt_data.ahk
; Persistent prompt registry for Utility Shortcuts (#!+U).
; Store: assets/data/prompts.ini  (same Load/Save mechanics as project_data_cursor.ahk)
; =============================================================================

global g_PromptEntries := []
global g_PromptDataCacheReady := false
global g_PromptDataCacheMtime := ""

PromptData_IniPath() {
    return A_ScriptDir "\assets\data\prompts.ini"
}

; Same assignment pool as g_HotstringCharSequence; 'l' is reserved for Gemini-arm, 'e' for edit.
global g_PromptCharSequence := ["1", "2", "3", "4", "5", "q", "w", "e", "r", "t", "a", "s", "d", "f", "g", "z", "x",
    "c", "v", "b", "6", "7", "8", "9", "0", "y", "u", "i", "o", "p", "h", "j", "k", "n", "m", ",", "."]
global g_PromptCharValid := Map()
for _promptChar in g_PromptCharSequence
    g_PromptCharValid[_promptChar] := true

PromptData_CharSequence() {
    global g_PromptCharSequence
    return g_PromptCharSequence
}

PromptData_IsValidChar(char) {
    global g_PromptCharValid
    if (char = "" || char = "l" || char = "e")
        return false
    return g_PromptCharValid.Has(char)
}

PromptData_NormalizeIniValue(val) {
    if (val = "" || val = "ERROR")
        return ""
    return val
}

PromptData_StripPathQuotes(path) {
    p := Trim(path)
    loop 8 {
        if (StrLen(p) < 2)
            break
        first := SubStr(p, 1, 1)
        last := SubStr(p, -1)
        if ((first = '"' && last = '"') || (first = "'" && last = "'"))
            p := Trim(SubStr(p, 2, StrLen(p) - 2))
        else
            break
    }
    return p
}

PromptData_ExtractQuotedOrPlainTokens(line) {
    tokens := []
    i := 1
    len := StrLen(line)
    while (i <= len) {
        while (i <= len) {
            ch := SubStr(line, i, 1)
            if (ch = " " || ch = "`t" || ch = "|")
                i++
            else
                break
        }
        if (i > len)
            break
        ch := SubStr(line, i, 1)
        if (ch = '"' || ch = "'") {
            q := ch
            i++
            start := i
            while (i <= len && SubStr(line, i, 1) != q)
                i++
            tokens.Push(PromptData_StripPathQuotes(SubStr(line, start, i - start)))
            if (i <= len)
                i++
        } else {
            start := i
            while (i <= len && SubStr(line, i, 1) != "|")
                i++
            tokens.Push(PromptData_StripPathQuotes(SubStr(line, start, i - start)))
            if (i <= len && SubStr(line, i, 1) = "|")
                i++
        }
    }
    return tokens
}

PromptData_ParsePathList(raw) {
    out := []
    if (IsObject(raw)) {
        for p in raw {
            n := PromptData_StripPathQuotes(p)
            if (n != "")
                out.Push(n)
        }
        return out
    }
    text := Trim(raw)
    if (text = "")
        return out
    text := StrReplace(StrReplace(text, "`r`n", "`n"), "`r", "`n")
    for line in StrSplit(text, "`n") {
        line := Trim(line)
        if (line = "")
            continue
        if (InStr(line, '"') || InStr(line, "'")) {
            for token in PromptData_ExtractQuotedOrPlainTokens(line) {
                if (token != "")
                    out.Push(token)
            }
        } else {
            for part in StrSplit(line, "|") {
                n := PromptData_StripPathQuotes(part)
                if (n != "")
                    out.Push(n)
            }
        }
    }
    return out
}

PromptData_JoinPathList(arr) {
    s := ""
    for p in PromptData_ParsePathList(arr) {
        if (s != "")
            s .= "|"
        s .= p
    }
    return s
}

PromptData_NewContextEntry(path, compact := 0, csvKeepFrom := 0, csvKeepTo := 0) {
    from := 0
    to := 0
    try from := Integer(csvKeepFrom)
    catch {
        from := 0
    }
    try to := Integer(csvKeepTo)
    catch {
        to := 0
    }
    if (from < 1 || to < 1 || from > to) {
        from := 0
        to := 0
    }
    return {
        path: PromptData_StripPathQuotes(path),
        compact: compact ? 1 : 0,
        csvKeepFrom: from,
        csvKeepTo: to
    }
}

PromptData_IsCsvPath(path) {
    if (path = "")
        return false
    SplitPath path, , , &ext
    return StrLower(ext) = "csv"
}

PromptData_ParseCsvKeepToken(token, &from, &to) {
    from := 0
    to := 0
    t := Trim(token)
    if (t = "")
        return
    if !RegExMatch(t, "^(\d+)-(\d+)$", &m)
        return
    try from := Integer(m[1])
    catch {
        from := 0
    }
    try to := Integer(m[2])
    catch {
        to := 0
    }
    if (from < 1 || to < 1 || from > to) {
        from := 0
        to := 0
    }
}

PromptData_FormatCsvKeep(from, to) {
    if (from < 1 || to < 1 || from > to)
        return ""
    return from . "-" . to
}

PromptData_SplitPipeTokens(raw) {
    text := Trim(raw)
    if (text = "")
        return []
    return StrSplit(text, "|")
}

PromptData_ParseContextEntries(raw) {
    out := []
    if (IsObject(raw)) {
        for item in raw {
            if (IsObject(item) && item.HasProp("path")) {
                p := PromptData_StripPathQuotes(item.path)
                if (p = "")
                    continue
                compact := (item.HasProp("compact") && item.compact) ? 1 : 0
                from := 0
                to := 0
                if (item.HasProp("csvKeepFrom"))
                    from := item.csvKeepFrom
                if (item.HasProp("csvKeepTo"))
                    to := item.csvKeepTo
                out.Push(PromptData_NewContextEntry(p, compact, from, to))
            } else {
                n := PromptData_StripPathQuotes(item)
                if (n != "")
                    out.Push(PromptData_NewContextEntry(n))
            }
        }
        return out
    }
    for n in PromptData_ParsePathList(raw) {
        if (n != "")
            out.Push(PromptData_NewContextEntry(n))
    }
    return out
}

PromptData_MergeContextEntries(paths, compactTokens, csvTokens) {
    out := []
    loop paths.Length {
        compact := 0
        if (A_Index <= compactTokens.Length && Trim(compactTokens[A_Index]) = "1")
            compact := 1
        from := 0
        to := 0
        if (A_Index <= csvTokens.Length)
            PromptData_ParseCsvKeepToken(csvTokens[A_Index], &from, &to)
        out.Push(PromptData_NewContextEntry(paths[A_Index], compact, from, to))
    }
    return out
}

PromptData_JoinContextPaths(entries) {
    s := ""
    for e in PromptData_ParseContextEntries(entries) {
        if (s != "")
            s .= "|"
        s .= e.path
    }
    return s
}

PromptData_JoinContextCompact(entries) {
    s := ""
    for e in PromptData_ParseContextEntries(entries) {
        if (s != "")
            s .= "|"
        s .= e.compact ? "1" : "0"
    }
    return s
}

PromptData_JoinContextCsvKeep(entries) {
    s := ""
    first := true
    for e in PromptData_ParseContextEntries(entries) {
        if (!first)
            s .= "|"
        first := false
        s .= PromptData_FormatCsvKeep(e.csvKeepFrom, e.csvKeepTo)
    }
    return s
}

PromptData_ResolveContextPath(path) {
    p := Trim(path)
    if (p = "")
        return ""
    if (RegExMatch(p, "^[a-zA-Z]:\\") || SubStr(p, 1, 2) = "\\")
        return p
    return A_ScriptDir "\" p
}

PromptData_ContextEntryPath(entry) {
    raw := ""
    if (!IsObject(entry))
        raw := Trim(entry)
    else
        raw := entry.HasProp("path") ? entry.path : ""
    return PromptData_ResolveContextPath(raw)
}

PromptData_ContextCompactLabel(entry) {
    if (!IsObject(entry) || !entry.HasProp("compact") || !entry.compact)
        return ""
    return "Yes"
}

PromptData_ContextCsvKeepLabel(entry) {
    if (!IsObject(entry))
        return ""
    if !PromptData_IsCsvPath(PromptData_ContextEntryPath(entry))
        return ""
    return PromptData_FormatCsvKeep(entry.HasProp("csvKeepFrom") ? entry.csvKeepFrom : 0, entry.HasProp("csvKeepTo") ?
        entry.csvKeepTo : 0)
}

PromptData_ContextEntryNeedsTransform(entry) {
    if (!IsObject(entry))
        return false
    if (entry.HasProp("compact") && entry.compact)
        return true
    path := PromptData_ContextEntryPath(entry)
    if !PromptData_IsCsvPath(path)
        return false
    from := entry.HasProp("csvKeepFrom") ? entry.csvKeepFrom : 0
    to := entry.HasProp("csvKeepTo") ? entry.csvKeepTo : 0
    return (from >= 1 && to >= 1 && from <= to)
}

PromptData_CompactedFileName(path) {
    SplitPath path, &name, , &ext, &nameNoExt
    if (nameNoExt = "")
        nameNoExt := name
    if (ext = "")
        return nameNoExt ".compacted"
    return nameNoExt ".compacted." ext
}

PromptData_UniqueCompactedName(path, usedMap) {
    name := PromptData_CompactedFileName(path)
    key := StrLower(name)
    if (!usedMap.Has(key)) {
        usedMap[key] := true
        return name
    }
    SplitPath name, , , &ext, &nameNoExt
    i := 2
    loop {
        cand := (ext != "") ? (nameNoExt "-" i "." ext) : (nameNoExt "-" i)
        ck := StrLower(cand)
        if (!usedMap.Has(ck)) {
            usedMap[ck] := true
            return cand
        }
        i += 1
    }
}

PromptData_ContextEntriesForCurrentEnv(prompt) {
    global IS_WORK_ENVIRONMENT
    if (!IsObject(prompt))
        return []
    arr := []
    if (IS_WORK_ENVIRONMENT)
        arr := prompt.HasProp("work_context_files") ? prompt.work_context_files : []
    else
        arr := prompt.HasProp("personal_context_files") ? prompt.personal_context_files : []
    return PromptData_ParseContextEntries(arr)
}

PromptData_ContextFilesForCurrentEnv(prompt) {
    paths := []
    for e in PromptData_ContextEntriesForCurrentEnv(prompt) {
        p := PromptData_ContextEntryPath(e)
        if (p != "")
            paths.Push(p)
    }
    return paths
}

PromptData_ReadIniRaw(iniPath, section, key) {
    raw := ""
    try raw := IniRead(iniPath, section, key, "")
    catch {
        raw := ""
    }
    return PromptData_NormalizeIniValue(raw)
}

PromptData_ReadContextFilesKey(iniPath, section, key) {
    return PromptData_ParsePathList(PromptData_ReadIniRaw(iniPath, section, key))
}

PromptData_ReadContextEntries(iniPath, section, prefix) {
    paths := PromptData_ReadContextFilesKey(iniPath, section, prefix . "ContextFiles")
    compactTokens := PromptData_SplitPipeTokens(PromptData_ReadIniRaw(iniPath, section, prefix . "ContextCompact"))
    csvTokens := PromptData_SplitPipeTokens(PromptData_ReadIniRaw(iniPath, section, prefix . "ContextCsvKeep"))
    return PromptData_MergeContextEntries(paths, compactTokens, csvTokens)
}

PromptData_WriteContextEntries(iniPath, section, prefix, entries) {
    list := PromptData_ParseContextEntries(entries)
    IniWrite(PromptData_JoinContextPaths(list), iniPath, section, prefix . "ContextFiles")
    IniWrite(PromptData_JoinContextCompact(list), iniPath, section, prefix . "ContextCompact")
    IniWrite(PromptData_JoinContextCsvKeep(list), iniPath, section, prefix . "ContextCsvKeep")
}

PromptData_ParseTags(raw) {
    out := []
    text := Trim(raw)
    if (text = "")
        return out
    for part in StrSplit(text, ",") {
        t := Trim(part)
        if (t != "")
            out.Push(t)
    }
    return out
}

PromptData_JoinTags(arr) {
    s := ""
    for t in PromptData_ParseTags(arr) {
        if (s != "")
            s .= ","
        s .= t
    }
    return s
}

PromptData_NormalizePasteMode(mode) {
    m := StrLower(Trim(mode))
    valid := Map("default", 1, "body_only", 1, "body_plus_clipboard", 1, "body_attach_clipboard", 1, "attach_only", 1,
        "auto_send", 1)
    if (m = "" || !valid.Has(m))
        return "default"
    return m
}

PromptData_NormalizeAttachAsTxt(val) {
    if (val = true || val = 1 || val = "1")
        return 1
    s := StrLower(Trim(val))
    if (s = "1" || s = "true" || s = "yes" || s = "on")
        return 1
    return 0
}

PromptData_AttachAsTxt(prompt) {
    if (!IsObject(prompt))
        return false
    return PromptData_NormalizeAttachAsTxt(prompt.HasProp("attachAsTxt") ? prompt.attachAsTxt : 0) = 1
}

PromptData_PasteMode(prompt) {
    if (!IsObject(prompt))
        return "default"
    return PromptData_NormalizePasteMode(prompt.HasProp("pasteMode") ? prompt.pasteMode : "")
}

PromptData_ResolveDraftPath(prompt) {
    if (!IsObject(prompt))
        return ""
    fp := prompt.HasProp("filePathDraft") ? prompt.filePathDraft : ""
    if (fp = "")
        return ""
    if (RegExMatch(fp, "^[a-zA-Z]:\\") || SubStr(fp, 1, 2) = "\\")
        return fp
    return A_ScriptDir "\" fp
}

PromptData_PromptRowMatches(prompt, query) {
    q := Trim(query)
    if (q = "")
        return true
    qLower := StrLower(q)
    if (InStr(StrLower(prompt.name), qLower))
        return true
    if (prompt.HasProp("tags")) {
        for t in PromptData_ParseTags(prompt.tags) {
            if (InStr(StrLower(t), qLower))
                return true
        }
    }
    if (prompt.HasProp("filePath") && InStr(StrLower(prompt.filePath), qLower))
        return true
    if (prompt.HasProp("category") && InStr(StrLower(prompt.category), qLower))
        return true
    return false
}

PromptData_NormalizeEntry(prompt) {
    if (!IsObject(prompt))
        return prompt
    prompt.tags := PromptData_JoinTags(prompt.HasProp("tags") ? prompt.tags : [])
    prompt.pasteMode := PromptData_NormalizePasteMode(prompt.HasProp("pasteMode") ? prompt.pasteMode : "")
    prompt.attachAsTxt := PromptData_NormalizeAttachAsTxt(prompt.HasProp("attachAsTxt") ? prompt.attachAsTxt : 0)
    prompt.variables := Trim(prompt.HasProp("variables") ? prompt.variables : "")
    prompt.filePathDraft := Trim(prompt.HasProp("filePathDraft") ? prompt.filePathDraft : "")
    prompt.personal_context_files := PromptData_ParseContextEntries(prompt.HasProp("personal_context_files") ?
        prompt.personal_context_files : [])
    prompt.work_context_files := PromptData_ParseContextEntries(prompt.HasProp("work_context_files") ?
        prompt.work_context_files : [])
    return prompt
}

PromptData_Invalidate() {
    global g_PromptEntries, g_PromptDataCacheReady, g_PromptDataCacheMtime
    g_PromptEntries := []
    g_PromptDataCacheReady := false
    g_PromptDataCacheMtime := ""
}

PromptData_FileMtime() {
    path := PromptData_IniPath()
    if (!FileExist(path))
        return ""
    mtime := ""
    try mtime := FileGetTime(path, "M")
    catch {
        mtime := ""
    }
    return mtime
}

PromptData_DefaultEntries() {
    list := [{ name: "✏️ Grammar & Spelling Corrector", char: "1", category: "General", author: "",
        filePath: "assets\prompt\grammar.txt", source: "file" }, { name: "🔲 Convert to Task", char: "2", category: "General",
            author: "",
            filePath: "assets\prompt\mtask.txt", source: "file" }, { name: "🤖 AI Text Optimizer", char: "3", category: "General",
                author: "",
                filePath: "assets\prompt\aiopt.txt", source: "file" }, { name: "📝 Summarize for Handoff", char: "r",
                    category: "General", author: "",
                    filePath: "assets\prompt\handoff-summary.txt", source: "file" }, { name: "📖 Creating mnemonic stories",
                        char: "4", category: "Mnemonic", author: "",
                        filePath: "story-prompt.txt", source: "technique" }, { name: "🎬 Transcript Youtube Video",
                            char: "5", category: "Mnemonic", author: "",
                            filePath: "video-transcription-prompt.txt", source: "technique" }, { name: "📝 Story reduction",
                                char: "a", category: "Mnemonic", author: "",
                                filePath: "story-reduction-prompt.txt", source: "technique" }, { name: "🧩 Punctual beast append",
                                    char: "p", category: "Mnemonic", author: "",
                                    filePath: "punctual-beast-append-prompt.txt", source: "technique" }, { name: "🛡️ Preserve background for image generation",
                                        char: "g", category: "Mnemonic", author: "",
                                        filePath: "image-background-preservation-prompt.txt", source: "technique" }, { name: "📊 PPT stage 1: content to slides CSV",
                                            char: "s", category: "General", author: "",
                                            filePath: "assets\prompt\ppt-content-to-slides-csv.txt", source: "file" }, { name: "🧩 PPT stage 2: slides to elements CSV",
                                                char: "i", category: "General", author: "",
                                                filePath: "assets\prompt\ppt-slides-to-elements-csv.txt", source: "file" }, { name: "📱 Prototype stage 1: content to screens CSV",
                                                    char: "w", category: "General", author: "",
                                                    filePath: "assets\prompt\proto-content-to-screens-csv.txt", source: "file" }, { name: "🧩 Prototype stage 2: screens to elements CSV",
                                                        char: "f", category: "General", author: "",
                                                        filePath: "assets\prompt\proto-screens-to-elements-csv.txt",
                                                        source: "file" }, { name: "🎨 Bosch brand-compliant image",
                                                            char: "q", category: "General", author: "",
                                                            filePath: "assets\prompt\bosch-brand-image.txt", source: "file" }, { name: "📋 Fill CSV from unstructured text",
                                                                char: "t", category: "General", author: "",
                                                                filePath: "assets\prompt\unstructured-to-csv.txt",
                                                                source: "file" }, { name: "📎 ClipAngel .cac export",
                                                                    char: "c", category: "General", author: "",
                                                                    filePath: "assets\prompt\clipangel-cac.txt", source: "file" }, { name: "📝 How-to steps CSV",
                                                                        char: "h", category: "General", author: "",
                                                                        filePath: "assets\prompt\howto-steps-csv.txt",
                                                                        source: "file" }
    ]
    for item in list {
        item.personal_context_files := []
        item.work_context_files := []
        item.tags := ""
        item.pasteMode := "default"
        item.attachAsTxt := 0
        item.variables := ""
        item.filePathDraft := ""
    }
    return list
}

; skipMtime: return in-memory cache without FileGetTime/IniRead (hotkey open on Google Drive).
PromptData_Load(force := false, skipMtime := false) {
    global g_PromptEntries, g_PromptDataCacheReady, g_PromptDataCacheMtime
    if (!force && skipMtime && g_PromptDataCacheReady)
        return g_PromptEntries
    path := PromptData_IniPath()
    if (!FileExist(path)) {
        if (!PromptData_Save(PromptData_DefaultEntries()))
            return g_PromptEntries
        force := true
    }
    mtime := PromptData_FileMtime()
    if (!force && g_PromptDataCacheReady && mtime = g_PromptDataCacheMtime)
        return g_PromptEntries

    list := []
    taken := Map()
    if (FileExist(path)) {
        idx := 1
        loop 200 {
            section := "Prompt_" . idx
            name := ""
            try name := IniRead(path, section, "Name", "")
            catch {
                break
            }
            if (name = "ERROR")
                break
            charVal := ""
            category := ""
            author := ""
            filePath := ""
            source := ""
            tags := ""
            pasteMode := ""
            attachAsTxt := ""
            variables := ""
            filePathDraft := ""
            try charVal := IniRead(path, section, "Char", "")
            try category := IniRead(path, section, "Category", "")
            try author := IniRead(path, section, "Author", "")
            try filePath := IniRead(path, section, "FilePath", "")
            try source := IniRead(path, section, "Source", "")
            try tags := IniRead(path, section, "Tags", "")
            try pasteMode := IniRead(path, section, "PasteMode", "")
            try attachAsTxt := IniRead(path, section, "AttachAsTxt", "0")
            try variables := IniRead(path, section, "Variables", "")
            try filePathDraft := IniRead(path, section, "FilePathDraft", "")
            personalFiles := PromptData_ReadContextEntries(path, section, "Personal")
            workFiles := PromptData_ReadContextEntries(path, section, "Work")
            name := PromptData_NormalizeIniValue(name)
            charVal := StrLower(PromptData_NormalizeIniValue(charVal))
            category := PromptData_NormalizeIniValue(category)
            author := PromptData_NormalizeIniValue(author)
            filePath := PromptData_NormalizeIniValue(filePath)
            source := StrLower(PromptData_NormalizeIniValue(source))
            tags := PromptData_NormalizeIniValue(tags)
            pasteMode := PromptData_NormalizeIniValue(pasteMode)
            attachAsTxt := PromptData_NormalizeIniValue(attachAsTxt)
            variables := PromptData_NormalizeIniValue(variables)
            filePathDraft := PromptData_NormalizeIniValue(filePathDraft)
            if (name = "" && filePath = "") {
                idx += 1
                continue
            }
            if (source = "")
                source := "file"
            if (category = "")
                category := "General"
            if (!PromptData_IsValidChar(charVal) || taken.Has(charVal))
                charVal := ""
            if (charVal != "")
                taken[charVal] := true
            list.Push(PromptData_NormalizeEntry({ name: name, char: charVal, category: category, author: author,
                filePath: filePath, source: source, tags: tags, pasteMode: pasteMode, attachAsTxt: attachAsTxt,
                variables: variables,
                filePathDraft: filePathDraft, personal_context_files: personalFiles,
                work_context_files: workFiles }))
            idx += 1
        }
    }

    g_PromptEntries := list
    g_PromptDataCacheReady := true
    g_PromptDataCacheMtime := mtime
    return g_PromptEntries
}

PromptData_Save(list) {
    global g_PromptEntries, g_PromptDataCacheReady, g_PromptDataCacheMtime
    path := PromptData_IniPath()
    try DirCreate(A_ScriptDir "\assets\data")
    catch {
    }
    try FileDelete(path)
    catch {
    }
    if (!IsObject(list) || list.Length = 0) {
        g_PromptEntries := []
        g_PromptDataCacheReady := true
        g_PromptDataCacheMtime := PromptData_FileMtime()
        return true
    }
    try {
        idx := 1
        for prompt in list {
            section := "Prompt_" . idx
            prompt := PromptData_NormalizeEntry(prompt)
            IniWrite(prompt.HasProp("name") ? prompt.name : "", path, section, "Name")
            IniWrite(prompt.HasProp("char") ? prompt.char : "", path, section, "Char")
            IniWrite(prompt.HasProp("category") ? prompt.category : "General", path, section, "Category")
            IniWrite(prompt.HasProp("author") ? prompt.author : "", path, section, "Author")
            IniWrite(prompt.HasProp("filePath") ? prompt.filePath : "", path, section, "FilePath")
            IniWrite(prompt.HasProp("source") ? prompt.source : "file", path, section, "Source")
            IniWrite(PromptData_JoinTags(prompt.HasProp("tags") ? prompt.tags : []), path, section, "Tags")
            IniWrite(prompt.pasteMode, path, section, "PasteMode")
            IniWrite(prompt.attachAsTxt, path, section, "AttachAsTxt")
            IniWrite(prompt.variables, path, section, "Variables")
            IniWrite(prompt.filePathDraft, path, section, "FilePathDraft")
            PromptData_WriteContextEntries(path, section, "Personal", prompt.personal_context_files)
            PromptData_WriteContextEntries(path, section, "Work", prompt.work_context_files)
            idx += 1
        }
    } catch {
        PromptData_Invalidate()
        return false
    }
    g_PromptEntries := list
    g_PromptDataCacheReady := true
    g_PromptDataCacheMtime := PromptData_FileMtime()
    return true
}

PromptData_ToStoredPath(absPath) {
    p := Trim(absPath)
    if (p = "")
        return ""
    scriptDir := RTrim(A_ScriptDir, "\")
    prefix := scriptDir "\"
    if (InStr(p, prefix) = 1)
        return SubStr(p, StrLen(prefix) + 1)
    return p
}

PromptData_ResolvePath(prompt) {
    if (!IsObject(prompt))
        return ""
    fp := prompt.HasProp("filePath") ? prompt.filePath : ""
    if (fp = "")
        return ""
    source := prompt.HasProp("source") ? StrLower(prompt.source) : "file"
    if (source = "technique")
        return GetTechniquePromptFilePath(fp)
    if (RegExMatch(fp, "^[a-zA-Z]:\\") || SubStr(fp, 1, 2) = "\\")
        return fp
    return A_ScriptDir "\" fp
}

PromptData_ReadBody(prompt) {
    path := PromptData_ResolvePath(prompt)
    if (path = "")
        return "[PROMPT FILE MISSING]"
    try {
        return ReadUtf8File(path)
    } catch {
        return "[PROMPT FILE MISSING: " path "]"
    }
}

PromptData_Sorted() {
    global g_PromptEntries
    PromptData_Load()
    list := []
    loop g_PromptEntries.Length {
        p := g_PromptEntries[A_Index]
        list.Push(PromptData_NormalizeEntry({
            name: p.name, char: p.char, category: p.category, author: p.author,
            filePath: p.filePath, source: p.source, listIndex: A_Index,
            tags: p.HasProp("tags") ? p.tags : "",
            pasteMode: p.HasProp("pasteMode") ? p.pasteMode : "default",
            attachAsTxt: p.HasProp("attachAsTxt") ? p.attachAsTxt : 0,
            variables: p.HasProp("variables") ? p.variables : "",
            filePathDraft: p.HasProp("filePathDraft") ? p.filePathDraft : "",
            personal_context_files: PromptData_ParseContextEntries(p.HasProp("personal_context_files") ?
                p.personal_context_files : []),
            work_context_files: PromptData_ParseContextEntries(p.HasProp("work_context_files") ? p.work_context_files :
                [])
        }))
    }
    PromptData_SortInPlace(list)
    return list
}

PromptData_CompareSorted(a, b) {
    c := StrCompare(a.category, b.category, "Locale")
    if (c != 0)
        return c
    return StrCompare(a.name, b.name, "Locale")
}

; Array.Sort is not available on this AHK v2 build; insertion-sort by Category then Name.
PromptData_SortInPlace(list) {
    if (!IsObject(list) || list.Length < 2)
        return
    i := 2
    while (i <= list.Length) {
        key := list[i]
        j := i - 1
        while (j >= 1 && PromptData_CompareSorted(list[j], key) > 0) {
            list[j + 1] := list[j]
            j -= 1
        }
        list[j + 1] := key
        i += 1
    }
}

PromptData_FindByChar(char) {
    ch := StrLower(char)
    for prompt in PromptData_Sorted() {
        if (prompt.char != "" && prompt.char = ch)
            return prompt
    }
    return ""
}

PromptData_Load()