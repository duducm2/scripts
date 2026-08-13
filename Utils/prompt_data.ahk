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

; Same assignment pool as g_HotstringCharSequence; 'l' is reserved for Gemini-arm in Prompts.
PromptData_CharSequence() {
    return ["1", "2", "3", "4", "5", "q", "w", "e", "r", "t", "a", "s", "d", "f", "g", "z", "x",
        "c", "v", "b", "6", "7", "8", "9", "0", "y", "u", "i", "o", "p", "h", "j", "k", "n", "m", ",", "."]
}

PromptData_IsValidChar(char) {
    if (char = "" || char = "l")
        return false
    for c in PromptData_CharSequence() {
        if (c = char)
            return true
    }
    return false
}

PromptData_NormalizeIniValue(val) {
    if (val = "" || val = "ERROR")
        return ""
    return val
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
    return [{ name: "✏️ Grammar & Spelling Corrector", char: "1", category: "General", author: "",
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
                                                char: "e", category: "General", author: "",
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
}

PromptData_Load(force := false) {
    global g_PromptEntries, g_PromptDataCacheReady, g_PromptDataCacheMtime
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
            try charVal := IniRead(path, section, "Char", "")
            try category := IniRead(path, section, "Category", "")
            try author := IniRead(path, section, "Author", "")
            try filePath := IniRead(path, section, "FilePath", "")
            try source := IniRead(path, section, "Source", "")
            name := PromptData_NormalizeIniValue(name)
            charVal := StrLower(PromptData_NormalizeIniValue(charVal))
            category := PromptData_NormalizeIniValue(category)
            author := PromptData_NormalizeIniValue(author)
            filePath := PromptData_NormalizeIniValue(filePath)
            source := StrLower(PromptData_NormalizeIniValue(source))
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
            list.Push({ name: name, char: charVal, category: category, author: author, filePath: filePath,
                source: source })
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
            IniWrite(prompt.HasProp("name") ? prompt.name : "", path, section, "Name")
            IniWrite(prompt.HasProp("char") ? prompt.char : "", path, section, "Char")
            IniWrite(prompt.HasProp("category") ? prompt.category : "General", path, section, "Category")
            IniWrite(prompt.HasProp("author") ? prompt.author : "", path, section, "Author")
            IniWrite(prompt.HasProp("filePath") ? prompt.filePath : "", path, section, "FilePath")
            IniWrite(prompt.HasProp("source") ? prompt.source : "file", path, section, "Source")
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
        list.Push({
            name: p.name, char: p.char, category: p.category, author: p.author,
            filePath: p.filePath, source: p.source, listIndex: A_Index
        })
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