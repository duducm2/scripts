; =============================================================================
; Utils module: hotstring_data.ahk
; Persistent pasteable-string registry for Utility Shortcuts (#!+U) Hotstrings.
; Store: assets/data/hotstrings.ini  (same Load/Save mechanics as project_data_cursor.ahk)
; =============================================================================

global g_HotstringEntries := []
global g_HotstringDataCacheReady := false
global g_HotstringDataCacheMtime := ""

HotstringData_IniPath() {
    return A_ScriptDir "\assets\data\hotstrings.ini"
}

global g_HotstringDataCharSequence := ["1", "2", "3", "4", "5", "q", "w", "e", "r", "t", "a", "s", "d", "f", "g", "z",
    "x", "c", "v", "b", "6", "7", "8", "9", "0", "y", "u", "i", "o", "p", "h", "j", "k", "l", "n", "m", ",", "."]
global g_HotstringDataCharValid := Map()
for _hsChar in g_HotstringDataCharSequence
    g_HotstringDataCharValid[_hsChar] := true

HotstringData_CharSequence() {
    global g_HotstringDataCharSequence
    return g_HotstringDataCharSequence
}

HotstringData_IsValidChar(char) {
    global g_HotstringDataCharValid
    if (char = "")
        return false
    return g_HotstringDataCharValid.Has(char)
}

HotstringData_NormalizeIniValue(val) {
    if (val = "" || val = "ERROR")
        return ""
    return val
}

HotstringData_Invalidate() {
    global g_HotstringEntries, g_HotstringDataCacheReady, g_HotstringDataCacheMtime
    g_HotstringEntries := []
    g_HotstringDataCacheReady := false
    g_HotstringDataCacheMtime := ""
}

HotstringData_FileMtime() {
    path := HotstringData_IniPath()
    if (!FileExist(path))
        return ""
    mtime := ""
    try mtime := FileGetTime(path, "M")
    catch {
        mtime := ""
    }
    return mtime
}

HotstringData_DefaultEntries() {
    return [{ name: "💼 Bosch Email", char: "1", text: "eduardo.figueiredo@br.bosch.com" }, { name: "📧 Gmail", char: "2",
        text: "edu.evangelista.figueiredo@gmail.com" }, { name: "🔗 my links", char: "m", text: "my links" }, { name: "📋 project management LA",
            char: "p", text: "project management LA" }, { name: "🔗 UX and CIP", char: "x", text: "UX and CIP" }, { name: "🎓 Trainings Management",
                char: "t", text: "GS_UX core team_Trainings Management" }
    ]
}

HotstringData_Load(force := false, skipMtime := false) {
    global g_HotstringEntries, g_HotstringDataCacheReady, g_HotstringDataCacheMtime
    if (!force && skipMtime && g_HotstringDataCacheReady)
        return g_HotstringEntries
    path := HotstringData_IniPath()
    if (!FileExist(path)) {
        if (!HotstringData_Save(HotstringData_DefaultEntries()))
            return g_HotstringEntries
        force := true
    }
    mtime := HotstringData_FileMtime()
    if (!force && g_HotstringDataCacheReady && mtime = g_HotstringDataCacheMtime)
        return g_HotstringEntries

    list := []
    taken := Map()
    if (FileExist(path)) {
        idx := 1
        loop 200 {
            section := "Hotstring_" . idx
            name := ""
            try name := IniRead(path, section, "Name", "")
            catch {
                break
            }
            if (name = "ERROR")
                break
            charVal := ""
            textVal := ""
            try charVal := IniRead(path, section, "Char", "")
            try textVal := IniRead(path, section, "Text", "")
            name := HotstringData_NormalizeIniValue(name)
            charVal := StrLower(HotstringData_NormalizeIniValue(charVal))
            textVal := HotstringData_NormalizeIniValue(textVal)
            if (name = "" && textVal = "") {
                idx += 1
                continue
            }
            if (!HotstringData_IsValidChar(charVal) || taken.Has(charVal))
                charVal := ""
            if (charVal != "")
                taken[charVal] := true
            list.Push({ name: name, char: charVal, text: textVal })
            idx += 1
        }
    }

    g_HotstringEntries := list
    g_HotstringDataCacheReady := true
    g_HotstringDataCacheMtime := mtime
    return g_HotstringEntries
}

HotstringData_Save(list) {
    global g_HotstringEntries, g_HotstringDataCacheReady, g_HotstringDataCacheMtime
    path := HotstringData_IniPath()
    try DirCreate(A_ScriptDir "\assets\data")
    catch {
    }
    try FileDelete(path)
    catch {
    }
    if (!IsObject(list) || list.Length = 0) {
        g_HotstringEntries := []
        g_HotstringDataCacheReady := true
        g_HotstringDataCacheMtime := HotstringData_FileMtime()
        return true
    }
    try {
        idx := 1
        for item in list {
            section := "Hotstring_" . idx
            IniWrite(item.HasProp("name") ? item.name : "", path, section, "Name")
            IniWrite(item.HasProp("char") ? item.char : "", path, section, "Char")
            IniWrite(item.HasProp("text") ? item.text : "", path, section, "Text")
            idx += 1
        }
    } catch {
        HotstringData_Invalidate()
        return false
    }
    g_HotstringEntries := list
    g_HotstringDataCacheReady := true
    g_HotstringDataCacheMtime := HotstringData_FileMtime()
    return true
}

HotstringData_FindByChar(char) {
    global g_HotstringEntries
    HotstringData_Load()
    ch := StrLower(char)
    for item in g_HotstringEntries {
        if (item.char != "" && item.char = ch)
            return item
    }
    return ""
}

HotstringData_Load()