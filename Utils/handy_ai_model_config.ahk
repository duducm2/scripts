; =============================================================================
; Utils module: handy_ai_model_config.ahk
; Handy AI model configuration map and persistence
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; AI Model Selection System for Handy
; =============================================================================
; Slots 1-3 are immutable builtins. Slots 4+ are user ExtraModels (pipe list in INI).
; Names are substrings used to find buttons in the UIA tree (Type 50000).

global HANDY_AI_BUILTIN_COUNT := 3

; Seed builtins; HandyAiModels_LoadCatalog() rebuilds g_HandyAiModels at include time.
global g_HandyAiBuiltinModels := Map(
    1, { name: "Parakeet Unified EN", desc: "Fast, accurate live English transcription." },
    2, { name: "Nemotron Streaming", desc: "Live multilingual transcription (Portuguese)." },
    3, { name: "Cohere Transcribe", desc: "Highest accuracy, 14 languages, slower." }
)

global g_HandyAiModels := Map()

; Picker indices for ^!#9 / ^!#b; update g_HandyAiBuiltinModels names if Handy renames models.
global HANDY_AI_SLOT_ENGLISH := 1
global HANDY_AI_SLOT_PORTUGUESE := 2
global HANDY_AI_SLOT_MULTILANG := 3

; Send dictation? [B]: toggle Parakeet Unified EN (1) <-> Cohere Transcribe (3); never Nemotron (2).
; From Nemotron/unknown -> Cohere (bilingual correction path).
Handy_GetDictationToggleTargetSlot() {
    global HANDY_AI_SLOT_ENGLISH, HANDY_AI_SLOT_MULTILANG
    current := Handy_GetPersistedAiModelSlot()
    if (current = HANDY_AI_SLOT_ENGLISH)
        return HANDY_AI_SLOT_MULTILANG
    if (current = HANDY_AI_SLOT_MULTILANG)
        return HANDY_AI_SLOT_ENGLISH
    return HANDY_AI_SLOT_MULTILANG
}

; INI schema: 2 = current 1-3 slot lineup; 3 = ExtraModels pipe list for user slots 4+.
; Legacy slot remap only when SchemaVersion < 2.
global HANDY_AI_LEGACY_SLOT_SCHEMA := 2
global HANDY_AI_INI_SCHEMA_VERSION := 3

; Model switch retry / quality gate (ExecuteHandyAiModelSelection).
global HANDY_AI_MODEL_MAX_ATTEMPTS := 3
global HANDY_AI_MODEL_RETRY_DELAY_MS := 500

; GUI state for AI model selector
global g_AiModelSelectorGui := false
global g_AiModelSelectorActive := false
global g_AiModelBannerGui := false

; Set true while ExecuteHandyAiModelSelection runs (guards overlapping #!+C / shortcuts).
global g_HandyModelSwitchBusy := false

; In-process slot cache (shortcuts and the model picker share this after INI load).
global g_HandyAiPersistedSlot := 0

; AppLaunchers.ahk is the single owner for Handy model hotkeys and the recording flag.
HandyAi_IsOwnerProcess() {
    return A_ScriptName = "AppLaunchers.ahk"
}

; Disable Handy model hotkeys in non-owner processes (called after hotkey modules load in Utils.ahk).
HandyAi_ConfigureProcessOwnership() {
    if (HandyAi_IsOwnerProcess())
        return
    try Hotkey("#!+C", "Off")
    try Hotkey("^!#9", "Off")
    try Hotkey("^!#b", "Off")
}

Handy_GetHandyAiModelIniPath() {
    return A_ScriptDir "\assets\data\handy_ai_model.ini"
}

Handy_ReadIniSchemaVersion() {
    path := Handy_GetHandyAiModelIniPath()
    s := ""
    try s := IniRead(path, "Handy", "SchemaVersion", "")
    if (s = "" || !IsInteger(s))
        return 0
    return Integer(s)
}

; Legacy migration for unversioned INI only (SchemaVersion missing or 1).
; Slots 1-3 are current schema; only old Cohere slots 3/4 are remapped.
Handy_MigrateLegacyAiModelSlot(n) {
    switch n {
        case 3: return 1   ; old Cohere English -> Parakeet Unified EN
        case 4: return 2   ; old Cohere Portuguese -> Nemotron Streaming
        default: return n
    }
}

; --- Catalog: builtins + ExtraModels ------------------------------------------------

; Letters for slots 10+; skip a/e (add/edit hotkeys in the picker).
HandyAiModels_ItemLetters() {
    return "bcdfghijklmnopqrstuvwxyz"
}

HandyAiModels_MaxSlots() {
    global HANDY_AI_BUILTIN_COUNT
    return 9 + StrLen(HandyAiModels_ItemLetters())
}

HandyAiModels_LabelForIndex(index) {
    if (index >= 1 && index <= 9)
        return String(index)
    letters := HandyAiModels_ItemLetters()
    offset := index - 9
    if (offset >= 1 && offset <= StrLen(letters))
        return SubStr(letters, offset, 1)
    return String(index)
}

HandyAiModels_IndexFromKey(key) {
    key := StrLower(Trim(key))
    if (RegExMatch(key, "^[1-9]$"))
        return Integer(key)
    letters := HandyAiModels_ItemLetters()
    pos := InStr(letters, key, true)
    if (pos)
        return 9 + pos
    return 0
}

HandyAiModels_Count() {
    global g_HandyAiModels
    return g_HandyAiModels.Count
}

HandyAiModels_IsBuiltin(slot) {
    global HANDY_AI_BUILTIN_COUNT
    return (slot >= 1 && slot <= HANDY_AI_BUILTIN_COUNT)
}

HandyAiModels_ParseExtrasPipe(raw) {
    models := []
    raw := Trim(raw)
    if (raw = "" || raw = "ERROR")
        return models
    for part in StrSplit(raw, "|") {
        name := Trim(part)
        if (name != "")
            models.Push(name)
    }
    return models
}

HandyAiModels_ExtrasToPipe(extras) {
    if (!IsObject(extras) || extras.Length = 0)
        return ""
    out := ""
    for name in extras {
        name := Trim(name)
        if (name = "")
            continue
        out .= (out = "" ? "" : "|") . name
    }
    return out
}

HandyAiModels_CollectExtras() {
    global g_HandyAiModels, HANDY_AI_BUILTIN_COUNT
    extras := []
    slot := HANDY_AI_BUILTIN_COUNT + 1
    while g_HandyAiModels.Has(slot) {
        extras.Push(g_HandyAiModels[slot].name)
        slot++
    }
    return extras
}

HandyAiModels_NameExists(name, excludeSlot := 0) {
    global g_HandyAiModels
    name := Trim(name)
    for slot, model in g_HandyAiModels {
        if (excludeSlot > 0 && slot = excludeSlot)
            continue
        if (model.name = name)
            return true
    }
    return false
}

HandyAiModels_LoadCatalog() {
    global g_HandyAiModels, g_HandyAiBuiltinModels, HANDY_AI_BUILTIN_COUNT

    g_HandyAiModels := Map()
    loop HANDY_AI_BUILTIN_COUNT {
        slot := A_Index
        if !g_HandyAiBuiltinModels.Has(slot)
            continue
        src := g_HandyAiBuiltinModels[slot]
        g_HandyAiModels[slot] := { name: src.name, desc: src.desc, builtin: true }
    }

    path := Handy_GetHandyAiModelIniPath()
    raw := ""
    try raw := IniRead(path, "Handy", "ExtraModels", "")
    catch {
        raw := ""
    }
    extras := HandyAiModels_ParseExtrasPipe(raw)
    maxSlots := HandyAiModels_MaxSlots()
    slot := HANDY_AI_BUILTIN_COUNT + 1
    for name in extras {
        if (slot > maxSlots)
            break
        if (HandyAiModels_NameExists(name))
            continue
        g_HandyAiModels[slot] := { name: name, desc: "", builtin: false }
        slot++
    }
}

HandyAiModels_SaveExtras() {
    global HANDY_AI_INI_SCHEMA_VERSION
    path := Handy_GetHandyAiModelIniPath()
    try {
        SplitPath(path, , &dataDir)
        if !DirExist(dataDir)
            DirCreate(dataDir)
        IniWrite(HandyAiModels_ExtrasToPipe(HandyAiModels_CollectExtras()), path, "Handy", "ExtraModels")
        IniWrite(String(HANDY_AI_INI_SCHEMA_VERSION), path, "Handy", "SchemaVersion")
        return true
    } catch {
        return false
    }
}

HandyAiModels_Add(name) {
    global g_HandyAiModels, HANDY_AI_BUILTIN_COUNT
    name := Trim(name)
    if (name = "")
        return false
    if (HandyAiModels_NameExists(name))
        return false
    if (HandyAiModels_Count() >= HandyAiModels_MaxSlots())
        return false
    slot := HANDY_AI_BUILTIN_COUNT + 1
    while g_HandyAiModels.Has(slot)
        slot++
    g_HandyAiModels[slot] := { name: name, desc: "", builtin: false }
    return HandyAiModels_SaveExtras()
}

HandyAiModels_Rename(slot, name) {
    global g_HandyAiModels
    if (HandyAiModels_IsBuiltin(slot))
        return false
    if !g_HandyAiModels.Has(slot)
        return false
    name := Trim(name)
    if (name = "")
        return false
    if (HandyAiModels_NameExists(name, slot))
        return false
    g_HandyAiModels[slot].name := name
    g_HandyAiModels[slot].desc := ""
    return HandyAiModels_SaveExtras()
}

HandyAiModels_Remove(slot) {
    global g_HandyAiModels, g_HandyAiPersistedSlot, HANDY_AI_BUILTIN_COUNT, HANDY_AI_SLOT_ENGLISH
    if (HandyAiModels_IsBuiltin(slot))
        return false
    if !g_HandyAiModels.Has(slot)
        return false

    wasCurrent := (Handy_GetPersistedAiModelSlot() = slot)
    extras := HandyAiModels_CollectExtras()
    extraIndex := slot - HANDY_AI_BUILTIN_COUNT
    if (extraIndex < 1 || extraIndex > extras.Length)
        return false
    extras.RemoveAt(extraIndex)

    ; Rebuild slots 4+ densely after remove.
    slotNum := HANDY_AI_BUILTIN_COUNT + 1
    while g_HandyAiModels.Has(slotNum) {
        g_HandyAiModels.Delete(slotNum)
        slotNum++
    }
    slotNum := HANDY_AI_BUILTIN_COUNT + 1
    for name in extras {
        g_HandyAiModels[slotNum] := { name: name, desc: "", builtin: false }
        slotNum++
    }

    if !HandyAiModels_SaveExtras()
        return false

    if (wasCurrent) {
        ; Fall back to English builtin when the active extra is removed.
        Handy_PersistSlotToIni(HANDY_AI_SLOT_ENGLISH)
    } else if (g_HandyAiPersistedSlot > HANDY_AI_BUILTIN_COUNT) {
        ; Remap persisted extra slot if it sat after the removed index.
        current := g_HandyAiPersistedSlot
        if (current > slot) {
            newSlot := current - 1
            if g_HandyAiModels.Has(newSlot)
                Handy_PersistSlotToIni(newSlot)
        }
    }
    return true
}

; Write Slot + SchemaVersion to INI and update in-process cache. Returns true on success.
Handy_PersistSlotToIni(slot) {
    global g_HandyAiModels, g_HandyAiPersistedSlot, HANDY_AI_INI_SCHEMA_VERSION
    if !g_HandyAiModels.Has(slot)
        return false
    path := Handy_GetHandyAiModelIniPath()
    try {
        SplitPath(path, , &dataDir)
        if !DirExist(dataDir)
            DirCreate(dataDir)
        IniWrite(String(slot), path, "Handy", "Slot")
        IniWrite(String(HANDY_AI_INI_SCHEMA_VERSION), path, "Handy", "SchemaVersion")
        g_HandyAiPersistedSlot := slot
        return true
    } catch {
        return false
    }
}

; Read slot from INI (bypasses in-process cache). Migrates legacy slots when SchemaVersion < 2.
Handy_ReadPersistedAiModelSlotFromIni() {
    global g_HandyAiModels, HANDY_AI_LEGACY_SLOT_SCHEMA
    path := Handy_GetHandyAiModelIniPath()
    s := ""
    try s := IniRead(path, "Handy", "Slot", "")
    if (s = "" || !IsInteger(s))
        return 0
    rawSlot := Integer(s)
    schemaVersion := Handy_ReadIniSchemaVersion()
    if (schemaVersion >= HANDY_AI_LEGACY_SLOT_SCHEMA) {
        n := rawSlot
    } else {
        n := Handy_MigrateLegacyAiModelSlot(rawSlot)
        if (g_HandyAiModels.Has(n))
            Handy_PersistSlotToIni(n)
    }
    if !g_HandyAiModels.Has(n)
        return 0
    return n
}

; Returns persisted slot, or 0 if missing / invalid / not in g_HandyAiModels.
Handy_GetPersistedAiModelSlot() {
    global g_HandyAiPersistedSlot, g_HandyAiModels
    if (g_HandyAiPersistedSlot >= 1 && g_HandyAiModels.Has(g_HandyAiPersistedSlot))
        return g_HandyAiPersistedSlot
    slot := Handy_ReadPersistedAiModelSlotFromIni()
    g_HandyAiPersistedSlot := g_HandyAiModels.Has(slot) ? slot : 0
    return g_HandyAiPersistedSlot
}

Handy_SetPersistedAiModelSlot(slot) {
    return Handy_PersistSlotToIni(slot)
}

; Build catalog then warm in-process slot cache (modal highlight before any take).
HandyAiModels_LoadCatalog()
Handy_GetPersistedAiModelSlot()