; =============================================================================

; Utils module: handy_ai_model_config.ahk

; Handy AI model configuration map and persistence

; Extracted verbatim from Utils.ahk; loaded via #include into the

; Utils.ahk orchestrator / shared library entry point.

; =============================================================================

; =============================================================================

; AI Model Selection System for Handy

; =============================================================================

; Configuration: Maps selection numbers (1-3) to AI model names.

; These are partial name prefixes used to find buttons in the UIA tree (Type 50000, botão).

; Descriptions match Handy Transcription Models UI for quick verification.

global g_HandyAiModels := Map(
    1, { name: "Parakeet Unified EN", desc: "Fast, accurate live English transcription." },
    2, { name: "Nemotron Streaming", desc: "Live multilingual transcription (Portuguese)." },
    3, { name: "Cohere Transcribe", desc: "Highest accuracy, 14 languages, slower." }
)

; Picker indices for ^!#9 / ^!#b; update g_HandyAiModels names if Handy renames models.

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

; INI schema: 2 = current 1-3 slot lineup (Slot=2 is Portuguese; no legacy remap on read).

global HANDY_AI_INI_SCHEMA_VERSION := 2

; Model switch retry / quality gate (ExecuteHandyAiModelSelection).

global HANDY_AI_MODEL_MAX_ATTEMPTS := 3

global HANDY_AI_MODEL_RETRY_DELAY_MS := 500

; GUI state for AI model selector

global g_AiModelSelectorGui := false

global g_AiModelSelectorActive := false

global g_AiModelBannerGui := false

; In-process slot cache (shortcuts, modal highlight, flag init share this after INI load).

global g_HandyAiPersistedSlot := 0

; Persistent language flag indicator (slot 1 = UK, slot 2 = Brazil, slot 3 = multi); see docs/standard_information_display.md "Persistent Indicators".

global g_LanguageFlagGuis := []

global g_LanguageFlagSlot := 0

global LANGUAGE_FLAG_WIDTH := 45                ; px (~30% smaller than 64); aspect kept via Picture h:-1

global LANGUAGE_FLAG_MARGIN := 20               ; px from work-area right/bottom

; AppLaunchers.ahk is the single owner for the language flag and Handy model hotkeys.

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

; Read slot from INI (bypasses in-process cache). Migrates and stamps SchemaVersion=2 once when needed.

Handy_ReadPersistedAiModelSlotFromIni() {

    global g_HandyAiModels, HANDY_AI_INI_SCHEMA_VERSION

    path := Handy_GetHandyAiModelIniPath()

    s := ""

    try s := IniRead(path, "Handy", "Slot", "")

    if (s = "" || !IsInteger(s))
        return 0

    rawSlot := Integer(s)

    schemaVersion := Handy_ReadIniSchemaVersion()

    if (schemaVersion >= HANDY_AI_INI_SCHEMA_VERSION) {

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

; Returns persisted slot 1-3, or 0 if missing / invalid / not in g_HandyAiModels.

Handy_GetPersistedAiModelSlot() {

    global g_HandyAiPersistedSlot, g_HandyAiModels

    if (g_HandyAiPersistedSlot >= 1 && g_HandyAiPersistedSlot <= 3 && g_HandyAiModels.Has(g_HandyAiPersistedSlot))
        return g_HandyAiPersistedSlot

    slot := Handy_ReadPersistedAiModelSlotFromIni()

    g_HandyAiPersistedSlot := (slot >= 1 && slot <= 3) ? slot : 0

    return slot

}

Handy_SetPersistedAiModelSlot(slot) {

    return Handy_PersistSlotToIni(slot)

}

; Warm in-process slot cache from INI at include time (modal correct before flag timer).

Handy_GetPersistedAiModelSlot()

; Restore the persistent flag on script load (Reload-safe). Deferred so the GUI

; subsystem is ready and any concurrent auto-execute side-effects settle first.

if (HandyAi_IsOwnerProcess())
    SetTimer(LanguageFlag_InitFromPersistedSlot, -250)