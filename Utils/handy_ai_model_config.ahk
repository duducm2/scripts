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

; Model switch retry / quality gate (ExecuteHandyAiModelSelection).
global HANDY_AI_MODEL_MAX_ATTEMPTS := 3
global HANDY_AI_MODEL_RETRY_DELAY_MS := 500

; GUI state for AI model selector
global g_AiModelSelectorGui := false
global g_AiModelSelectorActive := false
global g_AiModelBannerGui := false

; Persistent language flag indicator (slot 1 = UK, slot 2 = Brazil, slot 3 = multi); see docs/standard_information_display.md "Persistent Indicators".
global g_LanguageFlagGuis := []
global g_LanguageFlagSlot := 0
global LANGUAGE_FLAG_WIDTH := 45                ; px (~30% smaller than 64); aspect kept via Picture h:-1
global LANGUAGE_FLAG_MARGIN := 20               ; px from work-area right/bottom

; Restore the persistent flag on script load (Reload-safe). Deferred so the GUI
; subsystem is ready and any concurrent auto-execute side-effects settle first.
SetTimer(LanguageFlag_InitFromPersistedSlot, -250)

Handy_GetHandyAiModelIniPath() {
    return A_ScriptDir "\assets\data\handy_ai_model.ini"
}

; Map legacy persisted slots (pre 3-model lineup) to current slots.
Handy_MigrateLegacyAiModelSlot(n) {
    switch n {
        case 3: return 1   ; old Cohere English -> Parakeet Unified EN
        case 4: return 2   ; old Cohere Portuguese -> Nemotron Streaming
        case 2: return 3   ; old Parakeet V3 multi -> Cohere Transcribe
        default: return n
    }
}

; Returns persisted slot 1-3, or 0 if missing / invalid / not in g_HandyAiModels.
Handy_GetPersistedAiModelSlot() {
    global g_HandyAiModels
    path := Handy_GetHandyAiModelIniPath()
    s := ""
    try s := IniRead(path, "Handy", "Slot", "")
    if (s = "")
        return 0
    if !IsInteger(s)
        return 0
    n := Handy_MigrateLegacyAiModelSlot(Integer(s))
    if !g_HandyAiModels.Has(n)
        return 0
    return n
}

Handy_SetPersistedAiModelSlot(slot) {
    global g_HandyAiModels
    if !g_HandyAiModels.Has(slot)
        return
    path := Handy_GetHandyAiModelIniPath()
    try {
        dataDir := A_ScriptDir "\data"
        if !DirExist(dataDir)
            DirCreate(dataDir)
        IniWrite(String(slot), path, "Handy", "Slot")
    } catch {
    }
}
