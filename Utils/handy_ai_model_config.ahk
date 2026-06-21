; =============================================================================
; Utils module: handy_ai_model_config.ahk
; Handy AI model configuration map and persistence
; Extracted verbatim from Utils.ahk; loaded via #include into the
; Utils.ahk orchestrator / shared library entry point.
; =============================================================================

; =============================================================================
; AI Model Selection System for Handy
; =============================================================================
; Configuration: Maps selection numbers (1-4) to AI model names.
; These are partial name prefixes used to find buttons in the UIA tree (Type 50000, botão).
; Descriptions match Handy Transcription Models UI for quick verification.
; Slots 3-4: set Cohere Language on General tab before selecting the Cohere model (modelClickName).
global g_HandyAiModels := Map(
    1, { name: "Parakeet V2", desc: "English only. Best model for English speakers." },
    2, { name: "Parakeet V3", desc: "Fast and accurate. Multi-language." },
    3, { name: "Cohere English", desc: "Sets Cohere language to English (General), then activates Cohere.",
        cohereLanguage: "English", modelClickName: "Cohere" },
    4, { name: "Cohere Portuguese", desc: "Sets Cohere language to Portuguese (General), then activates Cohere.",
        cohereLanguage: "Portuguese", modelClickName: "Cohere" }
)

; Picker indices for ^!#9 / ^!#b; update g_HandyAiModels names if Handy renames models.
global HANDY_AI_SLOT_COHERE_PORTUGUESE := 4
global HANDY_AI_SLOT_COHERE_ENGLISH := 3

; GUI state for AI model selector
global g_AiModelSelectorGui := false
global g_AiModelSelectorActive := false
global g_AiModelBannerGui := false

; Persistent language flag indicator (slot 3 = UK, slot 4 = Brazil); see docs/standard_information_display.md "Persistent Indicators".
global g_LanguageFlagGuis := []
global g_LanguageFlagSlot := 0
global LANGUAGE_FLAG_WIDTH := 45                ; px (~30% smaller than 64); aspect kept via Picture h:-1
global LANGUAGE_FLAG_MARGIN := 20               ; px from work-area right/bottom

; Restore the persistent flag on script load (Reload-safe). Deferred so the GUI
; subsystem is ready and any concurrent auto-execute side-effects settle first.
SetTimer(LanguageFlag_InitFromPersistedSlot, -250)

Handy_GetHandyAiModelIniPath() {
    return A_ScriptDir "\data\handy_ai_model.ini"
}

; Returns persisted slot 1-4, or 0 if missing / invalid / not in g_HandyAiModels.
Handy_GetPersistedAiModelSlot() {
    global g_HandyAiModels
    path := Handy_GetHandyAiModelIniPath()
    s := ""
    try s := IniRead(path, "Handy", "Slot", "")
    if (s = "")
        return 0
    if !IsInteger(s)
        return 0
    n := Integer(s)
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

