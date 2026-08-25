; =============================================================================
; Lib: AiCompanionModels.ahk
; Per-companion Fast / Deep / Models persistence for Gemini, Gemini Enterprise,
; and Copilot Web. Included from Utils.ahk after CopilotWeb + GeminiEnterprise.
; =============================================================================

AI_COMPANION_GEMINI := "Gemini"
AI_COMPANION_ENTERPRISE := "GeminiEnterprise"
AI_COMPANION_COPILOT := "CopilotWeb"

; companion → { fast, deep, models (Array of names) }; filled on first read / save.
global g_AiCompanionModelsCache := Map()

AiCompanionModels_GetIniPath() {
    return A_ScriptDir "\assets\data\ai_companion_models.ini"
}

AiCompanionModels_IsValidCompanion(companion) {
    return companion = AI_COMPANION_GEMINI || companion = AI_COMPANION_ENTERPRISE || companion = AI_COMPANION_COPILOT
}

AiCompanionModels_DefaultConfig(companion) {
    if (companion = AI_COMPANION_GEMINI)
        return { fast: "1.5 Flash-Lite", deep: "1.5 Pro", models: ["1.5 Flash", "Extended thinking"] }
    if (companion = AI_COMPANION_ENTERPRISE)
        return { fast: "", deep: "3.1 Pro", models: [] }
    if (companion = AI_COMPANION_COPILOT)
        return { fast: "", deep: "Think deeper", models: [] }
    return { fast: "", deep: "", models: [] }
}

; Optional Models-list cleanup only. Never rewrite Fast/Deep — users set those via Shift+L f/d.
AiCompanionModels_MigrateObsoleteGeminiName(name) {
    name := Trim(name)
    if (name = "")
        return ""
    if (name = "Thinking level")
        return "Extended thinking"
    return name
}

AiCompanionModels_MigrateObsoleteGeminiConfig(fast, deep, models) {
    migrated := false
    ; Preserve Fast/Deep exactly as stored (do not map 3.1 Pro → 1.5 Pro, etc.).
    newFast := Trim(fast)
    newDeep := Trim(deep)
    newModels := []
    for name in models {
        mapped := AiCompanionModels_MigrateObsoleteGeminiName(name)
        if (mapped != name)
            migrated := true
        already := false
        for existing in newModels {
            if (existing = mapped) {
                already := true
                break
            }
        }
        if (!already && mapped != "")
            newModels.Push(mapped)
    }
    return { fast: newFast, deep: newDeep, models: newModels, migrated: migrated }
}

AiCompanionModels_ParseModelsPipe(raw) {
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

AiCompanionModels_ModelsToPipe(models) {
    if (!IsObject(models) || models.Length = 0)
        return ""
    out := ""
    for name in models {
        name := Trim(name)
        if (name = "")
            continue
        out .= (out = "" ? "" : "|") . name
    }
    return out
}

AiCompanionModels_InvalidateCache(companion := "") {
    global g_AiCompanionModelsCache
    if (companion = "") {
        g_AiCompanionModelsCache := Map()
        return
    }
    if g_AiCompanionModelsCache.Has(companion)
        g_AiCompanionModelsCache.Delete(companion)
}

AiCompanionModels_Load(companion) {
    global g_AiCompanionModelsCache
    if !AiCompanionModels_IsValidCompanion(companion)
        return AiCompanionModels_DefaultConfig(companion)
    if g_AiCompanionModelsCache.Has(companion)
        return g_AiCompanionModelsCache[companion]

    defaults := AiCompanionModels_DefaultConfig(companion)
    iniPath := AiCompanionModels_GetIniPath()
    fast := ""
    deep := ""
    modelsRaw := ""
    sectionMissing := false
    try {
        fast := IniRead(iniPath, companion, "Fast", "")
        deep := IniRead(iniPath, companion, "Deep", "")
        modelsRaw := IniRead(iniPath, companion, "Models", "")
    } catch {
        sectionMissing := true
    }
    ; Fresh / missing section → seed defaults to disk.
    if (sectionMissing || (fast = "" && deep = "" && (modelsRaw = "" || modelsRaw = "ERROR"))) {
        ; Enterprise/Copilot intentionally allow empty Fast; only seed when file section absent.
        if (sectionMissing || !FileExist(iniPath)) {
            cfg := defaults
            AiCompanionModels_Save(companion, cfg.fast, cfg.deep, cfg.models)
            return g_AiCompanionModelsCache[companion]
        }
    }
    if (fast = "ERROR")
        fast := ""
    if (deep = "ERROR")
        deep := ""
    if (modelsRaw = "ERROR")
        modelsRaw := ""
    ; If Deep empty, fall back to default deep so Shift+M still works.
    if (Trim(deep) = "")
        deep := defaults.deep
    models := AiCompanionModels_ParseModelsPipe(modelsRaw)
    if (companion = AI_COMPANION_GEMINI) {
        mig := AiCompanionModels_MigrateObsoleteGeminiConfig(fast, deep, models)
        fast := mig.fast
        deep := mig.deep
        models := mig.models
        if (mig.migrated)
            AiCompanionModels_Save(companion, fast, deep, models)
    }
    cfg := { fast: Trim(fast), deep: Trim(deep), models: models }
    g_AiCompanionModelsCache[companion] := cfg
    return cfg
}

AiCompanionModels_Save(companion, fast, deep, models) {
    global g_AiCompanionModelsCache
    if !AiCompanionModels_IsValidCompanion(companion)
        return false
    iniPath := AiCompanionModels_GetIniPath()
    try DirCreate(A_ScriptDir "\assets\data")
    catch {
    }
    try {
        IniWrite(Trim(fast), iniPath, companion, "Fast")
        IniWrite(Trim(deep), iniPath, companion, "Deep")
        IniWrite(AiCompanionModels_ModelsToPipe(models), iniPath, companion, "Models")
    } catch {
        return false
    }
    cfg := { fast: Trim(fast), deep: Trim(deep), models: IsObject(models) ? models : [] }
    g_AiCompanionModelsCache[companion] := cfg
    return true
}

AiCompanionModels_GetFast(companion) {
    return AiCompanionModels_Load(companion).fast
}

AiCompanionModels_GetDeep(companion) {
    return AiCompanionModels_Load(companion).deep
}

AiCompanionModels_GetModels(companion) {
    return AiCompanionModels_Load(companion).models
}

AiCompanionModels_SetRole(companion, role, modelName) {
    cfg := AiCompanionModels_Load(companion)
    modelName := Trim(modelName)
    role := StrLower(Trim(role))
    if (role = "fast")
        return AiCompanionModels_Save(companion, modelName, cfg.deep, cfg.models)
    if (role = "deep")
        return AiCompanionModels_Save(companion, cfg.fast, modelName, cfg.models)
    return false
}

; Append to Models list. roleTag is optional API sugar ("f"/"d"); Shift+L add uses "".
; Fast/Deep roles are normally set via Shift+L f/d → AiCompanionModels_SetRole.
AiCompanionModels_AddModel(companion, modelName, roleTag := "") {
    cfg := AiCompanionModels_Load(companion)
    modelName := Trim(modelName)
    if (modelName = "")
        return false
    already := false
    for name in cfg.models {
        if (name = modelName) {
            already := true
            break
        }
    }
    if (!already)
        cfg.models.Push(modelName)
    fast := cfg.fast
    deep := cfg.deep
    tag := StrLower(Trim(roleTag))
    if (tag = "f" || tag = "fast")
        fast := modelName
    else if (tag = "d" || tag = "deep")
        deep := modelName
    return AiCompanionModels_Save(companion, fast, deep, cfg.models)
}

AiCompanionModels_RemoveModel(companion, index) {
    cfg := AiCompanionModels_Load(companion)
    if (index < 1 || index > cfg.models.Length)
        return false
    removed := cfg.models[index]
    cfg.models.RemoveAt(index)
    fast := cfg.fast
    deep := cfg.deep
    if (fast = removed)
        fast := ""
    if (deep = removed)
        deep := ""
    return AiCompanionModels_Save(companion, fast, deep, cfg.models)
}

; Rename Models[index]; keep Fast/Deep in sync when they matched the old name.
AiCompanionModels_RenameModel(companion, index, newName) {
    cfg := AiCompanionModels_Load(companion)
    if (index < 1 || index > cfg.models.Length)
        return false
    newName := Trim(newName)
    if (newName = "")
        return false
    oldName := cfg.models[index]
    for i, name in cfg.models {
        if (i != index && name = newName)
            return false
    }
    cfg.models[index] := newName
    fast := cfg.fast
    deep := cfg.deep
    if (fast = oldName)
        fast := newName
    if (deep = oldName)
        deep := newName
    return AiCompanionModels_Save(companion, fast, deep, cfg.models)
}

; Max selectable numbered/lettered slots (1-9 + letters skipping a/e/f/d).
; a = add, e = edit (Utility Shortcuts style); f/d = Fast/Deep roles.
AiCompanionModels_ItemLetters() {
    return "bcghijklmnopqrstuvwxyz"
}

AiCompanionModels_IndexFromKey(key) {
    key := StrLower(Trim(key))
    if (RegExMatch(key, "^[1-9]$"))
        return Integer(key)
    letters := AiCompanionModels_ItemLetters()
    pos := InStr(letters, key, true)
    if (pos)
        return 9 + pos
    return 0
}

AiCompanionModels_LabelForIndex(index) {
    if (index >= 1 && index <= 9)
        return String(index)
    letters := AiCompanionModels_ItemLetters()
    offset := index - 9
    if (offset >= 1 && offset <= StrLen(letters))
        return SubStr(letters, offset, 1)
    return String(index)
}

AiCompanionModels_MaxSlots() {
    return 9 + StrLen(AiCompanionModels_ItemLetters())
}

AiCompanionModels_DisplayName(companion) {
    if (companion = AI_COMPANION_GEMINI)
        return "Gemini"
    if (companion = AI_COMPANION_ENTERPRISE)
        return "Gemini Enterprise"
    if (companion = AI_COMPANION_COPILOT)
        return "Copilot Web"
    return companion
}

AiCompanionModels_IsGeminiThinkingToggleName(modelName) {
    modelName := Trim(modelName)
    if (modelName = "")
        return false
    return (modelName = "Thinking level" || modelName = "Extended thinking"
        || RegExMatch(modelName, "i)extended\s*thinking|thinking\s*level"))
}

; Apply a model by exact UIA-visible name for the active companion window.
AiCompanionModels_Apply(companion, modelName) {
    modelName := Trim(modelName)
    if (modelName = "")
        return false
    if (companion = AI_COMPANION_GEMINI)
        return AiCompanionModels_ApplyGemini(modelName)
    if (companion = AI_COMPANION_ENTERPRISE)
        return AiCompanionModels_ApplyEnterprise(modelName)
    if (companion = AI_COMPANION_COPILOT)
        return AiCompanionModels_ApplyCopilot(modelName)
    return false
}

AiCompanionModels_ApplyGemini(modelName) {
    try {
        SetTitleMatchMode(2)
        geminiHwnd := 0
        try geminiHwnd := FindGeminiChromeHwnd()
        catch {
            geminiHwnd := 0
        }
        if (geminiHwnd) {
            if (!WinExist("ahk_id " geminiHwnd))
                return false
            WinActivate("ahk_id " geminiHwnd)
            if !WinWaitActive("ahk_id " geminiHwnd, , 1)
                return false
        } else if (WinExist("ahk_exe chrome.exe")) {
            WinActivate("ahk_exe chrome.exe")
            if !WinWaitActive("ahk_exe chrome.exe", , 2)
                return false
            geminiHwnd := WinExist("A")
        } else {
            return false
        }

        StandardLoadingBar_Show("🔄 Switching Gemini model…", BANNER_ACCENT_INTERMEDIATE, { centerOnHwnd: geminiHwnd })
        try {
            if (AiCompanionModels_IsGeminiThinkingToggleName(modelName)) {
                StandardLoadingBar_Update("🔄 Toggling Extended thinking…", BANNER_ACCENT_INTERMEDIATE)
                opened := EnsureGeminiExtendedThinkingToggle(geminiHwnd)
                if (opened) {
                    StandardLoadingBar_Update("✅ Extended thinking toggled", BANNER_ACCENT_INTERMEDIATE)
                    StandardLoadingBar_Hide(700)
                    return true
                }
                StandardLoadingBar_Hide(0)
                return false
            }
            verified := EnsureGeminiModelViaMenu(modelName, geminiHwnd)
            if (verified) {
                StandardLoadingBar_Update("✅ " . modelName . " model verified", BANNER_ACCENT_INTERMEDIATE)
                try FocusGeminiAskFieldForHwnd(geminiHwnd, false)
                StandardLoadingBar_Hide(700)
                return true
            }
            StandardLoadingBar_Hide(0)
            return false
        } finally {
            try StandardLoadingBar_Hide(0)
        }
    } catch {
        try StandardLoadingBar_Hide(0)
        return false
    }
}

AiCompanionModels_ApplyEnterprise(modelName) {
    ok := GeminiEnterprise_RunWithBusyBanner("⏳ Selecting " . modelName . "… Don't move the mouse", (*) =>
        GeminiEnterprise_SelectModelByName(modelName))
    if (ok)
        GeminiEnterprise_ReturnToComposer()
    return !!ok
}

AiCompanionModels_ApplyCopilot(modelName) {
    ok := CopilotWeb_RunWithBusyBanner("⏳ Selecting " . modelName . "… Don't move the mouse", (*) =>
        CopilotWeb_SelectModelByName(modelName))
    CopilotWeb_ReturnToComposer()
    return !!ok
}

; One-shot Fast / Deep. Returns true if already active or selection succeeded.
AiCompanionModels_SelectRole(companion, role) {
    role := StrLower(Trim(role))
    cfg := AiCompanionModels_Load(companion)
    modelName := (role = "fast") ? cfg.fast : ((role = "deep") ? cfg.deep : "")
    if (modelName = "") {
        try ShowCenteredOverlay_Utils((role = "fast" ? "Fast" : "Deep") .
        " model not configured — use Shift+L to set it", 2800, BANNER_ACCENT_INTERMEDIATE)
        return false
    }
    if (companion = AI_COMPANION_ENTERPRISE) {
        if (GeminiEnterprise_IsModelSelected(modelName)) {
            GeminiEnterprise_ReturnToComposer()
            return true
        }
    } else if (companion = AI_COMPANION_COPILOT) {
        root := CopilotWeb_GetActiveUia()
        btn := CopilotWeb_FindModelSelectorButton(root)
        if (btn && CopilotWeb_ModelLabelMatches(CopilotWeb_GetModelSelectorLabel(btn), modelName, role)) {
            CopilotWeb_ReturnToComposer()
            return true
        }
    } else if (companion = AI_COMPANION_GEMINI) {
        ; EnsureGeminiModelViaMenu no-ops when already selected.
    }
    ok := AiCompanionModels_Apply(companion, modelName)
    if (!ok) {
        try ShowCenteredOverlay_Utils("Could not select " . modelName, 2200, BANNER_ACCENT_ERROR)
    }
    return ok
}
