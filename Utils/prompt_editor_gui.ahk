; =============================================================================
; Utils module: prompt_editor_gui.ahk
; Single-form Add/Edit prompt dialog (identity + personal/work context file lists)
; =============================================================================

global g_PromptEditorGui := false
global g_PromptEditorResult := { saved: false }
global g_PromptEditorName := false
global g_PromptEditorCategory := false
global g_PromptEditorChar := false
global g_PromptEditorFile := false
global g_PromptEditorFilePath := ""
global g_PromptEditorSource := "file"
global g_PromptEditorAuthor := ""
global g_PromptEditorPersonalLv := false
global g_PromptEditorWorkLv := false
global g_PromptEditorIsEdit := false
global g_PromptEditorListIndex := 0
global g_PromptEditorPersonalPaths := []
global g_PromptEditorWorkPaths := []
global g_PromptEditorFlagCtrls := Map()
global g_PromptEditorSuppressFlagSync := false
global g_PromptEditorHelpGui := false
global g_PromptEditorTags := false
global g_PromptEditorVariables := false
global g_PromptEditorPasteMode := false
global g_PromptEditorAttachAsTxt := false
global g_PromptEditorExpectsDataOutput := false
global g_PromptEditorDataOutputFormat := false
global g_PromptEditorDataOutputHint := false
global g_PromptEditorDraftFile := false
global g_PromptEditorDraftPath := ""
global g_PromptEditorGitCommit := false
global g_PromptEditorPersonalPreset := false
global g_PromptEditorWorkPreset := false
global g_PromptEditorPersonalSelectablePaths := []
global g_PromptEditorWorkSelectablePaths := []
global g_PromptEditorPersonalSelectableLv := false
global g_PromptEditorWorkSelectableLv := false
global g_PromptEditorTabs := false
global g_PromptEditorLayout := false
global g_PromptEditorContextCtrls := []
global g_PromptEditorContextOrig := []
global g_PromptEditorContextScroll := 0
global g_PromptEditorContextScrollMax := 0
global g_PromptEditorContextScrollPage := 1
global g_PromptEditorContextSb := false
global g_PromptEditorContextScrollBound := false

PromptEditor_Show(existingPrompt := false, listIndex := 0) {
    global g_PromptEntries, g_PromptEditorGui, g_PromptEditorResult
    global g_PromptEditorIsEdit, g_PromptEditorListIndex, g_PromptEditorAuthor
    global g_PromptEditorFilePath, g_PromptEditorSource
    global g_PromptEditorPersonalPaths, g_PromptEditorWorkPaths
    global g_PromptEditorPersonalSelectablePaths, g_PromptEditorWorkSelectablePaths
    global g_PromptEditorDraftPath

    PromptData_Load()
    g_PromptEditorIsEdit := IsObject(existingPrompt)
    g_PromptEditorListIndex := listIndex
    g_PromptEditorResult := { saved: false }
    g_PromptEditorAuthor := (g_PromptEditorIsEdit && existingPrompt.HasProp("author")) ? existingPrompt.author : ""
    g_PromptEditorFilePath := (g_PromptEditorIsEdit && existingPrompt.HasProp("filePath")) ? existingPrompt.filePath :
        ""
    g_PromptEditorSource := (g_PromptEditorIsEdit && existingPrompt.HasProp("source")) ? existingPrompt.source : "file"
    g_PromptEditorPersonalPaths := PromptData_ParseContextEntries((g_PromptEditorIsEdit && existingPrompt.HasProp(
        "personal_context_files")) ? existingPrompt.personal_context_files : [])
    g_PromptEditorWorkPaths := PromptData_ParseContextEntries((g_PromptEditorIsEdit && existingPrompt.HasProp(
        "work_context_files")) ? existingPrompt.work_context_files : [])
    g_PromptEditorPersonalSelectablePaths := PromptData_ParseContextEntries((g_PromptEditorIsEdit && existingPrompt.HasProp(
        "personal_selectable_context_files")) ? existingPrompt.personal_selectable_context_files : [])
    g_PromptEditorWorkSelectablePaths := PromptData_ParseContextEntries((g_PromptEditorIsEdit && existingPrompt.HasProp(
        "work_selectable_context_files")) ? existingPrompt.work_selectable_context_files : [])
    g_PromptEditorDraftPath := (g_PromptEditorIsEdit && existingPrompt.HasProp("filePathDraft")) ? existingPrompt.filePathDraft :
        ""

    currentChar := (g_PromptEditorIsEdit && existingPrompt.HasProp("char")) ? existingPrompt.char : ""
    avail := UtilitySelector_AvailableChars(PromptData_CharSequence, PromptData_IsValidChar, g_PromptEntries, "char",
        listIndex)
    if (currentChar != "" && PromptData_IsValidChar(currentChar)) {
        already := false
        for c in avail {
            if (c = currentChar) {
                already := true
                break
            }
        }
        if (!already)
            avail.InsertAt(1, currentChar)
    }
    if (avail.Length = 0) {
        UtilitySelector_Notify("No free characters left. Delete an item first.")
        return g_PromptEditorResult
    }

    ownerOpt := "+AlwaysOnTop +ToolWindow"
    hwnd := UtilitySelector_SelectorHwnd()
    if (hwnd)
        ownerOpt .= " +Owner" . hwnd

    title := g_PromptEditorIsEdit ? "Edit prompt" : "Add prompt"
    g_PromptEditorGui := Gui(ownerOpt, title)
    g_PromptEditorGui.SetFont("s10", "Segoe UI")
    g_PromptEditorGui.MarginX := 22
    g_PromptEditorGui.MarginY := 14
    PromptEditor_BuildControls(existingPrompt, avail, currentChar)
    g_PromptEditorGui.OnEvent("Close", PromptEditor_OnCancel)
    g_PromptEditorGui.OnEvent("Escape", PromptEditor_OnCancel)

    UtilitySelector_DialogsBegin()
    PromptEditor_CenterOnSelector()
    PromptEditor_ContextScrollInit()
    try g_PromptEditorName.Focus()
    catch {
    }
    PromptEditor_BindEditorHotkeys(true)
    try WinWaitClose("ahk_id " g_PromptEditorGui.Hwnd)
    catch {
    }
    PromptEditor_ContextScrollTeardown()
    PromptEditor_BindEditorHotkeys(false)
    UtilitySelector_DialogsEnd()
    g_PromptEditorGui := false
    return g_PromptEditorResult
}

PromptEditor_Layout() {
    padX := 22
    padY := 14
    colGap := 16
    tabW := 760
    tabH := 720
    ; Insets inside the tab page (applied via Gui.Margin after UseTab — not absolute x/y).
    innerX := 18
    innerY := 14
    scrollW := 18
    ; Tab chrome eats a few px; reserve scrollbar gutter on the right so Apply/lists don't touch the edge.
    tabChrome := 12
    labelW := 88
    footerGap := 14
    footerY := padY + tabH + footerGap
    innerW := tabW - (2 * innerX) - scrollW - tabChrome
    colW := Floor((innerW - colGap) / 2)
    return Map(
        "padX", padX,
        "padY", padY,
        "innerX", innerX,
        "innerY", innerY,
        "scrollW", scrollW,
        "labelW", labelW,
        "colGap", colGap,
        "tabW", tabW,
        "tabH", tabH,
        "footerY", footerY,
        "innerW", innerW,
        "colW", colW,
        "workX", colW + colGap,
        "pathCol", colW - 96,
        "compactCol", 44,
        "csvCol", 44,
        "selPathCol", colW - 8,
        "hintW", innerW,
        "saveX", padX + tabW - 208
    )
}

; Tab3: never use xm inside tabs (window-relative → clips). First control: no x/y so
; MarginX/MarginY place it in the tab page; later rows use xs / xs y+N Section only.
PromptEditor_BuildControls(existingPrompt, avail, currentChar) {
    global g_PromptEditorGui, g_PromptEditorTabs, g_PromptEditorLayout
    L := PromptEditor_Layout()
    g_PromptEditorLayout := L
    g_PromptEditorTabs := g_PromptEditorGui.Add("Tab3",
        "x" . L["padX"] . " y" . L["padY"] . " w" . L["tabW"] . " h" . L["tabH"] . " Choose1",
        ["General", "Context", "Advanced"])
    g_PromptEditorTabs.OnEvent("Change", PromptEditor_OnTabChange)
    PromptEditor_BuildTabGeneral(existingPrompt, avail, currentChar, L)
    PromptEditor_BuildTabContext(existingPrompt, L)
    PromptEditor_BuildTabAdvanced(existingPrompt, L)
    g_PromptEditorTabs.UseTab()
    g_PromptEditorGui.MarginX := L["padX"]
    g_PromptEditorGui.MarginY := L["padY"]
    PromptEditor_BuildFooter(L)
}

PromptEditor_CtxTrack(ctrl) {
    global g_PromptEditorContextCtrls
    if (IsObject(ctrl))
        g_PromptEditorContextCtrls.Push(ctrl)
    return ctrl
}

PromptEditor_BuildTabGeneral(existingPrompt, avail, currentChar, L) {
    global g_PromptEditorGui, g_PromptEditorTabs
    global g_PromptEditorName, g_PromptEditorCategory, g_PromptEditorChar
    global g_PromptEditorFile, g_PromptEditorFilePath
    global g_PromptEditorTags, g_PromptEditorVariables, g_PromptEditorPasteMode
    global g_PromptEditorDraftFile, g_PromptEditorDraftPath

    innerW := L["innerW"]
    innerX := L["innerX"]
    innerY := L["innerY"]
    labelW := L["labelW"]
    fileEditW := Max(160, innerW - labelW - 184)
    tagsEditW := Floor((innerW - labelW - 16 - 70) * 0.45)
    varsEditW := innerW - labelW - 16 - 70 - tagsEditW
    draftEditW := Max(160, innerW - labelW - 166)

    g_PromptEditorTabs.UseTab(1)
    g_PromptEditorGui.MarginX := innerX
    g_PromptEditorGui.MarginY := innerY
    g_PromptEditorGui.Add("Text", "w" . labelW . " Section", "Name")
    nameVal := (IsObject(existingPrompt) && existingPrompt.HasProp("name")) ? existingPrompt.name : ""
    g_PromptEditorName := g_PromptEditorGui.Add("Edit", "yp w" . (innerW - labelW), nameVal)

    g_PromptEditorGui.Add("Text", "xs y+12 w" . labelW . " Section", "Category")
    catChoices := PromptEditor_CategoryChoices(IsObject(existingPrompt) ? existingPrompt.category : "General")
    g_PromptEditorCategory := g_PromptEditorGui.Add("ComboBox", "yp w220", catChoices)
    if (IsObject(existingPrompt) && existingPrompt.HasProp("category") && existingPrompt.category != "")
        g_PromptEditorCategory.Text := existingPrompt.category
    else
        g_PromptEditorCategory.Text := "General"

    g_PromptEditorGui.Add("Text", "x+16 yp w40", "Char")
    g_PromptEditorChar := g_PromptEditorGui.Add("DropDownList", "yp w80", avail)
    chooseChar := currentChar != "" ? currentChar : avail[1]
    try g_PromptEditorChar.Text := chooseChar
    catch {
        try g_PromptEditorChar.Choose(1)
        catch {
        }
    }

    g_PromptEditorGui.Add("Text", "xs y+12 w" . labelW . " Section", "Prompt file")
    g_PromptEditorFile := g_PromptEditorGui.Add("Edit", "yp w" . fileEditW . " ReadOnly", g_PromptEditorFilePath)
    g_PromptEditorGui.Add("Button", "yp w80", "Browse").OnEvent("Click", PromptEditor_OnBrowsePromptFile)
    g_PromptEditorGui.Add("Button", "x+8 yp w80", "History").OnEvent("Click", PromptEditor_OnShowHistory)

    g_PromptEditorGui.Add("Text", "xs y+12 w" . labelW . " Section", "Tags")
    tagsVal := (IsObject(existingPrompt) && existingPrompt.HasProp("tags")) ? existingPrompt.tags : ""
    g_PromptEditorTags := g_PromptEditorGui.Add("Edit", "yp w" . tagsEditW, tagsVal)
    g_PromptEditorGui.Add("Text", "x+16 yp w70", "Variables")
    varsVal := (IsObject(existingPrompt) && existingPrompt.HasProp("variables")) ? existingPrompt.variables : ""
    g_PromptEditorVariables := g_PromptEditorGui.Add("Edit", "yp w" . varsEditW, varsVal)

    g_PromptEditorGui.Add("Text", "xs y+12 w" . labelW . " Section", "Paste mode")
    pasteModes := ["default", "body_only", "body_plus_clipboard", "body_attach_clipboard", "attach_only", "auto_send"]
    g_PromptEditorPasteMode := g_PromptEditorGui.Add("DropDownList", "yp w220", pasteModes)
    pasteVal := (IsObject(existingPrompt) && existingPrompt.HasProp("pasteMode")) ? PromptData_NormalizePasteMode(
        existingPrompt.pasteMode) : "default"
    try g_PromptEditorPasteMode.Text := pasteVal
    catch {
        try g_PromptEditorPasteMode.Choose(1)
        catch {
        }
    }

    g_PromptEditorGui.Add("Text", "xs y+12 w" . labelW . " Section", "Draft file")
    g_PromptEditorDraftFile := g_PromptEditorGui.Add("Edit", "yp w" . draftEditW . " ReadOnly", g_PromptEditorDraftPath
    )
    g_PromptEditorGui.Add("Button", "yp w70", "Draft…").OnEvent("Click", PromptEditor_OnBrowseDraftFile)
    g_PromptEditorGui.Add("Button", "x+8 yp w80", "Promote").OnEvent("Click", PromptEditor_OnPromoteDraft)
}

PromptEditor_BuildTabContext(existingPrompt, L) {
    global g_PromptEditorGui, g_PromptEditorTabs
    global g_PromptEditorPersonalLv, g_PromptEditorWorkLv, g_PromptEditorFlagCtrls
    global g_PromptEditorPersonalPreset, g_PromptEditorWorkPreset
    global g_PromptEditorPersonalSelectableLv, g_PromptEditorWorkSelectableLv
    global g_PromptEditorContextCtrls

    g_PromptEditorContextCtrls := []
    track := PromptEditor_CtxTrack
    innerX := L["innerX"]
    innerY := L["innerY"]
    colW := L["colW"]
    colGap := L["colGap"]
    workX := L["workX"]
    pathCol := L["pathCol"]
    compactCol := L["compactCol"]
    csvCol := L["csvCol"]
    selPathCol := L["selPathCol"]
    hintW := L["hintW"]

    g_PromptEditorTabs.UseTab(2)
    g_PromptEditorGui.MarginX := innerX
    g_PromptEditorGui.MarginY := innerY
    g_PromptEditorGui.SetFont("s9 c808080 Norm", "Segoe UI")
    track(g_PromptEditorGui.Add("Text", "w" . hintW . " Wrap Section",
        "Static lists attach every time. Selectable files appear in a picker when you invoke this prompt (not auto-attached)."
    ))
    g_PromptEditorGui.SetFont("s10 Norm", "Segoe UI")

    track(g_PromptEditorGui.Add("Text", "xs y+12 w" . colW . " Section", "Personal static context"))
    track(g_PromptEditorGui.Add("Text", "x+" . colGap . " yp w" . colW, "Work static context"))
    g_PromptEditorPersonalLv := track(g_PromptEditorGui.Add("ListView", "xs w" . colW . " r5", ["Path", "Cmp", "CSV"]))
    g_PromptEditorWorkLv := track(g_PromptEditorGui.Add("ListView", "x+" . colGap . " yp w" . colW . " r5", ["Path",
        "Cmp",
        "CSV"]))
    for lv in [g_PromptEditorPersonalLv, g_PromptEditorWorkLv] {
        lv.ModifyCol(1, pathCol)
        lv.ModifyCol(2, compactCol)
        lv.ModifyCol(3, csvCol)
    }
    g_PromptEditorPersonalLv.OnEvent("ItemFocus", (ctrl, info) => PromptEditor_OnItemFocus("personal", info))
    g_PromptEditorWorkLv.OnEvent("ItemFocus", (ctrl, info) => PromptEditor_OnItemFocus("work", info))
    g_PromptEditorPersonalLv.OnEvent("Click", (ctrl, info) => PromptEditor_OnItemFocus("personal", info))
    g_PromptEditorWorkLv.OnEvent("Click", (ctrl, info) => PromptEditor_OnItemFocus("work", info))
    PromptEditor_ReloadList("personal")
    PromptEditor_ReloadList("work")

    presetLabels := PromptContextPresets_ChoiceLabels()
    track(g_PromptEditorGui.Add("Button", "xs y+10 w70 Section", "Add")).OnEvent("Click", (*) =>
        PromptEditor_OnAddFiles(
            "personal"
        ))
    track(g_PromptEditorGui.Add("Button", "x+6 yp w70", "Paste")).OnEvent("Click", (*) => PromptEditor_OnPastePaths(
        "personal"
    ))
    track(g_PromptEditorGui.Add("Button", "x+6 yp w70", "Remove")).OnEvent("Click", (*) => PromptEditor_OnRemove(
        "personal"))
    track(g_PromptEditorGui.Add("Button", "xs+" . workX . " ys w70", "Add")).OnEvent("Click", (*) =>
        PromptEditor_OnAddFiles("work"))
    track(g_PromptEditorGui.Add("Button", "x+6 yp w70", "Paste")).OnEvent("Click", (*) => PromptEditor_OnPastePaths(
        "work"))
    track(g_PromptEditorGui.Add("Button", "x+6 yp w70", "Remove")).OnEvent("Click", (*) => PromptEditor_OnRemove("work"
    ))

    track(g_PromptEditorGui.Add("Text", "xs y+10 w44 Section", "Preset"))
    g_PromptEditorPersonalPreset := track(g_PromptEditorGui.Add("DropDownList", "yp w210", presetLabels))
    try g_PromptEditorPersonalPreset.Text := "(none)"
    catch {
    }
    track(g_PromptEditorGui.Add("Button", "x+6 yp w70", "Apply")).OnEvent("Click", (*) => PromptEditor_OnApplyPreset(
        "personal"))
    track(g_PromptEditorGui.Add("Text", "xs+" . workX . " ys w44", "Preset"))
    g_PromptEditorWorkPreset := track(g_PromptEditorGui.Add("DropDownList", "yp w210", presetLabels))
    try g_PromptEditorWorkPreset.Text := "(none)"
    catch {
    }
    track(g_PromptEditorGui.Add("Button", "x+6 yp w70", "Apply")).OnEvent("Click", (*) => PromptEditor_OnApplyPreset(
        "work"))

    personalCompact := track(g_PromptEditorGui.Add("CheckBox", "xs w" . colW, "Compact"))
    workCompact := track(g_PromptEditorGui.Add("CheckBox", "x+" . colGap . " yp w" . colW, "Compact"))
    personalCompact.OnEvent("Click", (*) => PromptEditor_OnCompactClick("personal"))
    workCompact.OnEvent("Click", (*) => PromptEditor_OnCompactClick("work"))

    personalCsvFromLabel := track(g_PromptEditorGui.Add("Text", "xs w90 Hidden", "CSV keep from"))
    personalCsvFrom := track(g_PromptEditorGui.Add("Edit", "yp w50 Number Hidden"))
    personalCsvToLabel := track(g_PromptEditorGui.Add("Text", "yp w20 Hidden", "to"))
    personalCsvTo := track(g_PromptEditorGui.Add("Edit", "yp w50 Number Hidden"))
    workCsvFromLabel := track(g_PromptEditorGui.Add("Text", "x+162 yp w90 Hidden", "CSV keep from"))
    workCsvFrom := track(g_PromptEditorGui.Add("Edit", "yp w50 Number Hidden"))
    workCsvToLabel := track(g_PromptEditorGui.Add("Text", "yp w20 Hidden", "to"))
    workCsvTo := track(g_PromptEditorGui.Add("Edit", "yp w50 Number Hidden"))
    personalCsvFrom.OnEvent("Change", (*) => PromptEditor_OnCsvKeepChange("personal"))
    personalCsvTo.OnEvent("Change", (*) => PromptEditor_OnCsvKeepChange("personal"))
    workCsvFrom.OnEvent("Change", (*) => PromptEditor_OnCsvKeepChange("work"))
    workCsvTo.OnEvent("Change", (*) => PromptEditor_OnCsvKeepChange("work"))

    g_PromptEditorFlagCtrls := Map()
    g_PromptEditorFlagCtrls["personal"] := {
        compact: personalCompact,
        csvFromLabel: personalCsvFromLabel,
        csvFrom: personalCsvFrom,
        csvToLabel: personalCsvToLabel,
        csvTo: personalCsvTo,
        selected: 0
    }
    g_PromptEditorFlagCtrls["work"] := {
        compact: workCompact,
        csvFromLabel: workCsvFromLabel,
        csvFrom: workCsvFrom,
        csvToLabel: workCsvToLabel,
        csvTo: workCsvTo,
        selected: 0
    }
    PromptEditor_LoadFlagControls("personal", 0)
    PromptEditor_LoadFlagControls("work", 0)

    track(g_PromptEditorGui.Add("Button", "xs y+14 w120 Section", "Manage presets")).OnEvent("Click",
        PromptEditor_OnOpenPresetManager)
    track(g_PromptEditorGui.Add("Button", "x+8 yp w120", "Save as preset")).OnEvent("Click",
        PromptEditor_OnSaveAsPreset)

    g_PromptEditorGui.SetFont("s9 c808080 Norm", "Segoe UI")
    track(g_PromptEditorGui.Add("Text", "xs y+16 w" . hintW . " Wrap Section",
        "Selectable at paste — add paths below; choose 0–n when you invoke this prompt (not auto-attached)."))
    g_PromptEditorGui.SetFont("s10 Norm", "Segoe UI")

    track(g_PromptEditorGui.Add("Text", "xs y+14 w" . colW . " Section", "Personal selectable"))
    track(g_PromptEditorGui.Add("Text", "x+" . colGap . " yp w" . colW, "Work selectable"))
    g_PromptEditorPersonalSelectableLv := track(g_PromptEditorGui.Add("ListView", "xs w" . colW . " r5", ["Path"]))
    g_PromptEditorWorkSelectableLv := track(g_PromptEditorGui.Add("ListView", "x+" . colGap . " yp w" . colW . " r5", [
        "Path"]))
    g_PromptEditorPersonalSelectableLv.ModifyCol(1, selPathCol)
    g_PromptEditorWorkSelectableLv.ModifyCol(1, selPathCol)
    PromptEditor_ReloadSelectableList("personal")
    PromptEditor_ReloadSelectableList("work")

    track(g_PromptEditorGui.Add("Button", "xs y+8 w70 Section", "Add")).OnEvent("Click", (*) =>
        PromptEditor_OnAddSelectableFiles("personal"))
    track(g_PromptEditorGui.Add("Button", "x+6 yp w70", "Paste")).OnEvent("Click", (*) =>
        PromptEditor_OnPasteSelectablePaths("personal"))
    track(g_PromptEditorGui.Add("Button", "x+6 yp w70", "Remove")).OnEvent("Click", (*) =>
        PromptEditor_OnRemoveSelectable("personal"))
    track(g_PromptEditorGui.Add("Button", "xs+" . workX . " ys w70", "Add")).OnEvent("Click", (*) =>
        PromptEditor_OnAddSelectableFiles("work"))
    track(g_PromptEditorGui.Add("Button", "x+6 yp w70", "Paste")).OnEvent("Click", (*) =>
        PromptEditor_OnPasteSelectablePaths("work"))
    track(g_PromptEditorGui.Add("Button", "x+6 yp w70", "Remove")).OnEvent("Click", (*) =>
        PromptEditor_OnRemoveSelectable("work"))
}

PromptEditor_BuildTabAdvanced(existingPrompt, L) {
    global g_PromptEditorGui, g_PromptEditorTabs
    global g_PromptEditorGitCommit, g_PromptEditorAttachAsTxt
    global g_PromptEditorExpectsDataOutput, g_PromptEditorDataOutputFormat, g_PromptEditorDataOutputHint

    innerX := L["innerX"]
    innerY := L["innerY"]
    hintW := L["hintW"]

    g_PromptEditorTabs.UseTab(3)
    g_PromptEditorGui.MarginX := innerX
    g_PromptEditorGui.MarginY := innerY
    g_PromptEditorGui.Add("Text", "w" . hintW . " Section", "Output and save options")

    g_PromptEditorGitCommit := g_PromptEditorGui.Add("CheckBox", "xs y+14 w" . hintW . " Section", "Git commit on save"
    )
    try g_PromptEditorGitCommit.Value := 0
    catch {
    }
    g_PromptEditorAttachAsTxt := g_PromptEditorGui.Add("CheckBox", "xs y+12 w" . hintW . " Section",
        "Attach context as .txt")
    attachAsTxtVal := 0
    if (IsObject(existingPrompt) && existingPrompt.HasProp("attachAsTxt"))
        attachAsTxtVal := PromptData_NormalizeAttachAsTxt(existingPrompt.attachAsTxt)
    try g_PromptEditorAttachAsTxt.Value := attachAsTxtVal
    catch {
    }

    g_PromptEditorExpectsDataOutput := g_PromptEditorGui.Add("CheckBox", "xs y+12 w" . hintW . " Section",
        "Expects data output")
    expectsVal := 0
    if (IsObject(existingPrompt) && existingPrompt.HasProp("expectsDataOutput"))
        expectsVal := PromptData_NormalizeExpectsDataOutput(existingPrompt.expectsDataOutput)
    try g_PromptEditorExpectsDataOutput.Value := expectsVal
    catch {
    }
    g_PromptEditorExpectsDataOutput.OnEvent("Click", PromptEditor_OnExpectsDataOutputClick)

    g_PromptEditorGui.Add("Text", "xs y+10 w90 Section", "Data output")
    g_PromptEditorDataOutputFormat := g_PromptEditorGui.Add("DropDownList", "yp w100", ["file", "code"])
    fmtVal := "file"
    if (IsObject(existingPrompt) && existingPrompt.HasProp("dataOutputFormat"))
        fmtVal := PromptData_NormalizeDataOutputFormat(existingPrompt.dataOutputFormat)
    try g_PromptEditorDataOutputFormat.Text := fmtVal
    catch {
        try g_PromptEditorDataOutputFormat.Choose(1)
        catch {
        }
    }
    try g_PromptEditorDataOutputFormat.Enabled := (expectsVal = 1)
    catch {
    }
    g_PromptEditorDataOutputHint := g_PromptEditorGui.Add("Text", "xs y+10 w" . hintW . " Wrap c808080",
        "Convention: .txt — app transpiles after save")
}

PromptEditor_BuildFooter(L) {
    global g_PromptEditorGui
    fy := L["footerY"]
    px := L["padX"]
    g_PromptEditorGui.Add("Button", "x" . px . " y" . fy . " w80 Section", "Help").OnEvent("Click",
        PromptEditor_ShowHelp)
    g_PromptEditorGui.Add("Button", "x" . L["saveX"] . " yp w100 Default", "Save").OnEvent("Click", PromptEditor_OnSave
    )
    g_PromptEditorGui.Add("Button", "x+8 yp w100", "Cancel").OnEvent("Click", PromptEditor_OnCancel)
}

PromptEditor_OnTabChange(*) {
    PromptEditor_ContextScrollSyncVisibility()
}

PromptEditor_ContextScrollInit() {
    global g_PromptEditorGui, g_PromptEditorTabs, g_PromptEditorLayout
    global g_PromptEditorContextCtrls, g_PromptEditorContextOrig
    global g_PromptEditorContextScroll, g_PromptEditorContextScrollMax, g_PromptEditorContextScrollPage
    global g_PromptEditorContextSb, g_PromptEditorContextScrollBound

    g_PromptEditorContextOrig := []
    g_PromptEditorContextScroll := 0
    g_PromptEditorContextScrollMax := 0
    if (!IsObject(g_PromptEditorGui) || !IsObject(g_PromptEditorTabs) || !IsObject(g_PromptEditorLayout))
        return
    if (g_PromptEditorContextCtrls.Length = 0)
        return

    L := g_PromptEditorLayout
    contentBottom := 0
    for ctrl in g_PromptEditorContextCtrls {
        try {
            ctrl.GetPos(&cx, &cy, &cw, &ch)
            g_PromptEditorContextOrig.Push({ ctrl: ctrl, x: cx, y: cy, w: cw, h: ch })
            contentBottom := Max(contentBottom, cy + ch)
        } catch {
        }
    }

    g_PromptEditorTabs.GetPos(&tx, &ty, &tw, &th)
    tabHeader := 32
    viewBottom := ty + th - 6
    overflow := contentBottom - viewBottom
    g_PromptEditorContextScrollMax := Max(0, overflow + 12)
    g_PromptEditorContextScrollPage := Max(40, th - tabHeader - 40)

    try {
        if (IsObject(g_PromptEditorContextSb)) {
            g_PromptEditorContextSb.Visible := false
            g_PromptEditorContextSb := false
        }
    } catch {
        g_PromptEditorContextSb := false
    }

    if (g_PromptEditorContextScrollMax > 0) {
        sbX := tx + tw - L["scrollW"] - 4
        sbY := ty + tabHeader
        sbH := Max(80, th - tabHeader - 8)
        ; Native Slider (Custom msctls_scrollbar32 is unreliable as Gui.Add Custom).
        g_PromptEditorContextSb := g_PromptEditorGui.Add("Slider",
            "Vertical Invert Center NoTicks x" . sbX . " y" . sbY
            . " w" . L["scrollW"] . " h" . sbH
            . " Range0-" . g_PromptEditorContextScrollMax)
        try g_PromptEditorContextSb.Value := 0
        catch {
        }
        g_PromptEditorContextSb.OnEvent("Change", PromptEditor_ContextOnSliderChange)
    }
    PromptEditor_ContextScrollSyncVisibility()

    if (!g_PromptEditorContextScrollBound) {
        OnMessage(0x20A, PromptEditor_ContextOnMouseWheel) ; WM_MOUSEWHEEL
        g_PromptEditorContextScrollBound := true
    }
}

PromptEditor_ContextScrollTeardown() {
    global g_PromptEditorContextScrollBound, g_PromptEditorContextSb
    global g_PromptEditorContextCtrls, g_PromptEditorContextOrig
    global g_PromptEditorContextScroll, g_PromptEditorContextScrollMax

    if (g_PromptEditorContextScrollBound) {
        OnMessage(0x20A, PromptEditor_ContextOnMouseWheel, 0)
        g_PromptEditorContextScrollBound := false
    }
    g_PromptEditorContextSb := false
    g_PromptEditorContextCtrls := []
    g_PromptEditorContextOrig := []
    g_PromptEditorContextScroll := 0
    g_PromptEditorContextScrollMax := 0
}

PromptEditor_ContextScrollSyncVisibility() {
    global g_PromptEditorTabs, g_PromptEditorContextSb, g_PromptEditorContextScrollMax
    showSb := false
    try showSb := IsObject(g_PromptEditorTabs) && (g_PromptEditorTabs.Value = 2) && (g_PromptEditorContextScrollMax > 0
    )
    catch {
        showSb := false
    }
    try {
        if (IsObject(g_PromptEditorContextSb))
            g_PromptEditorContextSb.Visible := showSb
    } catch {
    }
}

PromptEditor_ContextOnSliderChange(ctrl, *) {
    try PromptEditor_ContextScrollTo(ctrl.Value)
    catch {
    }
}

PromptEditor_ContextScrollTo(pos) {
    global g_PromptEditorContextScroll, g_PromptEditorContextScrollMax, g_PromptEditorContextOrig
    global g_PromptEditorContextSb
    if (g_PromptEditorContextScrollMax <= 0)
        return
    pos := Max(0, Min(Integer(pos), g_PromptEditorContextScrollMax))
    if (pos = g_PromptEditorContextScroll)
        return
    g_PromptEditorContextScroll := pos
    for item in g_PromptEditorContextOrig {
        try item.ctrl.Move(item.x, item.y - pos)
        catch {
        }
    }
    try {
        if (IsObject(g_PromptEditorContextSb) && g_PromptEditorContextSb.Value != pos)
            g_PromptEditorContextSb.Value := pos
    } catch {
    }
}

PromptEditor_ContextScrollBy(delta) {
    global g_PromptEditorContextScroll
    PromptEditor_ContextScrollTo(g_PromptEditorContextScroll + delta)
}

PromptEditor_ContextOnMouseWheel(wParam, lParam, msg, hwnd) {
    global g_PromptEditorGui, g_PromptEditorTabs, g_PromptEditorContextScrollMax
    if (!IsObject(g_PromptEditorGui) || !IsObject(g_PromptEditorTabs))
        return
    try {
        if (g_PromptEditorTabs.Value != 2 || g_PromptEditorContextScrollMax <= 0)
            return
    } catch {
        return
    }
    mouseX := lParam & 0xFFFF
    mouseY := (lParam >> 16) & 0xFFFF
    if (mouseX >= 0x8000)
        mouseX -= 0x10000
    if (mouseY >= 0x8000)
        mouseY -= 0x10000
    g_PromptEditorTabs.GetPos(&tx, &ty, &tw, &th)
    WinGetClientPos(&gx, &gy, , , g_PromptEditorGui.Hwnd)
    tabScreenL := gx + tx
    tabScreenT := gy + ty
    tabScreenR := tabScreenL + tw
    tabScreenB := tabScreenT + th
    if (mouseX < tabScreenL || mouseX > tabScreenR || mouseY < tabScreenT || mouseY > tabScreenB)
        return
    delta := (wParam >> 16) & 0xFFFF
    if (delta >= 0x8000)
        delta -= 0x10000
    PromptEditor_ContextScrollBy(-delta // 120 * 32)
    return 0
}

PromptEditor_CategoryChoices(current) {
    global g_PromptEntries
    seen := Map()
    choices := ["General"]
    seen["general"] := true
    if (current != "" && !seen.Has(StrLower(current))) {
        choices.Push(current)
        seen[StrLower(current)] := true
    }
    for p in g_PromptEntries {
        cat := p.HasProp("category") ? Trim(p.category) : ""
        if (cat = "" || seen.Has(StrLower(cat)))
            continue
        seen[StrLower(cat)] := true
        choices.Push(cat)
    }
    return choices
}

PromptEditor_CenterOnSelector() {
    global g_PromptEditorGui
    mon := UtilitySelector_ActiveMonitorWorkArea()
    g_PromptEditorGui.Show("Hide")
    g_PromptEditorGui.GetPos(, , &gw, &gh)
    guiX := mon.left + (mon.width - gw) // 2
    guiY := mon.top + (mon.height - gh) // 2
    if (guiX < mon.left + 20)
        guiX := mon.left + 20
    if (guiY < mon.top + 20)
        guiY := mon.top + 20
    g_PromptEditorGui.Show("x" . guiX . " y" . guiY)
}

PromptEditor_BindEditorHotkeys(enable) {
    global g_PromptEditorGui
    hwnd := 0
    try {
        if (IsObject(g_PromptEditorGui))
            hwnd := g_PromptEditorGui.Hwnd
    } catch {
        hwnd := 0
    }
    if (!hwnd)
        return
    try HotIfWinActive("ahk_id " hwnd)
    catch {
        return
    }
    try Hotkey("Delete", PromptEditor_OnDeleteKey, enable ? "On" : "Off")
    catch {
    }
    try Hotkey("F1", PromptEditor_ShowHelp, enable ? "On" : "Off")
    catch {
    }
    try Hotkey("h", PromptEditor_OnShowHistory, enable ? "On" : "Off")
    catch {
    }
    try HotIf()
    catch {
    }
}

PromptEditor_FocusedSide() {
    global g_PromptEditorGui, g_PromptEditorPersonalLv, g_PromptEditorWorkLv
    global g_PromptEditorPersonalSelectableLv, g_PromptEditorWorkSelectableLv
    focused := 0
    try focused := ControlGetFocus("ahk_id " g_PromptEditorGui.Hwnd)
    catch {
        return "personal"
    }
    if (IsObject(g_PromptEditorWorkSelectableLv) && focused = g_PromptEditorWorkSelectableLv.Hwnd)
        return "work"
    if (IsObject(g_PromptEditorWorkLv) && focused = g_PromptEditorWorkLv.Hwnd)
        return "work"
    return "personal"
}

PromptEditor_FocusedListKind() {
    global g_PromptEditorGui, g_PromptEditorPersonalLv, g_PromptEditorWorkLv
    global g_PromptEditorPersonalSelectableLv, g_PromptEditorWorkSelectableLv
    focused := 0
    try focused := ControlGetFocus("ahk_id " g_PromptEditorGui.Hwnd)
    catch {
        return ""
    }
    if (IsObject(g_PromptEditorPersonalSelectableLv) && focused = g_PromptEditorPersonalSelectableLv.Hwnd)
        return "selectable"
    if (IsObject(g_PromptEditorWorkSelectableLv) && focused = g_PromptEditorWorkSelectableLv.Hwnd)
        return "selectable"
    if (IsObject(g_PromptEditorPersonalLv) && focused = g_PromptEditorPersonalLv.Hwnd)
        return "static"
    if (IsObject(g_PromptEditorWorkLv) && focused = g_PromptEditorWorkLv.Hwnd)
        return "static"
    return ""
}

PromptEditor_OnDeleteKey(*) {
    kind := PromptEditor_FocusedListKind()
    side := PromptEditor_FocusedSide()
    if (kind = "selectable")
        PromptEditor_OnRemoveSelectable(side)
    else if (kind = "static")
        PromptEditor_OnRemove(side)
}

PromptEditor_PathsRef(side) {
    global g_PromptEditorPersonalPaths, g_PromptEditorWorkPaths
    return (side = "work") ? g_PromptEditorWorkPaths : g_PromptEditorPersonalPaths
}

PromptEditor_Lv(side) {
    global g_PromptEditorPersonalLv, g_PromptEditorWorkLv
    return (side = "work") ? g_PromptEditorWorkLv : g_PromptEditorPersonalLv
}

PromptEditor_SetPaths(side, arr) {
    global g_PromptEditorPersonalPaths, g_PromptEditorWorkPaths
    if (side = "work")
        g_PromptEditorWorkPaths := arr
    else
        g_PromptEditorPersonalPaths := arr
}

PromptEditor_ReloadList(side) {
    lv := PromptEditor_Lv(side)
    if (!IsObject(lv))
        return
    selected := PromptEditor_SelectedRow(side)
    lv.Delete()
    for e in PromptEditor_PathsRef(side)
        lv.Add("", PromptData_ContextEntryPath(e), PromptData_ContextCompactLabel(e), PromptData_ContextCsvKeepLabel(e))
    if (selected >= 1 && selected <= PromptEditor_PathsRef(side).Length)
        lv.Modify(selected, "Select Focus Vis")
    else
        PromptEditor_LoadFlagControls(side, 0)
}

PromptEditor_RefreshRow(side, row) {
    lv := PromptEditor_Lv(side)
    arr := PromptEditor_PathsRef(side)
    if (!IsObject(lv) || row < 1 || row > arr.Length)
        return
    e := arr[row]
    lv.Modify(row, , PromptData_ContextEntryPath(e), PromptData_ContextCompactLabel(e), PromptData_ContextCsvKeepLabel(
        e))
}

PromptEditor_SelectedRow(side) {
    lv := PromptEditor_Lv(side)
    if (!IsObject(lv))
        return 0
    row := lv.GetNext(0, "F")
    if (!row)
        row := lv.GetNext(0)
    return row
}

PromptEditor_EventRow(info) {
    row := 0
    try {
        if (IsObject(info) && info.HasProp("EventInfo"))
            row := Integer(info.EventInfo)
    } catch {
        row := 0
    }
    return row
}

PromptEditor_OnItemFocus(side, info) {
    row := PromptEditor_EventRow(info)
    if (row < 1)
        row := PromptEditor_SelectedRow(side)
    PromptEditor_LoadFlagControls(side, row)
}

PromptEditor_FlagCtrls(side) {
    global g_PromptEditorFlagCtrls
    if (!IsObject(g_PromptEditorFlagCtrls) || !g_PromptEditorFlagCtrls.Has(side))
        return false
    return g_PromptEditorFlagCtrls[side]
}

PromptEditor_SetCsvVisible(side, visible) {
    ctrls := PromptEditor_FlagCtrls(side)
    if (!IsObject(ctrls))
        return
    opt := visible ? "-Hidden" : "+Hidden"
    try ctrls.csvFromLabel.Opt(opt)
    try ctrls.csvFrom.Opt(opt)
    try ctrls.csvToLabel.Opt(opt)
    try ctrls.csvTo.Opt(opt)
    catch {
    }
}

PromptEditor_LoadFlagControls(side, row) {
    global g_PromptEditorSuppressFlagSync, g_PromptEditorFlagCtrls
    ctrls := PromptEditor_FlagCtrls(side)
    if (!IsObject(ctrls))
        return
    g_PromptEditorSuppressFlagSync := true
    arr := PromptEditor_PathsRef(side)
    ctrls.selected := row
    if (row < 1 || row > arr.Length) {
        try ctrls.compact.Value := 0
        catch {
        }
        try ctrls.compact.Enabled := false
        catch {
        }
        try ctrls.csvFrom.Value := ""
        try ctrls.csvTo.Value := ""
        catch {
        }
        PromptEditor_SetCsvVisible(side, false)
        g_PromptEditorSuppressFlagSync := false
        return
    }
    e := arr[row]
    try ctrls.compact.Enabled := true
    catch {
    }
    try ctrls.compact.Value := (e.HasProp("compact") && e.compact) ? 1 : 0
    catch {
    }
    isCsv := PromptData_IsCsvPath(PromptData_ContextEntryPath(e))
    PromptEditor_SetCsvVisible(side, isCsv)
    if (isCsv) {
        from := (e.HasProp("csvKeepFrom") && e.csvKeepFrom >= 1) ? e.csvKeepFrom : ""
        to := (e.HasProp("csvKeepTo") && e.csvKeepTo >= 1) ? e.csvKeepTo : ""
        try ctrls.csvFrom.Value := from
        try ctrls.csvTo.Value := to
        catch {
        }
    } else {
        try ctrls.csvFrom.Value := ""
        try ctrls.csvTo.Value := ""
        catch {
        }
    }
    g_PromptEditorSuppressFlagSync := false
    g_PromptEditorFlagCtrls[side] := ctrls
}

PromptEditor_OnCompactClick(side) {
    global g_PromptEditorSuppressFlagSync
    if (g_PromptEditorSuppressFlagSync)
        return
    ctrls := PromptEditor_FlagCtrls(side)
    if (!IsObject(ctrls))
        return
    row := PromptEditor_SelectedRow(side)
    arr := PromptEditor_PathsRef(side)
    if (row < 1 || row > arr.Length)
        return
    arr[row].compact := ctrls.compact.Value ? 1 : 0
    PromptEditor_RefreshRow(side, row)
}

PromptEditor_OnCsvKeepChange(side) {
    global g_PromptEditorSuppressFlagSync
    if (g_PromptEditorSuppressFlagSync)
        return
    ctrls := PromptEditor_FlagCtrls(side)
    if (!IsObject(ctrls))
        return
    row := PromptEditor_SelectedRow(side)
    arr := PromptEditor_PathsRef(side)
    if (row < 1 || row > arr.Length)
        return
    from := 0
    to := 0
    try from := Integer(Trim(ctrls.csvFrom.Value))
    catch {
        from := 0
    }
    try to := Integer(Trim(ctrls.csvTo.Value))
    catch {
        to := 0
    }
    if (from < 1 || to < 1 || from > to) {
        from := 0
        to := 0
    }
    arr[row].csvKeepFrom := from
    arr[row].csvKeepTo := to
    PromptEditor_RefreshRow(side, row)
}

PromptEditor_HasPath(arr, path) {
    needle := StrLower(path)
    for e in arr {
        p := PromptData_ContextEntryPath(e)
        if (StrLower(p) = needle)
            return true
    }
    return false
}

PromptEditor_AppendPaths(side, incoming) {
    arr := PromptEditor_PathsRef(side)
    for raw in incoming {
        n := PromptData_StripPathQuotes(raw)
        if (n = "" || PromptEditor_HasPath(arr, n))
            continue
        arr.Push(PromptData_NewContextEntry(n))
    }
    PromptEditor_SetPaths(side, arr)
    PromptEditor_ReloadList(side)
}

PromptEditor_SelectablePathsRef(side) {
    global g_PromptEditorPersonalSelectablePaths, g_PromptEditorWorkSelectablePaths
    return (side = "work") ? g_PromptEditorWorkSelectablePaths : g_PromptEditorPersonalSelectablePaths
}

PromptEditor_SelectableLv(side) {
    global g_PromptEditorPersonalSelectableLv, g_PromptEditorWorkSelectableLv
    return (side = "work") ? g_PromptEditorWorkSelectableLv : g_PromptEditorPersonalSelectableLv
}

PromptEditor_SetSelectablePaths(side, arr) {
    global g_PromptEditorPersonalSelectablePaths, g_PromptEditorWorkSelectablePaths
    if (side = "work")
        g_PromptEditorWorkSelectablePaths := arr
    else
        g_PromptEditorPersonalSelectablePaths := arr
}

PromptEditor_ReloadSelectableList(side) {
    lv := PromptEditor_SelectableLv(side)
    if (!IsObject(lv))
        return
    lv.Delete()
    for e in PromptEditor_SelectablePathsRef(side)
        lv.Add("", PromptData_ContextEntryPath(e))
}

PromptEditor_AppendSelectablePaths(side, incoming) {
    arr := PromptEditor_SelectablePathsRef(side)
    for raw in incoming {
        n := PromptData_StripPathQuotes(raw)
        if (n = "" || PromptEditor_HasPath(arr, n))
            continue
        arr.Push(PromptData_NewContextEntry(n))
    }
    PromptEditor_SetSelectablePaths(side, arr)
    PromptEditor_ReloadSelectableList(side)
}

PromptEditor_OnAddSelectableFiles(side) {
    global g_PromptEditorGui
    if (IsObject(g_PromptEditorGui)) {
        try g_PromptEditorGui.Opt("-AlwaysOnTop")
        catch {
        }
    }
    selected := FileSelect("M1", , "Select selectable context files")
    if (IsObject(g_PromptEditorGui)) {
        try g_PromptEditorGui.Opt("+AlwaysOnTop")
        try WinActivate("ahk_id " g_PromptEditorGui.Hwnd)
        catch {
        }
    }
    PromptEditor_AppendSelectablePaths(side, PromptEditor_ParseFileSelect(selected))
}

PromptEditor_OnPasteSelectablePaths(side) {
    raw := ""
    try raw := A_Clipboard
    catch {
        return
    }
    PromptEditor_AppendSelectablePaths(side, PromptData_ParsePathList(raw))
}

PromptEditor_OnRemoveSelectable(side) {
    lv := PromptEditor_SelectableLv(side)
    if (!IsObject(lv))
        return
    selected := []
    row := 0
    loop {
        row := lv.GetNext(row)
        if (!row)
            break
        selected.Push(Integer(row))
    }
    if (selected.Length = 0)
        return
    arr := PromptEditor_SelectablePathsRef(side)
    keep := []
    skip := Map()
    for idx in selected
        skip[idx] := true
    loop arr.Length {
        if (!skip.Has(A_Index))
            keep.Push(arr[A_Index])
    }
    PromptEditor_SetSelectablePaths(side, keep)
    PromptEditor_ReloadSelectableList(side)
}

PromptEditor_ParseFileSelect(result) {
    paths := []
    if (result = "" || result = 0)
        return paths
    if (IsObject(result)) {
        if (result.Length = 0)
            return paths
        first := result[1]
        if (result.Length >= 2 && DirExist(first)) {
            second := result[2]
            if (!InStr(second, "\") && !InStr(second, "/")) {
                dir := RTrim(first, "\")
                i := 2
                while (i <= result.Length) {
                    name := Trim(result[i])
                    if (name != "")
                        paths.Push(dir "\" name)
                    i++
                }
                return paths
            }
        }
        for p in result {
            n := Trim(p)
            if (n != "")
                paths.Push(n)
        }
        return paths
    }
    text := StrReplace(StrReplace(result, "`r`n", "`n"), "`r", "`n")
    if (!InStr(text, "`n")) {
        paths.Push(result)
        return paths
    }
    lines := StrSplit(text, "`n")
    dir := RTrim(Trim(lines[1]), "\")
    i := 2
    while (i <= lines.Length) {
        name := Trim(lines[i])
        if (name != "")
            paths.Push(dir "\" name)
        i++
    }
    return paths
}

PromptEditor_OnBrowseDraftFile(*) {
    global g_PromptEditorGui, g_PromptEditorDraftFile, g_PromptEditorDraftPath
    start := g_PromptEditorDraftPath != "" ? PromptData_ResolveDraftPath({ filePathDraft: g_PromptEditorDraftPath }) :
        A_ScriptDir "\assets\prompt\"
    if (IsObject(g_PromptEditorGui)) {
        try g_PromptEditorGui.Opt("-AlwaysOnTop")
        catch {
        }
    }
    selected := FileSelect(1, start, "Select draft prompt file", "Text (*.txt)")
    if (IsObject(g_PromptEditorGui)) {
        try g_PromptEditorGui.Opt("+AlwaysOnTop")
        try WinActivate("ahk_id " g_PromptEditorGui.Hwnd)
        catch {
        }
    }
    if (selected = "")
        return
    g_PromptEditorDraftPath := PromptData_ToStoredPath(selected)
    try g_PromptEditorDraftFile.Value := g_PromptEditorDraftPath
    catch {
    }
}

PromptEditor_OnPromoteDraft(*) {
    global g_PromptEditorFilePath, g_PromptEditorDraftPath, g_PromptEditorSource
    if (g_PromptEditorDraftPath = "") {
        UtilitySelector_Notify("Set a draft file path first.")
        return
    }
    draftAbs := PromptData_ResolveDraftPath({ filePathDraft: g_PromptEditorDraftPath })
    liveAbs := PromptData_ResolvePath({ filePath: g_PromptEditorFilePath, source: g_PromptEditorSource })
    if (draftAbs = "" || !FileExist(draftAbs)) {
        UtilitySelector_Notify("Draft file missing.")
        return
    }
    if (liveAbs = "") {
        UtilitySelector_Notify("Live prompt file path is missing.")
        return
    }
    if (MsgBox("Copy draft over live prompt file?`n`nDraft: " g_PromptEditorDraftPath "`nLive: " g_PromptEditorFilePath,
        "Promote draft", "YesNo Icon!") != "Yes")
        return
    try {
        content := ReadUtf8File(draftAbs)
        WriteUtf8File(liveAbs, content)
        UtilitySelector_Notify("Draft promoted to live file.")
    } catch {
        UtilitySelector_Notify("Promote failed.")
    }
}

PromptEditor_OnShowHistory(*) {
    global g_PromptEditorFilePath, g_PromptEditorSource
    if (g_PromptEditorFilePath = "") {
        UtilitySelector_Notify("Set a prompt file first.")
        return
    }
    absPath := PromptData_ResolvePath({ filePath: g_PromptEditorFilePath, source: g_PromptEditorSource })
    PromptGit_ShowHistory(absPath)
}

PromptEditor_FindPathIndex(arr, path) {
    norm := StrLower(PromptData_StripPathQuotes(path))
    idx := 0
    for e in arr {
        idx += 1
        if (StrLower(PromptData_ContextEntryPath(e)) = norm)
            return idx
    }
    return 0
}

PromptEditor_MergeContextEntries(side, incoming) {
    arr := PromptEditor_PathsRef(side)
    for e in PromptData_ParseContextEntries(incoming) {
        p := PromptData_ContextEntryPath(e)
        if (p = "")
            continue
        idx := PromptEditor_FindPathIndex(arr, p)
        if (idx = 0)
            arr.Push(e)
        else {
            arr[idx].compact := e.HasProp("compact") ? e.compact : 0
            arr[idx].csvKeepFrom := e.HasProp("csvKeepFrom") ? e.csvKeepFrom : 0
            arr[idx].csvKeepTo := e.HasProp("csvKeepTo") ? e.csvKeepTo : 0
        }
    }
    PromptEditor_SetPaths(side, arr)
    PromptEditor_ReloadList(side)
}

PromptEditor_RefreshPresetDropdowns() {
    global g_PromptEditorPersonalPreset, g_PromptEditorWorkPreset
    if (IsObject(g_PromptEditorPersonalPreset))
        PromptContextPresets_RefreshDropdown(g_PromptEditorPersonalPreset)
    if (IsObject(g_PromptEditorWorkPreset))
        PromptContextPresets_RefreshDropdown(g_PromptEditorWorkPreset)
}

PromptEditor_OnOpenPresetManager(*) {
    global g_PromptEditorGui
    owner := 0
    try {
        if (IsObject(g_PromptEditorGui))
            owner := g_PromptEditorGui.Hwnd
    } catch {
    }
    PromptContextPresets_ShowManager(owner)
}

PromptEditor_OnSaveAsPreset(*) {
    global g_PromptEditorGui, g_PromptEditorPersonalPaths, g_PromptEditorWorkPaths
    owner := 0
    try {
        if (IsObject(g_PromptEditorGui))
            owner := g_PromptEditorGui.Hwnd
    } catch {
    }
    PromptContextPresets_ShowSaveDialog(g_PromptEditorPersonalPaths, g_PromptEditorWorkPaths, owner)
}

PromptEditor_OnApplyPreset(side) {
    global g_PromptEditorPersonalPreset, g_PromptEditorWorkPreset
    label := ""
    try label := (side = "work") ? g_PromptEditorWorkPreset.Text : g_PromptEditorPersonalPreset.Text
    catch {
        return
    }
    if (label = "" || label = "(none)")
        return
    id := PromptContextPresets_IdByLabel(label)
    if (id = "")
        return
    entries := PromptContextPresets_GetEntries(id, side)
    if (entries.Length = 0) {
        UtilitySelector_Notify("Preset has no " side " files.")
        return
    }
    PromptEditor_MergeContextEntries(side, entries)
}

PromptEditor_OnBrowsePromptFile(*) {
    global g_PromptEditorGui, g_PromptEditorFile, g_PromptEditorFilePath, g_PromptEditorSource
    start := g_PromptEditorFilePath != "" ? PromptData_ResolvePath({ filePath: g_PromptEditorFilePath,
        source: g_PromptEditorSource }) : A_ScriptDir "\assets\prompt\"
    if (IsObject(g_PromptEditorGui)) {
        try g_PromptEditorGui.Opt("-AlwaysOnTop")
        catch {
        }
    }
    selected := FileSelect(1, start, "Select prompt file", "Text (*.txt)")
    if (IsObject(g_PromptEditorGui)) {
        try g_PromptEditorGui.Opt("+AlwaysOnTop")
        try WinActivate("ahk_id " g_PromptEditorGui.Hwnd)
        catch {
        }
    }
    if (selected = "")
        return
    g_PromptEditorFilePath := PromptData_ToStoredPath(selected)
    g_PromptEditorSource := "file"
    try g_PromptEditorFile.Value := g_PromptEditorFilePath
    catch {
    }
}

PromptEditor_OnAddFiles(side) {
    global g_PromptEditorGui
    if (IsObject(g_PromptEditorGui)) {
        try g_PromptEditorGui.Opt("-AlwaysOnTop")
        catch {
        }
    }
    selected := FileSelect("M1", , "Select context files")
    if (IsObject(g_PromptEditorGui)) {
        try g_PromptEditorGui.Opt("+AlwaysOnTop")
        try WinActivate("ahk_id " g_PromptEditorGui.Hwnd)
        catch {
        }
    }
    PromptEditor_AppendPaths(side, PromptEditor_ParseFileSelect(selected))
}

PromptEditor_OnPastePaths(side) {
    raw := ""
    try raw := A_Clipboard
    catch {
        return
    }
    PromptEditor_AppendPaths(side, PromptData_ParsePathList(raw))
}

PromptEditor_OnRemove(side) {
    lv := PromptEditor_Lv(side)
    if (!IsObject(lv))
        return
    selected := []
    row := 0
    loop {
        row := lv.GetNext(row)
        if (!row)
            break
        selected.Push(Integer(row))
    }
    if (selected.Length = 0)
        return
    arr := PromptEditor_PathsRef(side)
    keep := []
    skip := Map()
    for idx in selected
        skip[idx] := true
    loop arr.Length {
        if (!skip.Has(A_Index))
            keep.Push(arr[A_Index])
    }
    PromptEditor_SetPaths(side, keep)
    PromptEditor_ReloadList(side)
    PromptEditor_LoadFlagControls(side, PromptEditor_SelectedRow(side))
}

PromptEditor_OnCancel(*) {
    global g_PromptEditorResult, g_PromptEditorGui
    g_PromptEditorResult := { saved: false }
    PromptEditor_Destroy()
}

PromptEditor_OnExpectsDataOutputClick(*) {
    global g_PromptEditorExpectsDataOutput, g_PromptEditorDataOutputFormat
    on := 0
    try on := g_PromptEditorExpectsDataOutput.Value
    catch {
    }
    try g_PromptEditorDataOutputFormat.Enabled := (on = 1)
    catch {
    }
}

PromptEditor_OnSave(*) {
    global g_PromptEditorResult, g_PromptEditorName, g_PromptEditorCategory, g_PromptEditorChar
    global g_PromptEditorFilePath, g_PromptEditorSource, g_PromptEditorAuthor, g_PromptEditorIsEdit
    global g_PromptEditorPersonalPaths, g_PromptEditorWorkPaths
    global g_PromptEditorTags, g_PromptEditorVariables, g_PromptEditorPasteMode, g_PromptEditorDraftPath
    global g_PromptEditorGitCommit, g_PromptEditorAttachAsTxt
    global g_PromptEditorExpectsDataOutput, g_PromptEditorDataOutputFormat

    name := Trim(g_PromptEditorName.Value)
    if (name = "") {
        UtilitySelector_Notify("Name is required.")
        try g_PromptEditorName.Focus()
        catch {
        }
        return
    }
    category := Trim(g_PromptEditorCategory.Text)
    if (category = "")
        category := "General"
    ch := StrLower(Trim(g_PromptEditorChar.Text))
    if (ch = "" || !PromptData_IsValidChar(ch)) {
        UtilitySelector_Notify("Character must be one of the assignment pool keys.")
        return
    }
    if (g_PromptEditorFilePath = "") {
        UtilitySelector_Notify("Prompt file is required.")
        return
    }
    PromptEditor_OnCompactClick("personal")
    PromptEditor_OnCompactClick("work")
    PromptEditor_OnCsvKeepChange("personal")
    PromptEditor_OnCsvKeepChange("work")
    tags := ""
    variables := ""
    pasteMode := "default"
    draftPath := g_PromptEditorDraftPath
    try tags := Trim(g_PromptEditorTags.Value)
    catch {
    }
    try variables := Trim(g_PromptEditorVariables.Value)
    catch {
    }
    try pasteMode := PromptData_NormalizePasteMode(g_PromptEditorPasteMode.Text)
    catch {
    }
    attachAsTxt := 0
    try attachAsTxt := PromptData_NormalizeAttachAsTxt(g_PromptEditorAttachAsTxt.Value)
    catch {
    }
    expectsDataOutput := 0
    try expectsDataOutput := PromptData_NormalizeExpectsDataOutput(g_PromptEditorExpectsDataOutput.Value)
    catch {
    }
    dataOutputFormat := "file"
    try dataOutputFormat := PromptData_NormalizeDataOutputFormat(g_PromptEditorDataOutputFormat.Text)
    catch {
    }
    if (expectsDataOutput = 0)
        dataOutputFormat := "file"
    personalEntries := PromptData_ParseContextEntries(g_PromptEditorPersonalPaths)
    workEntries := PromptData_ParseContextEntries(g_PromptEditorWorkPaths)
    personalSelectable := PromptData_ParseContextEntries(g_PromptEditorPersonalSelectablePaths)
    workSelectable := PromptData_ParseContextEntries(g_PromptEditorWorkSelectablePaths)
    draft := {
        name: name,
        char: ch,
        category: category,
        author: g_PromptEditorAuthor,
        filePath: g_PromptEditorFilePath,
        source: g_PromptEditorSource,
        tags: tags,
        pasteMode: pasteMode,
        attachAsTxt: attachAsTxt,
        expectsDataOutput: expectsDataOutput,
        dataOutputFormat: dataOutputFormat,
        variables: variables,
        filePathDraft: draftPath,
        selectContextCatalog: "",
        personal_context_files: personalEntries,
        work_context_files: workEntries,
        personal_selectable_context_files: personalSelectable,
        work_selectable_context_files: workSelectable
    }
    if (!PromptLint_ConfirmSave(draft, personalEntries, workEntries))
        return
    gitCommit := false
    try gitCommit := g_PromptEditorGitCommit.Value
    catch {
    }
    gitCommitMsg := ""
    if (gitCommit) {
        commitMsg := PromptGit_PromptCommitMessage(name)
        if (commitMsg = false)
            gitCommit := false
        else
            gitCommitMsg := commitMsg
    }
    g_PromptEditorResult := {
        saved: true,
        name: name,
        category: category,
        char: ch,
        filePath: g_PromptEditorFilePath,
        source: g_PromptEditorSource,
        author: g_PromptEditorAuthor,
        tags: tags,
        pasteMode: pasteMode,
        attachAsTxt: attachAsTxt,
        expectsDataOutput: expectsDataOutput,
        dataOutputFormat: dataOutputFormat,
        variables: variables,
        filePathDraft: draftPath,
        selectContextCatalog: "",
        personal_context_files: personalEntries,
        work_context_files: workEntries,
        personal_selectable_context_files: personalSelectable,
        work_selectable_context_files: workSelectable,
        gitCommit: gitCommit,
        gitCommitMsg: gitCommitMsg
    }
    PromptEditor_Destroy()
}

PromptEditor_Destroy() {
    global g_PromptEditorGui
    PromptEditor_CloseHelp()
    PromptEditor_ContextScrollTeardown()
    PromptEditor_BindEditorHotkeys(false)
    if (IsObject(g_PromptEditorGui)) {
        try g_PromptEditorGui.Destroy()
        catch {
        }
    }
    g_PromptEditorGui := false
}

PromptEditor_HelpText() {
    return "
(
Context file flags

Flags are per file (personal and work lists are separate). They apply only when this prompt is inserted; the original files on disk are never changed. Flagged files are attached as a temp copy whose name includes a compacted tag, e.g. transactions.compacted.csv.

Attach context as .txt (prompt-level)
• When on, every context attachment is staged as a local .txt copy before paste (Gemini-friendly for .ini and other types some AI runs reject).
• When off, files keep their extensions except .ini, which is still staged as .txt for upload safety.

Compact (any file)
• JSON: minify, drop null values, replace data: URIs and long http(s) URLs.
• Markdown / text: strip trailing spaces; collapse 3+ blank lines to one.
• CSV: strip trailing spaces and drop fully empty rows (header kept).

CSV keep from–to (CSV files only)
• 1-based inclusive line range. Empty From/To means the whole file (unless Compact is also on).
• Line 1 is the header and is always kept, even if From > 1.
• If From is 1, the header is not written twice.

Example — keep 3–4 on this file:
  1  name,qty
  2  apples,4
  3  bananas,2
  4  carrots,9
  5  dates,1

Attached copy:
  name,qty
  bananas,2
  carrots,9

The header stays (line 1). Lines 3–4 are kept. Lines 2 and 5 are dropped.

Variables and includes
• {{clipboard}}, {{dictation}}, {{date}}, {{time}}, {{datetime}}, {{env}}, {{selection}} fill at paste time.
• Other {{name}} placeholders open a fill dialog unless listed in Variables (comma-separated).
• {{include:assets/prompt/_shared/rules.txt}} inlines a file (cycle-safe).

Paste modes
• default — attach context files (if any) + paste body.
• body_only — paste body only.
• body_plus_clipboard — paste body then append clipboard.
• body_attach_clipboard — same as default (explicit).
• attach_only — attach context files only.
• auto_send — paste body/attach like default, then Enter (Gemini path).

Author notes
• Content after a --- line is stripped before send (human reminders).

Context tab (Edit prompt)
• General — name, category, char, prompt file, tags, variables, paste mode, draft.
• Context — static context (always attached), selectable pool (picker at paste), presets.
• Advanced — git commit, attach as .txt, data output flags.

Context presets
• A preset is a reusable bundle of attachment files (personal and/or work).
• Apply (per side) adds or updates files in this prompt; Save on this dialog keeps them.
• Manage presets opens the preset library; Save as preset stores the current lists.
• At paste time, only the static list for your environment (personal or work) is attached automatically.

Selectable context at paste (Context tab)
• Selectable lists are the only source for the paste-time picker pool.
• Add the files you want available; choose 0–n when you invoke the prompt (they attach in addition to static context).
)"
}

PromptEditor_ShowHelp(*) {
    global g_PromptEditorGui, g_PromptEditorHelpGui
    if (IsObject(g_PromptEditorHelpGui)) {
        try WinActivate("ahk_id " g_PromptEditorHelpGui.Hwnd)
        catch {
        }
        return
    }
    ownerOpt := "+AlwaysOnTop +ToolWindow"
    try {
        if (IsObject(g_PromptEditorGui))
            ownerOpt .= " +Owner" . g_PromptEditorGui.Hwnd
    } catch {
    }
    g_PromptEditorHelpGui := Gui(ownerOpt, "Prompt editor help")
    g_PromptEditorHelpGui.SetFont("s10", "Segoe UI")
    g_PromptEditorHelpGui.Add("Edit", "xm w520 r18 ReadOnly -WantReturn Wrap", PromptEditor_HelpText())
    g_PromptEditorHelpGui.Add("Button", "xm+220 w80 Default", "Close").OnEvent("Click", PromptEditor_CloseHelp)
    g_PromptEditorHelpGui.OnEvent("Close", PromptEditor_CloseHelp)
    g_PromptEditorHelpGui.OnEvent("Escape", PromptEditor_CloseHelp)
    PromptEditor_CenterHelp()
}

PromptEditor_CenterHelp() {
    global g_PromptEditorHelpGui
    if (!IsObject(g_PromptEditorHelpGui))
        return
    mon := UtilitySelector_ActiveMonitorWorkArea()
    g_PromptEditorHelpGui.Show("Hide")
    g_PromptEditorHelpGui.GetPos(, , &gw, &gh)
    guiX := mon.left + (mon.width - gw) // 2
    guiY := mon.top + (mon.height - gh) // 2
    if (guiX < mon.left + 20)
        guiX := mon.left + 20
    if (guiY < mon.top + 20)
        guiY := mon.top + 20
    g_PromptEditorHelpGui.Show("x" . guiX . " y" . guiY)
}

PromptEditor_CloseHelp(*) {
    global g_PromptEditorHelpGui
    if (IsObject(g_PromptEditorHelpGui)) {
        try g_PromptEditorHelpGui.Destroy()
        catch {
        }
    }
    g_PromptEditorHelpGui := false
}
