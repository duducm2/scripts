; AutoHotkey Mouse Navigation Tool - Mousemaster.ahk (AutoHotkey v2)
; Based on architectural plan in mousemaster_ahk_plan.md
#Requires AutoHotkey v2.0
#Warn

; Incluindo a biblioteca UIA-v2 que você já possui instalada
#Include %A_ScriptDir%\UIA-v2\Lib\UIA.ahk
#Include %A_ScriptDir%\UIA-v2\Lib\UIA_Browser.ahk

; ==============================================================================
; Phase 1: Activation and Initialization
; ==============================================================================

global MousemasterActive := false
global MousemasterOverlayGui := ""
global MousemasterElements := []
global UserInputBuffer := ""
global MM_InputHook := ""
global ActiveWinID := ""
; Cap hints after UIA FindElements (trade-off: faster overlay build vs incomplete hint coverage on huge trees).
global Mousemaster_MaxHints := 350

; ==============================================================================
; Phase 0.5: Double-Tap State (Ctrl+Alt+Win+C)
; ==============================================================================
global MM_DoubleTapArmed := false
global MM_LastPressTick := 0
global MM_DoubleTapThresholdMs := 330
global MM_DoubleTapTimer := 0
global MM_DoubleTapOverlayGui := ""

^!#c:: {
    global MousemasterActive, ActiveWinID, MM_DoubleTapArmed, MM_LastPressTick, MM_DoubleTapThresholdMs,
        MM_DoubleTapTimer

    if (MousemasterActive) {
        Mousemaster_Deactivate()
        return
    }

    try {
        ActiveWinID := WinGetID("A")
        if (!ActiveWinID) {
            ToolTip("❌ Nenhuma janela ativa encontrada!", 200, 200)
            SetTimer(() => ToolTip(), -2000)
            return
        }
    } catch as e {
        ToolTip("❌ Erro ao obter janela ativa: " e.Message, 200, 200)
        SetTimer(() => ToolTip(), -2000)
        return
    }

    ; === Double-tap detection ===
    local now := A_TickCount
    local elapsed := (MM_LastPressTick > 0) ? (now - MM_LastPressTick) : 9999

    if (MM_DoubleTapArmed && elapsed >= 0 && elapsed < MM_DoubleTapThresholdMs) {
        ; SECOND PRESS within threshold → execute text selection macro
        MM_DoubleTapArmed := false
        MM_LastPressTick := 0
        if (MM_DoubleTapTimer) {
            SetTimer(MM_DoubleTapTimer, 0)
            MM_DoubleTapTimer := 0
        }
        MM_DoubleTap_HideOverlay()
        Mousemaster_ExecuteDoubleTapSearch()
        return
    }

    ; FIRST PRESS → arm for double-tap, show overlay
    MM_LastPressTick := now
    MM_DoubleTapArmed := true
    MM_DoubleTap_ShowOverlay()

    ; Auto-disarm after threshold — fires Mousemaster_Activate if no second press
    MM_DoubleTapTimer := ObjBindMethod(MM_DoubleTapTimerObj, "OnSingleTapTimeout")
    SetTimer(MM_DoubleTapTimer, -MM_DoubleTapThresholdMs)
}

class MM_DoubleTapTimerObj {
    static OnSingleTapTimeout() {
        global MM_DoubleTapArmed, ActiveWinID, MM_DoubleTapTimer
        if (!MM_DoubleTapArmed)
            return
        MM_DoubleTapArmed := false
        MM_DoubleTapTimer := 0
        MM_DoubleTap_HideOverlay()
        Mousemaster_Activate(ActiveWinID)
    }
}

Mousemaster_Activate(WinID) {
    global MousemasterActive, MousemasterOverlayGui, MousemasterElements, UserInputBuffer, MM_InputHook, ActiveWinX,
        ActiveWinY

    MousemasterActive := true
    MousemasterElements := []
    UserInputBuffer := ""

    WinGetPos(&ActiveWinX, &ActiveWinY, &ActiveWinW, &ActiveWinH, "ahk_id " WinID)

    ; ==============================================================================
    ; Phase 2: UI Scanning and Element Detection (UIA-v2 Implementation)
    ; ==============================================================================
    try {
        RootElement := UIA.ElementFromHandle(WinID)

        if (!RootElement) {
            ToolTip("❌ UIA: Não foi possível obter o RootElement da janela.", 200, 200)
            SetTimer(() => ToolTip(), -2000)
            Mousemaster_Deactivate()
            return
        }

        allElements := RootElement.FindElements({ IsOffscreen: false, IsEnabled: true })

        if (allElements.Length = 0) {
            ToolTip("❌ UIA: Nenhum elemento interativo encontrado.", 200, 200)
            SetTimer(() => ToolTip(), -2000)
            Mousemaster_Deactivate()
            return
        }

        for uiaEl in allElements {
            try {
                cType := uiaEl.ControlType

                ; 50000=Button, 50004=Edit, 50005=Hyperlink, 50009=MenuItem, 50002=CheckBox, 50003=RadioButton
                if (cType == 50000 || cType == 50004 || cType == 50005 || cType == 50009 || cType == 50002 || cType ==
                    50003) {

                    loc := uiaEl.Location

                    if (loc.w > 5 && loc.h > 5) {
                        hint := Mousemaster_GenerateHint(MousemasterElements.Length + 1)

                        centerX := loc.x + (loc.w // 2)
                        centerY := loc.y + (loc.h // 2)

                        MousemasterElements.Push({ hint: hint, uiaElement: uiaEl, x: loc.x, y: loc.y, w: loc.w, h: loc.h,
                            cx: centerX, cy: centerY })
                        if (MousemasterElements.Length >= Mousemaster_MaxHints)
                            break
                    }
                }
            } catch {
                continue
            }
        }

        if (MousemasterElements.Length = 0) {
            ToolTip("❌ UIA: Nenhum elemento válido/clicável encontrado.", 200, 200)
            SetTimer(() => ToolTip(), -2000)
            Mousemaster_Deactivate()
            return
        }

    } catch as e {
        ToolTip("❌ UIA Erro (Ativação): " e.Message, 200, 200)
        SetTimer(() => ToolTip(), -2000)
        Mousemaster_Deactivate()
        return
    }

    ; ==============================================================================
    ; Phase 3: Visual Overlay Generation
    ; ==============================================================================
    if (MousemasterOverlayGui) {
        MousemasterOverlayGui.Destroy()
        MousemasterOverlayGui := ""
    }

    MousemasterOverlayGui := Gui("+ToolWindow +AlwaysOnTop -Caption +E0x20 -DPIScale", "MousemasterOverlay")
    MousemasterOverlayGui.BackColor := "00FF00"
    WinSetTransColor("00FF00", MousemasterOverlayGui.Hwnd)
    MousemasterOverlayGui.SetFont("s12 w700 cBlack", "Arial")

    for index, element in MousemasterElements {
        local hintX := element.x - ActiveWinX + 2
        local hintY := element.y - ActiveWinY + 2
        local hintW := 25
        local hintH := 20

        if (StrLen(element.hint) > 1)
            hintW := 40

        MousemasterOverlayGui.Add("Text", "x" hintX " y" hintY " w" hintW " h" hintH " 0x200 Border Background0xFFFFFF Center",
            element.hint)
    }

    MousemasterOverlayGui.Show("NoActivate x" ActiveWinX " y" ActiveWinY " w" ActiveWinW " h" ActiveWinH)

    ; ==============================================================================
    ; Phase 4: Input Interception using InputHook
    ; ==============================================================================
    MM_InputHook := InputHook("T10")
    MM_InputHook.KeyOpt("{Escape}", "E")
    MM_InputHook.OnChar := Mousemaster_OnChar
    MM_InputHook.OnEnd := Mousemaster_OnEnd
    MM_InputHook.Start()
}

Mousemaster_OnChar(hook, char) {
    global MousemasterElements, UserInputBuffer

    UserInputBuffer .= StrUpper(char)
    OutputDebug("Mousemaster_OnChar: Received '" char "', Buffer: '" UserInputBuffer "'")

    local foundPartialMatch := false
    local foundExactMatch := false
    local exactElement := ""
    local potentialCount := 0

    for index, element in MousemasterElements {
        if (SubStr(element.hint, 1, StrLen(UserInputBuffer)) = UserInputBuffer) {
            foundPartialMatch := true
            potentialCount++
            if (element.hint = UserInputBuffer) {
                foundExactMatch := true
                exactElement := element
            }
        }
    }

    if (foundExactMatch && potentialCount = 1) {
        hook.Stop()
        Mousemaster_PerformAction(exactElement)
    } else if (!foundPartialMatch) {
        ToolTip("❌ Não há correspondência para: " UserInputBuffer, 200, 200)
        SetTimer(() => ToolTip(), -2000)
        hook.Stop()
        Mousemaster_Deactivate()
    } else {
        ToolTip("❓ Correspondência parcial para: " UserInputBuffer, 200, 200)
        SetTimer(() => ToolTip(), -2000)
    }
}

Mousemaster_OnEnd(hook) {
    if (hook.EndReason = "EndKey" && hook.EndKey = "Escape") {
        ToolTip("❌ Mousemaster Cancelado", 200, 200)
        SetTimer(() => ToolTip(), -2000)
    } else if (hook.EndReason = "Timeout") {
        ToolTip("❌ Mousemaster Expirou", 200, 200)
        SetTimer(() => ToolTip(), -2000)
    }
    Mousemaster_Deactivate()
}

Mousemaster_Deactivate() {
    global MousemasterActive, MousemasterOverlayGui, UserInputBuffer, MM_InputHook, MM_DoubleTapArmed

    if (!MousemasterActive && !MM_DoubleTapArmed)
        return

    if (!MousemasterActive) {
        ; Only double-tap overlay cleanup needed
        MM_DoubleTapArmed := false
        MM_DoubleTap_HideOverlay()
        return
    }

    MousemasterActive := false
    UserInputBuffer := ""

    if (MM_InputHook) {
        MM_InputHook.Stop()
        MM_InputHook := ""
    }

    if (MousemasterOverlayGui) {
        MousemasterOverlayGui.Destroy()
        MousemasterOverlayGui := ""
    }
}

; ==============================================================================
; Phase 3.5: Double-Tap Overlay Helpers
; ==============================================================================
MM_GetTopCenterPos(&outX, &outY, width, height, topMargin := 8) {
    local monLeft, monTop, monRight, monBottom
    MonitorGetWorkArea(MonitorGetPrimary(), &monLeft, &monTop, &monRight, &monBottom)
    outX := monLeft + ((monRight - monLeft) - width) // 2
    outY := monTop + topMargin
}

MM_ShowTopInputBox(Prompt, Title) {
    local ibW := 300, ibH := 88
    local ibX, ibY
    MM_GetTopCenterPos(&ibX, &ibY, ibW, ibH)
    return InputBox(Prompt, Title, "x" ibX " y" ibY " w" ibW " h" ibH)
}

MM_DoubleTap_ShowOverlay() {
    global MM_DoubleTapOverlayGui
    if (MM_DoubleTapOverlayGui) {
        MM_DoubleTapOverlayGui.Destroy()
        MM_DoubleTapOverlayGui := ""
    }
    MM_DoubleTapOverlayGui := Gui("+ToolWindow +AlwaysOnTop -Caption +E0x20 -DPIScale", "MM_DoubleTapOverlay")
    MM_DoubleTapOverlayGui.BackColor := "2980B9"
    WinSetTransColor("2980B9", MM_DoubleTapOverlayGui.Hwnd)
    MM_DoubleTapOverlayGui.SetFont("s11 w700 cWhite", "Segoe UI")
    local ovlW := 280, ovlH := 36
    MM_DoubleTapOverlayGui.Add("Text", "x12 y8 w256 h20 0x200 Background2980B9 Center",
        "🔍 Double-tap for text selection")
    local ovlX, ovlY
    MM_GetTopCenterPos(&ovlX, &ovlY, ovlW, ovlH)
    MM_DoubleTapOverlayGui.Show("NoActivate x" ovlX " y" ovlY " w" ovlW " h" ovlH)
}

MM_DoubleTap_HideOverlay() {
    global MM_DoubleTapOverlayGui
    if (MM_DoubleTapOverlayGui) {
        MM_DoubleTapOverlayGui.Destroy()
        MM_DoubleTapOverlayGui := ""
    }
}

; ==============================================================================
; Phase 4.5: Double-Tap Text Selection Macro
; ==============================================================================
Mousemaster_ExecuteDoubleTapSearch() {
    ; Initialize logging (safe: creates dir, checks file exists, try/catch)
    static logFile := A_AppData "\..\Local\AutoHotkey\logs\mousemaster.log"

    ; 1. Dynamically extract the directory path from the logFile variable
    SplitPath(logFile, , &logDir)

    ; 2. If the folder structure doesn't exist, create it first
    if (!DirExist(logDir))
        DirCreate(logDir)

    ; 3. Only attempt to delete the file if it actually exists on disk
    if (FileExist(logFile))
        FileDelete(logFile)

    ; 4. Now line 313 can execute safely without throwing Error (3)
    FileAppend("Starting Mousemaster_ExecuteDoubleTapSearch...`n", logFile)

    global ActiveWinID

    ; 0. Validate target window still exists
    if (!ActiveWinID || !WinExist("ahk_id " ActiveWinID)) {
        ToolTip("❌ Target window no longer exists.", 200, 200)
        SetTimer(() => ToolTip(), -2000)
        return
    }

    ; 1. Prompt for start and end text
    local startSeq := ""
    local endSeq := ""
    local ibResult := 0

    ibResult := MM_ShowTopInputBox("Enter the START letters of the target text range:", "Double-Tap — Start Text")
    if (ibResult.Result = "Cancel")
        return
    startSeq := Trim(ibResult.Value)
    if (startSeq = "") {
        ToolTip("❌ Start text cannot be empty.", 200, 200)
        SetTimer(() => ToolTip(), -2000)
        return
    }

    ibResult := MM_ShowTopInputBox("Enter the END letters of the target text range:", "Double-Tap — End Text")
    if (ibResult.Result = "Cancel")
        return
    endSeq := Trim(ibResult.Value)
    if (endSeq = "") {
        ToolTip("❌ End text cannot be empty.", 200, 200)
        SetTimer(() => ToolTip(), -2000)
        return
    }

    ; 2. Save clipboard + capture full text via Ctrl+A → Ctrl+C
    local clipSaved := ""
    try clipSaved := ClipboardAll()

    local foundResult := ""
    local extractionSuccess := false

    try {
        WinActivate("ahk_id " ActiveWinID)
        Sleep(60)
        Send("^a")
        Sleep(80)
        local seqBefore := DllCall("GetClipboardSequenceNumber", "uint")
        Send("^c")

        ; Bounded wait for clipboard change
        local deadline := A_TickCount + 2000
        local changed := false
        while (A_TickCount < deadline) {
            local seqNow := DllCall("GetClipboardSequenceNumber", "uint")
            if (seqNow && seqNow != seqBefore) {
                changed := true
                break
            }
            Sleep(20)
        }
        if (!changed) {
            ToolTip("❌ Clipboard capture timed out.", 200, 200)
            SetTimer(() => ToolTip(), -2000)
            return
        }

        local fullText := A_Clipboard
        if (Trim(fullText) = "") {
            ToolTip("❌ No text content found in active window.", 200, 200)
            SetTimer(() => ToolTip(), -2000)
            return
        }

        ; 3. Search loop — up to 4 start-text occurrences
        local searchPos := 1
        local maxAttempts := 4
        loop maxAttempts {
            local startAt := InStr(fullText, startSeq, , searchPos)
            if (!startAt) {
                ToolTip("❌ Start text '" startSeq "' not found.", 200, 200)
                SetTimer(() => ToolTip(), -2000)
                return
            }

            local endAt := InStr(fullText, endSeq, , startAt + StrLen(startSeq))
            if (endAt) {
                foundResult := SubStr(fullText, startAt, endAt + StrLen(endSeq) - startAt)
                break
            }
            searchPos := startAt + 1
        }

        if (foundResult = "") {
            ToolTip("❌ End text '" endSeq "' not found after any start occurrence.", 200, 200)
            SetTimer(() => ToolTip(), -2000)
            return
        }

        if (StrLen(foundResult) < 3) {
            ToolTip("❌ Extracted text too short (" StrLen(foundResult) " chars).", 200, 200)
            SetTimer(() => ToolTip(), -2000)
            return
        }

        ; Success — extraction passed all checks
        extractionSuccess := true

    } finally {
        ; Restore original clipboard content (only here, not duplicated after)
        try A_Clipboard := clipSaved
    }

    ; 4. Visual selection — adaptive per app type
    if (extractionSuccess) {
        A_Clipboard := foundResult
        Sleep(80)
        WinActivate("ahk_id " ActiveWinID)
        Sleep(50)
        Send("{Escape}")
        Sleep(60)

        ; Detecta se é Cursor ou VS Code
        local isMonacoEditor := false
        try {
            local procName := WinGetProcessName("ahk_id " ActiveWinID)
            if (procName = "Cursor.exe" || procName = "Code.exe")
                isMonacoEditor := true
        }

        if (isMonacoEditor) {
            ; VS Code/Cursor: Ctrl+F, cola texto, clica checkbox "Find in Selection"
            ; via UIA (Alt+L tem bug no VS Code), depois Escape — seleção permanece.
            Send("^f")
            Sleep(80)
            Send("^v")
            Sleep(120)
            try {
                local mmUia := UIA_Browser("ahk_id " ActiveWinID)
                if (mmUia) {
                    local findCheckbox := mmUia.FindFirst({
                        Type: 50002,
                        Name: "Find in Selection (Alt+L)"
                    })
                    if (!findCheckbox) {
                        ; Tenta nome alternativo (EN vs locale)
                        findCheckbox := mmUia.FindFirst({
                            Type: 50002,
                            Name: "Find in Selection"
                        })
                    }
                    if (findCheckbox) {
                        findCheckbox.Click()
                        Sleep(60)
                    }
                }
            } catch {
                ; Fallback se UIA falhar — tenta Alt+L
                Send("!l")
            }
            Send("{Escape}")
        } else {
            ; Demais apps: Ctrl+F, cola o texto, mantém find bar aberta
            Send("^f")
            Sleep(80)
            Send("^v")
            Sleep(80)
        }

        ToolTip("✅ Texto capturado (" StrLen(foundResult) " chars)", 200, 200)
        SetTimer(() => ToolTip(), -2000)
    }
}
Mousemaster_PerformAction(elementObject) {
    global ActiveWinID

    ; 0. Destrói a interface gráfica para limpar o caminho
    Mousemaster_Deactivate()
    Sleep(80)

    ; 1. Garante que a janela correta está ativa e recebendo comandos
    if (ActiveWinID) {
        try WinActivate("ahk_id " ActiveWinID)
        Sleep(40)
    }

    cx := elementObject.cx
    cy := elementObject.cy

    ; 2. API Nativa do Windows: Move o cursor para as coordenadas físicas exatas.
    ; Ignora completamente a escala de DPI do Windows ou do AutoHotkey.
    DllCall("SetCursorPos", "int", cx, "int", cy)

    ; Pequena pausa para os listeners Javascript ('hover') da página web ou app detectarem o cursor
    Sleep(40)

    ; 3. API Nativa do Windows: Envia o evento de clique físico.
    ; 0x0002 = MOUSEEVENTF_LEFTDOWN
    ; 0x0004 = MOUSEEVENTF_LEFTUP
    DllCall("mouse_event", "uint", 0x0002, "int", 0, "int", 0, "uint", 0, "uptr", 0)
    Sleep(25) ; Duração realística da pressão de um dedo
    DllCall("mouse_event", "uint", 0x0004, "int", 0, "int", 0, "uint", 0, "uptr", 0)

    ToolTip("✅ Clique hardware nativo via DllCall.", 200, 200)
    SetTimer(() => ToolTip(), -2000)
}

; ==============================================================================
; Helper Functions
; ==============================================================================
Mousemaster_GenerateHint(index) {
    local hint := ""
    local alphabet := "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    local base := StrLen(alphabet)

    index--

    if (index < 0)
        return ""

    while (true) {
        local remainder := Mod(index, base)
        hint := SubStr(alphabet, remainder + 1, 1) . hint
        index := Floor(index / base) - 1
        if (index < 0)
            break
    }
    return hint
}
