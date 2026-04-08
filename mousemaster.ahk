; AutoHotkey Mouse Navigation Tool - Mousemaster.ahk (AutoHotkey v2)
; Based on architectural plan in mousemaster_ahk_plan.md
#Requires AutoHotkey v2.0
#Warn

; Incluindo a biblioteca UIA-v2 que você já possui instalada
#Include %A_ScriptDir%\UIA-v2\Lib\UIA.ahk

; ==============================================================================
; Phase 1: Activation and Initialization
; ==============================================================================

global MousemasterActive := false
global MousemasterOverlayGui := ""
global MousemasterElements := []
global UserInputBuffer := ""
global MM_InputHook := ""
global ActiveWinID := "" 

^!#c:: {
    global MousemasterActive, ActiveWinID
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

    Mousemaster_Activate(ActiveWinID)
}

Mousemaster_Activate(WinID) {
    global MousemasterActive, MousemasterOverlayGui, MousemasterElements, UserInputBuffer, MM_InputHook, ActiveWinX, ActiveWinY

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

        allElements := RootElement.FindElements({IsOffscreen: false, IsEnabled: true})

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
                if (cType == 50000 || cType == 50004 || cType == 50005 || cType == 50009 || cType == 50002 || cType == 50003) {
                    
                    loc := uiaEl.Location 
                    
                    if (loc.w > 5 && loc.h > 5) {
                        hint := Mousemaster_GenerateHint(MousemasterElements.Length + 1)
                        
                        centerX := loc.x + (loc.w // 2)
                        centerY := loc.y + (loc.h // 2)
                        
                        MousemasterElements.Push({hint: hint, uiaElement: uiaEl, x: loc.x, y: loc.y, w: loc.w, h: loc.h, cx: centerX, cy: centerY})
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

        MousemasterOverlayGui.Add("Text", "x" hintX " y" hintY " w" hintW " h" hintH " 0x200 Border Background0xFFFFFF Center", element.hint)
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
    global MousemasterActive, MousemasterOverlayGui, UserInputBuffer, MM_InputHook

    if (!MousemasterActive)
        return

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
; Phase 5: Action Execution (Raw Windows API Click)
; ==============================================================================
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
