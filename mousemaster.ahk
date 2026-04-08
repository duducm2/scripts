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

^!#c:: {
    global MousemasterActive
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

Mousemaster_Activate(ActiveWinID) {
    global MousemasterActive, MousemasterOverlayGui, MousemasterElements, UserInputBuffer, MM_InputHook

    MousemasterActive := true
    MousemasterElements := []
    UserInputBuffer := ""

    WinGetPos(&ActiveWinX, &ActiveWinY, &ActiveWinW, &ActiveWinH, "ahk_id " ActiveWinID)

    ; ==============================================================================
    ; Phase 2: UI Scanning and Element Detection (UIA-v2 Implementation)
    ; ==============================================================================
    try {
        ; Inicializa o RootElement usando a biblioteca UIA-v2 nativa
        RootElement := UIA.ElementFromHandle(ActiveWinID)

        if (!RootElement) {
            ToolTip("❌ UIA: Não foi possível obter o RootElement da janela.", 200, 200)
            SetTimer(() => ToolTip(), -2000)
            Mousemaster_Deactivate()
            return
        }

        ; Busca TODOS os elementos que estão visíveis e habilitados
        allElements := RootElement.FindElements({IsOffscreen: false, IsEnabled: true})

        if (allElements.Length = 0) {
            ToolTip("❌ UIA: Nenhum elemento interativo encontrado.", 200, 200)
            SetTimer(() => ToolTip(), -2000)
            Mousemaster_Deactivate()
            return
        }

        ; Filtra os tipos de controles que nos interessam e extrai as coordenadas
        for uiaEl in allElements {
            try {
                cType := uiaEl.ControlType
                
                ; 50000=Button, 50004=Edit, 50005=Hyperlink, 50009=MenuItem, 50002=CheckBox, 50003=RadioButton
                if (cType == 50000 || cType == 50004 || cType == 50005 || cType == 50009 || cType == 50002 || cType == 50003) {
                    
                    loc := uiaEl.Location ; Retorna um objeto com {x, y, w, h} absolutos na tela
                    
                    ; Filtra elementos invisíveis com tamanho muito pequeno
                    if (loc.w > 5 && loc.h > 5) {
                        hint := Mousemaster_GenerateHint(MousemasterElements.Length + 1)
                        
                        ; Calcula o centro do elemento para o clique físico ser preciso (coordenadas de tela absolutas)
                        centerX := loc.x + (loc.w // 2)
                        centerY := loc.y + (loc.h // 2)
                        
                        ; Armazena o objeto uiaEl completo!
                        MousemasterElements.Push({hint: hint, uiaElement: uiaEl, x: loc.x, y: loc.y, w: loc.w, h: loc.h, cx: centerX, cy: centerY})
                    }
                }
            } catch {
                ; Ignora elementos que possam ter desaparecido durante a varredura
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
    MousemasterOverlayGui.BackColor := "00FF00" ; Cor de fundo transparente para a GUI
    WinSetTransColor("00FF00", MousemasterOverlayGui.Hwnd)

    MousemasterOverlayGui.SetFont("s12 w700 cBlack", "Arial") ; Fonte preta para as dicas

    ; Desenha as dicas na tela
    for index, element in MousemasterElements {
        ; Cria um campo de texto com fundo branco para a dica
        ; As coordenadas são ajustadas para serem relativas à GUI
        local hintX := element.x - ActiveWinX + 2
        local hintY := element.y - ActiveWinY + 2
        local hintW := 25 ; Largura aproximada para uma letra
        local hintH := 20 ; Altura aproximada para uma letra

        ; Se a dica for de 2 caracteres (AA), ajusta a largura
        if (StrLen(element.hint) > 1)
            hintW := 40

        MousemasterOverlayGui.Add("Text", "x" hintX " y" hintY " w" hintW " h" hintH " 0x200 Border Background0xFFFFFF Center", element.hint)
    }

    MousemasterOverlayGui.Show("NoActivate x" ActiveWinX " y" ActiveWinY " w" ActiveWinW " h" ActiveWinH)

    ; ==============================================================================
    ; Phase 4: Input Interception using InputHook
    ; ==============================================================================
    MM_InputHook := InputHook("T10") ; 10 seconds timeout
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
    OutputDebug("Mousemaster_OnChar: foundExactMatch=" foundExactMatch ", potentialCount=" potentialCount ", exactElement: " IsObject(exactElement) ? exactElement.hint : "N/A")

    if (foundExactMatch && potentialCount = 1) {
        ; Única correspondência exata encontrada
        hook.Stop()
        ; Passa o objeto do elemento UIA para a ação.
        Mousemaster_PerformAction(exactElement)
    } else if (!foundPartialMatch) {
        ; Dica digitada incorretamente
        ToolTip("❌ Não há correspondência para: " UserInputBuffer, 200, 200)
        SetTimer(() => ToolTip(), -2000)
        hook.Stop()
        Mousemaster_Deactivate() ; Desativa se não houver correspondência
    } else {
        ; Correspondência parcial, continua bufferizando.
        ToolTip("❓ Correspondência parcial para: " UserInputBuffer, 200, 200)
        SetTimer(() => ToolTip(), -2000)
    }
}

Mousemaster_OnEnd(hook) {
    OutputDebug("Mousemaster_OnEnd: EndReason=" hook.EndReason ", EndKey=" hook.EndKey)
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

    if (!MousemasterActive) {
        OutputDebug("Mousemaster_Deactivate: Already inactive.")
        return
    }

    MousemasterActive := false
    UserInputBuffer := ""

    if (MM_InputHook) {
        MM_InputHook.Stop()
        MM_InputHook := ""
        OutputDebug("Mousemaster_Deactivate: InputHook stopped.")
    }

    if (MousemasterOverlayGui) {
        MousemasterOverlayGui.Destroy()
        MousemasterOverlayGui := ""
        OutputDebug("Mousemaster_Deactivate: Overlay GUI destroyed.")
    }
    OutputDebug("Mousemaster_Deactivate: Mousemaster is now inactive.")
}

; ==============================================================================
; Phase 5: Action Execution
; ==============================================================================
Mousemaster_PerformAction(elementObject) {
    OutputDebug("Mousemaster_PerformAction: Attempting to click element: " elementObject.hint)
    
    ; Primeiramente, tenta o método InvokePattern da UIA, que é o mais robusto
    try {
        if (elementObject.uiaElement.IsInvokePatternAvailable) {
            elementObject.uiaElement.Invoke()
            OutputDebug("Mousemaster_PerformAction: InvokePattern succeeded for " elementObject.hint)
            ToolTip("✅ Elemento '" elementObject.hint "' invocado com sucesso!", 200, 200)
            SetTimer(() => ToolTip(), -2000)
            Mousemaster_Deactivate()
            return
        }
    } catch as e {
        OutputDebug("Mousemaster_PerformAction: InvokePattern failed for " elementObject.hint ": " e.Message)
        ToolTip("❌ InvokePattern falhou para '" elementObject.hint "'. Tentando clique físico...", 200, 200)
        SetTimer(() => ToolTip(), -2000)
    }

    ; Fallback para clique físico
    OutputDebug("Mousemaster_PerformAction: Falling back to physical click for " elementObject.hint " at coords: " elementObject.cx ", " elementObject.cy)
    Mousemaster_Deactivate() ; Desativa a GUI e InputHook ANTES do clique físico
    Sleep(50) ; Pequena pausa para garantir que a GUI sumiu
    CoordMode("Mouse", "Screen")
    MouseMove(elementObject.cx, elementObject.cy, 0)
    Click()
    OutputDebug("Mousemaster_PerformAction: Physical click attempted.")
    ToolTip("✅ Clique físico em '" elementObject.hint "' realizado.", 200, 200)
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
