; =============================================================================
; Shift keys module: hotif_mercado_livre.ahk
; Mercado Livre hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

#HotIf IsMercadoLivreActive()

; Shift + S: Focus Mercado Livre search field
+s::
{
    ML_EnsureHotkeyReceptivity()
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 200

        ; Prefer the document root; fall back to browser element only if needed
        isDocRoot := false
        try {
            root := uia.GetCurrentDocumentElement()
            isDocRoot := true
        } catch {
            root := uia.BrowserElement
        }

        field := 0

        ; 1) Try AutomationId from the current document (or fallback root)
        try {
            field := root.FindElement({ AutomationId: "cb1-edit" })
        } catch {
        }

        ; 2) From the document root, try the numeric path 1,1,4,2 if available
        if (!field && isDocRoot) {
            try {
                field := root.ElementFromPath("1,1,4,2")
            } catch {
            }
        }

        ; 3) As a last resort, search from the browser element with Descendants scope
        if (!field) {
            try {
                field := uia.BrowserElement.FindElement({ AutomationId: "cb1-edit" }, UIA.TreeScope.Descendants)
            } catch {
            }
        }

        if (field) {
            focusOk := false
            try {
                field.SetFocus()
                focusOk := true
            } catch {
                try {
                    field.Click()
                    focusOk := true
                } catch {
                }
            }
            if (focusOk)
                return
        }

        ; Error-driven workaround: force right-click + close menu, then retry once
        ML_EnsureHotkeyReceptivity(true)
        Sleep 300

        field := 0
        try {
            try
                field := root.FindElement({ AutomationId: "cb1-edit" })
            catch
                field := 0
            if (!field && isDocRoot) {
                try
                    field := root.ElementFromPath("1,1,4,2")
                catch
                    field := 0
            }
            if (!field) {
                try
                    field := uia.BrowserElement.FindElement({ AutomationId: "cb1-edit" }, UIA.TreeScope.Descendants)
                catch
                    field := 0
            }
            if (field) {
                try
                    field.SetFocus()
                catch {
                    try
                        field.Click()
                    catch {
                    }
                }
                return
            }
        } catch {
        }

        MsgBox "Could not find Mercado Livre search field."
    } catch Error as e {
        MsgBox "An error occurred: " e.Message
    }
}

; Shift + C: Carrinho de compras (Cart)
+c::
{
    ML_EnsureHotkeyReceptivity()
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 200
        try {
            root := uia.GetCurrentDocumentElement()
        } catch {
            root := uia.BrowserElement
        }

        cart := 0
        ; Prefer AutomationId
        try {
            cart := root.FindElement({ AutomationId: "nav-cart" })
        } catch {
        }
        if (!cart) {
            ; Try by class name substring
            try {
                cart := root.FindElement({ ClassName: "nav-cart", matchmode: "Substring" })
            } catch {
            }
        }
        if (!cart) {
            ; Try by link name containing 'carrinho'
            try {
                cart := root.FindElement({ Type: 50005, Name: "carrinho", cs: false, matchmode: "Substring" })
            } catch {
            }
        }

        if (cart) {
            try cart.Invoke()
            catch {
                try cart.Click()
            }
            return
        }
        MsgBox "Could not find Mercado Livre cart link."
    } catch Error as e {
        MsgBox "An error occurred: " e.Message
    }
}

; Shift + P: Compras (Purchases)
+p::
{
    ML_EnsureHotkeyReceptivity()
    try {
        uia := UIA_Browser("ahk_exe chrome.exe")
        Sleep 200
        try {
            root := uia.GetCurrentDocumentElement()
        } catch {
            root := uia.BrowserElement
        }

        purchases := 0
        ; Try by class name first
        try {
            purchases := root.FindElement({ ClassName: "option-purchases" })
        } catch {
        }
        if (!purchases) {
            ; Try by link name 'Compras'
            try {
                purchases := root.FindElement({ Type: 50005, Name: "Compras", cs: false, matchmode: "Substring" })
            } catch {
            }
        }

        if (purchases) {
            try purchases.Invoke()
            catch {
                try purchases.Click()
            }
            return
        }
        MsgBox "Could not find Mercado Livre purchases link."
    } catch Error as e {
        MsgBox "An error occurred: " e.Message
    }
}

; Shift + Y: Chegará amanhã (filter toggle)
+y::
{
    ML_EnsureHotkeyReceptivity()
    if ML_FindAndInvoke([{ Type: 50000, AutomationId: "shipping_time_highlighted_nextday" }])
        return
    MsgBox "Filtro 'Chegará amanhã' não encontrado."
}

; Shift + F: Full (frete grátis Full)
+f::
{
    ML_EnsureHotkeyReceptivity()
    if ML_FindAndInvoke([{ Type: 50000, AutomationId: "shipping_highlighted_fulfillment" }])
        return
    MsgBox "Filtro 'Full' não encontrado."
}

; Shift + I: Compra Internacional
+i::
{
    ML_EnsureHotkeyReceptivity()
    if ML_FindAndInvoke([{ Type: 50000, AutomationId: "SHIPPING_ORIGIN_HIGHLIGHTED" }])
        return
    MsgBox "Filtro 'Internacional' não encontrado."
}

; Shift + N: Envio local / Produtos com frete nacional
+n::
{
    ML_EnsureHotkeyReceptivity()
    if ML_FindAndInvoke([{ Type: 50000, AutomationId: "SHIPPING_ORIGIN_LOCAL_HIGHLIGHTED" }, { Type: 50000, Name: "Envio local",
        cs: false, matchmode: "Substring" }])
        return
    MsgBox "Filtro 'Produtos com frete nacional' não encontrado."
}

; Shift + G: Frete grátis
+g::
{
    ML_EnsureHotkeyReceptivity()
    if ML_FindAndInvoke([{ Type: 50000, AutomationId: "shipping_cost_highlighted_free" }])
        return
    MsgBox "Filtro 'Frete grátis' não encontrado."
}

; Shift + O: Ordenar por – Handy-style menu and UIA execution
; Returns current option label from list selected item (Layer 1) or trigger (Layer 2). Single authority for quality-check reads.
ML_SortGetCurrentLabel(sortList, trigger) {
    currentLabel := ""
    if (sortList) {
        try {
            children := sortList.FindAll({ Type: 50007 })
            for ch in children {
                try {
                    if (ch.SelectionItemPattern.IsSelected) {
                        currentLabel := Trim(ch.Name)
                        if (currentLabel = "")
                            currentLabel := Trim(ch.Value)
                        if (currentLabel != "")
                            return currentLabel
                    }
                } catch {
                }
            }
        } catch {
        }
    }
    try {
        currentLabel := Trim(trigger.Name)
        if (currentLabel = "")
            currentLabel := Trim(trigger.Value)
    } catch {
    }
    return currentLabel
}

ML_SortClose() {
    try Hotkey("1", "Off")
    catch {
    }
    try Hotkey("2", "Off")
    catch {
    }
    try Hotkey("3", "Off")
    catch {
    }
    try Hotkey("Escape", ML_SortCancel, "Off")
    catch {
    }
    try Hotkey("Enter", "Off")
    catch {
    }
    global g_ML_SortGui, g_ML_SortLv
    if (g_ML_SortGui && IsObject(g_ML_SortGui) && g_ML_SortGui.Hwnd)
        try g_ML_SortGui.Destroy()
    g_ML_SortGui := 0
    g_ML_SortLv := false
}

ML_SortCancel(*) {
    ML_SortClose()
}

ML_SortSelect(idx) {
    ML_SortClose()
    ML_SortApply(idx)
}

ML_SortOnEnter(*) {
    global g_ML_SortLv
    if (!IsObject(g_ML_SortLv))
        return
    row := 0
    try row := g_ML_SortLv.GetNext(0, "Focused")
    catch {
        row := 0
    }
    if (row < 1) {
        try row := g_ML_SortLv.GetNext(0, "Selected")
        catch {
            row := 0
        }
    }
    if (row < 1)
        return
    ch := ""
    try ch := g_ML_SortLv.GetText(row, 1)
    catch {
        return
    }
    if (!RegExMatch(ch, "^\d+$"))
        return
    ML_SortSelect(Integer(ch))
}

ML_SortOnListActivate(*) {
    ML_SortOnEnter()
}

ML_SortApply(idx) {
    StandardLoadingBar_Show("⏳ Ordenando...", BANNER_ACCENT_INTERMEDIATE, { passive: false, centerOnHwnd: 0,
        textWidth: 380, fontSize: 17 })
    try {
        root := ML_GetDocRoot()
        if (!root) {
            StandardLoadingBar_Update("❌ Página do Mercado Livre não disponível.")
            Sleep 800
            return
        }
        trigger := ML_Find(root, { Type: 50003, AutomationId: "5clcjae", matchmode: "Substring" })
        if (!trigger)
            trigger := ML_Find(root, { Type: 50003, AutomationId: "trigger", matchmode: "Substring" })
        if (!trigger) {
            StandardLoadingBar_Update("❌ Botão 'Ordenar por' não encontrado.")
            Sleep 800
            return
        }
        ; Current selection: menu opens with this option highlighted. Order in menu: 1=Mais relevantes, 2=Menor preço, 3=Maior preço
        current := 1
        label := ""
        try {
            label := Trim(trigger.Name)
            if (label = "")
                label := Trim(trigger.Value)
            if (InStr(label, "Maior preço"))
                current := 3
            else if (InStr(label, "Menor preço"))
                current := 2
            else if (InStr(label, "Mais relevantes"))
                current := 1
        } catch {
        }
        StandardLoadingBar_Update("⏳ Abrindo menu...")
        clickOk := false
        try {
            trigger.Click()
            clickOk := true
        } catch {
        }
        if (!clickOk) {
            StandardLoadingBar_Update("❌ Não foi possível abrir o menu.")
            Sleep 800
            return
        }
        ; Bounded wait for menu: poll for sortList up to 1.2s (canon: condition-based waits with timeout).
        searchRoot := root
        sortList := 0
        loop 12 {
            Sleep 100
            try {
                dropdownSibling := UIA.TreeWalkerTrue.TryGetNextSiblingElement(trigger)
                if (dropdownSibling)
                    sortList := ML_Find(dropdownSibling, { Type: 50008, AutomationId: "menu-list", matchmode: "Substring" })
            } catch {
            }
            if (sortList)
                break
            sortList := ML_Find(searchRoot, { Type: 50008, AutomationId: "5clcjae_-menu-list", matchmode: "Substring" })
            if (sortList)
                break
            sortList := ML_Find(searchRoot, { Type: 50008, AutomationId: "menu-list", matchmode: "Substring" })
            if (sortList)
                break
            try {
                triggerParent := trigger.Parent
                if (triggerParent)
                    sortList := ML_Find(triggerParent, { Type: 50008, AutomationId: "menu-list", matchmode: "Substring" })
            } catch {
            }
            if (sortList)
                break
            sortList := ML_Find(searchRoot, { AutomationId: "menu-list", matchmode: "Substring" })
            if (sortList)
                break
        }
        StandardLoadingBar_Update("⏳ Selecionando opção...")
        optionSubstrings := ["menu-list-option-relevance", "menu-list-option-price_asc", "menu-list-option-price_desc"]
        sub := optionSubstrings[idx]
        optionNames := ["Mais relevantes", "Menor preço", "Maior preço"]
        item := 0
        if (sortList) {
            item := ML_Find(sortList, { Type: 50007, AutomationId: sub, matchmode: "Substring" })
            if (!item)
                item := ML_Find(sortList, { Type: 50007, Name: optionNames[idx], cs: false })
        }
        if (!item) {
            item := ML_Find(searchRoot, { Type: 50007, AutomationId: sub, matchmode: "Substring" })
            if (!item)
                item := ML_Find(searchRoot, { Type: 50007, Name: optionNames[idx], cs: false })
        }
        if (item) {
            try item.Click()
            catch
                try item.Invoke()
            Sleep 150
            StandardLoadingBar_Update("✅ Ordenação aplicada")
            Sleep 300
        } else {
            ; Keyboard: navigate from current to target (delta), then confirm with Enter after label check.
            try WinActivate("ahk_exe chrome.exe")
            catch {
            }
            Sleep 400
            delta := idx - current
            if (delta > 0) {
                loop delta {
                    Send "{Down}"
                    Sleep 100
                }
            } else if (delta < 0) {
                loop (-delta) {
                    Send "{Up}"
                    Sleep 100
                }
            }
            Sleep 200
            targetName := optionNames[idx]
            currentLabel := ML_SortGetCurrentLabel(sortList, trigger)
            labelMatches := (currentLabel != "" && (InStr(currentLabel, targetName) || InStr(targetName, currentLabel)))
            if (labelMatches) {
                Send "{Enter}"
                StandardLoadingBar_Update("✅ Ordenação aplicada")
                Sleep 300
            } else {
                currentIdx := 1
                if (InStr(currentLabel, "Maior preço"))
                    currentIdx := 3
                else if (InStr(currentLabel, "Menor preço"))
                    currentIdx := 2
                else if (InStr(currentLabel, "Mais relevantes"))
                    currentIdx := 1
                if (idx > currentIdx) {
                    Send "{Down}"
                    Sleep 100
                } else if (idx < currentIdx) {
                    Send "{Up}"
                    Sleep 100
                }
                currentLabel := ML_SortGetCurrentLabel(sortList, trigger)
                labelMatches := (currentLabel != "" && (InStr(currentLabel, targetName) || InStr(targetName,
                    currentLabel)))
                Send "{Enter}"
                StandardLoadingBar_Update("✅ Ordenação aplicada")
                Sleep 300
            }
        }
    } catch Error as err {
        try StandardLoadingBar_Update("❌ Erro: " SubStr(err.Message, 1, 40))
        catch {
        }
        Sleep 700
    } finally {
        try StandardLoadingBar_Hide(0)
        catch {
        }
    }
}

+o::
{
    ML_EnsureHotkeyReceptivity()
    root := ML_GetDocRoot()
    if (!root) {
        ; Retry once after workaround and delay (document may not have been ready on first load)
        ML_EnsureHotkeyReceptivity(true)
        Sleep 400
        root := ML_GetDocRoot()
    }
    if (!root) {
        MsgBox "Página do Mercado Livre não disponível."
        return
    }
    global g_ML_SortGui, g_ML_SortLv
    g_ML_SortGui := Gui("+AlwaysOnTop +ToolWindow", "Ordenar por")
    g_ML_SortGui.SetFont("s10", "Segoe UI")
    g_ML_SortGui.Add("Text", "w420",
        "Char = select   Enter/double-click = select   Esc = cancel")
    g_ML_SortLv := g_ML_SortGui.Add("ListView", "w420 h120 -Multi", ["Char", "Option", "Description"])
    g_ML_SortLv.OnEvent("DoubleClick", ML_SortOnListActivate)
    g_ML_SortGui.Add("Button", "w100", "Close").OnEvent("Click", ML_SortCancel)
    g_ML_SortGui.OnEvent("Close", ML_SortCancel)
    g_ML_SortGui.OnEvent("Escape", ML_SortCancel)
    g_ML_SortLv.Add("", "1", "Mais relevantes", "Relevância da busca")
    g_ML_SortLv.Add("", "2", "Menor preço", "Preço crescente")
    g_ML_SortLv.Add("", "3", "Maior preço", "Preço decrescente")
    try g_ML_SortLv.ModifyCol(1, 50)
    try g_ML_SortLv.ModifyCol(2, 140)
    try g_ML_SortLv.ModifyCol(3, 200)
    try g_ML_SortLv.Modify(1, "Select Focus Vis")
    catch {
    }
    activeWin := 0
    try
        activeWin := WinGetID("A")
    catch
        activeWin := 0
    MonitorGetWorkArea(1, &monitorLeft, &monitorTop, &monitorRight, &monitorBottom)
    monitorWidth := monitorRight - monitorLeft
    monitorHeight := monitorBottom - monitorTop
    if (activeWin && activeWin != 0) {
        rect := Buffer(16, 0)
        if (DllCall("GetWindowRect", "ptr", activeWin, "ptr", rect)) {
            centerX := NumGet(rect, 0, "int") + (NumGet(rect, 8, "int") - NumGet(rect, 0, "int")) // 2
            centerY := NumGet(rect, 4, "int") + (NumGet(rect, 12, "int") - NumGet(rect, 4, "int")) // 2
            monitorCount := MonitorGetCount()
            loop monitorCount {
                idx := A_Index
                MonitorGetWorkArea(idx, &l, &t, &r, &b)
                if (centerX >= l && centerX <= r && centerY >= t && centerY <= b) {
                    monitorLeft := l
                    monitorTop := t
                    monitorRight := r
                    monitorBottom := b
                    monitorWidth := r - l
                    monitorHeight := b - t
                    break
                }
            }
        }
    }
    guiW := 440
    guiH := 220
    cx := monitorLeft + (monitorWidth - guiW) // 2
    cy := monitorTop + (monitorHeight - guiH) // 2
    g_ML_SortGui.Show("x" . cx . " y" . cy . " w" . guiW . " h" . guiH . " NA")
    Hotkey("1", (*) => ML_SortSelect(1), "On")
    Hotkey("2", (*) => ML_SortSelect(2), "On")
    Hotkey("3", (*) => ML_SortSelect(3), "On")
    Hotkey("Enter", ML_SortOnEnter, "On")
    Hotkey("Escape", ML_SortCancel, "On")
}

; Shift + L: Paginação – Seguinte
+l::
{
    ML_EnsureHotkeyReceptivity()
    if ML_FindAndInvoke([{ Type: 50005, Name: "Seguinte", cs: false }, { Type: 50000, Name: "Seguinte", cs: false }])
        return
    MsgBox "Botão 'Seguinte' não encontrado."
}

; Shift + K: Paginação – Anterior
+k::
{
    ML_EnsureHotkeyReceptivity()
    if ML_FindAndInvoke([{ Type: 50005, Name: "Anterior", cs: false }, { Type: 50000, Name: "Anterior", cs: false }])
        return
    MsgBox "Botão 'Anterior' não encontrado."
}

; Shift + A: Adicionar ao carrinho (página do produto)
+a::
{
    ML_EnsureHotkeyReceptivity()
    if ML_FindAndInvoke([{ Type: 50000, Name: "Adicionar ao carrinho", cs: false }])
        return
    MsgBox "Botão 'Adicionar ao carrinho' não encontrado."
}

; Shift + V: Adicionar aos favoritos (coração)
+v::
{
    ML_EnsureHotkeyReceptivity()
    if ML_FindAndInvoke([{ Type: 50000, Name: "Adicionar aos favoritos", cs: false }, { Type: 50000, ClassName: "ui-pdp-bookmark",
        matchmode: "Substring" }])
        return
    MsgBox "Botão 'Adicionar aos favoritos' não encontrado."
}

; Shift + J: Continuar fluxo (Continuar a compra / Continuar / OK)
+j::
{
    ML_EnsureHotkeyReceptivity()
    conditions := [{ Type: 50005, Name: "Continuar a compra", cs: false }, { Type: 50000, Name: "Continuar", cs: false }, { Type: 50000,
        Name: "OK", cs: false }, { Type: 50000, AutomationId: "shipping_footer_confirm_button" }, { Type: 50005, Name: "Continuar",
            cs: false }, { Type: 50000, Name: "Seguinte", cs: false }
    ]
    if ML_FindAndInvoke(conditions)
        return
    MsgBox "Botão de continuar não encontrado."
}

;-------------------------------------------------------------------
; Shopee (Brazil) Shortcuts
;-------------------------------------------------------------------
