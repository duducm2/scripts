; =============================================================================
; Shift keys module: hotif_shopee.ahk
; Shopee hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

#HotIf IsShopeeActive()

; Shift + S: Focus Shopee search field
+s::
{
    try {
        root := Shopee_GetDocRoot()
        if (!root) {
            MsgBox "Página da Shopee não disponível."
            return
        }

        field := 0
        ; Prefer the main search combo box
        try {
            field := root.FindElement({ Type: 50003, Name: "Buscar na Shopee" })
        } catch {
        }
        ; Fallback: any control with LocalizedType = "search"
        if (!field) {
            try field := root.FindElement({ LocalizedType: "search" })
        }
        ; Fallback: numeric path from document root (see shopping uia3.md)
        if (!field) {
            try field := root.ElementFromPathExist("1,1,2,3,1")
        }

        if (field) {
            try field.SetFocus()
            catch {
                try field.Click()
            }
            return
        }
        MsgBox "Campo de busca da Shopee não encontrado."
    } catch Error as e {
        MsgBox "Erro ao focar busca da Shopee: " e.Message
    }
}

; Shift + C: Carrinho de compras (cart)
+c::
{
    if Shopee_FindAndInvoke([{ Type: 50005, AutomationId: "cart_drawer_target_id" }, { Type: 50005, Name: "Carrinho",
        cs: false, matchmode: "Substring" }, { Type: 50000, Name: "Carrinho", cs: false, matchmode: "Substring" }
    ])
        return
    MsgBox "Link/botão de carrinho da Shopee não encontrado."
}

; Shift + P: Minhas compras / pedidos (speculative)
+p::
{
    if Shopee_FindAndInvoke([{ Type: 50005, Name: "Minhas compras", cs: false, matchmode: "Substring" }, { Type: 50005,
        Name: "Meus pedidos", cs: false, matchmode: "Substring" }, { Type: 50005, Name: "Pedidos", cs: false, matchmode: "Substring" }
    ])
        return
    MsgBox "Link de compras/pedidos da Shopee não encontrado (atalho especulativo)."
}

; Shift + Y: Entrega Rápida (Chegará amanhã analog)
+y::
{
    if Shopee_FindAndInvoke([{ Type: 50002, Name: "Entrega Rápida", cs: false, matchmode: "Substring" }])
        return
    MsgBox "Filtro 'Entrega Rápida' não encontrado (atalho especulativo)."
}

; Shift + F: Promoções / produtos com desconto (Full analog, speculative)
+f::
{
    if Shopee_FindAndInvoke([{ Type: 50002, Name: "Produtos com Desconto", cs: false, matchmode: "Substring" }])
        return
    MsgBox "Filtro de promoções/produtos com desconto não encontrado (atalho especulativo)."
}

; Shift + I: Compra internacional
+i::
{
    if Shopee_FindAndInvoke([{ Type: 50002, Name: "Internacional", cs: false }])
        return
    MsgBox "Filtro 'Internacional' da Shopee não encontrado."
}

; Shift + N: Envio nacional
+n::
{
    if Shopee_FindAndInvoke([{ Type: 50002, Name: "Nacional", cs: false }])
        return
    MsgBox "Filtro 'Nacional' da Shopee não encontrado."
}

; Shift + G: Frete grátis (speculative)
+g::
{
    if Shopee_FindAndInvoke([{ Type: 50020, Name: "Frete grátis", cs: false, matchmode: "Substring" }])
        return
    MsgBox "Indicador/controle de 'Frete grátis' não encontrado (atalho especulativo)."
}

; Shift + O: Ordenar por (open sort menu)
+o::
{
    if Shopee_FindAndInvoke([{ Type: 50000, Name: "Classificar por relevância", cs: false, matchmode: "Substring" }, { Type: 50000,
        Name: "Classificar por", cs: false, matchmode: "Substring" }])
        return
    MsgBox "Botão 'Classificar por' da Shopee não encontrado."
}

; Shift + R: Faixa de preço (focus min/max price edits)
+r::
{
    try {
        root := Shopee_GetDocRoot()
        if (!root) {
            MsgBox "Página da Shopee não disponível."
            return
        }

        minimo := 0
        try minimo := root.FindElement({ Type: 50004, Name: "Preço mínimo" })
        catch {
        }
        maximo := 0
        try maximo := root.FindElement({ Type: 50004, Name: "Preço máximo" })
        catch {
        }

        target := minimo ? minimo : maximo
        if (target) {
            try target.SetFocus()
            catch {
                try target.Click()
            }
            Sleep 50
            Send "^a"
            return
        }

        MsgBox "Campos de faixa de preço da Shopee não encontrados."
    } catch Error as e {
        MsgBox "Erro ao focar faixa de preço da Shopee: " e.Message
    }
}

; Shift + L: Paginação – Seguinte (results)
+l::
{
    if Shopee_NavMove(1)
        return
    MsgBox "Navegação 'Seguinte' da Shopee não encontrada ou não aplicável."
}

; Shift + K: Paginação – Anterior (results)
+k::
{
    if Shopee_NavMove(-1)
        return
    MsgBox "Navegação 'Anterior' da Shopee não encontrada ou não aplicável."
}

; Shift + A: Adicionar ao carrinho (página do produto)
+a::
{
    if Shopee_FindAndInvoke([{ Type: 50000, Name: "Adicionar Ao Carrinho", cs: false, matchmode: "Substring" }])
        return
    MsgBox "Botão 'Adicionar Ao Carrinho' não encontrado na página da Shopee."
}

; Shift + V: Favoritar (coração)
+v::
{
    if Shopee_FindAndInvoke([{ Type: 50000, Name: "Favoritar", cs: false, matchmode: "Substring" }])
        return
    MsgBox "Botão de favoritos da Shopee não encontrado."
}

; Shift + J: Continuar fluxo (Continuar / Fazer pedido)
+j::
{
    if Shopee_FindAndInvoke([{ Type: 50000, Name: "Continuar", cs: false, matchmode: "Substring" }, { Type: 50000, Name: "Fazer pedido",
        cs: false, matchmode: "Substring" }, { Type: 50000, Name: "OK", cs: false }
    ])
        return
    MsgBox "Botão de continuar/fazer pedido da Shopee não encontrado."
}

#HotIf
