; =============================================================================
; Shift keys module: cheat_sheet_data.ahk
; cheatSheets map population (Mercado Livre, Shopee)
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

; Map that stores the pop-up text for each application.  Extend freely.
cheatSheets := Map()

; --- Mercado Livre (Brazil) -----------------------------------------------
cheatSheets["Mercado Livre"] := "
(
    Mercado Livre (Shift)
    [S] Focus search field
    [C] Carrinho de compras (cart)
    [P] Compras feitas (purchases)
    [Y] Filtro Chegará amanhã
    [F] Filtro Full
    [I] Filtro Internacional
    [N] Filtro Envio local / Nacional
    [G] Filtro Frete grátis
    [O] Ordenar por (menu)
    [R] Faixa de preço (Mín/Máx)
    [L] Seguinte (paginação)
    [K] Anterior (paginação)
    [A] Adicionar ao carrinho
    [V] Favoritos (coração)
    [J] Continuar (fluxo compra/endereço)
)"  ; end Mercado Livre

; --- Shopee (Brazil) -------------------------------------------------------
cheatSheets["Shopee"] := "
(
    Shopee (Shift)
    [S] Buscar na Shopee (campo de busca)
    [C] Carrinho de compras
    [P] Minhas compras / Pedidos (especulativo)
    [Y] Filtro Entrega Rápida (analogia Chegará amanhã)
    [F] Filtro Promoções / Full (especulativo)
    [I] Filtro Internacional
    [N] Filtro Envio Nacional
    [G] Filtro Frete grátis (especulativo)
    [O] Ordenar por (menu)
    [R] Faixa de preço (Mín/Máx)
    [L] Seguinte (paginação) – especulativo
    [K] Anterior (paginação) – especulativo
    [A] Adicionar ao carrinho (página do produto)
    [V] Favoritar produto (coração)
    [J] Continuar (carrinho/checkout, incl. \"Fazer pedido\")
)"  ; end Shopee
