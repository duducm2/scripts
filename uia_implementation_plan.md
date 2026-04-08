# Plano de Implementação: UIAutomation (UIA) Nativo

## Objetivo
Substituir as coordenadas de teste (dummy data) por coordenadas reais extraídas dos elementos interativos da janela ativa utilizando a API nativa do Windows (UIAutomation) via COM (Component Object Model) do AutoHotkey v2.

## Restrição Respeitada
Nenhum executável compilado de terceiros ou DLL externa não-nativa será utilizado. A comunicação será estritamente via interface COM do Windows (`UIAutomationClient`), garantindo conformidade com as políticas de segurança corporativa.

## Passos da Implementação Arquitetural

### Passo 1: Inicialização do Objeto UIA (Interface COM)
*   Instanciar o objeto principal de automação do Windows criando um `ComObject` usando o CLSID do UIAutomation (`{ff48dba4-60ef-4201-aa87-54103eef594e}`).
*   Isso fornece a interface base (`IUIAutomation`) para interagir com a árvore de acessibilidade do sistema operacional.

### Passo 2: Obter o Elemento Raiz (Janela Ativa)
*   Utilizar o método `ElementFromHandle` do objeto UIA, passando o `ActiveWinID` (o HWND capturado no momento em que o atalho é pressionado).
*   Isso define o escopo da nossa busca exclusivamente para a janela em foco.

### Passo 3: Criação de Filtros (Conditions)
Para garantir alta performance e evitar inundar a tela com letras inúteis (como painéis invisíveis ou textos estáticos), precisamos criar uma "Query" restritiva:
1.  **Filtro de Visibilidade:** Propriedade `IsOffscreen` deve ser `False`.
2.  **Filtro de Interação:** Propriedade `IsEnabled` deve ser `True`.
3.  **Filtro de Tipo de Controle (ControlType):** Criar uma condição "OU" (OR Condition) que aceite apenas elementos como:
    *   `Button` (Botões)
    *   `Hyperlink` (Links)
    *   `Edit` (Campos de texto/input)
    *   `MenuItem` (Itens de menu)
    *   `CheckBox` / `RadioButton`
*   Essas condições são combinadas usando `CreateAndCondition`.

### Passo 4: Varredura da Árvore (Tree Walking)
*   Acionar o método `FindAll` a partir do elemento raiz da janela (Passo 2).
*   Passar o escopo `TreeScope_Descendants` (procurar em todos os filhos e netos) junto com o filtro rigoroso criado no Passo 3.
*   Isso retornará uma matriz (Array) de elementos UIA que passaram no filtro.

### Passo 5: Extração das Coordenadas (BoundingRectangle)
*   Iterar pela matriz de elementos encontrados.
*   Para cada elemento, chamar a propriedade `CurrentBoundingRectangle`.
*   O Windows retornará um formato estruturado contendo as bordas `Left`, `Top`, `Right` e `Bottom` daquele elemento exato na tela.
*   Calcular o ponto ideal (ex: canto superior esquerdo do elemento + alguns pixels de margem) para desenhar a letra indicativa.

### Passo 6: Gerador Dinâmico de Dicas (Hint Generator)
*   Criar uma função que converte o índice do loop em letras sequenciais.
*   Exemplo de algoritmo matemático:
    *   Índices 1 a 26: A, B, C ... Z.
    *   Índices 27 a 702: AA, AB, AC ... ZZ.
*   Vincular a letra gerada às coordenadas extraídas e inserir no array `MousemasterElements`.

### Passo 7: Refatoração da Ação de Clique
*   Atualmente, usamos `MouseMove` e `Click`. Com o UIA, se preferirmos, podemos extrair o padrão de invocação (`InvokePattern`) nativo do elemento e ordenar que o Windows clique nele silenciosamente, ou manter a abordagem de mover o mouse e clicar fisicamente nas coordenadas precisas (esta última é mais confiável para UIAs web). Manteremos o clique físico nas coordenadas exatas do BoundingRectangle para maior compatibilidade.

---
**Resultado Esperado após esta fase:** Quando você pressionar a hotkey, o script congelará por alguns milissegundos, consultará o Windows sobre todos os botões visíveis na tela atual, e desenhará as letras do alfabeto precisamente sobre eles.