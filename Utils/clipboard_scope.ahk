; =============================================================================
; Utils module: clipboard_scope.ahk
; Save and restore the clipboard around a paste. Call Pop from finally.
; =============================================================================

global g_ClipboardScopeStack := []

ClipboardScope_Push() {
    global g_ClipboardScopeStack
    saved := ""
    try saved := ClipboardAll()
    catch {
        saved := ""
    }
    g_ClipboardScopeStack.Push(saved)
}

ClipboardScope_Pop() {
    global g_ClipboardScopeStack
    if (!g_ClipboardScopeStack.Length)
        return
    saved := g_ClipboardScopeStack.Pop()
    if (saved = "")
        return
    try A_Clipboard := saved
    catch {
    }
}
