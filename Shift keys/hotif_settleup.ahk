; =============================================================================
; Shift keys module: hotif_settleup.ahk
; Settle Up hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

; SettleUp Shortcuts
;-------------------------------------------------------------------
SettleUp_GetNewExpenseDialog() {
    try {
        uia := UIA_Browser()
        return uia.FindElement({ Name: "New expense", Type: "Group" })
    } catch
        return 0
}

#HotIf WinActive("Settle Up")

; Shift + A : Click Add Transaction button (UIA by Name, EN/PT)
+a:: {
    try {
        uia := UIA_Browser()
        Sleep 150
        btn := uia.FindElement({ Type: "Button", Name: "Add transaction", matchmode: "Substring" })
        if (!btn)
            btn := uia.FindElement({ Type: "Button", Name: "Adicionar transa", matchmode: "Substring" })
        if (btn) {
            btn.Click()
        } else {
            MsgBox "Could not find the Add Transaction button."
        }
    } catch Error as e {
        MsgBox "Error clicking Add Transaction: " e.Message
    }
}

; Shift + N : Focus expense name (Purpose) field in New expense dialog
+n:: {
    try {
        dialog := SettleUp_GetNewExpenseDialog()
        if (!dialog)
            return
        nameEdit := dialog.FindElement({ Type: 50004, Name: "e.g.", matchmode: "Substring" })
        if (nameEdit)
            nameEdit.SetFocus()
    } catch
        return
}

; Shift + V : Focus expense value (amount) field in New expense dialog
+v:: {
    try {
        dialog := SettleUp_GetNewExpenseDialog()
        if (!dialog)
            return
        valueEdit := dialog.FindElement({ Type: 50004 }, 4, 1)
        if (valueEdit)
            valueEdit.SetFocus()
    } catch
        return
}

#HotIf
