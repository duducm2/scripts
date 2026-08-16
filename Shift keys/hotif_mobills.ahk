; =============================================================================
; Shift keys module: hotif_mobills.ahk
; Mobills title WinActive hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

;-------------------------------------------------------------------
; Mobills Shortcuts
;-------------------------------------------------------------------
#HotIf Mobills_ShouldHandleAppKeys()

; Shift + D : Dashboard
+d:: {
    try {
        btn := GetMobillsButton("menu-dashboard-item", "Dashboard")
        if (btn) {
            btn.Click()
        } else {
            MsgBox "Could not find the Dashboard button.", "Mobills Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error navigating to Dashboard: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + A : Contas
+a:: {
    try {
        btn := GetMobillsButton("menu-accounts-item", "Accounts")
        if (btn) {
            btn.Click()
        } else {
            MsgBox "Could not find the Contas/Accounts button.", "Mobills Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error navigating to Contas/Accounts: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + T : TransaÃ§Ãµes
+t:: {
    try {
        btn := GetMobillsButton("menu-transactions-item", "Transactions")
        if (btn) {
            btn.Click()
        } else {
            MsgBox "Could not find the TransaÃ§Ãµes/Transactions button.", "Mobills Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error navigating to TransaÃ§Ãµes/Transactions: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + C : CartÃµes de crÃ©dito
+c:: {
    try {
        btn := GetMobillsButton("menu-creditCards-item", "Credit cards")
        if (btn) {
            btn.Click()
        } else {
            MsgBox "Could not find the CartÃµes de crÃ©dito/Credit cards button.", "Mobills Navigation",
                "IconX"
        }
    } catch Error as e {
        MsgBox "Error navigating to CartÃµes de crÃ©dito/Credit cards: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + P : Planejamento
+p:: {
    try {
        btn := GetMobillsButton("menu-budgets-item", "Budgets")
        if (btn) {
            btn.Click()
        } else {
            MsgBox "Could not find the Planejamento/Budgets button.", "Mobills Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error navigating to Planejamento/Budgets: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + R : RelatÃ³rios
+r:: {
    try {
        btn := GetMobillsButton("menu-reports-item", "Reports")
        if (btn) {
            btn.Click()
        } else {
            MsgBox "Could not find the RelatÃ³rios/Reports button.", "Mobills Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error navigating to RelatÃ³rios/Reports: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + M : Mais opÃ§Ãµes
+m:: {
    try {
        btn := GetMobillsButton("menu-moreOptions-item", "More options")
        if (btn) {
            btn.Click()
        } else {
            MsgBox "Could not find the Mais opÃ§Ãµes/More options button.", "Mobills Navigation", "IconX"
        }
    } catch Error as e {
        MsgBox "Error navigating to Mais opÃ§Ãµes/More options: " e.Message, "Mobills Error", "IconX"
    }
}

#HotIf