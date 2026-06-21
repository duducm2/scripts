; =============================================================================
; Shift keys module: mobills_running_banner.ahk
; Mobills running overlay banner
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

; =============================================================================
; Mobills "running" banner (non-blocking, overlay style; dark background, blue accent border)
; =============================================================================
global g_MobillsRunningBannerGui := 0
global g_MobillsRunningBannerBorderGui := 0

Mobills_ShowRunningBanner(dir) {
    global g_MobillsRunningBannerGui, g_MobillsRunningBannerBorderGui

    ; Close any existing banner first
    try {
        if IsObject(g_MobillsRunningBannerBorderGui)
            g_MobillsRunningBannerBorderGui.Destroy()
    } catch {
    }
    g_MobillsRunningBannerBorderGui := 0
    try {
        if IsObject(g_MobillsRunningBannerGui)
            g_MobillsRunningBannerGui.Destroy()
    } catch {
    }
    g_MobillsRunningBannerGui := 0

    text := "Mobills: " . ((dir = "Prev") ? "Previous" : "Next") . " (running...)"

    target := WinGetID("A")
    hasWindow := false
    if target && WinExist("ahk_id " target) {
        try {
            WinGetPos(&wx, &wy, &ww, &wh, target)
            hasWindow := (ww > 0 && wh > 0)
        } catch {
            hasWindow := false
        }
    }

    ov := Gui("+AlwaysOnTop -Caption +ToolWindow")
    ov.BackColor := "1E1E2E"
    ov.SetFont("s24 cFFFFFF Bold", "Segoe UI")
    ov.Add("Text", "w500 Center", text)
    ov.Show("AutoSize Hide")
    ov.GetPos(&gx, &gy, &gw, &gh)

    if hasWindow {
        cx := wx + (ww - gw) // 2
        cy := wy + (wh - gh) // 2
    } else {
        vx := SysGet(76)  ; SM_XVIRTUALSCREEN
        vy := SysGet(77)  ; SM_YVIRTUALSCREEN
        vw := SysGet(78)  ; SM_CXVIRTUALSCREEN
        vh := SysGet(79)  ; SM_CYVIRTUALSCREEN
        cx := vx + (vw - gw) // 2
        cy := vy + (vh - gh) // 2
    }

    borderWidth := 6
    borderGui := Gui("+AlwaysOnTop -Caption +ToolWindow")
    borderGui.BackColor := BANNER_ACCENT_INTERMEDIATE
    borderGui.Show("NA x" . (cx - borderWidth) . " y" . (cy - borderWidth) . " w" . (gw + 2 * borderWidth) .
    " h" . (gh +
        2 * borderWidth))
    g_MobillsRunningBannerBorderGui := borderGui

    ov.Show("x" . cx . " y" . cy . " NA")
    WinSetTransparent(178, ov)
    g_MobillsRunningBannerGui := ov
}

Mobills_HideRunningBanner() {
    global g_MobillsRunningBannerGui, g_MobillsRunningBannerBorderGui
    try {
        if IsObject(g_MobillsRunningBannerBorderGui)
            g_MobillsRunningBannerBorderGui.Destroy()
    } catch {
    }
    g_MobillsRunningBannerBorderGui := 0
    try {
        if IsObject(g_MobillsRunningBannerGui)
            g_MobillsRunningBannerGui.Destroy()
    } catch {
    }
    g_MobillsRunningBannerGui := 0
}

; One-shot pager resolution (no waiting/retries). Returns element or "".
Mobills_FindPagerOnce(uia, dir, context) {
    btn := ""
    ; Context-specific ordering (per plan)
    if (context = "transactions") {
        ; Prefer the arrows next to the month/year header (mobile/desktop), avoid table pagination.
        btn := Mobills_FindMonthNavByMonthYear(uia, dir)
        if !btn
            btn := Mobills_FindPagerByMonthHeader(uia, dir)
        if !btn
            btn := Mobills_FindPagerByName(uia, dir)
    } else if (context = "accounts") {
        btn := Mobills_FindPagerByPath(uia, dir, context)
        if !btn
            btn := Mobills_FindPagerByName(uia, dir)
    } else if (context = "budgets") {
        ; Budgets: ONLY use the known arrow buttons by index (avoid misclicking "Next").
        btn := Mobills_FindPagerByPath(uia, dir, context)
    } else if (context = "planning") {
        btn := Mobills_FindPagerByPath(uia, dir, context)
        if !btn
            btn := Mobills_FindPagerByName(uia, dir)
    } else {
        ; Generic fallback: try name -> month header -> legacy path
        btn := Mobills_FindPagerByName(uia, dir)
        if !btn
            btn := Mobills_FindPagerByMonthHeader(uia, dir)
        if !btn
            btn := Mobills_FindPagerByPath(uia, dir, context)
    }
    return btn
}

; Multi-layer verification to confirm absence (prevents false negatives).
; Returns an element if found in any layer; otherwise returns "".
Mobills_VerifyPagerMissing(dir, context, uiaCurrent := 0) {
    ; Layer 1: Re-check using the current attachment (cheap).
    try {
        if uiaCurrent {
            Sleep 120
            btn := Mobills_FindPagerOnce(uiaCurrent, dir, context)
            if btn
                return btn
        }
    } catch {
    }

    ; Layer 2: Fresh re-attach (new UIA tree) + wait.
    try {
        uia2 := TryAttachBrowser()
        if uia2 {
            Sleep 350
            btn := Mobills_FindPagerOnce(uia2, dir, context)
            if btn
                return btn
        }
    } catch {
    }

    ; Layer 3 (budgets only): re-attach + long wait + ONLY the known i:7/i:8 buttons.
    if (context = "budgets") {
        try {
            uia3 := TryAttachBrowser()
            if uia3 {
                Sleep 600
                if Mobills_GetBudgetsPrevNext(uia3, &prevBtn, &nextBtn) {
                    return (dir = "Prev") ? (prevBtn ? prevBtn : "") : (nextBtn ? nextBtn : "")
                }
            }
        } catch {
        }
    }

    return ""
}

Mobills_Navigate(dir) {
    Mobills_ShowRunningBanner(dir)
    try {
        uia := TryAttachBrowser()
        if !uia {
            MsgBox "Could not attach to the browser window.", "Mobills Navigation", "IconX"
            return
        }

        context := Mobills_GetContext(uia)
        maxRetries := 2
        retryDelay := 200

        loop maxRetries {
            if (A_Index > 1) {
                ; Refresh UIA tree on subsequent attempts (more reliable than just waiting).
                try uia := TryAttachBrowser()
            }
            btn := Mobills_FindPagerOnce(uia, dir, context)

            if btn && Mobills_ClickPager(btn)
                return

            Sleep retryDelay
        }

        ; Multi-layer verification before declaring "not present"
        verifiedBtn := Mobills_VerifyPagerMissing(dir, context, uia)
        if verifiedBtn {
            if Mobills_ClickPager(verifiedBtn)
                return
            MsgBox "Pager control was found but could not be clicked.", "Mobills Navigation", "IconX"
            return
        }

        MsgBox "Could not find the " . ((dir = "Prev") ? "previous" : "next") .
        " page/month control (verified missing).", "Mobills Navigation", "IconX"
    } catch Error as e {
        MsgBox "Error navigating Mobills:`n" e.Message, "Mobills Error", "IconX"
    } finally {
        Mobills_HideRunningBanner()
    }
}

; Backwards-compatible wrappers (old names kept, logic refactored)
PrevMobillsMonth() => Mobills_Navigate("Prev")
NextMobillsMonth() => Mobills_Navigate("Next")

; ---- New helper to jump from "Open" button ----
FocusViaOpenButton(tabs, pressSpace := false) {
    try {
        uia := TryAttachBrowser()
        if !uia
            return false
        ; Anchor = Button named "Open"
        openBtn := uia.FindElement({ Name: "Open", Type: "Button" })
        if !openBtn {
            ; fallback by class substring
            openBtn := uia.FindElement({ ClassName: "MuiAutocomplete-popupIndicator", Type: "Button", matchmode: "Substring" })
        }
        if !openBtn
            return false
        openBtn.SetFocus()
        Sleep 200
        ; Tab forward specified times
        loop tabs {
            Send "+{Tab}"
            Sleep 80
        }
        if pressSpace {
            Sleep 80
            Send "{Space}"
        }
        return true
    } catch Error {
        return false
    }
}

; Shift + I : Toggle "Ignore transaction"
+i:: {
    try {
        uia := TryAttachBrowser()
        if !uia {
            MsgBox "Could not attach to the browser window.", "Mobills Navigation", "IconX"
            return
        }

        ignoreToggle := ""
        try {
            label := uia.FindElement({ Name: "Ignore transaction", Type: 50020, matchmode: "Substring" })
            if label {
                parent := UIA.TreeWalkerTrue.GetParentElement(label)
                if parent {
                    for , cb in parent.FindAll({ Type: 50002 }) {
                        ignoreToggle := cb
                        break
                    }
                }
            }
        } catch {
        }
        if !ignoreToggle {
            ; Fallback: pick the last checkbox in the dialog, which is the ignore toggle in current forms.
            try {
                checkboxes := uia.FindAll({ Type: 50002 })
                if (checkboxes && checkboxes.Length > 0)
                    ignoreToggle := checkboxes[checkboxes.Length]
            } catch {
            }
        }
        if !ignoreToggle {
            MsgBox "Could not find the Ignore transaction toggle.", "Mobills Navigation", "IconX"
            return
        }

        try ignoreToggle.SetFocus()
        Sleep 80
        Send "{Space}"

    } catch Error as e {
        MsgBox "Error toggling Ignore transaction: " e.Message, "Mobills Error", "IconX"
    }
}

; Click "New" and select the requested creation menu item.
Mobills_SelectNewMenuItem(itemName) {
    try {
        uia := TryAttachBrowser()
        if !uia
            return false

        actionBtn := Mobills_FindElementByCandidates(uia, [{ Type: 50000, AutomationId: "action-button" }, { Type: 50000,
            Name: "New", matchmode: "Substring" }])
        if !actionBtn
            return false
        actionBtn.Click()
        Sleep 250

        menuItem := Mobills_FindElementByCandidates(uia, [{ Type: 50011, Name: itemName, matchmode: "Substring" }, { Type: 50000,
            Name: itemName, matchmode: "Substring" }])
        if !menuItem
            return false
        menuItem.Click()
        return true
    } catch {
        return false
    }
}

; ---- Helper to focus the Description field directly ----
FocusDescriptionField() {
    try {
        uia := TryAttachBrowser()
        if !uia
            return false

        descriptionElement := Mobills_FindElementByCandidates(uia, [{ Name: "Description", Type: 50004,
            matchmode: "Substring" }, { Name: "Description",
                Type: "Edit", matchmode: "Substring" }, { ClassName: "MuiAutocomplete-input", Type: "Edit",
                    matchmode: "Substring" }
        ])

        if !descriptionElement {
            MsgBox "Could not find the Description field.", "Mobills Navigation", "IconX"
            return false
        }

        try descriptionElement.Click()
        Sleep 100
        descriptionElement.SetFocus()
        return true
    } catch Error as e {
        MsgBox "Error focusing Description field: " e.Message, "Mobills Error", "IconX"
        return false
    }
}

; Focuses the first account/category "Open" picker and types MAIN.
Mobills_TypeMainInOpenPicker() {
    try {
        uia := TryAttachBrowser()
        if !uia
            return false

        openBtn := Mobills_FindOpenButton(uia, 1)
        if !openBtn
            return false

        try openBtn.Click()
        try openBtn.SetFocus()
        Sleep 120
        Send "MAIN"
        return true
    } catch {
        return false
    }
}

; Shift + N : Focus name/description field
+n:: FocusDescriptionField()

; Shift + E : Click action button then Expense menu item
+e:: {
    try {
        if !Mobills_SelectNewMenuItem("Expense")
            MsgBox "Could not open New > Expense.", "Mobills Navigation", "IconX"
    } catch Error as e {
        MsgBox "Error clicking action button and Expense menu: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + Y : Click action button then Income menu item
+y:: {
    try {
        if !Mobills_SelectNewMenuItem("Income")
            MsgBox "Could not open New > Income.", "Mobills Navigation", "IconX"
    } catch Error as e {
        MsgBox "Error clicking action button and Income menu: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + X : Click action button then Credit card expense menu item
+x:: {
    try {
        if !Mobills_SelectNewMenuItem("Credit card expense")
            MsgBox "Could not open New > Credit card expense.", "Mobills Navigation", "IconX"
    } catch Error as e {
        MsgBox "Error clicking action button and Credit card expense menu: " e.Message, "Mobills Error",
            "IconX"
    }
}

; Shift + F : Click action button then Transfer menu item
+f:: {
    try {
        if !Mobills_SelectNewMenuItem("Transfer")
            MsgBox "Could not open New > Transfer.", "Mobills Navigation", "IconX"
    } catch Error as e {
        MsgBox "Error clicking action button and Transfer menu: " e.Message, "Mobills Error", "IconX"
    }
}

; Shift + W : Focus "Open" picker and type "MAIN"
+w:: {
    try {
        if !Mobills_TypeMainInOpenPicker()
            MsgBox "Could not find the Open picker.", "Mobills Navigation", "IconX"
    } catch Error as e {
        MsgBox "Error finding Open picker: " e.Message, "Mobills Error", "IconX"
    }
}
