; =============================================================================
; Shift keys module: outlook_appointment_hotif_01.ahk
; Outlook appointment inspector hotkeys (part 1)
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

#HotIf IsOutlookAppointmentActive()

; -------------------------------------------------------------------
; New Outlook Appointment (New event) popover helpers
; - The date/time area opens a popover (fui-PopoverSurface) containing Start date/time, End time, All day, Recurring, Time suggestions.
; - Shortcuts must work whether the popover is open or closed.
; -------------------------------------------------------------------
Appt_LoadingShow(text) {
    try StandardLoadingBar_Show(text, BANNER_ACCENT_INTERMEDIATE, { passive: false, centerOnHwnd: 0, textWidth: 560,
        fontSize: 17 })
    catch {
    }
}

Appt_LoadingHide(delayMs := 0) {
    try StandardLoadingBar_Hide(delayMs)
    catch {
    }
}

Appt_RunWithLoading(label, fn) {
    Appt_LoadingShow("⏳ Appointment: " label "…")
    try {
        return fn.Call()
    } finally {
        Appt_LoadingHide(0)
    }
}

Appt_GetRootActive() {
    try return UIA.ElementFromHandle(WinExist("A"))
    catch {
        return 0
    }
}

Appt_IsPopoverSurface(el) {
    if !el
        return false
    cn := ""
    try cn := el.ClassName
    if !InStr(cn, "fui-PopoverSurface")
        return false
    try {
        ; Confirm it's the right popover by checking for Start date/time presence.
        if el.FindFirst({ Name: "Start date", ControlType: "ComboBox" })
            return true
        if el.FindFirst({ Name: "Start time", ControlType: "ComboBox" })
            return true
    } catch {
    }
    return false
}

Appt_FindOpenPopover() {
    ; Popovers can be hosted outside the window subtree, so search desktop root first.
    roots := []
    try roots.Push(UIA.GetRootElement())
    catch {
    }
    try roots.Push(Appt_GetRootActive())
    catch {
    }

    for root in roots {
        if !root
            continue
        try {
            pop := root.FindFirst({ ClassName: "fui-PopoverSurface", matchmode: "Substring" })
            if (Appt_IsPopoverSurface(pop))
                return pop
        } catch {
        }
        ; Fallback: search for any dialog window that contains Start date/time combo.
        try {
            w := root.FindFirst({ ControlType: "Window", LocalizedType: "dialog" })
            if (w && (w.FindFirst({ Name: "Start date", ControlType: "ComboBox" }) || w.FindFirst({ Name: "Start time",
                ControlType: "ComboBox" })))
                return w
        } catch {
        }
    }
    return 0
}

Appt_OpenPopoverIfNeeded() {
    ; Debug instrumentation removed (b96502).
    pop := Appt_FindOpenPopover()
    if pop
        return pop

    root := Appt_GetRootActive()
    if !root
        return 0

    ; Best trigger: click the Start time combo (or its caret button) to open the popover.
    try {
        trigger := root.FindFirst({ Name: "Start time", ControlType: "ComboBox" })
        if trigger {
            try trigger.SetFocus()
            Sleep 40
            try trigger.Click()
        } else {
            btn := root.FindFirst({ Name: "Start time", ControlType: "Button" })
            if btn {
                try btn.SetFocus()
                Sleep 40
                try btn.Click()
            }
        }
    } catch {
    }

    ; Fallback trigger: click the date/time range summary button (e.g. "Wed 4/1/2026 2:00 PM - 2:30 PM …")
    ; This is required in Scheduler view where Start time controls may not be present until expanded.
    try {
        days := ["Mon ", "Tue ", "Wed ", "Thu ", "Fri ", "Sat ", "Sun "]
        for _, d in days {
            rangeBtn := root.FindFirst({ Type: 50000, Name: d, matchmode: "Substring" })
            if rangeBtn {
                name := ""
                try name := rangeBtn.Name
                if (name != "" && InStr(name, " - ") && (InStr(name, " AM") || InStr(name, " PM"))) {
                    try rangeBtn.SetFocus()
                    Sleep 40
                    try rangeBtn.Click()
                    break
                }
            }
        }
    } catch {
    }

    ; Anchor-based fallback: focus a stable neighbor, Tab to the dynamic "Wed …" button, then Enter.
    ; In the captured tree, "Response options" immediately precedes the date/time range button.
    try {
        anchor := root.FindFirst({ AutomationId: "menur1qn" }) ; "Response options"
        if !anchor
            anchor := root.FindFirst({ Name: "Response options", ControlType: "Button" })
        if anchor {
            try anchor.SetFocus()
            Sleep 40
            Send "{Tab}"
            Sleep 40
            Send "{Enter}"
        }
    } catch {
    }

    ; Strategy A (advanced): sibling traversal from stable "Open scheduler" button to locate the dynamic date-range button.
    try {
        schedulerBtn := root.WaitElement({ Name: "Open scheduler", Type: 50000 }, 600)
        if schedulerBtn {
            dateBtn := ""
            try {
                ; Walk backwards among siblings until a button that looks like a time range is found.
                walker := UIA.RawViewWalker
                sib := walker.TryGetPreviousSiblingElement(schedulerBtn)
                tries := 0
                while (sib && tries < 8) {
                    tries += 1
                    n := ""
                    t := ""
                    try n := sib.Name
                    try t := sib.Type
                    if (t = UIA.Type.Button && n != "" && InStr(n, " - ") && (InStr(n, " AM") || InStr(n, " PM"))) {
                        dateBtn := sib
                        break
                    }
                    sib := walker.TryGetPreviousSiblingElement(sib)
                }
            } catch {
            }

            if dateBtn {
                try dateBtn.SetFocus()
                try dateBtn.Invoke()
                catch {
                    try dateBtn.Click()
                }
                ; If the popover opens, stop here (avoid toggling it closed with later strategies).
                Sleep 60
                popNow := Appt_FindOpenPopover()
                if popNow {
                    return popNow
                }
            }
        }
    } catch {
    }

    ; Strategy B (advanced): regex match button by time range (works even if day/date varies).
    try {
        ; Match e.g. "Wed 4/1/2026 2:00 PM - 2:30 PM" or "2:00 PM - 2:30 PM".
        re := "\d{1,2}:\d{2}\s*(AM|PM)?\s*-\s*\d{1,2}:\d{2}\s*(AM|PM)?"
        el := root.WaitElement({ Type: 50000, Name: re, matchmode: "RegEx" }, 600)
        if el {
            try el.SetFocus()
            try el.Invoke()
            catch {
                try el.Click()
            }
            Sleep 60
            popNow := Appt_FindOpenPopover()
            if popNow {
                return popNow
            }
        }
    } catch {
    }

    deadline := A_TickCount + 1200
    while (A_TickCount < deadline) {
        pop := Appt_FindOpenPopover()
        if pop
            return pop
        Sleep 60
    }
    return 0
}

Appt_PopoverFocusFirst(criteriaList) {
    pop := Appt_OpenPopoverIfNeeded()
    if !pop
        return false
    for crit in criteriaList {
        try {
            el := pop.FindFirst(crit)
            if el {
                el.SetFocus()
                return true
            }
        } catch {
        }
    }
    return false
}

Appt_PopoverInvokeFirst(criteriaList) {
    pop := Appt_OpenPopoverIfNeeded()
    if !pop {
        return false
    }
    for crit in criteriaList {
        try {
            el := pop.FindFirst(crit)
            if el {
                ok := false
                try {
                    el.Click()
                    ok := true
                } catch as err1 {
                    try {
                        el.Invoke()
                        ok := true
                    } catch as err2 {
                    }
                }
                if ok
                    return true
            }
        } catch {
        }
    }
    return false
}

Appt_PopoverToggleFirst(criteriaList) {
    pop := Appt_OpenPopoverIfNeeded()
    if !pop {
        return false
    }
    for crit in criteriaList {
        try {
            el := pop.FindFirst(crit)
            if !el
                continue
            tog := 0
            try tog := el.IsTogglePatternAvailable
            if tog {
                try {
                    el.TogglePattern.Toggle()
                    return true
                } catch as errT {
                }
            }

            ; Fallback: click/invoke if TogglePattern not available.
            try {
                el.Click()
                return true
            } catch as errC {
                try {
                    el.Invoke()
                    return true
                } catch as errI {
                }
            }
        } catch {
        }
    }
    return false
}

Appt_PopoverSelectTimeSuggestion(idx) {
    pop := Appt_OpenPopoverIfNeeded()
    if !pop
        return false
    try {
        list := pop.FindFirst({ Name: "Time suggestions", ControlType: "List" })
        if !list
            list := pop.FindFirst({ Name: "Time suggestions", Type: 50008 })
        if !list
            return false
        items := ""
        try items := list.FindAll({ ControlType: "ListItem" })
        catch {
            try items := list.FindAll({ Type: 50007 })
        }
        if (!IsObject(items) || items.Length < idx)
            return false
        li := items[idx]
        try li.Click()
        catch {
            try li.Invoke()
        }
        return true
    } catch {
        return false
    }
}

Appt_FocusBodyField_NewOutlook() {
    root := Appt_GetRootActive()
    if !root
        return false

    ; Prefer common body placeholders (best effort).
    needles := ["Add details", "Add description", "Description", "Message", "Details"]
    for n in needles {
        try {
            el := root.FindFirst({ Name: n, ControlType: "Edit" })
            if el {
                el.SetFocus()
                return true
            }
        } catch {
        }
        try {
            el := root.FindFirst({ Name: n, Type: 50004 })
            if el {
                el.SetFocus()
                return true
            }
        } catch {
        }
    }

    ; Fallback: pick the largest Edit/Document region and focus it.
    best := 0
    bestArea := 0
    candidates := []
    try candidates := root.FindAll({ ControlType: "Edit" })
    catch {
        candidates := []
    }
    if (!IsObject(candidates) || candidates.Length = 0) {
        try candidates := root.FindAll({ ControlType: "Document" })
        catch {
            candidates := []
        }
    }
    for c in candidates {
        try {
            rect := c.BoundingRectangle
            ; UIA-v2 typically returns {l,t,r,b} or an array-like; handle both.
            l := rect.l, t := rect.t, r := rect.r, b := rect.b
            area := Abs((r - l) * (b - t))
            if (area > bestArea) {
                bestArea := area
                best := c
            }
        } catch {
        }
    }
    if best {
        try best.SetFocus()
        return true
    }
    return false
}

; -------------------------------------------------------------------
; New Outlook Appointment command bar + selection modals
; -------------------------------------------------------------------
global g_ApptPickKey := ""

Appt_PickKey(key) {
    global g_ApptPickKey
    g_ApptPickKey := key
    try StandardLoadingBar_CloseKeysOverlay()
    try StandardLoadingBar_Hide(0)
}

Appt_PickTimeout() {
    Appt_PickKey("TIMEOUT")
}

Appt_SelectFromModal(title, options, promptKeys := "[1-9] Select  [Esc] Cancel", timeoutMs := 45000) {
    global g_ApptPickKey
    g_ApptPickKey := ""

    ; Ensure any loading indicator is cleared before showing interactive modal.
    try StandardLoadingBar_Hide(0)
    catch {
    }

    keyCallbacks := Map()
    msg := "❓ " title ":`n`n"
    loop options.Length {
        i := A_Index
        opt := options[i]
        k := opt.k
        label := opt.label
        msg .= k ") " label "`n"
        keyCallbacks.Set(k, Appt_PickKey.Bind(k))
    }
    keyCallbacks.Set("Escape", Appt_PickKey.Bind("ESC"))

    StandardLoadingBar_CloseKeysOverlay()
    StandardLoadingBar_ShowWithKeys(
        msg,
        keyCallbacks,
        timeoutMs,
        0,
        Appt_PickTimeout,
        "1E1E2E",
        760,
        17,
        BANNER_ACCENT_INTERMEDIATE,
        false,
        promptKeys,
        true
    )

    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        if (g_ApptPickKey != "")
            break
        Sleep 40
    }

    picked := g_ApptPickKey
    if (picked = "" || picked = "ESC" || picked = "TIMEOUT")
        return ""
    return picked
}

Appt_FindCommandBar() {
    root := Appt_GetRootActive()
    if !root
        return 0
    try {
        tb := root.FindFirst({ Name: "Event form commands", ControlType: "ToolBar" })
        if tb
            return tb
    } catch {
    }
    try {
        tb := root.FindFirst({ ControlType: "ToolBar" })
        if tb {
            ; Prefer a toolbar that contains Save or Send (meeting invites often show Send only).
            if tb.FindFirst({ Name: "Save", ControlType: "Button" }) || tb.FindFirst({ Name: "Send", ControlType: "Button" })
                return tb
        }
    } catch {
    }
    return 0
}

Appt_ClickInCommandBar(criteriaList) {
    tb := Appt_FindCommandBar()
    if !tb
        return false
    for crit in criteriaList {
        try {
            el := tb.FindFirst(crit)
            if el {
                try el.Click()
                catch {
                    try el.Invoke()
                }
                return true
            }
        } catch {
        }
    }
    return false
}

Appt_ClickAny(criteriaList) {
    root := Appt_GetRootActive()
    if !root
        return false
    for crit in criteriaList {
        try {
            el := root.FindFirst(crit)
            if el {
                try el.Click()
                catch {
                    try el.Invoke()
                }
                return true
            }
        } catch {
        }
    }
    return false
}

Appt_OpenMenuAndPick(menuButtonCriteriaList, menuItemName, preClickDelayMs := 0) {
    global APPT_MENU_OPEN_SETTLE_MS, APPT_MENU_ITEM_PRECLICK_MS
    if !IsSet(APPT_MENU_OPEN_SETTLE_MS)
        APPT_MENU_OPEN_SETTLE_MS := 520
    if !IsSet(APPT_MENU_ITEM_PRECLICK_MS)
        APPT_MENU_ITEM_PRECLICK_MS := 620

    ; Open menu (button), then pick the menu item.
    ; IMPORTANT: Searching the desktop root can be extremely expensive and can freeze the PC.
    if !Appt_ClickAny(menuButtonCriteriaList)
        return false
    if (preClickDelayMs <= 0)
        preClickDelayMs := APPT_MENU_ITEM_PRECLICK_MS
    try StandardLoadingBar_Update("🔄 Appointment: opening status menu…", BANNER_ACCENT_INTERMEDIATE)
    if (APPT_MENU_OPEN_SETTLE_MS > 0)
        Sleep APPT_MENU_OPEN_SETTLE_MS
    try {
        rootWin := Appt_GetRootActive()
        if !rootWin
            return false
        try StandardLoadingBar_Update("🔄 Appointment: selecting " menuItemName "…", BANNER_ACCENT_INTERMEDIATE)

        ; Try within the active appointment window first (fast).
        mi := 0
        try mi := rootWin.FindFirst({ Name: menuItemName, ControlType: "MenuItem" })
        catch {
        }
        if !mi {
            try mi := rootWin.FindFirst({ Name: menuItemName, ControlType: "RadioButton" })
            catch {
            }
        }
        if !mi {
            try mi := rootWin.FindFirst({ Name: menuItemName, ControlType: "Button" })
            catch {
            }
        }
        if !mi {
            try mi := rootWin.FindFirst({ Name: menuItemName, ControlType: "ListItem" })
            catch {
            }
        }
        if !mi {
            try mi := rootWin.FindFirst({ Name: menuItemName, ControlType: "CheckBox" })
            catch {
            }
        }
        if !mi {
            try mi := rootWin.FindFirst({ Name: menuItemName })
            catch {
            }
        }

        ; Fallback: ONE desktop attempt (avoid freezing the machine).
        if !mi {
            try {
                desktop := UIA.GetRootElement()
                if desktop {
                    try mi := desktop.FindFirst({ Name: menuItemName })
                }
            } catch {
            }
        }
        if !mi {
            ; Last resort: Portuguese category names may have extra state text; allow substring match.
            try {
                desktop := UIA.GetRootElement()
                if desktop {
                    try mi := desktop.FindFirst({ Name: menuItemName, matchmode: "Substring" })
                }
            } catch {
            }
        }
        if mi {
            if (preClickDelayMs > 0) {
                try StandardLoadingBar_Update("👁️ Appointment: about to click → " menuItemName,
                    BANNER_ACCENT_INTERMEDIATE)
                catch {
                }
                Sleep preClickDelayMs
            }
            try mi.Click()
            catch {
                try mi.Invoke()
            }
            StandardLoadingBar_Update("✅ Appointment: " menuItemName, BANNER_ACCENT_SUCCESS)
            return true
        }
    } catch as err {
    }
    return false
}
