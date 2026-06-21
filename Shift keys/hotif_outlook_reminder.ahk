; =============================================================================
; Shift keys module: hotif_outlook_reminder.ahk
; Outlook reminder/appointment hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

#HotIf (WinActive("ahk_exe OUTLOOK.EXE") || WinActive("ahk_exe olk.exe")) && RegExMatch(WinGetTitle("A"),
"i)Reminders?") && !IsFileDialogActive()

; ativa a janela de lembretes do Outlook
ActivateReminder() {
    if (!WinExist("ahk_exe OUTLOOK.EXE")) {
        ShowCenteredOverlay_Utils("❌ Error: Target window not found.", 2000, BANNER_ACCENT_ERROR)
        return
    }
    WinActivate("ahk_exe OUTLOOK.EXE")
    WinWaitActive("ahk_exe OUTLOOK.EXE", , 1)
}

; digita o tempo e aperta Alt+S
QuickSnooze(t) {
    ActivateReminder()
    Send("{Tab}")              ; chega ao combo
    Send("{Tab}")              ; chega ao combo
    Sleep 100
    Send("^a{Delete}" . t)     ; substitui o texto
    Sleep 120
    Send("!s")                 ; Alt+S = Snooze
    Sleep 200
    Send("{Tab}")
    Send("{Tab}")
    Send("{Tab}")
}

; caixa de confirmaÃ§Ã£o antes de executar
Confirm(t) {
    if MsgBox("Snooze for " t "?", "Confirm Snooze", "YesNo Icon?") = "Yes"
        QuickSnooze(t)
}

; ---------------------------------------------------------------------------
; New Outlook Reminders (keyboard-only)
; - Uses UIA list extraction + Standard Information Display selection banner
; - Executes item actions via Apps/Menu key + arrow navigation (per screenshots)
; ---------------------------------------------------------------------------
global g_RemindersPickKey := ""
global g_DebugBe11ecLogPath := "C:\Users\fie7ca\Documents\scripts\debug-be11ec.log"
global g_RemindersDebugEnabled := false
global g_RemindersTimingProfile := "very_slow"
global g_RemindersTimingProfiles := Map(
    "normal", Map(
        "context_focus_ms", 80,
        "focus_to_menu_ms", 25,
        "menu_render_ms", 45,
        "menu_home_ms", 60,
        "menu_scan_step_ms", 50,
        "submenu_open_ms", 80,
        "post_select_ms", 60,
        "modal_to_row_ms", 35
    ),
    "slow", Map(
        "context_focus_ms", 120,
        "focus_to_menu_ms", 70,
        "menu_render_ms", 170,
        "menu_home_ms", 100,
        "menu_scan_step_ms", 140,
        "submenu_open_ms", 200,
        "post_select_ms", 110,
        "modal_to_row_ms", 80
    ),
    "very_slow", Map(
        "context_focus_ms", 260,
        "focus_to_menu_ms", 180,
        "menu_render_ms", 420,
        "menu_home_ms", 220,
        "menu_scan_step_ms", 360,
        "submenu_open_ms", 520,
        "post_select_ms", 260,
        "modal_to_row_ms", 220
    )
)

Reminders_GetTimingProfile() {
    global g_RemindersTimingProfile, g_RemindersTimingProfiles
    p := "very_slow"
    try p := StrLower(Trim(g_RemindersTimingProfile))
    if !g_RemindersTimingProfiles.Has(p)
        p := "very_slow"
    return g_RemindersTimingProfiles[p]
}

Reminders_DelayValue(key, fallbackMs := 80) {
    try {
        profile := Reminders_GetTimingProfile()
        if profile.Has(key)
            return profile[key]
    } catch {
    }
    return fallbackMs
}

Reminders_LoadingShow(text) {
    try {
        StandardLoadingBar_Show(text, BANNER_ACCENT_INTERMEDIATE, { passive: false, centerOnHwnd: 0, textWidth: 560,
            fontSize: 17 })
    } catch {
    }
}

Reminders_LoadingHide(delayMs := 0) {
    try StandardLoadingBar_Hide(delayMs)
    catch {
    }
}

Reminders_DebugLog(location, message, data := "", hypothesisId := "A", runId := "pre-fix") {
    if !g_RemindersDebugEnabled
        return
    try {
        Debug_Escape(s) => StrReplace(StrReplace(StrReplace(String(s), "\", "\\"), "`"", "\`""), "`n", "\n")
        Debug_MapToJson(m) {
            out := ""
            for k, v in m {
                if (out != "")
                    out .= ","
                out .= "`"" Debug_Escape(k) "`":"
                if (v is Integer || v is Float)
                    out .= String(v)
                else
                    out .= "`"" Debug_Escape(v) "`""
            }
            return "{" out "}"
        }

        payload := Map()
        payload["sessionId"] := "be11ec"
        payload["id"] := "log_" A_TickCount "_" Random(1000, 9999)
        payload["timestamp"] := A_TickCount
        payload["location"] := location
        payload["message"] := message
        payload["runId"] := runId
        payload["hypothesisId"] := hypothesisId

        d := IsObject(data) ? data : Map("value", data)
        ; Encode only scalar data safely (stringify non-numeric as strings)
        payloadJson := Debug_MapToJson(payload)
        dataJson := Debug_MapToJson(d)
        line := SubStr(payloadJson, 1, StrLen(payloadJson) - 1) . ",`"data`":" . dataJson . "}"
        global g_DebugBe11ecLogPath
        FileAppend(line "`n", g_DebugBe11ecLogPath, "UTF-8")
    } catch {
    }
}

; #region agent log (session 6dacac — NDJSON to debug-6dacac.log; remove after verification)
Debug6dacac_Log(location, message, data := "", hypothesisId := "H1", runId := "post-fix") {
    if !g_RemindersDebugEnabled
        return
    try {
        Esc(s) => StrReplace(StrReplace(StrReplace(String(s), "\", "\\"), "`"", "\`""), "`n", "\n")
        MapToJson(m) {
            out := ""
            for k, v in m {
                if (out != "")
                    out .= ","
                out .= "`"" Esc(k) "`":"
                if (v is Integer || v is Float)
                    out .= String(v)
                else
                    out .= "`"" Esc(v) "`""
            }
            return "{" out "}"
        }
        d := IsObject(data) ? data : Map("value", data)
        id := "log_" A_TickCount "_" Random(1000, 9999)
        ts := A_TickCount
        payload := Map()
        payload["sessionId"] := "6dacac"
        payload["id"] := id
        payload["timestamp"] := ts
        payload["location"] := location
        payload["message"] := message
        payload["runId"] := runId
        payload["hypothesisId"] := hypothesisId
        payloadJson := MapToJson(payload)
        dataJson := MapToJson(d)
        line := SubStr(payloadJson, 1, StrLen(payloadJson) - 1) . ",`"data`":" . dataJson . "}`n"
        FileAppend(line, A_ScriptDir "\debug-6dacac.log", "UTF-8")
    } catch {
    }
}

; #endregion

; Diagnostic-only raw UIA dump for reminder window.
; Captures candidate counts and samples by control type so we can patch selectors based on facts.
Reminders_DebugDumpWindowTree(hwnd, tag := "") {
    if !g_RemindersDebugEnabled
        return
    if !hwnd
        return
    try {
        rootSource := ""
        root := Reminders_RootElementForHwnd(hwnd, &rootSource)
        if !root {
            try Reminders_DebugLog("Shift keys.ahk:Reminders_DebugDumpWindowTree", "Root not found", Map(
                "hwnd", hwnd,
                "tag", tag,
                "rootSource", rootSource
            ), "TREE0", "pre-fix")
            return
        }

        DumpByType(typeName, hyp) {
            arr := 0
            try arr := root.FindAll({ ControlType: typeName })
            cnt := arr ? arr.Length : 0
            sample := ""
            if arr {
                maxSample := Min(25, arr.Length)
                loop maxSample {
                    i := A_Index
                    n := ""
                    try n := arr[i].Name
                    if (n = "")
                        continue
                    if (sample != "")
                        sample .= " | "
                    sample .= n
                }
            }
            try Reminders_DebugLog("Shift keys.ahk:Reminders_DebugDumpWindowTree", "Tree slice", Map(
                "hwnd", hwnd,
                "tag", tag,
                "rootSource", rootSource,
                "type", typeName,
                "count", cnt,
                "sample", sample
            ), hyp, "pre-fix")
        }

        ; Types most likely to contain New Outlook reminder rows.
        DumpByType("Group", "TREE-G")
        DumpByType("Button", "TREE-B")
        DumpByType("List", "TREE-L")
        DumpByType("ListItem", "TREE-LI")
        DumpByType("DataItem", "TREE-DI")
        DumpByType("Text", "TREE-T")
        DumpByType("Hyperlink", "TREE-H")
    } catch {
    }
}

Reminders_IsNewOutlookWindow() {
    try {
        ; Reminders window can run under classic OUTLOOK.EXE or Store olk.exe.
        if !(WinActive("ahk_exe OUTLOOK.EXE") || WinActive("ahk_exe olk.exe"))
            return false
        t := WinGetTitle("A")
        ; NOTE: single backslash in regex. Using \\b would match literal "\b".
        ok := RegExMatch(t, "i)\bReminders\b")
        ; #region agent log
        try Reminders_DebugLog("Shift keys.ahk:Reminders_IsNewOutlookWindow", "Computed isNewReminders", Map(
            "ok", ok,
            "title", t
        ), "H1", "pre-fix")
        ; #endregion
        return ok
    } catch {
        return false
    }
}

Reminders_ItemsListSignature(items, maxLabels := 0) {
    sig := String(items.Length)
    n := (maxLabels > 0) ? Min(items.Length, maxLabels) : items.Length
    loop n {
        sig .= "`n" items[A_Index].label
    }
    return sig
}

; Tiny move + restore so Outlook refreshes the accessible tree (same effect as manually moving the window).
; Skip when maximized — WinMove is unreliable; user can restore the window first if needed.
Reminders_NudgeWindowForUiRefresh(hwnd) {
    if !hwnd || !WinExist("ahk_id " hwnd)
        return
    try {
        if (WinGetMinMax("ahk_id " hwnd) = 1)
            return
        WinGetPos(&x, &y, &w, &h, "ahk_id " hwnd)
        WinMove(x + 1, y, w, h, "ahk_id " hwnd)
        Sleep 20  ; brief beat so the shell/UIA sees the move before restore
        WinMove(x, y, w, h, "ahk_id " hwnd)
    } catch {
    }
}

; Two consecutive identical UIA snapshots (or maxPasses) — reduces races while Outlook refreshes the list.
; NOTE: If we get two consecutive EMPTY matches on early passes, keep trying instead of giving up,
; since a newly-visible window might just have UIA that hasn't rendered reminders yet.
Reminders_GetItemsStable(targetHwnd, delayMs := 60, maxPasses := 3) {
    lastSig := ""
    lastItems := []
    loop maxPasses {
        cur := Reminders_GetItems(targetHwnd)
        sig := Reminders_ItemsListSignature(cur, 0)
        if (lastSig != "" && sig = lastSig) {
            ; Don't accept empty-empty match as stable on early passes (pass 1-2); window might still be rendering.
            ; Only consider pass 3+ empty match as "stable" (window legitimately empty).
            emptyMatch := (cur.Length = 0 && lastItems.Length = 0)
            if (!emptyMatch || A_Index >= maxPasses) {
                try Reminders_DebugLog("Shift keys.ahk:Reminders_GetItemsStable", "Stable snapshot matched", Map(
                    "pass", A_Index,
                    "count", cur.Length,
                    "maxPasses", maxPasses,
                    "emptyMatch", emptyMatch ? 1 : 0,
                    "remHwnd", targetHwnd
                ), "ST1", "pre-fix")
                return cur
            } else {
                try Reminders_DebugLog("Shift keys.ahk:Reminders_GetItemsStable",
                    "Empty-empty match ignored, continuing retry", Map(
                        "pass", A_Index,
                        "maxPasses", maxPasses,
                        "remHwnd", targetHwnd
                    ), "ST-SKIP", "pre-fix")
            }
        }
        lastSig := sig
        lastItems := cur
        Sleep delayMs
    }
    try Reminders_DebugLog("Shift keys.ahk:Reminders_GetItemsStable", "Stable snapshot max passes; using last read",
        Map(
            "count", lastItems.Length,
            "remHwnd", targetHwnd
        ), "ST2", "pre-fix")
    return lastItems
}

Reminders_RootElementForHwnd(hwnd, &source := "") {
    source := ""
    if !hwnd
        return ""
    ; New Outlook surfaces often live under WebView2/Chromium subtree.
    try {
        el := UIA.ElementFromChromium("ahk_id " hwnd, 500)
        if el {
            source := "chromium"
            return el
        }
    } catch {
    }
    try {
        el := UIA.ElementFromHandle(hwnd)
        if el {
            source := "handle"
            return el
        }
    } catch {
    }
    return ""
}

Reminders_GetItems(targetHwnd := 0) {
    items := []
    hwnd := targetHwnd ? targetHwnd : WinExist("A")
    if !hwnd
        return items
    try {
        rootSource := ""
        root := Reminders_RootElementForHwnd(hwnd, &rootSource)
        if !root {
            try Reminders_DebugLog("Shift keys.ahk:Reminders_GetItems", "ERROR: Root element null", Map(
                "hwnd", hwnd,
                "rootSource", rootSource
            ), "X0-ERR", "pre-fix")
            return items
        }

        CollectFromCandidates(cands) {
            out := []
            dropped := 0
            dropSample := ""
            for b in cands {
                n := ""
                try n := b.Name
                if (n = "")
                    continue
                ; Exclude global/window controls
                if (n = "Settings" || n = "Dismiss all" || n = "Dismiss All" || n = "Minimize" || n = "Maximize" || n =
                    "Close")
                    continue
                ; Exclude action/menu-like items that can appear while context UI is open
                if RegExMatch(n, "i)^(Snooze reminder|Dismiss reminder|Join Teams meeting|Chat with participants)$")
                    continue

                ; Keep only actual reminder rows. In our UIA tree, these names include time/all-day + a relative marker.
                ; Examples: "Stretch All day Today", "CIM Journey 3:00 PM Microsoft Teams Meeting 18 hrs ago"
                ; Relative-age tokens vary (e.g. "1 hour ago", "7 days", "4 wks ago", "Today").
                ; NOTE: single backslash in regex. Using \\b would match literal "\b".
                hasTime := RegExMatch(n, "i)(\bAll day\b|\bAM\b|\bPM\b)")
                isAllDay := RegExMatch(n, "i)\bAll day\b")
                hasRel := RegExMatch(n,
                    "i)\b(Today|\d+\s*(min|mins|minute|minutes|hr|hrs|hour|hours|day|days|wk|wks|week|weeks)\b(\s+ago)?)\b"
                )
                hasClockTime := RegExMatch(n, "i)\b\d{1,2}:\d{2}\b")
                hasEventCue := RegExMatch(n,
                    "i)\b(Meeting|Teams|Appointment|Event|Call|Invite|Reminder)\b")
                ; New reminders can be "in 15 minutes" without AM/PM; keep broader candidates
                ; while still excluding known command/window controls above.
                isLikelyRow := ((hasTime || hasClockTime || hasRel || hasEventCue) && !RegExMatch(n,
                    "i)^(Snooze reminder|Dismiss reminder|Join Teams meeting|Chat with participants)$"))
                if !isLikelyRow {
                    dropped++
                    if (dropSample != "" && dropped <= 8)
                        dropSample .= " | "
                    if (dropped <= 8)
                        dropSample .= n
                    continue
                }
                out.Push({ el: b, label: n })
            }
            return Map("items", out, "dropped", dropped, "dropSample", dropSample, "totalCandidates", cands ? cands.Length :
                0)
        }

        ; Primary anchor: reminder summary group text varies ("There are ..." / "There is only ...").
        listGroup := root.FindFirst({ ControlType: "Group", Name: "reminder", matchmode: "Substring" })
        if !listGroup
            listGroup := root.FindFirst({ ControlType: "Group", Name: "There ", matchmode: "Substring" })
        if !listGroup
            listGroup := root.FindFirst({ Name: "reminder", matchmode: "Substring" })

        if !listGroup {
            try Reminders_DebugLog("Shift keys.ahk:Reminders_GetItems",
                "WARNING: 'There are X reminders' group not found in UIA", Map(
                    "hwnd", hwnd,
                    "rootSource", rootSource
                ), "X0-WRN", "pre-fix")
        }

        ; Always re-evaluate source: a short-lived cache reused "listGroup" without re-checking
        ; listSmallerThanRoot and could omit reminders (same class of bug as scanning only the group).
        listBtns := 0
        listBtnLen := 0
        if listGroup {
            try listBtns := listGroup.FindAll({ ControlType: "Button" })
            listBtnLen := listBtns ? listBtns.Length : 0
        }
        rootBtns := root.FindAll({ ControlType: "Button" })
        rootBtnLen := rootBtns ? rootBtns.Length : 0
        listSmallerThanRoot := (listGroup && listBtnLen > 0 && rootBtnLen > listBtnLen)

        ; Diagnostic logging for button search
        try Reminders_DebugLog("Shift keys.ahk:Reminders_GetItems", "Button search results", Map(
            "rootBtnLen", rootBtnLen,
            "listBtnLen", listBtnLen,
            "listGroupFound", listGroup ? 1 : 0,
            "hwnd", hwnd,
            "rootSource", rootSource
        ), "GETITEMS-BTN", "pre-fix")

        btns := 0
        source := ""
        if (listGroup && listBtnLen > 0 && !listSmallerThanRoot) {
            btns := listBtns
            source := "listGroup"
        } else {
            btns := rootBtns
            source := "root"
        }

        ; Debug: log if btns is empty before attempting to collect
        if (!btns || btns.Length = 0) {
            try Reminders_DebugLog("Shift keys.ahk:Reminders_GetItems", "No buttons found to process", Map(
                "source", source,
                "rootBtnLen", rootBtnLen,
                "listBtnLen", listBtnLen,
                "hwnd", hwnd,
                "rootSource", rootSource
            ), "GETITEMS-NOBTNS", "pre-fix")
        }

        r := CollectFromCandidates(btns)
        items := r["items"]
        dropped := r["dropped"]
        dropSample := r["dropSample"]

        ; Fallback for New Outlook builds where reminder rows are exposed as ListItem/DataItem, not Button.
        if (items.Length = 0) {
            listItems := 0
            dataItems := 0
            try listItems := root.FindAll({ ControlType: "ListItem" })
            try dataItems := root.FindAll({ ControlType: "DataItem" })
            liLen := listItems ? listItems.Length : 0
            diLen := dataItems ? dataItems.Length : 0
            try Reminders_DebugLog("Shift keys.ahk:Reminders_GetItems", "Fallback candidate search", Map(
                "listItemLen", liLen,
                "dataItemLen", diLen,
                "hwnd", hwnd,
                "rootSource", rootSource
            ), "GETITEMS-FB", "pre-fix")

            if (liLen > 0) {
                r := CollectFromCandidates(listItems)
                if (r["items"].Length > 0) {
                    items := r["items"]
                    dropped := r["dropped"]
                    dropSample := r["dropSample"]
                    source := "fallback-listitem"
                }
            }
            if (items.Length = 0 && diLen > 0) {
                r := CollectFromCandidates(dataItems)
                if (r["items"].Length > 0) {
                    items := r["items"]
                    dropped := r["dropped"]
                    dropSample := r["dropSample"]
                    source := "fallback-dataitem"
                }
            }
        }
        ; #region agent log (session 6dacac — H1 source + row count)
        try {
            Debug6dacac_Log("Shift keys.ahk:Reminders_GetItems", "source chosen", Map(
                "source", source,
                "keptCount", items.Length,
                "rootBtnLen", rootBtnLen,
                "listBtnLen", listBtnLen,
                "listSmallerThanRoot", listSmallerThanRoot ? 1 : 0,
                "remHwnd", hwnd,
                "rootSource", rootSource
            ), "H1", "post-fix")
        } catch {
        }
        ; #endregion
        ; #region agent log
        try {
            sample := ""
            maxSample := Min(12, items.Length)
            loop maxSample {
                i := A_Index
                if (sample != "")
                    sample .= " | "
                sample .= items[i].label
            }
            Reminders_DebugLog("Shift keys.ahk:Reminders_GetItems", "Extracted reminders sample", Map(
                "count", items.Length,
                "sample", sample,
                "source", source,
                "totalCandidates", r["totalCandidates"],
                "dropped", dropped,
                "dropSample", dropSample,
                "hwnd", hwnd,
                "rootSource", rootSource
            ), "X1", "pre-fix")
        } catch {
        }
        ; #endregion
    } catch {
        return items
    }
    return items
}

Reminders_PickKey(key) {
    global g_RemindersPickKey
    ; #region agent log
    try Reminders_DebugLog("Shift keys.ahk:Reminders_PickKey", "PickKey called", Map(
        "key", key,
        "priorKey", A_PriorKey,
        "thisHotkey", A_ThisHotkey
    ), "P1", "pre-fix")
    ; #endregion

    ; Guard: ignore accidental selection when Windows key (or other modifiers) is involved.
    ; This prevents the selection modal from disappearing due to unrelated Win-key chords.
    try {
        if (GetKeyState("LWin", "P") || GetKeyState("RWin", "P")
        || GetKeyState("Ctrl", "P") || GetKeyState("Alt", "P")) {
            ; #region agent log
            try Reminders_DebugLog("Shift keys.ahk:Reminders_PickKey", "Ignored pick due to modifier down", Map(
                "key", key,
                "lwin", GetKeyState("LWin", "P"),
                "rwin", GetKeyState("RWin", "P"),
                "ctrl", GetKeyState("Ctrl", "P"),
                "alt", GetKeyState("Alt", "P")
            ), "P3", "pre-fix")
            ; #endregion
            return
        }
        if (A_PriorKey = "LWin" || A_PriorKey = "RWin") {
            ; #region agent log
            try Reminders_DebugLog("Shift keys.ahk:Reminders_PickKey", "Ignored pick due to priorKey Win", Map(
                "key", key,
                "priorKey", A_PriorKey
            ), "P4", "pre-fix")
            ; #endregion
            return
        }
    } catch {
    }

    g_RemindersPickKey := key
    try StandardLoadingBar_CloseKeysOverlay()
    try StandardLoadingBar_Hide(0)
}

Reminders_PickTimeout() {
    ; #region agent log
    try Reminders_DebugLog("Shift keys.ahk:Reminders_PickTimeout", "Timeout fired", Map(), "P2", "pre-fix")
    ; #endregion
    Reminders_PickKey("TIMEOUT")
}

Reminders_SelectItem(actionLabel, &items, remHwnd, maxItems := 35) {
    global g_RemindersPickKey
    g_RemindersPickKey := ""

    ; Workaround: one-pixel nudge refreshes UIA before enumeration (not repeated in the modal refresh loop).
    Reminders_NudgeWindowForUiRefresh(remHwnd)
    Sleep 100  ; longer wait for reminder window to stabilize UIA after nudge (highly volatile window)

    ; Volatile window: refresh once right before showing the modal so we don't start with a stale/partial snapshot.
    ; Use longer delays (100ms between passes) to wait for Outlook to stabilize the reminder list in UIA.
    try items := Reminders_GetItemsStable(remHwnd, 100, 5)

    if (items.Length = 0) {
        ; One more try with even longer wait (UIA can lag when reminders are being added rapidly).
        try items := Reminders_GetItemsStable(remHwnd, 150, 6)
        if (items.Length = 0) {
            ShowCenteredOverlay_Utils("❌ No reminders found", 1600, BANNER_ACCENT_ERROR)
            return 0
        }
    }

    ; Stable key set (we'll keep callbacks stable and just refresh the displayed list).
    keys := []
    loop 9
        keys.Push(A_Index)
    for c in StrSplit("ABCDEFGHIJKLMNOPQRSTUVWXYZ")
        keys.Push(c)

    ClampCount(itemsLen) {
        c := itemsLen
        if (c > maxItems)
            c := maxItems
        if (c > keys.Length)
            c := keys.Length
        return c
    }

    BuildMsg(currentItems, currentCount) {
        ; Information-only copy inside the interactive banner (see docs/standard_information_display.md).
        m :=
            "ℹ️ The script nudges this window by one pixel before listing (Outlook UIA quirk). If a row is still missing, focus or move the window, then use the shortcut again.`n`n"
        m .= "❓ Select reminder to " actionLabel ":`n`n"
        loop currentCount {
            i := A_Index
            k := keys[i]
            label := currentItems[i].label
            m .= k ")`n    " label "`n`n"
        }
        if (currentItems.Length > currentCount)
            m .= "`n⚠ Showing first " currentCount " of " currentItems.Length " reminders"
        return m
    }

    ; Derive count using integer-only operations (avoid Float issues in ranges/loops).
    count := ClampCount(items.Length)
    msg := BuildMsg(items, count)

    keyCallbacks := Map()
    ; Register all potential selection keys once (prevents needing to rebuild callbacks on refresh).
    maxKeyCount := maxItems
    if (maxKeyCount > keys.Length)
        maxKeyCount := keys.Length
    loop maxKeyCount {
        i := A_Index
        k := keys[i]
        keyCallbacks.Set(k, Reminders_PickKey.Bind(k))
    }
    keyCallbacks.Set("Escape", Reminders_PickKey.Bind("ESC"))

    ; #region agent log
    try Reminders_DebugLog("Shift keys.ahk:Reminders_SelectItem", "About to show selection modal", Map(
        "actionLabel", actionLabel,
        "count", count,
        "priorKey", A_PriorKey,
        "thisHotkey", A_ThisHotkey,
        "remHwnd", remHwnd
    ), "S1", "pre-fix")
    ; #endregion

    ; Prevent immediate auto-selection / ignored picks caused by modifiers still being held from the trigger hotkey.
    ; If the trigger uses Win/Ctrl/Alt/Shift, wait for release before listening for selection keys.
    try {
        if (A_ThisHotkey != "") {
            th := A_ThisHotkey
            if InStr(th, "#") {
                try KeyWait "LWin"
                try KeyWait "RWin"
            }
            if InStr(th, "^")
                try KeyWait "Ctrl"
            if InStr(th, "!")
                try KeyWait "Alt"
            if InStr(th, "+")
                try KeyWait "Shift"

            hk := th
            hk := StrReplace(hk, "+", "")
            hk := StrReplace(hk, "^", "")
            hk := StrReplace(hk, "!", "")
            hk := StrReplace(hk, "#", "")
            if (StrLen(hk) = 1)
                KeyWait hk
        }
    } catch {
    }

    StandardLoadingBar_CloseKeysOverlay()
    ShowModal() {
        if (!IsBlackoutSuppressed()) {
            if (IsObject(keyCallbacks))
                keyCallbacks["D"] := DisableBlackout7Min
            StandardLoadingBar_ShowWithKeys(
                msg,
                keyCallbacks,
                45000,
                0,
                Reminders_PickTimeout,
                "1E1E2E",
                760,
                14,
                BANNER_ACCENT_INTERMEDIATE,
                false,
                "[1-9/A-Z] Select  [Esc] Cancel",
                true
            )
        }
    }
    lastSig := Reminders_ItemsListSignature(items, maxItems)
    ShowModal()
    try {
        latest := Reminders_GetItems(remHwnd)
        if (latest.Length = 0) {
            ; If empty immediately after show, use more aggressive retry
            latest := Reminders_GetItemsStable(remHwnd, 100, 4)
        }
        latestCount := ClampCount(latest.Length)
        sig := Reminders_ItemsListSignature(latest, maxItems)
        if (sig != lastSig) {
            lastSig := sig
            items := latest
            count := latestCount
            msg := BuildMsg(items, count)
            try StandardLoadingBar_Update(msg)
            ; #region agent log
            try Reminders_DebugLog("Shift keys.ahk:Reminders_SelectItem", "Post-show reminder sync", Map(
                "itemsCount", items.Length,
                "showing", count,
                "remHwnd", remHwnd
            ), "SR0", "pre-fix")
            ; #endregion
        }
    } catch {
    }
    ; #region agent log
    try Reminders_DebugLog("Shift keys.ahk:Reminders_SelectItem", "Selection modal shown", Map(
        "count", count,
        "remHwnd", remHwnd
    ), "S3", "pre-fix")
    ; #endregion

    deadline := A_TickCount + 45000
    reopens := 0
    lastRefreshTick := A_TickCount
    pollMs := 300
    while (A_TickCount < deadline) {
        if (g_RemindersPickKey != "")
            break

        ; Live refresh: if reminders change while the modal is open, update the displayed list.
        if (A_TickCount - lastRefreshTick >= pollMs) {
            lastRefreshTick := A_TickCount
            try {
                latest := Reminders_GetItems(remHwnd)
                if (latest.Length = 0) {
                    ; If empty, try with stability check and longer waits (volatile window)
                    latest := Reminders_GetItemsStable(remHwnd, 100, 4)
                }
                latestCount := ClampCount(latest.Length)
                sig := Reminders_ItemsListSignature(latest, maxItems)
                if (sig != lastSig) {
                    lastSig := sig
                    items := latest
                    count := latestCount
                    msg := BuildMsg(items, count)
                    try StandardLoadingBar_Update(msg)
                    ; #region agent log
                    try Reminders_DebugLog("Shift keys.ahk:Reminders_SelectItem", "Live-refreshed reminder list", Map(
                        "itemsCount", items.Length,
                        "showing", count,
                        "remHwnd", remHwnd
                    ), "SR1", "pre-fix")
                    ; #endregion
                }
            } catch {
            }
        }

        ; Detect unexpected overlay dismissal (another script/banner replaced it).
        try {
            global g_StandardLoadingBarIsKeysOverlay, g_StandardLoadingBarGui
            if (!g_StandardLoadingBarIsKeysOverlay || !IsObject(g_StandardLoadingBarGui)) {
                ; #region agent log
                try Reminders_DebugLog("Shift keys.ahk:Reminders_SelectItem",
                    "Selection modal disappeared unexpectedly", Map(
                        "isKeys", g_StandardLoadingBarIsKeysOverlay ? 1 : 0,
                        "hasGui", IsObject(g_StandardLoadingBarGui) ? 1 : 0,
                        "priorKey", A_PriorKey
                    ), "S4", "pre-fix")
                ; #endregion
                reopens++
                if (reopens >= 3) {
                    ; #region agent log
                    try Reminders_DebugLog("Shift keys.ahk:Reminders_SelectItem",
                        "Too many unexpected closes; giving up",
                        Map("reopens", reopens), "S6", "pre-fix")
                    ; #endregion
                    break
                }
                ; Re-show the modal; another overlay likely replaced it.
                ; #region agent log
                try Reminders_DebugLog("Shift keys.ahk:Reminders_SelectItem", "Reopening selection modal after close",
                    Map("reopens", reopens), "S5", "pre-fix")
                ; #endregion
                try StandardLoadingBar_CloseKeysOverlay()
                try StandardLoadingBar_Hide(0)
                Sleep 50
                ShowModal()
                try {
                    latest := Reminders_GetItems(remHwnd)
                    if (latest.Length = 0) {
                        ; If empty after reopen, use more aggressive retry
                        latest := Reminders_GetItemsStable(remHwnd, 100, 4)
                    }
                    latestCount := ClampCount(latest.Length)
                    sig := Reminders_ItemsListSignature(latest, maxItems)
                    if (sig != lastSig) {
                        lastSig := sig
                        items := latest
                        count := latestCount
                        msg := BuildMsg(items, count)
                        try StandardLoadingBar_Update(msg)
                        try Reminders_DebugLog("Shift keys.ahk:Reminders_SelectItem", "Post-reopen reminder sync", Map(
                            "itemsCount", items.Length,
                            "showing", count,
                            "remHwnd", remHwnd
                        ), "SR0b", "pre-fix")
                    }
                } catch {
                }
            }
        } catch {
        }
        Sleep 50
    }

    picked := g_RemindersPickKey
    ; #region agent log
    try Reminders_DebugLog("Shift keys.ahk:Reminders_SelectItem", "Selection modal resolved", Map(
        "picked", picked
    ), "S2", "pre-fix")
    ; #endregion
    if (picked = "" || picked = "ESC" || picked = "TIMEOUT")
        return 0

    ; Resolve key to index
    loop count {
        i := A_Index
        if (keys[i] = picked)
            return i
    }
    return 0
}

Reminders_OpenContextMenuForItem(itemEl) {
    remHwnd := WinExist("A")
    if remHwnd {
        WinActivate("ahk_id " remHwnd)
        WinWaitActive("ahk_id " remHwnd, , 1)
        Sleep 80  ; Give time for activation to settle
    }
    try itemEl.SetFocus()
    Sleep Reminders_DelayValue("context_focus_ms", 80)
    EnsureFocus()
    Sleep Reminders_DelayValue("focus_to_menu_ms", 25)  ; focus → menu open
    ; Apps/Menu key (keyboard-only)
    Send "{AppsKey}"
    Sleep Reminders_DelayValue("menu_render_ms", 70)  ; context menu rendered before Home/scan (dismiss / snooze / join online)
}

Reminders_MenuGetFocusedName() {
    try {
        fe := UIA.GetFocusedElement()
        if !fe
            return ""
        return fe.Name
    } catch {
        return ""
    }
}

Reminders_MenuFindItemContains(needle, maxSteps := 20, logId := "SN") {
    ; Sequentially navigate the context menu looking for an item whose focused text contains needle.
    ; Returns true when found (focus rests on that item).
    needle := StrLower(needle)
    Send "{Home}"
    Sleep Reminders_DelayValue("menu_home_ms", 60)
    loop maxSteps {
        name := Reminders_MenuGetFocusedName()
        shownName := (name != "") ? name : "(no focused text)"
        try StandardLoadingBar_Update("👁️ Menu scan " A_Index "/" maxSteps ": " shownName,
            BANNER_ACCENT_INTERMEDIATE)
        if (Mod(A_Index, 5) = 0) {
            ; #region agent log
            try Reminders_DebugLog("Shift keys.ahk:Reminders_MenuFindItemContains", "Menu scan step", Map(
                "logId", logId,
                "step", A_Index,
                "name", name
            ), "SN1", "pre-fix")
            ; #endregion
        }
        if (name != "" && InStr(StrLower(name), needle)) {
            try StandardLoadingBar_Update("✅ Menu selected: " name, BANNER_ACCENT_SUCCESS)
            ; #region agent log
            try Reminders_DebugLog("Shift keys.ahk:Reminders_MenuFindItemContains", "Menu item matched", Map(
                "logId", logId,
                "step", A_Index,
                "name", name,
                "needle", needle
            ), "SN2", "pre-fix")
            ; #endregion
            return true
        }
        Send "{Down}"
        Sleep Reminders_DelayValue("menu_scan_step_ms", 50)
    }
    ; #region agent log
    try Reminders_DebugLog("Shift keys.ahk:Reminders_MenuFindItemContains", "Menu item not found", Map(
        "logId", logId,
        "needle", needle,
        "maxSteps", maxSteps
    ), "SN3", "pre-fix")
    ; #endregion
    return false
}

Reminders_MenuOpenSubmenuRight() {
    ; Expand focused menu item.
    Send "{Right}"
    Sleep Reminders_DelayValue("submenu_open_ms", 80)
}

Reminders_MenuFindAndOpenSnooze(maxSteps := 20) {
    ; Find "Snooze reminder" regardless of position, then open its submenu.
    if !Reminders_MenuFindItemContains("snooze", maxSteps, "snooze")
        return false
    Reminders_MenuOpenSubmenuRight()
    return true
}

Reminders_NormalizeDurationNeedle(d) {
    d := StrLower(Trim(d))
    if (d = "30m" || d = "30 min" || d = "30 mins" || d = "30 minutes")
        return "30 minutes"
    if (d = "1h" || d = "1 hr" || d = "1 hour" || d = "one hour")
        return "1 hour"
    if (d = "4h" || d = "4 hrs" || d = "4 hours")
        return "4 hours"
    if (d = "1d" || d = "1 day")
        return "1 day"
    if (d = "1w" || d = "1 week")
        return "1 week"
    return d
}

Reminders_MenuFindAndSelectDuration(durationNeedle, maxSteps := 50) {
    durationNeedle := Reminders_NormalizeDurationNeedle(durationNeedle)
    ; After snooze submenu is open, scan items by focused text and press Enter on match.
    if !Reminders_MenuFindItemContains(durationNeedle, maxSteps, "dur:" durationNeedle)
        return false
    Send "{Enter}"
    Sleep Reminders_DelayValue("post_select_ms", 60)
    ; #region agent log
    try Reminders_DebugLog("Shift keys.ahk:Reminders_MenuFindAndSelectDuration", "Duration selected", Map(
        "duration", durationNeedle
    ), "SD1", "pre-fix")
    ; #endregion
    return true
}

Reminders_TryInvokeJoinOnlineMenuItem() {
    ; Context menus are often hosted outside the window subtree,
    ; so search from the UIA root element (desktop).
    try {
        rootDesktop := UIA.GetRootElement()
        rootWin := UIA.ElementFromHandle(WinExist("A"))

        roots := []
        if rootDesktop
            roots.Push({ root: rootDesktop, label: "desktop" })
        if rootWin
            roots.Push({ root: rootWin, label: "window" })

        for r in roots {
            root := r.root

            ; #region agent log
            try {
                menuItems := root.FindAll({ ControlType: "MenuItem" })
                cnt := menuItems ? menuItems.Length : 0
                sample := ""
                if menuItems {
                    maxSample := Min(20, menuItems.Length)
                    loop maxSample {
                        i := A_Index
                        n := ""
                        try n := menuItems[i].Name
                        if (n = "")
                            continue
                        if (sample != "")
                            sample .= " | "
                        sample .= n
                    }
                }
                Reminders_DebugLog("Shift keys.ahk:Reminders_TryInvokeJoinOnlineMenuItem", "MenuItems snapshot", Map(
                    "root", r.label,
                    "count", cnt,
                    "sample", sample
                ), "C1", "pre-fix")
            } catch {
            }
            ; #endregion

            candidates := [{ Name: "Join online", ControlType: "MenuItem" }, { Name: "Join Online", ControlType: "MenuItem" }, { Name: "Join online",
                ControlType: "Button" }, { Name: "Join Online", ControlType: "Button" }, { Name: "Join", matchmode: "Substring",
                    ControlType: "MenuItem", cs: false }, { Name: "Join", matchmode: "Substring", ControlType: "Button",
                        cs: false }
            ]

            for crit in candidates {
                mi := root.FindFirst(crit)
                if mi {
                    ; #region agent log
                    try Reminders_DebugLog("Shift keys.ahk:Reminders_TryInvokeJoinOnlineMenuItem",
                        "Found Join candidate", Map(
                            "root", r.label,
                            "name", mi.Name,
                            "type", mi.ControlType
                        ), "C2", "pre-fix")
                    ; #endregion
                    try {
                        if mi.GetPropertyValue(UIA.Property.IsOffscreen)
                            continue
                    } catch {
                    }
                    try {
                        if !mi.GetPropertyValue(UIA.Property.IsEnabled)
                            continue
                    } catch {
                    }
                    try mi.Click()
                    catch Error {
                        try mi.Invoke()
                    }
                    return true
                }
            }
        }
    } catch {
    }
    return false
}

Reminders_ExecuteItemAction(action) {
    ; Always hide the loading indicator, even on early returns or exceptions.
    ; (Prevents a stuck indicator if UIA/menu calls throw.)
    try {
        ; #region agent log
        try Reminders_DebugLog("Shift keys.ahk:Reminders_ExecuteItemAction", "Enter ExecuteItemAction", Map(
            "action", action,
            "title", WinGetTitle("A"),
            "class", WinGetClass("A"),
            "remHwnd", WinExist("A")
        ), "J1", "pre-fix")
        ; #endregion
        remHwnd := WinExist("A")

        ; Nudge window to refresh UIA before getting items (same as in SelectItem)
        Reminders_NudgeWindowForUiRefresh(remHwnd)
        Sleep 100

        ; Try 3 times with longer waits before accepting empty result
        items := []
        loop 3 {
            attempt := A_Index
            triedItems := Reminders_GetItemsStable(remHwnd, 100, 5)
            try Reminders_DebugLog("Shift keys.ahk:Reminders_ExecuteItemAction",
                "Attempt " attempt " retrieval", Map(
                    "attempt", attempt,
                    "itemsCount", triedItems.Length
                ), "J2-ATT", "pre-fix")
            if (triedItems.Length > 0) {
                items := triedItems
                break
            }
            if (attempt < 3) {
                Reminders_NudgeWindowForUiRefresh(remHwnd)
                Sleep 150  ; longer wait before retry
            }
        }

        ; Final fallback if still empty
        if (items.Length = 0) {
            try items := Reminders_GetItemsStable(remHwnd, 150, 6)
        }

        ; #region agent log
        remTitle := ""
        remClass := ""
        try remTitle := WinGetTitle("ahk_id " remHwnd)
        try remClass := WinGetClass("ahk_id " remHwnd)
        try Reminders_DebugLog("Shift keys.ahk:Reminders_ExecuteItemAction", "Items extracted", Map(
            "action", action,
            "itemsCount", items.Length,
            "remHwnd", remHwnd,
            "remTitle", remTitle,
            "remClass", remClass,
            "activeTitle", WinGetTitle("A"),
            "activeClass", WinGetClass("A")
        ), "J2", "pre-fix")
        ; #endregion
        actionLabel := ""
        if (action = "snooze_1h")
            actionLabel := "Snooze 1 hour"
        else if (action = "snooze_4h")
            actionLabel := "Snooze 4 hours"
        else if (action = "snooze_10m")
            actionLabel := "Snooze 10 minutes"
        else if (action = "snooze_1d")
            actionLabel := "Snooze 1 day"
        else if (action = "snooze_1w")
            actionLabel := "Snooze 1 week"
        else if (action = "dismiss_item")
            actionLabel := "Dismiss reminder"
        else if (action = "join_online")
            actionLabel := "Join online"
        else
            actionLabel := action

        ; Diagnostic: if items empty, show message and log more info
        if (items.Length = 0) {
            ; Capture raw UIA structure from the reminder window before returning.
            try Reminders_DebugDumpWindowTree(remHwnd, "empty-extraction")
            try {
                global g_DebugBe11ecLogPath
                FileAppend(
                    "{`"sessionId`":`"be11ec`",`"id`":`"diagnostic_" A_TickCount "_" Random(1000, 9999)
                    "`",`"timestamp`":" A_TickCount ",`"location`":`"ExecuteItemAction:empty-items`",`"message`":`"No items found after stable retrieval`",`"data`":{"
                    "`"action`":`"" action "`",`"remHwnd`":" remHwnd ",`"title`":`"" remTitle
                    "`",`"class`":`"" remClass "`",`"activeTitle`":`"" WinGetTitle("A") "`",`"activeClass`":`"" WinGetClass(
                        "A") "`"},`"runId`":`"pre-fix`",`"hypothesisId`":`"Z`"}`n",
                    g_DebugBe11ecLogPath,
                    "UTF-8"
                )
            } catch {
            }
            ShowCenteredOverlay_Utils(
                "❌ No reminders found in window. Check debug log at C:\Users\fie7ca\Documents\scripts\debug-be11ec.log",
                2500, BANNER_ACCENT_ERROR)
            return false
        }

        Reminders_LoadingShow("⏳ Reminders: " actionLabel "…")

        ; The selection modal is interactive; hide loading before it shows.
        Reminders_LoadingHide(0)
        idx := Reminders_SelectItem(actionLabel, &items, remHwnd)
        ; #region agent log
        try Reminders_DebugLog("Shift keys.ahk:Reminders_ExecuteItemAction", "Selection result", Map(
            "action", action,
            "selectedIndex", idx
        ), "J3", "pre-fix")
        ; #endregion
        if (!idx)
            return false

        ; Resume loading while executing the chosen action.
        Reminders_LoadingShow("⏳ Reminders: " actionLabel "…")

        el := items[idx].el
        Sleep Reminders_DelayValue("modal_to_row_ms", 35)  ; modal closed → row target ready before context menu (join / dismiss / snooze)
        ; #region agent log
        try Reminders_DebugLog("Shift keys.ahk:Reminders_ExecuteItemAction", "Selected reminder item", Map(
            "action", action,
            "selectedIndex", idx,
            "itemsCount", items.Length,
            "title", WinGetTitle("A")
        ), "A", "pre-fix")
        ; #endregion
        Reminders_OpenContextMenuForItem(el)
        ; #region agent log
        try Reminders_DebugLog("Shift keys.ahk:Reminders_ExecuteItemAction", "Context menu open attempt sent AppsKey",
            Map(
                "action", action
            ), "B", "pre-fix")
        ; #endregion

        ; Assume first menu item is highlighted (Snooze reminder) as per screenshots.
        if (action = "dismiss_item") {
            ; Menu order varies by reminder item (e.g. meeting reminders show Join/Chat first),
            ; so locate "Dismiss reminder" by focused UIA name instead of fixed offsets.
            ; #region agent log
            try Reminders_DebugLog("Shift keys.ahk:Reminders_ExecuteItemAction", "Dynamic dismiss requested", Map(),
            "DZ0",
            "pre-fix")
            ; #endregion
            ok := Reminders_MenuFindItemContains("dismiss", 20, "dismiss")
            if ok {
                Send "{Enter}"
                Sleep Reminders_DelayValue("post_select_ms", 60)
                ; #region agent log
                try Reminders_DebugLog("Shift keys.ahk:Reminders_ExecuteItemAction", "Dismiss invoked", Map(), "DZ1",
                "pre-fix")
                ; #endregion
                return true
            }
            ; #region agent log
            try Reminders_DebugLog("Shift keys.ahk:Reminders_ExecuteItemAction", "Dismiss not found in menu scan", Map(),
            "DZ2",
            "pre-fix")
            ; #endregion
            return false
        }

        if (action = "snooze_1h" || action = "snooze_4h" || action = "snooze_10m" || action = "snooze_1d" || action =
            "snooze_1w") {
            desired := ""
            if (action = "snooze_1h")
                desired := "1 hour"
            else if (action = "snooze_4h")
                desired := "4 hours"
            else if (action = "snooze_10m")
                desired := "10 minutes"
            else if (action = "snooze_1d")
                desired := "1 day"
            else if (action = "snooze_1w")
                desired := "1 week"
            ; #region agent log
            try Reminders_DebugLog("Shift keys.ahk:Reminders_ExecuteItemAction", "Dynamic snooze requested", Map(
                "desired", desired
            ), "SZ0", "pre-fix")
            ; #endregion
            if !Reminders_MenuFindAndOpenSnooze(20)
                return false
            return Reminders_MenuFindAndSelectDuration(desired, 60)
        }

        if (action = "join_online") {
            ; Preferred: direct UIA invoke (menu items are usually under UIA root).
            ok := false
            ; #region agent log
            try Reminders_DebugLog("Shift keys.ahk:Reminders_ExecuteItemAction", "Attempting UIA root Join invoke", Map(),
            "C",
            "pre-fix")
            ; #endregion
            ok := Reminders_TryInvokeJoinOnlineMenuItem()
            ; #region agent log
            try Reminders_DebugLog("Shift keys.ahk:Reminders_ExecuteItemAction", "UIA root Join invoke result", Map(
                "ok", ok), "C",
            "pre-fix")
            ; #endregion
            if ok
                return true

            ; Fallback 1: first-letter navigation (if supported)
            ; #region agent log
            try Reminders_DebugLog("Shift keys.ahk:Reminders_ExecuteItemAction", "Fallback: type 'j' then Enter", Map(),
            "D",
            "pre-fix")
            ; #endregion
            Send "j"
            Sleep 60
            Send "{Enter}"
            Sleep 80

            ; Fallback 2: bounded arrow scan (best-effort, no UIA reads)
            ; #region agent log
            try Reminders_DebugLog("Shift keys.ahk:Reminders_ExecuteItemAction",
                "Fallback: bounded arrow scan then Enter", Map(),
                "E", "pre-fix")
            ; #endregion
            Send "{Home}"
            loop 12 {
                Send "{Down}"
                Sleep 40
            }
            Send "{Enter}"

            return false
        }

        return false
    } finally {
        Reminders_LoadingHide(0)
    }
}

; Shift + H : Snooze 1 hour - Hour
+H:: {
    isNewRem := Reminders_IsNewOutlookWindow()
    ; #region agent log
    try Reminders_DebugLog("Shift keys.ahk:+H", "Shift+H pressed", Map(
        "isNewReminders", isNewRem,
        "title", WinGetTitle("A"),
        "class", WinGetClass("A")
    ), "H2", "pre-fix")
    ; #endregion
    if isNewRem {
        Reminders_ExecuteItemAction("snooze_1h")
        return
    }
    ; #region agent log
    try Reminders_DebugLog("Shift keys.ahk:+H", "Falling back to classic Confirm(1 hour)", Map(), "H3", "pre-fix")
    ; #endregion
    Confirm("1 hour")
}

; Shift + F : Snooze 4 hours - Four
+F:: {
    if Reminders_IsNewOutlookWindow() {
        Reminders_ExecuteItemAction("snooze_4h")
        return
    }
    Confirm("4 hours")
}

; Shift + T : Snooze 10 minutes - Ten
+T:: {
    if Reminders_IsNewOutlookWindow() {
        Reminders_ExecuteItemAction("snooze_10m")
        return
    }
    Confirm("10 minutes")
}

; Shift + Y : Snooze 1 day - daY
+Y:: {
    if Reminders_IsNewOutlookWindow() {
        Reminders_ExecuteItemAction("snooze_1d")
        return
    }
    Confirm("1 day")
}

; Shift + W : Snooze 1 week - Week
+W:: {
    if Reminders_IsNewOutlookWindow() {
        Reminders_ExecuteItemAction("snooze_1w")
        return
    }
    Confirm("1 week")
}

; Shift + D : Dismiss reminder (New Outlook) / Snooze 1 day (classic)
+D:: {
    if Reminders_IsNewOutlookWindow() {
        Reminders_ExecuteItemAction("dismiss_item")
        return
    }
    Confirm("1 day")
}

; Shift + X : Dismiss all reminders - Dismiss (confirm first: New Outlook + classic)
+X:: {
    if Reminders_IsNewOutlookWindow() {
        if MsgBox("Dismiss all reminders?", "Confirm Dismiss", "YesNo Icon?") != "Yes"
            return
        ; Ensure loading indicator can't get stuck on exceptions.
        try {
            Reminders_LoadingShow("⏳ Reminders: Dismiss all…")
            ; Global action: click "Dismiss all" button (UIA)
            try {
                root := UIA.ElementFromHandle(WinExist("A"))
                btn := root.FindFirst({ Name: "Dismiss all", ControlType: "Button" })
                if !btn
                    btn := root.FindFirst({ Name: "Dismiss All", ControlType: "Button" })
                if btn {
                    btn.Click()
                    return
                }
            } catch {
            }
            return
        } finally {
            Reminders_LoadingHide(0)
        }
    }
    ConfirmDismissAll()
}

; Shift + J : Join Online - Join
+J:: {
    isNewRem := Reminders_IsNewOutlookWindow()
    ; #region agent log
    try Reminders_DebugLog("Shift keys.ahk:+J", "Shift+J pressed", Map(
        "isNewReminders", isNewRem,
        "title", WinGetTitle("A"),
        "class", WinGetClass("A")
    ), "F", "pre-fix")
    ; #endregion
    if isNewRem {
        Reminders_ExecuteItemAction("join_online")
        return
    }
    try {
        win := WinExist("A")
        root := UIA.ElementFromHandle(win)

        ; Find the "Join Online" button (classic reminder UI)
        joinButton := root.FindFirst({ Name: "Join Online", Type: "50000", AutomationId: "8346" })
        if !joinButton
            joinButton := root.FindFirst({ Name: "Join Online", ControlType: "Button" })
        if !joinButton
            joinButton := root.FindFirst({ AutomationId: "8346" })

        if joinButton
            joinButton.Click()
    } catch {
    }
}

#HotIf
