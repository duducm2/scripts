; =============================================================================
; Utils module: import_watcher.ahk
; Always-on Desktop poller for canonical AI import packs.
; Detects stable FINANCE_*/PALACE_*/PLAN_*/TASK_PACK files, runs existing domain
; importers (preview → confirm → commit), then routes fresh AI-fix files to the
; active companion via import_watcher_companion.ahk.
; Owner process: AppLaunchers.ahk only (same pin as dictation / Handy AI).
; Agent docs: docs/prompt-data-output-and-finance-packs.md
; =============================================================================

global g_ImportWatcherEnabled := true
global g_ImportWatcherBusy := false
global g_ImportWatcherSeen := Map()       ; lowerPath → "mtime|size"
global g_ImportWatcherPending := Map()    ; lowerPath → { mtime, size, stable }
global g_ImportWatcherTimerArmed := false
global g_ImportWatcherPollMs := 2000
global IMPORT_WATCHER_STABLE_POLLS := 2

; AppLaunchers is the single long-lived owner (mirrors Dictation_IsOwnerProcess).
ImportWatcher_IsOwnerProcess() {
    return A_ScriptName = "AppLaunchers.ahk"
}

ImportWatcher_ResolveDesktopPath() {
    desktop := ""
    try desktop := GetDesktopToRecyclePath()
    catch {
        desktop := A_Desktop
    }
    if (!desktop || !DirExist(desktop))
        desktop := A_Desktop
    return RTrim(desktop, "\")
}

ImportWatcher_PollMs() {
    global g_ImportWatcherPollMs
    if (IsSet(IMPORT_WATCHER_POLL_MS) && Integer(IMPORT_WATCHER_POLL_MS) > 0)
        return Integer(IMPORT_WATCHER_POLL_MS)
    return g_ImportWatcherPollMs > 0 ? g_ImportWatcherPollMs : 2000
}

ImportWatcher_DefaultEnabled() {
    if (IsSet(IMPORT_WATCHER_ENABLED))
        return !!IMPORT_WATCHER_ENABLED
    return true
}

; Catalog: pattern → { id, label, canonical, run }
ImportWatcher_Catalog() {
    return [
        Map("id", "finance_daily", "label", "Finance daily",
            "patterns", ["FINANCE_DAILY*.txt", "FINANCE_DAILY*.csv"],
            "canonical", "FINANCE_DAILY.txt", "run", Finance_ImportDaily),
        Map("id", "finance_monthly", "label", "Finance monthly",
            "patterns", ["FINANCE_MONTHLY*.txt", "FINANCE_MONTHLY*.csv", "FINANCE_MONTHLY*.ini"],
            "canonical", "FINANCE_MONTHLY.txt", "run", Finance_ImportMonthly),
        Map("id", "palace_pack", "label", "Palace pack",
            "patterns", ["PALACE_PACK*.txt", "PALACE_PACK*.csv"],
            "canonical", "PALACE_PACK.txt", "run", Palace_ImportMnemonicsFromDesktop),
        Map("id", "plan_pack", "label", "Study plan pack",
            "patterns", ["PLAN_PACK*.txt", "PLAN_PACK*.csv"],
            "canonical", "PLAN_PACK.txt", "run", Palace_ImportPlanPackFromDesktop),
        Map("id", "task_pack", "label", "Task pack",
            "patterns", ["TASK_PACK*.txt", "TASK_PACK*.csv"],
            "canonical", "TASK_PACK.txt", "run", Task_ImportPackFromDesktop)
    ]
}

ImportWatcher_FileStamp(path) {
    mtime := ""
    size := 0
    try mtime := FileGetTime(path, "M")
    catch {
        return ""
    }
    try size := FileGetSize(path)
    catch {
        size := 0
    }
    return mtime . "|" . size
}

ImportWatcher_LowerPath(path) {
    return StrLower(RTrim(path, "\"))
}

; Seed seen-map with whatever is already on Desktop so startup does not re-import.
ImportWatcher_SeedSeen() {
    global g_ImportWatcherSeen
    g_ImportWatcherSeen := Map()
    desktop := ImportWatcher_ResolveDesktopPath()
    if (desktop = "")
        return
    for item in ImportWatcher_Catalog() {
        for pat in item["patterns"] {
            loop files desktop . "\" . pat, "F" {
                key := ImportWatcher_LowerPath(A_LoopFileFullPath)
                stamp := ImportWatcher_FileStamp(A_LoopFileFullPath)
                if (stamp != "")
                    g_ImportWatcherSeen[key] := stamp
            }
        }
    }
}

ImportWatcher_ScanCandidates() {
    global g_ImportWatcherSeen, g_ImportWatcherPending, IMPORT_WATCHER_STABLE_POLLS
    ready := []
    desktop := ImportWatcher_ResolveDesktopPath()
    if (desktop = "")
        return ready

    foundKeys := Map()
    for item in ImportWatcher_Catalog() {
        for pat in item["patterns"] {
            loop files desktop . "\" . pat, "F" {
                path := A_LoopFileFullPath
                key := ImportWatcher_LowerPath(path)
                foundKeys[key] := true
                stamp := ImportWatcher_FileStamp(path)
                if (stamp = "")
                    continue
                ; Already processed at this exact stamp.
                if (g_ImportWatcherSeen.Has(key) && g_ImportWatcherSeen[key] = stamp) {
                    if (g_ImportWatcherPending.Has(key))
                        g_ImportWatcherPending.Delete(key)
                    continue
                }
                parts := StrSplit(stamp, "|")
                mtime := parts.Length >= 1 ? parts[1] : ""
                size := parts.Length >= 2 ? parts[2] : "0"
                if (!g_ImportWatcherPending.Has(key)
                    || g_ImportWatcherPending[key].mtime != mtime
                    || g_ImportWatcherPending[key].size != size) {
                    g_ImportWatcherPending[key] := { mtime: mtime, size: size, stable: 1, path: path, item: item }
                    continue
                }
                g_ImportWatcherPending[key].stable += 1
                g_ImportWatcherPending[key].path := path
                g_ImportWatcherPending[key].item := item
                if (g_ImportWatcherPending[key].stable >= IMPORT_WATCHER_STABLE_POLLS)
                    ready.Push({ path: path, stamp: stamp, item: item, key: key })
            }
        }
    }

    ; Drop pending entries for files that disappeared.
    toDelete := []
    for key, _ in g_ImportWatcherPending {
        if (!foundKeys.Has(key))
            toDelete.Push(key)
    }
    for key in toDelete
        g_ImportWatcherPending.Delete(key)

    return ready
}

ImportWatcher_MarkSeen(key, stamp) {
    global g_ImportWatcherSeen, g_ImportWatcherPending
    if (key != "" && stamp != "")
        g_ImportWatcherSeen[key] := stamp
    if (key != "" && g_ImportWatcherPending.Has(key))
        g_ImportWatcherPending.Delete(key)
}

ImportWatcher_ShowDetected(label) {
    try StandardLoadingBar_Show("📥 Detected " . label . " pack on Desktop…", BANNER_ACCENT_INFO, {
        passive: false
    })
    catch {
    }
}

ImportWatcher_ShowPreparing(label) {
    try StandardLoadingBar_Update("⏳ Preparing " . label . " preview…", BANNER_ACCENT_INTERMEDIATE)
    catch {
        try StandardLoadingBar_Show("⏳ Preparing " . label . " preview…", BANNER_ACCENT_INTERMEDIATE, {
            passive: false
        })
        catch {
        }
    }
}

ImportWatcher_HideBanner(delayMs := 200) {
    try StandardLoadingBar_Hide(delayMs)
    catch {
    }
}

; Normalize discovered path to canonical Desktop name, then run domain importer.
ImportWatcher_RunItem(path, item) {
    canonical := item["canonical"]
    label := item["label"]
    runFn := item["run"]
    ImportWatcher_ShowPreparing(label)
    normalized := path
    try normalized := PackImport_NormalizeDesktopSource(path, canonical)
    catch {
        normalized := path
    }
    if (normalized = "")
        normalized := path
    ; Domain runners discover on Desktop themselves; hide banner once they take over.
    ; Brief settle so the user sees the detected/preparing state.
    Sleep 350
    ImportWatcher_HideBanner(0)
    try runFn.Call()
    catch as e {
        try ShowCenteredOverlay_Utils("Import failed: " . e.Message, 2800, BANNER_ACCENT_ERROR)
        catch {
            TrayTip("Import", "Import failed")
        }
    }
}

ImportWatcher_ProcessCandidate(cand) {
    global g_ImportWatcherBusy
    if (g_ImportWatcherBusy)
        return
    g_ImportWatcherBusy := true
    importStartStamp := FormatTime(, "yyyyMMddHHmmss")
    try {
        ImportWatcher_ShowDetected(cand.item["label"])
        Sleep 400
        ImportWatcher_RunItem(cand.path, cand.item)
        ; After import: route a freshly written AI-fix file to the companion.
        try ImportWatcher_CompanionHandleAiFixAfterImport(importStartStamp)
        catch {
        }
    } finally {
        ; Mark current stamp seen (file may be archived/gone — then stamp from cand).
        stamp := cand.stamp
        if (FileExist(cand.path)) {
            try stamp := ImportWatcher_FileStamp(cand.path)
            catch {
            }
        }
        ImportWatcher_MarkSeen(cand.key, stamp)
        ; Also mark canonical path if normalize moved the file.
        canonPath := PackImport_CanonicalDesktopPath(cand.item["canonical"])
        if (FileExist(canonPath)) {
            cStamp := ImportWatcher_FileStamp(canonPath)
            ImportWatcher_MarkSeen(ImportWatcher_LowerPath(canonPath), cStamp)
        } else {
            ImportWatcher_MarkSeen(ImportWatcher_LowerPath(canonPath), stamp)
        }
        g_ImportWatcherBusy := false
        ImportWatcher_HideBanner(0)
    }
}

ImportWatcher_Poll(*) {
    global g_ImportWatcherEnabled, g_ImportWatcherBusy
    if (!g_ImportWatcherEnabled || g_ImportWatcherBusy)
        return
    if (!ImportWatcher_IsOwnerProcess())
        return
    ready := []
    try ready := ImportWatcher_ScanCandidates()
    catch {
        return
    }
    if (!ready.Length)
        return
    ; One at a time; pick the first stable candidate.
    cand := ready[1]
    ImportWatcher_ProcessCandidate(cand)
}

ImportWatcher_Start() {
    global g_ImportWatcherEnabled, g_ImportWatcherTimerArmed, g_ImportWatcherBusy
    if (!ImportWatcher_IsOwnerProcess())
        return false
    g_ImportWatcherEnabled := ImportWatcher_DefaultEnabled()
    g_ImportWatcherBusy := false
    ImportWatcher_SeedSeen()
    if (!g_ImportWatcherEnabled) {
        ImportWatcher_StopTimer()
        return false
    }
    pollMs := ImportWatcher_PollMs()
    try SetTimer(ImportWatcher_Poll, pollMs)
    catch {
        return false
    }
    g_ImportWatcherTimerArmed := true
    ImportWatcher_RefreshTray()
    return true
}

ImportWatcher_StopTimer() {
    global g_ImportWatcherTimerArmed
    try SetTimer(ImportWatcher_Poll, 0)
    catch {
    }
    g_ImportWatcherTimerArmed := false
}

ImportWatcher_Stop() {
    global g_ImportWatcherEnabled
    g_ImportWatcherEnabled := false
    ImportWatcher_StopTimer()
    ImportWatcher_RefreshTray()
}

ImportWatcher_Toggle(*) {
    global g_ImportWatcherEnabled
    if (!ImportWatcher_IsOwnerProcess())
        return
    if (g_ImportWatcherEnabled) {
        ImportWatcher_Stop()
        try ShowCenteredOverlay_Utils("Import watcher off", 1500, BANNER_ACCENT_INTERMEDIATE)
        catch {
            TrayTip("Import", "Import watcher off")
        }
    } else {
        g_ImportWatcherEnabled := true
        ImportWatcher_SeedSeen()
        pollMs := ImportWatcher_PollMs()
        try SetTimer(ImportWatcher_Poll, pollMs)
        catch {
        }
        global g_ImportWatcherTimerArmed
        g_ImportWatcherTimerArmed := true
        ImportWatcher_RefreshTray()
        try ShowCenteredOverlay_Utils("Import watcher on", 1500, BANNER_ACCENT_SUCCESS)
        catch {
            TrayTip("Import", "Import watcher on")
        }
    }
}

ImportWatcher_TrayLabel() {
    global g_ImportWatcherEnabled
    return g_ImportWatcherEnabled ? "Import watcher: ON" : "Import watcher: OFF"
}

ImportWatcher_RefreshTray() {
    if (!ImportWatcher_IsOwnerProcess())
        return
    try {
        A_TrayMenu.Delete("Import watcher: ON")
    } catch {
    }
    try {
        A_TrayMenu.Delete("Import watcher: OFF")
    } catch {
    }
    try {
        A_TrayMenu.Add(ImportWatcher_TrayLabel(), ImportWatcher_Toggle)
    } catch {
    }
}

; Auto-init when this module is loaded from AppLaunchers via Utils.ahk.
ImportWatcher_Init() {
    if (!ImportWatcher_IsOwnerProcess())
        return
    ; Defer so the rest of Utils / AppLaunchers hotkeys finish registering.
    SetTimer(ImportWatcher_Start, -800)
}
