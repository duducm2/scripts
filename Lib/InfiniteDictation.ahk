; Infinite Dictation Module
; Isolated state and loop logic for Handy.exe lifecycle and ClipAngel merge.
; State is persisted to data/infinite_dictation.ini so only one instance is ever active (single process or across reloads).
; Requires Utils.ahk (or host script) to provide: DictationCleanup_StartCountdown,
; Handy_ActivateOrLaunch, DictationMerge_StartCountdown, MergeNonFavoriteClips,
; globals g_PendingDictationMerge, g_ProgrammaticDictationStop, g_DictationLoopActive,
; g_DictationLoopSound, IsSoundEnabled, ShowCenteredOverlay_Utils.

class InfiniteDictation {
    ; Static state
    static IsActive := false
    static LoopTimer := ""   ; one-shot stop timer
    static StartRetries := 0
    static MonitorTimer := "" ; high-frequency exit monitor (optional)

    static _IniPath() {
        return A_ScriptDir "\data\infinite_dictation.ini"
    }

    ; Read persisted state. Returns {Active: 0|1, Pid: number}. If Active=1 and Pid is another running process, another instance owns the loop.
    static _ReadPersistedState() {
        path := InfiniteDictation._IniPath()
        active := Integer(IniRead(path, "State", "Active", "0"))
        pid := Integer(IniRead(path, "State", "Pid", "0"))
        return { Active: active, Pid: pid }
    }

    static _WritePersistedState(active, pid := 0) {
        IniWrite(String(active), InfiniteDictation._IniPath(), "State", "Active")
        IniWrite(String(pid), InfiniteDictation._IniPath(), "State", "Pid")
    }

    ; Ensure only one Infinite Dictation can be active. Returns true if we may proceed to start, false if another instance owns it.
    static _ClaimOrRejectStart() {
        myPid := DllCall("GetCurrentProcessId", "UInt")  ; current script's PID (avoids #Warn on A_ProcessId)
        s := InfiniteDictation._ReadPersistedState()
        if (s.Active = 0)
            return true
        if (s.Pid = myPid)
            return true
        if (ProcessExist(s.Pid))
            return false
        InfiniteDictation._WritePersistedState(0, 0)
        return true
    }

    static Start() {
        if (!InfiniteDictation._ClaimOrRejectStart()) {
            try ShowCenteredOverlay_Utils("Infinite Dictation already active in another process", 3000)
            catch {
            }
            return
        }
        DictationCleanup_StartCountdown(5)
        InfiniteDictation.IsActive := true
        global g_DictationLoopActive := true
        InfiniteDictation._WritePersistedState(1, DllCall("GetCurrentProcessId", "UInt"))
        SetTimer(ObjBindMethod(InfiniteDictation, "StopHandyAndRestart"), 0)
        SetTimer(ObjBindMethod(InfiniteDictation, "LoopCycle"), 0)
        InfiniteDictation.LoopCycle()
    }

    ; Start without the 5s clipboard cleanup countdown (e.g. for ToggleDictationLoop / DictationStartWithClipboardOption).
    static StartWithoutCleanup() {
        if (!InfiniteDictation._ClaimOrRejectStart()) {
            try ShowCenteredOverlay_Utils("Infinite Dictation already active in another process", 3000)
            catch {
            }
            return
        }
        InfiniteDictation.IsActive := true
        global g_DictationLoopActive := true
        InfiniteDictation._WritePersistedState(1, DllCall("GetCurrentProcessId", "UInt"))
        SetTimer(ObjBindMethod(InfiniteDictation, "StopHandyAndRestart"), 0)
        SetTimer(ObjBindMethod(InfiniteDictation, "LoopCycle"), 0)
        InfiniteDictation.LoopCycle()
    }

    ; Stop Infinite Dictation: cancel timers, send stop, request merge after transcription.
    static Stop() {
        InfiniteDictation.IsActive := false
        global g_DictationLoopActive := false
        global g_PendingDictationMerge := true
        global g_ProgrammaticDictationStop := true
        InfiniteDictation._WritePersistedState(0, 0)
        SetTimer(ObjBindMethod(InfiniteDictation, "StopHandyAndRestart"), 0)
        SetTimer(ObjBindMethod(InfiniteDictation, "LoopCycle"), 0)
        SetTimer(ObjBindMethod(InfiniteDictation, "VerifyStart"), 0)
        SetTimer(ObjBindMethod(InfiniteDictation, "MonitorHandyExit"), 0)
        SendInput "#!+0"
    }

    ; One loop cycle: ensure Handy running, trigger start, schedule stop after 15s.
    static LoopCycle() {
        if (!InfiniteDictation.IsActive)
            return
        if (!ProcessExist("handy.exe")) {
            Handy_ActivateOrLaunch()
            Sleep 2000
        }
        if (WinExist("Recording ahk_exe handy.exe")) {
            SetTimer(ObjBindMethod(InfiniteDictation, "StopHandyAndRestart"), 0)
            SetTimer(ObjBindMethod(InfiniteDictation, "StopHandyAndRestart"), -15000)
            return
        }
        InfiniteDictation.StartRetries := 0
        global g_ProgrammaticDictationStop := true
        SendEvent "#!+0"
        if (!InfiniteDictation.IsActive)
            return
        SetTimer(ObjBindMethod(InfiniteDictation, "StopHandyAndRestart"), 0)
        SetTimer(ObjBindMethod(InfiniteDictation, "StopHandyAndRestart"), -15000)
        SetTimer(ObjBindMethod(InfiniteDictation, "VerifyStart"), -1500)
    }

    ; Trigger Win+Alt+Shift+0 (start/stop Handy dictation).
    static TriggerHandyToggle() {
        global g_ProgrammaticDictationStop := true
        SendEvent "#!+0"
    }

    ; Scheduled after 15s: stop recording (send #!+0), play sound; next cycle is scheduled by PlayDictationCompletionChime in Utils.
    static StopHandyAndRestart() {
        if (!InfiniteDictation.IsActive)
            return
        if (WinExist("Recording ahk_exe handy.exe")) {
            global g_ProgrammaticDictationStop := true
            global g_DictationLoopSound
            SendEvent "#!+0"
            try {
                if (IsSoundEnabled())
                    SoundPlay(g_DictationLoopSound)
            } catch {
            }
        } else {
            SetTimer(ObjBindMethod(InfiniteDictation, "LoopCycle"), -1000)
        }
        ; Next loop is started by Utils.PlayDictationCompletionChime -> SetTimer(DictationLoopStart, -2000) -> Utils must call back to InfiniteDictation.LoopCycle
    }

    ; Optional: wait for handy.exe to exit with high-frequency polling; force close after timeout, then trigger next cycle.
    static WaitForHandyExit(timeoutMs := 30000, pollIntervalMs := 50) {
        start := A_TickCount
        while (ProcessExist("handy.exe") && (A_TickCount - start < timeoutMs))
            Sleep(pollIntervalMs)
        if (ProcessExist("handy.exe")) {
            try
                ProcessClose("handy.exe")
            catch {
            }
        }
    }

    ; High-frequency monitor (50ms): when handy.exe exits, trigger next cycle. Used if we want strict exit-before-restart.
    static MonitorHandyExit() {
        if (!InfiniteDictation.IsActive)
            return
        if (ProcessExist("handy.exe"))
            return
        SetTimer(ObjBindMethod(InfiniteDictation, "MonitorHandyExit"), 0)
        SetTimer(ObjBindMethod(InfiniteDictation, "LoopCycle"), -2000)
    }

    ; Verify Recording window appeared; retry start up to 3 times.
    static VerifyStart() {
        if (!InfiniteDictation.IsActive)
            return
        if (WinExist("Recording ahk_exe handy.exe"))
            return
        InfiniteDictation.StartRetries++
        if (InfiniteDictation.StartRetries <= 3) {
            global g_ProgrammaticDictationStop := true
            SendEvent "#!+0"
            SetTimer(ObjBindMethod(InfiniteDictation, "StopHandyAndRestart"), 0)
            SetTimer(ObjBindMethod(InfiniteDictation, "StopHandyAndRestart"), -15000)
            SetTimer(ObjBindMethod(InfiniteDictation, "VerifyStart"), -1500)
        } else {
            try
                ShowCenteredOverlay_Utils("Failed to start dictation", 2000)
            catch {
            }
            InfiniteDictation.IsActive := false
            global g_DictationLoopActive := false
            InfiniteDictation._WritePersistedState(0, 0)
        }
    }

    ; Called by host when transcription is complete and loop is still active (schedule next cycle in 2s).
    static OnTranscriptionComplete() {
        if (!InfiniteDictation.IsActive)
            return
        SetTimer(ObjBindMethod(InfiniteDictation, "LoopCycle"), -2000)
    }
}
