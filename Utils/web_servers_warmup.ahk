; =============================================================================
; Utils/web_servers_warmup.ahk
; Act.ahk startup: pre-start Tasks (:8766) and Memory Palace (:8767) Python servers.
; Fire-and-forget; exits without stopping servers (g_*LauncherSkipOnExit).
; =============================================================================
#Requires AutoHotkey v2.0+
#SingleInstance Off

global g_TaskLauncherSkipOnExit := true
global g_PalaceLauncherSkipOnExit := true

#Include %A_ScriptDir%\..\env.ahk
#Include %A_ScriptDir%\standard_loading_bar.ahk

; Task_Notify / Palace_Notify fall back to TrayTip when this stub is present.
ShowCenteredOverlay_Utils(*) {
}

#Include %A_ScriptDir%\task_helpers.ahk
#Include %A_ScriptDir%\task_launcher.ahk
#Include %A_ScriptDir%\mnemonic_palace_helpers.ahk
#Include %A_ScriptDir%\mnemonic_palace_launcher.ahk

try Task_EnsureData()
catch {
}
try Task_WriteEnvDefaultFocus()
catch {
}
try Palace_EnsureData()
catch {
}

if (!Task_IsServerRunning())
    Task_EnsureServer(false)
if (!Palace_IsServerRunning())
    Palace_EnsureServer(false)

ExitApp