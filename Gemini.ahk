#Requires AutoHotkey v2.0
#SingleInstance Force

; --- Includes ----------------------------------------------------------------
#include vendor\UIA-v2\Lib\UIA.ahk
#include vendor\UIA-v2\Lib\UIA_Browser.ahk
#include %A_ScriptDir%\env.ahk
#include %A_ScriptDir%\Utils.ahk
try Hotkey("#!+Y", "Off")
try Hotkey("#!+X", "Off")
#include %A_ScriptDir%\infra\ipc\WMIPC.ahk

#include %A_ScriptDir%\infra\ipc\GeminiIPC.ahk

; -----------------------------------------------------------------------------
; MODULE MAP - Gemini.ahk stays the runnable entry point and #includes each
; module below. Early preamble includes (UIA, env, Utils, WMIPC, GeminiIPC) stay
; here. See Gemini/MODULARIZATION_PROGRESS.md for the full module list.
; -----------------------------------------------------------------------------

; [Gemini module] Threshold constants, feature flags, early helpers -> Gemini\config_constants.ahk
#include %A_ScriptDir%\Gemini\config_constants.ahk

; [Gemini module] GeminiState, copy-button UIA, tab banner, model picker -> Gemini\gemini_uia_core.ahk
#include %A_ScriptDir%\Gemini\gemini_uia_core.ahk

; [Gemini module] ShowNotification, background timers, copy chime -> Gemini\background_helpers.ahk
#include %A_ScriptDir%\Gemini\background_helpers.ahk

; [Gemini module] Small loading indicator and WaitForButton helpers -> Gemini\loading_wait.ahk
#include %A_ScriptDir%\Gemini\loading_wait.ahk

; [Gemini module] #!+P, empty #!+O stub, copy helper, read-aloud IPC -> Gemini\hotkey_read_copy.ahk
#include %A_ScriptDir%\Gemini\hotkey_read_copy.ahk

; [Gemini module] Language picker and #!+8 pronunciation hotkey -> Gemini\hotkey_pronunciation.ahk
#include %A_ScriptDir%\Gemini\hotkey_pronunciation.ahk

; [Gemini module] InitializeGeminiFirstTime and #!+I open/focus hotkey -> Gemini\gemini_open.ahk
#include %A_ScriptDir%\Gemini\gemini_open.ahk

; Gemini_FocusPromptSameAsOpenHotkey lives in Utils.ahk (shared with Shift keys Fast Copy).

; [Gemini module] GeminiAsyncReadAloud async read-aloud class -> Gemini\gemini_async_readaloud.ahk
#include %A_ScriptDir%\Gemini\gemini_async_readaloud.ahk

; [Gemini module] GeminiAsyncLookup pronunciation async class -> Gemini\gemini_async_lookup.ahk
#include %A_ScriptDir%\Gemini\gemini_async_lookup.ahk

; [Gemini module] GeminiDelayedSubmitMonitor and start/stop helpers -> Gemini\gemini_delayed_submit.ahk
#include %A_ScriptDir%\Gemini\gemini_delayed_submit.ahk

; [Gemini module] GeminiAsyncTTS class and GeminiQueueBackgroundTask -> Gemini\gemini_async_tts.ahk
#include %A_ScriptDir%\Gemini\gemini_async_tts.ahk