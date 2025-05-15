#Requires AutoHotkey v2.0

#Include env.ahk ; Ensures IS_WORK_ENVIRONMENT is available

; Define your hotkey for Hunt and Peck here if this script is standalone
; For example, if it was Alt + Shift + 0
CapsLock & 0:: ; Alt + Shift + 0
{
    if (IS_WORK_ENVIRONMENT) {
        Run "C:\Users\fie7ca\Documents\HuntAndPack\hap.exe"
    } else { ; Personal Environment
        Run "C:\Users\eduev\OneDrive\Documentos\HuntAndPeck\hap.exe"
    }
    ; If CapsLock state needs to be toggled or managed, add Send("{CapsLock}") or similar here
}

; If this script is ONLY run by Act.ahk and does not define its own hotkey,
; then the content should be just the conditional Run block, without the hotkey definition:
;
; #Include env.ahk
; if (IS_WORK_ENVIRONMENT) {
;     Run "C:\Users\fie7ca\Documents\HuntAndPack\hap.exe"
; } else { ; Personal Environment
;     Run "C:\Users\eduev\OneDrive\Documentos\HuntAndPeck\hap.exe"
; }

#UseHook  ; <-- This is important

; Change the default shortcuts (Let's say you want `ctrl + ,`
