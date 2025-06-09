; Environment Variables for AutoHotKey Scripts
; This file contains all the path configurations for different environments

; Base paths for scripts
global WORK_SCRIPTS_PATH := "C:\Users\fie7ca\Documents\01 - Scripts"
global PERSONAL_SCRIPTS_PATH := "G:\Meu Drive\12 - Scripts"

; Environment detection
global IS_WORK_ENVIRONMENT := true  ; Set this to true for work environment, false for personal

; Function to get the correct script path based on environment
GetScriptPath(scriptName) {
    if (IS_WORK_ENVIRONMENT) {
        return WORK_SCRIPTS_PATH . "\" . scriptName
    } else {
        return PERSONAL_SCRIPTS_PATH . "\" . scriptName
    }
}
