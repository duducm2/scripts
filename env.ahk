; Environment Variables for AutoHotKey Scripts
; This file contains all the path configurations for different environments

; Base paths
global WORK_SCRIPTS_PATH := "C:\Users\fie7ca\Documents\01 - Scripts"
global PERSONAL_SCRIPTS_PATH := "G:\Meu Drive\12 - Scripts"

; Environment detection
global IS_WORK_ENVIRONMENT := true  ; Set this to false for personal environment

; Function to get the correct path based on environment
GetScriptPath(scriptName) {
    if (IS_WORK_ENVIRONMENT) {
        return WORK_SCRIPTS_PATH . "\" . scriptName
    } else {
        return PERSONAL_SCRIPTS_PATH . "\" . scriptName
    }
}

; Example usage:
; Run GetScriptPath("Spotify - Open.ahk")
