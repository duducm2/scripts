; =============================================================================
; Shift keys module: hotif_settings.ahk
; Windows Settings hotkeys
; Extracted verbatim from Shift keys.ahk; loaded via #include into the
; Shift keys.ahk process, which remains the entry point / source of truth.
; =============================================================================

#HotIf WinActive("Settings") || WinActive("ConfiguraÃ§Ãµes")

; Shift + V : Set input volume (see SCRIPT_MICROPHONE_INPUT_SLIDER_PERCENT in Utils.ahk)
+V::
{
    try {
        ; Get the active Settings window
        settingsHwnd := WinExist("A")
        settingsRoot := UIA.ElementFromHandle(settingsHwnd)

        ; Try to find the input volume slider by AutomationId first (most reliable)
        volumeSlider := ""
        try {
            volumeSlider := settingsRoot.FindFirst({ AutomationId: "SystemSettings_Audio_Input_VolumeValue_Slider",
                ControlType: "Slider" })
        } catch {
            ; Fallback: Try by name (both English and Portuguese)
            sliderNames := ["Input volume", "Ajustar o volume de entrada"]
            for sliderName in sliderNames {
                try {
                    volumeSlider := settingsRoot.FindFirst({ Name: sliderName, ControlType: "Slider" })
                    if volumeSlider
                        break
                } catch {
                    continue
                }
            }
        }

        if volumeSlider {
            volumeSlider.SetValue(SCRIPT_MICROPHONE_INPUT_SLIDER_PERCENT)
            ; Optional: Brief confirmation
            ToolTip("Input volume set to " SCRIPT_MICROPHONE_INPUT_SLIDER_PERCENT "%")
            SetTimer(() => ToolTip(), -1000)

            ; Close the Settings window
            Sleep(300) ; Small delay to ensure setting is applied
            Send("!{F4}") ; Alt+F4 to close Settings
        } else {
            MsgBox("Input volume slider not found. Make sure you're on the microphone settings page.", "Error", "IconX"
            )
        }

    } catch Error as e {
        MsgBox("Error setting input volume: " . e.Message, "Error", "IconX")
    }
}

