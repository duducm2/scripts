; =============================================================================
; AppLaunchers module: wikipedia_entry.ahk
; SelectWikipediaInHandy and #!+k hotkey
; Extracted verbatim from AppLaunchers.ahk; loaded via #include into the
; AppLaunchers.ahk process, which remains the entry point / source of truth.
; =============================================================================

; =============================================================================
; SelectWikipediaInHandy() — Opens/closes selector or activates Wikipedia (same pattern as SelectAiModelInHandy in Utils.ahk)
; =============================================================================
SelectWikipediaInHandy() {
    global g_WikipediaSelectorActive
    if (g_WikipediaSelectorActive) {
        WikipediaSelector_Cancel()
        return
    }

    SetTitleMatchMode 2
    if WinExist("Wikipedia") {
        ; Window already exists - use reduced delay for activation
        windowAlreadyOpen := true
        WinActivate
        WinWaitActive("Wikipedia", , 2)
        ; Ensure Chrome is active (Wikipedia windows are Chrome windows)
        WinWaitActive("ahk_exe chrome.exe", , 2)
        Sleep(100)  ; Reduced delay since window is already open
        CenterMouse()
    } else {
        ; Window doesn't exist yet - it will be created by selector
        ; This path is handled by the selector logic below
        windowAlreadyOpen := false
    }

    ; If window was activated (already existed), proceed with fullscreen setup
    if (windowAlreadyOpen) {
        ; Check for saved scroll position immediately to prioritize feedback
        url := GetWikipediaURL()
        savedPercentage := 0.0
        if (url != "") {
            savedPercentage := LoadWikipediaScrollPosition(url)
        }

        if (savedPercentage > 0.0) {
            StandardLoadingBar_Show("📜 Restoring scroll position... Please wait", BANNER_ACCENT_INTERMEDIATE)
            AL_InstallInputGuard()

            ; Initialize UIA_Browser with retry logic
            uia := false
            uiaRetries := 3
            loop uiaRetries {
                try {
                    uia := UIA_Browser("ahk_exe chrome.exe")
                    if (uia) {
                        break
                    }
                } catch Error as uiaErr {
                    if (A_Index < uiaRetries) {
                        Sleep(200)
                    }
                }
            }

            if (!uia) {
                AL_RemoveInputGuard()
                StandardLoadingBar_Update("Error: Could not access browser")
                StandardLoadingBar_Hide(1000)
                Send("{F11}")
                Sleep(300)
            } else {
                ; Wait for page stability
                Sleep(500)

                ; Get current document height with retry logic and stabilization
                docHeight := ""
                docHeightRetries := 5
                lastDocHeight := 0
                stableCount := 0
                loop docHeightRetries {
                    try {
                        docHeight := uia.JSReturnThroughClipboard("document.documentElement.scrollHeight")
                        if (docHeight != "" && docHeight != "undefined" && docHeight != "null") {
                            docHeightFloat := Float(docHeight)
                            if (docHeightFloat = lastDocHeight) {
                                stableCount++
                                if (stableCount >= 2) {
                                    break
                                }
                            } else {
                                stableCount := 0
                                lastDocHeight := docHeightFloat
                            }
                        }
                    } catch Error as docErr {
                        if (A_Index < docHeightRetries) {
                            Sleep(200)
                        }
                    }
                    if (A_Index < docHeightRetries) {
                        Sleep(200)
                    }
                }

                if (docHeight != "" && docHeight != "undefined" && docHeight != "null") {
                    docHeightFloat := Float(docHeight)
                    if (docHeightFloat > 0) {
                        targetScrollY := savedPercentage * docHeightFloat
                        try {
                            scrollCommand := "window.scrollTo({top: " . targetScrollY . ", behavior: 'instant'});"
                            uia.JSExecute(scrollCommand)
                            Sleep(500)

                            ; Verify scroll position
                            verificationRetries := 3
                            loop verificationRetries {
                                try {
                                    actualScrollY := uia.JSReturnThroughClipboard("window.pageYOffset")
                                    actualScrollYFloat := Float(actualScrollY)
                                    if (Abs(actualScrollYFloat - targetScrollY) <= 5.0) {
                                        break
                                    }
                                } catch {
                                }
                                if (A_Index < verificationRetries) {
                                    Sleep(200)
                                }
                            }

                            StandardLoadingBar_Update("Scroll position restored!")
                        } catch Error as scrollErr {
                            StandardLoadingBar_Update("Error: Scroll failed")
                        }
                    }
                } else {
                    StandardLoadingBar_Update("Error: Page not ready")
                }

                AL_RemoveInputGuard()
                Send("{F11}")
                Sleep(300)
                StandardLoadingBar_Hide(1000)
            }
        } else {
            ; No saved position, just enter fullscreen
            Send("{F11}")
            Sleep(200)
        }

        ; Enable focus mode to darken other monitors
        EnableFocusMode()

        ; Start monitoring Wikipedia focus for automatic blackout cancellation
        StartWikipediaFocusMonitor()
    } else {
        ShowWikipediaSelector()
    }
}

; =============================================================================
; Open/Activate Wikipedia
; Hotkey: Win+Alt+Shift+K
; =============================================================================
#!+k::
{
}
