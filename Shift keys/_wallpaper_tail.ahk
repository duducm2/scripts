; Shift + B : Set selected image as desktop background (locale-flexible)
+b::
{
    Explorer_SetAsDesktopBackground()
}

Explorer_WaitShellContextMenuHwnd(timeoutMs := 2500) {
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        for hwnd in WinGetList("ahk_class #32768") {
            if hwnd
                return hwnd
        }
        Sleep 40
    }
    return 0
}

Explorer_MenuItemMatchesNeedles(name, needles) {
    if (name = "")
        return false
    nameLower := StrLower(name)
    for needle in needles {
        if InStr(nameLower, StrLower(needle))
            return true
    }
    return false
}

Explorer_InvokeMenuItem(menuItemEl) {
    if !IsObject(menuItemEl)
        return false
    try {
        if menuItemEl.GetPropertyValue(UIA.Property.IsInvokePatternAvailable) {
            menuItemEl.InvokePattern.Invoke()
            return true
        }
    } catch {
    }
    try {
        menuItemEl.Click()
        return true
    } catch {
    }
    return false
}

Explorer_FindWallpaperMenuItemInPopup(menuHwnd, needles) {
    if !menuHwnd
        return 0
    try {
        menuRoot := UIA.ElementFromHandle(menuHwnd)
        ; Shell popups usually expose MenuItem; also check ListItem/Button for Win11 variants.
        for typeName in ["MenuItem", "ListItem", "Button"] {
            try {
                items := menuRoot.FindAll({ ControlType: typeName })
            } catch {
                continue
            }
            for item in items {
                n := ""
                try n := item.Name
                if (n != "" && Explorer_MenuItemMatchesNeedles(n, needles))
                    return item
            }
        }
    } catch {
    }
    return 0
}

Explorer_SetAsDesktopBackground() {
    ; Classic context menu: locate #32768 popup and invoke locale-flexible wallpaper item.
    needles := [
        "desktop background",
        "as wallpaper",
        "set as background",
        "fundo da área de trabalho",
        "papel de parede",
        "como fundo",
        "fondo de escritorio",
        "fondo del escritorio",
        "desktophintergrund"
    ]

    StandardLoadingBar_Show("⏳ Preparing set as background...", BANNER_ACCENT_INTERMEDIATE, { passive: false })

    try {
        StandardLoadingBar_Update("📁 Ensuring focus on items view...", BANNER_ACCENT_INTERMEDIATE)
        EnsureItemsViewFocus()
        Sleep 300

        StandardLoadingBar_Update("📋 Opening classic context menu...", BANNER_ACCENT_INTERMEDIATE)
        ; Shift+F10 opens the classic shell menu where wallpaper lives (Win11 AppsKey often does not).
        Send "+{F10}"

        menuHwnd := Explorer_WaitShellContextMenuHwnd(2500)
        if !menuHwnd {
            ; Fallback: AppsKey then wait again (some Explorer builds differ).
            Send "{AppsKey}"
            menuHwnd := Explorer_WaitShellContextMenuHwnd(2000)
        }

        if !menuHwnd {
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ Could not open Explorer context menu", 2000, BANNER_ACCENT_ERROR)
            return
        }

        StandardLoadingBar_Update("🔍 Locating set-as-background option...", BANNER_ACCENT_INTERMEDIATE)
        menuItem := Explorer_FindWallpaperMenuItemInPopup(menuHwnd, needles)
        if !menuItem {
            Send "{Esc}"
            StandardLoadingBar_Hide(0)
            ShowCenteredOverlay_Utils("❌ Could not find set-as-background menu item", 2000, BANNER_ACCENT_ERROR)
            return
        }

        StandardLoadingBar_Update("✅ Setting as desktop background...", BANNER_ACCENT_SUCCESS)
        if !Explorer_InvokeMenuItem(menuItem) {
            try menuItem.SetFocus()
            Sleep 50
            Send "{Enter}"
        }
        Sleep 500
        StandardLoadingBar_Hide(500)
    } catch as e {
        try Send "{Esc}"
        StandardLoadingBar_Hide(0)
        ShowCenteredOverlay_Utils("❌ Set as background error: " . e.Message, 2000, BANNER_ACCENT_ERROR)
    }
}
