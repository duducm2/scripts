; Step-by-step Chrome tab-detach diagnostics. Focus a Chrome window, then press Ctrl+Alt+Shift+D.
; Does NOT detach by default — logs each stage via ToolTip. Press Ctrl+Alt+Shift+T to run full detach.

#Requires AutoHotkey v2.0

#SingleInstance Force

#Include %A_ScriptDir%\..\..\vendor\UIA-v2\Lib\UIA.ahk
#Include %A_ScriptDir%\..\..\vendor\UIA-v2\Lib\UIA_Browser.ahk
#Include %A_ScriptDir%\..\..\Utils.ahk

ChromeDetachDebug_Tip(msg, ms := 3500) {
    ToolTip msg, 10, 10
    SetTimer(() => ToolTip(), -ms)
}

ChromeDetachDebug_RunSteps(runDetach := false) {
    hwnd := WinExist("A")
    if !hwnd || WinGetProcessName("ahk_id " hwnd) != "chrome.exe" {
        ChromeDetachDebug_Tip("❌ Active window is not Chrome")
        return
    }

    session := Chrome_DetachSessionCreate(hwnd)
    if !IsObject(session.uia) {
        ChromeDetachDebug_Tip("❌ UIA_Browser failed to attach")
        return
    }

    tab := Chrome_DetachGetActiveTab(session)
    if !Chrome_IsValidTabElement(tab) {
        ChromeDetachDebug_Tip("❌ Active tab not found via UIA`nTitle: " Chrome_DetachGetWindowTitleForMatch(hwnd))
        return
    }

    try tabName := tab.Name
    catch {
        tabName := "(unknown)"
    }
    try tabType := tab.Type
    catch {
        tabType := "?"
    }
    ChromeDetachDebug_Tip("✅ Tab found`nType: " tabType "`nName: " SubStr(tabName, 1, 80), 4500)
    Sleep 4500

    session.baselinePopups := Chrome_DetachListMenuPopups()
    session.menuPopupHwnd := 0
    opened := Chrome_OpenActiveTabContextMenu(session)
    if !opened {
        ChromeDetachDebug_Tip("❌ Could not open tab context menu")
        Chrome_ContextMenuDismiss()
        return
    }

    isTabMenu := Chrome_ContextMenuLooksLikeTabMenu(session)
    popupInfo := session.menuPopupHwnd ? ("popup hwnd=" session.menuPopupHwnd) : "no popup hwnd"
    ChromeDetachDebug_Tip((isTabMenu ? "✅ Tab menu detected`n" : "❌ Page/other menu detected`n") popupInfo, 4500)
    Sleep 4500

    if !isTabMenu {
        Chrome_ContextMenuDismiss()
        return
    }

    flat := Chrome_ContextMenuFindFirst(session, CHROME_DETACH_MENU_EN_NAMES, false)
    parent := Chrome_ContextMenuFindFirst(session, CHROME_DETACH_MENU_PARENT_NAMES, true)
    menuInfo := flat ? "flat detach item found" : (parent ? "parent submenu found" :
        "no detach item (keyboard fallback)")
    ChromeDetachDebug_Tip("Menu scan: " menuInfo, 3500)
    Sleep 3500

    Chrome_ContextMenuDismiss()
    Sleep 200

    if !runDetach {
        ChromeDetachDebug_Tip("Done (menu closed). Ctrl+Alt+Shift+T = full detach test.", 4000)
        return
    }

    ok := Chrome_DetachActiveTabToNewWindow()
    ChromeDetachDebug_Tip(ok ? "✅ Full detach succeeded" : "❌ Full detach failed", 4000)
}

^!+d:: ChromeDetachDebug_RunSteps(false)
^!+t:: ChromeDetachDebug_RunSteps(true)

ChromeDetachDebug_Tip("Chrome Detach Debug loaded.`nCtrl+Alt+Shift+D = diagnose`nCtrl+Alt+Shift+T = full detach", 5000)