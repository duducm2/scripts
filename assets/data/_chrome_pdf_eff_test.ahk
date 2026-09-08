#Requires AutoHotkey v2.0
#SingleInstance Force
#Include C:\Users\eduev\Meu Drive\17 - Projects\scripts\vendor\UIA-v2\Lib\UIA.ahk
#Include C:\Users\eduev\Meu Drive\17 - Projects\scripts\vendor\UIA-v2\Lib\UIA_Browser.ahk
#Include C:\Users\eduev\Meu Drive\17 - Projects\scripts\Shift keys\predicates_chrome_pdf.ahk

outPath := "C:\Users\eduev\Meu Drive\17 - Projects\scripts\assets\data\_chrome_pdf_eff_test.txt"
try FileDelete outPath

ChromePdf_GetViewerRoot(uia) {
    if (!uia)
        return 0
    root := 0
    try root := uia.FindElement({ Type: 50030, Value: "chrome-extension://mhjfbmdgcfjbbpaeojofohoefgiehjai",
        matchmode: "Substring" })
    if (root)
        return root
    try root := uia.GetCurrentDocumentElement()
    if (root)
        return root
    try root := uia.BrowserElement
    return root
}
ChromePdf_FindByAutomationId(root, automationId, typeHint := 0) {
    if (!root || automationId = "")
        return 0
    el := 0
    if (typeHint) {
        try el := root.FindFirst({ Type: typeHint, AutomationId: automationId })
    }
    if (!el)
        try el := root.FindFirst({ Type: 50000, AutomationId: automationId })
    if (!el)
        try el := root.FindFirst({ AutomationId: automationId })
    return el
}

lines := []
chromeActive := WinActive("ahk_exe chrome.exe") ? 1 : 0
lines.Push("chromeActive=" chromeActive)
t0 := A_TickCount
r1 := IsChromePdfViewerActive()
t1 := A_TickCount - t0
lines.Push("pred1=" r1 " ms=" t1)
t0 := A_TickCount
r2 := IsChromePdfViewerActive()
t2 := A_TickCount - t0
lines.Push("pred2_cached=" r2 " ms=" t2)

chromeList := WinGetList("ahk_exe chrome.exe")
lines.Push("chromeWindows=" chromeList.Length)
pdfFound := 0
for hwnd in chromeList {
    try {
        uia := UIA_Browser("ahk_id " hwnd)
        if (uia.FindElement({ Type: 50030, Value: "chrome-extension://mhjfbmdgcfjbbpaeojofohoefgiehjai",
            matchmode: "Substring" })) {
            pdfFound := 1
            lines.Push("pdfHwnd=" hwnd)
            root := ChromePdf_GetViewerRoot(uia)
            lines.Push("root=" (root ? 1 : 0))
            for id in ["fit", "pageSelector", "sidenavToggle", "more"] {
                typeHint := (id = "pageSelector") ? 50004 : 50000
                el := ChromePdf_FindByAutomationId(root, id, typeHint)
                lines.Push(id "=" (el ? 1 : 0))
            }
            btn := 0
            try btn := root.FindFirst({ Type: 50000, Name: "Baixar", cs: false })
            if (!btn)
                try btn := root.FindFirst({ Type: 50000, Name: "Download", cs: false })
            lines.Push("download=" (btn ? 1 : 0))
            break
        }
    } catch as err {
        lines.Push("probeErr=" err.Message)
    }
}
lines.Push("pdfFound=" pdfFound)

text := ""
for line in lines
    text .= line "`n"
FileAppend text, outPath
ExitApp
