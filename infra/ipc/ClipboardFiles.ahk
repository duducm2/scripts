; CF_HDROP (15) clipboard helpers for file attachment paste/copy flows.

Clipboard_HasFileDrop() {
    try {
        return !!DllCall("IsClipboardFormatAvailable", "UInt", 15, "Int") ; CF_HDROP
    } catch {
        return false
    }
}

Clipboard_WaitForFileDrop(timeoutMs := 800) {
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        if (Clipboard_HasFileDrop())
            return true
        Sleep 50
    }
    return Clipboard_HasFileDrop()
}

Clipboard_NormalizeFilePath(path) {
    if (path = "")
        return ""
    normalized := Trim(Trim(path), Chr(34))
    normalized := StrReplace(normalized, "/", "\")
    return StrLower(RTrim(normalized, "\"))
}

Clipboard_PathIsExistingFile(path) {
    if (path = "")
        return false
    attr := FileExist(path)
    return (attr && !InStr(attr, "D"))
}

Clipboard_GetFilePaths() {
    paths := []
    opened := false
    try {
        if !DllCall("OpenClipboard", "Ptr", 0, "Int")
            return paths
        opened := true
        hDrop := DllCall("GetClipboardData", "UInt", 15, "Ptr") ; CF_HDROP
        if (!hDrop)
            return paths
        count := DllCall("shell32\DragQueryFileW", "Ptr", hDrop, "UInt", 0xFFFFFFFF, "Ptr", 0, "UInt", 0, "UInt")
        loop count {
            idx := A_Index - 1
            chars := DllCall("shell32\DragQueryFileW", "Ptr", hDrop, "UInt", idx, "Ptr", 0, "UInt", 0, "UInt")
            if (chars <= 0)
                continue
            buf := Buffer((chars + 1) * 2, 0)
            if DllCall("shell32\DragQueryFileW", "Ptr", hDrop, "UInt", idx, "Ptr", buf, "UInt", chars + 1, "UInt")
                paths.Push(StrGet(buf, "UTF-16"))
        }
    } catch {
    } finally {
        if (opened) {
            try DllCall("CloseClipboard")
        }
    }
    return paths
}

Clipboard_ContainsFilePath(expectedPath) {
    expected := Clipboard_NormalizeFilePath(expectedPath)
    if (expected = "")
        return false
    for path in Clipboard_GetFilePaths() {
        if (Clipboard_NormalizeFilePath(path) = expected)
            return true
    }
    return false
}

Clipboard_SetFiles(paths) {
    if (!paths || paths.Length = 0)
        return false

    dropFilesOffset := 20
    totalChars := 1 ; final extra NUL after the NUL-terminated file list.
    for path in paths {
        if (path = "")
            return false
        totalChars += StrLen(path) + 1
    }

    hMem := 0
    opened := false
    transferred := false
    try {
        hMem := DllCall("GlobalAlloc", "UInt", 0x42, "UPtr", dropFilesOffset + (totalChars * 2), "Ptr")
        if (!hMem)
            return false
        pMem := DllCall("GlobalLock", "Ptr", hMem, "Ptr")
        if (!pMem)
            return false

        NumPut("UInt", dropFilesOffset, pMem, 0) ; DROPFILES.pFiles
        NumPut("Int", 1, pMem, 16) ; DROPFILES.fWide
        charOffset := 0
        for path in paths {
            StrPut(path, pMem + dropFilesOffset + (charOffset * 2), StrLen(path) + 1, "UTF-16")
            charOffset += StrLen(path) + 1
        }
        DllCall("GlobalUnlock", "Ptr", hMem)

        if !DllCall("OpenClipboard", "Ptr", 0, "Int")
            return false
        opened := true
        if !DllCall("EmptyClipboard", "Int")
            return false
        if !DllCall("SetClipboardData", "UInt", 15, "Ptr", hMem, "Ptr")
            return false
        transferred := true
        hMem := 0
        return true
    } catch {
        return false
    } finally {
        if (opened) {
            try DllCall("CloseClipboard")
        }
        if (hMem && !transferred) {
            try DllCall("GlobalFree", "Ptr", hMem)
        }
    }
}
