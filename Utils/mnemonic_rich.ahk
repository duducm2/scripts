; =============================================================================
; Utils module: mnemonic_rich.ahk
; RichEdit helpers (mnemonic emphasis for selector-style ListView modals).
; =============================================================================

global g_MnemonicRichDll := 0

MnemonicRich_EnsureDll() {
    global g_MnemonicRichDll
    if (!g_MnemonicRichDll)
        g_MnemonicRichDll := DllCall("LoadLibrary", "str", "msftedit.dll", "ptr")
}

MnemonicRich_Utf16Units(s) {
    n := 0
    for c in StrSplit(s, "") {
        o := Ord(c)
        n += (o > 0xFFFF) ? 2 : 1
    }
    return n
}

MnemonicRich_SetPlainUtf16(ctrl, plain) {
    hwnd := ctrl.Hwnd
    flags := 8
    cp := 1200
    settextex := Buffer(8, 0)
    NumPut("uint", flags, settextex, 0)
    NumPut("uint", cp, settextex, 4)
    if (plain = "") {
        emptyBuf := Buffer(2, 0)
        SendMessage(0x461, settextex.Ptr, emptyBuf.Ptr, hwnd)
        return
    }
    textBuf := Buffer((StrLen(plain) + 1) * 2)
    StrPut(plain, textBuf, "UTF-16")
    SendMessage(0x461, settextex.Ptr, textBuf.Ptr, hwnd)
}

MnemonicRich_ThemingOff(ctrl) {
    hwnd := ctrl.Hwnd
    DllCall("uxtheme\SetWindowTheme", "ptr", hwnd, "wstr", "", "wstr", "")
    parent := DllCall("GetParent", "ptr", hwnd, "ptr")
    if (parent)
        DllCall("uxtheme\SetWindowTheme", "ptr", parent, "wstr", "", "wstr", "")
}

MnemonicRich_CharFormat2(faceName, pt, textColor, bold := false) {
    yh := Round(pt * 20)
    cf := Buffer(116, 0)
    NumPut("uint", 116, cf, 0)
    mask := 0x40000000 | 0x80000000 | 0x20000000 | 0x1
    NumPut("uint", mask, cf, 4)
    NumPut("uint", bold ? 0x1 : 0, cf, 8)
    NumPut("int", yh, cf, 12)
    NumPut("int", 0, cf, 16)
    NumPut("uint", textColor, cf, 20)
    NumPut("uchar", 1, cf, 24)
    NumPut("uchar", 0, cf, 25)
    StrPut(faceName, cf.Ptr + 26, 64, "UTF-16")
    return cf
}

MnemonicRich_ApplyCharFormat(ctrl, scopeAll, cfBuf) {
    w := scopeAll ? 4 : 1
    return SendMessage(0x444, w, cfBuf.Ptr, ctrl.Hwnd)
}

MnemonicRich_SetSel(hwnd, cpMin, cpMax) {
    return SendMessage(0xB1, cpMin, cpMax, hwnd)
}

MnemonicRich_Render(ctrl, lines, basePt, bumpPx := 6, faceName := "Segoe UI", rgbHex := "CDD6F4", bgHex := "1E1E2E") {
    MnemonicRich_EnsureDll()
    MnemonicRich_ThemingOff(ctrl)

    rr := Integer("0x" . SubStr(rgbHex, 1, 2))
    gg := Integer("0x" . SubStr(rgbHex, 3, 2))
    bb := Integer("0x" . SubStr(rgbHex, 5, 2))
    textColor := (bb << 16) | (gg << 8) | rr

    br := Integer("0x" . SubStr(bgHex, 1, 2))
    bg := Integer("0x" . SubStr(bgHex, 3, 2))
    bb2 := Integer("0x" . SubStr(bgHex, 5, 2))
    bgColor := (bb2 << 16) | (bg << 8) | br

    bumpPt := bumpPx * 72 / 96
    bigPt := basePt + bumpPt

    plain := ""
    spans := []
    subsectionSpans := []
    u16Pos := 0
    first := true

    RenderTitleKey(lineText, key, baseU16) {
        if (key = "")
            return
        rb := InStr(lineText, "]")
        if (!rb)
            return
        after := SubStr(lineText, rb + 1)
        tpos := InStr(after, key, false)
        if (!tpos)
            tpos := InStr(after, StrUpper(key), false)
        if (!tpos)
            return
        preNoLast := SubStr(lineText, 1, rb + tpos - 1)
        spans.Push({ start: baseU16 + MnemonicRich_Utf16Units(preNoLast), len: 1 })
    }

    for ln in lines {
        if (!first) {
            plain .= "`r"
            u16Pos += 1
        }
        first := false

        lineText := ln.text
        lineStartU16 := u16Pos
        key := ln.HasProp("key") ? ln.key : ""
        RenderTitleKey(lineText, key, u16Pos)

        if (ln.HasProp("keyRight") && ln.keyRight != "" && ln.HasProp("rightStartCharPos") && ln.rightStartCharPos > 1) {
            rightStart := ln.rightStartCharPos
            prefix := SubStr(lineText, 1, rightStart - 1)
            rightText := SubStr(lineText, rightStart)
            baseRightU16 := u16Pos + MnemonicRich_Utf16Units(prefix)
            RenderTitleKey(rightText, ln.keyRight, baseRightU16)
        }

        plain .= lineText
        u16Pos += MnemonicRich_Utf16Units(lineText)
        if (ln.HasProp("isMnemonicSubsection") && ln.isMnemonicSubsection && lineText != "") {
            subsectionSpans.Push({ start: lineStartU16, len: MnemonicRich_Utf16Units(lineText) })
        }
    }

    MnemonicRich_SetPlainUtf16(ctrl, plain)

    hwnd := ctrl.Hwnd
    SendMessage(0x4CF, 0, 0, hwnd)
    SendMessage(0x443, 0, bgColor, hwnd)

    baseCf := MnemonicRich_CharFormat2(faceName, basePt, textColor, false)
    MnemonicRich_SetSel(hwnd, 0, -1)
    MnemonicRich_ApplyCharFormat(ctrl, false, baseCf)

    if (subsectionSpans.Length > 0) {
        subRgb := "CBA6F7"
        srr := Integer("0x" . SubStr(subRgb, 1, 2))
        sgg := Integer("0x" . SubStr(subRgb, 3, 2))
        sbb := Integer("0x" . SubStr(subRgb, 5, 2))
        subColor := (sbb << 16) | (sgg << 8) | srr
        subCf := MnemonicRich_CharFormat2(faceName, basePt + 1, subColor, true)
        for ss in subsectionSpans {
            if (ss.len <= 0)
                continue
            MnemonicRich_SetSel(hwnd, ss.start, ss.start + ss.len)
            MnemonicRich_ApplyCharFormat(ctrl, false, subCf)
        }
    }

    bigCf := MnemonicRich_CharFormat2(faceName, bigPt, textColor, false)
    for sp in spans {
        if (sp.len <= 0)
            continue
        MnemonicRich_SetSel(hwnd, sp.start, sp.start + sp.len)
        MnemonicRich_ApplyCharFormat(ctrl, false, bigCf)
    }
    MnemonicRich_SetSel(hwnd, 0, 0)
    SendMessage(0xB7, 0, 0, hwnd)
    SendMessage(0x4CF, 1, 0, hwnd)
    SendMessage(0x443, 0, bgColor, hwnd)
}
