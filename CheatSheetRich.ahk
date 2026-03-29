#Requires AutoHotkey v2.0+
; RichEdit body for cheat sheet: mnemonics as bold + larger font (no square brackets in display).

global g_cheatSheetRichDll := 0

; #region agent log
CheatSheet_AgentDebugLog(hypothesisId, location, message, data) {
    logPath := A_ScriptDir "\debug-338fbf.log"
    ts := DllCall("GetTickCount64", "int64")
    if !IsObject(data)
        data := Map()
    data["ptrSize"] := A_PtrSize
    inner := ""
    first := true
    for k, v in data {
        inner .= (first ? "" : ",")
        first := false
        inner .= '"' k '":'
        if IsNumber(v)
            inner .= v
        else
            inner .= '"' StrReplace(StrReplace(String(v), "\", "\\"), '"', '\"') '"'
    }
    line := '{"sessionId":"338fbf","timestamp":' ts ',"hypothesisId":"' hypothesisId '","location":"' location '","message":"' message '","data":{' inner '}}`n'
    FileAppend line, logPath, "UTF-8"
}
; #endregion

CheatSheet_EnsureRichDll() {
    global g_cheatSheetRichDll
    if (!g_cheatSheetRichDll)
        g_cheatSheetRichDll := DllCall("LoadLibrary", "str", "riched20.dll", "ptr")
}

; UTF-16 code unit count for RichEdit character indices (BMP = 1, supplementary = 2).
CheatSheet_Utf16Units(s) {
    n := 0
    for c in StrSplit(s, "") {
        o := Ord(c)
        n += (o > 0xFFFF) ? 2 : 1
    }
    return n
}

; EM_SETTEXTEX = 0x461, ST_UNICODE = 8 — RichEdit’s native UTF-16 path (WM_SETTEXT can leave body empty in some hosts).
CheatSheet_RichSetPlainUtf16(ctrl, plain) {
    hwnd := ctrl.Hwnd
    flags := 8 ; ST_UNICODE
    cp := 1200
    settextex := Buffer(8, 0)
    NumPut("uint", flags, settextex, 0)
    NumPut("uint", cp, settextex, 4)
    if (plain = "") {
        emptyBuf := Buffer(2, 0)
        emRet := SendMessage(0x461, settextex.Ptr, emptyBuf.Ptr, hwnd)
        ; #region agent log
        CheatSheet_AgentDebugLog("B", "CheatSheet_RichSetPlainUtf16", "em_settextex_empty", Map("emRet", emRet,
            "hwnd", hwnd))
        ; #endregion
        return
    }
    textBuf := Buffer((StrLen(plain) + 1) * 2)
    StrPut(plain, textBuf, "UTF-16")
    emRet := SendMessage(0x461, settextex.Ptr, textBuf.Ptr, hwnd)
    ; #region agent log
    CheatSheet_AgentDebugLog("B", "CheatSheet_RichSetPlainUtf16", "em_settextex_plain", Map("emRet", emRet,
        "plainLen", StrLen(plain), "hwnd", hwnd))
    ; #endregion
}

; Small WM_GETTEXT sample (fixed cap) — do not use WM_GETTEXTLENGTH on RichEdit (can block).
CheatSheet_RichPeekPrefix(hwnd, maxTchars := 10) {
    cap := maxTchars + 1
    buf := Buffer(cap * 2, 0)
    copied := SendMessage(0xD, cap, buf.Ptr, hwnd)
    s := StrGet(buf.Ptr, maxTchars, "UTF-16")
    return Map("copied", copied, "prefix", SubStr(s, 1, maxTchars))
}

; EM_EXSETSEL = WM_USER + 55 — use CHARRANGE; EM_SETSEL(0,-1) via SendMessage can fail to select whole doc in AHK x64.
CheatSheet_RichExSetSel(hwnd, cpMin, cpMax) {
    cr := Buffer(8, 0)
    NumPut("int", cpMin, cr, 0)
    NumPut("int", cpMax, cr, 4)
    return SendMessage(0x437, 0, cr.Ptr, hwnd) ; 0x400+0x37 = WM_USER+55
}

; EM_SETCHARFORMAT = 0x444, SCF_ALL = 4, SCF_SELECTION = 1
CheatSheet_RichApplyCharFormat(ctrl, scopeAll, cfBuf) {
    w := scopeAll ? 4 : 1
    return SendMessage(0x444, w, cfBuf.Ptr, ctrl.Hwnd)
}

; Build CHARFORMAT2W: default body (yellow, Consolas, base pt) or mnemonic (bold, larger pt).
; Full CHARFORMAT2W is 116 bytes on Win32; cbSize must match or EM_SETCHARFORMAT ignores color/face.
; Layout: yHeight 12, yOffset 16, crTextColor 20, bCharSet 24, szFaceName 26.
CheatSheet_RichCharFormat2(basePt := 12, mnemonicPt := 15, bold := false) {
    yh := bold ? Round(mnemonicPt * 20) : Round(basePt * 20)
    cf := Buffer(116, 0)
    NumPut("uint", 116, cf, 0) ; cbSize = sizeof(CHARFORMAT2)
    ; Always set CFM_BOLD in mask so base (non-bold) clears bold; otherwise SCF_ALL can leave default black runs.
    mask := 0x40000000 | 0x80000000 | 0x20000000 | 0x1 ; CFM_COLOR | CFM_SIZE | CFM_FACE | CFM_BOLD
    NumPut("uint", mask, cf, 4) ; dwMask
    NumPut("uint", bold ? 0x1 : 0, cf, 8) ; dwEffects: CFE_BOLD or clear
    NumPut("int", yh, cf, 12) ; yHeight (twips)
    NumPut("int", 0, cf, 16) ; yOffset
    NumPut("uint", 0x0000FFFF, cf, 20) ; crTextColor yellow (RGB)
    NumPut("uchar", 1, cf, 24) ; bCharSet DEFAULT_CHARSET
    NumPut("uchar", 0, cf, 25) ; bPitchAndFamily
    StrPut("Consolas", cf.Ptr + 26, 64, "UTF-16")
    return cf
}

; Disable visual styles on RichEdit + parent (themed read-only / dark can ignore CHARFORMAT and draw “invisible” text).
CheatSheet_RichThemingOff(ctrl) {
    hwnd := ctrl.Hwnd
    DllCall("uxtheme\SetWindowTheme", "ptr", hwnd, "wstr", "", "wstr", "")
    parent := DllCall("GetParent", "ptr", hwnd, "ptr")
    if (parent)
        DllCall("uxtheme\SetWindowTheme", "ptr", parent, "wstr", "", "wstr", "")
}

; After plain text is set: default format all, then selection-format mnemonic spans.
CheatSheet_RichSetProcessedBody(ctrl, processedText) {
    CheatSheet_RichThemingOff(ctrl)
    if (processedText = "") {
        ; #region agent log
        CheatSheet_AgentDebugLog("A", "CheatSheet_RichSetProcessedBody", "processedText_empty_early_exit", Map("hwnd",
            ctrl.Hwnd))
        ; #endregion
        CheatSheet_RichSetPlainUtf16(ctrl, "")
        SendMessage(0x4CF, 1, 0, ctrl.Hwnd) ; EM_SETREADONLY = WM_USER+207 (control created without ES_READONLY so CHARFORMAT colors apply)
        return
    }
    plain := ""
    spans := [] ; { u16Start, u16Len } mnemonic ranges in UTF-16 units
    u16Pos := 0
    first := true
    for line in StrSplit(processedText, "`n", "`r") {
        if (!first) {
            plain .= "`r`n"
            u16Pos += 2
        }
        first := false
        segs := CheatSheet_ParseProcessedLine(line)
        for seg in segs {
            slen := CheatSheet_Utf16Units(seg.text)
            if (seg.mnemonic)
                spans.Push({ u16Start: u16Pos, u16Len: slen })
            plain .= seg.text
            u16Pos += slen
        }
    }
    ; #region agent log
    CheatSheet_AgentDebugLog("A", "CheatSheet_RichSetProcessedBody", "after_parse", Map("plainLen", StrLen(plain),
    "processedLen", StrLen(processedText), "spanCount", spans.Length, "hwnd", ctrl.Hwnd))
    ; #endregion
    CheatSheet_RichSetPlainUtf16(ctrl, plain)
    hwnd := ctrl.Hwnd
    ; #region agent log
    peekA := CheatSheet_RichPeekPrefix(hwnd)
    CheatSheet_AgentDebugLog("F", "CheatSheet_RichSetProcessedBody", "wmgettext_after_plain", Map("peekCopied",
        peekA["copied"], "peekPrefix", peekA["prefix"]))
    ; #endregion
    ; Do not call WM_GETTEXTLENGTH / EM_GETTEXTLENGTHEX here — on RichEdit20W they can block the UI thread for a long time.
    rc := Buffer(16, 0)
    DllCall("GetClientRect", "ptr", hwnd, "ptr", rc.Ptr)
    clientW := NumGet(rc, 8, "int") - NumGet(rc, 0, "int")
    clientH := NumGet(rc, 12, "int") - NumGet(rc, 4, "int")
    isVis := DllCall("IsWindowVisible", "ptr", hwnd)
    ; #region agent log
    CheatSheet_AgentDebugLog("C", "CheatSheet_RichSetProcessedBody", "after_settext_metrics", Map("expectedPlainLen",
        StrLen(plain), "clientW", clientW, "clientH", clientH, "hwnd", hwnd, "isVisible", isVis))
    ; #endregion
    SendMessage(0x443, 0, 0x000000, ctrl.Hwnd) ; EM_SETBKGNDCOLOR black (before char format so text is not default black on black)
    ; Base yellow on full body: select entire document via EM_EXSETSEL, then SCF_SELECTION (reliable vs EM_SETSEL 0,-1).
    baseCf := CheatSheet_RichCharFormat2(12, 12, false)
    CheatSheet_RichExSetSel(hwnd, 0, -1)
    emCfBase := CheatSheet_RichApplyCharFormat(ctrl, false, baseCf) ; SCF_SELECTION
    ; #region agent log
    rng := Buffer(8, 0)
    SendMessage(0x434, 0, rng.Ptr, hwnd) ; EM_EXGETSEL = WM_USER+52 — verify selection spans document
    CheatSheet_AgentDebugLog("D", "CheatSheet_RichSetProcessedBody", "em_setcharformat_base", Map("emCfBase", emCfBase,
        "cfBufSize", NumGet(baseCf, 0, "uint"), "exSelMin", NumGet(rng, 0, "int"), "exSelMax", NumGet(rng, 4, "int")))
    ; #endregion
    ; Mnemonic spans (bold + slightly larger) — reuse one CHARFORMAT buffer (many spans × alloc was slow).
    mnCf := CheatSheet_RichCharFormat2(12, 15, true)
    for sp in spans {
        if (sp.u16Len <= 0)
            continue
        SendMessage(0xB1, sp.u16Start, sp.u16Start + sp.u16Len, ctrl.Hwnd) ; EM_SETSEL
        CheatSheet_RichApplyCharFormat(ctrl, false, mnCf)
    }
    ; #region agent log
    peekB := CheatSheet_RichPeekPrefix(hwnd)
    CheatSheet_AgentDebugLog("F", "CheatSheet_RichSetProcessedBody", "wmgettext_after_spans", Map("peekCopied",
        peekB["copied"], "peekPrefix", peekB["prefix"]))
    ; #endregion
    ; Collapse selection to start (read-only display), then scroll caret into view + repaint
    SendMessage(0xB1, 0, 0, ctrl.Hwnd)
    SendMessage(0xB7, 0, 0, hwnd) ; EM_SCROLLCARET
    DllCall("InvalidateRect", "ptr", hwnd, "ptr", 0, "int", 1)
    DllCall("RedrawWindow", "ptr", hwnd, "ptr", 0, "ptr", 0, "uint", 0x105) ; RDW_INVALIDATE | RDW_UPDATENOW | RDW_ERASE
    SendMessage(0x4CF, 1, 0, hwnd) ; EM_SETREADONLY TRUE — use message instead of ES_READONLY to avoid gray read-only text on black
}

; Split processed line into display segments (no brackets in output).
CheatSheet_ParseProcessedLine(line) {
    segs := []
    if (line = "")
        return segs
    ; Headers / lines without standard mnemonic structure
    if !InStr(line, "[") {
        segs.Push({ text: line, mnemonic: false })
        return segs
    }
    prefix := ""
    rest := line
    if RegExMatch(line, "^(>>>\s*|---\s*)", &pm) {
        prefix := pm[0]
        rest := SubStr(line, StrLen(pm[0]) + 1)
    }
    if (prefix != "")
        segs.Push({ text: prefix, mnemonic: false })
    if !RegExMatch(rest, "^([^\[\]]*?)(\[.*?\])(.*)$", &m) {
        segs.Push({ text: rest, mnemonic: false })
        return segs
    }
    emojiPart := m[1]
    bracketPart := m[2]
    tail := m[3]
    if (emojiPart != "")
        segs.Push({ text: emojiPart, mnemonic: false })
    ; First column: padded bracket -> inner only, split mnemonic run
    if RegExMatch(bracketPart, "^\[(.+)\]$", &ib) {
        CheatSheet_AppendInnerMnemonicSegments(segs, ib[1])
    } else {
        segs.Push({ text: bracketPart, mnemonic: false })
    }
    if (tail != "")
        CheatSheet_ParseBracketTail(segs, tail)
    return segs
}

CheatSheet_AppendInnerMnemonicSegments(segs, inner) {
    trimmed := Trim(inner, " `t")
    if (trimmed = "") {
        segs.Push({ text: inner, mnemonic: false })
        return
    }
    p := InStr(inner, trimmed, false, , 1)
    if (p > 1)
        segs.Push({ text: SubStr(inner, 1, p - 1), mnemonic: false })
    segs.Push({ text: trimmed, mnemonic: true })
    rest := SubStr(inner, p + StrLen(trimmed))
    if (rest != "")
        segs.Push({ text: rest, mnemonic: false })
}

CheatSheet_ParseBracketTail(segs, tail) {
    pos := 1
    n := StrLen(tail)
    while pos <= n {
        if SubStr(tail, pos, 1) != "[" {
            nxt := InStr(tail, "[", , pos)
            if !nxt
                nxt := n + 1
            segs.Push({ text: SubStr(tail, pos, nxt - pos), mnemonic: false })
            pos := nxt
            continue
        }
        close := InStr(tail, "]", , pos)
        if !close {
            segs.Push({ text: SubStr(tail, pos), mnemonic: false })
            break
        }
        inner := SubStr(tail, pos + 1, close - pos - 1)
        segs.Push({ text: inner, mnemonic: true })
        pos := close + 1
    }
}
