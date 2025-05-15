/********************************************************************
 *   Win+Alt+Shift symbol layer shortcuts (AHK v2)
 *   • Provides system-wide symbol shortcuts
 ********************************************************************/

#Requires AutoHotkey v2.0+

; Function to send symbol characters
SendSymbol(sym) {
    SendText(sym)
}

; Symbol shortcuts using Win+Alt+Shift combinations
#!+y:: SendSymbol("?")   ; Win+Alt+Shift+Y → ?
#!+u:: SendSymbol("[")   ; Win+Alt+Shift+U → [
#!+i:: SendSymbol("]")   ; Win+Alt+Shift+I → ]
#!+o:: SendSymbol("|")   ; Win+Alt+Shift+O → |
#!+h:: SendSymbol("/")   ; Win+Alt+Shift+H → /
#!+j:: SendSymbol("\")   ; Win+Alt+Shift+J → \
