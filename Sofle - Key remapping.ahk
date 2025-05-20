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
+y:: SendSymbol("?")   ; Win+Alt+Shift+Y → ?