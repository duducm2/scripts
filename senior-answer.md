Ranked Hypotheses
1. +Multi is ignored on Custom AHK controls (The "Single-Line" Bug)

Why: In AHK, the +Multi option translates to the ES_MULTILINE (0x0004) window style. However, AHK only translates this keyword for built-in Edit controls. For Custom controls, AHK passes styles generically, meaning +Multi is completely ignored.
Impact: Your ClassRichEdit20W is being instantiated as a single-line edit control. Single-line edit controls either truncate text at the first newline or display everything superimposed on a single line. This perfectly matches your symptom: “sometimes a single line of yellow text... Most of the body looks empty” while WM_GETTEXT proves the buffer actually holds the whole document.
2. CRLF vs CR Character Indexing Drift (The Formatting Bug)

Why: In CheatSheet_RichSetProcessedBody, you join lines with plain .= "\r\n" and increment u16Pos += 2. When you use EM_SETTEXTEX, RichEdit internally normalizes \r\n to a single \r (1 character).
Impact: Your EM_SETSEL indices (u16Pos) drift by +1 character for every line break. By line 10, your bold/yellow spans are shifted right by 10 characters, formatting the wrong text or invisible spaces.
3. Antiquated RichEdit DLL Version

Why: riched20.dll (RichEdit20W) is the legacy Windows XP-era RichEdit 3.0. Modern Windows uses msftedit.dll (RICHEDIT50W), which handles High DPI, Unicode, and background/theme coloring far more predictably.
Concrete Next Experiments (Execute in this order)
Experiment 1: Force ES_MULTILINE and upgrade to RichEdit 5.0
Change the initialization to explicitly pass the hex codes for ES_MULTILINE (0x0004) and ES_AUTOVSCROLL (0x0040), combining to +0x44.

In Shift keys.ahk (Lines ~255 and ~278):
 g_helpCheatCtrl := g_helpGui.Add("Custom", "ClassRICHEDIT50W xs+0 y+4 +0x44 -E0x200 +VScroll -HScroll -Border Background000000 w1000 r12")

 g_globalCheatCtrl := g_globalGui.Add("Custom", "ClassRICHEDIT50W xs+0 y+4 +0x44 +VScroll -HScroll -Border Background000000 w1000 r12")

In CheatSheetRich.ahk (Line ~16):
 g_cheatSheetRichDll := DllCall("LoadLibrary", "str", "msftedit.dll", "ptr")

Experiment 2: Fix the Character Count Drift
Prevent your formatting spans from drifting by feeding RichEdit the exact newline character it expects.

In CheatSheetRich.ahk (Line ~102):

diff
-2
+2
   for line in StrSplit(processedText, "`n", "`r") {
       if (!first) {
           plain .= "`r"
           u16Pos += 1
       }

Experiment 3: Verify Row Height Parsing on Custom Controls
If the body is still black, check if AHK is failing to calculate the height of r12 on a custom control. Because AHK doesn't know the default font metrics of a Custom control, r12 might calculate to a tiny height initially.

Test: Add a debug log right after g_helpCheatCtrl.GetPos(,, &w, &h) in your CheatSheet_ResizeBody to ensure h isn't 0 or 12 pixels.
Risks and Observations in the Current Code
CHARFORMAT2W Size on x64: Your 116-byte size is 100% correct. The structure contains no pointers (COLORREF, LONG, and DWORD are all 4 bytes on both 32-bit and 64-bit Windows, and WCHAR szFaceName[32] takes 64 bytes). So cbSize = 116 is safe.
EM_EXSETSEL(0, -1): Using NumPut("int", -1, cr, 4) is correct. It writes 0xFFFFFFFF into the cpMax field. RichEdit interprets this as "end of document".
EM_SETREADONLY placement: You correctly apply this message after formatting. Be aware that on older RichEdit versions, switching to Read-Only sometimes causes the control to ignore EM_SETBKGNDCOLOR and revert to a gray background. If you see gray backgrounds after upgrading to RICHEDIT50W, apply EM_SETBKGNDCOLOR after setting Read-Only.
Theming disable: SetWindowTheme(hwnd, "", "") is a great defensive move, as dark mode Explorer themes love to hijack RichEdit scrollbars and text colors. Keep this.