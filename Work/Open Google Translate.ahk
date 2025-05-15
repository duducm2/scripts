#Requires AutoHotkey v2

#include C:\Users\fie7ca\Documents\UIA-v2\Lib\UIA.ahk
#include C:\Users\fie7ca\Documents\UIA-v2\Lib\UIA_Browser.ahk

CapsLock & t::
{

	Run "C:\Program Files\Google\Chrome\Application\chrome.exe --new-window --incognito https://translate.google.com.br/?hl=en&sl=en&tl=pt&op=translate"
	Send("{CapsLock}")
}