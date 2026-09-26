#Requires AutoHotkey v2.0+
#SingleInstance Force
#Include env.ahk
#Include %A_ScriptDir%\Utils\git_cli.ahk
#Include %A_ScriptDir%\Utils\standard_loading_bar.ahk
Utils_EnsureGlobalEscapeHotkey() {
}
#Include %A_ScriptDir%\Utils\clip_angel_act_bootstrap.ahk
p := GetQuickLookExePath()
out := A_ScriptDir "\.cursor\_act_boot_out.txt"
try FileDelete(out)
FileAppend("ok bootstrap`nql=[" . p . "]`nexist=" . ((p != "" && FileExist(p)) ? "1" : "0") . "`n", out)
ExitApp
