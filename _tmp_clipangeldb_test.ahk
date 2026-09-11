#Requires AutoHotkey v2.0
#include %A_ScriptDir%\Utils\clip_angel_db.ahk
f := A_Temp "\clipangeldb_ahk_test.txt"
try FileDelete(f)
exe := ClipAngelDb_ExePath()
maxId := ClipAngelDb_MaxId()
newest := ClipAngelDb_Newest()
payload := ClipAngelDb_MergePayload()
isfav := ""
if (maxId != "" && IsInteger(maxId))
    isfav := ClipAngelDb_IsFav(Integer(maxId))
lines := "exe=" exe "`n"
lines .= "maxId=" maxId "`n"
if newest
    lines .= "newest=" newest["id"] "`t" newest["type"] "`t" newest["favorite"] "`t" newest["title"] "`n"
else
    lines .= "newest=FAIL`n"
lines .= "isfav=" isfav "`n"
if payload
    lines .= "merge=" payload["boundaryId"] "`t" payload["total"] "`t" payload["skipped"] "`tlen=" StrLen(payload["text"]) "`n"
else
    lines .= "merge=FAIL`n"
; Also capture raw run stderr path for debug if fail
if (maxId = "") {
    stamp := "debug"
    outFile := A_Temp "\clipangeldb_debug_out.txt"
    errFile := A_Temp "\clipangeldb_debug_err.txt"
    full := A_ComSpec ' /c ""' exe '" maxid >"' outFile '" 2>"' errFile '""'
    ec := RunWait(full, A_ScriptDir, "Hide")
    lines .= "debug_ec=" ec "`n"
    try lines .= "debug_out=" FileRead(outFile) "`n"
    try lines .= "debug_err=" FileRead(errFile) "`n"
}
FileAppend(lines, f, "UTF-8")
ExitApp
