#Requires AutoHotkey v2.0
#Include Utils.ahk

ImportMgmt_EnsureData()
rows := ImportMgmt_Load()
if (!rows.Length) {
    rows.Push(Map(
        "id", "JOB_COCACOLA",
        "company", "Coca Cola",
        "role_title", "Engineer",
        "status", "applied",
        "status_date", "2026-08-01",
        "applied_date", "2026-08-01",
        "source", "LinkedIn",
        "notes", ""
    ))
    ImportMgmt_Save(rows)
}

pack := "===PREVIEW===`n1 row · Coca Cola rejected`n===END_PREVIEW===`n`n"
    . "===FILE: JOB_SEARCH_UPDATE.csv===`n"
    . "id,company,role_title,status,status_date,applied_date,source,notes`n"
    . "JOB_COCACOLA,Coca Cola,,rejected,2026-08-27,,,Unsuccessful`n"
    . "===END_FILE==="
testPath := A_Desktop . "\JOB_SEARCH_UPDATE_SMOKE.txt"
ImportMgmt_WriteUtf8(testPath, pack)

ok := ImportMgmt_ImportFromDesktop()
if (!ok) {
    FileAppend("FAIL: import returned false`n", "*")
    ExitApp 1
}

after := ImportMgmt_Load()
found := false
for r in after {
    if (r.Has("id") && r["id"] = "JOB_COCACOLA") {
        found := true
        if (r["status"] != "rejected") {
            FileAppend("FAIL: status=" . r["status"] . "`n", "*")
            ExitApp 2
        }
        if (Trim(r["notes"]) != "Unsuccessful") {
            FileAppend("FAIL: notes=" . r["notes"] . "`n", "*")
            ExitApp 3
        }
    }
}
if (!found) {
    FileAppend("FAIL: row not found`n", "*")
    ExitApp 4
}
FileAppend("OK`n", "*")
ExitApp 0