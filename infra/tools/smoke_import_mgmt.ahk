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
variantPath := A_Desktop . "\JOB_SEARCH_UPDATE_updated.txt"
canonicalPath := A_Desktop . "\JOB_SEARCH_UPDATE.txt"
try FileDelete(variantPath)
catch {
}
try FileDelete(canonicalPath)
catch {
}
ImportMgmt_WriteUtf8(variantPath, pack)

ok := ImportMgmt_ImportFromDesktop()
if (!ok) {
    FileAppend("FAIL: import returned false`n", "*")
    ExitApp 1
}

if (!FileExist(canonicalPath)) {
    FileAppend("FAIL: canonical JOB_SEARCH_UPDATE.txt missing after import`n", "*")
    ExitApp 5
}
if (FileExist(variantPath)) {
    FileAppend("FAIL: variant JOB_SEARCH_UPDATE_updated.txt still on Desktop`n", "*")
    ExitApp 6
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
