#Requires AutoHotkey v2.0
#Include Utils.ahk

ImportMgmt_EnsureData()
pack := "===PREVIEW===`n1 row`n===END_PREVIEW===`n`n"
    . "===FILE: JOB_SEARCH_UPDATE.csv===`n"
    . "id,company,role_title,job_url,job_description,status,status_date,applied_date,notes`n"
    . "JOB_PEPSI,Pepsi,Backend Engineer,https://linkedin.com/jobs/123,"
    . "" "Build APIs in Python. 5+ years experience. Remote Brazil." "", applied, 2026 - 08 - 28, , `n "
    . "===END_FILE==="
ImportMgmt_WriteUtf8(A_Desktop . "\JOB_SEARCH_UPDATE_SMOKE.txt", pack)
if (!ImportMgmt_ImportFromDesktop())
    ExitApp 1
for r in ImportMgmt_Load() {
    if (r.Has("id") && r["id"] = "JOB_PEPSI") {
        if (r.Has("source"))
            ExitApp 4
        if (InStr(r["job_description"], "Python"))
            ExitApp 0
        ExitApp 2
    }
}
ExitApp 3