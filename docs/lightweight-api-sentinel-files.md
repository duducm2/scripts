# Lightweight API sentinel files

This document records how to create and use **empty sentinel files** in the user **Documents** folder as local markers for lightweight APIs and cross-process workflows. It complements remote APIs (for example Google Apps Script in `StudyLinkHelpers.ahk`) and repo-local sentinels under `A_ScriptDir` (see [focus-mode-secondary-monitor-blackout.md](focus-mode-secondary-monitor-blackout.md)).

---

## Canonical sentinel: Manage Study Subtopic Link

| Property         | Value                                                                                                  |
| ---------------- | ------------------------------------------------------------------------------------------------------ |
| **Feature**      | Manage Study Subtopic Link (`Utils.ahk`, Study Topic selector option `[3]`)                            |
| **Filename**     | `manage, study, set, top, link` (comma-separated tokens from the UI label; no extension)               |
| **Location**     | User Documents folder (`A_MyDocuments` in AHK; `%UserProfile%\Documents\` on typical Windows profiles) |
| **Content**      | Empty (0 bytes)                                                                                        |
| **Remote state** | `StudyLink_Get` / `StudyLink_Set` via `StudyLinkHelpers.ahk` (HTTP to Apps Script, key `subtopic`)     |

The sentinel is a **local presence flag** only. It does not store the YouTube URL; the URL lives in the remote API. Scripts may check `FileExist` on this path in the future without a network round-trip.

### Environment-specific Documents paths

Resolve the folder at runtime; do not hardcode `OneDrive\Documentos`.

| Environment | Typical Documents path                                             |
| ----------- | ------------------------------------------------------------------ |
| Personal    | `C:\Users\eduev\OneDrive\Documentos\` (via `A_MyDocuments`)        |
| Work        | `C:\Users\fie7ca\Documents\` (via `A_MyDocuments` on work profile) |

Detection of work vs personal for scripts: [`env.ahk`](../env.ahk) (`IS_WORK_ENVIRONMENT`).

---

## Creation procedure

Always treat the filename as a **single literal path**. The commas are part of the name, not field separators.

### 1. Check existence

**PowerShell:**

```powershell
$path = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'manage, study, set, top, link'
Test-Path -LiteralPath $path
```

**AutoHotkey v2:**

```ahk
path := A_MyDocuments "\manage, study, set, top, link"
exists := FileExist(path)
```

### 2. Create empty file (only if missing)

**PowerShell** (use .NET when `New-Item` does not accept `-LiteralPath` on your shell version):

```powershell
$path = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'manage, study, set, top, link'
if (-not [System.IO.File]::Exists($path)) {
    $fs = [System.IO.File]::Create($path)
    $fs.Close()
}
```

Alternative:

```powershell
$path = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'manage, study, set, top, link'
if (-not (Test-Path -LiteralPath $path)) {
    New-Item -ItemType File -Path $path -Force | Out-Null
}
```

**AutoHotkey v2:**

```ahk
StudyLink_EnsureManageSubtopicSentinel() {
    path := A_MyDocuments "\manage, study, set, top, link"
    if FileExist(path)
        return path
    try {
        if !DirExist(A_MyDocuments)
            return ""
        f := FileOpen(path, "w")
        if IsObject(f)
            f.Close()
    } catch {
        return ""
    }
    return FileExist(path) ? path : ""
}
```

### 3. Verify zero-byte sentinel

**PowerShell:**

```powershell
(Get-Item -LiteralPath $path).Length -eq 0
```

**AutoHotkey v2:**

```ahk
path := A_MyDocuments "\manage, study, set, top, link"
ok := FileExist(path) && FileRead(path) = ""
```

---

## When to use this pattern

| Approach                                          | Use when                                                                                          |
| ------------------------------------------------- | ------------------------------------------------------------------------------------------------- |
| **Documents sentinel (this doc)**                 | Cross-machine user profile marker; optional “feature initialized” flag; no repo checkout required |
| **Repo sentinel** (`A_ScriptDir\.cursor\...`)     | Cross-script IPC on one machine with a shared scripts folder (see focus-mode disable request)     |
| **HTTP lightweight API** (`StudyLinkHelpers.ahk`) | Shared URL/state across phone (MacroDroid), work PC, and personal PC                              |

**Do not** parse the sentinel filename by splitting on commas when building paths. Use `Join-Path` / string concat with the full literal name, or `-LiteralPath` / `FileExist(fullPath)`.

---

## Naming rules for new sentinels

1. Derive tokens from the user-visible feature name (lowercase words, comma + space separated).
2. No file extension.
3. Keep the file empty unless a future contract requires content (then document the format here).
4. Document each new sentinel in this file with creation date and owning feature.

---

## Future integration (optional)

If hotkeys should auto-create the sentinel on first use, add helpers to `StudyLinkHelpers.ahk`:

- `StudyLink_ManageSubtopicSentinelPath()` → `A_MyDocuments "\manage, study, set, top, link"`
- `StudyLink_EnsureManageSubtopicSentinel()` → create if missing; return path or `""`

Wire only where a local marker is required; URL read/write stays on `StudyLink_GetResult` / `StudyLink_Set`.

---

## Related documentation

- [focus-mode-secondary-monitor-blackout.md](focus-mode-secondary-monitor-blackout.md) — repo-local empty sentinels for cross-process focus mode
- [efficiency-canon.md](efficiency-canon.md) — IPC and sentinel return-value patterns
- `StudyLinkHelpers.ahk` — remote GET/POST for `subtopic` links

---

## Creation log (reference)

| Date       | Machine            | Action                                                                                              |
| ---------- | ------------------ | --------------------------------------------------------------------------------------------------- |
| 2026-05-22 | Personal (`eduev`) | Created `manage, study, set, top, link` in `A_MyDocuments` (0 bytes) via `[System.IO.File]::Create` |

Work PC: repeat the creation procedure once on that profile if the file is not present.
