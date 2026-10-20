# TMP — How we built the Study Link Google API

Delete this file when you no longer need the refresher. Canonical docs stay in:

- [study-link-lightweight-api-setup.md](study-link-lightweight-api-setup.md)
- [study-link-macrodroid-same-url.md](study-link-macrodroid-same-url.md)
- Source: [study-link-api/Code.gs](study-link-api/Code.gs)

---

## What it is

A tiny **Google Apps Script web app** bound to a spreadsheet (sheet name / file: **my-study-database**). Four cells hold four URLs:

| Cell | API key             | Used for                                    |
| ---- | ------------------- | ------------------------------------------- |
| A1   | `subtopic`          | Video (YouTube)                             |
| A2   | `subtopic_article`  | Article                                     |
| A3   | `subtopic_favorite` | Favorite (any URL you set manually) ss      |
| A4   | `subtopic_practice` | Practice MD on GitHub (after palace import) |

PC (AHK) and phone (MacroDroid) share one `/exec` URL. Only the **key** in the POST body / GET query changes which cell is read or written.

Live endpoint (also `STUDY_LINKS_API_URL` in `Lib/study/StudyLinkHelpers.ahk`):

`https://script.google.com/macros/s/AKfycbzKDLbmzGF8iduyNpaUymONEkERi089rBjW0jrYUX4a8K9ornfGwYIOsgQP1K_dfaj5/exec`

---

## Step-by-step (recreate from scratch)

### 1. Spreadsheet

1. In Google Drive, create a Google Sheet (or open the existing **my-study-database**).
2. On the active tab, leave **A1 / A2 / A3 / A4** for the four link bodies (you can leave them empty at first).

### 2. Apps Script project

1. In that Sheet: **Extensions → Apps Script**.
2. Delete any default stub code.
3. Paste the full contents of `docs/study-link-api/Code.gs` into `Code.gs` (one `doPost`, one `doGet` only — no duplicates).
4. Save the project.

### 3. Deploy as web app

1. **Deploy → New deployment** (first time) or **Deploy → Manage deployments → Edit** (updates).
2. Type: **Web app**.
3. **Execute as:** Me.
4. **Who has access:** Anyone (MacroDroid / AHK need an unauthenticated GET/POST).
5. Deploy. Copy the **Web app URL** ending in `/exec`.
6. Put that URL into `STUDY_LINKS_API_URL` in `StudyLinkHelpers.ahk` and into the MacroDroid Set/Get macros.

**Important:** Saving the editor does **not** update the live URL. Every code change needs **Manage deployments → Edit → Version: New version → Deploy**.

### 4. Contract (how clients talk to it)

**Write (POST)** — body `application/x-www-form-urlencoded`, URL **not** percent-encoded after `url=` (literal, same as MacroDroid):

```text
key=subtopic_practice&url=https://github.com/duducm2/scripts/blob/main/mnemonics/output/practice/data-analyst.md
```

| Body contains          | Goes to |
| ---------------------- | ------- |
| `subtopic_article`     | A2      |
| `subtopic_favorite`    | A3      |
| `subtopic_practice`    | A4      |
| otherwise (`subtopic`) | A1      |

Response on success: plain `Saved`.

**Read (GET)**

| Query                            | Returns                                    |
| -------------------------------- | ------------------------------------------ |
| `?key=subtopic_practice`         | Full stored body (`key=…&url=…`) — for AHK |
| `?key=subtopic_practice&plain=1` | URL only — for MacroDroid Open Web Page    |

Same pattern for `subtopic`, `subtopic_article`, and `subtopic_favorite`.

### 5. Browser smoke test

1. POST or Set Practice with a known URL into A4.
2. Open in browser:  
   `…/exec?key=subtopic_practice&plain=1`  
   → should show **only** the URL.
3. Open without `plain=1` → full `key=…&url=…` body.
4. Confirm Favorite still works: `?key=subtopic_favorite&plain=1` reads **A3** only.
5. If Set Practice still writes **A1**, the deployment is stale — redeploy a **new version**.

### 6. PC (AHK)

- Shared HTTP: `StudyLinkHelpers.ahk` → `StudyLink_Get` / `StudyLink_Set`.
- Favorite UI: `StudyFavoriteLink.ahk` + Memory Palace **Links** / cheat sheet **Alt+F** (set) / **Shift+F** (open) — still **A3**.
- Practice: palace import → `Palace_PushPracticeMdLink` → **A4** (not exposed in `#\!+W` / Links UI).
- Smoke test script: `infra/tools/TestStudyLinkApi.ahk`.

### 7. Phone (MacroDroid)

Import the macros under `docs/macrodroid/` (Set/Get × Video, Article, Favorite, Practice). All use the **same** `/exec` URL; only the key differs. Details: [study-link-macrodroid-same-url.md](study-link-macrodroid-same-url.md).

Android 10+: Set macros must run **Clipboard Refresh** before reading `{clipboard}`.

---

## Palace import → practice (what we added)

On successful **memory palace pack import** (same place Quick Recall sync runs):

1. Resolve the imported study’s `notes_rel_path` (e.g. `data-analyst`).
2. Build:  
   `https://github.com/duducm2/scripts/blob/main/mnemonics/output/practice/{slug}.md`
3. `StudyLink_Set(subtopic_practice, thatUrl)` → cell **A4** (does **not** touch Favorite / A3).

Then:

- **Get Practice** (MacroDroid) opens that GitHub practice file (A4).
- **Shift+F** / **Get Favorite** still open whatever you set as Favorite on A3.
- **Win+Alt+Shift+9 tap** is still **Push scripts** (Utility `[G]`), not “open practice”. Push so GitHub has the latest `practice/*.md` before you open the link on the phone.

---

## Quick mental model

```text
Clipboard / palace import
        │
        ▼
 POST key=…&url=…  ──►  Apps Script  ──►  Sheet A1/A2/A3/A4
        ▲
 GET ?key=…[&plain=1] ◄────────────────────┘
        │
        ▼
 Chrome (PC) / Open Web Page (phone)
```
