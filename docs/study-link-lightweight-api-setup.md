# Study Link lightweight API — setup guides

Configure the **Google Apps Script** backend and **MacroDroid** for study links. PC code is split by module: YouTube in [`Utils.ahk`](../Utils.ahk), articles in [`StudyArticleLink.ahk`](../StudyArticleLink.ahk), shared HTTP in [`StudyLinkHelpers.ahk`](../StudyLinkHelpers.ahk).

## Study material menu (main)

Open via Study Topic selector (QuickLook flow). Keys **1–5**:

| Key | Module           | Action                                  |
| --- | ---------------- | --------------------------------------- |
| 1   | —                | Mnemonics                               |
| 2   | —                | Plans                                   |
| 3   | YouTube subtopic | Manage Study Subtopic Link → inner 1–2  |
| 4   | Article link     | Manage Study Article Link → inner 1–2   |
| 5   | Technique        | QuickLook `studies/technique/README.md` |

## API keys (same web app URL)

| Key                | Module      | PC file                |
| ------------------ | ----------- | ---------------------- |
| `subtopic`         | 3 — YouTube | `Utils.ahk`            |
| `subtopic_article` | 4 — Article | `StudyArticleLink.ahk` |

**Endpoint** (`STUDY_LINKS_API_URL` in `StudyLinkHelpers.ahk`):

`https://script.google.com/macros/s/AKfycbzzkjpT_47W0TwcjwEulzkV9l5xTtqcwWJmF0h-B-11SwiL_49SPhKXnj3PTsgFUZcp/exec`

**Contract**

- **GET** `?key=<key>` → `key=<key>&url=<url>` (URL may contain literal `&`)
- **POST** `key=<encoded_key>&url=<full url>` — Apps Script reads everything after the first `url=`

## Code map

| Module      | Source files                                               | GET / SET                                                                                 |
| ----------- | ---------------------------------------------------------- | ----------------------------------------------------------------------------------------- |
| 3 YouTube   | `StudyLinkHelpers.ahk`, `Utils.ahk` (UIA Share capture)    | `StudyLink_GetResult(STUDYLINK_KEY_YOUTUBE)`, `StudyLink_Set(STUDYLINK_KEY_YOUTUBE, url)` |
| 4 Article   | `StudyLinkHelpers.ahk`, `StudyArticleLink.ahk` (F6 + copy) | `StudyLink_GetResult(STUDYLINK_KEY_ARTICLE)`, `StudyLink_Set(STUDYLINK_KEY_ARTICLE, url)` |
| 5 Technique | `Utils.ahk` (`StudyTopicSelector_SelectTechnique`)         | — (local repo markdown)                                                                   |

---

## Module 3 — YouTube (`subtopic`)

### PC submenu (Study Topic → `[3]`)

| Inner key | Action                                           |
| --------- | ------------------------------------------------ |
| 1         | Open YouTube link in Chrome                      |
| 2         | Set YouTube link (Share panel + Start at + Copy) |

Sentinel: `StudyLink_EnsureManageSubtopicSentinel()` runs on open (see [lightweight-api-sentinel-files.md](lightweight-api-sentinel-files.md)).

### Guide A — Apps Script (YouTube — keep as-is)

Your existing handlers (cell **A1**) stay unchanged:

```javascript
function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange("A1").setValue(e.postData.contents);
  return ContentService.createTextOutput("Saved");
}

function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var link = sheet.getRange("A1").getValue();
  return ContentService.createTextOutput(link);
}
```

After adding article support (module 4), only add the two `if` lines at the top of each function — see module 4 wiring below.

### Guide B — MacroDroid (YouTube)

Use your existing **Set_Video** (or equivalent) macro: POST body `key=subtopic&url=` + clipboard URL (literal `&` in URL allowed).

### Code snippets — Module 3

```ahk
; GET (read current YouTube link)
yt := StudyLink_GetResult(STUDYLINK_KEY_YOUTUBE)
if (yt["ok"] && yt["url"] != "")
    StudyLink_OpenUrlInChrome(yt["url"])

; SET (after UIA capture in Utils — StudyLink_CaptureYoutubeTimestampUrl)
StudyLink_Set(STUDYLINK_KEY_YOUTUBE, timestampedUrl)
```

---

## Module 4 — Article (`subtopic_article`)

### PC submenu (Study Topic → `[4]`)

| Inner key | Action                                                    |
| --------- | --------------------------------------------------------- |
| 1         | Open article link in Chrome                               |
| 2         | Set article link — Chrome **F6**, **Ctrl+C**, POST to API |

All article UI and capture live in **`StudyArticleLink.ahk`** only (no YouTube/Share/UIA code in that file).

### Guide A — Apps Script (article — copy/paste)

YouTube keeps your existing **`doPost`** / **`doGet`** (cell **A1**) unchanged. Add the two functions below for articles (cell **A2**), same pattern.

**1. Paste these new functions** (anywhere in the same `.gs` file, e.g. below your YouTube functions):

```javascript
function doPostArticle(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Writes the raw string from your HTTP request to cell A2
  sheet.getRange("A2").setValue(e.postData.contents);
  return ContentService.createTextOutput("Saved");
}

function doGetArticle(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Reads cell A2 and returns it as plain text
  var link = sheet.getRange("A2").getValue();
  return ContentService.createTextOutput(link);
}
```

**2. Wire the web app** (one extra line at the top of each existing handler — your A1 logic stays as-is below):

```javascript
function doPost(e) {
  if (
    e.postData &&
    e.postData.contents &&
    e.postData.contents.indexOf("subtopic_article") !== -1
  )
    return doPostArticle(e);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange("A1").setValue(e.postData.contents);
  return ContentService.createTextOutput("Saved");
}

function doGet(e) {
  if (e.parameter && e.parameter.key === "subtopic_article")
    return doGetArticle(e);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var link = sheet.getRange("A1").getValue();
  return ContentService.createTextOutput(link);
}
```

| Cell | Module      | Functions                                |
| ---- | ----------- | ---------------------------------------- |
| A1   | 3 — YouTube | `doPost` / `doGet` (body below the `if`) |
| A2   | 4 — Article | `doPostArticle` / `doGetArticle`         |

**3. Redeploy** the web app (new version), then verify in the browser:

```
<EXEC>?key=subtopic_article
```

Expect the text stored in **A2** (same shape as A1: full POST body or URL string your clients write).

**MacroDroid / PC POST body for article:** `key=subtopic_article&url=` + your URL (stored as-is in A2, like YouTube in A1).

### Guide B — MacroDroid (article)

1. New macro: e.g. `Study — Set article link`.
2. Trigger: your choice.
3. Read clipboard (URL already copied on device).
4. POST to `/exec`, body:

   ```
   key=subtopic_article&url={clipUrl}
   ```

5. Optional: GET `?key=subtopic_article` + Open URL on phone.

### Code snippets — Module 4

```ahk
; GET (read current article link)
art := StudyLink_GetResult(STUDYLINK_KEY_ARTICLE)
if (art["ok"] && art["url"] != "")
    StudyLink_OpenUrlInChrome(art["url"])

; SET (Chrome address bar — StudyArticleLink.ahk)
url := StudyArticleLink_CaptureChromeUrlFromAddressBar(&errMsg)
if (url != "")
    StudyLink_Set(STUDYLINK_KEY_ARTICLE, url)
```

---

## Module 5 — Technique

Study Topic → **`[5] Technique`** opens the technique README in QuickLook (`StudyTopicSelector_SelectTechnique` in `Utils.ahk`). No StudyLink API.

---

## Verification

1. **Quick Update Scripts** after any code change.
2. Run `TestStudyLinkApi.ahk` — YouTube and article SET/GET (both keys).
3. Manual: `[3]` inner 1–2 (YouTube), `[4]` inner 1–2 (article), `[5]` (technique).

---

## Related files

- `StudyLinkHelpers.ahk` — HTTP, keys, `StudyLink_OpenUrlInChrome`, sentinel, functional test
- `StudyArticleLink.ahk` — module 4 GUI + F6 capture
- `Utils.ahk` — module 3 GUI + YouTube UIA capture
- `TestStudyLinkApi.ahk` — API smoke test
- [lightweight-api-sentinel-files.md](lightweight-api-sentinel-files.md) — module 3 sentinel
- [efficiency-canon.md](efficiency-canon.md) — clipboard restore patterns
