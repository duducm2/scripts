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

| Key                 | Module        | PC file                   |
| ------------------- | ------------- | ------------------------- |
| `subtopic`          | 3 — YouTube   | `Utils.ahk`               |
| `subtopic_article`  | 4 — Article   | `StudyArticleLink.ahk`    |
| `subtopic_favorite` | Favorite link | MacroDroid only (for now) |

**Endpoint** (`STUDY_LINKS_API_URL` in `StudyLinkHelpers.ahk`):

`https://script.google.com/macros/s/AKfycbzKDLbmzGF8iduyNpaUymONEkERi089rBjW0jrYUX4a8K9ornfGwYIOsgQP1K_dfaj5/exec`

**MacroDroid (same URL, three POST bodies):** see [study-link-macrodroid-same-url.md](study-link-macrodroid-same-url.md) — YouTube uses `key=subtopic&url=`; article uses `key=subtopic_article&url=`; favorite uses `key=subtopic_favorite&url=` (all POST to the same `/exec`).

**Contract**

- **GET** `?key=<key>` → `key=<key>&url=<url>` (URL may contain literal `&`)
- **POST** `key=<encoded_key>&url=<full url>` — Apps Script reads everything after the first `url=`

## Code map

| Module      | Source files                                               | GET / SET                                                                                 |
| ----------- | ---------------------------------------------------------- | ----------------------------------------------------------------------------------------- |
| 3 YouTube   | `StudyLinkHelpers.ahk`, `Utils.ahk` (UIA Share capture)    | `StudyLink_GetResult(STUDYLINK_KEY_YOUTUBE)`, `StudyLink_Set(STUDYLINK_KEY_YOUTUBE, url)` |
| 4 Article   | `StudyLinkHelpers.ahk`, `StudyArticleLink.ahk` (F6 + copy) | `StudyLink_GetResult(STUDYLINK_KEY_ARTICLE)`, `StudyLink_Set(STUDYLINK_KEY_ARTICLE, url)` |
| Favorite    | MacroDroid SET/GET macros only                             | — (PC/AHK deferred)                                                                       |
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

Import both macros in MacroDroid:

- **[`docs/macrodroid/Set_Video_(direct_link).macro`](<macrodroid/Set_Video_(direct_link).macro>)** — copy link → Clipboard Refresh → POST `key=subtopic&url=` + clip → A1
- **[`docs/macrodroid/Get_Video_(direct_link).macro`](<macrodroid/Get_Video_(direct_link).macro>)** — Open Web Page `?key=subtopic&open=1` (server redirects to stored URL; requires `doGet` `open=1` branch in consolidated script)

**Set workflow:** YouTube → **Share → Copy link** → run Set macro from drawer. Clipboard Refresh runs first (Android 10+), then POST. Toast shows API reply + clip — confirm it starts with `http`.

Full MacroDroid table and troubleshooting: [study-link-macrodroid-same-url.md](study-link-macrodroid-same-url.md).

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

Import both macros in MacroDroid:

- **[`docs/macrodroid/Set_Article_(direct_link).macro`](<macrodroid/Set_Article_(direct_link).macro>)** — copy URL → Clipboard Refresh → POST `key=subtopic_article&url=` + clip → A2
- **[`docs/macrodroid/Get_Article_(direct_link).macro`](<macrodroid/Get_Article_(direct_link).macro>)** — Open Web Page `?key=subtopic_article&open=1` (server redirects; requires `doGetArticle` `open=1` branch)

**Set workflow:** Browser → copy article URL → run Set macro from drawer. Toast shows API reply + clip — confirm it starts with `http`.

Full MacroDroid table and troubleshooting: [study-link-macrodroid-same-url.md](study-link-macrodroid-same-url.md).

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

## Favorite link (`subtopic_favorite`) — Android only

Stores a third link in sheet cell **A3**. Same `/exec` URL; PC/AHK support is deferred.

### Guide A — Apps Script (favorite — copy/paste)

**Important:** Your live web app must expose the favorite router. If Set Favorite overwrites **A1**, the deployed code still lacks the favorite branch (saving in the editor is not enough — you must **Deploy → Manage deployments → Edit → New version → Deploy**).

**Verify routing before using the phone macro** (in a browser):

| URL                              | Should read                            |
| -------------------------------- | -------------------------------------- |
| `.../exec?key=subtopic`          | A1 only                                |
| `.../exec?key=subtopic_favorite` | A3 only (empty until you set favorite) |

If both URLs return the **same** text, favorite GET routing is not deployed yet.

**Recommended:** replace your entire `Code.gs` with this single file (one `doPost` / one `doGet` — delete any duplicate handlers):

```javascript
function doPost(e) {
  var key = e.parameter && e.parameter.key ? String(e.parameter.key) : "";
  var contents = e.postData && e.postData.contents ? e.postData.contents : "";

  // Prefer parsed form key (MacroDroid POST); fall back to body substring
  if (key === "subtopic_article" || contents.indexOf("subtopic_article") !== -1)
    return doPostArticle(e);
  if (
    key === "subtopic_favorite" ||
    contents.indexOf("subtopic_favorite") !== -1
  )
    return doPostFavorite(e);

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange("A1").setValue(contents);
  return ContentService.createTextOutput("Saved");
}

function doGet(e) {
  var key = e.parameter && e.parameter.key ? String(e.parameter.key) : "";

  if (key === "subtopic_article") return doGetArticle(e);
  if (key === "subtopic_favorite") return doGetFavorite(e);

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var link = sheet.getRange("A1").getValue();
  if (e.parameter && e.parameter.open === "1") {
    return openStoredLinkResponse(
      link,
      "No video link stored. Run Set Video on your phone first.",
    );
  }
  return ContentService.createTextOutput(link);
}

function doPostArticle(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange("A2").setValue(e.postData.contents);
  return ContentService.createTextOutput("Saved");
}

function doGetArticle(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var link = sheet.getRange("A2").getValue();
  if (e.parameter && e.parameter.open === "1") {
    return openStoredLinkResponse(
      link,
      "No article link stored. Run Set Article on your phone first.",
    );
  }
  return ContentService.createTextOutput(link);
}

function doPostFavorite(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange("A3").setValue(e.postData.contents);
  return ContentService.createTextOutput("Saved");
}

function extractUrlFromStoredBody(stored) {
  var s = String(stored || "");
  var pos = s.indexOf("url=");
  if (pos !== -1) return s.substring(pos + 4);
  return s;
}

function openStoredLinkResponse(stored, emptyMessage) {
  var url = extractUrlFromStoredBody(stored);
  if (!url || url.indexOf("http") !== 0) {
    return HtmlService.createHtmlOutput("<p>" + emptyMessage + "</p>");
  }
  return HtmlService.createHtmlOutput(
    "<script>window.location.href=" + JSON.stringify(url) + ";</script>",
  ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doGetFavorite(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var link = sheet.getRange("A3").getValue();
  if (e.parameter && e.parameter.open === "1") {
    return openStoredLinkResponse(
      link,
      "No favorite link stored. Run Set Favorite on your phone first.",
    );
  }
  return ContentService.createTextOutput(link);
}
```

| Cell | Key                 | Functions                           |
| ---- | ------------------- | ----------------------------------- |
| A1   | `subtopic`          | `doPost` / `doGet` (default branch) |
| A2   | `subtopic_article`  | `doPostArticle` / `doGetArticle`    |
| A3   | `subtopic_favorite` | `doPostFavorite` / `doGetFavorite`  |

**Redeploy checklist:**

1. In Apps Script, press **Ctrl+F** and search `function doPost` — there must be **only one** definition.
2. **Deploy → Manage deployments → Edit (pencil) → Version: New version → Deploy**
3. Browser test: `.../exec?key=subtopic&open=1` redirects to A1 URL; `?key=subtopic_article&open=1` → A2; `?key=subtopic_favorite&open=1` → A3.
4. Plain GET (no `open=1`) still returns raw stored text for PC/AHK (`?key=subtopic`, etc.).
5. Run Set Favorite on phone — only **A3** should change; A1 stays `key=subtopic&url=...`

**POST body:** `key=subtopic_favorite&url=` + your URL (stored as-is in A3).

### Guide B — MacroDroid (favorite)

Import both macros:

- **[`docs/macrodroid/Set_Favorite_(direct_link).macro`](<macrodroid/Set_Favorite_(direct_link).macro>)** — copy URL → Clipboard Refresh → POST `key=subtopic_favorite&url=` + clip → A3
- **[`docs/macrodroid/Get_Favorite_(direct_link).macro`](<macrodroid/Get_Favorite_(direct_link).macro>)** — Open Web Page `?key=subtopic_favorite&open=1` (server redirects to stored URL; requires `doGetFavorite` `open=1` branch above)

Full table and troubleshooting: [study-link-macrodroid-same-url.md](study-link-macrodroid-same-url.md).

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
- [`macrodroid/Set_Video_(direct_link).macro`](<macrodroid/Set_Video_(direct_link).macro>) — YouTube SET macro
- [`macrodroid/Get_Video_(direct_link).macro`](<macrodroid/Get_Video_(direct_link).macro>) — YouTube GET macro (open via `open=1` redirect)
- [`macrodroid/Set_Article_(direct_link).macro`](<macrodroid/Set_Article_(direct_link).macro>) — article SET macro
- [`macrodroid/Get_Article_(direct_link).macro`](<macrodroid/Get_Article_(direct_link).macro>) — article GET macro (open via `open=1` redirect)
- [`macrodroid/Set_Favorite_(direct_link).macro`](<macrodroid/Set_Favorite_(direct_link).macro>) — favorite SET macro
- [`macrodroid/Get_Favorite_(direct_link).macro`](<macrodroid/Get_Favorite_(direct_link).macro>) — favorite GET macro (open via `open=1` redirect)
- [study-link-macrodroid-same-url.md](study-link-macrodroid-same-url.md) — one `/exec` URL, YouTube vs article POST/GET, Android workflow
- [lightweight-api-sentinel-files.md](lightweight-api-sentinel-files.md) — module 3 sentinel
- [efficiency-canon.md](efficiency-canon.md) — clipboard restore patterns
