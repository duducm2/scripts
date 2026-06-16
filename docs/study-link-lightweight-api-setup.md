# Study Link lightweight API — setup guides

Configure the **Google Apps Script** backend and **MacroDroid** for study links. PC code is split by module: YouTube in [`Utils.ahk`](../Utils.ahk), articles in [`StudyArticleLink.ahk`](../StudyArticleLink.ahk), favorites in [`StudyFavoriteLink.ahk`](../StudyFavoriteLink.ahk), shared HTTP in [`StudyLinkHelpers.ahk`](../StudyLinkHelpers.ahk).

## Study material menu (main)

Open via Study Topic selector (QuickLook flow). Keys **1–6**:

| Key | Module | Action |

| --- | ---------------- | --------------------------------------- |

| 1 | — | Mnemonics |

| 2 | — | Plans |

| 3 | YouTube subtopic | Manage Study Subtopic Link → inner 1–2 |

| 4 | Article link | Manage Study Article Link → inner 1–2 |

| 5 | Favorite link | Manage Study Favorite Link → inner 1–2 |

| 6 | Technique | QuickLook `studies/technique/README.md` |

## API keys (same web app URL)

| Key | Module | Sheet | PC file |

| ------------------- | ------------- | ----- | ------------------------- |

| `subtopic` | 3 — YouTube | A1 | `Utils.ahk` |

| `subtopic_article` | 4 — Article | A2 | `StudyArticleLink.ahk` |

| `subtopic_favorite` | 5 — Favorite | A3 | `StudyFavoriteLink.ahk` |

**Endpoint** (`STUDY_LINKS_API_URL` in `StudyLinkHelpers.ahk`):

`https://script.google.com/macros/s/AKfycbzKDLbmzGF8iduyNpaUymONEkERi089rBjW0jrYUX4a8K9ornfGwYIOsgQP1K_dfaj5/exec`

**Contract**

- **GET** `?key=<key>` → raw stored body (`key=<key>&url=<url>`; URL may contain literal `&`)

- **GET** `?key=<key>&plain=1` → URL only (MacroDroid Get macros — requires deployed `Code.gs`)

- **POST** `key=<encoded_key>&url=<full url>` — full body stored verbatim in the matching cell

**MacroDroid:** [study-link-macrodroid-same-url.md](study-link-macrodroid-same-url.md) — same `/exec` URL; POST body key selects A1 / A2 / A3.

## Apps Script

**Source:** [`study-link-api/Code.gs`](study-link-api/Code.gs) — paste into the Apps Script project bound to **my-study-database** (replace entire `Code.gs`; one `doPost` / one `doGet` only).

| Cell | Key | Handlers |

| ---- | ------------------- | ----------------------------------- |

| A1 | `subtopic` | `doPost` / `doGet` (default branch) |

| A2 | `subtopic_article` | `doPostArticle` / `doGetArticle` |

| A3 | `subtopic_favorite` | `doPostFavorite` / `doGetFavorite` |

**Deploy checklist**

1. Copy [`Code.gs`](study-link-api/Code.gs) into the Apps Script editor (delete duplicate `doPost` / `doGet` if present).

2. **Deploy → Manage deployments → Edit → Version: New version → Deploy** (editor save alone does not update the live URL).

3. Browser: `?key=subtopic_favorite&plain=1` returns **only** the URL (e.g. `https://youtu.be/...`) — required for Get macros.

4. Browser: `?key=subtopic`, `?key=subtopic_article`, `?key=subtopic_favorite` (no `plain=1`) return full stored body for PC/AHK.

If Set Favorite writes **A1** instead of **A3**, the deployed web app is still on old code — redeploy step 2.

## Code map

| Module | Source files | GET / SET |

| ----------- | ---------------------------------------------------------- | ----------------------------------------------------------------------------------------- |

| 3 YouTube | `StudyLinkHelpers.ahk`, `Utils.ahk` (UIA Share capture) | `StudyLink_GetResult(STUDYLINK_KEY_YOUTUBE)`, `StudyLink_Set(STUDYLINK_KEY_YOUTUBE, url)` |

| 4 Article | `StudyLinkHelpers.ahk`, `StudyArticleLink.ahk` (F6 + copy) | `StudyLink_GetResult(STUDYLINK_KEY_ARTICLE)`, `StudyLink_Set(STUDYLINK_KEY_ARTICLE, url)` |

| 5 Favorite | `StudyLinkHelpers.ahk`, `StudyFavoriteLink.ahk` (F6 + copy) | `StudyLink_GetResult(STUDYLINK_KEY_FAVORITE)`, `StudyLink_Set(STUDYLINK_KEY_FAVORITE, url)` |

| 6 Technique | `Utils.ahk` (`StudyTopicSelector_SelectTechnique`) | — (local repo markdown) |

---

## Module 3 — YouTube (`subtopic`)

### PC submenu (Study Topic → `[3]`)

| Inner key | Action |

| --------- | ------------------------------------------------ |

| 1 | Open YouTube link in Chrome |

| 2 | Set YouTube link (Share panel + Start at + Copy) |

Sentinel: `StudyLink_EnsureManageSubtopicSentinel()` runs on open (see [lightweight-api-sentinel-files.md](lightweight-api-sentinel-files.md)).

### MacroDroid

- **Set:** [`macrodroid/Set_Video_(direct_link).macro`](<macrodroid/Set_Video_(direct_link).macro>) — Clipboard Refresh → POST `key=subtopic&url=` + clip → **A1**

- **Get:** [`macrodroid/Get_Video_(direct_link).macro`](<macrodroid/Get_Video_(direct_link).macro>) — GET `?key=subtopic&plain=1` → Open Web Page `{lv=resp}`

Workflow: YouTube → **Share → Copy link** → run Set from drawer. Toast must show clip starting with `http`.

### AHK snippets

```ahk

yt := StudyLink_GetResult(STUDYLINK_KEY_YOUTUBE)

if (yt["ok"] && yt["url"] != "")

    StudyLink_OpenUrlInChrome(yt["url"])



StudyLink_Set(STUDYLINK_KEY_YOUTUBE, timestampedUrl)

```

---

## Module 4 — Article (`subtopic_article`)

### PC submenu (Study Topic → `[4]`)

| Inner key | Action |

| --------- | --------------------------------------------------------- |

| 1 | Open article link in Chrome |

| 2 | Set article link — Chrome **F6**, **Ctrl+C**, POST to API |

All article UI and capture live in **`StudyArticleLink.ahk`** only.

### MacroDroid

- **Set:** [`macrodroid/Set_Article_(direct_link).macro`](<macrodroid/Set_Article_(direct_link).macro>) — POST `key=subtopic_article&url=` + clip → **A2**

- **Get:** [`macrodroid/Get_Article_(direct_link).macro`](<macrodroid/Get_Article_(direct_link).macro>) — GET `?key=subtopic_article&plain=1` → Open Web Page `{lv=resp}`

### AHK snippets

```ahk

art := StudyLink_GetResult(STUDYLINK_KEY_ARTICLE)

if (art["ok"] && art["url"] != "")

    StudyLink_OpenUrlInChrome(art["url"])



url := StudyArticleLink_CaptureChromeUrlFromAddressBar(&errMsg)

if (url != "")

    StudyLink_Set(STUDYLINK_KEY_ARTICLE, url)

```

---

## Module 5 — Favorite (`subtopic_favorite`)

Third slot in **A3**. PC and MacroDroid share the same API key.

### PC submenu (Study Topic → `[5]`)

| Inner key | Action |

| --------- | ---------------------------------------------------------- |

| 1 | Open favorite link in Chrome |

| 2 | Set favorite link — Chrome **F6**, **Ctrl+C**, POST to API |

All favorite UI lives in **`StudyFavoriteLink.ahk`**; capture reuses `StudyArticleLink_CaptureChromeUrlFromAddressBar`.

### MacroDroid

- **Set:** [`macrodroid/Set_Favorite_(direct_link).macro`](<macrodroid/Set_Favorite_(direct_link).macro>) — POST `key=subtopic_favorite&url=` + clip → **A3**

- **Get:** [`macrodroid/Get_Favorite_(direct_link).macro`](<macrodroid/Get_Favorite_(direct_link).macro>) — GET `?key=subtopic_favorite&plain=1` → Open Web Page `{lv=resp}`

Full Android table and troubleshooting: [study-link-macrodroid-same-url.md](study-link-macrodroid-same-url.md).

### AHK snippets

```ahk

fav := StudyLink_GetResult(STUDYLINK_KEY_FAVORITE)

if (fav["ok"] && fav["url"] != "")

    StudyLink_OpenUrlInChrome(fav["url"])



url := StudyArticleLink_CaptureChromeUrlFromAddressBar(&errMsg)

if (url != "")

    StudyLink_Set(STUDYLINK_KEY_FAVORITE, url)

```

---

## Module 6 — Technique

Study Topic → **`[6] Technique`** opens the technique README in QuickLook (`StudyTopicSelector_SelectTechnique` in `Utils.ahk`). No StudyLink API.

---

## Verification

1. **Quick Update Scripts** after any AHK change.

2. Run `TestStudyLinkApi.ahk` — YouTube, article, and favorite SET/GET.

3. Manual: `[3]` inner 1–2, `[4]` inner 1–2, `[5]` inner 1–2, `[6]` (technique).

4. Phone: Set then Get for each link type after Apps Script redeploy.

---

## Related files

- [`study-link-api/Code.gs`](study-link-api/Code.gs) — Apps Script source (deploy to Google)

- `StudyLinkHelpers.ahk` — HTTP, keys, `StudyLink_OpenUrlInChrome`, sentinel, functional test

- `StudyArticleLink.ahk` — module 4 GUI + F6 capture

- `StudyFavoriteLink.ahk` — module 5 GUI + open/set (reuses F6 capture)

- `Utils.ahk` — module 3 GUI + YouTube UIA capture

- `TestStudyLinkApi.ahk` — API smoke test

- [`macrodroid/`](macrodroid/) — Set/Get macro exports (Video, Article, Favorite)

- [study-link-macrodroid-same-url.md](study-link-macrodroid-same-url.md) — one `/exec` URL, POST/GET routing, Android workflow

- [lightweight-api-sentinel-files.md](lightweight-api-sentinel-files.md) — module 3 sentinel

- [efficiency-canon.md](efficiency-canon.md) — clipboard restore patterns
