# MacroDroid — same web app URL for YouTube, article, favorite, and practice

One deployment URL for all link types. MacroDroid and the PC scripts use the **same** `/exec` address; only the **POST body** and **GET query** change.

**Web app URL (POST and GET base):**

`https://script.google.com/macros/s/AKfycbzKDLbmzGF8iduyNpaUymONEkERi089rBjW0jrYUX4a8K9ornfGwYIOsgQP1K_dfaj5/exec`

(Also set as `STUDY_LINKS_API_URL` in `StudyLinkHelpers.ahk`.)

---

## How the server picks YouTube vs article vs favorite vs practice

After you deploy [`study-link-api/Code.gs`](study-link-api/Code.gs) (see [study-link-lightweight-api-setup.md](study-link-lightweight-api-setup.md)):

| Request                                    | Routed to          | Sheet cell |
| ------------------------------------------ | ------------------ | ---------- |
| POST body contains `subtopic_article`      | `doPostArticle`    | **A2**     |
| POST body contains `subtopic_favorite`     | `doPostFavorite`   | **A3**     |
| POST body contains `subtopic_practice`     | `doPostPractice`   | **A4**     |
| Any other POST (e.g. `key=subtopic&url=…`) | `doPost` (YouTube) | **A1**     |
| GET `?key=subtopic_article`                | `doGetArticle`     | **A2**     |
| GET `?key=subtopic_favorite`               | `doGetFavorite`    | **A3**     |
| GET `?key=subtopic_practice`               | `doGetPractice`    | **A4**     |
| GET without that key (or `?key=subtopic`)  | `doGet` (YouTube)  | **A1**     |

Router order in `doPost`: `subtopic_article` → `subtopic_favorite` → `subtopic_practice` → default YouTube.

You do **not** need a second URL. Differentiate with the **key name in the POST body** (and `?key=` on GET).

---

## MacroDroid — SET (save link from clipboard)

**Canonical macro exports:**

- YouTube SET: [`macrodroid/Set_Video_(direct_link).macro`](<macrodroid/Set_Video_(direct_link).macro>)
- YouTube GET: [`macrodroid/Get_Video_(direct_link).macro`](<macrodroid/Get_Video_(direct_link).macro>)
- Article SET: [`macrodroid/Set_Article_(direct_link).macro`](<macrodroid/Set_Article_(direct_link).macro>)
- Article GET: [`macrodroid/Get_Article_(direct_link).macro`](<macrodroid/Get_Article_(direct_link).macro>)
- Favorite SET: [`macrodroid/Set_Favorite_(direct_link).macro`](<macrodroid/Set_Favorite_(direct_link).macro>)
- Favorite GET: [`macrodroid/Get_Favorite_(direct_link).macro`](<macrodroid/Get_Favorite_(direct_link).macro>)
- Practice SET: [`macrodroid/Set_Practice_(direct_link).macro`](<macrodroid/Set_Practice_(direct_link).macro>)
- Practice GET: [`macrodroid/Get_Practice_(direct_link).macro`](<macrodroid/Get_Practice_(direct_link).macro>)
- Practice poll: [`macrodroid/Poll_Practice_link_(5_min).macro`](<macrodroid/Poll_Practice_link_(5_min).macro>) — every 5 min GET A4; if changed → global `{v=practice_url}` + toast

Import the Set/Get macros (and the poll macro if you want background updates).

### Android workflow (YouTube)

1. In the **YouTube** app, open the video.
2. Tap **Share** → **Copy link** (not the video title, description, or transcript).
3. **Verify:** paste in Notes — must start with `http` (if you see the title, copy the link again).
4. Run **Set Video (direct link)** from your MacroDroid drawer.

The macro **refreshes the clipboard** (required on Android 10+), stores it in `clip`, POSTs `key=subtopic&url={lv=clip}` — same shape as PC `StudyLink_Set`. No URL validation; you must copy the link yourself (**Share → Copy link**, not the title).

**Macro actions (Set, in order):** Clipboard Refresh → Set `clip` = `{clipboard}` → HTTP POST → Toast `{lv=resp}` + `{lv=clip}`.

**Get:** run **Get Video (direct link)** → GET `?key=subtopic&plain=1` → `resp` is the URL → Open Web Page.

**Get macro actions (all three):** HTTP GET with `&plain=1` → Open Web Page `{lv=resp}`. No Text Manipulation (that action fails when imported from JSON). **Redeploy `Code.gs` first** so `plain=1` returns the URL only.

**Android 10+:** MacroDroid cannot read the system clipboard from the background without the **Clipboard Refresh** action first ([MacroDroid wiki](https://macrodroidforum.com/wiki/index.php/Action:_Clipboard_Refresh)). Without it, `{clipboard}` may be empty or stale.

### Android workflow (Article)

1. In your **browser** (or any app), copy the article URL from the address bar or **Share → Copy link**.
2. **Verify:** paste in Notes — must start with `http`.
3. Run **Set Article (direct link)** from your MacroDroid drawer.

Same macro actions as YouTube Set; POST body is `key=subtopic_article&url={lv=clip}` → Apps Script routes to **A2**.

**Get:** run **Get Article (direct link)** → GET `?key=subtopic_article&plain=1` → Open Web Page `{lv=resp}`.

### Android workflow (Favorite)

**Set:** copy any URL → verify in Notes (`http`) → run **Set Favorite (direct link)** → toast shows `Saved` + clip → **A3**.

**Get:** run **Get Favorite (direct link)** → GET `?key=subtopic_favorite&plain=1` → Open Web Page `{lv=resp}`.

### Android workflow (Practice)

Practice is written from the **PC** on palace import (`Palace_PushPracticeMdLink` → A4). On the phone you mainly need **Get**:

**Get:** run **Get Practice (direct link)** → GET `?key=subtopic_practice&plain=1` → Open Web Page `{lv=resp}` (GitHub practice MD).

**Set** (optional): same pattern as Favorite; POST `key=subtopic_practice&url=` → **A4**. Does not overwrite Favorite.

| Macro                                   | HTTP | URL                      | Body / query                                                       |
| --------------------------------------- | ---- | ------------------------ | ------------------------------------------------------------------ |
| **YouTube** (Set Video direct link)     | POST | `/exec` (full URL above) | `key=subtopic&url=` + `{lv=clip}` after Clipboard Refresh          |
| **YouTube** (Get Video direct link)     | GET  | **same** `/exec`         | GET `?key=subtopic&plain=1` → Open Web Page `{lv=resp}`            |
| **Article** (Set Article direct link)   | POST | **same** `/exec`         | `key=subtopic_article&url=` + `{lv=clip}` after Clipboard Refresh  |
| **Article** (Get Article direct link)   | GET  | **same** `/exec`         | GET `?key=subtopic_article&plain=1` → Open Web Page `{lv=resp}`    |
| **Favorite** (Set Favorite direct link) | POST | **same** `/exec`         | `key=subtopic_favorite&url=` + `{lv=clip}` after Clipboard Refresh |
| **Favorite** (Get Favorite direct link) | GET  | **same** `/exec`         | GET `?key=subtopic_favorite&plain=1` → Open Web Page `{lv=resp}`   |
| **Practice** (Set Practice direct link) | POST | **same** `/exec`         | `key=subtopic_practice&url=` + `{lv=clip}` after Clipboard Refresh |
| **Practice** (Get Practice direct link) | GET  | **same** `/exec`         | GET `?key=subtopic_practice&plain=1` → Open Web Page `{lv=resp}`   |

MacroDroid settings (both macros):

- Method: **POST**
- Content-Type: `application/x-www-form-urlencoded`
- Clipboard must hold the full link **before** you run the macro
- Success response: `Saved`
- Success toast: `{lv=resp}` then `{lv=clip}` — check the second line starts with `http`

**Important:** Article → `subtopic_article` (A2); Favorite → `subtopic_favorite` (A3); Practice → `subtopic_practice` (A4).

### API contract (Apps Script is working as designed)

Your `doPost` stores **`e.postData.contents` verbatim** in A1 and returns `"Saved"`. That is correct.

| Step            | What happens                                                                                     |
| --------------- | ------------------------------------------------------------------------------------------------ |
| MacroDroid POST | Body must be `key=subtopic&url=<clipboard text>` in the **request body** (not query string only) |
| Apps Script     | Writes full body string to A1, responds `Saved`                                                  |
| PC GET          | Reads A1, parses everything after `url=` as the link                                             |

If A1 shows `key=subtopic&url=Execute systematic literature reviews...`, the API **did save** — it saved the video **title** because that was the POST body. The toast `Saved` only means Apps Script accepted the request.

Live check: POST `key=subtopic&url=https://youtu.be/test` → GET `?key=subtopic` should return that same string.

Article POST: `key=subtopic_article&url=https://example.com/paper` → stored in **A2**; GET `?key=subtopic_article` returns that string.

Favorite POST: `key=subtopic_favorite&url=https://gemini.google.com/` → stored in **A3**; GET `?key=subtopic_favorite` returns that string.

Practice POST: `key=subtopic_practice&url=https://github.com/.../practice/data-analyst.md` → stored in **A4**; GET `?key=subtopic_practice` returns that string.

### Troubleshooting

| Symptom in Google Sheet cell A1                | Cause                                     | Fix                                                |
| ---------------------------------------------- | ----------------------------------------- | -------------------------------------------------- |
| `key=subtopic&url=<video title or plain text>` | Clipboard had the **title**, not the link | Use **Share → Copy link** before running the macro |
| Toast line 2 shows title, line 1 shows `Saved` | Same — wrong clipboard content was POSTed | Re-copy link; toast shows exactly what was sent    |
| Toast shows `Saved` but PC cannot open link    | API saved whatever was on the clipboard   | Paste clipboard in Notes first — must be a URL     |
| Sheet shows wrong text                         | Wrong thing was copied to clipboard       | Copy the link, not the title                       |

| Symptom in Google Sheet cell A2                    | Cause                                | Fix                                            |
| -------------------------------------------------- | ------------------------------------ | ---------------------------------------------- |
| `key=subtopic_article&url=<plain text, not a URL>` | Clipboard had wrong content          | Copy the article URL from the address bar      |
| Toast line 2 shows non-URL, line 1 shows `Saved`   | Same — wrong clipboard was POSTed    | Re-copy URL; toast shows exactly what was sent |
| Article saved to A1 instead of A2                  | POST body missing `subtopic_article` | Use **Set Article (direct link)** macro export |

| Symptom in Google Sheet cell A3                         | Cause                                                              | Fix                                                                                                      |
| ------------------------------------------------------- | ------------------------------------------------------------------ | -------------------------------------------------------------------------------------------------------- |
| `key=subtopic_favorite&url=<plain text, not a URL>`     | Clipboard had wrong content                                        | Copy the URL from the address bar                                                                        |
| Toast line 2 shows non-URL, line 1 shows `Saved`        | Same — wrong clipboard was POSTed                                  | Re-copy URL; toast shows exactly what was sent                                                           |
| Favorite saved to A1/A2 instead of A3                   | Favorite router **not in deployed web app** (editor save ≠ deploy) | Paste [`study-link-api/Code.gs`](study-link-api/Code.gs); **New version** deploy; verify GET URLs differ |
| A1 shows `key=subtopic_favorite&...` after Set Favorite | Same — default `doPost` wrote A1                                   | Redeploy; `GET ?key=subtopic_favorite` must not echo A1 when A3 empty                                    |

| Symptom in Google Sheet cell A4       | Cause                                        | Fix                                                                              |
| ------------------------------------- | -------------------------------------------- | -------------------------------------------------------------------------------- |
| Practice saved to A1/A3 instead of A4 | Practice router **not in deployed web app**  | Paste [`study-link-api/Code.gs`](study-link-api/Code.gs); **New version** deploy |
| A3 overwritten after palace import    | Old PC code still POSTed `subtopic_favorite` | Update AHK; import uses `STUDYLINK_KEY_PRACTICE` → A4 only                       |

| Get macro symptom                                             | Cause                                                | Fix                                                                |
| ------------------------------------------------------------- | ---------------------------------------------------- | ------------------------------------------------------------------ |
| Log: `Open Website ''` after HTTP 200                         | Text Manipulation dropped on import, or no `plain=1` | Redeploy [`Code.gs`](study-link-api/Code.gs); re-import Get macros |
| `resp` is full `key=…&url=…` not just URL                     | `plain=1` not deployed                               | **New version** deploy after pasting updated `Code.gs`             |
| Get opens `script.google.com` + connection refused to YouTube | Used `&open=1` (YouTube blocked in Google iframe)    | Use Get macro with `&plain=1`; re-import from `docs/macrodroid/`   |

---

## MacroDroid / browser — GET (read link)

**PC / AHK** — plain GET (full stored body):

| Type     | GET URL                          |
| -------- | -------------------------------- |
| YouTube  | `.../exec?key=subtopic`          |
| Article  | `.../exec?key=subtopic_article`  |
| Favorite | `.../exec?key=subtopic_favorite` |
| Practice | `.../exec?key=subtopic_practice` |

**MacroDroid Get macros** — append `&plain=1` (API returns URL only; Open Web Page uses `{lv=resp}`):

| Type     | GET URL                                  |
| -------- | ---------------------------------------- |
| YouTube  | `.../exec?key=subtopic&plain=1`          |
| Article  | `.../exec?key=subtopic_article&plain=1`  |
| Favorite | `.../exec?key=subtopic_favorite&plain=1` |
| Practice | `.../exec?key=subtopic_practice&plain=1` |

Full base: `https://script.google.com/macros/s/AKfycbzKDLbmzGF8iduyNpaUymONEkERi089rBjW0jrYUX4a8K9ornfGwYIOsgQP1K_dfaj5/exec`

Optional desktop browser bookmark: append `&open=1` for an HTML redirect (not used by MacroDroid).

---

## Quick checklist

1. Apps Script: paste [`study-link-api/Code.gs`](study-link-api/Code.gs) → **New version** deploy (must include A4 / `subtopic_practice` + `plain=1`).
2. Import all eight macros from `docs/macrodroid/`.
3. Browser test: `?key=subtopic_practice&plain=1` must return **only** `https://...` (not the full `key=…&url=…` body).
4. Phone test: Get Practice → log should show `Open Website 'https://...'`. Favorite Get must still read A3 only.

See also: [study-link-lightweight-api-setup.md](study-link-lightweight-api-setup.md), [`study-link-api/Code.gs`](study-link-api/Code.gs), and the Set/Get macros under `docs/macrodroid/` (Video, Article, Favorite, Practice).
