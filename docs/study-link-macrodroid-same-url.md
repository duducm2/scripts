# MacroDroid — same web app URL for YouTube, article, and favorite

One deployment URL for all link types. MacroDroid and the PC scripts use the **same** `/exec` address; only the **POST body** and **GET query** change.

**Web app URL (POST and GET base):**

`https://script.google.com/macros/s/AKfycbzKDLbmzGF8iduyNpaUymONEkERi089rBjW0jrYUX4a8K9ornfGwYIOsgQP1K_dfaj5/exec`

(Also set as `STUDY_LINKS_API_URL` in `StudyLinkHelpers.ahk`.)

---

## How the server picks YouTube vs article vs favorite

After you add the wiring in [study-link-lightweight-api-setup.md](study-link-lightweight-api-setup.md):

| Request                                    | Routed to          | Sheet cell |
| ------------------------------------------ | ------------------ | ---------- |
| POST body contains `subtopic_article`      | `doPostArticle`    | **A2**     |
| POST body contains `subtopic_favorite`     | `doPostFavorite`   | **A3**     |
| Any other POST (e.g. `key=subtopic&url=…`) | `doPost` (YouTube) | **A1**     |
| GET `?key=subtopic_article`                | `doGetArticle`     | **A2**     |
| GET `?key=subtopic_favorite`               | `doGetFavorite`    | **A3**     |
| GET without that key (or `?key=subtopic`)  | `doGet` (YouTube)  | **A1**     |

Router order in `doPost`: check `subtopic_article` first, then `subtopic_favorite`, then default YouTube.

You do **not** need a second URL. Differentiate with the **key name in the POST body** (and `?key=` on GET).

---

## MacroDroid — SET (save link from clipboard)

**Canonical macro exports:**

- YouTube: [`macrodroid/Set_Video_(direct_link).macro`](<macrodroid/Set_Video_(direct_link).macro>)
- Article: [`macrodroid/Set_Article_(direct_link).macro`](<macrodroid/Set_Article_(direct_link).macro>)
- Favorite SET: [`macrodroid/Set_Favorite_(direct_link).macro`](<macrodroid/Set_Favorite_(direct_link).macro>)
- Favorite GET: [`macrodroid/Get_Favorite_(direct_link).macro`](<macrodroid/Get_Favorite_(direct_link).macro>)

Import all in MacroDroid (replaces older Set_Video / manual duplicates).

### Android workflow (YouTube)

1. In the **YouTube** app, open the video.
2. Tap **Share** → **Copy link** (not the video title, description, or transcript).
3. **Verify:** paste in Notes — must start with `http` (if you see the title, copy the link again).
4. Run **Set Video (direct link)** from your MacroDroid drawer.

The macro **refreshes the clipboard** (required on Android 10+), stores it in `clip`, POSTs `key=subtopic&url={lv=clip}` — same shape as PC `StudyLink_Set`. No URL validation; you must copy the link yourself (**Share → Copy link**, not the title).

**Macro actions (in order):** Clipboard Refresh → Set `clip` = `{clipboard}` → HTTP POST → Toast `{lv=resp}` + `{lv=clip}`.

**Android 10+:** MacroDroid cannot read the system clipboard from the background without the **Clipboard Refresh** action first ([MacroDroid wiki](https://macrodroidforum.com/wiki/index.php/Action:_Clipboard_Refresh)). Without it, `{clipboard}` may be empty or stale.

### Android workflow (Article)

1. In your **browser** (or any app), copy the article URL from the address bar or **Share → Copy link**.
2. **Verify:** paste in Notes — must start with `http`.
3. Run **Set Article (direct link)** from your MacroDroid drawer.

Same macro actions as YouTube; POST body is `key=subtopic_article&url={lv=clip}` → Apps Script routes to **A2**.

### Android workflow (Favorite)

**Set:** copy any URL → verify in Notes (`http`) → run **Set Favorite (direct link)** → toast shows `Saved` + clip → **A3**.

**Get:** run **Get Favorite (direct link)** → fetches A3 → extracts URL after `url=` → opens in browser. If A3 is empty, toast: _No favorite link stored_.

| Macro                                   | HTTP | URL                      | Body / query                                                       |
| --------------------------------------- | ---- | ------------------------ | ------------------------------------------------------------------ |
| **YouTube** (Set Video direct link)     | POST | `/exec` (full URL above) | `key=subtopic&url=` + `{lv=clip}` after Clipboard Refresh          |
| **Article** (Set Article direct link)   | POST | **same** `/exec`         | `key=subtopic_article&url=` + `{lv=clip}` after Clipboard Refresh  |
| **Favorite** (Set Favorite direct link) | POST | **same** `/exec`         | `key=subtopic_favorite&url=` + `{lv=clip}` after Clipboard Refresh |
| **Favorite** (Get Favorite direct link) | GET  | **same** `/exec`         | `?key=subtopic_favorite` → extract `url=` → Open Web Page          |

MacroDroid settings (both macros):

- Method: **POST**
- Content-Type: `application/x-www-form-urlencoded`
- Clipboard must hold the full link **before** you run the macro
- Success response: `Saved`
- Success toast: `{lv=resp}` then `{lv=clip}` — check the second line starts with `http`

**Important:** The article body must include `subtopic_article`; the favorite body must include `subtopic_favorite`. That is what routes to A2 and A3.

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

| Symptom in Google Sheet cell A3                         | Cause                                                              | Fix                                                                                     |
| ------------------------------------------------------- | ------------------------------------------------------------------ | --------------------------------------------------------------------------------------- |
| `key=subtopic_favorite&url=<plain text, not a URL>`     | Clipboard had wrong content                                        | Copy the URL from the address bar                                                       |
| Toast line 2 shows non-URL, line 1 shows `Saved`        | Same — wrong clipboard was POSTed                                  | Re-copy URL; toast shows exactly what was sent                                          |
| Favorite saved to A1/A2 instead of A3                   | Favorite router **not in deployed web app** (editor save ≠ deploy) | Replace full `Code.gs` from setup guide; **New version** deploy; verify GET URLs differ |
| A1 shows `key=subtopic_favorite&...` after Set Favorite | Same — default `doPost` wrote A1                                   | Redeploy; `GET ?key=subtopic_favorite` must not echo A1 when A3 empty                   |
| Get Favorite toast: _No favorite link stored_           | A3 empty or GET failed                                             | Run Set Favorite first; redeploy Apps Script                                            |
| Get Favorite does not open browser                      | Text extract failed on GET response                                | Rebuild extract step in MacroDroid UI and re-export                                     |

---

## MacroDroid / browser — GET (read link)

| Type     | Paste in browser (or MacroDroid GET action)                                                                                              |
| -------- | ---------------------------------------------------------------------------------------------------------------------------------------- |
| YouTube  | `https://script.google.com/macros/s/AKfycbzKDLbmzGF8iduyNpaUymONEkERi089rBjW0jrYUX4a8K9ornfGwYIOsgQP1K_dfaj5/exec?key=subtopic`          |
| Article  | `https://script.google.com/macros/s/AKfycbzKDLbmzGF8iduyNpaUymONEkERi089rBjW0jrYUX4a8K9ornfGwYIOsgQP1K_dfaj5/exec?key=subtopic_article`  |
| Favorite | `https://script.google.com/macros/s/AKfycbzKDLbmzGF8iduyNpaUymONEkERi089rBjW0jrYUX4a8K9ornfGwYIOsgQP1K_dfaj5/exec?key=subtopic_favorite` |

Opening the URL with **no** `?key=` reads **A1** (YouTube), which matches what you see today when the sheet returns something like:

`key=subtopic&url=https://youtube.com/live/...`

---

## Quick checklist

1. Apps Script: `doPostFavorite` / `doGetFavorite` + router lines in `doPost` / `doGet` — then **redeploy**.
2. YouTube macro: import [`macrodroid/Set_Video_(direct_link).macro`](<macrodroid/Set_Video_(direct_link).macro>).
3. Article macro: import [`macrodroid/Set_Article_(direct_link).macro`](<macrodroid/Set_Article_(direct_link).macro>).
4. Favorite macros: import [`Set_Favorite_(direct_link).macro`](<macrodroid/Set_Favorite_(direct_link).macro>) and [`Get_Favorite_(direct_link).macro`](<macrodroid/Get_Favorite_(direct_link).macro>).
5. Test favorite GET in browser with `?key=subtopic_favorite` after saving once from Set Favorite.

See also: [study-link-lightweight-api-setup.md](study-link-lightweight-api-setup.md), [`Set_Video_(direct_link).macro`](<macrodroid/Set_Video_(direct_link).macro>), [`Set_Article_(direct_link).macro`](<macrodroid/Set_Article_(direct_link).macro>), [`Set_Favorite_(direct_link).macro`](<macrodroid/Set_Favorite_(direct_link).macro>), [`Get_Favorite_(direct_link).macro`](<macrodroid/Get_Favorite_(direct_link).macro>).
