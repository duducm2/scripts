# MacroDroid — same web app URL for YouTube and article

One deployment URL for both link types. MacroDroid and the PC scripts use the **same** `/exec` address; only the **POST body** and **GET query** change.

**Web app URL (POST and GET base):**

`https://script.google.com/macros/s/AKfycbzzkjpT_47W0TwcjwEulzkV9l5xTtqcwWJmF0h-B-11SwiL_49SPhKXnj3PTsgFUZcp/exec`

(Also set as `STUDY_LINKS_API_URL` in `StudyLinkHelpers.ahk`.)

---

## How the server picks YouTube vs article

After you add the article wiring in [study-link-lightweight-api-setup.md](study-link-lightweight-api-setup.md):

| Request                                    | Routed to          | Sheet cell |
| ------------------------------------------ | ------------------ | ---------- |
| POST body contains `subtopic_article`      | `doPostArticle`    | **A2**     |
| Any other POST (e.g. `key=subtopic&url=…`) | `doPost` (YouTube) | **A1**     |
| GET `?key=subtopic_article`                | `doGetArticle`     | **A2**     |
| GET without that key (or `?key=subtopic`)  | `doGet` (YouTube)  | **A1**     |

You do **not** need a second URL. Differentiate with the **key name in the POST body** (and `?key=` on GET).

---

## MacroDroid — SET (save link from clipboard)

**Canonical macro exports:**

- YouTube: [`macrodroid/Set_Video_(direct_link).macro`](<macrodroid/Set_Video_(direct_link).macro>)
- Article: [`macrodroid/Set_Article_(direct_link).macro`](<macrodroid/Set_Article_(direct_link).macro>)

Import both in MacroDroid (replaces older Set_Video / manual duplicates).

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

| Macro                                 | HTTP | URL                      | Body (paste pattern)                                              |
| ------------------------------------- | ---- | ------------------------ | ----------------------------------------------------------------- |
| **YouTube** (Set Video direct link)   | POST | `/exec` (full URL above) | `key=subtopic&url=` + `{lv=clip}` after Clipboard Refresh         |
| **Article** (Set Article direct link) | POST | **same** `/exec`         | `key=subtopic_article&url=` + `{lv=clip}` after Clipboard Refresh |

MacroDroid settings (both macros):

- Method: **POST**
- Content-Type: `application/x-www-form-urlencoded`
- Clipboard must hold the full link **before** you run the macro
- Success response: `Saved`
- Success toast: `{lv=resp}` then `{lv=clip}` — check the second line starts with `http`

**Important:** The article body must include the text `subtopic_article` (not `subtopic`). That is what sends the request to cell A2.

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

---

## MacroDroid / browser — GET (read link)

| Type    | Paste in browser (or MacroDroid GET action)                                                                                             |
| ------- | --------------------------------------------------------------------------------------------------------------------------------------- |
| YouTube | `https://script.google.com/macros/s/AKfycbzzkjpT_47W0TwcjwEulzkV9l5xTtqcwWJmF0h-B-11SwiL_49SPhKXnj3PTsgFUZcp/exec?key=subtopic`         |
| Article | `https://script.google.com/macros/s/AKfycbzzkjpT_47W0TwcjwEulzkV9l5xTtqcwWJmF0h-B-11SwiL_49SPhKXnj3PTsgFUZcp/exec?key=subtopic_article` |

Opening the URL with **no** `?key=` reads **A1** (YouTube), which matches what you see today when the sheet returns something like:

`key=subtopic&url=https://youtube.com/live/...`

---

## Quick checklist

1. Apps Script: `doPostArticle` / `doGetArticle` + two `if` lines in `doPost` / `doGet` — then **redeploy**.
2. YouTube macro: import [`macrodroid/Set_Video_(direct_link).macro`](<macrodroid/Set_Video_(direct_link).macro>).
3. Article macro: import [`macrodroid/Set_Article_(direct_link).macro`](<macrodroid/Set_Article_(direct_link).macro>).
4. Test article GET in browser with `?key=subtopic_article` after saving once from the article macro.

See also: [study-link-lightweight-api-setup.md](study-link-lightweight-api-setup.md) (modules 3 and 4), [`Set_Video_(direct_link).macro`](<macrodroid/Set_Video_(direct_link).macro>), [`Set_Article_(direct_link).macro`](<macrodroid/Set_Article_(direct_link).macro>).
