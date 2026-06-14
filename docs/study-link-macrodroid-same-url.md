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

**Canonical YouTube macro export:** [`macrodroid/Set_Video_(direct_link).macro`](<macrodroid/Set_Video_(direct_link).macro>) — import this file in MacroDroid (replaces older Set_Video macros).

### Android workflow (YouTube)

1. In the **YouTube** app, open the video.
2. Tap **Share** → **Copy link** (not the video title, description, or transcript).
3. Run **Set Video (direct link)** from your MacroDroid drawer (within ~20 seconds).

The macro **polls the clipboard**, sets `url` to the **full clipboard text**, then validates with **Compare Values** wildcards (`*youtu.be*` / `*youtube.com*`). No regex extract (MacroDroid extract was unreliable). If validation fails, the error toast includes `Got: {lv=clip}` so you can see what MacroDroid actually read.

Use **two macros** for YouTube + article (duplicate the YouTube macro; change only the POST body key).

| Macro                               | HTTP | URL                      | Body (paste pattern)                    |
| ----------------------------------- | ---- | ------------------------ | --------------------------------------- |
| **YouTube** (Set Video direct link) | POST | `/exec` (full URL above) | `key=subtopic&url=` + validated URL     |
| **Article** (duplicate)             | POST | **same** `/exec`         | `key=subtopic_article&url=` + clipboard |

MacroDroid settings (both macros):

- Method: **POST**
- Content-Type: `application/x-www-form-urlencoded`
- Clipboard must hold the full `http(s)://…` URL (YouTube macro extracts and validates automatically)
- Success response: `Saved`
- YouTube macro success toast shows the saved URL; error toasts if clipboard has no URL or a non-YouTube link

**Important:** The article body must include the text `subtopic_article` (not `subtopic`). That is what sends the request to cell A2.

Example article body in MacroDroid:

```
key=subtopic_article&url={clipboard}
```

(`{clipboard}` = your variable for clipboard text.)

### Troubleshooting

| Symptom in Google Sheet cell A1                | Cause                                                                          | Fix                                                                                                                |
| ---------------------------------------------- | ------------------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------ |
| `key=subtopic&url=<video title or plain text>` | Clipboard had the **title**, not the link; old macro POSTed any clipboard text | Re-import [`Set_Video_(direct_link).macro`](<macrodroid/Set_Video_(direct_link).macro>); use **Share → Copy link** |
| Macro toast: _Not YouTube. Got: …_             | Clipboard had text but not a YouTube URL                                       | **Share → Copy link**; check the `Got:` snippet in the toast                                                       |
| Macro toast: _Clipboard empty_                 | MacroDroid read empty clipboard                                                | Copy link, then run macro within ~20s (or grant MacroDroid clipboard access)                                       |
| Sheet unchanged, toast shows error             | Validation blocked POST (expected)                                             | Copy the correct YouTube link                                                                                      |

Filters and delays were **not** the root cause — the API accepted the POST; the clipboard simply did not contain a URL.

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
2. YouTube macro: import [`macrodroid/Set_Video_(direct_link).macro`](<macrodroid/Set_Video_(direct_link).macro>); body `key=subtopic&url=` + validated YouTube URL.
3. Article macro: **same URL**; body `key=subtopic_article&url=` + clipboard.
4. Test article GET in browser with `?key=subtopic_article` after saving once from the new macro.

See also: [study-link-lightweight-api-setup.md](study-link-lightweight-api-setup.md) (modules 3 and 4), [`macrodroid/Set_Video_(direct_link).macro`](<macrodroid/Set_Video_(direct_link).macro>).
