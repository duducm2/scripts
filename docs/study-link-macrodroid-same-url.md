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

Use **two macros** (duplicate your working YouTube macro; change only the body line).

| Macro                                  | HTTP | URL                      | Body (paste pattern)                    |
| -------------------------------------- | ---- | ------------------------ | --------------------------------------- |
| **YouTube** (existing, e.g. Set_Video) | POST | `/exec` (full URL above) | `key=subtopic&url=` + clipboard         |
| **Article** (new)                      | POST | **same** `/exec`         | `key=subtopic_article&url=` + clipboard |

MacroDroid settings (both macros):

- Method: **POST**
- Content-Type: `application/x-www-form-urlencoded`
- Clipboard must already hold the full `http(s)://…` URL before you run the macro
- Success response: `Saved`

**Important:** The article body must include the text `subtopic_article` (not `subtopic`). That is what sends the request to cell A2.

Example article body in MacroDroid:

```
key=subtopic_article&url={clipboard}
```

(`{clipboard}` = your variable for clipboard text.)

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
2. YouTube macro: unchanged URL; body still `key=subtopic&url=` + clipboard.
3. Article macro: **same URL**; body `key=subtopic_article&url=` + clipboard.
4. Test article GET in browser with `?key=subtopic_article` after saving once from the new macro.

See also: [study-link-lightweight-api-setup.md](study-link-lightweight-api-setup.md) (modules 3 and 4).
