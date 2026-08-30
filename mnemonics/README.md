# Memory Palace

Tasks-style Memory Palace manager: thin AutoHotkey launcher + Python server on **`127.0.0.1:8767`** + SPA under `mnemonics/web/`. CSV under `mnemonics/data/`. Practice/plan Markdown under `mnemonics/output/`.

Open via **Utility Shortcuts → [N] Memory Palace** (`#!+U`, then N), or **`#!+D` hold** (≥700 ms). Pack import stays **Import Management** (`#!+X` / Utility `[J]` → `[P]` / `[L]`). Git push stays Utility Shortcuts **`[G]`**.

Technique rules live in-repo at `mnemonics/technique/` (SSOT). The app does not change the mnemonic method. Technique docs may still say “street”; this software uses **Memory Palace** only.

## Requirements

1. **AutoHotkey v2** with `Utils.ahk` loaded (palace modules included).
2. **Python 3** (`py -3`, `py`, `python3`, or `python`).
3. Optional Plotly (legacy chart generator only; not required for the web app):

```powershell
py -3 -m pip install -r mnemonics\python\requirements.txt
```

## Host shape (primary)

| Piece | Path / port |
| --- | --- |
| Launcher | [`Utils/mnemonic_palace_launcher.ahk`](../Utils/mnemonic_palace_launcher.ahk) — ensure server, open/focus Chrome titled **Memory Palace** |
| Server | [`mnemonics/python/palace_server.py`](python/palace_server.py) on **`:8767`** (Tasks owns `:8766`) |
| Store | [`mnemonics/python/palace_store.py`](python/palace_store.py) |
| SPA | [`mnemonics/web/index.html`](web/index.html) |
| PID | `mnemonics/data/palace_server.pid` |

Deprecated as the primary open path: `file://` `%TEMP%\palace_dashboard.html` from `chart_generator.py`, and the write-only **`:8765`** `plan_save_server`. Those writes now live on `:8767` (`/api/plans/save`, `/api/palace/notes`, `/api/palace/images`).

## Vocabulary (also in-app Help)

| Term               | Meaning                                                                     |
| ------------------ | --------------------------------------------------------------------------- |
| **Study**          | Broad domain (english, german, …). Owns Memory Palaces.                     |
| **Memory Palace**  | Location; **one** generated image per palace.                               |
| **Character**      | From `characters.json`; one per Memory Palace.                              |
| **Beast**          | From `bestiary.json`; peg animal that carries a Knowledge Atom.             |
| **Knowledge Atom** | Discrete information on a Beast, made of Concept + Quote + Story + Sensory. |
| **Concept**        | Rehearsal definition of the fact.                                           |
| **Quote**          | Verbatim source payload.                                                    |
| **Story**          | Bizarre mnemonic narrative / action.                                        |
| **Sensory**        | Visual, auditory, tactile, olfactory, gustatory, or thermal channel.        |
| **Mapping**        | One atom per beast, or up to four zoned atoms (Z1–Z4).                      |

## Web app views

| View | Role |
| --- | --- |
| **Browse** | Studies → palaces → beasts → atoms CRUD |
| **Practice** | Study picker, palace cards, overlay (notes, prompt copy) |
| **Plans** | Checklist progress save, add items |
| **Links** | Study video / article / favorite (Google Docs API) |
| **Tools** | Regen Markdown, GitHub practice/plans (quick image is Import Management **`[Q]`**) |
| **Help** | Glossary + shortcuts / import pointers |
| **Method** | Technique README excerpt |

Keyboard: **Esc** back, letter hints in the nav bar (`B`/`P`/`L`/`T`/`1`). Study links: **Shift+V** / **Shift+A** / **Shift+F**.

Pack imports (PALACE_PACK / PLAN_PACK) and **Quick image**: Import Management (`#!+X` or Utility Shortcuts `[J]` → **`[P]`** / **`[L]`** / **`[Q]`**).

## Migrate legacy Markdown

One-time migrator for historical `mnemonics-*.md` archives (if you still have a copy elsewhere):

```powershell
py -3 mnemonics\python\migrate_md_to_csv.py `
  --notes-root "<path-to-folder-of-legacy-md>" `
  --out mnemonics\data
```

Writes `mnemonics/data/migration_report.md`. MD headings `## Street N` are still parsed; output is `palaces.csv` with empty `image_prompt`.

## Study plan Markdown (mobile / GitHub)

Canonical plan rows live in `mnemonics/data/plans.csv` (+ items/resources). Exports sync to `mnemonics/output/plans/{slug}.md` for mobile/GitHub.

Batch browse on GitHub:

`https://github.com/duducm2/scripts/tree/main/mnemonics/output/plans`

Regenerate all plan files:

```powershell
py -3 mnemonics\python\study_plans_md.py `
  --data-dir mnemonics\data `
  --studies-root mnemonics\studies `
  --output-dir mnemonics\output `
  --sync-all
```

**Save (web app):** Plans **Save** posts to `:8767` `/api/plans/save` (CSV + MD refresh). Palace overlay **notes** / **gallery** use `/api/palace/notes` and `/api/palace/images`. Push via Utility Shortcuts **`[G]`**.

## Technique (SSOT)

Edit method rules, prompts, and canon JSON under `mnemonics/technique/`. Optional one-way sync from an external technique folder:

```powershell
py -3 mnemonics\python\sync_technique.py `
  --source "<external-technique-folder>" `
  --dest mnemonics\technique
```

## Practice Markdown (mobile / GitHub)

Each active study gets a Markdown file under `mnemonics/output/practice/{notes_rel_path}.md`, with palace images under `mnemonics/output/practice/images/`. Files sync after browse CRUD, Import Management palace pack import **[P]**, Import Management quick image **[Q]**, and regen.

**Layout (GitHub mobile):** Collapsible Memory Palaces (`<details>`; newest open by default). Beasts as flat headings with **Concept / Quote / Story / Sensory**. Emoji markers match technique canon. Image prompts are omitted (recall-only).

Each palace block ends with **Notes** and **Gallery**. Hero scene images (`image_rel_path`, Import Management → **[Q]** Quick image) are unchanged.

Batch browse on GitHub:

`https://github.com/duducm2/scripts/tree/main/mnemonics/output/practice`

Regenerate all practice files:

```powershell
py -3 mnemonics\python\study_practice_md.py `
  --data-dir mnemonics\data `
  --output-dir mnemonics\output `
  --sync-all
```

Or use **Tools → Regen Markdown** in the web app (`POST /api/regen`).

## AI import

1. Technique prompts (Utility Shortcuts → Prompts): `5` transcript, `4` create mnemonic stories, `a` story reduction, `g` preserve background.
2. Stories / reduction deliver one downloadable **`PALACE_PACK.txt`**: human-readable `===PREVIEW===` plus three labeled CSV sections (`===FILE: PALACE_PALACES.csv===`, `BEASTS`, `ATOMS`). A `gemini-code-….txt` dump with the same markers also works. Edit the pack on Desktop if needed.
3. Separate Desktop files `PALACE_PALACES*` / `PALACE_BEASTS*` / `PALACE_ATOMS*` still import when present (preferred over pack if any exist).
4. Import Management (`#!+X` / Utility `[J]` → **[P]**) — one-shot import (palaces → beasts → atoms), combined preview, then archive under `data/imported/`. Beasts for palace ids in the pack replace existing beasts (and their atoms) for those palaces.
5. After generating a palace image: save PNG/JPG to Desktop → Import Management (`#!+X` / Utility `[J]` → **[Q]**) pick the palace missing `image_rel_path` (API `POST /api/quick-image` remains for scripts).

### Prompt context pack

See [`docs/prompt-data-output-and-finance-packs.md`](../docs/prompt-data-output-and-finance-packs.md).

## After pull

Data lives in the scripts repo. After pull, open Memory Palace once (`[N]` or hold-`#!+D`) so the launcher starts `:8767`. Use Utility **`[G]`** to commit and push.
