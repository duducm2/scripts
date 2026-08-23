# Memory Palace

Keyboard-first Memory Palace manager: AutoHotkey CRUD + CSV under `mnemonics/data/` + Python palace-image dashboard in `mnemonics/output/`.

Open via **Utility Shortcuts → [N] Memory Palace** (`#!+U`, then N).

Technique rules stay in notes `studies/technique/README.md` — this app does not change the mnemonic method. Technique docs may still say “street”; this software uses **Memory Palace** only.

## Requirements

1. **AutoHotkey v2** with `Utils.ahk` loaded (palace modules included).
2. **Python 3** (`py -3`, `py`, `python3`, or `python`).
3. Optional Plotly (dashboard is static HTML; requirements kept for parity):

```powershell
py -3 -m pip install -r mnemonics\python\requirements.txt
```

## Vocabulary (also in-app [H] Help)

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
| **Sensory**        | Modality the Story emphasizes (visual, auditory, …).                        |
| **Mapping**        | One Knowledge Atom, or up to four zoned Knowledge Atoms (Z1–Z4).            |

`atoms.csv` columns: `concept`, `quote`, `story`, `sensory` (plus ids/zones). `kind` is `single` or `zoned`. Beast is joined from `beasts.csv` via `beast_id`.

## Data layout

| Path                              | Role                                                                  |
| --------------------------------- | --------------------------------------------------------------------- |
| `mnemonics/data/*.csv`            | Source of truth (studies, palaces, beasts, atoms).                    |
| `mnemonics/data/settings.ini`     | Last study, optional NotesStudiesRoot override.                       |
| `mnemonics/data/imported/`        | Archived AI import CSVs.                                              |
| `mnemonics/output/dashboard.html` | Generated cockpit.                                                    |
| `mnemonics/output/practice/`      | Auto-synced study `.md` files + palace images (mobile/GitHub).        |
| `mnemonics/python/`               | Migrator, aggregator, chart generator, practice MD sync, prompt pack. |

`palaces.csv` columns: `id`, `study_id`, `palace_number`, `title`, `character_name`, `image_rel_path`, `depth_slots_used`, `image_prompt`.

`studies.csv` columns: `id`, `title`, `notes_rel_path`, `sort_order`, `active` (no separate slug; folder key is `notes_rel_path`).

`image_rel_path` is a relative path to the palace composite image. New attaches use `practice/images/{study}/{n}.ext` under `mnemonics/output/` (self-contained in the scripts repo). Legacy rows may still point at `notes/studies/…` until migrated. `image_prompt` stores the text used to generate that image; **empty is valid** (legacy rows migrated without prompts). Canon JSON: `studies/technique/characters.json`, `bestiary.json`.

`beasts.csv` FK is `palace_id` (row ids use `PALACE_*`).

## Main menu letters

| Key       | Module                                             |
| --------- | -------------------------------------------------- |
| D         | Dashboard (Python → Chrome)                        |
| B         | Browse (studies → palaces → beasts → atoms)        |
| I         | AI import (Desktop `PALACE_*` pack, one preview)   |
| Q         | Quick image (newest Desktop PNG/JPG → last palace) |
| G         | Practice on GitHub (synced mobile notes)           |
| H         | Glossary                                           |
| P         | Push scripts repo to cloud                         |
| Backspace | Return to Utility Shortcuts                        |
| Esc       | Close without reopening Utility Shortcuts          |

## Migrate legacy Markdown

```powershell
py -3 mnemonics\python\migrate_md_to_csv.py `
  --notes-root "C:\Users\eduev\Meu Drive\17 - Projects\notes\studies" `
  --out mnemonics\data `
  --dry-run

py -3 mnemonics\python\migrate_md_to_csv.py `
  --notes-root "C:\Users\eduev\Meu Drive\17 - Projects\notes\studies" `
  --out mnemonics\data
```

Writes `mnemonics/data/migration_report.md`. Legacy `mnemonics-*.md` files are left untouched. MD headings `## Street N` are still parsed; output is `palaces.csv` with empty `image_prompt`.

## Dashboard

Memory Palace **[D]** runs `chart_generator.py` with `--data-dir` / `--output-dir` / `--notes-root`, then opens Chrome. Pick a study in the page to view Memory Palace images (newest palace number first).

Click a palace card to open a fullscreen view: image, **Image prompt** (or empty state) with **Copy prompt**, Knowledge Atom count, **Close**, and a practice list that labels each atom’s **Beast**, **Concept**, **Quote**, **Story**, and **Sensory**.

## Practice Markdown (mobile / GitHub)

Each active study gets a Markdown file under `mnemonics/output/practice/{notes_rel_path}.md`, with palace images copied to `mnemonics/output/practice/images/`. Files sync automatically after browse CRUD, AI import **[I]**, and quick image attach **[Q]** (loading bar shown during generation).

Batch browse on GitHub:

`https://github.com/duducm2/scripts/tree/main/mnemonics/output/practice`

One-time backfill (copies images from notes and migrates `image_rel_path` in CSV):

```powershell
py -3 mnemonics\python\study_practice_md.py `
  --data-dir mnemonics\data `
  --output-dir mnemonics\output `
  --notes-root "C:\Users\eduev\Meu Drive\17 - Projects\notes\studies" `
  --sync-all --migrate-image-paths
```

Optional legacy cleanup in the notes repo (dry-run first):

```powershell
py -3 mnemonics\python\study_practice_md.py `
  --data-dir mnemonics\data `
  --output-dir mnemonics\output `
  --notes-root "C:\Users\eduev\Meu Drive\17 - Projects\notes\studies" `
  --prune-legacy-notes-md --dry-run
```

Remove `--dry-run` to delete `mnemonics-*.md` under notes studies (skips `technique/` and `portals/`).

## AI import

1. Technique prompts (Utility Shortcuts → Prompts): `5` transcript, `4` create mnemonic stories, `a` story reduction, `g` preserve background.
2. Stories / reduction deliver one downloadable **`PALACE_PACK.txt`**: human-readable `===PREVIEW===` plus three labeled CSV sections (`===FILE: PALACE_PALACES.csv===`, `BEASTS`, `ATOMS`). A `gemini-code-….txt` dump with the same markers also works. Edit the pack on Desktop if needed.
3. Separate Desktop files `PALACE_PALACES*` / `PALACE_BEASTS*` / `PALACE_ATOMS*` still import when present (preferred over pack if any exist).
4. Memory Palace **[I]** — one-shot import (palaces → beasts → atoms), combined preview, then archive under `data/imported/`. Beasts for palace ids in the pack replace existing beasts (and their atoms) for those palaces.
5. After generating a palace image: save PNG/JPG to Desktop → Memory Palace **[Q]** attaches the newest image to the last palace under `LastStudyId` (stored under `mnemonics/output/practice/images/`).

### Prompt context pack

```powershell
py -3 mnemonics\python\prompt_context_pack.py `
  --data-dir mnemonics\data `
  --study-id STUDY_ENGLISH
```

Writes slices under `mnemonics/python/packs/<study_id>/` (`palaces.csv`, `beasts.csv`, `atoms.csv`) for attach.

## Multi-PC

Data lives in the scripts repo. After pull, open Memory Palace → Dashboard once. Use **[P]** to commit and push.
