# Memory Palace

Keyboard-first Memory Palace manager: AutoHotkey CRUD + CSV under `mnemonics/data/` + Python palace-image dashboard in `mnemonics/output/`.

Open via **Win+Alt+Shift+X**, or **Utility Shortcuts → [N] Memory Palace** (`#!+U`, then N).

Technique rules live in-repo at `mnemonics/technique/` (SSOT). The app does not change the mnemonic method. Technique docs may still say “street”; this software uses **Memory Palace** only.

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
| `mnemonics/output/plans/`         | Auto-synced study plan `.md` files (mobile/GitHub).                   |
| `mnemonics/technique/`            | Method docs, canon JSON, prompts, research (SSOT).                    |
| `mnemonics/studies/`              | Study plans / portals still tied to active topics.                    |
| `mnemonics/_quarantine_review/`   | Suspected unused files staged for manual delete (not auto-removed).   |
| `mnemonics/python/`               | Migrator, aggregator, chart generator, practice MD sync, prompt pack. |

`palaces.csv` columns: `id`, `study_id`, `palace_number`, `title`, `character_name`, `image_rel_path`, `depth_slots_used`, `image_prompt`.

`studies.csv` columns: `id`, `title`, `notes_rel_path`, `sort_order`, `active` (no separate slug; folder key is `notes_rel_path`).

`image_rel_path` is a relative path to the palace composite image. New attaches use `practice/images/{study}/{n}.ext` under `mnemonics/output/`. `image_prompt` stores the text used to generate that image; **empty is valid**. Canon JSON: `mnemonics/technique/characters.json`, `bestiary.json`.

`NotesStudiesRoot` (optional) defaults to `mnemonics/studies/` when unset.

`beasts.csv` FK is `palace_id` (row ids use `PALACE_*`).

## Main menu

| Key       | Module                                             |
| --------- | -------------------------------------------------- |
| 1         | Study Video (open / set video link via API)        |
| 2         | Study Article (open / set article link via API)    |
| 3         | Favorite (open / set favorite link via API)        |
| D         | Dashboard (Python → Chrome)                        |
| B         | Browse (studies → palaces → beasts → atoms)        |
| I         | AI import (Desktop `PALACE_*` pack, one preview)   |
| Q         | Quick image (newest Desktop PNG/JPG → last palace) |
| G         | Practice on GitHub (synced mobile notes)           |
| O         | Plans on GitHub (synced study plan checklists)     |
| R         | Regen Markdown (force all practice + plan `.md`)   |
| H         | Glossary                                           |
| P         | Push scripts repo to cloud                         |
| Backspace | Return to Utility Shortcuts                        |
| Esc       | Close without reopening Utility Shortcuts          |

## Migrate legacy Markdown

One-time migrator for historical `mnemonics-*.md` archives (if you still have a copy elsewhere):

```powershell
py -3 mnemonics\python\migrate_md_to_csv.py `
  --notes-root "<path-to-folder-of-legacy-md>" `
  --out mnemonics\data
```

Writes `mnemonics/data/migration_report.md`. MD headings `## Street N` are still parsed; output is `palaces.csv` with empty `image_prompt`.

## Dashboard

Memory Palace **[D]** runs `chart_generator.py` with `--data-dir` / `--output-dir` (and optional `--notes-root`), then opens Chrome. Technique docs are loaded from `mnemonics/technique/` (no notes clone required). Pick a study in the page to view Memory Palace images (newest palace number first).

Click a palace card to open a fullscreen view: image, **Image prompt** (or empty state) with **Copy prompt**, Knowledge Atom count, **Close**, and a practice list that labels each atom’s **Beast**, **Concept**, **Quote**, **Story**, and **Sensory**.

**Method** button (or keyboard **M**) opens the technique docs in the same page: README (tables, mermaid workflow), research notes, prompt previews, and searchable Characters / Bestiary canon.

**Plans** button (or keyboard **P**) opens study plan checklists parsed from `mnemonics/studies/*/*-plan.md`: backlog, phased sections, checkbox todos, and collapsible resource links. Progress toggles are saved in the browser (localStorage). **Save** writes checkbox state back to the source plan `.md` files (and refreshes `output/plans/`); then use **[P] Push to cloud** for GitHub. **Reset to file** clears local-only progress.

## Study plan Markdown (mobile / GitHub)

Source plans live under `mnemonics/studies/{topic}/*-plan.md`. On each dashboard build, copies sync to `mnemonics/output/plans/{slug}.md` for batch mobile/GitHub access.

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

**Save (dashboard):** Opening the dashboard via **[D]** starts a local save server (`127.0.0.1:8765`). In the Plans panel, **Save** writes checkbox progress to `mnemonics/studies/*/*-plan.md` and refreshes `output/plans/`. Palace overlay **notes** and **gallery** saves use the same server (`/palace/notes`, `/palace/images`) and refresh `output/practice/*.md`. Then **[P] Push to cloud** commits for GitHub.

## Technique (SSOT)

Edit method rules, prompts, and canon JSON under `mnemonics/technique/`. Dashboard **[D]** uses that tree directly. Optional one-way sync from an external technique folder remains available:

```powershell
py -3 mnemonics\python\sync_technique.py `
  --source "<external-technique-folder>" `
  --dest mnemonics\technique
```

Prompt context files resolve from `mnemonics/technique/` first.

## Practice Markdown (mobile / GitHub)

Each active study gets a Markdown file under `mnemonics/output/practice/{notes_rel_path}.md`, with palace images copied to `mnemonics/output/practice/images/`. Files sync automatically after browse CRUD, AI import **[I]**, and quick image attach **[Q]** (loading bar shown during generation).

**Layout (GitHub mobile):** Export mirrors the dashboard hierarchy — collapsible Memory Palaces only (`<details>`; newest open by default). Beasts are flat headings (`### 🟧 …`, always expanded) with stacked **Concept / Quote / Story / Sensory** field blocks. Emoji markers match the dashboard/technique canon: `🟧` beast, `🟦` zone (when set), `💡` concept only, sensory channel map (`👁️👂✋👃👅🌡️`, emoji then word). Image prompts are omitted (recall-only). GitHub’s Markdown renderer does not apply dashboard CSS (no dark/gold theme); structural and label parity is intentional for phone recall.

Each palace block ends with **Notes** (`palace_notes` on `palaces.csv`) and **Gallery** (supplementary images in `palace_images.csv`, separate from the hero scene image). Order in export: Knowledge Atoms → Notes → Gallery.

### Dashboard notes and gallery

In the palace overlay (bottom sections): type notes (auto-save ~800ms or **Save notes**); manage a supplementary image gallery (**Add image**, caption, reorder, delete). Saves go to CSV via the local save server started with **[D]** (`127.0.0.1:8765`, routes `/palace/notes` and `/palace/images`) and refresh the study’s practice Markdown. Hero scene images (`image_rel_path`, **[Q]** attach) are unchanged.

Batch browse on GitHub:

`https://github.com/duducm2/scripts/tree/main/mnemonics/output/practice`

Regenerate all practice files:

```powershell
py -3 mnemonics\python\study_practice_md.py `
  --data-dir mnemonics\data `
  --output-dir mnemonics\output `
  --sync-all
```

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
