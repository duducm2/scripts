# Memory Palace

Keyboard-first Memory Palace manager: AutoHotkey CRUD + CSV under `mnemonics/data/` + Python street-image dashboard in `mnemonics/output/`.

Open via **Utility Shortcuts → [N] Memory Palace** (`#!+U`, then N).

Technique rules stay in notes `studies/technique/README.md` — this app does not change the mnemonic method.

## Requirements

1. **AutoHotkey v2** with `Utils.ahk` loaded (palace modules included).
2. **Python 3** (`py -3`, `py`, `python3`, or `python`).
3. Optional Plotly (dashboard is static HTML; requirements kept for parity):

```powershell
py -3 -m pip install -r mnemonics\python\requirements.txt
```

## Vocabulary (also in-app [H] Help)

| Term | Meaning |
|------|---------|
| **Study** | Broad domain (english, german, …). Owns streets. |
| **Memory Palace / Street** | Location; **one** generated image per street. |
| **Character** | From `characters.json`; one per street. |
| **Beast** | From `bestiary.json`; peg holder. |
| **Knowledge Atom / Topic / Subtopic** | Discrete information on a beast. |
| **Mapping** | One comprehensive atom **or** up to **four** smashed subtopics (Z1–Z4). |

## Data layout

| Path | Role |
|------|------|
| `mnemonics/data/*.csv` | Source of truth (studies, streets, beasts, atoms). |
| `mnemonics/data/settings.ini` | Last study, optional NotesStudiesRoot override. |
| `mnemonics/data/imported/` | Archived AI import CSVs. |
| `mnemonics/output/dashboard.html` | Generated cockpit. |
| `mnemonics/python/` | Migrator, aggregator, chart generator, prompt pack. |

Street images stay under **notes** `studies/<study>/images/N.png|jpg`. CSV stores paths relative to `notes/studies/`.

Canon JSON remains in notes: `studies/technique/characters.json`, `bestiary.json`.

## Main menu letters

| Key | Module |
|-----|--------|
| D | Dashboard (Python → Chrome) |
| Y | Studies |
| S | Streets |
| B | Beasts |
| A | Knowledge atoms |
| I | AI import (`PALACE_*.csv` on Desktop) |
| H | Glossary |
| P | Push scripts repo to cloud |

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

Writes `mnemonics/data/migration_report.md`. Legacy `mnemonics-*.md` files are left untouched.

## Dashboard

Memory Palace **[D]** runs `chart_generator.py` with `--data-dir` / `--output-dir` / `--notes-root`, then opens Chrome. Pick a study in the page to view street images.

## AI import

1. Technique prompts (Utility Shortcuts → Prompts): `5` transcript, `4` stories, `a` reduction, `g` background. **Ignore** punctual beast (`p`).
2. Import contract prompt **`k`** — `assets/prompt/mnemonic-atoms-import.txt`.
3. Save AI CSV to Desktop as `PALACE_ATOMS.csv` (or `PALACE_BEASTS` / `PALACE_STREETS`).
4. Memory Palace **[I]** → preview → commit → archive under `data/imported/`.

### Prompt context pack

```powershell
py -3 mnemonics\python\prompt_context_pack.py `
  --data-dir mnemonics\data `
  --study-id STUDY_ENGLISH
```

Writes slices under `mnemonics/python/packs/<study_id>/` for attach.

## Multi-PC

Data lives in the scripts repo. After pull, open Memory Palace → Dashboard once. Use **[P]** to commit and push.
