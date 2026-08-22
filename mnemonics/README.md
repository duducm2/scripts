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

| Path                              | Role                                                |
| --------------------------------- | --------------------------------------------------- |
| `mnemonics/data/*.csv`            | Source of truth (studies, palaces, beasts, atoms).  |
| `mnemonics/data/settings.ini`     | Last study, optional NotesStudiesRoot override.     |
| `mnemonics/data/imported/`        | Archived AI import CSVs.                            |
| `mnemonics/output/dashboard.html` | Generated cockpit.                                  |
| `mnemonics/python/`               | Migrator, aggregator, chart generator, prompt pack. |

`palaces.csv` columns: `id`, `study_id`, `palace_number`, `title`, `character_name`, `image_rel_path`, `depth_slots_used`, `image_prompt`.

`studies.csv` columns: `id`, `title`, `notes_rel_path`, `sort_order`, `active` (no separate slug; folder key is `notes_rel_path`).

`image_rel_path` is a relative path to the palace composite image. `image_prompt` stores the text used to generate that image; **empty is valid** (legacy rows migrated without prompts). Canon JSON: `studies/technique/characters.json`, `bestiary.json`.

`beasts.csv` FK is `palace_id` (row ids use `PALACE_*`).

## Main menu letters

| Key       | Module                                    |
| --------- | ----------------------------------------- |
| D         | Dashboard (Python → Chrome)               |
| Y         | Studies                                   |
| S         | Palaces (Memory Palaces and images)       |
| B         | Beasts                                    |
| A         | Knowledge atoms                           |
| I         | AI import (`PALACE_*.csv` on Desktop)     |
| H         | Glossary                                  |
| P         | Push scripts repo to cloud                |
| Backspace | Return to Utility Shortcuts               |
| Esc       | Close without reopening Utility Shortcuts |

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

## AI import

1. Technique prompts (Utility Shortcuts → Prompts): `5` transcript, `4` stories, `a` reduction, `g` background. **Ignore** punctual beast (`p`).
2. Import contract prompts:
   - **`k`** — `assets/prompt/mnemonic-atoms-import.txt` (atoms)
   - Palaces — `assets/prompt/mnemonic-palaces-import.txt` → Desktop `PALACE_PALACES.csv`
3. Save AI CSV to Desktop as `PALACE_ATOMS.csv` (or `PALACE_BEASTS` / `PALACE_PALACES`). Legacy `PALACE_STREETS*` filenames still import.
4. Memory Palace **[I]** → preview → commit → archive under `data/imported/`.

### Prompt context pack

```powershell
py -3 mnemonics\python\prompt_context_pack.py `
  --data-dir mnemonics\data `
  --study-id STUDY_ENGLISH
```

Writes slices under `mnemonics/python/packs/<study_id>/` (`palaces.csv`, `beasts.csv`, `atoms.csv`) for attach.

## Multi-PC

Data lives in the scripts repo. After pull, open Memory Palace → Dashboard once. Use **[P]** to commit and push.
