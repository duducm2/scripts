# MyNotes technique prompts (mnemonics)

This document describes how the **mnemonic / study technique prompt entries** are resolved from the **MyNotes** repository when available, mirrored into this scripts repo when desired, and exposed in AutoHotkey (**Win+Alt+Shift+U** Utility Shortcuts → **Prompts**).

The prompts are the single source of truth under the notes repo:

`studies/technique/prompts/`

---

## Purpose

- **MyNotes** holds the `.txt` files. You edit them there; scripts should not duplicate their body text in code.
- **This repo** can keep a **mirror** under `prompt/technique/` so machines without the notes clone still resolve files at runtime (see resolution order below).

---

## Resolution order (runtime)

Implemented in [`env.ahk`](../env.ahk) and [`Utils/hotstrings_core.ahk`](../Utils/hotstrings_core.ahk) (`GetTechniquePromptFilePath`):

1. **Environment override:** if `MYNOTES_TECHNIQUE_PROMPTS` is set to the full path of the `prompts` folder and that folder exists, it is used.
2. Else **`GetNotesRepoPath()`** + `studies\technique\prompts` (work vs personal is chosen from `NOTES_REPO_PATH_WORK` / `NOTES_REPO_PATH_PERSONAL` and `IS_WORK_ENVIRONMENT`).
3. Else the **other** machine’s notes root (if that clone exists), same subpath.

`GetTechniquePromptFilePath()` then resolves each file: **live MyNotes folder first**, then fallback **`A_ScriptDir\assets\prompt\technique\<filename>`**.

---

## Work vs personal paths

Same idea as [`Act.ahk`](../Act.ahk) (`notesFolder`):

| Environment | Notes root (typical)                                 |
| ----------- | ---------------------------------------------------- |
| Work        | `C:\Users\fie7ca\OneDrive - Bosch Group\14-my-notes` |
| Personal    | `C:\Users\eduev\Meu Drive\17 - Projects\notes`       |

Adjust in [`env.ahk`](../env.ahk) if your layout differs.

---

## Scripts-repo mirror and automation

| Mechanism                                                                                       | Role                                                                                                                                      |
| ----------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------- |
| [`infra/ipc/Sync-MyNotesTechniquePrompts.ps1`](../infra/ipc/Sync-MyNotesTechniquePrompts.ps1)   | Copies the registered technique prompt files into `assets/prompt/technique/`. Use **`-Commit`** to commit the mirror in the scripts repo. |
| [`infra/ipc/Watch-MyNotesTechniquePrompts.ps1`](../infra/ipc/Watch-MyNotesTechniquePrompts.ps1) | Optional **debounced** watcher on the MyNotes prompts folder; re-runs sync when files change.                                             |
| [`Act.ahk`](../Act.ahk)                                                                         | After **`git pull`** on the notes repo, runs the sync script with **`-Commit`** so the mirror stays aligned with MyNotes.                 |

---

## The registered technique files

| File                                       | Role (short)                                                             |
| ------------------------------------------ | ------------------------------------------------------------------------ |
| `story-prompt.txt`                         | Creating mnemonic stories                                                |
| `video-transcription-prompt.txt`           | YouTube transcript workflow                                              |
| `concept-curation-prompt.txt`              | Curate knowledge atoms from long transcripts/text (before story)         |
| `story-reduction-prompt.txt`               | Story reduction                                                          |
| `punctual-beast-append-prompt.txt`         | Append isolated beasts or small punctual batches into open streets       |
| `image-background-preservation-prompt.txt` | Preserve the locked background while adding mnemonic foreground elements |

Registration lives in [`assets/data/prompts.ini`](../assets/data/prompts.ini): each technique row uses `Source=technique` and `FilePath=<basename>`. The Utility Shortcuts Prompts view loads the file **at paste time** via `PromptData_ReadBody()` in [`Utils/prompt_data.ahk`](../Utils/prompt_data.ahk). Add/edit/delete of metadata is done in the Prompts ListView (Insert / E / Delete); the `.txt` file is never deleted.

Pack prompts (`story-prompt`, `story-reduction-prompt`, `plan-prompt`) also use `ExpectsDataOutput` / `DataOutputFormat` so AIB delivery is **file** or **code**. `concept-curation-prompt.txt` is seeded as **code** (one grab-able fence). See [`prompt-data-output-and-finance-packs.md`](prompt-data-output-and-finance-packs.md).

Before pasting from Utility Shortcuts → Prompts, a 3s banner appears **immediately** after you pick a prompt (before context attach): **[Y]** include human reminders (below `---`), **[Esc]** paste stripped (default), or **[S]** paste stripped and send Enter. Attach/paste run after your choice. Shift-keys composer strip still uses `ReplaceComposerWithStrippedReminders`.

---

## Operational note

Prompt **file contents** are read when you paste, so editing a `.txt` does not require a script reload. Adding or renaming entries in `prompts.ini` (or via the UI) is picked up the next time the selector opens (`PromptData_Load` reloads on file mtime).

---

## Utility Shortcuts (#!+U)

Under **Prompts**, entries are sorted by the INI `Category` attribute then `Name`. Migrated technique prompts use `Category=Mnemonic`; the other built-in prompts use `Category=General`.
