# MyNotes technique prompts (mnemonics)

This document describes how the **five mnemonic / study technique prompt entries** are resolved from the **MyNotes** repository when available, mirrored into this scripts repo when desired, and exposed in AutoHotkey (hotstrings and **Win+Alt+Shift+U** Utility Shortcuts → **Prompts**).

The prompts are the single source of truth under the notes repo:

`studies/technique/prompts/`

---

## Purpose

- **MyNotes** holds the `.txt` files. You edit them there; scripts should not duplicate their body text in code.
- **This repo** can keep a **mirror** under `prompt/technique/` so machines without the notes clone still resolve files at runtime (see resolution order below).

---

## Resolution order (runtime)

Implemented in [`env.ahk`](../env.ahk) and [`Utils.ahk`](../Utils.ahk):

1. **Environment override:** if `MYNOTES_TECHNIQUE_PROMPTS` is set to the full path of the `prompts` folder and that folder exists, it is used.
2. Else **`GetNotesRepoPath()`** + `studies\technique\prompts` (work vs personal is chosen from `NOTES_REPO_PATH_WORK` / `NOTES_REPO_PATH_PERSONAL` and `IS_WORK_ENVIRONMENT`).
3. Else the **other** machine’s notes root (if that clone exists), same subpath.

`GetTechniquePromptFilePath()` in [`Utils.ahk`](../Utils.ahk) then resolves each file: **live MyNotes folder first**, then fallback **`A_ScriptDir\assets\prompt\technique\<filename>`**.

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

## The five registered files and hotstring triggers

| File                                       | Hotstring           | Role (short)                                                             |
| ------------------------------------------ | ------------------- | ------------------------------------------------------------------------ |
| `story-prompt.txt`                         | `:o:mnemonic`       | Creating mnemonic stories                                                |
| `video-transcription-prompt.txt`           | `:o:ytranscript`    | YouTube transcript workflow                                              |
| `story-reduction-prompt.txt`               | `:o:storyreduction` | Story reduction                                                          |
| `punctual-beast-append-prompt.txt`         | `:o:punctualbeast`  | Append isolated beasts or small punctual batches into open streets       |
| `image-background-preservation-prompt.txt` | `:o:imgpreserve`    | Preserve the locked background while adding mnemonic foreground elements |

Registration lives in `InitTechniquePromptHotstrings()` in [`Utils.ahk`](../Utils.ahk).

---

## Operational note

After editing prompt files on disk, **reload** the AutoHotkey entry script (or restart scripts) if you do not use a workflow that reloads `Utils.ahk` automatically.

---

## Utility Shortcuts (#!+U)

Under **Prompts**, the five mnemonic-technique entries are grouped under a **Mnemonics technique** subsection in the selector UI (see `UtilitySelector_IsMnemonicTechniquePrompt` / reorder logic in [`Utils.ahk`](../Utils.ahk)).
