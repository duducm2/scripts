# Prompt data output and finance/mnemonic packs

Documentation for future agents working on Utility Shortcuts prompts, AIB delivery, and Finance / Memory Palace pack imports.

## Summary

1. **Prompt Manager metadata** decides whether AIB must emit structured data and whether delivery is a **downloadable `.txt` file** or a **single code fence**.
2. At paste/send time, AHK **injects a short DATA OUTPUT CONTRACT** from that metadata (runtime authority for file vs code).
3. Pack bodies use `===PREVIEW===` / `===FILE: …csv===` markers. Importers accept Desktop `.txt` packs and convert CSV sections locally.
4. Never claim the companion wrote to the user’s Desktop/disk; Gemini/Copilot sandboxes are not the PC.

---

## Prompt Manager metadata

**Store:** [`assets/data/prompts.ini`](../assets/data/prompts.ini)  
**Load/Save/Normalize:** [`Utils/prompt_data.ahk`](../Utils/prompt_data.ahk)  
**Editor:** [`Utils/prompt_editor_gui.ahk`](../Utils/prompt_editor_gui.ahk)  
**List column Out:** [`Utils/hotstring_selector_gui.ahk`](../Utils/hotstring_selector_gui.ahk), [`Utils/hotstring_selector_handlers_02.ahk`](../Utils/hotstring_selector_handlers_02.ahk)

| INI key             | Values          | Meaning                                                                       |
| ------------------- | --------------- | ----------------------------------------------------------------------------- |
| `ExpectsDataOutput` | `0` / `1`       | Prompt expects structured data from AIB                                       |
| `DataOutputFormat`  | `file` / `code` | Only when expects=1. `file` = download chip `.txt`; `code` = one marked fence |

ListView **Out** column: blank, `txt·file`, or `txt·code`.

**Seeded `ExpectsDataOutput=1` + `DataOutputFormat=file`:**

- `finance-daily-transactions.txt` (char `d`)
- `finance-monthly-investments.txt` (char `m`)
- `story-prompt.txt`, `story-reduction-prompt.txt`, `plan-prompt.txt`

Switch to `code` in the editor when Gemini cannot attach downloads.

### Injected contract (runtime)

`PromptData_AppendDataOutputDirective(body, prompt)` prepends a short mandatory block when expects=1.

Call sites:

- [`Utils/prompt_render.ahk`](../Utils/prompt_render.ahk) — `PromptRender_Prepare` (Utility Shortcuts paste / L-arm companion)
- [`Utils/d2c_flow_manager.ahk`](../Utils/d2c_flow_manager.ahk) — finance daily **[D]** after `ReadBody`, before combine with dictation

The long FILE DELIVERY PROTOCOL inside each `.txt` body is documentation/fallback; the injected block is the runtime authority for **file vs code**.

---

## Pack body layout (shared)

Used by finance and mnemonic pack prompts:

```
===PREVIEW===
…
===END_PREVIEW===

===FILE: <NAME>.csv===
<header + CSV rows>
===END_FILE===
```

(`---` markers also accepted by some importers.)

Extension convention for downloadable packs: **`.txt`**. App code extracts CSV and materializes temp CSV for parsing.

---

## Finance import (txt → CSV)

**Import:** [`Utils/finance_import.ahk`](../Utils/finance_import.ahk)  
**Helpers:** [`Utils/finance_helpers.ahk`](../Utils/finance_helpers.ahk) (`Finance_ImportAccountLabel`, `Finance_NameOrUnknown`)  
**Transactions list:** [`Utils/finance_transactions.ahk`](../Utils/finance_transactions.ahk)  
**Launcher copy:** [`Utils/finance_launcher.ahk`](../Utils/finance_launcher.ahk)

Flow:

1. Discover newest `FINANCE_DAILY*.txt` / `FINANCE_MONTHLY*.txt` on Desktop (then `.csv` / `.ini` / `gemini-code*.txt`).
2. `Finance_MaterializeAiCsv` — extract `===FILE: …===` (tries `.csv` then `.txt` section name), strip markdown fences via `Chr(96)` fence helper (do **not** put raw ` ``` ` inside AHK string literals — backtick is the escape char), or fall back to header row after stripping PREVIEW.
3. Parse CSV → editable/confirm UI with **resolved names and BRL amounts** (not raw ids).
4. Daily: `Finance_ImportConfirmEditable`; monthly: multi-column confirm.
5. Archive source `.txt`; write app CSVs.

Post-import daily opens Transactions; card expenses show **card name**; transfers show `source → dest`.

---

## Pack prompt delivery prose

Files (prefer download; if attach fails, one marked fence; never fake disk save):

- [`assets/prompt/finance-daily-transactions.txt`](../assets/prompt/finance-daily-transactions.txt)
- [`assets/prompt/finance-monthly-investments.txt`](../assets/prompt/finance-monthly-investments.txt)
- [`mnemonics/technique/prompts/story-prompt.txt`](../mnemonics/technique/prompts/story-prompt.txt)
- [`mnemonics/technique/prompts/story-reduction-prompt.txt`](../mnemonics/technique/prompts/story-reduction-prompt.txt)
- [`mnemonics/technique/prompts/plan-prompt.txt`](../mnemonics/technique/prompts/plan-prompt.txt)

Human footers: Quick Download on chip, or copy fence → save as pack name on Desktop → import.

Mnemonic import still uses [`Utils/mnemonic_palace_import.ahk`](../Utils/mnemonic_palace_import.ahk) (same marker style).

---

## Related docs

- Dictation **[D]** path: [`docs/dictation-to-gemini-cursor-flow.md`](dictation-to-gemini-cursor-flow.md)
- Technique prompt resolution: [`docs/mynotes-technique-prompts.md`](mynotes-technique-prompts.md)
- Companion routing: [`docs/global-ai-companion-routing.md`](global-ai-companion-routing.md)

---

## Pitfalls for agents

- **AHK strings:** never write markdown fence literals as `"```"`; build with `Chr(96)` (see `Finance_MdFence`).
- **Do not** treat “I saved to your Desktop” as success; companions have no disk write to the user PC.
- Changing `DataOutputFormat` in the Prompt Manager is enough to flip file vs code without rewriting every pack prompt body.
- Saving prompts via the editor rewrites `prompts.ini`; preserve `ExpectsDataOutput` / `DataOutputFormat` in Load/Save/Normalize/`PromptFromEditorResult`.
