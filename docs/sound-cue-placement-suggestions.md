# Sound cue placement map

Use this file as a scratchpad: which actions deserve a distinct chime, which asset name to add under `sounds/`, and where in code to hook it (always behind `IsSoundEnabled()` unless noted).

---

## Conventions in this repo

| Mechanism | Where |
|-----------|--------|
| Global sound on/off | `IsSoundEnabled()` / `ToggleSoundState()` in [`../Utils.ahk`](../Utils.ahk) (`data/settings.ini`, `SoundEnabled`) |
| Master volume target | `ApplyScriptMasterVolumeTarget()` / `SCRIPT_MASTER_VOLUME_PERCENT` in Utils |
| Quiet confirm chimes | `PlayCleaningDesktopSound()` pattern (WMP internal volume; no master ducking) |

---

## Inventory: sounds already wired (grep-backed)

Fill in filenames you keep locally if they differ from the list.

| Sound file (under `sounds/`) | Role |
|------------------------------|------|
| `gemini-focused.wav` | Gemini / Cursor focus chimes |
| `gemini-completion.wav` | Task / copy / completion |
| `copy.wav` | Copy actions |
| `handy-model-chosen.mp3` | Handy model selection |
| `pre-movement.wav` | Before mouse move / “hands off” style flows |
| `cleaning-desktop.wav` | Clipboard-clean / desktop-clean confirm (Y) |
| `robots-are-working.wav` / `no-robot-working.wav` | AI working check |
| `quick-update-success.wav` | After Quick Update (`/Updated`) |
| `print-screen.wav` | PrintScreen path |
| `speach-start.wav` / `speach-finished.wav` | Dictation |
| `into-cursor-textfield.wav` | WindowManagement → Cursor field |
| `pomodo-start.wav` | Pomodoro start (AppLaunchers) |
| `fastcopy-start.mp3` / `fastcopy-finish.mp3` | Fast Copy mode |
| `commit-start.wav` | Commit flows (Shift keys) |
| `favorite-set.wav` | After Alt+Q marks focused Clip Angel clip as favorite (`MarkLastClipAsFavorite`) |
| System `*16` / `*64` | Via `ScriptSoundPlaySystem` (gated) |

`quick-update-failure.wav` exists in tree; confirm whether any path plays it—if not, candidate for failed Quick Update.

---

## Suggested placements (no code added yet)

Check boxes when you add an asset and wire it.

### Win+Alt+Shift macro outcomes ([`../Utils.ahk`](../Utils.ahk) `InitMacros`)

- [ ] **Quick Update** — start vs success already have loading + `quick-update-success.wav`; optional: soft tick when PowerShell handoff starts (`QuickUpdateScripts`).
- [ ] **Add word to Handy** — success / failure (`AddWordToHandy`).
- [ ] **Toggle Outlook & Teams** — distinct open vs close (currently banners; `ToggleOutlookAndTeams`).
- [ ] **Clean clipboard** — differentiate: user pressed Y (already has cleaning chime path) vs timeout auto-run (`CleanClipboard_OnTimeout` has no chime—intentional; optional subtle cue).
- [ ] **Toggle sound** — very short earcon when toggling (`ToggleSoundState`) so you hear mode even if UI is missed.
- [ ] **AI working?** — already uses `PlayAiWorkingStateSound`; optional third “unknown/error” blip in catch.
- [x] **Mark last clip favorite** — `favorite-set.wav` after successful Alt+Q (`MarkLastClipAsFavorite`); errors stay banner-only; “already a favorite” has no chime.
- [ ] **Desktop to Recycle** — success vs cancel vs error (`DesktopToRecycle_*`).

### Clip Angel & clipboard

- [ ] Merge clips complete (`ShowCenteredOverlay_Utils` success paths in merge flow).
- [x] Newly marked favorite — `favorite-set.wav` in `MarkLastClipAsFavorite` (already-favorite branch stays silent).

### Gemini & Cursor bridges

- [ ] Failed IPC / “Gemini not running” vs timeout vs empty clipboard (several `ShowCenteredOverlay_Utils` error branches in Utils/Gemini).
- [ ] Dictation → Gemini confirm banner: Y / S / N / timeout (distinct optional; avoid noise).

### Dictation ([`../Utils.ahk`](../Utils.ahk))

- [ ] Clipboard changed after stop (completion chime already exists—only add if you want a second “pasted” cue).
- [ ] Hotkey ownership / mutex failure (`g_DictationHotkeyIsOwner` false)—optional low beep.

### Shift keys & IDE

- [ ] Fast Copy: mode enter/exit beyond existing mp3s (if you add more states).
- [ ] Any long-running VS Code / Cursor operation that already shows a banner but no sound.

### Window management

- [ ] Success vs failure when focusing Cursor / moving mouse ([`../WindowManagement.ahk`](../WindowManagement.ahk) near `into-cursor-textfield.wav`).

### App launchers

- [ ] Wikipedia / launcher FSM transitions (if you want audible phase changes; currently mostly silent).

---

## Blank rows (your ideas)

| Where (file:function or hotkey) | Event | Proposed filename | Notes |
|---------------------------------|-------|-------------------|-------|
| | | | |
| | | | |
| | | | |

---

## Implementation notes

- Prefer one shared helper (e.g. `SafePlayCue(fileName)`) with throttling if the same action can fire twice.
- For destructive actions, keep chimes subtle or reuse `cleaning-desktop`-style quiet playback.
- After adding files, list them here so you can avoid duplicate meanings.
