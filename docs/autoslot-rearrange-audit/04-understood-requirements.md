# Understood requirements (for your revision)

**Purpose:** Product requirements as understood from AutoSlot / Windows rearrangement **as coded today**. Use this to mark what is wrong, incomplete, or unintended.

**Gates:** AutoSlot **ON** and **more than one monitor**. When OFF or single-monitor, place / fill / rearrange are no-ops.

**Framing:** Follows code in [`AutoSlot/AutoSlot.ahk`](../../AutoSlot/AutoSlot.ahk), not stale Place tables in [`AutoSlot/README.md`](../../AutoSlot/README.md) (README still describes open-time 50/50; code does not).

**Related:** behavior map [`01-how-it-acts.md`](01-how-it-acts.md); risks [`03-main-risks.md`](03-main-risks.md).

---

## 1. Slot system (capacity)

| Rule                            | Understood meaning                                                                                                                                                                     |
| ------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Grid                            | **2 slots** per **ordinal** monitor; up to **4** ordinals → **8** slots max                                                                                                            |
| Lone maximized / work-area fill | Counts as **1 slot**, but is "physically" occupying two slots; the other half stays **free** for a 50/50 partner                                                                       |
| Windows behind a maximized one  | Do **not** make the monitor “full” for free-half / heal decisions                                                                                                                      |
| True full                       | Two half-panes, or two+ maximized / filled — no import                                                                                                                                 |
| Exclusions                      | Clip Angel, tool windows, desktop/taskbar, empty titles, excluded exe/title (dialogs, Teams chrome, etc.), own-script PID — **do not occupy slots** and **are not moved** by rearrange |

Ordinal order is left-to-right (suite monitor order), not necessarily Windows monitor index.

---

## 2. When a window is **created** (Place)

| Requirement          | Understood                                                                                                    |
| -------------------- | ------------------------------------------------------------------------------------------------------------- |
| Empty ordinal exists | Maximize the new window onto the **first empty** ordinal                                                      |
| No empty ordinal     | Maximize **in place** (“grid full”) — do **not** move it to another monitor                                   |
| Must not             | 50/50 or demaximize / shrink windows already in slots **on open**                                             |
| After place          | **Place freeze** + **claim** that monitor briefly: no background **import**; companion **heal** still allowed |
| Feedback             | INFO toast (`Auto-slotted → M…` or `Grid full — maximized`)                                                   |

**Verify note:** Older README still says origin/half 50/50 on open. Code = empty-only + maximize in place.

---

## 3. When a window is **closed**

| Requirement                              | Understood                                                                         |
| ---------------------------------------- | ---------------------------------------------------------------------------------- |
| Snap partner exists                      | **Heal**: maximize the leftover companion                                          |
| Monitor empty                            | Fill from background: two → 50/50; one → maximize                                  |
| Monitor half-full (one half or lone max) | SnapPair a background into the free half, or heal residual                         |
| Already-visible on other monitors        | **Not** reshuffled onto the freed monitor — only **background promote** + **heal** |
| Undo                                     | **No** undo modal on fill-on-close                                                 |

---

## 4. When a window is **minimized**

| Requirement                          | Understood                                                                                                 |
| ------------------------------------ | ---------------------------------------------------------------------------------------------------------- |
| Snap pair                            | **Break** the pair registry for that HWND                                                                  |
| Leftover companion                   | Heal (maximize) when applicable                                                                            |
| Underfilled monitors                 | **Rearrange** (same fill/heal policy as close), unless swap-quiet                                          |
| Restore from taskbar (`MINIMIZEEND`) | Does **not** run Place; arms JustRestored; clears replace-skip so the window can be a fill candidate later |

---

## 5. When a window is **maximized** or **moved**

| Requirement                                               | Understood                                                                         |
| --------------------------------------------------------- | ---------------------------------------------------------------------------------- |
| One half of a registered 50/50 maximized                  | Maximize the **companion** too (no orphaned half)                                  |
| Drag end / suite leave / AutoSlot maximize onto a monitor | Debounced **rearrange underfilled** ordinals (exclude the mover where applicable)  |
| Rearrange may                                             | Promote **background** into empty/half slots; heal lone companions                 |
| Rearrange must not                                        | Move already-visible slotted windows **between** monitors (no full grid reshuffle) |
| Self-triggered moves from fill/heal                       | Suppressed via pair suppress / claim so fill does not immediately re-fire          |

---

## 6. Explicit user fill

| Action                                              | Understood                                                                                                                                                                                  |
| --------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Ctrl+Alt+Win+Y / Window tools **[3]** (AutoSlot ON) | Fill free ordinal capacity from background: **empty monitors first**, then **halves** (including lone max). Bypasses place freeze / claim (`forceImport`). Does not break full 50/50 pairs. |
| Ctrl+Alt+Win+6 / list open (AutoSlot ON)            | Chosen window: empty → maximize; free half → 50/50; else restore in place                                                                                                                   |

---

## 7. Foreground swap (suite move when AutoSlot ON)

| Requirement     | Understood                                                                                               |
| --------------- | -------------------------------------------------------------------------------------------------------- |
| When            | Suite move-to-monitor / related snap path with AutoSlot ON may **swap** whole-monitor foreground layouts |
| During modal    | **Swap quiet** — do not fill/rearrange mid-exchange                                                      |
| **[F]** replace | Minimize displaced dest window(s); **block** promoting them back until restore / TTL                     |
| After quiet     | Post-quiet rearrange of underfilled ordinals                                                             |

---

## 8. Please verify (checklist)

Mark each as **intended** or **wrong**:

- [okay] **Place empty-only** — new windows never 50/50 into existing slots; only empty maximize or maximize in place
- [okay] **Rearrange does not reshuffle** already-visible windows across monitors — only background + heal
- [okay] **Lone maximized = free half** for fill-on-close, rearrange, and Y
- [okay] **Heal often silent**; **fill always toasted** (INFO)
- [okay] **Restore from minimize does not Place**
- [okay] **Y / Ctrl+6 may free-half SniapPair**; Place must not
- [okay] Cap **2 × ordinal (max 8)** and exclusion of ClipAngel / dialogs / Teams chrome from occupancy

---

## Out of scope in this doc

Implementation risks, Windows API details, and fix roadmaps — see other files in this folder.
