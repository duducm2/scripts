# Understood requirements (for your revision)

**Purpose:** Product requirements as understood from AutoSlot / Windows rearrangement **as coded today**.

**Gates:** AutoSlot **ON** and **more than one monitor**.

**Framing:** Follows code in [`AutoSlot/AutoSlot.ahk`](../../AutoSlot/AutoSlot.ahk). Background import policy: [`05-y-only-fill.md`](05-y-only-fill.md).

---

## 1. Slot system (capacity)

| Rule                       | Understood meaning                                                                 |
| -------------------------- | ---------------------------------------------------------------------------------- |
| Grid                       | **2 slots** per ordinal monitor; up to **4** ordinals → **8** max                  |
| Lone maximized / work-area | Counts as **1 slot**; free half available for Y / Ctrl+6                           |
| Behind maximized           | Do not make the monitor “full” for free-half checks                                |
| Exclusions                 | ClipAngel, tool windows, dialogs/Teams chrome, own PID — not moved / not occupancy |

---

## 2. When a window is **created** (Place)

| Requirement   | Understood                            |
| ------------- | ------------------------------------- |
| Empty ordinal | Maximize onto first empty             |
| No empty      | Maximize in place (“grid full”)       |
| Must not      | 50/50 or demax existing slots on open |
| Feedback      | INFO toast                            |

---

## 3. When a window is **closed**

| Requirement          | Understood                           |
| -------------------- | ------------------------------------ |
| Snap partner         | **Heal** maximize companion (silent) |
| Empty / half monitor | **No** automatic background import   |
| Fill from background | Only via **Ctrl+Alt+Win+Y**          |

---

## 4. When a window is **minimized**

| Requirement       | Understood                          |
| ----------------- | ----------------------------------- |
| Snap pair         | Break pair; heal leftover companion |
| Background import | **None**                            |
| Restore           | Does not Place                      |

---

## 5. When a window is **maximized** or **moved**

| Requirement                 | Understood                                                     |
| --------------------------- | -------------------------------------------------------------- |
| One half of 50/50 maximized | **Unregister pair**; companion stays half — not auto-maximized |
| Move / suite leave          | No automatic background rearrange                              |
| Explicit Y                  | May SnapPair into free halves                                  |

---

## 6. Explicit user fill

| Action                      | Understood                                                            |
| --------------------------- | --------------------------------------------------------------------- |
| Ctrl+Alt+Win+Y / menu `[3]` | **Only** full background→slot fill (empty then halves, `forceImport`) |
| Ctrl+Alt+Win+6 open         | Place chosen window into empty / free half                            |

---

## 7. Foreground swap

| Requirement     | Understood                                     |
| --------------- | ---------------------------------------------- |
| When            | Suite move/snap may swap whole-monitor layouts |
| After           | INFO toast + brief quiet; **no** `[F]` replace |
| Fill after swap | User hits **Y** if desired                     |

---

## 8. Please verify (checklist)

- [okay] Place empty-only
- [okay] No auto background import on close/move/minimize
- [okay] Y is sole scan-fill from background
- [okay] Maximize half does not maximize companion
- [okay] Heal on close/minimize stays (silent)
- [okay] No swap `[F]` replace-skip
