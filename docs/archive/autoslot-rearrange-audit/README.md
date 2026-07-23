# AutoSlot rearrange audit — ARCHIVED

**Do not treat this folder as current policy.**

Current Windows Rearrange behavior is documented in:

**[`docs/canon/windows-rearrange.md`](../../canon/windows-rearrange.md)**

This workbench is historical (findings, risks, pre/post Y-only notes). Some artifacts may predate free-half Place or other later changes. Prefer the canon + [`AutoSlot/AutoSlot.ahk`](../../../AutoSlot/AutoSlot.ahk) when implementing or explaining behavior.

---

# Original audit README (historical)

Workbench for documenting and later fixing Windows rearrangement problems in AutoSlot / WindowManagement.

**Method:** static code analysis only. No runtime instrumentation, debug banners, or UI output added to identify bugs.

**Primary surface:** [`AutoSlot/AutoSlot.ahk`](../../../AutoSlot/AutoSlot.ahk) plus WM call sites in WindowManagement.

## Steps

| Step | Status  | Artifact                                                         | Intent                                                                         |
| ---- | ------- | ---------------------------------------------------------------- | ------------------------------------------------------------------------------ |
| 0    | Done    | [`00-findings-report.md`](00-findings-report.md)                 | Full problem inventory (correctness, races, coupling, verbosity, inefficiency) |
| 0b   | Done    | [`01-how-it-acts.md`](01-how-it-acts.md)                         | As-implemented behavior map (triggers, guards, outcomes, banners)              |
| 0c   | Done    | [`02-windows-apis-influence.md`](02-windows-apis-influence.md)   | Windows APIs/events that influence rearrange (risks + opportunities)           |
| 0d   | Done    | [`03-main-risks.md`](03-main-risks.md)                           | Pinpoint main risks of rearrange                                               |
| 0e   | Done    | [`04-understood-requirements.md`](04-understood-requirements.md) | Understood requirements for your revision                                      |
| 0f   | Done    | [`05-y-only-fill.md`](05-y-only-fill.md)                         | Y-only background fill policy (UX 1/3/5/8)                                     |
| 1–4  | Planned | TBD placeholders                                                 | Not authoritative; see canon for current policy                                |

Note: auto fill-on-close / rearrange-import was removed in favor of Y-only (`05-y-only-fill.md`); older findings in `00` may describe pre-change behavior. Place free-half on open, all-monitor busy overlays, and **Y lone-half → maximize** (not BG 50/50) came later — see canon.

## Related

- [`docs/canon/windows-rearrange.md`](../../canon/windows-rearrange.md) — **current** policy
- [`docs/standard_information_display.md`](../../standard_information_display.md) — banners
- [`AutoSlot/README.md`](../../../AutoSlot/README.md) — enablement overview
