# AutoSlot rearrange audit

Workbench for documenting and later fixing Windows rearrangement problems in AutoSlot / WindowManagement.

**Method:** static code analysis only. No runtime instrumentation, debug banners, or UI output added to identify bugs.

**Primary surface:** [`AutoSlot/AutoSlot.ahk`](../../AutoSlot/AutoSlot.ahk) plus WM call sites in WindowManagement.

## Steps

| Step | Status  | Artifact                                                         | Intent                                                                         |
| ---- | ------- | ---------------------------------------------------------------- | ------------------------------------------------------------------------------ |
| 0    | Done    | [`00-findings-report.md`](00-findings-report.md)                 | Full problem inventory (correctness, races, coupling, verbosity, inefficiency) |
| 0b   | Done    | [`01-how-it-acts.md`](01-how-it-acts.md)                         | As-implemented behavior map (triggers, guards, outcomes, banners)              |
| 0c   | Done    | [`02-windows-apis-influence.md`](02-windows-apis-influence.md)   | Windows APIs/events that influence rearrange (risks + opportunities)           |
| 0d   | Done    | [`03-main-risks.md`](03-main-risks.md)                           | Pinpoint main risks of rearrange                                               |
| 0e   | Done    | [`04-understood-requirements.md`](04-understood-requirements.md) | Understood requirements for your revision                                      |
| 0f   | Done    | [`05-y-only-fill.md`](05-y-only-fill.md)                         | Y-only background fill policy (UX 1/3/5/8)                                     |
| 1    | Planned | `01-timer-consolidation` (TBD)                                   | Collapse overlapping fill/heal/rearrange timer pipelines                       |
| 2    | Planned | `02-toast-policy` (TBD)                                          | Reduce rearrange banner noise without losing mode identity                     |
| 3    | Planned | `03-occupancy-perf` (TBD)                                        | Cut repeated occupancy / background enumeration cost                           |
| 4    | Planned | `04-policy-alignment` (TBD)                                      | Align Place vs Y vs rearrange/fill-on-close semantics and docs                 |

Later fix steps are placeholders only; do not implement until a dedicated plan is approved. Note: auto fill-on-close / rearrange-import was removed in favor of Y-only (`05-y-only-fill.md`); older findings in `00` may describe pre-change behavior.

## Related

- [`docs/standard_information_display.md`](../standard_information_display.md) — AutoSlot rearrange accent (`BANNER_ACCENT_INFO`)
- [`AutoSlot/README.md`](../../AutoSlot/README.md) — feature overview
- [`05-y-only-fill.md`](05-y-only-fill.md) — current background-import policy
