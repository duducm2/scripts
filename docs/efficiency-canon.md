# Efficiency Canon

**Purpose:** Foundational context for future AI executions. This document synthesizes prior Investigator AI deep technical investigations and establishes strategic guidelines, standardized best practices, and recommended technologies for automation development in this repository.

**Scope:** Read-only reference. Source corpus: all `*-evaluation-report.md` and `*-improvement*.md` documents in this `docs/` directory (Gemini, Teams, Outlook, Spotify, ShiftKeys, WindowManagement, AppLaunchers). No modifications are made to scripts by this document.

---

## 1. Investigator Declaration

Deep technical investigations have already been completed on the automation scripts and workflows covered by the evaluation and improvement documents in this repository. Those investigations are to be treated as **authoritative baseline context** for:

- Performance and reliability issues (bottlenecks, blocking, race conditions, targeting fragility).
- Architectural and remediation strategies (native AHK hardening, polyglot offload, COM/UIA usage).
- Technology viability and constraints (e.g. deprecated APIs, latency requirements, state synchronization).

Future AI runs must **ingest this canon and the referenced reports** before proposing changes. Do not re-audit from scratch unless the scope explicitly excludes prior findings. When in conflict, precedence is: determinism and safety first, then behavior parity, then latency and maintainability.

---

## 2. Strategic Execution Doctrine for Future AI Runs

- **Preserve behavior parity first.** Refactors must not change observable hotkey or workflow behavior unless the task explicitly requests it. Use feature flags and staged rollout so legacy paths remain callable until parity is verified.
- **Prefer deterministic state over timing assumptions.** Replace fixed `Sleep` chains and “hope the UI is ready” logic with condition-based waits (`WinWaitActive`, `StatusBarWait`, UIA state checks) and bounded timeouts. Do not rely on synthetic keystrokes (e.g. Alt+Tab, taskbar cycling) for critical state transitions.
- **Enforce timeout-bounded failure paths.** Every activation, wait, or IPC call must have a defined deadline. On timeout or error: return a strict sentinel (e.g. `0` or `false`), show non-blocking user feedback, and never leave global state (clipboard, delay settings, `SetTitleMatchMode`) altered.
- **Use typed contracts at boundaries.** Target and activation functions must return a validated integer (HWND) on success and exactly `0` or `false` on failure. Callers must gate state mutations on explicit checks (e.g. `hwnd is Integer && hwnd > 0`) and must not infer success from object truthiness or side effects.
- **Isolate global state mutations.** Any routine that changes `A_Clipboard`, `SetWinDelay`, `SetKeyDelay`, `SetControlDelay`, or `SetTitleMatchMode` must capture prior values and restore them in a `try`/`finally` block so restoration runs regardless of early return or exception.

---

## 3. Unified Bottleneck Taxonomy

Recurring issues identified across evaluation reports, normalized into a single taxonomy:

| Category                        | Description                                                                                                       | Typical impact                                            |
| ------------------------------- | ----------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------- |
| **Blocking sleeps**             | Fixed `Sleep` on the hotkey/flow thread with no condition check.                                                  | High latency, race when UI is slow or fast.               |
| **Polling loops**               | High-frequency timers (e.g. 100–200 ms) or tight loops calling `FindAll`/`FindFirst` or `WinGetList` repeatedly.  | CPU load, battery drain, input lag.                       |
| **Repeated enumeration**        | Full `WinGetList` or equivalent per hotkey press with no cache; O(n) or O(n²) visibility/sort per call.           | Latency scales with window count.                         |
| **Full UIA tree scans**         | Repeated `root.FindAll` or long `FindFirst` fallback ladders per operation; no cache request or shared discovery. | Cross-process COM cost, hundreds of ms per flow.          |
| **Expensive #HotIf predicates** | Predicates that run full window list or UIA on every keystroke.                                                   | Perceived typing latency.                                 |
| **Silent catch blocks**         | Empty or minimal `catch` that swallows errors without logging or state restoration.                               | Hidden failures, leaked state, BlockInput left on.        |
| **Global state leaks**          | `SetTitleMatchMode`, delay settings, or clipboard changed without guaranteed restore.                             | Unpredictable behavior in other hotkeys/scripts.          |
| **Hardcoded paths/literals**    | User-specific paths, version-pinned executables, localized title fragments, magic numbers.                        | Portability and maintenance failure.                      |
| **Non-deterministic fallbacks** | Alt+Tab, taskbar cycling (#t + arrows + Enter) to “find” a window.                                                | Wrong window activated, input delivered to incorrect app. |
| **Unbounded or blind waits**    | `WinWaitActive`/`WinWait` with no timeout; keys sent without verifying foreground window.                         | Indefinite block or key injection into wrong window.      |

---

## 4. Standardized Best Practices and Design Patterns

- **Strict return contracts.** Resolvers (e.g. “get meeting HWND”, “get Outlook mailbox HWND”) return a positive integer HWND on success and `0` (or `false`) on failure. No composite objects used for success/failure when callers rely on truthiness.
- **try/finally for global state.** For any path that mutates clipboard or delay/title-match settings: capture state at entry, run logic in `try`, restore in `finally`. Use `ClipWait` after clipboard restore where appropriate.
- **Cache-first with validation.** Use a singleton or module-level cache for expensive targets (e.g. meeting HWND, mailbox HWND). On use: validate with `WinExist("ahk_id " hwnd)` and, where applicable, a style or process check (e.g. `WinGetStyle(...) & WS_VISIBLE`). On miss or invalid: invalidate, perform one controlled re-resolve, then update cache.
- **Event-driven hooks over polling.** Prefer `SetWinEventHook` (e.g. `EVENT_SYSTEM_FOREGROUND`, `EVENT_OBJECT_CREATE`/`DESTROY`) to maintain focus or window-state cache instead of a 100–200 ms `SetTimer` loop. For UIA, prefer `StructureChangedEventHandler` or `PropertyChangedEventHandler` with a bounded fallback poll only when events are unavailable.
- **Bounded wait contracts.** All waits use an explicit timeout. Prefer `WinWaitActive(win, , timeoutSeconds)` and condition-based loops with a deadline over open-ended `WinWait` or long fixed sleeps.
- **Non-blocking callback dispatch.** When using timers or IPC callbacks, enqueue lightweight completion work; do not run heavy UIA traversal or long loops inside the callback. Use request IDs or correlation so callbacks map to the correct action.
- **Process-bound targeting.** Window targeting must constrain by process and, where stable, by class (e.g. `ahk_exe OUTLOOK.EXE`, `ahk_class rctrl_renwnd32`). Avoid title-only or substring-only matching for activation.
- **Single authority for duplicated logic.** Consolidate repeated hotkey bodies, fallback ladders, and validation into shared helpers (e.g. one “activate with fallback” function, one “get toggle state” with parameters for automation id and name patterns).

---

## 5. Architecture Orientations by Execution Tier

- **Native AHK-only tier.** Use for low-risk, deterministic refactors: HWND caching, try/finally state restoration, removal of synthetic fallbacks, replacement of fixed sleeps with bounded condition waits, consolidation of duplicate logic. No new processes or IPC. Keeps core execution strictly AutoHotkey v2 + Win32/UIA.
- **Polyglot offload tier.** Consider when heavy enumeration, O(n²) visibility/sort, long-running UIA monitors, or expensive #HotIf predicates cannot be sufficiently reduced in AHK. Offload to a **persistent** external daemon (Python or C#). AHK remains the hotkey/router and UI shell; the daemon owns state tracking, event hooks, and heavy traversal. **No** `RunWait`, per-keystroke process spawn, or temp-file IPC. Use Named Pipes or Shared Memory with defined message contracts and synchronization (e.g. Mutex, EventWaitHandle).
- **Cloud/API tier.** Optional and **not** part of core execution unless the roadmap explicitly requires it (e.g. Graph for presence). Cloud APIs are unsuitable for real-time, low-latency peripheral or UI control due to polling delay and rate limits; use only where the use case accepts multi-second latency.

---

## 6. Recommended Technologies and Integration Strategies

- **AutoHotkey v2 (AHK).** Primary language for hotkeys, window targeting, and local UI feedback. Use Win32 APIs (`WinExist`, `WinActivate`, `WinGetList`, `WinGetStyle`, etc.) with process/class constraints. Use UIA (e.g. UIA-v2) for controls that lack standard Win32 handles; prefer **UIA cache requests** (`UIA.CreateCacheRequest`, `ElementFromHandleBuildCache`, `FindFirstBuildCache`) for bulk property fetches to reduce cross-process calls.
- **UI Automation (UIA).** Use for discovery and invocation of buttons, list items, and toggle state when Win32 is insufficient. Prefer cached property/pattern retrieval over repeated live COM calls. For UI-bound commands not exposed via COM (e.g. “Read Aloud”), consider WM_COMMAND or UIA `InvokePattern` instead of synthetic keystrokes when IDs or elements are stable.
- **COM Interop.** Use for application data and model operations (e.g. Outlook `ComObjActive("Outlook.Application")`, mail/folder access) where available. Keep a native AHK fallback when COM is unavailable. Do not rely on COM for features that require sub-second UI reaction unless the API supports it.
- **Python or C# daemon.** Use for polyglot offload: persistent process, event hooks (`SetWinEventHook`), O(1) state cache, heavy UIA/traversal in background threads. Python: asyncio or threading; C#: async/Task and native Win32 interop. Daemon must expose a clear operation set (e.g. GetForegroundWindowState, GetCursorWindows, WatchUIState) and lifecycle (startup, heartbeat, graceful shutdown).
- **IPC strategy hierarchy.**
  - **Shared Memory (Memory-Mapped File):** Lowest latency (~sub-millisecond), suitable for state arrays and FSM flags. Requires strict synchronization (named Mutex, optional EventWaitHandle for readiness). Define a fixed layout (version, seq, opCode, payload, status) and byte offsets.
  - **Named Pipes:** Duplex, kernel-managed; latency typically tens of microseconds. Use for request/response and streaming events when Shared Memory complexity is undesirable. Use length-prefixed or framed messages (e.g. 4-byte length + UTF-8 payload).
  - **Forbidden:** CLI execution (`RunWait`, spawning script per action), temp-file IPC, unbuffered HTTP/localhost for hot paths.
- **WASAPI / audio.** For **per-application volume** control without touching the UI, use Windows Audio Session API (e.g. IAudioSessionManager2, IAudioSessionEnumerator, ISimpleAudioVolume) by PID. Do not use WASAPI to “mute” for UI features that must stay in sync with the app’s own mute state (e.g. Teams mic indicator); use for hardware or mixer-level control only where state sync is not required.
- **Integration strategy for new shortcuts/workflows.** Prefer: (1) implement in AHK with the patterns above; (2) if bottlenecks are severe and localized, consider a small, focused daemon with a single IPC channel and feature-flagged cutover; (3) document rollout order and rollback (legacy path remains callable).

---

## 7. Canonical Build/Refactor Workflow

1. **Audit.** Align with existing evaluation report for the script (if any); identify hotkey/flow and bottleneck taxonomy entries.
2. **Baseline.** Note current behavior and, where relevant, latency or enumeration cost (no need for heavy profiling unless the task requires it).
3. **Feature flags.** Introduce toggles (e.g. `USE_HWND_CACHE`, `USE_DAEMON`) default-off so new behavior can be enabled incrementally.
4. **Incremental cutover.** Replace one logical unit at a time (e.g. activation, then state checks, then UIA cache). Keep legacy implementation callable behind the flag.
5. **Failure-path tests.** Validate: app closed, app loading, stale cache, timeout, and ensure no leaked global state and no unhandled exception swallowing in critical branches.
6. **Rollback gates.** Ensure disabling the flag restores prior behavior so rollout can be reverted without code revert.

---

## 8. Verification Matrix Template

For each refactor or new automation path, confirm:

| Criterion             | Check                                                                                                                |
| --------------------- | -------------------------------------------------------------------------------------------------------------------- |
| **Functional parity** | Same hotkeys and user-visible behavior; no regressions in edge cases.                                                |
| **Latency**           | Activation and hotkey paths meet expectations (e.g. no multi-second block on hotkey thread).                         |
| **Reliability**       | No synthetic keystrokes to wrong window; no unbounded waits; focus/state restoration deterministic where applicable. |
| **Cleanup**           | Clipboard, delay settings, and title-match mode restored on success and on error; timers/hooks unregistered on exit. |
| **No leaked state**   | No permanent change to global AHK or OS state visible to other scripts or hotkeys.                                   |

---

## 9. Anti-Patterns (Do-Not-Use List)

- **Non-deterministic synthetic fallbacks:** Alt+Tab, taskbar cycling (#t + arrow keys + Enter) to “find” or activate a window.
- **Unbounded waits:** `WinWaitActive` or `WinWait` without a timeout parameter.
- **Per-action process spawn:** Running an external script or executable on every hotkey press (e.g. `RunWait` for each shortcut).
- **Temp-file IPC:** Using temporary files to pass data between AHK and another process for hot-path or repeated operations.
- **Silent exception swallowing:** Empty or minimal `catch` blocks in code that mutates global state or input (e.g. after `BlockInput("On")`) without logging or guaranteed cleanup.
- **Blind key injection:** Sending keystrokes (e.g. Alt+1, Escape) without verifying that the target window is foreground or that the control exists.
- **Hardcoded user or version-specific paths** in launch or resolution logic without fallback to environment variables or registry.

---

## 10. Canon Maintenance Rules

- **Append, do not duplicate.** New evaluation or improvement documents should reference this canon and add only **new** findings or **refinements** that do not contradict sections 2–9. If a refinement conflicts (e.g. a new technology recommendation), document the exception and the rationale in the report; do not silently override the canon.
- **Scoped overrides.** Script- or app-specific plans (e.g. “Teams: do not use Graph for core flow”) remain in their plans/reports; the canon states **general** principles. When a script’s constraints are stricter than the canon, the script-specific constraint wins for that script only.
- **Vocabulary.** Use the canonical terms (e.g. _deterministic_, _bounded_, _cache-first_, _event-driven_, _typed-contract_) in new reports and implementation plans so that future AI runs can match intent without ambiguity.

---

## 11. Hot-path example: YouTube focus (Win+Alt+Shift+H, AppLaunchers)

Illustrates **minimizing** UIA cost on a sub-second hotkey path:

- **Bind** `UIA_Browser` to the **specific** window (`"ahk_id " hwnd`), not a generic `ahk_exe` match, so URL and document queries refer to the correct Chrome instance.
- **Prefer** a **single** cheap call (`GetCurrentURL`) plus a **documented assumption** (e.g. user enters focus with the watch-page video **paused**) and **one** `Send("k")` over repeated `FindFirst` / subtree scans that can cost hundreds of milliseconds per key press.
- **Trade-off** must be explicit in code comments: if the assumption is wrong (e.g. video already playing), the same key toggles playback — acceptable only when the workflow guarantees or accepts that state.
- **Do not** add loading banners or NDJSON logging on the same hotkey path unless required for UX or diagnostics; they add latency and I/O.
