# Efficiency Canon

**Purpose:** Foundational context for future AI executions. This document synthesizes prior Investigator AI deep technical investigations and establishes strategic guidelines, standardized best practices, and recommended technologies for automation development in this repository.

**Scope:** Evaluation and improvement reports under `docs/` are the read-only source corpus. **Sections 11+** record **proven in-repo hot-path patterns** (may be updated when those implementations change). This document does not modify scripts by itself.

---

## 1. Investigator Declaration

Deep technical investigations have already been completed on the automation scripts and workflows covered by the evaluation and improvement documents in this repository. Those investigations are to be treated as **authoritative baseline context** for:

- Performance and reliability issues (bottlenecks, blocking, race conditions, targeting fragility).
- Architectural and remediation strategies (native AHK hardening, polyglot offload, COM/UIA usage).
- Technology viability and constraints (e.g. deprecated APIs, latency requirements, state synchronization).

Future AI runs must **ingest this canon and the referenced reports** before proposing changes. Do not re-audit from scratch unmless the scope explicitly excludes prior findings. When in conflict, precedence is: determinism and safety first, then behavior parity, then latency and maintainability.

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
- **Documents-folder sentinels:** For user-profile-local markers (no repo path), use zero-byte files in `A_MyDocuments` with literal comma-containing names; see [lightweight-api-sentinel-files.md](lightweight-api-sentinel-files.md) (`manage, study, set, top, link` for Study Subtopic Link).
- **WASAPI / audio.** For **per-application volume** control without touching the UI, use Windows Audio Session API (e.g. IAudioSessionManager2, IAudioSessionEnumerator, ISimpleAudioVolume) by PID. Do not use WASAPI to “mute” for UI features that must stay in sync with the app’s own mute state (e.g. Teams mic indicator); use for hardware or mixer-level control only where state sync is not required. For **quiet confirmation sounds in the same AHK process**, do not combine WMP OCX volume with WASAPI targets — see **§14**.
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

## 11. Hot-path example: YouTube focus (historical; `#!+h` now opens Prompts)

Previously illustrated by Win+Alt+Shift+H in AppLaunchers (YouTube focus). That hotkey now opens Utility Shortcuts → Prompts. The principles below still apply to other sub-second UIA hotkeys:

- **Bind** `UIA_Browser` to the **specific** window (`"ahk_id " hwnd`), not a generic `ahk_exe` match, so URL and document queries refer to the correct Chrome instance.
- **Prefer** a **single** cheap call (`GetCurrentURL`) plus a **documented assumption** and **one** `Send(...)` over repeated `FindFirst` / subtree scans that can cost hundreds of milliseconds per key press.
- **Trade-off** must be explicit in code comments when an assumption can be wrong.
- **Do not** add loading banners or NDJSON logging on the same hotkey path unless required for UX or diagnostics; they add latency and I/O.

YouTube focus-session helpers remain in [`Utils/peek_pdf_study_01.ahk`](../Utils/peek_pdf_study_01.ahk) if revived under another binding.

---

## 12. Findings: Gemini read-aloud + WindowManagement daemon (polyglot async, 2026)

Interesting outcomes from integrating Python for `Gemini.ahk` / `WindowManagement.ahk` without moving fragile browser UIA into Python. Useful for later refactors and audits.

- **AHK has no user threads for hotkeys.** “Async” means **return quickly from the hotkey** and continue work via `SetTimer(..., -delay)` (one-shot) or periodic timers. Splitting a former single function with chained `Sleep` into **phased callbacks** reduces how long any one timer callback holds the script; it does not parallelize UIA work across CPU cores.

- **Hybrid boundary is deliberate.** **Python:** persistent named-pipe daemon for cheap state (foreground snapshot, optional task queue, cursor-suppression window). **AHK:** hotkeys, `WinActivate`, UIA for Gemini Chrome (“Read aloud”, copy last response). Rebuilding Gemini DOM automation in Python (Playwright/Selenium/UIA) trades latency and maintenance for little gain unless the UI contract is frozen.

- **Small task queue vs. heavy offload.** A minimal `QueueTask` / `GetTaskStatus` pattern lets the hotkey enqueue intent and poll until `ready` without blocking the initial press; the daemon can later grow real work (prefetch, logging) behind the same contract. Avoid `RunWait` or per-keystroke Python spawn (already in anti-patterns §9).

- **Cursor jumps are a policy problem, not only a timer frequency problem.** `MonitorActiveWindow` + `SetCursorPos` on foreground change will fight any automation that activates another window. Mitigations that compose well: (1) **local suppression** (`TickCount`-bounded “do not recenter”), (2) **daemon-backed suppression** so all scripts that call `WMIPC_GetForegroundWindowState` see `suppressCursorCentering`, (3) **gate explicit** `MoveMouseToCenter` behind `MaybeCenterMouse`-style helpers on Cursor/Gemini bridge paths.

- **Restore-focus quality improves with hook cache.** Tracking **last non-Gemini foreground** in the WM hook cache helps pick a sane `OriginalHwnd` when the user’s true prior window is not `WinExist("A")` at hotkey time (e.g. focus already in Gemini).

- **Feature flags must stay orthogonal.** `GEMINI_USE_PYTHON_IPC` gates the Gemini sidecar; WM daemon behavior uses `WM_USE_DAEMON`, `WM_USE_PIPE_IPC`, `WM_USE_EVENT_HOOK_CACHE`. Gemini can call `WMIPC_*` only when WM flags allow connection; otherwise fall back to AHK-only behavior.

- **IPC framing must match on both ends.** If the AHK client speaks length-prefixed JSON over `\\.\pipe\...`, the Python daemon must use the same framing (not a different transport for the same script without updating the client).

- **Repo hygiene:** This tree may **track** some `infra/python/__pycache__/*.pyc` files. Running `python -m py_compile` in the workspace can dirty or create bytecode artifacts; prefer restoring tracked `.pyc` from git or avoiding compile-in-place when only validating syntax elsewhere.

- **Verification reminder for this stack:** With daemons off, behavior should match legacy AHK paths; with daemons on, confirm no pointer recenter during Gemini→restore cycles, read-aloud still reaches the latest response, and IPC timeouts degrade without wedging the script.

---

## 13. Shift keys: Gemini paste pacing (condition-based wait, 2026)

- **Avoid fixed multi-second tails** after Clip Angel or screenshot paste into Gemini when the UI can signal “uploading” via `FastCopyMode_GeminiIsUploadingImage`. Prefer **`Gemini_WaitForUploadIdleWithRefocus`** (bounded loop, refocus while uploading) over `Sleep 2600` so fast uploads return early.
- **Tune `minNoIndicatorMs` per flow:** Clip Angel mixes text and media — use a **high** minimum (e.g. 2600 ms) when no indicator appeared so behavior stays close to the old fixed delay; screenshot paste can use a **lower** minimum (e.g. 800 ms) because image uploads usually surface the heuristic quickly.
- **Reuse cached `UIA_Browser` in `finally`:** when Gemini stays foreground, **`FastCopyMode_FocusGeminiPromptField(uia)`** avoids a second `UIA_Browser` attach versus always calling `FocusGeminiAskFieldForHwnd` (see `Shift keys.ahk`, `Gemini_PasteFromClipAngelSequential`).
- **Prompt Manager `[Y]` auto-send:** [`PromptContext_WaitForSendReady`](../Utils/hotstring_selector_handlers_01.ahk) uses **scoped** upload text (`Gemini_GetSearchRoot` / companion root) plus **stable** Send-enabled + composer text for `PROMPT_PASTE_SEND_READY_STABLE_POLLS` (default 2). When `attachCount > 0` and no upload label appears, require `PROMPT_PASTE_SEND_MIN_NO_INDICATOR_MS` before counting stable polls. Early return on pass (do not burn the wait budget). Banner phases: uploads → Send. **Rollback:** `PROMPT_PASTE_USE_STABLE_SEND_READY := false` restores single-poll legacy wait.

---

## 14. Same-process confirmation chimes (WASAPI + SoundPlay, 2026)

- **Do not** use `WMPlayer.OCX` `settings.volume` for quiet confirmation sounds in the **same** AutoHotkey process that also sets per-session level via WASAPI (`SCRIPT_MASTER_VOLUME_PERCENT`) — the host’s per-app mixer can stay at the attenuation step (~10%) after playback.
- **Do:** `ApplyAutoHotkeyAudioSessionsVolumePercent(lowPercent)` → `ScriptSoundPlay(path, true)` (or `SoundPlay` with wait) → **`ApplyScriptMasterVolumeTarget()` inside `try`/`finally`** so restore runs on success, failure, or early exit. If the mixer still shows the low step, add **one** short **one-shot** `SetTimer(..., -250)` to call the same restore again — not a blind multi-second sleep, but a second enumeration pass after the audio graph updates.
- **Avoid** relying on a **fixed-delay** `SetTimer` alone to fix mixer level after embedded WMP — non-deterministic vs session lifetime and enumeration.
- **Reference:** `PlayCleaningDesktopSound` in `Utils.ahk`; globals `SCRIPT_MASTER_VOLUME_PERCENT` and `CLEANING_CONFIRM_WMP_VOLUME_PERCENT` (name retained; level is applied via WASAPI, not WMP).
- **Quick Update (`/Updated`):** play `quick-update-success.wav` with **wait** (`ScriptSoundPlay(..., true)`) **before** `ScheduleApplyScriptMasterVolumeTargetAfterQuickUpdate()`. If the chime runs **async**, a new audio session can appear **after** the first WASAPI pass and show ~10% in the mixer until the next enumeration.

---

## 15. Efficiency iteration — 2026 (revision-aligned, non-parallelism)

**Non-goal:** This wave did **not** target “type in one app while mouse/GUI automation runs elsewhere” (see [revision/revision.md](../revision/revision.md) for why that stays out of scope here). Focus: shorter hot paths, fewer redundant operations, optional event-driven cache invalidation.

### Wave 1 — Quick wins

| Item               | Change                                                                                                                                                                                                                                                                                                                                                                                                        | Rollback                                                                                                    |
| ------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------- |
| **WASAPI volume**  | [`SpotifyWASAPI.ahk`](../Lib/SpotifyWASAPI.ahk): implemented `WASAPI_SetSessionScalar`, `AdjustProcessVolumeByPid`, `SetProcessPlaybackVolumePercent`, `ApplyAutoHotkeyAudioSessionsVolumePercent` via `IAudioSessionManager2::GetSessionEnumerator` (session enum + `ISimpleAudioVolume`). Restores **Ctrl+Volume** silent path for Spotify when `AL_USE_WASAPI` is true in [`Spotify.ahk`](../Spotify.ahk). | Revert `SpotifyWASAPI.ahk` to stub returns if COM/session enumeration misbehaves on a specific audio stack. |
| **Bridge log I/O** | [`GeminiToCursorBridge.ahk`](../Lib/GeminiToCursorBridge.ahk): `global BRIDGE_AGENT_LOG_ENABLED := false` — `Bridge_Log` is a no-op unless set true (reduces `FileAppend` on bridge paths).                                                                                                                                                                                                                   | Set `BRIDGE_AGENT_LOG_ENABLED := true` in that file (or a small include) for diagnosis.                     |

**Patterns:** §3 “repeated enumeration” / anti hot-path I/O; §6 WASAPI per-process volume.

### Wave 2 — Enumeration / cache

| Item                   | Change                                                                                                                                                                                                                                                                                                        | Rollback                                                  |
| ---------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------- |
| **Outlook HWND cache** | [`Outlook.ahk`](../Outlook.ahk): optional `SetWinEventHook` for `EVENT_OBJECT_DESTROY` (0x8001) to call `InvalidateMailbox` / `InvalidateCalendar` / `InvalidateReminder` when the matching cached HWND is destroyed. **Default:** `OUTLOOK_USE_WINEVENT_INVALIDATE := false`. `OnExit` unregisters the hook. | Set `OUTLOOK_USE_WINEVENT_INVALIDATE := false` (default). |
| **WindowManagement**   | No code change: [`WindowManagement.ahk`](../WindowManagement.ahk) `MonitorActiveWindow` already early-returns when foreground `hwnd` equals `lastHwnd` (canon cache-first / avoid redundant work).                                                                                                            | N/A.                                                      |

**Patterns:** §4 event-driven hooks vs polling; §7 incremental cutover with flag.

### Wave 3 — UIA / overlay cost

| Item               | Change                                                                                                                                                                                        | Rollback                                        |
| ------------------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------- |
| **mousemaster**    | [`mousemaster.ahk`](../mousemaster.ahk): `Mousemaster_MaxHints` (default 350) stops scanning after enough interactive elements; **trade-off** documented in file (fewer hints on huge trees). | Raise `Mousemaster_MaxHints` or remove the cap. |
| **Gemini / Utils** | No structural change: further `CacheRequest` / scoped `FindAll` refactors left for a follow-up if profiling shows a hot path.                                                                 | N/A.                                            |

**Patterns:** §3 full UIA tree scans; §11 explicit trade-off in comments.

### Wave 4 — Shift keys split

**Deferred:** No `#include` extract from `Shift keys.ahk` in this iteration (optional in plan; avoid load-order risk until a dedicated refactor task).

**Pattern:** §7 one logical unit per cutover; defer monolith split until Waves 1–3 are stable in daily use.

### Verification reminder

Use §8 matrix after enabling `OUTLOOK_USE_WINEVENT_INVALIDATE` or toggling `BRIDGE_AGENT_LOG_ENABLED`: parity, latency, clipboard/state cleanup, hook teardown on exit.

### Clip Angel activate / maximize (2026)

- **Same-monitor fast path:** [`ClipAngel_ApplyLayoutOnMonitor`](../Utils/clip_angel_favorite.ahk) skips `MoveWindowToMonitor` (restore + Sleep 80 + WinMove) when the hwnd is already on the target monitor and not a tiny bar — maximize + short activate (150 ms) only; already-maximized → activate only.
- **Native Alt+P/B settle:** [`CLIPANGEL_NATIVE_OPEN_SETTLE_MS`](../Utils/clip_angel_activate.ahk) **50** (not 300); if the foreground/max gate still fails after one layout, **one** retry at **100** ms (`g_ClipAngelNativeOpenRetryArmed`). Hotkey still returns immediately.
- **Do not** change global `MoveWindowToMonitor` (Peek and others keep the restore settle). Cross-monitor or tiny-bar opens still use the full move+max path.
- **UIA root reuse:** [`ClipAngel_LeaveFavoritesFilter`](../Utils/clip_angel_favorite.ahk) resolves `ElementFromHandle` once and passes `root` into MarkFilter / Shift+P / Row0 helpers (avoids re-attach every 15 ms poll). [`ClipAngel_WaitForListReady`](../Utils/clip_angel_favorite.ahk) layouts once (re-layout only if `NeedsLayoutCorrection`) and reuses `root` in `IsListReady`.
- **Condition waits over fixed Sleep:** D2C `[O]` uses `ClipWait` after `^c` and [`ClipAngel_UiaWaitPreviewFocused`](../Utils/clip_angel_favorite.ahk) after `F10` (not 5× Sleep 100). Ctrl+1–5 copy path same pattern. [`ClipAngel_HideWindow`](../Utils/clip_angel_favorite.ahk) polls iconic/hidden (~100 ms @ 15 ms) instead of Sleep 50×2.
- **Rollback:** restore settle 300 / remove fast path in `ClipAngel_ApplyLayoutOnMonitor` if maximize races on a specific DPI/monitor setup; restore fixed Sleep chains / per-poll layout in LeaveFavoritesFilter / WaitForListReady / HideWindow if CA UI races appear.

---

## 16. Study Topic QuickLook cold-start (Win+Alt+Shift+X, 2026)

- **Do not** gate post-open layout on `WinWait(..., 2)` alone — cold `QuickLook.exe` launch often exceeds 2s; use **`QuickLook_WaitForHwnd`** (bounded poll, default 10s) and a user-visible timeout overlay instead of a silent skip.
- **Single authority:** **`QuickLook_ApplyStudyLayout`** (activate, focus click, optional scroll) shared by **`QuickLook_OpenPath`** and the **`#!+X`** fast path when QuickLook is already running. After layout, **`QuickLook_ScheduleAutoSlotPlace`** defers IPC (~400 / 1200 / 2500 / 4000 ms), re-resolving `ahk_exe QuickLook.exe` each time. QL’s `PositionWindow` undoes external resize unless WPF `WindowState` is Maximized — so AutoSlot always sends `SC_MAXIMIZE`, uses **PostMessage IPC** (file backup on Google Drive paths), and **sticky-retries** place until full/half geometry sticks (snap-pair map alone is not enough).
- **Readiness gates:** title basename match (`QuickLook_WaitForOpenReady`) plus UIA **`Document`** enabled on two consecutive polls (`QuickLook_WaitForViewerReady`); scroll via **`ScrollPattern`** first, **`^{End}`** fallback with vertical-% verification when available.
- **Rollback:** `STUDY_TOPIC_QL_STRICT_LAYOUT := false` in [`Utils.ahk`](../Utils.ahk) restores legacy 2s `WinWait` + inline scroll; optional one-shot **`SetTimer(..., -800)`** deferred layout if the process appears just after the hwnd wait times out.

---

## 17. Editor Smart Nav — Explorer reveal wait (Alt+H / Alt+I, 2026)

- **Do not** use a fixed multi-second sleep after Explorer activation when Alt+H/Alt+I need a selected file in ItemsView; use **`Editor_WaitForActiveExplorerWindow`** (single activate poll loop, not separate `WinWait` + `WinWaitActive`) → **`Editor_WaitForExplorerItemsView`** → **`Editor_WaitForExplorerRevealReady`** (poll `Explorer_GetItemsViewSelection`; **any** highlighted item — IDE reveal pre-selects; **`EDITOR_REVEAL_STABLE_POLLS`** consecutive stable polls, default 2) with bounded timeout (default 3500 ms).
- **Normalize only for hints:** **`Editor_NormalizeRevealBasename`** strips Cursor binary-tab placeholder text (comma suffix); wait does **not** gate on basename match.
- **Sidebar focus:** replace fixed sleeps after `^+e` with **`Editor_WaitForSidebarExplorerFocus`** (~800 ms poll on `FocusCursorFilesExplorer`).
- **Copy fast path (Alt+I):** when **`EDITOR_COPY_USE_EDITOR_FASTPATH`** (default true), try **`Editor_TryCopyFileFromActiveEditor`** before sidebar/reveal: native Copy Path (`^2`), `FileExist`, **`Editor_SetClipboardFiles`** + CF_HDROP verify; success uses **`StandardLoadingBar_Hide(200)`** and skips Explorer entirely. Fallback to reveal flow on untitled paths, missing files, or verify failure.
- **Copy reveal fallback:** **`Editor_GatherRevealContext`** (single Shell + UIA gather for folder, selection, resolved path) → when **`EDITOR_COPY_PREFER_DIRECT_SET`** (default true) and path exists, **`Editor_CopyVerifiedFileToClipboard`** first; one keyboard `^c` retry only if direct set fails. **`Explorer_EnsureItemsViewFocusPreserveSelection`** uses bounded focus poll (40 ms steps, 400 ms cap), not fixed `Sleep 120` ladders.
- **Open (Alt+H):** **`Editor_EnsureRevealItemSelected`** → **`Editor_ResolveRevealFullPath`** → `Run` full path; `Enter` + **`Editor_WaitForShellDispatchedAfterOpen`** fallback; then `WinClose`.
- **Copy (Alt+I) CF_HDROP contract:** verify clipboard file list (full path or basename); **`ClipboardAll()`** for failure restore and keep **`WinActivate`** editor in `finally`. Clipboard wait tunables: **`EDITOR_COPY_DIRECT_CLIP_WAIT_MS`** (300), **`EDITOR_COPY_CLIP_WAIT_MS`** (500 keyboard retry).
- **Throttle:** `EDITOR_SMARTNAV_MIN_INTERVAL_MS` (~450 ms) in **`Editor_SmartNavReveal`** to ignore rapid double-presses.
- **Dev timing:** **`Editor_SmartNav_TimingLog`** gated by **`EDITOR_SMARTNAV_TIMING := false`** (default off); no NDJSON on hotkey path in production.
- **Rollback:** `EDITOR_COPY_USE_EDITOR_FASTPATH := false` restores always-reveal copy; `EDITOR_COPY_PREFER_DIRECT_SET := false` restores keyboard-first on reveal fallback; `EDITOR_USE_CONDITIONAL_EXPLORER_WAIT := false` restores legacy `Sleep 2500` after Explorer activation; `EDITOR_COPY_VERIFY_FILEDROP := false` restores legacy `^c` + `ClipWait` copy behavior only; tune **`EDITOR_REVEAL_STABLE_POLLS`** / clipboard wait globals without code revert.
- **Failure:** `Editor_SmartNavRevealShowExplorerTimeout` (hides loading bar first); no NDJSON on hotkey path.
- **Loading UI:** `StandardLoadingBar_Show` / `Update` / `Hide` across `Editor_SmartNavReveal` and Explorer wait/copy/open (`BANNER_ACCENT_INTERMEDIATE`; `centerOnHwnd` = editor); see [standard_information_display.md](standard_information_display.md).

---

## 18. Snap half-pair hot path (Ctrl+Alt+Win+X, 2026)

- **Gapless-only placement:** `^!#x` → `WM_SnapHalfPairActiveWindow` → `WM_SnapPairGaplessRects` → `WM_MoveHwndToRectGapless` per pane. No Win+Z UI automation, no Win+Left/Right synthetic tier (removed — OS snap ignores margin/gutter and forced a redundant validation pass).
- **Sizing:** `DwmGetWindowAttribute(DWMWA_EXTENDED_FRAME_BOUNDS)` converted via `PhysicalToLogicalPointForPerMonitorDPI`; split `SetWindowPos` (move then size) for cross-DPI stability; up to 6 measure-and-nudge iterations per window — acceptable on hotkey press for pixel-accurate fit; do not reduce without verification.
- **Geometry authority:** `WM_ComputeSnapPairPaneRects` applies `WM_SNAP_PAIR_MARGIN` (6 px) and `WM_SNAP_PAIR_GAP` (4 px) before halving; portrait/landscape axis from work-area aspect ratio.
- **Empty monitor:** `WM_ResolveSnapTargetMonitor` — when cursor is on a visually empty monitor (same rule as `CycleWindowsOnMonitor`), snap both windows there instead of the still-focused window's monitor.
- **Partner search:** `WM_EnumerateOpenHwndsGlobal` (global MRU z-order via `WinGetList`); validation uses `partnerHwnd` fast path when known — skip `GetVisibleWindowsOnMonitor` scan when partner already classifies in the opposite pane.
- **Bounded validation:** single post-placement `WM_WaitValidateSnapBipartitionStrict` (400 ms deadline, 25 ms poll); failure → `ShowNotification_WM`.
- **Abandoned companion heal:** use `WM_MaximizeHwndBackground` with `SetWindowPos` flags `SWP_NOACTIVATE | SWP_NOZORDER` (`0x0014`) — never `SWP_SHOWWINDOW` / `SC_MAXIMIZE` on heal paths (those activate and steal focus).
- **AutoSlot free-capacity (current):** no background import on close/move/minimize — heal leftover companion only. Explicit **Ctrl+Alt+Win+6**: same-monitor BG pairing first, then ordinal free-half fill; leftover lone halves expand to full. See [`docs/canon/windows-rearrange.md`](canon/windows-rearrange.md).
- **Foreground after mutations:** single `WM_EnsureForegroundHwnd` (`WinActivate` + bounded `WinWaitActive`) then one `WM_MaybeCenterMouse`; no focus-lock timers, no deferred reclaim, no NDJSON on the hotkey path.
- **Prepare order:** in `WM_SnapHalfPairActiveWindow`, `WM_PrepareHwndForTile(partner)` then `WM_PrepareHwndForTile(target)` so the target remains the foreground candidate before gapless placement.
- **Do not:** reintroduce synthetic snap keystrokes, NDJSON/logging on the hotkey path, or unbounded validation waits.
- **Rollback:** revert [`WindowManagement/tile_snap.ahk`](../WindowManagement/tile_snap.ahk) if snap regressions appear on any monitor/DPI combo.

---

## 19. Command Palette Edit Favorite (Shift+E, 2026)

- **Do not** walk menu items via `UIA.GetFocusedElement` after Ctrl+K — focus stays on `ContextFilterBox` (Edit/50004); `elFound` stays empty for the whole Down loop.
- **Do not** poll with repeated full-tree `FindAll` (Button + MenuItem + ListItem) per Down step — that is §3 “full UIA tree scans” on a hotkey path.
- **Do:** `Send "^k"` → bounded deadline (~1500 ms) with short Sleep (~40 ms); each tick **`FindFirst`** Name substring `"Editar favorito"` / `"Edit bookmark"` (then `"favorit"`/`"bookmark"` + edit-stem check); **`InvokePattern`/`Click`** when found. Send `{Down}` only while FindFirst still misses (lazy populate), capped (~12).
- **Rollback:** restore prior focus-walk / multi-`FindAll` loop in [`command_palette_helpers.ahk`](../Shift keys/command_palette_helpers.ahk) only if FindFirst misses localized labels on a specific PowerToys build.

---

## 20. Editor Git stash/fetch/pull (Alt+S, 2026)

- **Do not** hide git in a background PowerShell script or drive the command palette. The user must see a robot doing the work.
- **Do** open a **new editor terminal** (same chord as Shift+N / `Ctrl+Shift+'`) and **SendText** a PowerShell sequence: if there are stashable changes (`git diff` / untracked, ignoring dirty submodule _content_) then `git stash push -u` → wait until `.git/index.lock` is gone and the tree has no stashable changes → `git fetch` → `git pull --ff-only` → `git stash pop` only if that Alt+S stash is on top. Dirty submodule worktrees are not stashable; skip stash and pull anyway. If stash prints `No local changes to save`, skip stash rather than fail. `git -C` uses **`Editor_ResolveGitRepoDir`** when the workspace root is known. Play **`pull-successful.wav`** only after the terminal sequence reports success (including stash pop).
- **Do** print `=== ROBOT stash/fetch/pull/done ===` in the terminal so progress is visible. Watch the terminal for errors; do not celebrate a hidden exit code. Pull must not start on stash exit 0 alone.
- **Do** after a successful Alt+S on the project-selector **[s] Scripts** window (`CursorTransfer_GetMatchingProjectIndex` / `projects.ini` Char=s), call `QuickUpdateScripts()` so AHK relaunches from the pulled files. The `/Updated` overlay is the success signal. Do not send `Ctrl+Alt+Win+2` into the editor terminal. Other repos keep the Pull complete banner only.
- **Rollback:** restore [`infra/tools/Editor-GitStashFetchPull.ps1`](../infra/tools/Editor-GitStashFetchPull.ps1) CLI path in [`cursor_predicates.ahk`](../Shift%20keys/cursor_predicates.ahk) only if terminal typing is unreliable on a specific machine.
