# Efficiency iteration log

Running changelog for [efficiency-canon.md §15](efficiency-canon.md) work (revision-aligned improvements).

| Date    | Wave | Summary                                                                                                                                                                                                                                          |
| ------- | ---- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| 2026-05 | 1    | WASAPI implementations in `SpotifyWASAPI.ahk`; `BRIDGE_AGENT_LOG_ENABLED` in `GeminiToCursorBridge.ahk`.                                                                                                                                         |
| 2026-05 | 2    | Outlook optional `OUTLOOK_USE_WINEVENT_INVALIDATE` (default false); WM noted unchanged (already optimized early exit).                                                                                                                           |
| 2026-05 | 3    | `mousemaster.ahk` `Mousemaster_MaxHints` cap.                                                                                                                                                                                                    |
| 2026-05 | 4    | Deferred — no `Shift keys.ahk` extract this round.                                                                                                                                                                                               |
| 2026-06 | 1    | Chrome Shift+W detach (`Utils.ahk`): `CHROME_DETACH_DEBUG_LOG_ENABLED` (default false); single-pass `Chrome_ContextMenuInspectPopupHwnd`; cached `menuPopupClassify` / `activeTab`; tuned settle/poll timings; removed post-AppsKey fixed sleep. |
