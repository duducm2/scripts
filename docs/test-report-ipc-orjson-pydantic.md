# Test Report: Priority 1 — orjson + pydantic (Python IPC protocols)

## 1. Selected task (highest priority)

From [python-library-opportunities.md](python-library-opportunities.md) section 4:

1. **pydantic + orjson** on protocol encode/decode (low risk, clear win).

**Status:** Python changes are **in the repo** (`ipc_wire.py`, protocol modules, `requirements.txt`). Run `pip install -r python/requirements.txt` if you have not since pulling, then use **Section 2** for hands-on AutoHotkey validation.

---

## 2. Manual tests — which shortcut, what to expect

These checks confirm that JSON over the pipes still works end-to-end after **orjson + pydantic**. Modifiers below use **Win+Alt+Shift** where written as **MEH** (same as your repo’s “MEH” convention).

### Prerequisites (every manual test)

1. From the repo folder, install deps: `pip install -r python/requirements.txt`.
2. Start only the daemon(s) needed for the row you are testing (each in its own terminal, leave running):
   - ShiftKeys: `python python/shiftkeys_daemon.py`
   - Gemini (optional rows): `python python/gemini_daemon.py`
   - Window management (optional): `python python/wm_daemon.py`
   - App launcher (optional): `python python/applauncher_daemon.py`
3. Reload the matching AutoHotkey script after code changes (or restart via [Act.ahk](../Act.ahk) if that is how you usually boot).

### Test matrix (what to press)

| # | You do this (shortcut / action) | Where it lives | Daemon | What you should see if the change is OK |
|---|----------------------------------|----------------|--------|-------------------------------------------|
| A | **Reload [Shift keys.ahk](../Shift%20keys.ahk)** (or let Act start it) while `shiftkeys_daemon.py` is running. No keypress after load. | [Shift keys.ahk](../Shift%20keys.ahk) calls `ShiftKeysIPC_Bootstrap()` on load; [aux/ShiftKeysIPC.ahk](../aux/ShiftKeysIPC.ahk) defaults `USE_DAEMON` / `USE_PIPE_IPC` to **true**, so the script talks to the daemon every 200 ms for `ResolveContext`. | `shiftkeys_daemon` | No AutoHotkey error dialog. Normal desktop use for ~10 s: no freeze, no repeated error toasts. If the daemon is **off**, behavior is unchanged from before (reconnect / empty responses per existing logic), not a JSON regression. |
| B | **Optional — Gemini completion chime via daemon:** In [aux/ShiftKeysIPC.ahk](../aux/ShiftKeysIPC.ahk), set `USE_DAEMON_MONITOR_GEMINI := true` (default is **false**), reload Shift keys, start `shiftkeys_daemon`. Open Gemini in Chrome, focus the prompt, press **Enter** (not Shift+Enter) or **Ctrl+Enter** to send. | Hotkeys `Enter::` / `^Enter::` under the Gemini section in [Shift keys.ahk](../Shift%20keys.ahk) (~lines 22689–22720). | `shiftkeys_daemon` | Same as today: message sends; when the model finishes, your **completion chime** still plays (`PlayCompletionChime_Gemini`). With the flag **false** (default), this row does **not** use the daemon — use row A only unless you explicitly enable the flag. |
| C | **Optional — ChatGPT hotkeys using daemon context cache:** Set `USE_DAEMON_CONTEXT_CHATGPT := true` in [aux/ShiftKeysIPC.ahk](../aux/ShiftKeysIPC.ahk), reload Shift keys, run `shiftkeys_daemon`. Open ChatGPT in Chrome so the window is active. Press **Shift+I** (sidebar toggle). | `#HotIf IsChatGPTActiveForHotkey()` block in [Shift keys.ahk](../Shift%20keys.ahk) (~12974+); `+i::` sends `^+s`. | `shiftkeys_daemon` | Sidebar toggles (same as without the flag when ChatGPT was already the active window). If context cache is wrong, hotkeys might not fire when ChatGPT looks focused — after the change, behavior should match pre-change with the same flag. |
| D | **MEH+O — Read aloud when Gemini is not in front:** Run `gemini_daemon.py`, reload [Gemini.ahk](../Gemini.ahk). Put **another** window in front; leave Gemini open in Chrome. Press **Win+Alt+Shift+O** (`#!+o`). | [Gemini.ahk](../Gemini.ahk) ~780; `GEMINI_USE_PYTHON_IPC` defaults **true**; queue path uses the pipe when a connection already exists. | `gemini_daemon` | Read-aloud flow still completes (switch to Gemini, Listen, etc.) as before. If the daemon is down, script falls back without requiring the daemon — still not a JSON encode failure. |
| E | **Only if you enable WM IPC:** In [aux/WMIPC.ahk](../aux/WMIPC.ahk), set `WM_USE_DAEMON` / `WM_USE_PIPE_IPC` / `WM_USE_EVENT_HOOK_CACHE` as you use for cutover, run `wm_daemon.py`, reload [WindowManagement.ahk](../WindowManagement.ahk), then exercise your usual **window move / monitor** shortcuts. | [WindowManagement.ahk](../WindowManagement.ahk) | `wm_daemon` | Same window behavior as before enabling IPC (no new errors). Defaults are **off** — skip this row unless you opt in. |
| F | **Only if you enable App Launcher IPC:** Turn on `AL_USE_DAEMON` / `AL_USE_MMF_IPC` in [AppLaunchers.ahk](../AppLaunchers.ahk), run `applauncher_daemon.py`, reload AppLaunchers, then use the flow that resolves cursor targets over IPC (your usual shortcut that hits that path). | [AppLaunchers.ahk](../AppLaunchers.ahk) + [aux/AppLauncherIPC.ahk](../aux/AppLauncherIPC.ahk) | `applauncher_daemon` | Same result as before (primary/fallback hwnd behavior). Defaults are **off** — skip unless you opt in. |

**Summary:** After the code change, **everyone** should at least run **row A** with Shift keys + `shiftkeys_daemon`. Rows **B–F** only apply if you use those optional flags or daemons.

---

## 4. Intended code changes (single scope)

- Add [python/ipc_wire.py](../python/ipc_wire.py): `json_dumps` / `json_loads_dict` using **orjson**; `validate_ipc_request_envelope` using **pydantic** `IpcRequestEnvelope` (`id` / `op` required, `extra` ignored, `id`/`op` coerced with `str()` so JSON numbers from AHK still validate).
- Extend [python/requirements.txt](../python/requirements.txt): `orjson>=3.9`, `pydantic>=2.5`.
- Replace stdlib `json` in:
  - [python/wm_protocol.py](../python/wm_protocol.py) — `encode_message`, `decode_message`, `validate_request`
  - [python/shiftkeys_protocol.py](../python/shiftkeys_protocol.py) — same
  - [python/al_protocol.py](../python/al_protocol.py) — `encode_payload`, `decode_payload`, `validate_request`
  - [python/protocol.py](../python/protocol.py) — `encode_message`, `decode_message`, `read_frame`, `validate_request`

Wire JSON shape and framing are unchanged; only the Python serializer/validator implementation changes.

### 4.1 New file: `python/ipc_wire.py`

```python
"""Shared UTF-8 JSON and minimal request validation for daemon IPC (orjson + pydantic)."""

from __future__ import annotations

from typing import Any

import orjson
from pydantic import BaseModel, ConfigDict, ValidationError, field_validator


def json_dumps(obj: dict[str, Any]) -> bytes:
    """Serialize a dict to compact UTF-8 JSON (non-ASCII as UTF-8, not \\u escapes)."""
    return orjson.dumps(obj)


def json_loads_dict(raw: bytes | bytearray | memoryview) -> dict[str, Any] | None:
    """Parse JSON bytes; return a dict or None if invalid or root is not an object."""
    try:
        val = orjson.loads(raw)
    except orjson.JSONDecodeError:
        return None
    return val if isinstance(val, dict) else None


class IpcRequestEnvelope(BaseModel):
    """Wire requests must include id and op (AHK may send numeric ids as JSON numbers)."""

    model_config = ConfigDict(extra="ignore")

    id: str
    op: str

    @field_validator("id", "op", mode="before")
    @classmethod
    def _coerce_id_op(cls, v: object) -> str:
        if v is None:
            raise ValueError("missing")
        return str(v)


def validate_ipc_request_envelope(obj: object) -> bool:
    if not isinstance(obj, dict):
        return False
    try:
        IpcRequestEnvelope.model_validate(obj)
        return True
    except ValidationError:
        return False
```

### 4.2 `python/requirements.txt` (append two lines)

```
orjson>=3.9
pydantic>=2.5
```

### 4.3 Protocol edits (pattern)

In each of `wm_protocol.py`, `shiftkeys_protocol.py`:

- Remove `import json`.
- Add: `from ipc_wire import json_dumps, json_loads_dict, validate_ipc_request_envelope`
- `encode_message`: `payload = json_dumps(obj)` (replace `json.dumps(...).encode(...)`).
- `decode_message`: after bounds checks, `return json_loads_dict(memoryview(data)[4 : 4 + length])` or `return json_loads_dict(data[4 : 4 + length])` (avoid double UTF-8 decode).
- `validate_request`: `return validate_ipc_request_envelope(obj)`.

In `al_protocol.py`:

- Same imports; `encode_payload` → `return json_dumps(obj)`; `decode_payload` → `return json_loads_dict(data)` inside try/except replaced by `return json_loads_dict(data)` (orjson raises on bad UTF-8 in invalid bytes — same as before for non-UTF8).

In `protocol.py` (Gemini):

- Same pattern for `encode_message` / `decode_message`.
- `read_frame`: replace `json.loads(payload.decode(...))` with `json_loads_dict(payload)` where `payload` is `bytes`.

---

## 5. Automated verification (after applying code)

From the repository root (with venv or user Python that has new deps installed):

```text
pip install -r python/requirements.txt
cd python
python -c "from wm_protocol import encode_message, decode_message, make_request, validate_request; r=make_request('1','Ping'); b=encode_message(r); assert validate_request(decode_message(b)); print('wm ok')"
python -c "from shiftkeys_protocol import encode_message, decode_message, make_request, validate_request; r=make_request('x','HealthCheck'); b=encode_message(r); assert validate_request(decode_message(b)); print('sk ok')"
python -c "from al_protocol import encode_payload, decode_payload, validate_request; p={'id':'1','op':'Ping','context':'','payload':{},'ts':0,'deadlineMs':0}; b=encode_payload(p); assert validate_request(decode_payload(b)); print('al ok')"
python -c "from protocol import encode_message, decode_message, make_request, validate_request; r=make_request('1','Ping'); b=encode_message(r); assert validate_request(decode_message(b)); print('gemini ok')"
```

---

## 6. Pause for your validation

Please either:

1. **Enable Agent mode** and ask the assistant to “apply the ipc_wire + protocol changes from docs/test-report-ipc-orjson-pydantic.md”, then run the pip install + `python -c` checks above; or  
2. **Paste the new file and edits manually**, then run the same checks.

Reply with **approve for commit** after daemons + Shift keys (daemon on) behave as before, or request a **revision** (e.g. stricter validation, stdlib fallback without new deps).

---

## 7. Next task (after approval)

Proceed to priority 2 from the opportunities doc: **implement `shiftkeys_uia.py`** with comtypes/pywinauto (or chosen stack)—only after this item is merged to your satisfaction.
