# Finance app

Local personal-finance stack: AutoHotkey CRUD + CSV data under `finances/data/` + Plotly cockpit HTML in `finances/output/`.

Open via **Utility Shortcuts → [F] Finance** (`#!+u`, then F).

## Requirements (each PC)

1. **AutoHotkey v2** scripts already running (`Utils.ahk` includes the finance modules).
2. **Python 3** available as one of: `py -3`, `py`, `python3`, `python`, or a normal install under `%LOCALAPPDATA%\Programs\Python\…`.
   - The Windows Store “python” stub (exit **9009**) is not enough — install a real Python or the [py launcher](https://docs.python.org/3/using/windows.html#the-py-launcher-for-windows).
3. Plotly once per machine:

```powershell
py -3 -m pip install -r finances\python\requirements.txt
```

## Dashboard open (`[D]`)

`Finance_OpenDashboard` always:

1. Runs `finances/python/chart_generator.py` with `--data-dir` and `--output-dir` pointing at **this** clone (so work and personal paths stay correct). It does **not** rebuild or rewrite account/card balances.
2. Opens Chrome only if Python exits **0** and `dashboard.html` exists (avoids showing a stale file).
3. Copies HTML to a temp file with a cache-bust query so Chrome does not reuse an old `file://` page.

Manual balance/spent edits are preserved: accounts use `initial_balance`, cards use `initial_spent`, so an explicit Settings → Rebuild still keeps those totals.

## Important settings (`finances/data/settings.ini`)

| Key                        | Meaning                                                                                                            |
| -------------------------- | ------------------------------------------------------------------------------------------------------------------ |
| `General.DefaultAccountId` | Main checking / liquid account. Used for the “Main account after card” liquid bar. Set via Accounts → set primary. |
| `General.PrimaryCardId`    | Card whose spent is subtracted from that account. Set via Credit cards → set primary.                              |

IDs are preserved on rename. Do not hand-edit IDs unless you know the matching rows in `accounts.csv` / `credit_cards.csv`.

## Data layout

| Path                             | Role                                                                                    |
| -------------------------------- | --------------------------------------------------------------------------------------- |
| `finances/data/*.csv`            | Source of truth (accounts, cards, transactions, budgets, goals, categories, recurring). |
| `finances/data/settings.ini`     | Defaults + dashboard widget toggles.                                                    |
| `finances/data/imported/`        | Archived daily import CSVs.                                                             |
| `finances/output/dashboard.html` | Generated cockpit (safe to regenerate anytime).                                         |
| `finances/python/`               | Chart build (`chart_generator.py`, `data_aggregator.py`).                               |

Account balance edits set `initial_balance` so a later rebuild does not wipe the manual total. Card spent edits set `initial_spent` the same way.

## Multi-PC notes

- Data lives in the repo; after `git pull` on another machine, open Finance → Dashboard once so Python regenerates HTML for that PC.
- If Dashboard says Python failed / not found, fix the Python install on that machine (see Requirements). Do not rely on copying `dashboard.html` alone.
- Feature ideas / research: [docs/research-feature-ideas.md](docs/research-feature-ideas.md).
