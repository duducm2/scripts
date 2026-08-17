# Finance system — feature research notes

Research notes and guesses only. Not a build commitment. Written for a local stack: AutoHotkey CRUD + CSV data + Python Plotly cockpit (no bank sync).

**Date:** 2026-08-16

## Purpose

Compare what commercial personal-finance products and dashboard guides emphasize with what this project already does, then list ideas that could land in:

- **AHK UI** — fast daily actions, lists, confirm dialogs, banners
- **Web cockpit** — glanceable KPIs, charts, period filters

## Sources (web)

- [Personal Finance Dashboard (finhelp.io)](https://finhelp.io/glossary/how-to-create-a-personal-finance-dashboard/)
- [Household Financial KPI Dashboard Guide (finhelp.io)](https://finhelp.io/glossary/household-financial-kpi-dashboard-what-to-track-monthly/)
- [PocketSmith](https://www.pocketsmith.com/) — calendar + long-range balance forecast
- App roundups (Forbes Advisor / NerdWallet / PCMag, 2025–2026): Monarch Money, Quicken Simplifi, YNAB, PocketGuard — recurring bills, safe-to-spend, net worth, subscriptions, spending plans
- [Zerosum analytics](https://zerosum.so/analytics) — savings rate, runway, sankey-style cash flow
- [ProjXplorer](https://projxplorer.com/) — freedom runway / emergency-fund coverage in time units
- [FiHaven](https://fihaven.app/) — bills calendar, debt snowball/avalanche

## What you already cover

**Cockpit:** balance / income / expense / card-available KPIs; expense & income pies; spent-by-category and monthly balance charts; annual cash flow; goals bars; budgets (per category + highlighted total bar); credit cards; date-period filter; budget/card alerts.

**AHK:** accounts, cards, budgets, goals, categories, transactions; AI daily import; rebuild balances from transactions; dashboard widget toggles in Settings.

**Gaps vs industry "must-haves" (guess):** savings rate as a first-class KPI; emergency-fund months / runway; true net worth (assets − liabilities); budget *pace* vs calendar; recurring bills / subscriptions; merchant/payee analytics; forward cash-flow forecast.

---

## 1. Strong fits (mostly from current CSVs)

Compute from `accounts.csv`, `credit_cards.csv`, `transactions.csv`, `budgets.csv`, `goals.csv` with little or no schema change.

| Idea | Typical formula / behavior | Why apps push it | Cockpit vs AHK | Effort guess |
|------|----------------------------|------------------|----------------|--------------|
| **Savings rate %** | `(income − expense) / income` for the selected period (or saved-to-goals / income) | Core KPI next to cash flow on almost every dashboard guide | **Cockpit** KPI tile; optional AHK month banner | Low |
| **Net worth (simple)** | `sum(account current_balance) − sum(card current_spent)` | Single "am I getting richer?" number | **Cockpit** KPI + optional sparkline if you snapshot monthly | Low |
| **Budget pace** | Expected spend by day-of-month = `planned × (day / days_in_month)`; flag ahead/behind | Simplifi/Monarch "on track" pacing | **Cockpit** on total bar + category meta | Low–medium |
| **Top merchants / payees** | Aggregate by normalized `description` (or first N tokens) | Pattern spotting without new categories | **Cockpit** table or bar under period | Low–medium |
| **MoM / YoY category compare** | Same category spend vs prior period / prior year | Seasonality (already have period filter) | **Cockpit** compare strip or dual bars | Medium |
| **Liquid vs total balance** | Tag or heuristic: exclude long-term / FGTS-like accounts from "spendable" | Safe decisions need liquid cash, not paper wealth | Cockpit KPI + AHK header total | Low if you hardcode account ids; medium with `account_kind` |

---

## 2. Needs light new data

Small CSV columns or a new table unlock features competitors charge for.

| Idea | New data | Why useful | Cockpit vs AHK | Effort guess |
|------|----------|------------|----------------|--------------|
| **Emergency fund months / runway** | Which accounts count as liquid; optional "essential" category flag | `liquid / avg monthly essential expenses` → months of cover (finhelp, ProjXplorer) | **Cockpit** widget (months + traffic light); AHK alert if under 3 months | Medium |
| **Recurring bills / subscriptions** | `recurring.csv` or `is_recurring` + cadence on txs | PocketGuard/Simplifi differentiator; upcoming cash hits | **AHK** list + mark paid; **Cockpit** "next 14 days" | Medium |
| **Safe-to-spend** | Recurrings + budget leftovers + goal contributions | Income − bills − allocated budgets − goals = discretionary | Strong **AHK** "today" number; cockpit KPI | Medium (after recurrings) |
| **Account kind** | `liquid` / `investment` / `locked` on accounts | Enables net worth breakdown and runway without hacks | Settings in **AHK**; charts in cockpit | Low–medium |
| **Debt fields** | Interest rate, min payment on cards or a `debts.csv` | Avalanche vs snowball projections (FiHaven) | **AHK** payoff planner; cockpit debt-free date | Medium–high |
| **Cash-flow calendar / forecast** | Recurring schedule + known future txs | PocketSmith's main hook: projected balance days/months ahead | **Cockpit** calendar or projected balance line | High |
| **Weekly AI recap** | Prompt + existing CSVs | Monarch-style insights without new UI chrome | Extend **dictation/Gemini** path (you already import daily) | Low–medium |

---

## 3. Heavy / maybe skip (poor fit for local AHK + CSV)

| Idea | Why skip or defer |
|------|-------------------|
| Bank sync (Plaid etc.) | Complexity, security, cost; fights offline/local design |
| Live investment prices / brokerage APIs | Needs market data and continuous refresh |
| Multi-user household sharing | Auth, sync, conflict — overkill for single-machine utils |
| Full double-entry ledger / PFS generation | You already have pragmatic account + tx balances |
| Estate / tax filing suites | Out of scope vs Mobills-style replacement |
| Peer spending comparison | Needs external benchmarks and privacy tradeoffs |

---

## 4. Suggested first experiments (highest ROI for *this* stack)

1. **Savings rate + simple net worth KPIs** on the cockpit (pure compute from existing data).
2. **Budget pace** on the total (and maybe category) bars — day-of-month expected vs actual.
3. **Top payees** for the selected period — cheap insight from descriptions.
4. **`account_kind` + emergency-fund months** once liquid accounts are tagged.
5. **Recurring bills CSV** → AHK list + cockpit "upcoming" → unlocks safe-to-spend later.

Defer: full PocketSmith-style 30-year forecast, FIRE calculators, debt optimizers until recurrings and debt fields exist.

---

## 5. Placement cheat-sheet

| Prefer cockpit | Prefer AHK |
|----------------|------------|
| Savings rate, net worth, runway months | Recurring bill CRUD / mark paid |
| Budget pace visuals | Safe-to-spend banner after morning import |
| Top payees, MoM compare, forecast chart | Debt payoff what-if tool |
| Period-driven analytics | Quick adjust, rebuild, AI import |

---

## 6. Open questions (for you later)

- Which accounts are truly liquid for runway?
- Should savings rate use cash-flow surplus or only transfers into goals?
- Card `current_spent` as the only liability, or will you add loans?
- Recurrings: separate CSV vs flags on imported daily txs?

When you pick one row from section 4, turn it into a normal implementation plan.
