# Plan: Mobills Pagination Refactor

## Context
The current `Shift+K` (Previous Month/Page) and `Shift+L` (Next Month/Page) shortcuts in `Shift keys.ahk` are inconsistent across different Mobills modules (Transactions, Accounts, Planning) due to varying UI implementations.

## Objective
Unify the pagination logic into a single robust function `Mobills_Navigate(direction)` that detects the active context and applies the correct UIA selector strategy.

## Technical Analysis & Strategies

### 1. Context Detection
Determine the active module based on the URL or specific UI elements.
- **Transactions:** `web.mobills.com.br/transactions`
- **Accounts:** `web.mobills.com.br/accounts` (assumed based on context)
- **Planning:** `web.mobills.com.br/planning` (assumed based on context)

### 2. Element Selectors per Context

#### Scenario A: Transactions Page
- **Target:** Button (Type 50000)
- **Strategy:**
  1. Search by Name: "Go to previous page" / "Go to next page" (or localized variants).
  2. Fallback: Search for buttons adjacent to the Month/Year text header if pagination buttons are not found.

#### Scenario B: Accounts Page
- **Target:** Button (Type 50000)
- **Strategy:**
  1. Use UIA Path: `{T:30}, {T:26}, {T:0, i:7}` (Verify index for Prev vs Next).
  2. Fallback: Search for unnamed buttons in the top toolbar area.

#### Scenario C: Planning/Budgets Page
- **Target:** Text element acting as button? (Type 50020)
- **Strategy:**
  1. Use UIA Path: `{T:30}, {T:26}, {T:20, i:2}`.
  2. Fallback: Search by Name " " (Space) if unique in that container.

## Implementation Plan

1.  **Refactor `Shift keys.ahk`**:
    -   Remove existing `PrevMobillsMonth` and `NextMobillsMonth` logic.
    -   Implement `Mobills_Navigate(direction)` taking "Prev" or "Next".
    -   Implement `GetMobillsContext(uia)` to identify the current page.
    -   Implement specific handlers for each scenario inside `Mobills_Navigate`.

2.  **Unified Function Logic**:
    ```autohotkey
    Mobills_Navigate(dir) {
        uia := UIA_Browser()
        url := uia.GetCurrentURL()
        if InStr(url, "transactions")
            ; Strategy A
        else if InStr(url, "accounts")
            ; Strategy B
        else if InStr(url, "planning")
            ; Strategy C
        else
            ; Generic Fallback
    }
    ```

3.  **Verification**:
    -   Ensure `Shift+K` calls `Mobills_Navigate("Prev")`.
    -   Ensure `Shift+L` calls `Mobills_Navigate("Next")`.