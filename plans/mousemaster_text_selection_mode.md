# Mousemaster — Text Selection Mode: Implementation Plan

## Context

**File:** `mousemaster.ahk` (AutoHotkey v2 + UIA-v2 library at `UIA-v2\Lib\UIA.ahk`)

This script provides keyboard-driven UI element clicking via letter hints overlaid on screen.
The existing hotkey `Ctrl+Alt+Win+C` activates/deactivates the hint overlay.

**Goal:** Add a **double-tap** variant of the same hotkey that, instead of clicking, enters a
**Text Selection Mode** — a purely keyboard-driven workflow to select a range of words in
any on-screen text element.

---

## Existing Globals (do not rename/remove)

```ahk
global MousemasterActive := false
global MousemasterOverlayGui := ""
global MousemasterElements := []
global UserInputBuffer := ""
global MM_InputHook := ""
global ActiveWinID := ""
global Mousemaster_MaxHints := 350
```

The function `Mousemaster_GenerateHint(index)` generates alphabetic hint labels (A, B, ..., Z, AA, AB...).
The overlay GUI uses a **transparent green chroma-key** (`00FF00`) background with white `Text` controls.

---

## New Globals to Add

```ahk
global MM_TapCount := 0          ; how many taps fired in the current window
global MM_TapTimer := false      ; whether the 330ms window timer is active
global MM_TapInterval := 330     ; ms — max interval between taps to count as double
global MM_SelectMode := false    ; true when Text Selection Mode is active
global MM_SelectedLineRange := ""  ; IUIAutomationTextRange of the chosen line
global MM_WordRanges := []       ; array of {hint, range, rect} for each word
```

---

## Step 1 — Replace the Hotkey with a Double-Tap Dispatcher

**Replace** the current `^!#c::` block entirely:

```ahk
^!#c:: {
    global MM_TapCount, MM_TapTimer, ActiveWinID, MousemasterActive

    ; If Mousemaster is already showing overlay, any tap deactivates it
    if (MousemasterActive) {
        Mousemaster_Deactivate()
        MM_TapCount := 0
        MM_TapTimer := false
        return
    }

    try {
        ActiveWinID := WinGetID("A")
        if (!ActiveWinID) {
            ToolTip("❌ Nenhuma janela ativa!", 200, 200)
            SetTimer(() => ToolTip(), -2000)
            return
        }
    } catch as e {
        ToolTip("❌ Erro: " e.Message, 200, 200)
        SetTimer(() => ToolTip(), -2000)
        return
    }

    MM_TapCount++

    if (MM_TapCount = 1) {
        ; Start the 330ms window
        MM_TapTimer := true
        SetTimer(MM_TapFire, -MM_TapInterval)
    } else if (MM_TapCount >= 2) {
        ; Second tap arrived in time — cancel the pending single-tap and go to text mode
        MM_TapTimer := false
        MM_TapCount := 0
        SetTimer(MM_TapFire, 0)   ; cancel pending timer
        Mousemaster_ActivateTextMode(ActiveWinID)
    }
}

MM_TapFire() {
    global MM_TapCount, MM_TapTimer, ActiveWinID
    if (!MM_TapTimer)
        return
    MM_TapTimer := false
    MM_TapCount := 0
    Mousemaster_Activate(ActiveWinID)   ; single tap — original click behavior
}
```

---

## Step 2 — Text Mode: Phase A — Line Selection

Add a new function `Mousemaster_ActivateTextMode(WinID)`.

### What it does
1. Scans the UIA tree for elements that **support TextPattern** (i.e., contain readable text lines).
2. Builds and shows the hint overlay — same visual style as the existing one.
3. Starts an `InputHook` that calls `MM_TextMode_LineSelected(element)` when the user types a matching hint.

### Recommended element filter
Do **not** use `IsEnabled: true` — text/label elements are never enabled.
Instead, use `IsOffscreen: false` only and then test for TextPattern support per element:

```ahk
allElements := RootElement.FindElements({ IsOffscreen: false })
for uiaEl in allElements {
    try {
        tp := uiaEl.GetPattern(UIA.Pattern.Text)   ; returns "" or object
        if (!tp)
            continue
        loc := uiaEl.Location
        if (loc.w < 10 || loc.h < 10)
            continue
        ; also accept ControlType 50020 (Text), 50021 (Document), 50004 (Edit)
        ; even without TextPattern confirmation, log and skip gracefully
        MousemasterElements.Push({ hint: Mousemaster_GenerateHint(...), uiaElement: uiaEl,
            textPattern: tp, x: loc.x, y: loc.y, w: loc.w, h: loc.h,
            cx: loc.x + loc.w // 2, cy: loc.y + loc.h // 2 })
    } catch {
        continue
    }
}
```

> **Note on `UIA.Pattern.Text` constant:** In the Descolada UIA-v2 library, the TextPattern
> pattern ID is `10014`. Use `uiaEl.GetPattern(10014)` if `UIA.Pattern.Text` is not defined.
> Verify in `UIA-v2\Lib\UIA.ahk`.

### InputHook wiring
Same structure as `MM_InputHook` in `Mousemaster_Activate`. On exact match, call
`MM_TextMode_LineSelected(matchedElement)` instead of `Mousemaster_PerformAction`.

---

## Step 3 — Text Mode: Phase B — Word Selection (Start Word)

`MM_TextMode_LineSelected(element)` performs:

1. Destroy the line overlay.
2. Call `element.textPattern.GetVisibleRanges()` → array of `IUIAutomationTextRange`.
   - Each range is a visual line. If only one range, use it directly.
   - If multiple, display line sub-hints (A, B, C...) and wait for a second hint selection,
     then proceed with the chosen range. *(Simplification: use the first range if count = 1.)*
3. For the selected line range, enumerate words:

```ahk
MM_WordRanges := []
wordRange := lineRange.Clone()
wordRange.MoveEndpointByUnit("Start", "Word", 0)   ; anchor start

loop {
    ; advance end by one word unit
    moved := wordRange.MoveEndpointByUnit("End", "Word", 1)
    if (moved = 0)
        break
    text := Trim(wordRange.GetText(200))
    if (text = "")
        continue
    rects := wordRange.GetBoundingRectangles()   ; returns flat array [x,y,w,h, x,y,w,h...]
    if (rects.Length >= 4) {
        wx := rects[1], wy := rects[2], ww := rects[3], wh := rects[4]
        MM_WordRanges.Push({ hint: Mousemaster_GenerateHint(MM_WordRanges.Length+1),
            range: wordRange.Clone(), text: text, x: wx, y: wy, w: ww, h: wh })
    }
    ; collapse start to current end to advance
    wordRange.MoveEndpointByRange("Start", wordRange, "End")
}
```

4. Build a new overlay with word hints positioned using the bounding rect coordinates
   (these are **screen coordinates**, subtract `ActiveWinX/Y` for GUI-relative placement).
5. Start a new `InputHook`. On match, store `MM_StartWordRange := matchedWord.range` and call
   `MM_TextMode_SelectEndWord()`.

---

## Step 4 — Text Mode: Phase C — Word Selection (End Word)

`MM_TextMode_SelectEndWord()`:

1. Re-show the same word hint overlay (or a new one showing only words from start position onward).
2. On user selection of the end word, call `MM_TextMode_Commit(startRange, endRange)`.

---

## Step 5 — Commit the Selection

`MM_TextMode_Commit(startRange, endRange)`:

```ahk
MM_TextMode_Commit(startRange, endRange) {
    Mousemaster_Deactivate()
    Sleep(80)

    ; Build a combined range from start-word start → end-word end
    combined := startRange.Clone()
    combined.MoveEndpointByRange("End", endRange, "End")

    ; Select it using the accessibility API — no keyboard simulation needed
    combined.Select()

    ToolTip("✅ Text selected via UIA TextPattern.", 200, 200)
    SetTimer(() => ToolTip(), -2000)
}
```

---

## Fallback Strategy (if TextPattern is unavailable on a target element)

If `GetPattern(10014)` returns nothing for most elements in the target app (e.g., Electron, Qt):

**Fallback B — Keyboard simulation:**
1. Hint-select a focusable element (ControlType 50004 Edit or 50021 Document).
2. Focus it: `WinActivate` + `uiaEl.SetFocus()`.
3. `Send "{Ctrl down}a{Ctrl up}"` → `Send "{Ctrl down}c{Ctrl up}"` → read `A_Clipboard`.
4. Split clipboard text by whitespace → build word list.
5. Display word hints as evenly-distributed overlays within element bounding box (estimated, not pixel-accurate).
6. For start word at index N: `Send "{Home}"` then `Send "{Ctrl down}"` + N × `{Right}` + `"{Ctrl up}"`.
7. For end word at index M (from start): `Send "{Shift down}{Ctrl down}"` + (M-N) × `{Right}` + `"{Ctrl up}{Shift up}"`.

> This approach works universally but word hint positions are estimations, not exact.

---

## Simplified Scope Alternative (Recommended First Milestone)

Before building full word-level selection, implement **line-only selection**:

- Double-tap → hint to a TextPattern element → select a line → `lineRange.Select()` → done.
- No word enumeration at all.
- This validates the TextPattern detection pipeline before adding word complexity.

---

## Functions Summary Table

| Function | Purpose |
|---|---|
| `MM_TapFire()` | Fires single-tap action after 330ms if no second tap arrived |
| `Mousemaster_ActivateTextMode(WinID)` | Phase A: scan text elements, show line hints |
| `MM_TextMode_LineSelected(element)` | Phase B: enumerate words, show word hints |
| `MM_TextMode_SelectEndWord()` | Phase C: re-show word hints for end selection |
| `MM_TextMode_Commit(startRange, endRange)` | Build combined range and call `.Select()` |

---

## Key UIA-v2 API Calls Reference

| Operation | AHK Call |
|---|---|
| Get TextPattern | `uiaEl.GetPattern(10014)` |
| Get visible line ranges | `tp.GetVisibleRanges()` |
| Clone a range | `range.Clone()` |
| Move endpoint by word | `range.MoveEndpointByUnit("Start"/"End", "Word", n)` |
| Get word bounding boxes | `range.GetBoundingRectangles()` → flat array [x,y,w,h,...] |
| Set endpoint relative to other range | `range.MoveEndpointByRange("Start"/"End", otherRange, "Start"/"End")` |
| Get text of range | `range.GetText(maxLength)` |
| **Select the range** | `range.Select()` |

> Verify all method names against `UIA-v2\Lib\UIA.ahk` — the Descolada library may use
> slightly different casing or wrapper names.

---

## Implementation Order

1. Add new globals.
2. Replace hotkey block with double-tap dispatcher (`MM_TapFire`).
3. Implement `Mousemaster_ActivateTextMode` with TextPattern element scan + overlay + InputHook.
4. Implement `MM_TextMode_LineSelected` (line-only version first — just call `range.Select()`).
5. Test the line-only milestone against Chrome/Edge and a Win32 text editor.
6. Extend with word enumeration (`GetBoundingRectangles` + word hint overlay) once line mode is confirmed.
7. Implement `MM_TextMode_Commit` with combined range.
8. Add fallback keyboard path for apps without TextPattern.
