# Outlook Appointment Palette - Letter Mapping

## Overview
This document defines the letter-to-configuration mapping for the Outlook Appointment popup palette.

**Total Combinations:** 24 (3 Status × 2 Private × 2 Reminder × 2 All-day)

**Controls:**
- **Status**: Free, Busy, Out of office (3 options)
- **Private**: On, Off (2 options)
- **Reminder**: 15 minutes, 2 days (2 options)
- **All-day**: Yes, No (2 options)

---

## Mapping Strategy

The mapping uses a hierarchical grouping:
1. **Primary Group:** Status (Free, Busy, Out of office) - 3 groups
2. **Secondary Group:** All-day (Yes, No) - 2 subgroups per status
3. **Tertiary Group:** Private (On, Off) - 2 subgroups
4. **Final Group:** Reminder (15min, 2days) - 2 options

**Letter Allocation Pattern:**
- Each Status group gets 8 letters (2 All-day × 2 Private × 2 Reminder = 8)
- Letters are assigned in a 2×4 grid pattern (2 rows for All-day, 4 columns for Private×Reminder combinations)
- Uppercase letters for All-day=Yes, lowercase for All-day=No (within each status)

---

## Complete Letter Mapping

### Group 1: Free Status (8 combinations)

#### Free + All-day=Yes (4 combinations)

| Letter | Status | All-day | Private | Reminder |
|--------|--------|---------|---------|----------|
| **Q** | Free | Yes | Off | 15min |
| **W** | Free | Yes | Off | 2days |
| **E** | Free | Yes | On | 15min |
| **R** | Free | Yes | On | 2days |

#### Free + All-day=No (4 combinations)

| Letter | Status | All-day | Private | Reminder |
|--------|--------|---------|---------|----------|
| **A** | Free | No | Off | 15min |
| **S** | Free | No | Off | 2days |
| **D** | Free | No | On | 15min |
| **F** | Free | No | On | 2days |

**Free Status Letters:** `Q W E R / A S D F`

---

### Group 2: Busy Status (8 combinations)

#### Busy + All-day=Yes (4 combinations)

| Letter | Status | All-day | Private | Reminder |
|--------|--------|---------|---------|----------|
| **Z** | Busy | Yes | Off | 15min |
| **X** | Busy | Yes | Off | 2days |
| **C** | Busy | Yes | On | 15min |
| **V** | Busy | Yes | On | 2days |

#### Busy + All-day=No (4 combinations)

| Letter | Status | All-day | Private | Reminder |
|--------|--------|---------|---------|----------|
| **B** | Busy | No | Off | 15min |
| **N** | Busy | No | Off | 2days |
| **M** | Busy | No | On | 15min |
| **,** | Busy | No | On | 2days |

**Busy Status Letters:** `Z X C V / B N M ,`

---

### Group 3: Out of office Status (8 combinations)

#### Out of office + All-day=Yes (4 combinations)

| Letter | Status | All-day | Private | Reminder |
|--------|--------|---------|---------|----------|
| **U** | Out of office | Yes | Off | 15min |
| **I** | Out of office | Yes | Off | 2days |
| **O** | Out of office | Yes | On | 15min |
| **P** | Out of office | Yes | On | 2days |

#### Out of office + All-day=No (4 combinations)

| Letter | Status | All-day | Private | Reminder |
|--------|--------|---------|---------|----------|
| **J** | Out of office | No | Off | 15min |
| **K** | Out of office | No | Off | 2days |
| **L** | Out of office | No | On | 15min |
| **;** | Out of office | No | On | 2days |

**Out of office Status Letters:** `U I O P / J K L ;`

---

## Visual Layout

### Grid Layout (24 squares total)

Each status group displays as a **2×4 grid**:
- **2 rows**: All-day Yes (top) / All-day No (bottom)
- **4 columns**: Private Off + Reminder 15min | Private Off + Reminder 2days | Private On + Reminder 15min | Private On + Reminder 2days

**Visual representation:**
```
FREE STATUS
┌─────┬─────┬─────┬─────┐
│  Q  │  W  │  E  │  R  │  All-day=Yes
├─────┼─────┼─────┼─────┤
│  A  │  S  │  D  │  F  │  All-day=No
└─────┴─────┴─────┴─────┘
  Off  Off  On   On
 15m  2d  15m  2d

BUSY STATUS
┌─────┬─────┬─────┬─────┐
│  Z  │  X  │  C  │  V  │  All-day=Yes
├─────┼─────┼─────┼─────┤
│  B  │  N  │  M  │  ,  │  All-day=No
└─────┴─────┴─────┴─────┘
  Off  Off  On   On
 15m  2d  15m  2d

OUT OF OFFICE STATUS
┌─────┬─────┬─────┬─────┐
│  U  │  I  │  O  │  P  │  All-day=Yes
├─────┼─────┼─────┼─────┤
│  J  │  K  │  L  │  ;  │  All-day=No
└─────┴─────┴─────┴─────┘
  Off  Off  On   On
 15m  2d  15m  2d
```

**Alternative Layout:** All three status groups can be displayed side-by-side or stacked vertically, depending on screen space and readability preferences.

---

## Implementation Notes

1. **Letter Case Sensitivity:** The mapping uses both uppercase and lowercase letters. The palette should display letters exactly as shown in the mapping.

2. **Special Characters:** One combination uses a special character (`,` and `;`). These are standard keyboard characters that should be handled reliably.

3. **Reminder Values:**
   - 15min = 15 minutes
   - 2days = 2 days

4. **Private Values:**
   - On = Private enabled
   - Off = Private disabled

5. **All-day Values:**
   - Yes = All-day event
   - No = Timed event

6. **Status Values:**
   - Free
   - Busy
   - Out of office

---

## Column Meaning (Consistent Across All Status Groups)

The 4 columns have the same meaning for every status:
- **Column 1**: Private Off, Reminder 15min
- **Column 2**: Private Off, Reminder 2days
- **Column 3**: Private On, Reminder 15min
- **Column 4**: Private On, Reminder 2days

This consistent pattern makes the palette easier to learn and use.

---

## Validation Checklist

- [ ] Visual layout displays all 24 combinations correctly
- [ ] Letter selection works for all characters (including `,` and `;`)
- [ ] Grid layout is clear and readable
- [ ] Status groups are visually distinct
- [ ] Column meanings are consistent across all status groups
- [ ] Palette disappears after selection
- [ ] No confirmation required (immediate execution)
