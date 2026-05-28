# Shotcrete Ledger — Hide-State Sync Design

**Date:** 2026-05-28
**Status:** Approved by Yavuz (2026-05-28). Awaiting live UI verification before commit.

## Problem

`PreviewWindow` (main Excel pass) computes a hide list of cubes that the user
shouldn't act on — split into two reasons:

- `no_match` — the cube didn't match any sheet in the main Excel workbook
- `done` — both 7-day and 28-day checkboxes default to unchecked (no work)

These cubes are collapsed behind a `▸ N hidden (...)` toggle in the main
preview. However, when the user later opens
`ShotcreteLedgerPreviewWindow` (Books Excel → SHOTCRETE), those same cubes
re-appear as actionable cards, because the ledger window's hide partition
only knows its own local reasons (`fully_complete` and `not_found`).

User wants the main preview's hide decision to carry over: cards hidden in
main should also be hidden in ledger, with the same toggle, and they must
default to **unticked** so no accidental write happens. The user can opt-in
by manually re-ticking the per-card checkbox.

## Why the persistence approach works here

The app's flow guarantees `PreviewWindow` is always opened before the user
can reach `ShotcreteLedgerPreviewWindow`:

1. User picks PDF → `_last_cubes_data` populated → `PreviewWindow` auto-opens
   (`main.py` ~L4065).
2. User edits / writes / closes → `_persist_edits_to_cubes` runs.
3. User clicks Books → SHOTCRETE → `ShotcreteLedgerPreviewWindow` opens with
   the same `cubes_data` dict.

Therefore writing a per-cube flag in step 2 and reading it in step 3 is
sufficient. No need to recompute main's match logic inside the ledger
window.

## Changes (main.py)

### 1. `PreviewWindow._persist_edits_to_cubes` (~L1737)

After the existing `cube["_card_enabled"] = ...` line, compute the
current hide reason from live state (matched_sheet + check_7d/28d
BooleanVars) and persist it:

```python
if not entry.get("matched_sheet"):
    cube["_hidden_in_main"] = "no_match"
else:
    c7 = entry.get("check_7d")
    c28 = entry.get("check_28d")
    c7_on = c7.get() if c7 is not None else False
    c28_on = c28.get() if c28 is not None else False
    cube["_hidden_in_main"] = "done" if (not c7_on and not c28_on) else None
```

This runs in both the auto-save close path and the write path, so the
flag is fresh whenever the user goes to ledger next.

### 2. `ShotcreteLedgerPreviewWindow._load_ledger` (~L2877)

In the entry-build loop, store the flag on the entry:

```python
entry["hidden_in_main"] = bool(cube.get("_hidden_in_main"))
```

### 3. `ShotcreteLedgerPreviewWindow._build_body` (~L3055)

Replace the existing partition with a 3-way split:

```python
done_entries = [
    e for e in self.entries
    if not e.get("mismatch") and e.get("fully_complete")
]
main_hidden = [
    e for e in self.entries
    if e.get("hidden_in_main") and e not in done_entries
]
hidden = done_entries + main_hidden
actionable = [e for e in self.entries if e not in hidden]
hidden_total = len(hidden) + len(self.not_found)
```

Toggle label simplifies to: `▸ N hidden` (no breakdown). When the toggle
is open, `main_hidden` entries render as full interactive cards under the
same "Already filled in ledger" section, so the user can override.

### 4. `ShotcreteLedgerPreviewWindow._build_card` (~L3330)

Extend the existing auto-untick branch so hidden-in-main cards also
start unchecked:

```python
if not bad and (
    (not c7.get() and not c28.get())
    or entry.get("hidden_in_main")
):
    entry["enabled"].set(False)
```

The per-group checkbox already controls writing via `entry["enabled"]`
(`_do_write` skips disabled entries ~L3451), so unticking is enough — no
extra write-side change needed.

## Behavior

- Card hidden in main preview → also hidden in ledger.
- Toggle (`▸ N hidden`) shows total only.
- Opening the toggle reveals all hidden cards as interactive but unticked.
- User can manually tick to opt-in to writing — full override path.
- No regression for the "ledger already fully complete" untick (existing
  behavior preserved by OR-extending the same conditional).

## Rollback

Single file (`main.py`). Revert the four hunks. No new files. No data
migration — `_hidden_in_main` is just a transient dict key on cube
objects.

## Live verification

Per user preference, the app must be launched and the ledger window
shown live before any commit. The user will visually confirm:

1. Cubes hidden in main preview are absent from the actionable list in
   ledger.
2. `▸ N hidden` shows the correct total.
3. Opening the toggle reveals the cards unticked.
4. Manually ticking a card lets the write proceed.
