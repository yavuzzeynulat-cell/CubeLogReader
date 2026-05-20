# Shotcrete ledger (Books Excel)

**Date:** 2026-05-20

## Problem

The Books Excel ("ledger") flow currently writes only **concrete** cubes
to a single workbook whose sheet is named `Concrete`. Shotcrete cubes
are filtered out of that pass and have no ledger destination of their own.

A separate Excel file exists for shotcrete cores
(sheet tab: `Shotcrete Concrete Results`). Its layout is similar to the
concrete ledger but with two material differences:

- **5 specimens per age per set** (not 3) — block = 5 × 7-day + 5 × 28-day.
- **Four value columns per row** — Core Diameter, Core Height, Weight, Load —
  not just Weight + Load.

The shotcrete ledger must also support the same in-window file selector
the concrete ledger gained in v1.0.10 (multiple ledger files may be open
at once; the user picks).

## Decisions (approved by user)

- **Separate sibling class** `ShotcreteLedgerPreviewWindow` — copy of
  `LedgerPreviewWindow`, modified for shotcrete. The concrete class is
  not touched. Trade-off: future fixes may need to be applied twice;
  accepted for risk isolation.
- **Two grouped buttons on MainWindow** under a "Books Excel" label:
  `[ CONCRETE ]` (existing blue, opens concrete ledger) and
  `[ SHOTCRETE ]` (orange `#EF6C00`, opens shotcrete ledger).
- **All 5 specimens written to the ledger** — the per-sheet pass picks
  the top 3 by strength, but the ledger receives every specimen.
- **Header-based detection** with the shotcrete A7 text "Core No." — this
  naturally filters concrete files out of the shotcrete candidate list
  (concrete A7 says "Cube No.").
- **Card layout** for the shotcrete preview shows four value columns
  per row (Diameter | Height | Weight | Load) instead of two.

## Confirmed Excel layout

Source: real workbook screenshot, file "Shotcrete sample form",
tab "Shotcrete Concrete Results".

| Column | Field |
|--------|-------|
| A | Core No. |
| B | Site Sampling Mark/No. (e.g. `G26-CON-027`) |
| C | Shotcrete Supplier |
| D | Site Location / Comments |
| E | Section |
| F | Batch Ticket Number |
| G | C Grade |
| H | Date of Sampling |
| I | Sampled by |
| J | Date of Testing |
| **K** | **Age (days)** |
| **L** | **Core Diameter (mm)** |
| **M** | **Core Height (mm)** |
| **N** | **Weight (gr)** |
| **O** | **Load (Kn)** |
| P | Compressive Strength (N/mm²) — computed in Excel, not written |
| Q | Tested By (signature) |
| R | Engineer (signature) |

Header row: **row 7**. A7 contains "Core No.", B7 contains
"Site Sampling Mark/No:" (the concrete ledger's A7 is "Cube No." —
this single-character difference is the detection discriminator).

Block size: **10 rows per single set** (5 × 7-day + 5 × 28-day).
Two-set shotcrete is assumed to be a single 20-row block (parallel to
how concrete handles two-set 6+6); confirm against a real two-set
example before implementation closes.

## Design

### writer.py — additions (concrete code untouched)

Five new functions, sibling to the existing concrete ones:

- `find_shotcrete_ledger_candidates()` — same detection pattern as
  `find_ledger_candidates()` but matches A7 containing **"core no"** and
  B7 containing "sampling mark". Returns
  `list[(wb, ws, wb_name, sheet_name)]`. Raises `RuntimeError` on 0
  candidates with a shotcrete-specific message.
- `read_shotcrete_ledger_blocks(ws)` — like `read_ledger_blocks` but
  reads columns **A:K** (Age in column K). Returns block dicts with
  `rows_7d` / `rows_28d` populated by walking column K for `7` / `28`.
- `read_shotcrete_ledger_values(ws, blocks)` — like
  `read_ledger_values` but reads **L:O** for each block. Returns
  `{block_index: {"diameters": [...], "heights": [...], "weights":
  [...], "loads": [...]}}`.
- `merge_shotcrete_cubes_for_ledger(cubes_data)` — sibling of
  `merge_cubes_for_ledger` with **two inversions** and one carry-over:
  1. KEEP only cubes where `_shotcrete is True` (concrete dropped them).
  2. DO NOT skip tests where `_selected is False` — all 5 specimens go
     into `tests_7d` / `tests_28d`.
  3. Carry over: still skip cubes where `_card_enabled is False`
     (user unticked the master "Write this sample" in PreviewWindow).
- `write_shotcrete_ledger_cube(ws, cube, block, write_7d, write_28d)` —
  parallel to `write_ledger_cube`. For each specimen in the matched
  block rows, writes Diameter (L), Height (M), Weight (N), Load (O).
  Skips a cell when the ledger cell is already non-empty (no overwrite,
  identical policy to the concrete writer). Honours the per-group
  toggles `write_7d` / `write_28d`.

`match_cubes_to_blocks(merged, blocks)` is reused as-is — it already
checks generically that `len(tests_7d) <= len(rows_7d)` and likewise
for 28-day, so 5-into-5 fits exactly.

### main.py — new `ShotcreteLedgerPreviewWindow` class

Created by copying `LedgerPreviewWindow` and applying these
modifications:

- Constructor calls `writer.find_shotcrete_ledger_candidates()` (not the
  concrete function), and stores the result in `self._candidates`. The
  file-selector row (label vs `CTkOptionMenu`) and `_load_ledger` /
  `_refresh_view` logic are identical in shape; only the writer-function
  names change.
- `_load_ledger(candidate)` calls
  `writer.read_shotcrete_ledger_blocks`,
  `writer.read_shotcrete_ledger_values`,
  `writer.merge_shotcrete_cubes_for_ledger`, and the unchanged
  `writer.match_cubes_to_blocks`.
- The per-entry `vals` dict now carries four arrays
  (`diameters`, `heights`, `weights`, `loads`) instead of two.
- `_build_card` lays out a 4-column grid of value cells per row
  (Diameter | Height | Weight | Load). Reuses the same "green border =
  empty, gray = already filled" rendering — applied per-cell across all
  four columns. Header label additionally shows a small orange "SHOTCRETE"
  pill next to the cube title.
- `_do_write` calls `writer.write_shotcrete_ledger_cube` and passes the
  per-group `check_7d` / `check_28d` toggles.
- Title bar text: "Shotcrete Ledger Preview — {sheet_name}".

### main.py — MainWindow "Books Excel" group

The existing single `Go to Books Excel section` button is replaced by a
small group:

```
        Books Excel
   ┌────────────┐ ┌────────────┐
   │  CONCRETE  │ │  SHOTCRETE │
   └────────────┘ └────────────┘
```

- A small caption label "Books Excel" (size 11, gray55) above two
  buttons.
- `CONCRETE` — same blue `#1565C0` as today; opens `LedgerPreviewWindow`.
- `SHOTCRETE` — orange `#EF6C00`; opens `ShotcreteLedgerPreviewWindow`.
- Identical button width / height to the existing button so MainWindow
  layout footprint stays unchanged.

## Data flow

The first PreviewWindow already auto-persists every shotcrete row's
edits (weight, load, diameter, height, strength, plus the `_selected`
flag) back into `cubes_data["cubes"][i]["tests"]` on close
(`_persist_edits_to_cubes`, main.py:1531). The shotcrete ledger pass
reads the same `cubes_data`, so manual corrections in the first window
flow through automatically. No further plumbing is required.

The `_selected` flag is preserved on each test for the per-sheet pass
to honour the user's top-3 override, but the shotcrete ledger merge
ignores it (see decision above).

## Edge cases

- **No shotcrete ledger open** → `find_shotcrete_ledger_candidates`
  raises with a shotcrete-specific "open the file first" message; the
  window shows the error and stays.
- **No shotcrete cubes in PDF** → the merge returns an empty list; the
  window opens with the existing "No cubes to write" empty state.
- **Excel cell already filled** → the writer skips that cell, the card
  shows it in gray (concrete behaviour, reused).
- **Specimen value missing** (Gemini missed a row) → that one cell is
  left blank, the rest are written; the card shows blank input.
- **Master "Write this sample" unticked in first PreviewWindow** →
  carry-over filter drops the cube from the merge.
- **"WP" (work pending) rows in the ledger** — Age column is text, not
  7 or 28; the existing block walker only adds rows with `age == 7` or
  `age == 28` to `rows_7d` / `rows_28d`, so these rows are ignored
  naturally.
- **Two-set shotcrete (10 per age)** — assumed to share one 20-row
  block in the ledger; confirm against a real two-set sample before
  declaring complete. If layout turns out to be two distinct blocks
  instead, the merge can be relaxed not to combine sub-cubes
  (`_set_index` stays separate).
- **User picks the wrong workbook** in the shotcrete dropdown — header
  detection ensures only files with `Core No.` headers appear, so the
  concrete ledger never shows up in the shotcrete list (and vice versa).

## Risk / mitigation

- **Concrete ledger flow must not regress.** A local snapshot of the
  pre-change `main.py` and `writer.py` is stored under
  `_backup_pre_shotcrete/` (gitignored); git tag `v1.0.10` is the
  authoritative restore point.
- **MainWindow layout is being modified.** The two new buttons replace
  the one existing button, occupying the same footprint to avoid
  shifting other controls.
- **Risk-isolated sibling class** means the concrete preview's UI and
  writer chain are not edited — only added next to.

## Implementation plan

1. `writer.py`: add the five shotcrete functions next to the concrete
   ones; concrete functions untouched.
2. `main.py`: copy `LedgerPreviewWindow` → `ShotcreteLedgerPreviewWindow`,
   swap function calls and card layout to 4 columns.
3. `main.py` MainWindow: replace the single Books button with the
   `[CONCRETE] [SHOTCRETE]` group.
4. Update `preview_ledger_test.py` (or add a sibling
   `preview_shotcrete_ledger_test.py`) to exercise the new window with
   fake 5+5 blocks.
5. Verify concrete regression: open `LedgerPreviewWindow` end-to-end
   with one real concrete sample.
6. Verify shotcrete: open `ShotcreteLedgerPreviewWindow` with the
   shotcrete Excel and at least one real shotcrete cube.
7. Mirror changed files into `dist/CubeLogReader/src/`, bump
   `version.txt`, cut a release per `project_auto_update`.
