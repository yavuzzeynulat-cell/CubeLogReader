# Ledger Writer — Design

**Date:** 2026-04-17
**Status:** Approved (design phase). Implementation plan to follow via `writing-plans`.

## Goal

After the existing per-sample Excel write completes, allow the user to also
write the same weight/load values into a single "ledger" Excel workbook —
a long, downward-growing sheet that aggregates every cube across the
project. No Gemini re-read; ledger preview reuses the `cubes_data`
already produced for the first Excel flow.

## User flow

1. User loads a PDF, Gemini extracts `cubes_data`, first `PreviewWindow`
   opens — unchanged.
2. User writes to the per-sample Excel — unchanged.
3. When that write completes, MainWindow enables a new button
   **"Go to Books Excel section"** (disabled by default; re-disables if
   the user loads a new PDF or starts a new job).
4. User opens the ledger workbook manually in the same Excel instance
   (it is always already open in practice, but we do not assume it).
5. User clicks the button → new `LedgerPreviewWindow` opens.
6. The window detects the ledger `Concrete` sheet, finds the matching
   block for each cube in the in-memory `cubes_data`, and renders a
   preview.
7. User clicks **"Write to Ledger"** → M (weight) and N (load) cells are
   written for every selected card. Cells that are already filled are
   skipped silently (same rule as the first Excel).

The ledger file is shared with an external workflow; only the two
columns we write to (M, N) are our concern. L (age) and O (strength)
are formulas and update themselves once M and N are written.

## Data flow

### Input

`cubes_data` already in `MainWindow` memory after the first write. It is
the post-processing form: shotcrete flagged with `_shotcrete=True`, top-3
per age selected, multi-set cubes split into sub-cubes with
`_set_index` / `_set_total`.

### Pre-processing for ledger

Before matching against the ledger sheet:

1. **Drop shotcrete:** any sub-cube with `_shotcrete=True` is filtered
   out silently. A separate shotcrete ledger is deferred to a later
   feature.
2. **Merge sub-cubes back by `sample_mark`:** for ledger purposes, set 1
   and set 2 of the same cube form a single logical cube whose tests
   are the concatenation (in `_set_index` order) of the sub-cubes'
   `_selected` tests. (For normal non-multi-set cubes this is a no-op.)

The merged list is what we preview and write.

### Ledger sheet detection

- Iterate open workbooks (via existing `connect_to_excel()` COM handle).
- For each workbook, look for a sheet named `Concrete`.
- Validate by checking header row 7: `A7 == "Cube No."` and
  `B7 == "Site Sampling Mark/No:"` (tolerant of surrounding whitespace).
- 0 matches → error dialog, abort.
- 1 match → use it.
- 2+ matches → small dialog asking user to pick (rare case; shouldn't
  happen in the real workflow but must not crash).

### Block detection

Read column B in one COM call: `ws.Range("B8:B30000").Value` returns a
tuple of tuples. Parse in Python:

- A row whose B cell is non-empty starts a block.
- The block extends until the row before the next B-non-empty row (or
  end of data).
- Produce a list of `{sample_id_num, start_row, end_row, size}` where
  `sample_id_num` comes from `normalize_sample_mark(B_value)` — the
  same function that matches the first Excel's B14. This absorbs
  format differences like `G26-CON-522` ↔ `G-CON-0522`.

Only blocks whose `sample_id_num` is non-None participate; blocks with
unparseable marks (rare typos) are ignored.

### Cube → block matching

Mirror `match_cubes_to_sheets`: each merged cube finds its block by
`sample_id_num` equality. Since blocks are unique per sample in the
ledger, no "used blocks" bookkeeping is needed (we do not expect two
blocks with the same sample on the same sheet).

### Row split inside a block

- Assume a block's first half is 7-day rows, second half is 28-day rows
  (ledger convention observed across all existing entries).
- If `size` is odd, flag the card as **mismatched** and skip it on
  Write. Preview shows a red border + warning.
- If `size / 2` ≠ number of 7-day tests we have for that cube (e.g.
  block is 6 rows but we have 6 × 7-day tests from a 2-set cube) →
  flag as mismatched, same red treatment.

### Value writing

For each selected (card checkbox ON) and non-mismatched cube:

- 7-day tests: iterate in order, write `weight_gr` into
  `M{start_row + i}` and `load_kn` into `N{start_row + i}`.
- 28-day tests: iterate, write into rows `start_row + size/2 + i`.
- **Skip-if-filled:** before writing any cell, read its current value.
  If non-empty, leave it alone and mark it "skipped" in the run summary.

Writing is done through COM like the existing writer; one cell at a
time is fine for the volume involved (≤24 cells per cube, ≤300 cubes
per PDF in the worst case).

## UI — LedgerPreviewWindow

Matches the existing preview style: header bar, `corner_radius=12`
cards, `#2E7D32` accent, scroll region.

### Top of window
- Title: `Ledger Preview — Concrete Sheet`
- Red warning panel (only if there are any): `⚠️ N küp ledger'da
  bulunamadı: G-CON-XXX, G-CON-YYY, ...` — these are the merged cubes
  whose `sample_id_num` did not match any block.

### Cards (one per merged cube — per `sample_mark`)
- Title: `G-CON-522` with a subtitle badge when multi-set:
  `2 set · 12 test`
- Two-column layout: 7-günlük | 28-günlük
- Each cell shows `Weight (gr)` and `Load (kN)` pulled from
  `cubes_data`. Ledger's current M/N read (via the same Range batch)
  colors the cell:
  - Green border → ledger cell is empty (we'll write it)
  - Gray, muted text → ledger cell is already filled (we'll skip)
- Card-level checkbox: **"Write this sample"** (default ON). User may
  turn off to skip an entire card.
- Mismatched blocks (see "Row split" above) render with a red border
  and an explanatory line (`Blok boyutu 6, beklenen 12`); the checkbox
  is forced off and the card is skipped at write time.

### Bottom of window
- Big green **Write to Ledger** button.
- After write completes: summary message — e.g.
  `12 küp yazıldı · 14 hücre atlandı (zaten dolu) · 2 bulunamadı ·
   0 hata`.
- Close on dismiss.

## Error handling

- **Ledger not open:** dialog explaining the user must open the ledger
  workbook in Excel, then click again.
- **No cubes match any block:** show preview with only the "bulunamadı"
  red list, Write button disabled.
- **Block size mismatch per cube:** handled inline per card (red,
  skipped), does not abort the whole window.
- **COM error during write:** log full traceback to
  `gemini_debug.log` via `reader._log(...)` (same pattern as main),
  messagebox shows `str(err)` truncated to 500 chars + pointer to
  log.
- **User clicks the button twice:** button disables itself while the
  window is open to prevent double-open.

## Code organization

New code:

- `writer.py` — additions:
  - `find_ledger_sheet()` → returns `(workbook, sheet)` or raises.
  - `read_ledger_blocks(ws)` → returns `[{sample_id_num, start_row,
    end_row, size}, ...]` after one `Range("B8:B30000").Value` COM call.
  - `read_ledger_values(ws, blocks)` → returns a parallel dict of the
    current M/N values for each block's rows, one batched range read,
    used by preview to decide green/gray.
  - `match_cubes_to_blocks(merged_cubes, blocks)` → returns a list of
    `{cube, block, mismatch_reason}` entries.
  - `write_ledger_cube(ws, cube, block)` → the actual M/N write.
- `main.py` — additions:
  - `MainWindow`: new button, enable/disable hooks tied to first-write
    completion.
  - `LedgerPreviewWindow` class — a sibling of `PreviewWindow`, sharing
    the header/card styling but simpler internals (no shotcrete
    override, no per-group checkboxes, no 5-row logic). Kept as a
    separate class rather than a mode-flag on `PreviewWindow` to keep
    each window focused.

`reader.py` — no changes. The ledger flow consumes the same
post-processed `cubes_data` the first preview already uses.

## Testing (manual, user-led — no sample data in developer env)

Real-PDF checks:
1. Single-set normal cube (6-row block, 3×7d + 3×28d) → writes both sets.
2. 2-set normal cube (12-row block, 6×7d + 6×28d) → all 6 weights and
   loads per age go into the block in order.
3. 3-set normal cube (18-row block if the ledger has one pre-populated
   that way) → same as above, ceil-extended.
4. Pre-filled cell (user wrote a value manually) → skipped, summary
   reports `1 cell skipped (already filled)`.
5. Sample_mark present in PDF but absent from ledger → red warning,
   no card, no write.
6. Shotcrete cube mixed in the PDF → never appears in the ledger
   preview.
7. Ledger Excel not open at click time → dialog, no crash.

## Out of scope / deferred

- Shotcrete ledger (separate Excel, different cells including core
  diameter/height). Planned as a follow-up once this normal-cube
  ledger flow is validated.
- `Mix Designes` sheet writing. Not used by the user's current workflow.
- Batch/multi-file runs — the existing single-file flow stays.
- Edit dates / metadata in the ledger — only M and N.

## Open questions / assumptions to verify at implementation

- **Normalization:** `normalize_sample_mark` is assumed to already map
  `G-CON-0522` and `G26-CON-522` to the same numeric key. If not, the
  fix is localized to that one function and benefits the first-Excel
  flow too.
- **Block size heuristics:** a few observed blocks have odd sizes
  (3, 5, 11, 21). These are either incomplete blocks (only 7d or only
  28d present), shotcrete entries, or manual curiosities. The current
  design marks odd-sized and mismatched blocks as "skip with warning"
  rather than attempting to write — safest default. If the user wants
  a smarter rule later (e.g. write only the 7d half when block has
  only 7d rows), we can extend.
- **COM range read on 24k rows:** one `Range("B8:B30000").Value` call
  is fast enough in Excel; if profiling shows otherwise we can bound
  the read at `UsedRange.Rows.Count`.
