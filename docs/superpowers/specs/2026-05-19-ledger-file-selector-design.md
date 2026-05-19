# Ledger file selector (Books Excel)

**Date:** 2026-05-19

## Problem

`writer.find_ledger_sheet()` scans all open workbooks for a "Concrete"
ledger sheet. When **2+** open Excel files have one, it raises
`RuntimeError("Multiple ledger sheets found: ...")` and the user is stuck —
they must close a file. Users keep multiple ledgers open (e.g. last month +
this month) and want to pick which one to write to.

## Decisions (approved by user)

- **In-window selector only** — no separate modal pre-step.
- Selector lives at the top of `LedgerPreviewWindow`, just under the header.
- **1 ledger open** → static confirmation label ("Yazılacak dosya: X"),
  no choice. Behaviour otherwise identical to today.
- **2+ ledgers open** → small `CTkOptionMenu` dropdown listing file names;
  default = first candidate. Picking another reloads the window.
- UI must be **subtle and small**: size-11 gray text, not bold, tight pad.

## Design

### writer.py
- New `find_ledger_candidates()` → `list[(wb, ws, wb_name, sheet_name)]`.
  Same detection as `find_ledger_sheet` (sheet name + A7/B7 headers).
  - 0 candidates → `RuntimeError` ("open the ledger file first") — unchanged.
  - 1 or more → return the list; **the >1 error is removed**.
- Keep `find_ledger_sheet()` as-is for any other caller, or have it call
  `find_ledger_candidates()` and return `[0]` — minimal, low-risk.

### LedgerPreviewWindow (main.py)
- `__init__`: call `find_ledger_candidates()`, store `self._candidates`,
  `self._cand_index = 0`.
- Extract the data-load + card-build code currently inline in `__init__`
  (find sheet → `read_ledger_blocks` → `read_ledger_values` →
  `merge_cubes_for_ledger` → `match_cubes_to_blocks` → build cards) into a
  re-callable method `_load_ledger(candidate)`. It sets `self.wb_name`,
  `self.sheet_name`, `self._ws`, then **clears any existing cards** and
  rebuilds them.
- Add a small selector row under the header:
  - 1 candidate → `CTkLabel` "Yazılacak dosya: {name}" (size 11, gray55).
  - 2+ → `CTkLabel` "Ledger dosyası:" + small `CTkOptionMenu` (file names).
    Dropdown `command` → set `_cand_index`, call `_load_ledger(...)`.
- `_do_write` already writes via `self._ws`; `_load_ledger` keeps `_ws`
  current, so no change needed there (verify).

## Risk / mitigation

- **Single-candidate path = old behaviour** + a static label → the working
  flow is not changed.
- The reload path re-runs exactly what `__init__` did for the data section;
  factoring it into `_load_ledger` is the only structural change.
- Edge: a selected workbook closed mid-session → `_load_ledger` fails;
  catch and show a message, keep the previous view.

## Implementation plan

1. `writer.py`: add `find_ledger_candidates()`; keep `find_ledger_sheet`.
2. `LedgerPreviewWindow.__init__`: extract data-load + card-build into
   `_load_ledger(candidate)` (clears cards first, then rebuilds).
3. Add the selector row (label vs dropdown by candidate count), subtle style.
4. Wire dropdown change → `_load_ledger`.
5. Verify `_do_write` targets the currently selected worksheet.
6. Test: `preview_ledger_test.py` harness (mock) + real run — single file
   unchanged; two ledgers open → dropdown switches and reloads correctly.
7. Mirror changed files into `dist/CubeLogReader/src/`, cut a release.
