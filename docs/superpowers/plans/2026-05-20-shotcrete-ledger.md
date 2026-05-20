# Shotcrete Ledger Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a Books Excel pass for shotcrete cubes — a sibling
`ShotcreteLedgerPreviewWindow` that writes Diameter / Height / Weight /
Load (columns L/M/N/O) to a separate Excel ("Shotcrete sample form",
sheet "Shotcrete Concrete Results"). MainWindow gets a grouped
`[CONCRETE] [SHOTCRETE]` button pair under a "Books Excel" caption.

**Architecture:** Sibling-class approach — copy `LedgerPreviewWindow`
to `ShotcreteLedgerPreviewWindow` and adapt only the differences (4
value columns, K-not-L Age column, separate writer functions). Five
new writer functions live next to the existing concrete ones; concrete
code is not edited. Risk isolation is the priority; future bugfixes
may need to be applied in both places.

**Tech Stack:** Python 3.11+, customtkinter (CTk), pywin32 / win32com
(Excel COM), PIL (PreviewWindow only). No automated test framework —
visual verification through standalone harness scripts
(`preview_*_test.py`) is the project's convention.

**Reference spec:**
`docs/superpowers/specs/2026-05-20-shotcrete-ledger-design.md`

---

## File Structure

**New files:**
- `preview_shotcrete_ledger_test.py` — visual harness with mocked Excel
  data (sibling of `preview_ledger_test.py`)

**Modified files:**
- `writer.py` — add 5 shotcrete functions next to the existing concrete
  ledger functions (around lines 490–820). No edits to existing code.
- `main.py` — add `ShotcreteLedgerPreviewWindow` class after
  `LedgerPreviewWindow` (around line 2400+, before any module-level
  trailing code). Update `MainWindow._build_ui` (or the ledger-button
  region) to host the `[CONCRETE] [SHOTCRETE]` button group.

**Backup already in place:**
- `_backup_pre_shotcrete/main.py` + `writer.py` (gitignored)
- Git tag `v1.0.10` is the authoritative restore point

**Testing convention used in this plan:** the project has no pytest
suite. Each task ends with running the harness (`python
preview_shotcrete_ledger_test.py`) and a quick visual check. The
hand-on-keyboard live test in Task 10 is mandatory; harness tests are
fast iteration only.

---

## Task 1: writer.py — `find_shotcrete_ledger_candidates`

**Files:**
- Modify: `writer.py` — add new function after `find_ledger_sheet`
  (around line 525)

- [ ] **Step 1: Add the new function**

Add to `writer.py` immediately after `find_ledger_sheet` (after the
return at the end of that function):

```python
def find_shotcrete_ledger_candidates():
    """
    Locate every open workbook holding a shotcrete ledger sheet.
    Detection is by row-7 header text: A7 contains "core no" (note: the
    concrete ledger says "cube no" — single-word difference) and B7
    contains "sampling mark". Sheet tab name is NOT checked; the real
    workbook uses "Shotcrete Concrete Results" but other names are fine.
    Returns a list of (workbook_com, worksheet_com, workbook_name, sheet_name).
    Raises RuntimeError only when there are 0 candidates.
    """
    excel = connect_to_excel()
    candidates = []
    for wb in excel.Workbooks:
        for ws in wb.Worksheets:
            try:
                a7 = ws.Range("A7").Value
                b7 = ws.Range("B7").Value
            except Exception:
                continue
            a7s = str(a7 or "").lower()
            b7s = str(b7 or "").lower()
            if "core no" in a7s and "sampling mark" in b7s:
                candidates.append((wb, ws, str(wb.Name), str(ws.Name)))
    if not candidates:
        raise RuntimeError(
            "Shotcrete ledger not found: no open Excel workbook has a "
            "sheet with 'Core No.' / 'Site Sampling Mark/No:' headers on "
            "row 7. Open the shotcrete ledger file first."
        )
    return candidates
```

- [ ] **Step 2: Quick syntax check**

Run: `python -m py_compile writer.py`
Expected: no output (success)

- [ ] **Step 3: Commit**

```bash
git add writer.py
git commit -m "feat(writer): add find_shotcrete_ledger_candidates"
```

---

## Task 2: writer.py — `read_shotcrete_ledger_blocks`

**Files:**
- Modify: `writer.py` — add after `read_ledger_blocks` (around line 615)

- [ ] **Step 1: Add the new function**

Add to `writer.py` after the `return blocks` of `read_ledger_blocks`:

```python
def read_shotcrete_ledger_blocks(ws) -> list[dict]:
    """
    Shotcrete sibling of read_ledger_blocks. Scans columns A:K from row 8
    downward; identifies blocks by B-filled rows; collects 7-day and
    28-day row numbers from column K (Age). Empty K rows ("WP", blank,
    or other text) are ignored. Block size = 5×7d + 5×28d for a single
    set, 10×7d + 10×28d for a two-set sample.
    Returns [{sample_key, sample_id_num, sample_mark_raw, cube_no,
              cube_no_raw, start_row, end_row, size, rows_7d, rows_28d}, ...].
    """
    try:
        used = ws.UsedRange
        last_row = int(used.Row) + int(used.Rows.Count) - 1
    except Exception:
        last_row = LEDGER_MAX_SCAN_ROWS
    last_row = min(last_row, LEDGER_MAX_SCAN_ROWS)
    if last_row < 8:
        return []

    # A=0, B=1, ..., K=10
    data = ws.Range(f"A8:K{last_row}").Value
    if data is None:
        return []
    if not isinstance(data[0], tuple):
        data = (data,)

    # Trim trailing fully-empty rows
    last_idx = -1
    for i, row in enumerate(data):
        if not all(_is_empty(v) for v in row):
            last_idx = i
    if last_idx < 0:
        return []
    data = data[: last_idx + 1]

    head_indices = [i for i, row in enumerate(data) if not _is_empty(row[1])]
    blocks: list[dict] = []
    for j, hi in enumerate(head_indices):
        start = 8 + hi
        prov_end_idx = (
            head_indices[j + 1] - 1 if j + 1 < len(head_indices)
            else len(data) - 1
        )
        rows_7d: list[int] = []
        rows_28d: list[int] = []
        actual_end_idx = hi
        for k in range(hi, prov_end_idx + 1):
            age_val = data[k][10]  # column K
            if _is_empty(age_val):
                continue
            try:
                age_int = int(age_val)
            except (TypeError, ValueError):
                # "WP" or other text — skip but still trim end past it
                actual_end_idx = k
                continue
            if age_int == 7:
                rows_7d.append(8 + k)
                actual_end_idx = k
            elif age_int == 28:
                rows_28d.append(8 + k)
                actual_end_idx = k
            else:
                actual_end_idx = k

        end = 8 + actual_end_idx
        raw = data[hi][1]
        cube_no_raw = data[hi][0]
        blocks.append({
            "sample_key": ledger_sample_key(raw),
            "sample_id_num": normalize_sample_mark(raw),
            "sample_mark_raw": raw,
            "cube_no": normalize_cube_no(cube_no_raw),
            "cube_no_raw": cube_no_raw,
            "start_row": start,
            "end_row": end,
            "size": end - start + 1,
            "rows_7d": rows_7d,
            "rows_28d": rows_28d,
        })
    return blocks
```

- [ ] **Step 2: Quick syntax check**

Run: `python -m py_compile writer.py`
Expected: no output (success)

- [ ] **Step 3: Commit**

```bash
git add writer.py
git commit -m "feat(writer): add read_shotcrete_ledger_blocks (A:K, age=K)"
```

---

## Task 3: writer.py — `read_shotcrete_ledger_values`

**Files:**
- Modify: `writer.py` — add after `read_ledger_values` (around line 645)

- [ ] **Step 1: Add the new function**

Add to `writer.py` after the `return result` of `read_ledger_values`:

```python
def read_shotcrete_ledger_values(ws, blocks: list[dict]) -> dict:
    """
    Shotcrete sibling of read_ledger_values. Reads four columns per row:
    L (Core Diameter), M (Core Height), N (Weight), O (Load) for every
    row covered by `blocks`. One batched COM read.
    Returns {block_index: {"diameters":[...], "heights":[...],
                           "weights":[...], "loads":[...]}}.
    """
    if not blocks:
        return {}
    first = blocks[0]["start_row"]
    last = blocks[-1]["end_row"]
    data = ws.Range(f"L{first}:O{last}").Value
    if data is None:
        return {
            i: {
                "diameters": [None] * b["size"],
                "heights": [None] * b["size"],
                "weights": [None] * b["size"],
                "loads": [None] * b["size"],
            }
            for i, b in enumerate(blocks)
        }
    if not isinstance(data[0], tuple):
        data = (data,)
    result: dict = {}
    for i, b in enumerate(blocks):
        off = b["start_row"] - first
        n = b["size"]
        rows = data[off: off + n]
        result[i] = {
            "diameters": [r[0] for r in rows],
            "heights":   [r[1] for r in rows],
            "weights":   [r[2] for r in rows],
            "loads":     [r[3] for r in rows],
        }
    return result
```

- [ ] **Step 2: Quick syntax check**

Run: `python -m py_compile writer.py`
Expected: no output (success)

- [ ] **Step 3: Commit**

```bash
git add writer.py
git commit -m "feat(writer): add read_shotcrete_ledger_values (L:O)"
```

---

## Task 4: writer.py — `merge_shotcrete_cubes_for_ledger`

**Files:**
- Modify: `writer.py` — add after `merge_cubes_for_ledger` (around line 700)

- [ ] **Step 1: Add the new function**

Add to `writer.py` immediately after the `return` of
`merge_cubes_for_ledger`:

```python
def merge_shotcrete_cubes_for_ledger(cubes_data: dict) -> list[dict]:
    """
    Shotcrete sibling of merge_cubes_for_ledger. Two inversions:
      1. Keep ONLY cubes with _shotcrete=True (concrete dropped them).
      2. Do NOT skip tests where _selected is False — all 5 specimens
         per age are forwarded so the ledger receives every result.
    Carry-overs from the concrete merge:
      - Skip cubes where _card_enabled is False (master tick unset in
        the first PreviewWindow).
      - Merge multi-set sub-cubes back by (sample_key, cube_no);
        concatenate tests in _set_index order.

    Returns a list of {sample_key, sample_id_num, sample_mark, cube_no,
    tests_7d, tests_28d}.
    """
    by_num: dict = {}
    order: list = []
    for cube in cubes_data.get("cubes", []):
        if not cube.get("_shotcrete"):
            continue
        if cube.get("_card_enabled") is False:
            continue
        key = ledger_sample_key(cube.get("sample_mark"))
        if key is None:
            continue
        cube_no = normalize_cube_no(cube.get("cube_no"))
        set_idx = cube.get("_set_index") or 1
        merge_key = (key, cube_no)
        if merge_key not in by_num:
            by_num[merge_key] = {
                "sample_key": key,
                "sample_id_num": normalize_sample_mark(cube.get("sample_mark")),
                "sample_mark": cube.get("sample_mark"),
                "cube_no": cube_no,
                "_sets": {},
            }
            order.append(merge_key)
        entry = by_num[merge_key]
        sets = entry["_sets"]
        if set_idx not in sets:
            sets[set_idx] = ([], [])
        t7, t28 = sets[set_idx]
        for t in cube.get("tests", []):
            age = t.get("age_days")
            if age == 7:
                t7.append(t)
            elif age == 28:
                t28.append(t)

    out = []
    for merge_key in order:
        entry = by_num[merge_key]
        sets = entry.pop("_sets")
        tests_7d: list = []
        tests_28d: list = []
        for set_idx in sorted(sets.keys()):
            t7, t28 = sets[set_idx]
            tests_7d.extend(t7)
            tests_28d.extend(t28)
        entry["tests_7d"] = tests_7d
        entry["tests_28d"] = tests_28d
        out.append(entry)
    return out
```

- [ ] **Step 2: Quick syntax check**

Run: `python -m py_compile writer.py`
Expected: no output (success)

- [ ] **Step 3: Commit**

```bash
git add writer.py
git commit -m "feat(writer): add merge_shotcrete_cubes_for_ledger (keep _shotcrete, all 5)"
```

---

## Task 5: writer.py — `write_shotcrete_ledger_cube`

**Files:**
- Modify: `writer.py` — add after `write_ledger_cube` (around line 820)

- [ ] **Step 1: Add the new function**

Add to `writer.py` immediately after the `return` of
`write_ledger_cube`:

```python
def write_shotcrete_ledger_cube(
    ws,
    cube: dict,
    block: dict,
    write_7d: bool = True,
    write_28d: bool = True,
) -> dict:
    """
    Shotcrete sibling of write_ledger_cube. Writes four values per row
    to columns L (Diameter), M (Height), N (Weight), O (Load) for each
    specimen in `cube`'s tests_7d and tests_28d, aligned with the
    block's `rows_7d` / `rows_28d` row numbers.

    Skips a cell when the ledger cell is already non-empty (no overwrite
    — same policy as write_ledger_cube). Skips a cell when the source
    test value is None (so missing readings don't blank existing data).

    Returns {"wrote": [(row, col)...], "skipped": [(row, col, reason)...],
             "errors": [...]}.
    """
    wrote: list = []
    skipped: list = []
    errors: list = []

    def _write_one(row: int, col_letter: str, value):
        if value is None:
            skipped.append((row, col_letter, "no source value"))
            return
        cell = f"{col_letter}{row}"
        try:
            existing = ws.Range(cell).Value
        except Exception as e:
            errors.append(f"{cell}: read failed: {e}")
            return
        if not _is_empty(existing):
            skipped.append((row, col_letter, "already filled"))
            return
        try:
            ws.Range(cell).Value = value
        except Exception as e:
            errors.append(f"{cell}: write failed: {e}")
            return
        wrote.append((row, col_letter))

    def _walk(tests, rows, enabled):
        if not enabled:
            return
        for t, row in zip(tests, rows):
            _write_one(row, "L", t.get("core_diameter_mm"))
            _write_one(row, "M", t.get("core_height_mm"))
            _write_one(row, "N", t.get("weight_gr"))
            _write_one(row, "O", t.get("load_kn"))

    _walk(cube.get("tests_7d", []), block["rows_7d"], write_7d)
    _walk(cube.get("tests_28d", []), block["rows_28d"], write_28d)

    # Optional debug log (mirrors write_ledger_cube style)
    try:
        from reader import _log
        from datetime import datetime
        ts = datetime.now().strftime("%H:%M:%S")
        mark = cube.get("sample_mark", "?")
        _log(
            f"[{ts}] [SHOT-LEDGER] {mark} -> rows {block['start_row']}-"
            f"{block['end_row']}: wrote {len(wrote)}, skipped "
            f"{len(skipped)}, errors {len(errors)}"
        )
    except Exception:
        pass

    return {"wrote": wrote, "skipped": skipped, "errors": errors}
```

- [ ] **Step 2: Quick syntax check**

Run: `python -m py_compile writer.py`
Expected: no output (success)

- [ ] **Step 3: Commit**

```bash
git add writer.py
git commit -m "feat(writer): add write_shotcrete_ledger_cube (L/M/N/O writer)"
```

---

## Task 6: preview_shotcrete_ledger_test.py — visual harness

**Files:**
- Create: `preview_shotcrete_ledger_test.py`

- [ ] **Step 1: Create the harness file**

Create `preview_shotcrete_ledger_test.py` with this content:

```python
"""Quick visual test for ShotcreteLedgerPreviewWindow — no Excel needed.

Two fake open ledgers exercise the file-selector dropdown.
Four fake blocks cover: single 5+5 (clean), 5+5 with 7-day pre-filled,
two-set 10+10 interleaved, and an odd-count 4+3."""
import customtkinter as ctk
import main
import writer


class _FakeWS:
    def Range(self, _):
        class _R:
            Value = None
        return _R()


def _fake_find_candidates(*a, **k):
    return [
        (None, _FakeWS(), "Shotcrete sample form.xlsx", "Shotcrete Concrete Results"),
        (None, _FakeWS(), "Shotcrete sample form APRIL.xlsx", "Shotcrete Concrete Results"),
    ]


def _fake_blocks(ws):
    return [
        {  # 027 — clean 5+5
            "sample_key": "G26-CON-027", "sample_id_num": 27,
            "sample_mark_raw": "G26-CON-027", "cube_no": 2, "cube_no_raw": 2,
            "start_row": 2309, "end_row": 2318, "size": 10,
            "rows_7d":  [2309, 2310, 2311, 2312, 2313],
            "rows_28d": [2314, 2315, 2316, 2317, 2318],
        },
        {  # 082 — 7-day pre-filled in ledger (gray cells)
            "sample_key": "G26-CON-082", "sample_id_num": 82,
            "sample_mark_raw": "G26-CON-082", "cube_no": 3, "cube_no_raw": 3,
            "start_row": 2319, "end_row": 2328, "size": 10,
            "rows_7d":  [2319, 2320, 2321, 2322, 2323],
            "rows_28d": [2324, 2325, 2326, 2327, 2328],
        },
        {  # 200 — two-set: 10+10 interleaved across 20 rows
            "sample_key": "G26-CON-200", "sample_id_num": 200,
            "sample_mark_raw": "G26-CON-200", "cube_no": 5, "cube_no_raw": 5,
            "start_row": 2340, "end_row": 2359, "size": 20,
            "rows_7d":  [2340, 2341, 2342, 2343, 2344,
                         2350, 2351, 2352, 2353, 2354],
            "rows_28d": [2345, 2346, 2347, 2348, 2349,
                         2355, 2356, 2357, 2358, 2359],
        },
        {  # 333 — odd 4+3 block (partial data scenario)
            "sample_key": "G26-CON-333", "sample_id_num": 333,
            "sample_mark_raw": "G26-CON-333", "cube_no": 7, "cube_no_raw": 7,
            "start_row": 2370, "end_row": 2376, "size": 7,
            "rows_7d":  [2370, 2371, 2372, 2373],
            "rows_28d": [2374, 2375, 2376],
        },
    ]


def _fake_values(ws, blocks):
    out = {}
    for i, b in enumerate(blocks):
        n = b["size"]
        # 082 → 7-day rows already filled (so the card shows gray on them)
        if b["sample_id_num"] == 82:
            di = [94.0, 94.5, 94.8, 95.0, 95.3, None, None, None, None, None]
            he = [94.1, 94.6, 94.9, 95.1, 95.4, None, None, None, None, None]
            we = [1500.0, 1510, 1520, 1530, 1540, None, None, None, None, None]
            lo = [200.0, 210, 220, 230, 240, None, None, None, None, None]
        else:
            di = [None] * n
            he = [None] * n
            we = [None] * n
            lo = [None] * n
        out[i] = {"diameters": di, "heights": he, "weights": we, "loads": lo}
    return out


writer.find_shotcrete_ledger_candidates = _fake_find_candidates
writer.read_shotcrete_ledger_blocks = _fake_blocks
writer.read_shotcrete_ledger_values = _fake_values


def _shot_tests(n7, n28, base_w=1500, base_l=220, base_d=94.0, base_h=94.0):
    out = []
    for i in range(n7):
        out.append({
            "age_days": 7, "weight_gr": base_w + i,
            "load_kn": base_l + i, "core_diameter_mm": base_d + i * 0.1,
            "core_height_mm": base_h + i * 0.1, "strength_nmm2": 30 + i,
            "_selected": i < 3,
        })
    for i in range(n28):
        out.append({
            "age_days": 28, "weight_gr": base_w + 50 + i,
            "load_kn": base_l + 50 + i, "core_diameter_mm": base_d + 1 + i * 0.1,
            "core_height_mm": base_h + 1 + i * 0.1, "strength_nmm2": 40 + i,
            "_selected": i < 3,
        })
    return out


CUBES = {
    "cubes": [
        {"sample_mark": "G26-CON-027", "cube_no": 2, "_shotcrete": True,
         "tests": _shot_tests(5, 5)},
        {"sample_mark": "G26-CON-082", "cube_no": 3, "_shotcrete": True,
         "tests": _shot_tests(5, 5, base_w=1550, base_l=230)},
        {"sample_mark": "G26-CON-200", "cube_no": 5, "_shotcrete": True,
         "_set_index": 1, "tests": _shot_tests(5, 5, base_w=1600)},
        {"sample_mark": "G26-CON-200", "cube_no": 5, "_shotcrete": True,
         "_set_index": 2, "tests": _shot_tests(5, 5, base_w=1620)},
        {"sample_mark": "G26-CON-333", "cube_no": 7, "_shotcrete": True,
         "tests": _shot_tests(4, 3, base_w=1700)},
    ]
}


root = ctk.CTk()
root.geometry("400x100")
root.title("Shotcrete ledger test parent")
ctk.CTkLabel(root, text="ShotcreteLedgerPreviewWindow test").pack(pady=20)

main.ShotcreteLedgerPreviewWindow(root, CUBES)

root.mainloop()
```

- [ ] **Step 2: Quick syntax check**

Run: `python -m py_compile preview_shotcrete_ledger_test.py`
Expected: no output (success). The script imports `main.ShotcreteLedgerPreviewWindow`
which does not exist yet — that's fine, py_compile only checks syntax.

- [ ] **Step 3: Commit**

```bash
git add preview_shotcrete_ledger_test.py
git commit -m "test: add visual harness for ShotcreteLedgerPreviewWindow"
```

---

## Task 7: main.py — `ShotcreteLedgerPreviewWindow` class

**Files:**
- Modify: `main.py` — add new class after `LedgerPreviewWindow`
  (after its `_do_write` ends, around line 2280 — locate by searching
  for the `class LedgerPreviewWindow` and finding its trailing method)

This task creates the full sibling class. To keep step blocks small,
the class is built in three steps: scaffolding + data load, then UI,
then write handler.

- [ ] **Step 1: Find the insertion point and add the class scaffold**

Run: `grep -n "class LedgerPreviewWindow\|class MainWindow" main.py`
Expected: two line numbers, one for `LedgerPreviewWindow` and one for
`MainWindow`. The new class goes between them.

Insert this scaffold just before `class MainWindow`:

```python
# ---------- Shotcrete Ledger Preview Window ----------

class ShotcreteLedgerPreviewWindow:
    """
    Books Excel write pass for SHOTCRETE cubes. Sibling of
    LedgerPreviewWindow targeting a separate workbook
    (sheet "Shotcrete Concrete Results"). Reads/writes four value
    columns per row: L (Diameter), M (Height), N (Weight), O (Load).
    All 5 specimens per age are written — the per-sheet top-3 filter
    does NOT apply here.
    """

    def __init__(self, parent, cubes_data: dict, on_close=None):
        self.parent = parent
        self.cubes_data = cubes_data
        self._on_close_cb = on_close
        self._closed = False

        self.win = ctk.CTkToplevel(parent)
        self.win.title("Shotcrete Ledger — " + APP_TITLE)
        sw = self.win.winfo_screenwidth()
        sh = self.win.winfo_screenheight()
        ww = min(1400, sw - 60)
        wh = min(860, sh - 80)
        self.win.geometry(f"{ww}x{wh}+40+40")
        self.win.minsize(900, 600)
        try:
            self.win.state("zoomed")
        except Exception:
            pass
        self.win.protocol("WM_DELETE_WINDOW", self._close)

        _add_credit_footer(self.win)

        self._entry_font = ctk.CTkFont(family="Segoe UI", size=13)
        self._title_font = ctk.CTkFont(family="Segoe UI", size=14, weight="bold")
        self._hdr_font = ctk.CTkFont(family="Segoe UI", size=11, weight="bold")
        self._pill_font = ctk.CTkFont(family="Segoe UI", size=10, weight="bold")
        self._check_font = ctk.CTkFont(family="Segoe UI", size=12)
        self._age_font = ctk.CTkFont(family="Segoe UI", size=13, weight="bold")
        self._colhdr_font = ctk.CTkFont(family="Segoe UI", size=11, weight="bold")

        try:
            self.parent.withdraw()
        except Exception:
            pass

        self.ledger_error: str | None = None
        self.wb_name: str | None = None
        self.sheet_name: str | None = None
        self.entries: list[dict] = []
        self.not_found: list[str] = []
        self._ws = None

        self._cards: list = []
        self._selected_idx: int | None = None
        self._cube_list = None
        self._scroll_anim_id: str | None = None

        self._candidates: list = []
        self._cand_labels: list[str] = []
        self._cand_index = 0
        self._ledger_menu = None
        self._body = None
        self._title_label = None
        try:
            self._candidates = writer.find_shotcrete_ledger_candidates()
        except Exception as e:
            self.ledger_error = str(e)

        seen: dict[str, int] = {}
        for (_w, _ws, wb_name, _sh) in self._candidates:
            if wb_name in seen:
                seen[wb_name] += 1
                self._cand_labels.append(f"{wb_name} ({seen[wb_name]})")
            else:
                seen[wb_name] = 1
                self._cand_labels.append(wb_name)

        if self._candidates:
            self._load_ledger(self._candidates[0])

        self._build_ui()

    def _close(self):
        if self._closed:
            return
        self._closed = True
        self._ws = None
        try:
            self.win.destroy()
        except Exception:
            pass
        try:
            self.parent.deiconify()
        except Exception:
            pass
        if self._on_close_cb is not None:
            try:
                self._on_close_cb()
            except Exception:
                pass

    def _load_ledger(self, candidate) -> bool:
        """Read one shotcrete ledger workbook and populate self.entries.
        Returns False on failure (previous state preserved)."""
        wb, ws, wb_name, sh_name = candidate
        try:
            blocks = writer.read_shotcrete_ledger_blocks(ws)
            values = writer.read_shotcrete_ledger_values(ws, blocks)
            merged = writer.merge_shotcrete_cubes_for_ledger(self.cubes_data)
            matches = writer.match_cubes_to_blocks(merged, blocks)
        except Exception as e:
            self.ledger_error = str(e)
            return False

        try:
            from reader import _log
            _log(f"[SHOT-LEDGER-DEBUG] {wb_name}: {len(blocks)} blocks, {len(merged)} shotcrete cubes")
        except Exception:
            pass

        entries: list[dict] = []
        not_found: list[str] = []
        block_idx_by_id = {id(b): i for i, b in enumerate(blocks)}
        for m in matches:
            cube = m["cube"]
            block = m["block"]
            reason = m["mismatch_reason"]
            if block is None:
                not_found.append(cube.get("sample_mark", "?"))
                continue
            bi = block_idx_by_id[id(block)]
            vals = values.get(bi, {"diameters": [], "heights": [],
                                   "weights": [], "loads": []})
            entries.append({
                "cube": cube,
                "block": block,
                "mismatch": reason,
                "diameters_ledger": vals["diameters"],
                "heights_ledger":   vals["heights"],
                "weights_ledger":   vals["weights"],
                "loads_ledger":     vals["loads"],
                "enabled": BooleanVar(value=reason is None),
            })

        self.ledger_error = None
        self.wb_name = wb_name
        self.sheet_name = sh_name
        self._ws = ws
        self.entries = entries
        self.not_found = not_found
        self._cards = []
        self._selected_idx = None
        return True
```

- [ ] **Step 2: Add the UI builders to the same class**

Append these methods to `ShotcreteLedgerPreviewWindow` (right after
`_load_ledger`):

```python
    def _compute_title(self) -> str:
        if self.ledger_error or self.sheet_name is None:
            return "Shotcrete Ledger Preview — ERROR"
        ok_cnt = sum(1 for e in self.entries if e["mismatch"] is None)
        bad_cnt = sum(1 for e in self.entries if e["mismatch"] is not None)
        return (
            f"Shotcrete Ledger Preview — {self.sheet_name}  ·  "
            f"{ok_cnt} ready, {bad_cnt} mismatched, "
            f"{len(self.not_found)} not found"
        )

    def _build_ui(self):
        top = ctk.CTkFrame(self.win, corner_radius=0, height=60)
        top.pack(side="top", fill="x")
        top.pack_propagate(False)

        self._title_label = ctk.CTkLabel(
            top, text=self._compute_title(),
            font=ctk.CTkFont(size=14, weight="bold"),
        )
        self._title_label.pack(side="left", padx=20, pady=16)

        ctk.CTkButton(
            top, text="Cancel", width=90,
            fg_color="gray70", hover_color="gray60",
            command=self._close,
        ).pack(side="right", padx=(0, 10), pady=12)

        self.write_btn = ctk.CTkButton(
            top, text=">> WRITE TO LEDGER",
            width=200, height=36,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#EF6C00", hover_color="#E65100",
            command=self._do_write,
        )
        self.write_btn.pack(side="right", padx=(0, 10), pady=12)

        writable = any(e["mismatch"] is None for e in self.entries)
        if not writable:
            self.write_btn.configure(state="disabled")

        self._build_selector()
        self._build_body()

    def _build_selector(self):
        bar = ctk.CTkFrame(self.win, fg_color="transparent")
        bar.pack(side="top", fill="x", padx=22, pady=(6, 0))
        if not self._candidates:
            return
        sel_font = ctk.CTkFont(family="Segoe UI", size=11)
        if len(self._candidates) == 1:
            ctk.CTkLabel(
                bar, text=f"Yazılacak dosya: {self._cand_labels[0]}",
                font=sel_font, text_color="gray55",
            ).pack(side="left")
        else:
            ctk.CTkLabel(
                bar, text="Shotcrete ledger dosyası:",
                font=sel_font, text_color="gray55",
            ).pack(side="left", padx=(0, 6))
            self._ledger_menu = ctk.CTkOptionMenu(
                bar, values=self._cand_labels,
                command=self._on_select_ledger,
                font=sel_font, height=24, width=280,
                fg_color="gray25", button_color="gray30",
                button_hover_color="gray35", text_color="gray80",
                dropdown_font=sel_font,
            )
            self._ledger_menu.set(self._cand_labels[self._cand_index])
            self._ledger_menu.pack(side="left")

    def _on_select_ledger(self, choice: str):
        try:
            idx = self._cand_labels.index(choice)
        except ValueError:
            return
        if idx == self._cand_index:
            return
        prev = self._cand_index
        self._cand_index = idx
        if self._load_ledger(self._candidates[idx]):
            self._refresh_view()
        else:
            err = self.ledger_error or "unknown error"
            self.ledger_error = None
            self._cand_index = prev
            if self._ledger_menu is not None:
                self._ledger_menu.set(self._cand_labels[prev])
            messagebox.showwarning(
                "Shotcrete Ledger",
                "Could not load that ledger file:\n" + err,
                parent=self.win,
            )

    def _refresh_view(self):
        if self._title_label is not None:
            self._title_label.configure(text=self._compute_title())
        writable = any(e["mismatch"] is None for e in self.entries)
        self.write_btn.configure(state="normal" if writable else "disabled")
        if self._body is not None:
            try:
                self._body.destroy()
            except Exception:
                pass
        self._build_body()

    def _build_body(self):
        body = ctk.CTkFrame(self.win, fg_color="transparent")
        body.pack(side="top", fill="both", expand=True, padx=16, pady=12)
        self._body = body

        if self.ledger_error:
            ctk.CTkLabel(
                body,
                text="Shotcrete ledger error:\n" + self.ledger_error,
                text_color="#C62828",
                font=self._entry_font,
                justify="left", wraplength=900,
            ).pack(pady=40)
            return

        legend = ctk.CTkFrame(body, fg_color="transparent")
        legend.pack(fill="x", pady=(0, 6))
        ctk.CTkLabel(
            legend, text="Green = ledger cell empty (will write)",
            text_color=EMPTY_BORDER,
            font=ctk.CTkFont(size=11, weight="bold"),
        ).pack(side="left", padx=(0, 16))
        ctk.CTkLabel(
            legend, text="Gray = already filled (skip)",
            text_color="gray55",
            font=ctk.CTkFont(size=11),
        ).pack(side="left")

        if self.not_found:
            banner = ctk.CTkFrame(
                body, corner_radius=8, border_width=2,
                border_color="#C62828", fg_color="#2a1010",
            )
            banner.pack(fill="x", pady=(4, 10))
            joined = ", ".join(self.not_found)
            ctk.CTkLabel(
                banner,
                text=f"⚠ {len(self.not_found)} shotcrete cubes not found in ledger: {joined}",
                text_color="#FF8A80",
                font=ctk.CTkFont(size=12, weight="bold"),
                anchor="w", justify="left", wraplength=1100,
            ).pack(fill="x", padx=12, pady=8)

        scroll = ctk.CTkScrollableFrame(body, fg_color="transparent")
        scroll.pack(fill="both", expand=True)
        self._cube_list = scroll

        for entry in self.entries:
            self._build_card(scroll, entry)

        self.win.bind("<Down>", self._on_card_down)
        self.win.bind("<Up>", self._on_card_up)

        if not self.entries and not self.not_found:
            ctk.CTkLabel(
                body,
                text="No shotcrete cubes to write.",
                text_color="gray55",
                font=self._entry_font,
            ).pack(pady=20)

    def _on_card_down(self, _event=None):
        # Reuse the same navigation as LedgerPreviewWindow by delegating
        # to a shared free function if available; otherwise no-op.
        pass

    def _on_card_up(self, _event=None):
        pass
```

- [ ] **Step 3: Add the card builder + write handler**

Append these methods to `ShotcreteLedgerPreviewWindow`:

```python
    def _build_card(self, parent, entry: dict):
        cube = entry["cube"]
        block = entry["block"]
        mismatch = entry["mismatch"]
        bad = mismatch is not None

        card = ctk.CTkFrame(parent, corner_radius=12, border_width=1)
        card.pack(fill="x", padx=8, pady=8)
        self._cards.append(card)

        header = ctk.CTkFrame(card, fg_color="transparent")
        header.pack(fill="x", padx=15, pady=(12, 6))
        header.grid_columnconfigure(0, weight=1, uniform="h")
        header.grid_columnconfigure(1, weight=0)
        header.grid_columnconfigure(2, weight=1, uniform="h")

        cube_checkbox = ctk.CTkCheckBox(
            header, text="Write this sample",
            variable=entry["enabled"], font=self._check_font,
        )
        cube_checkbox.grid(row=0, column=0, sticky="w")
        if bad:
            cube_checkbox.configure(state="disabled")

        title_txt = (
            f"Core No {cube.get('cube_no', '?')}  ·  "
            f"{cube.get('sample_mark', '?')}"
        )
        ctk.CTkLabel(
            header, text=title_txt, font=self._title_font,
        ).grid(row=0, column=1, sticky="")

        # Orange SHOTCRETE pill on the right
        pill_text = "[MISMATCH]" if bad else "[SHOTCRETE]"
        pill_color = PILL_NO_COLOR if bad else "#EF6C00"
        ctk.CTkLabel(
            header, text=pill_text, font=self._pill_font,
            text_color="white", fg_color=pill_color,
            corner_radius=8, padx=10, pady=2,
        ).grid(row=0, column=2, sticky="e")

        if mismatch:
            ctk.CTkLabel(
                card, text=mismatch, text_color="#FF8A80",
                font=ctk.CTkFont(size=11),
                anchor="w", justify="left", wraplength=1100,
            ).pack(fill="x", padx=15, pady=(0, 6))

        # Per-group toggles (7d / 28d)
        toggles = ctk.CTkFrame(card, fg_color="transparent")
        toggles.pack(fill="x", padx=15, pady=(0, 4))
        entry["check_7d"] = BooleanVar(value=not bad)
        entry["check_28d"] = BooleanVar(value=not bad)
        ctk.CTkCheckBox(
            toggles, text="Write 7-day rows", variable=entry["check_7d"],
            font=self._check_font,
        ).pack(side="left", padx=(0, 16))
        ctk.CTkCheckBox(
            toggles, text="Write 28-day rows", variable=entry["check_28d"],
            font=self._check_font,
        ).pack(side="left")

        # Column header row: Age | Diameter | Height | Weight | Load
        col_hdr = ctk.CTkFrame(card, fg_color="transparent")
        col_hdr.pack(fill="x", padx=15, pady=(6, 0))
        widths = [60, 140, 140, 140, 140]
        labels = ["Age", "Diameter (mm)", "Height (mm)", "Weight (gr)", "Load (Kn)"]
        for i, (w, lbl) in enumerate(zip(widths, labels)):
            col_hdr.grid_columnconfigure(i, weight=0, minsize=w)
            ctk.CTkLabel(
                col_hdr, text=lbl, font=self._colhdr_font,
                text_color="gray60",
            ).grid(row=0, column=i, sticky="w", padx=(0, 6))

        tests_7 = cube.get("tests_7d", [])
        tests_28 = cube.get("tests_28d", [])
        rows_7d = block["rows_7d"]
        rows_28d = block["rows_28d"]
        n7 = max(len(tests_7), len(rows_7d), 1)
        n28 = max(len(tests_28), len(rows_28d), 1)

        diam_l = entry["diameters_ledger"]
        height_l = entry["heights_ledger"]
        weight_l = entry["weights_ledger"]
        load_l = entry["loads_ledger"]

        def _ledger_idx(rows, row_no):
            # Block-wide ledger arrays are indexed by offset from block start.
            try:
                return rows.index(row_no) if row_no in rows else None
            except Exception:
                return None

        def _row_cell(parent, text, filled_existing):
            """One value cell — entry box; green border if empty in ledger,
            gray text if already filled (skip-on-write)."""
            border = EMPTY_BORDER if not filled_existing else "gray40"
            text_color = ("black", "white") if not filled_existing else FILLED_TEXT
            e = ctk.CTkEntry(
                parent, width=130, height=30,
                border_color=border, border_width=2,
                font=self._entry_font, text_color=text_color,
            )
            if text is not None:
                e.insert(0, str(text))
            return e

        def _build_value_row(grid_parent, grid_row, age_text,
                             diam_val, height_val, weight_val, load_val,
                             diam_filled, height_filled, weight_filled,
                             load_filled):
            ctk.CTkLabel(
                grid_parent, text=age_text, font=self._age_font,
                width=widths[0],
            ).grid(row=grid_row, column=0, sticky="w")
            cells = []
            for col, (val, filled) in enumerate(
                [(diam_val, diam_filled), (height_val, height_filled),
                 (weight_val, weight_filled), (load_val, load_filled)],
                start=1,
            ):
                c = _row_cell(grid_parent, val, filled)
                c.grid(row=grid_row, column=col, sticky="w", padx=(0, 6))
                cells.append(c)
            return cells

        rows_frame = ctk.CTkFrame(card, fg_color="transparent")
        rows_frame.pack(fill="x", padx=15, pady=(2, 12))
        for i, w in enumerate(widths):
            rows_frame.grid_columnconfigure(i, weight=0, minsize=w)

        entry["diam_7d_entries"]   = []
        entry["height_7d_entries"] = []
        entry["weight_7d_entries"] = []
        entry["load_7d_entries"]   = []
        entry["diam_28d_entries"]   = []
        entry["height_28d_entries"] = []
        entry["weight_28d_entries"] = []
        entry["load_28d_entries"]   = []

        for i in range(n7):
            t = tests_7[i] if i < len(tests_7) else {}
            row_no = rows_7d[i] if i < len(rows_7d) else None
            li = _ledger_idx(block["rows_7d"], row_no) if row_no else None
            def _ledger(arr, idx):
                return arr[idx] if (idx is not None and idx < len(arr)) else None
            d_filled = not _is_none_or_empty(_ledger(diam_l, li))
            h_filled = not _is_none_or_empty(_ledger(height_l, li))
            w_filled = not _is_none_or_empty(_ledger(weight_l, li))
            l_filled = not _is_none_or_empty(_ledger(load_l, li))
            cells = _build_value_row(
                rows_frame, i, "7d",
                t.get("core_diameter_mm"), t.get("core_height_mm"),
                t.get("weight_gr"), t.get("load_kn"),
                d_filled, h_filled, w_filled, l_filled,
            )
            entry["diam_7d_entries"].append(cells[0])
            entry["height_7d_entries"].append(cells[1])
            entry["weight_7d_entries"].append(cells[2])
            entry["load_7d_entries"].append(cells[3])

        for i in range(n28):
            t = tests_28[i] if i < len(tests_28) else {}
            row_no = rows_28d[i] if i < len(rows_28d) else None
            li = _ledger_idx(block["rows_28d"], row_no) if row_no else None
            def _ledger(arr, idx):
                return arr[idx] if (idx is not None and idx < len(arr)) else None
            d_filled = not _is_none_or_empty(_ledger(diam_l, li))
            h_filled = not _is_none_or_empty(_ledger(height_l, li))
            w_filled = not _is_none_or_empty(_ledger(weight_l, li))
            l_filled = not _is_none_or_empty(_ledger(load_l, li))
            cells = _build_value_row(
                rows_frame, n7 + i, "28d",
                t.get("core_diameter_mm"), t.get("core_height_mm"),
                t.get("weight_gr"), t.get("load_kn"),
                d_filled, h_filled, w_filled, l_filled,
            )
            entry["diam_28d_entries"].append(cells[0])
            entry["height_28d_entries"].append(cells[1])
            entry["weight_28d_entries"].append(cells[2])
            entry["load_28d_entries"].append(cells[3])

    def _do_write(self):
        if self.ledger_error:
            return

        def _f(entry_widget):
            t = entry_widget.get().strip()
            if not t:
                return None
            try:
                return float(t)
            except ValueError:
                return None

        written_cubes = 0
        total_cells = 0
        total_skipped = 0
        all_errors: list[str] = []
        unchecked: list[str] = []

        for entry in self.entries:
            cube = entry["cube"]
            mark = cube.get("sample_mark", "?")
            if entry["mismatch"] is not None:
                continue
            if not entry["enabled"].get():
                unchecked.append(mark)
                continue
            do_7d = entry["check_7d"].get()
            do_28d = entry["check_28d"].get()
            if not do_7d and not do_28d:
                unchecked.append(mark)
                continue

            tests_7d = []
            for d, h, w, l in zip(
                entry["diam_7d_entries"], entry["height_7d_entries"],
                entry["weight_7d_entries"], entry["load_7d_entries"],
            ):
                tests_7d.append({
                    "age_days": 7,
                    "core_diameter_mm": _f(d),
                    "core_height_mm":   _f(h),
                    "weight_gr":        _f(w),
                    "load_kn":          _f(l),
                })
            tests_28d = []
            for d, h, w, l in zip(
                entry["diam_28d_entries"], entry["height_28d_entries"],
                entry["weight_28d_entries"], entry["load_28d_entries"],
            ):
                tests_28d.append({
                    "age_days": 28,
                    "core_diameter_mm": _f(d),
                    "core_height_mm":   _f(h),
                    "weight_gr":        _f(w),
                    "load_kn":          _f(l),
                })
            cube_for_write = dict(cube)
            cube_for_write["tests_7d"] = tests_7d
            cube_for_write["tests_28d"] = tests_28d

            try:
                result = writer.write_shotcrete_ledger_cube(
                    self._ws, cube_for_write, entry["block"],
                    write_7d=do_7d, write_28d=do_28d,
                )
                total_cells += len(result["wrote"])
                total_skipped += len(result["skipped"])
                all_errors.extend(
                    f"{mark}: {e}" for e in result["errors"]
                )
                if result["wrote"]:
                    written_cubes += 1
            except Exception as e:
                all_errors.append(f"{mark}: write error: {e}")

        lines = [
            f"{written_cubes} shotcrete cubes updated, {total_cells} cells written, "
            f"{total_skipped} skipped (already filled or no value)."
        ]
        if unchecked:
            lines.append("")
            lines.append("Skipped (unchecked):")
            for s in unchecked:
                lines.append("  - " + s)
        if all_errors:
            lines.append("")
            lines.append("Errors:")
            for e in all_errors:
                lines.append("  - " + e)

        messagebox.showinfo(
            "Shotcrete Ledger", "\n".join(lines), parent=self.win,
        )
        if not all_errors:
            self._close()
```

- [ ] **Step 4: Add the `_is_none_or_empty` helper (missing in main.py)**

This helper is referenced by `_build_card` but does NOT exist in
`main.py` (confirmed: zero hits during planning). Add it as a private
module-level helper near the top of `main.py`, just below the imports
and constants block:

```python
def _is_none_or_empty(v) -> bool:
    if v is None:
        return True
    if isinstance(v, str) and not v.strip():
        return True
    return False
```

- [ ] **Step 5: Quick syntax check**

Run: `python -m py_compile main.py`
Expected: no output (success)

- [ ] **Step 6: Run the visual harness**

Run: `python preview_shotcrete_ledger_test.py`
Expected: a window opens showing the Shotcrete Ledger preview with the
file-selector dropdown (two ledger files), and 4 cards (027 clean,
082 with gray 7-day cells, 200 two-set with 20 rows, 333 odd 4+3).
Each card has 4 value columns per row. Pill says SHOTCRETE in orange.
No errors in the console.

Close the window. If anything is visually wrong, fix and rerun.

- [ ] **Step 7: User visual approval**

Show the window to the user and confirm the layout matches the spec
before continuing.

- [ ] **Step 8: Commit**

```bash
git add main.py
git commit -m "feat(main): add ShotcreteLedgerPreviewWindow"
```

---

## Task 8: main.py — MainWindow "Books Excel" button group

**Files:**
- Modify: `main.py` — replace the existing "Go to Books Excel section"
  button construction inside `MainWindow._build_ui` (or wherever it
  lives — locate with grep).

**Concrete reference (from planning grep):** The existing button is
constructed inside `MainWindow._build_ui` (around line 2583) and looks
like:

```python
self.ledger_btn = ctk.CTkButton(
    body,
    text="Go to Books Excel section",
    height=40,
    font=ctk.CTkFont(size=13, weight="bold"),
    fg_color="#1565C0", hover_color="#0D47A1",
    command=self._on_open_ledger,
    state="disabled",
)
self.ledger_btn.pack(fill="x", pady=(12, 0))
```

The handler is `_on_open_ledger` at line 2653; it calls
`LedgerPreviewWindow(self.root, self._last_cubes_data)`. There is no
`on_close` callback used today.

- [ ] **Step 1: Replace the single ledger button with the group**

In `MainWindow._build_ui`, find the `self.ledger_btn = ctk.CTkButton(...)`
construction and its `.pack(...)`. Replace BOTH (the construction and
the pack) with this block, keeping the same parent frame `body` and
similar vertical placement:

```python
# ---- Books Excel group: Concrete + Shotcrete ----
self.books_group = ctk.CTkFrame(body, fg_color="transparent")
self.books_group.pack(fill="x", pady=(12, 0))

ctk.CTkLabel(
    self.books_group, text="Books Excel",
    font=ctk.CTkFont(family="Segoe UI", size=11),
    text_color="gray55",
).pack(anchor="w", pady=(0, 2))

books_row = ctk.CTkFrame(self.books_group, fg_color="transparent")
books_row.pack(fill="x")

self.concrete_books_btn = ctk.CTkButton(
    books_row, text="CONCRETE", height=40,
    font=ctk.CTkFont(size=13, weight="bold"),
    fg_color="#1565C0", hover_color="#0D47A1",
    command=self._on_open_ledger,
    state="disabled",
)
self.concrete_books_btn.pack(side="left", fill="x", expand=True, padx=(0, 4))

self.shotcrete_books_btn = ctk.CTkButton(
    books_row, text="SHOTCRETE", height=40,
    font=ctk.CTkFont(size=13, weight="bold"),
    fg_color="#EF6C00", hover_color="#E65100",
    command=self._on_open_shotcrete_ledger,
    state="disabled",
)
self.shotcrete_books_btn.pack(side="left", fill="x", expand=True, padx=(4, 0))

# Preserve the legacy name so existing enable/disable code keeps working.
self.ledger_btn = self.concrete_books_btn
```

(`self.ledger_btn = self.concrete_books_btn` is the cheap shim that
avoids hunting down every `self.ledger_btn.configure(...)` call
elsewhere — both names point to the same widget.)

- [ ] **Step 2: Add a sibling handler `_on_open_shotcrete_ledger`**

Just below the existing `_on_open_ledger` method (around line 2653),
add a sibling that opens the shotcrete window:

```python
def _on_open_shotcrete_ledger(self):
    if self._last_cubes_data is None:
        messagebox.showwarning(
            "Shotcrete Ledger",
            "Process a PDF and write the first Excel first, then try again.",
            parent=self.root,
        )
        return
    try:
        ShotcreteLedgerPreviewWindow(self.root, self._last_cubes_data)
    except Exception as e:
        tb = traceback.format_exc()
        reader._log(f"--- SHOT-LEDGER ERROR ---\n{tb}")
        short = str(e)
        if len(short) > 500:
            short = short[:500] + "..."
        messagebox.showerror(
            "Shotcrete Ledger Error",
            f"{short}\n\n(Full traceback saved to gemini_debug.log)",
            parent=self.root,
        )
```

- [ ] **Step 3: Enable BOTH buttons together after first write**

Find every place `self.ledger_btn.configure(state="normal")` (or
`state="disabled"`) is called. After each one, add the same call for
`self.shotcrete_books_btn`. Use grep to be exhaustive:

Run: `grep -n "ledger_btn.configure" main.py`
For each hit, follow with a matching
`self.shotcrete_books_btn.configure(state=...)` line so both buttons
toggle together.

- [ ] **Step 4: Quick syntax check**

Run: `python -m py_compile main.py`
Expected: no output (success)

- [ ] **Step 5: Run the app, smoke-test BOTH buttons**

Run: `python main.py`
Expected: MainWindow shows a "Books Excel" caption with the two
buttons side-by-side. Concrete button must open the existing
LedgerPreviewWindow exactly as before (this is the regression check).
Shotcrete button must open ShotcreteLedgerPreviewWindow (will show
"shotcrete ledger not found" unless the user has the shotcrete Excel
open — that's expected).

- [ ] **Step 6: Commit**

```bash
git add main.py
git commit -m "feat(main): replace Books button with [CONCRETE] [SHOTCRETE] group"
```

---

## Task 9: Regression test — concrete ledger end-to-end

**Files:**
- No code changes. Live test.

- [ ] **Step 1: Open a real concrete ledger Excel**

Open `Concrete sample form (ALL).xlsx` (or similar) in Excel. Confirm
the active sheet is "Concrete" and row 7 has the expected headers.

- [ ] **Step 2: Process a PDF with at least one concrete cube**

Run the app, drag a real notebook PDF into MainWindow, wait for Gemini
to finish, click "Write to Excel" on the per-sheet preview, then close.

- [ ] **Step 3: Click the new CONCRETE button**

Expected: LedgerPreviewWindow opens exactly as before — same cards,
same file selector, same write behaviour. Click ">> WRITE TO LEDGER",
verify Excel rows get filled.

If anything looks different from v1.0.10, restore main.py from
`_backup_pre_shotcrete/` or `git checkout v1.0.10 -- main.py` and
investigate.

- [ ] **Step 4: Note any regressions**

If clean, continue. If broken, fix before moving on — do not stack
shotcrete bugs on top of a regression.

---

## Task 10: Live test — shotcrete ledger end-to-end

**Files:**
- No code changes. Live test on the real shotcrete Excel.

- [ ] **Step 1: Open the real shotcrete Excel**

Open `Shotcrete sample form.xlsx` (the actual workbook). Active sheet
should be "Shotcrete Concrete Results"; row 7 should have "Core No."
in A7 and "Site Sampling Mark/No:" in B7.

- [ ] **Step 2: Process a shotcrete PDF**

Drag a notebook PDF that contains at least one shotcrete page (the
form title says "Core Record …"). Confirm the per-sheet preview shows
the SHOTCRETE pill on those cubes.

- [ ] **Step 3: Open the shotcrete ledger pass**

Close the per-sheet preview (so edits persist). Click MainWindow's
SHOTCRETE button.

Expected: ShotcreteLedgerPreviewWindow opens with the file selector
(label or dropdown), cards with 4 value columns, and SHOTCRETE pills.

- [ ] **Step 4: Inspect a card**

For a matched card, verify:
- The 4 value columns are populated with the correct readings.
- Already-filled ledger cells are shown gray; empty cells are green.
- Per-group 7-day / 28-day checkboxes are present and default to ON.

- [ ] **Step 5: Write and verify Excel**

Click ">> WRITE TO LEDGER". Confirm the success dialog. Switch to the
Excel window and verify columns L (Diameter), M (Height), N (Weight),
O (Load) are filled on the correct rows for one block.

- [ ] **Step 6: Capture any bugs and fix**

If write skips rows it shouldn't, or writes wrong columns, isolate and
fix in `write_shotcrete_ledger_cube` first.

- [ ] **Step 7: Two-set sanity check (if a two-set sample is available)**

If you have a two-set shotcrete cube in the PDF, verify it writes to
all 20 rows of its block. If the real Excel uses two distinct 10-row
blocks instead of one 20-row block, the assumption in the spec is
wrong — relax `merge_shotcrete_cubes_for_ledger` to keep sub-cubes
separate (don't merge by `(sample_key, cube_no)`).

---

## Task 11: Mirror to `dist/CubeLogReader/src/` and bump version

**Files:**
- Modify: `version.txt` (root + dist)
- Copy: `main.py`, `reader.py`, `writer.py`, `updater.py` →
  `dist/CubeLogReader/src/`

- [ ] **Step 1: Bump version.txt in both locations**

Read `version.txt`. Current value should be `1.0.10`. Bump to
`1.0.11`.

```bash
echo "1.0.11" > version.txt
echo "1.0.11" > dist/CubeLogReader/src/version.txt
```

- [ ] **Step 2: Copy edited source into dist**

```bash
cp main.py reader.py writer.py updater.py dist/CubeLogReader/src/
```

- [ ] **Step 3: Commit the version bump and mirrors**

```bash
git add version.txt main.py writer.py
git commit -m "chore: bump version to 1.0.11"
```

(`dist/` is gitignored — no add for it.)

---

## Task 12: Cut release v1.0.11

**Files:**
- Build artifact: `src.zip`
- GitHub release: tag `v1.0.11`

- [ ] **Step 1: Build src.zip + compute SHA256**

```bash
rm -f src.zip
cd dist/CubeLogReader/src && powershell.exe -NoProfile -Command "Compress-Archive -Path *.py,version.txt -DestinationPath ../../../src.zip -Force" && cd ../../..
SHA=$(powershell.exe -NoProfile -Command "(Get-FileHash src.zip -Algorithm SHA256).Hash.ToLower()" | tr -d '\r\n')
echo "SHA256: $SHA"
```

Expected: prints a 64-char hex SHA.

- [ ] **Step 2: Verify the zip contents**

```bash
python -c "import zipfile; z=zipfile.ZipFile('src.zip'); print(z.namelist()); z.close()"
```

Expected: `['main.py', 'reader.py', 'updater.py', 'writer.py', 'version.txt']`.

- [ ] **Step 3: Create the GitHub release**

Substitute the SHA from Step 1 into the notes:

```bash
"/c/Program Files/GitHub CLI/gh.exe" release create v1.0.11 src.zip \
  --title "v1.0.11 - Shotcrete Books Excel" \
  --notes "What's new:

- New SHOTCRETE button in the Books Excel group writes all 5 specimens
  per age to a separate shotcrete ledger workbook (Core Diameter, Core
  Height, Weight, Load in columns L/M/N/O).
- In-window file selector lets you pick which shotcrete ledger to write
  to when more than one is open.
- Concrete Books behaviour is unchanged.

SHA256: $SHA"
```

- [ ] **Step 4: Push branch + tag**

```bash
git push origin main
git push origin v1.0.11
```

- [ ] **Step 5: Verify release**

```bash
"/c/Program Files/GitHub CLI/gh.exe" release view v1.0.11 --json tagName,assets,url
```

Expected: `tagName: v1.0.11`, one asset named `src.zip`, digest matches
the SHA you printed in Step 1.

Hand the URL to the user.

---

## Done criteria

- [ ] All 12 tasks above complete with committed code.
- [ ] Concrete ledger flow opens and writes as in v1.0.10
  (regression-clean).
- [ ] Shotcrete ledger flow opens against the real Excel, shows correct
  cards, and writes columns L/M/N/O on real blocks.
- [ ] Release v1.0.11 published on GitHub with matching SHA256 in the
  notes.
- [ ] `_backup_pre_shotcrete/` and tag `v1.0.10` remain available as
  rollback points.
