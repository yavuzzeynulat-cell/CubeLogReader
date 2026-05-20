"""
writer.py — Writes cube test results into the open Excel workbook.

Target cells (per sheet):
  28-day weight -> W20, W21, W22
  28-day load   -> AA20, AA21, AA22
  7-day weight  -> W17, W18, W19  (only read, for cross-check)
  7-day load    -> AA17, AA18, AA19 (only read, for cross-check)
  Sample ID     -> B14  (used for sheet matching)
"""
import re
import sys
from datetime import datetime
from pathlib import Path

import win32com.client
from pythoncom import com_error

# Resolve log path relative to the frozen exe when bundled, else the source file.
if getattr(sys, "frozen", False):
    _BASE_DIR = Path(sys.executable).parent
else:
    _BASE_DIR = Path(__file__).parent
LOG_PATH = _BASE_DIR / "write_log.txt"


def get_log_path() -> Path:
    return LOG_PATH


SAMPLE_ID_CELL = "B14"

# Green tab color applied to a sheet after values are written.
# Excel's Tab.Color uses BGR packed as int (B<<16 | G<<8 | R).
# RGB(0, 176, 80) is Excel's standard "Green".
WRITTEN_TAB_COLOR = (80 << 16) | (176 << 8) | 0

WEIGHT_CELLS_28D = ["W20", "W21", "W22"]
LOAD_CELLS_28D = ["AA20", "AA21", "AA22"]

WEIGHT_CELLS_7D = ["W17", "W18", "W19"]
LOAD_CELLS_7D = ["AA17", "AA18", "AA19"]

# Shotcrete-only: core diameter / height (in mm) per specimen.
CORE_DIAM_CELLS_7D = ["R17", "R18", "R19"]
CORE_DIAM_CELLS_28D = ["R20", "R21", "R22"]
CORE_HEIGHT_CELLS_7D = ["V17", "V18", "V19"]
CORE_HEIGHT_CELLS_28D = ["V20", "V21", "V22"]


def normalize_sample_mark(mark) -> int | None:
    """
    G26-CON-395  -> 395
    G26-CON-0395 -> 395
    395          -> 395
    395.0 (COM float from Excel) -> 395
    None/''      -> None
    """
    if mark is None:
        return None
    # Excel often returns numeric cells as float via COM. Handle that first
    # so "395.0" doesn't get parsed as trailing "0".
    if isinstance(mark, (int, float)):
        try:
            return int(mark)
        except (TypeError, ValueError):
            return None
    s = str(mark).strip()
    if not s:
        return None
    # Pure numeric string (possibly "395.0" from str(float)) — parse as number
    try:
        return int(float(s))
    except ValueError:
        pass
    m = re.search(r"(\d+)\s*$", s)
    if m:
        try:
            return int(m.group(1))
        except ValueError:
            return None
    return None


def connect_to_excel():
    """Connect to the already-running Excel application."""
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except com_error as e:
        raise RuntimeError(
            "No open Excel instance found. Please open the Excel file "
            "first, then run the program."
        ) from e
    return excel


def list_open_sheets() -> list[dict]:
    """
    List all sheets across all open workbooks, with their Sample IDs.
    (Only used by test_pipeline.py — the main UI uses scan_sheets_for_cubes.)
    """
    excel = connect_to_excel()
    results: list[dict] = []
    for wb in excel.Workbooks:
        for ws in wb.Worksheets:
            try:
                raw = ws.Range(SAMPLE_ID_CELL).Value
            except Exception:
                raw = None
            num = normalize_sample_mark(raw)
            results.append(
                {
                    "workbook": wb.Name,
                    "sheet": ws.Name,
                    "sample_id_raw": raw,
                    "sample_id_num": num,
                }
            )
    return results


def scan_sheets_for_cubes(
    cubes_data: dict,
    max_forward: int = 25,
) -> dict:
    """
    Scan FORWARD starting from the active Excel sheet. Continue until
    every cube from the notebook has been matched (or until the
    max_forward safety cap is hit). Gaps are skipped — e.g. if the
    notebook wants 378, 379, 381, a missing 380 doesn't stop the scan.

    If the same Sample ID appears in multiple cubes (multi-set), the
    scan keeps going until enough sheets have been found.

    Returns:
      {
        "sheets":        [dict, ...],  # ready for match_cubes_to_sheets
        "scanned_count": int,          # how many sheets were checked
        "start_sheet":   str,          # name of the starting sheet
        "workbook":      str,          # active workbook name
        "found_all":     bool,         # whether every cube matched
      }
    """
    from collections import Counter

    # What are we looking for? (Sample ID -> count)
    needed: Counter = Counter()
    for cube in cubes_data.get("cubes", []):
        num = normalize_sample_mark(cube.get("sample_mark"))
        if num is not None:
            needed[num] += 1

    excel = connect_to_excel()
    active_wb = excel.ActiveWorkbook
    if active_wb is None:
        raise RuntimeError(
            "No active Excel workbook. Open a file and select the "
            "starting sheet."
        )
    active_ws = excel.ActiveSheet
    if active_ws is None:
        raise RuntimeError("No active sheet.")

    start_index = int(active_ws.Index)  # 1-based
    total_sheets = int(active_wb.Worksheets.Count)
    start_name = str(active_ws.Name)
    wb_name = str(active_wb.Name)

    sheets: list[dict] = []
    scanned = 0
    remaining = sum(needed.values())

    i = start_index
    while i <= total_sheets and scanned < max_forward and remaining > 0:
        ws = active_wb.Worksheets(i)
        try:
            raw = ws.Range(SAMPLE_ID_CELL).Value
        except Exception:
            raw = None
        num = normalize_sample_mark(raw)
        sheets.append(
            {
                "workbook": wb_name,
                "sheet": str(ws.Name),
                "sample_id_raw": raw,
                "sample_id_num": num,
            }
        )
        if num is not None and needed.get(num, 0) > 0:
            needed[num] -= 1
            remaining -= 1
        i += 1
        scanned += 1

    return {
        "sheets": sheets,
        "scanned_count": scanned,
        "start_sheet": start_name,
        "workbook": wb_name,
        "found_all": remaining == 0,
    }


def match_cubes_to_sheets(
    cubes_data: dict,
    sheets: list[dict] | None = None,
) -> list[dict]:
    """
    Match cubes from Gemini against the open sheets.
    Each cube gets either a matched_sheet (dict) or None.

    When multiple sheets share the same Sample ID, cubes are
    distributed 1-to-1 in order: first cube -> first sheet, second
    cube -> second sheet, etc. This ensures two sets of cubes that
    share a G26-CON-XXX number are written to the correct sheets.

    If `sheets` is None, it is scanned over COM; otherwise the given
    pre-scanned list is reused.
    """
    if sheets is None:
        sheets = list_open_sheets()
    used_sheet_keys: set[tuple[str, str]] = set()
    results: list[dict] = []
    for cube in cubes_data.get("cubes", []):
        num = normalize_sample_mark(cube.get("sample_mark"))
        match = None
        if num is not None:
            for s in sheets:
                if s["sample_id_num"] != num:
                    continue
                key = (s["workbook"], s["sheet"])
                if key in used_sheet_keys:
                    continue
                match = s
                used_sheet_keys.add(key)
                break
        results.append({"cube": cube, "matched_sheet": match})
    return results


def read_7day_values(workbook_name: str, sheet_name: str) -> dict:
    """Read the existing 7-day values from Excel (for cross-checking)."""
    excel = connect_to_excel()
    wb = excel.Workbooks(workbook_name)
    ws = wb.Worksheets(sheet_name)
    weights = [ws.Range(c).Value for c in WEIGHT_CELLS_7D]
    loads = [ws.Range(c).Value for c in LOAD_CELLS_7D]
    return {"weights": weights, "loads": loads}


def read_all_values(workbook_name: str, sheet_name: str) -> dict:
    """Read both 7-day and 28-day values from Excel.

    Used by the UI to visually distinguish cells that are already
    filled (skip) vs empty (needs fill).
    """
    excel = connect_to_excel()
    wb = excel.Workbooks(workbook_name)
    ws = wb.Worksheets(sheet_name)
    return {
        "weights_7d": [ws.Range(c).Value for c in WEIGHT_CELLS_7D],
        "loads_7d": [ws.Range(c).Value for c in LOAD_CELLS_7D],
        "weights_28d": [ws.Range(c).Value for c in WEIGHT_CELLS_28D],
        "loads_28d": [ws.Range(c).Value for c in LOAD_CELLS_28D],
    }


def is_cell_empty(val) -> bool:
    """True if an Excel cell value should be considered empty."""
    if val is None:
        return True
    if isinstance(val, str) and val.strip() == "":
        return True
    return False


def cross_check_7day(cube: dict, excel_7d: dict) -> list[str]:
    """
    Compare the 7-day values from the notebook with the existing
    7-day values in Excel. Returns a list of warning messages.
    """
    warnings: list[str] = []
    tests_7 = [t for t in cube.get("tests", []) if t.get("age_days") == 7]
    ex_w = excel_7d.get("weights", [None, None, None])
    ex_l = excel_7d.get("loads", [None, None, None])
    mark = cube.get("sample_mark", "?")

    for i in range(min(3, len(tests_7))):
        t = tests_7[i]
        w_def = t.get("weight_gr")
        l_def = t.get("load_kn")
        w_ex = ex_w[i] if i < len(ex_w) else None
        l_ex = ex_l[i] if i < len(ex_l) else None

        if w_ex is not None and w_def is not None:
            try:
                if abs(float(w_ex) - float(w_def)) > 0.5:
                    warnings.append(
                        f"{mark}: 7-day W{17+i} Excel={w_ex}, notebook={w_def}"
                    )
            except (TypeError, ValueError):
                pass

        if l_ex is not None and l_def is not None:
            try:
                if abs(float(l_ex) - float(l_def)) > 0.01:
                    warnings.append(
                        f"{mark}: 7-day AA{17+i} Excel={l_ex}, notebook={l_def}"
                    )
            except (TypeError, ValueError):
                pass
    return warnings


def write_cube(
    cube: dict,
    workbook_name: str,
    sheet_name: str,
    weights_7d: list,
    loads_7d: list,
    weights_28d: list,
    loads_28d: list,
    diameters_7d: list | None = None,
    heights_7d: list | None = None,
    diameters_28d: list | None = None,
    heights_28d: list | None = None,
) -> dict:
    """
    Write a cube's results into the target Excel sheet.
    Accepts 7-day and 28-day weights and loads from the UI.
    Empty / None values are skipped (existing Excel values are preserved).

    For shotcrete cubes (cube["_shotcrete"] is True), core diameter and
    height are also written to R/V columns. For normal cubes they stay None.
    """
    excel = connect_to_excel()
    wb = excel.Workbooks(workbook_name)
    ws = wb.Worksheets(sheet_name)

    wrote: list[dict] = []
    errors: list[str] = []

    def _write(cell: str, val):
        if val is None or val == "":
            return
        try:
            ws.Range(cell).Value = float(val) if isinstance(val, str) else val
            wrote.append({"cell": cell, "value": val})
        except Exception as e:
            errors.append(f"{cell}: {e}")

    is_shot = bool(cube.get("_shotcrete"))
    diameters_7d = diameters_7d or []
    heights_7d = heights_7d or []
    diameters_28d = diameters_28d or []
    heights_28d = heights_28d or []

    for i in range(3):
        _write(WEIGHT_CELLS_7D[i], weights_7d[i] if i < len(weights_7d) else None)
        _write(LOAD_CELLS_7D[i], loads_7d[i] if i < len(loads_7d) else None)
        _write(WEIGHT_CELLS_28D[i], weights_28d[i] if i < len(weights_28d) else None)
        _write(LOAD_CELLS_28D[i], loads_28d[i] if i < len(loads_28d) else None)
        if is_shot:
            _write(CORE_DIAM_CELLS_7D[i],
                   diameters_7d[i] if i < len(diameters_7d) else None)
            _write(CORE_HEIGHT_CELLS_7D[i],
                   heights_7d[i] if i < len(heights_7d) else None)
            _write(CORE_DIAM_CELLS_28D[i],
                   diameters_28d[i] if i < len(diameters_28d) else None)
            _write(CORE_HEIGHT_CELLS_28D[i],
                   heights_28d[i] if i < len(heights_28d) else None)

    if wrote:
        # Tint the sheet tab light green so the user can see at a glance
        # which sheets were filled. Shotcrete sheets are left untouched.
        if not is_shot:
            try:
                ws.Tab.Color = WRITTEN_TAB_COLOR
            except Exception as e:
                print(f"tab color failed: {e}", file=sys.stderr)
        try:
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            mark = cube.get("sample_mark", "?")
            cube_no = cube.get("cube_no", "?")
            set_idx = cube.get("_set_index")
            set_suffix = ""
            if set_idx:
                try:
                    n, m = set_idx
                    set_suffix = f" (set {n}/{m})"
                except Exception:
                    set_suffix = f" (set {set_idx})"
            lines = [
                f"[{ts}] {mark} (cube {cube_no}){set_suffix} -> {workbook_name} / {sheet_name}"
            ]
            for entry in wrote:
                lines.append(f"  {entry['cell']} = {entry['value']}")
            with open(LOG_PATH, "a", encoding="utf-8") as f:
                f.write("\n".join(lines) + "\n")
        except Exception as e:
            print(f"write_log failed: {e}", file=sys.stderr)

    return {
        "wrote": wrote,
        "errors": errors,
        "workbook": workbook_name,
        "sheet": sheet_name,
    }


# Backwards-compatible alias — test_pipeline.py uses the old name
def write_cube_28day(
    cube: dict,
    workbook_name: str,
    sheet_name: str,
    weights: list,
    loads: list,
) -> dict:
    """Legacy: write only 28-day values."""
    return write_cube(
        cube,
        workbook_name,
        sheet_name,
        weights_7d=[],
        loads_7d=[],
        weights_28d=weights,
        loads_28d=loads,
    )


# =====================================================================
# Ledger (downward-growing "book" Excel) support
# =====================================================================

LEDGER_SHEET_NAME = "Concrete"
LEDGER_MAX_SCAN_ROWS = 30000  # hard cap; real end is UsedRange


def _is_empty(v) -> bool:
    return v is None or (isinstance(v, str) and v.strip() == "")


def normalize_cube_no(val) -> int | None:
    """Cube no → int. None / '' / parse error → None."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        try:
            return int(val)
        except (TypeError, ValueError):
            return None
    s = str(val).strip()
    if not s:
        return None
    try:
        return int(float(s))
    except ValueError:
        return None


def ledger_sample_key(mark) -> str | None:
    """
    Prefix-preserving canonical key for ledger matching.
      G26-CON-522   -> "G26-CON-522"
      G26-CON-0522  -> "G26-CON-522"   (leading zeros stripped)
      G-CON-0522    -> "G-CON-522"     (different prefix → different key)
      522 / 522.0   -> "522"
      None / ''     -> None
    """
    if mark is None:
        return None
    if isinstance(mark, (int, float)):
        try:
            return str(int(mark))
        except (TypeError, ValueError):
            return None
    s = str(mark).strip()
    if not s:
        return None
    try:
        return str(int(float(s)))
    except ValueError:
        pass
    m = re.search(r"(\d+)\s*$", s)
    if not m:
        return None
    num = int(m.group(1))
    prefix = s[: m.start()].rstrip("- ").upper()
    # Known Gemini OCR misread: handwritten "G" at the start of a prefix is
    # often read as the digit "6", producing "626-CON-NNN" where the real
    # mark is "G26-CON-NNN". Normalize so the ledger match still works.
    if prefix.startswith("626-") or prefix == "626":
        prefix = "G26" + prefix[3:]
    return f"{prefix}-{num}" if prefix else str(num)


def find_ledger_candidates():
    """
    Locate every open workbook that holds a ledger sheet ("Concrete").
    Validation: A7 mentions "Cube No" and B7 mentions "Sampling Mark".
    Returns a list of (workbook_com, worksheet_com, workbook_name, sheet_name).
    Raises RuntimeError only when there are 0 candidates.
    """
    excel = connect_to_excel()
    candidates = []
    for wb in excel.Workbooks:
        for ws in wb.Worksheets:
            if str(ws.Name).strip().lower() != LEDGER_SHEET_NAME.lower():
                continue
            try:
                a7 = ws.Range("A7").Value
                b7 = ws.Range("B7").Value
            except Exception:
                continue
            a7s = str(a7 or "").lower()
            b7s = str(b7 or "").lower()
            if "cube no" in a7s and "sampling mark" in b7s:
                candidates.append((wb, ws, str(wb.Name), str(ws.Name)))
    if not candidates:
        raise RuntimeError(
            "Ledger not found: no open Excel workbook has a 'Concrete' "
            "sheet with 'Cube No.' / 'Site Sampling Mark/No:' headers "
            "on row 7. Open the ledger file first."
        )
    return candidates


def find_ledger_sheet():
    """
    Locate the single ledger sheet ("Concrete").
    Returns (workbook_com, worksheet_com, workbook_name, sheet_name).
    Raises RuntimeError on 0 / >1 candidates. Retained for callers that
    require exactly one ledger; the UI uses find_ledger_candidates().
    """
    candidates = find_ledger_candidates()
    if len(candidates) > 1:
        names = ", ".join(f"{w}/{s}" for _, _, w, s in candidates)
        raise RuntimeError(f"Multiple ledger sheets found: {names}")
    return candidates[0]


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


def read_ledger_blocks(ws) -> list[dict]:
    """
    Scan from row 8 downward; identify blocks by B-filled rows.
    One batched COM read of columns A:L from row 8 to UsedRange end.
    Block size is trimmed to the last row where L (Age) is non-empty —
    so empty trailing rows between samples don't inflate the block.
    Returns [{sample_key, sample_id_num, sample_mark_raw, cube_no,
              start_row, end_row, size}, ...].
    """
    try:
        used = ws.UsedRange
        last_row = int(used.Row) + int(used.Rows.Count) - 1
    except Exception:
        last_row = LEDGER_MAX_SCAN_ROWS
    last_row = min(last_row, LEDGER_MAX_SCAN_ROWS)
    if last_row < 8:
        return []

    # A=0, B=1, C=2, ..., L=11
    data = ws.Range(f"A8:L{last_row}").Value
    if data is None:
        return []
    if not isinstance(data[0], tuple):
        data = (data,)

    # Trim trailing fully-empty rows (any column non-empty keeps it)
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
        # Provisional end = row before next head (or last data row).
        prov_end_idx = (
            head_indices[j + 1] - 1 if j + 1 < len(head_indices)
            else len(data) - 1
        )
        # Walk rows; collect 7d and 28d row numbers from column L (Age).
        # Empty L rows are ignored. End = last L-filled row.
        actual_end_idx = hi
        rows_7d: list[int] = []
        rows_28d: list[int] = []
        for k in range(hi, prov_end_idx + 1):
            l_val = data[k][11]
            if _is_empty(l_val):
                continue
            actual_end_idx = k
            try:
                age = int(float(l_val))
            except (TypeError, ValueError):
                continue
            row_no = 8 + k
            if age == 7:
                rows_7d.append(row_no)
            elif age == 28:
                rows_28d.append(row_no)
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


def read_shotcrete_ledger_blocks(ws) -> list[dict]:
    """
    Shotcrete sibling of read_ledger_blocks. Scans columns A:K from row 8
    downward; identifies blocks by B-filled rows; collects 7-day and
    28-day row numbers from column K (Age). Empty K rows ("WP", blank,
    or other text) are ignored. Block size = 5x7d + 5x28d for a single
    set, 10x7d + 10x28d for a two-set sample.
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
                # "WP" or other text — skip but still extend block end
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


def read_ledger_values(ws, blocks: list[dict]) -> dict:
    """
    Read current M (weight) and N (load) values for the rows covered by
    `blocks`. One batched COM read. Returns {block_index: {"weights":[...], "loads":[...]}}.
    """
    if not blocks:
        return {}
    first = blocks[0]["start_row"]
    last = blocks[-1]["end_row"]
    data = ws.Range(f"M{first}:N{last}").Value
    if data is None:
        return {i: {"weights": [None] * b["size"], "loads": [None] * b["size"]}
                for i, b in enumerate(blocks)}
    if not isinstance(data[0], tuple):
        data = (data,)
    result: dict = {}
    for i, b in enumerate(blocks):
        off = b["start_row"] - first
        n = b["size"]
        rows = data[off: off + n]
        result[i] = {
            "weights": [r[0] for r in rows],
            "loads": [r[1] for r in rows],
        }
    return result


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


def merge_cubes_for_ledger(cubes_data: dict) -> list[dict]:
    """
    Reshape the post-processed `cubes_data` for ledger writing:
      - Drop shotcrete cubes (they go to a separate Excel later).
      - Merge multi-set sub-cubes back by sample_id_num.
      - Concatenate _selected tests in _set_index order (set 1 first, set 2, ...).

    Returns a list of {sample_id_num, sample_mark, cube_no, tests_7d, tests_28d}
    — one entry per distinct sample_mark.
    """
    by_num: dict = {}
    order: list = []
    for cube in cubes_data.get("cubes", []):
        if cube.get("_shotcrete"):
            continue
        # Card unticked in the first PreviewWindow — exclude from the ledger.
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
            if t.get("_selected") is False:
                continue
            age = t.get("age_days")
            if age == 7:
                t7.append(t)
            elif age == 28:
                t28.append(t)

    result: list[dict] = []
    for num in order:
        entry = by_num[num]
        sets = entry.pop("_sets")
        t7_all, t28_all = [], []
        for sidx in sorted(sets.keys()):
            t7, t28 = sets[sidx]
            t7_all.extend(t7)
            t28_all.extend(t28)
        entry["tests_7d"] = t7_all
        entry["tests_28d"] = t28_all
        result.append(entry)
    return result


def merge_shotcrete_cubes_for_ledger(cubes_data: dict) -> list[dict]:
    """
    Shotcrete sibling of merge_cubes_for_ledger.
      1. Keep ONLY cubes with _shotcrete=True (concrete dropped them).
      2. Do NOT skip tests where _selected is False - all 5 specimens
         per age are forwarded so the ledger receives every result.
      3. Do NOT skip cubes where _card_enabled is False — for shotcrete
         the master tick on the first PreviewWindow only controls the
         per-sheet write. The ledger is the user's only path to record
         shotcrete data, so it must include those cubes too. The
         in-window "Write this sample" checkbox here is the real gate.
    Carry-over: merge multi-set sub-cubes back by (sample_key, cube_no);
    concatenate tests in _set_index order.

    Returns a list of {sample_key, sample_id_num, sample_mark, cube_no,
    tests_7d, tests_28d}.
    """
    by_num: dict = {}
    order: list = []
    for cube in cubes_data.get("cubes", []):
        if not cube.get("_shotcrete"):
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
        for sidx in sorted(sets.keys()):
            t7, t28 = sets[sidx]
            tests_7d.extend(t7)
            tests_28d.extend(t28)
        entry["tests_7d"] = tests_7d
        entry["tests_28d"] = tests_28d
        out.append(entry)
    return out


def match_cubes_to_blocks(
    merged_cubes: list[dict], blocks: list[dict]
) -> list[dict]:
    """
    Match merged ledger cubes to ledger blocks by sample_id_num.
    Returns [{cube, block, mismatch_reason}] — one entry per merged cube.
    mismatch_reason is None (OK), "not_found", or a human-readable string.
    """
    by_pair: dict = {}      # (sample_key, cube_no) → block (strict)
    by_key_only: dict = {}  # sample_key → first block (fallback)
    for b in blocks:
        key = b.get("sample_key")
        if key is None:
            continue
        cn = b.get("cube_no")
        pair = (key, cn)
        if pair not in by_pair:
            by_pair[pair] = b
        if key not in by_key_only:
            by_key_only[key] = b

    results: list[dict] = []
    for cube in merged_cubes:
        key = cube.get("sample_key")
        cn = cube.get("cube_no")
        block = by_pair.get((key, cn))
        # Fallback: if no exact (key, cube_no) match but key matches uniquely,
        # accept it but flag as warning so user notices.
        cube_no_warn = None
        if block is None:
            block = by_key_only.get(key)
            if block is not None and cn is not None and block.get("cube_no") is not None:
                cube_no_warn = (
                    f"Cube no mismatch: PDF={cn}, "
                    f"ledger={block.get('cube_no')}"
                )
        reason: str | None = None
        if block is None:
            reason = "not_found"
        else:
            slots_7d = len(block.get("rows_7d", []))
            slots_28d = len(block.get("rows_28d", []))
            n7 = len(cube.get("tests_7d", []))
            n28 = len(cube.get("tests_28d", []))
            # Allow partial: it's fine to have fewer tests than slots
            # (e.g. only 7d done so far). Only flag if data exceeds slots.
            if n7 > slots_7d or n28 > slots_28d:
                reason = (
                    f"Data doesn't fit block: block has {slots_7d}x7d + {slots_28d}x28d slots, "
                    f"data has {n7}x7d + {n28}x28d"
                )
        if reason is None and cube_no_warn:
            reason = cube_no_warn
        results.append({"cube": cube, "block": block, "mismatch_reason": reason})
    return results


def write_ledger_cube(
    ws, cube: dict, block: dict,
    write_7d: bool = True, write_28d: bool = True,
) -> dict:
    """
    Write M (weight) and N (load) into rows the block actually maps to.
    Block carries `rows_7d` / `rows_28d` (Excel row numbers per age, from
    column L). 7d tests fill the 7d rows in order; 28d tests fill the 28d
    rows in order. Already-filled cells are skipped silently.
    Flags `write_7d` / `write_28d` skip an entire age group if False.
    """
    wrote: list[dict] = []
    errors: list[str] = []
    skipped: list[str] = []

    rows_7d = block.get("rows_7d", [])
    rows_28d = block.get("rows_28d", [])
    start = block["start_row"]

    def _write(cell: str, val):
        if val is None or val == "":
            return
        try:
            cur = ws.Range(cell).Value
            if not _is_empty(cur):
                skipped.append(cell)
                return
            ws.Range(cell).Value = float(val) if isinstance(val, str) else val
            wrote.append({"cell": cell, "value": val})
        except Exception as e:
            errors.append(f"{cell}: {e}")

    if write_7d:
        for t, row in zip(cube.get("tests_7d", []), rows_7d):
            _write(f"M{row}", t.get("weight_gr"))
            _write(f"N{row}", t.get("load_kn"))
    if write_28d:
        for t, row in zip(cube.get("tests_28d", []), rows_28d):
            _write(f"M{row}", t.get("weight_gr"))
            _write(f"N{row}", t.get("load_kn"))

    if wrote:
        try:
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            mark = cube.get("sample_mark", "?")
            lines = [f"[{ts}] [LEDGER] {mark} -> rows {start}-{block['end_row']}"]
            for entry in wrote:
                lines.append(f"  {entry['cell']} = {entry['value']}")
            with open(LOG_PATH, "a", encoding="utf-8") as f:
                f.write("\n".join(lines) + "\n")
        except Exception as e:
            print(f"ledger write_log failed: {e}", file=sys.stderr)

    return {"wrote": wrote, "errors": errors, "skipped": skipped}


def write_shotcrete_ledger_cube(
    ws, cube: dict, block: dict,
    write_7d: bool = True, write_28d: bool = True,
) -> dict:
    """
    Shotcrete sibling of write_ledger_cube. For each specimen, writes
    four values per row:
      L = Core Diameter, M = Core Height, N = Weight, O = Load.

    Block carries `rows_7d` / `rows_28d` (Excel row numbers per age,
    from column K). 7d tests fill 7d rows in order; 28d tests fill 28d
    rows in order. Already-filled cells are skipped silently (no
    overwrite, identical policy to write_ledger_cube). None source
    values are also skipped (so missing readings don't blank existing
    data). Flags `write_7d` / `write_28d` skip an entire age group if
    False.
    """
    wrote: list[dict] = []
    errors: list[str] = []
    skipped: list[str] = []

    rows_7d = block.get("rows_7d", [])
    rows_28d = block.get("rows_28d", [])
    start = block["start_row"]

    def _write(cell: str, val):
        if val is None or val == "":
            return
        try:
            cur = ws.Range(cell).Value
            if not _is_empty(cur):
                skipped.append(cell)
                return
            ws.Range(cell).Value = float(val) if isinstance(val, str) else val
            wrote.append({"cell": cell, "value": val})
        except Exception as e:
            errors.append(f"{cell}: {e}")

    def _walk(tests, rows):
        for t, row in zip(tests, rows):
            _write(f"L{row}", t.get("core_diameter_mm"))
            _write(f"M{row}", t.get("core_height_mm"))
            _write(f"N{row}", t.get("weight_gr"))
            _write(f"O{row}", t.get("load_kn"))

    if write_7d:
        _walk(cube.get("tests_7d", []), rows_7d)
    if write_28d:
        _walk(cube.get("tests_28d", []), rows_28d)

    if wrote:
        try:
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            mark = cube.get("sample_mark", "?")
            lines = [
                f"[{ts}] [SHOT-LEDGER] {mark} -> rows {start}-{block['end_row']}"
            ]
            for entry in wrote:
                lines.append(f"  {entry['cell']} = {entry['value']}")
            with open(LOG_PATH, "a", encoding="utf-8") as f:
                f.write("\n".join(lines) + "\n")
        except Exception as e:
            print(f"shotcrete ledger write_log failed: {e}", file=sys.stderr)

    return {"wrote": wrote, "errors": errors, "skipped": skipped}


if __name__ == "__main__":
    # Quick test: list open sheets
    print("Open sheets:")
    for s in list_open_sheets():
        print(f"  {s['workbook']} / {s['sheet']}  -> Sample ID: "
              f"{s['sample_id_raw']!r}  (num={s['sample_id_num']})")
