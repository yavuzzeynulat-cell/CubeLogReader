"""
test_mark_fix_retry.py — retrying a Sample Mark correction.

The first scan only walks the ACTIVE workbook forward from the active sheet
(writer.scan_sheets_for_cubes), so a sample that belongs in a different file
matches nothing however right the typed mark is. The window then tells the
user to open that workbook and try again — which only works if trying again
actually does something.

It used to do nothing at all: _apply_sample_mark_fix returned early when the
typed number equalled the one already on the cube, and after the first
correction that is exactly what the cube carries. No scan, no re-match, no
message — indistinguishable from the correction being refused.

The GUI is stubbed: these drive the unbound methods against a fake `self`.
"""
import types

import main
import writer


class _FakeVar:
    def __init__(self, v=True):
        self._v = v

    def get(self):
        return self._v


def _cube(mark, cube_no="109"):
    return {"sample_mark": mark, "cube_no": cube_no, "tests": []}


def _sheet(num, name=None, wb="Cube forms.xlsx"):
    return {"workbook": wb, "sheet": name or f"S{num}",
            "sample_id_raw": num, "sample_id_num": num}


def _window(cubes, open_sheets, scan_result):
    """A stand-in for PreviewWindow carrying only what the fix path touches.

    `scan_result` is a list of sheet lists — one per expected scan, so a test
    can make the first scan come back empty and the second one find the sheet.
    """
    w = types.SimpleNamespace()
    w.cubes_data = {"cubes": cubes}
    w.open_sheets = list(open_sheets)
    w.matched = []
    w.match_error = None
    w.win = None
    w.scans = []
    w.infos = []
    w._scan_queue = list(scan_result)
    w._persist_edits_to_cubes = lambda: None
    w._populate_cards = lambda: None
    # Real methods, fake object — the point is to exercise the shipped code.
    for name in ("_merge_scanned_sheets", "_rescan_and_rematch",
                 "_after_mark_fix"):
        setattr(w, name,
                types.MethodType(getattr(main.PreviewWindow, name), w))
    return w


class _Patch:
    """Swap the Excel scan and the message boxes for recorders."""

    def __init__(self, win):
        self.win = win

    def __enter__(self):
        self._scan = writer.scan_sheets_for_cubes
        self._box = main.messagebox
        win = self.win

        def fake_scan(cubes_data, **kw):
            win.scans.append(cubes_data)
            found = win._scan_queue.pop(0) if win._scan_queue else []
            return {"sheets": list(found)}

        writer.scan_sheets_for_cubes = fake_scan
        main.messagebox = types.SimpleNamespace(
            showwarning=lambda *a, **k: win.infos.append(("warn", a)),
            showinfo=lambda *a, **k: win.infos.append(("info", a)),
            showerror=lambda *a, **k: win.infos.append(("error", a)),
            askyesno=lambda *a, **k: True,
        )
        return win

    def __exit__(self, *exc):
        writer.scan_sheets_for_cubes = self._scan
        main.messagebox = self._box
        return False


def _fix(win, old_mark, new_text):
    main.PreviewWindow._apply_sample_mark_fix(win, old_mark, new_text)


def _search_again(win, num):
    main.PreviewWindow._rescan_and_rematch(win, num, force_scan=True)


def test_a_correction_that_still_finds_no_sheet_is_kept():
    # The sample belongs to a workbook that isn't open, so the scan finds
    # nothing. The mark must still be corrected — the ledger can use it.
    win = _window([_cube("G26-CON-1798")], [], scan_result=[[]])
    with _Patch(win):
        _fix(win, "G26-CON-1798", "1198")
    assert win.cubes_data["cubes"][0]["sample_mark"] == "G26-CON-1198"
    assert win.matched[0]["matched_sheet"] is None
    assert any(kind == "info" for kind, _ in win.infos), \
        "the user should be told the sheet wasn't found"


def test_retyping_the_same_number_scans_again():
    # THE REGRESSION. First try: workbook not open, no match. The user opens
    # it and retypes the same number — that has to scan a second time.
    win = _window([_cube("G26-CON-1798")], [],
                  scan_result=[[], [_sheet(1198)]])
    with _Patch(win):
        _fix(win, "G26-CON-1798", "1198")
        assert len(win.scans) == 1
        _fix(win, "G26-CON-1198", "1198")

    assert len(win.scans) == 2, \
        f"second attempt did not rescan (scans={len(win.scans)})"
    assert win.matched[0]["matched_sheet"]["sheet"] == "S1198"


def test_search_again_matches_without_touching_the_mark():
    win = _window([_cube("G26-CON-1198")], [],
                  scan_result=[[_sheet(1198)]])
    with _Patch(win):
        _search_again(win, 1198)
    assert win.cubes_data["cubes"][0]["sample_mark"] == "G26-CON-1198"
    assert win.matched[0]["matched_sheet"]["sheet"] == "S1198"


def test_search_again_scans_even_when_the_id_is_already_known():
    # The sheet is in open_sheets but was claimed by another cube; pressing
    # the button must still go and look rather than short-circuit.
    win = _window([_cube("G26-CON-1198")], [_sheet(1198)],
                  scan_result=[[_sheet(1198, "S1198b")]])
    with _Patch(win):
        _search_again(win, 1198)
    assert len(win.scans) == 1, "forced scan was skipped"


def test_a_blank_retype_changes_nothing():
    win = _window([_cube("G26-CON-1798")], [], scan_result=[[]])
    with _Patch(win):
        _fix(win, "G26-CON-1798", "   ")
    assert win.cubes_data["cubes"][0]["sample_mark"] == "G26-CON-1798"
    assert not win.scans
    assert any(kind == "warn" for kind, _ in win.infos)


if __name__ == "__main__":
    fns = [v for k, v in sorted(globals().items()) if k.startswith("test_")]
    failed = 0
    for fn in fns:
        try:
            fn()
            print(f"PASS {fn.__name__}")
        except AssertionError as e:
            failed += 1
            print(f"FAIL {fn.__name__}: {e}")
        except Exception as e:
            failed += 1
            print(f"ERROR {fn.__name__}: {type(e).__name__}: {e}")
    print(f"\n{len(fns) - failed}/{len(fns)} passed")
    raise SystemExit(1 if failed else 0)
