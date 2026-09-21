"""
test_ledger_mark_fix.py — correcting a misread Sample Mark from the ledger.

The preview window could already retype a misread mark, but a cube whose
Excel sheet lives in a workbook that isn't open never matches there, so the
user's route is the ledger: correct the mark, let the ledger find the block,
and copy the values on from there. That means the ledger needs the same
entry, and a ledger match is keyed on the prefix too (ledger_sample_key), so
the block is only found by reading the ledger again.

The GUI is stubbed: these drive the real methods against a fake `self`.
"""
import types

import main
import writer


def _cube(mark, cube_no=109):
    return {"sample_mark": mark, "cube_no": cube_no,
            "tests": [{"age_days": 28, "weight_gr": 8300, "load_kn": 1200}]}


def _window(cubes):
    w = types.SimpleNamespace()
    w.cubes_data = {"cubes": cubes}
    w.entries = []
    w.not_found = list(cubes)
    w.ledger_error = None
    w.win = None
    w._candidates = ["ledger-A"]
    w._cand_index = 0
    w.infos = []
    w.loads = []
    w.refreshes = []

    def fake_load(candidate):
        """Stand-in for _load_ledger: re-derives entries from cubes_data the
        way the real one does, so a corrected mark changes the outcome."""
        w.loads.append(candidate)
        if w.load_fails:
            w.ledger_error = "workbook closed"
            return False
        merged = writer.merge_cubes_for_ledger(w.cubes_data)
        w.entries = [{"cube": c, "mismatch": None}
                     for c in merged
                     if c.get("sample_key") in w.ledger_keys]
        w.not_found = [c for c in merged
                       if c.get("sample_key") not in w.ledger_keys]
        return True

    w.load_fails = False
    w.ledger_keys = set()
    w._load_ledger = fake_load
    w._refresh_view = lambda: w.refreshes.append(True)
    w._after_mark_fix = types.MethodType(
        main.LedgerMarkEditor._after_mark_fix, w)
    w._apply_sample_mark_fix = types.MethodType(
        main.LedgerMarkEditor._apply_sample_mark_fix, w)
    return w


class _Patch:
    def __init__(self, win):
        self.win = win

    def __enter__(self):
        self._box = main.messagebox
        win = self.win
        main.messagebox = types.SimpleNamespace(
            showwarning=lambda *a, **k: win.infos.append(("warn", a)),
            showinfo=lambda *a, **k: win.infos.append(("info", a)),
            showerror=lambda *a, **k: win.infos.append(("error", a)),
        )
        return win

    def __exit__(self, *exc):
        main.messagebox = self._box
        return False


def test_correcting_from_the_ledger_finds_the_block():
    win = _window([_cube("G26-CON-1798")])
    win.ledger_keys = {"G26-CON-1198"}
    with _Patch(win):
        win._apply_sample_mark_fix("G26-CON-1798", "1198")

    assert win.cubes_data["cubes"][0]["sample_mark"] == "G26-CON-1198"
    assert win.loads, "the ledger was not re-read"
    assert win.refreshes, "the cards were not redrawn"
    assert [e["cube"]["sample_key"] for e in win.entries] == ["G26-CON-1198"]
    assert not win.not_found
    assert not win.infos, f"unexpected message: {win.infos}"


def test_a_wrong_prefix_still_reports_not_found_but_keeps_the_mark():
    # The ledger holds G-CON-1198; typing the number keeps the G26-CON prefix,
    # which is a different sample as far as ledger_sample_key is concerned.
    win = _window([_cube("G26-CON-1798")])
    win.ledger_keys = {"G-CON-1198"}
    with _Patch(win):
        win._apply_sample_mark_fix("G26-CON-1798", "1198")

    assert win.cubes_data["cubes"][0]["sample_mark"] == "G26-CON-1198"
    assert any(kind == "info" for kind, _ in win.infos)


def test_typing_the_full_mark_fixes_the_prefix_too():
    win = _window([_cube("G26-CON-1798")])
    win.ledger_keys = {"G-CON-1198"}
    with _Patch(win):
        win._apply_sample_mark_fix("G26-CON-1798", "G-CON-1198")

    assert win.cubes_data["cubes"][0]["sample_mark"] == "G-CON-1198"
    assert [e["cube"]["sample_key"] for e in win.entries] == ["G-CON-1198"]
    assert not win.infos


def test_retyping_the_same_number_reads_the_ledger_again():
    # Same regression as the preview window: the second attempt must act.
    win = _window([_cube("G26-CON-1798")])
    win.ledger_keys = set()
    with _Patch(win):
        win._apply_sample_mark_fix("G26-CON-1798", "1198")
        assert len(win.loads) == 1
        win.ledger_keys = {"G26-CON-1198"}  # user fixed the ledger row
        win._apply_sample_mark_fix("G26-CON-1198", "1198")

    assert len(win.loads) == 2, f"second attempt did not reload ({win.loads})"
    assert [e["cube"]["sample_key"] for e in win.entries] == ["G26-CON-1198"]


def test_both_sets_of_one_sample_are_corrected_together():
    win = _window([_cube("G26-CON-1798", 109), _cube("G26-CON-1798", 110)])
    win.ledger_keys = {"G26-CON-1198"}
    with _Patch(win):
        win._apply_sample_mark_fix("G26-CON-1798", "1198")

    assert [c["sample_mark"] for c in win.cubes_data["cubes"]] == \
        ["G26-CON-1198", "G26-CON-1198"]


def test_a_failed_ledger_reread_warns_but_keeps_the_correction():
    win = _window([_cube("G26-CON-1798")])
    win.load_fails = True
    with _Patch(win):
        win._apply_sample_mark_fix("G26-CON-1798", "1198")

    assert win.cubes_data["cubes"][0]["sample_mark"] == "G26-CON-1198"
    assert any(kind == "warn" for kind, _ in win.infos)
    assert not win.refreshes


def test_an_unreadable_mark_is_refused_without_touching_the_ledger():
    win = _window([_cube("G26-CON-1798")])
    with _Patch(win):
        win._apply_sample_mark_fix("G26-CON-1798", "  ")

    assert win.cubes_data["cubes"][0]["sample_mark"] == "G26-CON-1798"
    assert not win.loads
    assert any(kind == "warn" for kind, _ in win.infos)


def test_the_correction_is_saved_to_disk():
    # The fix has to outlive the window: the Gemini read is cached, so a
    # re-read would otherwise replay the misread mark.
    saved = []
    win = _window([_cube("G26-CON-1798")])
    win.cubes_data["_source_digest"] = "abc"
    win.ledger_keys = {"G26-CON-1198"}
    real = main.reader.save_mark_fixes
    main.reader.save_mark_fixes = lambda d: saved.append(d) or 1
    try:
        with _Patch(win):
            win._apply_sample_mark_fix("G26-CON-1798", "1198")
    finally:
        main.reader.save_mark_fixes = real
    assert saved, "the correction was never persisted"
    assert saved[0] is win.cubes_data


def test_a_correction_that_finds_nothing_is_still_saved():
    saved = []
    win = _window([_cube("G26-CON-1798")])
    win.cubes_data["_source_digest"] = "abc"
    win.ledger_keys = set()          # nothing in the ledger matches
    real = main.reader.save_mark_fixes
    main.reader.save_mark_fixes = lambda d: saved.append(d) or 1
    try:
        with _Patch(win):
            win._apply_sample_mark_fix("G26-CON-1798", "1198")
    finally:
        main.reader.save_mark_fixes = real
    assert saved, "an unmatched correction must still be persisted"


def test_both_ledger_windows_expose_the_editor():
    for cls in (main.LedgerPreviewWindow, main.ShotcreteLedgerPreviewWindow):
        assert issubclass(cls, main.LedgerMarkEditor), cls.__name__
        assert hasattr(cls, "_build_notfound_title"), cls.__name__


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
