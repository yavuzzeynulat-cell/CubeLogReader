"""
test_ledger_file_picker.py — the ledger file dropdown.

Both ledger windows scan every open workbook and let the user pick which one
to write to. The picker used to collapse to a plain "Yazilacak dosya: X"
label whenever exactly one workbook was recognised — which is the normal
case, one concrete ledger and one shotcrete ledger open. The user then had
no list at all, and no way to see that their other open workbooks simply
weren't detected.

Widgets are stubbed; these check what _build_selector builds.
"""
import types

import main


class _FakeWidget:
    def __init__(self, master=None, **kw):
        self.kw = kw
        self.packed = False
        self.value = None

    def pack(self, **kw):
        self.packed = True

    def set(self, v):
        self.value = v


class _Recorder:
    """Stand-in for the ctk module inside _build_selector."""

    def __init__(self):
        self.frames, self.labels, self.menus = [], [], []

    def CTkFrame(self, master=None, **kw):
        w = _FakeWidget(master, **kw)
        self.frames.append(w)
        return w

    def CTkLabel(self, master=None, **kw):
        w = _FakeWidget(master, **kw)
        self.labels.append(w)
        return w

    def CTkOptionMenu(self, master=None, **kw):
        w = _FakeWidget(master, **kw)
        self.menus.append(w)
        return w

    def CTkFont(self, **kw):
        return object()


def _build(cls, n_files):
    rec = _Recorder()
    w = types.SimpleNamespace()
    w.win = object()
    w._candidates = [object() for _ in range(n_files)]
    w._cand_labels = [f"ledger{i}.xlsx" for i in range(n_files)]
    w._cand_index = 0
    w._ledger_menu = None
    w._on_select_ledger = lambda choice: None
    real = main.ctk
    main.ctk = rec
    try:
        types.MethodType(cls._build_selector, w)()
    finally:
        main.ctk = real
    return w, rec


def test_one_file_still_gets_a_dropdown():
    for cls in (main.LedgerPreviewWindow, main.ShotcreteLedgerPreviewWindow):
        w, rec = _build(cls, 1)
        assert rec.menus, f"{cls.__name__}: single file got no dropdown"
        assert w._ledger_menu is not None, cls.__name__
        assert rec.menus[0].packed, cls.__name__
        assert rec.menus[0].value == "ledger0.xlsx", cls.__name__


def test_several_files_are_all_listed():
    for cls in (main.LedgerPreviewWindow, main.ShotcreteLedgerPreviewWindow):
        w, rec = _build(cls, 3)
        assert rec.menus[0].kw["values"] == \
            ["ledger0.xlsx", "ledger1.xlsx", "ledger2.xlsx"], cls.__name__


def test_the_count_is_shown():
    for cls in (main.LedgerPreviewWindow, main.ShotcreteLedgerPreviewWindow):
        w, rec = _build(cls, 2)
        texts = " | ".join(str(l.kw.get("text", "")) for l in rec.labels)
        assert "2 dosya bulundu" in texts, f"{cls.__name__}: {texts}"


def test_no_files_builds_no_picker():
    for cls in (main.LedgerPreviewWindow, main.ShotcreteLedgerPreviewWindow):
        w, rec = _build(cls, 0)
        assert not rec.menus, cls.__name__
        assert w._ledger_menu is None, cls.__name__


def test_the_selection_is_wired_to_the_switch_handler():
    for cls in (main.LedgerPreviewWindow, main.ShotcreteLedgerPreviewWindow):
        w, rec = _build(cls, 2)
        assert rec.menus[0].kw["command"] == w._on_select_ledger, cls.__name__


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
