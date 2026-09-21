"""
test_card_tabs.py — the To write / Done / No match tabs.

Cards used to sit in one list with a collapsed "N hidden" button; opening it
dropped the hidden cards back among the matched ones, so the unmatched card —
the only one you could act on — was the hardest to find.

Now every card is still built and kept (ticks, edits and the write button see
all of them), and the tabs only decide which are packed.

The widgets are stubbed: these drive CardTabs against fakes that record
pack/pack_forget.
"""
import types

import main


class _FakeCard:
    def __init__(self):
        self.packed = False

    def pack(self, **kw):
        self.packed = True

    def pack_forget(self):
        self.packed = False


class _FakeButton(_FakeCard):
    def __init__(self):
        super().__init__()
        self.cfg = {}

    def configure(self, **kw):
        self.cfg.update(kw)


class _FakeBar(_FakeCard):
    pass


def _tabs(groups, active="todo"):
    """A CardTabs host whose cards fall into `groups` (a list of tab names)."""
    w = types.SimpleNamespace()
    w._cards = [_FakeCard() for _ in groups]
    w._card_tab = active
    w._card_groups = []
    w._tab_bar = _FakeBar()
    w._tab_buttons = {name: _FakeButton() for name, _ in main.CARD_TABS}
    w._card_group = lambda i: groups[i]
    w._selected_idx = None
    w.selected = []
    w._select_card = lambda i: (w.selected.append(i),
                                setattr(w, "_selected_idx", i))
    for name in ("_apply_card_tabs", "_set_card_tab", "_visible_card_indices",
                 "_step_card"):
        setattr(w, name, types.MethodType(getattr(main.CardTabs, name), w))
    return w


def _shown(w):
    return [i for i, c in enumerate(w._cards) if c.packed]


def test_only_the_active_tab_is_packed():
    w = _tabs(["todo", "no_match", "todo", "done"])
    w._apply_card_tabs()
    assert _shown(w) == [0, 2]

    w._set_card_tab("no_match")
    assert _shown(w) == [1]

    w._set_card_tab("done")
    assert _shown(w) == [3]


def test_unmatched_cards_never_mix_with_matched_ones():
    # The whole point: no tab shows both.
    w = _tabs(["todo", "no_match", "todo"])
    for tab in ("todo", "no_match"):
        w._set_card_tab(tab)
        groups = {w._card_groups[i] for i in _shown(w)}
        assert len(groups) == 1, f"tab {tab} mixed groups: {groups}"


def test_every_card_stays_alive_on_every_tab():
    # Cards are only unpacked, never destroyed — the write button still has
    # the ticks from cards the user isn't looking at.
    w = _tabs(["todo", "no_match", "done"])
    w._apply_card_tabs()
    assert len(w._cards) == 3
    w._set_card_tab("no_match")
    assert len(w._cards) == 3


def test_an_empty_tab_is_hidden():
    w = _tabs(["todo", "todo"])
    w._apply_card_tabs()
    assert not w._tab_buttons["no_match"].packed
    assert not w._tab_buttons["done"].packed


def test_the_strip_disappears_when_there_is_only_one_group():
    w = _tabs(["todo", "todo"])
    w._apply_card_tabs()
    assert not w._tab_bar.packed, "a single group needs no tab strip"


def test_the_strip_shows_when_a_second_group_exists():
    w = _tabs(["todo", "no_match"])
    w._apply_card_tabs()
    assert w._tab_buttons["todo"].packed
    assert w._tab_buttons["no_match"].packed


def test_the_counts_are_on_the_tabs():
    w = _tabs(["todo", "todo", "no_match"])
    w._apply_card_tabs()
    assert w._tab_buttons["todo"].cfg["text"].endswith("2")
    assert w._tab_buttons["no_match"].cfg["text"].endswith("1")


def test_emptying_the_active_tab_falls_back_to_a_full_one():
    # The last unmatched card just had its mark fixed: don't leave the user
    # staring at a blank "No match" tab.
    groups = ["todo", "no_match"]
    w = _tabs(groups, active="no_match")
    w._apply_card_tabs()
    assert _shown(w) == [1]

    groups[1] = "todo"          # corrected — it now matches
    w._apply_card_tabs()
    assert w._card_tab == "todo"
    assert _shown(w) == [0, 1]


def test_arrow_keys_stay_on_the_active_tab():
    w = _tabs(["todo", "no_match", "todo"])
    w._apply_card_tabs()
    w._step_card(1)
    assert w.selected[-1] == 0
    w._step_card(1)
    assert w.selected[-1] == 2, "arrow key jumped to a card on another tab"
    w._step_card(1)
    assert w.selected[-1] == 2, "should stop at the last visible card"
    w._step_card(-1)
    assert w.selected[-1] == 0


def test_preview_window_uses_the_tabs():
    assert issubclass(main.PreviewWindow, main.CardTabs)
    assert not hasattr(main.PreviewWindow, "_toggle_hidden"), \
        "the old hidden-cards toggle should be gone"


def test_both_ledger_windows_use_the_tabs():
    for cls in (main.LedgerPreviewWindow, main.ShotcreteLedgerPreviewWindow):
        assert issubclass(cls, main.CardTabs), cls.__name__
        assert not hasattr(cls, "_toggle_notfound"), cls.__name__


def test_the_ledger_groups_follow_the_build_order():
    w = types.SimpleNamespace()
    w._card_kinds = ["todo", "done", "no_match"]
    group = types.MethodType(main.LedgerPreviewWindow._card_group, w)
    assert [group(i) for i in range(3)] == ["todo", "done", "no_match"]
    # A not-found card built before the kinds were recorded must not blow up.
    assert group(9) == "todo"


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
