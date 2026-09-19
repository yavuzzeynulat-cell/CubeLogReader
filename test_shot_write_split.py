"""
test_shot_write_split.py — splitting a shotcrete card's ticked rows across
sheets.

One card now shows a whole age block in a single list: the core rows and any
appended cube rows together. An Excel sheet holds 3 specimen slots per age, so
the ticked rows are handed out in row order, three at a time — first three to
the matched sheet, next three to the sheet after it.
"""
import main


class _Var:
    """Stands in for a tkinter BooleanVar."""

    def __init__(self, value):
        self._value = value

    def get(self):
        return self._value


def _rows(*flags):
    return [{"sel": _Var(f), "tag": i} for i, f in enumerate(flags)]


def _tags(chunks):
    return [[r["tag"] for r in c] for c in chunks]


def test_three_ticked_rows_make_one_chunk():
    chunks = main._chunk_selected_rows(_rows(True, True, True, False, False))
    assert _tags(chunks) == [[0, 1, 2]]


def test_six_ticked_rows_split_into_two_chunks():
    chunks = main._chunk_selected_rows(
        _rows(True, True, True, True, True, True))
    assert _tags(chunks) == [[0, 1, 2], [3, 4, 5]]


def test_the_real_1198_pattern_puts_cores_first_cubes_second():
    # 5 cores then 3 cubes; the three strongest cores and every cube ticked.
    chunks = main._chunk_selected_rows(
        _rows(True, True, True, False, False, True, True, True))
    assert _tags(chunks) == [[0, 1, 2], [5, 6, 7]]


def test_a_short_final_chunk_is_kept():
    chunks = main._chunk_selected_rows(_rows(True, True, True, True))
    assert _tags(chunks) == [[0, 1, 2], [3]]


def test_unticked_rows_are_dropped():
    chunks = main._chunk_selected_rows(_rows(False, True, False, True))
    assert _tags(chunks) == [[1, 3]]


def test_nothing_ticked_gives_no_chunks():
    assert main._chunk_selected_rows(_rows(False, False)) == []


def test_no_rows_at_all_gives_no_chunks():
    assert main._chunk_selected_rows([]) == []


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
