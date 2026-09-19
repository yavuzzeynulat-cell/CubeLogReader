"""
test_sample_id_fix.py — correcting a Sample ID that Gemini misread.

When the handwritten Sample Mark is misread the cube matches no sheet, and the
fix used to be to edit the Sample ID in Excel until it matched the bad read —
changing the record to fit the mistake. Instead the user retypes the mark on
the unmatched card.

A page can hold two sets of the same sample, so both cards carry the same bad
mark; correcting one corrects them all.

Only the mark rewrite is covered here. Re-matching afterwards is just another
call to match_cubes_to_sheets, which test_sheet_demand.py covers.
"""
import writer


def _cube(mark, cube_no="109"):
    return {"sample_mark": mark, "cube_no": cube_no, "tests": []}


def test_the_typed_mark_replaces_the_misread_one():
    data = {"cubes": [_cube("G26-CON-1798")]}
    writer.apply_corrected_mark(data, "G26-CON-1798", "G26-CON-1198")
    assert data["cubes"][0]["sample_mark"] == "G26-CON-1198"


def test_both_sets_of_the_same_sample_are_corrected():
    data = {"cubes": [_cube("G26-CON-1798", "109"),
                      _cube("G26-CON-1798", "109")]}
    n = writer.apply_corrected_mark(data, "G26-CON-1798", "G26-CON-1198")
    assert n == 2
    assert [c["sample_mark"] for c in data["cubes"]] == \
        ["G26-CON-1198", "G26-CON-1198"]


def test_other_samples_are_left_alone():
    data = {"cubes": [_cube("G26-CON-1798"), _cube("G26-CON-1200")]}
    writer.apply_corrected_mark(data, "G26-CON-1798", "G26-CON-1198")
    assert data["cubes"][1]["sample_mark"] == "G26-CON-1200"


def test_a_bare_number_keeps_the_original_prefix():
    data = {"cubes": [_cube("G26-CON-1798")]}
    writer.apply_corrected_mark(data, "G26-CON-1798", "1198")
    assert data["cubes"][0]["sample_mark"] == "G26-CON-1198"


def test_matching_is_by_number_not_by_spelling():
    # The card shows "G26-CON-1798"; a zero-padded write of the same number
    # is the same sample.
    data = {"cubes": [_cube("G26-CON-01798")]}
    n = writer.apply_corrected_mark(data, "G26-CON-1798", "1198")
    assert n == 1
    assert data["cubes"][0]["sample_mark"] == "G26-CON-1198"


def test_nothing_matching_changes_nothing():
    data = {"cubes": [_cube("G26-CON-1200")]}
    n = writer.apply_corrected_mark(data, "G26-CON-1798", "1198")
    assert n == 0
    assert data["cubes"][0]["sample_mark"] == "G26-CON-1200"


def test_an_unreadable_new_mark_is_refused():
    data = {"cubes": [_cube("G26-CON-1798")]}
    try:
        writer.apply_corrected_mark(data, "G26-CON-1798", "   ")
    except ValueError:
        assert data["cubes"][0]["sample_mark"] == "G26-CON-1798"
    else:
        raise AssertionError("expected ValueError for a blank mark")


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
