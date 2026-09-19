"""
test_sheet_demand.py — how many sheets the forward scan must look for.

scan_sheets_for_cubes walks forward from the active sheet and stops the
moment every cube has been matched. It counted one sheet per cube, which is
wrong for a core page whose 28-day block carries an appended cube set: those
cube rows go to a SECOND sheet sharing the same Sample ID, and the scan used
to stop before ever reaching it.

Only the demand calculation is covered here — the walk itself needs live
Excel (COM) and is exercised by hand.
"""
import writer


def _core(age, w, load, s):
    return {"age_days": age, "mould_no": None, "weight_gr": w,
            "load_kn": load, "strength_nmm2": s,
            "core_diameter_mm": 94.00, "core_height_mm": 96.00}


def _cube(age, w, load, s, mould):
    return {"age_days": age, "mould_no": mould, "weight_gr": w,
            "load_kn": load, "strength_nmm2": s,
            "core_diameter_mm": None, "core_height_mm": None}


CORES_28 = [_core(28, 1552, 239.11, 34.46), _core(28, 1554, 268.56, 38.70),
            _core(28, 1558, 266.12, 38.35), _core(28, 1556, 229.39, 33.05),
            _core(28, 1549, 232.97, 33.57)]
CUBES_28 = [_cube(28, 7869, 941.73, 41.85, "72"),
            _cube(28, 7789, 932.03, 41.42, "145"),
            _cube(28, 7708, 965.66, 42.92, "129")]


def test_plain_concrete_cube_asks_for_one_sheet():
    data = {"cubes": [{"sample_mark": "G26-CON-395",
                       "tests": [_cube(7, 8360, 1102.34, 48.99, "55")]}]}
    assert writer.needed_sheet_counts(data) == {395: 1}


def test_core_page_with_an_appended_cube_set_asks_for_two():
    data = {"cubes": [{"sample_mark": "G26-CON-1198", "_shotcrete": True,
                       "cube_no": "109", "tests": CORES_28 + CUBES_28}]}
    assert writer.needed_sheet_counts(data) == {1198: 2}


def test_core_page_without_cube_rows_asks_for_one():
    data = {"cubes": [{"sample_mark": "G26-CON-1198", "_shotcrete": True,
                       "cube_no": "109", "tests": list(CORES_28)}]}
    assert writer.needed_sheet_counts(data) == {1198: 1}


def test_two_cubes_sharing_a_sample_id_add_up():
    one = {"sample_mark": "G26-CON-1198", "tests": list(CORES_28)}
    two = {"sample_mark": "G26-CON-1198", "tests": list(CORES_28)}
    assert writer.needed_sheet_counts({"cubes": [one, two]}) == {1198: 2}


def test_unreadable_sample_mark_is_not_counted():
    data = {"cubes": [{"sample_mark": None, "tests": list(CORES_28)}]}
    assert writer.needed_sheet_counts(data) == {}


# ---------- matching: the extra sheet a cube set needs ----------

def _sheet(name, num, wb="Book1.xlsx"):
    return {"workbook": wb, "sheet": name,
            "sample_id_raw": f"G26-CON-{num}", "sample_id_num": num}


def test_mixed_cube_also_claims_the_next_sheet_with_that_id():
    cube = {"sample_mark": "G26-CON-1198", "_shotcrete": True,
            "cube_no": "109", "tests": CORES_28 + CUBES_28}
    sheets = [_sheet("109", 1198), _sheet("109b", 1198)]
    res = writer.match_cubes_to_sheets({"cubes": [cube]}, sheets=sheets)
    assert res[0]["matched_sheet"]["sheet"] == "109"
    assert [s["sheet"] for s in res[0]["extra_sheets"]] == ["109b"]


def test_pure_core_cube_claims_no_extra_sheet():
    cube = {"sample_mark": "G26-CON-1198", "_shotcrete": True,
            "cube_no": "109", "tests": list(CORES_28)}
    sheets = [_sheet("109", 1198), _sheet("109b", 1198)]
    res = writer.match_cubes_to_sheets({"cubes": [cube]}, sheets=sheets)
    assert res[0]["extra_sheets"] == []


def test_missing_second_sheet_is_not_invented():
    cube = {"sample_mark": "G26-CON-1198", "_shotcrete": True,
            "cube_no": "109", "tests": CORES_28 + CUBES_28}
    res = writer.match_cubes_to_sheets({"cubes": [cube]},
                                       sheets=[_sheet("109", 1198)])
    assert res[0]["matched_sheet"]["sheet"] == "109"
    assert res[0]["extra_sheets"] == []


def test_an_extra_sheet_is_not_stolen_from_the_next_cube():
    # Two separate cubes share the ID (the old two-set case). Neither may
    # swallow the other's sheet.
    a = {"sample_mark": "G26-CON-1198", "tests": list(CORES_28)}
    b = {"sample_mark": "G26-CON-1198", "tests": list(CORES_28)}
    sheets = [_sheet("109", 1198), _sheet("109b", 1198)]
    res = writer.match_cubes_to_sheets({"cubes": [a, b]}, sheets=sheets)
    assert res[0]["matched_sheet"]["sheet"] == "109"
    assert res[1]["matched_sheet"]["sheet"] == "109b"


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
