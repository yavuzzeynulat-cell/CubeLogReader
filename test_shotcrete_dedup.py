"""
test_shotcrete_dedup.py — deterministic (no-Gemini) tests for the shotcrete
row-healing in reader._process_shotcrete_cubes.

Two real Gemini misreads on Core/shotcrete forms are covered:

  A) MISCLASSIFICATION (the main bug): the first 28-day row's age digit is
     read as "7", so a clean 5+5 set comes back 6+4. The row VALUE is correct,
     only its age label is wrong (total stays 10). Confirmed on the real PDF:
     cube 045 -> extra 7d row 1532/218.44/31.48 == the real first 28d row.

  B) DUPLICATION: a row is emitted twice (sometimes once per age), e.g. 6+5.

Form invariant (user-confirmed): a set's 7d and 28d counts are always equal
(5/5 per set, 10/10 two-set); 28d is either fully present or fully empty
(never a partial 5+3). The Age column is preprinted 5x"7" then 5x"28".

Fix: drop exact-duplicate rows, then for a clean 10-row set with 28d present
assign age by position (first 5 -> 7d, last 5 -> 28d) instead of the label.
"""
import reader


def _ages(cube):
    t7 = [t for t in cube["tests"] if t.get("age_days") == 7]
    t28 = [t for t in cube["tests"] if t.get("age_days") == 28]
    return len(t7), len(t28)


def _vals(cube, age):
    return {(t["weight_gr"], t["load_kn"], t["strength_nmm2"])
            for t in cube["tests"] if t.get("age_days") == age}


def _row(age, w, l, s):
    return {"age_days": age, "weight_gr": w, "load_kn": l, "strength_nmm2": s,
            "core_diameter_mm": 94, "core_height_mm": 96}


# Real 045 data from the misread PDF
REAL7 = [(1531, 210.10, 30.27), (1515, 209.39, 30.17), (1510, 208.72, 30.07),
         (1530, 195.25, 28.13), (1523, 218.30, 31.46)]
REAL28 = [(1532, 218.44, 31.48), (1534, 245.07, 35.31), (1538, 228.54, 32.93),
          (1541, 250.08, 36.03), (1517, 255.74, 36.85)]


def _run(cube):
    out = reader._process_shotcrete_cubes({"cubes": [cube]})
    return out["cubes"][0]


def test_misclassification_first28_read_as_7():
    # 6+4: the first 28d row is mislabelled age 7 (its value still correct).
    t7 = [_row(7, *v) for v in REAL7] + [_row(7, *REAL28[0])]   # 6 rows
    t28 = [_row(28, *v) for v in REAL28[1:]]                    # 4 rows
    res = _run({"_shotcrete_page": True, "cube_no": "045", "tests": t7 + t28})
    assert _ages(res) == (5, 5), _ages(res)
    # the misread value must end up classified as 28d, not 7d
    assert _vals(res, 7) == set(REAL7), _vals(res, 7)
    assert _vals(res, 28) == set(REAL28), _vals(res, 28)


def test_cross_age_duplicate_then_rebalance():
    # 6+5: first 28d row appears BOTH as a 7d row and a 28d row (dup), total 11.
    t7 = [_row(7, *v) for v in REAL7] + [_row(7, *REAL28[0])]   # 6
    t28 = [_row(28, *v) for v in REAL28]                        # 5 (incl dup)
    res = _run({"_shotcrete_page": True, "tests": t7 + t28})
    assert _ages(res) == (5, 5), _ages(res)
    assert _vals(res, 28) == set(REAL28), _vals(res, 28)


def test_within_age_duplicate_dropped():
    # 6+5: an exact copy of the 5th 7d row, 28d clean -> dedup to 5, rebalance.
    t7 = [_row(7, *v) for v in REAL7] + [_row(7, *REAL7[4])]
    t28 = [_row(28, *v) for v in REAL28]
    res = _run({"_shotcrete_page": True, "tests": t7 + t28})
    assert _ages(res) == (5, 5), _ages(res)


def test_clean_set_untouched():
    t7 = [_row(7, *v) for v in REAL7]
    t28 = [_row(28, *v) for v in REAL28]
    res = _run({"_shotcrete_page": True, "tests": t7 + t28})
    assert _ages(res) == (5, 5), _ages(res)
    assert _vals(res, 7) == set(REAL7) and _vals(res, 28) == set(REAL28)


def test_only_7d_present_not_rebalanced():
    # 28d not tested yet -> 5+0 must stay 5+0 (no rows invented for 28d).
    t7 = [_row(7, *v) for v in REAL7]
    res = _run({"_shotcrete_page": True, "cube_no": "047", "tests": t7})
    assert _ages(res) == (5, 0), _ages(res)


def test_two_set_untouched():
    # 10+10 (two sets, all distinct) must remain 10+10.
    t7 = [_row(7, 1500 + i, 200.0 + i, 30.0 + i) for i in range(10)]
    t28 = [_row(28, 1600 + i, 240.0 + i, 34.0 + i) for i in range(10)]
    res = _run({"_shotcrete_page": True, "tests": t7 + t28})
    assert _ages(res) == (10, 10), _ages(res)


def test_two_set_7d_only_untouched():
    # Two sets, 28d not tested -> 10 rows all 7d. Must NOT be split into 5+5.
    t7 = [_row(7, 1500 + i, 200.0 + i, 30.0 + i) for i in range(10)]
    res = _run({"_shotcrete_page": True, "tests": list(t7)})
    assert _ages(res) == (10, 0), _ages(res)


def test_unbalanced_5_plus_10_loses_no_data():
    # Abnormal read (violates the equal-count invariant): 5x7d + 10x28d.
    # This must NEVER silently drop the 7d rows; every row survives the
    # full pipeline (the ledger 'doesn't fit block' warning then surfaces it).
    t7 = [_row(7, 1500 + i, 200.0 + i, 30.0 + i) for i in range(5)]
    t28 = [_row(28, 1600 + i, 240.0 + i, 34.0 + i) for i in range(10)]
    cube = {"_shotcrete_page": True, "sample_mark": "G26-CON-999",
            "cube_no": "099", "tests": t7 + t28}
    data = reader._postprocess_cubes({"cubes": [cube]})
    kept7 = sum(1 for c in data["cubes"]
                for t in c["tests"] if t.get("age_days") == 7)
    kept28 = sum(1 for c in data["cubes"]
                 for t in c["tests"] if t.get("age_days") == 28)
    assert (kept7, kept28) == (5, 10), (kept7, kept28)


def test_concrete_untouched():
    # Not a shotcrete page -> never rebalanced, even with a duplicate row.
    t7 = [_row(7, 8360, 1102.34, 48.99), _row(7, 8360, 1102.34, 48.99),
          _row(7, 8332, 1196.46, 53.17)]
    out = reader._process_shotcrete_cubes(
        {"cubes": [{"cube_no": "378", "sample_mark": "G26-CON-7000",
                    "tests": list(t7)}]})
    res = out["cubes"][0]
    assert not res.get("_shotcrete")
    assert _ages(res) == (3, 0), _ages(res)


if __name__ == "__main__":
    fns = [v for k, v in sorted(globals().items()) if k.startswith("test_")]
    failed = 0
    for fn in fns:
        try:
            fn()
            print(f"PASS {fn.__name__}")
        except AssertionError as e:
            failed += 1
            print(f"FAIL {fn.__name__}: got {e}")
    print(f"\n{len(fns) - failed}/{len(fns)} passed")
    raise SystemExit(1 if failed else 0)
