"""
test_shotcrete_mixed.py — deterministic (no-Gemini) tests for core/shotcrete
pages that carry BOTH core rows and normal cube rows.

Observed on the real form (G26-CON-1198 / Core No 109, page of 18.09.2026):

  7-day  : 4 core rows filled, the preprinted 5th row left blank
  28-day : 5 core rows, then 3 NORMAL CUBE rows appended (mould 72/145/129)

The cube rows are physically different specimens: ~7700-7900 gr and
~930-970 kN (a 150 mm cube), against ~1550 gr / ~230-270 kN for a core.
On the notebook their Core Diameter / Core Height cells are left blank —
that blankness is what identifies them.

Two old assumptions break on this page:

  1. "5 rows per age" — the 7-day block legitimately holds 4.
  2. "a set's 7d and 28d counts are always equal" — here it is 4 vs 8.

Assumption 2 used to let _rebalance_shotcrete_ages reassign ages by
position for any 10-row cube. On a genuine 4+6 read that MOVES a real
28-day core into the 7-day group. The rebalance must therefore fire only
on the misread signature it was written for (6+4).

Selection must also stop ranking cores against cubes: cubes always score
higher (41-43 vs 33-38), so a plain "top 3 by strength" hands every slot
to the cubes and leaves the cores unwritten.
"""
import reader


def _core(age, w, load, s, diam=94.00, h=96.00):
    return {"age_days": age, "mould_no": None, "weight_gr": w,
            "load_kn": load, "strength_nmm2": s,
            "core_diameter_mm": diam, "core_height_mm": h}


def _cube(age, w, load, s, mould):
    return {"age_days": age, "mould_no": mould, "weight_gr": w,
            "load_kn": load, "strength_nmm2": s,
            "core_diameter_mm": None, "core_height_mm": None}


# --- G26-CON-1198, read straight off the page ---
REAL_7D = [
    _core(7, 1538, 203.20, 29.28, 94.00, 96.20),
    _core(7, 1539, 202.81, 29.22, 94.00, 96.15),
    _core(7, 1543, 155.30, 22.38, 94.00, 96.76),
    _core(7, 1550, 191.63, 27.61, 94.00, 96.74),
]
REAL_28D_CORES = [
    _core(28, 1552, 239.11, 34.46, 94.00, 96.66),
    _core(28, 1554, 268.56, 38.70, 94.00, 96.59),
    _core(28, 1558, 266.12, 38.35, 94.00, 96.84),
    _core(28, 1556, 229.39, 33.05, 94.00, 96.85),
    _core(28, 1549, 232.97, 33.57, 94.00, 96.32),
]
REAL_28D_CUBES = [
    _cube(28, 7869, 941.73, 41.85, "72"),
    _cube(28, 7789, 932.03, 41.42, "145"),
    _cube(28, 7708, 965.66, 42.92, "129"),
]


def _run(tests):
    out = reader._process_shotcrete_cubes({"cubes": [{
        "_shotcrete_page": True, "cube_no": "109",
        "sample_mark": "G26-CON-1198", "tests": list(tests),
    }]})
    return out["cubes"][0]


def _by_age(cube, age):
    return [t for t in cube["tests"] if t.get("age_days") == age]


def _selected(cube, age):
    return [t["strength_nmm2"] for t in _by_age(cube, age) if t.get("_selected")]


# ---------- cube-row identification ----------

def test_blank_core_dimensions_mark_a_cube_row():
    assert reader.is_cube_row(REAL_28D_CUBES[0]) is True


def test_filled_core_dimensions_mark_a_core_row():
    assert reader.is_cube_row(REAL_28D_CORES[0]) is False


# ---------- row survival ----------

def test_cube_rows_survive_the_28d_block():
    res = _run(REAL_7D + REAL_28D_CORES + REAL_28D_CUBES)
    assert len(_by_age(res, 7)) == 4, len(_by_age(res, 7))
    assert len(_by_age(res, 28)) == 8, len(_by_age(res, 28))


def test_cube_rows_keep_form_order_after_the_cores():
    res = _run(REAL_7D + REAL_28D_CORES + REAL_28D_CUBES)
    weights = [t["weight_gr"] for t in _by_age(res, 28)]
    assert weights == [1552, 1554, 1558, 1556, 1549, 7869, 7789, 7708], weights


# ---------- rebalance must not corrupt a genuine 4+6 read ----------

def test_four_plus_six_is_left_alone():
    # 4 real 7-day rows + 6 real 28-day rows = 10 total. The old
    # position-based rebalance relabelled the first 28-day core as 7-day.
    extra = _core(28, 1547, 235.10, 33.90, 94.00, 96.40)
    res = _run(REAL_7D + REAL_28D_CORES + [extra])
    assert len(_by_age(res, 7)) == 4, len(_by_age(res, 7))
    assert 34.46 not in [t["strength_nmm2"] for t in _by_age(res, 7)]


def test_rebalance_skips_a_block_holding_cube_rows():
    # 6+4 is the misread signature, but the position rule only describes the
    # preprinted CORE rows. With cube rows in the 28-day block it would
    # relabel a cube as 7-day.
    t7 = [_core(7, 1500 + i, 200.0 + i, 30.0 + i) for i in range(6)]
    t28 = [_core(28, 1560, 240.00, 34.00)] + REAL_28D_CUBES
    res = _run(t7 + t28)
    assert len(_by_age(res, 7)) == 6, len(_by_age(res, 7))
    assert all(not reader.is_cube_row(t) for t in _by_age(res, 7))


# ---------- selection ----------

def test_cubes_do_not_steal_the_core_slots():
    res = _run(REAL_7D + REAL_28D_CORES + REAL_28D_CUBES)
    picked = _selected(res, 28)
    # the three strongest cores, plus every cube
    assert sorted(picked) == sorted([34.46, 38.70, 38.35,
                                     41.85, 41.42, 42.92]), picked


def test_every_cube_row_is_selected():
    res = _run(REAL_7D + REAL_28D_CORES + REAL_28D_CUBES)
    cubes = [t for t in _by_age(res, 28) if reader.is_cube_row(t)]
    assert all(t.get("_selected") for t in cubes)


def test_four_core_rows_pick_the_best_three():
    res = _run(REAL_7D)
    assert sorted(_selected(res, 7)) == sorted([29.28, 29.22, 27.61]), \
        _selected(res, 7)


def test_pure_core_group_of_five_still_picks_three():
    res = _run(REAL_28D_CORES)
    assert len(_selected(res, 28)) == 3, _selected(res, 28)


def test_core_row_without_strength_is_never_selected():
    blank = _core(28, None, None, None)
    res = _run(REAL_28D_CORES + [blank])
    assert len(_selected(res, 28)) == 3


# ---------- two core sets plus a cube set ----------
# Cards map onto sheets in set order. The cube rows are last on the form, so
# they are the last set: cores set 1 -> sheet 1, cores set 2 -> sheet 2, cubes
# -> sheet 3. Without an index of their own the cubes fall into set 1 and take
# sheet 2, pushing core set 2 onto sheet 3.

TWO_CORE_SETS = ([_core(28, 1550 + i, 230.00 + i, 33.00 + i) for i in range(5)]
                 + [_core(28, 1570 + i, 250.00 + i, 36.00 + i) for i in range(5)])


def test_cubes_become_the_set_after_the_last_core_set():
    res = _run(TWO_CORE_SETS + REAL_28D_CUBES)
    idx = {t["strength_nmm2"]: t.get("_set_index")
           for t in _by_age(res, 28) if reader.is_cube_row(t)}
    assert set(idx.values()) == {3}, idx


def test_core_sets_keep_their_own_indices():
    res = _run(TWO_CORE_SETS + REAL_28D_CUBES)
    cores = [t.get("_set_index") for t in _by_age(res, 28)
             if not reader.is_cube_row(t)]
    assert cores == [1] * 5 + [2] * 5, cores


def test_a_single_core_set_leaves_the_cubes_unindexed():
    # 1198's shape: one card, one list — the cubes must NOT be split off.
    res = _run(REAL_7D + REAL_28D_CORES + REAL_28D_CUBES)
    assert all(t.get("_set_index") is None
               for t in _by_age(res, 28) if reader.is_cube_row(t))


def test_two_core_sets_and_cubes_split_into_three_cards():
    data = reader._postprocess_cubes({"cubes": [{
        "_shotcrete_page": True, "cube_no": "109",
        "sample_mark": "G26-CON-1198",
        "tests": TWO_CORE_SETS + REAL_28D_CUBES,
    }]})
    assert len(data["cubes"]) == 3, len(data["cubes"])
    assert [c["_set_index"] for c in data["cubes"]] == [1, 2, 3]


def test_the_third_card_holds_only_the_cube_rows():
    data = reader._postprocess_cubes({"cubes": [{
        "_shotcrete_page": True, "cube_no": "109",
        "sample_mark": "G26-CON-1198",
        "tests": TWO_CORE_SETS + REAL_28D_CUBES,
    }]})
    third = data["cubes"][2]["tests"]
    assert len(third) == 3, len(third)
    assert all(reader.is_cube_row(t) for t in third)
    assert [t["weight_gr"] for t in third] == [7869, 7789, 7708]


def test_a_single_core_set_and_cubes_stay_on_one_card():
    data = reader._postprocess_cubes({"cubes": [{
        "_shotcrete_page": True, "cube_no": "109",
        "sample_mark": "G26-CON-1198",
        "tests": REAL_7D + REAL_28D_CORES + REAL_28D_CUBES,
    }]})
    assert len(data["cubes"]) == 1, len(data["cubes"])


# ---------- how many Excel sheets the cube needs ----------
# An Excel sheet holds 3 specimen slots per age. Cores and cubes are separate
# sets and never share a sheet, so a mixed block needs a second sheet. The
# forward scan has to know this up front: it stops as soon as every cube has
# been matched, so a cube that silently needs two sheets would leave the
# second one unscanned.

def _shot_cube(tests):
    return {"_shotcrete": True, "sample_mark": "G26-CON-1198",
            "cube_no": "109", "tests": list(tests)}


def test_pure_core_cube_needs_one_sheet():
    assert reader.sheets_needed_for_cube(
        _shot_cube(REAL_7D + REAL_28D_CORES)) == 1


def test_cube_rows_in_the_28d_block_need_a_second_sheet():
    assert reader.sheets_needed_for_cube(
        _shot_cube(REAL_7D + REAL_28D_CORES + REAL_28D_CUBES)) == 2


def test_normal_concrete_cube_needs_one_sheet():
    plain = {"sample_mark": "G26-CON-395", "tests": [
        {"age_days": 7, "weight_gr": 8360, "load_kn": 1102.34,
         "strength_nmm2": 48.99, "core_diameter_mm": None,
         "core_height_mm": None},
    ]}
    assert reader.sheets_needed_for_cube(plain) == 1


def test_a_cube_only_28d_block_still_needs_one_sheet():
    # 28-day block holds nothing but cube rows — one group, one sheet.
    assert reader.sheets_needed_for_cube(_shot_cube(REAL_28D_CUBES)) == 1


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
