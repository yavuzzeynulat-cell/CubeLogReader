"""Visual test for ShotcreteLedgerPreviewWindow — hide-state sync.

Scenario:
  599 → visible, actionable
  602 → ledger DONE (fully filled in Excel)
  624 → NOT FOUND in ledger (no matching block)
  646 → ledger DONE (fully filled in Excel)
  701 → ledger EMPTY but _hidden_in_main="no_match" → should appear
        in "Hidden in main preview" group with checkbox OFF
  702 → ledger EMPTY but _hidden_in_main="done" → should appear in
        "Hidden in main preview" group with checkbox OFF

Expected: toggle reads "▸ 5 hidden". Open it → 701 + 702 visible
under "Hidden in main preview" section, but their per-card
checkboxes start UNTICKED. User can re-tick to opt-in to writing.
"""
import customtkinter as ctk
import main
import writer


class _FakeWS:
    def Range(self, _):
        class _R:
            Value = None
        return _R()


def _fake_find_candidates(*a, **k):
    return [
        (None, _FakeWS(), "Shotcrete sample form.xlsx", "Shotcrete Concrete Results"),
    ]


def _fake_blocks(ws):
    return [
        {
            "sample_key": "G26-CON-599", "sample_id_num": 599,
            "sample_mark_raw": "G26-CON-599", "cube_no": 41, "cube_no_raw": 41,
            "start_row": 2693, "end_row": 2702, "size": 10,
            "rows_7d":  [2693, 2694, 2695, 2696, 2697],
            "rows_28d": [2698, 2699, 2700, 2701, 2702],
        },
        {
            "sample_key": "G26-CON-602", "sample_id_num": 602,
            "sample_mark_raw": "G26-CON-602", "cube_no": 42, "cube_no_raw": 42,
            "start_row": 2703, "end_row": 2712, "size": 10,
            "rows_7d":  [2703, 2704, 2705, 2706, 2707],
            "rows_28d": [2708, 2709, 2710, 2711, 2712],
        },
        # 624 absent on purpose (not_found)
        {
            "sample_key": "G26-CON-646", "sample_id_num": 646,
            "sample_mark_raw": "G26-CON-646", "cube_no": 44, "cube_no_raw": 44,
            "start_row": 2719, "end_row": 2724, "size": 6,
            "rows_7d":  [2719, 2720, 2721],
            "rows_28d": [2722, 2723, 2724],
        },
        {
            "sample_key": "G26-CON-701", "sample_id_num": 701,
            "sample_mark_raw": "G26-CON-701", "cube_no": 45, "cube_no_raw": 45,
            "start_row": 2725, "end_row": 2734, "size": 10,
            "rows_7d":  [2725, 2726, 2727, 2728, 2729],
            "rows_28d": [2730, 2731, 2732, 2733, 2734],
        },
        {
            "sample_key": "G26-CON-702", "sample_id_num": 702,
            "sample_mark_raw": "G26-CON-702", "cube_no": 46, "cube_no_raw": 46,
            "start_row": 2735, "end_row": 2744, "size": 10,
            "rows_7d":  [2735, 2736, 2737, 2738, 2739],
            "rows_28d": [2740, 2741, 2742, 2743, 2744],
        },
    ]


def _fake_values(ws, blocks):
    """Excel state per cube:
      599 → 7-day filled, 28-day empty
      602 → fully filled
      646 → fully filled
      701 → fully empty (would be actionable without hidden_in_main)
      702 → fully empty (same)"""
    out = {}
    for i, b in enumerate(blocks):
        n = b["size"]
        sid = b["sample_id_num"]
        if sid == 599:
            di = [94.0, 94.0, 94.0, 94.0, 94.0] + [None] * 5
            he = [95.77, 95.67, 95.90, 96.29, 95.91] + [None] * 5
            we = [1501.0, 1500, 1498, 1500, 1494] + [None] * 5
            lo = [205.32, 189.58, 200.48, 191.47, 190.82] + [None] * 5
        elif sid in (602, 646):
            di = [94.0] * n
            he = [96.0 + j * 0.1 for j in range(n)]
            we = [1530.0 + j for j in range(n)]
            lo = [240.0 + j for j in range(n)]
        else:
            di = [None] * n
            he = [None] * n
            we = [None] * n
            lo = [None] * n
        out[i] = {"diameters": di, "heights": he, "weights": we, "loads": lo}
    return out


writer.find_shotcrete_ledger_candidates = _fake_find_candidates
writer.read_shotcrete_ledger_blocks = _fake_blocks
writer.read_shotcrete_ledger_values = _fake_values


def _shot_tests_full(base_w, base_l, base_d=94.0, base_h=95.5):
    out = []
    for i in range(5):
        out.append({
            "age_days": 7, "weight_gr": base_w + i, "load_kn": base_l + i,
            "core_diameter_mm": base_d, "core_height_mm": base_h + i * 0.1,
            "strength_nmm2": 30 + i, "_selected": i < 3,
        })
    for i in range(5):
        out.append({
            "age_days": 28, "weight_gr": base_w + 20 + i,
            "load_kn": base_l + 50 + i, "core_diameter_mm": base_d,
            "core_height_mm": base_h + 0.5 + i * 0.1,
            "strength_nmm2": 35 + i, "_selected": i < 3,
        })
    return out


def _shot_tests_7day_only(base_w, base_l, n=3, base_d=150.0, base_h=150.0):
    out = []
    for i in range(n):
        out.append({
            "age_days": 7, "weight_gr": base_w + i, "load_kn": base_l + i,
            "core_diameter_mm": base_d, "core_height_mm": base_h,
            "strength_nmm2": 39 + i, "_selected": True,
        })
    return out


CUBES = {
    "cubes": [
        {"sample_mark": "G26-CON-599", "cube_no": 41, "_shotcrete": True,
         "tests": _shot_tests_full(1494, 189.58)},
        {"sample_mark": "G26-CON-602", "cube_no": 42, "_shotcrete": True,
         "tests": _shot_tests_full(1518, 229.81)},
        {"sample_mark": "G26-CON-624", "cube_no": 43, "_shotcrete": True,
         "tests": _shot_tests_7day_only(7752, 865.96)},
        {"sample_mark": "G26-CON-646", "cube_no": 44, "_shotcrete": True,
         "tests": _shot_tests_7day_only(7808, 914.62)},
        # NEW — main preview marked these as no_match / done.
        # Ledger blocks for them are EMPTY, so without the new logic
        # they'd be actionable. Expected: pulled into hidden group,
        # per-card checkbox starts OFF.
        {"sample_mark": "G26-CON-701", "cube_no": 45, "_shotcrete": True,
         "_hidden_in_main": "no_match",
         "tests": _shot_tests_full(1480, 180.0)},
        {"sample_mark": "G26-CON-702", "cube_no": 46, "_shotcrete": True,
         "_hidden_in_main": "done",
         "tests": _shot_tests_full(1485, 185.0)},
    ]
}


root = ctk.CTk()
root.geometry("520x140")
root.title("Shotcrete ledger — hide-sync test")
ctk.CTkLabel(
    root,
    text=(
        "ShotcreteLedgerPreviewWindow — hide-state sync test\n"
        "Expected: 599 visible, ▸ 5 hidden  (602, 646 done · 624 not "
        "found · 701, 702 from main, unticked)"
    ),
    justify="left",
).pack(pady=20, padx=12)

main.ShotcreteLedgerPreviewWindow(root, CUBES)

root.mainloop()
