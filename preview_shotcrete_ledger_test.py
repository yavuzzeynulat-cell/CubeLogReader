"""Visual test for ShotcreteLedgerPreviewWindow — real-world scenario.

PDF: 28days 20.05.2026.pdf (4 cubes, all shotcrete).
Excel state simulated below — only G26-CON-599 should be visible at
the top (7-day already in Excel → 7-day group OFF, 28-day empty →
28-day group ON). The other three (602, 624, 646) are hidden behind
the toggle: 602 + 646 as "done" (Excel fully filled), 624 as "not
found" (no matching block)."""
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
        {  # 599 — 5+5, 7-day already filled in Excel, 28-day empty
            "sample_key": "G26-CON-599", "sample_id_num": 599,
            "sample_mark_raw": "G26-CON-599", "cube_no": 41, "cube_no_raw": 41,
            "start_row": 2693, "end_row": 2702, "size": 10,
            "rows_7d":  [2693, 2694, 2695, 2696, 2697],
            "rows_28d": [2698, 2699, 2700, 2701, 2702],
        },
        {  # 602 — 5+5, Excel fully filled (will be hidden as DONE)
            "sample_key": "G26-CON-602", "sample_id_num": 602,
            "sample_mark_raw": "G26-CON-602", "cube_no": 42, "cube_no_raw": 42,
            "start_row": 2703, "end_row": 2712, "size": 10,
            "rows_7d":  [2703, 2704, 2705, 2706, 2707],
            "rows_28d": [2708, 2709, 2710, 2711, 2712],
        },
        # 624 → INTENTIONALLY ABSENT — cube has no matching block
        {  # 646 — 3+3, Excel fully filled (will be hidden as DONE)
            "sample_key": "G26-CON-646", "sample_id_num": 646,
            "sample_mark_raw": "G26-CON-646", "cube_no": 44, "cube_no_raw": 44,
            "start_row": 2719, "end_row": 2724, "size": 6,
            "rows_7d":  [2719, 2720, 2721],
            "rows_28d": [2722, 2723, 2724],
        },
    ]


def _fake_values(ws, blocks):
    """Excel state simulation:
      599 → 7-day filled, 28-day empty
      602 → fully filled (done)
      646 → fully filled (done)"""
    out = {}
    for i, b in enumerate(blocks):
        n = b["size"]
        sid = b["sample_id_num"]
        if sid == 599:
            # 7-day filled (5 rows), 28-day empty (5 rows)
            di = [94.0, 94.0, 94.0, 94.0, 94.0] + [None] * 5
            he = [95.77, 95.67, 95.90, 96.29, 95.91] + [None] * 5
            we = [1501.0, 1500, 1498, 1500, 1494] + [None] * 5
            lo = [205.32, 189.58, 200.48, 191.47, 190.82] + [None] * 5
        elif sid in (602, 646):
            # Fully filled across all rows + all 4 columns
            di = [94.0] * n
            he = [96.0 + i * 0.1 for i in range(n)]
            we = [1530.0 + i for i in range(n)]
            lo = [240.0 + i for i in range(n)]
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
    """5 specimens at age 7 + 5 specimens at age 28, all values present."""
    out = []
    for i in range(5):
        out.append({
            "age_days": 7,
            "weight_gr": base_w + i,
            "load_kn": base_l + i,
            "core_diameter_mm": base_d,
            "core_height_mm": base_h + i * 0.1,
            "strength_nmm2": 30 + i,
            "_selected": i < 3,
        })
    for i in range(5):
        out.append({
            "age_days": 28,
            "weight_gr": base_w + 20 + i,
            "load_kn": base_l + 50 + i,
            "core_diameter_mm": base_d,
            "core_height_mm": base_h + 0.5 + i * 0.1,
            "strength_nmm2": 35 + i,
            "_selected": i < 3,
        })
    return out


def _shot_tests_7day_only(base_w, base_l, n=3, base_d=150.0, base_h=150.0):
    """Larger-core scenario (e.g. 150x150 mm) — only 7-day specimens."""
    out = []
    for i in range(n):
        out.append({
            "age_days": 7,
            "weight_gr": base_w + i,
            "load_kn": base_l + i,
            "core_diameter_mm": base_d,
            "core_height_mm": base_h,
            "strength_nmm2": 39 + i,
            "_selected": True,
        })
    return out


# Cubes from the real PDF "28days 20.05.2026.pdf":
CUBES = {
    "cubes": [
        # 041 — G26-CON-599 (the one we want to write today)
        {"sample_mark": "G26-CON-599", "cube_no": 41, "_shotcrete": True,
         "tests": _shot_tests_full(1494, 189.58)},
        # 042 — G26-CON-602 (Excel already fully filled → DONE)
        {"sample_mark": "G26-CON-602", "cube_no": 42, "_shotcrete": True,
         "tests": _shot_tests_full(1518, 229.81)},
        # 043 — G26-CON-624 (no matching block → NOT FOUND)
        {"sample_mark": "G26-CON-624", "cube_no": 43, "_shotcrete": True,
         "tests": _shot_tests_7day_only(7752, 865.96)},
        # 044 — G26-CON-646 (Excel already fully filled → DONE)
        {"sample_mark": "G26-CON-646", "cube_no": 44, "_shotcrete": True,
         "tests": _shot_tests_7day_only(7808, 914.62)},
    ]
}


root = ctk.CTk()
root.geometry("400x100")
root.title("Shotcrete ledger test parent")
ctk.CTkLabel(
    root,
    text="ShotcreteLedgerPreviewWindow — 599 visible, others hidden",
).pack(pady=20)

main.ShotcreteLedgerPreviewWindow(root, CUBES)

root.mainloop()
