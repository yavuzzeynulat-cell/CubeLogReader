"""Quick visual test for ShotcreteLedgerPreviewWindow — no Excel needed.

Two fake open ledgers exercise the file-selector dropdown.
Four fake blocks cover: single 5+5 (clean), 5+5 with 7-day pre-filled,
two-set 10+10 interleaved, and an odd-count 4+3."""
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
        (None, _FakeWS(), "Shotcrete sample form APRIL.xlsx", "Shotcrete Concrete Results"),
    ]


def _fake_blocks(ws):
    # NOTE: sample_key must match what ledger_sample_key() produces on the
    # cube's sample_mark — leading zeros are stripped to integers, e.g.
    # "G26-CON-027" → "G26-CON-27".
    return [
        {  # 027 — clean 5+5
            "sample_key": "G26-CON-27", "sample_id_num": 27,
            "sample_mark_raw": "G26-CON-027", "cube_no": 2, "cube_no_raw": 2,
            "start_row": 2309, "end_row": 2318, "size": 10,
            "rows_7d":  [2309, 2310, 2311, 2312, 2313],
            "rows_28d": [2314, 2315, 2316, 2317, 2318],
        },
        {  # 082 — 7-day pre-filled in ledger (gray cells)
            "sample_key": "G26-CON-82", "sample_id_num": 82,
            "sample_mark_raw": "G26-CON-082", "cube_no": 3, "cube_no_raw": 3,
            "start_row": 2319, "end_row": 2328, "size": 10,
            "rows_7d":  [2319, 2320, 2321, 2322, 2323],
            "rows_28d": [2324, 2325, 2326, 2327, 2328],
        },
        {  # 200 — two-set: 10+10 interleaved across 20 rows
            "sample_key": "G26-CON-200", "sample_id_num": 200,
            "sample_mark_raw": "G26-CON-200", "cube_no": 5, "cube_no_raw": 5,
            "start_row": 2340, "end_row": 2359, "size": 20,
            "rows_7d":  [2340, 2341, 2342, 2343, 2344,
                         2350, 2351, 2352, 2353, 2354],
            "rows_28d": [2345, 2346, 2347, 2348, 2349,
                         2355, 2356, 2357, 2358, 2359],
        },
        {  # 333 — odd 4+3 block (partial data scenario)
            "sample_key": "G26-CON-333", "sample_id_num": 333,
            "sample_mark_raw": "G26-CON-333", "cube_no": 7, "cube_no_raw": 7,
            "start_row": 2370, "end_row": 2376, "size": 7,
            "rows_7d":  [2370, 2371, 2372, 2373],
            "rows_28d": [2374, 2375, 2376],
        },
    ]


def _fake_values(ws, blocks):
    out = {}
    for i, b in enumerate(blocks):
        n = b["size"]
        # 082 → 7-day rows already filled (gray rendering)
        if b["sample_id_num"] == 82:
            di = [94.0, 94.5, 94.8, 95.0, 95.3] + [None] * (n - 5)
            he = [94.1, 94.6, 94.9, 95.1, 95.4] + [None] * (n - 5)
            we = [1500.0, 1510, 1520, 1530, 1540] + [None] * (n - 5)
            lo = [200.0, 210, 220, 230, 240] + [None] * (n - 5)
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


def _shot_tests(n7, n28, base_w=1500, base_l=220, base_d=94.0, base_h=94.0):
    out = []
    for i in range(n7):
        out.append({
            "age_days": 7, "weight_gr": base_w + i,
            "load_kn": base_l + i, "core_diameter_mm": base_d + i * 0.1,
            "core_height_mm": base_h + i * 0.1, "strength_nmm2": 30 + i,
            "_selected": i < 3,
        })
    for i in range(n28):
        out.append({
            "age_days": 28, "weight_gr": base_w + 50 + i,
            "load_kn": base_l + 50 + i, "core_diameter_mm": base_d + 1 + i * 0.1,
            "core_height_mm": base_h + 1 + i * 0.1, "strength_nmm2": 40 + i,
            "_selected": i < 3,
        })
    return out


CUBES = {
    "cubes": [
        {"sample_mark": "G26-CON-027", "cube_no": 2, "_shotcrete": True,
         "tests": _shot_tests(5, 5)},
        {"sample_mark": "G26-CON-082", "cube_no": 3, "_shotcrete": True,
         "tests": _shot_tests(5, 5, base_w=1550, base_l=230)},
        {"sample_mark": "G26-CON-200", "cube_no": 5, "_shotcrete": True,
         "_set_index": 1, "tests": _shot_tests(5, 5, base_w=1600)},
        {"sample_mark": "G26-CON-200", "cube_no": 5, "_shotcrete": True,
         "_set_index": 2, "tests": _shot_tests(5, 5, base_w=1620)},
        {"sample_mark": "G26-CON-333", "cube_no": 7, "_shotcrete": True,
         "tests": _shot_tests(4, 3, base_w=1700)},
    ]
}


root = ctk.CTk()
root.geometry("400x100")
root.title("Shotcrete ledger test parent")
ctk.CTkLabel(root, text="ShotcreteLedgerPreviewWindow test").pack(pady=20)

main.ShotcreteLedgerPreviewWindow(root, CUBES)

root.mainloop()
