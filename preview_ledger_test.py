"""Quick visual test for LedgerPreviewWindow — no Excel needed."""
import customtkinter as ctk
import main
import writer


# Stub out Excel-touching writer functions
class _FakeWS:
    def Range(self, _):
        class _R:
            Value = None
        return _R()


def _fake_find(*a, **k):
    return (None, _FakeWS(), "Concrete sample form (ALL).xlsx", "Concrete")


def _fake_find_candidates(*a, **k):
    # Two open ledgers → exercises the in-window file selector dropdown.
    return [
        (None, _FakeWS(), "Concrete sample form (ALL).xlsx", "Concrete"),
        (None, _FakeWS(), "Concrete sample form APRIL.xlsx", "Concrete"),
    ]


def _fake_blocks(ws):
    return [
        {  # 522 — full 3+3 normal cube, all empty in ledger
            "sample_key": "G26-CON-522", "sample_id_num": 522,
            "sample_mark_raw": "G26-CON-522", "cube_no": 496, "cube_no_raw": 496,
            "start_row": 24206, "end_row": 24211, "size": 6,
            "rows_7d": [24206, 24207, 24208],
            "rows_28d": [24209, 24210, 24211],
        },
        {  # 525 — 7-day rows already filled in ledger (gray)
            "sample_key": "G26-CON-525", "sample_id_num": 525,
            "sample_mark_raw": "G26-CON-525", "cube_no": 499, "cube_no_raw": 499,
            "start_row": 24218, "end_row": 24223, "size": 6,
            "rows_7d": [24218, 24219, 24220],
            "rows_28d": [24221, 24222, 24223],
        },
        {  # 530 — 12-row 2-set, interleaved (set1: 3×7g+3×28g, set2: 3×7g+3×28g)
            "sample_key": "G26-CON-530", "sample_id_num": 530,
            "sample_mark_raw": "G26-CON-530", "cube_no": 504, "cube_no_raw": 504,
            "start_row": 24230, "end_row": 24241, "size": 12,
            "rows_7d":  [24230, 24231, 24232, 24236, 24237, 24238],
            "rows_28d": [24233, 24234, 24235, 24239, 24240, 24241],
        },
        {  # 533 — 7-row odd-size block (4×7d + 3×28d)
            "sample_key": "G26-CON-533", "sample_id_num": 533,
            "sample_mark_raw": "G26-CON-533", "cube_no": 507, "cube_no_raw": 507,
            "start_row": 24250, "end_row": 24256, "size": 7,
            "rows_7d": [24250, 24251, 24252, 24253],
            "rows_28d": [24254, 24255, 24256],
        },
    ]


def _fake_values(ws, blocks):
    out = {}
    for i, b in enumerate(blocks):
        n = b["size"]
        # 525 → fill the 7-day rows with old values, leave 28-day empty
        if b["sample_id_num"] == 525:
            ws_w = [4500.0, 4520.0, 4480.0, None, None, None]
            ws_l = [320.0, 325.0, 318.0, None, None, None]
        # 522 → ALL filled in ledger → "done" card (hidden by default)
        elif b["sample_id_num"] == 522:
            ws_w = [4400.0, 4410, 4420, 4500, 4510, 4520]
            ws_l = [310, 312, 314, 820, 825, 830]
        else:
            ws_w = [None] * n
            ws_l = [None] * n
        out[i] = {"weights": ws_w, "loads": ws_l}
    return out


writer.find_ledger_sheet = _fake_find
writer.find_ledger_candidates = _fake_find_candidates
writer.read_ledger_blocks = _fake_blocks
writer.read_ledger_values = _fake_values


# Synthetic cubes_data
CUBES = {
    "cubes": [
        {
            "sample_mark": "G26-CON-522", "cube_no": 496,
            "tests_7d": [
                {"weight_gr": 8456, "load_kn": 612.4},
                {"weight_gr": 8438, "load_kn": 605.1},
                {"weight_gr": 8472, "load_kn": 618.9},
            ],
            "tests_28d": [
                {"weight_gr": 8520, "load_kn": 845.2},
                {"weight_gr": 8498, "load_kn": 838.7},
                {"weight_gr": 8505, "load_kn": 842.0},
            ],
        },
        {
            "sample_mark": "G26-CON-525", "cube_no": 499,
            "tests_7d": [
                {"weight_gr": 8400, "load_kn": 600.0},
                {"weight_gr": 8410, "load_kn": 602.0},
                {"weight_gr": 8420, "load_kn": 604.0},
            ],
            "tests_28d": [
                {"weight_gr": 8480, "load_kn": 820.0},
                {"weight_gr": 8490, "load_kn": 825.0},
                {"weight_gr": 8500, "load_kn": 830.0},
            ],
        },
        {
            "sample_mark": "G26-CON-530", "cube_no": 504,
            "tests_7d": [{"weight_gr": 8400 + i, "load_kn": 610.0 + i} for i in range(6)],
            "tests_28d": [{"weight_gr": 8500 + i, "load_kn": 845.0 + i} for i in range(6)],
        },
        {
            "sample_mark": "G26-CON-533", "cube_no": 507,
            "tests_7d": [
                {"weight_gr": 8455, "load_kn": 611.0},
                {"weight_gr": 8460, "load_kn": 613.0},
                {"weight_gr": 8465, "load_kn": 615.0},
                {"weight_gr": 8470, "load_kn": 617.0},
            ],
            "tests_28d": [
                {"weight_gr": 8530, "load_kn": 848.0},
                {"weight_gr": 8535, "load_kn": 850.0},
                {"weight_gr": 8540, "load_kn": 852.0},
            ],
        },
        # --- These three cubes have no matching block → "not found" toggle test
        {
            "sample_mark": "G26-CON-901", "cube_no": 700,
            "tests_7d": [{"weight_gr": 8000, "load_kn": 500.0}],
            "tests_28d": [],
        },
        {
            "sample_mark": "G26-CON-902", "cube_no": 701,
            "tests_7d": [{"weight_gr": 8010, "load_kn": 502.0}],
            "tests_28d": [],
        },
        {
            "sample_mark": "G26-CON-903", "cube_no": 702,
            "tests_7d": [{"weight_gr": 8020, "load_kn": 504.0}],
            "tests_28d": [],
        },
    ]
}


# Bypass merge: feed the cubes through directly
def _fake_merge(cubes_data):
    out = []
    for c in cubes_data["cubes"]:
        cc = dict(c)
        cc["sample_key"] = "G26-CON-" + str(c["sample_mark"].split("-")[-1])
        cc["sample_id_num"] = int(c["sample_mark"].split("-")[-1])
        out.append(cc)
    return out


writer.merge_cubes_for_ledger = _fake_merge


def _fake_match(merged, blocks):
    by_key = {b["sample_key"]: b for b in blocks}
    results = []
    for c in merged:
        block = by_key.get(c["sample_key"])
        reason = None if block is not None else "not_found"
        results.append({"cube": c, "block": block, "mismatch_reason": reason})
    return results


writer.match_cubes_to_blocks = _fake_match


root = ctk.CTk()
root.geometry("400x100")
root.title("Test parent")
ctk.CTkLabel(root, text="LedgerPreviewWindow test").pack(pady=20)

main.LedgerPreviewWindow(root, CUBES)

root.mainloop()
