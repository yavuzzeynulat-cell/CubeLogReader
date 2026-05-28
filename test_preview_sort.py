"""Visual test for the new card display order in PreviewWindow.

Expected order from top to bottom:
  1. 28-day concrete: G26-CON-7001, G26-CON-7002
  2. 28-day shotcrete: G26-CON-7003
  3. 7-day concrete: G26-CON-7004, G26-CON-7005
  4. 7-day shotcrete: G26-CON-7006

Cubes are intentionally added to cubes_data in a SCRAMBLED order so
we can prove the sort is working — if the cards come out in PDF
order (7003, 7001, 7006, ...) the change failed.
"""
import customtkinter as ctk
from PIL import Image

import reader
import writer
from main import PreviewWindow

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


def _concrete_tests(ages):
    return [
        {"age_days": a, "mould_no": str(10 + i),
         "weight_gr": 8300 + i, "load_kn": 1200 + i, "strength_nmm2": 53 + i * 0.1}
        for i, a in enumerate(ages)
    ]


def _shot_tests(ages):
    return [
        {"age_days": a, "core_diameter_mm": 94, "core_height_mm": 188,
         "weight_gr": 2100 + i, "load_kn": 270 + i, "strength_nmm2": 38 + i * 0.3}
        for i, a in enumerate(ages)
    ]


# Bigger, scrambled input — proves sort across many cards.
# Expected display order:
#   28d concrete: 7001, 7002, 7008, 7011, 7013, 7015
#   28d shotcrete: 7003, 7016
#   7d concrete:  7004, 7005, 7009, 7012, 7014
#   7d shotcrete: 7006, 7010, 7017

def _con28(num):
    return {"cube_no": str(num), "sample_mark": f"G26-CON-{num:04d}",
            "tests": _concrete_tests([7, 7, 7, 28, 28, 28])}

def _con7(num):
    return {"cube_no": str(num), "sample_mark": f"G26-CON-{num:04d}",
            "tests": _concrete_tests([7, 7, 7])}

def _shot28(num):
    return {"cube_no": str(num), "sample_mark": f"G26-CON-{num:04d}",
            "_shotcrete_page": True,
            "tests": _shot_tests([7, 7, 7, 7, 7, 28, 28, 28, 28, 28])}

def _shot7(num):
    return {"cube_no": str(num), "sample_mark": f"G26-CON-{num:04d}",
            "_shotcrete_page": True,
            "tests": _shot_tests([7, 7, 7, 7, 7])}

# One card of each type, scrambled input order. Expected display:
#   1. 28d concrete  (7001)
#   2. 28d shotcrete (7003)
#   3. 7d concrete   (7004)
#   4. 7d shotcrete  (7006)
CUBES = {"cubes": [
    _shot7(7006),    # last in display
    _shot28(7003),   # second
    _con7(7004),     # third
    _con28(7001),    # first
]}


def main():
    cubes_data = reader._split_multi_set_cubes(
        reader._process_shotcrete_cubes(CUBES)
    )

    sheets = [
        {"workbook": "(mock.xlsx)", "sheet": f"sheet-{n}",
         "sample_id_raw": str(n), "sample_id_num": n}
        for n in (7001, 7003, 7004, 7006)
    ]
    scan_result = {
        "sheets": sheets, "scanned_count": len(sheets),
        "start_sheet": "sheet-7001", "workbook": "(mock.xlsx)",
        "found_all": True,
    }

    img = Image.new("RGB", (800, 1100), "white")

    # Realistic Excel state: 28d cubes (7001/7002/7003) have their
    # 7-day rows already written in a previous session → those cells
    # should render GRAY. 7d-only cubes (7004/7005/7006) have nothing
    # in Excel yet → green.
    # 28d cubes' sheets have 7-day already filled in Excel → 7d cells
    # show gray. 7d-only cubes have nothing in Excel → green.
    SHEETS_WITH_7D_FILLED = {"sheet-7001", "sheet-7003"}
    def _read_all_values(wb, sh):
        if sh in SHEETS_WITH_7D_FILLED:
            return {
                "weights_7d": [8301, 8311, 8321],
                "loads_7d":   [1201, 1212, 1223],
                "weights_28d": [None] * 3, "loads_28d": [None] * 3,
            }
        return {
            "weights_7d": [None] * 3, "loads_7d": [None] * 3,
            "weights_28d": [None] * 3, "loads_28d": [None] * 3,
        }
    writer.read_all_values = _read_all_values
    writer.cross_check_7day = lambda *_a, **_k: []

    root = ctk.CTk()
    root.withdraw()
    PreviewWindow(root, img, cubes_data, scan_result)
    root.mainloop()


if __name__ == "__main__":
    main()
