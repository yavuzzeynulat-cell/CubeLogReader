"""
test_preview_shot.py — Open the real PreviewWindow with a synthetic
shotcrete cube so we can see the final UI without running Gemini or
needing an Excel file open.

Cube #502 is a two-set shotcrete (10+10) so the splitter also exercises
the "set 1/2" / "set 2/2" path.
"""
import customtkinter as ctk
from PIL import Image

import reader
from main import PreviewWindow

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


def _make_cubes():
    cubes = {
        "cubes": [
            # 1) Normal cube (unchanged path)
            {
                "cube_no": "378", "sample_mark": "G26-CON-7000",
                "tests": [
                    {"age_days": 7,  "mould_no": "110", "weight_gr": 8360, "load_kn": 1102.34, "strength_nmm2": 48.99},
                    {"age_days": 7,  "mould_no": "208", "weight_gr": 8332, "load_kn": 1196.46, "strength_nmm2": 53.17},
                    {"age_days": 7,  "mould_no": "48",  "weight_gr": 8315, "load_kn": 1202.46, "strength_nmm2": 53.44},
                    {"age_days": 28, "mould_no": "29",  "weight_gr": 8318, "load_kn": 1233.74, "strength_nmm2": 54.83},
                    {"age_days": 28, "mould_no": "21",  "weight_gr": 8383, "load_kn": 1305.96, "strength_nmm2": 58.04},
                    {"age_days": 28, "mould_no": "99",  "weight_gr": 8371, "load_kn": 1258.69, "strength_nmm2": 55.94},
                ],
            },
            # 2) Shotcrete single-set (5+5)
            {
                "cube_no": "501", "sample_mark": "G-CON-7001",
                "tests": [
                    {"age_days": 7,  "core_diameter_mm": 94, "core_height_mm": 188, "weight_gr": 2110, "load_kn": 275.20, "strength_nmm2": 39.60},
                    {"age_days": 7,  "core_diameter_mm": 94, "core_height_mm": 188, "weight_gr": 2098, "load_kn": 268.40, "strength_nmm2": 38.52},
                    {"age_days": 7,  "core_diameter_mm": 93, "core_height_mm": 187, "weight_gr": 2125, "load_kn": 281.10, "strength_nmm2": 40.89},
                    {"age_days": 7,  "core_diameter_mm": 94, "core_height_mm": 189, "weight_gr": 2085, "load_kn": 255.70, "strength_nmm2": 36.77},
                    {"age_days": 7,  "core_diameter_mm": 94, "core_height_mm": 188, "weight_gr": 2105, "load_kn": 272.80, "strength_nmm2": 39.12},
                    {"age_days": 28, "core_diameter_mm": 94, "core_height_mm": 188, "weight_gr": 2120, "load_kn": 345.50, "strength_nmm2": 49.75},
                    {"age_days": 28, "core_diameter_mm": 93, "core_height_mm": 187, "weight_gr": 2115, "load_kn": 352.80, "strength_nmm2": 51.24},
                    {"age_days": 28, "core_diameter_mm": 94, "core_height_mm": 188, "weight_gr": 2108, "load_kn": 338.20, "strength_nmm2": 48.60},
                    {"age_days": 28, "core_diameter_mm": 94, "core_height_mm": 189, "weight_gr": 2125, "load_kn": 360.10, "strength_nmm2": 52.03},
                    {"age_days": 28, "core_diameter_mm": 94, "core_height_mm": 188, "weight_gr": 2099, "load_kn": 342.00, "strength_nmm2": 49.10},
                ],
            },
            # 3) Shotcrete two-set (10+10) — splitter should produce 2 sub-cubes
            {
                "cube_no": "502", "sample_mark": "G-CON-7002",
                "tests":
                    [{"age_days": 7, "core_diameter_mm": 94, "core_height_mm": 188,
                      "weight_gr": 2100 + k, "load_kn": 270.0 + k, "strength_nmm2": 35.0 + k}
                     for k in range(5)] +
                    [{"age_days": 7, "core_diameter_mm": 94, "core_height_mm": 188,
                      "weight_gr": 2100 + k, "load_kn": 270.0 + k, "strength_nmm2": 42.0 + k}
                     for k in range(5)] +
                    [{"age_days": 28, "core_diameter_mm": 94, "core_height_mm": 188,
                      "weight_gr": 2100 + k, "load_kn": 340.0 + k, "strength_nmm2": 48.0 + k}
                     for k in range(5)] +
                    [{"age_days": 28, "core_diameter_mm": 94, "core_height_mm": 188,
                      "weight_gr": 2100 + k, "load_kn": 340.0 + k, "strength_nmm2": 52.0 + k}
                     for k in range(5)],
            },
        ]
    }
    return reader._split_multi_set_cubes(reader._process_shotcrete_cubes(cubes))


def main():
    cubes_data = _make_cubes()

    # Mock open-sheets so cubes "match" without Excel:
    # the normal cube resolves to 7000, the shotcrete ones to 7001 / 7002.
    sheets = [
        {"workbook": "(mock.xlsx)", "sheet": "sheet-7000",
         "sample_id_raw": "7000", "sample_id_num": 7000},
        {"workbook": "(mock.xlsx)", "sheet": "sheet-7001",
         "sample_id_raw": "7001", "sample_id_num": 7001},
        {"workbook": "(mock.xlsx)", "sheet": "sheet-7002-set1",
         "sample_id_raw": "7002", "sample_id_num": 7002},
        {"workbook": "(mock.xlsx)", "sheet": "sheet-7002-set2",
         "sample_id_raw": "7002", "sample_id_num": 7002},
    ]
    scan_result = {
        "sheets": sheets, "scanned_count": len(sheets),
        "start_sheet": "sheet-7000", "workbook": "(mock.xlsx)",
        "found_all": True,
    }

    # A placeholder image (real app shows the PDF page)
    img = Image.new("RGB", (800, 1100), "white")

    root = ctk.CTk()
    root.withdraw()

    # PreviewWindow.read_all_values is only called for NORMAL cubes.
    # To avoid it reaching Excel, unmatch the normal cube by clearing its
    # sheet so read_all_values is not invoked.
    # Simpler: monkey-patch writer.read_all_values to return empty cells.
    import writer
    writer.read_all_values = lambda wb, sh: {
        "weights_7d": [None]*3, "loads_7d": [None]*3,
        "weights_28d": [None]*3, "loads_28d": [None]*3,
    }
    writer.cross_check_7day = lambda *_a, **_k: []

    PreviewWindow(root, img, cubes_data, scan_result)
    root.mainloop()


if __name__ == "__main__":
    main()
