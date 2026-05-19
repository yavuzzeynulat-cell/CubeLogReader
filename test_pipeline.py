"""
Test the writer against a real Excel file.
This script: opens Excel, calls the writer module, checks correctness, closes.
"""
import json
import os
import sys
import time

import win32com.client
from pythoncom import com_error

import reader
import writer

TEST_XLSX = os.path.abspath("_test_copy.xlsx")
TEST_PDF = r"C:\Users\Yafka\Pictures\Camera Roll\28days 09.04.2026.pdf"


def main():
    print(f"Test Excel: {TEST_XLSX}")
    print(f"Test PDF:   {TEST_PDF}")

    # 1) Read the notebook page with Gemini
    print("\n[1] Gemini is reading the notebook...")
    data = reader.read_notebook(TEST_PDF)
    n = len(data.get("cubes", []))
    print(f"    -> {n} cubes read")

    # 2) Open Excel
    print("\n[2] Opening Excel...")
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(TEST_XLSX)
    print(f"    -> workbook: {wb.Name}")
    print(f"    -> sheet count: {wb.Worksheets.Count}")
    for ws in wb.Worksheets:
        b14 = ws.Range("B14").Value
        print(f"       sheet '{ws.Name}': B14 = {b14!r}")

    # 3) Match cubes to sheets
    print("\n[3] Matching cubes to sheets...")
    matched = writer.match_cubes_to_sheets(data)
    for m in matched:
        mark = m["cube"].get("sample_mark")
        sh = m["matched_sheet"]
        if sh:
            print(f"    OK  {mark} -> {sh['workbook']} / {sh['sheet']}")
        else:
            print(f"    --  {mark} -> no match")

    # 4) Write 28-day values for the first matched cube only (test)
    print("\n[4] Writing 28-day values for the first matched cube...")
    first_match = next((m for m in matched if m["matched_sheet"]), None)
    if not first_match:
        print("    ERROR: no matched cubes!")
        wb.Close(SaveChanges=False)
        excel.Quit()
        return 1

    cube = first_match["cube"]
    sh = first_match["matched_sheet"]
    tests_28 = [t for t in cube["tests"] if t.get("age_days") == 28]
    weights = [t.get("weight_gr") for t in tests_28[:3]]
    loads = [t.get("load_kn") for t in tests_28[:3]]
    print(f"    Will write: weights={weights}, loads={loads}")

    # Read existing values first
    ws_test = wb.Worksheets(sh["sheet"])
    before = {
        "W20": ws_test.Range("W20").Value,
        "W21": ws_test.Range("W21").Value,
        "W22": ws_test.Range("W22").Value,
        "AA20": ws_test.Range("AA20").Value,
        "AA21": ws_test.Range("AA21").Value,
        "AA22": ws_test.Range("AA22").Value,
    }
    print(f"    BEFORE: {before}")

    result = writer.write_cube_28day(
        cube, sh["workbook"], sh["sheet"], weights, loads
    )
    print(f"    write_cube_28day result: {len(result['wrote'])} cells, {len(result['errors'])} errors")
    for err in result["errors"]:
        print(f"      ERROR: {err}")

    after = {
        "W20": ws_test.Range("W20").Value,
        "W21": ws_test.Range("W21").Value,
        "W22": ws_test.Range("W22").Value,
        "AA20": ws_test.Range("AA20").Value,
        "AA21": ws_test.Range("AA21").Value,
        "AA22": ws_test.Range("AA22").Value,
    }
    print(f"    AFTER: {after}")

    # Are AB (compressive strength) and AC (average) auto-computed?
    print("\n[5] Checking auto-computed cells (AB20-22, AC20):")
    for c in ["AB20", "AB21", "AB22", "AC20"]:
        print(f"    {c} = {ws_test.Range(c).Value}")

    # Cleanup: close WITHOUT saving changes
    print("\n[6] Cleanup (close without saving)...")
    wb.Close(SaveChanges=False)
    excel.Quit()
    time.sleep(0.5)
    print("    OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
