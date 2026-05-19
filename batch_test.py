"""Headless test: run both PDFs through reader to see if core pipeline works."""
import traceback
import reader

PDFS = [
    r"C:\Users\Yafka\Desktop\7days .pdf",
    r"C:\Users\Yafka\Desktop\28days 14.04.2026.pdf",
]

def progress(stage, cur=0, tot=0):
    print(f"  [progress] {stage} {cur}/{tot}")

for i, path in enumerate(PDFS, 1):
    print(f"\n=== FILE {i}/{len(PDFS)}: {path} ===")
    try:
        data = reader.read_notebook(path, progress_cb=progress)
        cubes = data.get("cubes", [])
        print(f"  OK -> {len(cubes)} cubes")
        for c in cubes[:3]:
            print(f"    {c.get('sample_no')} cube={c.get('cube_no')}")
    except Exception as e:
        print(f"  ERROR: {e}")
        traceback.print_exc()

print("\n=== DONE ===")
