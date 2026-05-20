"""
reader.py — Reads a handwritten concrete-cube test notebook via Gemini Vision.
"""
import warnings

warnings.filterwarnings("ignore", category=FutureWarning)

import hashlib  # noqa: E402
import json  # noqa: E402
import os  # noqa: E402
import sys  # noqa: E402
from io import BytesIO  # noqa: E402
from pathlib import Path  # noqa: E402

import fitz  # noqa: E402  (PyMuPDF)
import google.generativeai as genai  # noqa: E402
from dotenv import load_dotenv  # noqa: E402
from PIL import Image  # noqa: E402

if getattr(sys, "frozen", False):
    _BASE_DIR = Path(sys.executable).parent
else:
    _BASE_DIR = Path(__file__).parent
load_dotenv(_BASE_DIR / ".env")

DEFAULT_GEMINI_MODEL = "gemini-2.5-flash"
FALLBACK_MODELS = ("gemini-2.0-flash", "gemini-2.5-pro")
MAX_IMAGE_SIDE = 2500


def _get_model_name() -> str:
    """Read model from env (set via Settings dialog), fall back to default."""
    return (os.getenv("GEMINI_MODEL") or "").strip() or DEFAULT_GEMINI_MODEL

PROMPT = """You are reading a handwritten concrete cube compressive strength test log.
The form has pre-printed English headers and handwritten entries in blue/black ink.

STRUCTURE:
- Each "Cube No" groups consecutive rows that share the same Site Sampling
  Mark (Sample ID).
- The typical case is 6 rows per cube: 3 rows at Age 7 and 3 rows at Age 28.
- BUT some samples have more — e.g. 12 rows (6 × Age 7 + 6 × Age 28) when the
  same sample is tested in two sets. Extract EVERY test row for that cube.
- A cube may also have ONLY 7-day rows (no 28-day yet) or ONLY 28-day rows.
  If a 7-day or 28-day group is simply not on the page for a cube, do NOT
  output any row for that missing group. Never fabricate rows.
- Each row represents one physical cube specimen. Include all readable rows
  in the cube's "tests" list with the correct age_days value (read it from
  the Age column on that row).

COLUMNS (left to right):
1. Cube No — 3-digit number (e.g. 378, 379)
2. Site Sampling Mark/No — e.g. G26-CON-395
3. Mould No — 1-3 digit number
4. Concrete Supplier — short code like "S2A BP", "CMD"
5. Site Location / Comments
6. Section — e.g. 2A, 2B, 1
7. Batch Ticket Number — 5-digit
8. C Grade — e.g. 35/45, 30/37
9. Date of Sampling — e.g. 12.03.26
10. Sampled by — initials
11. Date of Testing — e.g. 19.03.26 (7-day) or 09.04.26 (28-day)
12. Age (days) — 7 or 28
13. Weight (gr) — 4-digit integer, e.g. 8360, 8332
14. Load (kN) — decimal number, e.g. 1102.34, 1196.46
15. Compressive Strength (N/mm²) — decimal number, e.g. 48.99

SHOTCRETE / CORE FORM — occasional special case:
- This is a SEPARATE printed form titled "Core Record and Core
  Compressive Strength Test Forms" — note the word "Core". The normal
  cube form is titled "Concrete Sampling Log and Cube Compressive
  Strength Test Forms" — note the word "Cube". Both forms share the
  same form code (e.g. "T/02 - BEJV - MS - MKC - EN - 12390-3"); the
  ONLY reliable difference is "Core" vs "Cube" in the printed title.
- 5 rows per age instead of 3 (so 10 per set; 20 if two-set).
- Column 3 "Mould No" is replaced by TWO columns:
    3a. Core Diameter (mm)  — integer or decimal mm (e.g. 94, 93.5)
    3b. Core Height   (mm)  — integer or decimal mm (e.g. 95, 188)
- All other columns are the same. Extract core_diameter_mm and
  core_height_mm for core-form rows; null for normal-cube rows.
- A page is a CORE/shotcrete page ONLY if the printed form TITLE
  contains "Core" (e.g. "Core Record and Core Compressive Strength").
  A normal "Cube" page with extra handwritten rows is NOT a core page
  — never classify by row count alone, only by the printed title.

*** CRITICAL NOTATION STYLE ***
Some decimal numbers are written in a European shorthand where the fractional
part is written SMALLER and ELEVATED (superscript style). For example:
  - "48⁹⁹" means 48.99 (the "99" is smaller and above baseline)
  - "53⁴⁷" means 53.47
  - "1102³⁴" might mean 1102.34 if the "34" is elevated
Always reconstruct the proper decimal value. Small elevated digits are the
fractional part. If you see two groups of digits written close together with
the second group smaller/higher, interpret as: <whole>.<fraction>.

Normal decimals written with a visible dot (like "1102.34") are the same value.

TASK:
Extract EVERY cube and EVERY test row on the page into JSON.

OUTPUT FORMAT (JSON only, no explanation):
{
  "page_is_shotcrete": false,
  "cubes": [
    {
      "cube_no": "XXX",
      "sample_mark": "G26-CON-XXX",
      "tests": [
        {"age_days": 7,  "mould_no": "NNN", "weight_gr": 9999, "load_kn": 999.99, "strength_nmm2": 99.99, "core_diameter_mm": null, "core_height_mm": null}
      ]
    }
  ]
}

The top-level "page_is_shotcrete" field is true ONLY when the printed
form TITLE on the page contains the word "Core" (the "Core Record and
Core Compressive Strength Test Forms" form). Set it to false for
normal "Cube" pages ("...Cube Compressive Strength Test Forms"), even
if a cube has more than 3 rows per age.
(The placeholders above — 9999, 999.99, 99.99 — are NOT real. They only
show FIELD SHAPE. Do NOT echo them. Never emit 9999 / 999.99 / 99.99 in
your answer. Every numeric value must come from the handwritten entries
you can actually see on the page. If a value is unreadable, output null.)

RULES:
- weight_gr must be an integer (grams)
- load_kn must be a number with decimal (kN)
- strength_nmm2: decimal N/mm² (same superscript convention as load_kn).
- If you cannot read a value confidently, output null for that field. Do NOT guess.
- If an age group (7 or 28) has no rows on the page for a cube, just omit
  those rows entirely. Do not emit empty/null rows to pad the group.
- Output all cubes found on the page, in top-to-bottom order.
- Do not add cubes or tests that are not on the page.
- Output ONLY the JSON object, nothing else.
"""


def load_images(file_path: str) -> list[Image.Image]:
    """
    Load a file as a list of PIL Images.
      - PDF -> one image per page
      - JPG/PNG/etc -> single-element list
    """
    path = Path(file_path)
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        doc = fitz.open(str(path))
        # 200 DPI keeps handwriting legible; cap longest side so a big
        # source PDF can't blow the upload to 30+ MP.
        mat = fitz.Matrix(200 / 72, 200 / 72)
        images: list[Image.Image] = []
        for page in doc:
            pix = page.get_pixmap(matrix=mat)
            img = Image.open(BytesIO(pix.tobytes("png"))).convert("RGB")
            if max(img.size) > MAX_IMAGE_SIDE:
                ratio = MAX_IMAGE_SIDE / max(img.size)
                new_size = (int(img.width * ratio), int(img.height * ratio))
                img = img.resize(new_size, Image.LANCZOS)
            buf = BytesIO()
            img.save(buf, format="JPEG", quality=85, optimize=True)
            buf.seek(0)
            images.append(Image.open(buf))
        doc.close()
        return images
    elif suffix in (".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif"):
        return [Image.open(str(path)).convert("RGB")]
    else:
        raise ValueError(f"Unsupported file type: {suffix}")


def load_image(file_path: str) -> Image.Image:
    """Legacy single-image loader — returns the first page."""
    return load_images(file_path)[0]


def stack_images_vertically(
    images: list[Image.Image], gap: int = 20, bg: str = "white"
) -> Image.Image:
    """Combine a list of images into one tall image for preview."""
    if not images:
        raise ValueError("No images to stack")
    if len(images) == 1:
        return images[0]
    max_w = max(img.width for img in images)
    total_h = sum(img.height for img in images) + gap * (len(images) - 1)
    combined = Image.new("RGB", (max_w, total_h), bg)
    y = 0
    for img in images:
        # Center horizontally
        x = (max_w - img.width) // 2
        combined.paste(img, (x, y))
        y += img.height + gap
    return combined


def _log(msg: str) -> None:
    """Append a timestamped line to gemini_debug.log next to the exe."""
    try:
        from datetime import datetime
        line = f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n"
        with open(_BASE_DIR / "gemini_debug.log", "a", encoding="utf-8") as f:
            f.write(line)
    except Exception:
        pass


def _call_gemini_on_image(model_chain: list[str], image: Image.Image) -> dict:
    """Send one image to Gemini, trying models in order, return parsed JSON.

    On 503/429 (busy), falls over to the next model immediately. On
    DeadlineExceeded, retries up to 3x on the same model before falling over.
    """
    import time
    from google.api_core import exceptions as gax

    model_chain = [m.split("/")[-1] for m in model_chain]

    last_err = None
    for model_name in model_chain:
        model = genai.GenerativeModel(model_name)
        for attempt in range(3):
            try:
                t0 = time.time()
                _log(f"  [{model_name}] attempt {attempt+1}: sending image ({image.size})")
                response = model.generate_content(
                    [PROMPT, image],
                    generation_config={
                        "response_mime_type": "application/json",
                        "temperature": 0.0,
                    },
                    request_options={"timeout": 300},
                )
                _log(f"  [{model_name}] got response in {time.time()-t0:.1f}s")
                raw = response.text
                try:
                    data = json.loads(raw)
                except json.JSONDecodeError as e:
                    raise RuntimeError(f"Gemini JSON error: {e}\nResponse: {raw[:500]}")
                if "cubes" not in data or not isinstance(data["cubes"], list):
                    raise RuntimeError(f"Missing expected 'cubes' key: {raw[:500]}")
                return data
            except (gax.ServiceUnavailable, gax.ResourceExhausted) as e:
                last_err = e
                _log(f"  [{model_name}] busy ({type(e).__name__}); falling over to next model")
                break  # Don't burn retries on a busy model.
            except (gax.DeadlineExceeded, gax.InternalServerError) as e:
                last_err = e
                if attempt < 2:
                    time.sleep(2 ** attempt)
                    continue
                _log(f"  [{model_name}] timed out 3x; falling over to next model")
                break  # After 3 timeouts, try next model.
    raise RuntimeError(
        f"All models in fallback chain failed. Last error: {last_err}"
    )


def _is_blank(v) -> bool:
    """True for None or empty/whitespace string. Zero counts as a real value."""
    if v is None:
        return True
    if isinstance(v, str) and v.strip() == "":
        return True
    return False


def _test_is_empty(test: dict) -> bool:
    """A test row is empty if every payload field is blank — weight, load,
    strength, core diameter, core height. Such rows are almost certainly
    hallucinated by the model when it tries to pad an age group."""
    return all(
        _is_blank(test.get(k))
        for k in ("weight_gr", "load_kn", "strength_nmm2",
                  "core_diameter_mm", "core_height_mm")
    )


def _drop_hallucinated_groups(cubes_data: dict) -> dict:
    """Remove age groups where all rows share identical (weight, load,
    strength). Real concrete specimens always vary at least slightly;
    identical triplicates are a strong signal the model invented the
    group to pad an incomplete page.

    Also drop any row that matches the PLACEHOLDER values used in the
    prompt example (9999 / 999.99 / 99.99) — those are a direct echo."""
    PLACEHOLDER_W = 9999
    PLACEHOLDER_L = 999.99
    PLACEHOLDER_S = 99.99

    def _is_placeholder(t):
        try:
            return (
                (t.get("weight_gr") == PLACEHOLDER_W)
                or (float(t.get("load_kn") or 0) == PLACEHOLDER_L)
                or (float(t.get("strength_nmm2") or 0) == PLACEHOLDER_S)
            )
        except (TypeError, ValueError):
            return False

    for cube in cubes_data.get("cubes", []):
        tests = cube.get("tests", [])
        kept: list = []
        by_age: dict[int, list] = {}
        for t in tests:
            if _is_placeholder(t):
                continue
            by_age.setdefault(t.get("age_days"), []).append(t)

        for age, group in by_age.items():
            if len(group) >= 2:
                signatures = {
                    (t.get("weight_gr"), t.get("load_kn"),
                     t.get("strength_nmm2"))
                    for t in group
                }
                if len(signatures) == 1:
                    # Every row is a carbon copy → whole group is fake.
                    continue
            kept.extend(group)
        cube["tests"] = kept
    return cubes_data


def _clean_empty_rows(cubes_data: dict) -> dict:
    """Strip fully-empty tests and cubes with zero remaining tests.

    Run right after Gemini returns (before shotcrete detection) so that a
    hallucinated "age_days=28 with all nulls" row cannot trigger shotcrete
    detection or show up in the preview. Also normalizes empty strings
    ("" or whitespace) to None on numeric fields so downstream UI logic
    treats them uniformly."""
    numeric_fields = (
        "weight_gr", "load_kn", "strength_nmm2",
        "core_diameter_mm", "core_height_mm",
    )
    clean_cubes = []
    for cube in cubes_data.get("cubes", []):
        tests = []
        for t in cube.get("tests", []):
            for k in numeric_fields:
                if isinstance(t.get(k), str) and t[k].strip() == "":
                    t[k] = None
            if not _test_is_empty(t):
                tests.append(t)
        if not tests:
            continue
        cube["tests"] = tests
        clean_cubes.append(cube)
    return {**cubes_data, "cubes": clean_cubes}


def _strength_key(test: dict) -> float:
    """Sort key: higher strength first; missing strength ranks last."""
    s = test.get("strength_nmm2")
    try:
        return float(s)
    except (TypeError, ValueError):
        return float("-inf")


def _auto_pick_top3(group: list[dict]) -> None:
    """
    Tag `_selected` on each test in a single age-group:
    - 3 or fewer rows: all selected (normal-cube fallback).
    - 5 rows: top 3 by strength selected, bottom 2 unselected.
    - Multiples of 5 greater than 5 (10, 15, 20...): split into consecutive
      5-row chunks in row order (row order = set order on the form), each
      chunk picks its own top 3 and gets `_set_index` 1..N.
    - Any other count (e.g. 4, 6, 7): top 3 by strength. The splitter handles
      6-row multi-set normal cubes before this runs for shotcrete anyway; this
      is just a safe default.
    """
    n = len(group)
    if n == 0:
        return
    if n <= 3:
        for t in group:
            t["_selected"] = True
        return
    if n >= 10 and n % 5 == 0:
        for s in range(n // 5):
            chunk = group[s * 5:(s + 1) * 5]
            _auto_pick_top3(chunk)
            for t in chunk:
                t["_set_index"] = s + 1
        return
    # 5 rows (or any other count > 3 with no set boundary): top 3 by strength.
    # Rows with no strength value never get _selected.
    valid = [i for i in range(n) if _strength_key(group[i]) != float("-inf")]
    ranked = sorted(valid, key=lambda i: _strength_key(group[i]), reverse=True)
    top3 = set(ranked[:3])
    for i, t in enumerate(group):
        t["_selected"] = i in top3


def _process_shotcrete_cubes(cubes_data: dict) -> dict:
    """
    Mark every cube on a core/shotcrete page as shotcrete and pre-select
    the top 3 specimens per age by compressive strength. Sets
    `cube["_shotcrete"] = True` and a `_selected` flag on each test.

    GATE: the page TITLE is authoritative. A cube is shotcrete iff its
    source page was flagged `_shotcrete_page=True` by Gemini (the printed
    title contains "Core" — the "Core Record and Core Compressive Strength
    Test Forms" form). The row count is NOT used to gate: a core page may
    have any number of rows per cube (Gemini may misread 5 as 4 or 6), but
    if the title says "Core" the cube is shotcrete. This prevents core-page
    cubes from leaking into the normal multi-set splitter (which would
    wrongly split a 5-row group into "2 sets").

    For the 10-per-age two-set case `_auto_pick_top3` tags each test with
    `_set_index` (1 or 2) so the downstream splitter keeps the right rows
    together.
    """
    for cube in cubes_data.get("cubes", []):
        if not cube.get("_shotcrete_page"):
            continue
        cube["_shotcrete"] = True
        tests = cube.get("tests", [])
        t7 = [t for t in tests if t.get("age_days") == 7]
        t28 = [t for t in tests if t.get("age_days") == 28]
        _auto_pick_top3(t7)
        _auto_pick_top3(t28)
    return cubes_data


def _split_multi_set_cubes(cubes_data: dict) -> dict:
    """
    Split cubes with more than 3 tests at age 7 or 28 into multiple
    sub-cubes with the same sample_mark. Each sub-cube holds at most 3
    tests per age (so it fits one Excel sheet's slots). The matching
    logic in writer.match_cubes_to_sheets then distributes sub-cubes
    across the multiple sheets that share the sample ID.

    Cubes with ≤3 tests per age are returned unchanged (no behavior
    change for the common case).
    """
    out: list[dict] = []
    for cube in cubes_data.get("cubes", []):
        tests = cube.get("tests", [])
        tests_7 = [t for t in tests if t.get("age_days") == 7]
        tests_28 = [t for t in tests if t.get("age_days") == 28]
        other = [t for t in tests if t.get("age_days") not in (7, 28)]

        if cube.get("_shotcrete"):
            # Shotcrete: normal 3-at-a-time slicing would shred the 5-row
            # groups. Split by per-row _set_index (set by _auto_pick_top3
            # for 10-row groups); if no index is present it's a single set.
            set_indices = sorted(
                {t.get("_set_index") for t in tests
                 if t.get("_set_index") is not None}
            )
            if len(set_indices) <= 1:
                out.append(cube)
                continue
            total = len(set_indices)
            for sidx in set_indices:
                sub = {k: v for k, v in cube.items() if k != "tests"}
                sub_tests = [t for t in tests_7 if t.get("_set_index") == sidx]
                sub_tests += [t for t in tests_28 if t.get("_set_index") == sidx]
                if sidx == set_indices[0]:
                    sub_tests += other
                sub["tests"] = sub_tests
                sub["_set_index"] = sidx
                sub["_set_total"] = total
                out.append(sub)
            continue

        n_groups = max(
            1,
            -(-len(tests_7) // 3),   # ceil div
            -(-len(tests_28) // 3),
        )

        if n_groups <= 1:
            out.append(cube)
            continue

        for g in range(n_groups):
            sub = {k: v for k, v in cube.items() if k != "tests"}
            sub_tests: list = []
            sub_tests.extend(tests_7[g * 3:(g + 1) * 3])
            sub_tests.extend(tests_28[g * 3:(g + 1) * 3])
            if g == 0:
                # Unusual ages ride along with the first set
                sub_tests.extend(other)
            sub["tests"] = sub_tests
            sub["_set_index"] = g + 1
            sub["_set_total"] = n_groups
            out.append(sub)
    return {**cubes_data, "cubes": out}


def _cache_dir() -> Path:
    return _BASE_DIR / ".cache" / "gemini"


def _file_sha256(file_path: str) -> str:
    h = hashlib.sha256()
    with open(file_path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


# Bump when the Gemini PROMPT or cube post-processing changes shape;
# busts stale caches so old reads aren't reused with new logic.
_PROMPT_VERSION = "v3"


def _cache_path_for(digest: str, model_name: str) -> Path:
    safe_model = model_name.replace("/", "_").replace("\\", "_")
    return _cache_dir() / f"{digest}_{safe_model}_{_PROMPT_VERSION}.json"


def clear_gemini_cache() -> int:
    cache_dir = _cache_dir()
    if not cache_dir.exists():
        return 0
    removed = 0
    for p in cache_dir.iterdir():
        if p.is_file():
            try:
                p.unlink()
                removed += 1
            except OSError as e:
                print(f"[cache] failed to remove {p}: {e}", file=sys.stderr)
    return removed


def _fix_known_ocr_misreads(data: dict) -> dict:
    """Normalize cube sample marks that Gemini consistently misreads.

    Known pattern: a handwritten "G" at the start of the project prefix
    ("G26-CON-...") often gets read as the digit "6", producing
    "626-CON-...". Rewrite the sample_mark so every downstream consumer
    (preview cards, ledger match, debug output) sees the corrected
    string. Other prefixes are left untouched.
    """
    import re
    for cube in data.get("cubes", []) or []:
        mark = cube.get("sample_mark")
        if not isinstance(mark, str):
            continue
        new = re.sub(r"^626-", "G26-", mark)
        if new != mark:
            cube["sample_mark"] = new
            try:
                _log(f"[OCR-FIX] '{mark}' -> '{new}'")
            except Exception:
                pass
    return data


def _postprocess_cubes(data: dict) -> dict:
    """OCR misread fix → hallucination-drop → empty-row clean →
    shotcrete top-3 → multi-set split."""
    return _split_multi_set_cubes(
        _process_shotcrete_cubes(
            _clean_empty_rows(
                _drop_hallucinated_groups(
                    _fix_known_ocr_misreads(data)
                )
            )
        )
    )


def read_notebook(file_path: str, progress_cb=None) -> dict:
    """
    Send every page of a PDF (or the single image) to Gemini and return
    a merged dict with all cubes across all pages.

    progress_cb(stage: str, current: int, total: int) is called at key
    points: ("cache_hit", 0, 0), ("loading", 0, 0), ("page", i, n).
    """
    def _emit(stage, cur=0, tot=0):
        if progress_cb is not None:
            try:
                progress_cb(stage, cur, tot)
            except Exception:
                pass

    model_name = _get_model_name()
    digest = _file_sha256(file_path)
    cache_file = _cache_path_for(digest, model_name)

    if cache_file.exists():
        try:
            with open(cache_file, "r", encoding="utf-8") as f:
                cached = json.load(f)
            _emit("cache_hit")
            return _postprocess_cubes(cached)
        except (OSError, json.JSONDecodeError) as e:
            print(f"[cache] failed to read {cache_file}: {e}", file=sys.stderr)

    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise RuntimeError(
            "GEMINI_API_KEY not found. Check the .env file."
        )

    genai.configure(api_key=api_key)
    _emit("loading")
    images = load_images(file_path)

    model_chain = list(dict.fromkeys([model_name, *FALLBACK_MODELS]))

    n = len(images)
    _log(f"=== read_notebook: {n} pages, chain={model_chain}, file={Path(file_path).name} (parallel)")
    merged: dict = {"cubes": []}

    from concurrent.futures import ThreadPoolExecutor, as_completed

    def _run_page(idx_image):
        idx, image = idx_image
        _log(f"page {idx}/{n} starting...")
        data = _call_gemini_on_image(model_chain, image)
        _log(f"page {idx}/{n} done — {len(data.get('cubes', []))} cubes")
        return idx, data

    indexed = list(enumerate(images, start=1))
    results: dict[int, dict] = {}
    with ThreadPoolExecutor(max_workers=min(5, n)) as pool:
        futures = [pool.submit(_run_page, item) for item in indexed]
        for done_count, fut in enumerate(as_completed(futures), start=1):
            idx, data = fut.result()
            results[idx] = data
            _emit("page", done_count, n)

    # Preserve original page order so shotcrete set-splitting stays consistent.
    for idx in sorted(results.keys()):
        page_data = results[idx]
        page_is_shot = bool(page_data.get("page_is_shotcrete", False))
        for cube in page_data.get("cubes", []):
            cube["_page"] = idx
            cube["_shotcrete_page"] = page_is_shot
            merged["cubes"].append(cube)
    _log(f"=== all pages done, total {len(merged['cubes'])} cubes")

    try:
        _cache_dir().mkdir(parents=True, exist_ok=True)
        with open(cache_file, "w", encoding="utf-8") as f:
            json.dump(merged, f, ensure_ascii=False, indent=2)
    except OSError as e:
        print(f"[cache] failed to write {cache_file}: {e}", file=sys.stderr)

    return _postprocess_cubes(merged)


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python reader.py <pdf_or_image>")
        sys.exit(1)
    result = read_notebook(sys.argv[1])
    print(json.dumps(result, indent=2, ensure_ascii=False))
