"""
test_mark_fix_persistence.py — corrected Sample Marks survive a re-read.

The Gemini read is cached on disk by file digest, so reopening the same
notebook replays the original misread mark. A correction that only lived in
memory was lost every time, and the user had to retype it.

Fixes are stored against the mark as Gemini first read it (`_orig_mark`), not
against whatever the cube currently carries — correcting the same cube twice
then overwrites one entry instead of chaining 1798 -> 1198 -> 1200.
"""
import json
import shutil
import tempfile
from pathlib import Path

import reader


def _cube(mark, cube_no="109"):
    return {"sample_mark": mark, "cube_no": cube_no, "tests": []}


class _TempCache:
    """Point reader's .cache at a throwaway directory."""

    def __enter__(self):
        self.dir = Path(tempfile.mkdtemp(prefix="cubelog-test-"))
        self._old = reader._BASE_DIR
        reader._BASE_DIR = self.dir
        return self.dir

    def __exit__(self, *exc):
        reader._BASE_DIR = self._old
        shutil.rmtree(self.dir, ignore_errors=True)
        return False


def _read_back(digest, cubes):
    """Simulate a fresh read of the same file: original marks, fixes applied."""
    data = {"_source_digest": digest, "cubes": cubes}
    for c in data["cubes"]:
        c.setdefault("_orig_mark", c.get("sample_mark"))
    return reader.apply_saved_mark_fixes(data)


def test_a_saved_fix_comes_back_on_the_next_read():
    with _TempCache():
        data = {"_source_digest": "abc", "cubes": [_cube("G26-CON-1798")]}
        data["cubes"][0]["_orig_mark"] = "G26-CON-1798"
        data["cubes"][0]["sample_mark"] = "G26-CON-1198"
        reader.save_mark_fixes(data)

        fresh = _read_back("abc", [_cube("G26-CON-1798")])
        assert fresh["cubes"][0]["sample_mark"] == "G26-CON-1198"


def test_another_notebook_is_unaffected():
    with _TempCache():
        data = {"_source_digest": "abc", "cubes": [_cube("G26-CON-1798")]}
        data["cubes"][0]["_orig_mark"] = "G26-CON-1798"
        data["cubes"][0]["sample_mark"] = "G26-CON-1198"
        reader.save_mark_fixes(data)

        # A different page that genuinely reads 1798 keeps its mark.
        other = _read_back("zzz", [_cube("G26-CON-1798")])
        assert other["cubes"][0]["sample_mark"] == "G26-CON-1798"


def test_correcting_twice_overwrites_instead_of_chaining():
    with _TempCache():
        cubes = [_cube("G26-CON-1798")]
        cubes[0]["_orig_mark"] = "G26-CON-1798"
        data = {"_source_digest": "abc", "cubes": cubes}

        cubes[0]["sample_mark"] = "G26-CON-1198"
        reader.save_mark_fixes(data)
        cubes[0]["sample_mark"] = "G26-CON-1200"   # user corrects again
        reader.save_mark_fixes(data)

        fresh = _read_back("abc", [_cube("G26-CON-1798")])
        assert fresh["cubes"][0]["sample_mark"] == "G26-CON-1200"


def test_every_cube_sharing_the_misread_mark_is_restored():
    with _TempCache():
        cubes = [_cube("G26-CON-1798", "109"), _cube("G26-CON-1798", "110")]
        for c in cubes:
            c["_orig_mark"] = "G26-CON-1798"
            c["sample_mark"] = "G26-CON-1198"
        reader.save_mark_fixes({"_source_digest": "abc", "cubes": cubes})

        fresh = _read_back(
            "abc", [_cube("G26-CON-1798", "109"), _cube("G26-CON-1798", "110")])
        assert [c["sample_mark"] for c in fresh["cubes"]] == \
            ["G26-CON-1198", "G26-CON-1198"]


def test_an_uncorrected_cube_stores_nothing():
    with _TempCache() as d:
        cubes = [_cube("G26-CON-1200")]
        cubes[0]["_orig_mark"] = "G26-CON-1200"
        reader.save_mark_fixes({"_source_digest": "abc", "cubes": cubes})
        stored = json.loads(reader._mark_fix_path().read_text("utf-8")) \
            if reader._mark_fix_path().exists() else {}
        assert not stored.get("abc"), stored


def test_a_corrupt_store_does_not_break_the_read():
    with _TempCache():
        reader._mark_fix_path().parent.mkdir(parents=True, exist_ok=True)
        reader._mark_fix_path().write_text("{not json", encoding="utf-8")
        fresh = _read_back("abc", [_cube("G26-CON-1798")])
        assert fresh["cubes"][0]["sample_mark"] == "G26-CON-1798"


def test_the_stamp_survives_post_processing():
    # _orig_mark is stamped on the raw read, before the clean-up and the
    # multi-set split rebuild the cube dicts. If those drop it, every stored
    # fix silently stops applying.
    raw = {"cubes": [
        {"sample_mark": "G26-CON-1798", "cube_no": "109", "_orig_mark": "G26-CON-1798",
         "tests": [{"age_days": 7, "mould_no": "1", "weight_gr": 8300, "load_kn": 1200},
                   {"age_days": 28, "mould_no": "2", "weight_gr": 8400, "load_kn": 1300}]},
    ]}
    out = reader._postprocess_cubes(raw)
    assert out["cubes"], "the cube was dropped entirely"
    assert all(c.get("_orig_mark") == "G26-CON-1798" for c in out["cubes"])


def test_a_read_without_a_digest_is_left_alone():
    with _TempCache():
        data = {"cubes": [_cube("G26-CON-1798")]}
        out = reader.apply_saved_mark_fixes(data)
        assert out["cubes"][0]["sample_mark"] == "G26-CON-1798"


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
