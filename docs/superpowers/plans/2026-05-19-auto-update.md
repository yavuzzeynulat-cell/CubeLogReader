# CubeLogReader Auto-Update Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add antivirus-friendly auto-update to CubeLogReader so source `.py` files can be remotely updated via GitHub releases without changing the bundled exe.

**Architecture:** Replace PyInstaller entry point with a tiny `launcher.py` that loads application source from an external `src/` folder. Updates download a small zip from GitHub Releases, atomically swap `src/` with the new files, and keep the previous version as `src_backup/` for automatic rollback if the new version fails on next launch.

**Tech Stack:** Python 3 stdlib (`urllib`, `zipfile`, `hashlib`, `shutil`, `subprocess`, `threading`, `unittest`), CustomTkinter (existing UI), PyInstaller `--onedir` (existing build), Inno Setup (existing installer), GitHub Releases + `gh` CLI.

**Reference spec:** `docs/superpowers/specs/2026-05-19-auto-update-design.md`

---

## File Structure

**Project root (`C:\Users\Yafka\Desktop\CubeLogReader\`):**

| File | Status | Responsibility |
|------|--------|----------------|
| `main.py` | modify | Add startup update-check thread + SettingsDialog "Check for updates" button |
| `reader.py` | unchanged | — |
| `writer.py` | unchanged | — |
| `updater.py` | **create** | Update check / download / apply / restart logic |
| `version.txt` | **create** | Current version string (e.g., `1.0.0`) |
| `launcher.py` | **create** | PyInstaller entry; loads `src/`, rolls back on failure |
| `test_updater.py` | **create** | Unit tests for `updater.py` (uses stdlib `unittest`) |
| `CubeLogReader.spec` | modify | Entry → `launcher.py`; force-import deps; exclude src modules |
| `build_exe.bat` | modify | Post-build: copy app `.py` + `version.txt` to `dist/CubeLogReader/src/` |
| `installer.iss` | unchanged | `recursesubdirs` flag already picks up `src/` automatically |

**Runtime layout (installed app):**

```
CubeLogReader/
├── CubeLogReader.exe       (launcher; PyInstaller-built; never updated)
├── _internal/              (PyInstaller deps; never updated)
├── src/                    (UPDATABLE)
│   ├── main.py
│   ├── reader.py
│   ├── writer.py
│   ├── updater.py
│   └── version.txt
├── src_backup/             (auto-created on update; for rollback)
└── .env                    (user's API key)
```

---

## Conventions

- **Python module:** stdlib `unittest` (no new dependency). Run with `python -m unittest test_updater -v`.
- **Version format:** `MAJOR.MINOR.PATCH` plain string in `version.txt` (no `v` prefix). Git/release tag uses `vMAJOR.MINOR.PATCH`.
- **GitHub repo:** placeholder `OWNER/REPO` throughout this plan — replace with real `<github_user>/CubeLogReader` once repo is created.
- **Release zip name:** `src.zip` — contains all files inside `src/` flat (no top-level folder).
- **Release notes SHA256 line:** `SHA256: <hex>` on its own line; updater parses this.
- **Working directory for build:** project root.
- **Commits:** small, frequent; one per task at minimum.

---

## Task 1: Create `version.txt` and verify baseline build

**Goal:** Establish version-file convention; confirm current build still works before any changes.

**Files:**
- Create: `version.txt`

- [ ] **Step 1: Create `version.txt`**

```
1.0.0
```

(Single line, no trailing newline strictly required but harmless.)

- [ ] **Step 2: Verify current build still works**

Run: `build_exe.bat`
Expected: build completes; `dist\CubeLogReader\CubeLogReader.exe` exists; double-click opens app normally.

- [ ] **Step 3: Commit**

```bash
git add version.txt
git commit -m "chore: add version.txt baseline (1.0.0)"
```

---

## Task 2: `updater.check_for_update()` with tests

**Goal:** Implement and TDD the GitHub release polling. Function returns parsed `UpdateInfo` if remote version > local, else `None`. Network errors → `None` (silent).

**Files:**
- Create: `updater.py`
- Create: `test_updater.py`

- [ ] **Step 1: Write the failing tests**

Create `test_updater.py`:

```python
"""Unit tests for updater.py — uses unittest + tempfile, no network."""
import json
import os
import sys
import tempfile
import unittest
from unittest.mock import patch, MagicMock

# Add project root so we can import updater.py directly
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import updater


def _mock_release(tag="v1.0.1", body="Changes:\n- Bug fix\n\nSHA256: abc123",
                  asset_name="src.zip", asset_url="https://example.com/src.zip"):
    return {
        "tag_name": tag,
        "body": body,
        "assets": [{"name": asset_name, "browser_download_url": asset_url}],
    }


class CheckForUpdateTests(unittest.TestCase):

    def setUp(self):
        # Patch the local-version reader so tests don't depend on disk
        self._patcher = patch.object(updater, "_read_local_version", return_value="1.0.0")
        self._patcher.start()

    def tearDown(self):
        self._patcher.stop()

    @patch("updater._fetch_latest_release_json")
    def test_returns_update_info_when_remote_newer(self, mock_fetch):
        mock_fetch.return_value = _mock_release(tag="v1.0.1")
        info = updater.check_for_update()
        self.assertIsNotNone(info)
        self.assertEqual(info.version, "1.0.1")
        self.assertEqual(info.asset_url, "https://example.com/src.zip")
        self.assertEqual(info.sha256, "abc123")
        self.assertIn("Bug fix", info.notes)

    @patch("updater._fetch_latest_release_json")
    def test_returns_none_when_same_version(self, mock_fetch):
        mock_fetch.return_value = _mock_release(tag="v1.0.0")
        self.assertIsNone(updater.check_for_update())

    @patch("updater._fetch_latest_release_json")
    def test_returns_none_when_older_remote(self, mock_fetch):
        mock_fetch.return_value = _mock_release(tag="v0.9.0")
        self.assertIsNone(updater.check_for_update())

    @patch("updater._fetch_latest_release_json")
    def test_returns_none_on_network_error(self, mock_fetch):
        mock_fetch.side_effect = OSError("network down")
        self.assertIsNone(updater.check_for_update())

    @patch("updater._fetch_latest_release_json")
    def test_returns_none_on_bad_json(self, mock_fetch):
        mock_fetch.return_value = {"unexpected": "shape"}
        self.assertIsNone(updater.check_for_update())

    @patch("updater._fetch_latest_release_json")
    def test_handles_missing_sha256_line(self, mock_fetch):
        mock_fetch.return_value = _mock_release(body="Just a fix, no checksum")
        info = updater.check_for_update()
        self.assertIsNotNone(info)
        self.assertIsNone(info.sha256)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run tests to verify they fail with `ImportError`**

Run: `python -m unittest test_updater -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'updater'` (or import error since `updater.py` doesn't exist yet).

- [ ] **Step 3: Implement minimal `updater.py`**

Create `updater.py`:

```python
"""
updater.py — Auto-update logic for CubeLogReader.

Public API:
    check_for_update(timeout=5) -> Optional[UpdateInfo]
    download_update(info, dest_path) -> bool
    apply_update(zip_path) -> None
    restart_app() -> NoReturn

All network failures are silenced (return None / False) so the app
keeps running even when offline.
"""
from __future__ import annotations

import json
import os
import re
import sys
import urllib.request
from dataclasses import dataclass
from typing import Optional

# === CONFIG: change OWNER/REPO once the GitHub repo is created ===
GITHUB_OWNER = "OWNER"
GITHUB_REPO = "REPO"
RELEASE_API = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
ASSET_NAME = "src.zip"


@dataclass
class UpdateInfo:
    version: str          # e.g. "1.0.1" (no "v" prefix)
    notes: str            # raw release body (Turkish, multi-line)
    asset_url: str        # direct download URL for src.zip
    sha256: Optional[str] # parsed from notes; None if absent


def _app_dir() -> str:
    """Folder containing the running exe (or main.py during dev)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _read_local_version() -> str:
    """Read version.txt from src/ (frozen) or alongside this file (dev)."""
    candidates = [
        os.path.join(_app_dir(), "src", "version.txt"),
        os.path.join(_app_dir(), "version.txt"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "version.txt"),
    ]
    for p in candidates:
        if os.path.isfile(p):
            try:
                return open(p, "r", encoding="utf-8").read().strip()
            except OSError:
                pass
    return "0.0.0"


def _parse_version(s: str) -> tuple:
    """Parse '1.0.1' into (1,0,1). Invalid → (0,0,0)."""
    try:
        return tuple(int(x) for x in s.strip().lstrip("v").split("."))
    except (ValueError, AttributeError):
        return (0, 0, 0)


def _fetch_latest_release_json(timeout: int = 5) -> dict:
    """GET the latest release. Raises OSError on network failure."""
    req = urllib.request.Request(
        RELEASE_API,
        headers={"Accept": "application/vnd.github+json",
                 "User-Agent": "CubeLogReader-Updater"},
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _parse_sha256(body: str) -> Optional[str]:
    """Find 'SHA256: <hex>' anywhere in release notes. Case-insensitive."""
    if not body:
        return None
    m = re.search(r"SHA256\s*[:=]\s*([0-9a-fA-F]{64}|[0-9a-fA-F]+)", body)
    return m.group(1).lower() if m else None


def check_for_update(timeout: int = 5) -> Optional[UpdateInfo]:
    """Return UpdateInfo if a newer release exists; else None."""
    try:
        data = _fetch_latest_release_json(timeout=timeout)
        tag = data.get("tag_name")
        if not tag:
            return None
        remote_v = tag.lstrip("v")
        if _parse_version(remote_v) <= _parse_version(_read_local_version()):
            return None
        # Find the src.zip asset URL
        asset_url = None
        for a in data.get("assets", []):
            if a.get("name") == ASSET_NAME:
                asset_url = a.get("browser_download_url")
                break
        if not asset_url:
            return None
        return UpdateInfo(
            version=remote_v,
            notes=data.get("body", "") or "",
            asset_url=asset_url,
            sha256=_parse_sha256(data.get("body", "")),
        )
    except Exception:
        return None
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m unittest test_updater -v`
Expected: All 6 tests PASS.

- [ ] **Step 5: Commit**

```bash
git add updater.py test_updater.py
git commit -m "feat(updater): add check_for_update with tests"
```

---

## Task 3: `updater.download_update()` with tests

**Goal:** Download asset to a destination path, verify SHA256 (if provided) and zip integrity. Return `True` on success, `False` on any failure.

**Files:**
- Modify: `updater.py` (append new function)
- Modify: `test_updater.py` (append test class)

- [ ] **Step 1: Write the failing tests**

Append to `test_updater.py` (before `if __name__ == "__main__"`):

```python
import hashlib
import zipfile


def _make_test_zip(path, files):
    """Create a valid zip with given {name: content} dict."""
    with zipfile.ZipFile(path, "w") as zf:
        for name, content in files.items():
            zf.writestr(name, content)


def _sha256_of(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


class DownloadUpdateTests(unittest.TestCase):

    def setUp(self):
        self.tmpdir = tempfile.TemporaryDirectory()
        self.dest = os.path.join(self.tmpdir.name, "src.zip")
        # Create a "remote" valid zip in a temp file we'll point urlretrieve at
        self.remote = os.path.join(self.tmpdir.name, "remote_src.zip")
        _make_test_zip(self.remote, {"main.py": "print('hi')", "version.txt": "1.0.1"})
        self.remote_sha = _sha256_of(self.remote)

    def tearDown(self):
        self.tmpdir.cleanup()

    def _info(self, sha=None):
        return updater.UpdateInfo(
            version="1.0.1", notes="", asset_url="https://example.com/src.zip",
            sha256=sha,
        )

    @patch("updater._http_download")
    def test_download_success_with_correct_sha(self, mock_dl):
        mock_dl.side_effect = lambda url, dest, timeout=30: \
            __import__("shutil").copy(self.remote, dest)
        ok = updater.download_update(self._info(sha=self.remote_sha), self.dest)
        self.assertTrue(ok)
        self.assertTrue(os.path.isfile(self.dest))

    @patch("updater._http_download")
    def test_download_succeeds_when_sha_absent(self, mock_dl):
        mock_dl.side_effect = lambda url, dest, timeout=30: \
            __import__("shutil").copy(self.remote, dest)
        ok = updater.download_update(self._info(sha=None), self.dest)
        self.assertTrue(ok)

    @patch("updater._http_download")
    def test_download_fails_on_sha_mismatch(self, mock_dl):
        mock_dl.side_effect = lambda url, dest, timeout=30: \
            __import__("shutil").copy(self.remote, dest)
        ok = updater.download_update(self._info(sha="0" * 64), self.dest)
        self.assertFalse(ok)
        self.assertFalse(os.path.isfile(self.dest))  # bad file removed

    @patch("updater._http_download")
    def test_download_fails_on_corrupt_zip(self, mock_dl):
        def write_garbage(url, dest, timeout=30):
            with open(dest, "wb") as f:
                f.write(b"not a zip")
        mock_dl.side_effect = write_garbage
        ok = updater.download_update(self._info(sha=None), self.dest)
        self.assertFalse(ok)

    @patch("updater._http_download")
    def test_download_fails_on_network_error(self, mock_dl):
        mock_dl.side_effect = OSError("connection refused")
        ok = updater.download_update(self._info(), self.dest)
        self.assertFalse(ok)
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m unittest test_updater -v`
Expected: 5 new tests FAIL with `AttributeError: module 'updater' has no attribute 'download_update'`.

- [ ] **Step 3: Implement `download_update`**

Append to `updater.py`:

```python
import hashlib
import zipfile


def _http_download(url: str, dest: str, timeout: int = 30) -> None:
    """Stream-download `url` to `dest`. Raises OSError on failure."""
    req = urllib.request.Request(url, headers={"User-Agent": "CubeLogReader-Updater"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        with open(dest, "wb") as f:
            while True:
                chunk = resp.read(64 * 1024)
                if not chunk:
                    break
                f.write(chunk)


def _sha256_file(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(64 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def download_update(info: "UpdateInfo", dest_path: str, timeout: int = 30) -> bool:
    """Download src.zip, verify SHA256 (if given) and zip integrity.
    Returns True on success. On failure, removes any partial file."""
    try:
        _http_download(info.asset_url, dest_path, timeout=timeout)
    except Exception:
        if os.path.isfile(dest_path):
            try:
                os.remove(dest_path)
            except OSError:
                pass
        return False

    # SHA256 check (optional)
    if info.sha256:
        actual = _sha256_file(dest_path)
        if actual.lower() != info.sha256.lower():
            try:
                os.remove(dest_path)
            except OSError:
                pass
            return False

    # Zip integrity
    try:
        with zipfile.ZipFile(dest_path, "r") as zf:
            bad = zf.testzip()
            if bad is not None:
                raise zipfile.BadZipFile(f"corrupt entry: {bad}")
    except Exception:
        try:
            os.remove(dest_path)
        except OSError:
            pass
        return False

    return True
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m unittest test_updater -v`
Expected: All 11 tests PASS.

- [ ] **Step 5: Commit**

```bash
git add updater.py test_updater.py
git commit -m "feat(updater): add download_update with sha256 + zip integrity"
```

---

## Task 4: `updater.apply_update()` with tests

**Goal:** Replace existing `src/` with the new zip contents. Move old `src/` → `src_backup/` first (deleting any prior backup). Atomic-ish: rename is fast, extract failures still leave backup recoverable by launcher.

**Files:**
- Modify: `updater.py`
- Modify: `test_updater.py`

- [ ] **Step 1: Write the failing tests**

Append to `test_updater.py`:

```python
import shutil


class ApplyUpdateTests(unittest.TestCase):

    def setUp(self):
        self.tmpdir = tempfile.TemporaryDirectory()
        self.app_dir = self.tmpdir.name
        self.src = os.path.join(self.app_dir, "src")
        self.backup = os.path.join(self.app_dir, "src_backup")
        # Existing src/ with old content
        os.makedirs(self.src)
        with open(os.path.join(self.src, "main.py"), "w") as f:
            f.write("# old main")
        with open(os.path.join(self.src, "version.txt"), "w") as f:
            f.write("1.0.0")
        # New zip with updated content
        self.zip_path = os.path.join(self.app_dir, "update.zip")
        _make_test_zip(self.zip_path, {
            "main.py": "# new main",
            "reader.py": "# new reader",
            "version.txt": "1.0.1",
        })
        # Patch _app_dir so updater uses our temp dir
        self._patcher = patch.object(updater, "_app_dir", return_value=self.app_dir)
        self._patcher.start()

    def tearDown(self):
        self._patcher.stop()
        self.tmpdir.cleanup()

    def test_apply_moves_old_src_to_backup(self):
        updater.apply_update(self.zip_path)
        self.assertTrue(os.path.isdir(self.backup))
        self.assertEqual(open(os.path.join(self.backup, "main.py")).read(), "# old main")

    def test_apply_extracts_new_files_into_src(self):
        updater.apply_update(self.zip_path)
        self.assertEqual(open(os.path.join(self.src, "main.py")).read(), "# new main")
        self.assertEqual(open(os.path.join(self.src, "reader.py")).read(), "# new reader")
        self.assertEqual(open(os.path.join(self.src, "version.txt")).read(), "1.0.1")

    def test_apply_removes_pre_existing_backup(self):
        os.makedirs(self.backup)
        with open(os.path.join(self.backup, "stale.txt"), "w") as f:
            f.write("stale")
        updater.apply_update(self.zip_path)
        # The stale file from the old backup must be gone (replaced by old src)
        self.assertFalse(os.path.isfile(os.path.join(self.backup, "stale.txt")))
        self.assertTrue(os.path.isfile(os.path.join(self.backup, "main.py")))

    def test_apply_no_existing_src(self):
        # First-ever update scenario (shouldn't normally happen but be safe)
        shutil.rmtree(self.src)
        updater.apply_update(self.zip_path)
        self.assertTrue(os.path.isfile(os.path.join(self.src, "main.py")))
        self.assertFalse(os.path.isdir(self.backup))  # nothing to back up
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m unittest test_updater -v`
Expected: 4 new tests FAIL with `AttributeError: module 'updater' has no attribute 'apply_update'`.

- [ ] **Step 3: Implement `apply_update`**

Append to `updater.py`:

```python
import shutil


def apply_update(zip_path: str) -> None:
    """Replace src/ with the contents of zip_path. Move old src/ → src_backup/.

    Raises OSError if extraction fails (caller decides what to do).
    """
    app_dir = _app_dir()
    src = os.path.join(app_dir, "src")
    backup = os.path.join(app_dir, "src_backup")

    # 1. Clear any stale backup
    if os.path.isdir(backup):
        shutil.rmtree(backup)

    # 2. Move existing src → backup (if src exists)
    if os.path.isdir(src):
        os.rename(src, backup)

    # 3. Create fresh src and extract
    os.makedirs(src, exist_ok=True)
    try:
        with zipfile.ZipFile(zip_path, "r") as zf:
            zf.extractall(src)
    except Exception:
        # Extraction failed; try to restore backup
        if os.path.isdir(src):
            shutil.rmtree(src, ignore_errors=True)
        if os.path.isdir(backup):
            os.rename(backup, src)
        raise
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m unittest test_updater -v`
Expected: All 15 tests PASS.

- [ ] **Step 5: Commit**

```bash
git add updater.py test_updater.py
git commit -m "feat(updater): add apply_update with backup swap + rollback on extract fail"
```

---

## Task 5: `updater.restart_app()` + orchestration helper

**Goal:** Add `restart_app()` (spawns new process, exits current) and `run_update_flow(info, parent_window)` that ties check/download/apply/restart together with CTk dialogs. Restart is hard to unit-test; orchestrator is integration-tested manually in Task 11.

**Files:**
- Modify: `updater.py`

- [ ] **Step 1: Add `restart_app` and `run_update_flow`**

Append to `updater.py`:

```python
import subprocess
import tempfile


def restart_app() -> "NoReturn":
    """Spawn a fresh process of the current exe and exit. No return."""
    if getattr(sys, "frozen", False):
        exe = sys.executable
        subprocess.Popen([exe], close_fds=True)
    else:
        # Dev mode: re-run python with same script
        subprocess.Popen([sys.executable] + sys.argv, close_fds=True)
    sys.exit(0)


def run_update_flow(info: "UpdateInfo", parent_window=None) -> bool:
    """Download + apply + restart. Shows CTk messageboxes for progress/errors.
    Returns False if the user cancels or any step fails; True flow does not
    return (process exits via restart_app).

    `parent_window` is a CTk window for dialog parenting (optional).
    """
    # Local import so updater stays GUI-agnostic in tests
    from tkinter import messagebox

    tmp_zip = os.path.join(tempfile.gettempdir(), "CubeLogReader_update.zip")
    ok = download_update(info, tmp_zip)
    if not ok:
        messagebox.showerror(
            "Güncelleme başarısız",
            "İndirme veya doğrulama başarısız oldu. İnternet bağlantını kontrol et.",
            parent=parent_window,
        )
        return False

    try:
        apply_update(tmp_zip)
    except Exception as e:
        messagebox.showerror(
            "Güncelleme başarısız",
            f"Dosyalar yazılamadı: {e}\nEski sürüm korundu.",
            parent=parent_window,
        )
        return False
    finally:
        try:
            os.remove(tmp_zip)
        except OSError:
            pass

    messagebox.showinfo(
        "Güncelleme tamam",
        f"Sürüm {info.version} yüklendi. Uygulama yeniden başlatılacak.",
        parent=parent_window,
    )
    restart_app()
    return True  # unreachable
```

- [ ] **Step 2: Verify updater module still imports cleanly**

Run: `python -c "import updater; print(updater.check_for_update.__doc__)"`
Expected: prints the docstring (no errors).

- [ ] **Step 3: Re-run existing tests**

Run: `python -m unittest test_updater -v`
Expected: All 15 tests still PASS.

- [ ] **Step 4: Commit**

```bash
git add updater.py
git commit -m "feat(updater): add restart_app + run_update_flow orchestrator"
```

---

## Task 6: `launcher.py` with rollback logic

**Goal:** PyInstaller entry point. Adds `src/` to `sys.path`, runs `main.main()`. On any exception, falls back to `src_backup/` and relaunches. Includes forced imports so PyInstaller bundles all heavy deps.

**Files:**
- Create: `launcher.py`

- [ ] **Step 1: Create `launcher.py`**

```python
"""
launcher.py — PyInstaller entry point. Loads application from external src/.

This file is what PyInstaller builds into CubeLogReader.exe. It contains:
  1. Forced imports so PyInstaller bundles all runtime dependencies.
  2. Logic to load src/main.py at runtime from beside the exe.
  3. Rollback logic: if src/ crashes on import or main(), restore src_backup/
     and relaunch ONCE. Show error dialog if no backup or second attempt fails.

src/main.py, src/reader.py, src/writer.py, src/updater.py are NOT bundled.
They live alongside the exe and are updated by updater.apply_update().
"""
from __future__ import annotations

# === Force PyInstaller to bundle all runtime deps ============================
# These imports do nothing at runtime in production (they're already imported
# by src/main.py too), but they make PyInstaller's static analyzer include
# every package and DLL we need.
import customtkinter  # noqa: F401
import google.generativeai  # noqa: F401
import win32com.client  # noqa: F401
import pythoncom  # noqa: F401
import win32timezone  # noqa: F401
import fitz  # PyMuPDF  # noqa: F401
import PIL.Image  # noqa: F401
import dotenv  # noqa: F401
try:
    import tkinterdnd2  # noqa: F401
except ImportError:
    pass
# =============================================================================

import os
import shutil
import sys
import traceback


def _app_dir() -> str:
    return os.path.dirname(sys.executable) if getattr(sys, "frozen", False) \
        else os.path.dirname(os.path.abspath(__file__))


def _show_error(message: str) -> None:
    """Last-resort error dialog. Used when src/ AND src_backup/ both fail."""
    try:
        from tkinter import Tk, messagebox
        root = Tk()
        root.withdraw()
        messagebox.showerror("CubeLogReader — Kritik Hata", message)
        root.destroy()
    except Exception:
        # If even Tk fails, write a crash log
        try:
            with open(os.path.join(_app_dir(), "launcher_crash.log"), "a", encoding="utf-8") as f:
                f.write(message + "\n---\n")
        except OSError:
            pass


def _run_src() -> None:
    """Add src/ to sys.path and run main(). Raises on any failure."""
    src_dir = os.path.join(_app_dir(), "src")
    if not os.path.isdir(src_dir):
        raise RuntimeError(f"src/ klasörü bulunamadı: {src_dir}")
    if src_dir not in sys.path:
        sys.path.insert(0, src_dir)
    # Clear any cached imports from a previous (failed) attempt
    for mod in ("main", "reader", "writer", "updater"):
        sys.modules.pop(mod, None)
    import main as _main  # noqa: F401
    _main.main()


def _rollback() -> bool:
    """Restore src_backup/ → src/. Returns True if restored."""
    app = _app_dir()
    src = os.path.join(app, "src")
    backup = os.path.join(app, "src_backup")
    if not os.path.isdir(backup):
        return False
    if os.path.isdir(src):
        shutil.rmtree(src, ignore_errors=True)
    os.rename(backup, src)
    return True


def main() -> None:
    try:
        _run_src()
        return
    except SystemExit:
        raise  # normal exit, don't roll back
    except BaseException:
        first_err = traceback.format_exc()

    # First attempt failed → try rollback + one retry
    if not _rollback():
        _show_error(
            "Uygulama açılamadı ve geri yüklenecek önceki sürüm yok.\n\n"
            f"Hata:\n{first_err}\n\n"
            "Lütfen kurulum dosyasını tekrar çalıştır."
        )
        sys.exit(1)

    try:
        _run_src()
    except SystemExit:
        raise
    except BaseException:
        second_err = traceback.format_exc()
        _show_error(
            "Uygulama açılamadı (geri yükleme sonrası bile).\n\n"
            f"İlk hata:\n{first_err}\n\nİkinci hata:\n{second_err}"
        )
        sys.exit(1)


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Smoke test launcher in dev mode**

Run: `python launcher.py`
Expected: app opens normally (launcher finds `main.py` because in dev mode `src/` doesn't exist → raises RuntimeError → no backup → error dialog).

This will FAIL the way we want — it confirms the error path. We'll fix dev-mode behavior next.

- [ ] **Step 3: Add dev-mode fallback in `launcher._run_src`**

Replace the body of `_run_src` in `launcher.py`:

```python
def _run_src() -> None:
    """Add src/ to sys.path and run main(). In dev mode, fall back to
    project root if src/ doesn't exist."""
    app = _app_dir()
    src_dir = os.path.join(app, "src")
    target_dir = src_dir if os.path.isdir(src_dir) else app
    if target_dir not in sys.path:
        sys.path.insert(0, target_dir)
    for mod in ("main", "reader", "writer", "updater"):
        sys.modules.pop(mod, None)
    import main as _main  # noqa: F401
    _main.main()
```

- [ ] **Step 4: Re-run launcher in dev mode**

Run: `python launcher.py`
Expected: CubeLogReader main window opens normally.

- [ ] **Step 5: Commit**

```bash
git add launcher.py
git commit -m "feat: add launcher.py with rollback to src_backup on import failure"
```

---

## Task 7: Integrate update check into `main.py`

**Goal:** On startup, kick off a background thread that calls `updater.check_for_update()`. If a new version is found, show a modal asking the user; on accept, call `updater.run_update_flow()`. Add a "Güncellemeleri kontrol et" button to `SettingsDialog`.

**Files:**
- Modify: `main.py:2580-2600` (the `main()` function and `__main__` block — exact lines may differ; locate by class name)
- Modify: `main.py:199-...` (the `SettingsDialog` class — locate the button row)

- [ ] **Step 1: Import `updater` at top of `main.py`**

Find the existing imports near the top (around `import reader` / `import writer`, currently line ~40-42):

```python
import reader
import writer
```

Add right after:

```python
import updater
```

- [ ] **Step 2: Add background update-check call in `MainWindow.__init__` or `main()`**

Open `main.py`, find the `main()` function (around line 2580). It currently looks like:

```python
def main():
    ...
    root = ...
    MainWindow(root)
    root.mainloop()
```

After `MainWindow(root)` and before `root.mainloop()`, add:

```python
    # Kick off background update check
    _start_update_check_thread(root)
```

Then add this helper function just above `main()`:

```python
def _start_update_check_thread(root):
    """Background thread: poll GitHub, if new version, prompt user on UI thread."""
    def worker():
        info = updater.check_for_update(timeout=5)
        if info is None:
            return
        # Marshal to UI thread
        root.after(0, lambda: _prompt_user_for_update(root, info))

    t = threading.Thread(target=worker, daemon=True)
    t.start()


def _prompt_user_for_update(root, info):
    """Modal dialog: 'New version X — update now?'"""
    from tkinter import messagebox
    msg = f"Yeni sürüm mevcut: {info.version}\n\n{info.notes}\n\nŞimdi güncellensin mi?"
    if messagebox.askyesno("Güncelleme mevcut", msg, parent=root):
        updater.run_update_flow(info, parent_window=root)
```

(`threading` is already imported at top of `main.py`.)

- [ ] **Step 3: Add "Check for updates" button to `SettingsDialog`**

Locate the `SettingsDialog` class (around line 199). Find the row of buttons near the bottom of the dialog (typically a "Save" / "Cancel" row). Add a new button row above it:

```python
        # --- Update check row -----------------------------------------------
        update_row = ctk.CTkFrame(self.root, fg_color="transparent")
        update_row.pack(fill="x", padx=20, pady=(10, 0))
        ctk.CTkButton(
            update_row,
            text="Güncellemeleri kontrol et",
            fg_color="#2E7D32",
            hover_color="#1B5E20",
            corner_radius=8,
            command=self._on_check_updates,
        ).pack(side="right")
```

Then add the handler method inside `SettingsDialog`:

```python
    def _on_check_updates(self):
        from tkinter import messagebox
        info = updater.check_for_update(timeout=8)
        if info is None:
            messagebox.showinfo("Güncelleme yok", "En güncel sürümdesin.", parent=self.root)
            return
        msg = f"Yeni sürüm: {info.version}\n\n{info.notes}\n\nŞimdi güncellensin mi?"
        if messagebox.askyesno("Güncelleme mevcut", msg, parent=self.root):
            updater.run_update_flow(info, parent_window=self.root)
```

(Adjust `self.root` to whatever attribute `SettingsDialog` actually uses — could be `self.window`, `self.dialog`, etc. Check existing button bindings in the class to find the correct attribute name.)

- [ ] **Step 4: Run the app in dev mode**

Run: `python main.py`
Expected: app opens; no error in console even though `OWNER/REPO` GitHub call fails (silent). Settings dialog has the new button.

Test the button: click it → "En güncel sürümdesin" message appears (since GitHub call fails → `info is None` → no-update branch).

- [ ] **Step 5: Commit**

```bash
git add main.py
git commit -m "feat: integrate update check on startup and SettingsDialog button"
```

---

## Task 8: Update `CubeLogReader.spec`

**Goal:** Switch PyInstaller entry to `launcher.py`. Exclude `main`/`reader`/`writer`/`updater` from the bundled archive (they ship outside). Keep all existing `collect_all` calls for google/grpc.

**Files:**
- Modify: `CubeLogReader.spec`

- [ ] **Step 1: Replace `CubeLogReader.spec`**

Current spec has `['main.py']` as the script. Replace entire file with:

```python
# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = []
binaries = []
hiddenimports = ['pythoncom', 'win32com.client', 'win32timezone']
tmp_ret = collect_all('google')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('grpc')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['launcher.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # Exclude the updatable application modules — they live in src/, not in the exe
    excludes=['main', 'reader', 'writer', 'updater'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='CubeLogReader',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='CubeLogReader',
)
```

- [ ] **Step 2: Verify the change is correct**

Run: `grep "launcher.py\|excludes" CubeLogReader.spec`
Expected: shows `['launcher.py']` and `excludes=['main', 'reader', 'writer', 'updater']`.

- [ ] **Step 3: Commit**

```bash
git add CubeLogReader.spec
git commit -m "build: switch PyInstaller entry to launcher.py, exclude src modules"
```

---

## Task 9: Update `build_exe.bat` to copy `src/` files post-build

**Goal:** After `pyinstaller` completes, copy `main.py`, `reader.py`, `writer.py`, `updater.py`, `version.txt` into `dist/CubeLogReader/src/`.

**Files:**
- Modify: `build_exe.bat`

- [ ] **Step 1: Replace `build_exe.bat`**

```bat
@echo off
cd /d "%~dp0"
echo ================================================
echo  CubeLogReader - building EXE (launcher mode)
echo ================================================
echo.

if not exist "launcher.py" (
    echo ERROR: launcher.py not found
    pause
    exit /b 1
)
if not exist "main.py" (
    echo ERROR: main.py not found
    pause
    exit /b 1
)

pyinstaller --noconfirm CubeLogReader.spec

if errorlevel 1 (
    echo.
    echo ERROR: build failed
    pause
    exit /b 1
)

echo.
echo Copying src/ files into dist...
set SRCDIR=dist\CubeLogReader\src
if not exist "%SRCDIR%" mkdir "%SRCDIR%"
copy /Y main.py     "%SRCDIR%\" >nul
copy /Y reader.py   "%SRCDIR%\" >nul
copy /Y writer.py   "%SRCDIR%\" >nul
copy /Y updater.py  "%SRCDIR%\" >nul
copy /Y version.txt "%SRCDIR%\" >nul

echo.
echo ================================================
echo  Build complete!
echo ================================================
echo.
echo Exe folder: dist\CubeLogReader\
echo Main file:  dist\CubeLogReader\CubeLogReader.exe
echo Source dir: dist\CubeLogReader\src\
echo.
echo Copy this ENTIRE folder to the other computer and
echo double-click the exe. Python is not required.
echo.
pause
```

- [ ] **Step 2: Commit**

```bash
git add build_exe.bat
git commit -m "build: copy src/ files into dist after PyInstaller"
```

---

## Task 10: Verify `installer.iss` picks up `src/`

**Goal:** Confirm Inno Setup includes `src/` automatically (current line uses `recursesubdirs` so no change should be needed).

**Files:**
- Read: `installer.iss`

- [ ] **Step 1: Inspect existing `[Files]` block**

Run: `grep -A1 "\[Files\]" installer.iss`
Expected output: `Source: "dist\CubeLogReader\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs`

This already recurses into subdirectories, so `dist\CubeLogReader\src\*.py` will be included automatically. No edit needed.

- [ ] **Step 2: Update `MyAppVersion` define** (matches `version.txt`)

In `installer.iss`, find:

```
#define MyAppVersion "1.0.0"
```

Leave it at `1.0.0` for the first build. (Future releases bump both `version.txt` and this define — Task 12 release-process step.)

- [ ] **Step 3: No commit needed** (no file changes). Proceed.

---

## Task 11: Full build + local smoke test (no real release yet)

**Goal:** Build the new launcher-based exe, install it locally, confirm app still works, and confirm the "Check for updates" button shows "no update" (because the real release doesn't exist yet — GitHub call returns 404 → silent).

**Files:** none modified.

- [ ] **Step 1: Clean previous build artifacts**

Run: `rmdir /s /q build dist`
Expected: folders removed. (Skip if they don't exist.)

- [ ] **Step 2: Build**

Run: `build_exe.bat`
Expected: build completes; `dist\CubeLogReader\CubeLogReader.exe` exists; `dist\CubeLogReader\src\main.py`, `reader.py`, `writer.py`, `updater.py`, `version.txt` all present.

- [ ] **Step 3: Verify `src/` contents**

Run: `dir dist\CubeLogReader\src`
Expected: 5 files listed (4 `.py` + `version.txt`).

- [ ] **Step 4: Launch the built exe**

Double-click `dist\CubeLogReader\CubeLogReader.exe`.
Expected: app opens normally, identical to the old build.

- [ ] **Step 5: Test "Check for updates" button**

Open Settings → click "Güncellemeleri kontrol et".
Expected: "En güncel sürümdesin" message (because GitHub `OWNER/REPO` returns 404 → silent → no-update branch).

- [ ] **Step 6: Test write flow once**

Drag a real test PDF into the app, run the normal write flow into an open Excel.
Expected: still works exactly like before — proves we didn't break anything.

- [ ] **Step 7: Run rollback smoke test**

Manually break the build:
1. Open `dist\CubeLogReader\src\main.py` in a text editor.
2. On the first line add: `raise RuntimeError("test")`
3. Save.
4. Double-click the exe.

Expected: app opens an error dialog *or* nothing happens (no backup yet on first install, so error dialog is correct).

Now simulate a backup:
1. Copy `dist\CubeLogReader\src` → `dist\CubeLogReader\src_backup` (manually).
2. Edit `src\main.py` to still have the `raise RuntimeError("test")`.
3. Double-click exe.

Expected: app briefly tries, rolls back (`src_backup` → `src`), retries, opens normally.

- [ ] **Step 8: Clean up the deliberate break**

Restore `src/main.py` (re-copy from project root if needed) and remove `src_backup/`.

- [ ] **Step 9: Commit any incidental fixes discovered during smoke test**

If smoke test surfaced bugs, fix and commit them. Otherwise skip.

```bash
git status
# if no changes: do nothing
```

---

## Task 12: Create GitHub repo + first release `v1.0.0`

**Goal:** Set up the actual remote endpoint that the app polls.

**Prerequisites:** `gh` CLI installed on the dev machine (`winget install --id GitHub.cli`), then `gh auth login` completed once.

**Files:**
- Modify: `updater.py:15` — replace `GITHUB_OWNER = "OWNER"` and `GITHUB_REPO = "REPO"` with real values.

- [ ] **Step 1: Verify `gh` CLI**

Run: `gh --version`
Expected: prints `gh version 2.x.x`.

If not installed: run `winget install --id GitHub.cli`, restart shell, then `gh auth login` (choose GitHub.com → HTTPS → browser).

- [ ] **Step 2: Create a public GitHub repo**

Pick a name (e.g., `CubeLogReader`). Run:

```bash
gh repo create CubeLogReader --public --confirm --description "Concrete cube test notebook → Excel importer"
```

Note the resulting URL: `https://github.com/<your-user>/CubeLogReader`. Record `<your-user>` — needed in next step.

- [ ] **Step 3: Update `updater.py` with real owner/repo**

In `updater.py`, replace:

```python
GITHUB_OWNER = "OWNER"
GITHUB_REPO = "REPO"
```

with (example):

```python
GITHUB_OWNER = "yavuzzeynula"
GITHUB_REPO = "CubeLogReader"
```

- [ ] **Step 4: Commit and push initial code**

```bash
git init  # if not already a repo
git remote add origin https://github.com/<your-user>/CubeLogReader.git
git add .
git commit -m "chore: initial commit"
git branch -M main
git push -u origin main
```

If `.git` already exists, just push:

```bash
git remote add origin https://github.com/<your-user>/CubeLogReader.git
git push -u origin main
```

- [ ] **Step 5: Rebuild exe with updated owner/repo**

Run: `build_exe.bat`
Expected: build completes.

- [ ] **Step 6: Package `src.zip` for the first release**

Run from project root:

```bash
cd src 2>nul || cd dist\CubeLogReader\src
powershell Compress-Archive -Path *.py,version.txt -DestinationPath ..\..\..\src.zip -Force
cd ..\..\..
```

(Or use any zip tool — the requirement is a flat zip containing the 4 `.py` files and `version.txt` with NO top-level folder.)

- [ ] **Step 7: Compute SHA256 of `src.zip`**

Run:

```powershell
powershell -Command "(Get-FileHash src.zip -Algorithm SHA256).Hash.ToLower()"
```

Copy the hex string.

- [ ] **Step 8: Create the `v1.0.0` release**

```bash
gh release create v1.0.0 src.zip --title "v1.0.0 — Initial release" --notes "Initial release.\n\nSHA256: <paste-hash-here>"
```

Expected: prints release URL.

- [ ] **Step 9: Verify the release is reachable**

Run:

```bash
curl -s https://api.github.com/repos/<your-user>/CubeLogReader/releases/latest | findstr tag_name
```

Expected: `"tag_name": "v1.0.0",`

- [ ] **Step 10: Commit the updater.py owner change**

```bash
git add updater.py
git commit -m "chore: point updater at real GitHub repo"
git push
```

---

## Task 13: End-to-end update test (release `v1.0.1`)

**Goal:** Prove the full update loop works against the real GitHub repo.

**Files:**
- Modify: `version.txt` → `1.0.1`
- Modify: `main.py` (one visible cosmetic change — e.g., window title suffix)

- [ ] **Step 1: Make a visible cosmetic change in `main.py`**

Find where the main window title is set (search for `.title(` in `MainWindow.__init__`). Add a marker:

```python
root.title("CubeLogReader (v1.0.1)")
```

(Or change any other visible text.)

- [ ] **Step 2: Bump version**

Edit `version.txt`:

```
1.0.1
```

- [ ] **Step 3: Repackage `src.zip` from updated source files**

```bash
powershell Compress-Archive -Path main.py,reader.py,writer.py,updater.py,version.txt -DestinationPath src.zip -Force
```

- [ ] **Step 4: Compute SHA256**

```powershell
powershell -Command "(Get-FileHash src.zip -Algorithm SHA256).Hash.ToLower()"
```

- [ ] **Step 5: Release v1.0.1**

```bash
gh release create v1.0.1 src.zip --title "v1.0.1 — test update" --notes "Test update.\n\n- Title bar değişti\n\nSHA256: <paste-hash>"
```

- [ ] **Step 6: Test the installed app picks up the update**

Open the locally-installed `dist\CubeLogReader\CubeLogReader.exe` (this is still v1.0.0 because we haven't re-installed).

Expected: within ~5 seconds of startup, dialog appears: *"Yeni sürüm mevcut: 1.0.1 ... Şimdi güncellensin mi?"*

Click **Evet**.

Expected:
- "İndiriliyor" period (usually <1 sec for 50KB).
- "Güncelleme tamam, yeniden başlatılacak" dialog.
- App restarts.
- New window title now shows `(v1.0.1)`.
- `dist\CubeLogReader\src\version.txt` now contains `1.0.1`.
- `dist\CubeLogReader\src_backup\version.txt` contains `1.0.0`.

- [ ] **Step 7: Test "no update" path**

Click Settings → "Güncellemeleri kontrol et".
Expected: "En güncel sürümdesin" (because installed app is now v1.0.1, same as latest release).

- [ ] **Step 8: Commit the v1.0.1 source state**

```bash
git add main.py version.txt
git commit -m "test: v1.0.1 release for E2E update test"
git push
git tag v1.0.1 -f  # ensure tag matches GitHub
git push --tags -f
```

---

## Task 14: Rollback test (deliberately broken `v1.0.2`)

**Goal:** Prove the launcher rolls back to `src_backup/` when a release ships broken code.

**Files:**
- Modify: `version.txt` → `1.0.2`
- Modify: `main.py` (intentional break)

- [ ] **Step 1: Introduce a deliberate import error**

In `main.py`, near the top, add:

```python
import nonexistent_module_for_test  # noqa
```

(This will cause an `ImportError` when `from main import main` runs in the launcher.)

- [ ] **Step 2: Bump version**

Edit `version.txt`:

```
1.0.2
```

- [ ] **Step 3: Repackage and release**

```bash
powershell Compress-Archive -Path main.py,reader.py,writer.py,updater.py,version.txt -DestinationPath src.zip -Force
powershell -Command "(Get-FileHash src.zip -Algorithm SHA256).Hash.ToLower()"
gh release create v1.0.2 src.zip --title "v1.0.2 — deliberately broken" --notes "Test rollback. SHA256: <paste>"
```

- [ ] **Step 4: Open the installed app (currently v1.0.1)**

Double-click `dist\CubeLogReader\CubeLogReader.exe`.

Expected: update prompt appears for v1.0.2. Click **Evet**.

App downloads, applies, restarts.

On restart:
- Launcher tries `from main import main` → `ImportError: nonexistent_module_for_test`.
- Launcher catches, calls `_rollback()` → `src_backup/` (v1.0.1) → `src/`.
- Launcher relaunches via `_run_src()` → succeeds.
- App opens normally with v1.0.1 title.

Verify: check `dist\CubeLogReader\src\version.txt` — should be `1.0.1` (rolled back).

- [ ] **Step 5: Revert the deliberate break in source**

In `main.py`, remove the `import nonexistent_module_for_test` line.

- [ ] **Step 6: Delete the broken release from GitHub**

```bash
gh release delete v1.0.2 --yes
git push origin --delete v1.0.2 2>nul
```

(So future users don't get the broken update.)

- [ ] **Step 7: Commit the revert**

```bash
git add main.py
git commit -m "test: revert deliberate v1.0.2 break after rollback test passed"
git push
```

---

## Task 15: Update memory + finalize

**Goal:** Record the new architecture so future sessions know about it.

**Files:**
- Modify: `C:\Users\Yafka\.claude\projects\C--Users-Yafka-Desktop-CubeLogReader\memory\project_cubelogreader.md`

- [ ] **Step 1: Append auto-update info to the project memory file**

Add a new section to `project_cubelogreader.md`:

```markdown
**Auto-update (added 2026-05-XX):**
- `launcher.py` is the PyInstaller entry; loads `src/main.py` at runtime.
- `src/` next to exe contains updatable `.py` files + `version.txt`.
- `updater.py` polls `https://api.github.com/repos/<owner>/CubeLogReader/releases/latest`.
- New release → CTk dialog → download src.zip → swap `src/` ↔ `src_backup/` → restart.
- Launcher catches import/runtime errors and auto-rolls back to `src_backup/`.
- Release process: bump `version.txt`, zip `src/*.py + version.txt`, `gh release create vX.Y.Z src.zip --notes "...SHA256: ..."`
```

- [ ] **Step 2: Final smoke test of installed app**

Open the installed exe one more time. Confirm:
- Opens normally (v1.0.1 after the rollback test).
- Settings → "Güncellemeleri kontrol et" → "En güncel sürümdesin".

- [ ] **Step 3: All done**

The auto-update infrastructure is live. Future sessions just need to: edit a `.py`, bump `version.txt`, repackage, `gh release create`.

---

## Self-Review Notes

Verified against the spec:

- ✅ Section 1 (success criteria): all four criteria covered by Tasks 11–14.
- ✅ Section 2 (architecture): Tasks 6, 8, 9 implement launcher + src split.
- ✅ Section 3.1 (launcher): Task 6.
- ✅ Section 3.2 (updater API: check / download / apply / restart): Tasks 2–5.
- ✅ Section 3.3 (main.py integration: startup + button): Task 7.
- ✅ Section 3.4 (build system): Tasks 8 (spec), 9 (bat), 10 (installer verified).
- ✅ Section 4 (state diagram): rollback path covered by Task 6 + tested in Task 14.
- ✅ Section 5 (AV strategy): exe-unchanged property is structural; no separate task needed.
- ✅ Section 6 (risk matrix): all rows covered (rollback test, SHA256, silent network failure).
- ✅ Section 7 (test strategy): unit tests in Tasks 2–4; E2E in Tasks 11, 13, 14.
- ✅ Section 9 (prerequisites): Task 12 step 1.

No placeholders, no TBDs. Type names (`UpdateInfo`, `check_for_update`, `download_update`, `apply_update`, `restart_app`, `run_update_flow`) are consistent across all tasks.
