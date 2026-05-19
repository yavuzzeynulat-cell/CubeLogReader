"""
updater.py - Auto-update logic for CubeLogReader.

Public API:
    check_for_update(timeout=5) -> Optional[UpdateInfo]
    download_update(info, dest_path) -> bool
    apply_update(zip_path) -> None
    restart_app() -> NoReturn
    run_update_flow(info, parent_window=None) -> bool

All network failures are silenced (return None / False) so the app
keeps running even when offline.
"""
from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import stat
import subprocess
import sys
import tempfile
import time
import traceback
import urllib.request
import zipfile
from dataclasses import dataclass
from typing import Optional

# === CONFIG: change OWNER/REPO once the GitHub repo is created ===
GITHUB_OWNER = "yavuzzeynulat-cell"
GITHUB_REPO = "CubeLogReader"
RELEASE_API = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
ASSET_NAME = "src.zip"


@dataclass
class UpdateInfo:
    version: str            # e.g. "1.0.1" (no "v" prefix)
    notes: str              # raw release body (Turkish, multi-line)
    asset_url: str          # direct download URL for src.zip
    sha256: Optional[str]   # parsed from notes; None if absent


def _app_dir() -> str:
    """Folder containing the running exe (or this file during dev)."""
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
                with open(p, "r", encoding="utf-8") as f:
                    return f.read().strip()
            except OSError:
                pass
    return "0.0.0"


def _parse_version(s: str) -> tuple:
    """Parse '1.0.1' into (1,0,1). Invalid -> (0,0,0)."""
    try:
        return tuple(int(x) for x in s.strip().lstrip("v").split("."))
    except (ValueError, AttributeError):
        return (0, 0, 0)


def _fetch_latest_release_json(timeout: int = 5) -> dict:
    """GET the latest release. Raises OSError on network failure."""
    req = urllib.request.Request(
        RELEASE_API,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": "CubeLogReader-Updater",
        },
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
    """Return UpdateInfo if a newer release exists; else None.

    Silent on any failure (network, parse, missing fields) -> returns None.
    """
    try:
        data = _fetch_latest_release_json(timeout=timeout)
        tag = data.get("tag_name")
        if not tag:
            return None
        remote_v = tag.lstrip("v")
        if _parse_version(remote_v) <= _parse_version(_read_local_version()):
            return None
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

    Returns True on success. On failure, removes any partial file.
    """
    try:
        _http_download(info.asset_url, dest_path, timeout=timeout)
    except Exception:
        _safe_remove(dest_path)
        return False

    if info.sha256:
        actual = _sha256_file(dest_path)
        if actual.lower() != info.sha256.lower():
            _safe_remove(dest_path)
            return False

    try:
        with zipfile.ZipFile(dest_path, "r") as zf:
            bad = zf.testzip()
            if bad is not None:
                raise zipfile.BadZipFile(f"corrupt entry: {bad}")
    except Exception:
        _safe_remove(dest_path)
        return False

    return True


def _safe_remove(path: str) -> None:
    if os.path.isfile(path):
        try:
            os.remove(path)
        except OSError:
            pass


def _force_writable(path: str) -> None:
    """Clear the read-only attribute so a file can be deleted/overwritten."""
    try:
        os.chmod(path, stat.S_IWRITE)
    except OSError:
        pass


def _rmtree_onerror(func, path, _exc) -> None:
    """rmtree error handler: clear the read-only bit and retry the op once."""
    _force_writable(path)
    try:
        func(path)
    except OSError:
        pass


def _robust_rmtree(path: str, attempts: int = 5) -> None:
    """Delete a directory tree, tolerating read-only files and transient
    antivirus locks. Raises OSError if it still exists after every attempt."""
    for _ in range(attempts):
        if not os.path.isdir(path):
            return
        shutil.rmtree(path, onerror=_rmtree_onerror)
        if not os.path.isdir(path):
            return
        time.sleep(0.4)
    if os.path.isdir(path):
        raise OSError(f"Could not delete folder (locked?): {path}")


def _robust_rename(src: str, dst: str, attempts: int = 5) -> None:
    """os.rename with retries - antivirus often briefly locks fresh files."""
    last_err: Optional[OSError] = None
    for _ in range(attempts):
        try:
            os.rename(src, dst)
            return
        except OSError as e:
            last_err = e
            time.sleep(0.4)
    raise last_err if last_err else OSError(f"rename failed: {src} -> {dst}")


def _log_update_error() -> str:
    """Append the current traceback to update_error.log next to the exe.

    Returns the log path. Best-effort - never raises.
    """
    log_path = os.path.join(_app_dir(), "update_error.log")
    try:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write("=== update failed ===\n")
            f.write(traceback.format_exc())
            f.write("\n")
    except OSError:
        pass
    return log_path


def apply_update(zip_path: str) -> None:
    """Replace src/ with the contents of zip_path; keep the old src/ as
    src_backup/ for the launcher's rollback safety net.

    Hardened for Windows: extracts into a staging folder first (so a locked
    or failed extract never damages the running src/), then performs only
    fast directory renames. Tolerates read-only files and transient antivirus
    locks. Raises OSError if it ultimately fails, after restoring src/.
    """
    app_dir = _app_dir()
    src = os.path.join(app_dir, "src")
    backup = os.path.join(app_dir, "src_backup")
    staging = os.path.join(app_dir, "src_new")

    # 1. Extract into a fresh staging folder. If this fails, src/ is untouched.
    _robust_rmtree(staging)
    os.makedirs(staging, exist_ok=True)
    try:
        with zipfile.ZipFile(zip_path, "r") as zf:
            zf.extractall(staging)
    except Exception:
        _robust_rmtree(staging)
        raise

    # 2. Drop the previous backup, then swap: src -> backup, staging -> src.
    try:
        _robust_rmtree(backup)
        if os.path.isdir(src):
            _robust_rename(src, backup)
    except OSError:
        _robust_rmtree(staging)
        raise

    try:
        _robust_rename(staging, src)
    except OSError:
        # Swap failed mid-way: put the old src/ back so the app still runs.
        if not os.path.isdir(src) and os.path.isdir(backup):
            _robust_rename(backup, src)
        _robust_rmtree(staging)
        raise


def restart_app() -> "NoReturn":
    """Spawn a fresh process of the current exe and exit. No return."""
    if getattr(sys, "frozen", False):
        subprocess.Popen([sys.executable], close_fds=True)
    else:
        subprocess.Popen([sys.executable] + sys.argv, close_fds=True)
    sys.exit(0)


def run_update_flow(info: "UpdateInfo", parent_window=None) -> bool:
    """Download + apply + restart. Shows messageboxes for progress/errors.

    Returns False if download/apply fails. On success, this function does
    not return - the process exits via restart_app().
    """
    from tkinter import messagebox

    tmp_zip = os.path.join(tempfile.gettempdir(), "CubeLogReader_update.zip")
    if not download_update(info, tmp_zip):
        messagebox.showerror(
            "Update failed",
            "Download or verification failed. Check your internet connection.",
            parent=parent_window,
        )
        return False

    try:
        apply_update(tmp_zip)
    except Exception as e:
        log_path = _log_update_error()
        messagebox.showerror(
            "Update failed",
            f"Could not write files: {e}\n\n"
            f"Details saved to:\n{log_path}\n\n"
            "Previous version kept.",
            parent=parent_window,
        )
        return False
    finally:
        _safe_remove(tmp_zip)

    messagebox.showinfo(
        "Update complete",
        f"Version {info.version} installed. The app will now restart.",
        parent=parent_window,
    )
    restart_app()
    return True  # unreachable
