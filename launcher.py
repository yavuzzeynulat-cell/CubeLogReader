"""
launcher.py - PyInstaller entry point. Loads application from external src/.

This file is what PyInstaller builds into CubeLogReader.exe. It:
  1. Forces imports so PyInstaller bundles all runtime dependencies.
  2. Loads src/main.py at runtime from beside the exe (or project root in dev).
  3. Rolls back: if src/ crashes on import or main(), restore src_backup/
     and relaunch ONCE. Shows error dialog if no backup or second attempt fails.

src/main.py, src/reader.py, src/writer.py, src/updater.py are NOT bundled.
They live alongside the exe and are updated by updater.apply_update().
"""
from __future__ import annotations

# === Force PyInstaller to bundle all runtime deps ============================
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
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _show_error(message: str) -> None:
    """Last-resort error dialog. Used when src/ AND src_backup/ both fail."""
    try:
        from tkinter import Tk, messagebox
        root = Tk()
        root.withdraw()
        messagebox.showerror("CubeLogReader - Critical Error", message)
        root.destroy()
    except Exception:
        try:
            with open(os.path.join(_app_dir(), "launcher_crash.log"), "a",
                      encoding="utf-8") as f:
                f.write(message + "\n---\n")
        except OSError:
            pass


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
    import main as _main
    _main.main()


def _rollback() -> bool:
    """Restore src_backup/ -> src/. Returns True if restored."""
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
        raise
    except BaseException:
        first_err = traceback.format_exc()

    if not _rollback():
        _show_error(
            "The app could not start and there is no previous version "
            "to restore.\n\n"
            f"Error:\n{first_err}\n\n"
            "Please run the installer again."
        )
        sys.exit(1)

    try:
        _run_src()
    except SystemExit:
        raise
    except BaseException:
        second_err = traceback.format_exc()
        _show_error(
            "The app could not start (even after rollback).\n\n"
            f"First error:\n{first_err}\n\nSecond error:\n{second_err}"
        )
        sys.exit(1)


if __name__ == "__main__":
    main()
