"""
main.py — Notebook -> Excel importer (tkinter UI).

Flow:
  1. Open the Excel file (each sheet must have Sample ID in B14).
  2. Click the sheet of the FIRST cube from the notebook.
  3. Run the program, pick the notebook page (PDF/JPG/PNG).
  4. Gemini reads the page, a review window opens.
  5. Edit the values, click "Write to Excel" — it writes into the open Excel.
"""
import os
import queue
import sys
import threading
import traceback
from pathlib import Path
from tkinter import (
    BOTH,
    LEFT,
    RIGHT,
    Y,
    BooleanVar,
    Canvas,
    Frame,
    Label,
    Scrollbar,
    StringVar,
    Tk,
    Toplevel,
    filedialog,
    messagebox,
    ttk,
)

from tkinter import font as tkfont

import customtkinter as ctk
from PIL import Image, ImageTk

import reader
import updater
import writer

# Optional drag & drop support
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    _DND_AVAILABLE = True
except Exception:
    DND_FILES = None  # type: ignore
    TkinterDnD = None  # type: ignore
    _DND_AVAILABLE = False


class DnDCTk(ctk.CTk):
    """ctk.CTk with tkinterdnd2 drop-target support mixed in."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        if _DND_AVAILABLE:
            try:
                self.TkdndVersion = TkinterDnD._require(self)
            except Exception:
                pass


def _parse_dropped_paths(raw: str) -> list[str]:
    """tkinterdnd2 delivers a space-separated string; paths with spaces
    are wrapped in {braces}. Parse into a clean list."""
    paths: list[str] = []
    i, n = 0, len(raw)
    while i < n:
        while i < n and raw[i] == " ":
            i += 1
        if i >= n:
            break
        if raw[i] == "{":
            j = raw.find("}", i + 1)
            if j == -1:
                paths.append(raw[i + 1:])
                break
            paths.append(raw[i + 1:j])
            i = j + 1
        else:
            j = raw.find(" ", i)
            if j == -1:
                paths.append(raw[i:])
                break
            paths.append(raw[i:j])
            i = j + 1
    return [p for p in paths if p]


# ---- CustomTkinter theme (applies to the new-look preview window) ----
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


APP_TITLE = "Notebook -> Excel Importer"
APP_CREDIT = "Written by Yavuz Zeynula"


def _add_credit_footer(parent):
    """Pack a small credit label at the bottom of the given window/frame."""
    ctk.CTkLabel(
        parent, text=APP_CREDIT,
        font=ctk.CTkFont(size=10),
        text_color="gray55",
    ).pack(side="bottom", pady=(2, 4))
MODEL_PRESETS = [
    "gemini-2.5-flash",
    "gemini-2.5-pro",
    "gemini-3-pro-preview",
]
DEFAULT_MODEL = "gemini-2.5-flash"
SCAN_MODES = {
    "short": ("25 sheets forward from active (fast)", 25),
    "long": ("350 sheets forward from active (whole workbook)", 350),
}
DEFAULT_SCAN_MODE = "long"
IMAGE_CANVAS_WIDTH = 920
IMAGE_CANVAS_HEIGHT = 860
ZOOM_MIN = 0.2
ZOOM_MAX = 2.5
ZOOM_STEP = 1.25

if getattr(sys, "frozen", False):
    # Packaged with PyInstaller — .env lives next to the exe
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).parent
ENV_PATH = BASE_DIR / ".env"


# ---------- .env / API key management ----------

def get_current_api_key() -> str:
    """Read the current API key from .env."""
    if not ENV_PATH.exists():
        return ""
    try:
        for line in ENV_PATH.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line.startswith("GEMINI_API_KEY="):
                return line.split("=", 1)[1].strip().strip('"').strip("'")
    except Exception:
        return ""
    return ""


def _save_env_var(var_name: str, value: str) -> None:
    """Upsert a KEY=VALUE line in .env and update the current process env."""
    value = value.strip()
    lines: list[str] = []
    if ENV_PATH.exists():
        try:
            lines = ENV_PATH.read_text(encoding="utf-8").splitlines()
        except Exception:
            lines = []

    found = False
    for i, line in enumerate(lines):
        if line.strip().startswith(f"{var_name}="):
            lines[i] = f"{var_name}={value}"
            found = True
            break
    if not found:
        lines.append(f"{var_name}={value}")

    ENV_PATH.write_text("\n".join(lines) + "\n", encoding="utf-8")
    os.environ[var_name] = value


def save_api_key(key: str) -> None:
    """Write the API key to .env and update the current process env."""
    _save_env_var("GEMINI_API_KEY", key)


def get_current_model() -> str:
    """Read the currently selected Gemini model from .env."""
    if not ENV_PATH.exists():
        return DEFAULT_MODEL
    try:
        for line in ENV_PATH.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line.startswith("GEMINI_MODEL="):
                val = line.split("=", 1)[1].strip().strip('"').strip("'")
                return val or DEFAULT_MODEL
    except Exception:
        pass
    return DEFAULT_MODEL


def save_model(model: str) -> None:
    """Write the chosen model to .env."""
    _save_env_var("GEMINI_MODEL", model)


# ---------- Settings dialog ----------

class SettingsDialog:
    """API key settings dialog."""

    def __init__(self, parent):
        self.parent = parent
        self.win = ctk.CTkToplevel(parent)
        self.win.title("Settings — " + APP_TITLE)
        self.win.geometry("620x620")
        self.win.transient(parent)
        self.win.grab_set()
        self.win.resizable(False, False)

        _add_credit_footer(self.win)

        current = get_current_api_key()

        # ---- HEADER ----
        header = ctk.CTkFrame(self.win, corner_radius=0, height=60)
        header.pack(side="top", fill="x")
        header.pack_propagate(False)
        title_wrap = ctk.CTkFrame(header, fg_color="transparent")
        title_wrap.pack(side="left", padx=22, pady=8)
        ctk.CTkLabel(
            title_wrap, text="Settings",
            font=ctk.CTkFont(size=18, weight="bold"),
            anchor="w",
        ).pack(anchor="w")
        ctk.CTkLabel(
            title_wrap, text="API key and Gemini model used to read notebook pages.",
            font=ctk.CTkFont(size=11), text_color="gray55", anchor="w",
        ).pack(anchor="w")

        # ---- BODY ----
        body = ctk.CTkFrame(self.win, fg_color="transparent")
        body.pack(side="top", fill="both", expand=True, padx=16, pady=12)

        # ---- API key card ----
        key_card = ctk.CTkFrame(body, corner_radius=12, border_width=1)
        key_card.pack(fill="x", pady=(0, 12))
        ctk.CTkLabel(
            key_card, text="Google Gemini API Key",
            font=ctk.CTkFont(size=13, weight="bold"), anchor="w",
        ).pack(anchor="w", padx=16, pady=(12, 2))
        ctk.CTkLabel(
            key_card, text="Free key: https://aistudio.google.com/apikey",
            font=ctk.CTkFont(size=11), text_color="gray55", anchor="w",
        ).pack(anchor="w", padx=16)

        if current:
            masked = (
                current[:8] + "..." + current[-4:] if len(current) > 14 else "***"
            )
            ctk.CTkLabel(
                key_card, text=f"Current key: {masked}",
                text_color="#2E7D32",
                font=ctk.CTkFont(size=11, weight="bold"), anchor="w",
            ).pack(anchor="w", padx=16, pady=(6, 2))
        else:
            ctk.CTkLabel(
                key_card, text="No key set yet",
                text_color="#C62828",
                font=ctk.CTkFont(size=11, weight="bold"), anchor="w",
            ).pack(anchor="w", padx=16, pady=(6, 2))

        ctk.CTkLabel(
            key_card, text="New key:",
            font=ctk.CTkFont(size=11), anchor="w",
        ).pack(anchor="w", padx=16, pady=(8, 2))

        self.key_var = StringVar(value=current)
        self.entry = ctk.CTkEntry(
            key_card, textvariable=self.key_var, show="*", height=34,
            font=ctk.CTkFont(size=12),
        )
        self.entry.pack(fill="x", padx=16, pady=(0, 4))
        self.entry.focus_set()

        self.show_var = BooleanVar(value=False)
        ctk.CTkCheckBox(
            key_card, text="Show key",
            variable=self.show_var, command=self._toggle_show,
            font=ctk.CTkFont(size=11),
        ).pack(anchor="w", padx=16, pady=(4, 14))

        # ---- Model card ----
        model_card = ctk.CTkFrame(body, corner_radius=12, border_width=1)
        model_card.pack(fill="x", pady=(0, 12))
        ctk.CTkLabel(
            model_card, text="Gemini Model",
            font=ctk.CTkFont(size=13, weight="bold"), anchor="w",
        ).pack(anchor="w", padx=16, pady=(12, 2))
        ctk.CTkLabel(
            model_card,
            text=(
                "Flash = fast, large free quota (~1500/day).\n"
                "Pro = slower but more accurate, smaller free quota (~50/day)."
            ),
            font=ctk.CTkFont(size=11), text_color="gray55",
            justify="left", anchor="w",
        ).pack(anchor="w", padx=16, pady=(0, 6))

        self.model_var = StringVar(value=get_current_model())
        self.model_combo = ctk.CTkOptionMenu(
            model_card, variable=self.model_var,
            values=MODEL_PRESETS, width=280, height=34,
            font=ctk.CTkFont(size=12),
        )
        self.model_combo.pack(anchor="w", padx=16, pady=(0, 14))

        # ---- Info note ----
        ctk.CTkLabel(
            body,
            text=(
                "Settings are stored in the .env file next to the program.\n"
                "If you move the program to another computer, just reopen this "
                "dialog and paste the new key — no need to edit the file."
            ),
            font=ctk.CTkFont(size=11), text_color="gray55",
            justify="left", anchor="w",
        ).pack(anchor="w", pady=(0, 12))

        # ---- Action buttons ----
        btns = ctk.CTkFrame(self.win, corner_radius=0, height=56)
        btns.pack(side="bottom", fill="x")
        btns.pack_propagate(False)
        ctk.CTkButton(
            btns, text="Cancel", width=110, height=36,
            fg_color="gray70", hover_color="gray60",
            command=self.win.destroy,
        ).pack(side="right", padx=(6, 16), pady=10)
        ctk.CTkButton(
            btns, text=">> SAVE", width=160, height=36,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#2E7D32", hover_color="#1B5E20",
            command=self._save,
        ).pack(side="right", padx=6, pady=10)
        ctk.CTkButton(
            btns, text="Guncellemeleri kontrol et", width=180, height=36,
            fg_color="#2E7D32", hover_color="#1B5E20",
            command=self._on_check_updates,
        ).pack(side="left", padx=16, pady=10)

    def _on_check_updates(self):
        info = updater.check_for_update(timeout=8)
        if info is None:
            messagebox.showinfo(
                "Guncelleme yok", "En guncel surumdesin.", parent=self.win
            )
            return
        msg = (
            f"Yeni surum: {info.version}\n\n{info.notes}\n\n"
            "Simdi guncellensin mi?"
        )
        if messagebox.askyesno("Guncelleme mevcut", msg, parent=self.win):
            updater.run_update_flow(info, parent_window=self.win)

    def _toggle_show(self):
        self.entry.configure(show="" if self.show_var.get() else "*")

    def _save(self):
        key = self.key_var.get().strip()
        model = self.model_var.get().strip() or DEFAULT_MODEL
        if not key:
            messagebox.showwarning(
                "Error", "Key cannot be empty", parent=self.win
            )
            return
        try:
            save_api_key(key)
            save_model(model)
        except Exception as e:
            messagebox.showerror(
                "Could not save", f"Error: {e}", parent=self.win
            )
            return
        messagebox.showinfo(
            "Done",
            f"Saved. Using model: {model}",
            parent=self.win,
        )
        self.win.destroy()


# ---------- Preview / Edit Window (CustomTkinter redesign) ----------

# Card colors for Excel-state indication (approved design)
EMPTY_BORDER = "#00C853"      # bright green — cell empty in Excel
EMPTY_BG = "#E8F5E9"          # subtle green tint inside empty cells
FILLED_TEXT = "gray55"        # dim — cell already filled in Excel
AGE_7_COLOR = "#2E7D32"       # dark green for "7-day" label
AGE_28_COLOR = "#C62828"      # dark red for "28-day" label
PILL_OK_COLOR = "#2E7D32"
PILL_NO_COLOR = "#C62828"
PILL_SHOT_COLOR = "#C62828"   # red — "SHOTCRETE" badge
SELECTED_BORDER = "#1976D2"   # blue — currently selected card
DIM_BG = "gray90"             # unused shotcrete row background


class PreviewWindow:
    """Shows the extracted data, lets the user edit and write to Excel."""

    def __init__(
        self,
        parent: Tk,
        image: Image.Image,
        cubes_data: dict,
        scan_result: dict,
        on_close=None,
    ):
        self.parent = parent
        self.cubes_data = cubes_data
        self.orig_image = image
        self.scan_result = scan_result
        self._on_close_cb = on_close
        self._closed = False

        # Zoom state
        self.zoom: float = 1.0  # will be set by _zoom_fit on first render
        self._img_photo: ImageTk.PhotoImage | None = None
        self._img_canvas: Canvas | None = None
        self._img_item: int | None = None
        self.zoom_label = None
        self._hq_after_id: str | None = None

        # CTkFont objects for cards — grow when image panel is hidden so
        # values are easier to verify.
        self._entry_font = ctk.CTkFont(family="Segoe UI", size=15)
        self._age_font = ctk.CTkFont(family="Segoe UI", size=13, weight="bold")
        self._title_font = ctk.CTkFont(family="Segoe UI", size=14, weight="bold")
        self._pill_font = ctk.CTkFont(family="Segoe UI", size=11, weight="bold")
        self._colhdr_font = ctk.CTkFont(family="Segoe UI", size=11, weight="bold")
        self._check_font = ctk.CTkFont(family="Segoe UI", size=12)

        self.win = ctk.CTkToplevel(parent)
        self.win.title("Review — " + APP_TITLE)
        # Fallback geometry, then maximize (user wants full-screen previews).
        sw = self.win.winfo_screenwidth()
        sh = self.win.winfo_screenheight()
        ww = min(1640, sw - 60)
        wh = min(970, sh - 80)
        self.win.geometry(f"{ww}x{wh}+20+20")
        self.win.minsize(900, 600)
        try:
            self.win.state("zoomed")
        except Exception:
            pass

        _add_credit_footer(self.win)

        # Sheet matching (from forward-scan starting at active sheet)
        self.open_sheets = scan_result.get("sheets", [])
        self.match_error: str | None = None
        try:
            self.matched = writer.match_cubes_to_sheets(
                cubes_data, sheets=self.open_sheets
            )
        except Exception as e:
            self.match_error = str(e)
            self.matched = [
                {"cube": c, "matched_sheet": None}
                for c in cubes_data.get("cubes", [])
            ]

        # UI state: per-cube Entry widget references
        # self.entries[i] = {"weights": [Entry, Entry, Entry], "loads": [...]}
        self.entries: list[dict] = []

        # Card selection state (arrow-key navigation)
        self._cards: list = []
        self._selected_idx: int | None = None
        self._scroll_anim_id: str | None = None

        self._build_ui()
        self.win.protocol("WM_DELETE_WINDOW", self._close)

        # Hide the main window while preview is open so the user can't
        # accidentally close the wrong one.
        try:
            self.parent.withdraw()
        except Exception:
            pass

    def _close(self):
        if self._closed:
            return
        self._closed = True
        try:
            self.win.destroy()
        except Exception:
            pass
        try:
            self.parent.deiconify()
        except Exception:
            pass
        if self._on_close_cb is not None:
            try:
                self._on_close_cb()
            except Exception:
                pass

    def _build_ui(self):
        # ---- TOP BAR ----
        top = ctk.CTkFrame(self.win, corner_radius=0, height=60)
        top.pack(side="top", fill="x")

        total = len(self.matched)
        matched_count = sum(1 for m in self.matched if m["matched_sheet"])
        summary = f"Notebook -> Excel  ·  {total} cubes, {matched_count} matched"
        ctk.CTkLabel(
            top, text=summary,
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(side="left", padx=20, pady=12)

        scanned = self.scan_result.get("scanned_count", 0)
        start = self.scan_result.get("start_sheet", "?")
        found_all = self.scan_result.get("found_all", False)
        scan_info = f"Scan: {scanned} sheets from '{start}'"
        if not found_all:
            scan_info += "  (NOT ALL FOUND!)"
        ctk.CTkLabel(
            top, text=scan_info,
            text_color=("#2E7D32" if found_all else "#C62828"),
            font=ctk.CTkFont(size=11),
        ).pack(side="left")

        legend = ctk.CTkFrame(top, fg_color="transparent")
        legend.pack(side="left", padx=20)
        ctk.CTkLabel(
            legend, text="Green = empty in Excel (fill)",
            text_color=EMPTY_BORDER,
            font=ctk.CTkFont(size=11, weight="bold"),
        ).pack(side="left", padx=(0, 12))
        ctk.CTkLabel(
            legend, text="Gray = already filled (skip)",
            text_color="gray55",
            font=ctk.CTkFont(size=11),
        ).pack(side="left")

        if self.match_error:
            ctk.CTkLabel(
                top, text=f"  Excel error: {self.match_error}",
                text_color="#C62828",
            ).pack(side="left")

        ctk.CTkButton(
            top, text="Cancel", width=90,
            fg_color="gray70", hover_color="gray60",
            command=self._close,
        ).pack(side="right", padx=(0, 10), pady=12)

        self._toggle_img_btn = ctk.CTkButton(
            top, text="Hide image", width=110,
            command=self._toggle_image_panel,
        )
        self._toggle_img_btn.pack(side="right", padx=(0, 10), pady=12)

        self.top_write_btn = ctk.CTkButton(
            top, text=">> WRITE TO EXCEL",
            width=200, height=36,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#2E7D32", hover_color="#1B5E20",
            command=self._do_write,
        )
        self.top_write_btn.pack(side="right", padx=(10, 15), pady=12)

        # ---- BODY ----
        body = ctk.CTkFrame(self.win, corner_radius=0, fg_color="transparent")
        body.pack(side="top", fill="both", expand=True, padx=10, pady=10)

        # ---- LEFT: image panel ----
        self._img_panel = ctk.CTkFrame(body, width=720, corner_radius=10)
        self._img_panel.pack(side="left", fill="y", padx=(0, 10))
        self._img_panel.pack_propagate(False)
        self._img_visible = True

        img_toolbar = ctk.CTkFrame(self._img_panel, corner_radius=0, height=40)
        img_toolbar.pack(side="top", fill="x")
        ctk.CTkButton(
            img_toolbar, text="-", width=40, height=28,
            command=self._zoom_out,
        ).pack(side="left", padx=4, pady=6)
        self.zoom_label = ctk.CTkLabel(img_toolbar, text="100%", width=55)
        self.zoom_label.pack(side="left")
        ctk.CTkButton(
            img_toolbar, text="+", width=40, height=28,
            command=self._zoom_in,
        ).pack(side="left", padx=4)
        ctk.CTkButton(
            img_toolbar, text="Fit", width=55, height=28,
            command=self._zoom_fit,
        ).pack(side="left", padx=(8, 2))
        ctk.CTkButton(
            img_toolbar, text="Actual", width=65, height=28,
            command=self._zoom_100,
        ).pack(side="left", padx=2)
        ctk.CTkLabel(
            img_toolbar,
            text="Ctrl+wheel: zoom | Ctrl++/- | Ctrl+0: fit | Ctrl+1: 100%",
            text_color="gray50",
            font=ctk.CTkFont(size=10),
        ).pack(side="left", padx=10)

        # Canvas for the image (kept as tkinter Canvas — CTk doesn't wrap it)
        canvas_wrap = ctk.CTkFrame(
            self._img_panel, fg_color="gray85", corner_radius=6,
        )
        canvas_wrap.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        self._img_canvas = Canvas(
            canvas_wrap,
            bg="gray85",
            width=IMAGE_CANVAS_WIDTH,
            height=IMAGE_CANVAS_HEIGHT,
            highlightthickness=0,
            takefocus=True,
        )
        img_vsb = Scrollbar(
            canvas_wrap, orient="vertical", command=self._img_canvas.yview
        )
        img_hsb = Scrollbar(
            canvas_wrap, orient="horizontal", command=self._img_canvas.xview
        )
        self._img_canvas.configure(
            yscrollcommand=img_vsb.set, xscrollcommand=img_hsb.set
        )
        img_vsb.pack(side=RIGHT, fill=Y)
        img_hsb.pack(side="bottom", fill="x")
        self._img_canvas.pack(side=LEFT, fill=BOTH, expand=True)

        self._img_canvas.bind("<Enter>", self._bind_img_wheel)
        self._img_canvas.bind("<Leave>", self._unbind_img_wheel)
        self._img_canvas.bind("<Control-MouseWheel>", self._on_ctrl_wheel)

        self._img_key_bindings = [
            ("<Control-plus>", self._on_key_zoom_in),
            ("<Control-KP_Add>", self._on_key_zoom_in),
            ("<Control-equal>", self._on_key_zoom_in),
            ("<Control-minus>", self._on_key_zoom_out),
            ("<Control-KP_Subtract>", self._on_key_zoom_out),
            ("<Control-Key-0>", self._on_key_zoom_fit),
            ("<Control-Key-1>", self._on_key_zoom_100),
            ("<Prior>", self._on_key_page_up),
            ("<Next>", self._on_key_page_down),
            ("<Home>", self._on_key_home),
            ("<End>", self._on_key_end),
            ("<Up>", self._on_key_arrow("up")),
            ("<Down>", self._on_key_arrow("down")),
            ("<Left>", self._on_key_arrow("left")),
            ("<Right>", self._on_key_arrow("right")),
        ]
        self.win.after(80, self._zoom_fit)

        # ---- RIGHT: scrollable cube list ----
        self._cube_list = ctk.CTkScrollableFrame(body, corner_radius=10)
        self._cube_list.pack(side="right", fill="both", expand=True)

        # Faster wheel scrolling
        try:
            self._cube_list._parent_canvas.configure(yscrollincrement=20)
        except Exception:
            pass

        for i, m in enumerate(self.matched):
            self._build_cube_card(self._cube_list, i, m)

        # Arrow-key card navigation (Entry widgets don't use Up/Down so
        # they bubble up to the window binding).
        self.win.bind("<Down>", self._on_card_down)
        self.win.bind("<Up>", self._on_card_up)

    # ---------- Zoom & image canvas helpers ----------

    def _render_image(self, resample=Image.Resampling.LANCZOS):
        """Place the PIL image on the canvas at the current zoom level.

        Pass resample=Image.Resampling.BILINEAR during interactive zoom for
        responsiveness; the default LANCZOS is crisper but slower on big pages.
        """
        if self._img_canvas is None:
            return
        w, h = self.orig_image.size
        new_w = max(1, int(w * self.zoom))
        new_h = max(1, int(h * self.zoom))
        resized = self.orig_image.resize((new_w, new_h), resample)
        # Keep the PhotoImage reference alive (avoid garbage collection)
        self._img_photo = ImageTk.PhotoImage(resized)
        if self._img_item is not None:
            self._img_canvas.delete(self._img_item)
        self._img_item = self._img_canvas.create_image(
            0, 0, anchor="nw", image=self._img_photo
        )
        self._img_canvas.configure(scrollregion=(0, 0, new_w, new_h))
        if self.zoom_label is not None:
            self.zoom_label.configure(text=f"{int(self.zoom * 100)}%")

    def _render_fast_then_hq(self):
        """Fast render now, schedule a crisp re-render when the user stops."""
        self._render_image(resample=Image.Resampling.BILINEAR)
        if self._hq_after_id is not None:
            try:
                self.win.after_cancel(self._hq_after_id)
            except Exception:
                pass
        self._hq_after_id = self.win.after(200, self._hq_rerender)

    def _hq_rerender(self):
        self._hq_after_id = None
        self._render_image(resample=Image.Resampling.LANCZOS)

    def _zoom_fit(self):
        """Fit the whole image into the canvas."""
        if self._img_canvas is None:
            return
        cw = self._img_canvas.winfo_width()
        ch = self._img_canvas.winfo_height()
        # If the canvas hasn't been rendered yet, use the default size
        if cw < 50:
            cw = IMAGE_CANVAS_WIDTH
        if ch < 50:
            ch = IMAGE_CANVAS_HEIGHT
        w, h = self.orig_image.size
        self.zoom = min(cw / w, ch / h, 1.0)
        self._render_image()

    def _zoom_100(self):
        self.zoom = 1.0
        self._render_image()

    def _zoom_in(self, interactive: bool = False):
        new_zoom = min(self.zoom * ZOOM_STEP, ZOOM_MAX)
        if abs(new_zoom - self.zoom) > 0.001:
            self.zoom = new_zoom
            if interactive:
                self._render_fast_then_hq()
            else:
                self._render_image()

    def _zoom_out(self, interactive: bool = False):
        new_zoom = max(self.zoom / ZOOM_STEP, ZOOM_MIN)
        if abs(new_zoom - self.zoom) > 0.001:
            self.zoom = new_zoom
            if interactive:
                self._render_fast_then_hq()
            else:
                self._render_image()

    def _on_img_wheel(self, event):
        if self._img_canvas is None:
            return
        # Shift+wheel: horizontal scroll
        if event.state & 0x0001:  # Shift
            self._img_canvas.xview_scroll(
                int(-1 * (event.delta / 120)), "units"
            )
        else:
            self._img_canvas.yview_scroll(
                int(-1 * (event.delta / 120)), "units"
            )

    def _on_ctrl_wheel(self, event):
        if event.delta > 0:
            self._zoom_in(interactive=True)
        else:
            self._zoom_out(interactive=True)
        return "break"

    # ---- Keyboard shortcuts (Excel-like) ----

    def _on_key_zoom_in(self, _event=None):
        self._zoom_in(interactive=True)
        return "break"

    def _on_key_zoom_out(self, _event=None):
        self._zoom_out(interactive=True)
        return "break"

    def _on_key_zoom_fit(self, _event=None):
        self._zoom_fit()
        return "break"

    def _on_key_zoom_100(self, _event=None):
        self._zoom_100()
        return "break"

    def _on_key_page_down(self, _event=None):
        if self._img_canvas is not None:
            self._img_canvas.yview_scroll(1, "pages")
        return "break"

    def _on_key_page_up(self, _event=None):
        if self._img_canvas is not None:
            self._img_canvas.yview_scroll(-1, "pages")
        return "break"

    def _on_key_home(self, _event=None):
        if self._img_canvas is not None:
            self._img_canvas.yview_moveto(0.0)
        return "break"

    def _on_key_end(self, _event=None):
        if self._img_canvas is not None:
            self._img_canvas.yview_moveto(1.0)
        return "break"

    def _toggle_image_panel(self):
        """Show/hide the PDF preview panel so cards use full width.

        CTkScrollableFrame wraps an internal canvas, so `before=` fails.
        Instead, forget both and re-pack in the correct order.
        """
        if self._img_visible:
            self._img_panel.pack_forget()
            self._toggle_img_btn.configure(text="Show image")
            self._img_visible = False
            self._apply_card_font_sizes(big=True)
        else:
            self._cube_list.pack_forget()
            self._img_panel.pack(side="left", fill="y", padx=(0, 10))
            self._cube_list.pack(side="right", fill="both", expand=True)
            self._toggle_img_btn.configure(text="Hide image")
            self._img_visible = True
            self._apply_card_font_sizes(big=False)

    def _apply_card_font_sizes(self, big: bool):
        """Resize fonts used on cube cards so values are easier to verify
        when the image panel is hidden."""
        if big:
            self._entry_font.configure(size=19)
            self._age_font.configure(size=16)
            self._title_font.configure(size=18)
            self._pill_font.configure(size=13)
            self._colhdr_font.configure(size=13)
            self._check_font.configure(size=14)
        else:
            self._entry_font.configure(size=15)
            self._age_font.configure(size=13)
            self._title_font.configure(size=14)
            self._pill_font.configure(size=11)
            self._colhdr_font.configure(size=11)
            self._check_font.configure(size=12)

    def _on_key_arrow(self, direction: str):
        def handler(_event=None):
            if self._img_canvas is None:
                return "break"
            if direction == "up":
                self._img_canvas.yview_scroll(-3, "units")
            elif direction == "down":
                self._img_canvas.yview_scroll(3, "units")
            elif direction == "left":
                self._img_canvas.xview_scroll(-3, "units")
            elif direction == "right":
                self._img_canvas.xview_scroll(3, "units")
            return "break"
        return handler

    def _bind_img_wheel(self, _event):
        self.win.bind_all("<MouseWheel>", self._on_img_wheel)
        self.win.bind_all("<Shift-MouseWheel>", self._on_img_wheel)
        # Activate Excel-like keyboard shortcuts while hovering the image
        for seq, handler in self._img_key_bindings:
            self.win.bind_all(seq, handler)

    def _unbind_img_wheel(self, _event):
        try:
            self.win.unbind_all("<MouseWheel>")
            self.win.unbind_all("<Shift-MouseWheel>")
            for seq, _handler in self._img_key_bindings:
                self.win.unbind_all(seq)
        except Exception:
            pass

    # ---------- Cube cards (CustomTkinter) ----------

    def _build_cube_card(self, parent, index: int, match: dict):
        cube = match["cube"]
        if cube.get("_shotcrete"):
            self._build_shotcrete_card(parent, index, match)
            return
        sheet = match["matched_sheet"]
        mark = cube.get("sample_mark", "?")

        card = ctk.CTkFrame(parent, corner_radius=12, border_width=1)
        card.pack(fill="x", padx=8, pady=8)
        self._cards.append(card)

        # Click anywhere on card background → select this card
        def _on_card_click(_e, i=index):
            self._select_card(i)
        card.bind("<Button-1>", _on_card_click)

        # ---- Header: 3-column grid keeps title truly centered ----
        header = ctk.CTkFrame(card, fg_color="transparent")
        header.pack(fill="x", padx=15, pady=(12, 6))
        header.grid_columnconfigure(0, weight=1, uniform="h")
        header.grid_columnconfigure(1, weight=0)
        header.grid_columnconfigure(2, weight=1, uniform="h")

        cube_enabled = BooleanVar(value=True)
        cube_checkbox = ctk.CTkCheckBox(
            header, text="Write this sample",
            variable=cube_enabled, font=self._check_font,
        )
        cube_checkbox.grid(row=0, column=0, sticky="w")

        title_txt = f"Cube No {cube.get('cube_no','?')}  ·  {mark}"
        set_total = cube.get("_set_total")
        if set_total and set_total > 1:
            title_txt += f"  ·  set {cube.get('_set_index', '?')}/{set_total}"
        ctk.CTkLabel(
            header, text=title_txt, font=self._title_font,
        ).grid(row=0, column=1, sticky="")

        if sheet:
            pill_text = f"[OK]  {sheet['sheet']}"
            pill_color = PILL_OK_COLOR
        else:
            pill_text = "[NO MATCH]"
            pill_color = PILL_NO_COLOR
        ctk.CTkLabel(
            header, text=pill_text,
            fg_color=pill_color, text_color="white",
            corner_radius=14, font=self._pill_font,
            padx=10, pady=3,
        ).grid(row=0, column=2, sticky="e")

        # ---- Read Excel state (empty vs filled) for all 6 cells ----
        excel_w7_empty = [True, True, True]
        excel_l7_empty = [True, True, True]
        excel_w28_empty = [True, True, True]
        excel_l28_empty = [True, True, True]
        excel_vals = None
        if sheet:
            try:
                excel_vals = writer.read_all_values(
                    sheet["workbook"], sheet["sheet"]
                )
                excel_w7_empty = [writer.is_cell_empty(v)
                                  for v in excel_vals["weights_7d"]]
                excel_l7_empty = [writer.is_cell_empty(v)
                                  for v in excel_vals["loads_7d"]]
                excel_w28_empty = [writer.is_cell_empty(v)
                                   for v in excel_vals["weights_28d"]]
                excel_l28_empty = [writer.is_cell_empty(v)
                                   for v in excel_vals["loads_28d"]]
            except Exception:
                pass

        # ---- 7-day / 28-day test rows (pad to 3 each) ----
        tests_7 = [t for t in cube.get("tests", []) if t.get("age_days") == 7]
        tests_28 = [t for t in cube.get("tests", []) if t.get("age_days") == 28]
        while len(tests_7) < 3:
            tests_7.append({"weight_gr": None, "load_kn": None})
        while len(tests_28) < 3:
            tests_28.append({"weight_gr": None, "load_kn": None})

        # ---- Value grid ----
        grid = ctk.CTkFrame(card, fg_color="transparent")
        grid.pack(padx=15, pady=(2, 12))

        # Column headers
        ctk.CTkLabel(grid, text="", width=90).grid(row=0, column=0)
        ctk.CTkLabel(
            grid, text="Weight (gr)", width=180, font=self._colhdr_font,
        ).grid(row=0, column=1, padx=8)
        ctk.CTkLabel(
            grid, text="Load (kN)", width=180, font=self._colhdr_font,
        ).grid(row=0, column=2, padx=8)
        ctk.CTkLabel(grid, text="", width=70).grid(row=0, column=3)

        def _build_group(start_row, age_label, age_color, tests,
                         w_empty_list, l_empty_list):
            # Age label pinned left, spanning 3 rows
            ctk.CTkLabel(
                grid, text=age_label, text_color=age_color,
                font=self._age_font, width=90, anchor="w",
            ).grid(
                row=start_row, column=0, rowspan=3, sticky="w",
                padx=(0, 4), pady=(8, 4),
            )

            w_vars, l_vars = [], []
            w_widgets, l_widgets = [], []
            for i, t in enumerate(tests[:3]):
                row = start_row + i
                weight = t.get("weight_gr")
                load = t.get("load_kn")
                w_var = StringVar(
                    value="" if weight is None else str(weight)
                )
                l_var = StringVar(
                    value="" if load is None else str(load)
                )

                w_kwargs = dict(
                    textvariable=w_var, width=180, height=38,
                    justify="center", font=self._entry_font,
                )
                if w_empty_list[i] and weight is not None:
                    # Excel empty + notebook has a value → will be written
                    w_kwargs["border_color"] = EMPTY_BORDER
                    w_kwargs["border_width"] = 3
                    w_kwargs["fg_color"] = EMPTY_BG
                elif not w_empty_list[i]:
                    # Already filled in Excel
                    w_kwargs["text_color"] = FILLED_TEXT
                # else: empty in Excel + no notebook value → plain

                l_kwargs = dict(
                    textvariable=l_var, width=180, height=38,
                    justify="center", font=self._entry_font,
                )
                if l_empty_list[i] and load is not None:
                    l_kwargs["border_color"] = EMPTY_BORDER
                    l_kwargs["border_width"] = 3
                    l_kwargs["fg_color"] = EMPTY_BG
                elif not l_empty_list[i]:
                    l_kwargs["text_color"] = FILLED_TEXT

                w_entry = ctk.CTkEntry(grid, **w_kwargs)
                w_entry.grid(row=row, column=1, padx=8, pady=4)
                l_entry = ctk.CTkEntry(grid, **l_kwargs)
                l_entry.grid(row=row, column=2, padx=8, pady=4)
                w_vars.append(w_var)
                l_vars.append(l_var)
                w_widgets.append(w_entry)
                l_widgets.append(l_entry)

            # Per-GROUP checkbox — auto-ticked only if there is work to do:
            # at least one empty Excel cell AND at least one notebook value.
            any_excel_empty = any(w_empty_list) or any(l_empty_list)
            any_value_to_write = any(
                t.get("weight_gr") is not None or t.get("load_kn") is not None
                for t in tests[:3]
            )
            group_var = BooleanVar(value=any_excel_empty and any_value_to_write)
            ctk.CTkCheckBox(
                grid, text="", variable=group_var, width=24,
            ).grid(
                row=start_row, column=3, rowspan=3, sticky="",
                padx=(12, 0), pady=3,
            )
            return w_vars, l_vars, group_var, w_widgets, l_widgets

        w7, l7, c7, ww7, lw7 = _build_group(
            1, "7-day", AGE_7_COLOR, tests_7,
            excel_w7_empty, excel_l7_empty,
        )
        w28, l28, c28, ww28, lw28 = _build_group(
            4, "28-day", AGE_28_COLOR, tests_28,
            excel_w28_empty, excel_l28_empty,
        )

        # If both groups have no work, uncheck the cube master too
        if not c7.get() and not c28.get():
            cube_enabled.set(False)

        self.entries.append({
            "cube": cube,
            "matched_sheet": sheet,
            "cube_enabled": cube_enabled,
            "weights_7d": w7, "loads_7d": l7, "check_7d": c7,
            "weights_28d": w28, "loads_28d": l28, "check_28d": c28,
            # widget refs for focus navigation
            "weight_widgets_7d": ww7, "load_widgets_7d": lw7,
            "weight_empty_7d": excel_w7_empty, "load_empty_7d": excel_l7_empty,
            "weight_widgets_28d": ww28, "load_widgets_28d": lw28,
            "weight_empty_28d": excel_w28_empty, "load_empty_28d": excel_l28_empty,
        })

        # ---- 7-day cross-check warnings ----
        if sheet and excel_vals is not None:
            try:
                ex7 = {
                    "weights": excel_vals["weights_7d"],
                    "loads": excel_vals["loads_7d"],
                }
                warnings = writer.cross_check_7day(cube, ex7)
                if warnings:
                    wf = ctk.CTkFrame(card, fg_color="transparent")
                    wf.pack(fill="x", padx=15, pady=(0, 10))
                    ctk.CTkLabel(
                        wf, text="7-day check warnings:",
                        text_color="#E65100",
                        font=ctk.CTkFont(size=11, weight="bold"),
                    ).pack(anchor="w")
                    for w in warnings:
                        ctk.CTkLabel(
                            wf, text="  - " + w, text_color="#E65100",
                        ).pack(anchor="w")
            except Exception as e:
                ctk.CTkLabel(
                    card, text=f"7-day check failed: {e}",
                    text_color="gray55",
                ).pack(anchor="w", padx=15, pady=(0, 8))

    # ---------- Shotcrete card (5 rows per age, top-3 by strength) ----------

    def _build_shotcrete_card(self, parent, index: int, match: dict):
        cube = match["cube"]
        sheet = match["matched_sheet"]
        mark = cube.get("sample_mark", "?")

        card = ctk.CTkFrame(parent, corner_radius=12, border_width=1)
        card.pack(fill="x", padx=8, pady=8)
        self._cards.append(card)

        def _on_card_click(_e, i=index):
            self._select_card(i)
        card.bind("<Button-1>", _on_card_click)

        # ---- Header ----
        header = ctk.CTkFrame(card, fg_color="transparent")
        header.pack(fill="x", padx=15, pady=(12, 6))
        header.grid_columnconfigure(0, weight=1, uniform="h")
        header.grid_columnconfigure(1, weight=0)
        header.grid_columnconfigure(2, weight=1, uniform="h")

        left_box = ctk.CTkFrame(header, fg_color="transparent")
        left_box.grid(row=0, column=0, sticky="w")
        cube_enabled = BooleanVar(value=sheet is not None)
        ctk.CTkCheckBox(
            left_box, text="Write this sample",
            variable=cube_enabled, font=self._check_font,
        ).pack(side="left")
        ctk.CTkLabel(
            left_box, text="SHOTCRETE",
            fg_color=PILL_SHOT_COLOR, text_color="white", corner_radius=10,
            font=ctk.CTkFont(size=10, weight="bold"), padx=8, pady=2,
        ).pack(side="left", padx=(10, 0))

        title_txt = f"Cube No {cube.get('cube_no','?')}  ·  {mark}"
        set_total = cube.get("_set_total")
        if set_total and set_total > 1:
            title_txt += f"  ·  set {cube.get('_set_index', '?')}/{set_total}"
        ctk.CTkLabel(
            header, text=title_txt, font=self._title_font,
        ).grid(row=0, column=1, sticky="")

        if sheet:
            pill_text = f"[OK]  {sheet['sheet']}"
            pill_color = PILL_OK_COLOR
        else:
            pill_text = "[NO MATCH]"
            pill_color = PILL_NO_COLOR
        ctk.CTkLabel(
            header, text=pill_text,
            fg_color=pill_color, text_color="white",
            corner_radius=14, font=self._pill_font,
            padx=10, pady=3,
        ).grid(row=0, column=2, sticky="e")

        # ---- Rows ----
        tests_7 = [t for t in cube.get("tests", []) if t.get("age_days") == 7]
        tests_28 = [t for t in cube.get("tests", []) if t.get("age_days") == 28]
        # Pad to 5 so the UI always shows 5 slots even on partial reads.
        while len(tests_7) < 5:
            tests_7.append({})
        while len(tests_28) < 5:
            tests_28.append({})

        grid = ctk.CTkFrame(card, fg_color="transparent")
        grid.pack(padx=15, pady=(2, 12))

        ctk.CTkLabel(grid, text="", width=55).grid(row=0, column=0)
        ctk.CTkLabel(grid, text="", width=24).grid(row=0, column=1)
        ctk.CTkLabel(grid, text="Diameter", width=80,
                     font=self._colhdr_font).grid(row=0, column=2, padx=3)
        ctk.CTkLabel(grid, text="Height", width=80,
                     font=self._colhdr_font).grid(row=0, column=3, padx=3)
        ctk.CTkLabel(grid, text="Weight (gr)", width=100,
                     font=self._colhdr_font).grid(row=0, column=4, padx=3)
        ctk.CTkLabel(grid, text="Load (kN)", width=100,
                     font=self._colhdr_font).grid(row=0, column=5, padx=3)
        ctk.CTkLabel(grid, text="Strength", width=100,
                     font=self._colhdr_font).grid(row=0, column=6, padx=3)

        rows_7d = self._build_shotcrete_group(
            grid, 1, "7-day", AGE_7_COLOR, tests_7[:5]
        )
        rows_28d = self._build_shotcrete_group(
            grid, 7, "28-day", AGE_28_COLOR, tests_28[:5]
        )

        # Per-group enable checkbox (re-uses 7d / 28d lists)
        check_7d = BooleanVar(value=True)
        check_28d = BooleanVar(value=True)

        # For arrow-key focus navigation: flatten only the selected rows'
        # weight/load widgets so navigation behaves like normal cubes.
        def _selected_widgets(rows, field):
            out = []
            for r in rows:
                if r["sel"].get():
                    out.append(r[field])
            return out

        self.entries.append({
            "cube": cube,
            "matched_sheet": sheet,
            "shotcrete": True,
            "cube_enabled": cube_enabled,
            "shot_rows_7d": rows_7d,
            "shot_rows_28d": rows_28d,
            "check_7d": check_7d,
            "check_28d": check_28d,
            # Focus navigation fallbacks — point at the selected rows' entries.
            "weight_widgets_7d": [r["weight_ent"] for r in rows_7d],
            "load_widgets_7d": [r["load_ent"] for r in rows_7d],
            "weight_empty_7d": [True] * len(rows_7d),
            "load_empty_7d": [True] * len(rows_7d),
            "weight_widgets_28d": [r["weight_ent"] for r in rows_28d],
            "load_widgets_28d": [r["load_ent"] for r in rows_28d],
            "weight_empty_28d": [True] * len(rows_28d),
            "load_empty_28d": [True] * len(rows_28d),
        })

    def _build_shotcrete_group(self, grid, start_row, age_label, age_color, tests):
        """Build one 5-row shotcrete group. Returns a list of row dicts."""
        ctk.CTkLabel(
            grid, text=age_label, text_color=age_color,
            font=self._age_font, width=70, anchor="w",
        ).grid(
            row=start_row, column=0, rowspan=5, sticky="w",
            padx=(0, 4), pady=(6, 4),
        )

        def _strength_val(t):
            s = t.get("strength_nmm2")
            try:
                return float(s)
            except (TypeError, ValueError):
                return float("-inf")

        # Top 3 indices by strength (uses reader's _selected tag if set,
        # otherwise computes fresh from the strength values).
        # Rows with no strength value are never auto-selected.
        if any("_selected" in t for t in tests):
            top3 = {i for i, t in enumerate(tests) if t.get("_selected")}
        else:
            valid = [i for i in range(len(tests))
                     if _strength_val(tests[i]) != float("-inf")]
            ranked = sorted(valid, key=lambda i: _strength_val(tests[i]), reverse=True)
            top3 = set(ranked[:3])

        rows: list[dict] = []

        def _style_row(r: dict):
            on = r["sel"].get()
            for ent in (r["diam_ent"], r["height_ent"],
                        r["weight_ent"], r["load_ent"], r["strength_ent"]):
                if on:
                    ent.configure(
                        border_color=EMPTY_BORDER, border_width=3,
                        fg_color=EMPTY_BG, text_color=("gray14", "gray84"),
                    )
                else:
                    ent.configure(
                        border_color=("gray70", "gray30"), border_width=1,
                        fg_color=DIM_BG, text_color=FILLED_TEXT,
                    )

        def _on_toggle(idx):
            on_now = [i for i, r in enumerate(rows) if r["sel"].get()]
            if rows[idx]["sel"].get() and len(on_now) > 3:
                others = [i for i in on_now if i != idx]
                drop = min(others, key=lambda i: _strength_val(tests[i]))
                rows[drop]["sel"].set(False)
                _style_row(rows[drop])
            _style_row(rows[idx])

        for i, t in enumerate(tests):
            r_idx = start_row + i
            sel = BooleanVar(value=i in top3)

            ctk.CTkCheckBox(
                grid, text="", variable=sel, width=24,
                command=lambda ii=i: _on_toggle(ii),
            ).grid(row=r_idx, column=1, padx=(0, 4), pady=2)

            def _var(val):
                return StringVar(value="" if val is None else str(val))

            diam_var   = _var(t.get("core_diameter_mm"))
            height_var = _var(t.get("core_height_mm"))
            weight_var = _var(t.get("weight_gr"))
            load_var   = _var(t.get("load_kn"))
            strength_var = _var(t.get("strength_nmm2"))

            def _ent(var, w):
                return ctk.CTkEntry(
                    grid, textvariable=var, width=w, height=32,
                    justify="center", font=self._entry_font,
                )

            diam_ent     = _ent(diam_var, 110)
            height_ent   = _ent(height_var, 110)
            weight_ent   = _ent(weight_var, 130)
            load_ent     = _ent(load_var, 130)
            strength_ent = _ent(strength_var, 130)

            diam_ent.grid    (row=r_idx, column=2, padx=5, pady=2)
            height_ent.grid  (row=r_idx, column=3, padx=5, pady=2)
            weight_ent.grid  (row=r_idx, column=4, padx=5, pady=2)
            load_ent.grid    (row=r_idx, column=5, padx=5, pady=2)
            strength_ent.grid(row=r_idx, column=6, padx=5, pady=2)

            row = {
                "sel": sel,
                "diam_var": diam_var, "height_var": height_var,
                "weight_var": weight_var, "load_var": load_var,
                "strength_var": strength_var,
                "diam_ent": diam_ent, "height_ent": height_ent,
                "weight_ent": weight_ent, "load_ent": load_ent,
                "strength_ent": strength_ent,
            }
            rows.append(row)
            _style_row(row)

        return rows

    # ---------- Card selection & smooth scroll ----------

    def _on_card_down(self, _event=None):
        if not self._cards:
            return "break"
        if self._selected_idx is None:
            self._select_card(0)
        else:
            self._select_card(min(self._selected_idx + 1, len(self._cards) - 1))
        return "break"

    def _on_card_up(self, _event=None):
        if not self._cards:
            return "break"
        if self._selected_idx is None:
            self._select_card(0)
        else:
            self._select_card(max(self._selected_idx - 1, 0))
        return "break"

    def _select_card(self, idx: int):
        if idx < 0 or idx >= len(self._cards):
            return
        # Deselect previous
        if self._selected_idx is not None and self._selected_idx != idx:
            try:
                self._cards[self._selected_idx].configure(
                    border_color=("gray70", "gray30"), border_width=1,
                )
            except Exception:
                pass
        # Highlight new
        self._selected_idx = idx
        try:
            self._cards[idx].configure(
                border_color=SELECTED_BORDER, border_width=3,
            )
        except Exception:
            pass
        self._scroll_to_card_centered(idx)
        self._focus_first_empty_entry(idx)

    def _scroll_to_card_centered(self, idx: int):
        """Animate the cube list so card idx sits vertically centered."""
        canvas = getattr(self._cube_list, "_parent_canvas", None)
        if canvas is None:
            return
        card = self._cards[idx]
        self.win.update_idletasks()

        bbox = canvas.bbox("all")
        if not bbox:
            return
        total_h = bbox[3] - bbox[1]
        viewport_h = canvas.winfo_height()
        if total_h <= viewport_h:
            return  # nothing to scroll

        card_y = card.winfo_y()
        card_h = card.winfo_height()
        card_center = card_y + card_h / 2

        target_top = card_center - viewport_h / 2
        target_top = max(0, min(target_top, total_h - viewport_h))
        target_frac = target_top / total_h

        self._animate_scroll(canvas, target_frac, duration_ms=220)

    def _animate_scroll(self, canvas, target_frac: float, duration_ms: int = 220):
        """Ease-out animated yview_moveto."""
        if self._scroll_anim_id is not None:
            try:
                self.win.after_cancel(self._scroll_anim_id)
            except Exception:
                pass
            self._scroll_anim_id = None

        frame_ms = 16
        frames = max(1, duration_ms // frame_ms)
        try:
            start_frac = canvas.yview()[0]
        except Exception:
            start_frac = 0.0

        def step(i):
            t = i / frames
            ease = 1 - (1 - t) ** 3  # ease-out cubic
            pos = start_frac + (target_frac - start_frac) * ease
            try:
                canvas.yview_moveto(pos)
            except Exception:
                return
            if i < frames:
                self._scroll_anim_id = self.win.after(
                    frame_ms, lambda: step(i + 1)
                )
            else:
                self._scroll_anim_id = None

        step(1)

    def _focus_first_empty_entry(self, idx: int):
        """Focus the first Excel-empty Weight/Load entry of the card, so
        the user can start typing immediately after arrow-navigating."""
        entry = self.entries[idx]
        groups = [
            ("weight_widgets_7d", "weight_empty_7d"),
            ("load_widgets_7d", "load_empty_7d"),
            ("weight_widgets_28d", "weight_empty_28d"),
            ("load_widgets_28d", "load_empty_28d"),
        ]
        for widget_key, empty_key in groups:
            widgets = entry.get(widget_key, [])
            empties = entry.get(empty_key, [])
            for w, is_empty in zip(widgets, empties):
                if is_empty:
                    try:
                        w.focus_set()
                    except Exception:
                        pass
                    return
        # Nothing empty — focus first weight_7d widget just to lock focus
        widgets = entry.get("weight_widgets_7d", [])
        if widgets:
            try:
                widgets[0].focus_set()
            except Exception:
                pass

    def _do_write(self):
        """Write to Excel button handler."""
        written_sheets = 0
        total_cells = 0
        all_errors: list[str] = []
        skipped: list[str] = []

        def _parse_list(var_list, label, mark):
            out = []
            for v in var_list:
                txt = v.get().strip()
                if not txt:
                    out.append(None)
                    continue
                try:
                    out.append(float(txt))
                except ValueError:
                    all_errors.append(
                        f"{mark}: invalid {label} '{txt}'"
                    )
                    out.append(None)
            return out

        def _parse_shot_rows(rows, label, mark):
            """Extract (weight, load, diameter, height) for ONLY the 3 selected
            shotcrete rows in a group. Returns 4 lists of length 3 (None-padded)."""
            selected = [r for r in rows if r["sel"].get()]
            w_list, l_list, d_list, h_list = [], [], [], []

            def _to_float(var, fname):
                txt = var.get().strip()
                if not txt:
                    return None
                try:
                    return float(txt)
                except ValueError:
                    all_errors.append(f"{mark}: invalid {label} {fname} '{txt}'")
                    return None

            for r in selected[:3]:
                w_list.append(_to_float(r["weight_var"], "weight"))
                l_list.append(_to_float(r["load_var"], "load"))
                d_list.append(_to_float(r["diam_var"], "diameter"))
                h_list.append(_to_float(r["height_var"], "height"))
            while len(w_list) < 3:
                w_list.append(None); l_list.append(None)
                d_list.append(None); h_list.append(None)
            return w_list, l_list, d_list, h_list

        for entry in self.entries:
            sheet = entry["matched_sheet"]
            cube = entry["cube"]
            if not sheet:
                skipped.append(f"{cube.get('sample_mark','?')} (no match)")
                continue

            mark = cube.get("sample_mark", "?")

            # Master toggle: entire cube skipped (Excel left untouched)
            if not entry["cube_enabled"].get():
                skipped.append(f"{mark} (unchecked)")
                continue

            if entry.get("shotcrete"):
                w7, l7, d7, h7 = _parse_shot_rows(entry["shot_rows_7d"], "7d", mark)
                w28, l28, d28, h28 = _parse_shot_rows(entry["shot_rows_28d"], "28d", mark)

                if not entry["check_7d"].get():
                    w7 = [None]*3; l7 = [None]*3; d7 = [None]*3; h7 = [None]*3
                if not entry["check_28d"].get():
                    w28 = [None]*3; l28 = [None]*3; d28 = [None]*3; h28 = [None]*3

                try:
                    result = writer.write_cube(
                        cube,
                        sheet["workbook"], sheet["sheet"],
                        weights_7d=w7, loads_7d=l7,
                        weights_28d=w28, loads_28d=l28,
                        diameters_7d=d7, heights_7d=h7,
                        diameters_28d=d28, heights_28d=h28,
                    )
                    total_cells += len(result["wrote"])
                    all_errors.extend(result["errors"])
                    if result["wrote"]:
                        written_sheets += 1
                except Exception as e:
                    all_errors.append(f"{mark}: write error: {e}")
                continue

            weights_7d = _parse_list(entry["weights_7d"], "7d weight", mark)
            loads_7d = _parse_list(entry["loads_7d"], "7d load", mark)
            weights_28d = _parse_list(entry["weights_28d"], "28d weight", mark)
            loads_28d = _parse_list(entry["loads_28d"], "28d load", mark)

            # Per-group toggle: unchecking the 7-day or 28-day checkbox
            # blanks all 3 rows of that group so Excel is left untouched.
            if not entry["check_7d"].get():
                weights_7d = [None, None, None]
                loads_7d = [None, None, None]
            if not entry["check_28d"].get():
                weights_28d = [None, None, None]
                loads_28d = [None, None, None]

            try:
                result = writer.write_cube(
                    cube,
                    sheet["workbook"],
                    sheet["sheet"],
                    weights_7d=weights_7d,
                    loads_7d=loads_7d,
                    weights_28d=weights_28d,
                    loads_28d=loads_28d,
                )
                total_cells += len(result["wrote"])
                all_errors.extend(result["errors"])
                if result["wrote"]:
                    written_sheets += 1
            except Exception as e:
                all_errors.append(
                    f"{cube.get('sample_mark','?')}: write error: {e}"
                )

        # Summary message
        lines = [
            f"{written_sheets} sheets updated, {total_cells} cells written in total."
        ]
        if skipped:
            lines.append("")
            lines.append("Skipped (no match):")
            for s in skipped:
                lines.append("  - " + s)
        if all_errors:
            lines.append("")
            lines.append("Errors:")
            for e in all_errors:
                lines.append("  - " + e)

        messagebox.showinfo("Result", "\n".join(lines), parent=self.win)

        if not all_errors:
            self._close()


# ---------- Ledger Preview Window ----------

class LedgerPreviewWindow:
    """
    Second-pass preview for the downward-growing ledger Excel ("Concrete"
    sheet). Consumes the same in-memory cubes_data used by the first
    preview; no Gemini re-read.
    """

    def __init__(self, parent, cubes_data: dict, on_close=None):
        self.parent = parent
        self.cubes_data = cubes_data
        self._on_close_cb = on_close
        self._closed = False

        self.win = ctk.CTkToplevel(parent)
        self.win.title("Ledger — " + APP_TITLE)
        sw = self.win.winfo_screenwidth()
        sh = self.win.winfo_screenheight()
        ww = min(1200, sw - 60)
        wh = min(860, sh - 80)
        self.win.geometry(f"{ww}x{wh}+40+40")
        self.win.minsize(800, 600)
        try:
            self.win.state("zoomed")
        except Exception:
            pass
        self.win.protocol("WM_DELETE_WINDOW", self._close)

        _add_credit_footer(self.win)

        self._entry_font = ctk.CTkFont(family="Segoe UI", size=13)
        self._title_font = ctk.CTkFont(family="Segoe UI", size=14, weight="bold")
        self._hdr_font = ctk.CTkFont(family="Segoe UI", size=11, weight="bold")
        self._pill_font = ctk.CTkFont(family="Segoe UI", size=10, weight="bold")
        self._check_font = ctk.CTkFont(family="Segoe UI", size=12)
        self._age_font = ctk.CTkFont(family="Segoe UI", size=13, weight="bold")
        self._colhdr_font = ctk.CTkFont(family="Segoe UI", size=11, weight="bold")

        # Hide main window while this is open
        try:
            self.parent.withdraw()
        except Exception:
            pass

        # Load ledger data (synchronous — one UsedRange read)
        self.ledger_error: str | None = None
        self.wb_name: str | None = None
        self.sheet_name: str | None = None
        self.entries: list[dict] = []  # {cube, block, mismatch, weights_ledger, loads_ledger, enabled_var}
        self.not_found: list[str] = []

        # Arrow-key card navigation (same pattern as PreviewWindow)
        self._cards: list = []
        self._selected_idx: int | None = None
        self._cube_list = None  # assigned when the scrollable frame is built

        try:
            wb, ws, wb_name, sh_name = writer.find_ledger_sheet()
            self.wb_name = wb_name
            self.sheet_name = sh_name
            self._ws = ws  # keep COM reference alive while window is open
            blocks = writer.read_ledger_blocks(ws)
            values = writer.read_ledger_values(ws, blocks)
            merged = writer.merge_cubes_for_ledger(cubes_data)
            matches = writer.match_cubes_to_blocks(merged, blocks)

            # Diagnostic dump — every PDF cube key + every ledger block key
            try:
                from reader import _log
                _log(f"[LEDGER-DEBUG] {len(blocks)} ledger blocks, {len(merged)} PDF cubes")
                for c in merged:
                    _log(f"  PDF cube → key={c.get('sample_key')!r} cube_no={c.get('cube_no')!r} mark={c.get('sample_mark')!r}")
                # Only dump blocks whose key starts the same as any PDF key, plus any G26-CON
                pdf_keys = {c.get('sample_key') for c in merged}
                for b in blocks:
                    bk = b.get('sample_key')
                    if bk in pdf_keys or (isinstance(bk, str) and bk.upper().startswith('G26')):
                        _log(f"  LEDGER block → key={bk!r} cube_no={b.get('cube_no')!r} raw={b.get('sample_mark_raw')!r} rows={b['start_row']}-{b['end_row']}")
            except Exception as _e:
                pass

            block_idx_by_id = {id(b): i for i, b in enumerate(blocks)}
            for m in matches:
                cube = m["cube"]
                block = m["block"]
                reason = m["mismatch_reason"]
                if block is None:
                    self.not_found.append(cube.get("sample_mark", "?"))
                    continue
                bi = block_idx_by_id[id(block)]
                vals = values.get(bi, {"weights": [], "loads": []})
                self.entries.append({
                    "cube": cube,
                    "block": block,
                    "mismatch": reason,
                    "weights_ledger": vals["weights"],
                    "loads_ledger": vals["loads"],
                    "enabled": BooleanVar(value=reason is None),
                })
        except Exception as e:
            self.ledger_error = str(e)

        self._build_ui()

    def _close(self):
        if self._closed:
            return
        self._closed = True
        self._ws = None  # release COM reference
        try:
            self.win.destroy()
        except Exception:
            pass
        try:
            self.parent.deiconify()
        except Exception:
            pass
        if self._on_close_cb is not None:
            try:
                self._on_close_cb()
            except Exception:
                pass

    def _build_ui(self):
        # ---- TOP BAR ----
        top = ctk.CTkFrame(self.win, corner_radius=0, height=60)
        top.pack(side="top", fill="x")
        top.pack_propagate(False)

        if self.ledger_error:
            title = "Ledger Preview — ERROR"
        else:
            ok_cnt = sum(1 for e in self.entries if e["mismatch"] is None)
            bad_cnt = sum(1 for e in self.entries if e["mismatch"] is not None)
            title = (
                f"Ledger Preview — {self.sheet_name}  ·  "
                f"{ok_cnt} ready, {bad_cnt} mismatched, "
                f"{len(self.not_found)} not found"
            )
        ctk.CTkLabel(
            top, text=title,
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(side="left", padx=20, pady=16)

        ctk.CTkButton(
            top, text="Cancel", width=90,
            fg_color="gray70", hover_color="gray60",
            command=self._close,
        ).pack(side="right", padx=(0, 10), pady=12)

        self.write_btn = ctk.CTkButton(
            top, text=">> WRITE TO LEDGER",
            width=200, height=36,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#2E7D32", hover_color="#1B5E20",
            command=self._do_write,
        )
        self.write_btn.pack(side="right", padx=(0, 10), pady=12)

        # Disable write if there's nothing writable
        writable = any(e["mismatch"] is None for e in self.entries)
        if not writable:
            self.write_btn.configure(state="disabled")

        # ---- BODY ----
        body = ctk.CTkFrame(self.win, fg_color="transparent")
        body.pack(side="top", fill="both", expand=True, padx=16, pady=12)

        if self.ledger_error:
            ctk.CTkLabel(
                body,
                text="Ledger error:\n" + self.ledger_error,
                text_color="#C62828",
                font=self._entry_font,
                justify="left",
                wraplength=900,
            ).pack(pady=40)
            return

        # Legend + not-found banner
        legend = ctk.CTkFrame(body, fg_color="transparent")
        legend.pack(fill="x", pady=(0, 6))
        ctk.CTkLabel(
            legend, text="Green = ledger cell empty (will write)",
            text_color=EMPTY_BORDER,
            font=ctk.CTkFont(size=11, weight="bold"),
        ).pack(side="left", padx=(0, 16))
        ctk.CTkLabel(
            legend, text="Gray = already filled (skip)",
            text_color="gray55",
            font=ctk.CTkFont(size=11),
        ).pack(side="left")

        if self.not_found:
            banner = ctk.CTkFrame(
                body, corner_radius=8, border_width=2,
                border_color="#C62828", fg_color="#2a1010",
            )
            banner.pack(fill="x", pady=(4, 10))
            joined = ", ".join(self.not_found)
            ctk.CTkLabel(
                banner,
                text=f"⚠ {len(self.not_found)} cubes not found in ledger: {joined}",
                text_color="#FF8A80",
                font=ctk.CTkFont(size=12, weight="bold"),
                anchor="w", justify="left", wraplength=1100,
            ).pack(fill="x", padx=12, pady=8)

        # Scrollable card area
        scroll = ctk.CTkScrollableFrame(body, fg_color="transparent")
        scroll.pack(fill="both", expand=True)
        self._cube_list = scroll

        for entry in self.entries:
            self._build_card(scroll, entry)

        # Arrow-key navigation between cards
        self.win.bind("<Down>", self._on_card_down)
        self.win.bind("<Up>", self._on_card_up)

        if not self.entries and not self.not_found:
            ctk.CTkLabel(
                body,
                text="No normal cubes to write (shotcrete filtered out).",
                text_color="gray55",
                font=self._entry_font,
            ).pack(pady=20)

    def _build_card(self, parent, entry: dict):
        cube = entry["cube"]
        block = entry["block"]
        mismatch = entry["mismatch"]
        bad = mismatch is not None

        card = ctk.CTkFrame(parent, corner_radius=12, border_width=1)
        card.pack(fill="x", padx=8, pady=8)
        self._cards.append(card)

        # ---- Header: 3-column grid (checkbox | title | pill) ----
        header = ctk.CTkFrame(card, fg_color="transparent")
        header.pack(fill="x", padx=15, pady=(12, 6))
        header.grid_columnconfigure(0, weight=1, uniform="h")
        header.grid_columnconfigure(1, weight=0)
        header.grid_columnconfigure(2, weight=1, uniform="h")

        cube_checkbox = ctk.CTkCheckBox(
            header, text="Write this sample",
            variable=entry["enabled"], font=self._check_font,
        )
        cube_checkbox.grid(row=0, column=0, sticky="w")
        if bad:
            cube_checkbox.configure(state="disabled")

        title_txt = (
            f"Cube No {cube.get('cube_no', '?')}  ·  "
            f"{cube.get('sample_mark', '?')}"
        )
        n7 = len(cube.get("tests_7d", []))
        n28 = len(cube.get("tests_28d", []))
        if n7 > 3 or n28 > 3:
            if n7 % 3 == 0 and n28 % 3 == 0:
                sets = max(n7, n28) // 3
                if sets >= 2:
                    title_txt += f"  ·  {sets} set"
        ctk.CTkLabel(
            header, text=title_txt, font=self._title_font,
        ).grid(row=0, column=1, sticky="")

        if bad:
            pill_text = "[MISMATCH]"
            pill_color = PILL_NO_COLOR
        else:
            pill_text = f"[OK]  rows {block['start_row']}-{block['end_row']}"
            pill_color = PILL_OK_COLOR
        ctk.CTkLabel(
            header, text=pill_text,
            fg_color=pill_color, text_color="white",
            corner_radius=14, font=self._pill_font,
            padx=10, pady=3,
        ).grid(row=0, column=2, sticky="e")

        if bad:
            ctk.CTkLabel(
                card, text="⚠ " + mismatch,
                text_color="#FF8A80",
                font=ctk.CTkFont(size=11, weight="bold"),
                anchor="w",
            ).pack(anchor="w", padx=15, pady=(0, 6))

        # ---- Value grid (mirrors first preview style, no green border) ----
        grid = ctk.CTkFrame(card, fg_color="transparent")
        grid.pack(padx=15, pady=(2, 12))

        ctk.CTkLabel(grid, text="", width=90).grid(row=0, column=0)
        ctk.CTkLabel(
            grid, text="Weight (gr)", width=180, font=self._colhdr_font,
        ).grid(row=0, column=1, padx=8)
        ctk.CTkLabel(
            grid, text="Load (kN)", width=180, font=self._colhdr_font,
        ).grid(row=0, column=2, padx=8)
        ctk.CTkLabel(grid, text="", width=70).grid(row=0, column=3)

        weights_existing = entry["weights_ledger"]
        loads_existing = entry["loads_ledger"]
        block_start = block["start_row"]
        rows_7d = block.get("rows_7d", []) or []
        rows_28d = block.get("rows_28d", []) or []

        def _build_group(start_row, age_label, age_color, tests, age_rows):
            n_rows = max(len(tests), len(age_rows), 1)
            ctk.CTkLabel(
                grid, text=age_label, text_color=age_color,
                font=self._age_font, width=90, anchor="w",
            ).grid(
                row=start_row, column=0, rowspan=n_rows, sticky="w",
                padx=(0, 4), pady=(8, 4),
            )

            any_will_write = False
            for i in range(n_rows):
                row_no = start_row + i
                t = tests[i] if i < len(tests) else {}
                target_row = age_rows[i] if i < len(age_rows) else None

                weight = t.get("weight_gr")
                load = t.get("load_kn")

                existing_w = None
                existing_l = None
                if target_row is not None:
                    off = target_row - block_start
                    if 0 <= off < len(weights_existing):
                        existing_w = weights_existing[off]
                    if 0 <= off < len(loads_existing):
                        existing_l = loads_existing[off]

                w_empty = writer.is_cell_empty(existing_w)
                l_empty = writer.is_cell_empty(existing_l)

                w_var = StringVar(
                    value="" if weight is None else str(weight)
                )
                l_var = StringVar(
                    value="" if load is None else str(load)
                )

                w_kwargs = dict(
                    textvariable=w_var, width=180, height=38,
                    justify="center", font=self._entry_font,
                )
                l_kwargs = dict(
                    textvariable=l_var, width=180, height=38,
                    justify="center", font=self._entry_font,
                )
                if not w_empty:
                    w_kwargs["text_color"] = FILLED_TEXT
                if not l_empty:
                    l_kwargs["text_color"] = FILLED_TEXT

                if weight is not None and w_empty:
                    any_will_write = True
                if load is not None and l_empty:
                    any_will_write = True

                ctk.CTkEntry(grid, **w_kwargs).grid(
                    row=row_no, column=1, padx=8, pady=4,
                )
                ctk.CTkEntry(grid, **l_kwargs).grid(
                    row=row_no, column=2, padx=8, pady=4,
                )

            group_var = BooleanVar(value=any_will_write and not bad)
            group_chk = ctk.CTkCheckBox(
                grid, text="", variable=group_var, width=24,
            )
            group_chk.grid(
                row=start_row, column=3, rowspan=n_rows, sticky="",
                padx=(12, 0), pady=3,
            )
            if bad:
                group_chk.configure(state="disabled")
            return group_var, n_rows

        n7_rows_max = max(len(cube.get("tests_7d", [])), len(rows_7d), 1)
        c7, _ = _build_group(
            1, "7-day", AGE_7_COLOR,
            cube.get("tests_7d", []), rows_7d,
        )
        c28, _ = _build_group(
            1 + n7_rows_max, "28-day", AGE_28_COLOR,
            cube.get("tests_28d", []), rows_28d,
        )
        entry["check_7d"] = c7
        entry["check_28d"] = c28

        if not bad and not c7.get() and not c28.get():
            entry["enabled"].set(False)

    def _on_card_down(self, _event=None):
        if not self._cards:
            return "break"
        if self._selected_idx is None:
            self._select_card(0)
        else:
            self._select_card(
                min(self._selected_idx + 1, len(self._cards) - 1)
            )
        return "break"

    def _on_card_up(self, _event=None):
        if not self._cards:
            return "break"
        if self._selected_idx is None:
            self._select_card(0)
        else:
            self._select_card(max(self._selected_idx - 1, 0))
        return "break"

    def _select_card(self, idx: int):
        if idx < 0 or idx >= len(self._cards):
            return
        if self._selected_idx is not None and self._selected_idx != idx:
            try:
                self._cards[self._selected_idx].configure(
                    border_color=("gray70", "gray30"), border_width=1,
                )
            except Exception:
                pass
        self._selected_idx = idx
        try:
            self._cards[idx].configure(
                border_color=SELECTED_BORDER, border_width=3,
            )
        except Exception:
            pass
        self._scroll_to_card_centered(idx)

    def _scroll_to_card_centered(self, idx: int):
        canvas = getattr(self._cube_list, "_parent_canvas", None)
        if canvas is None:
            return
        card = self._cards[idx]
        self.win.update_idletasks()
        bbox = canvas.bbox("all")
        if not bbox:
            return
        total_h = bbox[3] - bbox[1]
        if total_h <= 0:
            return
        card_y = card.winfo_y()
        card_h = card.winfo_height()
        view_h = canvas.winfo_height()
        target = card_y + card_h / 2 - view_h / 2
        target = max(0, min(target, total_h - view_h))
        canvas.yview_moveto(target / total_h)

    def _do_write(self):
        if self.ledger_error:
            return

        written_cubes = 0
        total_cells = 0
        total_skipped = 0
        all_errors: list[str] = []
        unchecked: list[str] = []

        for entry in self.entries:
            cube = entry["cube"]
            mark = cube.get("sample_mark", "?")
            if entry["mismatch"] is not None:
                continue
            if not entry["enabled"].get():
                unchecked.append(mark)
                continue
            write_7d = entry.get("check_7d")
            write_28d = entry.get("check_28d")
            do_7d = write_7d.get() if write_7d is not None else True
            do_28d = write_28d.get() if write_28d is not None else True
            if not do_7d and not do_28d:
                unchecked.append(mark)
                continue
            try:
                result = writer.write_ledger_cube(
                    self._ws, cube, entry["block"],
                    write_7d=do_7d, write_28d=do_28d,
                )
                total_cells += len(result["wrote"])
                total_skipped += len(result["skipped"])
                all_errors.extend(
                    f"{mark}: {e}" for e in result["errors"]
                )
                if result["wrote"]:
                    written_cubes += 1
            except Exception as e:
                all_errors.append(f"{mark}: write error: {e}")

        lines = [
            f"{written_cubes} cubes written, {total_cells} cells.",
            f"{total_skipped} cells skipped (already filled).",
        ]
        if self.not_found:
            lines.append(f"{len(self.not_found)} cubes not found (not in ledger).")
        if unchecked:
            lines.append(f"{len(unchecked)} cubes manually disabled.")
        if all_errors:
            lines.append("")
            lines.append("Errors:")
            for e in all_errors:
                lines.append("  - " + e)

        messagebox.showinfo("Ledger — Result", "\n".join(lines), parent=self.win)

        if not all_errors:
            self._close()


# ---------- Main Window ----------

class MainWindow:
    # Status colors (map legacy names to CTk-friendly hexes)
    STATUS_COLORS = {
        "blue": "#1565C0",
        "darkgreen": "#2E7D32",
        "red": "#C62828",
        "gray": "gray55",
    }

    SUPPORTED_EXTS = (".pdf", ".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff")

    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        # Center the main window on the primary screen.
        ww, wh = 620, 680
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = max(0, (sw - ww) // 2)
        y = max(0, (sh - wh) // 2)
        self.root.geometry(f"{ww}x{wh}+{x}+{y}")
        self.root.minsize(560, 620)

        self.result_queue: queue.Queue = queue.Queue()
        self.selected_image: Image.Image | None = None

        # Batch state
        self._batch: list[str] = []   # remaining file paths to process
        self._batch_total: int = 0    # size of current batch
        self._busy: bool = False
        # Accumulates (image, cubes_data) per processed file; shown as
        # a single merged preview once the whole batch is read.
        self._pending_results: list[tuple[Image.Image, dict]] = []

        # Ledger state — cubes_data from the most recent successful
        # first-Excel write. Enables the "Go to Books Excel section" button.
        self._last_cubes_data: dict | None = None

        self._build_ui()
        self._register_dnd()

        # On first run with no key, auto-open the settings dialog
        if not get_current_api_key():
            self.root.after(200, self._first_run_prompt)

    def _register_dnd(self):
        if not _DND_AVAILABLE:
            return
        try:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind("<<Drop>>", self._on_drop)
        except Exception:
            pass

    def _on_drop(self, event):
        raw = getattr(event, "data", "") or ""
        paths = _parse_dropped_paths(raw)
        files = [p for p in paths
                 if Path(p).suffix.lower() in self.SUPPORTED_EXTS]
        if not files:
            self._set_status("Drop ignored: no supported files.", "red")
            return
        self._start_batch(files)

    def _build_ui(self):
        _add_credit_footer(self.root)
        # ---- HEADER ----
        header = ctk.CTkFrame(self.root, corner_radius=0, height=70)
        header.pack(side="top", fill="x")
        header.pack_propagate(False)

        title_wrap = ctk.CTkFrame(header, fg_color="transparent")
        title_wrap.pack(side="left", padx=22, pady=10)
        ctk.CTkLabel(
            title_wrap,
            text="Notebook -> Excel",
            font=ctk.CTkFont(size=18, weight="bold"),
            anchor="w",
        ).pack(anchor="w")
        ctk.CTkLabel(
            title_wrap,
            text="Read cube tests from a notebook page and fill them into Excel.",
            font=ctk.CTkFont(size=11),
            text_color="gray55",
            anchor="w",
        ).pack(anchor="w")

        ctk.CTkButton(
            header,
            text="Settings",
            width=90, height=32,
            fg_color="gray70", hover_color="gray60",
            command=self._open_settings,
        ).pack(side="right", padx=16, pady=18)

        # ---- BODY ----
        body = ctk.CTkFrame(self.root, fg_color="transparent")
        body.pack(side="top", fill="both", expand=True, padx=16, pady=12)

        # Steps card
        steps_card = ctk.CTkFrame(body, corner_radius=12, border_width=1)
        steps_card.pack(fill="x", pady=(0, 12))
        ctk.CTkLabel(
            steps_card,
            text="How it works",
            font=ctk.CTkFont(size=13, weight="bold"),
            anchor="w",
        ).pack(anchor="w", padx=16, pady=(12, 4))

        steps = [
            ("1.", "Open Excel and click the sheet of the FIRST cube in the notebook."),
            ("2.", "Pick the notebook page below (PDF / JPG / PNG)."),
            ("3.", "Gemini reads the page and matches cubes to sheets forward from the active one."),
            ("4.", "Verify the values in the review window, then click 'Write to Excel'."),
        ]
        for num, text in steps:
            row = ctk.CTkFrame(steps_card, fg_color="transparent")
            row.pack(fill="x", padx=16, pady=1)
            ctk.CTkLabel(
                row, text=num, width=22,
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color="#2E7D32", anchor="w",
            ).pack(side="left")
            ctk.CTkLabel(
                row, text=text, font=ctk.CTkFont(size=12),
                justify="left", anchor="w",
            ).pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(steps_card, text="").pack(pady=(0, 6))  # bottom spacer

        # Scan mode card
        scan_card = ctk.CTkFrame(body, corner_radius=12, border_width=1)
        scan_card.pack(fill="x", pady=(0, 12))
        ctk.CTkLabel(
            scan_card,
            text="Scan mode",
            font=ctk.CTkFont(size=13, weight="bold"),
            anchor="w",
        ).pack(anchor="w", padx=16, pady=(12, 6))

        self.scan_mode_var = StringVar(value=DEFAULT_SCAN_MODE)
        for key, (label, _cap) in SCAN_MODES.items():
            ctk.CTkRadioButton(
                scan_card,
                text=label,
                variable=self.scan_mode_var,
                value=key,
                font=ctk.CTkFont(size=12),
            ).pack(anchor="w", padx=18, pady=3)
        ctk.CTkLabel(scan_card, text="").pack(pady=(0, 4))

        # Primary action
        self.pick_btn = ctk.CTkButton(
            body,
            text=">> PICK NOTEBOOK PAGE(S)",
            height=46,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#2E7D32", hover_color="#1B5E20",
            command=self._on_pick,
        )
        self.pick_btn.pack(fill="x", pady=(4, 0))

        hint_text = (
            "Tip: drag & drop PDFs/images onto this window — batch supported."
            if _DND_AVAILABLE else
            "Tip: select multiple files to process them in sequence."
        )
        ctk.CTkLabel(
            body, text=hint_text,
            font=ctk.CTkFont(size=11),
            text_color="gray55",
        ).pack(pady=(6, 0))

        # Secondary action: open the ledger preview.
        # Disabled until the user finishes a first-Excel write.
        self.ledger_btn = ctk.CTkButton(
            body,
            text="Go to Books Excel section",
            height=40,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#1565C0", hover_color="#0D47A1",
            command=self._on_open_ledger,
            state="disabled",
        )
        self.ledger_btn.pack(fill="x", pady=(12, 0))

        # ---- FOOTER (status) ----
        footer = ctk.CTkFrame(self.root, corner_radius=0, height=38)
        footer.pack(side="bottom", fill="x")
        footer.pack_propagate(False)
        self.status = ctk.CTkLabel(
            footer, text="Ready.",
            text_color=self.STATUS_COLORS["blue"],
            font=ctk.CTkFont(size=12),
        )
        self.status.pack(side="left", padx=18)

    def _set_status(self, text: str, color: str = "blue"):
        self.status.configure(
            text=text,
            text_color=self.STATUS_COLORS.get(color, color),
        )
        self.root.update_idletasks()

    def _set_pick_enabled(self, enabled: bool):
        self.pick_btn.configure(state="normal" if enabled else "disabled")

    def _set_ledger_enabled(self, enabled: bool):
        if not hasattr(self, "ledger_btn"):
            return
        self.ledger_btn.configure(state="normal" if enabled else "disabled")

    def _on_open_ledger(self):
        if self._last_cubes_data is None:
            messagebox.showwarning(
                "Ledger",
                "Process a PDF and write the first Excel first, then try again.",
                parent=self.root,
            )
            return
        try:
            LedgerPreviewWindow(self.root, self._last_cubes_data)
        except Exception as e:
            tb = traceback.format_exc()
            reader._log(f"--- LEDGER ERROR ---\n{tb}")
            short = str(e)
            if len(short) > 500:
                short = short[:500] + "..."
            messagebox.showerror(
                "Ledger Error",
                f"{short}\n\n(Full traceback saved to gemini_debug.log)",
                parent=self.root,
            )

    def _open_settings(self):
        SettingsDialog(self.root)

    def _first_run_prompt(self):
        messagebox.showinfo(
            "API Key Required",
            "You need a Google Gemini API key to use this program.\n\n"
            "Get one for free: https://aistudio.google.com/apikey\n\n"
            "The settings dialog is opening now — paste your key and "
            "click 'Save'.",
            parent=self.root,
        )
        self._open_settings()

    def _on_pick(self):
        if not get_current_api_key():
            messagebox.showwarning(
                "No key",
                "Please set your Gemini API key first (Settings button).",
                parent=self.root,
            )
            return

        paths = filedialog.askopenfilenames(
            parent=self.root,
            title="Pick notebook page(s)",
            filetypes=[
                ("Supported files", "*.pdf *.jpg *.jpeg *.png *.bmp *.tif *.tiff"),
                ("PDF", "*.pdf"),
                ("Image", "*.jpg *.jpeg *.png *.bmp *.tif *.tiff"),
            ],
        )
        reader._log(f"[PICK] raw paths type={type(paths).__name__} "
                    f"repr={paths!r}")
        if not paths:
            return
        if isinstance(paths, str):
            paths = self.root.tk.splitlist(paths)
        paths = list(paths)
        reader._log(f"[PICK] final count={len(paths)} "
                    f"items={[str(p) for p in paths]}")
        self._start_batch(paths)

    def _start_batch(self, paths: list[str]):
        reader._log(f"[BATCH] _start_batch called: "
                    f"incoming={len(paths)} busy={self._busy} "
                    f"current_batch={len(self._batch)}")
        if self._busy:
            # Append to the current queue
            self._batch.extend(paths)
            self._batch_total += len(paths)
            self._set_status(
                f"Queued {len(paths)} more — {len(self._batch)} pending.",
                "blue",
            )
            reader._log(f"[BATCH] appended to busy queue, "
                        f"now batch={len(self._batch)} total={self._batch_total}")
            return
        if not get_current_api_key():
            messagebox.showwarning(
                "No key",
                "Please set your Gemini API key first (Settings button).",
                parent=self.root,
            )
            return
        self._batch = list(paths)
        self._batch_total = len(paths)
        reader._log(f"[BATCH] starting fresh batch: "
                    f"batch={len(self._batch)} total={self._batch_total}")
        self._process_next()

    def _process_next(self):
        reader._log(f"[BATCH] _process_next: remaining={len(self._batch)} "
                    f"total={self._batch_total} "
                    f"pending={len(self._pending_results)}")
        if not self._batch:
            if self._pending_results:
                # All files read → show single merged preview.
                self._show_merged_preview()
                return
            self._busy = False
            self._batch_total = 0
            self._set_pick_enabled(True)
            self._set_status("Ready.", "blue")
            return

        path = self._batch.pop(0)
        index = self._batch_total - len(self._batch)
        reader._log(f"[BATCH] processing {index}/{self._batch_total}: "
                    f"{Path(path).name}")
        self._busy = True
        self._set_pick_enabled(False)

        prefix = (f"[{index}/{self._batch_total}] "
                  if self._batch_total > 1 else "")
        self._set_status(
            f"{prefix}Gemini is reading: {Path(path).name} ...", "blue"
        )

        mode_key = self.scan_mode_var.get()
        max_forward = SCAN_MODES.get(mode_key, SCAN_MODES[DEFAULT_SCAN_MODE])[1]

        t = threading.Thread(
            target=self._worker, args=(path, max_forward, prefix), daemon=True
        )
        t.start()
        self.root.after(150, self._poll)

    def _worker(self, file_path: str, max_forward: int, prefix: str):
        """
        1. Gemini reads the notebook page
        2. Scan forward from the active Excel sheet based on notebook IDs
        """
        pythoncom = None
        try:
            import pythoncom as _pyc
            pythoncom = _pyc
            pythoncom.CoInitialize()
        except Exception:
            pass

        def _progress(stage, cur=0, tot=0):
            if stage == "cache_hit":
                msg = f"{prefix}Cache hit — no API call."
            elif stage == "loading":
                msg = f"{prefix}Loading pages..."
            elif stage == "page":
                msg = f"{prefix}Gemini reading page {cur}/{tot}..."
            else:
                msg = f"{prefix}{stage}"
            self.result_queue.put(("progress", msg))

        try:
            reader._log(f"[WORKER] START: {Path(file_path).name}")
            images = reader.load_images(file_path)
            reader._log(f"[WORKER] images loaded: {len(images)}")
            image = reader.stack_images_vertically(images)
            reader._log("[WORKER] stacked; calling read_notebook")
            data = reader.read_notebook(file_path, progress_cb=_progress)
            reader._log(f"[WORKER] read_notebook done: "
                        f"{len(data.get('cubes', []))} cubes")
            # Sheet scan runs once on the merged result in
            # _show_merged_preview — not per file.
            self.result_queue.put(("ok", image, data, None, prefix))
            reader._log("[WORKER] queued ok")
        except Exception as e:
            tb = traceback.format_exc()
            reader._log(f"[WORKER] EXCEPTION: {e}\n{tb}")
            self.result_queue.put(("err", e, tb))
        finally:
            if pythoncom is not None:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    def _poll(self):
        drained_final = False
        try:
            while True:
                item = self.result_queue.get_nowait()
                kind = item[0]
                if kind == "progress":
                    self._set_status(item[1], "blue")
                else:
                    # final event — handle and stop polling
                    self._handle_final(item)
                    drained_final = True
                    break
        except queue.Empty:
            pass

        if not drained_final:
            self.root.after(150, self._poll)

    def _show_merged_preview(self):
        """Merge all stashed (image, cubes_data) pairs and show one preview."""
        if not self._pending_results:
            return
        images = [img for (img, _) in self._pending_results]
        merged_image = (images[0] if len(images) == 1
                        else reader.stack_images_vertically(images))
        all_cubes: list[dict] = []
        for _, data in self._pending_results:
            all_cubes.extend(data.get("cubes", []))
        merged_data = {"cubes": all_cubes}

        mode_key = self.scan_mode_var.get()
        max_forward = SCAN_MODES.get(
            mode_key, SCAN_MODES[DEFAULT_SCAN_MODE]
        )[1]
        try:
            scan_result = writer.scan_sheets_for_cubes(
                merged_data, max_forward=max_forward
            )
        except Exception as e:
            reader._log(f"[BATCH] merged scan FAILED: {e}")
            scan_result = {"sheets": [], "scanned_count": 0}

        n = len(all_cubes)
        scanned = scan_result.get("scanned_count", 0)
        reader._log(f"[BATCH] merged preview: {n} cubes across "
                    f"{len(self._pending_results)} files, "
                    f"{scanned} sheets checked")
        self._set_status(
            f"Done: {n} cubes from {len(self._pending_results)} files, "
            f"{scanned} sheets checked.",
            "darkgreen",
        )
        self._last_cubes_data = merged_data
        self._set_ledger_enabled(True)
        self._pending_results = []

        try:
            PreviewWindow(
                self.root, merged_image, merged_data, scan_result,
                on_close=self._on_preview_closed,
            )
            reader._log("[BATCH] merged PreviewWindow created")
        except Exception as e:
            reader._log(f"[BATCH] merged PreviewWindow FAILED: {e}")
            import traceback as _tb
            reader._log(_tb.format_exc())
            self._on_preview_closed()

    def _handle_final(self, item):
        if item[0] == "ok":
            _, image, data, scan_result, prefix = item
            n = len(data.get("cubes", []))
            self._pending_results.append((image, data))
            reader._log(f"[BATCH] _handle_final OK: {n} cubes stashed "
                        f"(files done={len(self._pending_results)}, "
                        f"remaining={len(self._batch)})")
            self._set_status(
                f"{prefix}Read {n} cubes. "
                f"({len(self._pending_results)}/{self._batch_total})",
                "blue",
            )
            # Advance — _process_next decides whether to read next file
            # or show the merged preview.
            self.root.after(200, self._process_next)
            return
        else:
            _, err, tb = item
            self._set_status(f"Error: {err}", "red")
            reader._log(f"--- ERROR ---\n{tb}")
            short = str(err)
            if len(short) > 500:
                short = short[:500] + "..."
            messagebox.showerror(
                "Error",
                f"{short}\n\n(Full traceback saved to gemini_debug.log)",
                parent=self.root,
            )
            self._on_preview_closed()

    def _on_preview_closed(self):
        reader._log(f"[BATCH] _on_preview_closed: "
                    f"remaining={len(self._batch)}")
        # Small delay so Excel/COM releases before the next file starts.
        self.root.after(200, self._process_next)


def _start_update_check_thread(root):
    """Background thread: poll GitHub; if new version, prompt user on UI thread."""
    def worker():
        info = updater.check_for_update(timeout=5)
        if info is None:
            return
        root.after(0, lambda: _prompt_user_for_update(root, info))

    t = threading.Thread(target=worker, daemon=True)
    t.start()


def _prompt_user_for_update(root, info):
    msg = (
        f"Yeni surum mevcut: {info.version}\n\n{info.notes}\n\n"
        "Simdi guncellensin mi?"
    )
    if messagebox.askyesno("Guncelleme mevcut", msg, parent=root):
        updater.run_update_flow(info, parent_window=root)


def main():
    root = DnDCTk() if _DND_AVAILABLE else ctk.CTk()
    MainWindow(root)
    _start_update_check_thread(root)
    root.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        # Last-resort messagebox for errors that happen before the GUI is up
        try:
            from tkinter import Tk, messagebox
            root = Tk()
            root.withdraw()
            messagebox.showerror(
                "Critical Error",
                f"The program could not start:\n\n{type(e).__name__}: {e}\n\n"
                f"{traceback.format_exc()}",
            )
        except Exception:
            pass
        raise
