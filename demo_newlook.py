"""
demo_newlook.py — Mockup of the proposed CustomTkinter redesign.

Run:
    pip install customtkinter
    python demo_newlook.py

This is only a preview — no Gemini, no Excel. Just the look & feel.
"""
import customtkinter as ctk


# ---- Theme ----
# "light" / "dark" / "system"
ctk.set_appearance_mode("light")
# "blue" / "green" / "dark-blue"
ctk.set_default_color_theme("blue")


# ---- Fake data (stand-in for what Gemini would return) ----
SAMPLE_CUBES = [
    {
        "cube_no": "378",
        "sample_mark": "G26-CON-395",
        "sheet": "Sheet 395 — G26-CON-395",
        "matched": True,
        # rows: (weight, load, excel_weight_empty, excel_load_empty)
        "rows_7d": [
            (8360, 1102.34, False, False),
            (8332, 1196.46, False, False),
            (8315, 1202.46, False, False),
        ],
        "rows_28d": [
            (8318, 1233.74, True, True),
            (8383, 1305.96, True, True),
            (8371, 1258.69, True, True),
        ],
    },
    {
        "cube_no": "379",
        "sample_mark": "G26-CON-396",
        "sheet": "Sheet 396 — G26-CON-396",
        "matched": True,
        "rows_7d": [
            (8220, 1084.12, False, False),
            (8245, 1090.55, False, False),
            (8210, 1075.20, False, False),
        ],
        "rows_28d": [
            (8240, 1210.10, True, True),
            (8260, 1245.33, False, True),  # weight already in Excel
            (8235, 1198.80, True, False),  # load already in Excel
        ],
    },
    {
        "cube_no": "380",
        "sample_mark": "G26-CON-397",
        "sheet": None,
        "matched": False,
        "rows_7d": [
            (None, None, True, True),
            (None, None, True, True),
            (None, None, True, True),
        ],
        "rows_28d": [
            (8190, 1188.50, True, True),
            (8205, 1175.15, True, True),
            (8175, 1162.40, True, True),
        ],
    },
]


SELECTED_BORDER = "#1976D2"


class DemoPreview(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Preview (new design mockup) — CubeLogReader")
        self.geometry("1500x900")
        self.minsize(1000, 650)

        self._img_visible = True
        self._cards: list = []
        self._selected_idx = None
        self._scroll_anim_id = None
        self._build()

        self.bind("<Down>", self._on_card_down)
        self.bind("<Up>", self._on_card_up)

    def _build(self):
        # ---- TOP BAR ----
        top = ctk.CTkFrame(self, corner_radius=0, height=60)
        top.pack(side="top", fill="x")

        ctk.CTkLabel(
            top,
            text="Notebook -> Excel  ·  3 cubes found, 2 matched",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(side="left", padx=20, pady=12)

        legend = ctk.CTkFrame(top, fg_color="transparent")
        legend.pack(side="left", padx=20)
        ctk.CTkLabel(
            legend,
            text="Green border = empty in Excel (needs fill)",
            text_color="#00C853",
            font=ctk.CTkFont(size=11, weight="bold"),
        ).pack(side="left", padx=(0, 12))
        ctk.CTkLabel(
            legend,
            text="Gray text = already in Excel (skip)",
            text_color="gray55",
            font=ctk.CTkFont(size=11),
        ).pack(side="left")

        ctk.CTkButton(
            top,
            text="Cancel",
            width=90,
            fg_color="gray70",
            hover_color="gray60",
            command=self.destroy,
        ).pack(side="right", padx=(0, 10), pady=12)

        self._toggle_btn = ctk.CTkButton(
            top,
            text="Hide image",
            width=110,
            command=self._toggle_image,
        )
        self._toggle_btn.pack(side="right", padx=(0, 10), pady=12)

        ctk.CTkButton(
            top,
            text=">> WRITE TO EXCEL",
            width=200,
            height=36,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#2E7D32",
            hover_color="#1B5E20",
        ).pack(side="right", padx=(10, 15), pady=12)

        # ---- BODY (image left / cube list right) ----
        body = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        body.pack(side="top", fill="both", expand=True, padx=10, pady=10)

        # Left: "image" placeholder
        self._img_panel = ctk.CTkFrame(body, width=700, corner_radius=10)
        self._img_panel.pack(side="left", fill="y", padx=(0, 10))
        self._img_panel.pack_propagate(False)

        img_toolbar = ctk.CTkFrame(self._img_panel, corner_radius=0, height=40)
        img_toolbar.pack(side="top", fill="x")
        for txt in ["-", "100%", "+", "Fit", "Actual size"]:
            ctk.CTkButton(
                img_toolbar, text=txt, width=55, height=28
            ).pack(side="left", padx=4, pady=6)
        ctk.CTkLabel(
            img_toolbar,
            text="Ctrl+wheel: zoom | Ctrl++/- | Ctrl+0: fit",
            text_color="gray50",
            font=ctk.CTkFont(size=10),
        ).pack(side="left", padx=10)

        placeholder = ctk.CTkFrame(self._img_panel, fg_color="gray85")
        placeholder.pack(fill="both", expand=True, padx=10, pady=10)
        ctk.CTkLabel(
            placeholder,
            text="[ PDF page preview\n   would appear here ]",
            text_color="gray40",
            font=ctk.CTkFont(size=16),
        ).pack(expand=True)

        # Right: scrollable cube list
        self._list = ctk.CTkScrollableFrame(body, corner_radius=10)
        self._list.pack(side="right", fill="both", expand=True)

        # Faster wheel scrolling
        try:
            self._list._parent_canvas.configure(yscrollincrement=20)
        except Exception:
            pass

        for cube in SAMPLE_CUBES:
            self._build_card(self._list, cube)

    def _build_card(self, parent, cube):
        card = ctk.CTkFrame(parent, corner_radius=12, border_width=1)
        card.pack(fill="x", padx=8, pady=8)
        idx = len(self._cards)
        self._cards.append(card)
        card.bind("<Button-1>", lambda _e, i=idx: self._select_card(i))

        # ---- Header row (3-column grid so title stays truly centered) ----
        header = ctk.CTkFrame(card, fg_color="transparent")
        header.pack(fill="x", padx=15, pady=(12, 6))
        header.grid_columnconfigure(0, weight=1, uniform="h")
        header.grid_columnconfigure(1, weight=0)
        header.grid_columnconfigure(2, weight=1, uniform="h")

        # Auto-check logic: cube master only on if any group has real work
        def _group_has_work(rows):
            any_empty = any(w_empty or l_empty for _, _, w_empty, l_empty in rows)
            any_val = any(w is not None or l is not None for w, l, _, _ in rows)
            return any_empty and any_val

        master_var = ctk.BooleanVar(
            value=_group_has_work(cube["rows_7d"]) or _group_has_work(cube["rows_28d"])
        )
        ctk.CTkCheckBox(
            header,
            text="Write this sample",
            variable=master_var,
            font=ctk.CTkFont(size=12),
        ).grid(row=0, column=0, sticky="w")

        title = f"Cube No {cube['cube_no']}  ·  {cube['sample_mark']}"
        ctk.CTkLabel(
            header, text=title, font=ctk.CTkFont(size=15, weight="bold")
        ).grid(row=0, column=1, sticky="")  # centered in its cell

        # Status pill on the right
        if cube["matched"]:
            pill_text = f"[OK]  {cube['sheet']}"
            pill_color = "#2E7D32"
        else:
            pill_text = "[NO MATCH]"
            pill_color = "#C62828"

        ctk.CTkLabel(
            header,
            text=pill_text,
            fg_color=pill_color,
            text_color="white",
            corner_radius=14,
            font=ctk.CTkFont(size=11, weight="bold"),
            padx=10,
            pady=3,
        ).grid(row=0, column=2, sticky="e")

        # ---- Value grid ----
        grid_wrap = ctk.CTkFrame(card, fg_color="transparent")
        grid_wrap.pack(padx=15, pady=(2, 12))

        # Column headers
        hdr_font = ctk.CTkFont(size=11, weight="bold")
        ctk.CTkLabel(grid_wrap, text="", width=90).grid(row=0, column=0)
        ctk.CTkLabel(
            grid_wrap, text="Weight (gr)", width=180, font=hdr_font
        ).grid(row=0, column=1, padx=8)
        ctk.CTkLabel(
            grid_wrap, text="Load (kN)", width=180, font=hdr_font
        ).grid(row=0, column=2, padx=8)
        ctk.CTkLabel(grid_wrap, text="", width=70).grid(row=0, column=3)

        def _group(start_row, age_label, age_color, rows):
            # Age label pinned to the left of the group, spanning all 3 rows
            ctk.CTkLabel(
                grid_wrap,
                text=age_label,
                text_color=age_color,
                font=ctk.CTkFont(size=13, weight="bold"),
                width=90,
                anchor="w",
            ).grid(
                row=start_row, column=0, rowspan=3, sticky="w",
                padx=(0, 4), pady=(8, 4),
            )

            # Empty in Excel = bright green border (needs filling)
            # Already in Excel = dim gray text (don't touch)
            EMPTY_COLOR = "#00C853"
            FILLED_TEXT = "gray55"

            for i, (weight, load, w_empty, l_empty) in enumerate(rows):
                row = start_row + i
                w_var = ctk.StringVar(
                    value="" if weight is None else str(weight)
                )
                l_var = ctk.StringVar(
                    value="" if load is None else str(load)
                )

                w_kwargs = dict(
                    textvariable=w_var, width=180, height=38,
                    justify="center", font=ctk.CTkFont(size=16),
                )
                if w_empty and weight is not None:
                    # Will be written
                    w_kwargs["border_color"] = EMPTY_COLOR
                    w_kwargs["border_width"] = 3
                    w_kwargs["fg_color"] = "#E8F5E9"
                elif not w_empty:
                    # Already in Excel
                    w_kwargs["text_color"] = FILLED_TEXT
                # else: empty in Excel + no notebook value → no highlight

                l_kwargs = dict(
                    textvariable=l_var, width=180, height=38,
                    justify="center", font=ctk.CTkFont(size=16),
                )
                if l_empty and load is not None:
                    l_kwargs["border_color"] = EMPTY_COLOR
                    l_kwargs["border_width"] = 3
                    l_kwargs["fg_color"] = "#E8F5E9"
                elif not l_empty:
                    l_kwargs["text_color"] = FILLED_TEXT

                ctk.CTkEntry(grid_wrap, **w_kwargs).grid(
                    row=row, column=1, padx=8, pady=4
                )
                ctk.CTkEntry(grid_wrap, **l_kwargs).grid(
                    row=row, column=2, padx=8, pady=4
                )

            # Group checkbox — auto-off if no work (all filled or no values)
            any_empty = any(w_e or l_e for _, _, w_e, l_e in rows)
            any_val = any(w is not None or l is not None for w, l, _, _ in rows)
            group_var = ctk.BooleanVar(value=any_empty and any_val)
            ctk.CTkCheckBox(
                grid_wrap,
                text="",
                variable=group_var,
                width=24,
            ).grid(
                row=start_row, column=3, rowspan=3, sticky="",  # centered
                padx=(12, 0), pady=3,
            )

        _group(1, "7-day", "#2E7D32", cube["rows_7d"])
        _group(4, "28-day", "#C62828", cube["rows_28d"])

    # ---- Card selection + smooth scroll ----

    def _on_card_down(self, _e=None):
        if not self._cards:
            return "break"
        nxt = 0 if self._selected_idx is None else min(
            self._selected_idx + 1, len(self._cards) - 1
        )
        self._select_card(nxt)
        return "break"

    def _on_card_up(self, _e=None):
        if not self._cards:
            return "break"
        prv = 0 if self._selected_idx is None else max(
            self._selected_idx - 1, 0
        )
        self._select_card(prv)
        return "break"

    def _select_card(self, idx):
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

    def _scroll_to_card_centered(self, idx):
        canvas = getattr(self._list, "_parent_canvas", None)
        if canvas is None:
            return
        card = self._cards[idx]
        self.update_idletasks()
        bbox = canvas.bbox("all")
        if not bbox:
            return
        total_h = bbox[3] - bbox[1]
        viewport_h = canvas.winfo_height()
        if total_h <= viewport_h:
            return
        card_y = card.winfo_y()
        card_h = card.winfo_height()
        center = card_y + card_h / 2
        target_top = max(0, min(center - viewport_h / 2, total_h - viewport_h))
        self._animate_scroll(canvas, target_top / total_h, 220)

    def _animate_scroll(self, canvas, target_frac, duration_ms=220):
        if self._scroll_anim_id is not None:
            try:
                self.after_cancel(self._scroll_anim_id)
            except Exception:
                pass
            self._scroll_anim_id = None
        frame_ms = 16
        frames = max(1, duration_ms // frame_ms)
        start_frac = canvas.yview()[0]

        def step(i):
            t = i / frames
            ease = 1 - (1 - t) ** 3
            pos = start_frac + (target_frac - start_frac) * ease
            canvas.yview_moveto(pos)
            if i < frames:
                self._scroll_anim_id = self.after(
                    frame_ms, lambda: step(i + 1)
                )
            else:
                self._scroll_anim_id = None

        step(1)

    def _toggle_image(self):
        if self._img_visible:
            self._img_panel.pack_forget()
            self._toggle_btn.configure(text="Show image")
            self._img_visible = False
        else:
            # CTkScrollableFrame doesn't expose its real packed widget, so
            # "before=self._list" fails. Re-pack both in the correct order.
            self._list.pack_forget()
            self._img_panel.pack(side="left", fill="y", padx=(0, 10))
            self._list.pack(side="right", fill="both", expand=True)
            self._toggle_btn.configure(text="Hide image")
            self._img_visible = True


if __name__ == "__main__":
    app = DemoPreview()
    app.mainloop()
