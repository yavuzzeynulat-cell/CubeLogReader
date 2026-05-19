"""
demo_shotcrete.py — Mockup of the shotcrete preview card.

5 rows per age. Top 3 by strength auto-selected (green border + check on).
Bottom 2 dimmed (no check). Clicking a currently-off row auto-deselects
the currently-on row with the lowest strength in that age group.

Columns (matches writer cell layout):
    [chk]  Diameter (R)  Height (V)  Weight (W)  Load (AA)  Strength

Run:
    python demo_shotcrete.py
"""
import customtkinter as ctk

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

EMPTY_BORDER = "#00C853"
EMPTY_BG = "#E8F5E9"
DIM_TEXT = "gray55"
DIM_BG = "gray90"
SHOT_PILL = "#C62828"
OK_PILL = "#2E7D32"

SHOTCRETE_CUBE = {
    "cube_no": "501",
    "sample_mark": "G-CON-7000",
    "sheet": "2630 — G-CON-7000",
    "rows_7d": [
        # diameter, height, weight, load, strength
        (94, 188, 2110, 275.20, 39.60),
        (94, 188, 2098, 268.40, 38.52),
        (93, 187, 2125, 281.10, 40.89),
        (94, 189, 2085, 255.70, 36.77),
        (94, 188, 2105, 272.80, 39.12),
    ],
    "rows_28d": [
        (94, 188, 2120, 345.50, 49.75),
        (93, 187, 2115, 352.80, 51.24),
        (94, 188, 2108, 338.20, 48.60),
        (94, 189, 2125, 360.10, 52.03),
        (94, 188, 2099, 342.00, 49.10),
    ],
}


class ShotcreteDemo(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Shotcrete card preview — CubeLogReader")
        self.geometry("1250x780")
        self.minsize(1050, 680)

        top = ctk.CTkFrame(self, corner_radius=0, height=56)
        top.pack(side="top", fill="x")
        ctk.CTkLabel(
            top, text="Shotcrete cube  —  5 rows per age, top 3 by strength",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(side="left", padx=20, pady=12)
        ctk.CTkLabel(
            top,
            text="Green border = top 3, will be written   ·   Dim = unused specimen",
            text_color="gray40", font=ctk.CTkFont(size=11),
        ).pack(side="left", padx=20)
        ctk.CTkButton(
            top, text="Close", width=80, fg_color="gray70", hover_color="gray60",
            command=self.destroy,
        ).pack(side="right", padx=15, pady=10)

        wrap = ctk.CTkScrollableFrame(self, corner_radius=10)
        wrap.pack(fill="both", expand=True, padx=10, pady=10)
        try:
            wrap._parent_canvas.configure(yscrollincrement=20)
        except Exception:
            pass

        self._build_card(wrap, SHOTCRETE_CUBE)

    def _build_card(self, parent, cube):
        card = ctk.CTkFrame(parent, corner_radius=12, border_width=1)
        card.pack(fill="x", padx=8, pady=8)

        header = ctk.CTkFrame(card, fg_color="transparent")
        header.pack(fill="x", padx=15, pady=(12, 6))
        header.grid_columnconfigure(0, weight=1, uniform="h")
        header.grid_columnconfigure(1, weight=0)
        header.grid_columnconfigure(2, weight=1, uniform="h")

        left_box = ctk.CTkFrame(header, fg_color="transparent")
        left_box.grid(row=0, column=0, sticky="w")
        ctk.CTkCheckBox(
            left_box, text="Write this sample",
            variable=ctk.BooleanVar(value=True),
            font=ctk.CTkFont(size=12),
        ).pack(side="left")
        ctk.CTkLabel(
            left_box, text="SHOTCRETE",
            fg_color=SHOT_PILL, text_color="white", corner_radius=10,
            font=ctk.CTkFont(size=10, weight="bold"), padx=8, pady=2,
        ).pack(side="left", padx=(10, 0))

        title = f"Cube No {cube['cube_no']}  ·  {cube['sample_mark']}"
        ctk.CTkLabel(
            header, text=title, font=ctk.CTkFont(size=15, weight="bold"),
        ).grid(row=0, column=1, sticky="")

        ctk.CTkLabel(
            header, text=f"[OK]  {cube['sheet']}",
            fg_color=OK_PILL, text_color="white", corner_radius=14,
            font=ctk.CTkFont(size=11, weight="bold"), padx=10, pady=3,
        ).grid(row=0, column=2, sticky="e")

        grid = ctk.CTkFrame(card, fg_color="transparent")
        grid.pack(padx=15, pady=(2, 12))

        hdr_font = ctk.CTkFont(size=11, weight="bold")
        ctk.CTkLabel(grid, text="", width=70).grid(row=0, column=0)
        ctk.CTkLabel(grid, text="", width=30).grid(row=0, column=1)
        ctk.CTkLabel(grid, text="Diameter (mm)", width=120, font=hdr_font).grid(row=0, column=2, padx=6)
        ctk.CTkLabel(grid, text="Height (mm)",   width=120, font=hdr_font).grid(row=0, column=3, padx=6)
        ctk.CTkLabel(grid, text="Weight (gr)",   width=130, font=hdr_font).grid(row=0, column=4, padx=6)
        ctk.CTkLabel(grid, text="Load (kN)",     width=130, font=hdr_font).grid(row=0, column=5, padx=6)
        ctk.CTkLabel(grid, text="Strength (N/mm²)", width=140, font=hdr_font).grid(row=0, column=6, padx=6)

        self._build_group(grid, 1, "7-day",  "#2E7D32", cube["rows_7d"])
        self._build_group(grid, 7, "28-day", "#C62828", cube["rows_28d"])

    def _build_group(self, grid, start_row, age_label, age_color, rows):
        order = sorted(range(len(rows)), key=lambda i: rows[i][4], reverse=True)
        selected = set(order[:3])

        ctk.CTkLabel(
            grid, text=age_label, text_color=age_color,
            font=ctk.CTkFont(size=13, weight="bold"),
            width=70, anchor="w",
        ).grid(row=start_row, column=0, rowspan=5, sticky="w",
               padx=(0, 4), pady=(6, 4))

        vars_ = []
        widgets = []
        strength_map = {i: rows[i][4] for i in range(len(rows))}

        def _apply_style(idx):
            is_on = vars_[idx].get()
            for ent in widgets[idx]:
                if is_on:
                    ent.configure(border_color=EMPTY_BORDER, border_width=3,
                                  fg_color=EMPTY_BG,
                                  text_color=("gray14", "gray84"))
                else:
                    ent.configure(border_color=("gray70", "gray30"),
                                  border_width=1, fg_color=DIM_BG,
                                  text_color=DIM_TEXT)

        def _on_toggle(idx):
            on_now = [i for i in range(len(rows)) if vars_[i].get()]
            if vars_[idx].get() and len(on_now) > 3:
                others = [i for i in on_now if i != idx]
                drop = min(others, key=lambda i: strength_map[i])
                vars_[drop].set(False)
                _apply_style(drop)
            _apply_style(idx)

        ent_font = ctk.CTkFont(size=14)
        for i, (diam, height, weight, load, strength) in enumerate(rows):
            r = start_row + i
            var = ctk.BooleanVar(value=i in selected)
            vars_.append(var)

            ctk.CTkCheckBox(
                grid, text="", variable=var, width=24,
                command=lambda ii=i: _on_toggle(ii),
            ).grid(row=r, column=1, padx=(0, 6), pady=3)

            row_entries = []
            for col, val, w in [
                (2, diam,    120),
                (3, height,  120),
                (4, weight,  130),
                (5, load,    130),
                (6, strength, 140),
            ]:
                sv = ctk.StringVar(value="" if val is None else str(val))
                ent = ctk.CTkEntry(
                    grid, textvariable=sv, width=w, height=34,
                    justify="center", font=ent_font,
                )
                ent.grid(row=r, column=col, padx=6, pady=3)
                row_entries.append(ent)

            widgets.append(row_entries)
            _apply_style(i)


if __name__ == "__main__":
    app = ShotcreteDemo()
    app.mainloop()
