# CSV/Excel Scrambler — GUI Application
#
# This file is the **GUI app** entrypoint. It is intentionally written in
# "plain Python" (Tkinter) so it can be packaged into a portable Windows `.exe`.
#
# If you're maintaining this with AI help, search for:
#   - "PIPELINE STEP" (high-level steps)
#   - "COMMON EDITS"   (where to change common behaviors)
#
import datetime as dt
import os
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List, Optional, Tuple

from scrambler_core import (
    ScrambleMode,
    ShuffleValuesError,
    _validate_unique_headers,
    _xlsx_header_row,
    default_mode_for_header,
    is_dateish_header,
    is_program_project_stable_header,
    scramble_tabular_file,
)

# VERSIONING
# - Bump this when you ship a new build so staff can confirm they're running the right EXE.
# - It is shown in the window title and in the top-right corner of the app.
APP_VERSION = "1.4.0"


MODES: List[ScrambleMode] = [
    "full scramble",
    "date scramble",
    "shuffle values",
    "no scramble",
    "delete data",
]


def _try_read_headers(path: str) -> Tuple[List[str], str, List[str]]:
    # PIPELINE STEP 1: Read only the **headers** (row 1) so we can show the user a list
    # of columns and let them choose how each column should be handled.
    # For Excel, unions header names from every worksheet (first-seen order); sheet_titles lists worksheet tabs.
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        import csv

        # Try common Windows encodings.
        for enc in ["utf-8-sig", "utf-8", "cp1252"]:
            try:
                with open(path, "r", encoding=enc, newline="") as f:
                    reader = csv.reader(f)
                    headers = next(reader)
                    return headers, enc, []
            except Exception:
                continue
        raise ValueError("Could not read CSV headers. The file may be locked, not a CSV, or uses an unknown encoding.")

    if ext in [".xlsx", ".xlsm"]:
        try:
            import openpyxl  # type: ignore
        except Exception as e:
            raise ValueError("Excel support is not available. (Missing openpyxl)") from e

        wb = openpyxl.load_workbook(path, data_only=True)
        sheet_titles: List[str] = []
        all_headers: List[str] = []
        for ws in wb.worksheets:
            sheet_titles.append(ws.title)
            hdr_row = _xlsx_header_row(ws)
            max_col = ws.max_column or 0
            row_headers: List[str] = []
            for c in range(1, max_col + 1):
                v = ws.cell(row=hdr_row, column=c).value
                row_headers.append("" if v is None else str(v))
            _validate_unique_headers(row_headers, sheet_name=ws.title)
            for h in row_headers:
                all_headers.append(f"{h} ({ws.title})")
        return all_headers, "xlsx", sheet_titles

    raise ValueError("Unsupported file type. Please choose a .csv or .xlsx file.")


class ScrollFrame(ttk.Frame):
    def __init__(self, parent: tk.Widget) -> None:
        super().__init__(parent)
        self.canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)

        self.inner.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )
        self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # Mousewheel scrolling
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _on_mousewheel(self, event: tk.Event) -> None:
        try:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception:
            pass


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(f"CSV Scrambler v{APP_VERSION}")
        self.geometry("900x700")

        self.input_path: Optional[str] = None
        self.output_path: Optional[str] = None
        self.encoding: str = "utf-8"
        self.headers: List[str] = []
        self.excel_sheet_titles: List[str] = []

        self.mode_vars: Dict[str, tk.StringVar] = {}
        self.keyed_var = tk.BooleanVar(value=False)

        self.status_var = tk.StringVar(value="Choose a CSV or Excel file to begin.")

        # PIPELINE STEP 0: Build the UI layout (buttons, header list, dropdowns).
        self._build_ui()

    def _build_ui(self) -> None:
        # PIPELINE STEP 0A: Top bar = file selection + output selection + version label.
        top = ttk.Frame(self, padding=12)
        top.pack(fill="x")

        ttk.Button(top, text="Choose File…", command=self.choose_input).pack(side="left")
        ttk.Button(top, text="Choose Output…", command=self.choose_output).pack(side="left", padx=(8, 0))

        self.input_label = ttk.Label(top, text="(no file selected)")
        self.input_label.pack(side="left", padx=(12, 0), fill="x", expand=True)

        ver = ttk.Label(top, text=f"v{APP_VERSION}", foreground="#666")
        ver.pack(side="right")

        # PIPELINE STEP 0B: Options row (keyed original checkbox, as-of date note).
        opts = ttk.Frame(self, padding=(12, 0, 12, 8))
        opts.pack(fill="x")

        ttk.Checkbutton(
            opts,
            text="Create Keyed Normal Version (also saves a keyed original copy with Row ID)",
            variable=self.keyed_var,
        ).pack(side="left")

        note = ttk.Label(
            opts,
            text=f"Age buckets (DOB headers) preserved as-of today: {dt.date.today().isoformat()}",
            foreground="#555",
        )
        note.pack(side="right")

        ttk.Separator(self).pack(fill="x", pady=(8, 0))

        # PIPELINE STEP 0C: Middle area = scrollable list of headers with dropdown modes.
        mid = ttk.Frame(self, padding=12)
        mid.pack(fill="both", expand=True)

        ttk.Label(
            mid,
            text="Column rules (you can change these before running):",
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w")

        self.sheet_hint_label = ttk.Label(mid, text="", foreground="#555", wraplength=820)
        self.sheet_hint_label.pack(anchor="w", pady=(4, 0))

        self.scroll = ScrollFrame(mid)
        self.scroll.pack(fill="both", expand=True, pady=(8, 8))

        # PIPELINE STEP 0D: Bottom bar = run button + status text.
        bottom = ttk.Frame(self, padding=12)
        bottom.pack(fill="x")

        ttk.Button(bottom, text="Run Scramble", command=self.run).pack(side="right")
        ttk.Label(bottom, textvariable=self.status_var).pack(side="left", fill="x", expand=True)

    def choose_input(self) -> None:
        # PIPELINE STEP 1: User chooses the input file (.csv or .xlsx).
        path = filedialog.askopenfilename(
            title="Choose the file to scramble",
            filetypes=[
                ("CSV or Excel files", "*.csv;*.xlsx"),
                ("CSV files", "*.csv"),
                ("Excel files", "*.xlsx"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return
        self.input_path = path
        self.input_label.config(text=os.path.basename(path))

        try:
            headers, enc, sheet_titles = _try_read_headers(path)
        except Exception as e:
            messagebox.showerror("Could not read file", str(e))
            return

        self.encoding = enc
        self.headers = headers
        self.excel_sheet_titles = sheet_titles

        # Safety guard: our per-column mapping uses the header text as a key.
        # If headers repeat, we can't reliably know which column a rule applies to.
        if len(set(self.headers)) != len(self.headers):
            messagebox.showerror(
                "Unsupported file",
                "This file has duplicate (repeated) column headers. Please fix the headers so each column name is unique, then try again.",
            )
            self.headers = []
            return

        # If output path not chosen yet, set a sensible default next to input.
        if not self.output_path:
            stem = os.path.splitext(os.path.basename(path))[0]
            in_ext = os.path.splitext(path)[1].lower()
            out_ext = ".xlsx" if in_ext in [".xlsx", ".xlsm"] else ".csv"
            self.output_path = os.path.join(os.path.dirname(path), f"{stem} Scrambled{out_ext}")

        # PIPELINE STEP 2: Show the header list + dropdown modes with sensible defaults.
        self._render_header_controls()
        self.status_var.set("Ready. Review column rules, then click Run Scramble.")

    def choose_output(self) -> None:
        # PIPELINE STEP 1B: User chooses where the output file will be saved.
        # (We default the extension based on the input file type.)
        initialdir = os.path.dirname(self.input_path) if self.input_path else os.getcwd()
        if self.input_path:
            in_ext = os.path.splitext(self.input_path)[1].lower()
            out_ext = ".xlsx" if in_ext in [".xlsx", ".xlsm"] else ".csv"
            initialfile = os.path.splitext(os.path.basename(self.input_path))[0] + f" Scrambled{out_ext}"
            def_ext = out_ext
        else:
            initialfile = "Scrambled.csv"
            def_ext = ".csv"
        path = filedialog.asksaveasfilename(
            title="Choose where to save the scrambled CSV",
            defaultextension=def_ext,
            initialdir=initialdir,
            initialfile=initialfile,
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if not path:
            return
        self.output_path = path
        self.status_var.set(f"Output set to: {os.path.basename(path)}")

    def _clear_scroll(self) -> None:
        for w in self.scroll.inner.winfo_children():
            w.destroy()

    def _render_header_controls(self) -> None:
        # PIPELINE STEP 2: Render the header list grouped into:
        # - Date fields
        # - Program/Project fields
        # - Other fields
        #
        # Each row gets a dropdown with the column handling mode.
        self._clear_scroll()
        self.mode_vars = {}

        if self.excel_sheet_titles:
            listed = ", ".join(self.excel_sheet_titles)
            self.sheet_hint_label.config(
                text=f"{len(self.excel_sheet_titles)} sheet(s): {listed} — column rules apply to every worksheet tab."
            )
        else:
            self.sheet_hint_label.config(text="")

        groups: List[Tuple[str, List[str]]] = [
            ("Date fields", [h for h in self.headers if is_dateish_header(h)]),
            ("Program/Project fields", [h for h in self.headers if is_program_project_stable_header(h)]),
        ]
        assigned = set(groups[0][1]) | set(groups[1][1])
        groups.append(("Other fields", [h for h in self.headers if h not in assigned]))

        for group_name, headers in groups:
            if not headers:
                continue
            header_lbl = ttk.Label(self.scroll.inner, text=group_name, font=("Segoe UI", 10, "bold"))
            header_lbl.pack(anchor="w", pady=(10, 4))

            for h in headers:
                row = ttk.Frame(self.scroll.inner)
                row.pack(fill="x", pady=2)

                ttk.Label(row, text=h, width=50).pack(side="left", anchor="w")

                var = tk.StringVar(value=default_mode_for_header(h))
                self.mode_vars[h] = var

                cb = ttk.Combobox(row, textvariable=var, values=MODES, state="readonly", width=18)
                cb.pack(side="left", padx=(8, 0))

                # Small inline hint for date-ish headers
                if is_dateish_header(h):
                    ttk.Label(row, text="(date rules available)", foreground="#666").pack(side="left", padx=(8, 0))

        warning = ttk.Label(
            self.scroll.inner,
            text="Note: “Full scramble” will scramble dates/times too (no date rules). Use “Date scramble” for stable month/year and DOB bucket preservation.",
            foreground="#884400",
            wraplength=820,
        )
        warning.pack(anchor="w", pady=(14, 0))

    def run(self) -> None:
        # PIPELINE STEP 3: Validate inputs, collect the chosen dropdown modes,
        # then run the scramble operation in a background thread.
        if not self.input_path or not self.headers:
            messagebox.showinfo("Select a file", "Please choose a CSV or Excel file first.")
            return
        if not self.output_path:
            self.choose_output()
            if not self.output_path:
                return

        header_modes: Dict[str, ScrambleMode] = {}
        for h in self.headers:
            var = self.mode_vars.get(h)
            if var is None:
                header_modes[h] = default_mode_for_header(h)
            else:
                header_modes[h] = var.get()  # type: ignore[assignment]

        self.status_var.set("Running…")

        def worker() -> None:
            # PIPELINE STEP 4: Call the core pipeline which:
            # - reads rows
            # - applies per-column modes
            # - optionally creates keyed original
            # - writes output file(s)
            try:
                res = scramble_tabular_file(
                    input_path=self.input_path or "",
                    output_scrambled_path=self.output_path or "",
                    header_modes=header_modes,
                    create_keyed_normal_version=bool(self.keyed_var.get()),
                    encoding=self.encoding,
                )
                msg = f"Done.\n\nScrambled file:\n{res.scrambled_output_path}"
                if res.keyed_original_output_path:
                    msg += f"\n\nKeyed original file:\n{res.keyed_original_output_path}"
                msg += f"\n\nRows processed: {res.row_count}"

                self.after(0, lambda: self._on_success(msg))
            except ShuffleValuesError as e:
                self.after(0, lambda: self._on_error(f"Shuffle values error:\n\n{e}"))
            except Exception as e:
                details = traceback.format_exc()
                self.after(0, lambda: self._on_error(f"{e}\n\nDetails:\n{details}"))

        threading.Thread(target=worker, daemon=True).start()

    def _on_success(self, msg: str) -> None:
        self.status_var.set("Complete.")
        messagebox.showinfo("Scramble complete", msg)

    def _on_error(self, msg: str) -> None:
        self.status_var.set("Error.")
        messagebox.showerror("Error", msg)


def main() -> int:
    # PIPELINE STEP 0: Launch the app.
    app = App()
    app.mainloop()
    return 0


# COMMON EDITS (for maintainers)
# - Change the version shown to staff: update APP_VERSION above.
# - Change which file types are allowed: update choose_input() filetypes + scrambler_core.scramble_tabular_file().
# - Change the default dropdown selections: edit scrambler_core.default_mode_for_header().


if __name__ == "__main__":
    raise SystemExit(main())

