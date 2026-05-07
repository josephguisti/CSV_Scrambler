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
from typing import Dict, List, Optional, Set, Tuple

from scrambler_core import (
    ScrambleMode,
    ShuffleValuesError,
    EncryptedXlsxError,
    _col_looks_like_dates,
    _decrypt_xlsx,
    _is_excel_error,
    _validate_unique_headers,
    _xlsx_header_row,
    detect_cross_sheet_cf_breaks,
    is_program_project_stable_header,
    scramble_tabular_file,
)

# VERSIONING
# - Bump this when you ship a new build so staff can confirm they're running the right EXE.
# - It is shown in the window title and in the top-right corner of the app.
APP_VERSION = "1.12.0"


MODES: List[ScrambleMode] = [
    "full scramble",
    "date scramble",
    "shuffle values",
    "no scramble",
    "delete data",
]


_SAMPLE_ROWS = 20  # number of data rows to inspect for date detection
_MAX_SCAN_ROWS = 2000  # max data rows to scan for date detection + unique samples
_MAX_UNIQUE_SAMPLES = 3
_SAMPLE_DISPLAY_LEN = 72


def _truncate_sample(s: str, max_len: int = _SAMPLE_DISPLAY_LEN) -> str:
    s = " ".join(s.split())
    if len(s) <= max_len:
        return s
    return s[: max_len - 1] + "…"


def _samples_preview_text(values: List[str]) -> str:
    if not values:
        return "no values"
    return "\n".join(_truncate_sample(v) for v in values)


def _try_read_headers(path: str, password: Optional[str] = None) -> Tuple[List[str], str, List[str], Set[str], Dict[str, List[str]]]:
    # PIPELINE STEP 1: Read headers (row 1) and scan data rows for date detection and sample previews.
    # Returns (header_keys, encoding, sheet_titles, detected_date_keys, sample_values_by_key).
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        import csv

        # Try common Windows encodings.
        for enc in ["utf-8-sig", "utf-8", "cp1252"]:
            try:
                with open(path, "r", encoding=enc, newline="") as f:
                    reader = csv.reader(f)
                    headers = next(reader)
                    n = len(headers)
                    samples: Dict[int, List[str]] = {i: [] for i in range(n)}
                    uniques: Dict[str, List[str]] = {h: [] for h in headers}
                    seen_u: Dict[str, Set[str]] = {h: set() for h in headers}
                    rows_scanned = 0
                    for row in reader:
                        if rows_scanned >= _MAX_SCAN_ROWS:
                            break
                        rows_scanned += 1
                        if len(row) < n:
                            row = row + [""] * (n - len(row))
                        for i in range(n):
                            val = row[i]
                            raw = "" if val is None else str(val)
                            stripped = raw.strip()
                            if not stripped:
                                continue
                            if len(samples[i]) < _SAMPLE_ROWS:
                                samples[i].append(raw)
                            h = headers[i]
                            if len(uniques[h]) < _MAX_UNIQUE_SAMPLES and stripped not in seen_u[h]:
                                seen_u[h].add(stripped)
                                uniques[h].append(stripped)
                    detected: Set[str] = {
                        headers[i] for i, slist in samples.items() if _col_looks_like_dates(slist)
                    }
                    return headers, enc, [], detected, uniques
            except Exception:
                continue
        raise ValueError("Could not read CSV headers. The file may be locked, not a CSV, or uses an unknown encoding.")

    if ext in [".xlsx", ".xlsm"]:
        import zipfile

        try:
            import openpyxl  # type: ignore
        except Exception as e:
            raise ValueError("Excel support is not available. (Missing openpyxl)") from e

        try:
            if password is not None:
                wb = openpyxl.load_workbook(_decrypt_xlsx(path, password), data_only=True)
            else:
                wb = openpyxl.load_workbook(path, data_only=True)
        except zipfile.BadZipFile as e:
            raise EncryptedXlsxError(
                "This file is password-protected. Enter the password to continue."
            ) from e
        sheet_titles: List[str] = []
        all_headers: List[str] = []
        detected_xlsx: Set[str] = set()
        all_uniques: Dict[str, List[str]] = {}
        for ws in wb.worksheets:
            sheet_titles.append(ws.title)
            hdr_row = _xlsx_header_row(ws)
            max_col = ws.max_column or 0
            max_row = ws.max_row or hdr_row
            row_headers: List[str] = []
            for c in range(1, max_col + 1):
                v = ws.cell(row=hdr_row, column=c).value
                row_headers.append("" if (v is None or _is_excel_error(v)) else str(v))
            _validate_unique_headers(row_headers, sheet_name=ws.title)
            col_samples: Dict[int, List[str]] = {c: [] for c in range(max_col)}
            keys_order: List[str] = []
            for c_idx, h in enumerate(row_headers):
                key = f"{h} ({ws.title})"
                keys_order.append(key)
                all_uniques[key] = []
            seen_per_col: Dict[int, Set[str]] = {c_idx: set() for c_idx in range(max_col)}
            rows_in_sheet = 0
            for r in range(hdr_row + 1, max_row + 1):
                if rows_in_sheet >= _MAX_SCAN_ROWS:
                    break
                rows_in_sheet += 1
                for c_idx in range(max_col):
                    val = ws.cell(row=r, column=c_idx + 1).value
                    s_raw = "" if val is None else str(val)
                    stripped = s_raw.strip()
                    if not stripped:
                        continue
                    if len(col_samples[c_idx]) < _SAMPLE_ROWS:
                        col_samples[c_idx].append(s_raw)
                    key = keys_order[c_idx]
                    ul = all_uniques[key]
                    if len(ul) < _MAX_UNIQUE_SAMPLES and stripped not in seen_per_col[c_idx]:
                        seen_per_col[c_idx].add(stripped)
                        ul.append(stripped)
            for c_idx, h in enumerate(row_headers):
                key = f"{h} ({ws.title})"
                all_headers.append(key)
                if _col_looks_like_dates(col_samples[c_idx]):
                    detected_xlsx.add(key)
        return all_headers, "xlsx", sheet_titles, detected_xlsx, all_uniques

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

    def scroll_row_into_view(self, widget: tk.Widget) -> None:
        """Vertically scroll the canvas so *widget* (a descendant of *inner*) is centered."""
        inner = self.inner
        self.canvas.update_idletasks()
        self.update_idletasks()
        y = 0.0
        w: Optional[tk.Widget] = widget
        while w is not None and w != inner:
            y += float(w.winfo_y())
            w = w.master
        h = max(1, widget.winfo_height())
        y_center = y + h / 2.0
        ch = max(1, self.canvas.winfo_height())
        ih = max(1, inner.winfo_reqheight())
        max_scroll = max(0.0, float(ih - ch))
        target_top = y_center - ch / 2.0
        target_top = max(0.0, min(max_scroll, target_top))
        fraction = target_top / max_scroll if max_scroll > 0 else 0.0
        fraction = max(0.0, min(1.0, fraction))
        try:
            self.canvas.yview_moveto(fraction)
        except Exception:
            pass


class SheetPickerDialog(tk.Toplevel):
    """Modal dialog to choose which worksheet tabs to include in the export."""

    def __init__(self, parent: tk.Tk, sheet_titles: List[str]) -> None:
        super().__init__(parent)
        self.title("Select sheets to export")
        self.result: Optional[List[str]] = None
        self.transient(parent)
        self.grab_set()

        body = ttk.Frame(self, padding=12)
        body.pack(fill="both", expand=True)

        ttk.Label(
            body,
            text="Choose which worksheet tabs to include in the scrambled output:",
            wraplength=520,
        ).pack(anchor="w")

        canvas_h = min(420, max(140, 26 * len(sheet_titles) + 40))
        cf = ttk.Frame(body)
        cf.pack(fill="both", expand=True, pady=(8, 8))

        canvas = tk.Canvas(cf, height=canvas_h, highlightthickness=0)
        sb = ttk.Scrollbar(cf, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)

        def _on_inner_configure(_event: tk.Event) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))

        inner.bind("<Configure>", _on_inner_configure)
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)

        self._vars: Dict[str, tk.BooleanVar] = {}
        for title in sheet_titles:
            v = tk.BooleanVar(value=True)
            self._vars[title] = v
            ttk.Checkbutton(inner, text=title, variable=v).pack(anchor="w")

        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        def _wheel(event: tk.Event) -> None:
            try:
                canvas.yview_scroll(int(-event.delta / 120), "units")
            except Exception:
                pass

        canvas.bind("<MouseWheel>", _wheel)

        btn_row = ttk.Frame(body)
        btn_row.pack(fill="x")
        ttk.Button(btn_row, text="Select all", command=self._select_all).pack(side="left")
        ttk.Button(btn_row, text="Deselect all", command=self._deselect_all).pack(side="left", padx=(8, 0))

        ok_cancel = ttk.Frame(body)
        ok_cancel.pack(fill="x", pady=(12, 0))
        ttk.Button(ok_cancel, text="OK", command=self._ok).pack(side="right")
        ttk.Button(ok_cancel, text="Cancel", command=self._cancel).pack(side="right", padx=(0, 8))

        self.protocol("WM_DELETE_WINDOW", self._cancel)

    def _select_all(self) -> None:
        for v in self._vars.values():
            v.set(True)

    def _deselect_all(self) -> None:
        for v in self._vars.values():
            v.set(False)

    def _ok(self) -> None:
        sel = [t for t, v in self._vars.items() if v.get()]
        if not sel:
            messagebox.showwarning("No sheets selected", "Select at least one worksheet tab.", parent=self)
            return
        self.result = sel
        self.grab_release()
        self.destroy()

    def _cancel(self) -> None:
        self.result = None
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()


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
        self.detected_date_keys: Set[str] = set()
        self.sample_values_by_header: Dict[str, List[str]] = {}
        self.xlsx_password: Optional[str] = None
        self.selected_sheets: Optional[List[str]] = None

        self.mode_vars: Dict[str, tk.StringVar] = {}
        self.column_row_frames: Dict[str, tk.Widget] = {}
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

        self.total_fields_label = ttk.Label(mid, text="", foreground="#333", font=("Segoe UI", 10))
        self.total_fields_label.pack(anchor="w", pady=(4, 0))

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
        self.xlsx_password = None  # reset whenever a new file is chosen
        self.selected_sheets = None
        self.input_label.config(text=os.path.basename(path))

        try:
            headers, enc, sheet_titles, detected_date_keys, sample_values = _try_read_headers(path)
        except EncryptedXlsxError:
            from tkinter import simpledialog
            pw = simpledialog.askstring(
                "Password Required",
                f"{os.path.basename(path)} is password-protected.\nEnter the password:",
                show="*",
            )
            if not pw:
                return  # user cancelled
            try:
                headers, enc, sheet_titles, detected_date_keys, sample_values = _try_read_headers(path, password=pw)
            except Exception as e:
                messagebox.showerror("Wrong password or corrupt file", str(e))
                return
            self.xlsx_password = pw
        except Exception as e:
            messagebox.showerror("Could not read file", str(e))
            return

        self.encoding = enc
        self.headers = headers
        self.excel_sheet_titles = sheet_titles
        self.detected_date_keys = detected_date_keys
        self.sample_values_by_header = sample_values

        # Safety guard: our per-column mapping uses the header text as a key.
        # If headers repeat, we can't reliably know which column a rule applies to.
        if len(set(self.headers)) != len(self.headers):
            messagebox.showerror(
                "Unsupported file",
                "This file has duplicate (repeated) column headers. Please fix the headers so each column name is unique, then try again.",
            )
            self.headers = []
            self.selected_sheets = None
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

    def _filtered_headers(self) -> List[str]:
        """Column keys (`Header (SheetTitle)`) restricted to selected worksheets when set."""
        if not self.headers:
            return []
        if self.selected_sheets is None:
            return self.headers
        sel = set(self.selected_sheets)
        out: List[str] = []
        for h in self.headers:
            for st in sel:
                if h.endswith(f" ({st})"):
                    out.append(h)
                    break
        return out

    def _build_field_groups(self, hdr_subset: List[str]) -> List[Tuple[str, List[str]]]:
        """Date / Program-Project / Other groups over a subset of column keys."""
        groups: List[Tuple[str, List[str]]] = [
            ("Date fields (detected from data)", [h for h in hdr_subset if h in self.detected_date_keys]),
            ("Program/Project fields", [h for h in hdr_subset if is_program_project_stable_header(h)]),
        ]
        assigned = set(groups[0][1]) | set(groups[1][1])
        groups.append(("Other fields", [h for h in hdr_subset if h not in assigned]))
        return groups

    def _render_column_row(self, h: str, *, initial_mode: str = "no scramble") -> None:
        row = ttk.Frame(self.scroll.inner)
        row.pack(fill="x", pady=(4, 4))
        self.column_row_frames[h] = row

        left = ttk.Frame(row)
        left.pack(side="left", fill="both", expand=True, anchor="nw")
        ttk.Label(left, text=h, wraplength=540, justify="left").pack(anchor="nw")
        sample_text = _samples_preview_text(self.sample_values_by_header.get(h, []))
        ttk.Label(
            left,
            text=sample_text,
            font=("Segoe UI", 9),
            foreground="#666",
            wraplength=540,
            justify="left",
        ).pack(anchor="nw", pady=(2, 0))

        is_date = h in self.detected_date_keys
        var = tk.StringVar(value=initial_mode)
        self.mode_vars[h] = var

        cb = ttk.Combobox(row, textvariable=var, values=MODES, state="readonly", width=18)
        cb.pack(side="left", padx=(8, 0), anchor="n")

        if is_date:
            ttk.Label(row, text="(date rules available)", foreground="#666").pack(side="left", padx=(8, 0), anchor="n")

    def _clear_scroll(self) -> None:
        for w in self.scroll.inner.winfo_children():
            w.destroy()

    def _render_header_controls(self) -> None:
        # PIPELINE STEP 2: Render the header list grouped into:
        # - Multi-sheet Excel: Worksheet sections, each with Date / Program-Project / Other.
        # - CSV or single-sheet Excel: flat Date / Program / Other (no worksheet heading).
        #
        # Each row gets a dropdown with the column handling mode.
        # Preserve any mode values the user already set before this re-render.
        old_modes = {h: v.get() for h, v in self.mode_vars.items()}
        self._clear_scroll()
        self.mode_vars = {}
        self.column_row_frames = {}

        if self.excel_sheet_titles:
            if self.selected_sheets is not None:
                listed = ", ".join(self.selected_sheets)
                self.sheet_hint_label.config(
                    text=f"{len(self.selected_sheets)} sheet(s) selected for export: {listed} — "
                    "rules below apply only to included worksheet tabs."
                )
            else:
                listed = ", ".join(self.excel_sheet_titles)
                self.sheet_hint_label.config(
                    text=f"{len(self.excel_sheet_titles)} sheet(s): {listed} — column rules apply to every worksheet tab."
                )
        else:
            self.sheet_hint_label.config(text="")

        hdr_list = self._filtered_headers()
        n_total = len(hdr_list)
        self.total_fields_label.config(text=f"Total fields: {n_total}" if n_total else "")

        multi_sheet_excel = bool(self.excel_sheet_titles) and len(self.excel_sheet_titles) > 1

        def render_groups(groups: List[Tuple[str, List[str]]]) -> None:
            for group_name, headers in groups:
                if not headers:
                    continue
                header_lbl = ttk.Label(
                    self.scroll.inner,
                    text=f"{group_name} — {len(headers)} fields",
                    font=("Segoe UI", 10, "bold"),
                )
                header_lbl.pack(anchor="w", pady=(10, 4))

                for h in headers:
                    self._render_column_row(h, initial_mode=old_modes.get(h, "no scramble"))

        if multi_sheet_excel:
            sheet_order = self.selected_sheets if self.selected_sheets is not None else self.excel_sheet_titles
            first_sheet_section = True
            for st in sheet_order:
                sheet_headers = [h for h in hdr_list if h.endswith(f" ({st})")]
                if not sheet_headers:
                    continue
                ttk.Label(
                    self.scroll.inner,
                    text=f"Worksheet: {st}",
                    font=("Segoe UI", 11, "bold"),
                ).pack(anchor="w", pady=(22, 6) if not first_sheet_section else (8, 6))
                first_sheet_section = False
                groups = self._build_field_groups(sheet_headers)
                render_groups(groups)
        else:
            groups = self._build_field_groups(hdr_list)
            render_groups(groups)

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

        out_ext = os.path.splitext(self.output_path)[1].lower()

        if len(self.excel_sheet_titles) > 1:
            dlg = SheetPickerDialog(self, self.excel_sheet_titles)
            self.wait_window(dlg)
            if dlg.result is None:
                return
            selected = dlg.result
            try:
                breaks = detect_cross_sheet_cf_breaks(
                    self.input_path or "",
                    selected,
                    password=self.xlsx_password,
                )
            except EncryptedXlsxError:
                messagebox.showerror(
                    "Password required",
                    "Could not re-open the workbook to check sheet dependencies. Try selecting the file again.",
                )
                return
            except Exception as e:
                messagebox.showerror("Could not check sheet dependencies", str(e))
                return
            if breaks:
                warn_text = "\n".join(f"  • {b}" for b in breaks[:25])
                if len(breaks) > 25:
                    warn_text += f"\n  … and {len(breaks) - 25} more"
                proceed = messagebox.askyesno(
                    "Cross-sheet dependency warning",
                    "Excluding sheets may break these conditional formatting relationships:\n\n"
                    f"{warn_text}\n\nProceed anyway?",
                )
                if not proceed:
                    return
            self.selected_sheets = selected
            self._render_header_controls()
        else:
            self.selected_sheets = None

        n_export_sheets = (
            len(self.selected_sheets)
            if self.selected_sheets is not None
            else max(1, len(self.excel_sheet_titles))
        )
        if len(self.excel_sheet_titles) > 0 and out_ext == ".csv" and n_export_sheets > 1:
            messagebox.showerror(
                "Cannot use CSV output",
                "Cannot save more than one worksheet as a CSV file.\n\n"
                "Choose an output path ending in .xlsx, or select only one sheet in the sheet picker, "
                "or save sheets separately outside this tool.",
            )
            return

        hdr_for_modes = self._filtered_headers()
        if not hdr_for_modes:
            messagebox.showinfo("No columns", "There are no columns to scramble for the current selection.")
            return

        header_modes: Dict[str, ScrambleMode] = {}
        for h in hdr_for_modes:
            var = self.mode_vars.get(h)
            if var is None:
                header_modes[h] = "no scramble"
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
                    password=self.xlsx_password,
                    include_sheets=self.selected_sheets,
                )
                msg = f"Done.\n\nScrambled file:\n{res.scrambled_output_path}"
                if res.keyed_original_output_path:
                    msg += f"\n\nKeyed original file:\n{res.keyed_original_output_path}"
                msg += f"\n\nRows processed: {res.row_count}"
                if res.warnings:
                    skipped = "\n".join(f"  • {w}" for w in res.warnings[:10])
                    if len(res.warnings) > 10:
                        skipped += f"\n  … and {len(res.warnings) - 10} more"
                    msg += f"\n\nNote: {len(res.warnings)} conditional formatting rule(s) could not be shifted and were skipped:\n{skipped}"

                self.after(0, lambda: self._on_success(msg))
            except ShuffleValuesError as e:
                self.after(0, lambda err=e: self._on_shuffle_values_error(err))
            except Exception as e:
                details = traceback.format_exc()
                self.after(0, lambda: self._on_error(f"{e}\n\nDetails:\n{details}"))

        threading.Thread(target=worker, daemon=True).start()

    def _on_shuffle_values_error(self, e: ShuffleValuesError) -> None:
        key = getattr(e, "column_key", None)
        if key and key in self.column_row_frames:
            try:
                self.scroll.scroll_row_into_view(self.column_row_frames[key])
            except Exception:
                pass
        self._on_error(f"Shuffle values error:\n\n{e}")

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
# - Change the default dropdown selections: edit initial mode in _render_header_controls() (default "no scramble").


if __name__ == "__main__":
    raise SystemExit(main())

