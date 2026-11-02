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
import secrets
import threading
import traceback
import tkinter as tk
from dataclasses import dataclass, field
from tkinter import filedialog, messagebox, simpledialog, ttk
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
APP_VERSION = "1.13.0"


# Radios use the bare header as their value. This sentinel means nothing is selected.
_NO_LINK = "\x00"
_LINK_MODES = ("full scramble", "date scramble")


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


def _default_output_path(input_path: str) -> str:
    stem, in_ext = os.path.splitext(os.path.basename(input_path))
    out_ext = ".xlsx" if in_ext.lower() in (".xlsx", ".xlsm") else ".csv"
    return os.path.join(os.path.dirname(input_path), f"{stem} Scrambled{out_ext}")


def _bare_header(key: str, sheet_titles: List[str]) -> str:
    """Column name with its known sheet suffix removed.

    Matches the longest sheet title so a header that contains parentheses is kept intact.
    """
    best: Optional[str] = None
    for title in sheet_titles:
        suffix = f" ({title})"
        if key.endswith(suffix) and (best is None or len(suffix) > len(best)):
            best = suffix
    if best:
        return key[: -len(best)]
    return key


def _sheet_of_key(key: str, sheet_titles: List[str]) -> Optional[str]:
    """Worksheet title encoded in an Excel column key, if any."""
    bare = _bare_header(key, sheet_titles)
    if bare == key:
        return None
    suffix = key[len(bare) :]
    if suffix.startswith(" (") and suffix.endswith(")"):
        return suffix[2:-1]
    return None


@dataclass
class LoadedFile:
    path: str
    encoding: str
    password: Optional[str]
    sheet_titles: List[str]
    headers: List[str]
    detected_date_keys: Set[str]
    sample_values: Dict[str, List[str]]
    selected_sheets: Optional[List[str]] = None
    mode_vars: Dict[str, tk.StringVar] = field(default_factory=dict)


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

        self.files: List[LoadedFile] = []
        self.output_path: Optional[str] = None
        self.link_header_var = tk.StringVar(value=_NO_LINK)
        self._syncing_modes = False

        self.column_row_frames: Dict[Tuple[str, str], tk.Widget] = {}
        self.keyed_var = tk.BooleanVar(value=False)

        self.status_var = tk.StringVar(value="Choose CSV or Excel files to begin.")

        # PIPELINE STEP 0: Build the UI layout (buttons, header list, dropdowns).
        self._build_ui()

    def _build_ui(self) -> None:
        # PIPELINE STEP 0A: Top bar = file selection + output selection + version label.
        top = ttk.Frame(self, padding=12)
        top.pack(fill="x")

        ttk.Button(top, text="Choose files…", command=self.choose_files).pack(side="left")
        self.output_button = ttk.Button(top, text="Choose Output…", command=self.choose_output)
        self.output_button.pack(side="left", padx=(8, 0))
        self._update_output_controls()

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

        self.link_hint_label = ttk.Label(
            mid,
            text="Select one field to keep the same scrambled values in every file and sheet.",
            foreground="#333",
            wraplength=820,
        )
        self.link_hint_label.pack(anchor="w", pady=(4, 0))

        self.scroll = ScrollFrame(mid)
        self.scroll.pack(fill="both", expand=True, pady=(8, 8))

        # PIPELINE STEP 0D: Bottom bar = run button + status text.
        bottom = ttk.Frame(self, padding=12)
        bottom.pack(fill="x")

        ttk.Button(bottom, text="Run Scramble", command=self.run).pack(side="right")
        ttk.Label(bottom, textvariable=self.status_var).pack(side="left", fill="x", expand=True)

    def choose_files(self) -> None:
        # PIPELINE STEP 1: User chooses one or more input files (.csv or .xlsx).
        # Ctrl+click and Shift+click work in the Windows file dialog.
        paths = filedialog.askopenfilenames(
            parent=self,
            title="Choose files to scramble (Ctrl+click to select more than one)",
            filetypes=[
                ("CSV or Excel files", "*.csv;*.xlsx"),
                ("CSV files", "*.csv"),
                ("Excel files", "*.xlsx"),
                ("All files", "*.*"),
            ],
        )
        if not paths:
            return

        loaded: List[LoadedFile] = []
        failures: List[str] = []
        for path in paths:
            try:
                loaded.append(self._load_file(path))
            except Exception as e:
                failures.append(f"{os.path.basename(path)}: {e}")

        if not loaded:
            messagebox.showerror(
                "Could not read files",
                "None of the selected files could be opened.\n\n" + "\n".join(failures),
            )
            return

        if failures:
            messagebox.showerror(
                "Some files could not be opened",
                "These files were skipped:\n\n" + "\n".join(failures),
            )

        self.files = loaded
        self.link_header_var.set(_NO_LINK)
        self.output_path = None
        self._update_output_controls()
        self.input_label.config(text=self._input_label_text())
        self._render_header_controls()
        if len(self.files) > 1:
            self.status_var.set("Ready. Select the shared field, review column rules, then click Run Scramble.")
        else:
            self.status_var.set("Ready. Review column rules, then click Run Scramble.")

    def _load_file(self, path: str) -> LoadedFile:
        password: Optional[str] = None
        try:
            headers, enc, sheet_titles, detected_date_keys, sample_values = _try_read_headers(path)
        except EncryptedXlsxError:
            pw = simpledialog.askstring(
                "Password Required",
                f"{os.path.basename(path)} is password-protected.\nEnter the password:",
                show="*",
                parent=self,
            )
            if not pw:
                raise ValueError("password was not entered")
            headers, enc, sheet_titles, detected_date_keys, sample_values = _try_read_headers(path, password=pw)
            password = pw

        # Safety guard: per-column mapping uses the header text as a key.
        if len(set(headers)) != len(headers):
            raise ValueError(
                "duplicate column headers. Fix the headers so each column name is unique, then try again."
            )

        return LoadedFile(
            path=path,
            encoding=enc,
            password=password,
            sheet_titles=sheet_titles,
            headers=headers,
            detected_date_keys=detected_date_keys,
            sample_values=sample_values,
        )

    def _input_label_text(self) -> str:
        if not self.files:
            return "(no file selected)"
        names = [os.path.basename(f.path) for f in self.files]
        if len(names) == 1:
            return names[0]
        shown = ", ".join(names)
        if len(shown) > 80:
            shown = shown[:77] + "…"
        return f"{len(names)} files: {shown}"

    def _update_output_controls(self) -> None:
        # One file can still pick a custom output path. Several files each
        # write a sibling "Name Scrambled" file next to the source.
        if len(self.files) == 1:
            self.output_button.configure(state="normal")
            if not self.output_path:
                self.output_path = _default_output_path(self.files[0].path)
        else:
            self.output_button.configure(state="disabled")
            self.output_path = None

    def choose_output(self) -> None:
        # PIPELINE STEP 1B: Single-file runs can choose where the output is saved.
        if len(self.files) != 1:
            return
        source = self.files[0].path
        initialdir = os.path.dirname(source) if source else os.getcwd()
        in_ext = os.path.splitext(source)[1].lower()
        out_ext = ".xlsx" if in_ext in [".xlsx", ".xlsm"] else ".csv"
        initialfile = os.path.splitext(os.path.basename(source))[0] + f" Scrambled{out_ext}"
        path = filedialog.asksaveasfilename(
            title="Choose where to save the scrambled file",
            defaultextension=out_ext,
            initialdir=initialdir,
            initialfile=initialfile,
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if not path:
            return
        self.output_path = path
        self.status_var.set(f"Output set to: {os.path.basename(path)}")

    def _filtered_headers(self, loaded: LoadedFile) -> List[str]:
        """Column keys restricted to the worksheets chosen for this file."""
        if loaded.selected_sheets is None:
            return list(loaded.headers)
        sel = set(loaded.selected_sheets)
        return [h for h in loaded.headers if _sheet_of_key(h, loaded.sheet_titles) in sel]

    def _file_has_bare_header(self, loaded: LoadedFile, bare: str) -> bool:
        return any(_bare_header(h, loaded.sheet_titles) == bare for h in self._filtered_headers(loaded))

    def _link_mode(self, bare: str) -> Optional[str]:
        for loaded in self.files:
            for key, var in loaded.mode_vars.items():
                if _bare_header(key, loaded.sheet_titles) == bare:
                    return var.get()
        return None

    def _sync_link_modes(self, bare: str, mode: str) -> None:
        """Every column with this header uses the same dropdown mode."""
        if self._syncing_modes:
            return
        self._syncing_modes = True
        try:
            for loaded in self.files:
                for key, var in loaded.mode_vars.items():
                    if _bare_header(key, loaded.sheet_titles) == bare and var.get() != mode:
                        var.set(mode)
        finally:
            self._syncing_modes = False

    def _on_mode_selected(self, bare: str, mode: str) -> None:
        if self.link_header_var.get() != bare:
            return
        self._sync_link_modes(bare, mode)

    def _build_field_groups(self, loaded: LoadedFile, hdr_subset: List[str]) -> List[Tuple[str, List[str]]]:
        """Date / Program-Project / Other groups over a subset of column keys."""
        groups: List[Tuple[str, List[str]]] = [
            ("Date fields (detected from data)", [h for h in hdr_subset if h in loaded.detected_date_keys]),
            ("Program/Project fields", [h for h in hdr_subset if is_program_project_stable_header(h)]),
        ]
        assigned = set(groups[0][1]) | set(groups[1][1])
        groups.append(("Other fields", [h for h in hdr_subset if h not in assigned]))
        return groups

    def _render_column_row(self, loaded: LoadedFile, h: str) -> None:
        bare = _bare_header(h, loaded.sheet_titles)
        row = ttk.Frame(self.scroll.inner)
        row.pack(fill="x", pady=(4, 4))
        self.column_row_frames[(loaded.path, h)] = row

        var = loaded.mode_vars.get(h)
        if var is None:
            var = tk.StringVar(value="no scramble")
            loaded.mode_vars[h] = var

        ttk.Radiobutton(
            row,
            variable=self.link_header_var,
            value=bare,
            command=lambda b=bare, mv=var: self._sync_link_modes(b, mv.get()),
        ).pack(side="left", anchor="n", padx=(0, 6))

        left = ttk.Frame(row)
        left.pack(side="left", fill="both", expand=True, anchor="nw")
        ttk.Label(left, text=h, wraplength=520, justify="left").pack(anchor="nw")
        sample_text = _samples_preview_text(loaded.sample_values.get(h, []))
        ttk.Label(
            left,
            text=sample_text,
            font=("Segoe UI", 9),
            foreground="#666",
            wraplength=520,
            justify="left",
        ).pack(anchor="nw", pady=(2, 0))

        cb = ttk.Combobox(row, textvariable=var, values=MODES, state="readonly", width=18)
        cb.pack(side="left", padx=(8, 0), anchor="n")
        cb.bind(
            "<<ComboboxSelected>>",
            lambda _event, b=bare, mv=var: self._on_mode_selected(b, mv.get()),
        )

        if h in loaded.detected_date_keys:
            ttk.Label(row, text="(date rules available)", foreground="#666").pack(side="left", padx=(8, 0), anchor="n")

    def _clear_scroll(self) -> None:
        for w in self.scroll.inner.winfo_children():
            w.destroy()

    def _update_sheet_hint(self) -> None:
        if len(self.files) == 1 and self.files[0].sheet_titles:
            loaded = self.files[0]
            if loaded.selected_sheets is not None:
                listed = ", ".join(loaded.selected_sheets)
                self.sheet_hint_label.config(
                    text=f"{len(loaded.selected_sheets)} sheet(s) selected for export: {listed} — "
                    "rules below apply only to included worksheet tabs."
                )
            else:
                listed = ", ".join(loaded.sheet_titles)
                self.sheet_hint_label.config(
                    text=f"{len(loaded.sheet_titles)} sheet(s): {listed} — column rules apply to every worksheet tab."
                )
            return
        if len(self.files) > 1:
            self.sheet_hint_label.config(
                text=f"{len(self.files)} files. Each scrambled file is saved next to its source."
            )
            return
        self.sheet_hint_label.config(text="")

    def _render_header_controls(self) -> None:
        # PIPELINE STEP 2: Render the header list grouped into:
        # - Several files: a File section, then worksheet sections when a workbook has multiple tabs.
        # - One multi-sheet Excel file: Worksheet sections, each with Date / Program-Project / Other.
        # - One CSV or single-sheet Excel file: flat Date / Program / Other (no file or worksheet heading).
        #
        # Mode vars live on each LoadedFile, so a re-render keeps the user's dropdown choices.
        self._clear_scroll()
        self.column_row_frames = {}
        self._update_sheet_hint()

        n_total = sum(len(self._filtered_headers(loaded)) for loaded in self.files)
        self.total_fields_label.config(text=f"Total fields: {n_total}" if n_total else "")

        show_file_heading = len(self.files) > 1
        first_file_section = True
        for loaded in self.files:
            if show_file_heading:
                ttk.Label(
                    self.scroll.inner,
                    text=f"File: {os.path.basename(loaded.path)}",
                    font=("Segoe UI", 11, "bold"),
                ).pack(anchor="w", pady=(22, 6) if not first_file_section else (8, 6))
                first_file_section = False
            self._render_file_columns(loaded)

        if self.files:
            warning = ttk.Label(
                self.scroll.inner,
                text="Note: “Full scramble” will scramble dates/times too (no date rules). Use “Date scramble” for stable month/year and DOB bucket preservation.",
                foreground="#884400",
                wraplength=820,
            )
            warning.pack(anchor="w", pady=(14, 0))

    def _render_groups(self, loaded: LoadedFile, groups: List[Tuple[str, List[str]]]) -> None:
        for group_name, headers in groups:
            if not headers:
                continue
            ttk.Label(
                self.scroll.inner,
                text=f"{group_name} — {len(headers)} fields",
                font=("Segoe UI", 10, "bold"),
            ).pack(anchor="w", pady=(10, 4))
            for h in headers:
                self._render_column_row(loaded, h)

    def _render_file_columns(self, loaded: LoadedFile) -> None:
        hdr_list = self._filtered_headers(loaded)
        multi_sheet = len(loaded.sheet_titles) > 1
        if not multi_sheet:
            self._render_groups(loaded, self._build_field_groups(loaded, hdr_list))
            return

        sheet_order = loaded.selected_sheets if loaded.selected_sheets is not None else loaded.sheet_titles
        first_sheet_section = True
        for st in sheet_order:
            sheet_headers = [h for h in hdr_list if _sheet_of_key(h, loaded.sheet_titles) == st]
            if not sheet_headers:
                continue
            ttk.Label(
                self.scroll.inner,
                text=f"Worksheet: {st}",
                font=("Segoe UI", 11, "bold"),
            ).pack(anchor="w", pady=(22, 6) if not first_sheet_section else (8, 6))
            first_sheet_section = False
            self._render_groups(loaded, self._build_field_groups(loaded, sheet_headers))

    def _selected_link_header(self) -> Optional[str]:
        raw = self.link_header_var.get()
        if not raw or raw == _NO_LINK:
            return None
        return raw

    def _pick_sheets(self) -> bool:
        """Ask which tabs to export for each multi-sheet workbook. Return False if cancelled."""
        pending: List[Optional[List[str]]] = []
        for loaded in self.files:
            if len(loaded.sheet_titles) <= 1:
                pending.append(None)
                continue
            dlg = SheetPickerDialog(self, loaded.sheet_titles)
            dlg.title(f"Select sheets to export — {os.path.basename(loaded.path)}")
            self.wait_window(dlg)
            if dlg.result is None:
                return False
            selected = dlg.result
            try:
                breaks = detect_cross_sheet_cf_breaks(
                    loaded.path,
                    selected,
                    password=loaded.password,
                )
            except EncryptedXlsxError:
                messagebox.showerror(
                    "Password required",
                    f"Could not re-open {os.path.basename(loaded.path)} to check sheet dependencies. "
                    "Try selecting the file again.",
                )
                return False
            except Exception as e:
                messagebox.showerror("Could not check sheet dependencies", str(e))
                return False
            if breaks:
                warn_text = "\n".join(f"  • {b}" for b in breaks[:25])
                if len(breaks) > 25:
                    warn_text += f"\n  … and {len(breaks) - 25} more"
                proceed = messagebox.askyesno(
                    "Cross-sheet dependency warning",
                    f"{os.path.basename(loaded.path)}: excluding sheets may break these "
                    "conditional formatting relationships:\n\n"
                    f"{warn_text}\n\nProceed anyway?",
                )
                if not proceed:
                    return False
            pending.append(selected)

        for loaded, selected in zip(self.files, pending):
            loaded.selected_sheets = selected
        self._render_header_controls()
        return True

    def _output_path_for(self, loaded: LoadedFile) -> str:
        if len(self.files) == 1 and self.output_path:
            return self.output_path
        return _default_output_path(loaded.path)

    def _keyed_output_path(self, output_path: str) -> str:
        base, ext = os.path.splitext(output_path)
        return f"{base} (KEYED ORIGINAL){ext or '.xlsx'}"

    def _output_is_locked(self, path: str) -> bool:
        if not path or not os.path.exists(path):
            return False
        try:
            with open(path, "ab"):
                pass
        except PermissionError:
            messagebox.showerror(
                "Output file is open",
                f"Cannot write to:\n{path}\n\n"
                "The file is open in another program (Excel, OneDrive sync, etc.).\n\n"
                "Close the file and try again.",
            )
            return True
        return False

    def run(self) -> None:
        # PIPELINE STEP 3: Validate inputs, collect the chosen dropdown modes,
        # then run the scramble operation in a background thread.
        if not self.files:
            messagebox.showinfo("Select files", "Please choose a CSV or Excel file first.")
            return

        link_header = self._selected_link_header()
        if len(self.files) > 1 and link_header is None:
            messagebox.showinfo(
                "Select a shared field",
                "Choose the field that should keep the same scrambled values across files.\n\n"
                "Click the circle to the left of that column.",
            )
            return

        link_mode: Optional[ScrambleMode] = None
        if link_header is not None:
            chosen = self._link_mode(link_header)
            if chosen not in _LINK_MODES:
                shown = chosen if chosen else "no scramble"
                messagebox.showerror(
                    "Choose a scramble mode",
                    f"“{link_header}” is set to {shown}.\n\n"
                    "The shared field must use Full scramble or Date scramble so each original "
                    "value always becomes the same scrambled value.",
                )
                return
            link_mode = chosen  # type: ignore[assignment]

        if len(self.files) == 1 and not self.output_path:
            self.choose_output()
            if not self.output_path:
                return

        if not self._pick_sheets():
            return

        if len(self.files) == 1 and self.output_path:
            loaded = self.files[0]
            out_ext = os.path.splitext(self.output_path)[1].lower()
            n_export = (
                len(loaded.selected_sheets)
                if loaded.selected_sheets is not None
                else max(1, len(loaded.sheet_titles))
            )
            if loaded.sheet_titles and out_ext == ".csv" and n_export > 1:
                messagebox.showerror(
                    "Cannot use CSV output",
                    "Cannot save more than one worksheet as a CSV file.\n\n"
                    "Choose an output path ending in .xlsx, or select only one sheet in the sheet picker, "
                    "or save sheets separately outside this tool.",
                )
                return

        for loaded in self.files:
            if not self._filtered_headers(loaded):
                messagebox.showinfo(
                    "No columns",
                    f"{os.path.basename(loaded.path)} has no columns to scramble for the current selection.",
                )
                return

        missing: List[str] = []
        if link_header is not None:
            missing = [
                os.path.basename(loaded.path)
                for loaded in self.files
                if not self._file_has_bare_header(loaded, link_header)
            ]
            if missing:
                listed = "\n".join(f"  • {name}" for name in missing)
                proceed = messagebox.askyesno(
                    "Field not in every file",
                    f"“{link_header}” is not a column in:\n\n{listed}\n\n"
                    "The shared scrambled values will apply only where that column exists.\n\n"
                    "Proceed anyway?",
                )
                if not proceed:
                    return

        create_keyed = bool(self.keyed_var.get())
        jobs: List[Tuple[LoadedFile, str, Dict[str, ScrambleMode]]] = []
        for loaded in self.files:
            output_path = self._output_path_for(loaded)
            if self._output_is_locked(output_path):
                return
            if create_keyed and self._output_is_locked(self._keyed_output_path(output_path)):
                return
            header_modes: Dict[str, ScrambleMode] = {}
            for h in self._filtered_headers(loaded):
                var = loaded.mode_vars.get(h)
                header_modes[h] = var.get() if var is not None else "no scramble"  # type: ignore[assignment]
            jobs.append((loaded, output_path, header_modes))

        link_seed = secrets.token_bytes(32) if link_header is not None else None
        self.status_var.set("Running…")

        def worker() -> None:
            # PIPELINE STEP 4: Scramble each file. The shared field uses one seed
            # for the whole run. Every other column uses that file's own seed.
            try:
                parts = ["Done."]
                if link_header:
                    parts.append(f"\nShared field: {link_header}")
                warnings: List[str] = []
                saved_any = False
                for loaded, output_path, header_modes in jobs:
                    try:
                        res = scramble_tabular_file(
                            input_path=loaded.path,
                            output_scrambled_path=output_path,
                            header_modes=header_modes,
                            create_keyed_normal_version=create_keyed,
                            encoding=loaded.encoding,
                            password=loaded.password,
                            include_sheets=loaded.selected_sheets,
                            link_header=link_header,
                            link_mode=link_mode,
                            link_seed=link_seed,
                        )
                    except ShuffleValuesError as e:
                        err = e
                        path = loaded.path
                        prefix = ""
                        if saved_any:
                            prefix = "Files earlier in this run were already saved.\n\n"
                        self.after(0, lambda err=err, path=path, prefix=prefix: self._on_shuffle_values_error(err, path, prefix))
                        return
                    except Exception as e:
                        details = traceback.format_exc()
                        name = os.path.basename(loaded.path)
                        prefix = ""
                        if saved_any:
                            prefix = "Files earlier in this run were already saved.\n\n"
                        self.after(
                            0,
                            lambda name=name, err=e, details=details, prefix=prefix: self._on_error(
                                f"{prefix}{name}\n\n{err}\n\nDetails:\n{details}"
                            ),
                        )
                        return
                    saved_any = True
                    parts.append(f"\nScrambled file:\n{res.scrambled_output_path}")
                    if res.keyed_original_output_path:
                        parts.append(f"Keyed original file:\n{res.keyed_original_output_path}")
                    parts.append(f"Rows processed: {res.row_count}")
                    warnings.extend(res.warnings)

                if missing and link_header:
                    listed = "\n".join(f"  • {name}" for name in missing)
                    parts.append(
                        f"\nNo column “{link_header}” in:\n{listed}\n"
                        "Those files were scrambled without shared values."
                    )
                if warnings:
                    skipped = "\n".join(f"  • {w}" for w in warnings[:10])
                    if len(warnings) > 10:
                        skipped += f"\n  … and {len(warnings) - 10} more"
                    parts.append(
                        f"\nNote: {len(warnings)} conditional formatting rule(s) could not be shifted and were skipped:\n{skipped}"
                    )
                msg = "\n".join(parts)
                self.after(0, lambda msg=msg: self._on_success(msg))
            except Exception as e:
                details = traceback.format_exc()
                self.after(0, lambda err=e, details=details: self._on_error(f"{err}\n\nDetails:\n{details}"))

        threading.Thread(target=worker, daemon=True).start()

    def _on_shuffle_values_error(self, e: ShuffleValuesError, path: str, prefix: str = "") -> None:
        key = getattr(e, "column_key", None)
        widget = self.column_row_frames.get((path, key)) if key else None
        if widget is not None:
            try:
                self.scroll.scroll_row_into_view(widget)
            except Exception:
                pass
        self._on_error(f"{prefix}Shuffle values error in {os.path.basename(path)}:\n\n{e}")

    def _on_success(self, msg: str) -> None:
        self.status_var.set("Complete.")
        messagebox.showinfo("Scramble complete", msg)

    def _on_error(self, msg: str) -> None:
        self.status_var.set("Error.")
        messagebox.showerror("Error", msg)


def main() -> int:
    # PIPELINE STEP 0: Launch the app, then open the file picker.
    app = App()
    app.after(200, app.choose_files)
    app.mainloop()
    return 0


# COMMON EDITS (for maintainers)
# - Change the version shown to staff: update APP_VERSION above.
# - Change which file types are allowed: update choose_files() filetypes + scrambler_core.scramble_tabular_file().
# - Change the default dropdown selections: edit the "no scramble" default in _render_column_row().


if __name__ == "__main__":
    raise SystemExit(main())

