import datetime as dt
import hashlib
import os
import secrets
import sys
import argparse
import csv
from typing import List, Optional

from scrambler_core import (
    ColumnContext,
    ScrambleMode,
    default_mode_for_header,
    is_program_project_stable_header,
    scramble_cell,
)


def _seed_bytes_from_seed_string(seed_str: str) -> bytes:
    # Keep CLI deterministic behavior.
    return hashlib.sha256(seed_str.encode("utf-8")).digest()


def main(argv: Optional[List[str]] = None) -> int:
    p = argparse.ArgumentParser(
        description="Scramble PII in a CSV while keeping dates in the same month/year, and preserving minor/TAY/adult bucket for DOB-like columns (as-of today)."
    )
    p.add_argument("--input", default=None, help="Path to input CSV. If omitted, a file picker opens.")
    p.add_argument("--output", default=None, help="Path to output (scrubbed) CSV. If omitted, a save dialog opens.")
    p.add_argument(
        "--seed",
        default=None,
        help="Seed string for deterministic scrambling. If omitted, a random seed is generated and printed.",
    )
    p.add_argument(
        "--as-of",
        default=None,
        help="Override 'today' date for bucket preservation, in YYYY-MM-DD. Default: system today.",
    )
    p.add_argument("--encoding", default="utf-8", help="File encoding (default: utf-8)")
    args = p.parse_args(argv)

    as_of = dt.date.today() if not args.as_of else dt.date.fromisoformat(args.as_of)

    seed_str = args.seed or secrets.token_urlsafe(24)
    seed = _seed_bytes_from_seed_string(seed_str)

    in_path = args.input
    out_path = args.output

    # If paths weren't provided, use native file chooser dialogs (Windows-friendly).
    if not in_path or not out_path:
        try:
            import tkinter as tk
            from tkinter import filedialog
        except Exception:
            sys.stderr.write(
                "ERROR: No --input/--output provided, but the file picker is unavailable.\n"
                "Please run again with --input and --output paths.\n"
            )
            return 2

        root = tk.Tk()
        root.withdraw()
        root.update()

        if not in_path:
            in_path = filedialog.askopenfilename(
                title="Choose the CSV file to scramble",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            )
            if not in_path:
                return 1

        if not out_path:
            default_name = os.path.splitext(os.path.basename(in_path))[0] + " Scrambled.csv"
            out_path = filedialog.asksaveasfilename(
                title="Choose where to save the scrambled CSV",
                defaultextension=".csv",
                initialdir=os.path.dirname(os.path.abspath(in_path)),
                initialfile=default_name,
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            )
            if not out_path:
                return 1

        try:
            root.destroy()
        except Exception:
            pass

    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)

    with open(in_path, "r", encoding=args.encoding, newline="") as f_in, open(
        out_path, "w", encoding=args.encoding, newline=""
    ) as f_out:
        reader = csv.reader(f_in)
        writer = csv.writer(f_out)

        try:
            headers = next(reader)
        except StopIteration:
            return 0

        writer.writerow(headers)

        # CLI behavior: keep your original business rules as defaults.
        # - Program/Project (unless Date) stays stable
        # - DOB-like uses date rules
        # - Other date/time uses date rules
        header_modes = {h: default_mode_for_header(h) for h in headers}

        for row in reader:
            if len(row) < len(headers):
                row = row + [""] * (len(headers) - len(row))
            new_row: List[str] = []
            for idx, header in enumerate(headers):
                mode: ScrambleMode = header_modes[header]
                # Preserve historical CLI rule: stable program/project columns remain unchanged.
                if is_program_project_stable_header(header):
                    new_row.append(row[idx])
                    continue
                col = ColumnContext(header=header, mode=mode)
                new_row.append(scramble_cell(seed=seed, value=row[idx], col=col, as_of=as_of))
            writer.writerow(new_row)

    # Print seed so the user can reproduce the exact same scramble later.
    sys.stdout.write(f"Scramble complete.\nSeed: {seed_str}\nAs-of date: {as_of.isoformat()}\nOutput: {out_path}\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

