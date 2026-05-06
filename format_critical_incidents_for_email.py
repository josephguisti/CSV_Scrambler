import argparse
import csv
import os
import sys
from typing import Dict, Iterable, List, Optional, Tuple


# Canonical schema/order taken from `Critical Incident Report Pre-Scramble.csv`
SCHEMA_HEADERS: List[str] = [
    "Status",
    "Date Submitted",
    "Days Since Submission",
    "Assigned to",
    "Notes",
    "Resolution",
    "Client DOB",
    "Submitter's Name",
    "Submitter's Contact",
    "Supervisor",
    "Type of Incident",
    "Location of Incident",
    "Date & Time of Incident",
    "People Involved",
    "Police Involvement",
    "Witnesses",
    "Programs Involved",
    "Incident Description",
    "Actions Taken",
    "Submission Time",
]


def _norm(s: str) -> str:
    # Normalization for header matching: keep it conservative (case-insensitive + trim).
    # Also strip UTF-8 BOM and common zero-width chars that can sneak into CSV exports.
    # (Excel/Microsoft exports sometimes include these on the first header cell.)
    s = (s or "").replace("\u200b", "").replace("\u200c", "").replace("\u200d", "")
    s = s.lstrip("\ufeff")
    return s.strip().casefold()


def _schema_compare(actual_headers: List[str]) -> Tuple[List[str], List[str]]:
    # Returns (missing_from_csv, extra_in_csv) relative to SCHEMA_HEADERS.
    actual_norm = {_norm(h): h for h in actual_headers}
    schema_norm = [_norm(h) for h in SCHEMA_HEADERS]

    missing: List[str] = []
    for h in SCHEMA_HEADERS:
        if _norm(h) not in actual_norm:
            missing.append(h)

    schema_norm_set = set(schema_norm)
    extra: List[str] = []
    for h in actual_headers:
        if _norm(h) not in schema_norm_set:
            extra.append(h)

    return missing, extra


def _ordered_headers(actual_headers: List[str], include_extras: bool) -> List[Tuple[str, str]]:
    # Use schema order first; optionally append extra columns in their CSV order.
    #
    # Returns list of (display_header, lookup_key) so we can print clean canonical
    # headers even if the CSV has a BOM/hidden chars in the original header text.
    actual_by_norm: Dict[str, str] = {_norm(h): h for h in actual_headers}
    ordered: List[Tuple[str, str]] = []
    for schema_h in SCHEMA_HEADERS:
        k = _norm(schema_h)
        if k in actual_by_norm:
            ordered.append((schema_h, actual_by_norm[k]))
    if include_extras:
        schema_norm_set = {_norm(h) for h in SCHEMA_HEADERS}
        for h in actual_headers:
            if _norm(h) not in schema_norm_set:
                ordered.append((h, h))
    return ordered


def _format_value(v: str, indent: str) -> str:
    # Keep the original text, but make sure multiline cells remain readable.
    if v is None:
        v = ""
    v = str(v)
    v = v.replace("\r\n", "\n").replace("\r", "\n")
    if not v.strip():
        return ""
    lines = v.split("\n")
    if len(lines) == 1:
        return lines[0].strip()
    # For multiline, keep internal newlines and indent continuation lines.
    first = lines[0].rstrip()
    rest = [(indent + ln.rstrip()) if ln.strip() else "" for ln in lines[1:]]
    return "\n".join([first, *rest]).strip("\n")


def format_rows(
    *,
    reader: csv.DictReader,
    include_extras: bool,
    skip_empty_fields: bool,
    incident_separator: str,
    include_incident_header: bool,
    incident_header_template: str,
    indent_multiline: str,
) -> str:
    if reader.fieldnames is None:
        return ""

    actual_headers = [h for h in reader.fieldnames if h is not None]
    ordered_headers = _ordered_headers(actual_headers, include_extras=include_extras)

    blocks: List[str] = []
    for idx, row in enumerate(reader, start=1):
        lines: List[str] = []
        if include_incident_header:
            lines.append(incident_header_template.format(n=idx))
            lines.append("")  # blank line after title

        for display_h, lookup_h in ordered_headers:
            raw = row.get(lookup_h, "")
            val = _format_value(raw, indent=indent_multiline)
            if skip_empty_fields and not val.strip():
                continue
            lines.append(f"{display_h}:")
            lines.append(val)
            lines.append("")  # blank line between fields

        # Trim trailing whitespace-only lines.
        while lines and not lines[-1].strip():
            lines.pop()

        blocks.append("\n".join(lines))

    return f"\n{incident_separator}\n".join(blocks).rstrip() + "\n"


def main(argv: Optional[List[str]] = None) -> int:
    p = argparse.ArgumentParser(
        description="Format exported Microsoft List CSV rows into email-ready text blocks using the Critical Incident schema/order."
    )
    p.add_argument(
        "--input",
        default=None,
        help="Path to input CSV (export from Microsoft List). If omitted, a file picker will open.",
    )
    p.add_argument(
        "--output",
        default=None,
        help='Path to write output text file. If omitted, a "Save As" dialog will open (unless --stdout).',
    )
    p.add_argument(
        "--stdout",
        action="store_true",
        help="Print formatted output to stdout instead of creating a .txt file (no Save As dialog).",
    )
    p.add_argument("--encoding", default="utf-8", help="File encoding (default: utf-8)")
    p.add_argument(
        "--strict-schema",
        action="store_true",
        help="Fail if the CSV is missing any columns from the canonical schema.",
    )
    p.add_argument(
        "--include-extras",
        action="store_true",
        help="Include columns not in the canonical schema (appended after schema fields, in CSV order).",
    )
    p.add_argument(
        "--skip-empty-fields",
        action="store_true",
        help="Omit fields whose value is blank/whitespace.",
    )
    p.add_argument(
        "--incident-separator",
        default="----------",
        help='Text inserted between incidents (default: "----------")',
    )
    p.add_argument(
        "--no-incident-header",
        action="store_true",
        help='Do not include a per-incident title line like "INCIDENT 1".',
    )
    p.add_argument(
        "--incident-header-template",
        default="INCIDENT {n}",
        help='Template for per-incident title line. Use "{n}" for incident number (default: "INCIDENT {n}")',
    )
    p.add_argument(
        "--multiline-indent",
        default="  ",
        help="Indent used for continuation lines when a cell contains newlines (default: two spaces).",
    )
    args = p.parse_args(argv)

    def _pick_paths_interactively(
        *,
        provided_input: Optional[str],
        provided_output: Optional[str],
        use_stdout: bool,
    ) -> Tuple[Optional[str], Optional[str]]:
        """
        Returns (input_csv_path, output_txt_path). If use_stdout is True, output may be None.

        Uses tkinter file dialogs when available; falls back to console prompts if not.
        """
        in_path_local = provided_input
        out_path_local = provided_output

        if in_path_local and (use_stdout or out_path_local):
            return in_path_local, out_path_local

        try:
            import tkinter as tk  # stdlib
            from tkinter import filedialog

            root = tk.Tk()
            root.withdraw()
            try:
                root.attributes("-topmost", True)
            except Exception:
                # Not critical; continue without forcing topmost.
                pass

            if not in_path_local:
                in_path_local = filedialog.askopenfilename(
                    title="Select the Microsoft List CSV export",
                    filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                )
                if not in_path_local:
                    return None, None

            if not use_stdout and not out_path_local:
                default_name = "Critical Incident Email Output.txt"
                initial_dir = os.path.dirname(os.path.abspath(in_path_local)) if in_path_local else None
                out_path_local = filedialog.asksaveasfilename(
                    title="Save the email-ready text file as...",
                    defaultextension=".txt",
                    initialfile=default_name,
                    initialdir=initial_dir,
                    filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
                )
                if not out_path_local:
                    return None, None

            return in_path_local, out_path_local
        except Exception:
            # Fallback: console prompts (works even if tkinter isn't available).
            if not in_path_local:
                sys.stdout.write("Enter the path to the CSV export file:\n> ")
                in_path_local = (sys.stdin.readline() or "").strip().strip('"')
                if not in_path_local:
                    return None, None
            if not use_stdout and not out_path_local:
                sys.stdout.write('Enter the full path + filename for the output .txt (example: "C:\\Temp\\Critical Incident Email Output.txt"):\n> ')
                out_path_local = (sys.stdin.readline() or "").strip().strip('"')
                if not out_path_local:
                    return None, None
            return in_path_local, out_path_local

    in_path, out_path = _pick_paths_interactively(
        provided_input=args.input,
        provided_output=args.output,
        use_stdout=bool(args.stdout),
    )

    if not in_path:
        sys.stderr.write("No input CSV selected. Exiting.\n")
        return 2

    with open(in_path, "r", encoding=args.encoding, newline="") as f:
        reader = csv.DictReader(f)
        actual_headers = [h for h in (reader.fieldnames or []) if h is not None]
        missing, extra = _schema_compare(actual_headers)

        if missing:
            msg = (
                "Input CSV is missing one or more expected columns from the canonical schema:\n"
                + "\n".join([f"- {h}" for h in missing])
                + "\n"
            )
            if args.strict_schema:
                sys.stderr.write(msg)
                return 2
            # Non-strict: warn, then proceed with whatever columns are present.
            sys.stderr.write("WARNING: " + msg)

        if extra and not args.include_extras:
            sys.stderr.write(
                "NOTE: Input CSV contains extra columns not in the canonical schema. "
                "Run with --include-extras to include them in output.\n"
            )

        text = format_rows(
            reader=reader,
            include_extras=bool(args.include_extras),
            skip_empty_fields=bool(args.skip_empty_fields),
            incident_separator=str(args.incident_separator),
            include_incident_header=not bool(args.no_incident_header),
            incident_header_template=str(args.incident_header_template),
            indent_multiline=str(args.multiline_indent),
        )

    if args.stdout:
        sys.stdout.write(text)
        return 0

    if not out_path:
        sys.stderr.write("No output path selected. Exiting.\n")
        return 2

    if out_path:
        os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
        with open(out_path, "w", encoding=args.encoding, newline="\n") as f_out:
            f_out.write(text)
        sys.stdout.write(f"Format complete.\nOutput: {out_path}\n")
        return 0

    return 0


if __name__ == "__main__":
    raise SystemExit(main())

