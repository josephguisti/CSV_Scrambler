# CSV / Excel Scrambler

Privacy-oriented Windows tool for creating PII-safe copies of CSV and Excel files before sharing data with AI-enabled tools. The desktop app reads a source file, lets the user choose a rule for each column, and writes new output files without editing the original.

Current desktop app version: **1.13.0**

## Recommended Download / Use

For non-technical users, use the packaged Windows app:

- App file: `dist\CSV Scrambler.exe`
- End-user guide: `HOW_TO_USE_CSV_SCRAMBLER_APP_WINDOWS.md`
- Maintainer/service guide: `SERVICE_MANUAL.md`

To run from source:

```bash
python csv_scrambler_app.py
```

To rebuild the portable Windows `.exe`:

```bash
build_csv_scrambler_exe.bat
```

The build output is `dist\CSV Scrambler.exe`.

## What The Desktop App Supports

- Input files: `.csv` and `.xlsx`, including several files in one run (Ctrl+click in the file window that opens at startup)
- One radio-selected column keeps the same scrambled values across every selected file and sheet. Other columns are scrambled separately in each file
- Password-protected `.xlsx` files, using in-memory decryption
- Multi-sheet Excel workbooks with a sheet picker before export
- Header-row detection for Excel using AutoFilter first, then a density scan
- Per-column rules with sample values shown in the UI
- Optional keyed original output with matching `Row ID` values
- Static Excel cell styling copied where possible
- Conditional formatting copied where possible, with warnings for rules that cannot be shifted safely
- Pre-flight warning when the selected output file is already open or locked

The app always writes new output files. It does not modify the source file in place.

## Column Rules

Every column starts as **No scramble** in the desktop app. Users opt in to one of these modes:

- **Full scramble**: scrambles letters and digits everywhere in the cell, including dates and times.
- **Date scramble**: preserves recognized date/time structure while changing date days and scrambling other letters/digits.
- **Shuffle values**: reassigns non-blank values within the same column so each non-blank row gets a different value.
- **No scramble**: leaves values unchanged.
- **Delete data**: blanks cell values while keeping the column.

Date scramble recognizes `M/D/YYYY` and `YYYY-MM-DD`. Time-of-day strings such as `10:40AM` or `22:12` are preserved. Date scrambling keeps the same month and year and preserves minor/TAY/adult age buckets as of the run date.

## Excel Notes

- Excel output is values-first: formulas are written as their loaded values, not preserved as formulas.
- For multi-sheet workbooks, rules are keyed as `Column Name (Sheet Name)` so the same column name on different sheets can have different rules.
- The radio-selected header is the exception: that column uses one scramble everywhere it appears. The match is the exact header text.
- Shuffle values runs within each sheet, not across the entire workbook.
- Chart-only tabs are skipped because they are not grid worksheets.
- If exporting more than one worksheet, choose `.xlsx` output. CSV output is only valid when exporting a single sheet.
- Password-protected input can be opened, but output files are not re-encrypted.

## Developer / Script Entry Points

### `csv_scrambler_app.py`

Tkinter desktop app entry point. This is what gets packaged into the Windows `.exe`.

### `scrambler_core.py`

Shared business logic for CSV/Excel reading, column modes, date handling, shuffle handling, keyed output, and Excel formatting behavior.

### `scramble_csv.py`

CSV-only command-line tool. Unlike the desktop app, it uses automatic defaults:

- Program/Project columns stay unchanged unless the header also contains Date.
- Date-like columns use date scramble.
- Other columns use full scramble.

Usage:

```bash
python scramble_csv.py --input "input.csv" --output "output.csv"
```

Optional deterministic seed:

```bash
python scramble_csv.py --input "input.csv" --output "output.csv" --seed "my-fixed-seed"
```

### `age_bucket_report.py`

Reports minor/TAY/adult counts for DOB-like CSV columns and can compare an original CSV to a scrambled CSV.

```bash
python age_bucket_report.py --input "input.csv" --compare "output.csv"
```

### `format_critical_incidents_for_email.py`

Formats an exported Microsoft List CSV into email-ready text for Critical Incident workflows.

```bash
python format_critical_incidents_for_email.py --input "Critical Incident Report Pre-Scramble.csv" --output "Critical Incident Email Output.txt"
```

Useful flags:

- `--skip-empty-fields`: omit blank fields
- `--strict-schema`: error if the CSV is missing expected columns
- `--include-extras`: include extra columns not in the canonical schema
- `--stdout`: print output to the terminal instead of saving a `.txt` file

## Dependencies

Runtime dependencies are in `requirements.txt`:

- `openpyxl`
- `msoffcrypto-tool`

Build dependencies are in `requirements-dev.txt`:

- `pyinstaller`
