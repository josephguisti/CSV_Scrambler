# CSV Scrambler (PII-safe for AI tools)

This folder contains small Python scripts to **scramble PII** in CSV files before importing them into AI-enabled tools.

## Scripts

### `scramble_csv.py`
Scrambles **letters and digits** in all columns, with two special behaviors:

- **Dates**: any date-like substring stays in the **same month and year**, but the **day is scrambled**.
  - Supports formats like `M/D/YYYY` and `YYYY-MM-DD` (including when directly followed by time, e.g. `2026-03-0310:40AM`).
  - If a cell contains time, the **time portion is preserved** (only the date is changed).
- **DOB-like columns**: for columns whose header contains `age`, `birth`, or `dob` (case-insensitive),
  the scrambled day is chosen so the person remains in the same bucket **as-of today**:
  - minor: < 18
  - TAY: 18–24 inclusive
  - adult: 25+
- **Program/Project columns stay stable**: if a column header contains `Program` **or** `Project` but does **not** contain `Date`,
  the entire column is left unchanged.

Usage:

```bash
python scramble_csv.py --input "Critical Incident Report Pre-Scramble.csv" --output "Critical Incident Report Scrambled.csv"
```

Easiest (opens file picker windows if you omit `--input` and `--output`):

```bash
python scramble_csv.py
```

Optional: provide a seed for repeatable output:

```bash
python scramble_csv.py --input in.csv --output out.csv --seed "my-fixed-seed"
```

### `csv_scrambler_app.py` (GUI)
This is the **GUI app** version (no Terminal needed once packaged).

To run from Python (for developers):

```bash
python csv_scrambler_app.py
```

To build a portable Windows `.exe`:
- Run `build_csv_scrambler_exe.bat`
- The output will be `dist\\CSV Scrambler.exe`

Excel support:
- The GUI app supports **`.xlsx`** with **multiple worksheet tabs**: row 1 on each tab is headers; column rules apply by header name across all tabs.
- **Shuffle values** runs **within each sheet** separately.
- Chart-only tabs are omitted from input/output (openpyxl).
- Output is **values-only** and filters are removed.

### `age_bucket_report.py` (accessory)
Reports minor/TAY/adult counts for DOB-like columns (headers containing `age`, `birth`, or `dob`), **as-of today**.
You can also compare two CSVs (e.g., original vs scrambled) to confirm counts match.

```bash
python age_bucket_report.py --input "Critical Incident Report Pre-Scramble.csv"
python age_bucket_report.py --input "Critical Incident Report Pre-Scramble.csv" --compare "Critical Incident Report Scrambled.csv"
```

### `format_critical_incidents_for_email.py`
Formats an exported Microsoft List CSV (Critical Incident Reports) into **email-ready text**.
It uses the header row from `Critical Incident Report Pre-Scramble.csv` as the **canonical schema and field order**.

Default behavior:
- Outputs one incident per row
- Prints each field as:
  - `COLUMN HEADER:` then the value on the next line
- Inserts `----------` between incidents
- Handles quoted multi-line cells (keeps newlines; indents continuation lines)

Non-technical usage (recommended): run it and use the file picker + Save As dialogs:

```bash
python format_critical_incidents_for_email.py
```

Power-user usage (explicit paths):

```bash
python format_critical_incidents_for_email.py --input "Critical Incident Report Pre-Scramble.csv" --output "Critical Incident Email Output.txt"
```

Optional flags:
- `--skip-empty-fields`: omit blank fields
- `--strict-schema`: error if the CSV is missing expected columns
- `--include-extras`: include extra columns not in the canonical schema
- `--stdout`: print output to the terminal instead of saving a .txt file

## Notes
- The scripts use only the Python standard library (no dependencies).
- `scramble_csv.py` prints the **seed** it used so you can reproduce the same scramble later.

