# Service Manual — CSV/Excel Scrambler (Portable Windows App)

This document is written for a **non-coder** who may inherit this project. It explains:
- what the tool is for (scope),
- what it outputs,
- the **business rules** baked into it,
- how to **build** the portable `.exe`,
- and what to check first when something goes wrong.

---

## What this tool does (scope)

Staff often need to share spreadsheet exports with AI-enabled tools while preventing disclosure of **PII**. This tool:
- accepts a `.csv` or `.xlsx`,
- lets the user choose **per-column** handling rules,
- produces a **scrubbed/scrambled** output file,
- and can optionally produce a **keyed original** copy for recovery workflows.

This tool is intentionally built as a **single portable `.exe`** that can be copied and run without Terminal.

---

## Where the important files are

- **GUI app entrypoint**: `csv_scrambler_app.py`
- **Core business logic** (scrambling, dates, shuffle, keyed output): `scrambler_core.py`
- **Legacy CLI scrambler** (still works, uses core): `scramble_csv.py`
- **Accessory validation tool**: `age_bucket_report.py`
- **Build script**: `build_csv_scrambler_exe.bat`
- **Runtime dependency pins**: `requirements.txt`
- **Build-time dependency pins**: `requirements-dev.txt`
- **End-user app guide**: `HOW_TO_USE_CSV_SCRAMBLER_APP_WINDOWS.md`
- **Terminal-based guide (older flow)**: `HOW_TO_SCRAMBLE_CSV_WINDOWS.md`

The portable executable is created at:
- `dist\\CSV Scrambler.exe`

---

## Design decisions (why it’s built this way)

- **Tkinter GUI**: uses Python’s built-in GUI toolkit to keep the project simple and easy to package.
- **One-file EXE**: uses PyInstaller `--onefile` so staff can run it without Python/Terminal.
- **No seed exposure in GUI**: the GUI intentionally does **not** show or accept a seed (results are not meant to be reproducible).  
  (The older CLI script still supports `--seed` for reproducibility.)
- **CSV + XLSX support**:
  - CSV processing uses Python’s `csv` module.
  - Excel processing uses `openpyxl` (pinned in `requirements.txt`).
- **Excel output is values-only**: we write a brand-new workbook so filters/formatting/formulas are not preserved (by design).
- **Per-column dropdown modes**: the GUI exposes the rules directly so staff can keep particular columns stable for error-checking.

---

## Architecture overview (plain language pipeline)

This app is organized as a simple pipeline. If you get stuck, open the files below and search for the comment banners `PIPELINE STEP` and `BUSINESS RULE`.

### GUI pipeline (`csv_scrambler_app.py`)

- **PIPELINE STEP 0**: build the UI (buttons, header list, dropdowns)
- **PIPELINE STEP 1**: user chooses input file + output location
- **PIPELINE STEP 2**: app reads headers and shows grouped columns with dropdowns (defaults come from core)
- **PIPELINE STEP 3**: user clicks Run; app collects dropdown selections
- **PIPELINE STEP 4**: app calls the core pipeline and then shows success/error messages

Where to look:
- version label: `APP_VERSION`
- file picking + header scan: `_try_read_headers()` and `choose_input()`
- run action: `run()` (calls `scrambler_core.scramble_tabular_file()`)

### Core pipeline (`scrambler_core.py`)

- **PIPELINE STEP 0**: route by file extension (`scramble_tabular_file`)
- **CSV path** (`scramble_csv_file`)
  - **PIPELINE STEP 1**: first pass collects shuffle info + generates row IDs
  - **PIPELINE STEP 2**: compute “shuffle values” reassignment
  - **PIPELINE STEP 3**: second pass writes scrambled output + optional keyed original
- **Excel path** (`scramble_xlsx_file`)
  - **PIPELINE STEP 1**: load workbook + read header row on **each worksheet** (`wb.worksheets`; chart-only tabs are skipped by openpyxl)
  - **PIPELINE STEP 2**: collect shuffle info **per sheet** (non-blank only)
  - **PIPELINE STEP 3**: write brand-new values-only workbook with **one tab per worksheet** (filters removed)

Where the rules live:
- dropdown semantics: `scramble_cell()`
- default dropdown selection: `default_mode_for_header()`
- date parsing regexes: `ISO_DATE_RE`, `US_DATE_RE`, `TIME_RE`

---

## Business rules (critical)

These are the rules implemented in `scrambler_core.py`.

### Column handling modes (dropdown)

Each column is set to one of:

- **Full scramble**
  - Scrambles **letters and digits** everywhere in that column
  - **No date special handling**: dates/times are scrambled like normal text

- **Date scramble**
  - Scrambles letters and digits, **but** date-like substrings are protected so:
    - dates keep the **same month and year**; the **day changes**
    - time-of-day substrings (like `10:40AM`, `22:12`) are preserved
  - For DOB-like columns (see below), day scrambling is chosen to preserve bucket (minor/TAY/adult) **as-of today**

- **Shuffle values**
  - Designed for picklist/dropdown columns (e.g., Race/Ethnicity).
  - Blanks stay blank.
  - Non-blank values are reassigned across rows so **no row keeps its original value**.
  - Error cases:
    - if there are fewer than 2 non-blank rows
    - if all non-blank rows have the same value
    - if one value appears too frequently to guarantee every row changes (tool will show an error naming the column)

- **No scramble**
  - Leaves the column unchanged.

- **Delete data**
  - Blanks out the cells (column/header stays).

### What “scramble” means (text rules)

When scrambling letters/digits:
- **Letters and digits only** are changed.
- **Punctuation/spacing/newlines** are preserved exactly.
- **Length is preserved**.
- **Case is preserved** (uppercase stays uppercase).
- Hyphens remain in the same place (because punctuation is preserved).

### Program/Project stability rule (default behavior)

Default dropdown selection uses this rule:
- If a header contains **`Program`** or **`Project`** and does **not** contain `Date`, default is **No scramble**.

(Users can still override the dropdown.)

### DOB/age bucket preservation rule

DOB-like headers are detected by header text containing (case-insensitive):
- `age`
- `birth`
- `dob`

For those columns, **Date scramble** keeps the person in the same bucket **as-of today’s date**:
- minor: < 18
- TAY: 18–24 inclusive
- adult: 25+

### Date parsing rules (what counts as a date)

The tool recognizes date-like substrings in these formats:
- `M/D/YYYY` (example: `3/3/2026`)
- `YYYY-MM-DD` (example: `2026-03-03`)

Time-of-day recognition examples:
- `10:40AM`, `3:00pm`, `19:06`, `22:12`

Anything outside these patterns may be treated as normal text.

### Keyed normal version (checkbox)

If **Create Keyed Normal Version** is checked:
- the tool outputs **two files**:
  - the scrambled file
  - a keyed original (“normal”) file
- both files include a `Row ID` column (UUID per row) so they can be joined later.
- the keyed original is **unchanged** except for adding `Row ID` (Option A).

---

## Excel (XLSX) support rules

Excel files must be formatted like this:
- **Row 1** on **every worksheet tab** contains headers (same rule as CSV: no duplicate header text within a single sheet)
- The app builds a **union** of all column names across sheets; the column rules you set apply to **every** tab that has that column
- **Shuffle values** is computed **within each sheet** (not pooled across the whole file)
- Chart-only tabs (not grid worksheets) are not included in `wb.worksheets` and are omitted from the output
- Avoid merged cells / tables (filters are OK)

Output is:
- a new `.xlsx` with **values only** and the **same worksheet tabs** as the input (data sheets only)
- filters are effectively removed because the output workbook is newly created

---

## Build process (portable EXE)

This project builds the EXE using PyInstaller.

### Prerequisites

- Windows PC
- Python 3 installed
- Internet access (first time only, to install PyInstaller/openpyxl)

### Build steps

1. Open File Explorer and go to this project folder.
2. Double-click `build_csv_scrambler_exe.bat`
3. Wait until it finishes.
4. The EXE will be at:
   - `dist\\CSV Scrambler.exe`

### What “packaging” does (plain language)

The build script:
- installs pinned dependencies from `requirements.txt` (runtime) and `requirements-dev.txt` (build tools)
- runs PyInstaller to bundle:
  - `csv_scrambler_app.py` (the GUI entrypoint)
  - `scrambler_core.py` (the rule/pipeline engine)
  - Python + Tkinter + `openpyxl` into one portable `.exe`

### Why we don’t use `--clean`

The build script intentionally avoids PyInstaller `--clean` because OneDrive/antivirus can lock files in `build\\...` and cause “Access denied” errors during cleanup.

If the build gets stuck due to locks:
- Close Excel and any open files in this folder
- Wait for OneDrive sync to finish
- Re-run `build_csv_scrambler_exe.bat`

---

## Troubleshooting guide (non-coder friendly)

### The app won’t open
- If Windows SmartScreen blocks it: **More info → Run anyway**
- If your organization blocks EXEs from email: distribute via approved file share (OneDrive/SharePoint/Teams) or a password-protected zip.

### “Unsupported file type”
- The app supports `.csv` and `.xlsx`.

### “Duplicate headers not supported”
- On a CSV: two columns share the same header name.
- On Excel: the message may name the **sheet** (e.g. `Sheet 'Clients'`) when two columns on that tab share a name.
- Fix by making each column header unique **on that sheet**.

### “Shuffle values error”
- The message names the **sheet** and **column** when the input is Excel (shuffle is evaluated per sheet).
- Common fixes:
  - Don’t use Shuffle on a column where all rows are the same **on that sheet**, or where only one row has a value
  - Or use Full/Date scramble instead

### Dates didn’t scramble as expected
- The tool only detects `M/D/YYYY` and `YYYY-MM-DD`.
- If your date format is different (like `03-03-2026`), it may scramble as normal text.

---

## Where to change things (common enhancements)

### Add another dropdown mode
- Update the `ScrambleMode` type in `scrambler_core.py`
- Add implementation in `scramble_cell()` or the file-level pipelines
- Add it to `MODES` in `csv_scrambler_app.py`
- Rebuild the EXE

### Change default dropdowns
- Edit `default_mode_for_header()` in `scrambler_core.py`

### Add more date formats
- Add regex patterns near `ISO_DATE_RE` / `US_DATE_RE` in `scrambler_core.py`
- Update placeholder replacement logic in `_replace_protected_substrings_with_placeholders()`

### Change Excel sheet/header assumptions
- Multi-sheet behavior is implemented in `scramble_xlsx_file()` in `scrambler_core.py`; GUI header union lives in `_try_read_headers()` in `csv_scrambler_app.py`.

