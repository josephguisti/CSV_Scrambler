# CSV / Excel Scrambler — capability summary

**Purpose.** Privacy-oriented tool that reads a tabular **CSV** or **Excel (.xlsx/.xlsm)** file and writes **new** output files where cell values are transformed per column according to user-selected rules. It does **not** edit files in place.

**Architecture.** **`scrambler_core.py`** holds all transformation rules and file pipelines. **`csv_scrambler_app.py`** is a **Tkinter** GUI (Windows-friendly, packagable as `.exe`). A separate **`scramble_csv.py`** CLI exists for CSV-only batch use with different defaults.

---

## Inputs and constraints

- **Headers:** Column names must be **unique** per file (duplicate headers are rejected).
- **Excel:** Uses **openpyxl**. **Password-protected** workbooks require **`msoffcrypto-tool`** (decrypt in memory; output is not encrypted). **Chart-only** tabs are invisible to openpyxl and are skipped like today.
- **Header row detection (Excel):** Not always row 1—logic uses AutoFilter row / density scan (see `_xlsx_header_row`).

---

## Per-column modes (`ScrambleMode`)

| Mode | Behavior (high level) |
|------|-------------------------|
| **no scramble** | Copy value unchanged. |
| **delete data** | Blank the cell. |
| **full scramble** | HMAC-seeded substitution of letters/digits within alphanumeric runs; **same input string → same output** every time in one run (no row-specific randomness). |
| **date scramble** | Recognizes ISO/US date substrings and applies bucket-preserving DOB-style rules + month/year constraints; time-of-day preserved; remaining alnums go through full-style scrambling. |
| **shuffle values** | Keeps the **multiset** of non-blank values in the column but **reassigns** them across rows (within sheet / column scope); fails if reassignment is impossible (e.g. all rows identical). |

**GUI default:** Every column starts as **no scramble** until the user changes dropdowns. **Core/CLI default** differs (program/project stability, date detection → date scramble, else full scramble)—see `default_mode_for_column`.

**Special header rule:** Columns whose names match **program/project** (and not “date”) are classified as stable identifiers for grouping in the UI (`is_program_project_stable_header`).

---

## Outputs

- **Primary:** Scrambled CSV or Excel path chosen by the user.
- **Optional “keyed original”:** Second file with **same Row IDs** as scrambled output but **original** cell values (for linking scrambled ↔ source rows).

**Row IDs:** UUIDs per data row, prepended as a **Row ID** column.

---

## CSV pipeline (`scramble_csv_file`)

Two-pass: collect shuffle metadata + generate IDs; apply modes row-by-row.

---

## Excel pipeline (`scramble_xlsx_file`)

- Processes **each worksheet** (or **`include_sheets`** subset).
- Writes a **new workbook**: values per mode; **formulas not preserved** as formulas (values written).
- **Styles:** Static cell styling copied where possible; **conditional formatting** copied with coordinate shifting (+Row ID column, header-row offset). Rules using INDIRECT, external refs, or some edge cases may be skipped with warnings.
- **Cross-sheet CF:** Optional pre-export check warns when CF on an **included** sheet references an **excluded** sheet (`detect_cross_sheet_cf_breaks`).

---

## GUI-specific behavior

- **Date detection:** Samples early rows per column to suggest date-capable columns in the UI.
- **Multi-sheet Excel:** **Sheet picker** before run; **`selected_sheets`** filters which tabs export and which column keys appear in the rule list. Rules can be grouped **by worksheet** in the UI when the workbook has multiple tabs.
- **Column keys for Excel:** `"{header} ({sheet_title})"` so the same header name on two sheets is two distinct rules.

---

## Dependencies (typical)

- **CSV:** stdlib only for core path.
- **Excel:** `openpyxl`; encrypted files: `msoffcrypto-tool`.

---

## Integration hints for another codebase

If you port or embed this behavior elsewhere, treat these as **separable modules:**

1. **Cell-level transforms** — `scramble_cell`, `scramble_alnum_runs`, date placeholder pipeline, shuffle assignment logic.
2. **File orchestration** — CSV vs XLSX readers/writers, Row ID injection, optional second “keyed” output.
3. **Excel-only** — worksheet iteration, `include_sheets`, style + conditional formatting copy, header-row detection.
4. **UI layer** — file pickers, per-column mode storage, sheet picker + CF warning flow (optional outside a headless service).
