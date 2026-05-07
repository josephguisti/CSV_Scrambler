# Plan: Preserve Conditional Formatting in Output

## Overview

The scrambler currently builds output workbooks from scratch using `openpyxl.Workbook()`,
writing only cell values. All conditional formatting from the source file is discarded.
This plan adds logic to copy, adjust, and re-register conditional formatting rules from
each source worksheet onto its corresponding output worksheet.

---

## Background: Why coordinates need adjustment

Every output sheet has two layout differences vs. the source:

- **Column +1 shift** — a "Row ID" column is always prepended at column A, pushing all
  data columns right by one (source column A → output column B, etc.).
- **Row shift (conditional)** — when `_xlsx_header_row()` returns `2` (section-label row
  above headers), the output normalizes headers to row 1. Source data rows that start at
  row 3 map to output row 2 — a shift of `1 − hdr_row`.

Any `sqref` range and any formula string inside a rule must have its column letters and
row numbers adjusted to match the new layout.

---

## Files changed

- `scrambler_core.py` — all new logic; no other file needs to change.

---

## Implementation steps

### 1. Build a per-sheet `hdr_row` lookup before the main loop

```python
hdr_row_by_sheet: Dict[str, int] = {ws.title: _xlsx_header_row(ws) for ws in wb.worksheets}
```

This is used when rewriting cross-sheet row references inside formulas.

---

### 2. Add `_shift_sqref(sqref, col_delta, row_delta) -> str`

Adjusts a range string such as `"B2:G500"` or `"B2"` by incrementing each column letter
sequence and each row number.

```python
import copy, re
from openpyxl.utils import get_column_letter, column_index_from_string

_CELL_REF_RE = re.compile(r'(\$?)([A-Z]+)(\$?)(\d+)')

def _shift_sqref(sqref: str, col_delta: int, row_delta: int) -> str:
    def _shift_ref(m):
        col_abs, col_letters, row_abs, row_num = m.group(1), m.group(2), m.group(3), int(m.group(4))
        new_col = get_column_letter(column_index_from_string(col_letters) + col_delta)
        new_row = row_num + row_delta
        return f"{col_abs}{new_col}{row_abs}{new_row}"
    return _CELL_REF_RE.sub(_shift_ref, sqref)
```

`sqref` may contain multiple ranges separated by spaces (e.g. `"A1:C10 E1:E10"`); the
regex handles each match independently so this works automatically.

---

### 3. Add `_shift_formula(formula, col_delta, local_row_delta, hdr_row_by_sheet) -> Optional[str]`

Returns the adjusted formula string, or `None` if the formula cannot be safely rewritten
(skip-and-warn trigger).

```python
_CROSS_SHEET_REF_RE = re.compile(
    r"(?:'([^']+)'|([A-Za-z]\w*))!(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)"
)

def _shift_formula(
    formula: str,
    col_delta: int,
    local_row_delta: int,
    hdr_row_by_sheet: Dict[str, int],
) -> Optional[str]:
    # Skip rules that use INDIRECT() — references are dynamic strings, not static
    if re.search(r'\bINDIRECT\s*\(', formula, re.IGNORECASE):
        return None

    # Rewrite cross-sheet references first (they carry their own sheet name)
    def _shift_cross(m):
        sheet_name = m.group(1) or m.group(2)
        ref_part = m.group(3)
        other_hdr_row = hdr_row_by_sheet.get(sheet_name, 1)
        other_row_delta = 1 - other_hdr_row
        shifted_ref = _shift_sqref(ref_part, col_delta, other_row_delta)
        quoted = f"'{sheet_name}'" if ' ' in sheet_name else sheet_name
        return f"{quoted}!{shifted_ref}"

    result = _CROSS_SHEET_REF_RE.sub(_shift_cross, formula)

    # Rewrite remaining local cell references
    result = _shift_sqref(result, col_delta, local_row_delta)

    return result
```

---

### 4. Copy rules after writing cell values for each `out_ws`

Add this block at the end of the per-sheet loop in `scramble_xlsx_file`, after the cell
value writing loop:

```python
skipped_cf_rules: List[str] = []

col_delta = 1  # Row ID column prepended
local_row_delta = 1 - hdr_row

for cf in ws.conditional_formatting:
    new_sqref = _shift_sqref(str(cf.sqref), col_delta, local_row_delta)
    adjusted_rules = []
    skip_this_cf = False

    for rule in cf.rules:
        adjusted_rule = copy.deepcopy(rule)
        if adjusted_rule.formula:
            new_formulas = []
            for f in adjusted_rule.formula:
                shifted = _shift_formula(f, col_delta, local_row_delta, hdr_row_by_sheet)
                if shifted is None:
                    skipped_cf_rules.append(
                        f"Sheet '{ws.title}' / sqref {cf.sqref}: skipped rule with INDIRECT or named range"
                    )
                    skip_this_cf = True
                    break
                new_formulas.append(shifted)
            if skip_this_cf:
                break
            adjusted_rule.formula = new_formulas
        adjusted_rules.append(adjusted_rule)

    if not skip_this_cf:
        for adjusted_rule in adjusted_rules:
            out_ws.conditional_formatting.add(new_sqref, adjusted_rule)
        if orig_ws is not None:
            for adjusted_rule in adjusted_rules:
                orig_ws.conditional_formatting.add(new_sqref, copy.deepcopy(adjusted_rule))
```

---

### 5. Surface skipped rules in `RunResult`

`RunResult` already has a `warnings` field (or equivalent). Append any entries from
`skipped_cf_rules` to it so the GUI and CLI can surface them after the run.

If `RunResult` does not yet have a `warnings` list, add one:

```python
@dataclass
class RunResult:
    row_count: int
    keyed_path: Optional[str]
    warnings: List[str] = field(default_factory=list)  # add if missing
```

---

### 6. Version bump and documentation

- Bump `APP_VERSION` in `csv_scrambler_app.py` from `1.5.0` → `1.6.0`.
- Add a brief entry to `SERVICE_MANUAL.md` under the Excel pipeline section noting that
  conditional formatting is preserved, with a note about the skip-and-warn cases
  (`INDIRECT()`, named ranges, external workbook references).

---

## Rules that are fully preserved (no skip risk)

- ColorScale, DataBar, IconSet — no formulas, only `sqref` shift needed
- CellIsRule (`cellIs` type) — formula is a literal value like `["10"]`, no cell refs
- FormulaRule / expression rules using local `$COL ROW` patterns
- FormulaRule / expression rules using cross-sheet `SheetName!$COL ROW` patterns

## Rules that are skipped with a warning

- Any rule whose formula contains `INDIRECT()`
- Any rule whose formula references a named range (no `!` and no `$COL ROW` pattern found)
- Any rule referencing an external workbook (`[OtherFile.xlsx]Sheet1!...`)
