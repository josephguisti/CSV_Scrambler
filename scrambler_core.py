# CSV/Excel Scrambler — Core Logic (business rules live here)
#
# This module contains the **rules** and **pipelines** used by the GUI app.
# If you're maintaining this with AI help, search for:
#   - "PIPELINE STEP" (high-level file processing steps)
#   - "BUSINESS RULE" (important policy decisions)
#   - "COMMON EDITS"  (where to change common behaviors)
#
import calendar
import copy
import csv
import datetime as dt
import hashlib
import hmac
import os
import re
import secrets
import uuid
from dataclasses import dataclass, field
from typing import Dict, Iterable, List, Literal, Optional, Sequence, Set, Tuple


ScrambleMode = Literal["full scramble", "date scramble", "no scramble", "delete data", "shuffle values"]


class EncryptedXlsxError(ValueError):
    """Raised when an xlsx file is encrypted and no password was supplied."""


def _decrypt_xlsx(path: str, password: str) -> "io.BytesIO":
    """Decrypt a password-protected xlsx file in-memory and return a BytesIO.

    The returned object can be passed directly to openpyxl.load_workbook().
    No decrypted file is written to disk.
    """
    import io
    try:
        import msoffcrypto  # type: ignore
    except ImportError as e:
        raise RuntimeError(
            "msoffcrypto-tool is required to open encrypted files. "
            "Run: pip install msoffcrypto-tool"
        ) from e
    buf = io.BytesIO()
    with open(path, "rb") as f:
        office_file = msoffcrypto.OfficeFile(f)
        office_file.load_key(password=password)
        office_file.decrypt(buf)
    buf.seek(0)
    return buf


ISO_DATE_RE = re.compile(r"(?P<y>\d{4})-(?P<m>\d{2})-(?P<d>\d{2})")
US_DATE_RE = re.compile(r"\b(?P<m>\d{1,2})/(?P<d>\d{1,2})/(?P<y>\d{4})\b")
TIME_RE = re.compile(r"\b(?P<h>\d{1,2}):(?P<min>\d{2})(?P<ampm>[AaPp][Mm])?\b")


def _norm_header(s: str) -> str:
    return (s or "").casefold()


def header_contains_any(header: str, needles: Iterable[str]) -> bool:
    h = _norm_header(header)
    return any(n.casefold() in h for n in needles)


def is_program_project_stable_header(header: str) -> bool:
    # Stable if header contains "program" or "project" and does NOT contain "date"
    return header_contains_any(header, ["program", "project"]) and not header_contains_any(header, ["date"])


def _col_looks_like_dates(samples: List[str], threshold: float = 0.5) -> bool:
    """Return True if at least *threshold* fraction of non-blank sampled strings match a date pattern.

    Uses the same ISO_DATE_RE / US_DATE_RE patterns used during scrambling, so detection
    is consistent with what the date-scramble mode actually processes.
    """
    non_blank = [s for s in samples if s.strip()]
    if not non_blank:
        return False
    hits = sum(1 for s in non_blank if ISO_DATE_RE.search(s) or US_DATE_RE.search(s))
    return hits / len(non_blank) >= threshold


def default_mode_for_column(header: str, *, is_date: bool = False) -> ScrambleMode:
    # BUSINESS RULE (defaults) — used by CLI and as pipeline fallback when modes are missing:
    # - Program/Project columns (header contains "program"/"project" but not "date") stay stable.
    # - Columns whose cell values look like dates (detected by sampling) default to "date scramble".
    # - Everything else defaults to "full scramble".
    # The GUI app overrides this by initializing every dropdown to "no scramble" (see csv_scrambler_app.py).
    if is_program_project_stable_header(header):
        return "no scramble"
    if is_date:
        return "date scramble"
    return "full scramble"


def default_mode_for_header(header: str) -> ScrambleMode:
    # Backward-compat wrapper for callers that have no column samples (e.g. CLI fallback).
    return default_mode_for_column(header)


def _hmac_bytes(seed: bytes, msg: str) -> bytes:
    return hmac.new(seed, msg.encode("utf-8"), hashlib.sha256).digest()


def _pick_from(seed: bytes, msg: str, items: Sequence[int]) -> int:
    if not items:
        raise ValueError("No candidate items to pick from.")
    b = _hmac_bytes(seed, msg)
    n = int.from_bytes(b[:8], "big", signed=False)
    return items[n % len(items)]


def _month_last_day(year: int, month: int) -> int:
    return calendar.monthrange(year, month)[1]


def _parse_iso(y: str, m: str, d: str) -> dt.date:
    return dt.date(int(y), int(m), int(d))


def _parse_us(m: str, d: str, y: str) -> dt.date:
    return dt.date(int(y), int(m), int(d))


def _format_like_iso(new_date: dt.date) -> str:
    return f"{new_date.year:04d}-{new_date.month:02d}-{new_date.day:02d}"


def _format_like_us(new_date: dt.date) -> str:
    return f"{new_date.month}/{new_date.day}/{new_date.year}"


def age_years(dob: dt.date, as_of: dt.date) -> int:
    years = as_of.year - dob.year
    if (as_of.month, as_of.day) < (dob.month, dob.day):
        years -= 1
    return years


def age_bucket(dob: dt.date, as_of: dt.date) -> str:
    a = age_years(dob, as_of)
    if a < 18:
        return "minor"
    if 18 <= a <= 24:
        return "tay"
    return "adult"


def _scramble_day_same_month_year(*, seed: bytes, original: dt.date, scope_tag: str, extra_key: str) -> dt.date:
    last = _month_last_day(original.year, original.month)
    candidates = list(range(1, last + 1))
    day = _pick_from(seed, f"{scope_tag}|{original.isoformat()}|{extra_key}", candidates)
    return dt.date(original.year, original.month, day)


def _scramble_dob_preserve_bucket_same_month_year(
    *, seed: bytes, original_dob: dt.date, as_of: dt.date, extra_key: str
) -> dt.date:
    original_bucket = age_bucket(original_dob, as_of)
    last = _month_last_day(original_dob.year, original_dob.month)
    candidates: List[int] = []
    for day in range(1, last + 1):
        d = dt.date(original_dob.year, original_dob.month, day)
        if age_bucket(d, as_of) == original_bucket:
            candidates.append(day)
    if not candidates:
        return original_dob
    picked = _pick_from(seed, f"DOB|{original_dob.isoformat()}|ASOF|{as_of.isoformat()}|{extra_key}", candidates)
    return dt.date(original_dob.year, original_dob.month, picked)


def _scramble_run(seed: bytes, run: str) -> str:
    letters = "abcdefghijklmnopqrstuvwxyz"
    digits = "0123456789"
    out: List[str] = []
    for idx, ch in enumerate(run):
        if ch.isalpha():
            b = _hmac_bytes(seed, f"ALPHA|{run}|{idx}")
            new_ch = letters[b[0] % 26]
            out.append(new_ch.upper() if ch.isupper() else new_ch)
        elif ch.isdigit():
            b = _hmac_bytes(seed, f"DIGIT|{run}|{idx}")
            out.append(digits[b[0] % 10])
        else:
            out.append(ch)
    return "".join(out)


def scramble_alnum_runs(seed: bytes, text: str) -> str:
    out: List[str] = []
    i = 0
    n = len(text)
    while i < n:
        ch = text[i]
        if ch.isalnum():
            j = i + 1
            while j < n and text[j].isalnum():
                j += 1
            out.append(_scramble_run(seed, text[i:j]))
            i = j
        else:
            out.append(ch)
            i += 1
    return "".join(out)


@dataclass(frozen=True)
class ColumnContext:
    header: str
    mode: ScrambleMode


def _replace_protected_substrings_with_placeholders(
    *, seed: bytes, text: str, col: ColumnContext, as_of: dt.date
) -> Tuple[str, Dict[str, str]]:
    repl_map: Dict[str, str] = {}

    def make_placeholder(k: int) -> str:
        return "\uE000" + chr(0xE001 + k) + "\uE002"

    def add_replacement(new_value: str) -> str:
        ph = make_placeholder(len(repl_map))
        repl_map[ph] = new_value
        return ph

    def scramble_date(d: dt.date, scope: str, original_str: str) -> dt.date:
        extra_key = f"{col.header}|{original_str}"
        # BUSINESS RULE: all date-scrambled columns use bucket-preserving logic.
        # This ensures minor/TAY/adult age buckets are never exposed regardless of column type.
        return _scramble_dob_preserve_bucket_same_month_year(
            seed=seed, original_dob=d, as_of=as_of, extra_key=extra_key
        )

    def iso_sub(m: re.Match) -> str:
        original_str = m.group(0)
        d = _parse_iso(m.group("y"), m.group("m"), m.group("d"))
        new_d = scramble_date(d, "ISO", original_str)
        return add_replacement(_format_like_iso(new_d))

    def us_sub(m: re.Match) -> str:
        original_str = m.group(0)
        d = _parse_us(m.group("m"), m.group("d"), m.group("y"))
        new_d = scramble_date(d, "US", original_str)
        return add_replacement(_format_like_us(new_d))

    tmp = ISO_DATE_RE.sub(iso_sub, text)
    tmp = US_DATE_RE.sub(us_sub, tmp)
    # Always preserve time-of-day exactly for "date scramble".
    tmp = TIME_RE.sub(lambda m: add_replacement(m.group(0)), tmp)
    return tmp, repl_map


def scramble_cell(*, seed: bytes, value: str, col: ColumnContext, as_of: dt.date) -> str:
    # BUSINESS RULE (what each dropdown option means):
    # - no scramble: return the cell unchanged
    # - delete data: blank it out
    # - full scramble: scramble letters/digits everywhere, INCLUDING dates/times (no date handling)
    # - date scramble: apply date/DOB rules, then scramble remaining letters/digits
    if value is None:
        return value
    text = str(value)

    if col.mode == "no scramble":
        return text
    if col.mode == "delete data":
        return ""
    if col.mode == "full scramble":
        # Full means truly full: no special date/time handling.
        return scramble_alnum_runs(seed, text)
    if col.mode == "date scramble":
        tmp, repl_map = _replace_protected_substrings_with_placeholders(seed=seed, text=text, col=col, as_of=as_of)
        scrambled = scramble_alnum_runs(seed, tmp)
        for ph, rep in repl_map.items():
            scrambled = scrambled.replace(ph, rep)
        return scrambled

    raise ValueError(f"Unsupported mode for scramble_cell: {col.mode}")


class ShuffleValuesError(Exception):
    """Raised when shuffle-values reassignment is impossible for a column."""

    def __init__(self, message: str, *, column_key: Optional[str] = None) -> None:
        super().__init__(message)
        self.column_key = column_key


def make_shuffle_derangement(values: List[str]) -> List[str]:
    """
    Given a list of values (non-blank only), return a shuffled list of same multiset
    such that new[i] != old[i] for all i.

    Raises ShuffleValuesError if impossible.
    """
    # BUSINESS RULE (shuffle values):
    # We must reassign values so that NO ROW keeps its original value (derangement).
    # This is only guaranteed when the most-common value frequency is <= floor(n/2).
    n = len(values)
    if n < 2:
        raise ShuffleValuesError("Not enough non-blank rows to shuffle.")
    freq: Dict[str, int] = {}
    for v in values:
        freq[v] = freq.get(v, 0) + 1
    maxfreq = max(freq.values())
    if maxfreq == n:
        raise ShuffleValuesError("All rows have the same value; cannot shuffle.")
    if maxfreq > n // 2:
        raise ShuffleValuesError("One value appears too frequently to guarantee every row changes.")

    pairs = sorted([(v, i) for i, v in enumerate(values)], key=lambda t: t[0])
    sorted_values = [v for v, _ in pairs]

    for shift in range(maxfreq, n):
        rotated = sorted_values[shift:] + sorted_values[:shift]
        new_by_orig_index = ["" for _ in range(n)]
        ok = True
        for j, (_, orig_i) in enumerate(pairs):
            nv = rotated[j]
            if nv == values[orig_i]:
                ok = False
                break
            new_by_orig_index[orig_i] = nv
        if ok:
            return new_by_orig_index

    # Shouldn't happen if maxfreq <= n//2, but keep safe.
    raise ShuffleValuesError("Unable to find a shuffle where every row changes.")


@dataclass
class RunResult:
    scrambled_output_path: str
    keyed_original_output_path: Optional[str]
    row_count: int
    warnings: List[str] = field(default_factory=list)


def _validate_unique_headers(headers: List[str], *, sheet_name: Optional[str] = None) -> None:
    seen: Dict[str, int] = {}
    dups: List[str] = []
    for h in headers:
        seen[h] = seen.get(h, 0) + 1
    for h, c in seen.items():
        if c > 1:
            dups.append(h)
    if dups:
        detail = ", ".join(repr(d) for d in dups)
        if sheet_name is not None:
            raise ValueError(f"Sheet {sheet_name!r}: duplicate headers are not supported: {detail}")
        raise ValueError(f"Duplicate headers are not supported: {detail}")


_EXCEL_ERROR_RE = re.compile(r"^#[A-Z/]+[!?]?$")  # matches #REF!, #VALUE!, #N/A, #DIV/0!, etc.


def _is_excel_error(v: object) -> bool:
    return bool(v and _EXCEL_ERROR_RE.match(str(v)))


# ---------------------------------------------------------------------------
# Conditional-formatting coordinate helpers
# ---------------------------------------------------------------------------
_CELL_REF_RE = re.compile(r"(\$?)([A-Z]+)(\$?)(\d+)")
_CROSS_SHEET_REF_RE = re.compile(
    r"(?:'([^']+)'|([A-Za-z]\w*))!(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)"
)


def _shift_sqref(sqref: str, col_delta: int, row_delta: int) -> str:
    """Shift every cell reference in a sqref string by col_delta columns and row_delta rows."""
    def _shift_ref(m: re.Match) -> str:  # type: ignore[type-arg]
        col_abs, col_letters, row_abs, row_num = m.group(1), m.group(2), m.group(3), int(m.group(4))
        try:
            from openpyxl.utils import get_column_letter, column_index_from_string  # type: ignore
            new_col = get_column_letter(max(1, column_index_from_string(col_letters) + col_delta))
        except Exception:
            new_col = col_letters  # leave unchanged on import failure
        new_row = max(1, row_num + row_delta)
        return f"{col_abs}{new_col}{row_abs}{new_row}"
    return _CELL_REF_RE.sub(_shift_ref, sqref)


def _shift_formula(
    formula: str,
    col_delta: int,
    local_row_delta: int,
    hdr_row_by_sheet: Dict[str, int],
) -> Optional[str]:
    """Return a formula string with all cell references shifted, or None if it cannot be rewritten.

    Returns None for formulas containing INDIRECT() (dynamic references) or external workbook
    references ([File.xlsx]Sheet!...) which cannot be statically adjusted.
    """
    if re.search(r"\bINDIRECT\s*\(", formula, re.IGNORECASE):
        return None
    if re.search(r"\[[^\]]+\]", formula):  # external workbook reference
        return None

    def _shift_cross(m: re.Match) -> str:  # type: ignore[type-arg]
        sheet_name = m.group(1) or m.group(2)
        ref_part = m.group(3)
        other_hdr_row = hdr_row_by_sheet.get(sheet_name, 1)
        other_row_delta = 1 - other_hdr_row
        shifted_ref = _shift_sqref(ref_part, col_delta, other_row_delta)
        quoted = f"'{sheet_name}'" if " " in sheet_name else sheet_name
        return f"{quoted}!{shifted_ref}"

    result = _CROSS_SHEET_REF_RE.sub(_shift_cross, formula)
    result = _shift_sqref(result, col_delta, local_row_delta)
    return result


def _detect_cross_sheet_breaks(wb, include_sheets: List[str]) -> List[str]:  # type: ignore[no-untyped-def]
    """Return human-readable warnings when an included sheet's CF rules reference excluded sheets."""
    include_set = set(include_sheets)
    all_titles = {ws.title for ws in wb.worksheets}
    seen: Set[str] = set()
    out: List[str] = []

    for ws in wb.worksheets:
        if ws.title not in include_set:
            continue
        for cf in ws.conditional_formatting:
            for rule in cf.rules:
                formulas = getattr(rule, "formula", None) or []
                if not formulas:
                    continue
                for f_str in formulas:
                    if not isinstance(f_str, str):
                        continue
                    for m in _CROSS_SHEET_REF_RE.finditer(f_str):
                        ref_sheet = m.group(1) or m.group(2)
                        if ref_sheet not in all_titles:
                            continue
                        if ref_sheet not in include_set:
                            msg = (
                                f"Sheet '{ws.title}' has a conditional formatting rule that references "
                                f"'{ref_sheet}' (excluded from export)"
                            )
                            if msg not in seen:
                                seen.add(msg)
                                out.append(msg)
    return out


def detect_cross_sheet_cf_breaks(
    input_path: str,
    include_sheets: List[str],
    password: Optional[str] = None,
) -> List[str]:
    """Load workbook and return CF cross-sheet dependency warnings (for GUI pre-export check)."""
    import zipfile

    try:
        import openpyxl  # type: ignore
    except Exception as e:
        raise RuntimeError("Excel support requires openpyxl.") from e

    try:
        if password is not None:
            wb = openpyxl.load_workbook(_decrypt_xlsx(input_path, password), data_only=True)
        else:
            wb = openpyxl.load_workbook(input_path, data_only=True)
    except zipfile.BadZipFile as e:
        raise EncryptedXlsxError(
            "This file is password-protected. Select it again and enter the password when prompted."
        ) from e
    return _detect_cross_sheet_breaks(wb, include_sheets)


# ---------------------------------------------------------------------------


def _xlsx_header_row(ws) -> int:  # type: ignore[no-untyped-def]
    """Return the 1-based row number that contains per-column headers for this worksheet.

    Detection order:
    1. AutoFilter reference — Excel always places filter dropdowns on the header row.
       ws.auto_filter.ref is a range string like "A5:G200"; we extract the first row.
    2. Density scan — scan rows 1–20 and return the first row where more than half of
       its cells are non-empty and not Excel error strings (handles any number of
       preamble rows: titles, date stamps, notes, merged group labels, etc.).
    3. Fall back to row 1 if nothing else matches.
    """
    # Strategy 1: AutoFilter reference
    try:
        af_ref = ws.auto_filter.ref  # type: ignore[union-attr]
        if af_ref:
            from openpyxl.utils.cell import coordinate_to_tuple  # type: ignore
            first_cell = str(af_ref).split(":")[0]
            row, _ = coordinate_to_tuple(first_cell)
            if 1 <= row <= 20:
                return row
    except Exception:
        pass

    # Strategy 2: generalized density scan (up to 20 preamble rows)
    max_col = ws.max_column or 0
    if max_col == 0:
        return 1
    max_scan = min(20, ws.max_row or 1)
    for r in range(1, max_scan + 1):
        non_empty = sum(
            1 for c in range(1, max_col + 1)
            if ws.cell(r, c).value and not _is_excel_error(ws.cell(r, c).value)
        )
        if non_empty > max_col / 2:
            return r

    return 1


def scramble_csv_file(
    *,
    input_path: str,
    output_scrambled_path: str,
    header_modes: Dict[str, ScrambleMode],
    create_keyed_normal_version: bool,
    row_id_column_name: str = "Row ID",
    encoding: str = "utf-8",
    as_of: Optional[dt.date] = None,
) -> RunResult:
    """
    Scramble CSV at input_path into output_scrambled_path, using per-header modes.
    If create_keyed_normal_version is True, also writes a keyed original copy with Row ID column.

    Notes:
    - For "shuffle values": blanks are left blank; non-blank values are deranged so each row changes.
    - GUI should validate header_modes contains every header.
    """
    # PIPELINE STEP 1 (CSV): read headers + scan rows once to:
    # - generate Row IDs (if keyed output enabled)
    # - collect the values needed for any "shuffle values" columns
    as_of = as_of or dt.date.today()
    seed = secrets.token_bytes(32)  # GUI mode: no seed exposure

    # Pass 1: read headers and collect shuffle columns' values, plus row IDs.
    with open(input_path, "r", encoding=encoding, newline="") as f:
        reader = csv.reader(f)
        try:
            headers = next(reader)
        except StopIteration:
            # Empty file: just write headers.
            with open(output_scrambled_path, "w", encoding=encoding, newline="") as f_out:
                csv.writer(f_out).writerow([row_id_column_name] + [])
            return RunResult(scrambled_output_path=output_scrambled_path, keyed_original_output_path=None, row_count=0)

        _validate_unique_headers(headers)
        header_to_idx = {h: i for i, h in enumerate(headers)}
        shuffle_headers = [h for h in headers if header_modes.get(h) == "shuffle values"]

        shuffle_values_by_header: Dict[str, List[Tuple[int, str]]] = {h: [] for h in shuffle_headers}
        row_ids: List[str] = []
        row_count = 0
        for row in reader:
            row_count += 1
            row_ids.append(str(uuid.uuid4()))
            for h in shuffle_headers:
                idx = header_to_idx[h]
                val = row[idx] if idx < len(row) else ""
                if val is None:
                    val = ""
                if str(val).strip() == "":
                    continue  # blanks stay blank (Option B)
                shuffle_values_by_header[h].append((row_count - 1, str(val)))

    # Compute per-row shuffled assignments for each shuffle column.
    # PIPELINE STEP 2 (CSV): pre-compute the shuffle mapping so we can write output in one streaming pass.
    shuffled_assignment: Dict[str, Dict[int, str]] = {}
    for h, pairs in shuffle_values_by_header.items():
        row_idxs = [ri for (ri, _) in pairs]
        vals = [v for (_, v) in pairs]
        try:
            new_vals = make_shuffle_derangement(vals)
        except ShuffleValuesError as e:
            raise ShuffleValuesError(f"{h}: {e}", column_key=h) from e
        shuffled_assignment[h] = {row_idxs[i]: new_vals[i] for i in range(len(row_idxs))}

    keyed_original_path: Optional[str] = None
    if create_keyed_normal_version:
        base, ext = os.path.splitext(output_scrambled_path)
        if not ext:
            ext = ".csv"
        keyed_original_path = f"{base} (KEYED ORIGINAL){ext}"

    # Pass 2: stream input and write outputs.
    # PIPELINE STEP 3 (CSV): read rows again and write:
    # - scrambled output
    # - (optional) keyed original output
    with open(input_path, "r", encoding=encoding, newline="") as f_in, open(
        output_scrambled_path, "w", encoding=encoding, newline=""
    ) as f_out_scr:
        reader = csv.reader(f_in)
        writer_scr = csv.writer(f_out_scr)

        headers = next(reader, [])
        writer_scr.writerow([row_id_column_name] + headers)

        writer_orig = None
        f_out_orig = None
        if keyed_original_path:
            f_out_orig = open(keyed_original_path, "w", encoding=encoding, newline="")
            writer_orig = csv.writer(f_out_orig)
            writer_orig.writerow([row_id_column_name] + headers)

        try:
            for row_idx, row in enumerate(reader):
                # Original keyed output: row unchanged except Row ID.
                if writer_orig is not None:
                    writer_orig.writerow([row_ids[row_idx]] + row)

                # Scrambled output
                if len(row) < len(headers):
                    row = row + [""] * (len(headers) - len(row))
                new_row: List[str] = []
                for col_i, header in enumerate(headers):
                    mode = header_modes[header]
                    if mode == "shuffle values":
                        # blanks remain blank; non-blanks replaced from assignment
                        original_val = row[col_i] if col_i < len(row) else ""
                        if str(original_val).strip() == "":
                            new_row.append(original_val)
                        else:
                            new_row.append(shuffled_assignment[header].get(row_idx, original_val))
                        continue

                    col = ColumnContext(header=header, mode=mode)
                    new_row.append(scramble_cell(seed=seed, value=row[col_i], col=col, as_of=as_of))

                writer_scr.writerow([row_ids[row_idx]] + new_row)
        finally:
            if f_out_orig is not None:
                f_out_orig.close()

    return RunResult(
        scrambled_output_path=output_scrambled_path,
        keyed_original_output_path=keyed_original_path,
        row_count=row_count,
    )


def _copy_cell_static_style(src, dst) -> None:  # type: ignore[no-untyped-def]
    """Copy static cell styling from *src* to *dst* (fill, font, border, alignment, etc.).

    When copying across workbooks, copy individual style objects—not *_style* indices.
    """
    if not getattr(src, "has_style", False):
        return
    try:
        dst.fill = copy.copy(src.fill)
        dst.font = copy.copy(src.font)
        dst.border = copy.copy(src.border)
        dst.alignment = copy.copy(src.alignment)
        dst.number_format = src.number_format
        dst.protection = copy.copy(src.protection)
    except Exception:
        pass


def scramble_xlsx_file(
    *,
    input_path: str,
    output_scrambled_path: str,
    header_modes: Dict[str, ScrambleMode],
    create_keyed_normal_version: bool,
    row_id_column_name: str = "Row ID",
    as_of: Optional[dt.date] = None,
    password: Optional[str] = None,
    include_sheets: Optional[List[str]] = None,
) -> RunResult:
    """
    Scramble all worksheet tabs in an XLSX file.

    Assumptions:
    - Chart-only tabs are not in ``wb.worksheets`` (openpyxl); they are skipped like today.
    - Row 1 on each worksheet is headers; column rules apply by header name across sheets.
    - Output is values-only (no styles/tables/filters).
    """
    import zipfile

    try:
        import openpyxl  # type: ignore
    except Exception as e:
        raise RuntimeError("Excel support requires openpyxl. Install it, or use CSV.") from e

    as_of = as_of or dt.date.today()
    seed = secrets.token_bytes(32)  # GUI mode: no seed exposure

    try:
        if password is not None:
            wb = openpyxl.load_workbook(_decrypt_xlsx(input_path, password), data_only=True)
        else:
            wb = openpyxl.load_workbook(input_path, data_only=True)
    except zipfile.BadZipFile as e:
        raise EncryptedXlsxError(
            "This file is password-protected. Select it again and enter the password when prompted."
        ) from e
    if not wb.worksheets:
        raise ValueError("This Excel file has no worksheets.")

    all_sheet_titles = [ws.title for ws in wb.worksheets]
    if include_sheets is not None:
        unknown = set(include_sheets) - set(all_sheet_titles)
        if unknown:
            raise ValueError(f"Unknown sheet name(s) in include_sheets: {sorted(unknown)!r}")
        if len(include_sheets) == 0:
            raise ValueError("include_sheets cannot be empty.")

    sheets_export_count = len(include_sheets) if include_sheets is not None else len(all_sheet_titles)
    out_ext = os.path.splitext(output_scrambled_path)[1].lower()
    if out_ext == ".csv" and sheets_export_count > 1:
        raise ValueError(
            "Cannot save a multi-sheet Excel workbook as a CSV file. "
            "Choose an output path ending in .xlsx, or save each sheet separately outside this tool."
        )

    keyed_original_path: Optional[str] = None
    if create_keyed_normal_version:
        base, ext = os.path.splitext(output_scrambled_path)
        if not ext:
            ext = ".xlsx"
        keyed_original_path = f"{base} (KEYED ORIGINAL){ext}"

    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)

    orig_wb = None
    if keyed_original_path:
        orig_wb = openpyxl.Workbook()
        orig_wb.remove(orig_wb.active)

    total_row_count = 0
    skipped_cf_rules: List[str] = []

    # Pre-scan all sheets so cross-sheet formula references can be row-shifted correctly.
    hdr_row_by_sheet: Dict[str, int] = {ws.title: _xlsx_header_row(ws) for ws in wb.worksheets}

    for ws in wb.worksheets:
        if include_sheets is not None and ws.title not in include_sheets:
            continue
        max_col = ws.max_column or 0
        hdr_row = _xlsx_header_row(ws)
        headers: List[str] = []
        for c in range(1, max_col + 1):
            v = ws.cell(row=hdr_row, column=c).value
            headers.append("" if (v is None or _is_excel_error(v)) else str(v))

        _validate_unique_headers(headers, sheet_name=ws.title)

        max_row = ws.max_row or hdr_row
        sheet_row_count = max(0, max_row - hdr_row)
        total_row_count += sheet_row_count
        row_ids = [str(uuid.uuid4()) for _ in range(sheet_row_count)]

        shuffle_headers = [h for h in headers if header_modes.get(f"{h} ({ws.title})") == "shuffle values"]
        shuffle_values_by_header: Dict[str, List[Tuple[int, str]]] = {h: [] for h in shuffle_headers}

        header_to_colnum = {h: i + 1 for i, h in enumerate(headers)}
        for r_idx in range(hdr_row + 1, max_row + 1):
            row0 = r_idx - hdr_row - 1
            for h in shuffle_headers:
                cnum = header_to_colnum[h]
                val = ws.cell(row=r_idx, column=cnum).value
                s = "" if val is None else str(val)
                if s.strip() == "":
                    continue
                shuffle_values_by_header[h].append((row0, s))

        shuffled_assignment: Dict[str, Dict[int, str]] = {}
        for h, pairs in shuffle_values_by_header.items():
            row_idxs = [ri for (ri, _) in pairs]
            vals = [v for (_, v) in pairs]
            try:
                new_vals = make_shuffle_derangement(vals)
            except ShuffleValuesError as e:
                raise ShuffleValuesError(
                    f"Sheet {ws.title!r} / {h}: {e}",
                    column_key=f"{h} ({ws.title})",
                ) from e
            shuffled_assignment[h] = {row_idxs[i]: new_vals[i] for i in range(len(row_idxs))}

        out_ws = out_wb.create_sheet(title=ws.title)
        out_ws.cell(row=1, column=1).value = row_id_column_name
        _copy_cell_static_style(ws.cell(row=hdr_row, column=1), out_ws.cell(row=1, column=1))
        for i, h in enumerate(headers, start=2):
            out_cell = out_ws.cell(row=1, column=i)
            out_cell.value = h
            _copy_cell_static_style(ws.cell(row=hdr_row, column=i - 1), out_cell)

        orig_ws = None
        if orig_wb is not None:
            orig_ws = orig_wb.create_sheet(title=ws.title)
            orig_ws.cell(row=1, column=1).value = row_id_column_name
            _copy_cell_static_style(ws.cell(row=hdr_row, column=1), orig_ws.cell(row=1, column=1))
            for i, h in enumerate(headers, start=2):
                oc = orig_ws.cell(row=1, column=i)
                oc.value = h
                _copy_cell_static_style(ws.cell(row=hdr_row, column=i - 1), oc)

        for r_idx in range(hdr_row + 1, max_row + 1):
            row0 = r_idx - hdr_row - 1
            out_r = row0 + 2  # output row: row 1 is headers, data starts at row 2
            out_ws.cell(row=out_r, column=1).value = row_ids[row0]
            _copy_cell_static_style(ws.cell(row=r_idx, column=1), out_ws.cell(row=out_r, column=1))
            if orig_ws is not None:
                oc_rid = orig_ws.cell(row=out_r, column=1)
                oc_rid.value = row_ids[row0]
                _copy_cell_static_style(ws.cell(row=r_idx, column=1), oc_rid)

            for c in range(1, max_col + 1):
                header = headers[c - 1]
                mode = header_modes.get(f"{header} ({ws.title})", default_mode_for_header(header))
                val = ws.cell(row=r_idx, column=c).value
                s = "" if val is None else str(val)
                src_cell = ws.cell(row=r_idx, column=c)
                dst_out = out_ws.cell(row=out_r, column=c + 1)

                if orig_ws is not None:
                    oc = orig_ws.cell(row=out_r, column=c + 1)
                    oc.value = s
                    _copy_cell_static_style(src_cell, oc)

                if mode == "shuffle values":
                    if s.strip() == "":
                        dst_out.value = s
                    else:
                        dst_out.value = shuffled_assignment[header].get(row0, s)
                    _copy_cell_static_style(src_cell, dst_out)
                    continue

                col = ColumnContext(header=header, mode=mode)
                dst_out.value = scramble_cell(seed=seed, value=s, col=col, as_of=as_of)
                _copy_cell_static_style(src_cell, dst_out)

        # PIPELINE STEP 3B: Copy conditional formatting from source sheet, adjusting
        # all cell coordinates to match the output layout (Row ID column +1, hdr_row offset).
        col_delta = 1  # Row ID prepended at column A
        local_row_delta = 1 - hdr_row  # normalise source rows to 1-based output rows

        for cf in ws.conditional_formatting:
            try:
                new_sqref = _shift_sqref(str(cf.sqref), col_delta, local_row_delta)
            except Exception:
                skipped_cf_rules.append(
                    f"Sheet '{ws.title}' / sqref {cf.sqref}: could not shift sqref"
                )
                continue

            adjusted_rules: List[object] = []
            skip_this_cf = False

            for rule in cf.rules:
                adjusted_rule = copy.deepcopy(rule)
                if getattr(adjusted_rule, "formula", None):
                    new_formulas = []
                    for f_str in adjusted_rule.formula:
                        shifted = _shift_formula(f_str, col_delta, local_row_delta, hdr_row_by_sheet)
                        if shifted is None:
                            skipped_cf_rules.append(
                                f"Sheet '{ws.title}' / sqref {cf.sqref}: "
                                f"skipped rule containing INDIRECT(), external reference, or named range"
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
                    try:
                        out_ws.conditional_formatting.add(new_sqref, adjusted_rule)  # type: ignore[arg-type]
                    except Exception as exc:
                        skipped_cf_rules.append(
                            f"Sheet '{ws.title}' / sqref {cf.sqref}: could not apply rule ({exc})"
                        )
                if orig_ws is not None:
                    for adjusted_rule in adjusted_rules:
                        try:
                            orig_ws.conditional_formatting.add(new_sqref, copy.deepcopy(adjusted_rule))  # type: ignore[arg-type]
                        except Exception:
                            pass  # best-effort for keyed original

    out_wb.save(output_scrambled_path)
    if orig_wb is not None and keyed_original_path is not None:
        orig_wb.save(keyed_original_path)

    return RunResult(
        scrambled_output_path=output_scrambled_path,
        keyed_original_output_path=keyed_original_path,
        row_count=total_row_count,
        warnings=skipped_cf_rules,
    )


def scramble_tabular_file(
    *,
    input_path: str,
    output_scrambled_path: str,
    header_modes: Dict[str, ScrambleMode],
    create_keyed_normal_version: bool,
    row_id_column_name: str = "Row ID",
    encoding: str = "utf-8",
    as_of: Optional[dt.date] = None,
    password: Optional[str] = None,
    include_sheets: Optional[List[str]] = None,
) -> RunResult:
    # PIPELINE STEP 0: Pick the correct pipeline based on file extension.
    ext = os.path.splitext(input_path)[1].lower()
    if ext == ".csv":
        return scramble_csv_file(
            input_path=input_path,
            output_scrambled_path=output_scrambled_path,
            header_modes=header_modes,
            create_keyed_normal_version=create_keyed_normal_version,
            row_id_column_name=row_id_column_name,
            encoding=encoding,
            as_of=as_of,
        )
    if ext in [".xlsx", ".xlsm"]:
        return scramble_xlsx_file(
            input_path=input_path,
            output_scrambled_path=output_scrambled_path,
            header_modes=header_modes,
            create_keyed_normal_version=create_keyed_normal_version,
            row_id_column_name=row_id_column_name,
            as_of=as_of,
            password=password,
            include_sheets=include_sheets,
        )
    raise ValueError(f"Unsupported input type: {ext}. Please use .csv or .xlsx")


# COMMON EDITS (for maintainers)
# - Change default dropdown behavior (CLI / pipeline fallback): edit default_mode_for_column().
# - Add/adjust date formats: edit ISO_DATE_RE / US_DATE_RE / TIME_RE and placeholder logic.
# - Add more file types: extend scramble_tabular_file() and add a new pipeline function.

