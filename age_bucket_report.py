import argparse
import csv
import datetime as dt
from collections import Counter
from typing import Dict, List, Optional


def _norm_header(s: str) -> str:
    return (s or "").casefold()


def _is_dob_like_header(h: str) -> bool:
    hh = _norm_header(h)
    return ("age" in hh) or ("birth" in hh) or ("dob" in hh)


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


def _try_parse_us_date(s: str) -> Optional[dt.date]:
    # Supports M/D/YYYY
    if not s:
        return None
    s = s.strip()
    parts = s.split("/")
    if len(parts) != 3:
        return None
    try:
        m = int(parts[0])
        d = int(parts[1])
        y = int(parts[2])
        return dt.date(y, m, d)
    except Exception:
        return None


def report(path: str, as_of: dt.date, encoding: str = "utf-8") -> Dict[str, Counter]:
    with open(path, "r", encoding=encoding, newline="") as f:
        reader = csv.reader(f)
        try:
            headers = next(reader)
        except StopIteration:
            return {}

        dob_col_idxs: List[int] = [i for i, h in enumerate(headers) if _is_dob_like_header(h)]
        counts: Dict[str, Counter] = {headers[i]: Counter() for i in dob_col_idxs}

        for row in reader:
            for i in dob_col_idxs:
                val = row[i] if i < len(row) else ""
                d = _try_parse_us_date(val)
                if d is None:
                    if val.strip():
                        counts[headers[i]]["unparsed"] += 1
                    continue
                counts[headers[i]][age_bucket(d, as_of)] += 1
        return counts


def main(argv: Optional[List[str]] = None) -> int:
    p = argparse.ArgumentParser(
        description="Report minor/TAY/adult counts for DOB-like columns (headers containing age/birth/dob)."
    )
    p.add_argument("--input", required=True, help="Path to CSV")
    p.add_argument("--encoding", default="utf-8", help="File encoding (default: utf-8)")
    p.add_argument("--as-of", default=None, help="Override 'today' in YYYY-MM-DD. Default: system today.")
    p.add_argument("--compare", default=None, help="Optional second CSV to compare (e.g., original vs scrambled).")
    args = p.parse_args(argv)

    as_of = dt.date.today() if not args.as_of else dt.date.fromisoformat(args.as_of)
    a = report(args.input, as_of=as_of, encoding=args.encoding)

    def print_counts(title: str, counts: Dict[str, Counter]) -> None:
        print(title)
        if not counts:
            print("  (no DOB-like columns found)")
            return
        for col, c in counts.items():
            minor = c.get("minor", 0)
            tay = c.get("tay", 0)
            adult = c.get("adult", 0)
            unparsed = c.get("unparsed", 0)
            print(f"  {col}: minor={minor} tay={tay} adult={adult} unparsed={unparsed}")

    print_counts(f"As-of {as_of.isoformat()} — {args.input}", a)

    if args.compare:
        b = report(args.compare, as_of=as_of, encoding=args.encoding)
        print_counts(f"As-of {as_of.isoformat()} — {args.compare}", b)

        # Basic diff output
        print("Diff (compare - input):")
        all_cols = sorted(set(a.keys()) | set(b.keys()))
        for col in all_cols:
            ca = a.get(col, Counter())
            cb = b.get(col, Counter())
            for k in ["minor", "tay", "adult", "unparsed"]:
                da = ca.get(k, 0)
                db = cb.get(k, 0)
                if db - da != 0:
                    print(f"  {col} {k}: {db-da:+d}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())

