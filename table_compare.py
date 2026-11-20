#!/usr/bin/env python3
"""
bc_compare.py — local re-implementation of Beyond Compare's "Table Compare"
session for Excel (.xlsx) and CSV files.

Mirrors the BC model:
  * Excel files are converted to per-sheet cell tables (BC's "conversion").
    CSV files are parsed directly (delimited + text qualifier).
  * Sheets are aligned by name.
  * Columns are aligned by header name (like "Align by left name").
  * Rows are aligned by key columns (--key), or by an LCS/difflib
    alignment when no keys are given (BC's Myers/standard alignment).
  * Cells are compared per-column: numbers honour a numeric tolerance,
    text can ignore case/whitespace (reported as "unimportant"), dates
    honour a date tolerance. Columns can be flagged unimportant.

Usage:
  python bc_compare.py LEFT RIGHT [options]

Examples:
  python bc_compare.py a.xlsx b.xlsx --key Date,Product
  python bc_compare.py a.csv  b.xlsx --numeric-tolerance 0.001 --ignore-case
  python bc_compare.py a.xlsx b.xlsx --report diff.html --csv-diff diff.csv
"""

from __future__ import annotations

import argparse
import csv
import difflib
import html
import re
import sys
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path

try:
    from openpyxl import load_workbook
except ImportError:
    load_workbook = None

# ---------------------------------------------------------------- loading

Cell = str | int | float | datetime | date | None
Row = list[Cell]


@dataclass
class Table:
    name: str
    headers: list[str]          # display names; "" if unnamed
    rows: list[Row]
    has_header: bool = True

    @property
    def ncols(self) -> int:
        return max([len(self.headers)] + [len(r) for r in self.rows] or [0])

    def cell(self, row: Row, i: int) -> Cell:
        return row[i] if i < len(row) else None


def _trim_table(rows: list[Row]) -> list[Row]:
    """Drop fully-empty trailing rows/columns the way BC does."""
    rows = [r for r in rows]
    while rows and all(v is None or v == "" for v in rows[-1]):
        rows.pop()
    width = 0
    for r in rows:
        for i in range(len(r) - 1, -1, -1):
            if r[i] is not None and r[i] != "":
                width = max(width, i + 1)
                break
    return [(r + [None] * width)[:width] for r in rows]


def _is_scalar(v: Cell) -> bool:
    if isinstance(v, bool):
        return False
    if isinstance(v, (int, float, datetime, date)):
        return True
    return isinstance(v, str) and (to_number(v) is not None or to_datetime(v) is not None)


def _looks_like_header(first: Row, rest: list[Row]) -> bool:
    vals = [v for v in first if v not in (None, "")]
    if not vals:
        return False
    if not all(isinstance(v, str) and not _is_scalar(v) for v in vals):
        return False
    if not rest:
        return True
    return any(_is_scalar(v) for r in rest[:20] for v in r if v not in (None, ""))


def load_csv(path: Path) -> list[Table]:
    text = path.read_text(encoding="utf-8-sig", errors="replace")
    try:
        dialect = csv.Sniffer().sniff(text[:8192], delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel
    rows: list[Row] = [list(r) for r in csv.reader(text.splitlines(), dialect)]
    rows = _trim_table(rows)
    if not rows:
        return [Table(path.stem, [], [])]
    if _looks_like_header(rows[0], rows[1:]):
        return [Table(path.stem, [str(h or "") for h in rows[0]], rows[1:])]
    n = max(len(r) for r in rows)
    return [Table(path.stem, [""] * n, rows, has_header=False)]


def load_xlsx(path: Path) -> list[Table]:
    if load_workbook is None:
        sys.exit("openpyxl is required for .xlsx files: pip install openpyxl")
    wb = load_workbook(path, read_only=True, data_only=True)
    tables = []
    for ws in wb.worksheets:
        rows: list[Row] = [list(r) for r in ws.iter_rows(values_only=True)]
        rows = _trim_table(rows)
        if not rows:
            tables.append(Table(ws.title, [], []))
            continue
        if _looks_like_header(rows[0], rows[1:]):
            tables.append(Table(ws.title, [str(h or "") for h in rows[0]], rows[1:]))
        else:
            n = max(len(r) for r in rows)
            tables.append(Table(ws.title, [""] * n, rows, has_header=False))
    wb.close()
    return tables


def load_file(path: Path) -> list[Table]:
    ext = path.suffix.lower()
    if ext in (".xlsx", ".xlsm"):
        return load_xlsx(path)
    if ext in (".csv", ".tsv", ".txt"):
        return load_csv(path)
    sys.exit(f"Unsupported file type: {path} (use .xlsx/.xlsm/.csv/.tsv)")

# ------------------------------------------------------- normalization

_NUM_RE = re.compile(r"^[+-]?[\d,]*\.?\d+([eE][+-]?\d+)?$")
_DATE_FORMATS = (
    "%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S",
    "%m/%d/%Y", "%d/%m/%Y", "%m/%d/%y", "%d-%b-%Y", "%d-%b-%y",
    "%Y/%m/%d", "%d.%m.%Y",
)


def to_number(v: Cell) -> float | None:
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = v.strip().replace(",", "")
        s = re.sub(r"^[$£€¥]", "", s)
        if _NUM_RE.match(s):
            try:
                return float(s)
            except ValueError:
                return None
    return None


def to_datetime(v: Cell) -> datetime | None:
    if isinstance(v, datetime):
        return v
    if isinstance(v, date):
        return datetime(v.year, v.month, v.day)
    if isinstance(v, str):
        s = v.strip()
        for fmt in _DATE_FORMATS:
            try:
                return datetime.strptime(s, fmt)
            except ValueError:
                pass
    return None


def display(v: Cell) -> str:
    if v is None:
        return ""
    if isinstance(v, datetime):
        if (v.hour, v.minute, v.second, v.microsecond) == (0, 0, 0, 0):
            return v.strftime("%Y-%m-%d")
        return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


# ------------------------------------------------------ cell comparison

EQUAL, UNIMPORTANT, IMPORTANT = "equal", "unimportant", "important"


@dataclass
class CompareOptions:
    numeric_tolerance: float = 0.0
    date_tolerance_days: float = 0.0
    ignore_case: bool = False
    ignore_whitespace: bool = False
    unimportant_cols: set[str] = field(default_factory=set)
    sort_rows: bool = False


def _text_equal(a: str, b: str, o: CompareOptions) -> bool:
    if o.ignore_case:
        a, b = a.casefold(), b.casefold()
    if o.ignore_whitespace:
        a, b = " ".join(a.split()), " ".join(b.split())
    return a == b


def compare_cells(a: Cell, b: Cell, col_name: str, o: CompareOptions) -> str:
    if a is None:
        a = ""
    if b is None:
        b = ""
    if a == b:
        return EQUAL

    force_unimportant = col_name in o.unimportant_cols

    na, nb = to_number(a), to_number(b)
    if na is not None and nb is not None:
        if abs(na - nb) <= o.numeric_tolerance:
            return UNIMPORTANT if force_unimportant else EQUAL
        return UNIMPORTANT if force_unimportant else IMPORTANT

    da, db = to_datetime(a), to_datetime(b)
    if da is not None and db is not None:
        if abs(da - db) <= timedelta(days=o.date_tolerance_days):
            return UNIMPORTANT if force_unimportant else EQUAL
        return UNIMPORTANT if force_unimportant else IMPORTANT

    sa, sb = display(a), display(b)
    if sa == sb:
        return EQUAL
    if _text_equal(sa, sb, o):
        return UNIMPORTANT
    return UNIMPORTANT if force_unimportant else IMPORTANT


# ------------------------------------------------------ alignment/diff

@dataclass
class CellDiff:
    col: str
    left: str
    right: str
    kind: str  # IMPORTANT | UNIMPORTANT


@dataclass
class RowResult:
    kind: str                # equal | different | left_only | right_only
    left_rownum: int | None
    right_rownum: int | None
    left: Row | None
    right: Row | None
    diffs: list[CellDiff] = field(default_factory=list)


@dataclass
class SheetResult:
    name: str
    results: list[RowResult]
    left_only_cols: list[str]
    right_only_cols: list[str]
    colmap: list[tuple[int, int, str]]   # (left_col_idx, right_col_idx, name)
    stats: dict = field(default_factory=dict)
    note: str = ""

    @property
    def compared_cols(self) -> list[str]:
        return [n for _, _, n in self.colmap]


def align_columns(left: Table, right: Table) -> tuple[list[tuple[int, int | None]], list[str], list[str]]:
    """Match BC 'Align by left name': keep left order, pair right cols by name."""
    pairs: list[tuple[int, int | None]] = []
    used_right: set[int] = set()
    right_lookup: dict[str, int] = {}
    for j, h in enumerate(right.headers):
        right_lookup.setdefault(h, j)
        right_lookup.setdefault(h.casefold(), j)
    for i, h in enumerate(left.headers):
        j = right_lookup.get(h, right_lookup.get(h.casefold()))
        if j is not None and j not in used_right:
            pairs.append((i, j))
            used_right.add(j)
        else:
            pairs.append((i, None))
    right_only = [right.headers[j] or f"col{j + 1}" for j in range(len(right.headers)) if j not in used_right]
    left_only = [left.headers[i] or f"col{i + 1}" for i, j in pairs if j is None]
    return pairs, left_only, right_only


def _key_tuple(table: Table, row: Row, key_idx: list[int]) -> tuple:
    return tuple(display(table.cell(row, i)) for i in key_idx)


def compare_sheet(left: Table, right: Table, key_cols: list[str],
                  o: CompareOptions) -> SheetResult:
    pairs, left_only_cols, right_only_cols = align_columns(left, right)
    compared = [(i, j, left.headers[i] or f"col{i + 1}")
                for i, j in pairs if j is not None]

    note = ""
    use_keys = bool(key_cols)
    if use_keys:
        missing = [k for k in key_cols
                   if k not in left.headers or k not in right.headers]
        if missing:
            use_keys = False
            note = (f"key column(s) {missing} missing on one side; "
                    "fell back to content alignment")

    lrows, rrows = list(left.rows), list(right.rows)
    if o.sort_rows and use_keys:
        lk = [left.headers.index(k) for k in key_cols]
        rk = [right.headers.index(k) for k in key_cols]
        lrows.sort(key=lambda r: _key_tuple(left, r, lk))
        rrows.sort(key=lambda r: _key_tuple(right, r, rk))

    # ---- row alignment
    aligned: list[tuple[int | None, int | None]] = []  # indexes into lrows/rrows
    if use_keys:
        lk = [left.headers.index(k) for k in key_cols]
        rk = [right.headers.index(k) for k in key_cols]
        from collections import defaultdict, deque
        right_by_key: dict[tuple, deque[int]] = defaultdict(deque)
        for j, r in enumerate(rrows):
            right_by_key[_key_tuple(right, r, rk)].append(j)
        consumed: set[int] = set()
        for i, lr in enumerate(lrows):
            k = _key_tuple(left, lr, lk)
            if right_by_key[k]:
                j = right_by_key[k].popleft()
                consumed.add(j)
                aligned.append((i, j))
            else:
                aligned.append((i, None))
        for j in range(len(rrows)):
            if j not in consumed:
                aligned.append((None, j))
    else:
        sig_l = [tuple(display(left.cell(r, i)) for i, _ in pairs) for r in lrows]
        sig_r = [tuple(display(right.cell(r, j)) for _, j in pairs if j is not None) for r in rrows]
        sm = difflib.SequenceMatcher(None, sig_l, sig_r, autojunk=False)
        for tag, a0, a1, b0, b1 in sm.get_opcodes():
            if tag == "equal":
                aligned += [(a0 + k, b0 + k) for k in range(a1 - a0)]
            elif tag == "delete":
                aligned += [(a0 + k, None) for k in range(a1 - a0)]
            elif tag == "insert":
                aligned += [(None, b0 + k) for k in range(b1 - b0)]
            else:  # replace — pair up rows so cells can be diffed
                n = min(a1 - a0, b1 - b0)
                aligned += [(a0 + k, b0 + k) for k in range(n)]
                aligned += [(a0 + n + k, None) for k in range(a1 - a0 - n)]
                aligned += [(None, b0 + n + k) for k in range(b1 - b0 - n)]

    # ---- cell comparison
    results: list[RowResult] = []
    stats = dict(equal=0, different=0, left_only=0, right_only=0,
                 important=0, unimportant=0)
    for li, rj in aligned:
        if rj is None:
            results.append(RowResult("left_only", li + 2 if left.has_header else li + 1,
                                     None, lrows[li], None))
            stats["left_only"] += 1
            continue
        if li is None:
            results.append(RowResult("right_only", None,
                                     rj + 2 if right.has_header else rj + 1,
                                     None, rrows[rj]))
            stats["right_only"] += 1
            continue
        diffs: list[CellDiff] = []
        for i, j, name in compared:
            kind = compare_cells(left.cell(lrows[li], i), right.cell(rrows[rj], j), name, o)
            if kind != EQUAL:
                diffs.append(CellDiff(name, display(left.cell(lrows[li], i)),
                                      display(right.cell(rrows[rj], j)), kind))
                stats["important" if kind == IMPORTANT else "unimportant"] += 1
        status = "different" if diffs else "equal"
        stats[status] += 1
        results.append(RowResult(status,
                                 li + 2 if left.has_header else li + 1,
                                 rj + 2 if right.has_header else rj + 1,
                                 lrows[li], rrows[rj], diffs))
    return SheetResult(left.name, results, left_only_cols, right_only_cols,
                       compared, stats, note)


# ------------------------------------------------------------ reports

def print_summary(sheet_results, unmatched_left, unmatched_right):
    total = dict(equal=0, different=0, left_only=0, right_only=0,
                 important=0, unimportant=0)
    for sr in sheet_results:
        for k in total:
            total[k] += sr.stats[k]
        print(f"\nSheet: {sr.name}")
        print(f"  rows same.............. {sr.stats['equal']}")
        print(f"  rows with differences.. {sr.stats['different']} "
              f"({sr.stats['important']} important, {sr.stats['unimportant']} unimportant cells)")
        print(f"  left-only rows......... {sr.stats['left_only']}")
        print(f"  right-only rows........ {sr.stats['right_only']}")
        if sr.left_only_cols:
            print(f"  left-only columns...... {', '.join(sr.left_only_cols)}")
        if sr.right_only_cols:
            print(f"  right-only columns..... {', '.join(sr.right_only_cols)}")
        if sr.note:
            print(f"  note: {sr.note}")
    for n in unmatched_left:
        print(f"\nSheet only in LEFT : {n}")
    for n in unmatched_right:
        print(f"Sheet only in RIGHT: {n}")
    print("\n=== TOTAL ===")
    print(f"  rows same: {total['equal']}   rows different: {total['different']}   "
          f"left-only: {total['left_only']}   right-only: {total['right_only']}")
    print(f"  cell diffs: {total['important']} important, {total['unimportant']} unimportant")


def write_csv_diff(path: Path, sheet_results, key_cols):
    with path.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["sheet", "left_row", "right_row", "row_status",
                    "column", "left_value", "right_value", "severity"])
        for sr in sheet_results:
            for r in sr.results:
                if r.kind == "left_only":
                    w.writerow([sr.name, r.left_rownum, "", "left_only", "",
                                display_row(r.left), "", "important"])
                elif r.kind == "right_only":
                    w.writerow([sr.name, "", r.right_rownum, "right_only", "",
                                "", display_row(r.right), "important"])
                else:
                    for d in r.diffs:
                        w.writerow([sr.name, r.left_rownum, r.right_rownum,
                                    "different", d.col, d.left, d.right, d.kind])


def display_row(row):
    return " | ".join(display(v) for v in (row or []))


def write_html(path: Path, left_name, right_name, sheet_results,
               unmatched_left, unmatched_right, key_cols):
    COLORS = {"important": "#ffb3b3", "unimportant": "#b3d9ff",
              "left_only": "#ffe0b3", "right_only": "#b3ffd9"}
    parts = [f"""<!doctype html><meta charset="utf-8"><style>
body{{font-family:Menlo,monospace;font-size:13px;margin:20px}}
table{{border-collapse:collapse;margin-bottom:30px}}
td,th{{border:1px solid #ccc;padding:3px 8px;white-space:nowrap;max-width:400px;overflow:hidden;text-overflow:ellipsis}}
th{{background:#eee}}
.imp{{background:{COLORS['important']}}}.unimp{{background:{COLORS['unimportant']}}}
.lo{{background:{COLORS['left_only']}}}.ro{{background:{COLORS['right_only']}}}
.legend span{{padding:2px 10px;margin-right:10px;border:1px solid #999}}
h2{{margin-bottom:4px}}small{{color:#666}}
</style><h1>Beyond-Compare-style diff report</h1>
<small>LEFT: {html.escape(left_name)} &nbsp;|&nbsp; RIGHT: {html.escape(right_name)}
&nbsp;|&nbsp; keys: {html.escape(', '.join(key_cols) or '(positional alignment)')}</small>
<p class="legend">
<span class="imp">important difference</span><span class="unimp">unimportant difference</span>
<span class="lo">left-only row</span><span class="ro">right-only row</span></p>"""]
    for n in unmatched_left:
        parts.append(f"<h2>Sheet only in LEFT: {html.escape(n)}</h2>")
    for n in unmatched_right:
        parts.append(f"<h2>Sheet only in RIGHT: {html.escape(n)}</h2>")
    for sr in sheet_results:
        shown = [r for r in sr.results if r.kind != "equal"]
        parts.append(f"<h2>{html.escape(sr.name)}</h2><small>"
                     f"same {sr.stats['equal']} &middot; different {sr.stats['different']} "
                     f"&middot; left-only {sr.stats['left_only']} &middot; right-only {sr.stats['right_only']}"
                     f"{' &middot; left-only cols: ' + html.escape(', '.join(sr.left_only_cols)) if sr.left_only_cols else ''}"
                     f"{' &middot; right-only cols: ' + html.escape(', '.join(sr.right_only_cols)) if sr.right_only_cols else ''}"
                     f"</small>")
        if not shown:
            parts.append("<p>no differences</p>")
            continue
        parts.append("<table><tr><th>row&nbsp;L</th><th>row&nbsp;R</th>"
                     + "".join(f"<th>{html.escape(n)}</th>" for _, _, n in sr.colmap) + "</tr>")
        diff_by_col = {}
        for r in shown:
            cls = {"left_only": "lo", "right_only": "ro"}.get(r.kind, "")
            diff_by_col = {d.col: d for d in r.diffs}
            tds = []
            for li, rj, name in sr.colmap:
                if r.kind == "left_only":
                    tds.append(f"<td class='lo'>{html.escape(display(_get(r.left, li)))}</td>")
                elif r.kind == "right_only":
                    tds.append(f"<td class='ro'>{html.escape(display(_get(r.right, rj)))}</td>")
                else:
                    d = diff_by_col.get(name)
                    if d is None:
                        tds.append(f"<td>{html.escape(display(_get(r.left, li)))}</td>")
                    else:
                        k = "imp" if d.kind == IMPORTANT else "unimp"
                        tds.append(f"<td class='{k}'>{html.escape(d.left)} &#8596; {html.escape(d.right)}</td>")
            parts.append(f"<tr class='{cls}'><td>{r.left_rownum or ''}</td><td>{r.right_rownum or ''}</td>"
                         + "".join(tds) + "</tr>")
        parts.append("</table>")
    path.write_text("\n".join(parts), encoding="utf-8")


def _get(row, i):
    return row[i] if row is not None and i < len(row) else None


# ----------------------------------------------------------------- main

def main(argv=None):
    p = argparse.ArgumentParser(
        description="Beyond-Compare-style Table Compare for Excel/CSV files.")
    p.add_argument("left", type=Path)
    p.add_argument("right", type=Path)
    p.add_argument("--key", "-k", help="comma-separated key column name(s) that "
                   "identify a row (BC: mark column as Key). Without keys, rows "
                   "are aligned by content like BC's standard/Myers alignment.")
    p.add_argument("--sheet", help="compare only this sheet (Excel)")
    p.add_argument("--numeric-tolerance", "-nt", type=float, default=0.0,
                   help="numbers within this tolerance are equal")
    p.add_argument("--date-tolerance-days", "-dt", type=float, default=0.0,
                   help="dates within this many days are equal")
    p.add_argument("--ignore-case", "-ic", action="store_true",
                   help="case-only differences are unimportant")
    p.add_argument("--ignore-whitespace", "-iw", action="store_true",
                   help="whitespace-only differences are unimportant")
    p.add_argument("--unimportant-cols", "-uc", default="",
                   help="comma-separated columns whose diffs are never important")
    p.add_argument("--sort-rows", action="store_true",
                   help="sort rows by key before aligning (BC 'Sort rows')")
    p.add_argument("--no-header", action="store_true",
                   help="treat first line as data, not column names")
    p.add_argument("--report", type=Path, help="write HTML report to this path")
    p.add_argument("--csv-diff", type=Path, help="write machine-readable CSV diff")
    args = p.parse_args(argv)

    opts = CompareOptions(
        numeric_tolerance=args.numeric_tolerance,
        date_tolerance_days=args.date_tolerance_days,
        ignore_case=args.ignore_case,
        ignore_whitespace=args.ignore_whitespace,
        unimportant_cols={c.strip() for c in args.unimportant_cols.split(",") if c.strip()},
        sort_rows=args.sort_rows,
    )
    key_cols = [c.strip() for c in args.key.split(",")] if args.key else []

    left_tables = load_file(args.left)
    right_tables = load_file(args.right)
    if args.no_header:
        for t in left_tables + right_tables:
            t.rows.insert(0, t.headers)
            t.headers = [""] * max(1, t.ncols)
            t.has_header = False
    if args.sheet:
        left_tables = [t for t in left_tables if t.name == args.sheet]
        right_tables = [t for t in right_tables if t.name == args.sheet]
        if not left_tables or not right_tables:
            sys.exit(f"Sheet '{args.sheet}' not found on both sides")

    # align sheets by name (unmatched sheets reported); a lone sheet on
    # each side pairs directly (e.g. CSV vs CSV, or --sheet selection)
    if len(left_tables) == 1 and len(right_tables) == 1:
        pairs = [(left_tables[0], right_tables[0])]
        unmatched_left = unmatched_right = []
    else:
        right_by_name = {t.name: t for t in right_tables}
        pairs, unmatched_left = [], []
        for lt in left_tables:
            rt = right_by_name.pop(lt.name, None)
            if rt is None:
                unmatched_left.append(lt.name)
            else:
                pairs.append((lt, rt))
        unmatched_right = [t.name for t in right_by_name.values()]

    results = [compare_sheet(lt, rt, key_cols, opts) for lt, rt in pairs]
    print_summary(results, unmatched_left, unmatched_right)

    if args.csv_diff:
        write_csv_diff(args.csv_diff, results, key_cols)
        print(f"\nCSV diff written to {args.csv_diff}")
    if args.report:
        write_html(args.report, str(args.left), str(args.right), results,
                   unmatched_left, unmatched_right, key_cols)
        print(f"HTML report written to {args.report}")


if __name__ == "__main__":
    main()
