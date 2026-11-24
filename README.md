# table-compare — local Beyond Compare "Table Compare" for Excel/CSV

Replicates the parts of Beyond Compare you use for `.xlsx` / `.csv`
comparison — two ways to run it:

## Option A — `compare.html` (browser UI, no setup)

Double-click `compare.html` (or drag it into a browser tab) — or open the
hosted copy: <https://rjarupula1729.github.io/table-compare/>.
100% client-side, zero dependencies — **your files never leave your machine**
(safe for OneDrive / company data). Supports `.xlsx`, `.xlsm`, `.csv`, `.tsv`.
Sheets and CSVs are **streamed while parsing**, so files with hundreds of
thousands of rows (sheet XML or CSV larger than JS's ~512M-char string
limit) load fine — tested to ~7M rows. The practical limit is how many
parsed rows fit in browser memory.

### Quick start

1. Drop your **LEFT** and **RIGHT** files onto the two panels.
2. The tool **auto-detects the key column(s)** that uniquely identify
   matching records and runs the comparison immediately — no buttons needed.
3. Scroll to the bottom: records are aligned and sorted by key, differences
   highlighted.

Everything is live: change any option, column mapping, or key and the
comparison re-runs by itself.

### The panels, top to bottom

**Options bar**

- *Sheet* — pick which sheet to configure/preview, or `All / auto` to
  compare every sheet (matched by name).
- *Numeric / Date tolerance* — differences within tolerance count as
  unimportant (blue), never red.
- *ignore case / ignore whitespace* — same idea for text cells.
- *no header row* — treat the first row as data (applies once, safely).
- *Export diff CSV* — the full machine-readable diff, every row.
- *Reset* — start over.

**Columns strip** (horizontal cards, scroll left→right)

- Each card pairs a LEFT column with its matched RIGHT column;
  use the dropdowns to remap columns whose headers differ.
- **key** — tick one or *more* columns; together they form the composite
  key that matches records across the files. Cards marked `auto` were
  detected automatically — untick/change freely, your choice wins.
- **unimp** — hides the column from both files entirely and excludes it
  from comparison (removes noise columns).
- A column sample is shown under each dropdown so you can see the data
  you're mapping.

**Data preview**

- LEFT and RIGHT file previews, side by side, with source row numbers.
- **Click any column header to toggle it as a key** — the easiest way to
  pick keys while looking at real values.
- Scrolling is synchronized: vertical keeps records level, horizontal
  keeps the same field aligned on both sides.
- After the compare runs, this same panel shows the aligned result
  (first 500 rows) — click any row to inspect that record below.

**Compare results**

- Stats: same / different / left-only / right-only rows, important vs
  unimportant cells, per sheet.
- Filter buttons: **All / Differences / Same**.
- The bottom section is a **record inspector**: one record at a time,
  top row = left file's record, bottom row = right file's, sharing one
  union header so fields line up column-for-column. Navigate with
  `◀ prev` / `next ▶` / `go to`, or click a row in the preview.
  Columns only present in one file appear once, blank on the other side.

**Group-by summary**

- Pivot-style reconciliation: pick **Sheets** (multiple — their rows are
  combined), **Group by** columns (multiple — composite groups), and
  **Sum** columns (auto-detected numerics).
- **Group-by order is yours**: click the *Group by* chips in the order you
  want to group — numbered badges show the hierarchy (`1` = outermost
  group). Untick and re-tick a chip to move it to the end.
- Every group shows `Σ field L | Σ field R | Δ` plus row counts and a
  TOTAL row. **Δ: red = mismatch** (the group key cell flags red too),
  **green = reconciled**. Tiny float noise on huge sums is ignored
  (~1e-9 of the magnitude), plus your numeric tolerance applies.
- **Download Excel** exports the summary as `.xlsx` with the same
  red/green coloring, one column per group-by field (great for sharing or
  auditing). **Download HTML** exports the colored table as a webpage.
  Both run fully in the browser — no libraries, works offline.

### Colors

| Color | Meaning |
|---|---|
| Red cell | important difference |
| Blue cell | unimportant difference (within tolerance / case / whitespace) |
| Red row wash + bold row number | the record doesn't match |
| Yellow column | key column(s) — the values that paired the records |
| Orange row | exists only in LEFT |
| Green row | exists only in RIGHT |
| Tinted column header + count | that column contains differences |

## Option B — `table_compare.py` (terminal / scripts)

### Setup (once)

```bash
cd ~/Library/CloudStorage/OneDrive-Pepsico/Documents/table-compare
python3 -m venv .venv          # already done
.venv/bin/pip install openpyxl # already done
```

### Usage

```bash
.venv/bin/python table_compare.py LEFT.xlsx RIGHT.xlsx --key Date,Product
.venv/bin/python table_compare.py LEFT.csv RIGHT.xlsx --key id \
    --numeric-tolerance 0.001 --ignore-case \
    --report diff.html --csv-diff diff.csv
```

## How it maps to Beyond Compare

| BC concept | This tool |
|---|---|
| File format conversion (Excel → text) | openpyxl reads each sheet; dates/numbers kept as typed values |
| Delimited parse, text qualifier, first-line header detect | CSV sniffer + header heuristic (`--no-header` to force data) |
| Sheets aligned by name | matched by sheet name; unmatched sheets reported |
| Columns "align by left name" | columns matched by header name (remappable in the UI) |
| Key columns for row alignment | `--key ColA,ColB` / key checkboxes & header clicks in the UI |
| Standard/Myers row alignment (no keys) | difflib LCS alignment |
| Numeric tolerance | `--numeric-tolerance N` / Numeric tolerance field |
| Date tolerance | `--date-tolerance-days N` / Date tolerance field |
| Case / whitespace unimportant | `--ignore-case` / `--ignore-whitespace` / checkboxes |
| Column marked unimportant | `--unimportant-cols` / `unimp` checkbox (hides the column) |
| Sort rows before alignment | automatic whenever keys are set |
| BC red/blue/orange/teal colors | same palette in UI and `--report diff.html` |
| Report output | console summary, `--csv-diff` (machine-readable), `--report` (HTML) |

## Output (Python tool)

- **Console** — per-sheet counts: same rows, rows with differences
  (important vs unimportant cells), left-only/right-only rows and columns.
- **`--csv-diff diff.csv`** — one line per difference:
  `sheet,left_row,right_row,row_status,column,left_value,right_value,severity`
- **`--report diff.html`** — BC-style colored grid: red = important cell
  diff, blue = unimportant diff, orange = left-only row, green =
  right-only row. Only differing rows are shown.

## Example

```bash
.venv/bin/python table_compare.py \
  ~/Downloads/KSA_Digital_merged_data_snacks.xlsx \
  ~/Library/CloudStorage/OneDrive-Pepsico/Desktop/KSA_Data/Output/KSA_Digital_snacks_data.xlsx \
  --key Date,Product --report ksa_diff.html
```

(Small synthetic fixtures — `test_left.xlsx`, `test_right.xlsx`, etc. —
are handy for a quick sanity check; they are kept locally, not in the repo.)
