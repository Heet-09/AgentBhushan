"""Parse an Excel workbook and print every sheet.

The header row is detected automatically (it does not have to be row 1). Rows above
the header (tester, version, counts...) are printed as a summary, single-cell rows
inside the table are printed as section titles, and every other row is printed as a
block of "Header : value" lines so long multi-line cells stay readable.
Empty cells are skipped.

Usage:
    python parse_excel.py <path-to-workbook.xlsx> [--sheet NAME ...] [--max-rows N] [-o out.txt]
    python parse_excel.py <path-to-workbook.xlsx> --failed [-o failed.json | failed.xlsx]
"""

import argparse
import json
import sys
import textwrap

import pandas as pd

WIDTH = 100
HEADER_SCAN_ROWS = 30  # how many rows from the top to search for the header row
MIN_HEADER_CELLS = 3   # a header row needs at least this many filled text cells


def load_sheets(path, sheet_names=None):
    """Return a dict of {sheet_name: raw DataFrame} with no header applied."""
    return pd.read_excel(path, sheet_name=sheet_names or None, header=None)


def is_empty(value):
    return pd.isna(value) or str(value).strip() == ""


def cell_text(value):
    """Render a cell value, dropping the '.0' pandas adds to whole numbers."""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def find_header_row(df):
    """Return the index of the row with the most filled text cells, or None if no row qualifies."""
    best_index, best_score = None, MIN_HEADER_CELLS - 1
    for index, row in df.head(HEADER_SCAN_ROWS).iterrows():
        score = sum(isinstance(v, str) and v.strip() != "" for v in row)
        if score > best_score:
            best_index, best_score = index, score
    return best_index


def format_value(value, indent):
    """Wrap each line of a cell value to WIDTH, indenting continuation lines."""
    lines = []
    for line in cell_text(value).splitlines() or [""]:
        lines.extend(textwrap.wrap(line.strip(), WIDTH - indent) or [""])
    return ("\n" + " " * indent).join(lines)


def get_rows(df):
    """Return every table row of a raw sheet as a dict {header: value} (empty cells left out)."""
    df = df.dropna(how="all").dropna(axis=1, how="all")
    header_index = find_header_row(df)
    if header_index is None:
        return []

    header = df.loc[header_index]
    columns = [(col, cell_text(header[col])) for col in df.columns if not is_empty(header[col])]
    records = []
    for index, row in df.loc[header_index + 1 :].iterrows():
        if sum(not is_empty(v) for v in row) <= 1:
            continue  # blank row or a section title
        record = {label: cell_text(row[col]) for col, label in columns if not is_empty(row[col])}
        record["Excel row"] = index + 1
        records.append(record)
    return records


def get_failed_rows(sheets, column="Chrome", value="Fail"):
    """Return a list of dicts, one per row whose `column` equals `value`, with all the row's data."""
    failed = []
    for name, df in sheets.items():
        for record in get_rows(df):
            if record.get(column, "").strip().lower() == value.lower():
                failed.append({"Sheet": name.strip(), **record})
    return failed


def print_key_values(rows):
    """Print rows like 'Tester | Yogesh' as 'Tester : Yogesh' (first cell is the key)."""
    pairs = []
    for _, row in rows.iterrows():
        cells = [cell_text(v) for v in row if not is_empty(v)]
        if cells:
            pairs.append((cells[0], " | ".join(cells[1:])))
    if not pairs:
        return
    key_width = max(len(key) for key, _ in pairs)
    for key, value in pairs:
        print(f"{key.ljust(key_width)} : {format_value(value, key_width + 3)}")


def print_sheet(name, df, max_rows=None):
    df = df.dropna(how="all").dropna(axis=1, how="all")
    print("=" * WIDTH)
    print(f" Sheet: {name.strip()}")
    print("=" * WIDTH)
    if df.empty:
        print("(empty sheet)\n")
        return

    header_index = find_header_row(df)
    if header_index is None:
        # No table on this sheet, just label/value pairs
        print_key_values(df)
        print()
        return

    preamble = df.loc[: header_index - 1]
    if not preamble.empty:
        print_key_values(preamble)

    header = df.loc[header_index]
    table = df.loc[header_index + 1 :]
    columns = [(col, cell_text(header[col])) for col in df.columns if not is_empty(header[col])]
    label_width = max(len(label) for _, label in columns)
    indent = label_width + 3

    print(f"\nColumns: {', '.join(label for _, label in columns)}")
    print(f"Rows:    {len(table)}")

    shown = 0
    for index, row in table.iterrows():
        if max_rows is not None and shown >= max_rows:
            print(f"\n... {len(table) - shown} more rows not shown")
            break
        filled = [cell_text(v) for v in row if not is_empty(v)]
        if len(filled) == 1:
            # A lone cell inside the table is a section title like "SCAD-2 Search filters"
            print(f"\n### {filled[0]}")
            continue
        # +1 because pandas rows are 0-based and Excel rows are 1-based
        print(f"\n--- Excel row {index + 1} " + "-" * (WIDTH - 18))
        for col, label in columns:
            if not is_empty(row[col]):
                print(f"{label.ljust(label_width)} : {format_value(row[col], indent)}")
        shown += 1
    print()


def main():
    parser = argparse.ArgumentParser(description="Parse an Excel file and print all its sheets.")
    parser.add_argument("path", help="Path to the .xlsx/.xls file")
    parser.add_argument("--sheet", action="append", help="Only print this sheet (repeatable)")
    parser.add_argument("--max-rows", type=int, default=None, help="Limit rows printed per sheet")
    parser.add_argument("--failed", action="store_true",
                        help='Print a list of dicts (JSON) for every row whose "Chrome" column is "Fail"')
    parser.add_argument("-o", "--output", help="Write the output to this text file (UTF-8) instead of the screen")
    args = parser.parse_args()

    try:
        sheets = load_sheets(args.path, args.sheet)
    except FileNotFoundError:
        sys.exit(f"File not found: {args.path}")
    except ValueError as e:
        sys.exit(f"Could not read workbook: {e}")

    if args.failed and args.output and args.output.lower().endswith(".xlsx"):
        failed = get_failed_rows(sheets)
        pd.DataFrame(failed).to_excel(args.output, index=False)
        print(f"{len(failed)} failed rows written to {args.output}")
        return

    if args.output:
        sys.stdout = open(args.output, "w", encoding="utf-8")

    if args.failed:
        failed = get_failed_rows(sheets)
        print(json.dumps(failed, indent=2, ensure_ascii=False))
        print(f"{len(failed)} failed rows", file=sys.stderr)
    else:
        print_all(args, sheets)

    if args.output:
        sys.stdout.close()
        sys.stdout = sys.__stdout__
        print(f"Output written to {args.output}")


def print_all(args, sheets):
    print(f"Workbook: {args.path}")
    print(f"Sheets: {', '.join(name.strip() for name in sheets)}\n")
    for name, df in sheets.items():
        print_sheet(name, df, args.max_rows)


if __name__ == "__main__":
    main()
