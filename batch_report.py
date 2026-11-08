from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

INPUT_DIR = Path("input")
OUTPUT_DIR = Path("output")
REQUIRED = ["product", "price", "qty"]


def to_number(value):
    """Convert a value to a number, or return None if impossible."""
    if value is None or str(value).strip() == "":
        return None
    try:
        return float(str(value).replace(",", "").strip())
    except ValueError:
        return None


def process_file(path):
    """Process one Excel file. Returns (valid_count, issue_count) or None on failure."""
    # 1) Open the file
    try:
        wb_in = openpyxl.load_workbook(path)
    except Exception as e:
        print(f"  ERROR: could not open file: {e}")
        return None

    rows = list(wb_in.active.iter_rows(values_only=True))
    if len(rows) < 2:
        print("  ERROR: the file has no data rows.")
        return None

    # 2) Check required columns
    header = [str(h).strip().lower() if h is not None else "" for h in rows[0]]
    missing = [c for c in REQUIRED if c not in header]
    if missing:
        print(f"  ERROR: missing columns: {missing}. Found: {header}")
        return None
    pi, pr, qi = (header.index(c) for c in REQUIRED)
    n_cols = max(i for i, h in enumerate(header) if h) + 1

    # 3) Process rows and record problems
    good, issues = [], []
    for line_no, row in enumerate(rows[1:], start=2):
        if all(v is None or str(v).strip() == "" for v in row):
            continue  # completely empty row
        extra = [v for v in row[n_cols:] if v is not None and str(v).strip() != ""]
        if extra:
            issues.append((line_no, f"Extra values found: {extra} (is a number written like 2,600?)"))
            continue
        name = row[pi]
        price = to_number(row[pr])
        qty = to_number(row[qi])
        if name is None or str(name).strip() == "":
            issues.append((line_no, "Product name is empty"))
        elif price is None:
            issues.append((line_no, f"Invalid price: {row[pr]!r}"))
        elif qty is None:
            issues.append((line_no, f"Invalid quantity: {row[qi]!r}"))
        else:
            good.append((str(name).strip().title(), price, qty))

    # 4) Write the report
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Processed Sales"
    ws.append(["Product", "Price", "Qty", "Total"])
    for name, price, qty in good:
        r = ws.max_row + 1
        ws.append([name, price, qty, f"=B{r}*C{r}"])
    last = ws.max_row
    ws.append(["Grand Total", None, None, f"=SUM(D2:D{last})"])
    total_row = ws.max_row

    fill = PatternFill("solid", start_color="1F4E78")
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = fill
        c.alignment = Alignment(horizontal="center")
    for r in range(2, total_row + 1):
        ws[f"B{r}"].number_format = "#,##0.00"
        ws[f"D{r}"].number_format = "#,##0.00"
    for c in ws[total_row]:
        c.font = Font(bold=True)
    ws.column_dimensions["A"].width = 16
    for col in "BCD":
        ws.column_dimensions[col].width = 12

    # Summary sheet (only if there is valid data)
    if good:
        sm = wb.create_sheet("Summary", 1)
        rng = lambda col: f"'Processed Sales'!{col}2:{col}{last}"
        sm.append(["Metric", "Value"])
        sm.append(["Total sales", f"=SUM({rng('D')})"])
        sm.append(["Number of products", f"=COUNTA({rng('A')})"])
        sm.append(["Total units sold", f"=SUM({rng('C')})"])
        sm.append(["Average sale per product", f"=AVERAGE({rng('D')})"])
        sm.append(["Top product (by sales)",
                   f"=INDEX({rng('A')},MATCH(MAX({rng('D')}),{rng('D')},0))"])
        sm.append(["Top product sales", f"=MAX({rng('D')})"])
        for c in sm[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = fill
        sm["B2"].number_format = "#,##0.00"
        sm["B5"].number_format = "#,##0.00"
        sm["B7"].number_format = "#,##0.00"
        sm.column_dimensions["A"].width = 28
        sm.column_dimensions["B"].width = 18

    # 5) Issues sheet
    if issues:
        wi = wb.create_sheet("Issues")
        wi.append(["Row in source file", "Problem"])
        for line_no, msg in issues:
            wi.append([line_no, msg])
        for c in wi[1]:
            c.font = Font(bold=True)
        wi.column_dimensions["A"].width = 18
        wi.column_dimensions["B"].width = 60

    # 6) Save
    out_path = OUTPUT_DIR / f"processed_{path.name}"
    try:
        wb.save(out_path)
    except PermissionError:
        print(f"  ERROR: close {out_path} in Excel, then run again.")
        return None

    for line_no, msg in issues:
        print(f"  - Row {line_no}: {msg}")
    print(f"  Saved: {out_path}")
    return len(good), len(issues)


def main():
    INPUT_DIR.mkdir(exist_ok=True)
    OUTPUT_DIR.mkdir(exist_ok=True)

    # skip Excel temporary lock files that start with "~$"
    files = sorted(p for p in INPUT_DIR.glob("*.xlsx") if not p.name.startswith("~$"))
    if not files:
        print(f"No .xlsx files found in '{INPUT_DIR}'. Put your files there and run again.")
        return

    ok = failed = 0
    for path in files:
        print(f"Processing {path.name} ...")
        result = process_file(path)
        if result is None:
            failed += 1
        else:
            ok += 1
            print(f"  Done: {result[0]} valid rows, {result[1]} rows with problems.")

    print(f"\nFinished: {ok} file(s) processed, {failed} failed.")


main()