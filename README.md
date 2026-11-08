# Excel Sales Report Automation

A Python tool that automatically turns raw sales spreadsheets into clean,
formatted Excel reports. Built with Python and openpyxl.

## What it does

- Reads every `.xlsx` file in the `input/` folder (batch processing)
- Calculates the total for each product (`Price x Qty`) using live Excel formulas
- Adds a formatted report sheet with a Grand Total row
- Adds a **Summary** sheet: total sales, number of products, units sold,
  average sale, and top product
- Detects bad data (empty cells, text instead of numbers, extra values,
  missing columns) and lists the problems in an **Issues** sheet
  instead of crashing

## Project structure

```
project/
├── batch_report.py   # main script
├── main.py           # creates sample data in input/
├── input/            # put your .xlsx files here
├── output/           # processed reports are saved here
└── README.md
```

## Requirements

- Python 3.10+
- openpyxl

```
python -m pip install openpyxl
```

## How to use

1. Put your sales files in the `input/` folder.
   Required columns: `Product`, `Price`, `Qty`.
2. Run:
   ```
   python batch_report.py
   ```
3. Open the reports from the `output/` folder.

## Example output

| Product  | Price    | Qty | Total     |
|----------|----------|-----|-----------|
| Laptop   | 2,600.00 | 5   | 13,000.00 |
| Mouse    | 300.00   | 10  | 3,000.00  |
| Keyboard | 420.00   | 9   | 3,780.00  |
| **Grand Total** |   |     | **19,780.00** |

## Need a custom Excel automation?

I build Python scripts that automate repetitive Excel work: reports,
data cleaning, merging files, and more.