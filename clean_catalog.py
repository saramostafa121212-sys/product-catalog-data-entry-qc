"""Product Catalog Data Entry & QC
Transfers product records from CSV into a formatted Excel catalog,
keeping barcodes as TEXT (leading zeros preserved) and adding a QC sheet.
"""
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment

SRC = "data/source_products.csv"
OUT = "data/product_catalog.xlsx"
ATTRIBUTION = "Demo data created for this project (no third-party data)."

# Read EVERYTHING as text so barcodes keep their leading zeros
df = pd.read_csv(SRC, dtype=str)
df.columns = ["Product ID", "Barcode", "Name", "Brand", "Category", "Quantity", "Country"]

# Python-side QC (printed in the console)
print("Records:", len(df))
print("Duplicate barcodes:", int(df["Barcode"].duplicated().sum()))
print("Missing values:", int(df.isna().sum().sum()))

n = len(df) + 1  # last data row in Excel

with pd.ExcelWriter(OUT, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name="Catalog")
    ws = writer.sheets["Catalog"]
    for cell in ws[1]:
        cell.font = Font(name="Arial", bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center")
    for row in ws.iter_rows(min_row=2, max_row=n):
        for cell in row:
            cell.font = Font(name="Arial")
        row[1].number_format = "@"  # Barcode column = Text
    for col, w in zip("ABCDEFG", [12, 18, 28, 18, 16, 12, 12]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"

    qc = writer.book.create_sheet("QC")
    rows = [
        ("QC Check", "Result", "Status"),
        ("Total records", f"=COUNTA(Catalog!A2:A{n})", f'=IF(B2={len(df)},"PASS","FAIL")'),
        ("Duplicate barcodes", f"=SUMPRODUCT((COUNTIF(Catalog!B2:B{n},Catalog!B2:B{n})>1)*1)", '=IF(B3=0,"PASS","FAIL")'),
        ("Missing values", f"=COUNTBLANK(Catalog!A2:G{n})", '=IF(B4=0,"PASS","FAIL")'),
        ("Barcode length = 13", f"=SUMPRODUCT((LEN(Catalog!B2:B{n})<>13)*1)", '=IF(B5=0,"PASS","FAIL")'),
        ("", "", ""),
        ("Attribution", ATTRIBUTION, ""),
    ]
    for r in rows:
        qc.append(r)
    for cell in qc[1]:
        cell.font = Font(name="Arial", bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    for row in qc.iter_rows(min_row=2):
        for cell in row:
            cell.font = Font(name="Arial")
    qc.column_dimensions["A"].width = 24
    qc.column_dimensions["B"].width = 14
    qc.column_dimensions["C"].width = 10

print("Done:", OUT)
