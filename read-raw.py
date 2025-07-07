import pandas as pd
from openpyxl import load_workbook
from datetime import datetime

# Load your raw ICP-MS Excel file
raw_path = "AS250327_HB2412 ICPMS Raw Data.xlsx"
df = pd.read_excel(raw_path)

# Filter the calibration standards
std_names = ["Std1-5ppb", "Std2-20ppb", "Std3-50ppb", "Std4-200ppb", "Std5-500ppb"]
calibration_rows = df[df["Sample Id"].isin(std_names)].drop_duplicates(subset="Sample Id", keep="first")

# Keep only columns from "Na 23\n(cps)" onward
cols = calibration_rows.columns.tolist()
start_idx = cols.index("Na 23\n(cps)")
element_cols = cols[start_idx:]
calibration_df = calibration_rows[element_cols]

# Load your Excel template (DONT overwrite it)
wb = load_workbook("ICPMS Template Report2.xlsx")
ws = wb["Calibration"]  # Make sure this matches the tab name

# Automatically map Excel headers to column numbers
excel_headers = {}
for col in range(2, ws.max_column + 1):  # Start at col 2 (B)
    header = ws.cell(row=1, column=col).value
    if header in element_cols:
        excel_headers[header] = col

# Write calibration values into template
for element in calibration_df.columns:
    if element in excel_headers:
        col = excel_headers[element]
        for i, val in enumerate(calibration_df[element]):
            row = 3 + i  # Write from row 3 (Std1)
            ws.cell(row=row, column=col).value = val

# Saved to a new file
today = datetime.today().strftime("%Y-%m-%d")
output_path = f"ICPMS_Report_{today}.xlsx"
wb.save(output_path)

print(f"Saved new report as: {output_path}")