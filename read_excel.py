import pandas as pd

# Load all sheets in the Excel file into a dictionary
sheets = pd.read_excel("AS250204_HB2403 ICPMS.xlsx", sheet_name=None)

# Loop through each sheet and print its name and the first 5 rows

count_sheets = 0

for sheet_name, df in sheets.items():
    print(f"\n--- Sheet: {sheet_name} ---")
    print(df.head())
    count_sheets += 1

print(count_sheets)


