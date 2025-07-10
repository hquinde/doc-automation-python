import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
from openpyxl.utils import get_column_letter


# read excel file
file_path = 'Raw Data.xlsx'
sheet_name = 'Concentrations'




std_names = [
    'Std1-5ppb',
    'Std2-20ppb',
    'Std3-50ppb',
    'Std4-200ppb',
    'Std5-500ppb',
    'Std6-2000ppb'
]

element_columns = [
    'Na 23\n(ug/L)',
    'Mg 24\n(ug/L)',
    'Al 27\n(ug/L)',
    'Si 28\n(ug/L)',
    'P 31\n(ug/L)',
    'K 39\n(ug/L)',
    'Ca-43 43\nHelium KED\n(ug/L)',
    'Ca-43std 43\n(ug/L)',
    'Ca-44 44\nHelium KED\n(ug/L)',
    'Ca-44std 44\n(ug/L)',
    'Mn 55\n(ug/L)',
    'Fe 57\n(ug/L)',
    'Co 59\n(ug/L)',
    'Ni 60\n(ug/L)',
    'Cu 63\n(ug/L)',
    'Zn 68\n(ug/L)',
    'Se 78\n(ug/L)',
    'Se 82\n(ug/L)',
    'Sr 88\n(ug/L)',
    'Mo 96\n(ug/L)',
    'Cd 113\n(ug/L)',
    'Pb 206\n(ug/L)',
    'Pb 207\n(ug/L)',
    'Pb 208\n(ug/L)'
]


df = pd.read_excel(file_path, sheet_name=sheet_name)

filtered_df = df[df["Sample Id"].isin(std_names)]

calibration_df = filtered_df.drop_duplicates(subset="Sample Id", keep="first")

element_data = calibration_df[element_columns].values.tolist()  # this is a list of lists



# Load the template workbook and the "Calibration" sheet
template_path = "Template Report.xlsx"
output_path = f"Calibration_Report_Output.xlsx"  # Name of your output file

wb = load_workbook(template_path)
ws = wb["Calibration"]

# Write the element data starting from cell B2
start_row = 2
start_col = 2  # B is column 2

for i, row in enumerate(element_data):  # element_data is your list of lists
    for j, value in enumerate(row):
        ws.cell(row=start_row + i, column=start_col + j, value=value)

# Save the updated workbook
wb.save(output_path)
print(f"Calibration data successfully written to {output_path}")

