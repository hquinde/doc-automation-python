import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
from openpyxl.utils import get_column_letter

# activate virtual environment: .\venv\Scripts\Activate.ps1
# Phase1, hardcode automation, then for phase2 we automate the templates as well and figure out the equaton...


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

# template name : raw name
sample_map = {
    "QCS" : "QCS 200PPB 1%NO3",
    "Ca Chk 500 ppb" : "Ca check 500% NO3",
    "MDL" : "MDL",
    "CCV1 200 ppb" : "CCV1 200 ppb",
    "CCV2" : "CCV2",
    "CCV3" : "CCV3",
    "CCV4" : "CCV4",
    "CCV5" : "CCV5",
    "CCV6" : "CCV6",
    "CCV7" : "CCV7",
    "CCV8" : "CCV8",
    "CCV9" : "CCV9",
    "CCV10" : "CCV10",
    "CCV11" : "CCV11",

    "QCB" : "QCB",
#Blk will have unique code to find the data, since it requires getting the "Rinse" after QCB
    "CCB1" : "CCB1",
    "CCB2" : "CCB2",
    "CCB3" : "CCB3",
    "CCB4" : "CCB4",
    "CCB5" : "CCB5",
    "CCB6" : "CCB6",
    "CCB7" : "CCB7",
    "CCB8" : "CCB8",
    "CCB9" : "CCB9",
    "CCB10" : "CCB10",
    "CCB11" : "CCB11",

    "LCB1" : "LCB1",
    "LCB2" : "LCB2",
    "LCB3" : "LCB3",
    "LCB4" : "LCB4",
    "LCB5" : "LCB5"
}









#--------------------------------------------------------------------------------------------------
# Calibration
#--------------------------------------------------------------------------------------------------

df = pd.read_excel(file_path, sheet_name=sheet_name)

filtered_df = df[df["Sample Id"].isin(std_names)]

calibration_df = filtered_df.drop_duplicates(subset="Sample Id", keep="first")

element_data = calibration_df[element_columns].values.tolist()  # this is a list of lists



# Load the template workbook and the "Calibration" sheet
template_path = "Template Report.xlsx"
output_path = f"Report_Output.xlsx"  # Name of your output file

wb = load_workbook(template_path)
ws = wb["Calibration"]

# Write the element data starting from cell B2
start_row = 2
start_col = 2  # B is column 2

for i, row in enumerate(element_data):  # element_data is your list of lists
    for j, value in enumerate(row):
        ws.cell(row=start_row + i, column=start_col + j, value=value)



#--------------------------------------------------------------------------------------------------

# Find first "Rinse" row after "QCB"
sample_ids = df["Sample Id"].tolist()
blk_row = None

for i, sid in enumerate(sample_ids):
    if sid == "QCB":
        for j in range(i + 1, len(sample_ids)):
            if sample_ids[j] == "Rinse":
                blk_row = df.iloc[j]
                break
        break

# Extract the 24 element columns for blk
blk_values = blk_row[element_columns].tolist() if blk_row is not None else [None] * len(element_columns)




# Extracted data for the sample_map
qaqc_data = []  # List to hold all extracted rows in order
sample_names = list(sample_map.keys())  # Ordered keys from sample_map


for template_name, raw_sample_name in sample_map.items():
    match = df[df["Sample Id"] == raw_sample_name]
    if not match.empty:
        values = match[element_columns].iloc[0].tolist()
        qaqc_data.append(values)
    else:
        # If the sample isn't found, fill with blank values
        print(f"Warning: '{raw_sample_name}' not found in Concentrations")
        qaqc_data.append([None] * len(element_columns))

# Find index of "QCB" in the sample order
qcb_index = sample_names.index("QCB")

# Insert BLK values right after that
qaqc_data.insert(qcb_index + 1, blk_values)

# Also insert "Blk" into the sample_names list so it aligns with qaqc_data
sample_names.insert(qcb_index + 1, "Blk")


print(qaqc_data)

# ws = wb["QAQC"]

# # Define starting cell (we begin at row 3, column B)
# start_row = 3
# start_col = 2  # Column B is index 2

# # Write only the 24 element values (one row per sample)
# for i, values in enumerate(qaqc_data):
#     for j, val in enumerate(values):
#         ws.cell(row=start_row + i, column=start_col + j, value=val)

# # Save the updated workbook
# wb.save(output_path)
# print(f"Data written to {output_path}")
