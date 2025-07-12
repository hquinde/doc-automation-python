import pandas as pd
from openpyxl import load_workbook
from datetime import datetime

# activate virtual environment: .\venv\Scripts\Activate.ps1
# Phase1, hardcode automation, then for phase2 we automate the templates as well and figure out the equaton...
# get updates to current device: git pull origin main

#

# -----------------------------
# Reusable Writing Function
# -----------------------------
def write_sample_data_by_id(ws, sample_names, data_rows, element_columns, header_row=1, data_start_col=2):
    """
    Writes data to a worksheet by matching sample names in column A.
    """
    sample_row_map = {}
    for row in ws.iter_rows(min_row=header_row + 1, max_col=1):
        sample_id = row[0].value
        if sample_id:
            sample_row_map[sample_id.strip()] = row[0].row

    for sample_name, values in zip(sample_names, data_rows):
        target_row = sample_row_map.get(sample_name)
        if target_row:
            for j, val in enumerate(values):
                ws.cell(row=target_row, column=data_start_col + j, value=val)
        else:
            print(f"⚠️ Sample '{sample_name}' not found in sheet — skipping.")


# -----------------------------
# Input Setup
# -----------------------------
file_path = 'Raw Data.xlsx'
sheet_name = 'Concentrations'
template_path = 'Template Report.xlsx'
output_path = 'Report_Output.xlsx'

std_names = [
    'Std1-5ppb', 'Std2-20ppb', 'Std3-50ppb',
    'Std4-200ppb', 'Std5-500ppb', 'Std6-2000ppb'
]

element_columns = [
    'Na 23\n(ug/L)', 'Mg 24\n(ug/L)', 'Al 27\n(ug/L)', 'Si 28\n(ug/L)',
    'P 31\n(ug/L)', 'K 39\n(ug/L)', 'Ca-43 43\nHelium KED\n(ug/L)',
    'Ca-43std 43\n(ug/L)', 'Ca-44 44\nHelium KED\n(ug/L)', 'Ca-44std 44\n(ug/L)',
    'Mn 55\n(ug/L)', 'Fe 57\n(ug/L)', 'Co 59\n(ug/L)', 'Ni 60\n(ug/L)',
    'Cu 63\n(ug/L)', 'Zn 68\n(ug/L)', 'Se 78\n(ug/L)', 'Se 82\n(ug/L)',
    'Sr 88\n(ug/L)', 'Mo 96\n(ug/L)', 'Cd 113\n(ug/L)',
    'Pb 206\n(ug/L)', 'Pb 207\n(ug/L)', 'Pb 208\n(ug/L)'
]

sample_map = {
    "QCS": "QCS 200PPB 1%NO3",
    "Ca Chk 500 ppb": "Ca check 500% NO3",
    "MDL": "MDL",
    "CCV1 200 ppb": "CCV1 200 ppb",
    "CCV2": "CCV2", "CCV3": "CCV3", "CCV4": "CCV4", "CCV5": "CCV5",
    "CCV6": "CCV6", "CCV7": "CCV7", "CCV8": "CCV8", "CCV9": "CCV9",
    "CCV10": "CCV10", "CCV11": "CCV11",
    "QCB": "QCB",
    "CCB1": "CCB1", "CCB2": "CCB2", "CCB3": "CCB3", "CCB4": "CCB4", "CCB5": "CCB5",
    "CCB6": "CCB6", "CCB7": "CCB7", "CCB8": "CCB8", "CCB9": "CCB9", "CCB10": "CCB10", "CCB11": "CCB11",
    "LCB1": "LCB1", "LCB2": "LCB2", "LCB3": "LCB3", "LCB4": "LCB4", "LCB5": "LCB5"
}


# -----------------------------
# Read Data
# -----------------------------
df = pd.read_excel(file_path, sheet_name=sheet_name)

# -----------------------------
# CALIBRATION DATA
# -----------------------------
filtered_df = df[df["Sample Id"].isin(std_names)]
calibration_df = filtered_df.drop_duplicates(subset="Sample Id", keep="first")
element_data = calibration_df[element_columns].values.tolist()


# -----------------------------
# QAQC: Find BLK (Rinse after QCB)
# -----------------------------
sample_ids = df["Sample Id"].tolist()
blk_row = None

for i, sid in enumerate(sample_ids):
    if sid == "QCB":
        for j in range(i + 1, len(sample_ids)):
            if sample_ids[j] == "Rinse":
                blk_row = df.iloc[j]
                break
        break

blk_values = blk_row[element_columns].tolist() if blk_row is not None else [None] * len(element_columns)

# -----------------------------
# QAQC: Extract Values
# -----------------------------
qaqc_data = []
sample_names = list(sample_map.keys())

for template_name, raw_sample_name in sample_map.items():
    match = df[df["Sample Id"] == raw_sample_name]
    if not match.empty:
        values = match[element_columns].iloc[0].tolist()
        qaqc_data.append(values)
    else:
        print(f"Warning: '{raw_sample_name}' not found in Concentrations")
        qaqc_data.append([None] * len(element_columns))

# Insert BLK after QCB
qcb_index = sample_names.index("QCB")
sample_names.insert(qcb_index + 1, "Blk")
qaqc_data.insert(qcb_index + 1, blk_values)


# -----------------------------
# WRITE TO EXCEL
# -----------------------------
wb = load_workbook(template_path)
ws1 = wb["Calibration"]
ws2 = wb["QAQC"]

# Write Calibration and QAQC using shared logic
write_sample_data_by_id(ws1, std_names, element_data, element_columns)
write_sample_data_by_id(ws2, sample_names, qaqc_data, element_columns)

# Save output
wb.save(output_path)
print(f"Data written to {output_path}")