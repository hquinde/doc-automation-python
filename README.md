# DOC Automation with Python

This project automates the data processing workflow for Dissolved Organic Carbon (DOC) analysis, a crucial component in environmental and water quality research. The automation system replaces manual Excel macros with efficient, scalable Python scripts for data collection, cleaning, validation, and report generation.

## Project Overview

Traditional DOC data processing is time-consuming and error-prone due to reliance on manual Excel workflows. This tool provides a modular Python-based solution using:

- Pandas for structured data manipulation and validation
- OpenPyXL for automated, formatted Excel report generation

The system is designed with adaptability in mind and can be extended to automate workflows for other lab equipment.

## Features

- Automated Data Import: Handles DOC data from CSV/Excel files in various formats
- Data Cleaning and Validation: Identifies inconsistencies, fills missing values, and applies calibration rules
- Automated Reporting: Generates structured and formatted Excel reports using lab-specific templates
- Scalable Framework: Designed for reuse in other scientific workflows beyond DOC

## Project Structure

doc-automation-python/         ← Root folder (your GitHub repo)
│
├── data/                      ← Folder to store raw DOC data files (e.g., .csv, .xlsx)
├── scripts/                   ← Python scripts live here
│   ├── data_import.py         ← Script for importing and formatting raw data
│   ├── data_cleaning.py       ← Script to clean and validate the data
│   └── report_generator.py    ← Script to create formatted Excel reports
├── templates/                 ← Folder for reusable Excel or report templates
├── output/                    ← Folder where final reports get saved
└── README.md                  ← The README file you’re writing right now

## Dependencies

- Python 3.8+
- Pandas
- OpenPyXL

Install dependencies with:

```bash
pip install -r requirements.txt

```

## Usage

Run each script individually or combine them in a pipeline:

```bash
python scripts/data_import.py
python scripts/data_cleaning.py
python scripts/report_generator.py

```

