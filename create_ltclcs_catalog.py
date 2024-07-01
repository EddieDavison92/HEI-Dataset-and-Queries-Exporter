"""
This script processes dataset schemas and SQL scripts from CSV files
and exports the information into a structured Excel workbook.
The Excel workbook includes:
- An index sheet with links to dataset and script sheets.
- Separate sheets for each dataset schema.
- Separate sheets for each SQL script.
- A Tableau Calculated Fields section.
- A Refresh Instructions tab.

To run the script, simply execute it in your Python environment:
    python create_ltclcs_catalog.py

The output will be saved in the specified output folder as an Excel file.
"""

import os
from openpyxl import Workbook
from openpyxl.styles import NamedStyle, Font
from modules.readers import read_csv_schema, read_csv_scripts, read_csv_tableau, read_additional_v_catalog, read_all_scripts
from modules.excel_helpers import export_to_excel

# Constants for file paths and settings
CSV_FILE_TABLES = 'input/HEI_V_CATALOG_LTCLCS.csv'     # CSV file containing the dataset schema
CSV_FILE_SCRIPTS = 'input/HEI_LTCLCS_SCRIPTS.csv'      # CSV file containing the SQL scripts
CSV_FILE_TABLEAU = 'output/tableau.csv'                # CSV file containing the Tableau calculated fields
CSV_FILE_V_CATALOG = 'input/HEI_V_CATALOG.csv'         # CSV file containing V_CATALOG data
CSV_FILE_ALL_SCRIPTS = 'input/HEI_ALL_SCRIPTS.csv'     # CSV file containing all scripts
TABLEAU_WORKBOOK = 'LTC LCS Case Finding DEV V1.1.twb' # Constant for the workbook to filter
OUTPUT_FOLDER = './output'                             # Folder where the output will be saved
EXCEL_FILE_NAME = 'HEI_LTCLCS.xlsx'                    # Name of the Excel file to export
TABLE_STYLE = 'TableStyleLight8'                       # Excel table style to apply for data tables
INDEX_TABLE_STYLE = 'TableStyleLight8'                 # Excel table style to apply for index tables
FONT_NAME = 'Aptos'                                    # Font to be used globally

# Constants for header strings
HEADER_INSTRUCTIONS = "This document contains details of the datasets and scripts used to create the LTC LCS dashboard."
HEADER_NAVIGATION = "Click on the sheet names below to navigate to the respective sheet."
HEADER_DATASETS = "Datasets used in the LTC LCS following full dependency trace"
HEADER_SCRIPTS = "Scripts used to create each dataset in the LTC LCS Case Finding Workflow"
TABLEAU_HEADING = 'Tableau Calculated Fields'

def main():
    # Ensure output folder exists
    if not os.path.exists(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER)

    # Read data
    tables = read_csv_schema(CSV_FILE_TABLES)
    scripts = read_csv_scripts(CSV_FILE_SCRIPTS)
    tableau_fields = read_csv_tableau(CSV_FILE_TABLEAU, TABLEAU_WORKBOOK)
    additional_v_catalog = read_additional_v_catalog(CSV_FILE_V_CATALOG)
    all_scripts = read_all_scripts(CSV_FILE_ALL_SCRIPTS)

    # Define header style
    header_style = NamedStyle(name="header_style")
    header_style.font = Font(name=FONT_NAME, size=14, bold=True)

    # Export to Excel
    excel_file_path = os.path.join(OUTPUT_FOLDER, EXCEL_FILE_NAME)

    export_to_excel(tables, scripts, tableau_fields, additional_v_catalog, all_scripts, excel_file_path, header_style, INDEX_TABLE_STYLE, TABLE_STYLE, HEADER_INSTRUCTIONS, HEADER_NAVIGATION, HEADER_DATASETS, HEADER_SCRIPTS, TABLEAU_HEADING)

    print(f"Schema and scripts exported to {excel_file_path}")

if __name__ == "__main__":
    main()
