import os
from openpyxl.styles import NamedStyle, Font
from modules.readers import read_csv_schema, read_csv_scripts, read_csv_tableau, read_additional_v_catalog, read_all_scripts
from modules.excel_helpers import export_to_excel
import create_ltclcs_catalog

def main():
    # Ensure output folder exists
    if not os.path.exists(create_ltclcs_catalog.OUTPUT_FOLDER):
        os.makedirs(create_ltclcs_catalog.OUTPUT_FOLDER)

    # Read data
    tables = read_csv_schema(create_ltclcs_catalog.CSV_FILE_TABLES)
    scripts = read_csv_scripts(create_ltclcs_catalog.CSV_FILE_SCRIPTS)
    tableau_fields = read_csv_tableau(create_ltclcs_catalog.CSV_FILE_TABLEAU, create_ltclcs_catalog.TABLEAU_WORKBOOK)
    additional_v_catalog = read_additional_v_catalog(create_ltclcs_catalog.CSV_FILE_V_CATALOG)
    all_scripts = read_all_scripts(create_ltclcs_catalog.CSV_FILE_ALL_SCRIPTS)

    # Define header style
    header_style = NamedStyle(name="header_style")
    header_style.font = Font(name=create_ltclcs_catalog.FONT_NAME, size=14, bold=True)

    # Export to Excel
    excel_file_path = os.path.join(create_ltclcs_catalog.OUTPUT_FOLDER, create_ltclcs_catalog.EXCEL_FILE_NAME)
    export_to_excel(tables, scripts, tableau_fields, additional_v_catalog, all_scripts, excel_file_path, header_style, create_ltclcs_catalog.INDEX_TABLE_STYLE, create_ltclcs_catalog.TABLE_STYLE, create_ltclcs_catalog.HEADER_INSTRUCTIONS, create_ltclcs_catalog.HEADER_NAVIGATION, create_ltclcs_catalog.HEADER_DATASETS, create_ltclcs_catalog.HEADER_SCRIPTS, create_ltclcs_catalog.TABLEAU_HEADING)

    print(f"Schema and scripts exported to {excel_file_path}")

if __name__ == "__main__":
    main()
